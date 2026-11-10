#requires -Version 7.0
#requires -Modules Az.Accounts, Az.ResourceGraph

<#
.SYNOPSIS
Inventories Azure Front Door Classic, Standard, and Premium profiles.

.DESCRIPTION
Queries Azure Resource Graph across every enabled subscription in the current
Az context, or one subscription selected by name. The script counts endpoints,
links migrated Classic profiles to their Standard/Premium targets, and records
the Classic source on migrated Standard/Premium profiles. Scope is the current
tenant's enabled subscriptions, not every tenant associated with the account.

EndpointCount counts configured endpoints, including disabled or unused ones;
it does not measure traffic. Classic counts frontend endpoints (including custom
hostnames); Standard/Premium counts afdEndpoints, not custom-domain resources.
Missing or malformed Classic endpoint data terminates the inventory.

MigrationState is the business-facing status: ToMigrate or Migrated for Classic,
MigrationTarget for migrated Standard/Premium profiles, and AlreadyStandardPremium
for other Standard/Premium profiles. AzureResourceState preserves Azure's original
Enabled/Migrated/Active lifecycle values; they are not traffic or health indicators.
MigrationStatus separately distinguishes resolved links from counterparts not found in the
queried scope. Not found does not prove that the counterpart has been deleted:
scope, permissions, or Resource Graph indexing delays can also explain it.

MigrationSourceType distinguishes Front Door Classic from CDN profile migrations.
For Front Door Classic, MigrationDateUtc is the target's ARM createdTime (migration
preparation, not necessarily final commit). For CDN migrations it is MigratedOn;
the creation date of an in-place CDN resource is not a migration date.
Dates are also copied to linked Classic rows. Missing dates remain blank.

XLSX includes a Migration Statistics sheet with native Excel charts showing
monthly and cumulative migration counts. Prominent Classic progress counts use
only Classic resources that still exist in the queried scope. Migrated Classic
sources deleted after migration are excluded, so this is not an all-time total.
For monthly and cumulative charts only target rows are counted, once each,
with Front Door Classic and CDN series kept separate. Zero-activity months are
included; the current month is partial. These are dates of currently visible
targets, not a complete historical event log. Undated targets are disclosed and
excluded from timelines. CSV contains inventory rows only, with ISO 8601 UTC dates.

Exports are staged alongside the destination and published only after successful
generation. Empty inventories still produce column headers. Resource Graph
performs the cross-subscription query; continuation pages are fetched sequentially
to avoid unnecessary parallel requests and preserve fail-fast behavior. Creation
metadata is fetched once per relevant subscription (plus continuation pages),
rather than once per migration target.

Transient failures (HTTP 429/5xx, Resource Graph throttling) are retried with
backoff, honouring Retry-After; any other error still stops the run. More than
1,000 subscriptions are queried in chunks of 1,000, the Resource Graph limit.

.PARAMETER OutputFile
Destination file. Use a .csv extension for CSV or .xlsx for Excel. If .xlsx is
requested but ImportExcel is unavailable, the script writes a .csv instead.

.PARAMETER SubscriptionName
Optional exact subscription name. If omitted, all enabled subscriptions in the
current tenant are queried.

.PARAMETER PassThru
Also return the inventory rows to the pipeline. Console messages are written to
the host, so they never mix with the returned objects.

.PARAMETER SelfTest
Run the built-in offline checks of the migration, statistics, and retry logic
instead of querying Azure. No Azure login or output file is needed. Throws if
any check fails.

.EXAMPLE
.\Get-AzureFrontDoorInventory.ps1 -OutputFile .\FrontDoorInventory.xlsx

.EXAMPLE
.\Get-AzureFrontDoorInventory.ps1 -OutputFile .\FrontDoorInventory.csv `
    -SubscriptionName 'My Production Subscription'

.EXAMPLE
.\Get-AzureFrontDoorInventory.ps1 -SelfTest
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Console status must not mix with -PassThru output.')]
[CmdletBinding(DefaultParameterSetName = 'Inventory')]
param(
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Inventory')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        $extension = [System.IO.Path]::GetExtension($_)
        if ($extension -notin '.csv', '.xlsx') {
            throw "OutputFile must have a .csv or .xlsx extension."
        }

        $true
    })]
    [string] $OutputFile,

    [Parameter(ParameterSetName = 'Inventory')]
    [ValidateNotNullOrEmpty()]
    [string] $SubscriptionName,

    [Parameter(ParameterSetName = 'Inventory')]
    [switch] $PassThru,

    [Parameter(Mandatory, ParameterSetName = 'SelfTest')]
    [switch] $SelfTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Resource Graph accepts at most 1,000 subscriptions and returns at most 1,000 rows per page.
$resourceGraphLimit = 1000

# Keep both export formats, including header-only reports, on the same schema.
$columns = @(
    'Sku', 'ProfileName', 'SubscriptionName', 'SubscriptionId',
    'ResourceGroupName', 'EndpointCount', 'IsMigrated', 'MigrationStatus',
    'MigrationState', 'MigrationTargetProfileName', 'MigrationTargetResourceId',
    'ClassicSourceProfileName', 'ClassicSourceResourceId', 'ResourceId',
    'MigrationSourceType', 'MigrationSourceProfileName', 'MigrationSourceResourceId',
    'MigrationDateUtc', 'MigrationDateBasis', 'AzureResourceState'
)

# ARM resource IDs are case-insensitive. Normalize them before using them as
# lookup keys so migration links work despite casing differences in Azure data.
function Get-NormalizedResourceId {
    <#
    .SYNOPSIS
    Returns a lowercase, slash-trimmed ARM resource ID for use as a lookup key.
    .PARAMETER ResourceId
    ARM resource ID to normalize.
    .OUTPUTS
    System.String
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string] $ResourceId
    )

    return $ResourceId.TrimEnd('/').ToLowerInvariant()
}

# MigratedFrom is shared by AFD Classic and CDN migrations. Its resource type,
# not the presence of the property alone, determines which migration it describes.
function Get-MigrationSourceType {
    <#
    .SYNOPSIS
    Classifies a MigratedFrom resource ID as FrontDoorClassic, CdnProfile, or Other.
    .DESCRIPTION
    Returns $null when the ID is empty, meaning the profile is not a migration target.
    .PARAMETER ResourceId
    Value of extendedProperties.MigratedFrom.
    .OUTPUTS
    System.String, or $null.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string] $ResourceId
    )

    if ([string]::IsNullOrWhiteSpace($ResourceId)) {
        return $null
    }
    if ($ResourceId -match '^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/Microsoft.Network/frontDoors/[^/]+/?$') {
        return 'FrontDoorClassic'
    }
    if ($ResourceId -match '^/subscriptions/[^/]+/resourceGroups/[^/]+/providers/Microsoft.Cdn/profiles/[^/]+/?$') {
        return 'CdnProfile'
    }
    return 'Other'
}

# Normalize Azure JSON/SDK date representations to UTC. Missing metadata is
# allowed, but malformed values stop the run instead of fabricating a date.
function ConvertTo-MigrationUtcDate {
    <#
    .SYNOPSIS
    Converts an Azure date value to a UTC DateTime.
    .DESCRIPTION
    Returns $null for missing values and throws for malformed ones.
    .PARAMETER Value
    DateTime, DateTimeOffset, or ISO 8601 string.
    .PARAMETER ResourceId
    Resource ID used in the error message for malformed values.
    .OUTPUTS
    System.DateTime, or $null.
    #>
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [object] $Value,

        [Parameter(Mandatory)]
        [string] $ResourceId
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string] $Value)) {
        return $null
    }
    if ($Value -is [datetimeoffset]) {
        return $Value.UtcDateTime
    }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
            return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc)
        }
        return $Value.ToUniversalTime()
    }
    $parsed = [datetimeoffset]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
    if (-not [datetimeoffset]::TryParse(
        [string] $Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref] $parsed
    )) {
        throw "Invalid migration timestamp '$Value' on '$ResourceId'."
    }
    return $parsed.UtcDateTime
}

# Throttling and gateway errors are expected at tenant scale; everything else is
# a real failure and must keep the run fail-fast.
function Get-RetryDelay {
    <#
    .SYNOPSIS
    Calculates how long to wait before the next retry.
    .DESCRIPTION
    Honours a Retry-After header when the response carries one (clamped to 1-120
    seconds); otherwise uses exponential backoff capped at 30 seconds plus jitter.
    .PARAMETER Attempt
    One-based number of the attempt that just failed.
    .PARAMETER Response
    Optional HTTP response from Invoke-AzRestMethod.
    .OUTPUTS
    System.Double
    #>
    [CmdletBinding()]
    [OutputType([double])]
    param(
        [Parameter(Mandatory)]
        [int] $Attempt,

        [AllowNull()]
        [object] $Response
    )

    if ($null -ne $Response) {
        $headers = $Response.PSObject.Properties['Headers']
        $retryAfter = if ($headers -and $null -ne $headers.Value) { $headers.Value.PSObject.Properties['RetryAfter'] }
        if ($retryAfter -and $null -ne $retryAfter.Value) {
            $seconds = $null
            if ($null -ne $retryAfter.Value.Delta) {
                $seconds = $retryAfter.Value.Delta.TotalSeconds
            }
            elseif ($null -ne $retryAfter.Value.Date) {
                $seconds = ($retryAfter.Value.Date - [datetimeoffset]::UtcNow).TotalSeconds
            }
            if ($null -ne $seconds) {
                return [math]::Min(120.0, [math]::Max(1.0, $seconds))
            }
        }
    }
    return [math]::Min(30.0, [math]::Pow(2, $Attempt)) + (Get-Random -Minimum 0.0 -Maximum 1.0)
}

function Test-TransientError {
    <#
    .SYNOPSIS
    Returns $true when an error record describes throttling or a transient server failure.
    .PARAMETER ErrorRecord
    Error caught from an Az cmdlet.
    .OUTPUTS
    System.Boolean
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    return [bool] ($ErrorRecord.Exception.Message -match
        '\b(429|500|502|503|504)\b|TooManyRequests|throttl|ServiceUnavailable|GatewayTimeout|BadGateway')
}

function Invoke-WithRetry {
    <#
    .SYNOPSIS
    Runs a script block, retrying only transient failures.
    .DESCRIPTION
    Non-transient errors, and transient errors that persist past MaxAttempts, are
    rethrown unchanged so the run still stops at the first real failure.
    .PARAMETER ScriptBlock
    Operation to run.
    .PARAMETER Description
    Short label used in retry warnings.
    .PARAMETER MaxAttempts
    Maximum number of attempts, including the first.
    .OUTPUTS
    Whatever the script block returns.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock] $ScriptBlock,

        [Parameter(Mandatory)]
        [string] $Description,

        [ValidateRange(1, 10)]
        [int] $MaxAttempts = 5
    )

    for ($attempt = 1; ; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts -or -not (Test-TransientError -ErrorRecord $_)) {
                throw
            }
            $delay = Get-RetryDelay -Attempt $attempt -Response $null
            Write-Warning "$Description hit a transient error (attempt $attempt of $MaxAttempts): $($_.Exception.Message.TrimEnd('.')). Retrying in $([math]::Round($delay, 1))s."
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-AzRestMethodWithRetry {
    <#
    .SYNOPSIS
    Calls Invoke-AzRestMethod, retrying HTTP 429 and 5xx gateway/server responses.
    .DESCRIPTION
    Returns the final response, whatever its status, so callers keep their own
    status checks and error messages.
    .PARAMETER Request
    Splat table for Invoke-AzRestMethod.
    .PARAMETER Description
    Short label used in retry warnings.
    .PARAMETER MaxAttempts
    Maximum number of attempts, including the first.
    .OUTPUTS
    The Invoke-AzRestMethod response.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable] $Request,

        [Parameter(Mandatory)]
        [string] $Description,

        [ValidateRange(1, 10)]
        [int] $MaxAttempts = 5
    )

    for ($attempt = 1; ; $attempt++) {
        $response = Invoke-AzRestMethod @Request
        if ($response.StatusCode -notin 429, 500, 502, 503, 504 -or $attempt -ge $MaxAttempts) {
            return $response
        }
        $delay = Get-RetryDelay -Attempt $attempt -Response $response
        Write-Warning "$Description returned HTTP $($response.StatusCode) (attempt $attempt of $MaxAttempts). Retrying in $([math]::Round($delay, 1))s."
        Start-Sleep -Seconds $delay
    }
}

# The generic ARM resource list exposes createdTime even when systemData is absent
# from the Front Door provider. Batch by subscription and retain only target IDs.
function Get-FrontDoorCreationTime {
    <#
    .SYNOPSIS
    Fetches ARM createdTime for Front Door Classic migration targets.
    .DESCRIPTION
    Lists CDN profiles once per subscription that contains a target (following
    continuation pages) instead of making one call per target.
    .PARAMETER Profiles
    Resource Graph rows from Get-FrontDoorResource.
    .OUTPUTS
    System.Collections.Hashtable mapping normalized resource ID to a UTC DateTime.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Profiles
    )

    $creationTimes = @{}
    $targets = @($Profiles | Where-Object {
        $_.type -eq 'microsoft.cdn/profiles' -and
        (Get-MigrationSourceType -ResourceId $_.migratedFrom) -eq 'FrontDoorClassic'
    })
    foreach ($subscription in ($targets | Group-Object subscriptionId)) {
        $wantedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($target in $subscription.Group) {
            $null = $wantedIds.Add((Get-NormalizedResourceId -ResourceId $target.id))
        }
        # Do not add a name/resourceGroup filter: measured against a live tenant, ARM
        # evaluates it by scanning many empty pages, which was ~8x slower than this
        # single type-filtered listing. Resource Graph has no creation time either.
        $filter = [uri]::EscapeDataString("resourceType eq 'Microsoft.Cdn/profiles'")
        $nextLink = "/subscriptions/$($subscription.Name)/resources?api-version=2021-04-01" +
            '&$filter=' + $filter + '&$expand=createdTime'
        do {
            $request = @{ Method = 'GET' }
            if ($nextLink.StartsWith('/')) {
                $request.Path = $nextLink
            }
            else {
                $request.Uri = [uri] $nextLink
            }
            $response = Invoke-AzRestMethodWithRetry -Request $request -Description "Creation metadata request for subscription '$($subscription.Name)'"
            if ($response.StatusCode -ne 200) {
                throw "Creation metadata request for subscription '$($subscription.Name)' failed: HTTP $($response.StatusCode). $($response.Content)"
            }
            $body = $response.Content | ConvertFrom-Json -AsHashtable
            if (-not $body.ContainsKey('value')) {
                throw "Creation metadata response for '$($subscription.Name)' has no resource list."
            }
            foreach ($resource in $body['value']) {
                $id = Get-NormalizedResourceId -ResourceId $resource['id']
                if ($wantedIds.Contains($id)) {
                    $creationTimes[$id] = ConvertTo-MigrationUtcDate -Value $resource['createdTime'] -ResourceId $id
                }
            }
            $nextLink = $body['nextLink']
        } while ($nextLink)
    }
    return $creationTimes
}

# Resolve the actual format and check filesystem access before any Azure calls.
# Opening an existing report does not truncate it; it detects locks and read-only
# files early. A separate probe verifies permission to create staging files.
function Resolve-InventoryOutputPath {
    <#
    .SYNOPSIS
    Resolves the final output path and verifies it can be written.
    .DESCRIPTION
    Falls back from .xlsx to .csv when ImportExcel is unavailable, then checks for
    directory conflicts, locked or read-only existing files, and staging permission.
    .PARAMETER Path
    Requested output path.
    .OUTPUTS
    System.String. The absolute path that will actually be written.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $provider = $null
    $drive = $null
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(
        $Path, [ref] $provider, [ref] $drive
    )
    if ($provider.Name -ne 'FileSystem') {
        throw 'OutputFile must be a filesystem path.'
    }

    if ([System.IO.Path]::GetExtension($resolvedPath) -ieq '.xlsx') {
        if (Get-Module -ListAvailable ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
        }
        else {
            $resolvedPath = [System.IO.Path]::ChangeExtension($resolvedPath, '.csv')
            Write-Warning "ImportExcel is not available. Exporting CSV to '$resolvedPath' instead."
        }
    }

    if (Test-Path -LiteralPath $resolvedPath -PathType Container) {
        throw "OutputFile is a directory: '$resolvedPath'."
    }

    $directory = [System.IO.Path]::GetDirectoryName($resolvedPath)
    $null = [System.IO.Directory]::CreateDirectory($directory)
    if ([System.IO.File]::Exists($resolvedPath)) {
        $existing = [System.IO.File]::Open(
            $resolvedPath, [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Write, [System.IO.FileShare]::None
        )
        $existing.Dispose()
    }

    $probePath = Join-Path $directory ('.frontdoor-probe-' + [guid]::NewGuid().ToString('N'))
    $probe = [System.IO.File]::Open(
        $probePath, [System.IO.FileMode]::CreateNew,
        [System.IO.FileAccess]::Write, [System.IO.FileShare]::None
    )
    try {
        $probe.WriteByte(0)
    }
    finally {
        $probe.Dispose()
        [System.IO.File]::Delete($probePath)
    }

    return $resolvedPath
}

# Fetch profiles and endpoint child resources in one tenant-scale Resource Graph
# query. Server-side filtering is faster than making one ARM call per profile.
function Get-FrontDoorResource {
    <#
    .SYNOPSIS
    Retrieves Front Door profiles and Standard/Premium endpoints from Resource Graph.
    .DESCRIPTION
    Subscriptions are queried in chunks of at most 1,000 (the Resource Graph limit)
    and every continuation page is fetched sequentially. Throttled calls are retried.
    .PARAMETER SubscriptionId
    Subscription IDs to query.
    .OUTPUTS
    The Resource Graph rows, one object per profile or endpoint.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[object]])]
    param(
        [Parameter(Mandatory)]
        [string[]] $SubscriptionId
    )

    $query = @'
resources
| where type in~ (
    'microsoft.network/frontdoors',
    'microsoft.cdn/profiles',
    'microsoft.cdn/profiles/afdendpoints'
)
| where type !~ 'microsoft.cdn/profiles' or sku.name in~ (
    'Standard_AzureFrontDoor',
    'Premium_AzureFrontDoor'
)
| project id, name, type, subscriptionId, resourceGroup,
    skuName = tostring(sku.name),
    resourceState = tostring(properties.resourceState),
    migratedFrom = tostring(properties.extendedProperties.MigratedFrom),
    migratedOn = tostring(properties.extendedProperties.MigratedOn),
    classicEndpointCount = array_length(properties.frontendEndpoints)
| order by id asc
'@

    $resources = [System.Collections.Generic.List[object]]::new()

    # Resource Graph returns at most 1,000 rows per page. Continue until Azure
    # stops returning a skip token so large tenants are fully inventoried.
    for ($offset = 0; $offset -lt $SubscriptionId.Count; $offset += $resourceGraphLimit) {
        $chunk = @($SubscriptionId[$offset..([math]::Min($offset + $resourceGraphLimit, $SubscriptionId.Count) - 1)])
        $skipToken = $null
        do {
            $parameters = @{
                Query        = $query
                Subscription = $chunk
                First        = $resourceGraphLimit
            }
            if ($skipToken) {
                $parameters.SkipToken = $skipToken
            }

            $response = Invoke-WithRetry -Description 'Resource Graph query' -ScriptBlock {
                Search-AzGraph @parameters
            }
            foreach ($resource in $response.Data) {
                $resources.Add($resource)
            }

            $skipToken = $response.SkipToken
        } while ($skipToken)
    }

    return $resources
}

# Aggregate only visible migration targets, never their mirrored Classic rows.
# Missing dates are disclosed separately; zero-count calendar months are retained.
function Get-FrontDoorMigrationProgress {
    <#
    .SYNOPSIS
    Builds migration statistics from inventory rows.
    .DESCRIPTION
    Counts each dated Standard/Premium migration target once in a month-by-month
    timeline, and counts existing Classic profiles separately. The recent average
    divides by the complete months of the last three that the timeline actually
    covers, so a short history is not understated.
    .PARAMETER Inventory
    Rows produced by ConvertTo-FrontDoorInventory.
    .PARAMETER AsOfUtc
    Report timestamp; defaults to now. Later migration dates are rejected.
    .OUTPUTS
    System.Management.Automation.PSCustomObject
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Inventory,

        [datetime] $AsOfUtc = [datetime]::UtcNow
    )

    $AsOfUtc = ConvertTo-MigrationUtcDate -Value $AsOfUtc -ResourceId 'report-as-of'
    $reportMonth = [datetime]::new($AsOfUtc.Year, $AsOfUtc.Month, 1, 0, 0, 0, [DateTimeKind]::Utc)
    $targets = @($Inventory | Where-Object {
        $_.Sku -in 'Standard', 'Premium' -and $_.MigrationSourceResourceId -and
        $_.MigrationSourceType -in 'FrontDoorClassic', 'CdnProfile'
    })
    $datedTargets = @($targets | Where-Object { $null -ne $_.MigrationDateUtc })
    $byMonth = @{}
    $seenIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($target in $datedTargets) {
        if (-not $seenIds.Add((Get-NormalizedResourceId -ResourceId $target.ResourceId))) {
            throw "Duplicate migration target in statistics: '$($target.ResourceId)'."
        }
        $date = ConvertTo-MigrationUtcDate -Value $target.MigrationDateUtc -ResourceId $target.ResourceId
        if ($date -gt $AsOfUtc) {
            throw "Migration date for '$($target.ResourceId)' is later than the report timestamp."
        }
        $key = $date.ToString('yyyy-MM', [System.Globalization.CultureInfo]::InvariantCulture)
        if (-not $byMonth.ContainsKey($key)) {
            $byMonth[$key] = @{ FrontDoorClassic = 0; CdnProfile = 0 }
        }
        $byMonth[$key][$target.MigrationSourceType]++
    }

    # Fill from the earliest known date through the current (partial) UTC month.
    $monthly = [System.Collections.Generic.List[object]]::new()
    $cumulativeAfd = 0
    $cumulativeCdn = 0
    $lastThreeMonthsAfd = 0
    $averageMonths = 0
    if ($byMonth.Count -gt 0) {
        $firstKey = $byMonth.Keys | Sort-Object | Select-Object -First 1
        $month = [datetime]::ParseExact($firstKey, 'yyyy-MM', [System.Globalization.CultureInfo]::InvariantCulture)
        $month = [datetime]::SpecifyKind($month, [DateTimeKind]::Utc)
        while ($month -le $reportMonth) {
            $key = $month.ToString('yyyy-MM', [System.Globalization.CultureInfo]::InvariantCulture)
            $afdCount = 0
            $cdnCount = 0
            if ($byMonth.ContainsKey($key)) {
                $afdCount = $byMonth[$key].FrontDoorClassic
                $cdnCount = $byMonth[$key].CdnProfile
            }
            $cumulativeAfd += $afdCount
            $cumulativeCdn += $cdnCount
            if ($month -ge $reportMonth.AddMonths(-3) -and $month -lt $reportMonth) {
                $lastThreeMonthsAfd += $afdCount
                $averageMonths++
            }
            $monthly.Add([pscustomobject][ordered]@{
                Month                = $key
                AfdClassicMigrations = $afdCount
                CdnMigrations        = $cdnCount
                TotalMigrations      = $afdCount + $cdnCount
                CumulativeAfdClassic = $cumulativeAfd
                CumulativeCdn        = $cumulativeCdn
                CumulativeTotal      = $cumulativeAfd + $cumulativeCdn
                Period               = if ($month -eq $reportMonth) { 'Month to date' } else { 'Complete month' }
            })
            $month = $month.AddMonths(1)
        }
    }

    $afdTargets = @($targets | Where-Object MigrationSourceType -eq 'FrontDoorClassic')
    $cdnTargets = @($targets | Where-Object MigrationSourceType -eq 'CdnProfile')
    # Current Classic progress is a different population from the target timeline.
    # Once a migrated Classic source is deleted, it leaves these counts even if
    # its surviving Standard/Premium target is still represented in the charts.
    $existingClassic = @($Inventory | Where-Object Sku -eq 'Classic')
    [pscustomobject]@{
        AsOfUtc                      = $AsOfUtc
        Monthly                      = @($monthly)
        AfdMigrationTargets          = $afdTargets.Count
        CdnMigrationTargets          = $cdnTargets.Count
        DatedAfdTargets              = @($afdTargets | Where-Object { $null -ne $_.MigrationDateUtc }).Count
        DatedCdnTargets              = @($cdnTargets | Where-Object { $null -ne $_.MigrationDateUtc }).Count
        UndatedAfdTargets            = @($afdTargets | Where-Object { $null -eq $_.MigrationDateUtc }).Count
        UndatedCdnTargets            = @($cdnTargets | Where-Object { $null -eq $_.MigrationDateUtc }).Count
        OtherMigrationTargets       = @($Inventory | Where-Object { $_.Sku -ne 'Classic' -and $_.MigrationSourceType -eq 'Other' }).Count
        ExistingClassicProfiles      = $existingClassic.Count
        MigratedClassicProfiles      = @($existingClassic | Where-Object IsMigrated -eq $true).Count
        RemainingClassicProfiles     = @($existingClassic | Where-Object IsMigrated -eq $false).Count
        LastThreeCompleteMonthsAfd    = $lastThreeMonthsAfd
        AverageMonthsCovered          = $averageMonths
        LastThreeCompleteMonthAverage = if ($averageMonths -gt 0) { [math]::Round($lastThreeMonthsAfd / $averageMonths, 2) } else { 0 }
    }
}

# Add an XLSX-only dashboard with native, editable Excel charts. The table is
# placed below the charts, keeping chart ranges explicit and the sheet readable.
function Add-FrontDoorMigrationWorksheet {
    <#
    .SYNOPSIS
    Adds the "Migration Statistics" worksheet with count cards, metrics, and charts.
    .DESCRIPTION
    Vertical positions are derived from a few layout variables so adding a metric
    or note cannot silently overlap a chart or table.
    .PARAMETER Package
    Open ImportExcel/EPPlus package.
    .PARAMETER Inventory
    Rows produced by ConvertTo-FrontDoorInventory.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object] $Package,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Inventory
    )

    $statistics = Get-FrontDoorMigrationProgress -Inventory $Inventory
    # Sheet layout (1-based Excel rows). Charts use 0-based rows, hence the -1.
    $cardRow = 6
    $metricsRow = 16
    $chartSpacingRows = 24
    $firstChartRow = $metricsRow - 1
    $secondChartRow = $firstChartRow + $chartSpacingRows
    $tableRow = $secondChartRow + $chartSpacingRows + 1
    $sheet = $Package.Workbook.Worksheets.Add('Migration Statistics')
    $sheet.View.ShowGridLines = $false
    $sheet.View.FreezePanes(3, 1)
    # Put current Classic progress first; the target-based historical charts are
    # deliberately explained separately so their larger totals are not confused.
    $notes = [ordered]@{
        1 = 'Front Door migration statistics'
        2 = ('Snapshot (UTC): ' + $statistics.AsOfUtc.ToString('yyyy-MM-dd HH:mm:ss') + '. Scope: profiles in the inventory sheet.')
        3 = 'CLASSIC PROGRESS: counts below include only Classic profiles that still exist in the queried scope.'
        4 = 'Migrated Classic sources are gradually deleted. Deleted sources are excluded below; this is not an all-time migration total.'
        9 = 'AFD Classic chart dates: target ARM createdTime (migration preparation, not necessarily the final commit).'
        10 = 'CDN chart dates: extendedProperties.MigratedOn; original CDN creation dates are not used.'
        11 = 'Charts count each visible target once, even if its Classic source is absent. They are not a complete historical event log.'
        12 = 'Zero-activity months are included. The current month is partial; the recent average uses up to the last 3 complete months that the timeline covers.'
        13 = "Undated targets excluded from charts: AFD $($statistics.UndatedAfdTargets), CDN $($statistics.UndatedCdnTargets). Other source types excluded: $($statistics.OtherMigrationTargets)."
    }
    foreach ($note in $notes.GetEnumerator()) {
        $row = [int] $note.Key
        $sheet.Cells[$row, 1, $row, 8].Merge = $true
        $sheet.Cells[$row, 1].Value = $note.Value
        $sheet.Cells[$row, 1].Style.WrapText = $true
        $sheet.Row($row).Height = 30
    }
    $sheet.Cells[1, 1].Style.Font.Size = 20
    $sheet.Cells[1, 1].Style.Font.Bold = $true
    $sheet.Cells[1, 1].Style.Fill.PatternType = 'Solid'
    $sheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31, 78, 121))
    $sheet.Cells[1, 1].Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $sheet.Column(1).Width = 38
    for ($column = 2; $column -le 8; $column++) { $sheet.Column($column).Width = 23 }

    # Large, colored count cards use existing Classic rows only, including migrated
    # rows whose target cannot currently be resolved. No target-derived denominator.
    $cards = @(
        @{ Column = 1; Label = 'Migrated Classic (still existing)'; Value = $statistics.MigratedClassicProfiles; Color = [System.Drawing.Color]::FromArgb(226, 239, 218) }
        @{ Column = 3; Label = 'Classic still to migrate'; Value = $statistics.RemainingClassicProfiles; Color = [System.Drawing.Color]::FromArgb(255, 235, 156) }
        @{ Column = 5; Label = 'Total existing Classic profiles'; Value = $statistics.ExistingClassicProfiles; Color = [System.Drawing.Color]::FromArgb(221, 235, 247) }
    )
    foreach ($card in $cards) {
        $firstColumn = [int] $card.Column
        $lastColumn = $firstColumn + 1
        $sheet.Cells[$cardRow, $firstColumn, $cardRow, $lastColumn].Merge = $true
        $sheet.Cells[$cardRow, $firstColumn].Value = $card.Label
        $sheet.Cells[($cardRow + 1), $firstColumn, ($cardRow + 2), $lastColumn].Merge = $true
        $sheet.Cells[($cardRow + 1), $firstColumn].Value = $card.Value
        $range = $sheet.Cells[$cardRow, $firstColumn, ($cardRow + 2), $lastColumn]
        $range.Style.Fill.PatternType = 'Solid'
        $range.Style.Fill.BackgroundColor.SetColor($card.Color)
        $range.Style.Font.Bold = $true
        $range.Style.HorizontalAlignment = 'Center'
        $range.Style.VerticalAlignment = 'Center'
        $range.Style.WrapText = $true
        $sheet.Cells[($cardRow + 1), $firstColumn].Style.Font.Size = 30
        $sheet.Cells[($cardRow + 1), $firstColumn].Style.Numberformat.Format = '0'
    }
    $sheet.Row($cardRow).Height = 30
    $sheet.Row($cardRow + 1).Height = 28
    $sheet.Row($cardRow + 2).Height = 20

    # Coverage metrics keep missing dates and remaining Classic work visible instead
    # of quietly dropping them from the denominator of a completion percentage.
    $metrics = @(
        [pscustomobject]@{ Metric = 'Visible targets from AFD Classic'; Value = $statistics.AfdMigrationTargets }
        [pscustomobject]@{ Metric = 'CDN migration targets'; Value = $statistics.CdnMigrationTargets }
        [pscustomobject]@{ Metric = 'Dated AFD targets'; Value = $statistics.DatedAfdTargets }
        [pscustomobject]@{ Metric = 'Undated AFD targets'; Value = $statistics.UndatedAfdTargets }
        [pscustomobject]@{ Metric = 'Dated CDN targets'; Value = $statistics.DatedCdnTargets }
        [pscustomobject]@{ Metric = 'Undated CDN targets'; Value = $statistics.UndatedCdnTargets }
        [pscustomobject]@{ Metric = 'Existing Classic: migrated'; Value = $statistics.MigratedClassicProfiles }
        [pscustomobject]@{ Metric = 'Existing Classic: to migrate'; Value = $statistics.RemainingClassicProfiles }
        [pscustomobject]@{ Metric = 'AFD: last 3 complete months'; Value = $statistics.LastThreeCompleteMonthsAfd }
        [pscustomobject]@{ Metric = 'AFD: months in average (max 3)'; Value = $statistics.AverageMonthsCovered }
        [pscustomobject]@{ Metric = 'AFD: average/month (recent)'; Value = $statistics.LastThreeCompleteMonthAverage }
    )
    $lastMetricRow = $metricsRow + $metrics.Count
    $null = $metrics | Export-Excel -ExcelPackage $Package -WorksheetName $sheet.Name `
        -StartRow $metricsRow -TableName 'MigrationMetrics' -TableStyle 'Medium2' -PassThru
    $sheet.Cells[($metricsRow + 1), 1, $lastMetricRow, 1].Style.WrapText = $true
    for ($row = $metricsRow + 1; $row -le $lastMetricRow; $row++) { $sheet.Row($row).Height = 30 }

    if ($statistics.Monthly.Count -eq 0) {
        $emptyRow = $lastMetricRow + 3
        $sheet.Cells[$emptyRow, 1, $emptyRow, 8].Merge = $true
        $sheet.Cells[$emptyRow, 1].Value = 'No dated migrations are available; no timeline or charts have been fabricated.'
        return
    }

    # Native chart series reference numeric table cells; they are not screenshots.
    # Month labels are ISO strings, avoiding locale-dependent date-axis formatting.
    $firstDataRow = $tableRow + 1
    $lastDataRow = $tableRow + $statistics.Monthly.Count
    $null = $statistics.Monthly | Export-Excel -ExcelPackage $Package -WorksheetName $sheet.Name `
        -StartRow $tableRow -TableName 'MonthlyMigrations' -TableStyle 'Medium2' `
        -NoNumberConversion Month, Period -PassThru
    $sheet.Cells[$tableRow, 1, $tableRow, 8].Style.WrapText = $true
    $sheet.Row($tableRow).Height = 32

    $chartParameters = @{
        Worksheet      = $sheet
        XRange         = "A${firstDataRow}:A${lastDataRow}"
        SeriesHeader   = @('AFD Classic', 'CDN')
        Width          = 960
        Height         = 340
        Column         = 2
        LegendPosition = 'Bottom'
        TitleBold      = $true
        TitleSize      = 14
        XAxisTitleText = 'Month (UTC); current month is partial'
        YAxisNumberformat = '0'
        YMinValue      = 0
    }
    $null = Add-ExcelChart @chartParameters -ChartType ColumnStacked -Row $firstChartRow `
        -Title 'Migration targets by month' -YAxisTitleText 'Targets per month' `
        -YRange @("B${firstDataRow}:B${lastDataRow}", "C${firstDataRow}:C${lastDataRow}")
    $null = Add-ExcelChart @chartParameters -ChartType Line -Row $secondChartRow `
        -Title 'Cumulative migration targets' -YAxisTitleText 'Cumulative targets' `
        -YRange @("E${firstDataRow}:E${lastDataRow}", "F${firstDataRow}:F${lastDataRow}")
}

# Produce the inventory table plus the statistics dashboard. Always dispose the
# package, including when formatting, chart creation, or saving fails.
function Export-FrontDoorExcel {
    <#
    .SYNOPSIS
    Writes the inventory table and the Migration Statistics worksheet to an XLSX file.
    .PARAMETER Inventory
    Rows produced by ConvertTo-FrontDoorInventory; may be empty.
    .PARAMETER Path
    XLSX file to create.
    .PARAMETER Columns
    Ordered column names for the inventory sheet.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Inventory,

        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [string[]] $Columns
    )

    Import-Module ImportExcel -ErrorAction Stop

    $package = Open-ExcelPackage -Path $Path -Create
    try {
        if ($Inventory.Count -gt 0) {
            $null = $Inventory |
                Select-Object -Property $Columns |
                Export-Excel -ExcelPackage $package `
                    -WorksheetName 'FrontDoorProfiles' `
                    -TableName 'FrontDoorProfiles' `
                    -TableStyle 'Medium2' `
                    -NoNumberConversion ($Columns | Where-Object { $_ -notin 'EndpointCount', 'IsMigrated', 'MigrationDateUtc' }) `
                    -NoHyperLinkConversion '*' `
                    -BoldTopRow `
                    -FreezeTopRow `
                    -PassThru
        }
        else {
            $sheet = $package.Workbook.Worksheets.Add('FrontDoorProfiles')
            for ($index = 0; $index -lt $Columns.Count; $index++) {
                $sheet.Cells[1, ($index + 1)].Value = $Columns[$index]
            }
            $sheet.Cells[1, 1, 1, $Columns.Count].Style.Font.Bold = $true
            $sheet.View.FreezePanes(2, 1)
        }

        # Cap long ARM ID columns instead of making the workbook excessively wide.
        $sheet = $package.Workbook.Worksheets['FrontDoorProfiles']
        $sheet.Cells[$sheet.Dimension.Address].AutoFitColumns(12, 55)
        if ($Inventory.Count -gt 0) {
            # Excel stores a numeric date, explicitly displayed as UTC; CSV uses ISO.
            $dateColumn = [array]::IndexOf($Columns, 'MigrationDateUtc') + 1
            $sheet.Cells[2, $dateColumn, ($Inventory.Count + 1), $dateColumn].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm:ss "UTC"'
            $sheet.Column($dateColumn).Width = 27
        }
        Add-FrontDoorMigrationWorksheet -Package $package -Inventory $Inventory
        $package.Save()
    }
    finally {
        $package.Dispose()
    }
}

# Stage on the same filesystem, then replace the previous report in one operation.
# Export/publish errors propagate; finally removes only this invocation's stage.
function Export-FrontDoorReport {
    <#
    .SYNOPSIS
    Writes the inventory to a CSV or XLSX file via a staged, atomically published copy.
    .PARAMETER Inventory
    Rows produced by ConvertTo-FrontDoorInventory; may be empty.
    .PARAMETER Path
    Final destination; its extension selects the format.
    .PARAMETER Columns
    Ordered column names shared by both formats.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Inventory,

        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [string[]] $Columns
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $stagePath = Join-Path ([System.IO.Path]::GetDirectoryName($Path)) (
        '.frontdoor-stage-' + [guid]::NewGuid().ToString('N') + $extension
    )
    try {
        if ($extension -eq '.xlsx') {
            Export-FrontDoorExcel -Inventory $Inventory -Path $stagePath -Columns $Columns
        }
        elseif ($extension -eq '.csv') {
            if ($Inventory.Count -gt 0) {
                # Keep flat CSV dates portable instead of using the host's culture.
                $csvColumns = foreach ($column in $Columns) {
                    if ($column -eq 'MigrationDateUtc') {
                        @{
                            Name = $column
                            Expression = {
                                if ($null -ne $_.MigrationDateUtc) {
                                    $_.MigrationDateUtc.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", [System.Globalization.CultureInfo]::InvariantCulture)
                                }
                            }
                        }
                    }
                    else { $column }
                }
                $Inventory | Select-Object -Property $csvColumns |
                    Export-Csv -LiteralPath $stagePath -NoTypeInformation -Encoding utf8BOM -UseQuotes AsNeeded
            }
            else {
                # Serialize a schema-only object and retain its header, not a fake row.
                $emptyRow = [ordered]@{}
                foreach ($column in $Columns) {
                    $emptyRow[$column] = $null
                }
                $header = [pscustomobject] $emptyRow | ConvertTo-Csv -NoTypeInformation -UseQuotes AsNeeded
                Set-Content -LiteralPath $stagePath -Value $header[0] -Encoding utf8BOM
            }
        }
        else {
            throw "Unsupported output extension '$extension'."
        }

        if ([System.IO.File]::Exists($Path)) {
            # Pass a true .NET null backup path; PowerShell converts $null to "".
            [System.IO.File]::Replace($stagePath, $Path, [NullString]::Value)
        }
        else {
            [System.IO.File]::Move($stagePath, $Path)
        }
    }
    finally {
        if ([System.IO.File]::Exists($stagePath)) {
            [System.IO.File]::Delete($stagePath)
        }
    }
}

# Transform the graph snapshot locally. Keeping this separate from Azure calls
# and file I/O allows empty and incomplete snapshots to be tested offline.
function ConvertTo-FrontDoorInventory {
    <#
    .SYNOPSIS
    Converts Resource Graph rows into inventory rows with migration links.
    .DESCRIPTION
    Counts endpoints per profile, links Classic profiles to their migrated
    Standard/Premium targets by MigratedFrom, and assigns migration dates.
    Conflicting links and malformed data throw instead of being guessed.
    .PARAMETER Resources
    Rows from Get-FrontDoorResource.
    .PARAMETER Subscriptions
    Az subscriptions (Id, Name) used to resolve subscription names.
    .PARAMETER CreationTimes
    Normalized target resource ID to createdTime, from Get-FrontDoorCreationTime.
    .OUTPUTS
    Inventory rows sorted by Sku, subscription, resource group, and profile name.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Resources,

        [Parameter(Mandatory)]
        [object[]] $Subscriptions,

        [Parameter()]
        [hashtable] $CreationTimes = @{}
    )

    # Cache subscription names by ID to avoid repeated Az calls while creating rows.
    $subscriptionNames = @{}
    foreach ($subscription in $Subscriptions) {
        $subscriptionNames[$subscription.Id.ToLowerInvariant()] = $subscription.Name
    }

    $classicProfiles = @($Resources | Where-Object type -EQ 'microsoft.network/frontdoors')
    $modernProfiles = @($Resources | Where-Object type -EQ 'microsoft.cdn/profiles')
    $modernEndpoints = @($Resources | Where-Object type -EQ 'microsoft.cdn/profiles/afdendpoints')

    # Standard/Premium endpoints are child ARM resources. Aggregate them locally
    # by parent profile ID; both endpoint types count configuration, not traffic.
    $endpointCountByProfileId = @{}
    foreach ($endpoint in $modernEndpoints) {
        $profileId = $endpoint.id -replace '(?i)/afdEndpoints/[^/]+$', ''
        $normalizedProfileId = Get-NormalizedResourceId -ResourceId $profileId
        if (-not $endpointCountByProfileId.ContainsKey($normalizedProfileId)) {
            $endpointCountByProfileId[$normalizedProfileId] = 0
        }
        $endpointCountByProfileId[$normalizedProfileId]++
    }

    # Invert MigratedFrom to find each surviving Classic profile's target without
    # relying on profile names. Conflicting links are errors, not arbitrary picks.
    $migrationTargetByClassicId = @{}
    foreach ($modernProfile in $modernProfiles) {
        $migratedFrom = [string] $modernProfile.migratedFrom
        if ((Get-MigrationSourceType -ResourceId $migratedFrom) -ne 'FrontDoorClassic') {
            continue
        }
        $sourceId = Get-NormalizedResourceId -ResourceId $migratedFrom
        if ($migrationTargetByClassicId.ContainsKey($sourceId)) {
            throw "Multiple migration targets reference Classic profile '$migratedFrom'."
        }
        $migrationTargetByClassicId[$sourceId] = $modernProfile
    }

    $inventory = [System.Collections.Generic.List[object]]::new()
    $classicProfileIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($classicProfile in $classicProfiles) {
        $null = $classicProfileIds.Add((Get-NormalizedResourceId -ResourceId $classicProfile.id))
    }

    # IsMigrated is true when Azure marks the Classic profile Migrated or a
    # target explicitly references it. A missing target is reported separately.
    foreach ($classicProfile in $classicProfiles) {
        $normalizedId = Get-NormalizedResourceId -ResourceId $classicProfile.id
        $target = $migrationTargetByClassicId[$normalizedId]
        $resourceState = $classicProfile.resourceState
        # array_length returns null for absent or non-array data. Never turn
        # unknown metadata into a count; a real empty array returns zero.
        if ($null -eq $classicProfile.classicEndpointCount) {
            throw "Classic endpoint count is unavailable for '$($classicProfile.id)'; frontendEndpoints is missing or not an array."
        }
        $isMigrated = $resourceState -eq 'Migrated' -or $null -ne $target
        $migrationStatus = if ($target) {
            'TargetResolved'
        }
        elseif ($isMigrated) {
            'TargetNotFoundInScope'
        }
        else {
            'NotMigrated'
        }

        # Classic rows mirror the resolved target's date for filtering in either
        # direction. Statistics exclude these source rows to prevent double counting.
        $migrationDate = if ($target) { $CreationTimes[(Get-NormalizedResourceId -ResourceId $target.id)] } else { $null }
        $dateBasis = if (-not $isMigrated) {
            'NotApplicable'
        }
        elseif ($null -ne $migrationDate) {
            'TargetCreatedTime'
        }
        else {
            'Unavailable'
        }

        $inventory.Add([pscustomobject][ordered]@{
            Sku                         = 'Classic'
            ProfileName                 = $classicProfile.name
            SubscriptionName            = $subscriptionNames[$classicProfile.subscriptionId.ToLowerInvariant()]
            SubscriptionId              = $classicProfile.subscriptionId
            ResourceGroupName           = $classicProfile.resourceGroup
            EndpointCount               = [int] $classicProfile.classicEndpointCount
            IsMigrated                  = [bool] $isMigrated
            MigrationStatus             = $migrationStatus
            # A business label is clearer than renaming an Azure lifecycle value.
            MigrationState              = if ($isMigrated) { 'Migrated' } else { 'ToMigrate' }
            MigrationTargetProfileName  = if ($target) { $target.name } else { $null }
            MigrationTargetResourceId   = if ($target) { $target.id } else { $null }
            ClassicSourceProfileName    = $null
            ClassicSourceResourceId     = $null
            ResourceId                  = $classicProfile.id
            MigrationSourceType         = if ($isMigrated) { 'FrontDoorClassic' } else { $null }
            MigrationSourceProfileName  = $null
            MigrationSourceResourceId   = $null
            MigrationDateUtc            = $migrationDate
            MigrationDateBasis          = $dateBasis
            AzureResourceState          = $resourceState
        })
    }

    # A retained source ID is useful even when its Classic resource is absent
    # from the snapshot. Do not label absence as deletion or discard the ID.
    foreach ($modernProfile in $modernProfiles) {
        $sku = switch ($modernProfile.skuName) {
            'Standard_AzureFrontDoor' { 'Standard' }
            'Premium_AzureFrontDoor'  { 'Premium' }
            default { throw "Unexpected Azure Front Door SKU '$($modernProfile.skuName)' on '$($modernProfile.id)'." }
        }

        $normalizedId = Get-NormalizedResourceId -ResourceId $modernProfile.id
        $migratedFrom = [string] $modernProfile.migratedFrom
        $sourceType = Get-MigrationSourceType -ResourceId $migratedFrom
        $sourceName = if ([string]::IsNullOrWhiteSpace($migratedFrom)) {
            $null
        }
        else {
            ($migratedFrom.TrimEnd('/') -split '/')[-1]
        }
        $migrationStatus = if (-not $sourceName) {
            'NotMigrationTarget'
        }
        elseif ($sourceType -eq 'CdnProfile') {
            'CdnMigration'
        }
        elseif ($sourceType -eq 'Other') {
            'OtherMigrationSource'
        }
        elseif ($classicProfileIds.Contains((Get-NormalizedResourceId -ResourceId $migratedFrom))) {
            'SourceResolved'
        }
        else {
            'SourceNotFoundInScope'
        }

        # Only newly created AFD migration targets use createdTime. CDN migration
        # may reuse its resource ID, so its older creation time must never be used.
        $migrationDate = $null
        $dateBasis = if ($sourceType) { 'Unavailable' } else { 'NotApplicable' }
        if ($sourceType -eq 'FrontDoorClassic') {
            $migrationDate = $CreationTimes[$normalizedId]
            if ($null -ne $migrationDate) { $dateBasis = 'TargetCreatedTime' }
        }
        elseif ($sourceType -eq 'CdnProfile') {
            $migrationDate = ConvertTo-MigrationUtcDate -Value $modernProfile.migratedOn -ResourceId $modernProfile.id
            if ($null -ne $migrationDate) { $dateBasis = 'MigratedOn' }
        }

        $inventory.Add([pscustomobject][ordered]@{
            Sku                         = $sku
            ProfileName                 = $modernProfile.name
            SubscriptionName            = $subscriptionNames[$modernProfile.subscriptionId.ToLowerInvariant()]
            SubscriptionId              = $modernProfile.subscriptionId
            ResourceGroupName           = $modernProfile.resourceGroup
            EndpointCount               = if ($endpointCountByProfileId.ContainsKey($normalizedId)) {
                $endpointCountByProfileId[$normalizedId]
            }
            else {
                0
            }
            IsMigrated                  = $null
            MigrationStatus             = $migrationStatus
            MigrationState              = if ($sourceType) { 'MigrationTarget' } else { 'AlreadyStandardPremium' }
            MigrationTargetProfileName  = $null
            MigrationTargetResourceId   = $null
            ClassicSourceProfileName    = if ($sourceType -eq 'FrontDoorClassic') { $sourceName } else { $null }
            ClassicSourceResourceId     = if ($sourceType -eq 'FrontDoorClassic') { $migratedFrom } else { $null }
            ResourceId                  = $modernProfile.id
            MigrationSourceType         = $sourceType
            MigrationSourceProfileName  = $sourceName
            MigrationSourceResourceId   = if ($sourceType) { $migratedFrom } else { $null }
            MigrationDateUtc            = $migrationDate
            MigrationDateBasis          = $dateBasis
            AzureResourceState          = $modernProfile.resourceState
        })
    }

    # Keep output deterministic and easy to scan across repeated executions.
    $inventory | Sort-Object Sku, SubscriptionName, ResourceGroupName, ProfileName
}

# Emit all three SKU totals, even when a SKU (or the entire inventory) is empty.
# Explicit integer accumulators avoid missing Measure-Object properties in strict mode.
function Get-FrontDoorSummary {
    <#
    .SYNOPSIS
    Totals profiles, endpoints, and migrations for each SKU.
    .PARAMETER Inventory
    Rows produced by ConvertTo-FrontDoorInventory; may be empty.
    .OUTPUTS
    One object each for Classic, Standard, and Premium.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]] $Inventory
    )

    foreach ($sku in 'Classic', 'Standard', 'Premium') {
        $group = @($Inventory | Where-Object Sku -eq $sku)
        $endpointTotal = 0
        $migrationTotal = 0
        foreach ($row in $group) {
            $endpointTotal += $row.EndpointCount
            if ($row.IsMigrated -eq $true -or
                -not [string]::IsNullOrWhiteSpace([string] $row.MigrationSourceResourceId)) {
                $migrationTotal++
            }
        }
        [pscustomobject]@{
            Sku            = $sku
            ProfileCount   = $group.Count
            EndpointCount  = $endpointTotal
            MigrationCount = $migrationTotal
        }
    }
}

# Offline checks of the pure logic. Azure cmdlets and Start-Sleep are shadowed by
# local functions, so no login, network, or real delay is involved.
function Invoke-SelfTest {
    <#
    .SYNOPSIS
    Runs built-in offline checks and throws if any fail.
    .DESCRIPTION
    Covers ID classification, date parsing, migration linking, statistics (including
    the recent-average denominator), retry behavior, and subscription chunking.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Intentional test double for the real delay.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test double.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Test doubles mirror the real cmdlet signatures.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPositionalParameters', '', Justification = 'Compact assertions.')]
    [CmdletBinding()]
    param()

    $failures = [System.Collections.Generic.List[string]]::new()
    # Retry warnings are expected here and would only add noise.
    $WarningPreference = 'SilentlyContinue'
    $state = @{ Sleeps = 0; RestCalls = 0; Attempts = 0;     Chunks = [System.Collections.Generic.List[int]]::new() }

    function Assert-Equal($Name, $Actual, $Expected) {
        if ("$Actual" -ne "$Expected") { $failures.Add("$Name : expected '$Expected', got '$Actual'") }
    }
    function Assert-ThrowError($Name, [scriptblock] $Action, $Pattern) {
        try { & $Action }
        catch {
            if ($_.Exception.Message -notmatch $Pattern) { $failures.Add("$Name : unexpected error '$($_.Exception.Message)'") }
            return
        }
        $failures.Add("$Name : expected an error")
    }
    # Shadow the real delay and Azure calls (functions outrank cmdlets).
    function Start-Sleep { param($Seconds) $state.Sleeps++ }
    function Invoke-AzRestMethod {
        param($Method, $Path, $Uri)
        $state.RestCalls++
        if ($state.RestCalls -lt 3) { return [pscustomobject]@{ StatusCode = 429; Content = '' } }
        [pscustomobject]@{ StatusCode = 200; Content = '{}' }
    }
    function Search-AzGraph {
        param($Query, $Subscription, $First, $SkipToken)
        $state.Chunks.Add(@($Subscription).Count)
        [pscustomobject]@{ Data = @([pscustomobject]@{ id = 'x' }); SkipToken = $null }
    }

    $classicA = '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Network/frontDoors/classicA'
    $classicB = '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Network/frontDoors/classicB'
    $targetA = '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Cdn/profiles/targetA'
    $cdnTarget = '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Cdn/profiles/cdnTarget'
    $plain = '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Cdn/profiles/plain'

    #region Classification and dates
    Assert-Equal 'source type classic' (Get-MigrationSourceType -ResourceId $classicA) 'FrontDoorClassic'
    Assert-Equal 'source type cdn' (Get-MigrationSourceType -ResourceId $targetA) 'CdnProfile'
    Assert-Equal 'source type other' (Get-MigrationSourceType -ResourceId '/subscriptions/s1/foo') 'Other'
    Assert-Equal 'source type empty' ($null -eq (Get-MigrationSourceType -ResourceId '')) 'True'
    Assert-Equal 'utc date' ((ConvertTo-MigrationUtcDate -Value '2026-10-05T10:00:00+02:00' -ResourceId 'r').ToString('o')) '2026-10-05T08:00:00.0000000Z'
    Assert-ThrowError 'bad date' { ConvertTo-MigrationUtcDate -Value 'not a date' -ResourceId 'r' } 'Invalid migration timestamp'
    #endregion

    #region Inventory conversion
    $row = { param($id, $name, $type, $sku, $from, $on, $classicCount, $resourceState)
        [pscustomobject]@{
            id = $id; name = $name; type = $type; subscriptionId = 's1'; resourceGroup = 'rg'
            skuName = $sku; resourceState = $resourceState; migratedFrom = $from; migratedOn = $on
            classicEndpointCount = $classicCount
        }
    }
    $resources = @(
        (& $row $classicA 'classicA' 'microsoft.network/frontdoors' '' '' '' 3 'Migrated')
        (& $row $classicB 'classicB' 'microsoft.network/frontdoors' '' '' '' 0 'Enabled')
        (& $row $targetA 'targetA' 'microsoft.cdn/profiles' 'Premium_AzureFrontDoor' $classicA '' $null 'Active')
        (& $row $cdnTarget 'cdnTarget' 'microsoft.cdn/profiles' 'Standard_AzureFrontDoor' '/subscriptions/s1/resourceGroups/rg/providers/Microsoft.Cdn/profiles/old' '2026-10-20T00:00:00Z' $null 'Active')
        (& $row $plain 'plain' 'microsoft.cdn/profiles' 'Standard_AzureFrontDoor' '' '' $null 'Active')
        (& $row "$targetA/afdEndpoints/e1" 'e1' 'microsoft.cdn/profiles/afdendpoints' '' '' '' $null '')
        (& $row "$targetA/afdEndpoints/e2" 'e2' 'microsoft.cdn/profiles/afdendpoints' '' '' '' $null '')
    )
    $subscriptions = @([pscustomobject]@{ Id = 's1'; Name = 'Sub One' })
    $created = @{ (Get-NormalizedResourceId -ResourceId $targetA) = [datetime]::new(2026, 10, 5, 0, 0, 0, [DateTimeKind]::Utc) }
    $rows = @(ConvertTo-FrontDoorInventory -Resources $resources -Subscriptions $subscriptions -CreationTimes $created)
    $byName = @{}
    foreach ($item in $rows) { $byName[$item.ProfileName] = $item }
    Assert-Equal 'row count' $rows.Count 5
    Assert-Equal 'classicA status' $byName['classicA'].MigrationStatus 'TargetResolved'
    Assert-Equal 'classicA endpoints' $byName['classicA'].EndpointCount 3
    Assert-Equal 'classicA date basis' $byName['classicA'].MigrationDateBasis 'TargetCreatedTime'
    Assert-Equal 'classicA mirrors date' $byName['classicA'].MigrationDateUtc.Day 5
    Assert-Equal 'classicB state' $byName['classicB'].MigrationState 'ToMigrate'
    Assert-Equal 'classicB basis' $byName['classicB'].MigrationDateBasis 'NotApplicable'
    Assert-Equal 'targetA endpoints' $byName['targetA'].EndpointCount 2
    Assert-Equal 'targetA source' $byName['targetA'].ClassicSourceProfileName 'classicA'
    Assert-Equal 'cdn basis' $byName['cdnTarget'].MigrationDateBasis 'MigratedOn'
    Assert-Equal 'plain state' $byName['plain'].MigrationState 'AlreadyStandardPremium'
    $dupe = $resources + (& $row "$targetA-2" 'targetA2' 'microsoft.cdn/profiles' 'Premium_AzureFrontDoor' $classicA '' $null 'Active')
    Assert-ThrowError 'duplicate targets' { ConvertTo-FrontDoorInventory -Resources $dupe -Subscriptions $subscriptions } 'Multiple migration targets'
    Assert-Equal 'empty inventory' @(ConvertTo-FrontDoorInventory -Resources @() -Subscriptions $subscriptions).Count 0
    #endregion

    #region Statistics
    $asOf = [datetime]::new(2026, 11, 10, 0, 0, 0, [DateTimeKind]::Utc)
    $stats = Get-FrontDoorMigrationProgress -Inventory $rows -AsOfUtc $asOf
    Assert-Equal 'stats months' $stats.Monthly.Count 2
    Assert-Equal 'stats classic existing' $stats.ExistingClassicProfiles 2
    # One AFD migration in the only complete month covered: average is 1, not 1/3.
    Assert-Equal 'average months' $stats.AverageMonthsCovered 1
    Assert-Equal 'average' $stats.LastThreeCompleteMonthAverage 1
    $old = [pscustomobject]@{
        Sku = 'Premium'; ResourceId = $targetA; MigrationSourceResourceId = $classicA
        MigrationSourceType = 'FrontDoorClassic'; MigrationDateUtc = [datetime]::new(2026, 8, 3, 0, 0, 0, [DateTimeKind]::Utc)
    }
    $stats = Get-FrontDoorMigrationProgress -Inventory @($old) -AsOfUtc $asOf
    Assert-Equal 'gap months' $stats.Monthly.Count 4
    Assert-Equal 'full window average' $stats.LastThreeCompleteMonthAverage 0.33
    Assert-Equal 'empty stats' (Get-FrontDoorMigrationProgress -Inventory @() -AsOfUtc $asOf).Monthly.Count 0
    #endregion

    #region Retry and chunking
    Assert-Equal 'retry-after header' (Get-RetryDelay -Attempt 1 -Response ([pscustomobject]@{
        Headers = [pscustomobject]@{ RetryAfter = [pscustomobject]@{ Delta = [timespan]::FromSeconds(7); Date = $null } }
    })) 7
    $response = Invoke-AzRestMethodWithRetry -Request @{ Method = 'GET'; Path = '/x' } -Description 'test'
    Assert-Equal 'rest retried status' $response.StatusCode 200
    Assert-Equal 'rest calls' $state.RestCalls 3
    Assert-Equal 'rest sleeps' $state.Sleeps 2
    $state.Attempts = 0
    Assert-ThrowError 'non-transient not retried' { Invoke-WithRetry -Description 't' -ScriptBlock { $state.Attempts++; throw 'Forbidden' } } 'Forbidden'
    Assert-Equal 'non-transient attempts' $state.Attempts 1
    $state.Attempts = 0
    $result = Invoke-WithRetry -Description 't' -ScriptBlock { $state.Attempts++; if ($state.Attempts -lt 3) { throw 'HTTP 429 TooManyRequests' }; 'ok' }
    Assert-Equal 'transient recovered' "$result/$($state.Attempts)" 'ok/3'
    $originalLimit = $resourceGraphLimit
    $script:resourceGraphLimit = 2
    try {
        $null = Get-FrontDoorResource -SubscriptionId @('a', 'b', 'c', 'd', 'e')
    }
    finally {
        $script:resourceGraphLimit = $originalLimit
    }
    Assert-Equal 'chunk sizes' ($state.Chunks -join ',') '2,2,1'
    #endregion

    if ($failures.Count -gt 0) {
        throw "Self-test failed ($($failures.Count)):`n" + ($failures -join "`n")
    }
    Write-Host 'Self-test passed.'
}

if ($SelfTest) {
    Invoke-SelfTest
    return
}

# A bad path, locked destination, or unavailable Excel dependency is handled before
# spending time retrieving Azure subscriptions and resource pages.
$resolvedOutputFile = Resolve-InventoryOutputPath -Path $OutputFile

# Use the caller's existing Az login and restrict discovery to its current tenant.
$context = Get-AzContext
if (-not $context -or -not $context.Account -or -not $context.Tenant) {
    throw 'No established Azure context was found. Run Connect-AzAccount first.'
}

$subscriptions = @(
    Get-AzSubscription -TenantId $context.Tenant.Id |
        Where-Object State -eq 'Enabled'
)
if ($SubscriptionName) {
    $subscriptions = @($subscriptions | Where-Object Name -eq $SubscriptionName)
    if ($subscriptions.Count -eq 0) {
        throw "No enabled subscription named '$SubscriptionName' was found in tenant '$($context.Tenant.Id)'."
    }
    if ($subscriptions.Count -gt 1) {
        $ids = $subscriptions.Id -join ', '
        throw "More than one enabled subscription is named '$SubscriptionName'. Matching IDs: $ids"
    }
}
if ($subscriptions.Count -eq 0) {
    throw "No enabled subscriptions were found in tenant '$($context.Tenant.Id)'."
}

$resources = @(Get-FrontDoorResource -SubscriptionId $subscriptions.Id)
$creationTimes = Get-FrontDoorCreationTime -Profiles $resources
$inventory = @(ConvertTo-FrontDoorInventory -Resources $resources -Subscriptions $subscriptions -CreationTimes $creationTimes)

$unresolvedTargets = @($inventory | Where-Object MigrationStatus -eq 'TargetNotFoundInScope').Count
$unresolvedSources = @($inventory | Where-Object MigrationStatus -eq 'SourceNotFoundInScope').Count
if ($unresolvedTargets -gt 0 -or $unresolvedSources -gt 0) {
    Write-Warning "Migration counterparts not found in queried scope: $unresolvedTargets targets, $unresolvedSources Classic sources. This does not prove deletion."
}

$undatedTargets = @($inventory | Where-Object {
    $_.Sku -ne 'Classic' -and $_.MigrationSourceResourceId -and $null -eq $_.MigrationDateUtc
}).Count
if ($undatedTargets -gt 0) {
    Write-Warning "$undatedTargets migration target(s) lack a usable date. Dates remain blank and these targets are excluded from monthly charts."
}

# Prepare totals before publishing; report and summary use the same graph snapshot.
$summary = @(Get-FrontDoorSummary -Inventory $inventory)
Export-FrontDoorReport -Inventory $inventory -Path $resolvedOutputFile -Columns $columns

# Console text goes to the host so it never mixes with objects returned by -PassThru.
$stopwatch.Stop()
Write-Host "Scope: tenant $($context.Tenant.Id); $($subscriptions.Count) enabled subscription(s) queried."
Write-Host "Profiles: $($inventory.Count). Endpoint counts include configured but unused/disabled endpoints."
$summary | Format-Table -AutoSize | Out-Host
Write-Host "Inventory exported to: $resolvedOutputFile"
Write-Host ('Elapsed: {0:N2} seconds.' -f $stopwatch.Elapsed.TotalSeconds)

if ($PassThru) {
    $inventory
}
