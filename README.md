# AFDInventory

A PowerShell script that inventories **Azure Front Door Classic, Standard, and Premium** profiles across a tenant and tracks how far the Classic → Standard/Premium migration has progressed.

It queries Azure Resource Graph once for the whole tenant, links migrated Classic profiles to their Standard/Premium targets, and writes a CSV or an Excel workbook with a migration dashboard.

## Features

- Inventory of Classic, Standard, and Premium profiles with endpoint counts, subscription, and resource group.
- Migration linking in both directions: each migrated Classic profile points to its target, and each target records its Classic (or CDN) source.
- Migration dates, with the basis of each date recorded, and a **Migration Statistics** worksheet (XLSX only) with native Excel charts: monthly and cumulative migrations, current Classic progress, and coverage metrics.
- Fail-fast behavior: malformed data stops the run instead of producing guessed values.
- Transient throttling (HTTP 429/5xx) is retried with backoff, honouring `Retry-After`.
- Safe publishing: the report is staged next to the destination and only replaces the previous file after a successful export. Locked or unwritable destinations are detected before any Azure call.
- Built-in offline self-test.

## Requirements

- PowerShell 7.0 or later
- [`Az.Accounts`](https://www.powershellgallery.com/packages/Az.Accounts) and [`Az.ResourceGraph`](https://www.powershellgallery.com/packages/Az.ResourceGraph)
- [`ImportExcel`](https://www.powershellgallery.com/packages/ImportExcel) (optional, for `.xlsx`). If it is missing, the script writes a `.csv` instead and warns.
- An existing Azure sign-in with **Reader** access to the subscriptions to inventory

```powershell
Install-Module Az.Accounts, Az.ResourceGraph, ImportExcel -Scope CurrentUser
Connect-AzAccount
```

## Usage

```powershell
# Excel workbook with the migration dashboard, all enabled subscriptions in the current tenant
.\Get-AzureFrontDoorInventory.ps1 -OutputFile .\FrontDoorInventory.xlsx

# CSV, limited to one subscription (exact name)
.\Get-AzureFrontDoorInventory.ps1 -OutputFile .\FrontDoorInventory.csv -SubscriptionName 'My Production Subscription'

# Also return the rows to the pipeline
$rows = .\Get-AzureFrontDoorInventory.ps1 -OutputFile .\FrontDoorInventory.xlsx -PassThru
```

| Parameter | Description |
|---|---|
| `-OutputFile` | Destination `.csv` or `.xlsx` file. |
| `-SubscriptionName` | Optional exact subscription name. Default: all enabled subscriptions in the current tenant. |
| `-PassThru` | Also return the inventory rows. Console messages go to the host, so they never mix with the returned objects. |
| `-SelfTest` | Run the offline checks instead of querying Azure. |

Scope is the **current tenant's** enabled subscriptions, not every tenant associated with the account. Switch tenants with `Connect-AzAccount -Tenant <id>` first.

## Output

Both formats use the same columns:

`Sku`, `ProfileName`, `SubscriptionName`, `SubscriptionId`, `ResourceGroupName`, `EndpointCount`, `IsMigrated`, `MigrationStatus`, `MigrationState`, `MigrationTargetProfileName`, `MigrationTargetResourceId`, `ClassicSourceProfileName`, `ClassicSourceResourceId`, `ResourceId`, `MigrationSourceType`, `MigrationSourceProfileName`, `MigrationSourceResourceId`, `MigrationDateUtc`, `MigrationDateBasis`, `AzureResourceState`

Key semantics:

- **`EndpointCount`** counts configured endpoints, including disabled or unused ones. It is not a traffic measure. Classic counts frontend endpoints (including custom hostnames); Standard/Premium counts `afdEndpoints`, not custom-domain resources.
- **`MigrationState`** is the business-facing status: `ToMigrate` or `Migrated` (Classic), `MigrationTarget` (migrated Standard/Premium), `AlreadyStandardPremium` (other Standard/Premium).
- **`AzureResourceState`** keeps Azure's original lifecycle value. It is not a health indicator.
- **`MigrationStatus`** distinguishes resolved links from counterparts not found in the queried scope. *Not found* does not prove deletion: scope, permissions, or Resource Graph indexing delays can also explain it.
- **`MigrationDateUtc`** is the target's ARM `createdTime` for Front Door Classic migrations (migration preparation, not necessarily the final commit) and `MigratedOn` for CDN profile migrations. `MigrationDateBasis` records which one. Missing dates stay blank. CSV dates are ISO 8601 UTC.
- The **Migration Statistics** sheet counts each visible migration target once. Classic progress counts only Classic profiles that still exist, so deleted, already-migrated sources are excluded and the figures are not an all-time total. The current month is partial.

## Self-test

```powershell
.\Get-AzureFrontDoorInventory.ps1 -SelfTest
```

Runs offline, with no Azure login or output file, and throws if any check fails. It covers ID classification, date parsing, migration linking, statistics, retry behavior, and subscription chunking.

## Notes

- Resource Graph does not expose creation times, so the script reads them from the ARM resource list, once per subscription that contains a migration target. A name or resource group filter on that call was measured to be slower and is intentionally not used.
- Subscriptions are queried in chunks of 1,000 (the Resource Graph limit) and result pages are fetched sequentially.
- Generated reports are git-ignored (`output/` and `FrontDoorInventory*.xlsx|csv`). Inventory files contain subscription and resource names, so review them before sharing.
