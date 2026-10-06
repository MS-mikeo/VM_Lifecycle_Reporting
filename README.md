# Azure VM Retirement & Price-Change Impact Report

Generate multi-sheet Excel workbooks showing which virtual machines (and, where
relevant, storage) in an Azure tenant are affected by a Microsoft health
advisory — covering **retirements** (end-of-life SKUs), **price increases**, and
**regional price changes** — along with an estimated cost impact based on the
last N days of amortized spend.

The toolkit runs entirely from your own machine using the Azure PowerShell
modules and Python. It is **read-only**: it queries Azure Resource Graph and Cost
Management using your signed-in Azure context and never writes to your tenant.

## Advisories covered

Advisory definitions live in [`scripts/advisories.json`](scripts/advisories.json).
Two Microsoft notices ship out of the box:

| Tracking ID | Scope |
|-------------|-------|
| **PRFR-_4Z** | VM **retirements** (Dv3/Dsv3/Ev3/Esv3) plus **price increases** on v1 and v2 VM families, tenant-wide. |
| **JGW1-KG0** | **Regional** price increase on VM families *and* storage meters in a defined set of regions. |

Each advisory produces its own workbook so the two notices never overwrite each
other's data.

## What it produces

One `.xlsx` workbook per advisory. Sheets vary slightly by advisory:

| Sheet | Contents |
|-------|----------|
| **Summary** | Impact counts per family/region plus the estimated cost impact (trailing window + annualized). |
| **Retirement Inventory** *(PRFR-_4Z only)* | Every VM in a retiring family, with subscription, resource group, name, size, power state, and all tags. |
| **Price Increase Inventory** | Every price-impacted VM, with the change wave or region plus the same detail columns. |
| **VM Cost Impact** | Amortized VM cost grouped by meter family/region, with the `+X%` increase and projected cost. |
| **Storage Cost Impact** *(JGW1-KG0 only)* | Amortized storage cost by region and meter, with notice-excluded services removed. |
| **Failed Subscriptions** | Any subscription whose cost query could not be completed, with the HTTP code and a plain-language reason. Shows "None" when every subscription returned data. |

## Prerequisites

- **PowerShell 7+** with the Azure modules:
  ```powershell
  Install-Module Az.Accounts, Az.ResourceGraph -Scope CurrentUser
  ```
- **Python 3.9+** with openpyxl:
  ```powershell
  pip install -r requirements.txt
  ```
- **Azure permissions**: *Reader* on the subscriptions you want to scan, and
  *Cost Management Reader* (or equivalent) for the cost step.

## Setup

1. Copy the example config and fill in your tenant:
   ```powershell
   Copy-Item config.example.json config.json
   ```
   Edit `config.json` and set at least `tenantId` and `organizationName`.
   `config.json` is git-ignored so your tenant ID never gets committed.

2. Sign in to Azure for that tenant:
   ```powershell
   Connect-AzAccount -TenantId <your-tenant-id>
   ```

## Running the reports

The single supported entry point is the advisory runner. With no switch, it runs
both notices and creates two separately labeled workbooks:

```powershell
pwsh ./scripts/run_advisories.ps1
```

Run one notice only when needed:

```powershell
pwsh ./scripts/run_advisories.ps1 -PRFR4Z
pwsh ./scripts/run_advisories.ps1 -JGW1KG0
```

Use `-Both` for an explicit both-advisories run. Use `-SkipInventory` or
`-SkipCost` to reuse data already present in `data/` (for example, to rebuild the
workbook after a formatting change without re-querying Azure).

Under the hood the runner performs these steps per advisory:

1. `1_inventory.ps1` — tenant-wide VM inventory via Azure Resource Graph (runs once).
2. `2_build_advisory_rows.py` — classify VMs into the advisory's impact rows.
3. `3_cost_advisory.ps1` — query amortized VM/storage cost across the tenant.
4. `4_build_advisory_report.py` — assemble the Excel workbook.

### Output

Workbooks are written to the repository root. Filenames derive from
`organizationName` in your config, so with `"organizationName": "Contoso"` you get:

- `Contoso_VM_Impact_PRFR__4Z.xlsx`
- `Contoso_VM_Impact_JGW1_KG0.xlsx`

Set the `OUTNAME` environment variable before the report step to override the
output path.

## Cost scope, throttling, and recovery

- **Cost scope is always tenant-wide.** The cost step enumerates every enabled
  subscription, queries `AmortizedCost` over the trailing `costWindowDays`, and
  excludes `PricingModel = Reservation`. SavingsPlan usage remains included.
- The Cost Management API returns **HTTP 429** under load, especially on large
  tenants (hundreds of subscriptions). The cost step handles this automatically:
  - **Per-request retry** with exponential backoff and jitter, honoring the
    `Retry-After` header, and refreshing the access token on 401.
  - **Multi-pass recovery**: after the main sweep, still-failing subscriptions
    are re-queried in additional passes separated by a cooldown.
- Any subscription that still fails after all recovery passes is listed on the
  **Failed Subscriptions** tab (never silently dropped) and triggers a warning
  banner on the Summary sheet. Some subscriptions that run their own cost
  automation can saturate their per-subscription rate limit and may need a
  later re-run (`-SkipInventory`) to clear.

Tune the behavior for larger tenants via `config.json` (all optional):

| Key | Default | Purpose |
|-----|---------|---------|
| `costMaxAttempts` | `12` | Max retry attempts per subscription within a single query. |
| `costRecoveryPasses` | `3` | Number of additional recovery passes over still-failing subscriptions. |
| `costRecoveryCooldownSeconds` | `90` | Cooldown between recovery passes. |

## Configuration reference (`config.json`)

| Key | Purpose |
|-----|---------|
| `tenantId` | The Azure AD tenant to scan. **Required.** |
| `organizationName` | Branding shown on the Summary sheet and used for the output filename. |
| `reportTitle`, `advisoryReference` | Free-text labels shown on the Summary sheet. |
| `costWindowDays` | Trailing window (in days) for the amortized cost query. |
| `priceIncreasePercent` | Default price increase to model where an advisory does not override it. |
| `retirementEndOfLife`, `priceIncreaseEffective` | Dates shown on the Summary sheet. |
| `retirementFamilies`, `priceFamiliesV1`, `priceFamiliesV2` | Family labels used for Summary breakdown ordering. |
| `costMaxAttempts`, `costRecoveryPasses`, `costRecoveryCooldownSeconds` | Throttling/recovery tuning (see above). |
| `tagColumns` | **Optional.** Extra inventory columns pulled from specific VM tags. Empty by default. |

### Pulling specific tags into their own columns (optional)

By default the inventory sheets include an **All Tags** column that lists every
tag on each VM, so no tag data is lost. To break particular tags out into their
own dedicated, filterable columns, add entries to `tagColumns`:

```json
"tagColumns": [
  { "header": "CostCenter", "keys": ["CostCenter", "cost_center"] },
  { "header": "Owner",      "keys": ["Owner", "owner"] }
]
```

Each entry lists candidate tag keys (case-insensitive); the first non-empty match
wins. Leave `tagColumns` as `[]` to rely solely on the **All Tags** column.

## Adapting to a different advisory

Add a new profile to [`scripts/advisories.json`](scripts/advisories.json) — its
`mode` (`prfr` for tenant-wide family-based impact, `jgw` for region-based
impact), affected families/regions, excluded families, excluded storage terms,
title, and effective date. The VM size parser and impact classification live in
[`scripts/2_build_advisory_rows.py`](scripts/2_build_advisory_rows.py); billing
meter canonicalization (so the cost sheet shows one row per family group) lives
in [`scripts/classify.py`](scripts/classify.py). Prefer adding a profile over
creating a second script set.

## Security notes

- `config.json` and the entire `data/` folder (raw inventory and cost data) are
  git-ignored, as are generated `.xlsx` files. Only the sanitized scripts are
  intended to be committed.
- The scripts are **read-only** against Azure. They use your signed-in Azure
  context and never store credentials.
- Review the generated `.xlsx` before sharing — it contains subscription IDs,
  resource names, and tag values from your environment.

## License

MIT — see [LICENSE](LICENSE).
