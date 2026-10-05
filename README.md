# Azure VM Retirement & Price-Change Impact Report

Generate a multi-sheet Excel report showing which virtual machines in an Azure
tenant are affected by a Microsoft health advisory — both **retirements**
(end-of-life SKUs) and **price increases** — along with an estimated cost impact
based on the last N days of amortized spend.

The toolkit runs entirely from your own machine using the Azure PowerShell
modules and Python. It reads only from Azure (Resource Graph + Cost Management);
it never writes to your tenant.

## What it produces

A single `.xlsx` workbook with four sheets:

| Sheet | Contents |
|-------|----------|
| **Summary** | VM counts per family, and the estimated cost impact of the price increase (trailing window + annualized). |
| **Retirement Inventory** | Every VM in a retiring family, with subscription, resource group, name, size, power state, and all tags. |
| **Price Increase Inventory** | Every price-impacted VM, with the change wave (v1/v2) plus the same detail columns. |
| **Cost Impact** | Amortized cost grouped by billing meter family, with the `+X%` increase and projected cost. |

## Prerequisites

- **PowerShell 7+** with the Azure modules:
  ```powershell
  Install-Module Az.Accounts, Az.ResourceGraph -Scope CurrentUser
  ```
- **Python 3.9+** with openpyxl:
  ```bash
  pip install -r requirements.txt
  ```
- **Azure permissions**: Reader on the subscriptions you want to scan, and Cost
  Management read access (e.g. *Cost Management Reader*) for the cost step.

## Setup

1. Copy the example config and fill in your tenant ID:
   ```bash
   cp config.example.json config.json
   ```
   Edit `config.json` and set at least `tenantId`. `config.json` is
   git-ignored so your tenant ID never gets committed.

2. Sign in to Azure for that tenant:
   ```powershell
   Connect-AzAccount -TenantId <your-tenant-id>
   ```

## Running the report

Run the four steps in order **from the repository root**. Each step writes its
output into a local `data/` folder (also git-ignored).

```powershell
# 1. Tenant-wide VM inventory (Azure Resource Graph)
pwsh ./scripts/1_inventory.ps1

# 2. Classify VMs into retirement / price-impacted rows
python ./scripts/2_build_rows.py

# 3. Query amortized cost for the price-impacted subscriptions
pwsh ./scripts/3_cost.ps1

# 4. Build the Excel workbook
python ./scripts/4_build_report.py
```

The workbook is written to the repository root; its filename is derived from
`organizationName` and `reportTitle` in your config. To override the output
path, set the `OUTNAME` environment variable before step 4.

> **Tip:** If you only need the inventory (no cost figures), you can skip step 3.
> Step 4 will still run and simply leave the cost numbers at zero.

## Configuration reference (`config.json`)

| Key | Purpose |
|-----|---------|
| `tenantId` | The Azure AD tenant to scan. **Required.** |
| `organizationName`, `reportTitle` | Branding shown on the Summary sheet and used for the output filename. |
| `advisoryReference` | Free-text reference to the Microsoft advisory, shown on the Summary sheet. |
| `priceIncreasePercent` | The price increase to model (e.g. `25`). |
| `costWindowDays` | Trailing window (in days) for the amortized cost query. |
| `retirementEndOfLife`, `priceIncreaseEffective` | Dates shown on the Summary sheet. |
| `retirementFamilies`, `priceFamiliesV1`, `priceFamiliesV2` | Family labels used for the Summary breakdown ordering. |
| `tagColumns` | **Optional.** Extra inventory columns pulled from specific VM tags. Empty by default. |

### Pulling specific tags into their own columns (optional)

By default the inventory sheets include an **All Tags** column that lists every
tag on each VM, so no tag data is lost. If you want particular tags broken out
into their own dedicated, filterable columns, add entries to `tagColumns`:

```json
"tagColumns": [
  { "header": "CostCenter", "keys": ["CostCenter", "cost_center"] },
  { "header": "Owner",      "keys": ["Owner", "owner"] }
]
```

Each entry lists candidate tag keys (case-insensitive); the first non-empty
match wins. Leave `tagColumns` as `[]` to rely solely on the **All Tags** column.

## Adapting to a different advisory

The SKU taxonomy — which size families count as retirements vs. price wave 1 vs.
wave 2 — lives in [`scripts/classify.py`](scripts/classify.py). It is the single
place to edit when a new advisory covers different VM families. The dates,
percentage, and tenant are all config-driven and need no code changes.

## Security notes

- `config.json` and the entire `data/` folder (raw inventory and cost data) are
  git-ignored. Only the sanitized scripts are intended to be committed.
- The scripts are **read-only** against Azure. They use your signed-in Azure
  context and never store credentials.
- Review the generated `.xlsx` before sharing — it contains subscription IDs,
  resource names, and tag values from your environment.

## License

MIT — see [LICENSE](LICENSE).
