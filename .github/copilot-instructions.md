# Copilot instructions — Azure VM Impact Report

Project-specific guidance for working in this repository. Read this before editing
any script.

## What this project is

A small, **read-only** toolkit that generates multi-sheet Excel workbooks showing
which VMs (and, for some advisories, storage) in an Azure tenant are affected by a
Microsoft health advisory — retirements, price increases, and regional price
changes — plus an estimated cost impact. It is run manually by a user from their
own machine and is published as a standalone, customer-agnostic GitHub repo.

It does **not** write to Azure. It only queries Azure Resource Graph and Cost
Management using the user's existing signed-in Azure context.

## Architecture: one advisory-driven pipeline

The single supported entry point is [`scripts/run_advisories.ps1`](../scripts/run_advisories.ps1).
It defaults to both notices and accepts `-PRFR4Z`, `-JGW1KG0`, or `-Both`, plus
`-SkipInventory` / `-SkipCost`. It runs inventory once, then builds
advisory-specific rows, cost data, and a workbook per advisory. Files are named
with the tracking ID (`rows_<safe>.json`, `cost_<safe>.json`, etc., where `safe`
replaces `-` with `_`) so the two notices never overwrite each other's data.

Each step reads/writes JSON in a local `data/` folder (git-ignored). Run order
matters; later steps depend on earlier outputs. Run **from the repo root**.

| Step | Script | Reads | Writes |
|------|--------|-------|--------|
| 1 | [scripts/1_inventory.ps1](../scripts/1_inventory.ps1) | `config.json` | `data/vms.json`, `data/submap.json` (VM location included) |
| 2 | [scripts/2_build_advisory_rows.py](../scripts/2_build_advisory_rows.py) | `scripts/advisories.json`, `data/vms.json`, `data/submap.json` | `data/rows_<safe>.json` |
| 3 | [scripts/3_cost_advisory.ps1](../scripts/3_cost_advisory.ps1) | `config.json`, `data/rows_<safe>.json`, `data/submap.json` | `data/cost_<safe>.json`, `data/cost_window_<safe>.json`, `data/cost_errors_<safe>.json`, (JGW) `data/storage_cost_<safe>.json` |
| 4 | [scripts/4_build_advisory_report.py](../scripts/4_build_advisory_report.py) | `config.json`, `scripts/advisories.json`, the `data/*_<safe>.json` outputs | `<Org>_VM_Impact_<safe>.xlsx` at repo root |

There is **no legacy pipeline**. Do not reintroduce non-advisory duplicate
scripts (e.g. `2_build_rows.py`, `3_cost.ps1`, `4_build_report.py`).

## The two layers of configuration — keep them separate

1. **`config.json`** (copied from `config.example.json`, git-ignored): the
   *per-tenant knobs* — `tenantId`, `organizationName` (branding + output
   filename), `costWindowDays`, dates, family-label ordering, throttling/recovery
   tuning, and optional `tagColumns`.
2. **`scripts/advisories.json`**: the *advisory ruleset* — one profile per
   tracking ID. `mode` is `prfr` (tenant-wide family-based impact) or `jgw`
   (region-based impact). Holds affected families/regions, excluded families,
   excluded storage terms, percentage, title, and effective date. Adding a new
   advisory should be a profile edit here, not a new script set.

## Key domain facts (don't re-derive these)

- **Retirements** (default Dv3/Dsv3/Ev3/Esv3) are modelled separately from price
  increases and **excluded** from the cost estimate — they're being removed, not
  re-priced.
- **Price waves (PRFR-_4Z)**: `v1` = versionless families (Bv1, D, Ds, F, Fs, G,
  Gs, Ls, NP, HC); `v2` = v2-series (Av2, Amv2, Dv2, Dsv2, Fsv2, Lsv2).
- **JGW1-KG0 is regional**: VM impact is limited to the regions and excluded VM
  families in `advisories.json`. Its Storage scope is estimated separately from
  `MeterCategory = Storage`, with the notice-listed excluded storage services
  removed. Keep VM and Storage totals separate; the JGW workbook has a dedicated
  `Storage Cost Impact` sheet.
- **Cost scope is always tenant-wide**. Cost collection enumerates every enabled
  subscription from `data/submap.json`, uses `AmortizedCost` over
  `costWindowDays`, and excludes `PricingModel = Reservation`. SavingsPlan usage
  stays included.
- **Billing meter naming**: many Dv2/Dsv2 sizes bill on a single *combined* meter
  (e.g. `D4 v2/DS4 v2`); Spot usage bills on standalone meters. `classify.py`'s
  `meter_group()` canonicalizes these into one row per family group so the VM
  Cost Impact sheet doesn't show confusing duplicate rows — without
  double-counting.

## VM size classification

The size parser lives **inline in `scripts/2_build_advisory_rows.py`**
(`SIZE_RE` + `family()` + `classify_prfr()`). `SIZE_RE` parses
`Standard_<prefix><number>[-<constrained>]<suffix>[_v<N>][_Promo]`; a trailing
`S` on the letter prefix (e.g. `DS`) means premium storage. `scripts/classify.py`
is a **separate** concern: it canonicalizes *billing meter* names for cost
grouping (`meter_group()`), imported by step 4. Validate any taxonomy change
against real sizes before trusting it.

## Cost throttling and recovery (step 3)

The Cost Management API returns HTTP 429 under load at large-tenant scale. Step 3
is built to survive this:

- `Invoke-SubCostQuery` retries per request with exponential backoff + jitter,
  honors `Retry-After`, and refreshes the token on 401. Capped by `costMaxAttempts`
  (config, default 12).
- After the main sweep, a **multi-pass recovery** loop re-queries only the
  still-failing subscriptions, `costRecoveryPasses` times (default 3) separated by
  `costRecoveryCooldownSeconds` (default 90).
- Subscriptions that still fail are written to `data/cost_errors_<safe>.json` —
  **never silently dropped** — and surface on the workbook's **Failed
  Subscriptions** tab plus a Summary warning banner. Some subs that run their own
  cost automation persistently self-throttle and clear only on a later re-run.

## Conventions and gotchas

- **Run from the repo root.** Scripts resolve paths relative to the script file,
  so `config.json`/`data/` resolve correctly only from the repo root.
- **PowerShell `ConvertTo-Json` array traps** (these bit us repeatedly): a
  single-element array serializes as a bare object, and piping an *empty*
  collection writes nothing (leaving a stale file). Always write arrays as
  `Set-Content $path -Value (ConvertTo-Json -InputObject @($items) -Depth N -AsArray)`.
- **Tags**: every VM tag is always captured in the **All Tags** column. `tagColumns`
  is optional (`[]` by default) and only promotes specific tags into their own
  filterable columns. Do NOT hardcode customer-specific tag names in the scripts.
- **No customer-specific values in scripts.** The org name on the title and
  filename comes from `config.organizationName`; advisory specifics come from
  `advisories.json`. Keep it generic and publishable.
- **PowerShell token handling**: `Get-AzAccessToken` can return a `SecureString`;
  the cost script converts it. Keep that guard if you touch it.
- **Output filename** derives from `organizationName`; override with the `OUTNAME`
  environment variable. If the target `.xlsx` is open in Excel, `wb.save()` throws
  `PermissionError` — ask the user to close it or use a different `OUTNAME`.
- **Never commit** `config.json`, the `data/` folder, or generated `.xlsx` files.
  `.gitignore` already excludes them; they can contain tenant/subscription IDs,
  resource names, and tag values.

## Validating changes

There's no automated test suite. To validate without live Azure calls, stage real
`data/vms.json` + `data/submap.json`, then run the Python steps:

```powershell
python scripts/2_build_advisory_rows.py --advisory PRFR-_4Z
python scripts/2_build_advisory_rows.py --advisory JGW1-KG0
python scripts/4_build_advisory_report.py --advisory PRFR-_4Z
```

Compile-check Python (`python -m py_compile scripts\4_build_advisory_report.py`)
and parse-check PowerShell before a live run. A live run requires the user's
existing `Connect-AzAccount` session. Clean up staged `data/`, `config.json`, and
generated `.xlsx` afterward so the repo stays publish-ready.

## Scope discipline

Keep this tool small and read-only. Do not add write operations against Azure,
new dependencies beyond `openpyxl` + the Az modules, or customer-specific logic in
the scripts. Customer specifics belong in `config.json`; advisory specifics belong
in `advisories.json`.
