# Copilot instructions — Azure VM Impact Report

Project-specific guidance for working in this repository. Read this before editing
any script.

## What this project is

A small, **read-only** toolkit that generates a multi-sheet Excel report showing
which VMs in an Azure tenant are affected by a Microsoft health advisory
(retirements + price increases) and the estimated cost impact. It is designed to
be run manually by a user from their own machine, then published as a standalone
GitHub repo for other customers to reuse.

It does **not** write to Azure. It only queries Azure Resource Graph and Cost
Management using the user's existing signed-in Azure context.

## Architecture: advisory runner plus legacy pipeline

The supported publishing workflow is `scripts/run_advisories.ps1`. It defaults
to both notices and accepts `-PRFR4Z`, `-JGW1KG0`, or `-Both`. It runs the
inventory once, then creates advisory-specific rows, cost data, and workbooks.
The advisory-specific files are named with the tracking ID so the two notices
cannot overwrite each other's local data.

The original four-step pipeline remains available for backward compatibility:

Each step reads/writes JSON in a local `data/` folder (git-ignored). Run order
matters; later steps depend on earlier outputs. All commands run **from the repo
root**.

| Step | Script | Reads | Writes |
|------|--------|-------|--------|
| 1 | [scripts/1_inventory.ps1](scripts/1_inventory.ps1) | `config.json` | `data/vms.json`, `data/submap.json` (including VM location) |
| 2 | [scripts/2_build_rows.py](scripts/2_build_rows.py) | `config.json`, `data/vms.json`, `data/submap.json` | `data/rows.json` |
| 3 | [scripts/3_cost.ps1](scripts/3_cost.ps1) | `config.json`, `data/rows.json` | `data/cost.json`, `data/cost_window.json`, `data/cost_errors.json` |
| 4 | [scripts/4_build_report.py](scripts/4_build_report.py) | `config.json`, `data/rows.json`, `data/cost.json`, `data/cost_window.json`, `data/vms.json`, `data/submap.json` | `<Org>_<Title>.xlsx` at repo root |

[scripts/classify.py](scripts/classify.py) is a **shared module** imported by both
step 2 and step 4. It is the single source of truth for the SKU taxonomy.

## The two layers of configuration — keep them separate

1. **`config.json`** (copied from `config.example.json`, git-ignored): the
   *easy knobs* — `tenantId`, percentage, window days, dates, display/branding,
   family-label ordering, and optional `tagColumns`. Changing the advisory's
   dates/percentage/tenant should **never** require a code edit.
2. **`scripts/classify.py`**: the *SKU ruleset* — which VM size families count as
   `retire` / `price_v1` / `price_v2`. This IS the encoding of a specific
   advisory. Edit here only when adapting to an advisory that covers different VM
   families.

Supported advisory rules for the publishing runner live in
[`scripts/advisories.json`](../scripts/advisories.json). Prefer adding a
profile there over creating a second customer-specific script set.

## Key domain facts (don't re-derive these)

- **Retirements** are modelled separately from **price increases**. Retiring
  families (default Dv3/Dsv3/Ev3/Esv3) are **excluded** from the cost estimate —
  they're being removed, not re-priced.
- **Price waves**: `v1` = versionless families (Bv1, D, Ds, F, Fs, G, Gs, Ls, NP,
  HC); `v2` = v2-series (Av2, Amv2, Dv2, Dsv2, Fsv2, Lsv2).
- **Cost = AmortizedCost**, `MeterCategory = "Virtual Machines"`, trailing
  `costWindowDays`. The `+X%` is applied on top. Amortized cost includes
  RI-covered usage, so the increase figure is an **upper bound** for RI VMs
  (existing Reserved Instances are NOT affected by the price change).
- **Azure billing meter naming**: many Dv2 and Dsv2 sizes bill on a single
  *combined* meter (e.g. `D4 v2/DS4 v2`), while Spot usage bills on standalone
  meters (e.g. `DS3 v2 Spot`). `classify.meter_group()` canonicalizes these into
  one row per family group via the `_CANON` map (e.g. both collapse to
  `Dv2/Dsv2`) so the Cost Impact sheet doesn't show confusing duplicate rows.
  All variants are still counted once in the wave total — there is no
  double-counting.

- **JGW1-KG0 is regional**: its VM impact is limited to the seven regions and
  excluded VM families listed in `scripts/advisories.json`. Its Storage scope
  is not included in this repository.

## VM size classification (how `classify.py` works)

`SIZE_RE` parses `Standard_<prefix><number>[-<constrained>]<suffix>[_v<N>][_Promo]`.
A trailing `S` on the letter prefix (e.g. `DS`) means premium storage → base `D` +
`has_s=True`. `classify(size)` returns `(family_label, category)` where category is
`retire` / `price_v1` / `price_v2`, or `(None, None)` when not covered. Validate
any taxonomy change against real sizes before trusting it.

## Conventions and gotchas

- **Run Python steps from the repo root** (`python scripts/2_build_rows.py`). The
  scripts resolve paths relative to the script file, so cwd must be repo root for
  `config.json`/`data/` to resolve.
- **Tags**: every VM tag is always captured in the **All Tags** column, so no tag
  data is lost. `tagColumns` is optional and defaults to `[]`; it only promotes
  specific tags into their own filterable columns. Do NOT hardcode customer-specific
  tag names (e.g. AppID/Vendor) in the scripts — that's what `tagColumns` is for.
- **PowerShell token handling**: `Get-AzAccessToken` returns a `SecureString` on
  newer Az. Step 3 already converts it via `[System.Net.NetworkCredential]`. Keep
  that guard if you touch the cost script.
- **Cost query throttling**: the Cost Management API returns HTTP 429 under load.
  Step 3 retries with exponential backoff inside `Invoke-CostQuery`. Persistently
  failing subs are written to `data/cost_errors.json`, not silently dropped.
- **Output filename** is derived from `organizationName` + `reportTitle`. Override
  with the `OUTNAME` environment variable before running step 4. If the target
  `.xlsx` is open in Excel, `wb.save()` throws `PermissionError` — the file is
  locked; ask the user to close it (or use a different `OUTNAME`).
- **Never commit** `config.json`, the `data/` folder, or generated `.xlsx` files.
  `.gitignore` already excludes them. These can contain tenant IDs, subscription
  IDs, resource names, and tag values.

## Validating changes

There's no automated test suite. To validate a change, stage real inventory JSON
into `data/` (renamed to `vms.json`, `submap.json`, `cost.json`,
`cost_window.json`), copy `config.example.json` to `config.json`, then run steps 2
and 4 and confirm the printed counts and cost totals are unchanged. Clean up the
staged `data/`, `config.json`, and generated `.xlsx` afterward so the repo stays
publish-ready.

For advisory changes, first run the classifiers without Azure calls by using
existing `data/vms.json` and `data/submap.json`:

```powershell
python scripts/2_build_advisory_rows.py --advisory PRFR-_4Z
python scripts/2_build_advisory_rows.py --advisory JGW1-KG0
```

Then compile-check Python and parse-check PowerShell before a live run. A live
run requires the user's existing `Connect-AzAccount` session.

## Scope discipline

Keep this tool small and read-only. Do not add write operations against Azure,
new external dependencies beyond `openpyxl` + the Az modules, or
customer-specific logic in the scripts. Customer specifics belong in `config.json`.
