<#
.SYNOPSIS
    Step 3 — Query amortized VM cost for the price-impacted subscriptions.

.DESCRIPTION
    Reads the price-impacted subscription list from data/rows.json and queries
    the Azure Cost Management API (AmortizedCost, MeterCategory = "Virtual
    Machines") over a trailing window. Throttled (HTTP 429) subscriptions are
    retried automatically with exponential backoff. Writes:
        data/cost.json          - cost rows (one per MeterCategory/SubCategory/Meter)
        data/cost_window.json    - the from/to window actually used
        data/cost_errors.json    - any subscriptions that still failed

.NOTES
    Requires: Az.Accounts
    Sign in first:  Connect-AzAccount -TenantId <your-tenant-id>
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config.json'),
    [string]$DataDir    = (Join-Path $PSScriptRoot '..\data')
)

$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$windowDays = if ($cfg.costWindowDays) { [int]$cfg.costWindowDays } else { 30 }

$rows = Get-Content (Join-Path $DataDir 'rows.json') -Raw | ConvertFrom-Json
$subIds = @($rows.price_subs)
"Querying amortized cost for $($subIds.Count) subscriptions over $windowDays days"

# Trailing window, ending yesterday
$to   = (Get-Date).Date.AddDays(-1)
$from = $to.AddDays(-1 * ($windowDays - 1))
$fromStr = $from.ToString('yyyy-MM-ddT00:00:00+00:00')
$toStr   = $to.ToString('yyyy-MM-ddT23:59:59+00:00')
"Window: $fromStr  ->  $toStr"

# --- Acquire a management-plane token (handle SecureString on newer Az) ---
$tokObj = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
if ($tokObj.Token -is [System.Security.SecureString]) {
    $token = [System.Net.NetworkCredential]::new('', $tokObj.Token).Password
} else {
    $token = $tokObj.Token
}
$headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }

$body = @{
    type       = 'AmortizedCost'
    timeframe  = 'Custom'
    timePeriod = @{ from = $fromStr; to = $toStr }
    dataset    = @{
        granularity = 'None'
        aggregation = @{ totalCost = @{ name = 'Cost'; function = 'Sum' } }
        grouping    = @(
            @{ type = 'Dimension'; name = 'MeterCategory' },
            @{ type = 'Dimension'; name = 'MeterSubCategory' },
            @{ type = 'Dimension'; name = 'Meter' }
        )
        filter = @{ dimensions = @{ name = 'MeterCategory'; operator = 'In'; values = @('Virtual Machines') } }
    }
} | ConvertTo-Json -Depth 8

function Invoke-CostQuery {
    param([string]$Sub, [int]$MaxAttempts = 6)
    $uri = "https://management.azure.com/subscriptions/$Sub/providers/Microsoft.CostManagement/query?api-version=2023-11-01"
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $resp = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -Body $script:body
            $cols = $resp.properties.columns.name
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($r in $resp.properties.rows) {
                $obj = @{}
                for ($c = 0; $c -lt $cols.Count; $c++) { $obj[$cols[$c]] = $r[$c] }
                $obj['SubscriptionId'] = $Sub
                $out.Add([pscustomobject]$obj)
            }
            return @{ ok = $true; rows = $out }
        }
        catch {
            $code = $_.Exception.Response.StatusCode.value__
            if ($code -eq 429 -and $attempt -lt $MaxAttempts) {
                Start-Sleep -Seconds (5 * $attempt)
                continue
            }
            return @{ ok = $false; code = $code; msg = $_.Exception.Message }
        }
    }
}

$results = New-Object System.Collections.Generic.List[object]
$errors  = New-Object System.Collections.Generic.List[object]
$i = 0
foreach ($sub in $subIds) {
    $i++
    $r = Invoke-CostQuery -Sub $sub
    if ($r.ok) {
        foreach ($row in $r.rows) { $results.Add($row) }
    } else {
        $errors.Add([pscustomobject]@{ sub = $sub; code = $r.code; msg = $r.msg })
    }
    if ($i % 10 -eq 0) { "  processed $i/$($subIds.Count)" }
}

"Total cost rows: $($results.Count) | errors: $($errors.Count)"
$results | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $DataDir 'cost.json') -Encoding utf8
@{ from = $fromStr; to = $toStr } | ConvertTo-Json | Set-Content (Join-Path $DataDir 'cost_window.json') -Encoding utf8
if ($errors.Count) {
    $errors | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $DataDir 'cost_errors.json') -Encoding utf8
    "WARNING: $($errors.Count) subscriptions failed; see data/cost_errors.json"
}
"Wrote $(Join-Path $DataDir 'cost.json')"
