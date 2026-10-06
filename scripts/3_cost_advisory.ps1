[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('PRFR-_4Z', 'JGW1-KG0')]
    [string]$Advisory,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config.json'),
    [string]$DataDir = (Join-Path $PSScriptRoot '..\data')
)
$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'
$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$days = if ($cfg.costWindowDays) { [int]$cfg.costWindowDays } else { 30 }
$safe = $Advisory -replace '-', '_'
$rows = Get-Content (Join-Path $DataDir "rows_$safe.json") -Raw | ConvertFrom-Json
$profiles = Get-Content (Join-Path $PSScriptRoot 'advisories.json') -Raw | ConvertFrom-Json
$profile = $profiles.PSObject.Properties[$Advisory].Value
# Always query both categories so the filter 'values' serializes as a JSON array
# (PowerShell ConvertTo-Json unwraps a single-element array to a scalar, which the
# Cost Management API rejects with 'values is missing or empty'). The Python report
# filters to the categories each advisory needs.
$queryCategories = @('Virtual Machines', 'Storage')
$subIds = @((Get-Content (Join-Path $DataDir 'submap.json') -Raw | ConvertFrom-Json).PSObject.Properties.Name)
$to = (Get-Date).Date.AddDays(-1); $from = $to.AddDays(-1 * ($days - 1))
$fromStr = $from.ToString('yyyy-MM-ddT00:00:00+00:00'); $toStr = $to.ToString('yyyy-MM-ddT23:59:59+00:00')
if ($subIds.Count -eq 0) {
    '[]' | Set-Content (Join-Path $DataDir "cost_$safe.json") -Encoding utf8
    @{ from=$fromStr; to=$toStr } | ConvertTo-Json | Set-Content (Join-Path $DataDir "cost_window_$safe.json") -Encoding utf8
    return
}
function Get-MgmtHeaders {
    $t = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
    $tk = if ($t.Token -is [System.Security.SecureString]) { [System.Net.NetworkCredential]::new('', $t.Token).Password } else { $t.Token }
    return @{ Authorization = "Bearer $tk"; 'Content-Type' = 'application/json' }
}
$script:headers = Get-MgmtHeaders
$maxAttempts = if ($cfg.costMaxAttempts) { [int]$cfg.costMaxAttempts } else { 12 }
# Number of extra full sweeps over the still-failing subscriptions after the first pass,
# each preceded by a cooldown. On large tenants transient 429s are common; multi-pass
# recovery lets the per-subscription rate limits reset instead of giving up on pass 1.
$recoveryPasses = if ($cfg.costRecoveryPasses) { [int]$cfg.costRecoveryPasses } else { 3 }
$recoveryCooldown = if ($cfg.costRecoveryCooldownSeconds) { [int]$cfg.costRecoveryCooldownSeconds } else { 90 }
$grouping = @(
    @{ type='Dimension'; name='ResourceLocation' },
    @{ type='Dimension'; name='MeterCategory' },
    @{ type='Dimension'; name='MeterSubCategory' },
    @{ type='Dimension'; name='Meter' },
    @{ type='Dimension'; name='PricingModel' }
)
$body = @{ type='AmortizedCost'; timeframe='Custom'; timePeriod=@{from=$fromStr;to=$toStr}; dataset=@{granularity='None';aggregation=@{totalCost=@{name='Cost';function='Sum'}};grouping=$grouping;filter=@{dimensions=@{name='MeterCategory';operator='In';values=$queryCategories}}} } | ConvertTo-Json -Depth 10

# Query one subscription with in-request retry/backoff. Returns @{ok=$true;rows=...}
# or @{ok=$false;code=;msg=}. Refreshes the shared token on 401 or periodically.
function Invoke-SubCostQuery {
    param([string]$Sub)
    $uri = "https://management.azure.com/subscriptions/$Sub/providers/Microsoft.CostManagement/query?api-version=2023-11-01"
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $resp = Invoke-RestMethod -Method Post -Uri $uri -Headers $script:headers -Body $body
            $cols = $resp.properties.columns.name
            $out = New-Object System.Collections.Generic.List[object]
            foreach ($r in $resp.properties.rows) {
                $o = @{}; for ($c = 0; $c -lt $cols.Count; $c++) { $o[$cols[$c]] = $r[$c] }
                if ([string]$o.PricingModel -ne 'Reservation') { $o.SubscriptionId = $Sub; $out.Add([pscustomobject]$o) }
            }
            return @{ ok = $true; rows = $out }
        }
        catch {
            $code = $null
            try { $code = [int]$_.Exception.Response.StatusCode.value__ } catch {}
            if ($code -eq 401 -and $attempt -lt $maxAttempts) { $script:headers = Get-MgmtHeaders; continue }
            $retryable = ($code -eq 429 -or ($code -ge 500 -and $code -le 599) -or -not $code)
            if ($retryable -and $attempt -lt $maxAttempts) {
                # Exponential backoff capped at 120s, plus jitter to de-sync parallel throttles.
                $delay = [math]::Min(120, [math]::Pow(2, $attempt)) + (Get-Random -Minimum 0 -Maximum 5)
                try {
                    $retryAfter = $_.Exception.Response.Headers['Retry-After']
                    if ($retryAfter -match '^\d+$') { $delay = [math]::Min(180, [int]$retryAfter + 2) }
                } catch {}
                Start-Sleep -Seconds $delay
                continue
            }
            $message = $_.Exception.Message
            if ($_.ErrorDetails.Message) { $message = $_.ErrorDetails.Message }
            return @{ ok = $false; code = $code; msg = $message }
        }
    }
}

$results = New-Object System.Collections.Generic.List[object]
$errorMap = @{}            # sub -> @{code;msg} for subs still failing
$failed = New-Object System.Collections.Generic.List[string]
$i = 0
foreach ($sub in $subIds) {
    $i++
    if ($i % 150 -eq 0) { $script:headers = Get-MgmtHeaders }
    $res = Invoke-SubCostQuery -Sub $sub
    if ($res.ok) { foreach ($row in $res.rows) { $results.Add($row) } }
    else { $failed.Add($sub); $errorMap[$sub] = @{ code = $res.code; msg = $res.msg } }
    if ($i % 10 -eq 0) { "processed $i/$($subIds.Count) (failed so far: $($failed.Count))" }
}

# Multi-pass recovery: re-query only the still-failing subs after a cooldown so that
# per-subscription throttle windows can reset. Persistent failures (e.g. a sub whose
# own FinOps tooling saturates its Cost Management rate limit) are reported, not hidden.
for ($pass = 1; $pass -le $recoveryPasses -and $failed.Count -gt 0; $pass++) {
    "recovery pass $pass/$recoveryPasses for $($failed.Count) sub(s); cooldown ${recoveryCooldown}s"
    Start-Sleep -Seconds $recoveryCooldown
    $script:headers = Get-MgmtHeaders
    $stillFailed = New-Object System.Collections.Generic.List[string]
    foreach ($sub in $failed) {
        $res = Invoke-SubCostQuery -Sub $sub
        if ($res.ok) { foreach ($row in $res.rows) { $results.Add($row) }; $errorMap.Remove($sub) }
        else { $stillFailed.Add($sub); $errorMap[$sub] = @{ code = $res.code; msg = $res.msg } }
    }
    $failed = $stillFailed
    "  still failing after pass ${pass}: $($failed.Count)"
}

$errors = New-Object System.Collections.Generic.List[object]
foreach ($sub in $failed) { $errors.Add([pscustomobject]@{ sub = $sub; code = $errorMap[$sub].code; msg = $errorMap[$sub].msg }) }

$results | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $DataDir "cost_$safe.json") -Encoding utf8
@{advisory=$Advisory;from=$fromStr;to=$toStr} | ConvertTo-Json | Set-Content (Join-Path $DataDir "cost_window_$safe.json") -Encoding utf8
# Write as an argument (not pipeline) so an empty collection still serializes as "[]".
Set-Content (Join-Path $DataDir "cost_errors_$safe.json") -Value (ConvertTo-Json -InputObject @($errors) -Depth 4 -AsArray) -Encoding utf8
if ($Advisory -eq 'JGW1-KG0') {
    $storageRows = @($results | Where-Object { $_.MeterCategory -eq 'Storage' })
    $storageRows | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $DataDir "storage_cost_$safe.json") -Encoding utf8
    "Storage cost rows: $($storageRows.Count)"
}
"$Advisory cost rows: $($results.Count) | errors: $($errors.Count)"
