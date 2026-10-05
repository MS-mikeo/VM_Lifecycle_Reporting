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
$subIds = @($rows.price_subs)
$to = (Get-Date).Date.AddDays(-1); $from = $to.AddDays(-1 * ($days - 1))
$fromStr = $from.ToString('yyyy-MM-ddT00:00:00+00:00'); $toStr = $to.ToString('yyyy-MM-ddT23:59:59+00:00')
if ($subIds.Count -eq 0) {
    '[]' | Set-Content (Join-Path $DataDir "cost_$safe.json") -Encoding utf8
    @{ from=$fromStr; to=$toStr } | ConvertTo-Json | Set-Content (Join-Path $DataDir "cost_window_$safe.json") -Encoding utf8
    return
}
$tok = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
$token = if ($tok.Token -is [System.Security.SecureString]) { [System.Net.NetworkCredential]::new('', $tok.Token).Password } else { $tok.Token }
$headers = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
$grouping = @(
    @{ type='Dimension'; name='ResourceLocation' },
    @{ type='Dimension'; name='MeterCategory' },
    @{ type='Dimension'; name='MeterSubCategory' },
    @{ type='Dimension'; name='Meter' }
)
$body = @{ type='AmortizedCost'; timeframe='Custom'; timePeriod=@{from=$fromStr;to=$toStr}; dataset=@{granularity='None';aggregation=@{totalCost=@{name='Cost';function='Sum'}};grouping=$grouping;filter=@{dimensions=@{name='MeterCategory';operator='In';values=@('Virtual Machines')}}} } | ConvertTo-Json -Depth 8
$results = New-Object System.Collections.Generic.List[object]; $errors = New-Object System.Collections.Generic.List[object]; $i=0
foreach ($sub in $subIds) {
    $i++; $uri="https://management.azure.com/subscriptions/$sub/providers/Microsoft.CostManagement/query?api-version=2023-11-01"; $attempt=0
    while ($true) { $attempt++
        try {
            $resp=Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body; $cols=$resp.properties.columns.name
            foreach ($r in $resp.properties.rows) { $o=@{}; for($c=0;$c -lt $cols.Count;$c++){$o[$cols[$c]]=$r[$c]};$o.SubscriptionId=$sub;$results.Add([pscustomobject]$o) }; break
        } catch { $code=$_.Exception.Response.StatusCode.value__; if($code -eq 429 -and $attempt -lt 6){Start-Sleep -Seconds (5*$attempt);continue};$errors.Add([pscustomobject]@{sub=$sub;code=$code;msg=$_.Exception.Message});break }
    }
    if($i % 10 -eq 0){"processed $i/$($subIds.Count)"}
}
$results | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $DataDir "cost_$safe.json") -Encoding utf8
@{advisory=$Advisory;from=$fromStr;to=$toStr} | ConvertTo-Json | Set-Content (Join-Path $DataDir "cost_window_$safe.json") -Encoding utf8
if($errors.Count){$errors|ConvertTo-Json -Depth 4|Set-Content (Join-Path $DataDir "cost_errors_$safe.json") -Encoding utf8}
"$Advisory cost rows: $($results.Count) | errors: $($errors.Count)"
