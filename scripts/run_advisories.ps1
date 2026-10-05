[CmdletBinding()]
param(
    [switch]$PRFR4Z,
    [switch]$JGW1KG0,
    [switch]$Both,
    [switch]$SkipInventory,
    [switch]$SkipCost
)
$ErrorActionPreference = 'Stop'
if ($Both -or (-not $PRFR4Z -and -not $JGW1KG0)) { $advisories=@('PRFR-_4Z','JGW1-KG0') }
else { $advisories=@(); if($PRFR4Z){$advisories+='PRFR-_4Z'};if($JGW1KG0){$advisories+='JGW1-KG0'} }
if (-not $SkipInventory) { & (Join-Path $PSScriptRoot '1_inventory.ps1'); if($LASTEXITCODE){throw 'Inventory failed.'} }
foreach($advisory in $advisories){
    & python (Join-Path $PSScriptRoot '2_build_advisory_rows.py') --advisory $advisory
    if($LASTEXITCODE){throw "Row build failed for $advisory."}
    if(-not $SkipCost){& (Join-Path $PSScriptRoot '3_cost_advisory.ps1') -Advisory $advisory;if($LASTEXITCODE){throw "Cost query failed for $advisory."}}
    & python (Join-Path $PSScriptRoot '4_build_advisory_report.py') --advisory $advisory
    if($LASTEXITCODE){throw "Report build failed for $advisory."}
}
