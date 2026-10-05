<#
.SYNOPSIS
    Step 1 — Collect a tenant-wide inventory of all virtual machines via Azure
    Resource Graph.

.DESCRIPTION
    Reads the target tenant ID from config.json (or -TenantId), enumerates every
    enabled subscription in that tenant, and runs a paged Resource Graph query
    for all VMs. Writes:
        data/vms.json      - raw VM records (size, power state, tags)
        data/submap.json   - subscription id -> name map

.NOTES
    Requires: Az.Accounts, Az.ResourceGraph
        Install-Module Az.Accounts, Az.ResourceGraph -Scope CurrentUser
    Sign in first:
        Connect-AzAccount -TenantId <your-tenant-id>
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config.json'),
    [string]$DataDir    = (Join-Path $PSScriptRoot '..\data')
)

$WarningPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

# --- Resolve configuration ---
if (-not $TenantId) {
    if (-not (Test-Path $ConfigPath)) {
        throw "No -TenantId supplied and config not found at $ConfigPath. Copy config.example.json to config.json and set your tenantId."
    }
    $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $TenantId = $cfg.tenantId
}
if (-not $TenantId -or $TenantId -eq '00000000-0000-0000-0000-000000000000') {
    throw "Set a real tenant ID in config.json (tenantId) or pass -TenantId."
}

New-Item -ItemType Directory -Force -Path $DataDir | Out-Null

# --- 1. Build subscription id -> name map for the target tenant ---
$subs = Get-AzSubscription | Where-Object { $_.State -eq 'Enabled' -and $_.TenantId -eq $TenantId }
$subIds = @($subs.Id)
if ($subIds.Count -eq 0) {
    throw "No enabled subscriptions found for tenant $TenantId. Run Connect-AzAccount -TenantId $TenantId first."
}
"Enabled subscriptions in tenant: $($subIds.Count)"

$subMap = @{}
foreach ($s in $subs) { $subMap[$s.Id] = $s.Name }
$subMap | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $DataDir 'submap.json') -Encoding utf8

# --- 2. Paged Resource Graph query for ALL VMs across those subscriptions ---
$query = @"
resources
| where type =~ 'microsoft.compute/virtualmachines'
| extend vmSize = tostring(properties.hardwareProfile.vmSize)
| extend powerState = tostring(properties.extended.instanceView.powerState.code)
| project subscriptionId, resourceGroup, name, vmSize, powerState, tags
"@

$all = New-Object System.Collections.Generic.List[object]
$skip = $null
$page = 0
do {
    if ($skip) {
        $res = Search-AzGraph -Query $query -Subscription $subIds -First 1000 -SkipToken $skip
    } else {
        $res = Search-AzGraph -Query $query -Subscription $subIds -First 1000
    }
    foreach ($r in $res) { $all.Add($r) }
    $skip = $res.SkipToken
    $page++
    "Page $page : fetched $($res.Count), total $($all.Count)"
} while ($skip)

"TOTAL VMs: $($all.Count)"

# --- 3. Export raw VM records to JSON (tags preserved as an object) ---
$export = foreach ($v in $all) {
    [pscustomobject]@{
        subscriptionId = $v.subscriptionId
        resourceGroup  = $v.resourceGroup
        name           = $v.name
        vmSize         = $v.vmSize
        powerState     = $v.powerState
        tags           = $v.tags
    }
}
$export | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $DataDir 'vms.json') -Encoding utf8
"Wrote $(Join-Path $DataDir 'vms.json')"
