<#
.SYNOPSIS
    Removes leftover OneDrive and File Explorer entries from an old tenant after a
    tenant-to-tenant migration.

.DESCRIPTION
    After a tenant-to-tenant move, users often keep a dead "OneDrive - OldCompany" entry
    in the File Explorer navigation pane, or OneDrive keeps trying to sync the old account.
    This script cleans that up for the current user:

      1. Removes OneDrive account keys whose display name matches the old tenant.
      2. Removes matching tenant subkeys under the surviving accounts.
      3. Removes Explorer navigation pane CLSIDs and their Desktop NameSpace entries.
      4. Restarts Explorer so the pane refreshes.

    Supports -WhatIf, so you can see exactly what would be removed first.

.PARAMETER OldTenantName
    Part of the old tenant's display name as it appears in Explorer, e.g. "Fabrikam".
    Matched with wildcards on both sides.

.PARAMETER NoExplorerRestart
    Skip restarting Explorer at the end.

.EXAMPLE
    .\Repair-OneDriveTenantMigration.ps1 -OldTenantName Fabrikam -WhatIf

.EXAMPLE
    .\Repair-OneDriveTenantMigration.ps1 -OldTenantName Fabrikam

.NOTES
    Runs in the user context (HKCU). Deploy it as an Intune or RMM script that runs
    as the signed-in user, not as SYSTEM.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$OldTenantName,

    [switch]$NoExplorerRestart
)

$pattern = "*$OldTenantName*"
$removed = [System.Collections.Generic.List[string]]::new()

# 1 and 2. OneDrive account config
Get-ChildItem 'HKCU:\Software\Microsoft\OneDrive\Accounts' -ErrorAction SilentlyContinue | ForEach-Object {
    $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue

    if ($props.DisplayName -like $pattern) {
        if ($PSCmdlet.ShouldProcess($_.Name, 'Remove OneDrive account key')) {
            Remove-Item $_.PSPath -Recurse -Force
            $removed.Add("Account: $($_.PSChildName)")
        }
        return
    }

    $tenantsPath = Join-Path $_.PSPath 'Tenants'
    if (Test-Path $tenantsPath) {
        Get-ChildItem $tenantsPath | Where-Object { $_.PSChildName -like $pattern } | ForEach-Object {
            if ($PSCmdlet.ShouldProcess($_.Name, 'Remove OneDrive tenant key')) {
                Remove-Item $_.PSPath -Recurse -Force
                $removed.Add("Tenant key: $($_.PSChildName)")
            }
        }
    }
}

# 3. Explorer navigation pane entries
Get-ChildItem 'HKCU:\Software\Classes\CLSID' -ErrorAction SilentlyContinue | ForEach-Object {
    $name = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).'(default)'
    if ($name -like $pattern) {
        $guid = $_.PSChildName
        if ($PSCmdlet.ShouldProcess("$name ($guid)", 'Remove Explorer navigation pane entry')) {
            Remove-Item $_.PSPath -Recurse -Force
            Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$guid" `
                -Force -ErrorAction SilentlyContinue
            $removed.Add("Nav pane: $name $guid")
        }
    }
}

if ($removed.Count -eq 0) {
    Write-Output "Nothing matching '$OldTenantName' found. No changes made."
    return
}

$removed | ForEach-Object { Write-Output "Removed $_" }

# 4. Refresh Explorer
if (-not $NoExplorerRestart -and $PSCmdlet.ShouldProcess('explorer.exe', 'Restart')) {
    Stop-Process -Name explorer -Force
}
