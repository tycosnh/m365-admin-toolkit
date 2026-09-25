# Repair-OneDriveTenantMigration

Cleans up the leftovers a tenant-to-tenant migration leaves on each PC.

After a migration, users often see a dead "OneDrive - OldCompany" entry in File Explorer, or OneDrive keeps trying to sign in to the old account. This script removes those traces for the signed-in user.

## What it removes

1. OneDrive account keys tied to the old tenant.
2. Old tenant subkeys under the new account.
3. The old tenant's File Explorer navigation pane entry.

Then it restarts Explorer so the change shows right away.

## Usage

Preview first:

```powershell
.\Repair-OneDriveTenantMigration.ps1 -OldTenantName Fabrikam -WhatIf
```

Then run it:

```powershell
.\Repair-OneDriveTenantMigration.ps1 -OldTenantName Fabrikam
```

Add `-NoExplorerRestart` to skip the Explorer restart.

## Deploying at scale

The script works on the current user's registry. Deploy it through Intune or your RMM as a script that runs **as the signed-in user**, not as SYSTEM.
