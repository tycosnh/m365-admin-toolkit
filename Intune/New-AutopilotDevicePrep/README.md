# New-AutopilotDevicePrep

Builds a complete Windows Autopilot device preparation (Autopilot v2) baseline in one run.

Setting up Autopilot v2 by hand takes a dozen portal clicks, and one missed step leaves the policy showing "0 groups assigned." The most common miss: the device group must be owned by the Intune Provisioning Client service principal. This script gets every step right, every time.

## What it creates

1. **Device group.** An assigned security group owned by the Intune Provisioning Client. The script creates the service principal if the tenant lacks it.
2. **User group.** Dynamic by default. It includes only enabled members with an active Intune license, so unlicensed users never hit a failed enrollment.
3. **RMM install script.** An Intune platform script that installs the Datto RMM agent during setup. It runs as SYSTEM and 64-bit, and it waits for the installer to finish.
4. **Device preparation policy.** Wired to the script, with the device group bound for enrollment-time grouping and the user group assigned.

## Built to be safe

- **Idempotent.** It finds existing groups, scripts, and policies by name and skips them.
- **Checks consent first.** If any Graph scope wasn't granted, it stops before creating anything.
- **Confirms the tenant.** It shows the tenant ID and asks before making changes.
- **Retries** assignments that fail while Graph catches up on new objects.
- **One module.** It uses only `Microsoft.Graph.Authentication`, with raw Graph calls, so version mismatches between Graph sub-modules can't break it.

## Usage

```powershell
.\New-AutopilotDevicePrep.ps1 -ClientPrefix CONTOSO `
    -DattoSiteGuid 00000000-0000-0000-0000-000000000000 `
    -DattoPlatformHost concord.rmm.datto.com
```

Sign in as a Global Administrator of the target tenant when prompted.

| Parameter | Purpose |
| --- | --- |
| `-ClientPrefix` | Short name used for every object, e.g. `CONTOSO - Autopilot Devices`. |
| `-DattoSiteGuid` | Site GUID from the Datto RMM agent download URL. |
| `-DattoPlatformHost` | Your Datto platform host, e.g. `concord.rmm.datto.com`. |
| `-UserGroupType` | `Dynamic` (default) or `Assigned`. |
| `-UserAccountType` | `StandardUser` (default) or `Administrator`. |
| `-TimeoutMinutes` | Setup timeout, 15–720. Default 60. |
| `-AllowUserToSkip` | Let users skip setup if it fails. |
| `-HideDiagnosticsLink` | Hide the diagnostics link on failure. |

## Requirements

- PowerShell 7 or Windows PowerShell 5.1
- `Microsoft.Graph.Authentication`
- Delegated scopes: `Group.ReadWrite.All`, `DeviceManagementConfiguration.ReadWrite.All`, `DeviceManagementScripts.ReadWrite.All`, `Directory.Read.All`, `Application.ReadWrite.All`

## After it runs

- Check **Devices > Enrollment > Device platform restrictions**. Personally owned Windows MDM must be allowed, or you need corporate identifiers.
- Open the policy and confirm **1 group assigned** under Device group.
