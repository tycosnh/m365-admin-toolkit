# M365 Admin Toolkit

PowerShell tools I use to run Microsoft 365 for small businesses. Each one came out of a real ticket and runs in production tenants.

I'm a systems engineer at a managed service provider, holding MD-102 (Endpoint Administrator) and MS-102 (Microsoft 365 Administrator Expert).

| Tool | What it does |
| --- | --- |
| [New-AutopilotDevicePrep](Intune/New-AutopilotDevicePrep) | Builds a full Windows Autopilot device preparation (v2) baseline through Microsoft Graph in one run. |
| [Get-SitePermissionReport](SharePoint/Get-SitePermissionReport) | Shows who can actually open every SharePoint folder, with groups expanded and sharing links decoded. CSV and HTML output. |
| [Get-WeeklyEmailActivity](Exchange/Get-WeeklyEmailActivity) | Rebuilds a user's sent and received email counts week by week from message trace. |
| [Repair-OneDriveTenantMigration](OneDrive/Repair-OneDriveTenantMigration) | Cleans up dead OneDrive accounts and Explorer entries left over from a tenant-to-tenant migration. |

## How I build these

- **Safe to rerun.** Scripts check for what already exists before creating anything.
- **Fail early.** Missing permissions or wrong module versions stop the run before it changes anything.
- **Plain output.** Reports are made for the client to read, not just the admin.

## Need help with your tenant?

I take freelance Microsoft 365, Intune, and Entra ID work. Reach me through [Upwork](https://www.upwork.com/) or open an issue here.

## License

MIT. Use these freely. Test in a non-production tenant first.
