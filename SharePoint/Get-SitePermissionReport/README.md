# Get-SitePermissionReport

Answers the question every client eventually asks: **"Who can actually open this folder?"**

SharePoint's own permission pages show group names, not people, and they hide sharing links inside system groups. This script walks a whole site and lists every person who has access to every folder that breaks inheritance.

## What it does

- Scans the site, every visible document library, and every folder. Add `-IncludeFiles` to catch individually shared files too.
- **Expands groups to people.** SharePoint groups, Entra security groups, and Microsoft 365 groups all resolve to named users, including groups nested inside SharePoint groups.
- **Decodes sharing links.** Hidden `SharingLinks.*` groups become readable types: Anyone, Organization, or Specific people.
- **Flags risk.** "Anyone with the link" grants get highlighted in red in the HTML report.
- **Cuts noise.** It drops "Limited Access" entries by default, since they grant nothing on their own.

## Output

- **CSV:** full detail, one row per person per location.
- **HTML:** a client-ready summary of only the locations with unique permissions.

## Usage

```powershell
.\Get-SitePermissionReport.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
    -ClientId "00000000-0000-0000-0000-000000000000"
```

| Parameter | Purpose |
| --- | --- |
| `-SiteUrl` | The site to scan. |
| `-ClientId` | Your PnP app registration in the tenant. |
| `-OutputFolder` | Where reports go. Default: `.\Reports`. |
| `-IncludeFiles` | Also report files with unique permissions. Slower. |
| `-SkipGroupExpansion` | List group names only. |
| `-IncludeLimitedAccess` | Keep Limited Access entries. |

## Requirements

- `PnP.PowerShell`
- Site Collection Administrator on the target site
- An Entra app registration with SharePoint delegated permissions. Add Graph `Group.Read.All` and `User.Read.All` for group expansion. Without them, the script warns once and lists group names instead.
