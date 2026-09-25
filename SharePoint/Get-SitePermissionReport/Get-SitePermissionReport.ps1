<#
.SYNOPSIS
    Produces a folder-level permission report for a SharePoint Online site.

.DESCRIPTION
    Walks the site, every non-hidden document library, and every folder within
    those libraries, recording which objects break permission inheritance and who
    has access to them.

    SharePoint groups, Entra security groups and M365 groups are expanded down to
    named users so the report answers "who can actually open this folder" rather
    than just naming a group. Hidden SharingLinks.* system groups are decoded into
    readable sharing-link types (Anyone / Organization / Specific people), since
    those are a frequent cause of folders having unexpected access.

    Outputs a CSV (full detail) and an HTML summary (unique-permission items only).

.PARAMETER SiteUrl
    Full URL of the site, e.g. https://contoso.sharepoint.com/sites/PeopleandCulture

.PARAMETER ClientId
    Entra app registration (client) ID used for the PnP connection. PnP.PowerShell 2.x
    requires this even for -Interactive. The app needs SharePoint delegated permissions;
    for group expansion it also needs Microsoft Graph Group.Read.All + User.Read.All.

.PARAMETER OutputFolder
    Where the CSV and HTML land. Created if missing.

.PARAMETER IncludeFiles
    Also report individual files that have unique permissions. Slower on large libraries.

.PARAMETER SkipGroupExpansion
    Report principal names only, without resolving group membership.

.PARAMETER IncludeLimitedAccess
    Include role assignments that grant only "Limited Access". These are plumbing
    entries SharePoint creates automatically and are excluded by default as noise.

.EXAMPLE
    .\Get-SitePermissionReport.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/PeopleandCulture" -ClientId "00000000-0000-0000-0000-000000000000"

.NOTES
    Requires PnP.PowerShell and Site Collection Administrator rights on the target site.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    # Optional. Omit to use PnP's built-in app registration; supply your own if the
    # tenant has not consented to it, or if you need Graph scopes for group expansion.
    [string]$ClientId,

    [string]$OutputFolder = (Join-Path (Get-Location).Path "Reports"),

    [switch]$IncludeFiles,

    [switch]$SkipGroupExpansion,

    [switch]$IncludeLimitedAccess
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell -ErrorAction Stop

# Caches so a group shared across 50 folders is only resolved once
$Script:GroupCache = @{}
$Script:GraphAvailable = $true
$Script:Results = [System.Collections.Generic.List[object]]::new()

#region Helpers

function Write-Step {
    param([string]$Message, [string]$Colour = "Cyan")
    Write-Host $Message -ForegroundColor $Colour
}

# Folder names and display names can contain & < >, which would break the HTML table.
# Hand-rolled rather than using System.Web so this works on both PS 5.1 and 7.
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

<#
    Decodes a SharingLinks.* group title into something a customer can read.
    Format is: SharingLinks.<docGuid>.<LinkKind>.<linkGuid>
#>
function Get-SharingLinkType {
    param([string]$Title)

    if ($Title -notlike "SharingLinks.*") { return $null }

    $kind = ($Title -split '\.')[2]
    switch ($kind) {
        "Flexible"     { "Sharing link - Specific people" }
        "Organization" { "Sharing link - Anyone in the organisation" }
        "AnonymousEdit"{ "Sharing link - ANYONE with the link (edit)" }
        "AnonymousView"{ "Sharing link - ANYONE with the link (view)" }
        "Anonymous"    { "Sharing link - ANYONE with the link" }
        default        { "Sharing link - $kind" }
    }
}

<#
    Pulls the Entra object GUID out of a SharePoint claims login name.
    c:0t.c|tenant|<guid>                                 -> security group
    c:0o.c|federateddirectoryclaimprovider|<guid>        -> M365 group members
    c:0o.c|federateddirectoryclaimprovider|<guid>_o      -> M365 group owners
#>
function Get-EntraObjectId {
    param([string]$LoginName)

    if ($LoginName -match '\|([0-9a-fA-F-]{36})(_o)?$') { return $Matches[1] }
    return $null
}

function Expand-EntraGroup {
    param([string]$ObjectId, [string]$GroupTitle)

    if (-not $Script:GraphAvailable) {
        return @([PSCustomObject]@{ Name = "$GroupTitle (not expanded - no Graph access)"; Email = "" })
    }

    try {
        $members = Get-PnPAzureADGroupMember -Identity $ObjectId -ErrorAction Stop
        if (-not $members) {
            return @([PSCustomObject]@{ Name = "$GroupTitle (empty group)"; Email = "" })
        }
        return $members | ForEach-Object {
            [PSCustomObject]@{
                Name  = if ($_.DisplayName) { $_.DisplayName } else { $_.UserPrincipalName }
                Email = $_.UserPrincipalName
            }
        }
    }
    catch {
        # First failure almost always means the app registration lacks Graph consent.
        # Warn once, then stop retrying for every subsequent group.
        if ($Script:GraphAvailable) {
            Write-Warning "Could not read Entra group membership ($GroupTitle): $($_.Exception.Message)"
            Write-Warning "Group membership will not be expanded. Grant the app Group.Read.All to enable this."
            $Script:GraphAvailable = $false
        }
        return @([PSCustomObject]@{ Name = "$GroupTitle (not expanded - no Graph access)"; Email = "" })
    }
}

function Expand-SharePointGroup {
    param([string]$GroupTitle)

    try {
        $members = Get-PnPGroupMember -Identity $GroupTitle -ErrorAction Stop
        if (-not $members) {
            return @([PSCustomObject]@{ Name = "$GroupTitle (empty group)"; Email = "" })
        }

        $expanded = [System.Collections.Generic.List[object]]::new()
        foreach ($m in $members) {
            # SharePoint groups can themselves contain Entra groups - recurse one level
            $nestedId = Get-EntraObjectId -LoginName $m.LoginName
            if ($nestedId -and $m.PrincipalType -ne "User") {
                Expand-EntraGroup -ObjectId $nestedId -GroupTitle $m.Title | ForEach-Object { $expanded.Add($_) }
            }
            else {
                $expanded.Add([PSCustomObject]@{ Name = $m.Title; Email = $m.Email })
            }
        }
        return $expanded
    }
    catch {
        Write-Warning "Could not expand SharePoint group '$GroupTitle': $($_.Exception.Message)"
        return @([PSCustomObject]@{ Name = "$GroupTitle (could not expand)"; Email = "" })
    }
}

<#
    Turns a role assignment Member into one or more concrete people.
    Returns objects with Name / Email so the caller can emit a row per person.
#>
function Resolve-Principal {
    param($Member)

    $login = $Member.LoginName
    $title = $Member.Title

    if ($Script:GroupCache.ContainsKey($login)) { return $Script:GroupCache[$login] }

    $resolved = switch -Regex ($login) {
        # Individual user
        '^i:0#\.f\|membership\|' {
            @([PSCustomObject]@{ Name = $title; Email = ($login -split '\|')[-1] })
            break
        }
        # Everyone except external users
        '^c:0-\.f\|rolemanager\|spo-grid-all-users' {
            @([PSCustomObject]@{ Name = "** Everyone except external users **"; Email = "" })
            break
        }
        # Everyone, including external
        '^c:0\(\.s\|true' {
            @([PSCustomObject]@{ Name = "** Everyone (includes external users) **"; Email = "" })
            break
        }
        default {
            if ($SkipGroupExpansion) {
                @([PSCustomObject]@{ Name = $title; Email = "" })
            }
            elseif ($Member.PrincipalType -eq "SharePointGroup") {
                Expand-SharePointGroup -GroupTitle $title
            }
            else {
                $objectId = Get-EntraObjectId -LoginName $login
                if ($objectId) {
                    Expand-EntraGroup -ObjectId $objectId -GroupTitle $title
                }
                else {
                    @([PSCustomObject]@{ Name = $title; Email = "" })
                }
            }
        }
    }

    $Script:GroupCache[$login] = $resolved
    return $resolved
}

<#
    Reads the role assignments off any securable object (web, list, list item)
    and appends a row to $Script:Results for every person who ends up with access.
#>
function Add-PermissionRows {
    param(
        $Securable,
        [string]$Scope,
        [string]$ItemName,
        [string]$ItemPath,
        [bool]$HasUnique
    )

    if (-not $HasUnique) {
        $Script:Results.Add([PSCustomObject]@{
            Scope            = $Scope
            ItemName         = $ItemName
            ItemPath         = $ItemPath
            Inheritance      = "Inherited"
            PrincipalName    = "(inherits from parent)"
            PrincipalType    = ""
            GrantedVia       = ""
            MemberName       = ""
            MemberEmail      = ""
            PermissionLevels = ""
            SharingLinkType  = ""
        })
        return
    }

    $roleAssignments = Get-PnPProperty -ClientObject $Securable -Property RoleAssignments

    foreach ($ra in $roleAssignments) {
        Get-PnPProperty -ClientObject $ra -Property Member, RoleDefinitionBindings | Out-Null

        $levels = @($ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name)

        # "Limited Access" is scaffolding SharePoint adds so a user can traverse to
        # something deeper. It grants nothing on its own, so drop it by default.
        if (-not $IncludeLimitedAccess) {
            $meaningful = $levels | Where-Object { $_ -ne "Limited Access" }
            if (-not $meaningful) { continue }
            $levels = $meaningful
        }

        $member      = $ra.Member
        $linkType    = Get-SharingLinkType -Title $member.Title
        $displayName = if ($linkType) { $linkType } else { $member.Title }

        foreach ($person in (Resolve-Principal -Member $member)) {
            $Script:Results.Add([PSCustomObject]@{
                Scope            = $Scope
                ItemName         = $ItemName
                ItemPath         = $ItemPath
                Inheritance      = "UNIQUE"
                PrincipalName    = $displayName
                PrincipalType    = $member.PrincipalType
                GrantedVia       = if ($person.Name -eq $member.Title) { "Direct" } else { $displayName }
                MemberName       = $person.Name
                MemberEmail      = $person.Email
                PermissionLevels = ($levels -join ", ")
                SharingLinkType  = $linkType
            })
        }
    }
}

<#
    Returns every folder in a library. Uses a single recursive CAML query, which is
    fast, and falls back to walking the folder tree if the library trips the 5,000
    item list view threshold.
#>
function Get-AllFolders {
    param($List)

    $camlQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq></Where>
  </Query>
  <ViewFields>
    <FieldRef Name='ID' /><FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' />
  </ViewFields>
  <RowLimit Paged='TRUE'>2000</RowLimit>
</View>
"@

    try {
        return @(Get-PnPListItem -List $List -Query $camlQuery -ErrorAction Stop)
    }
    catch {
        Write-Warning "Recursive query failed on '$($List.Title)' ($($_.Exception.Message)). Falling back to tree walk."
        return @(Get-PnPListItem -List $List -PageSize 500 -ErrorAction Stop |
                    Where-Object { $_.FileSystemObjectType -eq "Folder" })
    }
}

function Get-UniqueFilesInList {
    param($List)

    $camlQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq></Where>
  </Query>
  <ViewFields>
    <FieldRef Name='ID' /><FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' />
  </ViewFields>
  <RowLimit Paged='TRUE'>2000</RowLimit>
</View>
"@

    try   { return @(Get-PnPListItem -List $List -Query $camlQuery -ErrorAction Stop) }
    catch { Write-Warning "Could not enumerate files in '$($List.Title)': $($_.Exception.Message)"; return @() }
}

#endregion Helpers

#region Connect

Write-Step "Connecting to $SiteUrl ..."
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -ErrorAction Stop
}
catch {
    Write-Error "Connection failed: $($_.Exception.Message)"
    exit 1
}

$web = Get-PnPWeb -Includes HasUniqueRoleAssignments, Title, ServerRelativeUrl
Write-Step "Connected to: $($web.Title)" "Green"

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$safeName  = ($web.Title -replace '[^\w\-]', '')
$csvPath   = Join-Path $OutputFolder "$($safeName)_Permissions_$timestamp.csv"
$htmlPath  = Join-Path $OutputFolder "$($safeName)_Permissions_$timestamp.html"

#endregion Connect

#region Site level

Write-Step "Reading site-level permissions ..."
Add-PermissionRows -Securable $web -Scope "Site" -ItemName $web.Title `
                   -ItemPath $web.ServerRelativeUrl -HasUnique $true

#endregion Site level

#region Libraries and folders

# BaseTemplate 101 = document library. Skip hidden and catalog lists.
$libraries = Get-PnPList -Includes RootFolder, HasUniqueRoleAssignments |
    Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden -and $_.Title -ne "Form Templates" }

Write-Step "Found $($libraries.Count) document librarie(s)." "Green"

$libIndex = 0
foreach ($lib in $libraries) {
    $libIndex++
    Write-Progress -Activity "Scanning libraries" -Status $lib.Title `
                   -PercentComplete (($libIndex / $libraries.Count) * 100)
    Write-Step "`n[$libIndex/$($libraries.Count)] Library: $($lib.Title)"

    Add-PermissionRows -Securable $lib -Scope "Library" -ItemName $lib.Title `
                       -ItemPath $lib.RootFolder.ServerRelativeUrl `
                       -HasUnique $lib.HasUniqueRoleAssignments

    $folders = Get-AllFolders -List $lib | Where-Object { $_["FileLeafRef"] -ne "Forms" }
    Write-Host "    $($folders.Count) folder(s)" -ForegroundColor DarkGray

    $folderIndex = 0
    foreach ($folder in $folders) {
        $folderIndex++
        $path = $folder["FileRef"]
        $name = $folder["FileLeafRef"]

        Write-Progress -Activity "  Folders in $($lib.Title)" -Status $name `
                       -PercentComplete (($folderIndex / [Math]::Max($folders.Count, 1)) * 100) -Id 1

        try {
            $hasUnique = Get-PnPProperty -ClientObject $folder -Property HasUniqueRoleAssignments
            if ($hasUnique) {
                Write-Host "    UNIQUE  $path" -ForegroundColor Yellow
            }
            Add-PermissionRows -Securable $folder -Scope "Folder" -ItemName $name `
                               -ItemPath $path -HasUnique $hasUnique
        }
        catch {
            Write-Warning "Failed on folder '$path': $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity "  Folders in $($lib.Title)" -Completed -Id 1

    if ($IncludeFiles) {
        $files = Get-UniqueFilesInList -List $lib
        Write-Host "    checking $($files.Count) file(s) for unique permissions" -ForegroundColor DarkGray

        foreach ($file in $files) {
            try {
                $hasUnique = Get-PnPProperty -ClientObject $file -Property HasUniqueRoleAssignments
                # Only report files that break inheritance - inherited files are just noise
                if ($hasUnique) {
                    Write-Host "    UNIQUE  $($file["FileRef"])" -ForegroundColor Yellow
                    Add-PermissionRows -Securable $file -Scope "File" -ItemName $file["FileLeafRef"] `
                                       -ItemPath $file["FileRef"] -HasUnique $true
                }
            }
            catch {
                Write-Warning "Failed on file '$($file["FileRef"])': $($_.Exception.Message)"
            }
        }
    }
}
Write-Progress -Activity "Scanning libraries" -Completed

#endregion Libraries and folders

#region Export

Write-Step "`nWriting report ..."
$Script:Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$unique = @($Script:Results | Where-Object { $_.Inheritance -eq "UNIQUE" -and $_.Scope -ne "Site" })
$uniqueLocations = @($unique | Select-Object -ExpandProperty ItemPath -Unique)
$externalLinks   = @($unique | Where-Object { $_.SharingLinkType -like "*ANYONE*" })

$rowsHtml = ($unique | Sort-Object ItemPath, PrincipalName, MemberName | ForEach-Object {
    $rowClass = if ($_.SharingLinkType -like "*ANYONE*") { ' class="alert"' } else { "" }
    "<tr$rowClass><td>$($_.Scope)</td><td>$(ConvertTo-HtmlSafe $_.ItemPath)</td>" +
    "<td>$(ConvertTo-HtmlSafe $_.PrincipalName)</td>" +
    "<td>$(ConvertTo-HtmlSafe $_.MemberName)</td>" +
    "<td>$(ConvertTo-HtmlSafe $_.MemberEmail)</td>" +
    "<td>$(ConvertTo-HtmlSafe $_.PermissionLevels)</td></tr>"
}) -join "`n"

$html = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Permission Report - $($web.Title)</title>
<style>
 body { font-family: Segoe UI, Arial, sans-serif; margin: 30px; color: #222; }
 h1 { font-size: 22px; margin-bottom: 4px; }
 .meta { color: #666; font-size: 13px; margin-bottom: 24px; }
 .summary { background: #f4f6f8; border-left: 4px solid #0078d4; padding: 12px 16px; margin-bottom: 24px; }
 .summary strong { display: inline-block; min-width: 260px; }
 table { border-collapse: collapse; width: 100%; font-size: 13px; }
 th { background: #0078d4; color: #fff; text-align: left; padding: 8px; position: sticky; top: 0; }
 td { border-bottom: 1px solid #e1e1e1; padding: 6px 8px; vertical-align: top; }
 tr:nth-child(even) td { background: #fafafa; }
 tr.alert td { background: #fff4f4; font-weight: 600; }
 .footer { margin-top: 24px; color: #888; font-size: 12px; }
</style></head><body>
<h1>Folder Permission Report</h1>
<div class="meta">$($web.Title) &mdash; $SiteUrl<br/>Generated $(Get-Date -Format "dd MMM yyyy HH:mm")</div>

<div class="summary">
  <strong>Libraries scanned:</strong> $($libraries.Count)<br/>
  <strong>Locations with unique permissions:</strong> $($uniqueLocations.Count)<br/>
  <strong>Total access entries listed below:</strong> $($unique.Count)<br/>
  <strong>"Anyone with the link" grants:</strong> $($externalLinks.Count)
</div>

<p>Only items that <em>break inheritance</em> are shown. Everything not listed inherits its
permissions from the site. Rows highlighted in red are anonymous sharing links, which are
accessible without signing in.</p>

<table>
<tr><th>Scope</th><th>Path</th><th>Granted To</th><th>Resolved User</th><th>Email</th><th>Permission</th></tr>
$rowsHtml
</table>

<div class="footer">Full detail, including inherited items, is in the accompanying CSV.</div>
</body></html>
"@

$html | Out-File -FilePath $htmlPath -Encoding UTF8

Write-Host "`n===== Summary =====" -ForegroundColor Green
Write-Host "Libraries scanned                : $($libraries.Count)"
Write-Host "Locations with unique permissions: $($uniqueLocations.Count)"
Write-Host "Total rows in CSV                : $($Script:Results.Count)"
Write-Host "'Anyone with the link' grants    : $($externalLinks.Count)" -ForegroundColor $(if ($externalLinks.Count) { "Red" } else { "Green" })
Write-Host "`nCSV : $csvPath" -ForegroundColor Yellow
Write-Host "HTML: $htmlPath" -ForegroundColor Yellow

Disconnect-PnPOnline

#endregion Export
