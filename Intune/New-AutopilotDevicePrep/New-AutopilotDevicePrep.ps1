<#
.SYNOPSIS
    Provisions a baseline Windows Autopilot device preparation (Autopilot v2) configuration
    in a target tenant: device group, user group, RMM platform script, and the device
    preparation policy with both assignments wired up.

.DESCRIPTION
    Creates the following, idempotently:

      1. Assigned (static) device security group, owned by the Intune Provisioning Client
         service principal (AppId f1346770-5b25-470b-88bd-d5744ab7952c). Required: the
         Autopilot service writes device objects into this group at enrollment time.
      2. User security group (dynamic all-users by default) used as the policy assignment target.
      3. Intune platform script containing the RMM installer, configured for system context
         and 64-bit, assigned to the device group.
      4. Device preparation policy referencing the script, with the device group bound via
         assignJustInTimeConfiguration and the user group bound via assign.

    Uses Invoke-MgGraphRequest exclusively, so only Microsoft.Graph.Authentication is required.
    No dependency on the Beta cmdlet modules.

.PARAMETER TenantId
    Optional. Omit it and the tenant is resolved from the account you sign in as. Only supply
    it if you have cached tokens for several tenants and want to force the sign-in prompt to
    a specific one.

.PARAMETER ClientPrefix
    Short client identifier used to name the groups, script, and policy.
    Example: "CONTOSO" produces "CONTOSO - Autopilot Devices", etc.

.PARAMETER DattoSiteGuid
    The per-site GUID from the Datto RMM agent download URL. The installer script is generated
    from this rather than being read off disk.

.PARAMETER DattoPlatformHost
    Datto RMM platform hostname for your region, e.g. concord.rmm.datto.com,
    merlot.rmm.datto.com, or pinotage.rmm.datto.com. Copy it from the agent download URL.

.PARAMETER UserGroupType
    Dynamic (users holding an Intune service plan) or Assigned (empty group you populate yourself).

.EXAMPLE
    .\New-AutopilotDevicePrep.ps1 -ClientPrefix CONTOSO `
        -DattoSiteGuid 00000000-0000-0000-0000-000000000000 `
        -DattoPlatformHost concord.rmm.datto.com

.NOTES
    Requires: Microsoft.Graph.Authentication
    Delegated scopes: Group.ReadWrite.All, DeviceManagementConfiguration.ReadWrite.All,
                      Directory.Read.All, Application.ReadWrite.All (only if the Intune
                      Provisioning Client service principal is absent and must be created)
#>

[CmdletBinding()]
param(
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientPrefix,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
    [string]$DattoSiteGuid,

    [Parameter(Mandatory)]
    [string]$DattoPlatformHost,

    [ValidateSet('Dynamic', 'Assigned')]
    [string]$UserGroupType = 'Dynamic',

    [ValidateSet('StandardUser', 'Administrator')]
    [string]$UserAccountType = 'StandardUser',

    [ValidateRange(15, 720)]
    [int]$TimeoutMinutes = 60,

    [string]$ErrorMessageText = "Setup could not complete. Contact your IT support team for assistance.",

    [switch]$AllowUserToSkip,

    [switch]$HideDiagnosticsLink
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Fixed AppId of the service principal that must own the device group.
# May surface as "Intune Provisioning Client" or "Intune Autopilot ConfidentialClient".
$ProvisioningClientAppId = 'f1346770-5b25-470b-88bd-d5744ab7952c'
$DevicePrepTemplateId    = '80d33118-b7b4-40d8-b15f-81be745e053f_1'

# Directory objects go through v1.0. Only the Intune device preparation endpoints
# require beta, and mixing versions in reference URLs causes resolution failures.
$GraphV1                 = 'https://graph.microsoft.com/v1.0'
$GraphBeta               = 'https://graph.microsoft.com/beta'

#region Helpers ---------------------------------------------------------------

function Write-Step {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Cyan
}

function Write-Skip {
    param([string]$Message)
    Write-Host "[=] $Message" -ForegroundColor DarkGray
}

function New-MailNickname {
    param([string]$DisplayName)
    $clean = ($DisplayName -replace '[^a-zA-Z0-9]', '')
    if ($clean.Length -gt 60) { $clean = $clean.Substring(0, 60) }
    return $clean.ToLower()
}

function Get-GroupByName {
    param([string]$DisplayName)
    $escaped = $DisplayName.Replace("'", "''")
    $uri = "$GraphV1/groups?`$filter=displayName eq '$escaped'&`$select=id,displayName,groupTypes,membershipRule"
    $result = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    $matched = @($result.value)
    if ($matched.Count -gt 0) { return $matched[0] }
    return $null
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 6,
        [int]$DelaySeconds = 10,
        [string]$Activity = 'operation'
    )
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($i -eq $MaxAttempts) { throw }
            Write-Host "    Attempt $i of $MaxAttempts for $Activity failed. Retrying in ${DelaySeconds}s. ($($_.Exception.Message))" -ForegroundColor Yellow
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

#endregion --------------------------------------------------------------------

#region Connect ---------------------------------------------------------------

$scopes = @(
    'Group.ReadWrite.All'
    'DeviceManagementConfiguration.ReadWrite.All'
    'DeviceManagementScripts.ReadWrite.All'   # platform scripts are behind their own granular scope
    'Directory.Read.All'
    'Application.ReadWrite.All'
)

$connectParams = @{
    Scopes     = $scopes
    NoWelcome  = $true
}
if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Write-Step 'Connecting to Microsoft Graph'
Write-Host '    Sign in with a Global Administrator in the target tenant.' -ForegroundColor DarkGray

# Clear any cached token first so a previous client tenant is never picked up by accident.
Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

Connect-MgGraph @connectParams

$context = Get-MgContext
Write-Host "    Signed in as $($context.Account)" -ForegroundColor DarkGray
Write-Host "    Tenant: $($context.TenantId)" -ForegroundColor DarkGray

# Fail before creating anything if consent came back short. Partially consented sessions
# otherwise blow up mid-run, leaving groups behind without the policy that uses them.
$grantedScopes = @($context.Scopes)
$missingScopes = @($scopes | Where-Object { $grantedScopes -notcontains $_ })
if ($missingScopes.Count -gt 0) {
    throw "The following scopes were not granted: $($missingScopes -join ', '). " +
          "Run Disconnect-MgGraph and sign in again as Global Administrator to re-trigger the consent prompt."
}

$confirm = Read-Host "Create the Autopilot baseline for '$ClientPrefix' in tenant $($context.TenantId)? (y/n)"
if ($confirm -notmatch '^(y|yes)$') {
    Write-Host 'Aborted.' -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

#endregion --------------------------------------------------------------------

#region 1. Intune Provisioning Client service principal -----------------------

Write-Step 'Resolving the Intune Provisioning Client service principal'

$spUri = "$GraphV1/servicePrincipals?`$filter=appId eq '$ProvisioningClientAppId'&`$select=id,displayName,appId"
$spResult = Invoke-MgGraphRequest -Method GET -Uri $spUri -OutputType PSObject

# @() forces an array. Without it a single-element collection can flatten and [0] returns null.
$spMatches = @($spResult.value)

if ($spMatches.Count -eq 0) {
    Write-Host '    Service principal not present in this tenant. Creating it.' -ForegroundColor Yellow
    $spBody = @{ appId = $ProvisioningClientAppId } | ConvertTo-Json
    $sp = Invoke-MgGraphRequest -Method POST -Uri "$GraphV1/servicePrincipals" `
        -Body $spBody -ContentType 'application/json' -OutputType PSObject
    Start-Sleep -Seconds 15
}
else {
    $sp = $spMatches[0]
}

$provisioningClientObjectId = $sp.id

if ([string]::IsNullOrWhiteSpace($provisioningClientObjectId)) {
    throw "Could not resolve an object ID for the Intune Provisioning Client service principal (AppId $ProvisioningClientAppId). " +
          "Without it the device group cannot be owned correctly and the device preparation policy will not bind. " +
          "Check that the signed-in account holds Directory.Read.All and that the service principal exists in this tenant."
}

Write-Host "    $($sp.displayName) resolved to object $provisioningClientObjectId" -ForegroundColor DarkGray

#endregion --------------------------------------------------------------------

#region 2. Device group (assigned, SP-owned) ----------------------------------

$deviceGroupName = "$ClientPrefix - Autopilot Devices"
Write-Step "Device group: $deviceGroupName"

$deviceGroup = Get-GroupByName -DisplayName $deviceGroupName

if ($null -eq $deviceGroup) {
    $deviceGroupBody = @{
        displayName         = $deviceGroupName
        description         = 'Enrollment time grouping target for Windows Autopilot device preparation. Membership is written by the Autopilot service.'
        mailEnabled         = $false
        mailNickname        = New-MailNickname -DisplayName $deviceGroupName
        securityEnabled     = $true
    } | ConvertTo-Json -Depth 5

    $deviceGroup = Invoke-MgGraphRequest -Method POST -Uri "$GraphV1/groups" `
        -Body $deviceGroupBody -ContentType 'application/json' -OutputType PSObject

    Write-Host "    Created $($deviceGroup.id)" -ForegroundColor DarkGray
}
else {
    Write-Skip "Already exists: $($deviceGroup.id)"
}

$deviceGroupId = $deviceGroup.id

if ([string]::IsNullOrWhiteSpace($deviceGroupId)) {
    throw 'Device group was created or found but no object ID was returned. Cannot continue.'
}

# Owner is set as a separate call so a failure here is attributable rather than surfacing
# as a generic 404 on group creation. The policy will not bind without this owner.
Write-Step 'Ensuring the provisioning client owns the device group'

function Test-GroupOwner {
    param([string]$GroupId, [string]$OwnerObjectId)

    # GET /groups/{id}/owners does not reliably return service principal owners.
    # The derived-type cast is required. Both are checked so a user owner would also match.
    $uris = @(
        "$GraphV1/groups/$GroupId/owners/microsoft.graph.servicePrincipal"
        "$GraphV1/groups/$GroupId/owners"
    )

    foreach ($uri in $uris) {
        try {
            $result = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
            $owners = @($result.value)
            if ($owners.Count -gt 0) {
                if (@($owners | ForEach-Object { $_.id }) -contains $OwnerObjectId) { return $true }
            }
        }
        catch {
            # A cast that returns nothing can 404 on some tenants. Fall through to the next form.
            continue
        }
    }

    return $false
}

if (Test-GroupOwner -GroupId $deviceGroupId -OwnerObjectId $provisioningClientObjectId) {
    Write-Skip 'Provisioning client is already an owner.'
}
else {
    $refCandidates = @(
        "$GraphV1/directoryObjects/$provisioningClientObjectId"
        "$GraphV1/servicePrincipals/$provisioningClientObjectId"
    )

    $ownerConfirmed = $false
    $lastFailure    = $null

    foreach ($ref in $refCandidates) {
        try {
            $body = @{ '@odata.id' = $ref } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST `
                -Uri "$GraphV1/groups/$deviceGroupId/owners/`$ref" `
                -Body $body -ContentType 'application/json' | Out-Null

            $ownerConfirmed = $true
            Write-Host '    Owner added.' -ForegroundColor DarkGray
            break
        }
        catch {
            $detail = $_.ErrorDetails.Message
            if (-not $detail) { $detail = $_.Exception.Message }
            $lastFailure = $detail

            # Graph returns 400 Request_BadRequest when the reference is already there.
            # That is the desired end state, not a failure.
            if ($detail -match 'already exist') {
                $ownerConfirmed = $true
                Write-Skip 'Owner reference already present.'
                break
            }

            Write-Host "    Reference form rejected, trying next. ($detail)" -ForegroundColor Yellow
        }
    }

    if (-not $ownerConfirmed -and -not (Test-GroupOwner -GroupId $deviceGroupId -OwnerObjectId $provisioningClientObjectId)) {
        throw "Could not make the Intune Provisioning Client ($provisioningClientObjectId) an owner of " +
              "device group $deviceGroupId. The device preparation policy will show 0 groups assigned " +
              "without it. Last response: $lastFailure"
    }
}

#endregion --------------------------------------------------------------------

#region 3. User group ---------------------------------------------------------

$userGroupName = "$ClientPrefix - Autopilot Users"
Write-Step "User group: $userGroupName ($UserGroupType)"

$userGroup = Get-GroupByName -DisplayName $userGroupName

if ($null -eq $userGroup) {
    $userGroupHash = @{
        displayName         = $userGroupName
        description         = 'Assignment target for the Windows Autopilot device preparation policy.'
        mailEnabled         = $false
        mailNickname        = New-MailNickname -DisplayName $userGroupName
        securityEnabled     = $true
    }

    if ($UserGroupType -eq 'Dynamic') {
        # Scoped to users holding the Microsoft Intune service plan (INTUNE_A,
        # c1ec4a95-1f05-45b3-a911-aa3fa01094f5). E1 users are excluded automatically
        # since they cannot enroll. This also keeps the Entra ID P1 requirement for
        # dynamic membership aligned with users who already carry P1 via E3/E5/BP.
        $intuneServicePlanId = 'c1ec4a95-1f05-45b3-a911-aa3fa01094f5'
        $rule = "(user.accountEnabled -eq true) " +
                "-and (user.userType -eq `"Member`") " +
                "-and (user.assignedPlans -any (assignedPlan.servicePlanId -eq `"$intuneServicePlanId`" -and assignedPlan.capabilityStatus -eq `"Enabled`"))"

        $userGroupHash['groupTypes']                    = @('DynamicMembership')
        $userGroupHash['membershipRule']                = $rule
        $userGroupHash['membershipRuleProcessingState'] = 'On'
        $userGroupHash['description']                   = 'Assignment target for the Windows Autopilot device preparation policy. Members are enabled member accounts with an active Microsoft Intune service plan.'
    }

    $userGroup = Invoke-MgGraphRequest -Method POST -Uri "$GraphV1/groups" `
        -Body ($userGroupHash | ConvertTo-Json -Depth 5) -ContentType 'application/json' -OutputType PSObject

    Write-Host "    Created $($userGroup.id)" -ForegroundColor DarkGray
}
else {
    Write-Skip "Already exists: $($userGroup.id)"
}

$userGroupId = $userGroup.id

#endregion --------------------------------------------------------------------

#region 4. RMM platform script ------------------------------------------------

$scriptDisplayName = "$ClientPrefix - Install RMM Agent"
Write-Step "Platform script: $scriptDisplayName"

$existingScripts = Invoke-MgGraphRequest -Method GET `
    -Uri "$GraphBeta/deviceManagement/deviceManagementScripts?`$select=id,displayName" -OutputType PSObject
$platformScript  = @($existingScripts.value) | Where-Object { $_.displayName -eq $scriptDisplayName } | Select-Object -First 1

if ($null -eq $platformScript) {

    # Generated rather than read from disk so the only per-client variable is the site GUID.
    # -Wait is important: without it the process returns immediately and device preparation
    # marks the script complete while the agent is still installing.
    $installerSource = @"
`$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (Get-Service -Name 'CagService' -ErrorAction SilentlyContinue) {
    Write-Output 'Datto RMM agent already present. Nothing to do.'
    exit 0
}

`$installer = Join-Path `$env:TEMP 'AgentInstall.exe'

try {
    (New-Object System.Net.WebClient).DownloadFile(
        'https://$DattoPlatformHost/download-agent/windows/$DattoSiteGuid',
        `$installer
    )
}
catch {
    Write-Error "Agent download failed: `$(`$_.Exception.Message)"
    exit 1
}

`$proc = Start-Process -FilePath `$installer -Wait -PassThru
Remove-Item `$installer -Force -ErrorAction SilentlyContinue

if (`$proc.ExitCode -ne 0) {
    Write-Error "Agent installer returned exit code `$(`$proc.ExitCode)."
    exit `$proc.ExitCode
}

Write-Output 'Datto RMM agent installed.'
exit 0
"@

    $scriptContent = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($installerSource))

    $scriptBody = @{
        '@odata.type'         = '#microsoft.graph.deviceManagementScript'
        displayName           = $scriptDisplayName
        description           = "Installs the Datto RMM agent during Autopilot device preparation. Site $DattoSiteGuid on $DattoPlatformHost."
        scriptContent         = $scriptContent
        runAsAccount          = 'system'     # required: OOBE runs with no user signed in
        runAs32Bit            = $false       # required: 64-bit
        enforceSignatureCheck = $false
        fileName              = 'Install-DattoRMMAgent.ps1'
        roleScopeTagIds       = @('0')
    } | ConvertTo-Json -Depth 5

    $platformScript = Invoke-MgGraphRequest -Method POST -Uri "$GraphBeta/deviceManagement/deviceManagementScripts" `
        -Body $scriptBody -ContentType 'application/json' -OutputType PSObject

    Write-Host "    Created $($platformScript.id)" -ForegroundColor DarkGray
}
else {
    Write-Skip "Already exists: $($platformScript.id)"
}

$scriptId = $platformScript.id

# The script must ALSO be assigned to the device group or it is reported as Skipped during OOBE.
Write-Step 'Assigning platform script to the device group'

$scriptAssignBody = @{
    deviceManagementScriptAssignments = @(
        @{
            target = @{
                '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                groupId       = $deviceGroupId
            }
        }
    )
} | ConvertTo-Json -Depth 6

Invoke-WithRetry -Activity 'script assignment' -ScriptBlock {
    Invoke-MgGraphRequest -Method POST `
        -Uri "$GraphBeta/deviceManagement/deviceManagementScripts/$scriptId/assign" `
        -Body $scriptAssignBody -ContentType 'application/json'
} | Out-Null

#endregion --------------------------------------------------------------------

#region 5. Device preparation policy ------------------------------------------

$policyName = "$ClientPrefix - Autopilot Device Preparation"
Write-Step "Device preparation policy: $policyName"

$policyFilter = "(technologies has 'enrollment') and (templateReference/templateFamily eq 'enrollmentConfiguration')"
$existingPolicies = Invoke-MgGraphRequest -Method GET `
    -Uri "$GraphBeta/deviceManagement/configurationPolicies?`$select=id,name,templateReference&`$filter=$([uri]::EscapeDataString($policyFilter))" `
    -OutputType PSObject

$policy = @($existingPolicies.value) | Where-Object { $_.name -eq $policyName } | Select-Object -First 1

if ($null -ne $policy) {
    Write-Skip "Already exists: $($policy.id). Delete it manually to rebuild."
}
else {
    $accountTypeValue = if ($UserAccountType -eq 'Administrator') { 'enrollment_autopilot_dpp_accountype_0' } else { 'enrollment_autopilot_dpp_accountype_1' }
    $allowSkipValue   = if ($AllowUserToSkip)   { 'enrollment_autopilot_dpp_allowskip_1' }        else { 'enrollment_autopilot_dpp_allowskip_0' }
    $diagnosticsValue = if ($HideDiagnosticsLink) { 'enrollment_autopilot_dpp_allowdiagnostics_0' } else { 'enrollment_autopilot_dpp_allowdiagnostics_1' }
    $errorMessageJson = $ErrorMessageText | ConvertTo-Json

    $policyBody = @"
{
  "name": "$policyName",
  "description": "Baseline Autopilot device preparation policy. Created by New-AutopilotDevicePrep.ps1.",
  "platforms": "windows10",
  "technologies": "enrollment",
  "roleScopeTagIds": [ "0" ],
  "templateReference": { "templateId": "$DevicePrepTemplateId" },
  "settings": [
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_deploymentmode",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "5180aeab-886e-4589-97d4-40855c646315" },
        "choiceSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
          "children": [],
          "settingValueTemplateReference": { "settingValueTemplateId": "5874c2f6-bcf1-463b-a9eb-bee64e2f2d82" },
          "value": "enrollment_autopilot_dpp_deploymentmode_0"
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_deploymenttype",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "f4184296-fa9f-4b67-8b12-1723b3f8456b" },
        "choiceSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
          "children": [],
          "settingValueTemplateReference": { "settingValueTemplateId": "e0af022f-37f3-4a40-916d-1ab7281c88d9" },
          "value": "enrollment_autopilot_dpp_deploymenttype_0"
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_jointype",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "6310e95d-6cfa-4d2f-aae0-1e7af12e2182" },
        "choiceSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
          "children": [],
          "settingValueTemplateReference": { "settingValueTemplateId": "1fa84eb3-fcfa-4ed6-9687-0f3d486402c4" },
          "value": "enrollment_autopilot_dpp_jointype_0"
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_accountype",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "d4f2a840-86d5-4162-9a08-fa8cc608b94e" },
        "choiceSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
          "children": [],
          "settingValueTemplateReference": { "settingValueTemplateId": "bf13bb47-69ef-4e06-97c1-50c2859a49c2" },
          "value": "$accountTypeValue"
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_timeout",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "6dec0657-dfb8-4906-a7ee-3ac6ee1edecb" },
        "simpleSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue",
          "settingValueTemplateReference": { "settingValueTemplateId": "0bbcce5b-a55a-4e05-821a-94bf576d6cc8" },
          "value": $TimeoutMinutes
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_customerrormessage",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "2ddf0619-2b7a-46de-b29b-c6191e9dda6e" },
        "simpleSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationStringSettingValue",
          "settingValueTemplateReference": { "settingValueTemplateId": "fe5002d5-fbe9-4920-9e2d-26bfc4b4cc97" },
          "value": $errorMessageJson
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_allowskip",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "2a71dc89-0f17-4ba9-bb27-af2521d34710" },
        "choiceSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
          "children": [],
          "settingValueTemplateReference": { "settingValueTemplateId": "a2323e5e-ac56-4517-8847-b0a6fdb467e7" },
          "value": "$allowSkipValue"
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_allowdiagnostics",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "e2b7a81b-f243-4abd-bce3-c1856345f405" },
        "choiceSettingValue": {
          "@odata.type": "#microsoft.graph.deviceManagementConfigurationChoiceSettingValue",
          "children": [],
          "settingValueTemplateReference": { "settingValueTemplateId": "c59d26fd-3460-4b26-b47a-f7e202e7d5a3" },
          "value": "$diagnosticsValue"
        }
      }
    },
    {
      "@odata.type": "#microsoft.graph.deviceManagementConfigurationSetting",
      "settingInstance": {
        "@odata.type": "#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance",
        "settingDefinitionId": "enrollment_autopilot_dpp_allowedscriptids",
        "settingInstanceTemplateReference": { "settingInstanceTemplateId": "1bc67702-800c-4271-8fd9-609351cc19cf" },
        "simpleSettingCollectionValue": [
          {
            "@odata.type": "#microsoft.graph.deviceManagementConfigurationStringSettingValue",
            "value": "$scriptId"
          }
        ]
      }
    }
  ]
}
"@

    $policy = Invoke-MgGraphRequest -Method POST -Uri "$GraphBeta/deviceManagement/configurationPolicies" `
        -Body $policyBody -ContentType 'application/json' -OutputType PSObject

    Write-Host "    Created $($policy.id)" -ForegroundColor DarkGray
}

$policyId = $policy.id

#endregion --------------------------------------------------------------------

#region 6. Bind the device group (enrollment time grouping) -------------------

Write-Step 'Binding device group via assignJustInTimeConfiguration'

$jitBody = @{
    justInTimeAssignments = @{
        targetType = 'entraSecurityGroup'
        target     = @($deviceGroupId)
    }
} | ConvertTo-Json -Depth 5

Invoke-WithRetry -Activity 'JIT device group binding' -ScriptBlock {
    Invoke-MgGraphRequest -Method POST `
        -Uri "$GraphBeta/deviceManagement/configurationPolicies('$policyId')/assignJustInTimeConfiguration" `
        -Body $jitBody -ContentType 'application/json'
} | Out-Null

#endregion --------------------------------------------------------------------

#region 7. Assign the policy to the user group --------------------------------

Write-Step 'Assigning policy to the user group'

$assignBody = @{
    assignments = @(
        @{
            id     = ''
            source = 'direct'
            target = @{
                '@odata.type'                              = '#microsoft.graph.groupAssignmentTarget'
                groupId                                    = $userGroupId
                deviceAndAppManagementAssignmentFilterType = 'none'
            }
        }
    )
} | ConvertTo-Json -Depth 6

Invoke-WithRetry -Activity 'policy assignment' -ScriptBlock {
    Invoke-MgGraphRequest -Method POST `
        -Uri "$GraphBeta/deviceManagement/configurationPolicies('$policyId')/assign" `
        -Body $assignBody -ContentType 'application/json'
} | Out-Null

#endregion --------------------------------------------------------------------

#region Summary ---------------------------------------------------------------

Write-Host ''
Write-Host 'Autopilot device preparation baseline complete.' -ForegroundColor Green
[pscustomobject]@{
    Tenant          = $context.TenantId
    DeviceGroup     = $deviceGroupName
    DeviceGroupId   = $deviceGroupId
    UserGroup       = $userGroupName
    UserGroupId     = $userGroupId
    PlatformScript  = $scriptDisplayName
    ScriptId        = $scriptId
    Policy          = $policyName
    PolicyId        = $policyId
} | Format-List

Write-Host 'Remaining manual checks:' -ForegroundColor Yellow
Write-Host '  - Devices > Enrollment > Device platform restrictions: Windows MDM personally owned must be Allow,'
Write-Host '    otherwise you must upload corporate identifiers before enrollment will succeed.'
Write-Host '  - Confirm the policy shows 1 group assigned under Device group. If it shows 0, the provisioning'
Write-Host '    client ownership did not stick and the group must be recreated.'
Write-Host ''

Disconnect-MgGraph | Out-Null

#endregion --------------------------------------------------------------------