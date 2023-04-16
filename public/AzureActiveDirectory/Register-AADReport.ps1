<# Customer infos#>
function Get-OrganizationReport {
    $Organization = Get-MgOrganization -Property DisplayName, Id
    $script:CustomerName = $Organization.DisplayName
    return  "<h2>$($Organization.DisplayName) ($($Organization.Id))</h2>"
}

<# User settings policy section #>
function Get-UserSettingsReport {
    param(
        [System.Boolean]$DisableUserConsent,
        [System.Boolean]$DisableUsersToCreateAppRegistrations,
        [System.Boolean]$DisableUsersToReadOtherUsers,
        [System.Boolean]$DisableUsersToCreateSecurityGroups,
        [System.Boolean]$DisableUsersToCreateUnifiedGroups,
        [System.Boolean]$CreateUnifiedGroupCreationAllowedGroup,
        [System.Boolean]$EnableBlockMsolPowerShell
    )
    if ($DisableUserConsent) { Disable-ApplicationUserConsent }
    if ($DisableUsersToCreateAppRegistrations) { Disable-UsersToCreateAppRegistrations }
    if ($DisableUsersToReadOtherUsers) { Disable-UsersToReadOtherUsers }
    if ($DisableUsersToCreateSecurityGroups) { Disable-UsersToCreateSecurityGroups }
    if ($DisableUsersToCreateUnifiedGroups) { Disable-UsersToCreateUnifiedGroups }
    if ($CreateUnifiedGroupCreationAllowedGroup) { New-UnifiedGroupCreationAllowedGroup }
    if ($EnableBlockMsolPowerShell) { Enable-BlockMsolPowerShell }
    Write-Host "Checking user settings"
    $Policy = Get-MgPolicyAuthorizationPolicy -Property BlockMsolPowerShell, DefaultUserRolePermissions
    $Report = $Policy | Select-Object -Property @{Name = "PermissionGrantPoliciesAssigned"; Expression = { [string]$_.DefaultUserRolePermissions.PermissionGrantPoliciesAssigned } },
    @{Name = "AllowedToCreateApps"; Expression = { [string]$_.DefaultUserRolePermissions.AllowedToCreateApps } },
    @{Name = "AllowedToCreateSecurityGroups"; Expression = { [string]$_.DefaultUserRolePermissions.AllowedToCreateSecurityGroups } },
    @{Name = "AllowedToCreateUnifiedGroups"; Expression = { Request-AllowedToCreateUnifiedGroups } },
    @{Name = "AllowedToCreateUnifiedGroupsGroupName"; Expression = { Request-UnifiedGroupCreationAllowedGroup } },
    @{Name = "AllowedToReadOtherUsers"; Expression = { [string]$_.DefaultUserRolePermissions.AllowedToReadOtherUsers } },
    @{Name = "AllowedToCreateTenants"; Expression = { Request-AllowedToCreateTenants } },
    BlockMsolPowerShell | ConvertTo-Html -As List -Fragment -PreContent "<h3 id='AAD_USER_SETTINGS'>User settings</h3>" -PostContent "<p>PermissionGrantPoliciesAssigned: empty (user consent not allowed), microsoft-user-default-legacy (user consent allowed for all apps), microsoft-user-default-low (user consent allowed for low permission apps)</p><p>Unified groups = Microsoft 365 groups</p><p>AllowedToReadOtherUsers: Should only be disabled if you do not use Microsoft Planner</p><p>BlockMsolPowerShell: Should only be disabled if you do not use Azure AD Connect Sync</p>"

    $Report = $Report -Replace "<td>PermissionGrantPoliciesAssigned:</td><td>ManagePermissionGrantsForSelf.microsoft-user-default-legacy</td>", "<td>PermissionGrantPoliciesAssigned:</td><td class='red'>microsoft-user-default-legacy</td>"
    $Report = $Report -Replace "<td>PermissionGrantPoliciesAssigned:</td><td>ManagePermissionGrantsForSelf.microsoft-user-default-low</td>", "<td>PermissionGrantPoliciesAssigned:</td><td class='orange'>microsoft-user-default-low</td>"
    $Report = $Report -Replace "<td>AllowedToCreateApps:</td><td>True</td>", "<td>AllowedToCreateApps:</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>AllowedToCreateSecurityGroups:</td><td>True</td>", "<td>AllowedToCreateSecurityGroups:</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>AllowedToCreateUnifiedGroups:</td><td>True</td>", "<td>AllowedToCreateUnifiedGroups:</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>AllowedToCreateUnifiedGroups:</td><td>false</td>", "<td>AllowedToCreateUnifiedGroups:</td><td>False</td>"
    $Report = $Report -Replace "<td>AllowedToReadOtherUsers:</td><td>True</td>", "<td>AllowedToReadOtherUsers:</td><td class='orange'>True</td>"
    $Report = $Report -Replace "<td>AllowedToCreateTenants:</td><td>True</td>", "<td>AllowedToCreateTenants:</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>BlockMsolPowerShell:</td><td>False</td>", "<td>BlockMsolPowerShell:</td><td class='orange'>False</td>"
    return $Report
}
function Disable-ApplicationUserConsent {
    Write-Host "Disable enterprise application user consent"
    Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "PermissionGrantPoliciesAssigned" = @() }
}
function Disable-UsersToCreateAppRegistrations {
    Write-Host "Disable users to create app registrations"
    Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "AllowedToCreateApps" = $false }
}
function Disable-UsersToReadOtherUsers {
    Write-Host "Disable users to read other users"
    Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "AllowedToReadOtherUsers" = $false }
}
function Disable-UsersToCreateSecurityGroups {
    Write-Host "Disable users to create security groups"
    Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "AllowedToCreateSecurityGroups" = $false }
}
function Enable-BlockMsolPowerShell {
    Write-Host "Disable legacy MsolPowerShell access"
    Update-MgPolicyAuthorizationPolicy -BlockMsolPowerShell
}
function Request-AllowedToCreateUnifiedGroups {
    if ($GroupSettingsUnified = Get-GroupSettingsUnified) {
        return ($GroupSettingsUnified.values | Where-Object name -EQ "EnableGroupCreation").value
    }
    return $true
}
function Request-UnifiedGroupCreationAllowedGroup {
    $GroupSettingsUnified = Get-GroupSettingsUnified
    if ($GroupId = ($GroupSettingsUnified.values | Where-Object name -EQ "GroupCreationAllowedGroupId").value) {
        return (Get-MgGroup -GroupId $GroupId -Property DisplayName).DisplayName
    }
}
function New-UnifiedGroupCreationAllowedGroup {
    Write-Host "Creating UnifiedGroupCreationAllowed group:" $script:UnifiedGroupCreationAllowedGroupName
    if (Request-UnifiedGroupCreationAllowedGroup) {
        Write-Host "UnifiedGroupCreationAllowed group already assigned"
        return
    }
    if ($GroupId = (Get-MgGroup -Property Id, DisplayName -Filter "DisplayName eq '$($script:UnifiedGroupCreationAllowedGroupName)'").Id) { Write-Host "UnifiedGroupCreationAllowed group already exists" } else { $GroupId = (New-MgGroup -DisplayName $script:UnifiedGroupCreationAllowedGroupName -MailEnabled:$False -MailNickname $script:UnifiedGroupCreationAllowedGroupName -SecurityEnabled).Id }
    $Body = @{
        templateId = (Get-GroupSettingsTemplateUnified).id
        values     = @( @{ Name = "GroupCreationAllowedGroupId" ; Value = $GroupId } )
    }
    if ($GroupSettingsUnified = Get-GroupSettingsUnified) { Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/groupSettings/$($GroupSettingsUnified.id)" -Body $Body }
    else { Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groupSettings" -Body $Body }
}
function Get-GroupSettingsTemplateUnified {
    $GroupSettingTemplates = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groupSettingTemplates?$select=value"
    return $GroupSettingTemplates.value | Where-Object { $_.displayName -eq "Group.Unified" } | Select-Object -Property id, DisplayName
}
function Get-GroupSettingsUnified {
    $GroupSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groupSettings?$select=value" 
    return $GroupSettings.value | Where-Object { $_.templateId -eq (Get-GroupSettingsTemplateUnified).id } | Select-Object -Property id, templateId, values
}
function Disable-UsersToCreateUnifiedGroups {
    $Body = @{
        templateId = (Get-GroupSettingsTemplateUnified).id
        values     = @( @{ Name = "EnableGroupCreation" ; Value = "false" } )
    }
    if ($GroupSettingsUnified = Get-GroupSettingsUnified) { Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/groupSettings/$($GroupSettingsUnified.id)" -Body $Body }
    else { Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groupSettings" -Body $Body }
}
function Request-AllowedToCreateTenants {
    $DefaultUserRolePermissions = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/authorizationPolicy/authorizationPolicy").defaultUserRolePermissions
    return $DefaultUserRolePermissions["allowedToCreateTenants"]
}

<# Device join settings#>
function Get-DeviceJoinSettingsReport {
    Write-Host "Checking device join settings"
    $DeviceSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy?$select=azureAdJoin, userDeviceQuota, multiFactorAuthConfiguration"
    $Report = $DeviceSettings | Select-Object -Property @{Name = "Require MFA"; Expression = { if ($_.multiFactorAuthConfiguration -eq 1) { return $true } return $false } },
    @{Name = "Allowed to join"; Expression = { if ($_.azureAdJoin.appliesTo -eq 1) { return "All users" } if ($_.azureAdJoin.appliesTo -eq 2) { return "Selected users" } return "No users" } },
    @{Name = "Users"; Expression = { $Users = @() ; foreach ($UserId in $_.azureAdJoin.allowedUsers) { $Users += (Get-MgUser -UserId $UserId -Property UserPrincipalName).UserPrincipalName } return $Users -join "; " } },
    @{Name = "Groups"; Expression = { $Groups = @() ; foreach ($GroupId in $_.azureAdJoin.allowedGroups) { $Groups += (Get-MgGroup -GroupId $GroupId -Property DisplayName).DisplayName } return $Groups -join "; " } },
    userDeviceQuota | ConvertTo-Html -As List -Fragment -PreContent "<br><h3 id='AAD_DEVICE_JOIN_SETTINGS'>Device join settings</h3>"
    $Report = $Report -Replace "<td>Require MFA:</td><td>False</td>", "<td>Require MFA:</td><td class='red'>False</td>"
    $Report = $Report -Replace "<td>Allowed to join:</td><td>All users</td>", "<td>Allowed to join:</td><td class='red'>All users</td>"
    return $Report
}

<# License SKU section#>
function Get-UsedSKUReport {
    Write-Host "Checking licenses"
    $SKU = Get-MgSubscribedSku -Property SkuPartNumber, ConsumedUnits, PrepaidUnits, AppliesTo
    return $SKU | Select-Object -Property @{Name = "Name"; Expression = { 
            if ($_.SkuPartNumber -eq "EXCHANGESTANDARD") { return "Exchange Online (Plan 1)" }
            if ($_.SkuPartNumber -eq "EXCHANGEENTERPRISE") { return "Exchange Online (PLAN 2)" }
            if ($_.SkuPartNumber -eq "EXCHANGEARCHIVE_ADDON") { return "Exchange Online Archiving for Exchange Online" }
            if ($_.SkuPartNumber -eq "EXCHANGE_S_ESSENTIALS") { return "Exchange Online Essentials" }

            if ($_.SkuPartNumber -eq "SHAREPOINTSTANDARD") { return "SharePoint Online (Plan 1)" }
            if ($_.SkuPartNumber -eq "SHAREPOINTENTERPRISE") { return "SharePoint Online (Plan 2)" }

            if ($_.SkuPartNumber -eq "AAD_BASIC") { return "Azure Active Directory Basic" }
            if ($_.SkuPartNumber -eq "AAD_PREMIUM") { return "Azure Active Directory Premium P1" }
            if ($_.SkuPartNumber -eq "AAD_PREMIUM_P2") { return "Azure Active Directory Premium P2" }

            if ($_.SkuPartNumber -eq "EMS") { return "Enterprise Mobility + Security E3" }
            if ($_.SkuPartNumber -eq "EMSPREMIUM") { return "Enterprise Mobility + Security E5" }

            if ($_.SkuPartNumber -eq "INTUNE_A") { return "Intune" }
            if ($_.SkuPartNumber -eq "INTUNE_A_D") { return "Microsoft Intune Device" }
            if ($_.SkuPartNumber -eq "INTUNE_SMB") { return "Microsoft Intune SMB" }

            if ($_.SkuPartNumber -eq "WINDOWS_STORE") { return "Windows Store for Business" }
            if ($_.SkuPartNumber -eq "RMSBASIC") { return "Rights Management Service Basic Content Protection" }
            if ($_.SkuPartNumber -eq "RIGHTSMANAGEMENT_ADHOC") { return "Rights Management Adhoc" }

            if ($_.SkuPartNumber -eq "VISIO_PLAN1_DEPT") { return "Visio Plan 1" }
            if ($_.SkuPartNumber -eq "VISIO_PLAN2_DEPT	") { return "Visio Plan 2" }
            if ($_.SkuPartNumber -eq "VISIOONLINE_PLAN1") { return "Visio Online Plan 1" }
            if ($_.SkuPartNumber -eq "VISIOCLIENT") { return "Visio Online Plan 2" }

            if ($_.SkuPartNumber -eq "PROJECTESSENTIALS") { return "Project Online Essentials" }
            if ($_.SkuPartNumber -eq "PROJECTPREMIUM") { return "Project Online Premium" }
            if ($_.SkuPartNumber -eq "PROJECT_P1") { return "Project Plan 1" }
            if ($_.SkuPartNumber -eq "PROJECTPROFESSIONAL") { return "Project Plan 3" }

            if ($_.SkuPartNumber -eq "MS_TEAMS_IW") { return "Microsoft Teams Trial" }
            if ($_.SkuPartNumber -eq "MCOCAP") { return "Microsoft Teams Shared Devices" }
            if ($_.SkuPartNumber -eq "MCOEV") { return "Microsoft Teams Phone Standard" }
            if ($_.SkuPartNumber -eq "MCOEV_DOD") { return "Microsoft Teams Phone Standard for DOD" }
            if ($_.SkuPartNumber -eq "MCOTEAMS_ESSENTIALS") { return "Teams Phone with Calling Plan" }
            if ($_.SkuPartNumber -eq "TEAMS_FREE") { return "Microsoft Teams (Free)" }
            if ($_.SkuPartNumber -eq "Teams_Ess") { return "Microsoft Teams Essentials" }
            if ($_.SkuPartNumber -eq "Microsoft_Teams_Premium") { return "Microsoft Teams Premium" }
            if ($_.SkuPartNumber -eq "TEAMS_EXPLORATORY") { return "Microsoft Teams Exploratory" }
            if ($_.SkuPartNumber -eq "BUSINESS_VOICE_DIRECTROUTING") { return "Microsoft 365 Business Voice (without calling plan)" }
            if ($_.SkuPartNumber -eq "PHONESYSTEM_VIRTUALUSER") { return "Microsoft Teams Phone Resoure Account" }
            if ($_.SkuPartNumber -eq "Microsoft_Teams_Rooms_Basic_without_Audio_Conferencing") { return "Microsoft Teams Rooms Basic without Audio Conferencing" }
            if ($_.SkuPartNumber -eq "Microsoft_Teams_Rooms_Pro") { return "Microsoft Teams Rooms Pro" }
            if ($_.SkuPartNumber -eq "BUSINESS_VOICE_MED2") { return "Microsoft 365 Business Voice" }
            if ($_.SkuPartNumber -eq "MCOPSTN_5") { return "Microsoft 365 Domestic Calling Plan (120 Minutes)" }

            if ($_.SkuPartNumber -eq "POWER_BI_PRO") { return "Power BI Pro" }
            if ($_.SkuPartNumber -eq "POWERAPPS_VIRAL") { return "Microsoft Power Apps Plan 2 Trial" }
            if ($_.SkuPartNumber -eq "SPZA_IW") { return "App Connect IW" }
            if ($_.SkuPartNumber -eq "FLOW_FREE") { return "Microsoft Flow Free" }
            if ($_.SkuPartNumber -eq "CCIBOTS_PRIVPREV_VIRAL") { return "Power Virtual Agents Viral Trial" }
            if ($_.SkuPartNumber -eq "VIRTUAL_AGENT_BASE") { return "Power Virtual Agent" }

            if ($_.SkuPartNumber -eq "WIN_DEF_ATP") { return "Microsoft Defender for Endpoint" }
            if ($_.SkuPartNumber -eq "ADALLOM_STANDALONE") { return "Microsoft Cloud App Security" }
            if ($_.SkuPartNumber -eq "DEFENDER_ENDPOINT_P1") { return "Microsoft Defender for Endpoint P1" }
            if ($_.SkuPartNumber -eq "MDATP_Server") { return "Microsoft Defender for Endpoint Server" }
            if ($_.SkuPartNumber -eq "ATP_ENTERPRISE_FACULTY") { return "Microsoft Defender for Office 365 (Plan 1) Faculty" }
            if ($_.SkuPartNumber -eq "ATA") { return "Microsoft Defender for Identity" }
            if ($_.SkuPartNumber -eq "ATP_ENTERPRISE") { return "Microsoft Defender for Office 365 (Plan 1)" }

            if ($_.SkuPartNumber -eq "M365_F1") { return "Microsoft 365 F1" }
            if ($_.SkuPartNumber -eq "SPE_F1") { return "Microsoft 365 F3" }
            if ($_.SkuPartNumber -eq "DESKLESSPACK") { return "Office 365 F3" }

            if ($_.SkuPartNumber -eq "SPE_E3") { return "Microsoft 365 E3" }
            if ($_.SkuPartNumber -eq "SPE_E5") { return "Microsoft 365 E5" }
            if ($_.SkuPartNumber -eq "SPE_E5_CALLINGMINUTES") { return "Microsoft 365 E5 with Calling Minutes" }
            if ($_.SkuPartNumber -eq "INFORMATION_PROTECTION_COMPLIANCE") { return "Microsoft 365 E5 Compliance" }
            if ($_.SkuPartNumber -eq "IDENTITY_THREAT_PROTECTION") { return "Microsoft 365 E5 Security" }
            if ($_.SkuPartNumber -eq "SPE_E5_NOPSTNCONF") { return "Microsoft 365 E5 without Audio Conferencing" }
        
            if ($_.SkuPartNumber -eq "STANDARDPACK") { return "Office 365 E1" }
            if ($_.SkuPartNumber -eq "STANDARDWOFFPACK") { return "Office 365 E2" }
            if ($_.SkuPartNumber -eq "ENTERPRISEPACK") { return "Office 365 E3" }
            if ($_.SkuPartNumber -eq "ENTERPRISEWITHSCAL") { return "Office 365 E4" }
            if ($_.SkuPartNumber -eq "ENTERPRISEPREMIUM") { return "Office 365 E5" }
            if ($_.SkuPartNumber -eq "ENTERPRISEPREMIUM_NOPSTNCONF") { return "Office 365 E5 without Audio Conferencing" }

            if ($_.SkuPartNumber -eq "SPB") { return "Microsoft 365 Business Premium" }
            if ($_.SkuPartNumber -eq "O365_BUSINESS_PREMIUM") { return "Microsoft 365 Business Standard" }
            if (($_.SkuPartNumber -eq "O365_BUSINESS") -or ($_.SkuPartNumber -eq "SMB_BUSINESS")) { return "Microsoft 365 Apps for Business" }
            if ($_.SkuPartNumber -eq "OFFICESUBSCRIPTION") { return "Microsoft 365 Apps for enterprise" }
            if (($_.SkuPartNumber -eq "O365_BUSINESS_ESSENTIALS") -or ($_.SkuPartNumber -eq "SMB_BUSINESS_ESSENTIALS")) { return "Microsoft 365 Business Basic" }
            else { return $_.SkuPartNumber }
        } 
    }, @{Name = "Total"; Expression = { $_.PrepaidUnits.Enabled } }, @{Name = "Assigned"; Expression = { $_.ConsumedUnits } } , @{Name = "Available"; Expression = { ($_.PrepaidUnits.Enabled) - ($_.ConsumedUnits) } } , AppliesTo | ConvertTo-Html -As Table -Fragment -PreContent "<br><h3 id='AAD_SKU'>Licenses</h3>"
}

<# User Account section #>
function Disable-UserAccount {
    param (
        $Users
    )
    $params = @{
        AccountEnabled = "false"
    }
    Write-Host "Disable user accounts"
    foreach ($User in $Users) {
        Update-MgUser -UserId $User.UserPrincipalName -BodyParameter $params
    }
}
function Request-UserAccountStatus {
    param(
        $UserId
    )
    return (Get-MgUser -UserId $UserId -Property AccountEnabled).AccountEnabled
}

<# Admin role section #>
function Get-AdminRoleReport {
    Write-Host "Checking admin role assignments"
    $Assignments = Get-MgRoleManagementDirectoryRoleAssignment -Property PrincipalId, RoleDefinitionId
    foreach ($Assignment in $Assignments) {
        $ProcessedCount++
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Assignment.PrincipalId)"
        if ($User = Get-MgUser -UserId $Assignment.PrincipalId -Property DisplayName, UserPrincipalName -ErrorAction SilentlyContinue) {
            $Assignment | Add-Member -NotePropertyName "DisplayName" -NotePropertyValue $User.DisplayName
            $Assignment | Add-Member -NotePropertyName "UserPrincipalName" -NotePropertyValue $User.UserPrincipalName
            $Assignment | Add-Member -NotePropertyName "RoleName" -NotePropertyValue (Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $Assignment.RoleDefinitionId -Property DisplayName).DisplayName
        }
    }
    Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Assignment.PrincipalId)" -Status "Ready" -Completed
    return $Assignments | Where-Object { -not ($null -eq $_.DisplayName) } | Sort-Object -Property UserPrincipalName | ConvertTo-Html -Property DisplayName, UserPrincipalName, RoleName -As Table -Fragment -PreContent "<br><h3 id='AAD_ADMINS'>Admin role assignments</h3>"
}

<# BreakGlass account Section #>
function Get-BreakGlassAccountReport {
    param (
        $Create
    )
    if ($BGAccount = Get-BreakGlassAccount) {
        $Report = $BGAccount | ConvertTo-Html -Property DisplayName, UserPrincipalName, AccountEnabled, GlobalAdmin, LastSignIn -As Table -Fragment -PreContent "<br><h3 id='AAD_BG'>BreakGlass account</h3>"
        $Report = $Report -Replace "<td>False</td>", "<td class='red'>False</td>"
        return $Report

    }
    if ($create) {
        New-BreakGlassAccount
        $Report = Get-BreakGlassAccount | ConvertTo-Html -Property DisplayName, UserPrincipalName, AccountEnabled, GlobalAdmin, LastSignIn -As Table -Fragment -PreContent "<br><h3 id='AAD_BG'>BreakGlass account</h3>" -PostContent "<p>Check console log for credentials</p>"
        $Report = $Report -Replace "<td>False</td>", "<td class='red'>False</td>"
        return $Report
    }
    return "<br><h3 id='AAD_BG'>BreakGlass account</h3><p>Not found</p>"
}
function Get-BreakGlassAccount {
    Write-Host "Checking BreakGlass account"
    Select-MgProfile -Name "beta"
    $BGAccounts = Get-MgUser -Filter "startswith(displayName, 'BreakGlass ')" -Property Id, DisplayName, UserPrincipalName, AccountEnabled, SignInActivity
    Select-MgProfile -Name "v1.0"
    if (-not $bgAccounts) { 
        $BGAccounts = Get-MgUser -Filter "startswith(displayName, 'BreakGlass ')" -Property Id, DisplayName, UserPrincipalName, AccountEnabled
    }
    if (-not $bgAccounts) { return }
    foreach ($BGAccount in $BGAccounts) {
        Add-Member -InputObject $BGAccount -NotePropertyName "GlobalAdmin" -NotePropertyValue (Request-GlobalAdminRole $BGAccount.Id)
        Add-Member -InputObject $BGAccount -NotePropertyName "LastSignIn" -NotePropertyValue $BGAccount.SignInActivity.LastSignInDateTime
    }
    return $BGAccounts
}
function Get-GlobalAdminRoleId {
    return (Get-MgDirectoryRole -Filter "DisplayName eq 'Global Administrator'" -Property Id).Id
}
function Request-GlobalAdminRole {
    param (
        $AccountId
    )
    if (Get-MgDirectoryRoleMember -DirectoryRoleId (Get-GlobalAdminRoleId) -Filter "id eq '$($AccountId)'") {
        return $true
    }
}
function New-BreakGlassAccount {
    Write-Host "Creating BreakGlass account:"
    $Name = -join ((97..122) | Get-Random -Count 64 | ForEach-Object { [char]$_ })
    $DisplayName = "BreakGlass $Name"
    $Domain = (Get-MgDomain -Property id, IsInitial | Where-Object { $_.IsInitial -eq $true }).Id
    $UPN = "$Name@$Domain"
    $PasswordProfile = @{
        ForceChangePasswordNextSignIn        = $false
        ForceChangePasswordNextSignInWithMfa = $false
        Password                             = New-Password
    }
    $BGAccount = New-MgUser -DisplayName $DisplayName -UserPrincipalName $UPN -MailNickname $Name -PasswordProfile $PasswordProfile -PreferredLanguage "en-US" -AccountEnabled
    $DirObject = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($BGAccount.id)"
    }
    New-MgDirectoryRoleMemberByRef -DirectoryRoleId (Get-GlobalAdminRoleId) -BodyParameter $DirObject
    Add-Member -InputObject $BGAccount -NotePropertyName "Password" -NotePropertyValue $PasswordProfile.Password
    Write-Host ($BGAccount | Select-Object -Property Id, DisplayName, UserPrincipalName, Password | Format-List | Out-String)
}
function New-Password {
    param (
        [ValidateRange(4, [int]::MaxValue)]
        [int] $Length = 64,
        [int] $Upper = 4,
        [int] $Lower = 4,
        [int] $Numeric = 4,
        [int] $Special = 4
    )
    if ($Upper + $Lower + $Numeric + $Special -gt $Length) {
        throw "number of Upper/Lower/Numeric/Special char must be Lower or equal to Length"
    }
    $UpperCaseCharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $LowerCaseCharSet = "abcdefghijklmnopqrstuvwxyz"
    $NumberCharSet = "0123456789"
    $SpecialCharCharSet = "/*-+, !?=()@; :._"
    $CharSet = ""
    if ($Upper -gt 0) { $CharSet += $UpperCaseCharSet }
    if ($Lower -gt 0) { $CharSet += $LowerCaseCharSet }
    if ($Numeric -gt 0) { $CharSet += $NumberCharSet }
    if ($Special -gt 0) { $CharSet += $SpecialCharCharSet }
    $CharSet = $CharSet.ToCharArray()
    $RNG = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $Bytes = New-Object byte[]($Length)
    $RNG.GetBytes($Bytes)
    $Result = New-Object char[]($Length)
    for ($i = 0 ; $i -lt $Length ; $i++) {
        $Result[$i] = $CharSet[$Bytes[$i] % $CharSet.Length]
    }
    $Password = (-join $Result)
    $Valid = $true
    if ($Upper -gt ($Password.ToCharArray() | Where-Object { $_ -cin $UpperCaseCharSet.ToCharArray() }).Count) { $Valid = $false }
    if ($Lower -gt ($Password.ToCharArray() | Where-Object { $_ -cin $LowerCaseCharSet.ToCharArray() }).Count) { $Valid = $false }
    if ($Numeric -gt ($Password.ToCharArray() | Where-Object { $_ -cin $NumberCharSet.ToCharArray() }).Count) { $Valid = $false }
    if ($Special -gt ($Password.ToCharArray() | Where-Object { $_ -cin $SpecialCharCharSet.ToCharArray() }).Count) { $Valid = $false }
    if (!$Valid) {
        $Password = Get-RandomPassword $Length $Upper $Lower $Numeric $Special
    }
    return $Password
}

<# User MFA section#>
function Get-UserMfaStatusReport {
    Write-Host "Checking user MFA status"
    $Users = Get-MgUser -All -Filter "UserType eq 'Member'" -Property Id, DisplayName, UserPrincipalName, AssignedLicenses, AccountEnabled
    $Users | ForEach-Object {
        $ProcessedCount++
        if (($_.AssignedLicenses).Count -ne 0) {
            $LicenseStatus = "Licensed"
        }
        else {
            $LicenseStatus = "Unlicensed"
        }
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($_.DisplayName)"
        [array]$MFAData = Get-MgUserAuthenticationMethod -UserId $_.Id
        $AuthenticationMethod = @()
        $AdditionalDetails = @()
        foreach ($MFA in $MFAData) {
            switch ($MFA.AdditionalProperties["@odata.type"]) { 
                "#microsoft.graph.passwordAuthenticationMethod" {
                    $AuthMethod = 'PasswordAuthentication'
                    $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
                } 
                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                    $AuthMethod = 'AuthenticatorApp'
                    $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
                }
                "#microsoft.graph.phoneAuthenticationMethod" {
                    $AuthMethod = 'PhoneAuthentication'
                    $AuthMethodDetails = $MFA.AdditionalProperties["phoneType", "phoneNumber"] -join ' '
                } 
                "#microsoft.graph.fido2AuthenticationMethod" {
                    $AuthMethod = 'Fido2'
                    $AuthMethodDetails = $MFA.AdditionalProperties["model"] 
                }  
                "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                    $AuthMethod = 'WindowsHelloForBusiness'
                    $AuthMethodDetails = $MFA.AdditionalProperties["displayName"] 
                }                        
                "#microsoft.graph.emailAuthenticationMethod" {
                    $AuthMethod = 'EmailAuthentication'
                    $AuthMethodDetails = $MFA.AdditionalProperties["emailAddress"] 
                }               
                "microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                    $AuthMethod = 'TemporaryAccessPass'
                    $AuthMethodDetails = 'Access pass lifetime (minutes): ' + $MFA.AdditionalProperties["lifetimeInMinutes"] 
                }
                "#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
                    $AuthMethod = 'PasswordlessMSAuthenticator'
                    $AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
                }
                "#microsoft.graph.softwareOathAuthenticationMethod" {
                    $AuthMethod = 'SoftwareOath'
                    $AuthMethodDetails = 'Third-party OTP app'
                }
            }
            $AuthenticationMethod += $AuthMethod
            if ($null -ne $AuthMethodDetails) {
                $AdditionalDetails += "$($AuthMethod): $AuthMethodDetails"
            }
        }
        $AuthenticationMethod = $AuthenticationMethod | Sort-Object | Get-Unique
        $AdditionalDetail = $AdditionalDetails -join ', '
        [array]$StrongMFAMethods = ("Fido2", "PasswordlessMSAuthenticator", "AuthenticatorApp", "WindowsHelloForBusiness", "SoftwareOath")
        $MFAStatus = "Disabled"
        foreach ($StrongMFAMethod in $StrongMFAMethods) {
            if ($AuthenticationMethod -contains $StrongMFAMethod) {
                $MFAStatus = "Strong"
                break
            }
        }
        if ( ($AuthenticationMethod -contains "PhoneAuthentication") -or ($AuthenticationMethod -contains "EmailAuthentication")) {
            $MFAStatus = "Weak"
        }
        Add-Member -InputObject $_ -NotePropertyName "LicenseStatus" -NotePropertyValue $LicenseStatus
        Add-Member -InputObject $_ -NotePropertyName "MFAStatus" -NotePropertyValue $MFAStatus
        Add-Member -InputObject $_ -NotePropertyName "AdditionalDetail" -NotePropertyValue $AdditionalDetail
    }
    Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($_.DisplayName)" -Status "Ready" -Completed
    $Report = $Users | Sort-Object -Property UserPrincipalName | ConvertTo-Html -Property DisplayName, UserPrincipalName, LicenseStatus, AccountEnabled, MFAStatus, AdditionalDetail -As Table -Fragment -PreContent "<br><h3 id='AAD_MFA'>User MFA status</h3>" -PostContent "<p>Weak: PhoneAuthentication, EmailAuthentication</p><p>Strong: Fido2, PasswordlessMSAuthenticator, AuthenticatorApp, WindowsHelloForBusiness, SoftwareOath</p>"
    $Report = $Report -Replace "<td>True</td><td>Disabled</td>", "<td>True</td><td class='red'>Disabled</td>"
    $Report = $Report -Replace "<td>True</td><td>Weak</td>", "<td>True</td><td class='orange'>Weak</td>"
    return $Report
}

<# Guest user section#>
function Get-GuestUserReport {
    Write-Host "Checking guest accounts"
    Select-MgProfile -Name "beta"
    $Users = Get-MgUser -All -Filter "UserType eq 'Guest'" -Property Id, DisplayName, UserPrincipalName, AccountEnabled, SignInActivity
    Select-MgProfile -Name "v1.0"
    if (-not $Users) {
        return "<br><h3 id='AAD_GUEST'>Guest accounts</h3><p>Not found</p>"
    }
    return $Users | Select-Object -Property DisplayName, UserPrincipalName, AccountEnabled, @{Name = "LastSignIn"; Expression = { $_.SignInActivity.LastSignInDateTime } } | Sort-Object -Property LastSignIn | ConvertTo-Html -As Table -Fragment -PreContent "<br><h3 id='AAD_GUEST'>Guest accounts</h3>"
}

<# Security defaults section #>
function Get-SecurityDefaultsReport {
    param (
        [System.Boolean]$EnableSecurityDefaults,
        [System.Boolean]$DisableSecurityDefaults
    )
    if ($EnableSecurityDefaults -and (-not $DisableSecurityDefaults)) {
        Update-SecurityDefaults -Enable $true
    }  
    if ($DisableSecurityDefaults -and (-not $EnableSecurityDefaults)) {
        Update-SecurityDefaults -Enable $false
    }
    if (Request-SecurityDefaults) {
        return "<br><h3 id='AAD_SEC_DEFAULTS'>Security defaults</h3><p>Enabled</p>"
    }
    return "<br><h3 id='AAD_SEC_DEFAULTS'>Security defaults</h3><p>Disabled</p>"
}
function Request-SecurityDefaults {
    Write-Host "Checking security defaults"
    return (Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -Property "isEnabled").IsEnabled
}
function Update-SecurityDefaults {
    param ([System.Boolean]$Enable)
    $params = @{
        IsEnabled = $Enable
    }
    Write-Host "Updating security defaults enable:" $Enable
    Update-MgPolicyIdentitySecurityDefaultEnforcementPolicy -BodyParameter $params
}

<# Conditional access section #>
function Get-ConditionalAccessPolicyReport {
    Write-Host "Checking conditional access policies"
    if ($Policy = Get-MgIdentityConditionalAccessPolicy -Property Id, DisplayName, State) {
        return $Policy | ConvertTo-Html -Property DisplayName, Id, State -As Table -Fragment -PreContent "<br><h3 id='AAD_CA'>Conditional access policies</h3>"
    }
    return "<br><h3 id='AAD_CA'>Conditional access policies</h3><p>Not found</p>"
}
function Get-NamedLocationReport {
    Write-Host "Checking named locations"
    if ($Locations = Get-MgIdentityConditionalAccessNamedLocation) {
        return $Locations | Select-Object -Property DisplayName, @{Name = "Trusted"; Expression = { $_.additionalProperties["isTrusted"] } }, @{Name = "IPRange"; Expression = {
                $IpRangesReport = @()
                foreach ($IpRange in ($_.additionalProperties["ipRanges"])) {
                    $IpRangesReport += $IpRange["cidrAddress"]
                }
                return $IpRangesReport
            }
        }, @{Name = "Countries"; Expression = { $_.additionalProperties["countriesAndRegions"] } } | ConvertTo-Html -As Table -Fragment -PreContent "<br><h3 id='AAD_CA_LOCATIONS'>Named locations</h3>"
    }
    return "<br><h3 id='AAD_CA_LOCATIONS'>Named locations</h3><p>Not found</p>"
}

<# Application protection polices section#>
function Get-AppProtectionPolicesReport {
    Write-Host "Checking app protection policies"
    if ($Polices = Get-AppProtectionPolices) {
        return $Polices | ConvertTo-Html -As Table -Property DisplayName, IsAssigned -Fragment -PreContent "<br><h3 id='AAD_APP_POLICY'>App protection policies</h3>"
    }
    return "<br><h3 id='AAD_APP_POLICY'>App protection policies</h3><p>Not found</p>"
}
function Get-AppProtectionPolices {
    $IOSPolicies = Get-MgDeviceAppManagementiOSManagedAppProtection -Property DisplayName, IsAssigned
    $AndroidPolicies = Get-MgDeviceAppManagementAndroidManagedAppProtection -Property DisplayName, IsAssigned
    $Policies = @()
    $Policies += $IOSPolicies
    $Policies += $AndroidPolicies
    return $Policies
}