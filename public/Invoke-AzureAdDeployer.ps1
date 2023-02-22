function Invoke-AzureAdDeployer {
    [CmdletBinding()]
    Param(
        [switch]$UseExistingExoSession,
        [switch]$KeepExoSessionAlive,
        [switch]$UseExistingGraphSession,
        [switch]$KeepGraphSessionAlive,
        [switch]$UseExistingSpoSession,
        [switch]$KeepSpoSessionAlive,
        [switch]$AddExchangeOnlineReport,
        [switch]$AddSharePointOnlineReport,
        [switch]$CreateBreakGlassAccount,
        [switch]$EnableSecurityDefaults,
        [switch]$DisableSecurityDefaults,
        [switch]$DisableEnterpiseApplicationUserConsent,
        [switch]$DisableUsersToCreateAppRegistrations,
        [switch]$DisableUsersToReadOtherUsers,
        [switch]$DisableUsersToCreateSecurityGroups,
        [switch]$DisableUsersToCreateUnifiedGroups,
        [switch]$CreateUnifiedGroupCreationAllowedGroup,
        [switch]$EnableBlockMsolPowerShell,
        [switch]$SetMailboxLanguage,
        [switch]$DisableSharedMailboxLogin,
        [switch]$EnableSharedMailboxCopyToSent,
        [switch]$HideUnifiedMailboxFromOutlookClient,
        [switch]$DisableAddToOneDrive
    )
    $ReportTitle = "Microsoft 365 Security Report"
    $Version = "2.15.5"
    $script:VersionMessage = "AzureAdDeployer version: $($Version)"

    $ReportImageUrl = "https://cdn-icons-png.flaticon.com/512/3540/3540926.png"

    $script:InteractiveMode = $false
    $script:MailboxLanguageCode = "de-CH"
    $script:MailboxTimeZone = "W. Europe Standard Time"

    $script:UnifiedGroupCreationAllowedGroupName = "M365_GROUP_CREATORS"

    $script:CustomerName = ""

    $script:CreateBreakGlassAccount = $CreateBreakGlassAccount
    $script:EnableSecurityDefaults = $EnableSecurityDefaults
    $script:DisableSecurityDefaults = $DisableSecurityDefaults
    $script:DisableEnterpiseApplicationUserConsent = $DisableEnterpiseApplicationUserConsent
    $script:DisableUsersToCreateAppRegistrations = $DisableUsersToCreateAppRegistrations
    $script:DisableUsersToReadOtherUsers = $DisableUsersToReadOtherUsers
    $script:DisableUsersToCreateSecurityGroups = $DisableUsersToCreateSecurityGroups
    $script:DisableUsersToCreateUnifiedGroups = $DisableUsersToCreateUnifiedGroups
    $script:CreateUnifiedGroupCreationAllowedGroup = $CreateUnifiedGroupCreationAllowedGroup
    $script:EnableBlockMsolPowerShell = $EnableBlockMsolPowerShell

    $script:SetMailboxLanguage = $SetMailboxLanguage
    $script:DisableSharedMailboxLogin = $DisableSharedMailboxLogin
    $script:EnableSharedMailboxCopyToSent = $EnableSharedMailboxCopyToSent
    $script:HideUnifiedMailboxFromOutlookClient = $HideUnifiedMailboxFromOutlookClient

    $script:DisableAddToOneDrive = $DisableAddToOneDrive

    $script:AddExchangeOnlineReport = $AddExchangeOnlineReport
    $script:AddSharePointOnlineReport = $AddSharePointOnlineReport

    <# Interactive inputs section #>
    function CheckInteractiveMode {
        Param(
            $Parameters
        )
        if ($Parameters.Count) {
            Write-Host $script:VersionMessage
            return
        }
        $script:InteractiveMode = $true
    }
    function InteractiveMenu {
        $script:AddExchangeOnlineReport = $true
        $script:AddSharePointOnlineReport = $true
        mainMenu
    }
    function mainMenu {
        $StartOptionValue = 0
        while (($result -ne $StartOptionValue) -or ($result -ne 1)) {
            Clear-Host
            $Status = @"
$($script:VersionMessage)
Main menu:

S: Start
C: Configure options

1: Add SharePoint Online report: $($script:AddSharePointOnlineReport)
2: Add Exchange Online report: $($script:AddExchangeOnlineReport)

"@
            $StartOption = New-Object System.Management.Automation.Host.ChoiceDescription "&START", "Start"
            $ConfigureOption = New-Object System.Management.Automation.Host.ChoiceDescription "&CONFIGURE", "Add SharePoint Online report"
            $AddSharePointOnlineReportOption = New-Object System.Management.Automation.Host.ChoiceDescription "&1 SPO", "Add SharePoint Online report"
            $AddExchangeOnlineReportOption = New-Object System.Management.Automation.Host.ChoiceDescription "&2 EXO", "Add Exchange Online report"
            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($StartOption, $ConfigureOption, $AddSharePointOnlineReportOption, $AddExchangeOnlineReportOption )
            Write-Host $Status
            $result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
            switch ($result) {
                0 { return }
                1 { configMenu }
                2 { $script:AddSharePointOnlineReport = ! $script:AddSharePointOnlineReport }
                3 { $script:AddExchangeOnlineReport = ! $script:AddExchangeOnlineReport }
            }
        }
    }
    function configMenu {
        $StartOptionValue = 0
        Clear-Host
        $Status = @"
$($script:VersionMessage)
Configure menu:

1: Azure Active Directory
2: SharePoint Online
3: Exchange Online

B: Back to main menu

"@
        $BackOption = New-Object System.Management.Automation.Host.ChoiceDescription "&BACK", "Back to main menu"
        $AADOption = New-Object System.Management.Automation.Host.ChoiceDescription "&1 AAD", "Azure Active Directory options"
        $SPOOption = New-Object System.Management.Automation.Host.ChoiceDescription "&2 SPO", "SharePoint Online options"
        $EXOOption = New-Object System.Management.Automation.Host.ChoiceDescription "&3 EXO", "Exchange Online options"
        $Options = [System.Management.Automation.Host.ChoiceDescription[]]($BackOption, $AADOption, $SPOOption, $EXOOption)
        Write-Host $Status
        $result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
        switch ($result) {
            0 { return }
            1 { AADMenu }
            2 { SPOMenu }
            3 { EXOMenu }
        }
    }
    function AADMenu {
        $StartOptionValue = 0
        while ($result -ne $StartOptionValue) {
            Clear-Host
            $Status = @"
$($script:VersionMessage)
Azure Active Directory options:

1: Create BreakGlass account: $($script:CreateBreakGlassAccount)
2: Enable security defaults: $($script:EnableSecurityDefaults)
3: Disable security defaults: $($script:DisableSecurityDefaults)
4: Disable enterprise application user consent: $($script:DisableEnterpiseApplicationUserConsent)
5: Disable user to create app registrations: $($script:DisableUsersToCreateAppRegistrations)
6: Disable user to read other users: $($script:DisableUsersToReadOtherUsers)
7: Disable users to create security groups: $($script:DisableUsersToCreateSecurityGroups)
8: Disable users to create unified groups: $($script:DisableUsersToCreateUnifiedGroups)
9: Create UnifiedGroupCreationAllowed group: $($script:CreateUnifiedGroupCreationAllowedGroup)
0: Disable legacy MsolPowerShell access: $($script:EnableBlockMsolPowerShell)

B: Back to main menu

"@
            $BackOption = New-Object System.Management.Automation.Host.ChoiceDescription "&BACK", "Back to main menu"
            $CreateBreakGlassAccountOption = New-Object System.Management.Automation.Host.ChoiceDescription "&1", "Create BreakGlass account"
            $EnableSecurityDefaultsOption = New-Object System.Management.Automation.Host.ChoiceDescription "&2", "Enable security defaults"
            $DisableSecurityDefaultsOption = New-Object System.Management.Automation.Host.ChoiceDescription "&3", "Disable security defaults"
            $DisableEnterpiseApplicationUserConsentOption = New-Object System.Management.Automation.Host.ChoiceDescription "&4", "Disable enterprise application user consent"
            $DisableUsersToCreateAppRegistrationsOption = New-Object System.Management.Automation.Host.ChoiceDescription "&5", "Disable user to create app registrations"
            $DisableUsersToReadOtherUsersOption = New-Object System.Management.Automation.Host.ChoiceDescription "&6", "Disable user to read other users"
            $DisableUsersToCreateSecurityGroupsOption = New-Object System.Management.Automation.Host.ChoiceDescription "&7", "Disable users to create security groups"
            $DisableUsersToCreateUnifiedGroupsOption = New-Object System.Management.Automation.Host.ChoiceDescription "&8", "Disable users to create unified groups"
            $CreateUnifiedGroupCreationAllowedGroupOption = New-Object System.Management.Automation.Host.ChoiceDescription "&9", "Create UnifiedGroupCreationAllowed group"
            $EnableBlockMsolPowerShellOption = New-Object System.Management.Automation.Host.ChoiceDescription "&0", "Disable legacy MsolPowerShell access"

            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($BackOption, $CreateBreakGlassAccountOption, $EnableSecurityDefaultsOption, $DisableSecurityDefaultsOption, $DisableEnterpiseApplicationUserConsentOption, $DisableUsersToCreateAppRegistrationsOption, $DisableUsersToReadOtherUsersOption, $DisableUsersToCreateSecurityGroupsOption, $DisableUsersToCreateUnifiedGroupsOption, $CreateUnifiedGroupCreationAllowedGroupOption, $EnableBlockMsolPowerShellOption)
            Write-Host $Status
            $result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
            switch ($result) {
                0 { return }
                1 { $script:CreateBreakGlassAccount = ! $script:CreateBreakGlassAccount }
                2 { $script:EnableSecurityDefaults = ! $script:EnableSecurityDefaults }
                3 { $script:DisableSecurityDefaults = ! $script:DisableSecurityDefaults }
                4 { $script:DisableEnterpiseApplicationUserConsent = ! $script:DisableEnterpiseApplicationUserConsent }
                5 { $script:DisableUsersToCreateAppRegistrations = ! $script:DisableUsersToCreateAppRegistrations }
                6 { $script:DisableUsersToReadOtherUsers = ! $script:DisableUsersToReadOtherUsers }
                7 { $script:DisableUsersToCreateSecurityGroups = ! $script:DisableUsersToCreateSecurityGroups }
                8 { $script:DisableUsersToCreateUnifiedGroups = ! $script:DisableUsersToCreateUnifiedGroups }
                9 { $script:CreateUnifiedGroupCreationAllowedGroup = ! $script:CreateUnifiedGroupCreationAllowedGroup }
                10 { $script:EnableBlockMsolPowerShell = ! $script:EnableBlockMsolPowerShell }
            }
        }
    }
    function SPOMenu {
        $StartOptionValue = 0
        while ($result -ne $StartOptionValue) {
            Clear-Host
            $Status = @"
$($script:VersionMessage)
SharePoint Online options:

1: Disable add to OneDrive button: $($script:DisableAddToOneDrive)

B: Back to main menu

"@
            $BackOption = New-Object System.Management.Automation.Host.ChoiceDescription "&BACK", "Back to main menu"
            $DisableAddToOneDriveOption = New-Object System.Management.Automation.Host.ChoiceDescription "&1}", "Disable add to OneDrive button"
            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($BackOption, $DisableAddToOneDriveOption)
            Write-Host $Status
            $result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
            switch ($result) {
                0 { return }
                1 { $script:DisableAddToOneDrive = ! $script:DisableAddToOneDrive }
            }
        }
    }
    function EXOMenu {
        $StartOptionValue = 0
        while ($result -ne $StartOptionValue) {
            Clear-Host
            $Status = @"
$($script:VersionMessage)
Exchange Online options:

1: Set mailbox language: $($script:SetMailboxLanguage)
2: Disable shared mailbox login: $($script:DisableSharedMailboxLogin)
3: Enable shared mailbox copy to sent: $($script:EnableSharedMailboxCopyToSent)
4: Hide unified mailbox from outlook client: $($script:HideUnifiedMailboxFromOutlookClient)

B: Back to main menu
"@
            $BackOption = New-Object System.Management.Automation.Host.ChoiceDescription "&BACK", "Back to main menu"
            $SetMailboxLanguageOption = New-Object System.Management.Automation.Host.ChoiceDescription "&1", "Set mailbox language"
            $DisableSharedMailboxLoginOption = New-Object System.Management.Automation.Host.ChoiceDescription "&2", "Disable shared mailbox login"
            $EnableSharedMailboxCopyToSentOption = New-Object System.Management.Automation.Host.ChoiceDescription "&3", "Enable shared mailbox copy to sent"
            $HideUnifiedMailboxFromOutlookClientOption = New-Object System.Management.Automation.Host.ChoiceDescription "&4", "Hide unified mailbox from outlook client"

            $Options = [System.Management.Automation.Host.ChoiceDescription[]]($BackOption, $SetMailboxLanguageOption, $DisableSharedMailboxLoginOption, $EnableSharedMailboxCopyToSentOption, $HideUnifiedMailboxFromOutlookClientOption)
            Write-Host $Status
            $result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
            switch ($result) {
                0 { return }
                1 { $script:SetMailboxLanguage = ! $script:SetMailboxLanguage }
                2 { $script:DisableSharedMailboxLogin = ! $script:DisableSharedMailboxLogin }
                3 { $script:EnableSharedMailboxCopyToSent = ! $script:EnableSharedMailboxCopyToSent }
                4 { $script:HideUnifiedMailboxFromOutlookClient = ! $script:HideUnifiedMailboxFromOutlookClient }
            }
        }
    }

    <# Connect sessions section #>
    function connectGraph {
        if (checkGraphSession) { disconnectGraph }
        Write-Host "Connecting Graph API PowerShell"
        Connect-MgGraph -Scopes "Policy.Read.All, Policy.ReadWrite.ConditionalAccess, Application.Read.All,
User.Read.All, User.ReadWrite.All, Domain.Read.All, Directory.Read.All, Directory.ReadWrite.All,
RoleManagement.ReadWrite.Directory, DeviceManagementApps.Read.All, DeviceManagementApps.ReadWrite.All,
Policy.ReadWrite.Authorization, Sites.Read.All, AuditLog.Read.All, UserAuthenticationMethod.Read.All, Organization.Read.All" | Out-Null
        if ( -not (checkGraphSession)) { exit }
    }
    function connectSpo {
        if (checkSpoSession) { disconnectSpo }
        Write-Host "Connecting SharePoint Online PowerShell"
        if ($PSVersionTable.PSEdition -eq "Core") { Connect-PnPOnline -Url (getSpoAdminUrl) -Interactive -LaunchBrowser }
        if ($PSVersionTable.PSEdition -eq "Desktop") { Connect-PnPOnline -Url (getSpoAdminUrl) -Interactive }
        if ( -not (checkSpoSession)) { exit }
    }
    function getSpoAdminUrl {
        return ((Invoke-MgGraphRequest -Method GET -Uri https://graph.microsoft.com/v1.0/sites/root).siteCollection.hostname) -replace ".sharepoint.com", "-admin.sharepoint.com"
    }
    function connectExo {
        if (checkExoSession) { disconnectExo }
        Write-Host "Connecting Exchange Online PowerShell"
        Connect-ExchangeOnline -ShowBanner:$false
        if ( -not (checkExoSession)) { exit }
    }

    <# Check session section#>
    function checkGraphSession {
        if (Get-MgContext) {
            Write-Host "Connected to Graph API PowerShell using $((Get-MgContext).Account) account"
            return $true
        }
        Write-Host "Not connected to Graph API PowerShell"
        return $false
    }
    function checkSpoSession {
        if (Get-PnPConnection) {
            Write-Host "Connected to SharePoint Online PowerShell tenant $((Get-PnPConnection).Url)"
            return $true
        }
        Write-Host "Not connected to SharePoint Online PowerShell"
        return $false
    }
    function checkExoSession {
        if ((Get-ConnectionInformation).State -eq "Connected") {
            Write-Host "Connected to Exchange Online PowerShell using $((Get-ConnectionInformation).UserPrincipalName) account"
            return $true
        }
        Write-Host "Not connected to Exchange Online PowerShell"
        return $false
    }
    
    <# Disconect session section #>
    function disconnectGraph {
        Write-Host "Disconnecting existing Graph API PowerShell"
        Disconnect-Graph | Out-Null
    }
    function disconnectSpo {
        Write-Host "Disconnecting existing SharePoint Online PowerShell"
        Disconnect-PnPOnline
    }
    function disconnectExo {
        Write-Host "Disconnecting existing Exchange Online PowerShell"
        Disconnect-ExchangeOnline -Confirm:$false
    }

    <# Customer infos#>
    function organizationReport {
        $Organization = Get-MgOrganization -Property DisplayName, Id
        $script:CustomerName = $Organization.DisplayName
        return  "<h2>$($Organization.DisplayName) ($($Organization.Id))</h2>"
    }

    <# User settings policy section #>
    function checkTenanUserSettingsReport {
        param(
            [System.Boolean]$DisableUserConsent,
            [System.Boolean]$DisableUsersToCreateAppRegistrations,
            [System.Boolean]$DisableUsersToReadOtherUsers,
            [System.Boolean]$DisableUsersToCreateSecurityGroups,
            [System.Boolean]$DisableUsersToCreateUnifiedGroups,
            [System.Boolean]$CreateUnifiedGroupCreationAllowedGroup,
            [System.Boolean]$EnableBlockMsolPowerShell
        )
        if ($DisableUserConsent) { disableApplicationUserConsent }
        if ($DisableUsersToCreateAppRegistrations) { disableUsersToCreateAppRegistrations }
        if ($DisableUsersToReadOtherUsers) { disableUsersToReadOtherUsers }
        if ($DisableUsersToCreateSecurityGroups) { disableUsersToCreateSecurityGroups }
        if ($DisableUsersToCreateUnifiedGroups) { disableUsersToCreateUnifiedGroups }
        if ($CreateUnifiedGroupCreationAllowedGroup) { createUnifiedGroupCreationAllowedGroup }
        if ($EnableBlockMsolPowerShell) { enableBlockMsolPowerShell }
        Write-Host "Checking user settings"
        $Policy = Get-MgPolicyAuthorizationPolicy -Property BlockMsolPowerShell, DefaultUserRolePermissions
        $Report = $Policy | Select-Object -Property @{Name = "PermissionGrantPoliciesAssigned"; Expression = { [string]$_.DefaultUserRolePermissions.PermissionGrantPoliciesAssigned } },
        @{Name = "AllowedToCreateApps"; Expression = { [string]$_.DefaultUserRolePermissions.AllowedToCreateApps } },
        @{Name = "AllowedToCreateSecurityGroups"; Expression = { [string]$_.DefaultUserRolePermissions.AllowedToCreateSecurityGroups } },
        @{Name = "AllowedToCreateUnifiedGroups"; Expression = { checkAllowedToCreateUnifiedGroups } },
        @{Name = "AllowedToCreateUnifiedGroupsGroupName"; Expression = { checkUnifiedGroupCreationAllowedGroup } },
        @{Name = "AllowedToReadOtherUsers"; Expression = { [string]$_.DefaultUserRolePermissions.AllowedToReadOtherUsers } },
        @{Name = "AllowedToCreateTenants"; Expression = { checkAllowedToCreateTenants } },
        BlockMsolPowerShell | ConvertTo-Html -As List -Fragment -PreContent "<h3 id='AAD_USER_SETTINGS'>User settings</h3>" -PostContent "<p>PermissionGrantPoliciesAssigned: empty (user consent not allowed), microsoft-user-default-legacy (user consent allowed for all apps), microsoft-user-default-low (user consent allowed for low permission apps)</p><p>Unified groups = Microsoft 365 groups</p>"

        $Report = $Report -Replace "<td>PermissionGrantPoliciesAssigned:</td><td>ManagePermissionGrantsForSelf.microsoft-user-default-legacy</td>", "<td>PermissionGrantPoliciesAssigned:</td><td class='red'>microsoft-user-default-legacy</td>"
        $Report = $Report -Replace "<td>PermissionGrantPoliciesAssigned:</td><td>ManagePermissionGrantsForSelf.microsoft-user-default-low</td>", "<td>PermissionGrantPoliciesAssigned:</td><td class='orange'>microsoft-user-default-low</td>"
        $Report = $Report -Replace "<td>AllowedToCreateApps:</td><td>True</td>", "<td>AllowedToCreateApps:</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>AllowedToCreateSecurityGroups:</td><td>True</td>", "<td>AllowedToCreateSecurityGroups:</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>AllowedToCreateUnifiedGroups:</td><td>True</td>", "<td>AllowedToCreateUnifiedGroups:</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>AllowedToCreateUnifiedGroups:</td><td>false</td>", "<td>AllowedToCreateUnifiedGroups:</td><td>False</td>"
        $Report = $Report -Replace "<td>AllowedToReadOtherUsers:</td><td>True</td>", "<td>AllowedToReadOtherUsers:</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>AllowedToCreateTenants:</td><td>True</td>", "<td>AllowedToCreateTenants:</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>BlockMsolPowerShell:</td><td>False</td>", "<td>BlockMsolPowerShell:</td><td class='red'>False</td>"
        return $Report
    }
    function disableApplicationUserConsent {
        Write-Host "Disable enterprise application user consent"
        Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "PermissionGrantPoliciesAssigned" = @() }
    }
    function disableUsersToCreateAppRegistrations {
        Write-Host "Disable users to create app registrations"
        Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "AllowedToCreateApps" = $false }
    }
    function disableUsersToReadOtherUsers {
        Write-Host "Disable users to read other users"
        Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "AllowedToReadOtherUsers" = $false }
    }
    function disableUsersToCreateSecurityGroups {
        Write-Host "Disable users to create security groups"
        Update-MgPolicyAuthorizationPolicy -DefaultUserRolePermissions @{ "AllowedToCreateSecurityGroups" = $false }
    }
    function enableBlockMsolPowerShell {
        Write-Host "Disable legacy MsolPowerShell access"
        Update-MgPolicyAuthorizationPolicy -BlockMsolPowerShell
    }
    function checkAllowedToCreateUnifiedGroups {
        if ($GroupSettingsUnified = getGroupSettingsUnified) {
            return ($GroupSettingsUnified.values | Where-Object name -EQ "EnableGroupCreation").value
        }
        return $true
    }
    function checkUnifiedGroupCreationAllowedGroup {
        $GroupSettingsUnified = getGroupSettingsUnified
        if ($GroupId = ($GroupSettingsUnified.values | Where-Object name -EQ "GroupCreationAllowedGroupId").value) {
            return (Get-MgGroup -GroupId $GroupId -Property DisplayName).DisplayName
        }
    }
    function createUnifiedGroupCreationAllowedGroup {
        Write-Host "Creating UnifiedGroupCreationAllowed group:" $script:UnifiedGroupCreationAllowedGroupName
        if (checkUnifiedGroupCreationAllowedGroup) {
            Write-Host "UnifiedGroupCreationAllowed group already assigned"
            return
        }
        if ($GroupId = (Get-MgGroup -Property Id, DisplayName -Filter "DisplayName eq '$($script:UnifiedGroupCreationAllowedGroupName)'").Id) { Write-Host "UnifiedGroupCreationAllowed group already exists" } else { $GroupId = (New-MgGroup -DisplayName $script:UnifiedGroupCreationAllowedGroupName -MailEnabled:$False -MailNickname $script:UnifiedGroupCreationAllowedGroupName -SecurityEnabled).Id }
        $Body = @{
            templateId = (getGroupSettingsTemplateUnified).id
            values     = @( @{ Name = "GroupCreationAllowedGroupId" ; Value = $GroupId } )
        }
        if ($GroupSettingsUnified = getGroupSettingsUnified) { Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/groupSettings/$($GroupSettingsUnified.id)" -Body $Body }
        else { Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groupSettings" -Body $Body }
    }
    function getGroupSettingsTemplateUnified {
        $GroupSettingTemplates = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groupSettingTemplates?$select=value"
        return $GroupSettingTemplates.value | Where-Object { $_.displayName -eq "Group.Unified" } | Select-Object -Property id, DisplayName
    }
    function getGroupSettingsUnified {
        $GroupSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groupSettings?$select=value" 
        return $GroupSettings.value | Where-Object { $_.templateId -eq (getGroupSettingsTemplateUnified).id } | Select-Object -Property id, templateId, values
    }
    function disableUsersToCreateUnifiedGroups {
        $Body = @{
            templateId = (getGroupSettingsTemplateUnified).id
            values     = @( @{ Name = "EnableGroupCreation" ; Value = "false" } )
        }
        if ($GroupSettingsUnified = getGroupSettingsUnified) { Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/groupSettings/$($GroupSettingsUnified.id)" -Body $Body }
        else { Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groupSettings" -Body $Body }
    }
    function checkAllowedToCreateTenants {
        $DefaultUserRolePermissions = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/authorizationPolicy/authorizationPolicy").defaultUserRolePermissions
        return $DefaultUserRolePermissions["allowedToCreateTenants"]
    }

    <# Device join settings#>
    function checkDeviceJoinSettingsReport {
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
    function checkUsedSKUReport {
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
    function disableUserAccount {
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
    function checkUserAccountStatus {
        param(
            $UserId
        )
        return (Get-MgUser -UserId $UserId -Property AccountEnabled).AccountEnabled
    }

    <# Admin role section #>
    function checkAdminRoleReport {
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
    function checkBreakGlassAccountReport {
        param (
            $Create
        )
        if ($BgAccount = getBreakGlassAccount) {
            $Report = $BgAccount | ConvertTo-Html -Property DisplayName, UserPrincipalName, AccountEnabled, GlobalAdmin, LastSignIn -As Table -Fragment -PreContent "<br><h3 id='AAD_BG'>BreakGlass account</h3>"
            $Report = $Report -Replace "<td>False</td>", "<td class='red'>False</td>"
            return $Report

        }
        if ($create) {
            createBreakGlassAccount
            $Report = getBreakGlassAccount | ConvertTo-Html -Property DisplayName, UserPrincipalName, AccountEnabled, GlobalAdmin, LastSignIn -As Table -Fragment -PreContent "<br><h3 id='AAD_BG'>BreakGlass account</h3>" -PostContent "<p>Check console log for credentials</p>"
            $Report = $Report -Replace "<td>False</td>", "<td class='red'>False</td>"
            return $Report
        }
        return "<br><h3 id='AAD_BG'>BreakGlass account</h3><p>Not found</p>"
    }
    function getBreakGlassAccount {
        Write-Host "Checking BreakGlass account"
        Select-MgProfile -Name "beta"
        $BgAccounts = Get-MgUser -Filter "startswith(displayName, 'BreakGlass ')" -Property Id, DisplayName, UserPrincipalName, AccountEnabled, SignInActivity
        Select-MgProfile -Name "v1.0"
        if (-not $bgAccounts) { 
            $BgAccounts = Get-MgUser -Filter "startswith(displayName, 'BreakGlass ')" -Property Id, DisplayName, UserPrincipalName, AccountEnabled
        }
        if (-not $bgAccounts) { return }
        foreach ($BgAccount in $BgAccounts) {
            Add-Member -InputObject $BgAccount -NotePropertyName "GlobalAdmin" -NotePropertyValue (checkGlobalAdminRole $BgAccount.Id)
            Add-Member -InputObject $BgAccount -NotePropertyName "LastSignIn" -NotePropertyValue $BgAccount.SignInActivity.LastSignInDateTime
        }
        return $BgAccounts
    }
    function getGlobalAdminRoleId {
        return (Get-MgDirectoryRole -Filter "DisplayName eq 'Global Administrator'" -Property Id).Id
    }
    function checkGlobalAdminRole {
        param (
            $AccountId
        )
        if (Get-MgDirectoryRoleMember -DirectoryRoleId (getGlobalAdminRoleId) -Filter "id eq '$($AccountId)'") {
            return $true
        }
    }
    function createBreakGlassAccount {
        Write-Host "Creating BreakGlass account:"
        $Name = -join ((97..122) | Get-Random -Count 64 | ForEach-Object { [char]$_ })
        $DisplayName = "BreakGlass $Name"
        $Domain = (Get-MgDomain -Property id, IsInitial | Where-Object { $_.IsInitial -eq $true }).Id
        $UPN = "$Name@$Domain"
        $PasswordProfile = @{
            ForceChangePasswordNextSignIn        = $false
            ForceChangePasswordNextSignInWithMfa = $false
            Password                             = generatePassword
        }
        $BgAccount = New-MgUser -DisplayName $DisplayName -UserPrincipalName $UPN -MailNickname $Name -PasswordProfile $PasswordProfile -PreferredLanguage "en-US" -AccountEnabled
        $DirObject = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($BgAccount.id)"
        }
        New-MgDirectoryRoleMemberByRef -DirectoryRoleId (getGlobalAdminRoleId) -BodyParameter $DirObject
        Add-Member -InputObject $BgAccount -NotePropertyName "Password" -NotePropertyValue $PasswordProfile.Password
        Write-Host ($BgAccount | Select-Object -Property Id, DisplayName, UserPrincipalName, Password | Format-List | Out-String)
    }
    function generatePassword {
        param (
            [ValidateRange(4, [int]::MaxValue)]
            [int] $length = 64,
            [int] $upper = 4,
            [int] $lower = 4,
            [int] $numeric = 4,
            [int] $special = 4
        )
        if ($upper + $lower + $numeric + $special -gt $length) {
            throw "number of upper/lower/numeric/special char must be lower or equal to length"
        }
        $uCharSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        $lCharSet = "abcdefghijklmnopqrstuvwxyz"
        $nCharSet = "0123456789"
        $sCharSet = "/*-+, !?=()@; :._"
        $charSet = ""
        if ($upper -gt 0) { $charSet += $uCharSet }
        if ($lower -gt 0) { $charSet += $lCharSet }
        if ($numeric -gt 0) { $charSet += $nCharSet }
        if ($special -gt 0) { $charSet += $sCharSet }
        $charSet = $charSet.ToCharArray()
        $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
        $bytes = New-Object byte[]($length)
        $rng.GetBytes($bytes)
        $result = New-Object char[]($length)
        for ($i = 0 ; $i -lt $length ; $i++) {
            $result[$i] = $charSet[$bytes[$i] % $charSet.Length]
        }
        $password = (-join $result)
        $valid = $true
        if ($upper -gt ($password.ToCharArray() | Where-Object { $_ -cin $uCharSet.ToCharArray() }).Count) { $valid = $false }
        if ($lower -gt ($password.ToCharArray() | Where-Object { $_ -cin $lCharSet.ToCharArray() }).Count) { $valid = $false }
        if ($numeric -gt ($password.ToCharArray() | Where-Object { $_ -cin $nCharSet.ToCharArray() }).Count) { $valid = $false }
        if ($special -gt ($password.ToCharArray() | Where-Object { $_ -cin $sCharSet.ToCharArray() }).Count) { $valid = $false }
        if (!$valid) {
            $password = Get-RandomPassword $length $upper $lower $numeric $special
        }
        return $password
    }

    <# User MFA section#>
    function checkUserMfaStatusReport {
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
                Switch ($MFA.AdditionalProperties["@odata.type"]) { 
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
                    }
                }
                $AuthenticationMethod += $AuthMethod
                if ($null -ne $AuthMethodDetails) {
                    $AdditionalDetails += "$AuthMethod : $AuthMethodDetails"
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
    function checkGuestUserReport {
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
    function checkSecurityDefaultsReport {
        param (
            [System.Boolean]$EnableSecurityDefaults,
            [System.Boolean]$DisableSecurityDefaults
        )
        if ($EnableSecurityDefaults -and (-not $DisableSecurityDefaults)) {
            updateSecurityDefaults -Enable $true
        }  
        if ($DisableSecurityDefaults -and (-not $EnableSecurityDefaults)) {
            updateSecurityDefaults -Enable $false
        }
        if (checkSecurityDefaults) {
            return "<br><h3 id='AAD_SEC_DEFAULTS'>Security defaults</h3><p>Enabled</p>"
        }
        return "<br><h3 id='AAD_SEC_DEFAULTS'>Security defaults</h3><p>Disabled</p>"
    }
    function checkSecurityDefaults {
        Write-Host "Checking security defaults"
        return (Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -Property "isEnabled").IsEnabled
    }
    function updateSecurityDefaults {
        param ([System.Boolean]$Enable)
        $params = @{
            IsEnabled = $Enable
        }
        Write-Host "Updating security defaults enable:" $Enable
        Update-MgPolicyIdentitySecurityDefaultEnforcementPolicy -BodyParameter $params
    }

    <# Conditional access section #>
    function checkConditionalAccessPolicyReport {
        Write-Host "Checking conditional access policies"
        if ($Policy = Get-MgIdentityConditionalAccessPolicy -Property Id, DisplayName, State) {
            return $Policy | ConvertTo-Html -Property DisplayName, Id, State -As Table -Fragment -PreContent "<br><h3 id='AAD_CA'>Conditional access policies</h3>"
        }
        return "<br><h3 id='AAD_CA'>Conditional access policies</h3><p>Not found</p>"
    }
    function checkNamedLocationReport {
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
    function checkAppProtectionPolicesReport {
        Write-Host "Checking app protection policies"
        if ($Polices = getAppProtectionPolices) {
            return $Polices | ConvertTo-Html -As Table -Property DisplayName, IsAssigned -Fragment -PreContent "<br><h3 id='AAD_APP_POLICY'>App protection policies</h3>"
        }
        return "<br><h3 id='AAD_APP_POLICY'>App protection policies</h3><p>Not found</p>"
    }
    function getAppProtectionPolices {
        $IOSPolicies = Get-MgDeviceAppManagementiOSManagedAppProtection -Property DisplayName, IsAssigned
        $AndroidPolicies = Get-MgDeviceAppManagementAndroidManagedAppProtection -Property DisplayName, IsAssigned
        $Policies = @()
        $Policies += $IOSPolicies
        $Policies += $AndroidPolicies
        return $Policies
    }

    <# SharePoint Tenant section #>
    function checkSpoTenantReport {
        param(
            [System.Boolean]$DisableAddToOneDrive
        )
        Write-Host "Checking tenant settings"
        if ($DisableAddToOneDrive) {
            Write-Host "Disable add to OneDrive button"
            Set-PnPTenant -DisableAddToOneDrive $True
        }
        $Report = Get-PnPTenant | ConvertTo-Html -As List -Property LegacyAuthProtocolsEnabled, DisableAddToOneDrive, ConditionalAccessPolicy, SharingCapability, ODBMembersCanShare, PreventExternalUsersFromResharing, DefaultSharingLinkType, DefaultLinkPermission, FolderAnonymousLinkType, FileAnonymousLinkType, RequireAnonymousLinksExpireInDays -Fragment -PreContent "<h3 id='SPO_SETTINGS'>Tenant settings</h3>" -PostContent "<p>ConditionalAccessPolicy: AllowFullAccess, AllowLimitedAccess, BlockAccess</p>
    <p>SharingCapability: Disabled, ExternalUserSharingOnly, ExternalUserAndGuestSharing, ExistingExternalUserSharingOnly</p>
    <p>DefaultSharingLinkType: None, Direct, Internal, AnonymousAccess</p>"
        $Report = $Report -Replace "<td>LegacyAuthProtocolsEnabled:</td><td>True</td>", "<td>LegacyAuthProtocolsEnabled:</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>DisableAddToOneDrive:</td><td>False</td>", "<td>DisableAddToOneDrive:</td><td class='red'>False</td>"
        $Report = $Report -Replace "<td>ConditionalAccessPolicy:</td><td>AllowFullAccess</td>", "<td>ConditionalAccessPolicy:</td><td class='red'>AllowFullAccess</td>"
        $Report = $Report -Replace "<td>SharingCapability:</td><td>ExternalUserAndGuestSharing</td>", "<td>SharingCapability:</td><td class='red'>ExternalUserAndGuestSharing</td>"
        $Report = $Report -Replace "<td>PreventExternalUsersFromResharing:</td><td>False</td>", "<td>PreventExternalUsersFromResharing:</td><td class='red'>False</td>"
        $Report = $Report -Replace "<td>DefaultSharingLinkType:</td><td>AnonymousAccess</td>", "<td>DefaultSharingLinkType:</td><td class='red'>AnonymousAccess</td>"
        return $Report
    }

    <# Mail Domain section #>
    function checkMailDomainReport {
        Write-Host "Checking domains"
        $Domains = Get-DkimSigningConfig | Select-Object -Property Id, @{Name = "Default"; Expression = { $_.IsDefault } }, @{Name = "DKIM"; Expression = { $_.Enabled } }
        if (-not ($Domains)) { $Domains = Get-AcceptedDomain | Select-Object -Property Id, "Default", @{Name = "DKIM"; Expression = { $false } } }
        $DomainsReport = @()
        foreach ($Domain in $Domains) {
            $ProcessedCount++
            Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Domain.Id)"
            $Domain = checkDMARC -Domain $Domain
            $Domain = checkSPF -Domain $Domain
            $DomainsReport += $Domain
        }
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Domain.Id)" -Status "Ready" -Completed
        $Report = $DomainsReport | ConvertTo-Html -As Table -Property Id, DKIM, DMARC, SPF, "DMARC record", "SPF record", "DMARC hint", "SPF hint", "Default" -Fragment -PreContent "<h3 id='EXO_DOMAIN'>Domains</h3>"
        $Report = $Report -Replace "<td>False</td><td>False</td><td>False</td>", "<td class='red'>False</td><td class='red'>False</td><td class='red'>False</td>"
        $Report = $Report -Replace "<td>False</td><td>False</td><td>True</td>", "<td class='red'>False</td><td class='red'>False</td><td>True</td>"
        $Report = $Report -Replace "<td>True</td><td>False</td><td>False</td>", "<td>True</td><td class='red'>False</td><td class='red'>False</td>"
        $Report = $Report -Replace "<td>True</td><td>False</td><td>True</td>", "<td>True</td><td class='red'>False</td><td>True</td>"
        $Report = $Report -Replace "<td>False</td><td>True</td><td>False</td>", "<td class='red'>False</td><td>True</td><td class='red'>False</td>"
        $Report = $Report -Replace "<td>False</td><td>True</td><td>True</td>", "<td class='red'>False</td><td>True</td><td>True</td>"
        $Report = $Report -Replace "<td>Should be p=reject</td>", "<td class='orange'>Should be p=reject</td>"
        $Report = $Report -Replace "<td>Not sufficiently stricth</td>", "<td class='orange'>Not sufficiently strict</td>"
        $Report = $Report -Replace "<td>Not effective enough</td>", "<td class='red'>Not effective enough</td>"
        $Report = $Report -Replace "<td>Does not protect</td>", "<td class='red'>Does not protect</td>"
        $Report = $Report -Replace "<td>No qualifier found</td>", "<td class='red'>No qualifier found</td>"
        return $Report
    }
    function checkDMARC {
        param($Domain)
        if ($PSVersionTable.Platform -eq "Unix") { $DMARCRecord = (Resolve-Dns -Query "_dmarc.$($Domain.Id)" -QueryType TXT | Select-Object -Expand Answers).Text }
        else { $DMARCRecord = Resolve-DnsName -Name "_dmarc.$($Domain.Id)" -Type TXT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty strings }
        if ($null -eq $DMARCRecord ) {
            $DMARC = $false
        }
        else {
            switch -Regex ($DMARCRecord ) {
                ('p=none') {
                    $DmarcHint = "Does not protect"
                    $DMARC = $true
                }
                ('p=quarantine') {
                    $DmarcHint = "Should be p=reject"
                    $DMARC = $true
                }
                ('p=reject') {
                    $DmarcHint = "Will protect"
                    $DMARC = $true
                }
                ('sp=none') {
                    $DmarcHint += "Does not protect"
                    $DMARC = $true
                }
                ('sp=quarantine') {
                    $DmarcHint += "Should be p=reject"
                    $DMARC = $true
                }
                ('sp=reject') {
                    $DmarcHint += "Will protect"
                    $DMARC = $true
                }
            }
        }
        $Domain | Add-Member NoteProperty "DMARC" $DMARC
        $Domain | Add-Member NoteProperty "DMARC record" "$($DMARCRecord )"
        $Domain | Add-Member NoteProperty "DMARC hint" $DmarcHint
        return $Domain
    }
    function checkSPF {
        param($Domain)
        if ($PSVersionTable.Platform -eq "Unix") { $SPFRecord = (Resolve-Dns -Query $Domain.Id -QueryType TXT | Select-Object -Expand Answers).Text | Where-Object { $_ -match "v=spf1" } }
        else { $SPFRecord = Resolve-DnsName -Name $Domain.Id -Type TXT -ErrorAction SilentlyContinue | Where-Object { $_.strings -match "v=spf1" } | Select-Object -ExpandProperty strings }
        if ($SPFRecord -match "redirect") {
            $redirect = $SPFRecord.Split(" ")
            $RedirectName = $redirect -match "redirect" -replace "redirect="
            if ($PSVersionTable.Platform -eq "Unix") { $SPFRecord = (Resolve-Dns -Query $RedirectName -QueryType TXT | Select-Object -Expand Answers).Text | Where-Object { $_ -match "v=spf1" } }
            else { $SPFRecord = Resolve-DnsName -Name $RedirectName -Type TXT -ErrorAction SilentlyContinue | Where-Object { $_.strings -match "v=spf1" } | Select-Object -ExpandProperty strings }
        }
        if ($null -eq $SPFRecord) {
            $SPF = $false
        }
        if ($SPFRecord -is [array]) {
            $SPFHint = "More than one SPF-record"
            $SPF = $true
        }
        Else {
            switch -Regex ($SPFRecord) {
                '~all' {
                    $SPFHint = "Not sufficiently strict"
                    $SPF = $true
                }
                '-all' {
                    $SPFHint = "Sufficiently strict"
                    $SPF = $true
                }
                "\?all" {
                    $SPFHint = "Not effective enough"
                    $SPF = $true
                }
                '\+all' {
                    $SPFHint = "Not effective enough"
                    $SPF = $true
                }
                Default {
                    $SPFHint = "No qualifier found"
                    $SPF = $true
                }
            }
        }
        $Domain | Add-Member NoteProperty "SPF" "$($SPF)"
        $Domain | Add-Member NoteProperty "SPF record" "$($SPFRecord)"
        $Domain | Add-Member NoteProperty "SPF hint" $SPFHint
        return $Domain
    }

    <# Mail connector section#>
    function checkMailConnectorReport {
        Write-Host "Checking mail connectors"
        if (-not ($Inbound = Get-InboundConnector)) { $InboundReport = "<br><h3 id='EXO_CONNECTOR_IN'>Inbound mail connector</h3><p>Not found</p>" }
        else { $InboundReport = $Inbound | ConvertTo-Html -As Table -Property Name, SenderDomains, SenderIPAddresses, Enabled -Fragment -PreContent "<br><h3 id='EXO_CONNECTOR_IN'>Inbound mail connector</h3>" }
        if (-not ($Outbound = Get-OutboundConnector -IncludeTestModeConnectors:$true)) { $OutboundReport = "<br><h3 id='EXO_CONNECTOR_OUT'>Outbound mail connector</h3><p>Not found</p>" }
        else { $OutboundReport = $Outbound | ConvertTo-Html -As Table -Property Name, RecipientDomains, SmartHosts, Enabled -Fragment -PreContent "<br><h3 id='EXO_CONNECTOR_OUT'>Outbound mail connector</h3>" }
        $Report = @()
        $Report += $InboundReport
        $Report += $OutboundReport
        return $Report
    }

    <# User mailbox section #>
    function checkMailboxReport {
        param(
            [System.Boolean]$Language
        )
        Write-Host "Checking user mailboxes"
        if ( -not ($Mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize:Unlimited -Properties DisplayName, UserPrincipalName)) {
            return "<br><h3 id='EXO_USER'>User mailbox</h3><p>Not found</p>"
        }
        if ($Language) {
            setMailboxLang -Mailbox $Mailboxes
        }
        $MailboxReport = @()
        foreach ($Mailbox in $Mailboxes) {
            $ProcessedCount++
            Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)"
            $MailboxReport += checkMailboxLoginAndLocation $Mailbox
        }
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)" -Status "Ready" -Completed
        return $MailboxReport | ConvertTo-Html -As Table -Property UserPrincipalName, DisplayName, Language, TimeZone, LoginAllowed `
            -Fragment -PreContent "<br><h3 id='EXO_USER'>User mailbox</h3>"
    }
    function setMailboxLang {
        param(
            $Mailbox
        )
        Write-Host "Setting mailboxes language:" $script:MailboxLanguageCode "timezone:" $script:MailboxTimeZone
        $Mailbox | Set-MailboxRegionalConfiguration -LocalizeDefaultFolderName:$true -Language $script:MailboxLanguageCode -TimeZone $script:MailboxTimeZone
    }

    <# Shared mailbox section #>
    function checkSharedMailboxReport {
        param(
            [System.Boolean]$Language,
            [System.Boolean]$DisableLogin,
            [System.Boolean]$EnableCopy
        )
        Write-Host "Checking shared mailboxes"
        if ( -not ($Mailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited -Properties DisplayName,
                UserPrincipalName, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled)) {
            return "<br><h3 id='EXO_SHARED'>Shared mailbox</h3><p>Not found</p>"
        }
        if ($Language) { setMailboxLang -Mailbox $Mailboxes }
        if ($DisableLogin) { disableUserAccount $Mailboxes }
        if ($EnableCopy) {
            setSharedMailboxEnableCopyToSent $Mailboxes
            $Mailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited -Properties DisplayName,
            UserPrincipalName, MessageCopyForSentAsEnabled, MessageCopyForSendOnBehalfEnabled
        }
        $MailboxReport = @()
        foreach ($Mailbox in $Mailboxes) {
            $ProcessedCount++
            Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)"
            $MailboxReport += checkMailboxLoginAndLocation $Mailbox
        }
        Write-Progress -Activity "Processed count: $ProcessedCount; Currently processing: $($Mailbox.DisplayName)" -Status "Ready" -Completed
        $Report = $MailboxReport | ConvertTo-Html -As Table -Property UserPrincipalName, DisplayName, Language, TimeZone, MessageCopyForSentAsEnabled,
        MessageCopyForSendOnBehalfEnabled, LoginAllowed -Fragment -PreContent "<br><h3 id='EXO_SHARED'>Shared mailbox</h3>"
        $Report = $Report -Replace "<td>True</td><td>True</td><td>True</td>", "<td>True</td><td>True</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>False</td><td>False</td><td>True</td>", "<td>False</td><td>False</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>True</td><td>False</td><td>True</td>", "<td>True</td><td>False</td><td class='red'>True</td>"
        $Report = $Report -Replace "<td>False</td><td>True</td><td>True</td>", "<td>False</td><td>True</td><td class='red'>True</td>"
        return $Report
    }
    function checkMailboxLoginAndLocation {
        param (
            $Mailbox
        )
        $ReginalConfig = $Mailbox | Get-MailboxRegionalConfiguration
        Add-Member -InputObject $Mailbox -NotePropertyName "Language" -NotePropertyValue $ReginalConfig.Language
        Add-Member -InputObject $Mailbox -NotePropertyName "TimeZone" -NotePropertyValue $ReginalConfig.TimeZone
        Add-Member -InputObject $Mailbox -NotePropertyName "LoginAllowed" -NotePropertyValue (checkUserAccountStatus $Mailbox.UserPrincipalName)
        return $Mailbox
    }
    function setSharedMailboxEnableCopyToSent {
        param(
            $Mailbox
        )
        Write-Host "Enable shared mailbox copy to sent"
        $Mailbox | Set-Mailbox -MessageCopyForSentAsEnabled $True -MessageCopyForSendOnBehalfEnabled $True
    }

    <# Unified mailbox section #>
    function checkUnifiedMailboxReport {
        param(
            [System.Boolean]$HideFromClient
        )
        Write-Host "Checking unified mailboxes"
        if ( -not ($Mailboxes = Get-UnifiedGroup -ResultSize Unlimited)) {
            return "<br><h3 id='EXO_UNIFIED'>Unified mailbox</h3><p>Not found</p>"
        }
        if ($HideFromClient) {
            Write-Host "Hiding unified mailboxes from outlook client"
            $Mailboxes | Set-UnifiedGroup -HiddenFromExchangeClientsEnabled:$true -HiddenFromAddressListsEnabled:$false
            $Mailboxes = Get-UnifiedGroup -ResultSize Unlimited 
        }
        return $Mailboxes | Sort-Object -Property PrimarySmtpAddress | ConvertTo-Html -As Table -Property DisplayName, PrimarySmtpAddress, HiddenFromAddressListsEnabled, HiddenFromExchangeClientsEnabled -Fragment -PreContent "<br><h3 id='EXO_UNIFIED'>Unified mailbox</h3>" -PostContent "<p>Unified groups = Microsoft 365 groups</p>"
    }

    <# HTML table of content section #>
    $TableOfContents = @()
    $TableOfContents += "<br><hr><h2>Contents</h2>"
    $AADTableOfContents = @"
<h3 class='TOC'><a href="#AAD">Azure Active Directory</a></h3>
<ul>
    <li><a href="#AAD_USER_SETTINGS">User settings</a></li>
    <li><a href="#AAD_DEVICE_JOIN_SETTINGS">Device join settings</a></li>
    <li><a href="#AAD_SKU">Licenses</a></li>
    <li><a href="#AAD_ADMINS">Admin role assignments</a></li>
    <li><a href="#AAD_BG">BreakGlass account</a></li>
    <li><a href="#AAD_MFA">User MFA status</a></li>
    <li><a href="#AAD_GUEST">Guest accounts</a></li>
    <li><a href="#AAD_SEC_DEFAULTS">Security defaults</a></li>
    <li><a href="#AAD_CA">Conditional access policies</a></li>
    <li><a href="#AAD_CA_LOCATIONS">Named locations</a></li>
    <li><a href="#AAD_APP_POLICY">App protection policies</a></li>
</ul>
"@
    $SPOTableOfContents = @"
<h3 class='TOC'><a href="#SPO">SharePoint Online</a></h3>
<ul>
    <li><a href="#SPO_SETTINGS">Tenant settings</a></li>
</ul>
"@
    $EXOTabelOfContents = @"
<h3 class='TOC'><a href="#EXO">Exchange Online</a></h3>
<ul>
    <li><a href="#EXO_DOMAIN">Domains</a></li>
    <li><a href="#EXO_CONNECTOR_IN">Inbound mail connector</a></li>
    <li><a href="#EXO_CONNECTOR_OUT">Outbound mail connector</a></li>
    <li><a href="#EXO_USER">User mailbox</a></li>
    <li><a href="#EXO_SHARED">Shared mailbox</a></li>
    <li><a href="#EXO_UNIFIED">Unified mailbox</a></li>
</ul>
"@

    <# Script logic start section #>
    CheckInteractiveMode -Parameters $PSBoundParameters
    if ($script:InteractiveMode) {
        InteractiveMenu
    }

    if (-not $UseExistingGraphSession) { connectGraph }
    else { if ( -not (checkGraphSession)) { exit } }
    $TableOfContents += $AADTableOfContents

    if ($script:AddSharePointOnlineReport -or $script:DisableAddToOneDrive) { 
        if (-not $UseExistingSpoSession) { connectSpo }
        else { if ( -not (checkSpoSession)) { exit } }
        $TableOfContents += $SPOTableOfContents
    }

    if ($script:AddExchangeOnlineReport -or $script:SetMailboxLanguage -or $script:DisableSharedMailboxLogin -or $script:EnableSharedMailboxCopyToSent -or $script:HideUnifiedMailboxFromOutlookClient) {
        if (-not $UseExistingExoSession) { connectExo }
        else { if ( -not (checkExoSession)) { exit } }
        $TableOfContents += $EXOTabelOfContents
    }

    $Report = @()
    $Report += organizationReport
    $Report += $TableOfContents
    $Report += "<br><hr><h2 id='AAD'>Azure Active Directory</h2>"
    Write-Host "Azure Active Directory"
    $Report += checkTenanUserSettingsReport -DisableUserConsent $script:DisableEnterpiseApplicationUserConsent -DisableUsersToCreateAppRegistrations $script:DisableUsersToCreateAppRegistrations -DisableUsersToReadOtherUsers $script:DisableUsersToReadOtherUsers -DisableUsersToCreateSecurityGroups $script:DisableUsersToCreateSecurityGroups -DisableUsersToCreateUnifiedGroups $script:DisableUsersToCreateUnifiedGroups -CreateUnifiedGroupCreationAllowedGroup $script:CreateUnifiedGroupCreationAllowedGroup -EnableBlockMsolPowerShell $script:EnableBlockMsolPowerShell
    $Report += checkDeviceJoinSettingsReport
    $Report += checkUsedSKUReport
    $Report += checkAdminRoleReport
    $Report += checkBreakGlassAccountReport -Create $script:CreateBreakGlassAccount
    $Report += checkUserMfaStatusReport
    $Report += checkGuestUserReport
    $Report += checkSecurityDefaultsReport -Enable $script:EnableSecurityDefaults -Disable $script:DisableSecurityDefaults
    $Report += checkConditionalAccessPolicyReport
    $Report += checkNamedLocationReport
    $Report += checkAppProtectionPolicesReport

    if ($script:AddSharePointOnlineReport -or $script:DisableAddToOneDrive) {
        $Report += "<br><hr><h2 id='SPO'>SharePoint Online</h2>"
        Write-Host "SharePoint Online"
        $Report += checkSpoTenantReport -DisableAddToOneDrive $script:DisableAddToOneDrive
    }
    if ($script:AddExchangeOnlineReport -or $script:SetMailboxLanguage -or $script:DisableSharedMailboxLogin -or $script:EnableSharedMailboxCopyToSent -or $script:HideUnifiedMailboxFromOutlookClient) {
        $Report += "<br><hr><h2 id='EXO'>Exchange Online</h2>"
        Write-Host "Exchange Online"
        $Report += checkMailDomainReport
        $Report += checkMailConnectorReport
        $Report += checkMailboxReport -Language $script:SetMailboxLanguage
        $Report += checkSharedMailboxReport -Language $script:SetMailboxLanguage -DisableLogin $script:DisableSharedMailboxLogin -EnableCopy $script:EnableSharedMailboxCopyToSent
        $Report += checkUnifiedMailboxReport -HideFromClient $script:HideUnifiedMailboxFromOutlookClient
    }
    if (-not $KeepGraphSessionAlive) {
        disconnectGraph
    }
    if (-not $KeepSpoSessionAlive) {
        if ($script:AddSharePointOnlineReport -or $script:DisableAddToOneDrive) { 
            disconnectSpo
        }
    }
    if (-not $KeepExoSessionAlive) {
        if ($script:AddExchangeOnlineReport -or $script:SetMailboxLanguage -or $script:DisableSharedMailboxLogin -or $script:EnableSharedMailboxCopyToSent -or $script:HideUnifiedMailboxFromOutlookClient) {
            disconnectExo
        }
    }

    <# CSS styles section #>
    $Header = @"
<title>$($ReportTitle)</title>
<link rel="icon" type="image/png" href="$($ReportImageUrl)">
<style>
html {
    display: table;
    margin: auto;
}
body {
    display: table-cell;
    vertical-align: middle;
    padding-right: 200px;
    padding-left: 200px;
}
h1 {
    font-family: Arial, Helvetica, sans-serif;
    color: #666666;
    font-size: 32px;
}
h2 {
    font-family: Arial, Helvetica, sans-serif;
    color: #666666;
    font-size: 24px;
}
h3 {
    font-family: Arial, Helvetica, sans-serif;
    color: #666666;
    font-size: 16px;

}
p {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 14px;
}
a {
    font-family: Arial, Helvetica, sans-serif;
    font-size: 16px;
    text-decoration: none;
    color: #666666;
}
ul {
    list-style-type: none;
    margin-top: 5px;
}
li {
    padding: 5px;
}
table {
    font-size: 14px;
    border: 0px;
    font-family: Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    margin: 25px 0;
    min-width: 400px;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
}
th,
td {
    padding: 4px;
    margin: 0px;
    border: 0;
    padding: 12px 15px;
}
th {
    background: #666666;
    color: #fff;
    font-size: 11px;
    padding: 10px 15px;
    vertical-align: middle;
}
tbody tr:nth-child(even) {
    background: #f0f0f2;
}
thead tr {
    color: #ffffff;
    text-align: left;
}
tbody tr {
    border-bottom: 1px solid #dddddd;
}
tbody tr:nth-of-type(even) {
    background-color: #f3f3f3;
}
.red {
    color: red;
}
.orange {
    color: orange;
}
.TOC {
    margin: 5px;
}
#FootNote {
font-family: Arial, Helvetica, sans-serif;
color: #666666;
font-size: 12px;
}
</style>
"@

    <# HTML report section #>
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $ReportTitleHtml = "<h1>" + $ReportTitle + "</h1>"
    $ReportName = ("Microsoft365-Report-$($script:CustomerName).html").Replace(" ", "")
    $PostContentHtml = @"
<p id='FootNote'>$($script:VersionMessage)</p>
<p id='FootNote'>Creation date: $(Get-Date -Format "dd.MM.yyyy HH:mm")</p>
"@
    Write-Host "Generating HTML report:" $ReportName
    $Report = ConvertTo-Html -Body "$ReportTitleHtml $Report" -Title $ReportTitle -Head $Header -PostContent $PostContentHtml
    $Report | Out-File $Desktop\$ReportName
    Invoke-Item $Desktop\$ReportName
    if ($script:InteractiveMode) { Read-Host "Click [ENTER] key to exit AzureAdDeployer" }
}
Set-Alias aaddepl -Value Invoke-AzureAdDeployer