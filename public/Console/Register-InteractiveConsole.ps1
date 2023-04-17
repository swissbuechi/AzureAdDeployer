<# Interactive inputs section #>
function Request-InteractiveMode {
    Param(
        $Parameters
    )
    if ($Parameters.Count) {
        return
    }
    $script:InteractiveMode = $true
}
function Show-InteractiveMenu {
    $script:AddExchangeOnlineReport = $true
    $script:AddSharePointOnlineReport = $true
    if (Get-ModuleUpdateNeeded) { 
        Get-ModuleUpdateMessageGUI
    }
    Show-MainMenu
}
function Show-MainMenu {
    $StartOptionValue = 0
    while (($Result -ne $StartOptionValue) -or ($Result -ne 1)) {
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
        $Result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
        switch ($Result) {
            0 { return }
            1 { Show-ConfigMenu }
            2 { $script:AddSharePointOnlineReport = ! $script:AddSharePointOnlineReport }
            3 { $script:AddExchangeOnlineReport = ! $script:AddExchangeOnlineReport }
        }
    }
}
function Show-ConfigMenu {
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
    $Result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
    switch ($Result) {
        0 { return }
        1 { Show-AADMenu }
        2 { Show-SPOMenu }
        3 { Show-EXOMenu }
    }
}
function Show-AADMenu {
    $StartOptionValue = 0
    while ($Result -ne $StartOptionValue) {
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
        $Result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
        switch ($Result) {
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
function Show-SPOMenu {
    $StartOptionValue = 0
    while ($Result -ne $StartOptionValue) {
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
        $Result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
        switch ($Result) {
            0 { return }
            1 { $script:DisableAddToOneDrive = ! $script:DisableAddToOneDrive }
        }
    }
}
function Show-EXOMenu {
    $StartOptionValue = 0
    while ($Result -ne $StartOptionValue) {
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
        $Result = $host.ui.PromptForChoice("", "", $Options, $StartOptionValue)
        switch ($Result) {
            0 { return }
            1 { $script:SetMailboxLanguage = ! $script:SetMailboxLanguage }
            2 { $script:DisableSharedMailboxLogin = ! $script:DisableSharedMailboxLogin }
            3 { $script:EnableSharedMailboxCopyToSent = ! $script:EnableSharedMailboxCopyToSent }
            4 { $script:HideUnifiedMailboxFromOutlookClient = ! $script:HideUnifiedMailboxFromOutlookClient }
        }
    }
}