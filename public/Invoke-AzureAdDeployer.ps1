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
        [switch]$DisableAddToOneDrive,
        [switch]$InstallDesktopIcon,
        [switch]$Version,
        [switch]$Help
    )
    $script:ReportTitle = "Microsoft 365 Security Report"
    $VersionNumber = $script:ModuleInfos.ModuleVersion
    $script:VersionMessage = "AzureAdDeployer version: $($VersionNumber)"
    $Repository = "https://github.com/swissbuechi/AzureAdDeployer"

    $script:ReportImageUrl = "https://cdn-icons-png.flaticon.com/512/3540/3540926.png"

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

    <# Script logic start section #>
    if ($Version) {
        Write-Host $VersionNumber
        return
    }
    if ($Help) {
        Write-Host "Checkout the documentation: $($Repository)#arguments"
        return
    }
    if ($InstallDesktopIcon) { 
        Install-DesktopIcon 
        return
    }
    Request-InteractiveMode -Parameters $PSBoundParameters
    if ($script:InteractiveMode) {
        Show-InteractiveMenu
    }

    if (-not $UseExistingGraphSession) { Connect-GraphSession }
    else { if ( -not (Request-GraphSession)) { return } }
    $TableOfContents = @()
    $TableOfContents += "<br><hr><h2>Contents</h2>"
    $TableOfContents += Get-AADTableOfContents

    if ($script:AddSharePointOnlineReport -or $script:DisableAddToOneDrive) { 
        if (-not $UseExistingSpoSession) { Connect-SPOSession }
        else { if ( -not (Request-SPOSession)) { return } }
        $TableOfContents += Get-SPOTableOfContents
    }

    if ($script:AddExchangeOnlineReport -or $script:SetMailboxLanguage -or $script:DisableSharedMailboxLogin -or $script:EnableSharedMailboxCopyToSent -or $script:HideUnifiedMailboxFromOutlookClient) {
        if (-not $UseExistingExoSession) { Connect-EXO }
        else { if ( -not (Request-EXOSession)) { return } }
        $TableOfContents += Get-EXOTableOfContents
    }

    $Report = @()
    $Report += Get-OrganizationReport
    $Report += $TableOfContents
    $Report += "<br><hr><h2 id='AAD'>Azure Active Directory</h2>"
    Write-Host "Azure Active Directory"
    $Report += Get-UserSettingsReport -DisableUserConsent $script:DisableEnterpiseApplicationUserConsent -DisableUsersToCreateAppRegistrations $script:DisableUsersToCreateAppRegistrations -DisableUsersToReadOtherUsers $script:DisableUsersToReadOtherUsers -DisableUsersToCreateSecurityGroups $script:DisableUsersToCreateSecurityGroups -DisableUsersToCreateUnifiedGroups $script:DisableUsersToCreateUnifiedGroups -CreateUnifiedGroupCreationAllowedGroup $script:CreateUnifiedGroupCreationAllowedGroup -EnableBlockMsolPowerShell $script:EnableBlockMsolPowerShell
    $Report += Get-DeviceJoinSettingsReport
    $Report += Get-UsedSKUReport
    $Report += Get-AdminRoleReport
    $Report += Get-BreakGlassAccountReport -Create $script:CreateBreakGlassAccount
    $Report += Get-UserMfaStatusReport
    $Report += Get-GuestUserReport
    $Report += Get-SecurityDefaultsReport -Enable $script:EnableSecurityDefaults -Disable $script:DisableSecurityDefaults
    $Report += Get-ConditionalAccessPolicyReport
    $Report += Get-NamedLocationReport
    $Report += Get-AppProtectionPolicesReport

    if ($script:AddSharePointOnlineReport -or $script:DisableAddToOneDrive) {
        $Report += "<br><hr><h2 id='SPO'>SharePoint Online</h2>"
        Write-Host "SharePoint Online"
        $Report += Get-SPOTenantReport -DisableAddToOneDrive $script:DisableAddToOneDrive
    }
    if ($script:AddExchangeOnlineReport -or $script:SetMailboxLanguage -or $script:DisableSharedMailboxLogin -or $script:EnableSharedMailboxCopyToSent -or $script:HideUnifiedMailboxFromOutlookClient) {
        $Report += "<br><hr><h2 id='EXO'>Exchange Online</h2>"
        Write-Host "Exchange Online"
        $Report += Get-MailDomainReport
        $Report += Get-MailConnectorReport
        # $Report += Get-UserMailboxReport -Language $script:SetMailboxLanguage
        $Report += Get-SharedMailboxReport -Language $script:SetMailboxLanguage -DisableLogin $script:DisableSharedMailboxLogin -EnableCopy $script:EnableSharedMailboxCopyToSent
        $Report += Get-UnifiedMailboxReport -HideFromClient $script:HideUnifiedMailboxFromOutlookClient
    }
    if (-not $KeepGraphSessionAlive) {
        Disconnect-GraphSession
    }
    if (-not $KeepSpoSessionAlive) {
        if ($script:AddSharePointOnlineReport -or $script:DisableAddToOneDrive) { 
            Disconnect-SPOSession
        }
    }
    if (-not $KeepExoSessionAlive) {
        if ($script:AddExchangeOnlineReport -or $script:SetMailboxLanguage -or $script:DisableSharedMailboxLogin -or $script:EnableSharedMailboxCopyToSent -or $script:HideUnifiedMailboxFromOutlookClient) {
            Disconnect-EXOSession
        }
    }

    <# HTML report section #>
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $ReportTitleHtml = "<h1>" + $ReportTitle + "</h1>"
    $ReportName = ("Microsoft365-Report-$($script:CustomerName).html").Replace(" ", "")
    $PostContentHtml = @"
<a id='FootNote' href="$($Repository)" target="blank">$($script:VersionMessage)</a>
<p id='FootNote'>Creation date: $(Get-Date -Format "dd.MM.yyyy HH:mm")</p>
"@
    Write-Host "Generating HTML report:" $ReportName
    $Report = ConvertTo-Html -Body "$ReportTitleHtml $Report" -Title $ReportTitle -Head (Get-Header) -PostContent $PostContentHtml
    $Report | Out-File $Desktop\$ReportName -Force
    Invoke-Item $Desktop\$ReportName
    if ($script:InteractiveMode) { Read-Host "Click [ENTER] key to exit AzureAdDeployer" }
}
Set-Alias aaddepl -Value Invoke-AzureAdDeployer