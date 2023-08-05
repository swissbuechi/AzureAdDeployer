<# SharePoint Tenant section #>
function Get-SPOTenantReport {
    param(
        [System.Boolean]$DisableAddToOneDrive
    )
    Write-Host "Checking tenant settings"
    if ($DisableAddToOneDrive) {
        Write-Host "Disable add to OneDrive button"
        Set-PnPTenant -DisableAddToOneDrive $True
    }
    $Report = Get-PnPTenant | ConvertTo-Html -As List -Property LegacyAuthProtocolsEnabled, DisableAddToOneDrive, ConditionalAccessPolicy, SharingCapability, RequireAcceptingAccountMatchInvitedAccount, PreventExternalUsersFromResharing, DefaultSharingLinkType -Fragment -PreContent "<h3 id='SPO_SETTINGS'>Tenant settings</h3>" -PostContent "<p>ConditionalAccessPolicy: AllowFullAccess, AllowLimitedAccess, BlockAccess</p>
<p>SharingCapability: Disabled, ExternalUserSharingOnly, ExternalUserAndGuestSharing, ExistingExternalUserSharingOnly</p>
<p>DefaultSharingLinkType: None, Direct, Internal, AnonymousAccess</p>"
    $Report = $Report -Replace "<td>RequireAcceptingAccountMatchInvitedAccount:</td><td>False</td>", "<td>RequireAcceptingAccountMatchInvitedAccount:</td><td class='red'>False</td>"
    $Report = $Report -Replace "<td>LegacyAuthProtocolsEnabled:</td><td>True</td>", "<td>LegacyAuthProtocolsEnabled:</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>DisableAddToOneDrive:</td><td>False</td>", "<td>DisableAddToOneDrive:</td><td class='red'>False</td>"
    # $Report = $Report -Replace "<td>DisplayStartASiteOption:</td><td>True</td>", "<td>DisplayStartASiteOption:</td><td class='red'>True</td>"
    $Report = $Report -Replace "<td>ConditionalAccessPolicy:</td><td>AllowFullAccess</td>", "<td>ConditionalAccessPolicy:</td><td class='red'>AllowFullAccess</td>"
    $Report = $Report -Replace "<td>SharingCapability:</td><td>ExternalUserAndGuestSharing</td>", "<td>SharingCapability:</td><td class='red'>ExternalUserAndGuestSharing</td>"
    $Report = $Report -Replace "<td>PreventExternalUsersFromResharing:</td><td>False</td>", "<td>PreventExternalUsersFromResharing:</td><td class='red'>False</td>"
    $Report = $Report -Replace "<td>DefaultSharingLinkType:</td><td>AnonymousAccess</td>", "<td>DefaultSharingLinkType:</td><td class='red'>AnonymousAccess</td>"
    return $Report
}