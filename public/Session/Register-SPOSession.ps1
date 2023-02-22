function Connect-SPOSession {
    # if (Request-SPOSession) { Disconnect-SPOSession }
    Disconnect-SPOSession
    Write-Host "Connecting SharePoint Online PowerShell"
    if ($PSVersionTable.PSEdition -eq "Core") { Connect-PnPOnline -Url (Get-SPOAdminURL) -Interactive -LaunchBrowser }
    if ($PSVersionTable.PSEdition -eq "Desktop") { Connect-PnPOnline -Url (Get-SPOAdminURL) -Interactive }
    if ( -not (Request-SPOSession)) { exit }
}
function Get-SPOAdminURL {
    return ((Invoke-MgGraphRequest -Method GET -Uri https://graph.microsoft.com/v1.0/sites/root).siteCollection.hostname) -replace ".sharepoint.com", "-admin.sharepoint.com"
}
function Disconnect-SPOSession {
    Write-Host "Disconnecting existing SharePoint Online PowerShell"
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
function Request-SPOSession {
    if (Get-PnPConnection) {
        Write-Host "Connected to SharePoint Online PowerShell tenant $((Get-PnPConnection).Url)"
        return $true
    }
    Write-Host "Not connected to SharePoint Online PowerShell"
    return $false
}