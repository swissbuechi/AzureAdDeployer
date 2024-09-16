function Connect-SPOSession {
    # if (Request-SPOSession) { Disconnect-SPOSession }
    Disconnect-SPOSession
    Write-Host "Connecting SharePoint Online PowerShell"
    Connect-PnPOnline -Url (Get-SPOAdminURL) -Interactive -LaunchBrowser -ClientId "acecb20a-2e59-48c2-8a3c-aa9a9bd3352c"
    if ( -not (Request-SPOSession)) { return }
}
function Get-SPOAdminURL {
    return ((Invoke-MgGraphRequest -Method GET -Uri https://graph.microsoft.com/v1.0/sites/root).siteCollection.hostname) -replace ".sharepoint.com", "-admin.sharepoint.com"
}
function Disconnect-SPOSession {
    Write-Host "Disconnecting existing SharePoint Online PowerShell"
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
    }
}
function Request-SPOSession {
    if (Get-PnPConnection) {
        Write-Host "Connected to SharePoint Online PowerShell tenant $((Get-PnPConnection).Url)"
        return $true
    }
    Write-Host "Not connected to SharePoint Online PowerShell"
    return $false
}