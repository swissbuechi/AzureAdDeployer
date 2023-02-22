function Connect-GraphSession {
    # if (Request-GraphSession) { Disconnect-GraphSession }
    Disconnect-GraphSession
    Write-Host "Connecting Graph API PowerShell"
    Connect-MgGraph -Scopes "Policy.Read.All, Policy.ReadWrite.ConditionalAccess, Application.Read.All,
User.Read.All, User.ReadWrite.All, Domain.Read.All, Directory.Read.All, Directory.ReadWrite.All,
RoleManagement.ReadWrite.Directory, DeviceManagementApps.Read.All, DeviceManagementApps.ReadWrite.All,
Policy.ReadWrite.Authorization, Sites.Read.All, AuditLog.Read.All, UserAuthenticationMethod.Read.All, Organization.Read.All" | Out-Null
    if ( -not (Request-GraphSession)) { exit }
}
function Disconnect-GraphSession {
    Write-Host "Disconnecting existing Graph API PowerShell"
    Disconnect-Graph -ErrorAction SilentlyContinue | Out-Null
}
function Request-GraphSession {
    if (Get-MgContext) {
        Write-Host "Connected to Graph API PowerShell using $((Get-MgContext).Account) account"
        return $true
    }
    Write-Host "Not connected to Graph API PowerShell"
    return $false
}