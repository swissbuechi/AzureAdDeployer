function Connect-GraphSession {
    # if (Request-GraphSession) { Disconnect-GraphSession }
    Disconnect-GraphSession
    Write-Host "Connecting Graph API PowerShell"
    Connect-MgGraph -ContextScope Process -ClientId "acecb20a-2e59-48c2-8a3c-aa9a9bd3352c" -Scopes "Policy.Read.All, Policy.ReadWrite.ConditionalAccess, Policy.ReadWrite.Authorization, Application.Read.All, User.ReadWrite.All, Domain.Read.All, Directory.ReadWrite.All, RoleManagement.ReadWrite.Directory, DeviceManagementApps.ReadWrite.All, Sites.Read.All, AuditLog.Read.All, UserAuthenticationMethod.Read.All, Organization.Read.All" | Out-Null
    if ( -not (Request-GraphSession)) { return }
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