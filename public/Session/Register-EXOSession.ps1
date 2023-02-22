function Connect-EXO {
    # if (Request-EXOSession) { Disconnect-EXOSession }
    Disconnect-EXOSession
    Write-Host "Connecting Exchange Online PowerShell"
    Connect-ExchangeOnline -ShowBanner:$false
    if ( -not (Request-EXOSession)) { exit }
}
function Disconnect-EXOSession {
    Write-Host "Disconnecting existing Exchange Online PowerShell"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}
function Request-EXOSession {
    if ((Get-ConnectionInformation).State -eq "Connected") {
        Write-Host "Connected to Exchange Online PowerShell using $((Get-ConnectionInformation).UserPrincipalName) account"
        return $true
    }
    Write-Host "Not connected to Exchange Online PowerShell"
    return $false
}