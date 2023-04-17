function Install-DesktopIcon {
    if ($PSVersionTable.Platform -eq "Unix") {
        Write-Host "Creating a Desktop shortcut is currently not supported on Unix"
        return
    }
    Write-Host "Creating a Desktop icon: AzureAdDeployer.lnk"
    New-DesktopShortcut -ShortcutTargetPath 'pwsh' -ShortcutDisplayName 'AzureAdDeployer' -ShortcutArguments '-NoExit -NoProfile -Command Invoke-AzureAdDeployer' -IconFile (Get-IconPath) -PinToStart
}
function Get-IconPath {
    $Path = (Get-InstalledModule "AzureAdDeployer").InstalledLocation
    $Path = $Path -replace ('\\', '\')
    return "$($Path)\logo\logo.ico"
}