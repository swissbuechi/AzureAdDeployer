function Install-DesktopIcon {
    if ($PSVersionTable.Platform -eq "Unix") {
        Write-Host "Creating a Desktop shortcut is currently not supported on Unix"
        return
    }
    Write-Host "Creating a Desktop icon: AzureAdDeployer.lnk"
    New-DesktopShortcut -ShortcutTargetPath 'powershell.exe' -ShortcutDisplayName 'AzureAdDeployer' -ShortcutArguments '-noexit -ExecutionPolicy Bypass -Command "Invoke-AzureAdDeployer"' -IconFile (Get-IconPath) -PinToStart
}
function Get-IconPath {
    $Path = (Get-InstalledModule "AzureAdDeployer").InstalledLocation
    $Path = $Path -replace ('\\', '\')
    return "$($Path)\logo\logo.ico"
}