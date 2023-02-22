function Install-DesktopIcon {
    if ($PSVersionTable.Platform -eq "Unix") {
        Write-Host "Creating a desktop shortcut is currently not supported on Unix"
        return
    }
    New-DesktopShortcut -ShortcutTargetPath 'powershell.exe' -ShortcutDisplayName 'AzureAdDeployer' -ShortcutArguments '-noexit -ExecutionPolicy Bypass -Command "Invoke-AzureAdDeployer"' -IconFile (Get-IconPath) -PinToStart
}
function Get-IconPath {
    $Path = (Get-Module -ListAvailable -Name "AzureAdDeployer").Path
    $Directory = (Get-Item $Path).DirectoryName
    return "$Directory\logo\logo.ico"
}