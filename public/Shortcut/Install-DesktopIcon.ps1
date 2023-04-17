function Install-DesktopIcon {
    if ($PSVersionTable.Platform -eq "Unix") {
        Write-Host "Creating desktop icons is currently not supported on Unix"
        return
    }
    $LoadingMessage = "$($script:VersionMessage) importing dependend PowerShell modules. This may take a few seconds..."
    Write-Host "Creating desktop icon: $($script:ModuleName).lnk"
    New-DesktopShortcut -ShortcutTargetPath 'pwsh' -ShortcutDisplayName $script:ModuleName -ShortcutArguments "-NoExit -NoProfile -Command (Write-Host $($LoadingMessage)), (Invoke-$($script:ModuleName))" -IconFile (Get-IconPath) -PinToStart
}
function Get-IconPath {
    $Path = (Get-InstalledModule $script:ModuleName).InstalledLocation
    $Path = $Path -replace ('\\', '\')
    return "$($Path)\logo\logo.ico"
}