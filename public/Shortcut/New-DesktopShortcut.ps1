function New-DesktopShortcut {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [String]$ShortcutTargetPath,

        [Parameter(Mandatory = $true)]
        [String]$ShortcutDisplayName,

        [Parameter(Mandatory = $false)]
        [Switch]$PinToStart = $false,

        [Parameter(Mandatory = $false)]
        [String]$IconFile = $null,

        [Parameter(Mandatory = $false)]
        [String]$ShortcutArguments = $null,

        [Parameter(Mandatory = $false)]
        [String]$WorkingDirectory = $null
    )
    function Add-Shortcut {
        param (
            [Parameter(Mandatory)]
            [String]$ShortcutTargetPath,
            [Parameter(Mandatory)]
            [String] $DestinationPath,
            [Parameter()]
            [String] $WorkingDirectory
        )
        process {
            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut($destinationPath)
            $Shortcut.TargetPath = $ShortcutTargetPath
            $Shortcut.Arguments = $ShortcutArguments
            $Shortcut.WorkingDirectory = $WorkingDirectory
            if ($IconFile) {
                $Shortcut.IconLocation = $IconFile
            }
            $Shortcut.Save()
            [Runtime.InteropServices.Marshal]::ReleaseComObject($WshShell) | Out-Null
        }
    }
    function Test-RunningAsSystem {
        [CmdletBinding()]
        param()
        process {
            return ($(whoami -user) -match "S-1-5-18")
        }
    }
    function Get-DesktopDir {
        [CmdletBinding()]
        param()
        process {
            if (Test-RunningAsSystem) {
                $desktopDir = Join-Path -Path $env:PUBLIC -ChildPath "Desktop"
            }
            else {
                $desktopDir = $([Environment]::GetFolderPath("Desktop"))
            }
            return $desktopDir
        }
    }
    function Get-StartDir {
        [CmdletBinding()]
        param()
        process {
            if (Test-RunningAsSystem) {
                $startMenuDir = Join-Path $env:ALLUSERSPROFILE "Microsoft\Windows\Start Menu\Programs"
            }
            else {
                $startMenuDir = "$([Environment]::GetFolderPath("StartMenu"))\Programs"
            }
            return $startMenuDir
        }
    }

    $destinationPath = Join-Path -Path $(Get-DesktopDir) -ChildPath "$shortcutDisplayName.lnk"
    Add-Shortcut -DestinationPath $destinationPath -ShortcutTargetPath $ShortcutTargetPath -WorkingDirectory $WorkingDirectory

    if ($PinToStart.IsPresent -eq $true) {
        $destinationPath = Join-Path -Path $(Get-StartDir) -ChildPath "$shortcutDisplayName.lnk"
        Add-Shortcut -DestinationPath $destinationPath -ShortcutTargetPath $ShortcutTargetPath -WorkingDirectory $WorkingDirectory
    }
}