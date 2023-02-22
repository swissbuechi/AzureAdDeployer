$script:ModuleInfos = Import-PowerShellDataFile -Path "$PsScriptRoot/AzureAdDeployer.psd1"
$Public = @(Get-ChildItem -Path $PSScriptRoot\public\*.ps1 -Recurse -ErrorAction SilentlyContinue)
foreach ($import in @($Public)) {
    try { . $import.FullName }
    catch { Write-Error -Message "Failed to import function $($import.FullName): $_" }
}