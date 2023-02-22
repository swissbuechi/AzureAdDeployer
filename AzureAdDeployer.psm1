$Public = @(Get-ChildItem -Path $PSScriptRoot\public\*.ps1 -ErrorAction SilentlyContinue)
foreach ($import in @($Public)) {
    try { . $import.FullName }
    catch { Write-Error -Message "Failed to import function $($import.FullName): $_" }
}