Function Get-AllModulesInstalled {
    param (
        [switch]$Install
    )
    $AllModulesInstalled = $true
    $AvailableModules = Get-Module -ListAvailable
    foreach ($Dependency in $script:Dependencies) {
        $Module = $AvailableModules | Where-Object { $_.Name -eq $Dependency.ModuleName -and $_.Version -eq $Dependency.ModuleVersion }
        if ($null -eq $Module) {
            if ($Install) {
                Write-Host "Module $($Dependency.ModuleName) with version $($Dependency.ModuleVersion) is not installed. Installing it now." -ForegroundColor Yellow
                Install-Module -Name $Dependency.ModuleName -RequiredVersion $Dependency.ModuleVersion -Force
            }
            else {
                Write-Host "Module $($Dependency.ModuleName) with version $($Dependency.ModuleVersion) is not installed. Please install it first." -ForegroundColor Red
                Write-Host "Install-Module -Name $($Dependency.ModuleName) -RequiredVersion $($Dependency.ModuleVersion) -Force"
                $AllModulesInstalled = $false
            }
        }
    }
    return $AllModulesInstalled
}