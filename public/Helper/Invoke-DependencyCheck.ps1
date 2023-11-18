Function Get-AllModulesInstalled {
    param (
        [switch]$Install
    )

    $Dependencies = @(
        @{
            ModuleName    = "PnP.PowerShell"; 
            ModuleVersion = "2.2.0"; 
        }, @{
            ModuleName    = "ExchangeOnlineManagement"; 
            ModuleVersion = "3.2.0"; 
        }, @{
            ModuleName    = "DnsClient-PS"; 
            ModuleVersion = "1.1.1"; 
        }, @{
            ModuleName    = "Microsoft.Graph.Authentication"; 
            ModuleVersion = "1.28.0"; 
        }, @{
            ModuleName    = "Microsoft.Graph.Users"; 
            ModuleVersion = "1.28.0"; 
       }, @{
            ModuleName    = "Microsoft.Graph.Groups"; 
            ModuleVersion = "1.28.0"; 
        }, @{
            ModuleName    = "Microsoft.Graph.Identity.DirectoryManagement"; 
            ModuleVersion = "1.28.0"; 
        }, @{
            ModuleName    = "Microsoft.Graph.DeviceManagement.Enrolment"; 
            ModuleVersion = "1.28.0"; 
        }, @{
            ModuleName    = "Microsoft.Graph.Identity.SignIns"; 
            ModuleVersion = "1.28.0"; 
        }, @{
            ModuleName    = "Microsoft.Graph.Devices.CorporateManagement"; 
            ModuleVersion = "1.28.0"; 
        }
    )

    $AllModulesInstalled = $true
    $AvailableModules = Get-Module -ListAvailable
    foreach ($Dependency in $Dependencies) {
        $Module = $AvailableModules | Where-Object { $_.Name -eq $Dependency.ModuleName -and $_.Version -eq $Dependency.ModuleVersion }
        if ($null -eq $Module) {
            if ($Install) {
                Write-Host "Module $($Dependency.ModuleName) with version $($Dependency.ModuleVersion) is not installed. Installing it now." -ForegroundColor Yellow
                Install-Module -Name $Dependency.ModuleName -RequiredVersion $Dependency.ModuleVersion -Force
                Import-Module -Name $Dependency.ModuleName -RequiredVersion $Dependency.ModuleVersion -Force
            }
            else {
                Write-Host "Module $($Dependency.ModuleName) with version $($Dependency.ModuleVersion) is not installed. Please install it first." -ForegroundColor Red
                $AllModulesInstalled = $false
            }
        }
        else {
            Import-Module -Name $Dependency.ModuleName -RequiredVersion $Dependency.ModuleVersion -Force
        }
    }
    if (-not ($AllModulesInstalled)) {
        Write-Host "Install dependencies: aaddepl -InstallDependencies"
    }
    if (-not ($Install)) {
        return $AllModulesInstalled
    }
}