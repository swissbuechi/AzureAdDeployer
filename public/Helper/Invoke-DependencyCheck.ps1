Function Get-AllModulesInstalled {
    param (
        [switch]$Install
    )

    $Dependencies = @(
        @{
            ModuleName    = "PnP.PowerShell"; 
            ModuleVersion = "2.1.1"; 
            Guid          = "0b0430ce-d799-4f3b-a565-f0dca1f31e17"
        }, @{
            ModuleName    = "ExchangeOnlineManagement"; 
            ModuleVersion = "3.1.0"; 
            Guid          = "b5eced50-afa4-455b-847a-d8fb64140a22"
        }, @{
            ModuleName    = "DnsClient-PS"; 
            ModuleVersion = "1.1.1"; 
            Guid          = "698438cc-f80d-4b88-aa04-16e302c1f326"
        }, @{
            ModuleName    = "Microsoft.Graph.Authentication"; 
            ModuleVersion = "1.27.0"; 
            Guid          = "883916f2-9184-46ee-b1f8-b6a2fb784cee"
        }, @{
            ModuleName    = "Microsoft.Graph.Users"; 
            ModuleVersion = "1.27.0"; 
            Guid          = "71150504-37a3-48c6-82c7-7a00a12168db"
        }, @{
            ModuleName    = "Microsoft.Graph.Identity.DirectoryManagement"; 
            ModuleVersion = "1.27.0"; 
            Guid          = "c767240d-585c-42cb-bb2f-6e76e6d639d4"
        }, @{
            ModuleName    = "Microsoft.Graph.DeviceManagement.Enrolment"; 
            ModuleVersion = "1.27.0"; 
            Guid          = "447dd5b5-a01b-45bb-a55c-c9ecce3e820f"
        }, @{
            ModuleName    = "Microsoft.Graph.Identity.SignIns"; 
            ModuleVersion = "1.27.0"; 
            Guid          = "60f889fa-f873-43ad-b7d3-b7fc1273a44f"
        }, @{
            ModuleName    = "Microsoft.Graph.Devices.CorporateManagement"; 
            ModuleVersion = "1.27.0"; 
            Guid          = "39dbb3bc-1a84-424a-9efe-683be70a1810"
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
            }
            else {
                Write-Host "Module $($Dependency.ModuleName) with version $($Dependency.ModuleVersion) is not installed. Please install it first." -ForegroundColor Red
                $AllModulesInstalled = $false
            }
        }
    }
    if (-not ($AllModulesInstalled)) {
        Write-Host "Install dependencies: aaddepl -InstallDependencies"
    }
    if (-not ($Install)) {
        return $AllModulesInstalled
    }
}