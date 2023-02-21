# AzureAdDeployer

## Features

### General

- Generates a HTML report to your desktop called `Microsoft365-Report-<customer_name>.html`
- Interactive console GUI

### Azure Active Directory

- User settings
  - Enterprise Application user consent: show, disable
  - Allowed to create apps: show, disable
  - Allowed to create secutity groups: show, disable
  - Allowed to create unified groups (Microsoft 365 groups): show, disable, create group
  - Allowed to read other users: show, disable
  - Allowed to create tenants: show
  - BlockMsolPowerShell: show, enable
- Device join settings: show
- Licenses: show
- Admin role assignments: show
- User mfa status: show
- Guest accounts: show
- BreakGlass account: show, create
- Security sefaults: show, enable, disable
- Conditional access policies: show, list locations
- App protection policies: show

### SharePoint Online

- Tenant settings:
  - Legacy authentication protocols enabled: show
  - Add to OneDrive button: show, disable
  - Conditional access policy: show
  - Sharing capability: show
  - Prevent external users from resharing: show
  - Default sharing link type: show

### Exchange Online

- Domains: show, check DKIM/DMARC/SPF
- Mail connector: show
- User mailbox: show, set language
- Shared mailbox: show, set language, disable login, enable copy to sent
- Unified mailbox: show, hide from client

## Infos

- Works on PowerShell Windows and PowerShell Core

## Installation

### Uninstall previously installed modules

PowerShell as user:

```PowerShell
$Documents = [Environment]::GetFolderPath("MyDocuments") 
Remove-Item $Documents\PowerShell\Modules\ -Recurse -Force
Remove-Item $Documents\WindowsPowerShell\Modules\ -Recurse -Force
```

Windows PowerShell 5.1 (not Core!) as administrator:

```PowerShell
Uninstall-Module -Name Microsoft365DSC

Uninstall-Module Microsoft.Graph
Get-InstalledModule Microsoft.Graph.* | %{ if($_.Name -ne "Microsoft.Graph.Authentication"){ Uninstall-Module $_.Name } }
Uninstall-Module Microsoft.Graph.Authentication

Uninstall-Module -Name PnP.PowerShell

Uninstall-Module -Name ExchangeOnlineManagement
```

### Installation

Windows PowerShell 5.1 (not Core!) as administrator:

```PowerShell
Install-Module -Name Microsoft.Graph -Scope AllUsers
Install-Module -Name PnP.PowerShell -Scope AllUsers
Install-Module -Name ExchangeOnlineManagement -Scope AllUsers
Install-Module -Name DnsClient-PS -Scope AllUsers #Only on Mac and Linux required
```

### Updating

Windows PowerShell 5.1 (not Core!) as administrator:

```PowerShell
Update-Module -Name Microsoft.Graph -Scope AllUsers
Update-Module -Name PnP.PowerShell -Scope AllUsers
Update-Module -Name ExchangeOnlineManagement -Scope AllUsers
Update-Module -Name DnsClient-PS -Scope AllUsers #Only on Mac and Linux required
```

## Arguments

### Azure Active directory

| Argument                                  | Description                                    |
| ----------------------------------------- | ---------------------------------------------- |
| `-CreateBreakGlassAccount`                | Create a BreakGlass Account if no one is found |
| `-EnableSecurityDefaults`                 | Enable security defaults                       |
| `-DisableSecurityDefaults`                | Disable security defaults                      |
| `-DisableEnterpiseApplicationUserConsent` | Disable enterprise application user consent    |
| `-DisableUsersToCreateAppRegistrations`   | Disable users to create app registrations      |
| `-DisableUsersToReadOtherUsers`           | Disable users to read other users              |
| `-DisableUsersToCreateSecurityGroups`     | Disable users to create security groups        |
| `-DisableUsersToCreateUnifiedGroups`      | Disable users to create unified groups         |
| `-CreateUnifiedGroupCreationAllowedGroup` | Create UnifiedGroupCreationAllowed group       |
| `-EnableBlockMsolPowerShell`              | Disable legacy MsolPowerShell access           |

### SharePoint Online

| Argument                | Description             |
| ----------------------- | ----------------------- |
| `-DisableAddToOneDrive` | Disable add to OneDrive |

### Exchange Online

| Argument                               | Description                                |
| -------------------------------------- | ------------------------------------------ |
| `-AddExchangeOnlineReport`             | Add a report section for Exchange Online   |
| `-SetMailboxLanguage`                  | Set mailbox language and location          |
| `-DisableSharedMailboxLogin`           | Disable direct login to shared mailbox     |
| `-EnableSharedMailboxCopyToSent`       | Enable shared mailbox copy to sent e-mails |
| `-HideUnifiedMailboxFromOutlookClient` | Hide unified mailbox from outlook client   |

### Advanced

| Argument                   | Description                                                              |
| -------------------------- | ------------------------------------------------------------------------ |
| `-UseExistingGraphSession` | Do not create a new Graph SDK PowerShell session                         |
| `-UseExistingSpoSession`   | Do not create a new SharePoint Online PowerShell session                 |
| `-UseExistingExoSession`   | Do not create a new Exchange Online PowerShell session                   |
| `-KeepGraphSessionAlive`   | Do not disconnect the Graph SDK PowerShell session after execution       |
| `-KeepSpoSessionAlive`     | Do not disconnect the SharePoint Online session after execution          |
| `-KeepExoSessionAlive`     | Do not disconnect the Exchange Online PowerShell session after execution |

## Usage

### Interactive GUI

`.\AzureAdDeployer.ps1`

### Azure Active Directory + Exchange Online HTML report

`.\AzureAdDeployer.ps1 -AddExchangeOnlineReport`

### Create a BreakGlass account

`.\AzureAdDeployer.ps1 -CreateBreakGlassAccount`

### Disable Security Defaults

`.\AzureAdDeployer.ps1 -DisableSecurityDefaults`

## ToDo

- Manage Self-service password reset (no API available)
- Manage Authentication methods available for users / Manage migration till Jan 24 <https://learn.microsoft.com/en-us/azure/active-directory/authentication/how-to-authentication-methods-manage>
- Enable enterpise state roaming (no API available)
- Create default Conditinal Access polices
- Find external e-mail forwardings
- Restrict access to PowerShell for edu tenants <https://learn.microsoft.com/en-us/schooldatasync/blocking-powershell-for-edu>
- Password policy <https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-password-ban-bad>
- Create default application protectin policy for iOS and Android
- Manage Enterprise application admin consent request policy <https://learn.microsoft.com/en-us/graph/api/adminconsentrequestpolicy-get?view=graph-rest-1.0&tabs=powershell>
- Check if required Modules are installed and imported -> `#require` is causing performance issues, long script startup times
- List externally shared files
- Inegrate Azure PowerShell module <https://learn.microsoft.com/en-us/powershell/azure/install-az-ps?view=azps-9.4.0>
  - Check if budget is set
  - Set usage location
- Set language of admin account to en-US
- Add Teams best-practice
  - Block 3.rd party apps
- Exchange Online
  - Switch to graph api where possible
  - Show security settings (spam, quarantine, safe links, etc)

## Credits

### Icons

- <a href="https://www.flaticon.com/free-icons/error" title="error icons">Error icons created by Smashicons - Flaticon</a>

### DMARC and SPF check

- <https://github.com/T13nn3s/Invoke-SpfDkimDmarc>

## Template code for later functions

### Conditional Access policies

```PowerShell
function deleteConditionalAccessPolicy {
    param (
        [Parameter(Mandatory = $true)]
        $Policies
    )
    foreach ($Policy in $Policies) {
        Write-Host "Removing existing Conditional Access policies"
        Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $Policy.Id
    }
}

function cleanUpConditionalAccessPolicy {
    $Policies = getConditionalAccessPolicy
    deleteConditionalAccessPolicy $Policies
}

function getNamedLocations {
    return Get-MgIdentityConditionalAccessNamedLocation -Property Id, DisplayName
}

function createConditionalAccessPolicy {
    $params = @{
        DisplayName   = "Require MFA from all unknown locations"
        State         = "enabled"
        Conditions    = @{
            Applications = @{
                IncludeApplications = @(
                    "All"
                )
            }
            Users        = @{
                IncludeUsers = @(
                    "All"
                )
                ExcludeUsers = @(
                    getBreakGlassAccount.Id
                )
            }
            Locations    = @{
                IncludeLocations = @(
                    "All"
                )
                ExcludeLocations = @(
                    "AllTrusted"
                )
            }
        }
        GrantControls = @{
            Operator        = "OR"
            BuiltInControls = @(
                "mfa"
            )
        }
    }
    New-MgIdentityConditionalAccessPolicy -BodyParameter $params
}
```

### Application protection policies

```PowerShell
function createAndroidAppProtectionPolicy {
    $Body = @{
        "@odata.type" = "#microsoft.graph.androidManagedAppProtection"
        displayName = "Test"
    }
    New-MgDeviceAppManagementAndroidManagedAppProtection -BodyParameter $Body
}
```
