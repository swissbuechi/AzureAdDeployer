<p align="center">
  <a href="https://www.powershellgallery.com/packages/AzureAdDeployer/"><img src="https://img.shields.io/powershellgallery/v/AzureAdDeployer"></a>
  <a href="https://www.powershellgallery.com/packages/AzureAdDeployer/"><img src="https://img.shields.io/badge/platform-windows-green"></a>
  <a href="https://www.powershellgallery.com/packages/AzureAdDeployer/"><img src="https://img.shields.io/badge/platform-macos-green"></a>
  <a href="https://www.powershellgallery.com/packages/AzureAdDeployer/"><img src="https://img.shields.io/badge/platform-linux-green"></a>
  <a href="https://www.powershellgallery.com/packages/AzureAdDeployer/"><img src="https://img.shields.io/github/languages/code-size/swissbuechi/AzureAdDeployer"></a>
  <a href="https://www.powershellgallery.com/packages/AzureAdDeployer/"><img src="https://img.shields.io/powershellgallery/dt/AzureAdDeployer"></a>
</p>

<p align="center">

# AzureAdDeployer

 Tool to analyze and remediate Microsoft 365 according to current security best practices.

<img src="./logo/logo.png"  width="200" height="200">

## System requirements

- Required PowerShell 5.1 or higher
- Windows PowerShell and PowerShell Core (Windows, macOS, Linux) supported
- Install required [Dependencies](#dependencies)

## Installation

### Install via PowerShellGallery (recommended)

The module is published on the PowerShellGallery. You can install this module directly from the PowerShellGallery with the following command:

`Install-Module -Name AzureAdDeployer -Scope AllUsers`

### Update via PowerShellGallery

```PowerShell
Update-Module -Name AzureAdDeployer -Scope AllUsers
$Latest = Get-InstalledModule AzureAdDeployer; 
Get-InstalledModule $ModuleName -AllVersions | ? {$_.Version -ne $Latest.Version} | Uninstall-Module
```

### Uninstall via PowerShellGallery

`Uninstall-Module -Name AzureAdDeployer -Scope AllUsers`

## Dependencies

### Uninstall previously installed dependencies (optional)

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

### Install dependencies

Windows PowerShell 5.1 (not Core!) as administrator:

```PowerShell
Install-Module -Name Microsoft.Graph -Scope AllUsers
Install-Module -Name PnP.PowerShell -Scope AllUsers
Install-Module -Name ExchangeOnlineManagement -Scope AllUsers
Install-Module -Name DnsClient-PS -Scope AllUsers #Only on Mac and Linux required
```

### Update dependencies

Windows PowerShell 5.1 (not Core!) as administrator:

```PowerShell
Update-Module -Name Microsoft.Graph -Scope AllUsers
Update-Module -Name PnP.PowerShell -Scope AllUsers
Update-Module -Name ExchangeOnlineManagement -Scope AllUsers
Update-Module -Name DnsClient-PS -Scope AllUsers #Only on Mac and Linux required
```

## Usage

### Interactive GUI

`Invoke-AzureAdDeployer`

Alias: `aaddepl`

### Azure Active Directory + Exchange Online HTML report

`aaddepl -AddExchangeOnlineReport`

### Create a BreakGlass account

`aaddepl -CreateBreakGlassAccount`

### Disable Security Defaults

`aaddepl -DisableSecurityDefaults`

## Features

### General

- Generates a HTML report to your desktop called `Microsoft365-Report-<customer_name>.html`
- Interactive console interface

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

## Upcoming Features

Checkout the [AzureAdDeployer project board](https://github.com/users/swissbuechi/projects/4)

## Credits

### Icon

- <a href="https://www.flaticon.com/free-icons/error" title="error icons">Error icons created by Smashicons - Flaticon</a>

### DMARC and SPF check

- Functions inspired by: <https://github.com/T13nn3s/Invoke-SpfDkimDmarc>

### User MFA check

- Functions inspired by: <https://o365reports.com/2022/04/27/get-mfa-status-of-office-365-users-using-microsoft-graph-powershell/>
