# o365recon
Script to retrieve information via O365 and Azure AD with valid credentials using Microsoft Graph API.

## Setup
Install the Microsoft Graph PowerShell module:
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

The script uses the following Microsoft Graph sub-modules (automatically installed with the main module):
- Microsoft.Graph.Users
- Microsoft.Graph.Groups
- Microsoft.Graph.Devices
- Microsoft.Graph.Identity.DirectoryManagement
- Microsoft.Graph.Applications
- Microsoft.Graph.Organization

### Requirements
- PowerShell 5.1 or PowerShell 7+
- Microsoft Graph PowerShell SDK
- Valid Office 365/Azure AD credentials
- Appropriate permissions to read directory data

#### Usage:
```powershell
.\o365recon.ps1 -azure
```

There is only one flag (`-azure`) and it is optional. You will be prompted to authenticate via Microsoft Graph. You may be prompted twice if MFA is enabled.

**Note:** This script has been updated to use the modern Microsoft Graph API, replacing the deprecated MSOnline and AzureAD modules which are no longer supported by Microsoft.

![ScreenShot](https://raw.github.com/nyxgeek/o365recon/master/screenshot.png?)
