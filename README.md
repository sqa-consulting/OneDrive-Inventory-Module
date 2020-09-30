# OneDrive Inventory Module
[![Website](https://img.shields.io/badge/Website-SQA--Consulting.com-blue?style=flat&logo=WordPress&link=https://sqa-consulting.com/home/)](https://sqa-consulting.com)
![Code Quality](https://github.com/sqa-consulting/OneDrive-Inventory-Module/workflows/Code%20Quality/badge.svg)

Table of Contents
============
<!-- toc -->
- [Introduction](#introduction)
- [Requirements](#requirements)
- [Cmdlets](#cmdlets)
- [Usage](#usage)

Introduction
============

We developed this module in house to allow us to inventory client OneDrive environments and provide management dashboards on the level of adoption across the enterprise.  

The module requires administrative access to the Microsoft365 SharepointOnline instance.  As a Sharepoint Admin does not have access (by default) to OneDrive sites, 
the module will first add explicit administrative permissions to each OneDrive site for the account provided.  It then inventories all OneDrive sites in parallel, before
revoking the explicit permissions to tidy up.  The addition and removal of permissions is the only change the module will make to your instance, other actions are
read-only.  As a convenience, a [Cmdlet](#cmdlets) is provided to only remove these permissions.

Parallel processing is achieved by creating child PowerShell processes for each OneDrive site.

The script will output CSV files to the configured [outputPath](#parameters).  The number of users within each CSV is determined by the paging size.  A larger paging size will require more memory but will allow additional parallel processing.

It is recommended that you create a specific account to conduct the audit as this process can be noisy on the audit logs.  

This method of API authentication does not support MFA, so either utilise an account without it or use an [application password](https://docs.microsoft.com/en-us/azure/active-directory/user-help/multi-factor-authentication-end-user-app-passwords).

This code is provided without Warranty.


Requirements
============
This module utilises two Microsoft published PowerShell modules and will fail if you have not got them installed.  Install them with:

### Microsoft.Online.SharePoint.PowerShell
```powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
```
### SharePointPnPPowerShellOnline
```powershell
Install-Module -Name SharePointPnPPowerShellOnline
```
### Windows
Whilst every effort has been made to make this code PowerShell Core compatible for future use, the required modules (specifically SharePointPnPPowerShellOnline) do not 
work fully in PowerShell Core so for now you will have to run on Win10.  Core parts of this code will be swapped out for direct REST interactions with the Graph API to overcome these issues and ensure pwsh compatibility.

Cmdlets
============
### Get-OneDriveInventory
This Cmdlet will request parameters if none are provided and outputs verbosely to the screen and to an output file.  Write-Host is utilised to colour the screen
output but is Windows specific.  Verbose logging will be written to the current working directory.

Parameters
- Site

  The Sharepoint site address.  This is the FQDN given when navigating to the Sharepoint admin centre in the Office365 admin portal
- Credential

  The PSCredential object to authenticate with Sharepoint.  Use Get-Credential to capture.
- Page

  The number of users to fetch in a single batch, defaults to 100.  A larger paging size requires additional memory.
- OutputPath

  The base directory for CSV output, defaults to the current working directory.  A directory will be created at this location following the format OneDriveInventory_<timestamp>
  
### Remove-SiteAdminPermissions
This Cmdlet is provided as a convenience.  It will connect to SharepointOnline, pull a list of all OneDrive sites and then attempt to revoke any administrative
permissions assigned to the provided credentials.  This can be used following a failed inventory run to force a clean-up.

Parameters
- Site

  The Sharepoint site address.  This is the FQDN given when navigating to the Sharepoint admin centre in the Office365 admin portal
- Credential

  The PSCredential object to authenticate with Sharepoint.  Use Get-Credential to capture.

Usage
============
```powershell
Import-Module ./OneDriveInventory
$Credentials = Get-Credential
Get-OneDriveInventory -Site http://sharepointcustomer-admin.sharepoint.com -Page 200 -Credential $Credentials -OutputPath "C:\Temp"
```
or
```powershell
Import-Module ./OneDriveInventory
Get-OneDriveInventory -Site http://sharepointcustomer-admin.sharepoint.com
```
and
```powershell
Import-Module ./OneDriveInventory
Remove-OneDriveAdminPermissions
```
PSScriptAnalyzer
=============
The module has been checked with [PSScriptAnalyzer](https://github.com/PowerShell/PSScriptAnalyzer) which is a great opensource project to audit PowerShell code for poor practice.  A configuration file is included in this repository.
