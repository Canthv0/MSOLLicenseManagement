:warning: **MSOL Cmdlets for managment of license are being removed!  I have started migrating these cmdlets to MS Graph but it will be some time before this is complete.** :warning:


# About MSOLLicenseManagement
PowerShell Module that provides functions to simplify the management of License Assignment, Swapping, Updating, and Reporting in Office 365.

# Recommended Install
Install the powershell module from the powershell gallery using `Install-Module MSOLLicenseManagement` from the server where data will be gathered.

Additional information can be found on how to install powershell modules here:
https://docs.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-7

# Manual Install
If the server where data is being gathered doesn't have access to the internet directly then either:

* Install the module to another machine in the org that does have access
* Download the module from https://github.com/Canthv0/MSOLLicenseManagement/releases
  * Extract the Zip file to a known location
  * From powershell in the extracted path run `Import-Module .\MSOLLicenseManagement.psd1`
  * Change to a different directory than the one with the psd1 file

# How to use
This module provides the following cmdlets to help with the management of licenses

### `Get-MGUserLicenseReport`
Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none


### `Get-MSOLUserLicenseReport`
Gets a csv report file with users and their assigned licenses.

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
OverWrite|Overwrite existing output file.|False
Logfile|Logfile location|None
IncludeDeletedUsers|Includes deleted users if reporting on ALL users|False

### `Add-MSOLUserLicense`
Adds a SKU to a collection of users.
Can set enabled OR disabled plans when added SKU

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
SKU|SKU to be added|None
Location|Sets user location|None
PlansToDisable|Comma seperated list of plans to Disable|None
PlansToEnable|Comma seperate list of plans to Enable|None
Logfile|Logfile location|None

### `Convert-MSOLUserLicenseToExplicit`
Converts a SKU from inherited (group assigned) to Explicitly assigned

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
SKU|SKU to be converted|None
Logfile|Logfile location|None


### `Remove-MSOLUserLicense`
Removes the provided SKU from the list of users.
Only works on explictly assigned SKUs

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
SKU|SKU to be removed|None
Logfile|Logfile location|None

### `Set-MSOUserLicensePlan`
OVERWRITES the current plan settings for an assigned SKU with the new plan settings provided.

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
SKU|SKU to be modified|None
PlansToDisable|Comma seperated list of plans to Disable|None
PlansToEnable|Comma seperate list of plans to Enable|None
Logfile|Logfile location|None

### `Switch-MSOLUserLicense`
Switch the user from SKU A to SKU B without interrupting license assignment.

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
SKU|SKU to be added|None
SKUToReplace|SKU that will be replaced|None
Location|Sets user location|None
PlansToDisable|Comma seperated list of plans to Disable|None
PlansToEnable|Comma seperate list of plans to Enable|None
Logfile|Logfile location|None
AttemptToMaintainPlans|Will attempt to maintain existing plan settings|False

### `Update-MSOLUserLicensePlan`
Updates the current plan settings for an assigned SKU while maintaining existing plan settings.

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Command Seperated List</br>Array of objects| none
SKU|SKU to be modified|None
PlansToDisable|Comma seperated list of plans to Disable|None
PlansToEnable|Comma seperate list of plans to Enable|None
Logfile|Logfile location|None

## Common Usage cases

### SKU report on all users

`Get-MSOLUserLicenseReport -logfile C:\temp\license.log`

Returns a CSV file with all users license and plan assignments in the C:\temp Directory

### Switch a user from one license to another with some plans disabled

`Switch-MSOLUserLicense -Users $NewUsers -Logfile C:\temp\switch_license.log -SKU company:ENTERPRISEPACK -SKUTOReplace company:STANDARDPACK -PlanstoDisable Deskless,Sway,Teams1`

Replaces the STANDARDPACK with the ENTERPRISEPACK and disables Deskless, Sway, and Teams.

### Set a collection of users to given set of plans

`Set-MSOLUserLicensePlan -Users $Promoted -Logfile C:\temp\set_license.log -SKU company:ENTERPRISEPACK -PlanstoEnable SWAY,TEAMS1,EXCHANGE_S_ENTERPRISE,YAMMER_ENTERPRISE,OFFICESUBSCRIPTION`

Set the enabled Plans for the ENTERPRISEPACK to be SWAY,TEAMS1,EXCHANGE_S_ENTERPRISE,YAMMER_ENTERPRISE,OFFICESUBSCRIPTION for all users in $Promoted.  This will overwrite any current plan settings for the users.

# Important Notes
* Set-MSOLUserLicense WILL overwrite any existing plan settings with the new provided values.  To modify a user without impacting their current plans use Update-MSOLUserLicensePlan