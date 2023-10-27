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
Users|UserPrincipalName</br>Comma Seperated List</br>Array of objects| none


### `Get-MSOLUserLicenseReport`
Gets a csv report file with users and their assigned licenses.

Switch | Description|Default
-------|-------|-------
Users|UserPrincipalName</br>Comma Seperated List</br>Array of objects| none
OverWrite|Overwrite existing output file.|False


## Common Usage cases

### Set a collection of users to given set of plans

# Important Notes
