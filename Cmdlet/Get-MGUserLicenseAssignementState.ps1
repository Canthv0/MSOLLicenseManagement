# Generates an MSOL User License Report

Function Get-MGUserLicenseAssignmentState {
	<#
 
	.SYNOPSIS
	Gathers a users license state using a direct Graph call.
	There is no cmdlet for this one.

	.DESCRIPTION
	Uses Invoke-MgGraphRequest to request the licnese assignement information for a user

	What comes back is JSON encoded data that is converted to PS objects and returned to the caller.
		
	.PARAMETER UsersID
	UserID of the user to gather the information for.

	.OUTPUTS
	PS Objects with license assignment information.

	.EXAMPLE
	Get-MGUserLicenseAssignmentState -UserID 123456

	Returns the license assignement information to the calling function.

	#>
    
	param 
	(
		[Parameter(Mandatory = $true)]
		[array]$UserID
	)

	# Make sure we have the connection to MSOL
	Test-MGServiceConnection
	
	Write-SimpleLogFile "Gathering user license assignment state"

	$Uri = "https://graph.microsoft.com/v1.0/users/" + $UserID + "?`$select=licenseAssignmentStates"
  
	# clear any errors
	$error.clear()
	
	# Try to retrieve the state
	Try { $state = Invoke-MgGraphRequest -Uri $Uri -OutputType JSON -ErrorAction Stop }
	catch {
		Write-SimpleLogfile ("Error Encountered: " + $_)
		Write-Error -message $_ -ErrorAction Stop 
		break
 }
  
	Return ($state | ConvertFrom-JSON).licenseAssignmentStates
}