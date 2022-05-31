# Generates an MSOL User License Report

Function Get-MGUserLicenseAssignmentState {
	<#
 
	.SYNOPSIS
	Generates a comprehensive license report.

	.DESCRIPTION
	Generates a license report on a all users or a specified group of users.
	By Default it will generate a new report file each time it is run.

	Report includes the following:
	* All licenses assigned to each user provided.
	* State of all plans inside of each license assignment
	
	.PARAMETER Users
	Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.

	.PARAMETER OverWrite
	If specified it will Over Write the current output CSV file instead of generating a new one.
    
	.OUTPUTS
	Log file showing all actions taken by the function.
	CSV file named License_Report_YYYYMMDD_X.csv that contains the report.

	.EXAMPLE
	Get-MSOLUserLicenseReport -LogFile C:\temp\report.log

	Creates a new License_Report_YYYYMMDD_X.csv file with the license report in the c:\temp directory for all users.

	.EXAMPLE
	Get-MSOLUserLicenseReport -LogFile C:\temp\report.log -Users $SalesUsers -Overwrite

	OverWrites the existing c:\temp\License_Report_YYYYMMDD.csv file with a license report for all users in $SalesUsers.
	
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
    
	$error.clear()
	$state = Invoke-MgGraphRequest -Uri $Uri -OutputType JSON 2>&1

	if ($error.count -eq 0) {
		Write-SimpleLogfile ("Found User, graph connected")
	} else {
		Write-SimpleLogfile ("Error Encountered: " + $error[0].tostring())
		Write-Error -message $error[0].tostring() -ErrorAction Stop 
	}
  
	Return ($state | ConvertFrom-JSON).licenseAssignmentStates
}