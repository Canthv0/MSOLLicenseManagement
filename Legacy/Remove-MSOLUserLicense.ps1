# Removes a SKU from a specified collection of users
Function Remove-MSOLUserLicense {
    <#
 
	.SYNOPSIS
	Removes a license SKU from users.

	.DESCRIPTION
	Removes a licese SKU from a user or collection of users.

	* Prompts for SKU if none provided
	* Logs all activity to a log file
	
	.PARAMETER Users
    Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.
    
    .PARAMETER LogFile
	File to log all actions taken by the function.

	.PARAMETER SKU
	SKU that should be removed.

	.OUTPUTS
	Log file showing all actions taken by the function.

	.EXAMPLE
	Remove-MSOLUserLicense -Users $LeavingEmployees -Logfile C:\temp\remove_license.log -SKU company:ENTERPRISEPACK
	
	Removes the SKU ENTERPRISEPACK from all users in the $LeavingEmployees variable

    #>
    
    Param
    (
        [Parameter(Mandatory)]
        [array]$Users,
        [Parameter(Mandatory)]
        [string]$LogFile,
        [string]$SKU
    )
	
    # Make sure we have a valid log file path
    Test-LogPath -LogFile $LogFile

    # Make sure we have the connection to MSOL
    Test-MSOLServiceConnection

    # Make user our Users object is valid
    [array]$Users = Test-UserObject -ToTest $Users

    # If no value of SKU passed in then call Select-Sku to allow one to be picked
    if ([string]::IsNullOrEmpty($SKU)) {
        $SKU = Select-SKU -Message "Select SKU to remove from the Users:"
    }
    # If a value has been passed in verify it
    else {
        $SKU = Select-SKU -SKUToCheck $SKU
    }

    # "Zero" out the user counter
    [int]$i = 1
    [int]$ErrorCount = 0
	
    # Remove License to the users passed in
    Foreach ($Account in $Users) {
		
        Write-Log ("==== Processing User " + $account.UserPrincipalName + " ====")

        # Build and run our license set command
        [string]$Command = ("Set-MsolUserLicense -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -RemoveLicense " + $SKU + " -ErrorAction Stop -ErrorVariable CatchError")
        Write-Log ("Running: " + $Command)
        
        # Try our command
        try { 
            Invoke-Expression $Command 
        }
        # If we have any error write out and stop
        # Doing this so I can customize the error later
        ## TODO: Update error with resume information!
        catch { 
            Write-Log ("[ERROR] - " + $CatchError.ErrorRecord)
            Write-Log ("[ERROR] - " + $Account.UserPrincipalName + " Failed to remove SKU " + $SKU)
            Write-Error ("Failed to successfully remove license from user " + $account.UserPrincipalName)
            $ErrorCount++
        }
		
        # Update the progress bar and increment our counter
        Update-Progress -CurrentCount $i -MaxCount $Users.Count -Message "Removing SKU"

        # Update our user counter
        $i++
    }

    Write-Log ("Finished Removing SKU " + $SKU + " from " + $Users.Count + " Users.")
    If ($ErrorCount -gt 0) {
        Write-Log ($ErrorCount.ToString() + " ERRORS DURING PROCESSING PLEASE REVIEW ENTRIES WITH '[ERROR]' FOR MORE INFORMATION")
    }
}