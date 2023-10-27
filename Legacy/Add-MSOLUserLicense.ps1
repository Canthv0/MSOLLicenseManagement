# Adds a SKU to a specified collection of users
Function Add-MSOLUserLicense {
    <#

 	.SYNOPSIS
	Adds licenses to users.

	.DESCRIPTION
	Adds a license SKU to a user or collection of users.

	* Can specify plans to enable or disable
	* Logs all activity to a log file
	* Can set location if desired

	.PARAMETER Users
	Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.

	.PARAMETER SKU
	SKU that should be added.

	.PARAMETER Location
	If provided will set the location of the user.

	.PARAMETER PlansToDisable
	Comma seperated list of SKU plans to Disable.

	.PARAMETER PlansToEnable
	Comma seperated list of SKU plans to Enable.

	.PARAMETER LogFile
	File to log all actions taken by the function.

	.OUTPUTS
	Log file showing all actions taken by the function.

	.EXAMPLE
	Add-MSOLUserLicense -Users $NewUsers -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK
	
	Adds the ENTERPRISEPACK SKU to all users in $NewUsers with all plans turned on.

	.EXAMPLE
	Add-MSOLUserLicense -Users $NewUsers -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK -PlanstoEnable Deskless,Sway,Teams1

	Adds the ENTERPRISEPACK SKU to all users in $NewUsers with ONLY the Deskless, Sway, and Teams1 Plans turned on.
		
    #>	
    
    Param
    (
        [Parameter(Mandatory = $true)]
        [array]$Users,
        [Parameter(Mandatory = $true)]
        [string]$LogFile,
        [string]$SKU,
        [string]$Location,
        [string[]]$PlansToDisable,
        [string[]]$PlansToEnable
    )

    # Make sure we have a valid log file path
    Test-LogPath -LogFile $LogFile

    # Make sure we have the connection to MSOL
    Test-MSOLServiceConnection

    # Make sure we didn't get both enable and disable
    if (!([string]::IsNullOrEmpty($PlansToDisable)) -and !([string]::IsNullOrEmpty($PlansToEnable))) {
        Write-Log "[ERROR] - Cannot use both -PlansToDisable and -PlansToEnable at the same time"
        Write-Error "Cannot use both -PlansToDisable and -PlansToEnable at the same time" -ErrorAction Stop
    }
	
    # Make user our Users object is valid
    [array]$Users = Test-UserObject -ToTest $Users

    # If no value of SKU passed in then call Select-Sku to allow one to be picked
    if ([string]::IsNullOrEmpty($SKU)) {
        $SKU = Select-SKU -Message "Select SKU to Add to the Users:"
    }
    # If a value has been passed in verify it
    else {
        $SKU = Select-SKU -SKUToCheck $SKU
    }

    # Testing the plan inputs to make sure they are valid
    if (!([string]::IsNullOrEmpty($PlansToDisable))) {
        Test-Plan -Sku $SKU -Plan $PlansToDisable
		
        # Get the license options
        $LicenseOption = Set-LicenseOption -DisabledPlansArray $PlansToDisable -SKU $SKU	

    }
    # If plans to enable has a value then we test them
    elseif (!([string]::IsNullOrEmpty($PlansToEnable))) {
        Test-Plan -Sku $SKU -Plan $PlansToEnable

        # Get the disabled plans and License options
        [string[]]$CalculatedPlansToDisable = Set-EnabledPlan -SKU $SKU -Plan $PlansToEnable
        $LicenseOption = Set-LicenseOption -DisabledPlansArray $CalculatedPlansToDisable -SKU $SKU
    }
    # If neither has been provided then just set default options
    else {
        $LicenseOption = Set-LicenseOption -SKU $SKU	
    }

    # "Zero" out the user counter
    [int]$i = 1
    [int]$ErrorCount = 0
	
    # Add License to the users passed in
    Foreach ($Account in $Users) {

        Write-Log ("==== Processing User " + $account.UserPrincipalName + " ====")

        # Get the MSOL User object
        $MSOLUser = $Null
        $MSOLUser = Get-MSOLUser -UserPrincipalName $account.UserPrincipalName -ErrorAction SilentlyContinue

        # Test the object / output to validate we have what we need
        if ($null -eq $MSOLUser){
            Write-Log ("[ERROR] - MSOL user not found " + $Account.UserPrincipalName)
            Write-Error "Unable to find provided user"
            $ErrorCount++
        }
        # Check to see if the sku we are trying to assign is already on the user is so break out of the foreach
        elseif (!($null -eq (($MSOLUser).licenses | Where-Object { $_.accountskuid -eq $SKU }))) {
            Write-Log ("[WARNING] - " + $SKU + " is already assigned to the user.")
            Write-Warning "User already has $SKU assigned. Please use Set-MSOLUserLicensePlans or Update-MSOLUserLicensePlans to modify existing Plans."
        }
        else {
            # If location has a value then we need to set it
            if (!([string]::IsNullOrEmpty($Location))) {
                $command = ("Set-MsolUser -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -UsageLocation " + $Location)
                Write-Log ("Running: " + $Command)
                Invoke-Expression $Command
            }

            # Build and run our license set command
            [string]$Command = ("Set-MsolUserLicense -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -AddLicenses " + $SKU + " -LicenseOptions `$LicenseOption -ErrorAction Stop -ErrorVariable CatchError")
            Write-Log ("Running: " + $Command)
            
            # Try our command
            try { 
                Invoke-Expression $Command 
            }
            # Doing this so I can customize the error later
            catch { 
                Write-Log ("[ERROR] - " + $CatchError.ErrorRecord)
                Write-Log ("[ERROR] - " + $Account.UserPrincipalName + " Failed to add SKU " + $SKU)
                Write-Error ("Failed to successfully add license to user " + $account.UserPrincipalName)
                $ErrorCount++
            }
        }        

        # Update the progress bar and increment our counter
        Update-Progress -CurrentCount $i -MaxCount $Users.Count -Message ("Adding " + $SKU + "to Users.")

        # Update our user counter
        $i++
    }

    Write-Log ("Finished Adding SKU " + $SKU + " to " + $Users.Count + " Users.")
    If ($ErrorCount -gt 0) {
        Write-Log ($ErrorCount.ToString() + " ERRORS DURING PROCESSING PLEASE REVIEW ENTRIES WITH '[ERROR]' FOR MORE INFORMATION")
    }
}