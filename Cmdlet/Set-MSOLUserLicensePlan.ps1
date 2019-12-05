# Overwrites plans on the specified SKU
Function Set-MSOLUserLicensePlan {
    
    <#
 
	.SYNOPSIS
    SET the current Enabled or Disabled Plans for a User.
    OVERWRITES the current values.

	.DESCRIPTION
	OVERWRITES the current plan settings for an assigned SKU with the new plan settings provided.

    * OVERWRITES existing plan settings and makes them match the provided values
    * Use Update-MSOLUserLicensePlan to modify existing plan configurations
	* Can specify plans to enable or disable
	* Logs all activity to a log file
	
	.PARAMETER Users
	Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.

	.PARAMETER SKU
	SKU that should be added.

	.PARAMETER PlansToDisable
	Comma seperated list of SKU plans to Disable.

	.PARAMETER PlansToEnable
	Comma seperated list of SKU plans to Enable.

	.PARAMETER LogFile
	File to log all actions taken by the function.

	.OUTPUTS
	Log file showing all actions taken by the function.

	.EXAMPLE
	Set-MSOLUserLicensePlan -Users $Promoted -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK
	
	Enable all of the plans in the ENTERPRISEPACK SKU for all users in $Promoted.  Overwriting any current settings.

	.EXAMPLE
	Set-MSOLUserLicensePlan -Users $Promoted -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK -PlanstoEnable SWAY,TEAMS1,EXCHANGE_S_ENTERPRISE,YAMMER_ENTERPRISE,OFFICESUBSCRIPTION

	Set the enabled Plans for the ENTERPRISEPACK to be SWAY,TEAMS1,EXCHANGE_S_ENTERPRISE,YAMMER_ENTERPRISE,OFFICESUBSCRIPTION for all users in $Promoted.  This will overwrite any current plan settings for the user.
	
    #>
    
    Param
    (
        [Parameter(Mandatory = $true)]
        [array]$Users,
        [Parameter(Mandatory = $true)]
        [string]$LogFile,
        [string]$SKU,
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
        $SKU = Select-SKU -Message "Select SKU to Set:"
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
	
    # Set License on the users passed in
    Foreach ($Account in $Users) {

        Write-Log ("==== Processing User " + $account.UserPrincipalName + " ====")
        
        # Build and run our license set command
        [string]$Command = ("Set-MsolUserLicense -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -LicenseOptions `$LicenseOption -ErrorAction Stop -ErrorVariable CatchError")
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
            Write-Error ("Failed to successfully add license to user " + $account.UserPrincipalName)
            $ErrorCount++
        }
		

        # Update the progress bar and increment our counter
        Update-Progress -CurrentCount $i -MaxCount $Users.Count -Message "Setting License Plan"

        # Update our user counter
        $i++
    }
    
    Write-Log ("Finished Adding SKU " + $SKU + " to " + $Users.Count + " Users.")
    If ($ErrorCount -gt 0) {
        Write-Log ($ErrorCount.ToString() + " ERRORS DURING PROCESSING PLEASE REVIEW ENTRIES WITH '[ERROR]' FOR MORE INFORMATION")
    }
}