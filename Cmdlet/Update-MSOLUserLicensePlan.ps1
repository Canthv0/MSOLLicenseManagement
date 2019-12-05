Function Update-MSOLUserLicensePlan {
    <#
 
	.SYNOPSIS
	UPDATES the current plan settings for an assigned SKU while maintaining existing plan settings.

	.DESCRIPTION
	Updates the current plan settings for an assigned SKU while maintaining existing plan settings.

    * UPDATES existing plan settings with the new enabled or disabled plans
    * All current plan setting are left in place.
    * To change plans without maininting the current state use Set-MSOLUserLicensePlan
	* Can specify plans to enable or disable
	* Logs all activity to a log file
	
	.PARAMETER Users
	Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.

	.PARAMETER SKU
	SKU that should be modified.

	.PARAMETER PlansToDisable
	Comma seperated list of SKU plans to Disable.

	.PARAMETER PlansToEnable
	Comma seperated list of SKU plans to Enable.

	.PARAMETER LogFile
	File to log all actions taken by the function.

	.OUTPUTS
	Log file showing all actions taken by the function.

	.EXAMPLE
	Update-MSOLUserLicensePlan -Users $Promoted -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK -PlanstoEnable SWAY,TEAMS1,EXCHANGE_S_ENTERPRISE,YAMMER_ENTERPRISE,OFFICESUBSCRIPTION

	Adds the enabled Plans SWAY,TEAMS1,EXCHANGE_S_ENTERPRISE,YAMMER_ENTERPRISE,OFFICESUBSCRIPTION to the ENTERPRISEPACK for all users in $Promoted. Maintaining any existing plan settings.
	
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
    
    # Takes in a list of plans and creates a new list based currently disabled plans + list
    Function Update-DisabledPlan {

        param 
        (
            [string]$SKU,
            [string[]]$PlansToDisable,
            $MSOLUser
        )

        Write-Log "Determining Plans to Disable"

        # Null out our array
        [array]$CurrentDisabledPlans = $null
	
        # Get the License object
        $license = $MSOLUser.licenses | Where-Object { $_.accountskuid -eq $SKU }
	
        # Make sure we found the license on the user
        if ($null -eq $license) {
            Write-Log ("[ERROR] - Cannot find SKU " + $SKU + " assigned to " + $MSOLUser.UserPrincipalName)
            Write-Error -Message ("Cannot find SKU " + $SKU + " assigned to " + $MSOLUser.UserPrincipalName)
            
            $Output = "Error"
            Return $Output
        }
        else {

            # Get currently disabled plans
            [array]$CurrentDisabledPlans = ($license.servicestatus | Where-Object { $_.provisioningstatus -eq "disabled" })

            # Make sure we got back some disabled plans
            if ($CurrentDisabledPlans.Count -le 0) {
		
                Write-Log "No Currently Disabled Plans."

                # Just return the value of -planstodisable that was submitted
                [string[]]$Output = $PlansToDisable.ToUpper()
		
                Return $Output

            }
            # There are currently disabled plans we need to combine the current with the new
            else {				

                # Show the currently Disabled Plans and then determine the new list
                foreach ($plan in $CurrentDisabledPlans) {
                    [string[]]$CurrentDisabledPlansNames = $CurrentDisabledPlansNames + ($plan.serviceplan.servicename).toupper()
                }

                Write-Log ("Currently Disabled plans:" + $CurrentDisabledPlansNames)
	
                # Combine the two lists of disabled plans
                [string[]]$Output = $CurrentDisabledPlansNames + $PlansToDisable

                # Make all of them uppercase for comparison
                $Output = $Output.toupper()

                # Throw out all of the duplicates
                $Output = ($Output | Select-Object -Unique)

                Return $Output
            }
        }
    }

    # Takes in a list of plans and creates a new list based on currently enabled plans + list
    Function Update-EnabledPlan {
        param 
        (
            [string]$SKU,
            [string[]]$PlansToEnable,
            $MSOLUser
        )

        Write-Log "Determining Plans to Enable"

        # Null out our values
        [string[]]$CurrentDisabledPlansNames = $null
        [array]$CurrentDisabledPlans = $null
	
        # Get the License object
        $license = $MSOLUser.licenses | Where-Object { $_.accountskuid -eq $SKU }
	
        # Make sure we found the license on the user
        if ($null -eq $license) {
            Write-Log ("[ERROR] - Cannot find SKU " + $SKU + " assigned to " + $MSOLUser.UserPrincipalName)
            Write-Error -Message ("Cannot find SKU " + $SKU + " assigned to " + $MSOLUser.UserPrincipalName)

            $Output = "Error"
            Return $Output
        }
        # If we do process it
        else {
            
            # Get currently disabled plans
            [array]$CurrentDisabledPlans = ($license.servicestatus | Where-Object { $_.provisioningstatus -eq "disabled" })
	
            # If there are no currently disabled plans then return null since there are no plans to update
            if ($null -eq $CurrentDisabledPlans) {
                Write-Log "No Currently Disabled Plans; No Plan updates needed."
                $Output = "Enabled"
                Return $Output
            }
            # Otherwise we need to show what is currently disabled and then calculate the new list to disable
            else {
                # Pull out the plan names
                foreach ($plan in $CurrentDisabledPlans) {
                    [string[]]$CurrentDisabledPlansNames = $CurrentDisabledPlansNames + ($plan.serviceplan.servicename).toupper()
                }

                Write-Log ("Currently Disabled plans:" + $CurrentDisabledPlansNames) 
	
                # Go thru each plan that needs to be enabled and remove it from the list of currently disabled plans
                foreach ($Plan in $PlansToEnable) {		
                    # Remove the plans we want to enable from the list of disabled plans
                    $CurrentDisabledPlansNames = ($CurrentDisabledPlansNames | Where-Object { $_ -ne $Plan })			
                }	

                # Make sure we havn't pulled out all disabled plans and if we have just return a null
                if ($null -eq $CurrentDisabledPlansNames) {
                    $Output = "Enabled"
                    Return $Output
                }
                # As long as we have at least one return that one
                else {
                    Return $CurrentDisabledPlansNames
                }
            }
        }
    }

    ### MAIN ###

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
        $SKU = Select-SKU -Message "Select SKU to be updated:"
    }
    # If a value has been passed in verify it
    else {
        $SKU = Select-SKU -SKUToCheck $SKU
    }

    # Testing the plan inputs to make sure they are valid
    if (!([string]::IsNullOrEmpty($PlansToDisable))) {
        Test-Plan -Sku $SKU -Plan $PlansToDisable
    }
    # If plans to enable has a value then we test them
    elseif (!([string]::IsNullOrEmpty($PlansToEnable))) {
        Test-Plan -Sku $SKU -Plan $PlansToEnable
    }
    # If neither has been provided then we don't need to do anything
    else { }

    # "Zero" out the user counter
    [int]$i = 1
    [int]$ErrorCount = 0
	
    # Update the Plans for the Users Passed in
    Foreach ($Account in $Users) {

        # Ensure that our MSOLUser object is Null at the start of each loop
        $MSOLUser = $null

        # For if we determine we need to skip processing the user
        $Skip = $false

        Write-Log ("==== Processing User " + $account.UserPrincipalName + " ====")

        # Get our user object
        Try {
            $MSOLUser = Get-MsolUser -UserPrincipalName $account.UserPrincipalName -ErrorAction Stop
        }
        Catch {
            Write-Log ("[ERROR] - Unable to Find User" + $Account.UserPrincipalName)
            Write-Error ("Unable to Find User" + $Account.UserPrincipalName) -ErrorAction Continue
            [string[]]$CalculatedPlansToDisable = "Error"
            continue
        }
        
        # Check that the user object we found contains the SKU we are looking to update
        if (!($MSOLUser.Licenses.AccountSkuId -contains $SKU)) {
            Write-log "[ERROR] - Unable to verify the SKU to update is assigned"
            Write-Log ("Assigned SKUs: " + $MSOLUser.Licenses.AccountSkuId)
            Write-Error "Unable to verify SKU is Assigned to User" -ErrorAction Continue
            [string[]]$CalculatedPlansToDisable = "Error"    
        }
        # If -planstodisable is set then read it in and set license options
        elseif (!([string]::IsNullOrEmpty($PlansToDisable))) {		
            # Get the license options
            [string[]]$CalculatedPlansToDisable = Update-DisabledPlan -SKU $SKU -Plan $PlansToDisable -MSOLUser $MSOLUser
        }
        # If -planstoenable has a value then we need to determine what plans to enable and set them
        elseif (!([string]::IsNullOrEmpty($PlansToEnable))) {
            # Get the disabled plans and License options
            [string[]]$CalculatedPlansToDisable = Update-EnabledPlan -SKU $SKU -Plan $PlansToEnable -MSOLUser $MSOLUser
        }
        # If Neither disable or enable were passed then we turn on all plans
        else {
            Write-Log ("Turning on all Plans for " + $Sku + " on user " + $Account.UserPrincipalName)
            [string[]]$CalculatedPlansToDisable = "Enabled"
        }

        # Set the License Options based ont he value of $CalculatedPlansToDisable

        switch ($CalculatedPlansToDisable[0]) {
            Error {
                Write-Log ("[ERROR] - Skipping User")
                $Skip = $true
                $ErrorCount++
            }
            Enabled {
                Write-Log ("Setting License Options to all Enabled")
                $LicenseOption = Set-LicenseOption -SKU $SKU
            }
            Default {
                Write-Log ("Setting License options for Plans")
                $LicenseOption = Set-LicenseOption -DisabledPlansArray $CalculatedPlansToDisable -SKU $SKU
            }
        }
		
        # If we determined no actions are needed then skip the user
        if ($Skip) { Write-Log "Skipping user" }
        # Process the user
        else {

            # Build and run our license set command
            [string]$Command = ("Set-MsolUserLicense -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -LicenseOptions `$LicenseOption -ErrorAction Stop -ErrorVariable CatchError")
            Write-Log ("Running: " + $Command)
            # Try our command
            try { 
                Invoke-Expression $Command 
            }
            # If we have any error write out and stop
            catch { 
                Write-Log ("[ERROR] - " + $CatchError.ErrorRecord)
                Write-Error ("Failed to successfully add license to user " + $account.UserPrincipalName)
                $ErrorCount++
            }
        }	

        # Update the progress bar and increment our counter
        Update-Progress -CurrentCount $i -MaxCount $Users.Count -Message "Updating License Plan"

        # Update our user counter
        $i++
    }

    Write-Log ("Finished Adding SKU " + $SKU + " to " + $Users.Count + " Users.")
    If ($ErrorCount -gt 0) {
        Write-Log ($ErrorCount.ToString() + " ERRORS DURING PROCESSING PLEASE REVIEW ENTRIES WITH '[ERROR]' FOR MORE INFORMATION")
    }
}