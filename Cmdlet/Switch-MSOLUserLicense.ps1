#############################################################################################
# DISCLAIMER:																				#
#																							#
# THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT					#
# PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY				#
# OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT		#
# LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR		#
# PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS		#
# AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR			#
# ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE	#
# FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS	#
# PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)	#
# ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,		#
# EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES						#
#############################################################################################
# Changes from one SKU to another SKU for a collection of users
Function Switch-MSOLUserLicense {
    <#

 	.SYNOPSIS
	Switches a user from one SKU to another SKU

	.DESCRIPTION
	Replaces one SKU with another SKU on a collection of users.

	* Prompts for SKUs if none are provided
	* Allows for Disabling or Enabling Plans on the new SKU
	* -AttemptToMaintainPlans will try to keep the current plan setting for the user on the new SKU

	.PARAMETER Users
	Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.

	.PARAMETER SKU
	SKU that should be added.

	.PARAMETER SKUToReplace
	SKU that will be removed.

	.PARAMETER Location
	If provided will set the location of the user.

	.PARAMETER PlansToDisable
	Comma seperated list of SKU plans to Disable.

	.PARAMETER PlansToEnable
	Comma seperated list of SKU plans to Enable.

	.PARAMETER AttemptToMaintainPlans
	Tries to keep the same plan states on the new SKU that is being assigned.  For plans that are unique
	to the new SKU it will default to enabling them.

	.PARAMETER LogFile
	File to log all actions taken by the function.

	.OUTPUTS
	Log file showing all actions taken by the function.

	.EXAMPLE
	Switch-MSOLUserLicense -Users $NewUsers -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK -SKUToReplace company:STANDARDPACK
	
	Replaces the STANDARDPACK with the ENTERPRISEPACK and enables all plans.

	.EXAMPLE
	Add-MSOLUserLicense -Users $NewUsers -Logfile C:\temp\add_license.log -SKU company:ENTERPRISEPACK -SKUTOReplace company:STANDARDPACK -PlanstoDisable Deskless,Sway,Teams1 

	Replaces the STANDARDPACK with the ENTERPRISEPACK and disables Deskless, Sway, and Teams.
		
	#>	
	
    Param
    (
        [Parameter(Mandatory = $true)]
        [array]$Users,
        [Parameter(Mandatory = $true)]
        [string]$LogFile,
        [string]$SKU,
        [string]$SKUToReplace,
        [string[]]$PlansToDisable,
        [string[]]$PlansToEnable,
        [switch]$AttemptToMaintainPlans
    )
	
    # Takes in a list of plans and creates a new list based on currently enabled plans + list
    Function Get-PlansToMaintain {
        param 
        (
            [Parameter(Mandatory = $true)]
            [string]$SKU,
            [Parameter(Mandatory = $true)]
            [string]$SKUToReplace,
            [Parameter(Mandatory = $true)]
            $MSOLUser
        )

        Write-Log "Determining Plans to Maintain"

        # Null out our values
        [string[]]$CurrentDisabledPlansNames = $null
        [array]$CurrentDisabledPlans = $null
	
        # Get the Plan names for our new SKU
        [string[]]$NewPlans = ((Get-MsolAccountSku | Where-Object { $_.accountskuid -eq $SKU }).servicestatus.serviceplan.servicename).toupper()

        # Get the License object
        $license = $MSOLUser.licenses | Where-Object { $_.accountskuid -eq $SKUToReplace }
	
        # Get currently disabled plans
        [array]$CurrentDisabledPlans = ($license.servicestatus | Where-Object { $_.provisioningstatus -eq "disabled" })
	
        # If there are no currently disabled plans then return null since there are no plans to update
        if ($null -eq $CurrentDisabledPlans) {
            Write-Log "No Currently Disabled Plans; All Plans will be enabled."
            $Output = "Enabled"
            Return $Output
        }
        # Otherwise we need to show what is currently disabled and then calculate the new list to disable
        else {
			
            # Pull out the plan names
            foreach ($plan in $CurrentDisabledPlans) {
                [string[]]$CurrentDisabledPlansNames = $CurrentDisabledPlansNames + ($plan.serviceplan.servicename).toupper()
            }

            Write-Log ("Plans disabled in current SKU: " + $CurrentDisabledPlansNames)

            # Go thru each plan that needs to be disabled and if it is there remove it from the list of plans in the new SKU
            foreach ($Plan in $CurrentDisabledPlansNames) {			
                # Builds a list of plans to enable from the new SKU
                $NewPlans = ($NewPlans | Where-Object { $_ -ne $Plan })				
            }	

            # If we didn't get back any plans to enable then something went wrong and we need to error out
            if ($null -eq $NewPlans) {
                Write-Log "[ERROR] - All plans are on the new SKU are set to be disabled"
                Write-Log "[ERROR] - Enabling Plans instead."
                Write-Error "All plans will be Enabled."
                $Output = "Enabled"
                Return $Output
            }
            # Take back our list of plans that need to be enabled and get back a list of plans to disable
            else {
                [string[]]$Output = Set-EnabledPlan -SKU $SKU -PlansToEnable $NewPlans
                Return $Output
            }
        }
    }
	
    ### Main ####

    # Make sure we didn't get both enable and disable
    if (!([string]::IsNullOrEmpty($PlansToDisable)) -and !([string]::IsNullOrEmpty($PlansToEnable))) {
        Write-Log "[ERROR] - Cannot use both -PlansToDisable and -PlansToEnable at the same time"
        Write-Error "Cannot use both -PlansToDisable and -PlansToEnable at the same time" -ErrorAction Stop
    }
    elseif ((!([string]::IsNullOrEmpty($PlansToDisable)) -or !([string]::IsNullOrEmpty($PlansToEnable))) -and $AttemptToMaintainPlans){
        Write-Log "[ERROR] - Cannot use -AttemptToMaintainPlans with -PlansToDisable or -PlansToEnable"
        Write-Error "Cannot use -AttemptToMaintainPlans with -PlansToDisable or -PlansToEnable" -ErrorAction Stop
    }

    # Make sure we have a valid log file path
    Test-LogPath -LogFile $LogFile

    # Make sure we have the connection to MSOL
    Test-MSOLServiceConnection

    # Make user our Users object is valid
    [array]$Users = Test-UserObject -ToTest $Users

    # If no value of SKU was passed in then call Select-Sku to allow one to be picked
    if ([string]::IsNullOrEmpty($SKU)) {
        $SKU = Select-SKU -Message "Select New SKU For Users:"
    }
    # If a value has been passed in verify it
    else {
        $SKU = Select-SKU -SKUToCheck $SKU
    }
	
    # If no value of SKUToReplace was passed in then call Select-Sku to allow one to be picked
    if ([string]::IsNullOrEmpty($SKUToReplace)) {
        $SKUToReplace = Select-SKU -Message "Select SKU to be replaced"
    }
    # If a value has been passed in verify it
    else {
        $SKUToReplace = Select-SKU -SKUToCheck $SKUToReplace
    }

    ## Make sure skutoreplace and sku don't match
    if ($SKUToReplace.toupper() -eq $SKU.ToUpper()) {
        Write-Log "[ERROR] - `$SKU and `$SKUToReplace match.  Unable to replace license with itself."
        Write-Log "[ERROR] - Please use Update-MSOLUserLicensePlan or Set-MSOLUserLicensePlan to modify license plans."
        Write-Error -Message "-SKU and -SKUToReplace can not have the same value." -ErrorAction Stop
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
    elseif ($AttemptToMaintainPlans) {
        Write-Log "Will attempt to maintain existing user plans when moving to new License"
		
        $ExistingPlans = (Get-MsolAccountSku | Where-Object { $_.accountskuid -eq $SKUToReplace }).servicestatus.serviceplan.servicename
        $NewPlans = (Get-MsolAccountSku | Where-Object { $_.accountskuid -eq $SKU }).servicestatus.serviceplan.servicename

        foreach ($Plan in $NewPlans) {
            # If the new plan name isn't in the existing list of plans then we will end up enabling it
            if ($ExistingPlans -contains $Plan) { }
            else {
                # Add to our list of plans we will be enabling by default
                [string[]]$PlansEnabledByDefault = $PlansEnabledByDefault + $Plan
            }			
        }

        # If we generated at least one plan that didn't match between the two SKUs then inform the user and get consent to continue
        if ($PlansEnabledByDefault.Count -gt 0) {
            Write-Log "Cannot match the following plans between the two SKUs so they will be enabled by default:"
            Write-Log $PlansEnabledByDefault

            # Prompt the user to upgrade or not
            $title = "Agree to Enable"
            $message = "When the SKU switch is complete the plans listed above will be enabled by default. `nIs this OK?"
            $Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Continues Switching SKUs enabling the listed plans."
            $No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Stops the function."
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($Yes, $No)
            $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

            # Check to see what the user choose
            switch ($result) {
				
                0 {
                    Write-Log "Agreed to enabled listed plans by default."
                }
                1 {
                    Write-Log "[ERROR] - Did not accept default plan enablement."
                    Write-Error "User terminated function.  Unable to enable default plans." -ErrorAction Stop
                }
            }
        }
    }
    # If none are provided then we will enable all plans
    else {
        Write-Log "Enabling all plans"
        $LicenseOption = Set-LicenseOption -SKU $SKU
    }

    # "Zero" out the user counter
    [int]$i = 1
    [int]$ErrorCount = 0

    foreach ($Account in $Users) {
        Write-Log ("==== Processing User " + $account.UserPrincipalName + " ====")
		
        # Null out our reused variables
        $CalculatedPlansToDisable = $null
		
        # Get the current user object so we can pass it in once and not ask for it multiple times
        $MSOLUser = Get-MsolUser -UserPrincipalName $Account.UserPrincipalName
	
        # Check to see if the sku we are trying to assign is already on the user is so break out of the foreach
        if (!($null -eq (($MSOLUser).licenses | Where-Object { $_.accountskuid -eq $SKU }))) {
            Write-Log ("[WARNING] - " + $SKU + " is already assigned to the user.")
            Write-Warning "User already has $SKU assigned."
        }
        # Make sure we have the SKU we are replacing assigned
        elseif ($null -eq (($MSOLUser).licenses | Where-Object { $_.accountskuid -eq $SKUToReplace })) {
            Write-Log ("[WARNING] - " + $SKUToReplace + " is not assigned to the user.")
            Write-Warning "User doesn't have $SKUToReplace Assigned."
        }
        else {

            # Since attempt to maintain is per user we need to calculate per user license options
            if ($AttemptToMaintainPlans) {

                # Get the disabled plans and License options
                [string[]]$CalculatedPlansToDisable = Get-PlansToMaintain -SKU $SKU -SKUToReplace $SKUToReplace -MSOLUser $MSOLUser
				

                # If we found no disabled plans then turn on everything
                if ($CalculatedPlansToDisable[0] -eq "Enabled") {
                    Write-Log ("Turning on all Plans for " + $Sku + " on user " + $Account.UserPrincipalName)                    	
                }
                # Turn on only the plans we found that need to be disable in the new SKU
                else {
                    $LicenseOption = Set-LicenseOption -DisabledPlansArray $CalculatedPlansToDisable -SKU $SKU
                }			
            }
            # If we are not trying to maintain plans we can use the license options that were calculated before the foreach
            else { }

            [string]$Command = ("Set-MsolUserLicense -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -AddLicenses " + $SKU + " -RemoveLicenses " + $SKUToReplace + " -LicenseOptions `$LicenseOption -ErrorAction Stop -ErrorVariable CatchError")
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
                Write-Error ("Failed to successfully switch licenses for user " + $account.UserPrincipalName)
                $ErrorCount++
            }	
        }

        # Update the progress bar and increment our counter
        Update-Progress -CurrentCount $i -MaxCount $Users.Count -Message "Switching Licenses"

        # Update our user counter
        $i++
    }

    Write-Log ("Finished Swapping SKU " + $SKUToReplace + " to " + $SKU + " for " + $Users.count + " Users.")
    If ($ErrorCount -gt 0) {
        Write-Log ($ErrorCount.ToString() + " ERRORS DURING PROCESSING PLEASE REVIEW ENTRIES WITH '[ERROR]' FOR MORE INFORMATION")
    }
}