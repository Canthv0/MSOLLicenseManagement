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
# Converts from an inherited SKU to an Explicit SKU with the same plan settings
Function Convert-MSOLUserLicenseToExplicit {
    <#
 
	.SYNOPSIS
	Converts a specified SKU from Group assiged to Explicitly assigned

	.DESCRIPTION
	Will explicitly apply the SKU specified to a user if that user is inheriting it from Group based licensing.
	https://docs.microsoft.com/en-us/azure/active-directory/active-directory-licensing-whatis-azure-portal

	* Will skip any users that don't have the specified SKU
	* Will skip any users that already have the SKU explicitly assigned
	* Will maintain the current plan state for the SKU on that user
	
	.PARAMETER Users
	Single UserPrincipalName, Comma seperated List, or Array of objects with UserPrincipalName property.

	.PARAMETER SKU
	SKU that should be converted to explict from inherited.

	.PARAMETER LogFile
	File to log all actions taken by the function.

	.OUTPUTS
	Log file showing all actions taken by the function.

	.EXAMPLE
	Convert-MSOLUserLicenseToExplicit -Users $UsersToConvert -SKU company:AAD_PREMIUM -logfile C:\temp\conversion.log
	
	Converts inherited AAD_PREMIUM licenses into explicit licenses for all users in $UsersToConvert and logs all actions in C:\temp\conversion.log
	
	#>
	
    param 
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
        $SKU = Select-SKU -Message "Select SKU to Convert to Explicit:"
    }
    # If a value has been passed in verify it
    else {
        $SKU = Select-SKU -SKUToCheck $SKU
    }	
	
    # Start Processing
    Write-Log ("Converting " + $SKU + " from inherited to explicit")

    # "Zero" out of counter
    [int]$i = 1
    [int]$ErrorCount = 0

	
    Foreach ($Account in $Users) {

        Write-Log ("==== Processing User " + $account.UserPrincipalName + " ====")

        # Set our skip variable to false so we process the user
        # Will change it to true if we need to skip the actual change
        $Skip = $false		
		
        # Null out the variables we use
        [array]$CurrentDisabledPlansArray = $null
        $CurrentLicenseDetails = $null
        $MSOLUser = $null

        # Verify that the SKU is inherited and isn't already explicit ... don't want to overwrite any explicit settings
        try {
            # Need stop here to ensure that the error is caught in the try/catch
            $MSOLUser = Get-MsolUser -UserPrincipalName $Account.UserPrincipalName -ErrorAction Stop
        }
        catch {
            Write-Log ("[Error] - Unable to find user:`n " + $_.Exception)
            Write-Error ("Problem finding user " + $Account.UserPrincipalName)
            $ErrorCount++
            $Skip = $true
            $MSOLUser = $null
        }

        # If the user object is null then we don't need to / can't get its license information
        if ($null -eq $MSOLUser) { }
        # Pull the license details and validate them
        else {
            $CurrentLicenseDetails = ($MSOLUser.licenses | Where-Object { $_.accountskuid -eq $SKU })

            # Make sure we found the SKU on the user
            if ($Null -eq $CurrentLicenseDetails) {
                Write-Log ("[Warning] - SKU " + $SKU + " not Assigned to user " + $Account.UserPrincipalName)
                Write-Warning ("SKU " + $SKU + " not Assigned to user " + $Account.UserPrincipalName)
				
                # Since we can't set this we need to skip it
                $Skip = $true
            }
            # Make sure the SKU is inherited
            elseif (($CurrentLicenseDetails.GroupsAssigningLicense -contains $Users.objectid) -or ($CurrentLicenseDetails.GroupsAssigningLicense.count -le 0)) {
                Write-Log ("[Warning] - User " + $Account.UserPrincipalName + " already has explicit assignment for " + $SKU)
                Write-Warning ("User " + $Account.UserPrincipalName + " already has explicit assignment for " + $SKU)
				
                # Since it is already explict we need to skip the actual set
                $Skip = $true
				
            }
            else {
								
                # Get the currently disabled plans
                [array]$CurrentDisabledPlansArray = (($CurrentLicenseDetails).ServiceStatus | Where-Object { $_.ProvisioningStatus -eq "Disabled" }).serviceplan | Select-Object -ExpandProperty ServiceName
                Write-Log ("Disabling Plans: " + [string]$CurrentDisabledPlansArray)
            }
        }	

        # If we set SKIP to true then we are not processing this user log it and move on
        if ($Skip -eq $true) {
            Write-Log ("Skipping user " + $Account.UserPrincipalName)
        }
        # Else we need to set the license
        else {
            # Create our License Options from the disabled plan array
            $LicenseOptions = Set-LicenseOption -DisabledPlansArray $CurrentDisabledPlansArray -SKU $SKU

            # Build, log, and run the command
            $cmd = "Set-MsolUserLicense -UserPrincipalName `"" + $Account.UserPrincipalName + "`" -AddLicenses " + $SKU + " -LicenseOptions `$LicenseOptions"
            Write-Log ("Running: " + $cmd)
            Invoke-Expression $cmd
        }

        # Update the progress bar and increment our counter
        Update-Progress -CurrentCount $i -MaxCount $Users.Count -Message "Coverting from Inherited to Explicit"
        $i++
    }
	
    Write-Log ("Finished Adding SKU " + $SKU + " to " + $Users.Count + " Users.")
    If ($ErrorCount -gt 0) {
        Write-Log ($ErrorCount.ToString() + " ERRORS DURING PROCESSING PLEASE REVIEW ENTRIES WITH '[ERROR]' FOR MORE INFORMATION")
    }

    
}
