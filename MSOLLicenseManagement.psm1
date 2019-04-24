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

#==========================
# Utility Functions
#==========================
# Writes output to a log file with a time date stamp
Function Write-Log {
    Param ([string]$string)
	
    # Get the current date
    [string]$date = Get-Date -Format G
    # Write everything to our log file
    ( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	
    # If NonInteractive true then suppress host output
    if (!($NonInteractive)) {
        ( "[" + $date + "] - " + $string) | Write-Host
    }
}

# Make sure that we are connected to the MSOL Server if not connect us
#TODO: Need to move from $error.clear() to try-catch
Function Test-MSOLServiceConnection {
    # Set the Error Action Prefernce to stop
    $ErrorActionPreference = "Stop"

    # Make sure that the MSOL Module is installed
    if ($null -eq (Get-Module -ListAvailable MSOnline)) {
        Write-Error "MSOL Module Not installed.  Please install from https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell" -ErrorAction Stop
    }

    # Make sure we are connected to the MSOL service
    $error.clear()
    
    # Call Get-msolaccountsku if we are not connected we will get an error if we are we will not
    $null = Get-MSOLAccountSKU -ErrorVariable error -ErrorAction SilentlyContinue

    # Check to see if we threw an error
    if ($error.count -gt 0) {
        # If we have the expected error for not being connected the call connect-msolservice
        if ($error.Exception -like "*You must call the Connect-MsolService cmdlet*") {
            Write-Log "Not Connected to MSOLService calling Connect-MSOLService"
            import-module MSOnline
            Connect-MsolService
        }
        # If we get any other error then throw and error and stop
        else {
            Write-log ("Unexpected Error encountered" + $error)
            Write-Error "Unexpected Error stopping execution" -ErrorAction Stop
        }
    }
    else {
        # Do nothing because we are connected
    }
}

# Updates a progress bar
Function Update-Progress {
    Param 
    (
        [Parameter(Mandatory = $true)]
        [int]$CurrentCount,
        [Parameter(Mandatory = $true)]
        [int]$MaxCount,
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    # If currentcount = maxcount we are done so we need to kill the progress bar
    if ($CurrentCount -ge $MaxCount) {
        Write-Progress -Completed -Activity "User Modification"
    }
    # Every 25 update the progress bar
    elseif ($CurrentCount % 25 -eq 0) {

        [int]$Percent = (($CurrentCount / $MaxCount) * 100)
        [string]$Operation = ([string]$Percent + "% " + $Message)
        Write-Progress -Activity "User Modification" -CurrentOperation $Operation -PercentComplete $Percent
    }
    else {
        # Nothing to do if not divisible by 25
    }
}

# Test the logfile path to ensure the dir exists and that a file was provided
Function Test-LogPath {
    param (
        # Log File e.g. c:\temp\out.log
        [Parameter(Mandatory)]
        [string]
        $LogFile
    )

    #Get the parent of the provided path
    $parent = split-path $LogFile -Parent

    # If the parent is null then the user provided the root of a drive for the log file
    if ([string]::IsNullOrEmpty($parent)) {
        Write-Error -Message "Log File path appears to be a drive root. Please provide a folder and file name with -logfile e.g. C:\Temp\out.log" -ErrorAction Stop
        break
    }

    # Verify that the parent folder exists since we won't create it if it doesn't
    if (!(test-path $parent)) {
        Write-Error -Message ("Path not found " + $parent + " ;please provide a valid path.") -ErrorAction Stop
        break
    }

    # Check to make sure the path contains a file with an extension
    if ((($LogFile).split('\'))[-1] -notmatch '\.') {
        Write-Error -Message ("Log file path does not appear to contain a file extension. Please provide a path and file name e.g. c:\temp\out.log")
        break
    }
	
    # Check to see if the value passed in a container (folder) if so we need to abort
    if (test-path $LogFile -PathType Container) {
        Write-Error  -Message "Log File Path appears to be a folder. Please provide a file name with -logfile e.g. C:\Temp\out.log" -ErrorAction Stop
        break
    }
}

# Determine if we have an array with UPNs or just a single UPN / UPN array unlabeled
Function Test-UserObject {
    param ([array]$ToTest)

    # See if we can get the UserPrincipalName property off of the input object
    # If we can't get the property then test the input and see if we can generate an acceptable object array
    if ($null -eq $ToTest[0].UserPrincipalName) {
        # Very basic check to see if this is a UPN
        # This deals with the input user@company.com,user2@company.com ....
        
        if ($ToTest[0] -match '@') {
            [array]$Output = $ToTest | Select-Object -Property @{Name = "UserPrincipalName"; Expression = { $_ } }
            Return $Output
        }
        # If we couldn't find the UserPrincipalName and we didn't see an @ sign then we need to abort
        else {
            Write-Log "[ERROR] - Unable to determine if input is a UserPrincipalName"
            Write-Log "Please provide a UPN or array of objects with the property UserPrincipalName populated"
            Write-Error "Unable to determine if input is a User Principal Name" -ErrorAction Stop
        }
    }
    # If we can pull the value of UserPrincipalName then just return the same object back
    else {
        Return $ToTest
    }
}

# Convert SkuID into Common Name
Function Get-SKUCommonName {
    param 
    (
        [Parameter(Mandatory = $true)]
        [string]$Skuid,
        [Parameter(Mandatory = $true)]
        [array]$SkuArray
    )

    # Find the Common name from the SkuID
    [string]$CN = ($Skuarray | where-object { $_.skuid -eq $Skuid }).CommonName

    # If we can't find the SkuID then we set it to "Not found"
    if ([string]::IsNullOrEmpty($CN)) { [string]$CN = "Not Found" }
	
    Return $CN
}

#==========================
# Common SKU Related Functions
#==========================

# Validate we have provided a valid SKU or collect one if we have provided none
Function Select-SKU {
    param 
    (
        [string]$SKUToCheck,                
        [string]$Message
    )	
	
    # If we don't have a value for the SKU then we need to ask for one.
    if ([string]::IsNullOrEmpty($SKUToCheck)) {
		
        Write-Log ("No SKU value provided.  Please select from the availible account SKUs")
	
        # Setup a counter	
        $i = 1
		
        # Write out the provided message
        Write-Host "`n$Message`n"
		
		
        # Get all of the SKUs and display them with a "selection number"
        $TenantSKU = Get-MSOLAccountSKU
        ForEach ($sku in $TenantSKU) {	
            Write-Host ([string]$i + " - " + $sku.SKUPartNumber)	
            $i++
        }
		
        # Collect the selected item from the user
        [int]$Selected = 0
        While (($Selected -ge $i) -or ($Selected -le 0) -or ($null -eq $Selected)) {
            [int]$Selected = Read-Host ("`nSelect the Number for the SKU to use.")	
        }
		
        Return ($TenantSKU[$Selected - 1].AccountSkuID)
	
    }
    # We need to verify that the submitted SKU is valid
    else {
	
        Write-Log ("Verifying SKU " + $SKUToCheck)
		
        # If we can't find it then throw and error and abort
        if ($null -eq (Get-MSOLAccountSKU | Where-Object { $_.AccountSKUID -eq $SKUToCheck })) {
            Write-Log ("[ERROR] - Did not find SKU with ID: " + $SKUToCheck)
            Write-Error ("Unable to locate SKU " + $SKUToCheck) -ErrorAction Stop
        }
        # If we found it log it
        else {
            Write-Log ("Found SKU")
            Return $SKUToCheck.ToUpper()
        }
    }	
}

# sets list of enabled plans to provided list
Function Set-EnabledPlan {
    param
    (
        [string]$SKU,
        [string[]]$PlansToEnable
    )

    Write-Log "Determining plans to Disable"
    [string[]]$Output = (Get-MsolAccountSku | Where-Object { $_.accountskuid -eq $SKU }).servicestatus.serviceplan.servicename
	
    Foreach ($PlanToRemove in $PlansToEnable) {
        # Take out each of the plans that were provided to us
        $Output = $Output | Where-Object { $_ -ne $PlanToRemove }
    }

    Return $Output
}

# Creates the LicenseOptions variable for setting diabled plans
Function Set-LicenseOption {

    param 
    (
        [string[]]$DisabledPlansArray,
        [Parameter(Mandatory = $true)]
        [string]$SKU	
    )

    # if there are no plans to disable we create the default option set
    if ($null -eq $DisabledPlansArray) {
        Write-log "Setting all SKU Plans to Enabled"
        $licenseOptions = New-MsolLicenseOptions -AccountSKUId $SKU
    }
    # Otherwise add the disabled plans to the option set
    else {
        Write-Log "Setting Disabled Plans License Options"
        Write-Log ("Disabled Plans: " + $DisabledPlansArray)
        $licenseOptions = New-MsolLicenseOptions -AccountSKUId $SKU -DisabledPlans $DisabledPlansArray
    }
	
    Return $licenseOptions
}

# Tests a list of plans agiast availible plans in the SKU
# Return the first plan that fails or null
Function Test-Plan {
    param 
    (
        [string[]]$Plan,
        [string]$Sku
    )

    Write-Log "Validating Provided Plans"
    $PlansInSKU = (Get-MsolAccountSku | Where-Object { $_.accountskuid -eq $sku }).servicestatus.serviceplan.servicename

    # Take each of the provided plans and check it against the plans we found in the SKU
    foreach ($PlanToCheck in $Plan) {
        if ($PlansInSKU -contains $PlanToCheck) { }
        # If we don't find it then throw an error and stop
        else { 
            Write-Log ("[ERROR] - Failed to find plan: " + $PlanToCheck)
            Write-Error ("Invalid plan " + $PlanToCheck + " provided in plan list. Please correct input and try again.") -ErrorAction Stop
        }
    }

    # If we validate them all then log it
    Write-Log "All Plans Valid"
}