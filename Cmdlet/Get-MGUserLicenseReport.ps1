# Generates an MSOL User License Report

Function Get-MSOLUserLicenseReport {
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

	.PARAMETER LogFile
    File to log all actions taken by the function.
    
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
        [array]$Users,
        [Parameter(Mandatory)]
        [string]$LogFile,
        [switch]$OverWrite = $false,
    )
	
    # Make sure we have a valid log file path
    Test-LogPath -LogFile $LogFile

    # Make sure we have the connection to MSOL
    Test-MGServiceConnection
	
    Write-Log "Generating Sku and Plan Report"

    # Get all of the availible SKUs
    $AllSku = Get-MgSubscribedSku
    if ($AllSku.count -le 0){
        Write-Error ("No SKU found.  Do you have permissions to run Get-MGSubscribedSKU? `n Connect-MGGraph -scopes Organization.Read.All, Directory.Read.All, Organization.ReadWrite.All, Directory.ReadWrite.All")
    }
    else {
         Write-Log ("Found " + $AllSku.count + " SKUs in the tenant")
    }
	
    # Make sure out plan array is null
    [array]$Plans = $null

    # Build a list of all of the plans from all of the SKUs
    foreach ($Sku in $AllSku) {
        $SKU.serviceplans.serviceplanname | ForEach-Object { [array]$Plans = $Plans + $_.servicename }
    }

    # Need just the unique plans
    $Plans = $Plans | Select-Object -Unique | Sort-Object
    Write-Log ("Found " + $Plans.count + " Unique plans in the tenant")

    # Make sure the output array is null
    $Output = $null

    # Create a false SKU object so we populate the first entry in the array with all needed values so we get a good CSV export
    # Basically a cheat so we can easily use export-csv we will remove this entry from the actual output
    $Object = $Null
    $Object = New-Object -TypeName PSobject
    $Object | Add-Member -MemberType NoteProperty -Name DisplayName -Value "BASELINE_IGNORE"
    $Object | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value "BASELINE_IGNORE@contoso.com"
    $Object | Add-Member -MemberType NoteProperty -Name UsageLocation -Value "BASELINELOCATION_IGNORE"
    $Object | Add-Member -MemberType NoteProperty -Name IsDeleted -Value "BASELINEDELETED_IGNORE"
    $Object | Add-Member -MemberType NoteProperty -Name SkuID -Value "BASELINESKU_IGNORE"
    $Object | Add-Member -MemberType NoteProperty -Name CommonName -Value "BASELINESKUNAME_IGNORE"
    $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value "BASELINEINHERIT_IGNORE"

    # Populate a value into all of the plan names
    foreach ($value in $Plans) {

        $Object | Add-Member -MemberType NoteProperty -Name $value -Value "---"
    }

    # Create the output list array
    $Output = New-Object System.Collections.ArrayList

    # Populate in the generic object to the list array
    $Output.Add($Object) | Out-Null

    # Make sure our UserToProcess array is created and is null
    [array]$UserToProcess = $null

    # See if our user array is null and pull all users if needed
    if ($null -eq $Users) {

        Write-Log "Getting all users in the tenant."
        Write-Log "This can take some time."

        # Get all of the users in the tenant
        [array]$UserToProcess = Get-MsolUser -All
        Write-log ("Found " + $UserToProcess.count + " users in the tenant")

        # Gather the deleted users as well if we want them
        if ($IncludeDeletedUsers){
            [array]$UserToProcess += Get-MsolUser -ReturnDeletedUsers -All
            Write-Log ("Found " + $UserToProcess.count + " users and deleted users in tenant")
        }

    }
    
    # Gather just the users provided
    else {
		
        # Make user our Users object is valid
        [array]$Users = Test-UserObject -ToTest $Users

        Write-Log "Gathering License information for provided users"
        $i = 0
        
        # Get the data for each user to use in generating the report
        foreach ($account in $Users) {
            $i++
            [array]$UserToProcess = [array]$UserToProcess + (Get-MsolUser -UserPrincipalName $Account.UserPrincipalName)
            if (!($i % 100)) { Write-log ("Gathered Data for " + $i + " Users") }
        }

        Write-log ("Found " + $UserToProcess.count + " users to report on")
    }

    # Now that we have all of the user objects we need to start processing them and building our output object

    # Setup a counter for informing the user of progress
    $i = 0
	
    # Null out the group cache array
    # Array holds group names and GUIDs so we do not have a lookup a group more than once
    [array]$GroupCache = $Null

    # Pull in our SkuID to Common Name map
    $SkuArray = Get-Content (join-path (Split-path (((get-module MSOLLicenseManagement)[0]).path) -Parent) "SkuMap.json") | convertfrom-json

    # Process each user
    Write-Log "Generating Report"
    foreach ($UserObject in $UserToProcess) {
	
        # Increase the counter
        $i++
        
        # Output every 100 users for progress
        if (!($i % 100)) { Write-log ("Finished " + $i + " Users") }
		
        # If no license is assigned we need to create that object
        if ($UserObject.isLicensed -eq $false) {
		
            $Object = $null
		
            # Create our object and populate in our values
            $Object = New-Object -TypeName PSobject
            $Object | Add-Member -MemberType NoteProperty -Name DisplayName -Value $UserObject.displayname
            $Object | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserObject.userprincipalname
            $Object | Add-Member -MemberType NoteProperty -Name UsageLocation -Value $UserObject.UsageLocation
            $Object | Add-Member -MemberType NoteProperty -Name IsDeleted -value ([bool]$UserObject.SoftDeletionTimestamp)
            $Object | Add-Member -MemberType NoteProperty -Name SkuID -Value "UNASSIGNED"
            $Object | Add-Member -MemberType NoteProperty -Name CommonName -Value "UNASSIGNED"
            $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value "Explicit"
		
            # Add the object to the output array
            #[array]$Output = $Output + $Object
            $Output.Add($Object) | Out-Null
        }
		
        #If we have a license then add that information
        else {
            # We can have more than one license on a user so we need to do each one seperatly
            foreach ($license in $UserObject.licenses) {
                $Object = $null
				
                # Get the common name of the licnese from our skuid.json
                $CommonName = Get-SKUCommonName -Skuid ((($license.accountskuid).split(":"))[-1]) -SkuArray $SkuArray
				
                # Create our object and populate in our values
                $Object = New-Object -TypeName PSobject
                $Object | Add-Member -MemberType NoteProperty -Name DisplayName -Value $UserObject.displayname
                $Object | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserObject.userprincipalname
                $Object | Add-Member -MemberType NoteProperty -Name UsageLocation -Value $UserObject.UsageLocation
                $Object | Add-Member -MemberType NoteProperty -Name IsDeleted -value ([bool]$UserObject.SoftDeletionTimestamp)
                $Object | Add-Member -MemberType NoteProperty -Name SkuID -Value $license.accountskuid
                $Object | Add-Member -MemberType NoteProperty -Name CommonName -Value $CommonName

                # Add each of the plans for the license in along with its status
                foreach ($value in $license.servicestatus) {
				
                    $Object | Add-Member -MemberType NoteProperty -Name $value.serviceplan.servicename -Value $value.ProvisioningStatus
				
                }
				
                # Get the inherited from / status of the license based on GBL
                # If there are no groups in GroupsAssigningLicense then it is explicitly assigned
                if ($license.GroupsAssigningLicense.count -le 0) {
					
                    $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value "Explicit"				
				
                }
                # If it is populated then we need to process ALL values
                else {
				
                    # Null out our assigned by string
                    [string]$assignedby = $null
					
                    # Process each group entry and work to resolve them
                    foreach ($entry in $license.GroupsAssigningLicense) {
					
                        # If the GUID = the ObjectID then it is direct assigned
                        if ($entry.guid -eq $UserObject.objectid) { [string]$assignedby = $assignedby + "Explicit;" }
						
                        # If it doesn't match the objectid of the user then it is a group and we need to resolve it
                        else {
						
                            # if groupcache isn't populated yet then we know we need to look it up
                            if ($null -eq $GroupCache) {
                                [array]$GroupCache = $GroupCache + (Get-MsolGroup -objectid $entry.guid | Select-Object -Property objectid, displayname)
                            }
                            # If not null check if we have a cache hit
                            else {
                                # If the guid is in the groupcache then no action needed
                                if ($GroupCache.objectid -contains $entry.guid) { }
								
                                # If we have a cache miss then we need to populate it into the cache
                                else { [array]$GroupCache = $GroupCache + (Get-MsolGroup -objectid $entry.guid | Select-Object -Property objectid, displayname) }
                            }
							
                            # Add the group information from the cache to the output
                            [string]$assignedby = $assignedby + ($GroupCache | Where-Object { $_.objectid -eq $entry.guid }).DisplayName + ";"
                        }
					
                    }
				
                    # Add the value of assignment to the object
                    $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value ($assignedby.trimend(";"))
                }
				
				
                # Add the object created from the user information to the output array
                #[array]$Output = $Output + $Object
                $Output.Add($Object) | Out-Null
            }
        }
    }
	
    # Now that we have all of our output objects in an array we need to output them to a file
    # If overwrite has been set then just take the file name as the date and force overwrite the file
    if ($OverWrite) {
        # Build our file name
        $path = Join-path (Split-Path $LogFile -Parent) ("License_Report_" + [string](Get-Date -UFormat %Y%m%d) + ".csv")
    }
	
    # Default behavior will be to create a new incremental file name
    else {
        # Build our file name
        $RootPath = Join-path (Split-Path $LogFile -Parent) ("License_Report_" + [string](Get-Date -UFormat %Y%m%d))
        $TryPath = $RootPath + "*"
		
        # Find any existing files that start with our planed file name
        $FilesInPath = Get-ChildItem $TryPath

        # Found files so we need to increment our file name with _X
        if ($FilesInPath.count -gt 0) {
            $Path = $RootPath + "_" + ($FilesInPath.count) + ".csv"
        }
        # Didn't find anything so we are good with the base file name
        else {
            $Path = $RootPath + ".csv"
        }
    }

    Write-Log ("Exporting report to " + $path)

    # Export everything to the CSV file
    $Output | Export-Csv $path -Force -NoTypeInformation
		
    # Pull it back in so we can remove our fake object
    $Temp = Import-Csv $path
    $Temp | Where-Object { $_.UserPrincipalName -ne "BASELINE_IGNORE@contoso.com" } | Export-Csv $path -Force -NoTypeInformation
	
    Write-log "Report generation finished"
}

