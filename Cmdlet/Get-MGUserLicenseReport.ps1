# Generates an MSOL User License Report

Function Get-MGUserLicenseReport {
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
    
	.OUTPUTS
	Log file showing all actions taken by the function $env:LOCALAPPDATA\
	CSV file named License_Report_YYYYMMDD_X.csv that contains the report.

	.EXAMPLE
	Get-MGUserLicenseReport -LogFile C:\temp\report.log

	Creates a new License_Report_YYYYMMDD_X.csv file with the license report in the c:\temp directory for all users.

	.EXAMPLE
	Get-MGUserLicenseReport -LogFile C:\temp\report.log -Users $SalesUsers -Overwrite

	OverWrites the existing c:\temp\License_Report_YYYYMMDD.csv file with a license report for all users in $SalesUsers.
	
    #>
    
    param 
    (
        [array]$Users,
        [switch]$OverWrite = $false
    )

    # Make sure we have the connection to MSOL
    Test-MGServiceConnection
	
    Write-SimpleLogFile "Generating Sku and Plan Report" -OutHost

    # Get all of the availible SKUs
    $AllSku = Get-MgSubscribedSku 2>%1
    if ($AllSku.count -le 0) {
        Write-Error ("No SKU found! Do you have permissions to run Get-MGSubscribedSKU? `nSuggested Command: Connect-MGGraph -scopes Organization.Read.All, Directory.Read.All, Organization.ReadWrite.All, Directory.ReadWrite.All")
    } else {
        Write-SimpleLogFile ("Found " + $AllSku.count + " SKUs in the tenant") -OutHost
    }
	
    # Make sure our plan array is null
    [array]$Plans = $null

    # Build a list of all of the plans from all of the SKUs
    foreach ($Sku in $AllSku) {
        $SKU.ServicePlans.ServicePlanName | ForEach-Object { [array]$Plans = $Plans + $_ }
    }

    # Need just the unique plans
    $Plans = $Plans | Select-Object -Unique | Sort-Object
    Write-SimpleLogFile ("Found " + $Plans.count + " Unique plans in the tenant") -OutHost

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
    $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value "BASELINEINHERIT_IGNORE"

    # Populate a value into all of the plan names
    foreach ($value in $Plans) {

        $Object | Add-Member -MemberType NoteProperty -Name $value -Value "---"
    }

    # Create the output list arrays
    $Output = New-Object System.Collections.ArrayList
    $UserToProcess = New-Object System.Collections.ArrayList

    # Populate in the generic object to the list array
    $Output.Add($Object) | Out-Null

    # Make sure our UserToProcess array is created and is null
    $UserToProcess = $null

    # See if our user array is null and pull all users if needed
    if ($null -eq $Users) {

        Write-SimpleLogFile "Getting all users in the tenant." -OutHost
        Write-SimpleLogFile "This can take some time." -OutHost

        # Get all of the users in the tenant
        $UserToProcess = Get-MgUser -All
        Write-SimpleLogFile ("Found " + $UserToProcess.count + " users in the tenant") -OutHost

    }
    
    # Gather just the users provided
    else {
		
        # Make user our Users object is valid
        [array]$Users = Test-MGUserObject -ToTest $Users

        Write-SimpleLogFile "Gathering License information for provided users" -OutHost
        $i = 0
        
        # Get the data for each user to use in generating the report
        foreach ($account in $Users) {
            $i++
            Update-MGProgress -CurrentCount $i -MaxCount $users.count -Message "Gathering License information for provided users"
            ### TODO: Need some error handling here for "bad inputs"
            $UserToProcess.Add((get-mguser -ConsistencyLevel eventual -Search ("Userprincipalname:" + $account.userprincipalname)))            
            if (!($i % 100)) { Write-SimpleLogFile ("Gathered Data for " + $i + " Users") }
        }

        Write-SimpleLogFile ("Found " + $UserToProcess.count + " users to report on")
    }

    ### Now that we have all of the user objects we need to start processing them and building our output object  ###

    # Setup a counter for informing the user of progress
    $i = 0
	
    # Process each user
    Write-SimpleLogFile "Generating Report" -OutHost
    foreach ($UserObject in $UserToProcess) {
	
        # Increase the counter
        $i++
        Update-MGProgress -CurrentCount $i -MaxCount $UserToProcess.count -Message "Generating Report"
        
        # Output every 100 users for progress
        if (!($i % 100)) { Write-SimpleLogFile ("Finished " + $i + " Users") }

        # Get the license assignments for the user
        $licenseOptions = Get-MgUserLicenseDetail -UserId $UserObject.Id
		
        # If no license is assigned we need to create that object
        if ($licenseOptions.count -eq 0) {
		
            $Object = $null
		
            # Create our object and populate in our values
            $Object = New-Object -TypeName PSobject
            $Object | Add-Member -MemberType NoteProperty -Name DisplayName -Value $UserObject.displayname
            $Object | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserObject.userprincipalname
            $Object | Add-Member -MemberType NoteProperty -Name UsageLocation -Value $UserObject.UsageLocation
            $Object | Add-Member -MemberType NoteProperty -Name IsDeleted -value ([bool]$UserObject.SoftDeletionTimestamp)
            $Object | Add-Member -MemberType NoteProperty -Name SkuID -Value "UNASSIGNED"
            $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value "Explicit"
		
            # Add the object to the output array
            $Output.Add($Object) | Out-Null
        }
		
        #If we have a license then add that information
        else {
            
            # Get how the licenses are assigned
            $UserLicenseAssignment = Get-MGUserLicenseAssignmentState -UserId $UserObject.Id
            
            # We can have more than one license on a user so we need to do each one seperatly
            foreach ($license in $licenseOptions) {
                $Object = $null

                # Determine how the license is assigned
                $assignment = Get-MGLicenseAssignmentMethod -UserAssignementArray $UserLicenseAssignment -SkuToResolve $license.SkuId
                			
                # Create our object and populate in our values
                $Object = New-Object -TypeName PSobject
                $Object | Add-Member -MemberType NoteProperty -Name DisplayName -Value $UserObject.displayname
                $Object | Add-Member -MemberType NoteProperty -Name UserPrincipalName -Value $UserObject.userprincipalname
                $Object | Add-Member -MemberType NoteProperty -Name UsageLocation -Value $UserObject.UsageLocation
                $Object | Add-Member -MemberType NoteProperty -Name IsDeleted -value ([bool]$UserObject.SoftDeletionTimestamp)
                $Object | Add-Member -MemberType NoteProperty -Name SkuID -Value $license.SKUPartNumber
                $Object | Add-Member -MemberType NoteProperty -Name Assignment -Value $assignment

                # Add each of the plans for the license in along with its status
                foreach ($value in $license.ServicePlans) {
				
                    $Object | Add-Member -MemberType NoteProperty -Name $value.ServicePlanName -Value $value.ProvisioningStatus
				
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
        $RootPath = Join-path $env:LOCALAPPDATA ("License_Report_" + [string](Get-Date -UFormat %Y%m%d))
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

    Write-SimpleLogFile ("Exporting report to " + $path) -OutHost

    # Export everything to the CSV file
    $Output | Export-Csv $path -Force -NoTypeInformation
		
    # Pull it back in so we can remove our fake object
    $Temp = Import-Csv $path
    $Temp | Where-Object { $_.UserPrincipalName -ne "BASELINE_IGNORE@contoso.com" } | Export-Csv $path -Force -NoTypeInformation
	
    Write-SimpleLogFile "Report generation finished" -OutHost
}

