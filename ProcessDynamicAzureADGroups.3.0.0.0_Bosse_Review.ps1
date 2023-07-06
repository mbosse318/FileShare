<# 
.SYNOPSIS
Creates one or more dynamic Azure AD groups based on metat data from a csv input file by usine Ms graph PowerShell

.OUTPUT
A log file with the summary of all the operations performed

.EXAMPLE 
ProcessDynameicAzureADGroups.ps1 - Inputfile "C:Data\Inpute\CreateDyanmicAzureGroups.csv" - LogFilesPath "C:Data\Output\"

Description
-----------
This command reads the meatadata from the input file and creates dynamic Azure AD groups for each row in the input file. The log file 
is written to the folder specified.

.EXAMPLE 
ProcessDyanmicAzureADGroups.ps1 -InputFile "C:\Data\Input\CreateDynamicAzureADGroups.csv"

Description
-----------
This command reads the metadata from the input file and creates dyanmic Azure AD groups for each row in the input file. The logfile is written
to the PowerShell script working location
#>

Param(
	[parameter(Mandatory = $true)]
	[String]
	$InputFile,

	[parameter(Mandatory = $false)]
	[String]
	$LogFilePath = "",

	[parameter(Mandatory = $false)]
	[Switch]
	$WhatIf = $false
)

$scriptStartTime = Get-Date
# Bosse Review - $todaysdate variable was misspelled - resulting in the log filename not having the date/time added
$todaysdate = Get-Date -Format "MM_dd_yyyy_hh_mm_ss"
$currentLocation = Get-Location

Function AssembleFilePath() {
	Param(
		[String] $filePath,
		[String] $baseFilename,
		[String] $extension,
		[bool] $appendCurrentDate = $true
	)

	# Bosse Review - this is declared a second time and never used. The variable that is used is not spelled the same, so it's actually blank.
	#$todaysdate= Get-Date -Format "MM-dd-yyyy-HH-mm-ss"
	#$currentLocation = Get-Location
		
	#user the current location for the log file path
	if ($filePath -eq "") { $filePath = $currentLocation.Path }
		
	#make sure log file path ends in backlash
	if ($filePath -notmatch '\\$') { $filePath += '\' }
		
	#assemble the file path 
	$filePath += $baseFilename
	# Bosse Review - the use of $todaysdate is always blank because a different variable is used to get the date string.  The $todaysdate variable used here will always be null.
	if ($appendCurrentDate) { $filePath += ("_" + $todaysdate) }
	$filePath += ("." + $extension)
	return $filePath
}

Function ParseParameters() {
	Param(
		[String] $inputFile,
		[String] $logFilePath,
		[Bool] $whatIf
	)

	$logFilePath = AssembleFilePath $logFilepath "DynamicAzureADGroupCreation-Log" "log" $true
	
	$paramObj = New-Object PSObject
	Add-Member -InputObject $paramObj -MemberType NoteProperty -Name InputFile -Value $inputFile 
	Add-Member -InputObject $paramObj -MemberType NoteProperty -Name LogFilePath -Value $logFilePath 
	Add-Member -InputObject $paramObj -MemberType NoteProperty -Name WhatIf -Value $whatIf 
	
	Return $paramObj
}

Function CheckParameters() {
	Param(
		[String] $inputFile,
		[String] $logFilePath,
		[Bool] $whatIf
	)
		
	#check for writability to log file 
	#do this because Start-Transcript does not thow an exception if the path does not exist 
	try {
		Out-File -FilePath $logFilePath -InputObject "Starting Process" -Encoding ASCII -ErrorAction Stop
	}
	catch {
		throw "Could not create the log file: $($logFilePath)."
	}
	# check if the input file exists 
	$CheckInputFile = Test-Path -Path $inputFile 
	if ($CheckInputFile -eq $false) {
		throw "Could not find the specified input file : $($inputFile)."
	}
		
}

Function EnsureMicrosoftGraphModuleInstalled() {
	try {
		# avoid use of the MsGraph module (need to check if it required ?)
		<#Uninstall-Module Microsoft.Graph ## 1st Unistall Main Module ##Then,Remove all of the dependency modules through following commands
        Get-InstalledModule Microsoft.Graph.* | %{ if($_.Name -ne "Microsoft.Graph.Authentication"){ Uninstall-Module $_.Name } }
        Uninstall-Module Microsoft.Graph.Authentication#>

		#install module if it's not already installed
		if (Get-InstalledModule Microsoft.Graph) {
			Write-Host "Microsoft Graph PowerShell module Already Installed" -ForegroundColor Green

		}
		else {
			Write-Host "Microsoft Graph PowerShell module Not Installed. Installing........." -ForegroundColor Red
			Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
			Y
			#verifying installation
			Write-Host "verifying installation Microsoft Graph PowerShell module ........."
			Get-InstalledModule Microsoft.Graph
			Write-Host "verifying installation Microsoft Graph PowerShell sub-module ........."
			#verifying Sub - Module installed
			Get-InstalledModule
			Write-Host "Updating module ........."
			#Updating module
			Update-Module Microsoft.Graph
			Write-Host "Microsoft Graph PowerShell module Installed" -ForegroundColor Green
		}

		Import-Module Microsoft.Graph.Groups
	}
	catch {
		throw "Could not_install/import the AzureADPreview module."
	}
}

Function OwnerStringToArray() {
	Param(
		[String] $ownerString
	)

	$option = [System.StringSplitoptions]::RemoveEmptyEntries

	$ownerArray = $ownerString.Split(";", $option)
	return $ownerArray
}

Function CompareOwners() {
	param(
		[string] $newOwners,
		[Object[]] $existingOwners
	)

	$ownersToAdd = @()
	$ownersToRemove = @()
	
	#new owners comes from the input file as a semicolon separated set of values, so translate to an array
	$newOwnerArray = OwnerStringToArray $newowners
	
	#loop through each expected value and compare against each actual value, ignoring case
	foreach ($newOwer in $newOwnerArray) {
		$found = $false
        
		foreach ($owner in $existingOwners) {
			$existingOwner = Get-MgUser -UserId $owner.Id | Select-Object -ExpandProperty UserPrincipalName
			if ($existingOwner.Trim().ToLower() -eq $newOwer.Trim().ToLower()) {
				$found = $true
				break
			}
		}
		if ($found -eq $false) {
			#the new owner is not an existing owner, owner, so add it to the list of owner to add
			$ownersToAdd += $newOwer
		}
	}

	#loop through each expected value and compare against each actual value, ignoring case
	foreach ($owner in $existingOwners) {
		$existingOwner = Get-MgUser -UserId $owner.Id | Select-Object -ExpandProperty UserPrincipalName
		$found = $false
		foreach ($newOwner in $newOwnerArray) {
			# $existingOwner=Get-MgUser -UserId $owner.Id | Select-Object -ExpandProperty UserPrincipalName
			
			if ($existingOwner.Trim().ToLower() -eq $newOwner.Trim().ToLower()) {
				$found = $true
				break
			}
		}
		
		if ($found -eq $false) {	
			#the existing owner is not an new owner array, so add it to the list of owners to remove
			#$ownersToRemove += $owner.UserPrincipalName
			$ownersToRemove += $existingOwner
		}
	}

	$resultobj = New-Object PSObject
	Add-Member -Inputobject $resultobj -MemberType NoteProperty -Name OwnersToAdd -Value $ownersToAdd
	Add-Member -InputObject $resultobj -MemberType NoteProperty -Name OwnersToRemove -Value $ownersToRemove
	#all good
	return $resultobj
}

Function AddOwner() {
	Param (
		[Object] $c,
		[String] $owner
	)

	Write-Host "Adding owner $($owner) to group $($group.DisplayName)" -ForegroundColor Green
	$groupIdStr = [string]$group.Id
  
	#get the user object from Azure AD
	$ownerobject = Get-MgUser -UserId "$($owner.Trim())"
	
	if ($ownerobject -eq $null) {
		Write-Host "Error: $($owner) was not found." -ForegroundColor Red
	}
	else {
		$newGroupOwner = @{
			"@odata.id" = "https://graph.microsoft.com/v1.0/users/$owner"
		}

		#New-MgGroupOwnerByRef -GroupId $groupIdStr -BodyParameter $newGroupOwner

		New-MgGroupOwnerByRef -GroupId $groupIdStr.Trim() -AdditionalProperties @{"@odata.id" = "https://graph.microsoft.com/v1.0/users/$owner" }

		Write-Host "Succesfully added owner $($owner) to group $($group.DisplayName)" -ForegroundColor Green
	}
}

Function AddOwners() {
	Param (
		[Object] $group,
		[String[]] $owners
	)

	foreach ($owner in $owners) {
		AddOwner $group $owner.Trim()
	}
}

Function RemoveOwner() {
	Param(
		[Object] $group,
		[String] $owner
	)

	# Bosse Review - import this just once when module is loaded
	#Import-Module Microsoft.Graph.Groups

	Write-Host "Removing owner $($owner) to group $($group.DisplayName)" -ForegroundColor Green
	
	#get the user object from Azure AD
	$ownerobject = Get-MgUser -UserId $owner.Trim()
	
	if ($ownerobject -eq $null) {
		Write-Host "Error: $($owner) was not found." -ForegroundColor Red
	}

	else {
		Remove-MgGroupOwnerByRef -GroupId $($group.ObjectID) -DirectoryObjectId $($ownerobject.Id)

		Write-Host "Succesfully removed owner $($owner) from group $($group.DisplayName)" -ForegroundColor Green
	}
}

Function RemoveOwners() {
	Param (
		[Object] $group,
		[String[]] $owners
	)

	foreach ($owner in $owners) {
		RemoveOwner $group $owner.Trim()
	}
}

Function GetExistingGroupByDisplayName() {
	Param(
		[String] $displayName
	)

	try {

		$getAllGroups = Get-MgGroup -ConsistencyLevel eventual -Count groupCount -Search "DisplayName:$($displayName.Trim())"  | Select *
		$groupFound = $false
	
		if ($getAllGroups -ne $null) {
			if ($getAllGroups -is [array]) {
				#more than one group found with the display name - which is possible, since SearchString operates as "startswith"
				foreach ($foundGroup in $getAllGroups) {
					# check for an exact match
					if ($foundGroup.DisplayName -eq $displayName) {
						#this is the exact group
						$existingGroup = $foundGroup
						$groupFound = $true
						break
					}
				}
			}
			else {
				#still need to make sure this group matches exactly and not just starts with
				if ($getAllGroups.DisplayName -eq $displayName) {
					$existingGroup = $getAllGroups
					$groupFound = $true
				}
			}
		}

		if ($groupFound -eq $false) {
			$existingGroup = $null
		}
	}
	catch {
		Write-Host $PSItem.Exception.Message -ForegroundColor Red
	}

	return $existingGroup
}

Function GetExistingGroupByObjectId() {
	Param(
		[String] $objectId
	)
	
	$existingGroup = Get-MgGroup -GroupId $objectId | Select *
	
	return $existingGroup
}

Function GetExistingGroup() {
	Param (
		[Object] $group
	)

	if ($group.Action -eq "update") {
		$groupIdStr = [string]$group.objectID
		$existingGroup = GetExistingGroupByobjectId  $groupIdStr
	}
	elseif ($group.Action -eq "Add") {
		$existingGroup = GetExistingGroupByDisplayName $group.DisplayName
	}
	else {
		$existingGroup = $null
	}

	return $existingGroup
}

Function AddGroup() {
	# Bosse Review - shouldn't we import the module earlier in the code so all functions have it in scope?
	#Import-Module Microsoft.Graph.Groups

	# Bosse Review - seems like Param () is missing around this object.
	Param (
		[Object] $group
	)
	Write-Host "Adding group $($group.DisplayName) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green

	#check to see if group already exists with the specified DisplayName
	$existingGroup = GetExistingGroup $group
	if ($existingGroup -ne $null) {
		#group with the DisplayName already exists, so don't attempt to create another group and write to the output
		throw "Group with DisplayName $($group.DisplayName) already exists."
	}
	else {

		#create the new group
		$newGroup = New-MgGroup -DisplayName "$($group.DisplayName)" `
			-Description "$($group.Pbl)" `
			-MailEnabled:$False `
			-MailNickName 'MailNickname' `
			-SecurityEnabled:$true `
			-GroupTypes DynamicMembership `
			-MembershipRule "$($group.Rule)"  `
			-MembershipRuleProcessingState On `
			-IsAssignableToRole: $false `
			-Visibility 'Public'



		Write-Host " Succesfully created the dynamic AD Group: $($group.DisplayName) as ObjectId $($newGroup.Id) from row
		$($currentRow.ToString()) of the input file." -ForegroundColor Green
	}

	return $newGroup
}

Function ValidateGroup() {
	Param(
		[Object] $group
	)

	#validate the data
	if ($group.Action -eq "Update" -and $group.ObjectId.Length -eq 0) {
		throw "ObjectId is blank for an Update action in row $($currentRow.ToString()) of the input file."
	}
	if ($group.DisplayName.Trim().Length -eq 0) {
		throw "DisplayName is blank in row $($currentRow.ToString()) of the input file."

	}
	if ($group.Owner.Length -eq 0) {
		throw "Owner is blank in row $($currentRow.ToString()) of thè input file."
	}
	if ($group.Pbl.Length -eq 0) {
		throw "PBL is blank in $($currentRow.ToString()) of the input file."
	}
	if ($group.Rule.Length -eq 0) {
		throw "Rule is blank in row $($currentRow.ToString()) of the input file."
	}

	#if we make it this far, all good
}

Function UpdateGroup() {
	# Bosse Review - import this just once when module is loaded
	#Import-Module Microsoft.Graph.Groups
	
	#Bosse Review - missing Param()
	Param (
		[Object] $group
	)
    
	$updatesMade = $false
    
	Write-Host "Updating Group' ID $($group.ObjectId) from row $($currentRow.ToString())-of the input file." -ForegroundColor Green
	#check to see if group already exists with the specified DisplayName
	$existingGroup = GetExistingGroup $group
	if ($existingGroup -eq $null) {
		#group with the ObjectId does not exist
		throw "Group with ID $($group.objectId) does not exist."
	}
	else {
		if ($group.DisplayName -ne $existingGroup.DisplayName) {
			#if we're changing the display name, make sure the new display name does not exist already
			$checkGroupName = GetExistingGroupByDisplayName $group.DisplayName
			if ($checkGroupName -ne $null) {
				# a group with the new DisplayName already exists
				throw "Can't update the DisplayName of Group ID $($group.ObjectId) because another group ($($checkGroupName.Id) DisplayName."
			}
		}
		# compare values to see if we need to change anything
		if ($group.DisplayName -ne $existingGroup.DisplayName -or
			$group.Pbl -ne $existingGroup.Description -or
			$group.Rule -ne $existingGroup.MembershipRule) {
			#update the existing group

			$params = @{
				"DisplayName"    = "$($group.DisplayName)"
				"Description"    = "$($group.Pbl)"
				"MembershipRule" = "$($group.Rule)"
				"MailNickname"   = "MailNickname"
				"Visibility"     = "Public"
			}

			Update-MgGroup -GroupId $group.ObjectId -BodyParameter $params
					
			$updatesMade = $true
		}

		#check to see if we need to change the owpers
		#geting UserPrincipalNames as array
		$owerObjs = Get-MgGroupOwner -GroupId $group.ObjectId -All
		$ownerResults = CompareOwners $group.Owner $owerObjs
			
		if ($ownerResults.OwnersToAdd.Length -gt 0) {
			AddOwners $existingGroup $ownerResults.OwnersToAdd
			$updatesMade = $true
		}
		if ($ownerResults.OwnersToRemove.Length -gt 0) {
			RemoveOwners $group $ownerResults.OwnersToRemove
			$updatesMade = $true
		}
		if ($updatesMade -eq $true) {
			Write-Host " Succesfully updated the group $ ($group.DisplayName) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green
		}			
		else {
			Write-Host "No updates needed for group $($group.DisplayName) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green
		}
	}
}

Function ProcessGroups() {
	Param(
		[Object[]] $groupsToProcess,
		[Object] $params
	)

	$currentRow
	$totalGroupsProcessed = 0
	$totalGroupsAdded = 0
	$totalGroupsUpdated = 0
	$totalGroupsFailed = 0

	#loop through each item from the input file and create the groups
	foreach ($group in $groupsToProcess) {

		$currentRow += 1
		$totalGroupsProcessed += 1
		try {
			ValidateGroup $group
			if ($group.Action -eq "Add") {
				$addedGroup = AddGroup $group
				# Bosse Review - total for added groups was not being incremented
				$totalGroupsAdded += 1
				
				$ownerArray = OwnerStringToArray $group.Owner
				AddOwners $addedGroup $ownerArray
				
			}
			elseif ($group.Action -eq "Update") {
				UpdateGroup $group
				$totalGroupsUpdated += 1
			}
			else {
				throw "Invalid Action in row $($currentRow.ToString()) of the input file."
			}
			
		}
		catch {
			# log the error and move on to the next item
			Write-Host "Error processing the dynamic AD Group: $($group.DisplayName) in row $($currentRow.ToString()) of the input file."-ForegroundColor Red
			Write-Host $PSItem.Exception.Message -ForegroundColor Red
			Write-Host $PSItem.InvocationInfo.PositionMessage -ForegroundColor Red
			$totalGroupsFailed += 1
		}	
	}

	# Bosse Review
	# There is nothing returned by this method.  The caller of this function expects a collection of processedGroups back.
	# how could testing of this script have ever worked?
	$results = New-Object psobject
	Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsProcessed -Value $totalGroupsProcessed
	Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsAdded -Value $totalGroupsAdded
	Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsUpdated -Value $totalGroupsUpdated
	Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsFailed -Value $totalGroupsFailed

	return $results
}

Function OutputContextInfo() {
	Param (
		[Object[]] $context
	)

	# Bosse Review - after getting better context info, we can add more details to this output
	#Write-Host $context 
	Write-Host "Connected Account: $($context.Account)" -ForegroundColor Yellow
	#Write-Host "Connected TenantDomain: $($context.TenantDomain)" -ForegroundColor Yellow
	#Write-Host "Connected Environment: $($context.Environment)" -ForegroundColor Yellow
	Write-Host "Connected TenantId: $($context.TenantId)" -ForegroundColor Yellow
}

Function Main() {
	$scriptStartTime = Get-Date
	$params = ParseParameters $InputFile $LogFilePath $WhatIf
	CheckParameters $params.InputFile Sparams.LogFilePath $params.WhatIf
	
	Start-Transcript -Path $params.LogFilePath
	
	Write-Host "Script Version: $($scriptVersion) " -ForegroundColor yellow
	Write-Host "WhatIf: $($params.WhatIf.ToString())" -ForegroundColor yellow
	
	EnsureMicrosoftGraphModuleInstalled

	#connect to MSGrapph AD - this will prompt for credentials
	# Bosse Review - I don't think this script needs User.ReadWrite.All
	# looks like the process to update owners requires User.ReadWrite.All
	#$context = 
	Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"
	# Bosse Review - to get better context info about this current session, use Get-MgContext
	# Bosse Review - don't actually really need to do this as the call to Disconnect-MgGraph at the end of this script automatically outputs the context details. 
	#	$context = Get-MgContext

	#output conext info
	#	OutputContextInfo $context
	#read the input file
	$groupsToprocess = Import-Csv -Path $params.InputFile
	#$groupsToprocess = Import-Csv -Path "D:\MR\OfficeWork\PowerShellScrips\ProcessAzureGroups.csv"
	
	#main processing
	$processedGroups = ProcessGroups $groupsToprocess

	#write summary information
	$scriptDuration = (New-TimeSpan -Start $scriptStartTime -End (Get-Date)).ToString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'")
	
	# Bosse Review - misspelling of ForegroundColor in several places
	Write-Host "************************" -ForegroundColor Yellow
	Write-Host "Script Summary" -ForegroundColor Yellow
	Write-Host "Script duration: $($scriptDuration)" -ForegroundColor Yellow
	Write-Host "Total groups added: $($processedGroups.TotalGroupsAdded)" -ForegroundColor Yellow
	Write-Host "Total groups updated: $($processedGroups.TotalGroupsUpdated)" -ForegroundColor Yellow
	Write-Host "Total groups failed: $($processedGroups.TotalGroupsFailed)" -ForegroundColor Yellow
	Write-Host "Total groups processed: $($processedGroups.TotalGroupsProcessed)" -ForegroundColor Yellow

	Disconnect-MgGraph

	Stop-Transcript
}

# start of script processing
$scriptVersion = "3.0.0.0"

Main