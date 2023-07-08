<# 
.SYNOPSIS
Creates one or more dynamic Azure AD Groups based on meta data from a CSV input file

.OUTPUTS
A log file with the summary of all the operations performed

.EXAMPLE
ProcessDynamicAzureADGroups.ps1 -InputFile "C:\Data\Input\CreateDynamicAzureADGroups.csv" -LogFilePath "C:\Data\Output\"

Description
-----------
This command reads the metadata from the input file and creates dynamic Azure AD groups for each row in the input file.  The log file is written to the folder specified.

.EXAMPLE
ProcessDynamicAzureADGroups.ps1 -InputFile "C:\Data\Input\CreateDynamicAzureADGroups.csv"

Description
-----------
This command reads the metadata from the input file and creates dynamic Azure AD groups for each row in the input file.  The log file is written to the PowerShell script working location.

#>

Param (
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
$todaysdate = Get-Date -Format "MM_dd_yyyy_hh_mm_ss"
$currentLocation = Get-Location

Function CheckOwners()
{
    param (
        [string]$ExpectedOwners,
        [Object[]]$ActualOwners
    )

    # expected owners comes from the input file as a semicolon separated set of values, so translate to an array
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $ExpectedOwnerArray = $ExpectedOwners.Split(";", $option)

    # check the array length first
    if ($ExpectedOwnerArray.Length -ne $ActualOwners.Length)
    {
        return $false;
    }

    # loop through each expected value and compare against each actual value, ignoring case
    $matchCount = 0
    foreach ($expectedOwner in $ExpectedOwnerArray)
    {
        foreach ($owner in $ActualOwners)
        {
            if ($owner.UserPrincipalName.Trim().ToLower() -eq $expectedOwner.Trim().ToLower())
            {
                $matchCount += 1
                break
            }
        }
    }

    if ($matchCount -ne $ExpectedOwnerArray.Length)
    {
        return $false;
    }

    # all good
    return $true
}

Function AssembleFilePath()
{
    Param (
        [String] $filePath,
        [String] $baseFilename,
        [String] $extension,
        [bool] $appendCurrentDate=$true
    )

    $todaysdate = Get-Date -Format "MM-dd-yyyy-HH-mm-ss"
    $currentLocation = Get-Location

    # use the current location for the log file path
    if ($filePath -eq "") { $filePath = $currentLocation.Path }

    # make sure log file path ends in backslash
    if ($filePath -notmatch '\\$') { $filePath += '\' }

    # assemble the file path
    $filePath += $baseFilename
    if ($appendCurrentDate) { $filePath += ("_" + $todaysdate) }
    $filePath += ("." + $extension)

    return $filePath
}

Function ParseParameters()
{
    Param (
        [String] $inputFile,
        [String] $logFilePath,
        [Bool] $whatIf
    )

    $logFilePath = AssembleFilePath $logFilePath "DynamicAzureADGoupCreation-Log" "log" $true

    $paramObj = New-Object PSObject
    Add-Member -InputObject $paramObj -MemberType NoteProperty -Name InputFile -Value $inputFile
    Add-Member -InputObject $paramObj -MemberType NoteProperty -Name LogFilePath -Value $logFilePath
    Add-Member -InputObject $paramObj -MemberType NoteProperty -Name WhatIf -Value $whatIf

    Return $paramObj
}

Function CheckParameters()
{
    Param (
        [String] $inputFile,
        [String] $logFilePath,
        [Bool] $whatIf
    )

    # check for writability to log file
    # do this because Start-Transcript does not throw an exception if the path does not exist
    try 
    {
        Out-File -FilePath $logFilePath -InputObject "Starting Process" -Encoding ASCII -ErrorAction Stop
    }
    catch 
    {
        throw "Could not create the log file: $($LogFilePath)."
    }

    # check if the input file exists
    $CheckInputFile = Test-Path -Path $inputFile
    if ($CheckInputFile -eq $false)
    {
        throw "Could not find the specified input file: $($inputFile)."
    }
}

Function EnsureModuleInstalled()
{
    try
    {
        # install the preview Microsoft Graph module if it's not already installed
        if (Get-Module -ListAvailable -Name Microsoft.Graph) 
        {
            Write-Host "Microsoft Graph Module Already Installed" -ForegroundColor Green
        } 
        else 
        {
            Write-Host "Microsoft Graph Module Not Installed. Installing........." -ForegroundColor Red
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -AllowClobber -Force 
            Write-Host "Microsoft Graph Module Installed" -ForegroundColor Green
        }

        # and import the Preview module to use it
        Import-Module Microsoft.Graph.Groups
    }
    catch
    {
        throw "Could not install/import the Microsoft Graph module." 
    }
}

Function ValidateGroup()
{
    Param (
        [Object] $group
    )

    # validate the data
    if ($group.Action -eq "Update" -and $group.ObjectId.Length -eq 0)
    {
        throw "ObjectId is blank for an Update action in row $($currentRow.ToString()) of the input file."
    }
    if ($group.DisplayName.Trim().Length -eq 0) 
    {
        throw "DisplayName is blank in row $($currentRow.ToString()) of the input file."
    }
   if ($group.Owner.Length -eq 0) 
    {
        throw "Owner is blank in row $($currentRow.ToString()) of the input file."
    }
    if ($group.Pbl.Length -eq 0) 
    {
        throw "PBL is blank in row $($currentRow.ToString()) of the input file."
    }
    if ($group.Rule.Length -eq 0) 
    {
        throw "Rule is blank in row $($currentRow.ToString()) of the input file."
    }

    # if we make it this far, all good

}

Function OwnerStringToArray()
{
    Param (
        [String] $ownerString
    )

    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $ownerArray = $ownerString.Split(";", $option)

    return $ownerArray
}

Function CompareOwners()
{
    param (
        [string] $newOwners,
        [Object[]] $existingOwners
    )

    $ownersToAdd = @()
    $ownersToRemove = @()

    # new owners comes from the input file as a semicolon separated set of values, so translate to an array
    $newOwnerArray = OwnerStringToArray $newOwners

    # loop through each expected value and compare against each actual value, ignoring case
    foreach ($newOwner in $newOwnerArray)
    {
        $found = $false
        foreach ($owner in $existingOwners)
        {
            $existingOwner = Get-MgUser -UserId $owner.Id | Select-Object -ExpandProperty UserPrincipalName
            if ($existingOwner.Trim().ToLower() -eq $newOwner.Trim().ToLower())
            {
                $found = $true
                break
            }
        }

        if ($found -eq $false)
        {
            # the new owner is not an existing owner, so add it to the lsit of owners to add
            $ownersToAdd += $newOwner
        }
    }

    # loop through each expected value and compare against each actual value, ignoring case
    foreach ($owner in $existingOwners)
    {
        $existingOwner = Get-MgUser -UserId $owner.Id | Select-Object -ExpandProperty UserPrincipalName
        $found = $false
        foreach ($newOwner in $newOwnerArray)
        {
            if ($existingOwner.Trim().ToLower() -eq $newOwner.Trim().ToLower())
            {
                $found = $true
                break
            }
        }

        if ($found -eq $false)
        {
            # the existing owner is not an new owner array, so add it to the list of owners to remove
            $ownersToRemove += $existingOwner
        }
    }

    $resultObj = New-Object PSObject
    Add-Member -InputObject $resultObj -MemberType NoteProperty -Name OwnersToAdd -Value $ownersToAdd
    Add-Member -InputObject $resultObj -MemberType NoteProperty -Name OwnersToRemove -Value $ownersToRemove

    # all good
    return $resultObj
}

Function GetExistingGroupByDisplayName()
{
    Param (
        [String] $displayName
    )

    $existingGroup = Get-MgGroup -Filter "DisplayName eq '$($displayname)'"

    if ($null -ne $existingGroup) {
        if ($existingGroup -is [array]) {
            # more than one group has the speicified Display name, so just return the first one
            $existingGroup = $existingGroup[0]	
        }
    }
    return $existingGroup
}

Function GetExistingGroupByObjectId()
{
    Param (
        [String] $objectId
    )

    return Get-MgGroup -GroupId $objectId | Select-Object *
}

Function GetExistingGroup()
{
    Param (
        [Object] $group
    )

    if ($group.Action -eq "Update")
    {
        $existingGroup = GetExistingGroupByObjectId $group.ObjectId
    }
    elseif ($group.Action -eq "Add")
    {
        $existingGroup = GetExistingGroupByDisplayName $group.DisplayName
    }
    else
    {
        $existingGroup = $null
    }

    return $existingGroup
}

Function AddOwner()
{
    Param (
        [Object] $group,
        [String] $owner
    )

    Write-Host "    Adding owner $($owner) to group $($group.DisplayName)" -ForegroundColor Green

    # get the user object from Azure AD
    $ownerObject = Get-MgUser -UserId "$($owner.Trim())"

    if ($null -eq $ownerObject)
    {
        Write-Host "    Error: $($owner) was not found." -ForegroundColor Red
    }
    else
    {
        # set the owner of the new group
        $odataId = "https://graph.microsoft.com/v1.0/users/" + $ownerObject.Id
        New-MgGroupOwnerByRef -GroupId $group.Id.Trim() -BodyParameter @{"@odata.id" = $odataId }

        Write-Host "    Succesfully added owner $($owner) to group $($group.DisplayName)" -ForegroundColor Green
    }
}

Function AddOwners()
{
    Param (
        [Object] $group,
        [String[]] $owners
    )

    foreach ($owner in $owners)
    {
        AddOwner $group $owner
    }
}

Function RemoveOwner()
{
    Param (
        [Object] $group,
        [String] $owner
    )

    Write-Host "    Removing owner $($owner) to group $($group.DisplayName)" -ForegroundColor Green

    # get the user object from Azure AD
    $ownerObject = Get-MgUser -UserId $owner.Trim()

    if ($null -eq $ownerObject)
    {
        Write-Host "    Error: $($owner) was not found." -ForegroundColor Red
    }
    else
    {
        # remove the owner of the new group
        Remove-MgGroupOwnerByRef -GroupId $($group.ObjectID) -DirectoryObjectId $($ownerobject.Id)

        Write-Host "    Succesfully removed owner $($owner) from group $($group.DisplayName)" -ForegroundColor Green
    }
}

Function RemoveOwners()
{
    Param (
        [Object] $group,
        [String[]] $owners
    )

    foreach ($owner in $owners)
    {
        RemoveOwner $group $owner
    }
}

Function AddGroup()
{
    Param (
        [Object] $group
    )

    Write-Host "Adding group $($group.DisplayName) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green

    # check to see if group already exists with the specified DisplayName.
    $existingGroup = GetExistingGroup $group

    if ($null -ne $existingGroup)
    {
        # group with the DisplayName already exists, so don't attempt to create another group and write to the output
        throw "Group with DisplayName $($group.DisplayName) already exists."
    }
    else
    {
        # create the new group
        $newGroup = New-MgGroup `
            -DisplayName "$($group.DisplayName)" `
            -Description "$($group.Pbl)" `
            -MembershipRule "$($group.Rule)" `
            -GroupTypes "DynamicMembership" `
            -MailEnabled:$False `
            -MailNickname "MailNickname" `
            -SecurityEnabled:$True `
            -MembershipRuleProcessingState "On" `
            -IsAssignableToRole:$False `
            -Visibility "Public" `
            -ErrorAction Stop

        Write-Host "  Succesfully created the dynamic AD Group: $($group.DisplayName) as ObjectId $($newGroup.Id) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green
    }

    return $newGroup
}

Function UpdateGroup()
{
    Param (
        [Object] $group
    )

    $updatesMade = $false

    Write-Host "Updating Group ID $($group.ObjectId) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green

    # check to see if group already exists with the specified DisplayName.
    $existingGroup = GetExistingGroup $group

    if ($null -eq $existingGroup)
    {
        # group with the ObjectId does not exist
        throw "Group with ID $($group.ObjectId) does not exist."
    }
    else
    {
        if ($group.DisplayName -ne $existingGroup.DisplayName)
        {
            # if we're changing the display name, make sure the new display name does not exist already
            $checkGroupName = GetExistingGroupByDisplayName $group.DisplayName
            if ($null -ne $checkGroupName)
            {
                # a group with the new DisplayName already exists
                throw "Can't update the DisplayName of Group ID $($group.ObjectId) because another group ($($checkGroupName.Id)) already has the DisplayName."
            }
        }

        # compare values to see if we need to change anything
        $params = @{}
        if ($group.DisplayName -ne $existingGroup.DisplayName) {
            $params.Add("DisplayName",$group.DisplayName)
        }
        if ($group.Pbl -ne $existingGroup.Description) {
            $params.Add("Description",$group.Pbl)
        }
        if ($group.Rule -ne $existingGroup.MembershipRule) {
            $params.Add("MembershipRule",$group.Rule)
        }
<#        {
            # update the existing group
            $params = @{
				"DisplayName"    = "$($group.DisplayName)"
				"Description"    = "$($group.Pbl)"
				"MembershipRule" = "$($group.Rule)"
				"MailNickname"   = "MailNickname"
				"Visibility"     = "Public"
			}
#>
        if ($params.Count -gt 0) {
			Update-MgGroup -GroupId $group.ObjectId -BodyParameter $params -ErrorAction Stop
            $updatesMade = $true
        }

        # check to see if we need to change the owners
        $ownerObjs = Get-MgGroupOwner -GroupId $group.ObjectId
        $ownerResults = CompareOwners $group.Owner $ownerObjs

        if ($ownerResults.OwnersToAdd.Length -gt 0)
        {
            AddOwners $existingGroup $ownerResults.OwnersToAdd
            $updatesMade = $true
        }
        if ($ownerResults.OwnersToRemove.Length -gt 0)
        { 
            RemoveOwners $group $ownerResults.OwnersToRemove
            $updatesMade = $true
        }

        if ($updatesMade -eq $true)
        {
            Write-Host "  Succesfully updated the group $($group.DisplayName) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green
        }
        else
        {
            Write-Host "  No updates needed for group $($group.DisplayName) from row $($currentRow.ToString()) of the input file." -ForegroundColor Green
        }
    }
}

Function ProcessGroups()
{
    Param (
        [Object[]] $groupsToProcess,
        [Object] $params
    )

    $currentRow = 1
    $totalGroupsProcessed = 0
    $totalGroupsAdded = 0
    $totalGroupsUpdated = 0
    $totalGroupsFailed = 0

    # loop through each item from the input file and create the groups
    foreach ($group in $groupsToProcess)
    {
        $currentRow += 1
        $totalGroupsProcessed += 1
        try
        {
            ValidateGroup $group

            if ($group.Action -eq "Add")
            {
                $addedGroup = AddGroup $group    
                $totalGroupsAdded += 1

                $ownerArray = OwnerStringToArray $group.Owner
                AddOwners $addedGroup $ownerArray
            }
            elseif ($group.Action -eq "Update")
            {
                UpdateGroup $group
                $totalGroupsUpdated += 1
            }
            else
            {
                throw "Invalid Action in row $($currentRow.ToString()) of the input file."
            }
        }
        catch
        {
            # log the error and move on to the next item
            Write-Host "Error processing the dynamic AD Group: $($group.DisplayName) in row $($currentRow.ToString()) of the input file." -ForegroundColor Red
            Write-Host $PSItem.Exception.Message -ForegroundColor Red
            Write-Host $PSItem.InvocationInfo.PositionMessage -ForegroundColor Red

            $totalGroupsFailed += 1
        }
    }

    $results = New-Object PSObject
    Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsProcessed -Value $totalGroupsProcessed
    Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsAdded -Value $totalGroupsAdded
    Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsUpdated -Value $totalGroupsUpdated
    Add-Member -InputObject $results -MemberType NoteProperty -Name TotalGroupsFailed -Value $totalGroupsFailed

    Return $results
}

Function OutputContextInfo()
{
    Param (
        [Object[]] $context
    )

    Write-Host "Connected Account: $($context.Account)" -ForegroundColor Yellow
    Write-Host "Connected TenantDomain: $($context.TenantDomain)" -ForegroundColor Yellow
    Write-Host "Connected Environment: $($context.Environment)" -ForegroundColor Yellow
    Write-Host "Connected TenantId: $($context.TenantId)" -ForegroundColor Yellow
}

Function Main()
{
    $scriptStartTime = Get-Date

    $params = ParseParameters $InputFile $LogFilePath $WhatIf
    CheckParameters $params.InputFile $params.LogFilePath $params.WhatIf
    
    Start-Transcript -Path $params.LogFilePath

    Write-Host "Script Version: $($scriptVersion)" -ForegroundColor yellow
    Write-Host "WhatIf: $($params.WhatIf.ToString())" -ForegroundColor yellow

    # make sure the necessary module is installed
    EnsureModuleInstalled

    # connect to Azure AD - this will prompt for credentials
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"

    # output conext info
#    OutputContextInfo $context

    # read the input file
    $groupsToProcess = Import-Csv -Path $params.InputFile

    # main processing
    $processedGroups = ProcessGroups $groupsToProcess

    # write summary information
    $scriptDuration = (New-TimeSpan -Start $scriptStartTime -End (Get-Date)).ToString("dd' days 'hh' hours 'mm' minutes 'ss' seconds'")

    Write-Host "**********************" -ForegroundColor Yellow
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
$scriptVersion = "Minimal Changes"

Main