function Connect-SCCM {
    <#
    .SYNOPSIS
        Connects to the SCCM site. Based on default Connect with ISE Configuration Manager Console option.
    
    .DESCRIPTION
        This function connects to the SCCM site by importing the ConfigurationManager module,
        creating a PSDrive for the site, and setting the location to the site code.
    
    .PARAMETER SiteCode
        The SCCM site code (three characters).
    
    .PARAMETER ProviderMachineName
        The SCCM provider machine name (SCCM server name).
    
    .PARAMETER InitParams
        Additional parameters to pass to the Import-Module and New-PSDrive cmdlets.
    
    .EXAMPLE
        Connect-SCCM -SiteCode "XX1" -ProviderMachineName "server.company.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SiteCode, # Three characters site code
        
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName, # SCCM server name
        
        [Parameter(Mandatory = $false)]
        [hashtable]$InitParams = @{}
    )
    
    # Import the ConfigurationManager.psd1 module if not already loaded
    if ($null -eq (Get-Module ConfigurationManager)) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @InitParams 
    }

    # Connect to the site's drive if it is not already present
    if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @InitParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @InitParams
}

function Test-SCCMConnected {
    <#
    .SYNOPSIS
        Function to ensure that session connected to SCCM and current location is set to SCCM provider path.
    
    .DESCRIPTION
        This function ensures that current session connected to SCCM and location is set to SCCM provider path (SiteCode:\).
        
        If SiteCode or ProviderMachineName are not provided, the function will prompt the user
        to enter these values interactively.
    
    .PARAMETER SiteCode
        The SCCM site code (three characters).
    
    .PARAMETER ProviderMachineName
        The SCCM provider machine name (SCCM server name).
    
    .EXAMPLE
        Test-SCCMConnected -SiteCode "XX1" -ProviderMachineName "server.company.com"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$SiteCode,
    
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName
    )

    # Prompt for SiteCode if not provided
    if ([string]::IsNullOrEmpty($SiteCode)) {
        $SiteCode = Read-Host -Prompt "Enter the SCCM site code (three characters)"
    }
    
    # Prompt for ProviderMachineName if not provided
    if ([string]::IsNullOrEmpty($ProviderMachineName)) {
        $ProviderMachineName = Read-Host -Prompt "Enter the SCCM provider machine name (server name)"
    }

    if ([bool] (Get-Module ConfigurationManager) -ne $true) {
        Connect-SCCM -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName
    }
    if ((Get-Location).Path -ne "$($SiteCode):\") {
        Set-Location "$($SiteCode):\"
    }
}

########### Functions to export Collections from SCCM to clixml file

# Function to convert relative collection path to full path
function Convert-CollectionRelativeToFullPath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $CollectionID,
        
        [Parameter(Mandatory = $true)]
        [string] $RelativePath,
        
        [Parameter(Mandatory = $false)]
        [string]$SiteCode,
        
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName
    )
    
    # Connect to SCCM if ConfigurationManager module is not loaded
    Test-SCCMConnected -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName
    
    $Collection = Get-CMCollection -Id $CollectionID
    $SiteCode = (Get-CMSite).SiteCode
    $FullPath = "$($SiteCode):\"
    switch ($Collection.CollectionType) {
        # UserCollection
        1 {
            $FullPath += "UserCollection"
        }
        # DeviceCollection
        2 {
            $FullPath += "DeviceCollection"
        }
        Default {
            throw "CollectionType: $_ unknown!"
        }
    }
    $FullPath += $RelativePath
    return $FullPath
}

# Function to get collection path from CIM
function Get-xCMObjectLocation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $InstanceKey,
        
        [Parameter(Mandatory = $false)]
        [string] $SiteCode = (Get-CMSite).SiteCode,
        
        [Parameter(Mandatory = $false)]
        [string] $SiteServer = ((Get-CMSiteSystemServer -SiteCode $SiteCode | Where-Object { $_.Type -eq "2" }) | Select-Object -ExpandProperty NetworkOSPath).TrimStart("\"),
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("RelativePath", "FullPath", IgnoreCase = $true)]
        [string] $Format = 'RelativePath',
        
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName
    )

    # Powered by:
    # https://petervanderwoude.nl/post/get-the-folder-location-of-an-object-in-configmgr-2012-via-powershell/
    # https://github.com/pvanderwoude/blog/blob/main/Get-ObjectLocation.ps1

    $ObjectContainerItem = Get-CimInstance -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServer -Query "SELECT * FROM SMS_ObjectContainerItem WHERE InstanceKey='$InstanceKey'"
    $ContainerNode = Get-CimInstance -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServer -Query "SELECT * FROM SMS_ObjectContainerNode WHERE ContainerNodeID='$($ObjectContainerItem.ContainerNodeID)'"

    if ($null -ne $ContainerNode) {
        $ObjectFolder = $ContainerNode.Name
        if ($ContainerNode.ParentContainerNodeID -eq 0) {
            $ParentFolder = $false
        }
        else {
            $ParentFolder = $true
            $ParentContainerNodeID = $ContainerNode.ParentContainerNodeID
        }
        while ($ParentFolder -eq $true) {
            $ParentContainerNode = Get-CimInstance -Namespace root/SMS/site_$($SiteCode) -ComputerName $SiteServer -Query "SELECT * FROM SMS_ObjectContainerNode WHERE ContainerNodeID = '$ParentContainerNodeID'"
            $ObjectFolder = $ParentContainerNode.Name + "\" + $ObjectFolder
            if ($ParentContainerNode.ParentContainerNodeID -eq 0) {
                $ParentFolder = $false
            }
            else {
                $ParentContainerNodeID = $ParentContainerNode.ParentContainerNodeID
            }
        }
        $ObjectFolder = "\" + $ObjectFolder
    }
    else {
        $ObjectFolder = "\"
    }
    # Convert relative path to full path if required
    if ($Format -eq "FullPath") {
        $ObjectFolder = Convert-CollectionRelativeToFullPath -CollectionID $InstanceKey -RelativePath $ObjectFolder
    }
    return $ObjectFolder
}

# Function to gather metadata about collections to recreate them later
function Get-CMCollectionsMetadata {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$SiteCode,
        
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName
    )
    
    # Connect to SCCM if ConfigurationManager module is not loaded
    Test-SCCMConnected -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName

    $AllNonBuiltInCollections = Get-CMCollection | Where-Object { $_.IsBuiltIn -eq $false }
    
    $CollectionsToExport = @()
    foreach ($Collection in $AllNonBuiltInCollections) {
        $PrunedCollection = $Collection | Select-Object -Property SmsProviderObjectPath, CollectionID, CollectionRules, CollectionType, Comment, LimitToCollectionID, LimitToCollectionName, MemberClassName, Name, RefreshSchedule, RefreshType, ReplicateToSubSites, ManagedObject, OverridingObjectClass, DisplayString, DisplayDescription, HelpTopic, ObjectClass, UniqueIdentifier, ParentResultObject, GlobalDisplayString
        $PrunedCollection | Add-Member -MemberType NoteProperty -Name ObjectPath -Value $(Get-xCMObjectLocation -InstanceKey $Collection.CollectionID -Format FullPath)
        $CollectionsToExport += $PrunedCollection
    }
    
    return $CollectionsToExport

}

function Export-CMCollectionsMetadata {
    [CmdletBinding()]
    param (
        # Path to export .clixml file
        [Parameter(Mandatory = $true)]
        [string]
        $Path,
        
        [Parameter(Mandatory = $false)]
        [string]$SiteCode,
        
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName
    )
    
    # Connect to SCCM if ConfigurationManager module is not loaded
    Test-SCCMConnected -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName
        
    $col = Get-CMCollectionsMetadata
    $col | Export-Clixml -Path $Path -Encoding UTF8
}

<#
.SYNOPSIS
    Ensures that a folder exists in SCCM, creating it and its parent folders if necessary.

.DESCRIPTION
    The Test-CMFolderExists function checks if a specified folder exists in SCCM.
    If the folder doesn't exist, it creates the folder and any necessary parent folders
    recursively. This function automatically determines the collection type from the path.

.PARAMETER FolderPath
    The full path of the folder to check/create. This should be in the format:
    "SiteCode:\FolderType\Path\To\Folder" where FolderType is either UserCollection or DeviceCollection.

.EXAMPLE
    Test-CMFolderExists -FolderPath "ABC:\DeviceCollection\Folder1\Folder2"

    Checks if the folder "Folder2" exists in the path "ABC:\DeviceCollection\Folder1".
    If it doesn't exist, creates the folder and any necessary parent folders.

.EXAMPLE
    Test-CMFolderExists -FolderPath "ABC:\UserCollection\Department\HR"

    Checks if the folder "HR" exists in the path "ABC:\UserCollection\Department".
    If it doesn't exist, creates the folder and any necessary parent folders.
#>
function Test-CMFolderExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )
    
    # Parse the folder path
    if ($FolderPath -match '^([A-Za-z0-9]{3}):\\(UserCollection|DeviceCollection)(.*)$') {
        $SiteCode = $matches[1]
        $RootFolder = $matches[2]
        $RelativePath = $matches[3]
        
        # Determine the collection type and ObjectTypeName based on the root folder
        $CollectionType = if ($RootFolder -eq "UserCollection") { 1 } else { 2 }
        $ObjectTypeName = if ($CollectionType -eq 1) { "SMS_Collection_User" } else { "SMS_Collection_Device" }
    }
    else {
        Write-Error "Invalid folder path format: $FolderPath. Expected format: SiteCode:\FolderType\Path\To\Folder"
        return $false
    }
    
    # If the path is just the root folder, it always exists
    if ([string]::IsNullOrEmpty($RelativePath) -or $RelativePath -eq "\") {
        return $true
    }
    
    # Split the relative path into folder names
    $FolderNames = $RelativePath.Trim('\').Split('\')
    
    # Start with the root folder
    $CurrentPath = "$($SiteCode):\$RootFolder"
    
    # Process each folder in the path
    foreach ($FolderName in $FolderNames) {
        # Check if the current folder exists
        $Folder = Get-CMFolder -Name $FolderName -ObjectType $ObjectTypeName -ErrorAction SilentlyContinue
        
        if ($null -eq $Folder) {
            # Folder doesn't exist, create it
            Write-Verbose "Creating folder: $FolderName in $CurrentPath"
            $Folder = New-CMFolder -Name $FolderName -ParentContainerNode $CurrentPath
            
            if ($null -eq $Folder) {
                Write-Error "Failed to create folder: $FolderName in $CurrentPath"
                return $false
            }
        }
        
        # Update current path and parent folder ID for the next iteration
        $CurrentPath = "$CurrentPath\$FolderName"
    }
    
    return $true
}

<#
.SYNOPSIS
    Imports SCCM collections from a metadata file exported by Export-CMCollectionsMetadata.

.DESCRIPTION
    The Import-CMCollectionsFromMetadata function imports SCCM collections from a clixml file
    that was previously exported using the Export-CMCollectionsMetadata function. It creates
    new collections in the current SCCM instance based on the metadata in the file.

    The function automatically detects and replaces the site code from the source SCCM instance
    with the site code of the current SCCM instance in all relevant fields:
    - CollectionID
    - LimitToCollectionID
    - ObjectPath (folder path)
    - IncludeCollectionID in include collection rules
    - ExcludeCollectionID in exclude collection rules
    - Query expressions that might contain references to the site code

    This allows for seamless migration of collections between different SCCM instances with
    different site codes.

.PARAMETER Path
    The path to the clixml file containing the exported collection metadata.

.PARAMETER NameSuffix
    An optional suffix to append to the names of the imported collections.
    This is useful to distinguish imported collections from existing ones.

.PARAMETER Force
    If specified, existing collections with the same name will be removed before
    creating new ones. Without this switch, collections that already exist will be skipped.

.PARAMETER WhatIf
    If specified, the function will show what would happen if it ran without
    actually making any changes.

.EXAMPLE
    Import-CMCollectionsFromMetadata -Path "C:\Temp\CollectionsExport.clixml"

    Imports all collections from the specified file, using their original names.

.EXAMPLE
    Import-CMCollectionsFromMetadata -Path "C:\Temp\CollectionsExport.clixml" -NameSuffix "_Migrated"

    Imports all collections from the specified file, appending "_Migrated" to their names.

.EXAMPLE
    Import-CMCollectionsFromMetadata -Path "C:\Temp\CollectionsExport.clixml" -Force

    Imports all collections from the specified file, overwriting any existing collections with the same names.

.EXAMPLE
    Import-CMCollectionsFromMetadata -Path "C:\Temp\CollectionsExport.clixml" -WhatIf

    Shows what would happen if the function ran without actually making any changes.
#>
function Import-CMCollectionsFromMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path,
        
        [Parameter(Mandatory = $false)]
        [string] $NameSuffix = "",
        
        [Parameter(Mandatory = $false)]
        [switch] $Force,
        
        [Parameter(Mandatory = $false)]
        [switch] $WhatIf,
        
        [Parameter(Mandatory = $false)]
        [string]$SiteCode,
        
        [Parameter(Mandatory = $false)]
        [string]$ProviderMachineName
    )
    
    # Connect to SCCM if ConfigurationManager module is not loaded
    Test-SCCMConnected -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName
    
    # Check if the file exists
    if (-not (Test-Path -Path $Path)) {
        Write-Error "The specified file does not exist: $Path"
        return
    }
    
    try {
        # Import the collection metadata from the clixml file
        Write-Host "Importing collection metadata from $Path..." -ForegroundColor Cyan
        $ImportedCollections = Import-Clixml -Path $Path
        
        # Get the current site code
        $CurrentSiteCode = (Get-CMSite).SiteCode
        
        # Determine the source site code from the first collection's ID
        if ($ImportedCollections.Count -gt 0) {
            # Extract the site code (typically the first 3 characters of the CollectionID)
            $SourceSiteCode = $null
            
            # Try to extract from CollectionID
            if (-not [string]::IsNullOrEmpty($ImportedCollections[0].CollectionID)) {
                # Most SCCM collection IDs are in the format "XXX00001" where XXX is the site code
                if ($ImportedCollections[0].CollectionID -match '^([A-Za-z0-9]{3})') {
                    $SourceSiteCode = $matches[1]
                }
            }
            
            # If we couldn't extract from CollectionID, try from ObjectPath
            if ([string]::IsNullOrEmpty($SourceSiteCode) -and -not [string]::IsNullOrEmpty($ImportedCollections[0].ObjectPath)) {
                # ObjectPath is typically in the format "XXX:\DeviceCollection\..."
                if ($ImportedCollections[0].ObjectPath -match '^([A-Za-z0-9]{3}):') {
                    $SourceSiteCode = $matches[1]
                }
            }
            
            if (-not [string]::IsNullOrEmpty($SourceSiteCode)) {
                Write-Host "Detected source site code: $SourceSiteCode" -ForegroundColor Cyan
                
                # Replace site code in all collections and their rules
                Write-Host "Replacing source site code '$SourceSiteCode' with current site code '$CurrentSiteCode' in all collections..." -ForegroundColor Cyan
                
                foreach ($Collection in $ImportedCollections) {
                    # Replace site code in CollectionID
                    if (-not [string]::IsNullOrEmpty($Collection.CollectionID)) {
                        $Collection.CollectionID = $Collection.CollectionID -replace "^$SourceSiteCode", $CurrentSiteCode
                    }
                    
                    # Replace site code in LimitToCollectionID
                    if (-not [string]::IsNullOrEmpty($Collection.LimitToCollectionID)) {
                        $Collection.LimitToCollectionID = $Collection.LimitToCollectionID -replace "^$SourceSiteCode", $CurrentSiteCode
                    }
                    
                    # Replace site code in ObjectPath
                    if (-not [string]::IsNullOrEmpty($Collection.ObjectPath)) {
                        $Collection.ObjectPath = $Collection.ObjectPath -replace "^$($SourceSiteCode):", "$($CurrentSiteCode):"
                    }
                    
                    # Replace site code in collection rules
                    if ($Collection.CollectionRules) {
                        foreach ($rule in $Collection.CollectionRules) {
                            # Replace site code in IncludeCollectionID
                            if ($rule.SmsProviderObjectPath -eq "SMS_CollectionRuleIncludeCollection" -and -not [string]::IsNullOrEmpty($rule.IncludeCollectionID)) {
                                $rule.IncludeCollectionID = $rule.IncludeCollectionID -replace "^$SourceSiteCode", $CurrentSiteCode
                            }
                            
                            # Replace site code in ExcludeCollectionID
                            if ($rule.SmsProviderObjectPath -eq "SMS_CollectionRuleExcludeCollection" -and -not [string]::IsNullOrEmpty($rule.ExcludeCollectionID)) {
                                $rule.ExcludeCollectionID = $rule.ExcludeCollectionID -replace "^$SourceSiteCode", $CurrentSiteCode
                            }
                            
                            # Replace site code in query expressions if they contain references to the site code
                            if ($rule.SmsProviderObjectPath -eq "SMS_CollectionRuleQuery" -and -not [string]::IsNullOrEmpty($rule.QueryExpression)) {
                                # This is a simplistic approach - complex queries might need more sophisticated parsing
                                $rule.QueryExpression = $rule.QueryExpression -replace $SourceSiteCode, $CurrentSiteCode
                            }
                        }
                    }
                }
            }
            else {
                Write-Warning "Could not detect source site code from the imported collections. Site code replacement will not be performed."
            }
        }
        
        # Define built-in collection names (e.g. "All Systems","All Users","All User Groups","All Users and User Groups","All Custom Resources","All Unknown Computers","All Provisioning Devices","Co-management Eligible Devices","All Mobile Devices","All Desktop and Server Clients")
        $BuiltInCollectionNames = Get-CMCollection | Where-Object { $_.IsBuiltIn -eq $true } | Select-Object -ExpandProperty name
        
        # Create a dictionary to store created collections
        $CreatedCollections = @{}
        
        # PHASE 1: Create collections without rules in the correct order
        Write-Host "PHASE 1: Creating collections without rules..." -ForegroundColor Green
        
        # First, remove existing collections if Force is specified
        if ($Force -and -not $WhatIf) {
            foreach ($Collection in $ImportedCollections) {
                $CollectionName = $Collection.Name + $NameSuffix
                $ExistingCollection = Get-CMCollection -Name $CollectionName -ErrorAction SilentlyContinue
                
                if ($ExistingCollection) {
                    Write-Host "Removing existing collection '$CollectionName'..." -ForegroundColor Yellow
                    Remove-CMCollection -Name $CollectionName -Force
                }
            }
        }
        
        # Step 1: Create collections that have built-in collections as their limiting collection
        Write-Host "Step 1: Creating collections with built-in limiting collections..." -ForegroundColor Cyan
        foreach ($Collection in $ImportedCollections) {
            $CollectionName = $Collection.Name + $NameSuffix
            
            # Skip if collection already exists and Force is not specified
            $ExistingCollection = Get-CMCollection -Name $CollectionName -ErrorAction SilentlyContinue
            if ($ExistingCollection -and -not $Force) {
                Write-Warning "Collection '$CollectionName' already exists. Use -Force to overwrite."
                $CreatedCollections[$CollectionName] = $ExistingCollection
                continue
            }
            elseif ($ExistingCollection -and $Force -and $WhatIf) {
                Write-Host "WhatIf: Would remove existing collection '$CollectionName'" -ForegroundColor Yellow
                continue
            }
            
            # Check if the limiting collection is a built-in collection
            if ($BuiltInCollectionNames -contains $Collection.LimitToCollectionName) {
                if (-not $WhatIf) {
                    Write-Host "Creating collection '$CollectionName' with built-in limiting collection '$($Collection.LimitToCollectionName)'..." -ForegroundColor Green
                    
                    # Create the collection
                    $SplatParams = @{
                        Name                  = $CollectionName
                        Comment               = $Collection.Comment
                        CollectionType        = $Collection.CollectionType # 1 = UserCollection; 2 = DeviceCollection
                        LimitToCollectionName = $Collection.LimitToCollectionName
                    }
                    
                    $NewCollection = New-CMCollection @SplatParams
                    $CreatedCollections[$CollectionName] = $NewCollection
                    
                    # Ensure the folder exists before moving the collection
                    try {
                        $FolderPath = $Collection.ObjectPath
                        
                        # Check if the folder path is just the root folder (UserCollection or DeviceCollection)
                        if ($FolderPath -match '^[A-Za-z0-9]{3}:\\(UserCollection|DeviceCollection)(\\)?$') {
                            # Skip moving if the target is the root folder - the collection is already there
                            Write-Host "Collection is already in the root folder: $FolderPath" -ForegroundColor Cyan
                        }
                        else {
                            # Ensure the folder exists and move the collection
                            Write-Host "Ensuring folder exists: $FolderPath" -ForegroundColor Cyan
                            $FolderExists = Test-CMFolderExists -FolderPath $FolderPath
                            
                            if ($FolderExists) {
                                Write-Host "Moving collection to folder: $FolderPath" -ForegroundColor Cyan
                                Move-CMObject -InputObject $NewCollection -FolderPath $FolderPath -ErrorAction SilentlyContinue
                            }
                            else {
                                Write-Warning "Failed to ensure folder exists: $FolderPath. Collection will be created in the root folder."
                            }
                        }
                    }
                    catch {
                        Write-Warning "Failed to move collection to folder: $FolderPath. Error: $_"
                    }
                    
                    # Set refresh schedule if available
                    if ($Collection.RefreshSchedule) {
                        try {
                            Write-Host "Setting refresh schedule for collection '$CollectionName'..." -ForegroundColor Cyan
                            
                            # Extract schedule parameters from the RefreshSchedule object
                            $scheduleData = $Collection.RefreshSchedule
                            
                            # Create a new schedule object using New-CMSchedule based on the parameters
                            $newSchedule = $null
                            
                            if ($scheduleData.DaySpan -gt 0) {
                                $newSchedule = New-CMSchedule -RecurInterval Days -RecurCount $scheduleData.DaySpan
                            }
                            elseif ($scheduleData.HourSpan -gt 0) {
                                $newSchedule = New-CMSchedule -RecurInterval Hours -RecurCount $scheduleData.HourSpan
                            }
                            elseif ($scheduleData.MinuteSpan -gt 0) {
                                $newSchedule = New-CMSchedule -RecurInterval Minutes -RecurCount $scheduleData.MinuteSpan
                            }
                            
                            # Set the collection's refresh schedule and type
                            if ($newSchedule) {
                                #TODO: if collection have no `Use incremental update for this collection` or `Schedule a full update on this collection` checks, then warning will be shown and default refresh schedule would be set to collection
                                Set-CMCollection -InputObject $NewCollection -RefreshSchedule $newSchedule -RefreshType $Collection.RefreshType -ErrorAction Stop
                                Write-Host "Successfully set refresh schedule and type for collection '$CollectionName'" -ForegroundColor Green
                            }
                            else {
                                Write-Warning "Could not create a valid schedule from the imported data for collection '$CollectionName'"
                            }
                        }
                        catch {
                            Write-Warning "Failed to set refresh schedule for collection '$CollectionName'. Error: $_"
                        }
                    }
                }
                else {
                    Write-Host "WhatIf: Would create collection '$CollectionName' with built-in limiting collection '$($Collection.LimitToCollectionName)'" -ForegroundColor Green
                }
            }
        }
        
        # Step 2: Create remaining collections in iterations
        Write-Host "Step 2: Creating remaining collections in iterations..." -ForegroundColor Cyan
        $RemainingCollections = $ImportedCollections | Where-Object { 
            $CollectionName = $_.Name + $NameSuffix
            -not $CreatedCollections.ContainsKey($CollectionName) -and 
            $BuiltInCollectionNames -notcontains $_.LimitToCollectionName
        }
        
        $NoProgressCount = 0
        $MaxNoProgressIterations = 3
        
        while ($RemainingCollections.Count -gt 0 -and $NoProgressCount -lt $MaxNoProgressIterations) {
            $CollectionsCreatedThisIteration = 0
            
            foreach ($Collection in $RemainingCollections) {
                $CollectionName = $Collection.Name + $NameSuffix
                $LimitToCollectionName = $Collection.LimitToCollectionName
                
                # Check if the limiting collection exists (either built-in or already created)
                $LimitingCollectionExists = $BuiltInCollectionNames -contains $LimitToCollectionName -or 
                $CreatedCollections.ContainsKey($LimitToCollectionName) -or
                (Get-CMCollection -Name $LimitToCollectionName -ErrorAction SilentlyContinue)
                
                if ($LimitingCollectionExists) {
                    if (-not $WhatIf) {
                        Write-Host "Creating collection '$CollectionName' with limiting collection '$LimitToCollectionName'..." -ForegroundColor Green
                        
                        # Create the collection
                        $SplatParams = @{
                            Name                  = $CollectionName
                            Comment               = $Collection.Comment
                            CollectionType        = $Collection.CollectionType
                            LimitToCollectionName = $LimitToCollectionName
                        }
                        
                        $NewCollection = New-CMCollection @SplatParams
                        $CreatedCollections[$CollectionName] = $NewCollection
                        $CollectionsCreatedThisIteration++
                        
                        # Ensure the folder exists before moving the collection
                        try {
                            $FolderPath = $Collection.ObjectPath
                            
                            # Check if the folder path is just the root folder (UserCollection or DeviceCollection)
                            if ($FolderPath -match '^[A-Za-z0-9]{3}:\\(UserCollection|DeviceCollection)(\\)?$') {
                                # Skip moving if the target is the root folder - the collection is already there
                                Write-Host "Collection is already in the root folder: $FolderPath" -ForegroundColor Cyan
                            }
                            else {
                                # Ensure the folder exists and move the collection
                                Write-Host "Ensuring folder exists: $FolderPath" -ForegroundColor Cyan
                                $FolderExists = Test-CMFolderExists -FolderPath $FolderPath
                                
                                if ($FolderExists) {
                                    Write-Host "Moving collection to folder: $FolderPath" -ForegroundColor Cyan
                                    Move-CMObject -InputObject $NewCollection -FolderPath $FolderPath -ErrorAction SilentlyContinue
                                }
                                else {
                                    Write-Warning "Failed to ensure folder exists: $FolderPath. Collection will be created in the root folder."
                                }
                            }
                        }
                        catch {
                            Write-Warning "Failed to move collection to folder: $FolderPath. Error: $_"
                        }
                        
                        # Set refresh schedule if available
                        if ($Collection.RefreshSchedule) {
                            try {
                                Write-Host "Setting refresh schedule for collection '$CollectionName'..." -ForegroundColor Cyan
                                
                                # Extract schedule parameters from the RefreshSchedule object
                                $scheduleData = $Collection.RefreshSchedule
                                
                                # Create a new schedule object using New-CMSchedule based on the parameters
                                $newSchedule = $null
                                
                                if ($scheduleData.DaySpan -gt 0) {
                                    $newSchedule = New-CMSchedule -RecurInterval Days -RecurCount $scheduleData.DaySpan
                                }
                                elseif ($scheduleData.HourSpan -gt 0) {
                                    $newSchedule = New-CMSchedule -RecurInterval Hours -RecurCount $scheduleData.HourSpan
                                }
                                elseif ($scheduleData.MinuteSpan -gt 0) {
                                    $newSchedule = New-CMSchedule -RecurInterval Minutes -RecurCount $scheduleData.MinuteSpan
                                }
                                
                                # Set the collection's refresh schedule and type
                                if ($newSchedule) {
                                    #TODO: if collection have no `Use incremental update for this collection` or `Schedule a full update on this collection` checks, then warning will be shown and default refresh schedule would be set to collection
                                    Set-CMCollection -InputObject $NewCollection -RefreshSchedule $newSchedule -RefreshType $Collection.RefreshType -ErrorAction Stop
                                    Write-Host "Successfully set refresh schedule and type for collection '$CollectionName'" -ForegroundColor Green
                                }
                                else {
                                    Write-Warning "Could not create a valid schedule from the imported data for collection '$CollectionName'"
                                }
                            }
                            catch {
                                Write-Warning "Failed to set refresh schedule for collection '$CollectionName'. Error: $_"
                            }
                        }
                    }
                    else {
                        Write-Host "WhatIf: Would create collection '$CollectionName' with limiting collection '$LimitToCollectionName'" -ForegroundColor Green
                        $CollectionsCreatedThisIteration++
                    }
                }
            }
            
            # Update the remaining collections list
            $RemainingCollections = $RemainingCollections | Where-Object { 
                $CollectionName = $_.Name + $NameSuffix
                -not $CreatedCollections.ContainsKey($CollectionName)
            }
            
            # Check if we made progress in this iteration
            if ($CollectionsCreatedThisIteration -eq 0) {
                $NoProgressCount++
                Write-Warning "No collections created in this iteration. Attempt $NoProgressCount of $MaxNoProgressIterations."
            }
            else {
                $NoProgressCount = 0
                Write-Host "Created $CollectionsCreatedThisIteration collections in this iteration. Remaining: $($RemainingCollections.Count)" -ForegroundColor Cyan
            }
        }
        
        # Check if there are still remaining collections
        if ($RemainingCollections.Count -gt 0) {
            Write-Warning "Could not create $($RemainingCollections.Count) collections due to missing dependencies:"
            foreach ($Collection in $RemainingCollections) {
                Write-Warning "  - '$($Collection.Name)' (depends on '$($Collection.LimitToCollectionName)')"
            }
        }
        
        # PHASE 2: Add collection rules to all created collections
        if (-not $WhatIf) {
            Write-Host "PHASE 2: Adding collection rules..." -ForegroundColor Green
            
            foreach ($Collection in $ImportedCollections) {
                $CollectionName = $Collection.Name + $NameSuffix
                
                # Skip if the collection wasn't created
                if (-not $CreatedCollections.ContainsKey($CollectionName)) {
                    continue
                }
                
                $NewCollection = $CreatedCollections[$CollectionName]
                
                # Add collection rules
                if ($Collection.CollectionRules) {
                    Write-Host "Adding rules to collection '$CollectionName'..." -ForegroundColor Cyan
                    
                    foreach ($rule in $Collection.CollectionRules) {
                        Write-Host "Processing rule: $($rule.RuleName)" -ForegroundColor Magenta
                        
                        switch ($rule.SmsProviderObjectPath) {
                            # Direct Rule
                            "SMS_CollectionRuleDirect" {
                                # Get the resource name from RuleName
                                $resourceName = $rule.RuleName
                                $resourceClassName = $rule.ResourceClassName
                                
                                Write-Host "Direct rule for resource: $resourceName (Type: $resourceClassName)" -ForegroundColor Cyan
                                
                                if ($resourceName) {
                                    try {
                                        $newResourceId = $null
                                        
                                        switch ($resourceClassName) {
                                            "SMS_R_System" {
                                                # Device
                                                # Extract device name from RuleName
                                                $deviceName = $resourceName
                                                Write-Host "Looking up device by name: $deviceName" -ForegroundColor Cyan
                                                
                                                # Find the device in the new environment
                                                $device = Get-CMDevice -Name $deviceName -ErrorAction SilentlyContinue
                                                if ($device) {
                                                    $newResourceId = $device.ResourceID
                                                    Write-Host "Found device '$deviceName' with ResourceID: $newResourceId" -ForegroundColor Green
                                                }
                                                else {
                                                    Write-Warning "Device '$deviceName' not found in the current environment. Skipping rule."
                                                    continue
                                                }
                                            }
                                            "SMS_R_User" {
                                                # User
                                                # Extract user name from RuleName (format: "DOMAIN\samaccountname (Display Name)")
                                                if ($resourceName -match '^([^(]+)') {
                                                    $userName = $matches[1].Trim()
                                                    Write-Host "Looking up user by name: $userName" -ForegroundColor Cyan
                                                    
                                                    # Find the user in the new environment
                                                    $user = Get-CMUser -Name $userName -ErrorAction SilentlyContinue
                                                    if ($user) {
                                                        $newResourceId = $user.ResourceID
                                                        Write-Host "Found user '$userName' with ResourceID: $newResourceId" -ForegroundColor Green
                                                    }
                                                    else {
                                                        Write-Warning "User '$userName' not found in the current environment. Skipping rule."
                                                        continue
                                                    }
                                                }
                                                else {
                                                    Write-Warning "Could not parse user name from RuleName: $resourceName. Skipping rule."
                                                    continue
                                                }
                                            }
                                            Default {
                                                Write-Warning "Unknown resource class name: $resourceClassName. Skipping rule."
                                                continue
                                            }
                                        }
                                        
                                        # Add the direct membership rule with the new resource ID
                                        if ($newResourceId) {
                                            switch ($Collection.CollectionType) {
                                                1 {
                                                    # User Collection
                                                    Add-CMUserCollectionDirectMembershipRule -CollectionId $NewCollection.CollectionID -ResourceId $newResourceId -ErrorAction Stop
                                                }
                                                2 {
                                                    # Device Collection
                                                    Add-CMDeviceCollectionDirectMembershipRule -CollectionId $NewCollection.CollectionID -ResourceId $newResourceId -ErrorAction Stop
                                                }
                                                Default {
                                                    Write-Warning "Unknown collection type: $($Collection.CollectionType)"
                                                }
                                            }
                                            Write-Host "Successfully added direct rule for resource '$resourceName'" -ForegroundColor Green
                                        }
                                    }
                                    catch {
                                        Write-Warning "Failed to add direct rule for resource '$resourceName'. Error: $_"
                                    }
                                }
                                else {
                                    Write-Warning "No resource name found in RuleName for direct rule. Skipping rule."
                                }
                            }
                            
                            # Query Rule
                            "SMS_CollectionRuleQuery" {
                                Write-Host "Query rule: $($rule.QueryExpression)" -ForegroundColor Cyan
                                
                                try {
                                    switch ($Collection.CollectionType) {
                                        1 {
                                            # User Collection
                                            Add-CMUserCollectionQueryMembershipRule -CollectionId $NewCollection.CollectionID -QueryExpression $rule.QueryExpression -RuleName $rule.RuleName -ErrorAction Stop
                                        }
                                        2 {
                                            # Device Collection
                                            Add-CMDeviceCollectionQueryMembershipRule -CollectionId $NewCollection.CollectionID -QueryExpression $rule.QueryExpression -RuleName $rule.RuleName -ErrorAction Stop
                                        }
                                        Default {
                                            Write-Warning "Unknown collection type: $($Collection.CollectionType)"
                                        }
                                    }
                                }
                                catch {
                                    Write-Warning "Failed to add query rule. Error: $_"
                                }
                            }
                            
                            # Include Collections
                            "SMS_CollectionRuleIncludeCollection" {
                                # Get the collection name from the source environment using the old ID
                                $includeCollectionName = $ImportedCollections | 
                                Where-Object { $_.CollectionID -eq $rule.IncludeCollectionID } | 
                                Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
                                
                                if ($includeCollectionName) {
                                    Write-Host "Looking up include collection by name: '$includeCollectionName'" -ForegroundColor Cyan
                                    
                                    # Look up the collection by name in the target environment
                                    $includeCollection = Get-CMCollection -Name $includeCollectionName -ErrorAction SilentlyContinue
                                    
                                    if ($includeCollection) {
                                        try {
                                            switch ($Collection.CollectionType) {
                                                1 {
                                                    # User Collection
                                                    Add-CMUserCollectionIncludeMembershipRule -CollectionId $NewCollection.CollectionID -IncludeCollectionId $includeCollection.CollectionID -ErrorAction Stop
                                                }
                                                2 {
                                                    # Device Collection
                                                    Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $NewCollection.CollectionID -IncludeCollectionId $includeCollection.CollectionID -ErrorAction Stop
                                                }
                                                Default {
                                                    Write-Warning "Unknown collection type: $($Collection.CollectionType)"
                                                }
                                            }
                                            Write-Host "Successfully added include rule for collection '$includeCollectionName'" -ForegroundColor Green
                                        }
                                        catch {
                                            Write-Warning "Failed to add include collection rule for '$includeCollectionName'. Error: $_"
                                        }
                                    }
                                    else {
                                        Write-Warning "Include collection with name '$includeCollectionName' not found. Skipping rule."
                                    }
                                }
                                else {
                                    Write-Warning "Could not determine name for include collection with ID $($rule.IncludeCollectionID). Skipping rule."
                                }
                            }
                            
                            # Exclude Collections
                            "SMS_CollectionRuleExcludeCollection" {
                                # Get the collection name from the source environment using the old ID
                                $excludeCollectionName = $ImportedCollections | 
                                Where-Object { $_.CollectionID -eq $rule.ExcludeCollectionID } | 
                                Select-Object -ExpandProperty Name -ErrorAction SilentlyContinue
                                
                                if ($excludeCollectionName) {
                                    Write-Host "Looking up exclude collection by name: '$excludeCollectionName'" -ForegroundColor Cyan
                                    
                                    # Look up the collection by name in the target environment
                                    $excludeCollection = Get-CMCollection -Name $excludeCollectionName -ErrorAction SilentlyContinue
                                    
                                    if ($excludeCollection) {
                                        try {
                                            switch ($Collection.CollectionType) {
                                                1 {
                                                    # User Collection
                                                    Add-CMUserCollectionExcludeMembershipRule -CollectionId $NewCollection.CollectionID -ExcludeCollectionId $excludeCollection.CollectionID -ErrorAction Stop
                                                }
                                                2 {
                                                    # Device Collection
                                                    Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $NewCollection.CollectionID -ExcludeCollectionId $excludeCollection.CollectionID -ErrorAction Stop
                                                }
                                                Default {
                                                    Write-Warning "Unknown collection type: $($Collection.CollectionType)"
                                                }
                                            }
                                            Write-Host "Successfully added exclude rule for collection '$excludeCollectionName'" -ForegroundColor Green
                                        }
                                        catch {
                                            Write-Warning "Failed to add exclude collection rule for '$excludeCollectionName'. Error: $_"
                                        }
                                    }
                                    else {
                                        Write-Warning "Exclude collection with name '$excludeCollectionName' not found. Skipping rule."
                                    }
                                }
                                else {
                                    Write-Warning "Could not determine name for exclude collection with ID $($rule.ExcludeCollectionID). Skipping rule."
                                }
                            }
                            
                            Default {
                                Write-Warning "Unknown rule type: $($rule.SmsProviderObjectPath). Skipping."
                            }
                        }
                    }
                }
                else {
                    Write-Host "No collection rules found for collection '$CollectionName'." -ForegroundColor Gray
                }
            }
        }
        
        Write-Host "Import completed successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "An error occurred during import: $_"
    }
}
