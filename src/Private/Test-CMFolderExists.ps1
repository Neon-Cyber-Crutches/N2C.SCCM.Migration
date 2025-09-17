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