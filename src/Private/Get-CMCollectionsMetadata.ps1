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