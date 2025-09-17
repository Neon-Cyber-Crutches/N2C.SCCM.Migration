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