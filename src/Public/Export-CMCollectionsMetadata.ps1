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