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