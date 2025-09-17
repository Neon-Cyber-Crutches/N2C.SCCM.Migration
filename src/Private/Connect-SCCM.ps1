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