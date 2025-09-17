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