<#
 .SYNOPSIS
    create a HTML report of Azure resources

 .DESCRIPTION
    build inventory data of Azure resources based on subscription and/or resourcegroup scope

 .PARAMETER outdir
    Directory where all the output will land (will be created if not found). 
    Different HTML pages will link to each other relatively, all in that directory.
    If left out, user TEMP directory with "report" and timestamp will be created

#>

Param(
    [parameter(mandatory=$false)][string]$outdir
)

$displayname=@{}
$displayname.Add('Microsoft.Compute/virtualMachines ','VM')
$displayname.add('Microsoft.Compute/disks','Disk')
$displayname.add('Microsoft.Compute/virtualMachines/extensions','VM extension')
$displayname.add('Microsoft.Network/networkInterfaces','NIC')
$displayname.add('Microsoft.Network/networkSecurityGroups','NSG')
$displayname.add('Microsoft.Network/publicIPAddresses','Public IP')
$displayname.add('Microsoft.Network/virtualNetworks','VNet')
$displayname.add('Microsoft.Storage/storageAccounts','Storage account')
$displayname.add('Microsoft.RecoveryServices/vaults','ARS vault')
$displayname.add('Microsoft.OperationalInsights/workspaces','OMS Workspace')
$displayname.add('Microsoft.OperationsManagement/solutions','OMS solution')
$displayname.add('Microsoft.KeyVault/vaults','KeyVault')
$displayname.add('Microsoft.ClassicStorage/storageAccounts','Classic storage account')
$displayname.add('Microsoft.Web/serverFarms','Websever farm')
$displayname.add('Microsoft.Web/sites','Website')
# $displayname.add('','')
# $displayname.add('','')


$ErrorActionPreference = "SilentlyContinue"

# save the resourcegroups to be inspected in a hashtable.
$RGSelected = @{}

# we should stop on this:
$FatalError= 0


# already logged in?
try
{
    Get-AzureRMContext | out-null 
}
catch
{
    $ErrorMessage = $_.Exception.Message
#    $FailedItem = $_.Exception.ItemName
}

# Houston, do we have a problem?

if ($ErrorMessage -like '*login*'){
    # PS module installed, but not login.
    Login-AzureRmAccount
} elseif ($ErrorMessage -like '*credentials*'){
    Login-AzureRmAccount
} elseif ($ErrorMessage -like '*not recognized as the name of a cmdlet*'){
    # PS module not installed
    write-host "Looks like Azure PS module not installed. Please visit: https://azure.microsoft.com/en-us/documentation/articles/powershell-install-configure/"
    $FatalError = 1
} elseif ($ErrorMessage -like '*Object reference not set to an instance of an object*'){
    # no internet?
    write-host "Problem with your internet connection?"
    $FatalError = 1
} elseif ($ErrorMessage){
    #anything else?
    write-host "uups. something went wrong. please check:"
    write-host $ErrorMessage -ForegroundColor "Red"
    $FatalError = 1
}





# If nothing fatal occurred we will go forward.
if ($FatalError -eq 0){
    #pick the subscription first
    Get-AzureRmSubscription | select-object Name,ID,state |Out-GridView -Title "Select subscription" -OutputMode Single | ForEach-Object {
        $sub = Get-AzureRmSubscription -Name $_.Name
    }

    "working on {0}" -f $sub.Name
    #pick the resourcegroups to examine
    $Resourcegroups = Get-AzureRmResourceGroup
    $Resourcegroups | Select-Object ResourceGroupName,ResourceId | Out-GridView -Title "Select Resourcegroups (use Ctrl)" -PassThru | ForEach-Object {
       $RGSelected.Add($_.ResourceGroupName, $_.ResourceID)
    }
    $nrRGSelected = $RGSelected.Count
    $nrRG = $resourcegroups.Count

    # at least one should be selected...
    if ($nrRGSelected -gt 0){
        
        #so finally here we go
        $now=(get-date -UFormat "%Y%m%d%H%M%S").ToString()

        if ($outdir.Length -eq 0){
            $outdir=$env:TEMP + "\report"+$now
        }
        
        "start inventory at {0}\main.htm" -f $outdir

        if(!(Test-Path -Path $outdir )){
            $dummy=New-Item -ItemType directory -Path $outdir
        }

        $outputmainfile="{0}\main.htm" -f $outdir


        $Resources=Get-AzureRmResource 
        $nrResources=$Resources.count

        #define global HEAD
        $outputhead = "
        <html>
        <head>
        <style>
        #inventory {
            font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;
            border-collapse: collapse;
        }
        
        body { 
            font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;
        }

        #inventory td, #inventory th {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        
        #inventory tr:nth-child(even){background-color: #f2f2f2;}
        
        #inventory tr:hover {background-color: #ddd;}
        
        #inventory th {
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: #4CAF50;
            color: white;
        }
        </style>
        <title>
                Azure Inventory
            </title>
        </head>
        "

        #start with the overview page
        $outputmain = $outputhead + "
    <body>
        <h1 id='top'>Azure Inventory</h1>
        <table id=inventory width='50%'>
            <tr>
                <td>
                    created:
                </td>
                <td>
                    {0}
                </td>
            </tr>
            <tr>
                <td>
                    Tenant-ID:
                </td>
                <td>
                    {1}
                </td>
            </tr>
            <tr>
                <td>
                    Subscription-ID:
                </td>
                <td>
                    {2}
                </td>
            </tr>
            <tr>
                <td>
                    Subscription name:
                </td>
                <td>
                    {3}
                </td>
            </tr>
            <tr>
                <td>
                    Resourcegroups:
                </td>
                <td>
                    {4}
                </td>
            </tr>
            <tr>
                <td>
                    Total Resources:
                </td>
                <td>
                    {5}
                </td>
            </tr>
        </table>
        " -f $now,$sub.TenantId,$sub.SubscriptionId,$sub.Name,$nrRG,$nrResources

        $outputmain += "
        <h1>Resourcegroups</h1>
        <p>Total: {0}, Selected {1}</p>

        <table id=inventory width='50%'>
            <tr>
                <th>
                    Resourcegroupname
                </th>
                <th>
                    Location
                </th>
            </tr>
        " -f $nrRG,$nrRGSelected

        $rgnumber=0;
        $rnumber=0

        foreach ($RG in $RGSelected.Keys) {
            $thisRG=Get-AzureRmResourceGroup -Name $RG
            $outputmain += "
            <tr>
                <td>
                    <a href='{0}.htm'>{0}</a>
                </td>
                <td>
                    {1}
                </td>
            </tr>
            " -f $RG,$thisRG.Location


            $rgnumber++
            $rnumber=0
            $outputRG = $outputhead +"
    <h1>
        {0}. Resourcegroup {1}
    </h1>
            " -f $rgnumber,$RG

            if ($thisRG.tags.keys.length -gt 0){
                $outputRG +"
        <table id=inventory width='50%'>
            <tr>
                <th>
                    Tag
                </th>
                <th>
                    Value
                </th>
            </tr>
            "
        
                foreach ($key in $thisRG.tags.keys){
                    $outputRG +="
            <tr>
                <td>
                    {0}
                </td>
                <td>
                    {1}
                </td>
            </tr>
                " -f $key, $resourcegroup.tags[$key]
                }

                $outputRG += "
        </table>
                "
            } #if any tags

            $outputRG += "
        <h2>
            All resources in resourcegroup {0}
        </h2>
        <table id=inventory width='75%'>
            <tr>
                <th>
                    Resourcename
                </th>
                <th>
                    Type
                </th>
            </tr>
            " -f $RG
    
            $resourcetable = ""
            $detailtable = ""
            $Resources | Where-Object {$_.resourcegroupname -eq $RG}| ForEach-Object {
                $rnumber++
                $thisresource=$_

                $resourcetable += "
            <tr>
                <td>
                    <a href='#{0}'>{1}</a>
                </td>
                <td>
                    {2}
                </td>
            </tr>
                " -f $thisresource.Name,$thisresource.Name,$thisresource.resourcetype

                if ($displayname.ContainsKey($thisresource.resourcetype)){
                    $display=$displayname.($thisresource.resourcetype)
                } else {
                    $display="Resource"
                }

                $detailtable += "
        <h3>
            <a name='{0}'>{1}.{2} {3} '{4}' in resourcegroup {5}</a>
        </h3>
        <table id=inventory width='50%'>
            <tr>
                <th>
                    Attribute
                </th>
                <th colspan=2>
                    Value
                </th>
            </tr>
            <tr>
                <td>
                    ResourceType
                </td>
                <td colspan=2>
                    {6}
                </td>
            </tr>
                " -f $thisresource.name,$rgnumber,$rnumber,$display,$thisresource.Name,$RG,$thisresource.resourcetype

                foreach ($key in $thisresource.tags.keys){
                    $detailtable +="
            <tr>
                <td>
                    {0}
                </td>
                <td colspan=2>
                    {1}
                </td>
            </tr>
                    " -f $key, $thisresource.tags[$key]
                }
    
                #now go for the real details. place a handler for each resourcetype here
                switch ($thisresource.resourcetype) {

                    "Microsoft.Compute/virtualMachines" {
                        $vm=get-azurermvm -Name $thisresource.Resourcename -ResourceGroupName $RG -WarningAction "SilentlyContinue"
                        $detailtable+="
                <tr>
                    <td>
                        VM size
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                        " -f $vm.HardwareProfile.vmSize
                
                        if ($vm.StorageProfile.ImageReference){
                            $detailtable+="
                <tr>
                    <td>
                        Image offer
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                <tr>
                    <td>
                        Image SKU
                    </td>
                    <td colspan=2>
                        {1}
                    </td>
                </tr>
                <tr>
                    <td>
                        Image publisher
                    </td>
                    <td colspan=2>
                        {2}
                    </td>
                </tr>
                            " -f $vm.storageProfile.Imagereference.offer,$vm.storageProfile.ImageReference.Sku,$vm.storageProfile.ImageReference.publisher
                        }
                        $temp=Get-AzureRmVM -ResourceGroupName $RG -Name $thisresource.Resourcename -status -InformationAction "SilentlyContinue" -WarningAction "SilentlyContinue"
                        ForEach ($VMStatus in $temp.Statuses){
                            if ($VMStatus.Code -like "PowerState/*"){
                                $status=$VMStatus.Code.split("/")[1]
                                $detailtable+="
                <tr>
                    <td>
                        PowerState
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                                " -f $status
                            }
                        }
                    }

                    "Microsoft.Storage/storageAccounts" {
                        $detailtable+="
                <tr>
                    <td>
                        SKU name
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                <tr>
                    <td>
                        SKU tier
                    </td>
                    <td colspan=2>
                        {1}
                    </td>
                </tr>
                        " -f $thisresource.sku.name,$thisresource.sku.tier
                    }

                    "Microsoft.Web/sites" {
                        $website=Get-AzureRmWebApp -ResourceGroupName $RG -Name $thisresource.name
                        $detailtable+="
                <tr>
                    <td>
                        State:
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                        " -f $website.state

                        foreach ($hostname in $website.hostNames){
                            $detailtable+="
                <tr>
                    <td>
                        hostname
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                            " -f $hostname
                        }
                    }

                    "Microsoft.Sql/servers" {
                        $detailtable+="
                <tr>
                    <td>
                        Kind
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                        " -f $thisresource.Kind
                    }

                    "Microsoft.Sql/servers/databases" {
                        $detailtable+="
                <tr>
                    <td>
                        Kind
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                        " -f $thisresource.Kind
                    }

                    "Microsoft.Network/networkInterfaces" {
                        $nic = Get-AzureRmNetworkInterface -Name $thisresource.Name -ResourceGroupName $RG 
                        $linkedVMId = $nic.VirtualMachine.Id
                        $detailtable += "
                <tr>
                    <td>
                        attachedTo
                    </td>
                        "
                        if ($linkedVMId){
                            $linkedVM = Get-AzureRmResource -ResourceId $linkedVMId
                            $detailtable+="
                    <td colspan=2>
                        <a href='#{0}'>{1} (in {2})</a>
                    </td>
                </tr>
                            " -f $linkedVM.Name,$linkedVM.Name,$linkedVM.Resourcegroupname
                        } else {
                            $detailtable+="
                    <td colspan=2>
                        nothing
                    </td>
                </tr>
                            "
                        }

                        foreach ($ipconfig in $nic.IpConfigurations){
                            $detailtable += "
                <tr>
                    <td>
                        IPconfig
                    </td>
                    <td>
                        Name
                    </td>
                    <td>
                        {0}
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        PrivateIP
                    </td>
                    <td>
                        {1}
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp;
                    </td>
                    <td>
                        AllocationMethodPrivateIP
                    </td>
                    <td>
                        {2}
                    </td>
                </tr>
                            " -f $ipconfig.name,$ipconfig.PrivateIpAddress,$ipconfig.PrivateIPAllocationMethod
                            
                            $subnetID=$ipconfig.Subnet.Id
                            $subnetparts=$subnetid.split("/")
                            
                        }
        
                    }

                    "Microsoft.Network/publicIPAddresses" {
                        $pip = Get-AzureRmPublicIpAddress -Name $thisresource.Name -ResourceGroupName $RG 
                        $attached=$pip.IpConfiguration
                        if ($attached.length -gt 0){
                            $temp=$attached.id.split("/")
                            $attachedText = "ipconfig {2} on <a href='{0}#{1}'>NIC {3}</a>" -f $temp[4],$temp[8],$temp[10],$temp[8]
                        } else {
                            $attachedText="nothing"
                        }
                        $detailtable += "
                <tr>
                    <td>
                        IP address
                    </td>
                    <td colspan=2>
                        {0}
                    </td>
                </tr>
                <tr>
                    <td>
                        Allocation method
                    </td>
                    <td colspan=2>
                        {1}
                    </td>
                </tr>
                <tr>
                    <td>
                        SKU
                    </td>
                    <td colspan=2>
                        {2}
                    </td>
                </tr>
                <tr>
                    <td>
                        DNS FQDN
                    </td>
                    <td colspan=2>
                        {3}
                    </td>
                </tr>
                <tr>
                    <td>
                        attached to
                    </td>
                    <td colspan=2>
                        {4}
                    </td>
                </tr>

                        " -f $pip.IpAddress, $pip.PublicIpAllocationMethod, $pip.Sku.Name, $pip.DnsSettings.Fqdn,$attachedText

                        
                    }





                    Default {
                        $detailtable +="
                <tr>
                    <td colspan=2>
                        no handler found
                    </td>
                </tr>
                        "
                    }
                }

                $detailtable+="</table>"
            }
            $outputRG += "
            {0}
            </table>
            {1}
        </body>
        </html>
            " -f $resourcetable,$detailtable

            $outputfile="{0}\{1}.htm" -f $outdir,$RG
            $outputRG | Out-File -filepath $outputfile
    
        }

        $outputmain +="
    </table>
</body>
</html>
        "
        $outputmain | Out-File -filepath $outputmainfile
    }
} else {
    write-host "nothing done" -ForegroundColor "Red"
}
