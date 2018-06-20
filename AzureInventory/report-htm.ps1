<#
 .SYNOPSIS
    create a HTML report of Azure resources

 .DESCRIPTION
    build inventory data of Azure resources based on subscription and/or resourcegroup scope. 
    select your subscription and the resource groups to be listed and there you go.

 .PARAMETER outdir
    Directory where all the output will go to (will be created if not found). 
    Different HTML pages will link to each other relatively, all in that directory.
    If left out, a directory in user TEMP will be created (with "report" and timestamp)

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
        $sub = Get-AzureRmSubscription -subscriptionName $_.Name
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

        #define a standard table row
        $intlink = "<a href='#{0}'>{1}</a>"
        $extlink = "<a href='{0}.htm'>{1}</a>"

        $table = "
        <table id=inventory width='{0}%'>
        "

        $rowhead="
        <tr>
            <th>
                {0}
            </th>
            <th>
                {1}
            </th>
        </tr>
        "

        $rowdetailhead="
        <tr>
            <th>
                {0}
            </th>
            <th colspan=2>
                {1}
            </th>
        </tr>
        "

        $row="
        <tr>
            <td>
                {0}
            </td>
            <td>
                {1}
            </td>
        </tr>
        "

        $rowdetail="
        <tr>
            <td>
                {0}
            </td>
            <td colspan=2>
                {1}
            </td>
        </tr>
        "

        $row3detail="
        <tr>
            <td>
                {0}
            </td>
            <td>
                {1}
            </td>
            <td>
                {2}
            </td>
        </tr>
        "

         
###### start with the overview page
        $outputmain = $outputhead 
        $outputmain += "<body>"
        $outputmain += "<h1 id='top'>Azure Inventory</h1>"
        $outputmain += $table -f "50"
        $outputmain += $row -f "created",$now
        $outputmain += $row -f "Tenant-ID",$sub.TenantId
        $outputmain += $row -f "Subscription-ID",$sub.SubscriptionId
        $outputmain += $row -f "Subscription name",$sub.Name
        $outputmain += $row -f "Resourcegroups",$nrRG
        $outputmain += $row -f "Total Resources",$nrResources
        $outputmain += "</table>"
        
###### we build a list of resourcegroups and - for each of them - a separate html file with the resources
        $outputmain += "<h1>Resourcegroups</h1><p>Total: {0}, Selected {1}</p>" -f $nrRG,$nrRGSelected
        $outputmain += $table -f "50"
        $outputmain += $rowhead -f "Resourcegroupname", "Location"

        $rgnumber=0;
        $rnumber=0

###### here we start walking through all RGs and 
###### add a line to the main html page ($outputmain) and 
###### new html content with the resources of that RG (in $outputRG)

        foreach ($RG in $RGSelected.Keys) {
            $thisRG=Get-AzureRmResourceGroup -Name $RG
            $link = $extlink -f $RG,$RG
            $outputmain += $row -f $link,$thisRG.Location

            $rgnumber++
            $rnumber=0


            $outputRG = $outputhead +"<h1>{0}. Resourcegroup {1}</h1>" -f $rgnumber,$RG

            ### do we have a RG with Tags?
            if ($thisRG.tags.keys.length -gt 0){
                $outputRG += $table -f "50"
                $outputRG += $rowhead -f "Tag","Value"
                        
                foreach ($key in $thisRG.tags.keys){
                    $outputRG += $row -f $key, $resourcegroup.tags[$key]
                }

                $outputRG += "</table>"
            } #tags

            $outputRG += "<h2>All resources in resourcegroup {0}</h2>" -f $RG
            $outputRG += $table -f "75"
            $outputRG += $rowhead -f "Resourcename","Type"

###### we create a table with the resource details (in $resourcetable)
            $resourcetable = ""
            $detailtable = ""

###### and we walk through all Resources and filter out those in the current RG
###### this seems to be faster than Find-AzureRmResource -ResourceGroupNameContains

            $Resources | Where-Object {$_.resourcegroupname -eq $RG}| ForEach-Object {
                $rnumber++
                $thisresource=$_
                $link = $intlink -f $thisresource.Name,$thisresource.Name
                $resourcetable += $row -f $link,$thisresource.resourcetype

                ### for the sake of readbility we replace the resourcetype with a more friendly name (see top)
                if ($displayname.ContainsKey($thisresource.resourcetype)){
                    $display=$displayname.($thisresource.resourcetype)
                } else {
                    $display="Resource"
                }

                $detailtable += "<h3><a name='{0}'>{1}.{2} {3} '{4}' in resourcegroup {5}</a></h3>" -f $thisresource.name,$rgnumber,$rnumber,$display,$thisresource.Name,$RG
                $detailtable += $table -f "50"
                $detailtable += $rowdetailhead -f "Attribute","Value"
                $detailtable += $rowdetail -f "ResourceType",$thisresource.resourcetype

                foreach ($key in $thisresource.tags.keys){
                    $detailtable += $rowdetail -f $key, $thisresource.tags[$key]
                }

###### this is where the magic happens
###### now we go for the real details. place a handler for each resourcetype here


                switch ($thisresource.resourcetype) {

######## Microsoft.Compute/virtualMachines
                    "Microsoft.Compute/virtualMachines" {
                        $vm=get-azurermvm -Name $thisresource.Resourcename -ResourceGroupName $RG -WarningAction "SilentlyContinue"
                        $detailtable += $rowdetail -f "VM Size",$vm.HardwareProfile.vmSize
                
                        if ($vm.StorageProfile.ImageReference){
                            $detailtable += $rowdetail -f "Image offer",$vm.storageProfile.Imagereference.offer
                            $detailtable += $rowdetail -f "ImageSKU", $vm.storageProfile.ImageReference.Sku
                            $detailtable += $rowdetail -f "Image publisher",$vm.storageProfile.ImageReference.publisher
                        }

                        $temp=Get-AzureRmVM -ResourceGroupName $RG -Name $thisresource.Resourcename -status -InformationAction "SilentlyContinue" -WarningAction "SilentlyContinue"
                        ForEach ($VMStatus in $temp.Statuses){
                            if ($VMStatus.Code -like "PowerState/*"){
                                $status=$VMStatus.Code.split("/")[1]
                                $detailtable += $rowdetail -f "PowerState",$status
                            }
                        }#status
                    }#vm handler

######## Microsoft.Storage/storageAccounts
                    "Microsoft.Storage/storageAccounts" {
                        $detailtable += $rowdetail -f "SKU name",$thisresource.sku.name
                        $detailtable += $rowdetail -f "SKU tier",$thisresource.sku.tier
                    }#storageAccount handler


######## Microsoft.Web/sites
                    "Microsoft.Web/sites" {
                        $website=Get-AzureRmWebApp -ResourceGroupName $RG -Name $thisresource.name
                        $detailtable += $rowdetail -f "State",$website.state

                        foreach ($hostname in $website.hostNames){
                            $detailtable += $rowdetail -f "hostname",$hostname
                        }
                    }#website handler


######## Microsoft.Sql/servers
                    "Microsoft.Sql/servers" {
                        $detailtable += $rowdetail -f "Kind",$thisresource.Kind
                    }#sql server handler


######## Microsoft.Sql/servers/databases
                    "Microsoft.Sql/servers/databases" {
                        $detailtable += $rowdetail -f "Kind",$thisresource.Kind
                    }#database handler


######## Microsoft.Network/networkInterfaces
                    "Microsoft.Network/networkInterfaces" {
                        $nic = Get-AzureRmNetworkInterface -Name $thisresource.Name -ResourceGroupName $RG 
                        $linkedVMId = $nic.VirtualMachine.Id
                        if ($linkedVMId){
                            $linkedVM = Get-AzureRmResource -ResourceId $linkedVMId
                            $link = $intlink -f $linkedVM.Name,$linkedVM.Name
                            $link2 = $extlink -f $linkedVM.Resourcegroupname,$linkedVM.Resourcegroupname
                            $attachedText = "{0} (in {1})" -f $link,$link2
                            $detailtable += $rowdetail -f "attachedTo",$attachedText
                        } else {
                            $detailtable += $rowdetail -f "nothing"
                        }

                        foreach ($ipconfig in $nic.IpConfigurations){
                            $detailtable += $row3detail -f "IPconfig","Name",$ipconfig.name
                            $detailtable += $row3detail -f "&nbsp;","PrivateIP",$ipconfig.PrivateIpAddress
                            $detailtable += $row3detail -f "&nbsp;","AllocationMethodPrivateIP",$ipconfig.PrivateIPAllocationMethod
                            
                            $subnetID=$ipconfig.Subnet.Id
                            $subnetparts=$subnetid.split("/")
                        }
                    }#nic handler

                    
######## Microsoft.Network/publicIPAddresses
                    "Microsoft.Network/publicIPAddresses" {
                        $pip = Get-AzureRmPublicIpAddress -Name $thisresource.Name -ResourceGroupName $RG 
                        $attached=$pip.IpConfiguration
                        if ($attached.length -gt 0){
                            $temp=$attached.id.split("/")
                            $link = $linkint -f $temp[8],$temp[8]
                            $link2 = $linkext -f $temp[4],$temp[4]
                            $attachedText = "{0} on NIC {1} (in {2}" -f $temp[10],$link,$link2
                        } else {
                            $attachedText="nothing"
                        }
                        $detailtable += $rowdetail -f "IPAddress",$pip.IpAddress
                        $detailtable += $rowdetail -f "AllocationMethod",$pip.PublicIpAllocationMethod
                        $detailtable += $rowdetail -f "SKU",$pip.Sku.Name
                        $detailtable += $rowdetail -f "DNS FQDN",$pip.DnsSettings.Fqdn
                        $detailtable += $rowdetail -f "attachedTo",$attachedText
                    }#publicIP handler





                    Default {
                        $detailtable += $rowdetail -f "no handler found","&nbsp;"
                    }
                }

                $detailtable += "</table>"
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
