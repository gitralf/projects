<#
 .SYNOPSIS
    create a HTML report of Azure resources

 .DESCRIPTION
    build inventory data of Azure resources based on subscription and/or resourcegroup scope. 
    select your subscription and the resource groups to be listed and there you go.
    This is version 2.1

 .PARAMETER outdir
    Directory where all the output will go to (will be created if not found). 
    Different HTML pages will link to each other relatively, all in that directory.
    If left out, a directory in user TEMP will be created (with "report" and timestamp)

 .PARAMETER subscription
    Define a subscription. If not defined, try all subscriptions

 .PARAMETER resourcegroup
    Define a Resourcegroup. If not defined, try all resourcegroups
#>

Param(
    [parameter(mandatory=$false)][string]$outdir,
    [parameter(mandatory=$false)][string]$subscription,
    [parameter(mandatory=$false)][string]$resourcegroup
)


function add-line {
    Param (
        [parameter(mandatory="true")][string]$table,
        [parameter(mandatory="true")][string]$left,
        [parameter(mandatory="true")][string]$right,
        [parameter(mandatory="false")][boolean]$isheader
    )
}

$now=(get-date -UFormat "%Y%m%d%H%M%S").ToString()

if ($outdir.Length -eq 0){
    # no outdir defined. use temp folder
    $outdir=$env:TEMP + "\report"+$now
}

if(!(Test-Path -Path $outdir )){
    $dummy=New-Item -ItemType directory -Path $outdir
}

#define the start of inventory
$outmain="main"
$outputmainfile="{0}\{1}.htm" -f $outdir,$outmain

#tell it to the user
write-host "writing inventory to $outputmainfile" -ForegroundColor "yellow"

#define nicer names for the resource providers
$displayname=@{}
$displayname.Add('Microsoft.Compute/virtualMachines','VM')
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
$WarningActionPreference = "SilentlyContinue"

# save the resourcegroups to be inspected in a hashtable.
$RGSelected = @{}

# we should stop on this:
$FatalError= 0

#define some HTML stuff here...

#define global HEAD
$outputhead = "
<html>
    <head>
        <style>
        #inventory, #inventory2 {
            font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;
            border-collapse: collapse;
        }
        
        body { 
            font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;
        }

        #inventory td, #inventory th, #inventory2 td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        
        #inventory tr:nth-child(even), #inventory2 tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        
        #inventory tr:hover, #inventory2 tr:hover {
            background-color: #ddd;
        }
        
        #inventory th {
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: #4CAF50;
            color: white;
        }

        #inventory2 th {
            padding-top: 12px;
            padding-bottom: 12px;
            text-align: left;
            background-color: #4c004c;
            color: white;
        }
        </style>
        <title>
                Azure Inventory
        </title>
    </head>
"

#define link format
$linkext = "<a href='{0}.htm'>{1}</a>"
$linkint = "<a href='#{0}'>{1}</a>"
$linkabs = "<a href='{0}.htm#{1}'>{2}</a>"

$table = "
        <table id=inventory width='{0}%'>
"
$table2= "
        <table id=inventory2 width='{0}%'>
"

$rowhead="
        <tr>
            <th width='33%'>
                {0}
            </th>
            <th width='67%'>
                {1}
            </th>
        </tr>
"

$rowdetailhead="
        <tr>
            <th width='33%'>
                {0}
            </th>
            <th colspan=2 width='67%'>
                {1}
            </th>
        </tr>
"

#normal table with 2 rows, 1:2
$row="
        <tr>
            <td width='33%'>
                {0}
            </td>
            <td width='67%'>
                {1}
            </td>
        </tr>
"

#table with 3 rows, 2+3 spanning, 1:2
$rowdetail="
        <tr>
            <td width='33%'>
                {0}
            </td>
            <td colspan=2 width='67%'>
                {1}
            </td>
        </tr>
"

#table with 3 rows, all equal
$row3detail="
        <tr>
            <td width='33%'>
                {0}
            </td>
            <td width='33%'>
                {1}
            </td>
            <td width='33%'>
                {2}
            </td>
        </tr>
        "


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
    #pick the subscription first. if there is only one, take that. 
    #If a subscriptionID was given as parameter, take this without checking
    if ($subscription){
        "Subscription defined on commandline, skipping selection..."
        $sub = Get-AzureRmSubscription -SubscriptionId $subscription
    } else {
        if ((Get-AzureRmSubscription).count -eq 1){
            "only one subscription found, skipping selection..." 
            $sub= Get-AzureRmSubscription
        } else {
            Get-AzureRmSubscription | select-object Name,ID,state |Out-GridView -Title "Select one subscription" -OutputMode Single | ForEach-Object {
                $sub = Get-AzureRmSubscription -subscriptionName $_.Name
            }
        }
    }

    $Resourcegroups = Get-AzureRmResourceGroup
    if ($resourcegroup){
        "Resourcegroup defined on commandline, skipping selection..."
        $temp = Get-AzureRmResourceGroup -name $resourcegroup
        $RGSelected.Add($temp.Resourcegroupname, $temp.ResourceID)
    } else {
        "Please select at least one resourcegroup in subscription '{0}'" -f $sub.Name
        #pick the resourcegroups to examine

        $Resourcegroups | Select-Object ResourceGroupName,ResourceId | Out-GridView -Title "Select Resourcegroups (use Ctrl)" -PassThru | ForEach-Object {
            $RGSelected.Add($_.ResourceGroupName, $_.ResourceID)
        }
    }
    $nrRGSelected = $RGSelected.Count
    $nrRG = $resourcegroups.Count

    # at least one should be selected...
    if ($nrRGSelected -gt 0){
        
        # we collect all resources at once instead of iterating through each
        $Resources=Get-AzureRmResource 
        $nrResources=$Resources.count


         
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
            $link = $linkext -f $RG,$RG
            $outputmain += $row -f $link,$thisRG.Location

            $rgnumber++
            $rnumber=0


            $outputRG = $outputhead +"<h1><a name='top'>{0}</a>. Resourcegroup {1}</h1>" -f $rgnumber,$RG

            ### do we have a RG with Tags?
            if ($thisRG.tags.keys.length -gt 0){
                $outputRG += $table -f "50"
                $outputRG += $rowhead -f "Tag","Value"
                        
                foreach ($key in $thisRG.tags.keys){
                    $outputRG += $row -f $key, $thisRG.tags[$key]
                }

                $outputRG += "</table>"
            } #tags

            $outputRG += "<h2>All resources in resourcegroup {0}</h2>" -f $RG
            $outputRG += $table -f "50"
            $outputRG += $rowhead -f "Resourcename","Type"

###### we create a table with the resource details (in $resourcetable)
            $resourcetable = ""
            $detailtable = ""

###### and we walk through all Resources and filter out those in the current RG
###### this seems to be faster than Find-AzureRmResource -ResourceGroupNameContains

            $Resources | Where-Object {$_.resourcegroupname -eq $RG}| ForEach-Object {
                # first build a list of all resources in this RG on top of the page 
                # do this for each resource as it is inspected
                $rnumber++
                $thisresource=$_
                $link = $linkint -f $thisresource.Name,$thisresource.Name
                $resourcetable += $row -f $link,$thisresource.resourcetype

                ### for better readability we replace the resourcetype with a more friendly name (see top)
                if ($displayname.ContainsKey($thisresource.resourcetype)){
                    $display=$displayname.($thisresource.resourcetype)
                } else {
                    $display="Resource"
                }

                ### create a detailtable for each resource. header is equal for all resource types, 
                ### the different atrributes follow after the switch further down
                $detailtable += "<h3><a name='{0}'>{1}.{2} {3} '{4}' in resourcegroup {5}</a></h3>" -f $thisresource.name,$rgnumber,$rnumber,$display,$thisresource.Name,$RG
                $detailtable += $table -f "50"
                $detailtable += $rowdetailhead -f "Attribute","Value"
                $detailtable += $rowdetail -f "ResourceType",$thisresource.resourcetype

                ###do we have tags on the resource? if so, put them at the beginning

                foreach ($key in $thisresource.tags.keys){
                    $thistag="Tag: {0}" -f $key
                    $detailtable += $rowdetail -f $thistag, $thisresource.tags[$key]
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
                        $sql = Get-AzureSqlDatabaseServer -ServerName $thisresource.name 
                        $detailtable += $rowdetail -f "Kind",$thisresource.Kind
                        $detailtable += $rowdetail -f "Location",$sql.Location
                        $detailtable += $rowdetail -f "Version",$sql.Version
                        $detailtable += $rowdetail -f "State",$sql.State
                        $detailtable += $rowdetail -f "AdminLogin",$sql.AdministratorLogin
                        $sqlfw=Get-AzureSqlDatabaseServerFirewallRule -ServerName $thisresource.Name
                        if ($sqlfw.length -gt 0){
                            $detailtable += "</table>"
                            $detailtable += $table2 -f "50"
                            $detailtable += $rowdetailhead -f "firewallRules","&nbsp;"
                            $detailtable += $row3detail -f "Name","StartIpAddress","EndIpAddress"
                            $sqlfw | foreach-object {
                                $detailtable += $row3detail -f $_.Rulename, $_.StartIpAddress, $_.EndIpAddress
                            }
                        }
                        $sqldb = Get-AzureSqlDatabase -ServerName $thisresource.Name 
                        if ($sqldb.length -gt 0){
                            $detailtable += "</table>"
                            $detailtable += $table2 -f "50"
                            $detailtable += $rowdetailhead -f "Databases","&nbsp;"
                            $detailtable += $row3detail -f "Name","Edition","Collation"
                            $sqldb | foreach-object {
                                $edition = "{0}, {1}" -f $_.ServiceObjectiveName,$_.$_.Edition
                                $detailtable += $row3detail -f $_.Name,$edition,$_.CollationName
                            }
                        }
                    }#sql server handler


######## Microsoft.Sql/servers/databases
                    "Microsoft.Sql/servers/databases" {
                        $names = $thisresource.name.split("/")
                        $sql = Get-AzureSqlDatabase -ServerName $names[0] -DatabaseName $names[1]
                        $detailtable += $rowdetail -f "Kind",$thisresource.Kind
                        $detailtable += $rowdetail -f "DBName",$sql.name
                        $detailtable += $rowdetail -f "Edition",$sql.edition
                        $detailtable += $rowdetail -f "ServiceObjectiveName",$sql.ServiceObjectiveName
                        $detailtable += $rowdetail -f "Collation",$sql.CollationName
                    }#database handler


######## Microsoft.Network/networkInterfaces
                    "Microsoft.Network/networkInterfaces" {
                        $nic = Get-AzureRmNetworkInterface -Name $thisresource.Name -ResourceGroupName $RG 
                        $linkedVMId = $nic.VirtualMachine.Id
                        if ($linkedVMId){
                            $linkedVM = Get-AzureRmResource -ResourceId $linkedVMId
                            $link = $linkint -f $linkedVM.Name,$linkedVM.Name
                            $link2 = $linkext -f $linkedVM.Resourcegroupname,$linkedVM.Resourcegroupname
                            $attachedText = "{0} (in {1})" -f $link,$link2
                            $detailtable += $rowdetail -f "attachedTo",$attachedText
                        } else {
                            $detailtable += $rowdetail -f "nothing"
                        }

                        if (($nic.IpConfigurations).length -gt 0){
                            foreach ($ipconfig in $nic.IpConfigurations){
                                $detailtable += "</table>"
                                $detailtable += $table2 -f "50"
                                $detailtable += $rowdetailhead -f "IPConfiguration",$ipconfig.name
                                $detailtable += $rowdetail -f "PrivateIP",$ipconfig.PrivateIpAddress
                                $detailtable += $rowdetail -f "AllocationMethodPrivateIP",$ipconfig.PrivateIPAllocationMethod
                                
                                $subnetID=$ipconfig.Subnet.Id
                                $subnetparts=$subnetid.split("/")
                            }
                        }
                    }#nic handler

                    
######## Microsoft.Network/publicIPAddresses
                    "Microsoft.Network/publicIPAddresses" {
                        $pip = Get-AzureRmPublicIpAddress -Name $thisresource.Name -ResourceGroupName $RG 
                        $attached=$pip.IpConfiguration
                        if ($attached.length -gt 0){
                            $temp=$attached.id.split("/")
                            $link = $linkext -f $temp[8],$temp[8]
                            $link2 = $linkext -f $temp[4],$temp[4]
                            $attachedText = "{0} on NIC {1} (in {2})" -f $temp[10],$link,$link2
                        } else {
                            $attachedText="nothing"
                        }
                        $detailtable += $rowdetail -f "IPAddress",$pip.IpAddress
                        $detailtable += $rowdetail -f "AllocationMethod",$pip.PublicIpAllocationMethod
                        $detailtable += $rowdetail -f "SKU",$pip.Sku.Name
                        $detailtable += $rowdetail -f "DNS FQDN",$pip.DnsSettings.Fqdn
                        $detailtable += $rowdetail -f "attachedTo",$attachedText
                    }#publicIP handler

###### Microsoft.Compute/virtualMachines/extensions
                    "Microsoft.Compute/virtualMachines/extensions"{
                        $temp=$thisresource.name.split("/")
                        $vmext=Get-AzureRmVMExtension -ResourceGroupName $rg -VMName $temp[0] -Name $temp[1]
                        $link = $linkext -f $RG,$vmext.VMName,$vmext.VMName
                        $detailtable += $rowdetail -f "attached to",$link
                        $detailtable += $rowdetail -f "ExtensionType",$vmext.ExtensionType
                        $detailtable += $rowdetail -f "TypeHandlerVersion",$vmext.TypeHandlerVersion
                    }#vm extension

###### Microsoft.KeyVault/vaults
                    "Microsoft.KeyVault/vaults"{
                        $vault=Get-AzureRmKeyVault -VaultName $thisresource.Name
                        $link= $linkext -f $vault.VaultUri,$vault.VaultUri
                        $detailtable += $rowdetail -f "VaultURI",$link
                        $detailtable += $rowdetail -f "SKU",$vault.Sku
                        $detailtable += $rowdetail -f "EnabledForDeployment",$vault.EnabledForDeployment
                        $detailtable += $rowdetail -f "EnabledForTemplateDeployment",$vault.EnabledForTemplateDeployment
                        $detailtable += $rowdetail -f "EnabledForDiskEncryption",$vault.EnabledForDiskEncryption
                    }#AKV

###### Microsoft.Compute/disks
                    "Microsoft.Compute/disks" {
                        $disk = Get-AzureRmDisk -ResourceGroupName $RG -DiskName $thisresource.Name -WarningAction "SilentlyContinue"
                        $temp = $disk.ManagedBy.split("/")
                        $linkr = $linkabs -f $temp[4],$temp[8],$temp[8]
                        $linkrg  = $linkabs -f $temp[4],"top",$temp[4]
                        $attachedText = "{0} (in {1})" -f $linkr,$linkrg
                        $detailtable += $rowdetail -f "managedBy",$attachedText 
                        $detailtable += $rowdetail -f "OSType",$disk.OsType
                        $detailtable += $rowdetail -f "Disksize [GB]",$disk.DiskSizeGB
                        $detailtable += $rowdetail -f "EncryptionSettings",$disk.EncryptionSettings
                        # sku to be added later. output will change...
                        #$detailtable += $rowdetail -f "sku",$disk.Sku
                    }                    

###### Microsoft.Network/networkSecurityGroups
                    "Microsoft.Network/networkSecurityGroups" {
                        $nsg = Get-AzureRmNetworkSecurityGroup -ResourceGroupName $RG -Name $thisresource.Name
                        if ($nsg.Subnets.Count+$nsg.NetworkInterfaces.count -lt 1){
                            $detailtable += $rowdetail -f "associated with ","nothing"
                        } else {
                            if ($nsg.Subnets.count -gt 0){
                                foreach ($subnet in $nsg.Subnets) {
                                    $temp=$subnet.id.split("/")
                                    $link= $linkabs -f $temp[4],$temp[8],$temp[8]
                                    $attachedText="subnet {0} (in vnet {1})" -f $temp[10],$link 
                                    $detailtable += $rowdetail -f "associated with",$attachedText
                                }
                            }
                            if ($nsg.NetworkInterfaces.Count -gt 0){
                                foreach ($nic in $nsg.NetworkInterfaces) {
                                    $temp=$nic.id.split("/")
                                    $link=$linkabs -f $temp[4],$temp[8],$temp[8]
                                    $attachedText="NIC {0}" -f $link
                                    $detailtable += $rowdetail -f "associated with",$attachedText
                                }
                            }
                        }
                        $detailtable += "</table>"
                        $detailtable += $table2 -f "50"
                        $detailtable += $rowdetailhead -f "SecurityRules","&nbsp;"
                        $rules=Get-AzureRmNetworkSecurityRuleConfig -NetworkSecurityGroup $nsg|sort-object -Property Direction,Priority
                        foreach ($rule in $rules){
                            $detailtable += $row3detail -f $rule.Name,"Direction",$rule.Direction
                            $detailtable += $row3detail -f "&nbsp;","Priority",$rule.Priority
                            $detailtable += $row3detail -f "&nbsp;","Access",$rule.Access
                            $detailtable += $row3detail -f "&nbsp;","Protocol",$rule.Protocol
                            $temp=$rule.SourceAddressPrefix -join ","
                            $detailtable += $row3detail -f "&nbsp;","SourceAddress",$temp
                            $temp=$rule.SourcePortRange -join ","
                            $detailtable += $row3detail -f "&nbsp;","SourcePort",$temp
                            $temp=$rule.DestinationAddressPrefix -join ","
                            $detailtable += $row3detail -f "&nbsp;","DestAddress",$temp
                            $temp=$rule.DestinationPortRange -join ","
                            $detailtable += $row3detail -f "&nbsp;","DestPort",$temp
                        }
                    }
###### Microsoft.Network/virtualNetworks
                    "Microsoft.Network/virtualNetworks" {
                        $vnet = Get-AzureRmVirtualNetwork -ResourceGroupName $RG -name $thisresource.Name
                        $detailtable += $rowdetail -f "EnableDDoSProtection",$vnet.EnableDDoSProtection
                        $detailtable += $rowdetail -f "EnableVmProtection",$vnet.EnableVmProtection
                        $temp=$vnet.AddressSpace.AddressPrefixes -join ","
                        $detailtable += $rowdetail -f "AddressPrefixes",$temp
                        $detailtable += "</table>"
                        $detailtable += $table2 -f "50"
                        $detailtable += $rowdetailhead -f "Subnets","&nbsp;"
                        
                        foreach ($subnet in $vnet.Subnets){
                            $detailtable += $row3detail -f $subnet.Name,"AddressPrefix",$subnet.AddressPrefix
                            $temp=$subnet.networkSecurityGroup.id.split("/")
                            $detailtable += $row3detail -f "&nbsp;","NSG",$temp[8]
                            foreach ($ipconfig in $subnet.IpConfigurations){
                                $temp=$subnet.IpConfigurations.id.split("/")
                                $link=$linkabs -f $temp[4],$temp[8],$temp[8]
                                $attachedText="{0} from NIC {1}" -f $temp[10],$link
                                $detailtable += $row3detail -f "&nbsp;","associated with",$attachedText
                            }
                        }
                    }


###### Default
                    Default {
                        $detailtable += $rowdetail -f "no handler found","&nbsp;"
                    }
                }

                $detailtable += "</table>"
                $detailtable += $linkint -f "top","[top]"
            }
            $link=$linkext -f $outmain,"[main]"
            $outputRG += "
            {0}
            </table>
            <p>{1}</p>
            {2}
    </body>
</html>
            " -f $resourcetable,$link,$detailtable

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
