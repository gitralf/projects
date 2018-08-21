# Azure Inventory

PowerShell script to read the details of ARM resources and reports them in HTML format.

Usage:

```
report-htm.ps1 [-outdir <path>] [-subscription <subscriptionID>] [-resourcegroup <resourcegroup>]
```

- default path is a subdirectory in user temp
- if subscription is not provided, will present list of subscriptions to pick one
- if resourcegroup is not provided, will present a list to select at least one resourcegroup

Creates a main HTML file and one for each resourcegroup. Download the content from demo folder and open in browser to see a demo output.

Here is a list of what's currently reported, check back, the list will grow continuously (and attributes might be added to the report).

- [x] Microsoft.Compute/virtualMachines
- [x] Microsoft.Compute/virtualMachines/extensions
- [x] Microsoft.Compute/disks
- [x] Microsoft.Network/networkInterfaces
- [x] Microsoft.Network/publicIPAddresses
- [x] Microsoft.Network/networkSecurityGroups
- [x] Microsoft.Storage/storageAccounts
- [x] Microsoft\.Web/serverFarms
- [x] Microsoft\.Web/sites
- [x] Microsoft.KeyVault/vaults
- [x] Microsoft.Network/virtualNetworks

Maybe to come:

- list of AAD users and groups in extra file
- RBAC for each resourcegroup and resource
  - maybe with a switch for extended
- other resource providers

## Version history

### 2.4 (8/21/2018)

- fixed display of creation date
- added demo html files
- sort selected resources alphabetically

### 2.3 (8/17/2018)

- added extra file (resources.htm) with a list of all resources in all selected resourcegroups. Sort by Type, Groupname and Name.
- added links in resources file back to resourcegroups and resources

### 2.2 (8/15/2018)

- fixed Public Ip link bug

### 2.1 (8/14/2018)

- added Virtual Networks handler
- fixed display of tags for resourcegroup
- fixed table width for "all resources"

### 2.0 (8/13/2018)

- added NetworkSecurityGroups including SecurityRules and associations
- fixed commandline selection of subscription (was ignored)