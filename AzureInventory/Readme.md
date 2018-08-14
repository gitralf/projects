# Azure Inventory

reads all the details of your resources and reports them in HTML format.

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
- [X] Microsoft.KeyVault/vaults
- [ ] Microsoft.ClassicStorage/storageAccounts
- [ ] Microsoft.Network/virtualNetworks


also to come (?):
- list of AAD users and groups in extra file
- RBAC for each resourcegroup and resource
  - maybe with a switch for extended
- build CSV overview of all resources additionally
- other resource providers

## Version history

### 2.0

- added NetworkSecurityGroups including SecurityRules and associations
- fixed commandline selection of subscription (was ignored)