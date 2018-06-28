# Azure Inventory

reads all the details of your resources and reports them in HTML format.

Here is a list of what's currently reported, check back, the list will grow continuously (and attributes might be added to the report).

- [x] Microsoft.Compute/virtualMachines
- [x] Microsoft.Compute/virtualMachines/extensions
- [x] Microsoft.Compute/disks
- [x] Microsoft.Network/networkInterfaces
- [x] Microsoft.Network/publicIPAddresses
- [ ] Microsoft.Network/networkSecurityGroups
- [ ] Microsoft.Network/virtualNetworks
- [x] Microsoft.Storage/storageAccounts
- [ ] Microsoft.ClassicStorage/storageAccounts
- [x] Microsoft\.Web/serverFarms
- [x] Microsoft\.Web/sites
- [ ] Microsoft.RecoveryServices/vaults
- [X] Microsoft.KeyVault/vaults
- [ ] Microsoft.OperationalInsights/workspaces
- [ ] Microsoft.OperationsManagement/solutions


also to come:
- list of AAD users and groups in extra file
- RBAC for each resourcegroup and resource
  - maybe with a switch for extended
- build CSV additionally