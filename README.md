# PSInternetConnectionSharing
Based on code from a [superuser.com forum post](https://superuser.com/questions/470319/how-to-enable-internet-connection-sharing-using-command-line/649183) this PowerShell Module provides simple functions to control Windows Internet Connection Sharing (ICS) from command line.

The module includes three cmdlet functions:
* Set-ICS
* Get-ICS
* Disable-ICS 

It requires to run with administrative rights.

The module has been tested on Windows 10.

## Installation

Download the PS Module file (psm1) and copy it to your PSModulePath, usually `C:\<User>\Documents\WindowsPowerShell\Modules\<Module Folder>\<Module Files>` if installing for a specific user or, if installing for all users, `C:\Program Files\WindowsPowerShell\Modules\<Module Folder>\<Module Files>`. Name the `<Module Folder>` exactly as the psm1 file, in this case "PSInternetConnectionSharing". PowerShell will now automatically find the module and its cmdlets.
  
## Functions (cmdlets)
 
### Set-ICS
#### Description
Set-ICS lets you share the internet connection of a network connection (called the public connection) with another
 network connection (called the private connection). The specified network connections must exist beforehand.
 In order to be able to set ICS, the function will first disable ICS for any existing network connections.
 It will also check for if ICS is already enabled for the specified network connection pair.
#### Parameters
##### PublicConnectionName
The name of the network connection that internet connection will be shared from.
##### PrivateConnectionName
The name of the network connection that internet connection will be shared with.
#### Usage examples
1. `Set-ICS -PublicConnectionName Ethernet -PrivateConnectionName 'VM Host-Only Network'`
2. `Set-ICS Ethernet 'VM Host-Only Network'`
