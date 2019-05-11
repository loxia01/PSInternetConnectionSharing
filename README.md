# PSInternetConnectionSharing
Based on code from a [superuser.com forum post](https://superuser.com/questions/470319/how-to-enable-internet-connection-sharing-using-command-line/649183) this PowerShell Module provides simple functions to control Windows Internet Connection Sharing (ICS) from command line.

The module includes three cmdlet functions:
* Set-ICS
* Get-ICS
* Disable-ICS 

It requires to run with administrative rights.

The module has been tested on Windows 10.

## Installation

Download the PS Module file (psm1) and copy it to your PSModulePath, usually `C:\<User>\Documents\WindowsPowerShell\Modules\<Module Folder>\<Module Files>` if installing for a specific user or, if installing for all users, C:\Program Files\WindowsPowerShell\Modules\<Module Folder>\<Module Files>. Name the <Module Folder> exactly as the psm1 file, in this case "PSInternetConnectionSharing". PowerShell will now automatically find the module and its cmdlets.
  
  
