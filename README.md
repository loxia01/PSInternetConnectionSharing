# PSInternetConnectionSharing
This PowerShell module provides simple functions to control Windows Internet Connection Sharing (ICS) from command line.

The module includes three functions:
* `Set-Ics`
* `Get-Ics`
* `Disable-Ics` 

All functions are required to run with administrative rights and works with both Powershell Desktop (v5.1) and Core editions. PowerShell execution policy must be set to `RemoteSigned`, `Unrestricted` or `Bypass`.

The module has been tested on Windows 10.
## Installation
The module can now be downloaded from PowerShell Gallery using PowerShell command line:

`Install-Module -Name PSInternetConnectionSharing`

OR manually downloaded here on GitHub:

Download the module files (extensions `.psm1` and `.psd1`) and then create a new module folder in your `PSModulePath`. Default `PSModulePath` is:

- for a specific user: `%USERPROFILE%\Documents\WindowsPowerShell\Modules`
- for all users: `%ProgramFiles%\WindowsPowerShell\Modules`

Name the new module folder exactly as the filename without the extension, in this case `PSInternetConnectionSharing`, and then copy the downloaded module files to that folder. PowerShell will now automatically find the module and its functions.
## Function: Set-Ics
### Syntax
```
Set-Ics [-PublicConnectionName] <string> [-PrivateConnectionName] <string> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### Description
`Set-Ics` lets you share the internet connection of a network connection (called the public connection) with another network connection (called the private connection). The specified network connections must exist beforehand. In order to be able to set ICS, the function will first disable ICS for any existing network connections.
### Parameters
#### -PublicConnectionName
The name of the network connection that internet connection will be shared from.
#### -PrivateConnectionName
The name of the network connection that internet connection will be shared with.
#### -PassThru
If this parameter is specified `Set-Ics` returns an output with the set connections. Optional. By default `Set-Ics` does not generate any output.
#### -WhatIf
Shows what would happen if the function runs. The function is not run.
#### -Confirm
Prompts you for confirmation before each change the function makes.
### Usage examples
#### Example 1: Set ICS for the specified public and private connections.
`Set-Ics -PublicConnectionName Ethernet -PrivateConnectionName 'VM Host-Only Network'`

`Set-Ics Ethernet 'VM Host-Only Network'`
#### Example 2: Set ICS for the specified public and private connections and generate an output.
`Set-Ics Ethernet 'VM Host-Only Network' -PassThru`

## Function: Get-Ics
### Syntax
```
Get-Ics [[-ConnectionNames] <string[]>] [-AllConnections] [<CommonParameters>]
```
### Description
Lists network connections where ICS is enabled, or optionally ICS status for the specified network connections. Outputs a PSCustomObject table.
### Parameters
#### -ConnectionNames
Name(s) of the network connection(s) to get ICS status for. Optional.
#### -AllConnections
If parameter **ConnectionNames** is omitted, `Get-Ics` by default only lists network connections where ICS is enabled. To list ICS status for all network connections, add the switch parameter **AllConnections**.
Cannot be combined with parameter **ConnectionNames**.
### Usage examples
#### Example 1: Gets ICS status for all network connections where ICS is enabled.
`Get-Ics`
#### Example 2: Gets ICS status for all network connections.
`Get-Ics -All`
#### Example 3: Gets ICS status for the specified network connections.
`Get-Ics -ConnectionNames Ethernet, Ethernet2, 'VM Host-Only Network'`

`Get-Ics Ethernet, Ethernet2, 'VM Host-Only Network'`

## Function: Disable-Ics
### Syntax
```
Disable-Ics [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```
### Description
Checks for if ICS is enabled for any network connection and, if so, disables ICS for those connections.
### Parameters
#### -PassThru
If this parameter is specified `Disable-Ics` returns an output with the disabled connections. Optional. By default `Disable-Ics` does not generate any output.
#### -Whatif
Shows what would happen if the function runs. The function is not run.
#### -Confirm
Prompts you for confirmation before each change the function makes.
### Usage examples
#### Example 1: Disable ICS for all connections.
`Disable-Ics`
#### Example 2: Disable ICS for all connections and generate an output.
`Disable-Ics -PassThru`
