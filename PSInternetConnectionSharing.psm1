using namespace System.Security.Principal
using namespace System.Management.Automation

function Set-Ics
{
<#
.SYNOPSIS
 Enables Internet Connection Sharing (ICS) for a specified network connection pair.

.DESCRIPTION
 Set-Ics lets you share the internet connection of a network connection (called the public
 connection) with another network connection (called the private connection).
 The specified network connections must exist beforehand. In order to be able to set ICS,
 the function will first disable ICS for any existing network connections.

.PARAMETER PublicConnectionName
 The name of the network connection that internet connection will be shared from.

.PARAMETER PrivateConnectionName
 The name of the network connection that internet connection will be shared with.

.PARAMETER PassThru
 If this parameter is specified Set-Ics returns an object representing the set connections.
 Optional. By default Set-Ics does not generate any output.

.PARAMETER WhatIf
 Shows what would happen if the function runs. The function is not run.

.PARAMETER Confirm
 Prompts you for confirmation before each change the function makes.

.EXAMPLE
 Set-Ics -PublicConnectionName 'Ethernet' -PrivateConnectionName 'VM Host-Only Network'

 # Sets ICS for the specified public and private connections.

.EXAMPLE
 Set-Ics Ethernet 'VM Host-Only Network'

 # Sets ICS for the specified public and private connections.

.EXAMPLE
 Set-Ics Ethernet 'VM Host-Only Network' -PassThru

 # Sets ICS for the specified public and private connections and generates an output.

.INPUTS
 Set-Ics does not take pipeline input.

.OUTPUTS
 Default is no output. If parameter PassThru is specified Set-Ics returns a PSCustomObject.

.NOTES
 Set-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online version: https://github.com/loxia01/PSInternetConnectionSharing#set-ics
 Get-Ics
 Disable-Ics
#>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [string]$PublicConnectionName,

        [Parameter(Mandatory)]
        [string]$PrivateConnectionName,

        [switch]$PassThru
    )

    begin
    {
        if (-not ([WindowsPrincipal][WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]'Administrator'))
        {
            $exception = "This function requires administrator rights."
            $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'AdminPrivilegeRequired', 18, $null))
        }

        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare

        $connections = @($netShare.EnumEveryConnection)
        $connectionsProps = $connections | ForEach-Object {$netShare.NetConnectionProps.Invoke($_)}

        Get-Variable PublicConnectionName, PrivateConnectionName | ForEach-Object {
            if ($connectionsProps.Name -notcontains $_.Value)
            {
                $exception = New-Object PSArgumentException "Cannot find a network connection with name '$($_.Value)'."
                $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'ConnectionNotFound', 13, $null))
            }
            else { $_.Value = $connectionsProps | Where-Object Name -EQ $_.Value | Select-Object -ExpandProperty Name }
        }

        if ($PrivateConnectionName -eq $PublicConnectionName)
        {
            $exception = New-Object PSArgumentException "The private connection cannot be the same as the public connection."
            $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'InvalidConnection', 5, $null))
        }
        if (($connectionsProps | Where-Object Name -EQ $PrivateConnectionName).Status -eq 0)
        {
            $exception = "Private connection '${PrivateConnectionName}' must be enabled to set ICS."
            $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'ConnectionNotEnabled', 31, $null))
        }
    }
    process
    {
        foreach ($connection in $connections)
        {
            try
            {
                if ($netShare.NetConnectionProps.Invoke($connection).Name -eq $PublicConnectionName)
                {
                    $publicConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
                }
                elseif ($netShare.NetConnectionProps.Invoke($connection).Name -eq $PrivateConnectionName)
                {
                    $privateConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
                }
                else { continue }
            }
            catch
            {
                $exception = New-Object RuntimeException -Args (
                    "ICS is not possible to set for connection '$($netShare.NetConnectionProps.Invoke($connection).Name)'.", $_.Exception
                )
                $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, $_.CategoryInfo.Reason, $_.CategoryInfo.Category, $null))
            }
        }

        if (-not (($publicConnectionConfig.SharingEnabled -and $publicConnectionConfig.SharingConnectionType -eq 0) -and
            ($privateConnectionConfig.SharingEnabled -and $privateConnectionConfig.SharingConnectionType -eq 1)))
        {
            $icsConnections = foreach ($method in "EnumPublicConnections","EnumPrivateConnections")
            {
                try
                {
                    [pscustomobject]@{
                          Name = $netshare.$method.Invoke(0) | ForEach-Object {$netShare.NetConnectionProps.Invoke($_).Name}
                        Config = $netshare.$method.Invoke(0) | ForEach-Object {$netShare.INetSharingConfigurationForINetConnection.Invoke($_)}
                    }
                }
                catch { continue }
            }
            if ($icsConnections -and $PSCmdlet.ShouldProcess(($icsConnections.Name -join ", "), 'Disable-Ics'))
            {
                $icsConnections.Config | ForEach-Object DisableSharing
            }

            if ($PSCmdlet.ShouldProcess(($PublicConnectionName, $PrivateConnectionName) -join ", "))
            {
                $publicConnectionConfig.EnableSharing(0)
                $privateConnectionConfig.EnableSharing(1)
            }
        }
    }
    end
    {
        if ($PassThru -and -not $WhatIfPreference) { Get-Ics }
    }
}

function Get-Ics
{
<#
.SYNOPSIS
 Gets Internet Connection Sharing (ICS) status for network connections.

.DESCRIPTION
 Gets network connections where ICS is enabled, or optionally for all specified network connections.
 Output is a PSCustomObject representing the connections.

.PARAMETER ConnectionNames
 Name(s) of the network connection(s) to get ICS status for. Optional.

.PARAMETER AllConnections
 If parameter ConnectionNames is omitted, Get-Ics by default only lists network connections where ICS is enabled.
 To list ICS status for all network connections, add the switch parameter AllConnections.
 Cannot be combined with parameter ConnectionNames.

.EXAMPLE
 Get-Ics

 # Gets ICS status for network connections where ICS is enabled.

.EXAMPLE
 Get-Ics -AllConnections

 # Gets ICS status for all network connections.

.EXAMPLE
 Get-Ics -ConnectionNames Ethernet, Ethernet2, 'VM Host-Only Network'

 # Gets ICS status for the specified network connections.

.EXAMPLE
 Get-Ics Ethernet, Ethernet2, 'VM Host-Only Network'

 # Gets ICS status for the specified network connections.

.INPUTS
 Get-Ics does not take pipeline input.

.OUTPUTS
 PSCustomObject.

.NOTES
 Get-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online version: https://github.com/loxia01/PSInternetConnectionSharing#get-ics
 Set-Ics
 Disable-Ics
#>
    [CmdletBinding(DefaultParameterSetName='ConnectionNames')]
    param (
        [Parameter(ParameterSetName='ConnectionNames', Position=0)]
        [SupportsWildcards()]
        [string[]]$ConnectionNames,

        [Parameter(ParameterSetName='AllConnections')]
        [switch]$AllConnections
    )

    begin
    {
        if ($PSCmdlet.MyInvocation.PSCommandPath -notmatch 'PSInternetConnectionSharing.psm1$')
        {
            if (-not ([WindowsPrincipal][WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]'Administrator'))
            {
                $exception = "This function requires administrator rights."
                $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'AdminPrivilegeRequired', 18, $null))
            }

            regsvr32 /s hnetcfg.dll
            $netShare = New-Object -ComObject HNetCfg.HNetShare

            $connections = @($netShare.EnumEveryConnection)
            $connectionsProps = $connections | ForEach-Object {$netShare.NetConnectionProps.Invoke($_)}

            if ($ConnectionNames)
            {
                $ConnectionNames = foreach ($connectionName in $ConnectionNames)
                {
                    if (-not ($connectionsProps.Name -like $connectionName))
                    {
                        $exception = New-Object PSArgumentException "Cannot find a network connection with name '$($_.Value)'."
                        $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'ConnectionNotFound', 13, $null))
                    }
                    else
                    {
                        $connectionsProps.Name -like $connectionName
                    }
                }
            }
            elseif ($AllConnections) { $ConnectionNames = $connectionsProps.Name }
            else {}
        }
        else
        {
            if (-not $connections) { $connections = @($netShare.EnumEveryConnection) }
        }
    }
    process
    {
        if ($ConnectionNames)
        {
            $output = foreach ($connectionName in $ConnectionNames)
            {
                $connection = $connections | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                try   { $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection) }
                catch
                { 
                    $connectionConfig = $null
                    if (-not $AllConnections) { Write-Warning "ICS is not settable for connection '${connectionName}'." }
                }

                if ($connectionConfig.SharingEnabled -eq 1)
                {
                    [pscustomobject]@{
                        ConnectionName = $connectionName
                            ICSEnabled = $true
                        ConnectionType = if ($connectionConfig.SharingConnectionType -eq 0) { 'Public' } else { 'Private' }
                    }
                }
                else
                {
                    [pscustomobject]@{
                        ConnectionName = $connectionName
                            ICSEnabled = if ($connectionConfig.SharingEnabled -eq 0) { $false } else { $null }
                    }
                }
            }
        }
        else
        {
            $output = foreach ($method in "EnumPublicConnections","EnumPrivateConnections")
            {
                try
                {
                    [pscustomobject]@{
                          ConnectionName = $netshare.$method.Invoke(0) | ForEach-Object {$netShare.NetConnectionProps.Invoke($_).Name}
                              ICSEnabled = $true
                          ConnectionType = if ($method -match 'Public') { 'Public' } else { 'Private' }
                    }
                }
                catch { continue }
            }
        }
    }
    end
    {
        if ($PSCmdlet.MyInvocation.PSCommandPath -notmatch 'PSInternetConnectionSharing.psm1$')
        {
            $output | Sort-Object ConnectionType -Descending
        }
        else { $output }
    }
}

function Disable-Ics
{
<#
.SYNOPSIS
 Disables Internet Connection Sharing (ICS) for all network connections.

.DESCRIPTION
 Disable-Ics checks for if ICS is enabled for any network connections and disables ICS for those connections.

.PARAMETER PassThru
 If this parameter is specified Disable-Ics returns an object representing the disabled connections.
 Optional. By default Disable-Ics does not generate any output.

.PARAMETER WhatIf
 Shows what would happen if the function runs. The function is not run.

.PARAMETER Confirm
 Prompts you for confirmation before each change the function makes.

.EXAMPLE
 Disable-Ics

 # Disables ICS for all connections.

.EXAMPLE
 Disable-Ics -PassThru

 # Disables ICS for all connections and generates an output.

.INPUTS
 Disable-Ics does not take pipeline input.

.OUTPUTS
 Default is no output. If parameter PassThru is specified Disable-Ics returns a PSCustomObject.

.NOTES
 Disable-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online version: https://github.com/loxia01/PSInternetConnectionSharing#disable-ics
 Set-Ics
 Get-Ics
#>
    [CmdletBinding(SupportsShouldProcess)]
    param ([switch]$PassThru)

    begin
    {
        if (-not ([WindowsPrincipal][WindowsIdentity]::GetCurrent()).IsInRole([WindowsBuiltInRole]'Administrator'))
        {
            $exception = "This function requires administrator rights."
            $PSCmdlet.ThrowTerminatingError((New-Object ErrorRecord -Args $exception, 'AdminPrivilegeRequired', 18, $null))
        }

        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
    }
    process
    {
        $icsConnections = foreach ($method in "EnumPublicConnections","EnumPrivateConnections")
        {
            try
            {
                [pscustomobject]@{
                      Name = $netShare.$method.Invoke(0) | ForEach-Object {$netShare.NetConnectionProps.Invoke($_).Name}
                    Config = $netShare.$method.Invoke(0) | ForEach-Object {$netShare.INetSharingConfigurationForINetConnection.Invoke($_)}
                }
            }
            catch { continue }
        }
        if ($icsConnections -and $PSCmdlet.ShouldProcess($icsConnections.Name -join ", "))
        {
            $icsConnections.Config | ForEach-Object DisableSharing
        }
    }
    end
    {
        if ($PassThru -and -not $WhatIfPreference) { Get-Ics -ConnectionNames $icsConnections.Name }
    }
}
