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
        [SupportsWildcards()]
        [string]$PublicConnectionName,
        
        [Parameter(Mandatory)]
        [SupportsWildcards()]
        [string]$PrivateConnectionName,
        
        [switch]$PassThru
    )
    
    begin
    {
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
            GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator'))
        {
            Write-Error "This function requires administrator rights. Restart PowerShell using the Run as Administrator option."`
                -Category PermissionDenied -ErrorAction Stop
        }
        
        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $connectionsProps = $netShare.EnumEveryConnection | ForEach-Object {$netShare.NetConnectionProps.Invoke($_)} |
            Where-Object Status -NE $null
        
        Get-Variable PublicConnectionName, PrivateConnectionName | ForEach-Object {
            if (-not ($connectionsProps.Name -like $_.Value))
            {
                Write-Error "'$($_.Value)' is not a valid network connection name." -Category InvalidArgument -ErrorAction Stop
            }
            elseif (($connectionsProps.Name -like $_.Value).Count -gt 1)
            {
                Write-Error "'$($_.Value)' resolved to multiple connection names: `n$(($connectionsProps.Name -like $_.Value) -join "`n")`n"`
                    -Category InvalidArgument -ErrorAction Stop
            }
            else
            {
                $_.Value = $connectionsProps.Name -like $_.Value
            }
        }
        
        if ($connectionsProps.Where({$_.Name -eq $PrivateConnectionName}).Status -eq 0)
        {
            Write-Error "Private connection '${PrivateConnectionName}' must be enabled to set ICS." -Category NotEnabled -ErrorAction Stop
        }
    }
    process
    {
        $publicConnection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $PublicConnectionName}
        $publicConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($publicConnection)
        $privateConnection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $PrivateConnectionName}
        $privateConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($privateConnection)
        
        
        if (-not (($publicConnectionConfig.SharingEnabled -and $publicConnectionConfig.SharingConnectionType -eq 0) -and
            ($privateConnectionConfig.SharingEnabled -and $privateConnectionConfig.SharingConnectionType -eq 1)))
        {
            foreach ($connectionName in $connectionsProps.Name)
            {
                $connection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)

                if ($connectionConfig.SharingEnabled)
                {
                    if ($PSCmdlet.ShouldProcess($connectionName, "Disable-Ics")) { $connectionConfig.DisableSharing() }
                }
            }
            if ($PSCmdlet.ShouldProcess($PublicConnectionName)) { $publicConnectionConfig.EnableSharing(0) }
            if ($PSCmdlet.ShouldProcess($PrivateConnectionName)) { $privateConnectionConfig.EnableSharing(1) }
        }
    }
    end
    {
        if ($PassThru -and $WhatIfPreference -eq $false) { Get-Ics -ConnectionNames $PublicConnectionName, $PrivateConnectionName }
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
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
            GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator'))
        {
            Write-Error "This function requires administrator rights. Restart PowerShell using the Run as Administrator option."`
                -Category PermissionDenied -ErrorAction Stop
        }
        
        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        if ($MyInvocation.PSCommandPath -notmatch 'PSInternetConnectionSharing.psm1')
        {
            $connectionsProps = $netShare.EnumEveryConnection | ForEach-Object {$netShare.NetConnectionProps.Invoke($_)} |
                Where-Object Status -NE $null
            
            if ($ConnectionNames)
            {
                $ConnectionNames = foreach ($connectionName in $ConnectionNames)
                {
                    if (-not ($connectionsProps.Name -like $connectionName))
                    {
                        Write-Error "'${connectionName}' is not a valid network connection name." -Category InvalidArgument
                    }
                    else
                    {
                        $connectionsProps.Name -like $connectionName
                    }
                }
            }
        }
        if (-not $ConnectionNames) { $connectionNames = $connectionsProps.Name }
    }
    process
    {
        $output = foreach ($connectionName in $connectionNames)
        {   
            $connection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
            
            if (-not $connectionConfig.SharingEnabled)
            {
                [pscustomobject]@{ConnectionName = $connectionName; ICSEnabled = $false}
            }
            if ($connectionConfig.SharingEnabled -and $connectionConfig.SharingConnectionType -eq 0)
            {
                [pscustomobject]@{ConnectionName = $connectionName; ICSEnabled = $true; ConnectionType = 'Public'}
            }
            if ($connectionConfig.SharingEnabled -and $connectionConfig.SharingConnectionType -eq 1)
            {
                [pscustomobject]@{ConnectionName = $connectionName; ICSEnabled = $true; ConnectionType = 'Private'}
            }
        }
    }
    end
    {
        if ($AllConnections -or $PSBoundParameters.ContainsKey('ConnectionNames'))
        {
             return $output | Sort-Object ICSEnabled, ConnectionType -Descending
        }
        else
        {
             return $output | Where-Object ICSEnabled | Sort-Object ConnectionType -Descending
        }
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
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
            GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator'))
        {
            Write-Error "This function requires administrator rights. Restart PowerShell using the Run as Administrator option."`
                -Category PermissionDenied -ErrorAction Stop
        }
        
        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $connectionsProps = $netShare.EnumEveryConnection | ForEach-Object {$netShare.NetConnectionProps.Invoke($_)} |
            Where-Object Status -NE $null
    }
    process
    {
        $disabledNames = foreach ($connectionName in $connectionsProps.Name)
        {
            $connection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
            if ($connectionConfig.SharingEnabled)
            {
                if ($PSCmdlet.ShouldProcess($connectionName))
                {
                    $connectionConfig.DisableSharing()
                    $connectionName
                }
            }
        }
    }
    end
    {
        if ($PassThru -and $WhatIfPreference -eq $false)
        {
            if ($disabledNames) { Get-Ics -ConnectionNames $disabledNames }
            else                { Get-Ics -AllConnections }
        }
    }
}
