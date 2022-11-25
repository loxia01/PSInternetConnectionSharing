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
 If this parameter is specified Set-ICS returns an output with the set connections.
 Optional. By default Set-ICS does not generate any output.
 
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
 Set-ICS does not take pipeline input.
 
.OUTPUTS
 Default is no output. If parameter PassThru is specified Set-ICS returns a PSCustomObject.
 
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
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
            GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
        {
            Write-Error "This function requires administrator rights. Restart PowerShell using the Run as Administrator option."`
                -Category PermissionDenied -ErrorAction Stop
        }
        
        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $connectionsProps = $netShare.EnumEveryConnection | ForEach-Object {
            $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
        
        if ($connectionsProps.Name -notcontains $PublicConnectionName)
        {
            Write-Error "'${PublicConnectionName}' is not a valid network connection name." -Category InvalidArgument -ErrorAction Stop
        }
        if ($connectionsProps.Name -notcontains $PrivateConnectionName)
        {
            Write-Error "'${PrivateConnectionName}' is not a valid network connection name." -Category InvalidArgument -ErrorAction Stop
        }
        if ($connectionsProps.Name.Where({$_.Name -eq $PrivateConnectionName}).Status -eq 0)
        {
            Write-Error "Private connection '${PrivateConnectionName}' must be enabled to set ICS." -Category NotEnabled -ErrorAction Stop
        }
        
        $publicConnectionName = $connectionsProps.Name.Where({$_ -eq $PublicConnectionName})
        $privateConnectionName = $connectionsProps.Name.Where({$_ -eq $PrivateConnectionName})
    }
    process
    {
        $publicConnection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $publicConnectionName}
        $publicConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($publicConnection)
        $privateConnection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $privateConnectionName}
        $privateConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($privateConnection)
        
        
        if (-not (($publicConnectionConfig.SharingEnabled -eq $true -and $publicConnectionConfig.SharingConnectionType -eq 0) -and
            ($privateConnectionConfig.SharingEnabled -eq $true -and $privateConnectionConfig.SharingConnectionType -eq 1)))
        {
            foreach ($connectionName in $connectionsProps.Name)
            {
                $connection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)

                if ($connectionConfig.SharingEnabled -eq $true)
                {
                    if ($PSCmdlet.ShouldProcess($connectionName, "DisableICS")) { $connectionConfig.DisableSharing() }
                }
            }
            if ($PSCmdlet.ShouldProcess($publicConnectionName)) { $publicConnectionConfig.EnableSharing(0) }
            if ($PSCmdlet.ShouldProcess($privateConnectionName)) { $privateConnectionConfig.EnableSharing(1) }
        }
    }
    end
    {
        if ($PassThru -and ($WhatIfPreference -eq $false)) { Get-Ics -ConnectionNames $publicConnectionName, $privateConnectionName }
    }
}

function Get-Ics
{
<#
.SYNOPSIS
 Retrieves status of Internet Connection Sharing (ICS) for all network connections,
 or optionally for the specified network connections.
 
.DESCRIPTION
 Retrieves status of Internet Connection Sharing (ICS) for all network connections,
 or optionally for the specified network connections. Output is in the form of a PSCustomObject.
 
.PARAMETER ConnectionNames
 Name(s) of the network connection(s) to get ICS status for. Optional.
 
.PARAMETER ShowEnabled 
 By default Get-Ics lists ICS status for all network connections if parameter ConnectionNames is omitted.
 When adding parameter ShowEnabled, Get-Ics only lists connections where ICS is enabled.
 
.EXAMPLE
 Get-Ics
 
 # Gets status for all network connections.
 
.EXAMPLE
 Get-Ics
 
 # Gets status for all network connections with ICS enabled.
 
.EXAMPLE
 Get-Ics -ConnectionNames Ethernet, Ethernet2, 'VM Host-Only Network'
 
 # Gets status for the specified network connections.
 
.EXAMPLE
 Get-Ics Ethernet, Ethernet2, 'VM Host-Only Network'
 
 # Gets status for the specified network connections.
 
.INPUTS
 Get-ICS does not take pipeline input.
 
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
    [CmdletBinding()]
    param (
        [string[]]$ConnectionNames,
        
        [switch]$ShowEnabled
    )
    
    begin
    {
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
            GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
        {
            Write-Error "This function requires administrator rights. Restart PowerShell using the Run as Administrator option."`
                -Category PermissionDenied -ErrorAction Stop
        }
        
        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        if ((Get-PSCallStack)[1].Command -notmatch '(Disable|Set)-Ics')
        {
            $connectionsProps = $netShare.EnumEveryConnection | ForEach-Object {
                $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
            
            if ($ConnectionNames)
            {
                foreach ($ConnectionName in $ConnectionNames)
                {
                    if ($connectionsProps.Name -notcontains $ConnectionName)
                    {
                        Write-Error "'${ConnectionName}' is not a valid network connection name." -Category InvalidArgument -ErrorAction Stop
                    }
                }
                $connectionNames = $connectionsProps.Name.Where({$_ -in $ConnectionNames})
            }
            else
            {
                $connectionNames = $connectionsProps.Name
            }
        }
        else
        {
            if (-not $ConnectionNames) { $connectionNames = $connectionsProps.Name }
        }
    }
    process
    {
        $output = @()
        foreach ($connectionName in $connectionNames)
        {   
            $connection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
            
            if ($connectionConfig.SharingEnabled -eq $false)
            {
                $output += [pscustomobject]@{NetworkConnectionName=$connectionName + "    "; StatusICS='Disabled'; Type=$null}
            }
            if ($connectionConfig.SharingEnabled -eq $true -and $connectionConfig.SharingConnectionType -eq 0)
            {
                $output += [pscustomobject]@{NetworkConnectionName=$connectionName + "    "; StatusICS='Enabled'; Type='Public'}
            }
            if ($connectionConfig.SharingEnabled -eq $true -and $connectionConfig.SharingConnectionType -eq 1)
            {
                $output += [pscustomobject]@{NetworkConnectionName=$connectionName + "    "; StatusICS='Enabled'; Type='Private'}
            }
        }
    }
    end
    {
        if ($ShowEnabled)
        {
            $output | Sort-Object Type -Descending | Where-Object StatusICS -EQ Enabled | Out-Host
        }
        elseif ((Get-PSCallStack)[1].Command -ne 'Disable-Ics')
        {
            $output | Sort-Object StatusICS, Type -Descending | Out-Host
        }
        else
        {
            $output | Select-Object NetworkConnectionName, StatusICS | Out-Host
        }
    }
}

function Disable-Ics
{
<#
.SYNOPSIS
 Disables Internet Connection Sharing (ICS) for all network connections.
 
.DESCRIPTION
 Disable-Ics checks for if ICS is enabled for any network connections and, if so,
 disables ICS for those connections.
 
.PARAMETER PassThru
 If this parameter is specified Disable-ICS returns an output with the disabled connections.
 Optional. By default Disable-ICS does not generate any output.
 
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
 Disable-ICS does not take pipeline input.
 
.OUTPUTS
 Default is no output. If parameter PassThru is specified Disable-ICS returns a PSCustomObject.
 
.NOTES
 Disable-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.
 
.LINK
 Online version: https://github.com/loxia01/PSInternetConnectionSharing#disable-ics
 Set-Ics
 Get-Ics
#>
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [switch]$PassThru
    )
    
    begin
    {
        if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
            GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
        {
            Write-Error "This function requires administrator rights. Restart PowerShell using the Run as Administrator option."`
                -Category PermissionDenied -ErrorAction Stop
        }
        
        regsvr32 /s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $connectionsProps = $netShare.EnumEveryConnection | ForEach-Object {
            $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
    }
    process
    {
        $disabledNames = @()
        foreach ($connectionName in $connectionsProps.Name)
        {
            $connection = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connection)
            if ($connectionConfig.SharingEnabled -eq $true)
            {
                if ($PSCmdlet.ShouldProcess($connectionName))
                {
                    $connectionConfig.DisableSharing()
                    $disabledNames += $connectionName
                }
            }
        }
    }
    end
    {
        if ($PassThru -and ($WhatIfPreference -eq $false)) { Get-Ics $disabledNames }
    }
}
