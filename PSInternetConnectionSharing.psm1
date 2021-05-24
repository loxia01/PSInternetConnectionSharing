#Requires -Version 3.0

$TestAdmin = {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
        GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
    {
        Write-Error "This function requires administrator rights. Restart PowerShell using the Run as administrator option."`
            -Category PermissionDenied -ErrorAction Stop
    }
}

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
 It will also check for if ICS is already enabled for the specified network connection pair.

.PARAMETER PublicConnectionName
 The name of the network connection that internet connection will be shared from.

.PARAMETER PrivateConnectionName
 The name of the network connection that internet connection will be shared with.

.EXAMPLE
 Set-Ics -PublicConnectionName Ethernet -PrivateConnectionName 'VM Host-Only Network'

.EXAMPLE
 Set-Ics Ethernet 'VM Host-Only Network'

.NOTES
 Set-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online Version: https://github.com/loxia01/PSInternetConnectionSharing#set-ics
 Get-Ics
 Disable-Ics
#>
    
    param(
        [Parameter(Mandatory)]
        [String]$PublicConnectionName,
        
        [Parameter(Mandatory)]
        [String]$PrivateConnectionName
    )
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $connections = $netShare.EnumEveryConnection | ForEach-Object { $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
        if ($connections.Name -notcontains $PublicConnectionName)
        {
            Write-Error "$PublicConnectionName is not a valid network connection name." -Category InvalidData -ErrorAction Stop
        }
        if ($connections.Name -notcontains $PrivateConnectionName)
        {
            Write-Error "$PrivateConnectionName is not a valid network connection name." -Category InvalidData -ErrorAction Stop
        }
        if ($connections.Where({$_.Name -eq $PrivateConnectionName}).Status -eq 0)
        {
            Write-Error "$PrivateConnectionName connection is not enabled." -Category NotEnabled -ErrorAction Stop
        }
        
        $publicConnectionName = $connections.Name.Where({$_ -eq $PublicConnectionName})
        $privateConnectionName = $connections.Name.Where({$_ -eq $PrivateConnectionName})
    }
    
    process
    {
        $publicConnectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $publicConnectionName}
        $publicConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($publicConnectionProps)
        $privateConnectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $privateConnectionName}
        $privateConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($privateConnectionProps)
        
        if (($publicConnectionConfig.SharingEnabled -eq $true -and $publicConnectionConfig.SharingConnectionType -eq 0) -and
            ($privateConnectionConfig.SharingEnabled -eq $true -and $privateConnectionConfig.SharingConnectionType -eq 1))
        {
            Write-Host "`nICS is already set for $publicConnectionName (public connection) and $privateConnectionName (private connection).`n"`
                -ForegroundColor Yellow -BackgroundColor Black
        }
        else
        {
            foreach ($connectionName in $connections.Name)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                if ($connectionConfig.SharingEnabled -eq $true)
                {
                    $connectionConfig.DisableSharing()    
                }
            }
            
            $publicConnectionConfig.EnableSharing(0)
            $privateConnectionConfig.EnableSharing(1)
            
            Get-Ics -ConnectionNames $publicConnectionName, $privateConnectionName -NoValidation
        }
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
 or optionally for the specified network connections. Output is printed as a PSCustomObject table.

.PARAMETER ConnectionNames
 Name(s) of the network connection(s) to get ICS status for. Optional.
 
.PARAMETER NoValidation
 Optional switch parameter intended only for in-function use when
 calling Get-Ics with parameter ConnectionNames already validated.

.EXAMPLE
 # Gets status for ALL network connections.
 Get-Ics 

.EXAMPLE
 # Gets status for the specified network connections.
 Get-Ics -ConnectionNames Ethernet, Ethernet2, 'VM Host-Only Network'

.EXAMPLE
 # Gets status for the specified network connections. 
 Get-Ics Ethernet, Ethernet2, 'VM Host-Only Network'

.NOTES
 Get-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online Version: https://github.com/loxia01/PSInternetConnectionSharing#get-ics
 Set-Ics
 Disable-Ics
#>
    
    param(
        [Parameter()]
        [String[]]$ConnectionNames,
        
        [Parameter()]
        [Switch]$NoValidation
    )
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        if (-not $NoValidation)
        {
            $connections = $netShare.EnumEveryConnection | ForEach-Object { $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
            if ($ConnectionNames)
            {
                foreach ($ConnectionName in $ConnectionNames)
                {
                    if ($connections.Name -notcontains $ConnectionName)
                    {
                        Write-Error "$ConnectionName is not a valid network connection name." -Category InvalidData -ErrorAction Stop
                    }
                }
                $connectionNames = $connections.Name.Where({$_ -in $ConnectionNames})
            }
        }
    }
    
    process
    {
        if ($ConnectionNames)
        {
            $output = @()
            foreach ($connectionName in $connectionNames)
            {   
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                if ($connectionConfig.SharingEnabled -ne $true)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Disabled'}
                }
                if ($connectionConfig.SharingEnabled -eq $true -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (public)'}
                }
                if ($connectionConfig.SharingEnabled -eq $true -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (private)'}
                }
            }
            Sort-Object NetworkConnectionName -InputObject $output
        }
        else
        {
            $output = @()
            foreach ($connectionName in $connections.Name)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                if ($connectionConfig.SharingEnabled -ne $true)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Disabled'}
                }
                if ($connectionConfig.SharingEnabled -eq $true -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (public)'}
                }
                if ($connectionConfig.SharingEnabled -eq $true -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (private)'}
                }
            }
            Sort-Object NetworkConnectionName -InputObject $output
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

.EXAMPLE
 Disable-Ics

.NOTES
 Disable-Ics requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online Version: https://github.com/loxia01/PSInternetConnectionSharing#disable-ics
 Set-Ics
 Get-Ics
#>
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $connections = $netShare.EnumEveryConnection | ForEach-Object { $netShare.NetConnectionProps.Invoke($_) } | Where-Object Status -NE $null
    }
    
    process
    {
        $connectionsDisabled = @()
        foreach ($connectionName in $connections.Name)
        {
            $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
            if ($connectionConfig.SharingEnabled -eq $true)
            {
                $connectionConfig.DisableSharing()
                if ($? -eq $true) { $connectionsDisabled += $connectionName }
            }
        }
        if ($Error.Count -eq 0 -and $connectionsDisabled.Count -eq 0)
        {
            Write-Host "`nICS was already disabled for all network connections.`n" -ForegroundColor Yellow -BackgroundColor Black
        }
        else
        {
            Get-Ics -ConnectionNames $connectionsDisabled -NoValidation
        }
    }
}
