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
        [ValidateScript({
            if ((Get-NetAdapter -Name $_).Name -eq $_ -and (Get-NetAdapter -Name $_).Status -notin 'Not Present', 'Disabled', $null) { $true }
            else { throw "$_ is either not a valid network connection name or $_ connection is not enabled." }
        })]
        [String]$PublicConnectionName,
        
        [Parameter(Mandatory)]
        [ValidateScript({
            if ((Get-NetAdapter -Name $_).Name -eq $_ -and (Get-NetAdapter -Name $_).Status -notin 'Not Present', 'Disabled', $null) { $true }
            else { throw "$_ is either not a valid network connection name or $_ connection is not enabled." }
        })]
        [String]$PrivateConnectionName
    )
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
        
        $publicConnectionName = (Get-NetAdapter $PublicConnectionName).Name
        $privateConnectionName = (Get-NetAdapter $PrivateConnectionName).Name
    }
    
    process
    {
        $publicConnectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $PublicConnectionName}
        $publicConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($publicConnectionProps)
        $privateConnectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $PrivateConnectionName}
        $privateConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($privateConnectionProps)
        
        if (($publicConnectionConfig.SharingEnabled -eq 'True' -and $publicConnectionConfig.SharingConnectionType -eq 0) -and
            ($privateConnectionConfig.SharingEnabled -eq 'True' -and $privateConnectionConfig.SharingConnectionType -eq 1))
        {
            Write-Host "`nICS is already enabled for $publicConnectionName (public connection) and $privateConnectionName (private connection).`n"
        }
        else
        { 
            $netAdapters = Get-NetAdapter | Where-Object {$_.Status -ne $null}
            foreach ($connectionName in $netAdapters.Name)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                
                if ($connectionConfig.SharingEnabled -eq 'True')
                {
                    $connectionConfig.DisableSharing()    
                }
            }
            foreach ($connectionName in $netAdapters.Name)
            {
                if ($connectionName -eq $PublicConnectionName)
                {
                    $publicConnectionConfig.EnableSharing(0)          
                }
                if ($connectionName -eq $PrivateConnectionName)
                {
                    $privateConnectionConfig.EnableSharing(1)
                }
            }
            if ($Error.Count -eq 0)
            {
                Write-Host "`nICS was enabled for $publicConnectionName (public connection) and $privateConnectionName (private connection).`n"
            }
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
        [ValidateScript({
            foreach ($ConnectionName in $_) {
                if (((Get-NetAdapter -Name $_).Name -eq $_) -and ((Get-NetAdapter -Name $_).Status -ne $null)) { $true }
                else { throw "$ConnectionName is not a valid network connection name." }}
        })]
        [String[]]$ConnectionNames
    )
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
    }
    
    process
    {
        if ($ConnectionNames)
        {
            $output = @()
            foreach ($ConnectionName in $ConnectionNames)
            {
                $connectionName = (Get-NetAdapter -Name $ConnectionName).Name
                
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                
                if ($connectionConfig.SharingEnabled -ne 'True')
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Disabled'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (public)'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (private)'}
                }
            }
            $output | Sort-Object NetworkConnectionName | Format-Table
        }
        else
        {
            $netAdapters = Get-NetAdapter | Where-Object {$_.Status -ne $null}
            $output = @()
            foreach ($connectionName in $netAdapters.Name)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                
                if ($connectionConfig.SharingEnabled -ne 'True')
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Disabled'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (public)'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $output += [PSCustomObject]@{'NetworkConnectionName'=$connectionName;'StatusICS'='Enabled (private)'}
                }
            }
            $output | Sort-Object NetworkConnectionName | Format-Table
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
    }
    
    process
    {
        $netAdapters = Get-NetAdapter | Where-Object {$_.Status -ne $null}
        
        $connectionsDisabled = @()
        foreach ($connectionName in $netAdapters.Name)
        {
            $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
            
            if ($connectionConfig.SharingEnabled -eq 'True')
            {
                $connectionConfig.DisableSharing()
                $connectionsDisabled += $connectionName
            }
        }
        if ($Error.Count -eq 0)
        {
            if ($connectionsDisabled.Count -eq 0)
            {
                Write-Host "`nICS was already disabled for all network connections.`n"
            }
            if ($connectionsDisabled.Count -eq 1)
            {
                Write-Host "`nICS was disabled for network connection $($connectionsDisabled[0]).`n"
            }
            if ($connectionsDisabled.Count -eq 2)
            {
                Write-Host "`nICS was disabled for network connections $($connectionsDisabled[0]) and $($connectionsDisabled[1]).`n"
            }
        }
    }
}
