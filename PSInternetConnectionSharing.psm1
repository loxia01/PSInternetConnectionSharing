#Requires -Version 3.0

$TestAdmin = {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::
        GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
    {
        Write-Error "This function requires administrator rights. Restart PowerShell using the Run as administrator option."`
            -Category PermissionDenied -ErrorAction Stop
    }
}

function Set-ICS
{
<#
.SYNOPSIS
 Enables Internet Connection Sharing (ICS) for a specified network connection pair.

.DESCRIPTION
 Set-ICS lets you share the internet connection of a network connection (called the public
 connection) with another network connection (called the private connection).
 The specified network connections must exist beforehand. In order to be able to set ICS,
 the function will first disable ICS for any existing network connections.
 It will also check for if ICS is already enabled for the specified network connection pair.

.PARAMETER PublicConnectionName
 The name of the network connection that internet connection will be shared from.

.PARAMETER PrivateConnectionName
 The name of the network connection that internet connection will be shared with.

.EXAMPLE
 Set-ICS -PublicConnectionName Ethernet -PrivateConnectionName 'VM Host-Only Network'

.EXAMPLE
 Set-ICS Ethernet 'VM Host-Only Network'

.NOTES
 Set-ICS requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online Version: https://github.com/loxia01/PSInternetConnectionSharing#set-ics
 Get-ICS
 Disable-ICS
#>
    
    param(
        [Parameter(Mandatory)]
        [ValidateScript({
            if (((Get-NetAdapter -Name $_).Name -eq $_) -and ((Get-NetAdapter -Name $_).Status -notin ('Not Present','Disabled'))) { $true }
            else { throw "$_ is either not a valid network connection name or $_ connection is not enabled." }
        })]
        [String]$PublicConnectionName,
        
        [Parameter(Mandatory)]
        [ValidateScript({
            if (((Get-NetAdapter -Name $_).Name -eq $_) -and ((Get-NetAdapter -Name $_).Status -notin ('Not Present','Disabled'))) { $true }
            else { throw "$_ is either not a valid network connection name or $_ connection is not enabled." }
        })]
        [String]$PrivateConnectionName
    )
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
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
            $publicConnectionName = (Get-NetAdapter $PublicConnectionName).Name
            $privateConnectionName = (Get-NetAdapter $PrivateConnectionName).Name
            Write-Host "`nICS is already set for network connections $publicConnectionName (public) and $privateConnectionName (private).`n"
        }
        else
        { 
            $netAdapters = Get-NetAdapter
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
                    if ($? -eq 'True')
                    {
                        Write-Host "`nICS was enabled for network connection $connectionName (public connection)."
                    }
                }
                if ($connectionName -eq $PrivateConnectionName)
                {
                    $privateConnectionConfig.EnableSharing(1)
                    if ($? -eq 'True')
                    {
                        Write-Host "ICS was enabled for network connection $connectionName (private connection).`n"
                    }
                }
            }
        }
    }
}

function Get-ICS
{
<#
.SYNOPSIS
 Retrieves status of Internet Connection Sharing (ICS) for all network connections,
 or optionally for the specified network connections.

.DESCRIPTION
 Retrieves status of Internet Connection Sharing (ICS) for all network connections,
 or optionally for the specified network connections. Output is printed in the form of a hash table.

.PARAMETER ConnectionNames
 Name(s) of the network connection(s) to get ICS status for. Optional.

.EXAMPLE
 # Gets status for ALL network connections.
 Get-ICS 

.EXAMPLE
 # Gets status for the specified network connections.
 Get-ICS -ConnectionNames Ethernet, Ethernet2, 'VM Host-Only Network'

.EXAMPLE
 # Gets status for the specified network connections. 
 Get-ICS Ethernet, Ethernet2, 'VM Host-Only Network'

.NOTES
 Get-ICS requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online Version: https://github.com/loxia01/PSInternetConnectionSharing#get-ics
 Set-ICS
 Disable-ICS
#>
    
    param(
        [Parameter()]
        [ValidateScript({
            foreach ($ConnectionName in $_) {
                if ((Get-NetAdapter -Name $_).Name -eq $_) { $true }
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
                    $output += [PSCustomObject]@{NetworkConnectionName=$connectionName; StatusICS='Disabled'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $output += [PSCustomObject]@{NetworkConnectionName=$connectionName; StatusICS='Enabled (public)'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $output += [PSCustomObject]@{NetworkConnectionName=$connectionName; StatusICS='Enabled (private)'}
                }
            }
            $output | Sort-Object NetworkConnectionName | Format-Table
        }
        else
        {
            $netAdapters = Get-NetAdapter
            $output = @()
            foreach ($connectionName in $netAdapters.Name)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                
                if ($connectionConfig.SharingEnabled -ne 'True')
                {
                    $output += [PSCustomObject]@{NetworkConnectionName=$connectionName; StatusICS='Disabled'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $output += [PSCustomObject]@{NetworkConnectionName=$connectionName; StatusICS='Enabled (public)'}
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $output += [PSCustomObject]@{NetworkConnectionName=$connectionName; StatusICS='Enabled (private)'}
                }
            }
            $output | Sort-Object NetworkConnectionName | Format-Table
        }
    }
}

function Disable-ICS
{
<#
.SYNOPSIS
 Disables Internet Connection Sharing (ICS) for all network connections.

.DESCRIPTION
 Disable-ICS checks for if ICS is enabled for any network connections and, if so,
 disables ICS for those connections.

.EXAMPLE
 Disable-ICS

.NOTES
 Disable-ICS requires elevated permissions. Use the Run as administrator option when starting PowerShell.
 Testing for administrator rights is done in the beginning of function.

.LINK
 Online Version: https://github.com/loxia01/PSInternetConnectionSharing#disable-ics
 Set-ICS
 Get-ICS
#>
    
    begin
    {
        Invoke-Command -ScriptBlock $TestAdmin
        regsvr32.exe -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
    }
    
    process
    {
        $netAdapters = Get-NetAdapter
        $icsDisabled = 0
        
        foreach ($connectionName in $netAdapters.Name)
        {
            $connectionProps = $netShare.EnumEveryConnection | Where-Object {$netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName}
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
            
            if ($connectionConfig.SharingEnabled -eq 'True')
            {
                $connectionConfig.DisableSharing()
                $icsDisabled++
                if ($? -eq 'True' -and $icsDisabled -eq 1)
                {
                    Write-Host "`nICS was disabled for network connection $connectionName."
                }
                if ($? -eq 'True' -and $icsDisabled -eq 2)
                {
                    Write-Host "ICS was disabled for network connection $connectionName."
                }
            }
        }
        if ($icsDisabled -eq 0)
        {
            Write-Host "`nICS was already disabled for all network connections."
        }
        Write-Host ""
    }
}
