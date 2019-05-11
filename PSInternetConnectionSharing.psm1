#Requires -RunAsAdministrator

Function Set-ICS
{   
<#
.SYNOPSIS
 Enables Internet Connection Sharing (ICS) for a specified network connection pair.

.DESCRIPTION
 Set-ICS lets you share the internet connection of a network connection (called the public connection) with another
 network connection (called the private connection). The specified network connections must exist beforehand.
 In order to be able to set ICS, the function will first disable ICS for any existing network connections.
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
 Version: 1.0
 Author:  Per Allner
 Website: https://github.com/loxia01/PSInternetConnectionSharing
#>
    [CmdletBinding(PositionalBinding = $true)]
    Param
    (
        [Parameter(Mandatory = $true)]
	[ValidateScript({Get-Netadapter -Name $_
        if ($? -eq 'True') { $true }
	else { throw "$_ is not a valid network connection name." }})]
        [string]
        $PublicConnectionName,
		
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Get-Netadapter -Name $_
        if ($? -eq 'True') { $true }
	else { throw "$_ is not a valid network connection name." }})]
        [string]
        $PrivateConnectionName
    )
    
    Begin
    {
        regsvr32 -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
    }

    Process
    {
        $publicConnectionProps = $netShare.EnumEveryConnection | Where-Object { $netShare.NetConnectionProps.Invoke($_).Name -eq $PublicConnectionName }
        $publicConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($publicConnectionProps)
        $privateConnectionProps = $netShare.EnumEveryConnection | Where-Object { $netShare.NetConnectionProps.Invoke($_).Name -eq $PrivateConnectionName }
        $privateConnectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($privateConnectionProps)
        
        if (($publicConnectionConfig.SharingEnabled -eq 'True' -and $publicConnectionConfig.SharingConnectionType -eq 0) -and `
            ($privateConnectionConfig.SharingEnabled -eq 'True' -and $privateConnectionConfig.SharingConnectionType -eq 1))
        { 
	    Write-Host "ICS is already set for network connections $PublicConnectionName (public) and $PrivateConnectionName (private)."
        }
        else
        { 
	    $netAdapters = Get-NetAdapter
	    foreach ($connectionName in $netAdapters.Name)
            {
	    	$connectionProps = $netShare.EnumEveryConnection | Where-Object { $netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName }
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
                    if ($? -eq 'True') { Write-Host "ICS was enabled for network connection $connectionName (public connection)." }
                }
            }
            foreach ($connectionName in $netAdapters.Name)
            {
                if ($connectionName -eq $PrivateConnectionName)
                {
                    $privateConnectionConfig.EnableSharing(1)
                    if ($? -eq 'True') { Write-Host "ICS was enabled for network connection $connectionName (private connection)." }
               	}
	    }
        }
    }
}

Function Disable-ICS
{
<#
.SYNOPSIS
 Disables Internet Connection Sharing (ICS) for all network connections.

.DESCRIPTION
 Disable-ICS checks for if ICS is enabled for any network connections and, if so, disables ICS for all connections. 

.PARAMETER
 None

.EXAMPLE
 Get-ICS

.NOTES
 Version: 1.0
 Author:  Per Allner
 Website: https://github.com/loxia01/PSInternetConnectionSharing
#>   
    [CmdletBinding()]
    Param()

    Begin
    {
        regsvr32 -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
    }

    Process
    {
        $netAdapters = Get-NetAdapter
        foreach ($connectionName in $netAdapters.Name)
        {
            $connectionProps = $netShare.EnumEveryConnection | Where-Object { $netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName }
            $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
            
            if ($connectionConfig.SharingEnabled -eq 'True')
            {
	        $connectionConfig.DisableSharing()
                if ($? -eq 'True') { Write-Host "ICS was disabled for network connection $connectionName." }       
            }
        }
    }
}

Function Get-ICS
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
 Get-ICS -ConnectionNames Ethernet,Ethernet2,'VM Host-Only Network'

.EXAMPLE
 # Gets status for the specified network connections. 
 Get-ICS Ethernet,Ethernet2,'VM Host-Only Network'

.NOTES
 Version: 1.0
 Author:  Per Allner
 Website: https://github.com/loxia01/PSInternetConnectionSharing
#>  
    [CmdletBinding(PositionalBinding = $true)]
    Param
    (
        [Parameter(Mandatory = $false)]
	[ValidateScript({foreach ($connectionName in $_) {
        Get-Netadapter -Name $connectionName
        if ($? -eq 'True') { $true }
	else { throw "$connectionName is not a valid network connection name." }}})]
        [string[]]
        $ConnectionNames
    )

    Begin
    {
        regsvr32 -s hnetcfg.dll
        $netShare = New-Object -ComObject HNetCfg.HNetShare
    }

    Process
    {
        if (!$ConnectionNames)
        {  
            $netAdapters = Get-NetAdapter
            foreach ($connectionName in $netAdapters.Name)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object { $netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName }
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                
                $connectionConfig | Add-Member -Type NoteProperty -Name NetworkConnectionName -Value $connectionName
            
                if ($connectionConfig.SharingEnabled -ne 'True')
                {
                    $connectionConfig | Add-Member -Type NoteProperty -Name InternetConnectionSharing  -Value 'Disabled'
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $connectionConfig | Add-Member -Type NoteProperty -Name InternetConnectionSharing -Value 'Enabled (public)'
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $connectionConfig | Add-Member -Type NoteProperty -Name InternetConnectionSharing -Value 'Enabled (private)'
                }
            
                Write-Output $connectionConfig | Select-Object -Property NetworkConnectionName,InternetConnectionSharing
            }
        }
        else
        {
            foreach ($connectionName in $ConnectionNames)
            {
                $connectionProps = $netShare.EnumEveryConnection | Where-Object { $netShare.NetConnectionProps.Invoke($_).Name -eq $connectionName }
                $connectionConfig = $netShare.INetSharingConfigurationForINetConnection.Invoke($connectionProps)
                
                $networkConnection = Get-NetAdapter -Name $connectionName
                $connectionConfig | Add-Member -Type NoteProperty -Name NetworkConnectionName -Value $networkConnection.Name
            
                if ($connectionConfig.SharingEnabled -ne 'True')
                {
                    $connectionConfig | Add-Member -Type NoteProperty -Name InternetConnectionSharing  -Value 'Disabled'
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 0)
                {
                    $connectionConfig | Add-Member -Type NoteProperty -Name InternetConnectionSharing -Value 'Enabled (public)'
                }
                if ($connectionConfig.SharingEnabled -eq 'True' -and $connectionConfig.SharingConnectionType -eq 1)
                {
                    $connectionConfig | Add-Member -Type NoteProperty -Name InternetConnectionSharing -Value 'Enabled (private)'
                }
            
                Write-Output $connectionConfig | Select-Object -Property NetworkConnectionName,InternetConnectionSharing
            }
        }   
    }
}
