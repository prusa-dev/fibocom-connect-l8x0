function Get-NetworkInterface {
    param(
        [Parameter(Mandatory)]
        [string] $Mac
    )

    $ncm1ifindex = Get-NetAdapter | Where-Object { $_.MacAddress -eq $Mac } | Select-Object -ExpandProperty InterfaceIndex
    if ($ncm1ifindex) {
        return $ncm1ifindex
    }
}

function Initialize-Network {
    param(
        [Parameter(Mandatory)]
        [uint32] $InterfaceIndex,
        [Parameter(Mandatory)]
        [string] $IpAddress,
        [Parameter(Mandatory)]
        [string] $IpMask,
        [Parameter(Mandatory)]
        [string] $IpGateway,
        [Parameter(Mandatory)]
        [string] $IpDns1,
        [Parameter(Mandatory)]
        [string] $IpDns2
    )
    ### Setup IPv4 Network

    $ipPrefixLength = ([Convert]::ToString(([ipaddress]$IpMask).Address, 2) -replace 0, $null).Length
    $mac = Get-NetAdapter -ifIndex $InterfaceIndex | Select-Object -ExpandProperty MacAddress

    #### Adapter init
    Get-NetAdapter -ifIndex $InterfaceIndex | Enable-NetAdapter -Confirm:$false | Out-Null
    Get-NetAdapter -ifIndex $InterfaceIndex | Select-Object -Property name | Disable-NetAdapterBinding | Out-Null
    Get-NetAdapter -ifIndex $InterfaceIndex | Select-Object -Property name | Enable-NetAdapterBinding -ComponentID ms_tcpip | Out-Null

    #### Address cleanup
    Start-Sleep -Milliseconds 100
    Get-NetIPAddress -ifIndex $InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Remove-NetIPAddress -Confirm:$false | Out-Null
    Get-NetNeighbor -ifIndex $InterfaceIndex -LinkLayerAddress $Mac -ErrorAction SilentlyContinue | Remove-NetNeighbor -Confirm:$false | Out-Null
    Get-NetRoute -ifIndex $InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Remove-NetRoute -Confirm:$false | Out-Null
    Set-DnsClientServerAddress -InterfaceIndex $InterfaceIndex -ResetServerAddresses -Confirm:$false | Out-Null

    ##### Address assign
    Start-Sleep -Milliseconds 100
    Set-NetIPInterface -ifIndex $InterfaceIndex -Dhcp Disabled
    New-NetIPAddress -ifIndex $InterfaceIndex -AddressFamily IPv4 -IPAddress $IpAddress -PrefixLength $ipPrefixLength -PolicyStore ActiveStore | Out-Null
    New-NetNeighbor -ifIndex $InterfaceIndex -AddressFamily IPv4 -IPAddress $IpAddress -LinkLayerAddress $Mac | Out-Null
    New-NetNeighbor -ifIndex $InterfaceIndex -AddressFamily IPv4 -IPAddress $IpGateway -LinkLayerAddress $Mac | Out-Null

    #### Add route
    Start-Sleep -Milliseconds 100
    New-NetRoute -ifIndex $InterfaceIndex -NextHop $IpGateway -DestinationPrefix "0.0.0.0/0" -RouteMetric 0 -PolicyStore ActiveStore | Out-Null

    #### Add DNS
    Start-Sleep -Milliseconds 100
    Set-DNSClient -InterfaceIndex $InterfaceIndex -RegisterThisConnectionsAddress $false | Out-Null
    Set-DnsClientServerAddress -InterfaceIndex $InterfaceIndex -ServerAddresses @("$($IpDns1)", "$($IpDns2)") | Out-Null
}
