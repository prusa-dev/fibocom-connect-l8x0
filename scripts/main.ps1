#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch] $OnlyMonitor = $false
)

$ErrorActionPreference = 'Stop'

$MAC = "00-00-11-12-13-14"
$NAM = "*acm2*"
$APN = "internet.mts.ru"
$APN_NAME = ""
$APN_PASS = ""

Clear-Host

### Import modules
if (-Not(Get-Command | Where-Object { $_.Name -like 'Start-ThreadJob' })) {
    Import-Module ./ThreadJob/ThreadJob.psd1
}
Import-Module ./common.psm1
Import-Module ./serial-port.psm1
Import-Module ./converters.psm1

### Find modem
Write-Host "Find modem control port..."
$acm2name = Get-PnpDevice -class ports -FriendlyName $NAM -ErrorAction SilentlyContinue | Where-Object { $_.Status -eq "ok" } | Select-Object -ExpandProperty Name
if (-Not $acm2name) {
    Write-Error2 "No modem control port found. Exiting."
    exit 1
}
$modemport_match = [regex]::Match($acm2name, '(COM\d{1,3})')
if (-Not $modemport_match.Success) {
    Write-Error2 "No modem control port found. Exiting."
    exit 1
}
$modemport = $modemport_match.Groups[1].Value
Write-Host "Found modem control port '$acm2name', '$modemport'"

Write-Host "Check NCM CDC availablity and status..."
$ncm1ifindex = Get-NetAdapter | Where-Object { $_.MacAddress -eq $MAC } | Select-Object -ExpandProperty InterfaceIndex
if (-Not $ncm1ifindex) {
    Write-Error2 "No NCM CDC with MAC '$MAC' found. Exiting."
    exit 1
}
Write-Host "NCM CDC available, ifindex = $ncm1ifindex"


$modem = New-SerialPort -Name $modemport

try {

    Open-SerialPort -Port $modem

    Send-ATCommand -Port $modem -Command "ATE0" | Out-Null

    ### Get modem information
    $response = Send-ATCommand -Port $modem -Command "AT+CGMI?; +FMM?; +GTPKGVER?; +CFSN?; +CGSN?"

    $manufacturer = $response | Awk -Split '[:,]' -Filter '\+CGMI:' -Action { $args[1] -replace '"|^\s', '' }
    $model = $response | Awk -Split '[:,]' -Filter '\+FMM:' -Action { $args[1] -replace '"|^\s', '' }

    $firmwareVer = $response | Awk -Filter '\+GTPKGVER:' -Action { $args[1] -replace '"', '' }
    $serialNumber = $response | Awk -Filter '\+CFSN:' -Action { $args[1] -replace '"', '' }

    $imei = $response | Awk -Filter '\+CGSN:' -Action { $args[1] -replace '"', '' }

    Write-Host
    Write-Host "=== Modem information ==="
    Write-Host "Manufacturer: $manufacturer"
    Write-Host "Model: $model"
    Write-Host "Firmware: $firmwareVer"
    Write-Host "Serial: $serialNumber"
    Write-Host "IMEI: $imei"

    ### TODO: add sim pin and status check

    ### Get SIM information
    $response = Send-ATCommand -Port $modem -Command "AT+CGSN?; +CIMI?; +CCID?"

    $imsi = $response | Awk -Filter '\+CIMI:' -Action { $args[1] -replace '"', '' }
    $ccid = $response | Awk -Filter '\+CCID:' -Action { $args[1] -replace '"', '' }

    Write-Host "IMSI: $imsi"
    Write-Host "ICCID: $ccid"


    if (-not $OnlyMonitor) {
        ### Connect
        Write-Host
        Wait-Action -Message "Initialize connection" -Action {
            $response = Send-ATCommand -Port $modem -Command "AT+CFUN=1"
            $response = Send-ATCommand -Port $modem -Command "AT+CMEE=1"
            $response = Send-ATCommand -Port $modem -Command "AT+CGPIAF=1,0,0,0"
            $response = Send-ATCommand -Port $modem -Command "AT+CREG=0"
            $response = Send-ATCommand -Port $modem -Command "AT+CEREG=0"
            $response = Send-ATCommand -Port $modem -Command "AT+CGATT=0"
            $response = Send-ATCommand -Port $modem -Command "AT+COPS=2"
            $response = Send-ATCommand -Port $modem -Command "AT+XCESQRC=1"
            $response = Send-ATCommand -Port $modem -Command "AT+XACT=2,,,0"
            $response = Send-ATCommand -Port $modem -Command "AT+CGDCONT=0,`"IP`""
            $response = Send-ATCommand -Port $modem -Command "AT+CGDCONT=0"
            $response = Send-ATCommand -Port $modem -Command "AT+CGDCONT=1,`"IP`",`"$APN`""
            $response = Send-ATCommand -Port $modem -Command "AT+XGAUTH=1,0,`"$APN_NAME`",`"$APN_PASS`""
            $response = Send-ATCommand -Port $modem -Command "AT+XDATACHANNEL=1,1,`"/USBCDC/0`",`"/USBHS/NCM/0`",2,1"
            $response = Send-ATCommand -Port $modem -Command "AT+XDNS=1,1"
            $response = Send-ATCommand -Port $modem -Command "AT+CGACT=1,1"
            $response = Send-ATCommand -Port $modem -Command "AT+COPS=0,0"
            $response = Send-ATCommand -Port $modem -Command "AT+CGATT=1"
            $response = Send-ATCommand -Port $modem -Command "AT+CGDATA=M-RAW_IP,1"
        }

        Wait-Action -Message "Establish connection" -Action {
            while ($true) {
                $response = Send-ATCommand -Port $modem -Command "AT+CGATT?; +CSQ?"

                $cgatt = $response | Awk -Split '[:,]' -Filter '\+CGATT:' -Action { [int]$args[1] }
                $csq = $response | Awk -Split '[:,]' -Filter '\+CSQ:' -Action { [int]$args[1] }

                if ($cgatt -eq 1 -and $csq -ne 99) {
                    break
                }

                Start-Sleep -Seconds 2
            }
        }
    }

    Write-Host
    Write-Host "=== Connection information ==="

    $ip_addr = "--"
    $ip_mask = "--"
    $ip_prefix_length = "--"
    $ip_gw = "--"
    $ip_dns1 = "--"
    $ip_dns2 = "--"

    $response = Send-ATCommand -Port $modem -Command "AT+CGCONTRDP=1"

    if ($response -match "`r`nOK") {
        $ip_addr = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[4] -replace '"', '' }
        $m = [regex]::Match($ip_addr, '(?<ip>(?:\d{1,3}\.){3}\d{1,3})\.(?<mask>(?:\d{1,3}\.){3}\d{1,3})')
        if (-Not $m.Success) {
            Write-Error2 "Could not get ip address from '$ip_addr'"
            exit 1
        }
        $ip_addr = $m.Groups['ip'].Value
        $ip_mask = $m.Groups['mask'].Value
        $ip_prefix_length = ([Convert]::ToString(([ipaddress]$ip_mask).Address, 2) -replace 0, $null).Length
        $ip_gw = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[5] -replace '"', '' }
        $ip_dns1 = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[6] -replace '"', '' }
        $ip_dns2 = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[7] -replace '"', '' }
    }

    Write-Host "IP: $ip_addr"
    Write-Host "MASK: $ip_mask"
    Write-Host "GW: $ip_gw"
    Write-Host "DNS1: $ip_dns1"
    Write-Host "DNS2: $ip_dns2"

    if (-Not $OnlyMonitor) {
        Wait-Action -Message "Setup network" -Action {
            ### Setup IPv4 Network

            #### Adapter init
            Get-NetAdapter -ifIndex $ncm1ifindex | Enable-NetAdapter -Confirm:$false | Out-Null
            Get-NetAdapter -ifIndex $ncm1ifindex | Select-Object -Property name | Disable-NetAdapterBinding | Out-Null
            Get-NetAdapter -ifIndex $ncm1ifindex | Select-Object -Property name | Enable-NetAdapterBinding -ComponentID ms_tcpip | Out-Null

            #### Address cleanup
            Start-Sleep -Milliseconds 100
            Get-NetIPAddress -ifIndex $ncm1ifindex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Remove-NetIPAddress -Confirm:$false | Out-Null
            Get-NetNeighbor -ifIndex $ncm1ifindex -LinkLayerAddress $MAC -ErrorAction SilentlyContinue | Remove-NetNeighbor -Confirm:$false | Out-Null
            Get-NetRoute -ifIndex $ncm1ifindex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Remove-NetRoute -Confirm:$false | Out-Null
            Set-DnsClientServerAddress -InterfaceIndex $ncm1ifindex -ResetServerAddresses -Confirm:$false | Out-Null

            ##### Address assign
            Start-Sleep -Milliseconds 100
            Set-NetIPInterface -ifIndex $ncm1ifindex -Dhcp Disabled
            New-NetIPAddress -ifIndex $ncm1ifindex -AddressFamily IPv4 -IPAddress $ip_addr -PrefixLength $ip_prefix_length -PolicyStore ActiveStore | Out-Null
            New-NetNeighbor -ifIndex $ncm1ifindex -AddressFamily IPv4 -IPAddress $ip_addr -LinkLayerAddress $MAC | Out-Null
            New-NetNeighbor -ifIndex $ncm1ifindex -AddressFamily IPv4 -IPAddress $ip_gw -LinkLayerAddress $MAC | Out-Null

            #### Add route
            Start-Sleep -Milliseconds 100
            New-NetRoute -ifIndex $ncm1ifindex -NextHop $ip_gw -DestinationPrefix "0.0.0.0/0" -RouteMetric 0 -PolicyStore ActiveStore | Out-Null

            #### Add DNS
            Start-Sleep -Milliseconds 100
            Set-DNSClient -InterfaceIndex $ncm1ifindex -RegisterThisConnectionsAddress $false | Out-Null
            Set-DnsClientServerAddress -InterfaceIndex $ncm1ifindex -ServerAddresses @("$($ip_dns1)", "$($ip_dns2)") | Out-Null
        }
    }


    ### Monitoring
    Write-Host
    Write-Host "=== Status ==="
    $cursorSize = $Host.UI.RawUI.CursorSize; $Host.UI.RawUI.CursorSize = 0
    try {
        $currentLine = $Host.UI.RawUI.CursorPosition

        while ($true) {
            $response = ''

            $response += Send-ATCommand -Port $modem -Command "AT+MTSM=1"
            $response += Send-ATCommand -Port $modem -Command "AT+COPS?"
            $response += Send-ATCommand -Port $modem -Command "AT+CSQ?"
            $response += Send-ATCommand -Port $modem -Command "AT+XLEC?; +XCCINFO?; +XMCI=1"

            $tech = $response | Awk -Split '[:,]' -Filter '\+COPS:' -Action { [int]$args[4] }
            $mode = '--'
            switch ($tech) {
                0 { $mode = 'EDGE' }
                2 { $mode = 'UMTS' }
                3 { $mode = 'LTE' }
                4 { $mode = 'HSDPA' }
                5 { $mode = 'HSUPA' }
                6 { $mode = 'HSPA' }
                7 { $mode = 'LTE' }
            }

            $oper = $response | Awk -Split '[:,]' -Filter '\+COPS:' -Action { $args[3] -replace '"', '' }
            $temp = $response | Awk -Split '[:,]' -Filter '\+MTSM:' -Action { [int]$args[1] }

            $csq = $response | Awk -Split '[:,]' -Filter '\+CSQ:' -Action { [int]$args[1] }
            $csq_perc = 0
            if ($csq -ge 0 -and $csq -le 31) {
                $csq_perc = $csq * 100 / 31
            }
            $cqs_rssi = 2 * $csq - 113

            $rsrp = $response | Awk -Split '[:,]' -Filter '\+XMCI: 4' -Action { ([int]$args[10]) - 141 }
            $rsrq = $response | Awk -Split '[:,]' -Filter '\+XMCI: 4' -Action { ([int]$args[11]) / 2 - 20 }
            $sinr = $response | Awk -Split '[:,]' -Filter '\+XMCI: 4' -Action { ([int]$args[12]) / 2 }

            $bw = $response | Awk -Split '[:,]' -Filter '\+XLEC:' -Action { [int]$args[3] }

            $rssi = Convert-RsrpToRssi $rsrp $bw

            $dluarfnc = $response | Awk -Split '[:,]' -Filter '\+XMCI: 4' -Action { [int]($args[7] -replace '"', '') }

            [int[]]$ci_x = $response | Awk -Split '[:,]' -Filter '\+XMCI:' -Action { [int]($args[5] -replace '"', '') }
            [int[]]$pci_x = $response | Awk -Split '[:,]' -Filter '\+XMCI:' -Action { [int]($args[6] -replace '"', '') }
            [int[]]$dluarfnc_x = $response | Awk -Split '[:,]' -Filter '\+XMCI:' -Action { [int]($args[7] -replace '"', '') }
            [string[]]$band_x = $dluarfnc_x | Get-BandLte
            [int[]]$rsrp_x = $response | Awk -Split '[:,]' -Filter '\+XMCI:' -Action { ([int]$args[10]) - 141 }
            [int[]]$rsrq_x = $response | Awk -Split '[:,]' -Filter '\+XMCI:' -Action { ([int]$args[11]) / 2 - 20 }

            $ca_match = [regex]::Match($response, "\+XLEC: (?:\d+),(?<no_of_cells>\d+),(?:(?<bw>\d+),*)+(?:BAND_LTE_(?:(?<band>\d+),*)+)?")
            if ($ca_match.Success) {
                $ca_number = $ca_match.Groups['no_of_cells'].Value

                [int[]]$ca_bands = $ca_match.Groups['band'].Captures | ForEach-Object { [int]$_.Value } | Where-Object { $_ -ne 0 }
                [int[]]$ca_bws = $ca_match.Groups['bw'].Captures | ForEach-Object { [int]$_.Value }

                $band = ''
                for (($i = 0); $i -lt $ca_number; $i++) {
                    $band += "B{0}@{1}MHz " -f $ca_bands[$i], (Get-BandwidthFrequency $ca_bws[$i])
                }
            }
            else {
                $band = "{0}@{1}MHz" -f (Get-BandLte $dluarfnc), (Get-BandwidthFrequency $bw)
            }

            ### Display
            $Host.UI.RawUI.CursorPosition = $currentLine

            $lineWidth = $Host.UI.RawUI.BufferSize.Width
            $titleWidth = 17

            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0} $([char]0xB0)C" -f "Temp:", $temp))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1} ({2})" -f "Operator:", $oper, $mode))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0}%   {2}" -f "Signal:", $csq_perc, (Get-Bars -Value $csq_perc -Min 0 -Max 100)))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0}dBm {2}" -f "RSSI:", $rssi, (Get-Bars -Value $rssi -Min -110 -Max -25)))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0}dB  {2}" -f "SINR:", $sinr, (Get-Bars -Value $sinr -Min -10 -Max 30)))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0}dBm {2}" -f "RSRP:", $rsrp, (Get-Bars -Value $rsrp -Min -120 -Max -50)))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0}dB  {2}" -f "RSRQ:", $rsrq, (Get-Bars -Value $rsrq -Min -25 -Max -1)))

            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1}" -f "Band:", $band))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1}" -f "EARFCN:", $dluarfnc))

            $currentLine1 = $Host.UI.RawUI.CursorPosition
            for (($i = 0); $i -lt $carriers_count; $i++) {
                Write-Host ("{0,-$lineWidth}" -f ' ')
            }
            $Host.UI.RawUI.CursorPosition = $currentLine1

            $carriers_count = $pci_x.Length
            for (($i = 0); $i -lt $carriers_count; $i++) {
                Write-Host -NoNewline ("{0} " -f "===Carrier $($i + 1):")
                Write-Host -NoNewline ("{0} {1,9} " -f "CI:", $ci_x[$i])
                Write-Host -NoNewline ("{0} {1,5} " -f "PCI:", $pci_x[$i])
                Write-Host -NoNewline ("{0} {1,3} ({2,5}) " -f "Band (EARFCN):", $band_x[$i], $dluarfnc_x[$i])
                Write-Host -NoNewline ("{0} {1,4:f0}dBm {2} " -f "RSRP:", $rsrp_x[$i], (Get-Bars -Value $rsrp_x[$i] -Min -120 -Max -50))
                Write-Host -NoNewline ("{0} {1,4:f0}dB  {2} " -f "RSRQ:", $rsrq_x[$i], (Get-Bars -Value $rsrq_x[$i] -Min -25 -Max -1))
                Write-Host
            }

            Start-Sleep -Seconds 2
        }
    }
    finally {
        $Host.UI.RawUI.CursorSize = $cursorSize
    }
}
finally {
    Close-SerialPort -Port $modem
}
