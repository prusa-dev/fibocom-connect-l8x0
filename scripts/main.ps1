#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch] $OnlyMonitor = $false
)

$ErrorActionPreference = 'Stop'

# NCM intrface MAC address
$MAC = "00-00-11-12-13-14"
# COM port display name search string. Could be '*COM7*' if acm2 does not exists on your machine
$COM_NAME = "*acm2*"

$APN = "internet"
$APN_USER = ""
$APN_PASS = ""

Clear-Host

### Ublock files
Get-ChildItem -Recurse -Path .\ -Include *.ps1, *.psm1, *.psd1, *.dll | Unblock-File

### Import modules
if (-Not(Get-Command | Where-Object { $_.Name -like 'Start-ThreadJob' })) {
    Import-Module -Global ./modules/ThreadJob/ThreadJob.psd1
}
Import-Module ./modules/common.psm1
Import-Module ./modules/serial-port.psm1
Import-Module ./modules/converters.psm1
Import-Module ./modules/network.psm1

try {
    while ($true) {
        Clear-Host

        $modem_port = Wait-Action -Message 'Find modem control port' -Action {
            while ($true) {
                $port = Get-SerialPort -FriendlyName $COM_NAME
                if ($port) {
                    return $port
                }
                Start-Sleep -Seconds 5 | Out-Null
            }
        }

        if ($modem) {
            $modem.Dispose()
        }

        $modem = New-SerialPort -Name $modem_port

        Open-SerialPort -Port $modem

        Send-ATCommand -Port $modem -Command "ATE1" | Out-Null
        Send-ATCommand -Port $modem -Command "AT+CMEE=2" | Out-Null

        ### Get modem information
        Write-Host
        Write-Host "=== Modem information ==="

        $response = Send-ATCommand -Port $modem -Command "AT+CGMI?; +FMM?; +GTPKGVER?; +CFSN?; +CGSN?"

        $manufacturer = $response | Awk -Split '[:,]' -Filter '\+CGMI:' -Action { $args[1] -replace '"|^\s', '' }
        $model = $response | Awk -Split '[:,]' -Filter '\+FMM:' -Action { $args[1] -replace '"|^\s', '' }
        $firmwareVer = $response | Awk -Filter '\+GTPKGVER:' -Action { $args[1] -replace '"', '' }
        $serialNumber = $response | Awk -Filter '\+CFSN:' -Action { $args[1] -replace '"', '' }
        $imei = $response | Awk -Filter '\+CGSN:' -Action { $args[1] -replace '"', '' }

        Write-Host "Manufacturer: $manufacturer"
        Write-Host "Model: $model"
        Write-Host "Firmware: $firmwareVer"
        Write-Host "Serial: $serialNumber"
        Write-Host "IMEI: $imei"

        ### Check SIM Card
        $response = Send-ATCommand -Port $modem -Command "AT+CPIN?"
        if (-Not($response -match '\+CPIN: READY')) {
            Write-Error2 "Check SIM card."
            Write-Error2 ($response -join "`r`n")
            exit 1
        }

        ### Get SIM information
        $response = Send-ATCommand -Port $modem -Command "AT+CIMI?; +CCID?"

        $imsi = $response | Awk -Filter '\+CIMI:' -Action { $args[1] -replace '"', '' }
        $ccid = $response | Awk -Filter '\+CCID:' -Action { $args[1] -replace '"', '' }

        Write-Host "IMSI: $imsi"
        Write-Host "ICCID: $ccid"

        if (-Not $OnlyMonitor) {
            ### Connect
            Write-Host
            Wait-Action -Message "Initialize connection" -Action {
                Start-Sleep -Seconds 5 | Out-Null
                $response = ''
                $response += Send-ATCommand -Port $modem -Command "AT+CFUN=1"
                $response += Send-ATCommand -Port $modem -Command "AT+CGPIAF=1,0,0,0"
                $response += Send-ATCommand -Port $modem -Command "AT+CREG=0"
                $response += Send-ATCommand -Port $modem -Command "AT+CEREG=0"
                $response += Send-ATCommand -Port $modem -Command "AT+CGATT=0"
                $response += Send-ATCommand -Port $modem -Command "AT+COPS=2"
                $response += Send-ATCommand -Port $modem -Command "AT+XCESQRC=1"
                $response += Send-ATCommand -Port $modem -Command "AT+XACT=2,,,0"
                $response += Send-ATCommand -Port $modem -Command "AT+CGDCONT=0,`"IP`""
                $response += Send-ATCommand -Port $modem -Command "AT+CGDCONT=0"
                $response += Send-ATCommand -Port $modem -Command "AT+CGDCONT=1,`"IP`",`"$APN`""
                $response += Send-ATCommand -Port $modem -Command "AT+XGAUTH=1,0,`"$APN_USER`",`"$APN_PASS`""
                $response += Send-ATCommand -Port $modem -Command "AT+XDATACHANNEL=1,1,`"/USBCDC/0`",`"/USBHS/NCM/0`",2,1"
                $response += Send-ATCommand -Port $modem -Command "AT+XDNS=1,1"
                $response += Send-ATCommand -Port $modem -Command "AT+CGACT=1,1"
                $response += Send-ATCommand -Port $modem -Command "AT+COPS=0,0"
                $response += Send-ATCommand -Port $modem -Command "AT+CGATT=1"
                $response += Send-ATCommand -Port $modem -Command "AT+CGDATA=M-RAW_IP,1"
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
        $ip_gw = "--"
        $ip_dns1 = "--"
        $ip_dns2 = "--"

        $response = Send-ATCommand -Port $modem -Command "AT+CGCONTRDP=1"

        if (Test-AtCommandSuccess $response) {
            $ip_addr = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[4] -replace '"', '' }
            $m = [regex]::Match($ip_addr, '(?<ip>(?:\d{1,3}\.){3}\d{1,3})\.(?<mask>(?:\d{1,3}\.){3}\d{1,3})')
            if (-Not $m.Success) {
                Write-Error2 "Could not get ip address from '$ip_addr'"
                exit 1
            }
            $ip_addr = $m.Groups['ip'].Value
            $ip_mask = $m.Groups['mask'].Value
            $ip_gw = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[5] -replace '"', '' }
            $ip_dns1 = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[6] -replace '"', '' }
            $ip_dns2 = $response | Awk -Split '[:,]' -Filter '\+CGCONTRDP:' -Action { $args[7] -replace '"', '' }
        }
        elseif (-Not $OnlyMonitor) {
            Write-Error2 "Could not get ip address."
            Write-Error2 $response
            exit 1
        }

        Write-Host "IP: $ip_addr"
        Write-Host "MASK: $ip_mask"
        Write-Host "GW: $ip_gw"
        Write-Host "DNS1: $ip_dns1"
        Write-Host "DNS2: $ip_dns2"

        if (-Not $OnlyMonitor) {
            Wait-Action -ErrorAction SilentlyContinue -Message "Setup network" -Action {
                $ncm1ifindex = Get-NetworkInterface -Mac $MAC
                if (-Not $ncm1ifindex) {
                    Write-Error2 "Could not find interface with mac '$MAC'"
                    exit 1
                }
                Initialize-Network -InterfaceIndex $ncm1ifindex -IpAddress $ip_addr -IpMask $ip_mask -IpGateway $ip_gw -IpDns1 $ip_dns1 -IpDns2 $ip_dns2
            }
        }


        ## Watchdog

        $watchdogEventSource = "WatchdogEvent"
        Start-SerialPortMonitoring -WatchdogSourceIdentifier $watchdogEventSource -FriendlyName $COM_NAME

        ### Monitoring
        Write-Host
        Write-Host "=== Status ==="
        $cursorSize = $Host.UI.RawUI.CursorSize; $Host.UI.RawUI.CursorSize = 0
        try {
            $currentLine = $Host.UI.RawUI.CursorPosition

            while ($true) {
                if ((Get-Event -SourceIdentifier $watchdogEventSource -ErrorAction SilentlyContinue)) {
                    break
                }

                $response = ''

                $response += Send-ATCommand -Port $modem -Command "AT+MTSM=1"
                $response += Send-ATCommand -Port $modem -Command "AT+COPS?"
                $response += Send-ATCommand -Port $modem -Command "AT+CSQ?"
                $response += Send-ATCommand -Port $modem -Command "AT+XLEC?; +XCCINFO?; +XMCI=1"

                if ([string]::IsNullOrEmpty($response)) {
                    continue
                }

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

        Stop-SerialPortMonitoring
        Get-Event -SourceIdentifier $watchdogEventSource -ErrorAction SilentlyContinue | Remove-Event
        Close-SerialPort -Port $modem
    }
}
finally {
    Stop-SerialPortMonitoring
    Close-SerialPort -Port $modem
}
