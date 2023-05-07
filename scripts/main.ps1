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

$modem = [System.IO.Ports.SerialPort] $null
$modemEventSourceIdentifier = "SerialPort.DataReceived"

function Write-Error2 {
    param (
        [Parameter(Position = 0)]
        [string]$Message
    )
    Write-Host -BackgroundColor $Host.PrivateData.ErrorBackgroundColor -ForegroundColor $Host.PrivateData.ErrorForegroundColor $Message
}

function Start-WaitMessage {
    param(
        [Parameter(Mandatory)]
        [string] $Message,
        [Parameter(Mandatory)]
        [scriptblock] $Action
    )

    try {
        $cursorSize = $Host.UI.RawUI.CursorSize; $Host.UI.RawUI.CursorSize = 0
        $messageLine = $Host.UI.RawUI.CursorPosition

        Write-Host

        $job = Start-ThreadJob -StreamingHost $Host -ScriptBlock {
            $messageLine = $using:messageLine
            $counter = 0
            while ($true) {
                $frame = $using:Message + ''.PadRight($counter % 4, '.')

                $currentLine = $Host.UI.RawUI.CursorPosition
                $Host.UI.RawUI.CursorPosition = $messageLine

                Write-Host "$frame".PadRight($Host.UI.RawUI.BufferSize.Width, ' ')

                $Host.UI.RawUI.CursorPosition = $currentLine

                $counter += 1
                Start-Sleep -Milliseconds 300
            }
        }

        & $Action

        $currentLine = $Host.UI.RawUI.CursorPosition
        $Host.UI.RawUI.CursorPosition = $messageLine
        Write-Host "$Message DONE!"
        $Host.UI.RawUI.CursorPosition = $currentLine
    }
    catch {
        $currentLine = $Host.UI.RawUI.CursorPosition
        $Host.UI.RawUI.CursorPosition = $messageLine
        Write-Host -BackgroundColor $Host.PrivateData.ErrorBackgroundColor -ForegroundColor $Host.PrivateData.ErrorForegroundColor "$Message ERROR!"
        $Host.UI.RawUI.CursorPosition = $currentLine
        throw
    }
    finally {
        $job | Stop-Job -PassThru -ErrorAction SilentlyContinue | Remove-Job
        $Host.UI.RawUI.CursorSize = $cursorSize
    }
}

function Awk {
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [string] $InputValue,
        [Parameter()]
        [regex] $Split = '\s',
        [Parameter(Mandatory)]
        [regex] $Filter,
        [Parameter(Mandatory)]
        [scriptblock] $Action
    )

    $InputValue -split "`r|`n" | Where-Object { $_ } | Select-String -Pattern $Filter | ForEach-Object {
        $actionArgs = $_ -split $Split
        Invoke-Command -ScriptBlock $Action -ArgumentList $actionArgs
    }
}

function Send-ATCommand {
    param (
        [string] $atCommand
    )

    $response = ''
    $modem.WriteLine($atCommand)

    while ($true) {
        $e = Wait-Event -SourceIdentifier $modemEventSourceIdentifier -Timeout ($modem.ReadTimeout / 1000)
        if (-Not $e) {
            return $null
        }
        Remove-Event -EventIdentifier $e.EventIdentifier
        $response += $modem.ReadExisting()
        if ($response -match "`r`n(OK|ERROR)") {
            break;
        }
    }

    $response
}

Write-Host "Check NCM CDC availablity and status..."
$ncm1ifindex = Get-NetAdapter | Where-Object { $_.MacAddress -eq $MAC } | Select-Object -ExpandProperty InterfaceIndex
if (-Not $ncm1ifindex) {
    Write-Error2 "No NCM CDC with MAC '$MAC' found. Exiting."
    exit 1
}
Write-Host "NCM CDC available, ifindex = $ncm1ifindex"

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

$modem = new-Object System.IO.Ports.SerialPort $modemport, 115200, None, 8, one
$modem.ReadBufferSize = 8192
$modem.ReadTimeout = 1000
$modem.WriteBufferSize = 8192
$modem.WriteTimeout = 1000
$modem.DtrEnable = $true
$modem.NewLine = "`r"

try {
    Register-ObjectEvent -SourceIdentifier $modemEventSourceIdentifier -InputObject $modem -EventName "DataReceived"

    $modem.Open();
    $modem.DiscardInBuffer()
    $modem.DiscardOutBuffer()


    Send-ATCommand "ATE0" | Out-Null

    ### Get modem information
    $response = Send-ATCommand "AT+CGMI?; +FMM?; +GTPKGVER?; +CFSN?; +CGSN?"

    $manufacturer = $response | Awk -Split ':|,' -Filter '\+CGMI:' -Action { $args[1] -replace '"|^\s', '' }
    $model = $response | Awk -Split ':|,' -Filter '\+FMM:' -Action { $args[1] -replace '"|^\s', '' }

    $firmwareVer = $response | Awk -Filter '\+GTPKGVER:' -Action { $args[1] -replace '"', '' }
    $serialNumber = $response | Awk -Filter '\+CFSN:' -Action { $args[1] -replace '"', '' }

    $imei = $response | Awk -Filter '\+CGSN:' -Action { $args[1] -replace '"', '' }

    Write-Host "Manufacturer: $manufacturer"
    Write-Host "Model: $model"
    Write-Host "Firmware: $firmwareVer"
    Write-Host "Serial: $serialNumber"
    Write-Host "IMEI: $imei"

    ### TODO: add sim pin and status check

    ### Get SIM information
    $response = Send-ATCommand "AT+CGSN?; +CIMI?; +CCID?"

    $imsi = $response | Awk -Filter '\+CIMI:' -Action { $args[1] -replace '"', '' }
    $ccid = $response | Awk -Filter '\+CCID:' -Action { $args[1] -replace '"', '' }

    Write-Host "IMSI: $imsi"
    Write-Host "ICCID: $ccid"


    if (-not $OnlyMonitor) {
        ### Connect
        Write-Host
        Start-WaitMessage -Message "Initialize connection" -Action {
            $response = Send-ATCommand "AT+CFUN=1"
            $response = Send-ATCommand "AT+CMEE=1"
            $response = Send-ATCommand "AT+CGPIAF=1,0,0,0"
            $response = Send-ATCommand "AT+CREG=0"
            $response = Send-ATCommand "AT+CEREG=0"
            $response = Send-ATCommand "AT+CGATT=0"
            $response = Send-ATCommand "AT+COPS=2"
            $response = Send-ATCommand "AT+CGDCONT=0,`"IP`""
            $response = Send-ATCommand "AT+CGDCONT=0"
            $response = Send-ATCommand "AT+XACT=2,,,0"
            $response = Send-ATCommand "AT+CGDCONT=1,`"IP`",`"$APN`""
            $response = Send-ATCommand "AT+XGAUTH=1,0,`"$APN_NAME`",`"$APN_PASS`""
            $response = Send-ATCommand "AT+XDATACHANNEL=1,1,`"/USBCDC/0`",`"/USBHS/NCM/0`",2,1"
            $response = Send-ATCommand "AT+XDNS=1,1"
            $response = Send-ATCommand "AT+CGACT=1,1"
            $response = Send-ATCommand "AT+COPS=0,0"
            $response = Send-ATCommand "AT+CGATT=1"
            $response = Send-ATCommand "AT+CGDATA=M-RAW_IP,1"
        }
    }

    Start-WaitMessage -Message "Establish connection" -Action {
        while ($true) {
            $response = Send-ATCommand "AT+CGATT?; +CSQ?"

            $cgatt = $response | Awk -Split ':|,' -Filter '\+CGATT:' -Action { 1 * $args[1] }
            $csq = $response | Awk -Split ':|,' -Filter '\+CSQ:' -Action { 1 * $args[1] }

            if ($cgatt -eq 1 -and $csq -ne 99) {
                break
            }

            Start-Sleep -Seconds 2
        }
    }

    Write-Host "=== Connection information ==="

    $ip_addr = "--"
    $ip_mask = "--"
    $ip_prefix_length = "--"
    $ip_gw = "--"
    $ip_dns1 = "--"
    $ip_dns2 = "--"

    $response = Send-ATCommand "AT+CGCONTRDP=1"

    if ($response -match "`r`nOK") {

        $ip_addr = $response | Awk -Split ':|,' -Filter '\+CGCONTRDP:' -Action { $args[4] -replace '"', '' }
        $m = [regex]::Match($ip_addr, '(?<ip>(?:\d{1,3}\.){3}\d{1,3})\.(?<mask>(?:\d{1,3}\.){3}\d{1,3})')
        if (-Not $m.Success) {
            Write-Error2 "Could not get ip address from '$ip_addr'"
            exit 1
        }
        $ip_addr = $m.Groups['ip'].Value
        $ip_mask = $m.Groups['mask'].Value
        $ip_prefix_length = ([Convert]::ToString(([ipaddress]$ip_mask).Address, 2) -replace 0, $null).Length
        $ip_gw = $response | Awk -Split ':|,' -Filter '\+CGCONTRDP:' -Action { $args[5] -replace '"', '' }
        $ip_dns1 = $response | Awk -Split ':|,' -Filter '\+CGCONTRDP:' -Action { $args[6] -replace '"', '' }
        $ip_dns2 = $response | Awk -Split ':|,' -Filter '\+CGCONTRDP:' -Action { $args[7] -replace '"', '' }
    }

    Write-Host "IP: $ip_addr"
    Write-Host "MASK: $ip_mask"
    Write-Host "GW: $ip_gw"
    Write-Host "DNS1: $ip_dns1"
    Write-Host "DNS2: $ip_dns2"

    if (-Not $OnlyMonitor) {
        ### Setup IPv4 Network
        Write-Host "Setup network"

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


    ### Monitoring
    Write-Host
    Write-Host "=== Status ==="
    $cursorSize = $Host.UI.RawUI.CursorSize; $Host.UI.RawUI.CursorSize = 0
    try {
        $currentLine = $Host.UI.RawUI.CursorPosition
        while ($true) {
            $response = ''

            $response += Send-ATCommand "AT+MTSM=1"
            $response += Send-ATCommand "AT+COPS?"
            $response += Send-ATCommand "AT+CSQ?"
            #$response += Send-ATCommand "AT+XCESQ?; +RSRP?; +RSRQ?"
            $response += Send-ATCommand "AT+XLEC?; +XCCINFO?; +XMCI=1"

            $tech = $response | Awk -Split ':|,' -Filter '\+COPS:' -Action { 1 * $args[4] }
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

            $oper = $response | Awk -Split ':|,' -Filter '\+COPS:' -Action { $args[3] -replace '"', '' }
            $temp = $response | Awk -Split ':|,' -Filter '\+MTSM:' -Action { 1 * $args[1] }

            $csq = $response | Awk -Split ':|,' -Filter '\+CSQ:' -Action { 1 * $args[1] }
            $csq_perc = 0
            if ($csq -ge 0 -and $csq -le 31) {
                $csq_perc = $csq * 100 / 31
            }
            $cqs_rssi = 2 * $csq - 113

            $rsrp = $response | Awk -Split ':|,' -Filter '\+XMCI: 4' -Action { (1 * $args[10]) - 141 }
            $rsrq = $response | Awk -Split ':|,' -Filter '\+XMCI: 4' -Action { (1 * $args[11]) / 2 - 20 }
            $sinr = $response | Awk -Split ':|,' -Filter '\+XMCI: 4' -Action { (1 * $args[12]) / 2 }

            $bw = $response | Awk -Split ':|,' -Filter '\+XLEC:' -Action { 1 * $args[3] }

            $bw_freq = switch ($bw) {
                0 { 1.4 }
                1 { 3 }
                2 { 5 }
                3 { 10 }
                4 { 15 }
                5 { 20 }
                default { 0 }
            }

            $np = switch ($bw) {
                0 { 6 }
                1 { 15 }
                2 { 25 }
                3 { 50 }
                4 { 75 }
                5 { 100 }
                default { 0 }
            }

            $rssi = '--'
            if ($np -ne 0) {
                $rssi = $rsrp + 10 * [Math]::Log(12 * $np) / [Math]::Log(10)
            }

            ### Display
            $Host.UI.RawUI.CursorPosition = $currentLine

            $lineWidth = 50
            $titleWidth = 10

            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0} $([char]0xB0)C" -f "Temp:", $temp))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1} ({2})" -f "Operator:", $oper, $mode))

            $csq_bar_size = 5
            $csq_bar_fill = [Math]::Round($csq_perc / (100 / $csq_bar_size))
            $csq_bar_empty = $csq_bar_size - $csq_bar_fill;
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,3:f0}% [{2}{3}]" -f "Signal:", $csq_perc, ("$([char]0x2588)" * $csq_bar_fill), ("$([char]0x2591)" * $csq_bar_empty)))


            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0} dBm" -f "RSSI:", $rssi))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0} dB" -f "SINR:", $sinr))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0} dBm" -f "RSRP:", $rsrp))
            Write-Host ("{0,-$lineWidth}" -f ("{0,-$titleWidth} {1,4:f0} dB" -f "RSRQ:", $rsrq))

            Start-Sleep -Seconds 2
        }
    }
    finally {
        $Host.UI.RawUI.CursorSize = $cursorSize
    }
}
finally {
    Unregister-Event -SourceIdentifier $modemEventSourceIdentifier -Force -ErrorAction SilentlyContinue
    if ($modem.IsOpen) {
        $modem.Close()
    }
}
