function Get-SerialPort {
    [OutputType([string])]
    param (
        [Parameter(Mandatory)]
        [string] $FriendlyName
    )
    $pnpDevice = Get-PnpDevice -class ports -FriendlyName $FriendlyName -Status OK -ErrorAction SilentlyContinue

    if ($pnpDevice -and $pnpDevice.Name) {
        $port_match = [regex]::Match($pnpDevice.Name, '(COM\d{1,3})')
        if ($port_match.Success) {
            return $port_match.Groups[1].Value
        }
    }
}

function New-SerialPort {
    [CmdletBinding()]
    [OutputType([System.IO.Ports.SerialPort])]
    param(
        [Parameter(Mandatory)]
        [string] $Name
    )

    $port = new-Object System.IO.Ports.SerialPort $Name, 115200, None, 8, one
    $port.ReadBufferSize = 8192
    $port.ReadTimeout = 1000
    $port.WriteBufferSize = 8192
    $port.WriteTimeout = 1000
    $port.DtrEnable = $true
    $port.NewLine = "`r"

    return $port
}


function Open-SerialPort {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.Ports.SerialPort] $Port
    )

    Register-ObjectEvent -InputObject $Port -EventName "DataReceived" -SourceIdentifier "$($Port.PortName)_DataReceived"

    $Port.Open();
    $Port.DiscardInBuffer()
    $Port.DiscardOutBuffer()
}

function Close-SerialPort {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.IO.Ports.SerialPort] $Port
    )

    Unregister-Event -SourceIdentifier "$($Port.PortName)_DataReceived" -Force -ErrorAction SilentlyContinue
    if ($Port.IsOpen) {
        try {
            $Port.Close()
        }
        catch {}
    }
}

function Send-ATCommand {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.IO.Ports.SerialPort] $Port,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Command
    )

    if (-Not($Port.IsOpen)) {
        throw "Can't send command. Modem port is not opened"
    }

    $sourceIdentifier = "$($Port.PortName)_DataReceived"
    $timeout = $Port.ReadTimeout / 1000

    $response = ''
    $Port.WriteLine($Command)

    while ($true) {
        $e = Wait-Event -SourceIdentifier $sourceIdentifier -Timeout $timeout
        if (-Not $e) {
            return $null
        }
        Remove-Event -EventIdentifier $e.EventIdentifier
        $response += $Port.ReadExisting()
        if ($response -match "`r`n(OK|ERROR)") {
            break;
        }
    }

    $response
}

function Start-SerialPortMonitoring {
    param(
        [Parameter(Mandatory)]
        [string] $SourceIdentifier,
        [Parameter(Mandatory)]
        [string] $FriendlyName
    )
    $null = Start-Job -Name "SerialPortMonitoring" -ArgumentList $SourceIdentifier, $FriendlyName -ScriptBlock {
        param (
            [string] $SourceIdentifier,
            [string] $FriendlyName
        )
        Import-Module ./modules/serial-port.psm1

        Register-EngineEvent -SourceIdentifier $SourceIdentifier -Forward
        Register-WMIEvent -SourceIdentifier "DeviceChangeEvent" -Query "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 OR EventType = 3 GROUP WITHIN 2"

        try {
            while ($true) {
                try {
                    $e = Wait-Event -SourceIdentifier "DeviceChangeEvent"
                    if (-Not $e) {
                        Start-Sleep -Seconds 1
                        continue
                    }
                    Remove-Event -EventIdentifier $e.EventIdentifier

                    $portName = Get-SerialPort -FriendlyName $FriendlyName
                    if (-Not ($portName)) {
                        Write-Host "Send event Connected"
                        New-Event -SourceIdentifier $SourceIdentifier -Sender "SerialPortMonitoring"  -MessageData "Disconnected"
                    }
                }
                catch {
                    New-Event -SourceIdentifier $SourceIdentifier -Sender "SerialPortMonitoring"  -MessageData "Error $_"
                }
            }
        }
        finally {
            Unregister-Event -SourceIdentifier "DeviceChangeEvent" -Force -ErrorAction SilentlyContinue
        }
    }
}

function Stop-SerialPortMonitoring {
    Stop-Job -Name "SerialPortMonitoring" -PassThru -ErrorAction SilentlyContinue | Remove-Job | Out-Null
}
