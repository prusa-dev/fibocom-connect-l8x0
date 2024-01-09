function Get-SerialPort {
    param (
        [Parameter(Mandatory)]
        [string] $FriendlyName
    )
    $pnpDevice = Get-PnpDevice -class ports -FriendlyName $FriendlyName -Status OK -ErrorAction SilentlyContinue

    if ($pnpDevice -and $pnpDevice.Name) {
        $port_match = [regex]::Match($pnpDevice.Name, '(COM\d{1,3})')
        if ($port_match.Success) {
            $port = $port_match.Groups[1].Value
            $containerId = $pnpDevice | Get-PnpDeviceProperty -KeyName DEVPKEY_Device_ContainerId | Select-Object -ExpandProperty Data
            return @($port, $containerId)
        }
    }
    else {
        $null
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

function Test-SerialPort {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.IO.Ports.SerialPort] $Port
    )

    if (-Not($Port.IsOpen)) {
        throw "Modem port is not opened."
    }

    if (-Not($Port.DsrHolding)) {
        throw "Modem port is not available."
    }
}

function Test-AtResponseError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string] $Result
    )

    return ($Result -match "`r`n(ERROR|\+CME ERROR|\+CMS ERROR)")
}

function Test-AtResponseSuccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string] $Result
    )

    return ($Result -match "`r`nOK")
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

    $response = ''
    Write-Verbose "--> `"$Command`""

    Test-SerialPort -Port $Port

    $sourceIdentifier = "$($Port.PortName)_DataReceived"

    $Port.WriteLine($Command)

    $waitEventAttempt = 0
    while ($true) {
        Test-SerialPort -Port $Port
        $e = Wait-Event -SourceIdentifier $sourceIdentifier -Timeout ([Math]::Max(1, [Math]::Ceiling($Port.ReadTimeout / 1000)))
        if (-Not($e)) {
            $waitEventAttempt += 1
            if ($waitEventAttempt -gt 10) {
                throw "Attempts to read the data from modem have been exhausted.`nCOMMAND: '$Command'.`nRESPONSE: '$response'";
            }
            continue;
        }
        Remove-Event -EventIdentifier $e.EventIdentifier
        $response += $Port.ReadExisting()
        Write-Verbose "<-- $response"
        if ((Test-AtResponseSuccess $response) -or (Test-AtResponseError $response)) {
            break;
        }
    }

    $response
}

function Start-SerialPortMonitoring {
    param(
        [Parameter(Mandatory)]
        [string] $WatchdogSourceIdentifier,
        [Parameter(Mandatory)]
        [string] $FriendlyName
    )

    $null = Start-Job -Name "SerialPortMonitoring" -ArgumentList $WatchdogSourceIdentifier, $FriendlyName -InitializationScript $functions -ScriptBlock {
        param (
            [string] $WatchdogSourceIdentifier,
            [string] $FriendlyName
        )

        Import-Module "$($using:PWD)/modules/serial-port.psm1"

        Register-EngineEvent -SourceIdentifier $WatchdogSourceIdentifier -Forward
        Register-WMIEvent -SourceIdentifier "DeviceChangeEvent" -Query "SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 OR EventType = 3 GROUP WITHIN 2"

        try {
            while ($true) {
                try {
                    $e = Wait-Event -SourceIdentifier "DeviceChangeEvent"
                    if (-Not($e)) {
                        Start-Sleep -Seconds 1
                        continue
                    }
                    Remove-Event -EventIdentifier $e.EventIdentifier

                    $foundPort = Get-SerialPort -FriendlyName $FriendlyName
                    if (-Not($foundPort)) {
                        New-Event -SourceIdentifier $WatchdogSourceIdentifier -Sender "SerialPortMonitoring"  -MessageData "Disconnected"
                    }
                }
                catch { }
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
