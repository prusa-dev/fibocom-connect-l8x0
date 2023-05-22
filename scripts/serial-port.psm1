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
        $Port.Close()
    }
}

function Send-ATCommand {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.IO.Ports.SerialPort] $Port,

        [Parameter(Mandatory)]
        [string] $Command
    )

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

