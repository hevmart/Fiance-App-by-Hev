$ErrorActionPreference = "Stop"

$appRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonPath = Join-Path $appRoot ".venv\Scripts\python.exe"
$appPath = Join-Path $appRoot "app.py"
$url = "http://127.0.0.1:5000"
$healthUrl = "http://127.0.0.1:5000/healthz"
$logPath = Join-Path $appRoot "launch-log.txt"
$runtimeOutPath = Join-Path $appRoot "launch-runtime-out.log"
$runtimeErrPath = Join-Path $appRoot "launch-runtime-err.log"
$port = 5000

function Write-LaunchLog {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $Message"
    for ($attempt = 0; $attempt -lt 3; $attempt++) {
        try {
            Add-Content -Path $logPath -Value $line
            return
        }
        catch {
            Start-Sleep -Milliseconds 80
        }
    }
}

function Show-LaunchError {
    param([string]$Message)
    try {
        $shell = New-Object -ComObject WScript.Shell
        $shell.Popup($Message, 0, "H-Queex Hub", 16) | Out-Null
    }
    catch {
    }
}

function Test-LocalPortListening {
    param(
        [string]$HostName,
        [int]$Port,
        [int]$TimeoutMs = 250
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($HostName, $Port, $null, $null)
        $connected = $async.AsyncWaitHandle.WaitOne($TimeoutMs)
        if (-not $connected) {
            return $false
        }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Close()
    }
}

function Wait-ForLocalPort {
    param(
        [string]$HostName,
        [int]$Port,
        [int]$TimeoutMs
    )

    $deadline = (Get-Date).AddMilliseconds($TimeoutMs)
    while ((Get-Date) -lt $deadline) {
        if (Test-LocalPortListening -HostName $HostName -Port $Port -TimeoutMs 300) {
            return $true
        }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

function Test-HttpReady {
    param(
        [string]$TargetUrl,
        [int]$TimeoutSec = 2
    )

    try {
        $response = Invoke-WebRequest -Uri $TargetUrl -UseBasicParsing -TimeoutSec $TimeoutSec
        return ($response.StatusCode -ge 200 -and $response.StatusCode -lt 500)
    }
    catch {
        return $false
    }
}

function Get-AppProcessIds {
    $processes = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            ($_.Name -eq "python.exe" -or $_.Name -eq "pythonw.exe") -and
            $_.CommandLine -like ("*" + $appPath + "*")
        }
    return @($processes | Select-Object -ExpandProperty ProcessId -Unique)
}

function Stop-AppProcessesExcept {
    param([int]$KeepProcessId = 0)

    $processIds = Get-AppProcessIds
    foreach ($processId in $processIds) {
        if ($KeepProcessId -gt 0 -and $processId -eq $KeepProcessId) {
            continue
        }
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }
}

function Get-PortOwnerProcessId {
    param([int]$Port)

    $listener = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -eq $listener) {
        return 0
    }
    return [int]$listener.OwningProcess
}

try {
    Write-LaunchLog "Launch requested"

    if (Test-HttpReady -TargetUrl $healthUrl -TimeoutSec 2) {
        Write-LaunchLog "Existing app HTTP endpoint is responsive"
        Start-Process -FilePath "explorer.exe" -ArgumentList $url | Out-Null
        Write-LaunchLog "Browser open command issued (existing healthy app)"
        return
    }

    $initialProcessIds = Get-AppProcessIds
    if ($initialProcessIds.Count -gt 0) {
        Write-LaunchLog "App processes detected ($($initialProcessIds.Count)) without a healthy endpoint; terminating all to reset"
        Stop-AppProcessesExcept
        Start-Sleep -Milliseconds 500
    }

    if (Test-LocalPortListening -HostName "127.0.0.1" -Port $port -TimeoutMs 300) {
        Write-LaunchLog "Process(es) still listening on port $port; terminating owning process(es)"
        $stalePids = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty OwningProcess -Unique
        foreach ($stalePid in $stalePids) {
            Stop-Process -Id $stalePid -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Milliseconds 600
    }

    if (Test-LocalPortListening -HostName "127.0.0.1" -Port $port -TimeoutMs 300) {
        throw "Port $port is still in use after attempting to terminate existing processes."
    }

    if (-not (Test-Path $pythonPath)) {
        throw "Python virtual environment not found at $pythonPath"
    }

    if (-not (Test-Path $appPath)) {
        throw "Flask app entry point not found at $appPath"
    }

    if (Test-Path $runtimeOutPath) {
        Remove-Item -Path $runtimeOutPath -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $runtimeErrPath) {
        Remove-Item -Path $runtimeErrPath -Force -ErrorAction SilentlyContinue
    }

    $oldDebug = $env:HQ_FINANCE_DEBUG
    $env:HQ_FINANCE_DEBUG = "0"
    $proc = Start-Process -FilePath $pythonPath -ArgumentList @('"' + $appPath + '"') -WorkingDirectory $appRoot -WindowStyle Hidden -RedirectStandardOutput $runtimeOutPath -RedirectStandardError $runtimeErrPath -PassThru
    if ($null -ne $oldDebug) {
        $env:HQ_FINANCE_DEBUG = $oldDebug
    }
    else {
        Remove-Item Env:HQ_FINANCE_DEBUG -ErrorAction SilentlyContinue
    }

    Write-LaunchLog "Python start command issued (PID: $($proc.Id))"

    $ready = $false
    $deadline = (Get-Date).AddSeconds(20)
    while ((Get-Date) -lt $deadline) {
        if (Test-HttpReady -TargetUrl $healthUrl -TimeoutSec 2) {
            $ready = $true
            break
        }
        Start-Sleep -Milliseconds 500
    }

    if (-not $ready) {
        $outTail = ""
        $errTail = ""
        if (Test-Path $runtimeOutPath) {
            $outTail = (Get-Content -Path $runtimeOutPath -Tail 20 | Out-String)
        }
        if (Test-Path $runtimeErrPath) {
            $errTail = (Get-Content -Path $runtimeErrPath -Tail 20 | Out-String)
        }
        throw "App did not become HTTP-responsive at $healthUrl within 20s. Stdout tail:`n$outTail`nStderr tail:`n$errTail"
    }

    Write-LaunchLog "App HTTP endpoint is responsive"

    try {
        Start-Process -FilePath "explorer.exe" -ArgumentList $url | Out-Null
        Write-LaunchLog "Browser open command issued (explorer)"
    }
    catch {
        try {
            Start-Process -FilePath "cmd.exe" -ArgumentList ('/c start "" "' + $url + '"') -WindowStyle Hidden | Out-Null
            Write-LaunchLog "Browser open command issued (cmd fallback)"
        }
        catch {
            Start-Process $url | Out-Null
            Write-LaunchLog "Browser open command issued (powershell fallback)"
        }
    }
}
catch {
    $errorMessage = "Launch failed: $($_.Exception.Message)"
    Write-LaunchLog $errorMessage
    Show-LaunchError "$errorMessage`n`nSee launch-log.txt, launch-runtime-out.log, and launch-runtime-err.log in the app folder."
}