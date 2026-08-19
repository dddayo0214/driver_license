$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$backendRoot = Join-Path $projectRoot "backend"
$frontendRoot = Join-Path $projectRoot "frontend"
$logRoot = Join-Path $projectRoot ".logs"
$runtimeRoot = Join-Path $projectRoot ".runtime"
$processStateFile = Join-Path $runtimeRoot "processes.json"
$backend = $null
$frontend = $null
$exitSubscription = $null
$trackedProcessIds = @()

function Get-ListenerProcessId([int]$Port) {
    try {
        $connection = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction Stop | Select-Object -First 1
        if ($connection) { return [int]$connection.OwningProcess }
    }
    catch { }
    foreach ($line in (& netstat -ano -p TCP)) {
        if ($line -match "^\s*TCP\s+127\.0\.0\.1:$Port\s+\S+\s+LISTENING\s+(\d+)\s*$") {
            return [int]$Matches[1]
        }
    }
    return $null
}

function Stop-ServiceTrees([int[]]$ProcessIds) {
    foreach ($processId in @($ProcessIds | Where-Object { $_ -gt 0 } | Sort-Object -Unique)) {
        & taskkill /PID $processId /T /F 2>$null | Out-Null
    }
}

function Save-ProcessState([int[]]$ProcessIds) {
    $records = foreach ($processId in @($ProcessIds | Where-Object { $_ -gt 0 } | Sort-Object -Unique)) {
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($process) {
            [PSCustomObject]@{ id = $process.Id; started = $process.StartTime.ToUniversalTime().ToString("o") }
        }
    }
    @($records) | ConvertTo-Json | Set-Content -LiteralPath $processStateFile -Encoding UTF8
}

function Stop-SavedProcesses {
    if (-not (Test-Path -LiteralPath $processStateFile)) { return }
    try {
        $records = @(Get-Content -LiteralPath $processStateFile -Raw | ConvertFrom-Json)
        foreach ($record in $records) {
            $process = Get-Process -Id ([int]$record.id) -ErrorAction SilentlyContinue
            if (-not $process) { continue }
            $savedStart = [DateTime]::Parse([string]$record.started).ToUniversalTime()
            if ([Math]::Abs(($process.StartTime.ToUniversalTime() - $savedStart).TotalSeconds) -lt 2) {
                & taskkill /PID $process.Id /T /F 2>$null | Out-Null
            }
        }
    }
    finally { Remove-Item -LiteralPath $processStateFile -Force -ErrorAction SilentlyContinue }
}

New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
Stop-SavedProcesses

# Clean listeners left by an older launcher that did not yet create a PID state file.
$occupied = @()
foreach ($port in @(5173, 8000)) {
    $listener = Get-ListenerProcessId $port
    if ($listener) {
        Stop-ServiceTrees @($listener)
        Start-Sleep -Milliseconds 300
        $remaining = Get-ListenerProcessId $port
        if ($remaining) { $occupied += "port $port (PID $remaining)" }
    }
}
if ($occupied.Count -gt 0) {
    throw "Ports are still occupied: $($occupied -join ', '). Close the previous PowerShell window once, then retry."
}

try {
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    Write-Host "Preparing backend..."
    Push-Location $backendRoot
    try {
        $env:UV_CACHE_DIR = Join-Path $backendRoot ".uv-cache"
        uv sync --locked
    }
    finally { Pop-Location }

    Write-Host "Preparing frontend..."
    Push-Location $frontendRoot
    try { npm install }
    finally { Pop-Location }

    $backend = Start-Process -WindowStyle Hidden -PassThru -WorkingDirectory $backendRoot -FilePath "uv" -ArgumentList "run", "uvicorn", "app.main:app", "--host", "127.0.0.1", "--port", "8000" -RedirectStandardOutput (Join-Path $logRoot "backend.out.log") -RedirectStandardError (Join-Path $logRoot "backend.err.log")
    $frontend = Start-Process -WindowStyle Hidden -PassThru -WorkingDirectory $frontendRoot -FilePath "npm.cmd" -ArgumentList "run", "dev" -RedirectStandardOutput (Join-Path $logRoot "frontend.out.log") -RedirectStandardError (Join-Path $logRoot "frontend.err.log")
    $trackedProcessIds = @($backend.Id, $frontend.Id)
    Save-ProcessState $trackedProcessIds

    Write-Host "Starting local web app..."
    $ready = $false
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        Start-Sleep -Milliseconds 500
        try {
            $frontendResponse = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:5173"
            $backendResponse = Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:8000/api/health"
            if ($frontendResponse.StatusCode -eq 200 -and $backendResponse.StatusCode -eq 200) {
                $ready = $true
                break
            }
        }
        catch { }
    }
    if (-not $ready) {
        $backendError = Get-Content (Join-Path $logRoot "backend.err.log") -Raw -ErrorAction SilentlyContinue
        $frontendError = Get-Content (Join-Path $logRoot "frontend.err.log") -Raw -ErrorAction SilentlyContinue
        throw "Frontend or backend failed to start.`nBackend: $backendError`nFrontend: $frontendError"
    }

    $trackedProcessIds += Get-ListenerProcessId 5173
    $trackedProcessIds += Get-ListenerProcessId 8000
    $trackedProcessIds = @($trackedProcessIds | Where-Object { $_ -gt 0 } | Sort-Object -Unique)
    Save-ProcessState $trackedProcessIds

    $exitSubscription = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -MessageData @{ ids = $trackedProcessIds; stateFile = $processStateFile } -Action {
        foreach ($processId in $event.MessageData.ids) {
            & taskkill /PID $processId /T /F 2>$null | Out-Null
        }
        Remove-Item -LiteralPath $event.MessageData.stateFile -Force -ErrorAction SilentlyContinue
    }

    Start-Process "http://127.0.0.1:5173"
    Read-Host "Services are running. Press Enter to stop all local services"
}
finally {
    Write-Host "Stopping local services..."
    Stop-ServiceTrees $trackedProcessIds
    Remove-Item -LiteralPath $processStateFile -Force -ErrorAction SilentlyContinue
    if ($exitSubscription) {
        Unregister-Event -SourceIdentifier PowerShell.Exiting -ErrorAction SilentlyContinue
        Remove-Job -Id $exitSubscription.Id -Force -ErrorAction SilentlyContinue
    }
    Write-Host "All local services stopped."
}
