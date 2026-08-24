$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$frontendRoot = "$projectRoot\frontend"
$backendPort = 8000
$frontendPort = 3002
$backendUrl = "http://127.0.0.1:$backendPort"
$frontendUrl = "http://127.0.0.1:$frontendPort"
$logRoot = "$projectRoot\logs"
$dataRoot = "$projectRoot\backend\data"

function Test-TcpPort([int] $port) {
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $task = $client.ConnectAsync("127.0.0.1", $port)
        if (-not $task.Wait(500)) { return $false }
        return $client.Connected
    } catch {
        return $false
    } finally {
        $client.Dispose()
    }
}

function Wait-TcpPort([int] $port, [string] $serviceName) {
    for ($i = 0; $i -lt 40; $i++) {
        if (Test-TcpPort $port) { return }
        Start-Sleep -Milliseconds 500
    }
    throw "$serviceName startup timed out. Check the logs folder."
}

if (-not (Test-Path "$frontendRoot\node_modules")) {
    throw "Missing frontend\node_modules. This launcher does not download dependencies."
}

$pythonExe = "$projectRoot\.venv-codex\Scripts\python.exe"
if (-not (Test-Path $pythonExe)) {
    $pythonExe = "$projectRoot\.venv\Scripts\python.exe"
}
if (-not (Test-Path $pythonExe)) {
    throw "Missing project Python environment (.venv-codex or .venv)."
}

$nodeExe = ""
$bundledNode = "C:\Users\keen\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
if (Test-Path $bundledNode) {
    $nodeExe = $bundledNode
} else {
    $nodeCommand = Get-Command node -ErrorAction SilentlyContinue
    if ($nodeCommand) { $nodeExe = $nodeCommand.Source }
}
if (-not $nodeExe) {
    throw "Missing Node.js runtime."
}

New-Item -ItemType Directory -Force -Path $logRoot, $dataRoot | Out-Null
$env:PAYROLL_DATA_DIR = $dataRoot

if (-not (Test-TcpPort $backendPort)) {
    Start-Process -FilePath $pythonExe `
        -ArgumentList @("-m", "uvicorn", "backend.main:app", "--host", "127.0.0.1", "--port", "$backendPort") `
        -WorkingDirectory $projectRoot `
        -RedirectStandardOutput "$logRoot\backend.log" `
        -RedirectStandardError "$logRoot\backend-error.log" `
        -WindowStyle Hidden | Out-Null
}
Wait-TcpPort $backendPort "Backend"

if (-not (Test-TcpPort $frontendPort)) {
    Start-Process -FilePath $nodeExe `
        -ArgumentList @("node_modules\next\dist\bin\next", "dev", "-p", "$frontendPort") `
        -WorkingDirectory $frontendRoot `
        -RedirectStandardOutput "$logRoot\frontend.log" `
        -RedirectStandardError "$logRoot\frontend-error.log" `
        -WindowStyle Hidden | Out-Null
}
Wait-TcpPort $frontendPort "Frontend"

Write-Output "Project started: $frontendUrl"
Write-Output "Backend: $backendUrl"
