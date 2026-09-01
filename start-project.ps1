$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$frontendRoot = "$projectRoot\frontend"
$backendPort = 8000
$frontendPort = 3000
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

# Load ignored local .env values for this process only. Explicit environment
# variables win, and the file is never copied into logs or repository files.
$localEnv = Join-Path $projectRoot ".env"
if (Test-Path $localEnv) {
    foreach ($line in Get-Content -LiteralPath $localEnv) {
        if ($line -match '^\s*#' -or $line -notmatch '^\s*([A-Za-z_][A-Za-z0-9_]*)=(.*)$') { continue }
        $name = $Matches[1]
        $value = $Matches[2].Trim().Trim('"').Trim("'")
        if ($value -and ($name -like 'PAYROLL_MODEL*' -or $name -like 'DEEPSEEK*')) {
            Set-Item -Path "Env:$name" -Value $value
        } elseif (-not [Environment]::GetEnvironmentVariable($name, "Process") -and $value) {
            Set-Item -Path "Env:$name" -Value $value
        }
    }
}

foreach ($modelVariable in @(
    "PAYROLL_MODEL_PROVIDER",
    "PAYROLL_MODEL_BASE_URL",
    "PAYROLL_MODEL_API_KEY",
    "PAYROLL_MODEL_NAME",
    "PAYROLL_MODEL_API_STYLE",
    "PAYROLL_MODEL_REASONING_EFFORT",
    "PAYROLL_MODEL_TIMEOUT_SECONDS",
    "PAYROLL_MODEL_MAX_OUTPUT_TOKENS",
    "PAYROLL_MODEL_ENABLE_THINKING",
    "PAYROLL_MODEL_FALLBACK_PROVIDER",
    "PAYROLL_MODEL_FALLBACK_BASE_URL",
    "PAYROLL_MODEL_FALLBACK_API_KEY",
    "PAYROLL_MODEL_FALLBACK_NAME",
    "PAYROLL_MODEL_FALLBACK_API_STYLE",
    "PAYROLL_MODEL_FALLBACK_REASONING_EFFORT",
    "PAYROLL_MODEL_FALLBACK_TIMEOUT_SECONDS",
    "PAYROLL_MODEL_FALLBACK_MAX_OUTPUT_TOKENS",
    "PAYROLL_MODEL_FALLBACK_ENABLE_THINKING",
    "DEEPSEEK_API_KEY",
    "DEEPSEEK_BASE_URL",
    "DEEPSEEK_MODEL",
    "DEEPSEEK_FALLBACK_MODEL",
    "DEEPSEEK_REASONING_EFFORT",
    "DEEPSEEK_TIMEOUT_SECONDS",
    "DEEPSEEK_MAX_OUTPUT_TOKENS"
)) {
    $modelValue = [Environment]::GetEnvironmentVariable($modelVariable, "User")
    if ($modelValue -and -not [Environment]::GetEnvironmentVariable($modelVariable, "Process")) {
        Set-Item -Path "Env:$modelVariable" -Value $modelValue
    }
}
if ($env:PAYROLL_MODEL_BASE_URL -like '*api.deepseek.com*' -and $env:DEEPSEEK_API_KEY) {
    $env:PAYROLL_MODEL_API_KEY = $env:DEEPSEEK_API_KEY
}

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
