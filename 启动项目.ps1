param(
    # 常驻守护模式：进程退出后自动重启（等价 PM2，零额外依赖）
    [switch]$Watchdog,
    # 守护巡检间隔（秒）
    [int]$WatchdogIntervalSeconds = 5
)

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
        # Model settings in the project .env are authoritative for this app;
        # do not let stale desktop-shell variables silently select another model.
        if ($value -and ($name -like 'PAYROLL_MODEL*' -or $name -like 'DEEPSEEK*')) {
            Set-Item -Path "Env:$name" -Value $value
        } elseif (-not [Environment]::GetEnvironmentVariable($name, "Process") -and $value) {
            Set-Item -Path "Env:$name" -Value $value
        }
    }
}
# When DeepSeek is the selected primary profile, never inherit a stale
# PAYROLL_MODEL_API_KEY from the desktop shell (it may belong to another
# provider). Reuse the project-scoped DeepSeek key instead.
if ($env:PAYROLL_MODEL_BASE_URL -like '*api.deepseek.com*' -and $env:DEEPSEEK_API_KEY) {
    $env:PAYROLL_MODEL_API_KEY = $env:DEEPSEEK_API_KEY
}

# Model credentials are stored in the current Windows user's environment,
# never in this repository. Explicitly load them because an already-running
# desktop shell may not inherit values created after it started.
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
# Apply this after loading user-level variables so a stale desktop-shell key
# cannot overwrite the project DeepSeek credential.
if ($env:PAYROLL_MODEL_BASE_URL -like '*api.deepseek.com*' -and $env:DEEPSEEK_API_KEY) {
    $env:PAYROLL_MODEL_API_KEY = $env:DEEPSEEK_API_KEY
}

function Start-Backend {
    Start-Process -FilePath $pythonExe `
        -ArgumentList @("-m", "uvicorn", "backend.main:app", "--host", "127.0.0.1", "--port", "$backendPort") `
        -WorkingDirectory $projectRoot `
        -RedirectStandardOutput "$logRoot\backend.log" `
        -RedirectStandardError "$logRoot\backend-error.log" `
        -WindowStyle Hidden | Out-Null
}

function Start-Frontend {
    Start-Process -FilePath $nodeExe `
        -ArgumentList @("node_modules\next\dist\bin\next", "dev", "-p", "$frontendPort") `
        -WorkingDirectory $frontendRoot `
        -RedirectStandardOutput "$logRoot\frontend.log" `
        -RedirectStandardError "$logRoot\frontend-error.log" `
        -WindowStyle Hidden | Out-Null
}

function Write-WatchdogLog([string]$Message) {
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path "$logRoot\watchdog.log" -Value $line
}

if (-not (Test-TcpPort $backendPort)) { Start-Backend }
Wait-TcpPort $backendPort "Backend"

if (-not (Test-TcpPort $frontendPort)) { Start-Frontend }
Wait-TcpPort $frontendPort "Frontend"

Write-Output "Project started: $frontendUrl"
Write-Output "Backend: $backendUrl"

if (-not $Watchdog) { exit 0 }

# ---- 常驻守护 ----
Write-Output ""
Write-Output "Watchdog running (interval ${WatchdogIntervalSeconds}s). Keep this window open; exited services restart automatically."
Write-WatchdogLog "watchdog started (interval ${WatchdogIntervalSeconds}s)"

$restartHistory = @{ backend = @(); frontend = @() }
while ($true) {
    Start-Sleep -Seconds $WatchdogIntervalSeconds
    foreach ($name in @("backend", "frontend")) {
        $port = if ($name -eq "backend") { $backendPort } else { $frontendPort }
        if (Test-TcpPort $port) { continue }
        # 崩溃循环保护：2 分钟内连续崩溃 3 次则退避 60 秒，避免无限拉起
        $now = Get-Date
        $recent = @($restartHistory[$name] | Where-Object { ($now - $_).TotalSeconds -lt 120 })
        if ($recent.Count -ge 3) {
            Write-WatchdogLog "$name crashed 3 times within 2 minutes; backing off 60s. Check logs\${name}-error.log"
            Start-Sleep -Seconds 60
            $restartHistory[$name] = @()
        }
        if ($name -eq "backend") { Start-Backend } else { Start-Frontend }
        $restartHistory[$name] = @($recent + (Get-Date))
        Write-WatchdogLog "$name on port $port was down; restarted"
    }
}
