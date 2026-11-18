param([Parameter(Mandatory=$true)][string]$Python, [int]$Port=8570, [string]$EnvFile)
$ErrorActionPreference='Stop'
if ($EnvFile) {
    foreach ($line in Get-Content -LiteralPath $EnvFile) {
        if ($line -match '^([A-Z][A-Z0-9_]*)=(.*)$') {
            [Environment]::SetEnvironmentVariable($matches[1], $matches[2], 'Process')
        }
    }
}
if (Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction SilentlyContinue) {
    throw "Port $Port is already in use. Select another -Port; no service was stopped."
}
$taskBackend=Join-Path $PSScriptRoot '..\backend'
Push-Location $taskBackend
try { & $Python -m uvicorn app.main:app --host 127.0.0.1 --port $Port --no-access-log }
finally { Pop-Location }
