param([Parameter(Mandatory=$true)][string]$Python)
Push-Location (Join-Path $PSScriptRoot '..\backend')
try { & $Python -m unittest discover -s tests -v; if ($LASTEXITCODE -ne 0) {throw 'Tests failed'} }
finally { Pop-Location }
