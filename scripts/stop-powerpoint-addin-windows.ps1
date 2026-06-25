param(
  [int]$Port = 3000
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$manifestPath = Join-Path $repoRoot "packages\addin\manifest.xml"
$certDir = Join-Path $repoRoot ".certs"
$serverPidPath = Join-Path $certDir "mdpr-ppt-addin-server.pid"

$stopArgs = @("--yes", "office-addin-debugging", "stop", $manifestPath)
& npx @stopArgs

if (Test-Path -LiteralPath $serverPidPath) {
  $oldPid = Get-Content -LiteralPath $serverPidPath -ErrorAction SilentlyContinue | Select-Object -First 1
  if ($oldPid) {
    Stop-Process -Id ([int]$oldPid) -Force -ErrorAction SilentlyContinue
  }
  Remove-Item -LiteralPath $serverPidPath -Force -ErrorAction SilentlyContinue
}

Get-CimInstance Win32_Process |
  Where-Object { $_.Name -eq "node.exe" -and $_.CommandLine -like "*scripts/serve-addin.mjs*" -and $_.CommandLine -like "*--port $Port*" } |
  ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }

Write-Output "Stopped mdpr-ppt PowerPoint sideload and HTTPS asset server."
