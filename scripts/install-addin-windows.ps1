param(
  [string]$CatalogPath = "$env:LOCALAPPDATA\mdpr-ppt\AddinCatalog",
  [string]$ShareName = "mdpr-ppt-addins",
  [switch]$TryShare
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$manifestPath = Join-Path $repoRoot "packages\addin\manifest.xml"
$targetManifest = Join-Path $CatalogPath "mdpr-ppt.xml"
$instructionsPath = Join-Path $CatalogPath "install-next-steps.txt"

if (-not (Test-Path -LiteralPath $manifestPath)) {
  throw "Add-in manifest not found: $manifestPath"
}

New-Item -ItemType Directory -Force -Path $CatalogPath | Out-Null
Copy-Item -LiteralPath $manifestPath -Destination $targetManifest -Force

$catalogUrl = $CatalogPath
$shareStatus = "not requested"
if ($TryShare) {
  try {
    $existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if (-not $existing) {
      New-SmbShare -Name $ShareName -Path $CatalogPath -ReadAccess $env:USERNAME | Out-Null
    }
    $catalogUrl = "\\$env:COMPUTERNAME\$ShareName"
    $shareStatus = "created or already available"
  } catch {
    $shareStatus = "failed: $($_.Exception.Message)"
  }
}

$instructions = @"
mdpr-ppt PowerPoint add-in registration

Manifest copied to:
$targetManifest

Catalog folder:
$CatalogPath

Shared folder status:
$shareStatus

Catalog path prepared by this script:
$catalogUrl

PowerPoint's SHARED FOLDER flow works best with a UNC shared-folder catalog,
for example \\COMPUTER\mdpr-ppt-addins. If the value above is a local folder
path, share that folder manually or rerun this script from an elevated
PowerShell session with -TryShare.

PowerPoint steps:
1. Open PowerPoint.
2. Go to File > Options > Trust Center > Trust Center Settings.
3. Open Trusted Add-in Catalogs.
4. Add the Catalog URL above.
5. Enable Show in Menu for that catalog.
6. Restart PowerPoint.
7. Go to Home > Add-ins > Advanced.
8. Choose SHARED FOLDER.
9. Add mdpr-ppt.

Taskpane asset server:
npm run serve:addin -- --cert .\certs\localhost.crt --key .\certs\localhost.key

Office add-in SourceLocation requires an HTTPS endpoint for the manifest URL.
"@

Set-Content -LiteralPath $instructionsPath -Value $instructions -Encoding UTF8

Write-Output "Manifest copied: $targetManifest"
Write-Output "Catalog path: $catalogUrl"
Write-Output "Instructions: $instructionsPath"
if ($TryShare -and $shareStatus.StartsWith("failed")) {
  Write-Warning "Shared folder creation failed. Add the local catalog folder manually in PowerPoint Trusted Add-in Catalogs, or rerun from an elevated PowerShell session."
}
