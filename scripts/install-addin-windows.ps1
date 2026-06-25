param(
  [string]$CatalogPath = "$env:LOCALAPPDATA\mdpr-ppt\AddinCatalog",
  [string]$ShareName = "mdpr-ppt-addins",
  [switch]$TryShare,
  [switch]$RegisterTrustCatalog
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

$trustCatalogStatus = "not requested"
$trustedCatalogsKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
if ($RegisterTrustCatalog -and -not $catalogUrl.StartsWith("\\")) {
  $trustCatalogStatus = "skipped: trusted shared-folder catalog requires a UNC path; rerun from elevated PowerShell with -TryShare or add the catalog manually"
} elseif ($RegisterTrustCatalog) {
  try {
    New-Item -Path $trustedCatalogsKey -Force | Out-Null
    $existingCatalog = Get-ChildItem -Path $trustedCatalogsKey -ErrorAction SilentlyContinue |
      Where-Object { (Get-ItemProperty -LiteralPath $_.PSPath -Name Url -ErrorAction SilentlyContinue).Url -eq $catalogUrl } |
      Select-Object -First 1

    if ($existingCatalog) {
      $catalogKey = $existingCatalog.PSPath
      $catalogId = $existingCatalog.PSChildName
    } else {
      $catalogId = "{" + (New-Guid).Guid + "}"
      $catalogKey = Join-Path $trustedCatalogsKey $catalogId
      New-Item -Path $catalogKey -Force | Out-Null
    }

    New-ItemProperty -Path $catalogKey -Name "Id" -Value $catalogId -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $catalogKey -Name "Url" -Value $catalogUrl -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $catalogKey -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null
    $trustCatalogStatus = "registered under $catalogKey"
  } catch {
    $trustCatalogStatus = "failed: $($_.Exception.Message)"
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

Trusted Add-in Catalogs registry status:
$trustCatalogStatus

PowerPoint's SHARED FOLDER flow works best with a UNC shared-folder catalog,
for example \\COMPUTER\mdpr-ppt-addins. If the value above is a local folder
path, share that folder manually or rerun this script from an elevated
PowerShell session with -TryShare. To register the catalog for the current
Windows user, rerun with -RegisterTrustCatalog. For the closest in-PowerPoint
experience, use both -TryShare and -RegisterTrustCatalog, then restart
PowerPoint.

PowerPoint steps:
1. Open PowerPoint.
2. Go to File > Options > Trust Center > Trust Center Settings.
3. Open Trusted Add-in Catalogs.
4. Confirm the catalog URL above is listed.
5. Enable Show in Menu for that catalog if it is not already enabled.
6. Restart PowerPoint.
7. Go to the MDPR tab and choose Inspect Selection. If the tab is not visible
   yet, go to Home > Add-ins > Advanced > SHARED FOLDER and add mdpr-ppt once.

Taskpane asset server:
npm run serve:addin -- --cert .\certs\localhost.crt --key .\certs\localhost.key

Office add-in SourceLocation requires an HTTPS endpoint for the manifest URL.
"@

Set-Content -LiteralPath $instructionsPath -Value $instructions -Encoding UTF8

Write-Output "Manifest copied: $targetManifest"
Write-Output "Catalog path: $catalogUrl"
Write-Output "Trusted catalog: $trustCatalogStatus"
Write-Output "Instructions: $instructionsPath"
if ($TryShare -and $shareStatus.StartsWith("failed")) {
  Write-Warning "Shared folder creation failed. Add the local catalog folder manually in PowerPoint Trusted Add-in Catalogs, or rerun from an elevated PowerShell session."
}
if ($RegisterTrustCatalog -and $trustCatalogStatus.StartsWith("failed")) {
  Write-Warning "Trusted catalog registry registration failed. Add the catalog manually in PowerPoint Trust Center."
}
