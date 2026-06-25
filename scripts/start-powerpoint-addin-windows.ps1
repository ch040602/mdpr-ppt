param(
  [int]$Port = 3000,
  [string]$PfxPassword = "mdpr-ppt-localhost",
  [switch]$Debug
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$manifestPath = Join-Path $repoRoot "packages\addin\manifest.xml"
$certDir = Join-Path $repoRoot ".certs"
$pfxPath = Join-Path $certDir "localhost.pfx"
$serverPidPath = Join-Path $certDir "mdpr-ppt-addin-server.pid"
$serverLogPath = Join-Path $certDir "mdpr-ppt-addin-server.log"
$serverErrorLogPath = Join-Path $certDir "mdpr-ppt-addin-server.err.log"

function Stop-ExistingAssetServer {
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
}

function Ensure-LocalhostCertificate {
  New-Item -ItemType Directory -Force -Path $certDir | Out-Null

  $certProviderAvailable = $null -ne (Get-PSDrive -Name Cert -ErrorAction SilentlyContinue)
  if (-not $certProviderAvailable) {
    Ensure-LocalhostCertificateWithoutCertProvider
    return
  }

  $cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq "CN=localhost" -and $_.FriendlyName -eq "mdpr-ppt localhost" } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

  if (-not $cert) {
    $cert = New-SelfSignedCertificate `
      -DnsName "localhost" `
      -CertStoreLocation "Cert:\CurrentUser\My" `
      -FriendlyName "mdpr-ppt localhost" `
      -NotAfter (Get-Date).AddYears(2)
  }

  $rootCert = Get-ChildItem Cert:\CurrentUser\Root |
    Where-Object { $_.Thumbprint -eq $cert.Thumbprint } |
    Select-Object -First 1
  if (-not $rootCert) {
    Export-Certificate -Cert $cert -FilePath (Join-Path $certDir "localhost.cer") | Out-Null
    Import-Certificate -FilePath (Join-Path $certDir "localhost.cer") -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null
  }

  $securePassword = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
  Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword | Out-Null
}

function Ensure-LocalhostCertificateWithoutCertProvider {
  $cerPath = Join-Path $certDir "localhost.cer"
  if (-not (Test-Path -LiteralPath $pfxPath) -or -not (Test-Path -LiteralPath $cerPath)) {
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
      "CN=localhost",
      $rsa,
      [System.Security.Cryptography.HashAlgorithmName]::SHA256,
      [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )
    $sanBuilder = [System.Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder]::new()
    $sanBuilder.AddDnsName("localhost")
    $request.CertificateExtensions.Add($sanBuilder.Build())
    $request.CertificateExtensions.Add(
      [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature,
        $false
      )
    )
    $serverAuthOid = [System.Security.Cryptography.Oid]::new("1.3.6.1.5.5.7.3.1")
    $enhancedUsages = [System.Security.Cryptography.OidCollection]::new()
    $enhancedUsages.Add($serverAuthOid) | Out-Null
    $request.CertificateExtensions.Add(
      [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($enhancedUsages, $false)
    )
    $certificate = $request.CreateSelfSigned(
      [System.DateTimeOffset]::Now.AddDays(-1),
      [System.DateTimeOffset]::Now.AddYears(2)
    )
    [System.IO.File]::WriteAllBytes($pfxPath, $certificate.Export(
      [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx,
      $PfxPassword
    ))
    [System.IO.File]::WriteAllBytes($cerPath, $certificate.Export(
      [System.Security.Cryptography.X509Certificates.X509ContentType]::Cert
    ))
  }

  & certutil -user -addstore Root $cerPath | Out-Null
}

function Start-AssetServer {
  $arguments = @(
    (Join-Path $repoRoot "scripts\serve-addin.mjs"),
    "--port", "$Port",
    "--pfx", $pfxPath,
    "--passphrase", $PfxPassword
  )
  $process = Start-Process `
    -FilePath "node" `
    -ArgumentList $arguments `
    -WorkingDirectory $repoRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $serverLogPath `
    -RedirectStandardError $serverErrorLogPath `
    -PassThru
  Set-Content -LiteralPath $serverPidPath -Value $process.Id -Encoding ASCII
}

function Wait-AssetServer {
  $url = "https://localhost:$Port/taskpane/index.html"
  for ($attempt = 0; $attempt -lt 40; $attempt += 1) {
    try {
      $response = Invoke-WebRequest -UseBasicParsing -Uri $url -TimeoutSec 2
      if ($response.StatusCode -eq 200) {
        return
      }
    } catch {
      Start-Sleep -Milliseconds 500
    }
  }
  throw "Timed out waiting for HTTPS localhost asset server: $url"
}

Stop-ExistingAssetServer
Ensure-LocalhostCertificate
Start-AssetServer
Wait-AssetServer

$stopArgs = @("--yes", "office-addin-debugging", "stop", $manifestPath)
& npx @stopArgs | Out-Null

$debugArgs = @("--yes", "office-addin-debugging", "start", $manifestPath, "desktop", "--app", "powerpoint")
if (-not $Debug) {
  $debugArgs += "--no-debug"
}
& npx @debugArgs

Write-Output "HTTPS localhost asset server: https://localhost:$Port/taskpane/index.html"
Write-Output "Server PID file: $serverPidPath"
