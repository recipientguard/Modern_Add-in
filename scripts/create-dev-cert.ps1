# Creates and trusts a localhost development certificate for the POC server.
# This modifies the CurrentUser Root certificate store. Run only if you accept that local trust change.

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$certDir = Join-Path $root ".certs"
$pfxPath = Join-Path $certDir "localhost.pfx"
$cerPath = Join-Path $certDir "localhost.cer"
$password = ConvertTo-SecureString -String "recipientguard" -Force -AsPlainText

New-Item -ItemType Directory -Path $certDir -Force | Out-Null

$cert = New-SelfSignedCertificate `
  -DnsName "localhost" `
  -CertStoreLocation "Cert:\CurrentUser\My" `
  -NotAfter (Get-Date).AddYears(2)

Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $password | Out-Null
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\CurrentUser\Root" | Out-Null

Write-Host "Created trusted localhost certificate: $pfxPath"
Write-Host "Password: recipientguard"
