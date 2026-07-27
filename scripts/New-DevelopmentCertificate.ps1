[CmdletBinding()]
param(
    [string]$Subject = "CN=IME Automation Report Automation Development",
    [int]$ValidYears = 3
)

$certificate = New-SelfSignedCertificate `
    -Type CodeSigningCert `
    -Subject $Subject `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy NonExportable `
    -NotAfter (Get-Date).AddYears($ValidYears)

Write-Host "Development certificate created."
Write-Host "Thumbprint: $($certificate.Thumbprint)"
Write-Host "Use this certificate for local testing only. Use an organisation-trusted code-signing certificate for deployment."
