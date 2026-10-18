<#
  Makes the test certificate that signs the add-in's ClickOnce manifests (a VSTO build needs one) and the settings the
  build reads, in temp\signing\ (git-ignored): signing.pfx and signing.local.props (SignManifests, ManifestKeyFile,
  ManifestCertificateThumbprint, ManifestPassword; pptcom.csproj imports it). Run it once on a new machine, or again
  when temp\signing\ is gone. The certificate is self-signed, for code signing, in CurrentUser\My, and is not made
  trusted; the password is random and is not shown. The release and check workflows make one the same way for each run.

  With a new certificate the development build registers itself again when it is built, and the next MSI carries the
  new key (an MSI built with the old one is installed again from a new build).

  usage: powershell -ExecutionPolicy Bypass -File scripts\build\new-signing.ps1 [-Subject <name>] [-Days <n>] [-Force]
         -Force: make a new one when temp\signing\ has one already (the old certificate stays in CurrentUser\My).
#>
param([string]$Subject = "CN=pptcom test build", [int]$Days = 3650, [switch]$Force)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$dir = Join-Path $root "temp\signing"
$props = Join-Path $dir "signing.local.props"
if ((Test-Path $props) -and -not $Force) {
    throw "$props is there already: the build uses it. Use -Force to make a new certificate."
}

New-Item -ItemType Directory -Force $dir | Out-Null
$cert = New-SelfSignedCertificate -Subject $Subject -Type CodeSigningCert -CertStoreLocation Cert:\CurrentUser\My `
    -NotAfter (Get-Date).AddDays($Days)
# RandomNumberGenerator.Create() works in Windows PowerShell 5.1 and in PowerShell 7 alike.
$bytes = New-Object byte[] 24
$rng = [Security.Cryptography.RandomNumberGenerator]::Create()
$rng.GetBytes($bytes)
$rng.Dispose()
$password = [Convert]::ToBase64String($bytes)
$pfx = Join-Path $dir "signing.pfx"
Export-PfxCertificate -Cert $cert -FilePath $pfx -Password (ConvertTo-SecureString $password -AsPlainText -Force) | Out-Null

$xml = @"
<Project xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <!-- Made by scripts\build\new-signing.ps1: the ClickOnce signing of pptcom on this machine. -->
  <PropertyGroup>
    <SignManifests>true</SignManifests>
    <ManifestKeyFile>`$(MSBuildThisFileDirectory)signing.pfx</ManifestKeyFile>
    <ManifestCertificateThumbprint>$($cert.Thumbprint)</ManifestCertificateThumbprint>
    <ManifestPassword>$password</ManifestPassword>
  </PropertyGroup>
</Project>
"@
[IO.File]::WriteAllText($props, $xml, (New-Object Text.UTF8Encoding($false)))
"made $pfx and $props (thumbprint $($cert.Thumbprint), valid until $($cert.NotAfter.ToString('yyyy-MM-dd')))"
