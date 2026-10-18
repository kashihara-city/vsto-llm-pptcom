<#
  Builds the release MSI: release\vsto-llm-pptcom-<version>.msi (release\ is git-ignored and holds only the MSI).

  1. Restores and builds the add-in (Release) into temp\build\pptcom\bin\Release\. Needs temp\signing\ (ClickOnce signing; made by
     scripts\build\new-signing.ps1).
  2. Reads the signing public key from the signed pptcom.vsto and writes temp\build\pptcom.Installer\obj\generated\Signing.wxi
     (the key belongs to this machine's certificate, so it is never committed).
  3. Builds the WiX v5 project (WixToolset.Sdk is restored from nuget.org) and copies the MSI to release\.

  Version: <ApplicationVersion> in pptcom.csproj (e.g. 0.0.1.0 -> MSI 0.0.1).

  usage: powershell -ExecutionPolicy Bypass -File scripts\build\build.ps1 [-WarningsLog <file>]
         -WarningsLog: the add-in build's warnings also go to <file>, for scripts\test\build-warnings.py.
#>
param([string]$WarningsLog)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$csproj = Join-Path $root "pptcom\pptcom.csproj"
$bin = Join-Path $root "temp\build\pptcom\bin\Release\"
$generated = Join-Path $root "temp\build\pptcom.Installer\obj\generated\"
$release = Join-Path $root "release"

# $env:MSBUILD when set (the release workflow sets it), otherwise the newest Visual Studio or Build Tools
# (-products *: without it vswhere leaves Build Tools out).
$msbuild = $env:MSBUILD
if (-not $msbuild) {
    $msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
}
if (-not $msbuild) { throw "MSBuild (Visual Studio or Build Tools) not found" }
"MSBuild: $msbuild"

# 1. add-in. A VSTO build also registers the add-in for this user (HKCU ...\Addins\pptcom) at its
#    output folder; put back whatever was registered before (usually the developer's Debug build).
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\pptcom"
$before = if (Test-Path $addinKey) { (Get-ItemProperty $addinKey).Manifest } else { $null }
$log = @()
if ($WarningsLog) { $log += "-flp:logfile=$WarningsLog;warningsonly" }
& $msbuild $csproj -restore -p:Configuration=Release -p:Platform=AnyCPU -nologo -v:minimal @log
$buildExit = $LASTEXITCODE
if ($before) { Set-ItemProperty $addinKey -Name Manifest -Value $before }
elseif (Test-Path $addinKey) { Remove-Item $addinKey -Recurse }
if ($buildExit -ne 0) { throw "add-in build failed" }

# version from the csproj (0.0.1.0 -> 0.0.1)
[xml]$proj = Get-Content $csproj -Raw
$appVersion = ($proj.Project.PropertyGroup | ForEach-Object { $_.ApplicationVersion } | Where-Object { $_ } | Select-Object -First 1)
if (-not $appVersion) { throw "ApplicationVersion not found in the csproj" }
$parts = $appVersion.Split(".")
$version = ($parts[0..2] -join ".")

# 2. signing public key -> Signing.wxi
$vsto = Get-Content (Join-Path $bin "pptcom.vsto") -Raw
$m = [regex]::Match($vsto, "<RSAKeyValue>.*?</RSAKeyValue>", "Singleline")
if (-not $m.Success) { throw "no RSAKeyValue in pptcom.vsto (is the manifest signed?)" }
New-Item -ItemType Directory -Force $generated | Out-Null
# The value goes into a preprocessor instruction verbatim (not XML-unescaped), so no escaping here.
$wxi = "<?xml version=`"1.0`" encoding=`"utf-8`"?>`r`n<Include>`r`n  <?define PublicKey = `"$($m.Value)`" ?>`r`n</Include>`r`n"
[IO.File]::WriteAllText((Join-Path $generated "Signing.wxi"), $wxi, (New-Object Text.UTF8Encoding($false)))

# 3. MSI
dotnet build (Join-Path $PSScriptRoot "pptcom.Installer.wixproj") -c Release -nologo -v:minimal `
    "-p:ProductVersion=$version" "-p:AddinBinDir=$bin" "-p:GeneratedDir=$generated"
if ($LASTEXITCODE -ne 0) { throw "MSI build failed" }

$msi = Get-ChildItem (Join-Path $root "temp\build\pptcom.Installer\bin\Release") -Recurse -Filter "vsto-llm-pptcom.msi" | Select-Object -First 1
if (-not $msi) { throw "vsto-llm-pptcom.msi not found" }
if (Test-Path $release) { Remove-Item -Recurse -Force $release }
New-Item -ItemType Directory -Force $release | Out-Null
$out = Join-Path $release "vsto-llm-pptcom-$version.msi"
Copy-Item $msi.FullName $out
"built: $out ($([math]::Round((Get-Item $out).Length / 1KB)) KB)"
