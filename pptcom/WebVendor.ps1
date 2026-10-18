<#
  Prepares the task pane's third-party files from deps\ (called by WebVendor.targets during the build; it can
  also be run by hand). npm is not used anywhere.

  deps\web-sources.lock.json pins everything (scripts\deps\fetch-web.py writes it and the archives next to it):
    pdfjs   Mozilla's release archive (deps\pdfjs.zip): the files the pane loads are copied out of it
    anydoc  the source of the tagged commit (deps\anydoc.tar.gz), built here to WebAssembly the way anydoc's
            own release does it (wasm-pack: cargo, then wasm-bindgen, then wasm-opt), with each tool pinned:
              rust          the toolchain in the lock, with the wasm32-unknown-unknown target (rustup)
              wasm-bindgen  the prebuilt Windows CLI, downloaded and checked against the lock's SHA256
              binaryen      the prebuilt wasm-opt, likewise
  Each archive is checked against the lock before anything is extracted; a mismatch stops the build.

  Output: -OutDir\files\pdfjs\..., -OutDir\files\anydoc\anydoc_wasm.js and anydoc_wasm_bg.wasm.
  The tools stay in -OutDir\.cache; the cargo build (about 800 MB) is deleted once anydoc is built, so it is
  built again only when the lock or this script changes (WebVendor.targets).

  usage: powershell -ExecutionPolicy Bypass -File WebVendor.ps1 -Lock ..\deps\web-sources.lock.json -OutDir ..\temp\build\pptcom\obj\webvendor\
#>
param(
    [Parameter(Mandatory = $true)][string]$Lock,
    [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"   # Invoke-WebRequest's progress bar makes large downloads crawl
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Lock = [IO.Path]::GetFullPath($Lock)
$deps = Split-Path -Parent $Lock
$OutDir = [IO.Path]::GetFullPath($OutDir)
$files = Join-Path $OutDir "files"
$cache = Join-Path $OutDir ".cache"
$tar = "$env:SystemRoot\System32\tar.exe"   # bsdtar: reads .tar.gz and .zip
$spec = Get-Content $Lock -Raw -Encoding UTF8 | ConvertFrom-Json

# Not Get-FileHash: Windows PowerShell started from PowerShell 7 (MSBuild run from pwsh) inherits pwsh's
# module path and cannot load the module that defines it.
function Sha256([string]$path) {
    $stream = [IO.File]::OpenRead($path)
    try { return -join ([Security.Cryptography.SHA256]::Create().ComputeHash($stream) | ForEach-Object { $_.ToString("x2") }) }
    finally { $stream.Dispose() }
}

function Checked([string]$name) {
    $entry = $spec.$name
    $path = Join-Path $deps $entry.archive
    if (-not (Test-Path $path)) { throw "$name`: $path is missing (run scripts\deps\fetch-web.py)" }
    if ((Sha256 $path) -ne $entry.sha256) { throw "$name`: $path does not match deps\web-sources.lock.json" }
    return $path
}

# Native tools write progress to stderr; only the exit code decides.
function Run([string]$exe, [string[]]$arguments) {
    $saved = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    & $exe @arguments
    $code = $LASTEXITCODE
    $ErrorActionPreference = $saved
    if ($code -ne 0) { throw "$exe failed ($code)" }
}

# A prebuilt tool from the lock: downloaded once, checked, and only the one executable kept.
function Tool([string]$name, [string]$exeInArchive) {
    $entry = $spec.$name
    $dir = Join-Path $cache "$name-$($entry.version)"
    $exe = Join-Path $dir (Split-Path -Leaf $exeInArchive)
    if (Test-Path $exe) { return $exe }
    New-Item -ItemType Directory -Force $dir | Out-Null
    $archive = Join-Path $dir "download.tar.gz"
    Write-Host "WebVendor: downloading $name $($entry.version)"
    Invoke-WebRequest -Uri $entry.url -OutFile $archive -UseBasicParsing
    if ((Sha256 $archive) -ne $entry.sha256) { Remove-Item $archive; throw "$name`: the download does not match deps\web-sources.lock.json" }
    Run $tar @("-xzf", $archive, "-C", $dir, "--strip-components", (($exeInArchive -split "/").Count - 1), $exeInArchive)
    Remove-Item $archive
    if (-not (Test-Path $exe)) { throw "$name`: $exeInArchive not found in the archive" }
    return $exe
}

if (Test-Path $files) { Remove-Item -Recurse -Force $files }
New-Item -ItemType Directory -Force $files, $cache | Out-Null

# ---- PDF.js: copy the files the pane loads out of the release archive ----------------------------------
$zip = Checked "pdfjs"
$x = Join-Path $cache "p"
if (Test-Path $x) { Remove-Item -Recurse -Force $x }
New-Item -ItemType Directory $x | Out-Null
# Left out: the viewer, the standard fonts (the Liberation fonts are GPL-based), the decoders' JavaScript
# fallbacks and the QuickJS sandbox for PDF scripts.
Run $tar @("-xf", $zip, "-C", $x, "build/pdf.mjs", "build/pdf.worker.mjs", "web/cmaps/*.bcmap", "web/iccs/*.icc",
           "web/wasm/openjpeg.wasm", "web/wasm/jbig2.wasm", "web/wasm/qcms_bg.wasm")
$pdfjs = Join-Path $files "pdfjs"
New-Item -ItemType Directory -Force $pdfjs | Out-Null
Copy-Item (Join-Path $x "build\*.mjs") $pdfjs
foreach ($sub in "cmaps", "iccs", "wasm") { Copy-Item -Recurse (Join-Path $x "web\$sub") (Join-Path $pdfjs $sub) }
Remove-Item -Recurse -Force $x
Write-Host "WebVendor: PDF.js $($spec.pdfjs.version) ok"

# ---- anydoc: build the WebAssembly module from the pinned source ----------------------------------------
$toolchain = $spec.rust.toolchain
$targets = ""
if (Get-Command rustup -ErrorAction SilentlyContinue) {
    $saved = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $targets = (& rustup target list --installed --toolchain $toolchain 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) { $targets = "" }
    $ErrorActionPreference = $saved
}
if ($targets -notmatch "wasm32-unknown-unknown") {
    throw "anydoc is built from source with Rust $toolchain and the wasm32-unknown-unknown target. Install them with: rustup toolchain install $toolchain --profile minimal --target wasm32-unknown-unknown"
}
$wasmBindgen = Tool "wasm-bindgen" "wasm-bindgen-$($spec.'wasm-bindgen'.version)-x86_64-pc-windows-msvc/wasm-bindgen.exe"
$wasmOpt = Tool "binaryen" "binaryen-$($spec.binaryen.version)/bin/wasm-opt.exe"

$source = Checked "anydoc"
# short folder names: cargo's build folders are deep
$src = Join-Path $cache "a"
$target = Join-Path $cache "t"
foreach ($d in $src, $target) { if (Test-Path $d) { Remove-Item -Recurse -Force $d } }
New-Item -ItemType Directory $src | Out-Null
try {
    Run $tar @("-xzf", $source, "-C", $src, "--strip-components", "1")
    Write-Host "WebVendor: building anydoc $($spec.anydoc.version) with Rust $toolchain (a few minutes the first time)"
    $env:CARGO_TARGET_DIR = $target
    Run "rustup" @("run", $toolchain, "cargo", "build", "--locked", "--release", "--lib",
                   "--target", "wasm32-unknown-unknown", "--manifest-path", (Join-Path $src "wasm\Cargo.toml"))
    $pkg = Join-Path $cache "pkg"
    if (Test-Path $pkg) { Remove-Item -Recurse -Force $pkg }
    Run $wasmBindgen @((Join-Path $target "wasm32-unknown-unknown\release\anydoc_wasm.wasm"),
                       "--out-dir", $pkg, "--out-name", "anydoc_wasm", "--target", "web")
    # the optimization anydoc's crate asks wasm-pack for ([package.metadata.wasm-pack.profile.release])
    $toml = Get-Content (Join-Path $src "wasm\Cargo.toml") -Raw
    $m = [regex]::Match($toml, 'wasm-opt\s*=\s*\[([^\]]*)\]')
    if (-not $m.Success) { throw "anydoc: no wasm-opt settings in wasm\Cargo.toml" }
    $optFlags = @([regex]::Matches($m.Groups[1].Value, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    $wasm = Join-Path $pkg "anydoc_wasm_bg.wasm"
    Run $wasmOpt (@($wasm, "-o", $wasm) + $optFlags)
    $anydoc = Join-Path $files "anydoc"
    New-Item -ItemType Directory -Force $anydoc | Out-Null
    Copy-Item (Join-Path $pkg "anydoc_wasm.js"), $wasm $anydoc
    Remove-Item -Recurse -Force $pkg
}
finally {
    Remove-Item Env:\CARGO_TARGET_DIR -ErrorAction SilentlyContinue
    foreach ($d in $src, $target) { if (Test-Path $d) { Remove-Item -Recurse -Force $d } }
}
Write-Host "WebVendor: anydoc $($spec.anydoc.version) ok"
