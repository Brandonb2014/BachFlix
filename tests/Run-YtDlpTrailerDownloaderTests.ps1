param([string]$Compiler)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
if (!$Compiler) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $install = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
    $Compiler = Join-Path $install 'MSBuild/Current/Bin/Roslyn/csc.exe'
}
$out = Join-Path $PSScriptRoot 'bin'
New-Item -ItemType Directory -Force $out | Out-Null
$sources = @('BachFlixNfo/Features/YtDlpTrailerDownloader.cs', 'tests/YtDlpTrailerDownloaderTests.cs') | ForEach-Object { Join-Path $repo $_ }
& $Compiler /nologo /target:exe /langversion:7.3 "/out:$out/YtDlpTrailerDownloaderTests.exe" $sources
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& "$out/YtDlpTrailerDownloaderTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'Trailer downloader tests failed.' }
