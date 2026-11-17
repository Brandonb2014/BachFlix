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
$json = Join-Path $repo 'packages/Newtonsoft.Json.13.0.3/lib/net45/Newtonsoft.Json.dll'
Copy-Item $json $out -Force
$sources = @('BachFlixNfo/Helpers/SheetHeaderMapping.cs', 'BachFlixNfo/Helpers/VideoResolutionClassifier.cs', 'BachFlixNfo/Features/TvEpisodeMediaScanner.cs', 'BachFlixNfo/Features/TvMediaScanReport.cs', 'tests/TvEpisodeMediaTests.cs') | ForEach-Object { Join-Path $repo $_ }
& $Compiler /nologo /target:exe /langversion:7.3 "/out:$out/TvEpisodeMediaTests.exe" "/reference:$json" $sources
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& "$out/TvEpisodeMediaTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'TV episode media tests failed.' }
