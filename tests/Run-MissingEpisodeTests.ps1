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
$rest = Join-Path $repo 'packages/RestSharp.106.11.4/lib/net452/RestSharp.dll'
Copy-Item $json,$rest $out -Force
$sources = @('BachFlixNfo/TvdbApi.cs', 'BachFlixNfo/EmbyApi.cs', 'BachFlixNfo/Helpers/ApiRetry.cs', 'BachFlixNfo/Helpers/LocalEnvironment.cs', 'BachFlixNfo/Helpers/SheetHeaderMapping.cs', 'BachFlixNfo/Features/MissingEpisodePreview.cs', 'BachFlixNfo/Features/MissingEpisodeWorkflow.cs', 'tests/MissingEpisodeTests.cs', 'tests/MissingEpisodeWorkflowTests.cs') | ForEach-Object { Join-Path $repo $_ }
& $Compiler /nologo /target:exe /langversion:7.3 "/out:$out/MissingEpisodeTests.exe" "/reference:$json" "/reference:$rest" /reference:System.Net.Http.dll /reference:Microsoft.CSharp.dll $sources
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& "$out/MissingEpisodeTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'Missing-episode tests failed.' }
