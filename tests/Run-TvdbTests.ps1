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
& $Compiler /nologo /target:exe /langversion:7.3 "/out:$out/TvdbApiTests.exe" "/reference:$json" "/reference:$rest" /reference:Microsoft.CSharp.dll (Join-Path $repo 'BachFlixNfo/TvdbApi.cs') (Join-Path $repo 'BachFlixNfo/Helpers/ApiRetry.cs') (Join-Path $repo 'BachFlixNfo/Helpers/LocalEnvironment.cs') (Join-Path $PSScriptRoot 'TvdbApiTests.cs')
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& "$out/TvdbApiTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'TVDB tests failed.' }
