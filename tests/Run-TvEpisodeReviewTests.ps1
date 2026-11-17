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
$refs = @(
    'packages/Newtonsoft.Json.13.0.3/lib/net45/Newtonsoft.Json.dll',
    'packages/Google.Apis.1.68.0/lib/net462/Google.Apis.dll',
    'packages/Google.Apis.Core.1.68.0/lib/net462/Google.Apis.Core.dll',
    'packages/Google.Apis.Sheets.v4.1.48.0.2037/lib/net45/Google.Apis.Sheets.v4.dll'
) | ForEach-Object { Join-Path $repo $_ }
Copy-Item $refs $out -Force
$sources = @(
    'BachFlixNfo/Features/TvEpisodeMatcher.cs',
    'BachFlixNfo/Features/TvEpisodeShadowRun.cs',
    'BachFlixNfo/Features/FileBotSubtitleNaming.cs',
    'BachFlixNfo/Helpers/SheetHeaderMapping.cs',
    'BachFlixNfo/Features/TvEpisodeReview.cs',
    'BachFlixNfo/Features/TvShowDbUpdate.cs',
    'tests/TvEpisodeReviewTests.cs'
) | ForEach-Object { Join-Path $repo $_ }
$refArgs = $refs | ForEach-Object { "/reference:$_" }
& $Compiler /nologo /target:exe /langversion:7.3 "/out:$out/TvEpisodeReviewTests.exe" $refArgs /reference:System.Net.Http.dll $sources
if ($LASTEXITCODE -ne 0) { throw 'Test compilation failed.' }
& "$out/TvEpisodeReviewTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'TV episode review tests failed.' }
