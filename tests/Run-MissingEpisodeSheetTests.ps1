param([string]$Compiler)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
if (!$Compiler) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $install = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
    $Compiler = Join-Path $install 'MSBuild/Current/Bin/Roslyn/csc.exe'
}
$out = Join-Path $PSScriptRoot 'bin/MissingEpisodeSheets'
New-Item -ItemType Directory -Force $out | Out-Null
$relative = @(
    'packages/Newtonsoft.Json.13.0.3/lib/net45/Newtonsoft.Json.dll',
    'packages/RestSharp.106.11.4/lib/net452/RestSharp.dll',
    'packages/Google.Apis.Core.1.68.0/lib/netstandard2.0/Google.Apis.Core.dll',
    'packages/Google.Apis.1.68.0/lib/netstandard2.0/Google.Apis.dll',
    'packages/Google.Apis.Sheets.v4.1.48.0.2037/lib/net45/Google.Apis.Sheets.v4.dll'
)
$references = @($relative | ForEach-Object { Join-Path $repo $_ })
Copy-Item -LiteralPath $references -Destination $out -Force
$facade = Join-Path ${env:ProgramFiles(x86)} 'Reference Assemblies/Microsoft/Framework/.NETFramework/v4.8/Facades/netstandard.dll'
$arguments = @('/nologo', '/target:exe', '/langversion:7.3', "/out:$out/MissingEpisodeSheetAdapterTests.exe", '/reference:System.Net.Http.dll', '/reference:Microsoft.CSharp.dll', "/reference:$facade")
$arguments += $references | ForEach-Object { "/reference:$_" }
$sources = @('BachFlixNfo/TvdbApi.cs', 'BachFlixNfo/EmbyApi.cs', 'BachFlixNfo/Helpers/ApiRetry.cs', 'BachFlixNfo/Helpers/LocalEnvironment.cs', 'BachFlixNfo/Helpers/SheetHeaderMapping.cs', 'BachFlixNfo/Features/MissingEpisodePreview.cs', 'BachFlixNfo/Features/MissingEpisodeWorkflow.cs', 'BachFlixNfo/Features/MissingEpisodeSheetRunner.cs', 'tests/MissingEpisodeSheetAdapterTests.cs') | ForEach-Object { Join-Path $repo $_ }
& $Compiler $arguments $sources
if ($LASTEXITCODE -ne 0) { throw 'Sheet adapter test compilation failed.' }
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'MissingEpisodeSheetAdapterTests.config') -Destination "$out/MissingEpisodeSheetAdapterTests.exe.config" -Force
& "$out/MissingEpisodeSheetAdapterTests.exe"
if ($LASTEXITCODE -ne 0) { throw 'Sheet adapter tests failed.' }
