$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vsRoot = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vsRoot) { throw 'Visual Studio MSBuild was not found.' }
$builder = Join-Path $vsRoot 'MSBuild/Current/Bin/MSBuild.exe'
foreach ($configuration in @('Release', 'Debug')) {
    & $builder (Join-Path $repoRoot 'Sheaside.slnx') /t:Build "/p:Configuration=$configuration" '/p:Platform=Any CPU' /p:VisualStudioVersion=18.0 /v:minimal /clp:Summary /nologo
    if ($LASTEXITCODE -ne 0) { throw "$configuration build failed." }
}
& (Join-Path $PSScriptRoot 'Test-Sheaside.ps1')
& (Join-Path $PSScriptRoot 'Test-Sheaside.ps1') -Excel
