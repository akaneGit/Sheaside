[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repoRoot 'Sheaside/obj/ReleaseCandidate' }
$publishPath = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vsRoot = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vsRoot) { throw 'Visual Studio MSBuild was not found.' }
$builder = Join-Path $vsRoot 'MSBuild/Current/Bin/MSBuild.exe'
& $builder (Join-Path $repoRoot 'Sheaside/Sheaside.csproj') /t:Publish /p:Configuration=Release /p:Platform=AnyCPU /p:VisualStudioVersion=18.0 "/p:PublishDir=$publishPath" "/p:PublishUrl=$publishPath" /p:ApplicationVersion=1.0.0.0 /p:IsWebBootstrapper=false /v:minimal /clp:Summary /nologo
if ($LASTEXITCODE -ne 0) { throw 'Release Candidate publish failed.' }
foreach ($required in @('setup.exe', 'Sheaside.vsto', 'Application Files/Sheaside_1_0_0_0/Sheaside.dll.manifest', 'Application Files/Sheaside_1_0_0_0/Sheaside.dll.deploy')) {
    if (-not (Test-Path -LiteralPath (Join-Path $publishPath $required))) { throw "Missing publish artifact: $required" }
}
Write-Output "Local RC package: $publishPath"
Write-Output 'This uses the project signing certificate. No installation or external publication was performed.'
