[CmdletBinding()]
param([switch]$SkipBuild)
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$payload = Join-Path $repoRoot 'Sheaside/bin/Release'
if (-not $SkipBuild) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    $vsRoot = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
    if (-not $vsRoot) { throw 'Visual Studio MSBuild was not found.' }
    & (Join-Path $vsRoot 'MSBuild/Current/Bin/MSBuild.exe') (Join-Path $repoRoot 'Sheaside.slnx') /t:Build /p:Configuration=Release '/p:Platform=Any CPU' /p:VisualStudioVersion=18.0 /v:minimal /nologo
    if ($LASTEXITCODE -ne 0) { throw 'Release build failed.' }
}
foreach ($name in @('Sheaside.dll','Sheaside.dll.manifest','Sheaside.vsto','Microsoft.Office.Tools.Common.v4.0.Utilities.dll')) {
    if (-not (Test-Path -LiteralPath (Join-Path $payload $name))) { throw "Missing payload: $name" }
}
$toolDir = Join-Path $repoRoot 'Sheaside/obj/wix'
$wix = Join-Path $toolDir 'wix.exe'
if (-not (Test-Path -LiteralPath $wix)) {
    & dotnet tool install wix --version 4.0.6 --tool-path $toolDir --allow-roll-forward
    if ($LASTEXITCODE -ne 0) { throw 'WiX 4.0.6 installation failed (.NET SDK and NuGet access required).' }
}
$artifacts = Join-Path $repoRoot 'artifacts'
[void][IO.Directory]::CreateDirectory($artifacts)
$msi = Join-Path $artifacts 'Sheaside-1.0.0-rc.2.msi'
& $wix build (Join-Path $repoRoot 'installer/Sheaside.wxs') -arch x86 -d "Payload=$payload" -pdbtype none -out $msi
if ($LASTEXITCODE -ne 0) { throw 'MSI compilation or ICE validation failed.' }
& (Join-Path $PSScriptRoot 'Test-Msi.ps1') -Path $msi
$extracted = Join-Path $repoRoot ('Sheaside/obj/msi-check-' + [Guid]::NewGuid().ToString('N'))
& $wix msi decompile $msi -out (Join-Path $extracted 'package.wxs') -x (Join-Path $extracted 'files')
if ($LASTEXITCODE -ne 0) { throw 'MSI payload extraction failed.' }
$fileMap = @{ AddinDll='Sheaside.dll'; ApplicationManifest='Sheaside.dll.manifest'; DeploymentManifest='Sheaside.vsto'; Utilities='Microsoft.Office.Tools.Common.v4.0.Utilities.dll' }
foreach ($id in $fileMap.Keys) {
    $actual = (Get-FileHash -LiteralPath (Join-Path $extracted "files/File/$id") -Algorithm SHA256).Hash
    $expected = (Get-FileHash -LiteralPath (Join-Path $payload $fileMap[$id]) -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw "MSI payload hash mismatch: $id" }
}
Write-Output 'PASS: 4 extracted MSI payload hashes match Release output'
Copy-Item -LiteralPath (Join-Path $repoRoot 'docs/deployment.md') -Destination (Join-Path $artifacts 'INSTALL.md') -Force
$hash = (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash
[IO.File]::WriteAllText((Join-Path $artifacts 'SHA256SUMS.txt'), "$hash  Sheaside-1.0.0-rc.2.msi`r`n")
Compress-Archive -LiteralPath $msi,(Join-Path $artifacts 'INSTALL.md'),(Join-Path $artifacts 'SHA256SUMS.txt') -DestinationPath (Join-Path $artifacts 'Sheaside-1.0.0-rc.2.zip') -Force
Write-Output "Created: $msi"
