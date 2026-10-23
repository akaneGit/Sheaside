[CmdletBinding()]
param([switch]$Excel, [switch]$BatchOnly)
$ErrorActionPreference = 'Stop'
$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$vsRoot = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
if (-not $vsRoot) { throw 'Visual Studio MSBuild was not found.' }
$compiler = Join-Path $vsRoot 'MSBuild/Current/Bin/Roslyn/csc.exe'
$output = Join-Path $repoRoot 'Sheaside/obj/SmokeTests'
[void][IO.Directory]::CreateDirectory($output)
$interop = Join-Path $repoRoot 'Sheaside/obj/Debug/Interop.Microsoft.Office.Interop.Excel.dll'
if (-not (Test-Path -LiteralPath $interop)) { throw 'Build the Debug configuration first.' }
$sources = @(
    'tests/SmokeTests.cs',
    'tests/ReorderTests.cs',
    'tests/AppearanceTests.cs',
    'tests/ChartTests.cs',
    'tests/SelectionTests.cs',
    'tests/BatchTests.cs',
    'Sheaside/Models/WorksheetItem.cs',
    'Sheaside/Models/SheetSelectionState.cs',
    'Sheaside/Models/SheetBatch.cs',
    'Sheaside/Services/SheetBatchRunner.cs',
    'Sheaside/Services/ExcelSheetBatchService.cs',
    'Sheaside/Controls/BatchDetailsDialog.cs',
    'Sheaside/Models/WorksheetDragEventArgs.cs',
    'Sheaside/Controls/ReorderableSheetListBox.cs',
    'Sheaside/Services/WorksheetMoveRules.cs',
    'Sheaside/Controls/SheetListControl.cs',
    'Sheaside/Controls/RenameSheetDialog.cs',
    'Sheaside/Services/WorksheetNameRules.cs',
    'Sheaside/Services/ExcelWorksheetService.cs'
    'Sheaside/Services/ExcelSheet.cs'
) | ForEach-Object { Join-Path $repoRoot $_ }
$executable = Join-Path $output 'SmokeTests.exe'
& $compiler /nologo /target:exe /debug:full /langversion:7.3 "/out:$executable" "/link:$interop" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:Microsoft.CSharp.dll @sources
if ($LASTEXITCODE -ne 0) { throw 'Smoke test compilation failed.' }
$testArguments = @()
if ($Excel) { $testArguments += '--excel' }
if ($BatchOnly) { $testArguments += '--batch-only' }
& $executable @testArguments
if ($LASTEXITCODE -ne 0) { throw 'Smoke tests failed.' }
