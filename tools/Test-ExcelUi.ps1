# Operates only on a new unsaved Excel instance. Requires Debug add-in registration.
# No SendKeys or global mouse input. Visibility preferences are restored in finally.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
if (-not ('SheasideUiProbe' -as [type])) {
 Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class SheasideUiProbe {
 [DllImport("user32.dll", EntryPoint="SendMessageTimeoutW", CharSet=CharSet.Unicode, SetLastError=true)] public static extern IntPtr SendText(IntPtr hwnd, uint msg, IntPtr w, string text, uint flags, uint timeout, out IntPtr result);
 [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
}
"@
}
$probeExcel=$null; $probeBook=$null; $checks=0
$prefPath=Join-Path $env:LOCALAPPDATA 'Sheaside/settings.xml'
$prefExists=Test-Path -LiteralPath $prefPath
$prefBytes=if($prefExists){[IO.File]::ReadAllBytes($prefPath)}else{$null}
function Assert($condition,$name) { if(-not $condition){throw "FAIL: $name"}; $script:checks++; Write-Output "PASS: $name" }
function Elements {
 $script:root=[System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$probeExcel.Hwnd)
 $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,[System.Windows.Automation.Condition]::TrueCondition)
}
function Named($name) { Elements | Where-Object { $_.Current.Name -eq $name -and -not $_.Current.IsOffscreen } | Select-Object -First 1 }
function PaneButtons { Elements | Where-Object { $_.Current.ClassName -like 'WindowsForms10.BUTTON*' -and -not $_.Current.IsOffscreen } }
function Select-View { $tab=Elements | Where-Object { $_.Current.Name -eq '表示' -and $_.Current.ControlType -eq [System.Windows.Automation.ControlType]::TabItem } | Select-Object -First 1; $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep -Milliseconds 200 }
function Toggle($name) { (Named $name).GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle(); Start-Sleep -Milliseconds 500 }
function Is-On($name) { (Named $name).GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On }
function Set-Toggle($name,$value) { if((Is-On $name) -ne $value){Toggle $name} }
function Click-Button($button) {
 $result=[IntPtr]::Zero
 if(-not $button -or [SheasideUiProbe]::SendMessageTimeout([IntPtr]$button.Current.NativeWindowHandle,0xF5,[IntPtr]::Zero,[IntPtr]::Zero,2,2000,[ref]$result) -eq [IntPtr]::Zero){throw 'Pane click timed out'}
 Start-Sleep -Milliseconds 200
}
try {
 $probeExcel=New-Object -ComObject Excel.Application
 $probeExcel.Visible=$true
 $probeBook=$probeExcel.Workbooks.Add(-4167)
 $visible=$probeBook.Worksheets.Item(1); $visible.Name='VisibleProbe'
 $hidden=$probeBook.Worksheets.Add(); $hidden.Name='OrdinaryHiddenProbe'; $hidden.Visible=0
 $very=$probeBook.Worksheets.Add(); $very.Name='VeryHiddenProbe'; $very.Visible=2
 $visible.Activate()
 Assert $probeExcel.COMAddIns.Item('Sheaside').Connect 'VSTO connected in Excel'
 Select-View
 Assert ((Named 'Sheaside').Current.ControlType -eq [System.Windows.Automation.ControlType]::Group) 'Sheaside group on built-in View tab'
 Assert (-not (Is-On 'シート一覧') -and @(PaneButtons).Count -eq 0) 'new workbook pane starts hidden'
 Toggle 'シート一覧'
 $add=PaneButtons | Where-Object {$_.Current.Name -eq '追加'} | Select-Object -First 1
 $delete=PaneButtons | Where-Object {$_.Current.Name -eq '削除'} | Select-Object -First 1
 Assert ($add -and $delete -and $add.Current.BoundingRectangle.Height -eq $delete.Current.BoundingRectangle.Height) 'ribbon opens aligned panel actions'
 Assert (@(PaneButtons).Count -eq 2) 'obsolete filter and settings controls absent'
 Set-Toggle '非表示も表示' $false; Set-Toggle 'VeryHiddenも表示' $false
 Assert ([bool](Named '1件選択中 / 1件表示')) 'default visible-only rows'
 Toggle '非表示も表示'
 Assert ((Is-On '非表示も表示') -and -not (Is-On 'VeryHiddenも表示') -and [bool](Named '1件選択中 / 2件表示')) 'Hidden switch independent and reflected in rows'
 Toggle 'VeryHiddenも表示'
 Assert ([bool](Named '1件選択中 / 3件表示')) 'both switches include all visibility states'
 $edit=Elements | Where-Object {$_.Current.ClassName -like 'WindowsForms10.EDIT*'} | Select-Object -First 1
 $textResult=[IntPtr]::Zero
 if([SheasideUiProbe]::SendText([IntPtr]$edit.Current.NativeWindowHandle,0xC,[IntPtr]::Zero,'VeryHiddenProbe',2,2000,[ref]$textResult) -eq [IntPtr]::Zero){throw 'Search input timed out'}
 Start-Sleep -Milliseconds 300
 Assert ([bool](Named '解除') -and [bool](Named '0件選択中 / 1件表示')) 'search filters rows and exposes clear'
 foreach($hiddenName in @('OrdinaryHiddenProbe','VeryHiddenProbe')) {
  [void][SheasideUiProbe]::SendText([IntPtr]$edit.Current.NativeWindowHandle,0xC,[IntPtr]::Zero,$hiddenName,2,2000,[ref]$textResult)
  Start-Sleep -Milliseconds 200
  $list=Elements | Where-Object {$_.Current.ClassName -like 'WindowsForms10.LISTBOX*'} | Select-Object -First 1
  $clickResult=[IntPtr]::Zero
  foreach($message in @(0x201,0x202)) {
   if([SheasideUiProbe]::SendMessageTimeout([IntPtr]$list.Current.NativeWindowHandle,$message,[IntPtr]::Zero,[IntPtr](8 -bor (8 -shl 16)),2,2000,[ref]$clickResult) -eq [IntPtr]::Zero){throw 'Hidden selection blocked by a modal notification'}
  }
  Assert ($probeExcel.ActiveSheet.Name -eq 'VisibleProbe' -and [bool](Named '1件選択中 / 1件表示')) "silent local selection without activation: $hiddenName"
 }
 Click-Button (PaneButtons | Where-Object {$_.Current.Name -eq '解除'} | Select-Object -First 1)
 Assert (-not [bool](Named '解除') -and (Is-On '非表示も表示') -and (Is-On 'VeryHiddenも表示')) 'search clear preserves visibility switches'
 Toggle 'シート一覧'
 Assert (-not (Is-On 'シート一覧') -and @(PaneButtons).Count -eq 0) 'ribbon hides pane and state matches'
 Toggle 'シート一覧'
 $first=$probeBook.Windows.Item(1)
 $second=$probeBook.NewWindow()
 Select-View
 Assert (-not (Is-On 'シート一覧') -and @(PaneButtons).Count -eq 0) 'second window starts hidden independently'
 $first.Activate(); Select-View
 Assert ((Is-On 'シート一覧') -and @(PaneButtons).Count -eq 2) 'return to first window preserves open pane'
 Write-Output "PASS: $checks Excel $($probeExcel.Version) host UI checks"
} finally {
 if($probeBook){$probeBook.Close($false)}
 if($probeExcel){$probeExcel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($probeExcel)}
 if($prefExists){[IO.File]::WriteAllBytes($prefPath,$prefBytes)}elseif(Test-Path -LiteralPath $prefPath){Remove-Item -LiteralPath $prefPath}
}