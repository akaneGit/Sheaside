# Opens only a new, unsaved Excel instance. Requires the development add-in registration
# and StartupVisible=true. Does not alter user preferences, certificates, or existing workbooks.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
if (-not ('SheasideUiProbe' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SheasideUiProbe {
 [DllImport("user32.dll", SetLastError=true)] public static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
}
'@
}
$probeExcel = $null; $probeBook = $null
try {
    $probeExcel = New-Object -ComObject Excel.Application
    $probeExcel.Visible = $true
    $probeBook = $probeExcel.Workbooks.Add()
    if (-not $probeExcel.COMAddIns.Item('Sheaside').Connect) { throw 'Sheaside is not connected in Excel.' }
    $root = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$probeExcel.Hwnd)
    function Get-PaneButtons {
        $root.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition) |
            Where-Object { $_.Current.ClassName -like 'WindowsForms10.BUTTON*' -and -not $_.Current.IsOffscreen }
    }
    function Click-PaneButton($button) {
        $result = [IntPtr]::Zero
        if (-not $button -or [SheasideUiProbe]::SendMessageTimeout([IntPtr]$button.Current.NativeWindowHandle, 0xF5, [IntPtr]::Zero, [IntPtr]::Zero, 2, 2000, [ref]$result) -eq [IntPtr]::Zero) { throw 'Pane button click failed or timed out.' }
    }
    $buttons = @(Get-PaneButtons)
    $add = $buttons | Where-Object { $_.Current.Name -eq '追加' } | Select-Object -First 1
    $delete = $buttons | Where-Object { $_.Current.Name -eq '削除' } | Select-Object -First 1
    if (-not $add -or -not $delete -or $add.Current.BoundingRectangle.Height -ne $delete.Current.BoundingRectangle.Height) { throw 'Action buttons missing or not aligned.' }
    $filter = $buttons | Where-Object { $_.Current.Name -eq '' } | Select-Object -First 1
    Click-PaneButton $filter
    $state = Get-PaneButtons | Where-Object { $_.Current.Name -eq '表示' } | Select-Object -First 1
    if (-not $state) { throw 'Popup did not open.' }
    Click-PaneButton $state
    $clear = Get-PaneButtons | Where-Object { $_.Current.Name -eq '解除' } | Select-Object -First 1
    if (-not $clear) { throw 'Applied facet has no external reset button.' }
    Click-PaneButton $clear
    if (@(Get-PaneButtons | Where-Object { $_.Current.Name -eq '解除' }).Count) { throw 'Reset button still visible after clearing facets.' }
    Click-PaneButton $filter
    Click-PaneButton $filter
    if (@(Get-PaneButtons | Where-Object { $_.Current.Name -eq '表示' }).Count) { throw 'Popup did not close on repeated filter click.' }
    Write-Output "PASS: 6 Excel $($probeExcel.Version) host UI checks (connection, aligned actions, popup open, facet apply, reset, toggle close)"
} finally {
    if ($probeBook) { $probeBook.Close($false) }
    if ($probeExcel) { $probeExcel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($probeExcel) }
}
