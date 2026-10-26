param([Parameter(Mandatory=$true)][string]$Path)
$ErrorActionPreference = 'Stop'
$installer = New-Object -ComObject WindowsInstaller.Installer
$installer.UILevel = 2
$session = $null
$checks = 0
function Check($ok, $message) {
    if (-not $ok) { throw $message }
    $script:checks++
}
try {
    # Ignore installed product state; only explicitly invoke read-only search and
    # launch checks. Never run INSTALL, InstallInitialize, or InstallFinalize.
    $session = $installer.OpenPackage([IO.Path]::GetFullPath($Path), 1)
    $session.Property('Installed') = ''
    $session.Property('WIX_DOWNGRADE_DETECTED') = ''
    Check ($session.DoAction('AppSearch') -eq 1) 'Actual MSI AppSearch failed'
    $actual32 = $session.Property('NET48_32')
    $actual64 = $session.Property('NET48_64')
    Check (($actual32 -match '^#\d+$') -or ($actual64 -match '^#\d+$')) 'No actual DWORD Release found'
    Check ($session.DoAction('LaunchConditions') -eq 1) 'Actual machine prerequisites rejected'
    Write-Output "PASS: real AppSearch + LaunchConditions; NET32=$actual32 NET64=$actual64"
    $db = $installer.OpenDatabase([IO.Path]::GetFullPath($Path), 0)
    $view = $db.OpenView('SELECT `Condition` FROM `LaunchCondition`')
    $view.Execute()
    $conditions=@()
    while($r=$view.Fetch()) { $conditions += $r.StringData(1); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($r) }
    $view.Close()
    $net=@($conditions | Where-Object {$_ -match 'NET48_32'})[0]
    $vsto=@($conditions | Where-Object {$_ -match 'VSTO32'})[0]
    Check ([bool]$net -and [bool]$vsto) 'Required conditions missing from generated MSI'
    foreach($property in @('NET48_32','NET48_64')) {
        foreach($value in @(528039,528040,533320,533509)) {
            $session.Property('NET48_32')=''; $session.Property('NET48_64')=''
            $session.Property($property)="#$value"
            $session.Property('VSTO32')='#1'; $session.Property('VSTO64')=''
            $expected=if($value -ge 528040){1}else{0}
            Check ($session.EvaluateCondition($net) -eq $expected) "MSI condition: $property/$value"
            $action=$session.DoAction('LaunchConditions')
            Check (($action -eq 1) -eq ($expected -eq 1)) "MSI LaunchConditions action: $property/$value"
        }
    }
    $session.Property('NET48_32')='';$session.Property('NET48_64')=''
    Check ($session.EvaluateCondition($net) -eq 0) 'Missing .NET accepted'
    foreach($property in @('VSTO32','VSTO64')) {
        foreach($value in @('','#0','#1')) {
            $session.Property('VSTO32')='';$session.Property('VSTO64')=''
            $session.Property($property)=$value
            Check (($session.EvaluateCondition($vsto) -eq 1) -eq ($value -eq '#1')) "VSTO raw DWORD $property"
        }
    }
    $session.Property('Installed')='1'
    Check ($session.EvaluateCondition($net) -eq 1 -and $session.EvaluateCondition($vsto) -eq 1) 'Maintenance must not block removal'
    Write-Output "PASS: $checks prerequisite checks; actual MSI actions, no product installed"
} finally {
    if($view){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view)}
    if($db){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($db)}
    if($session){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($session)}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
}
