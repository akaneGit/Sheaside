[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$Path, [string]$ExpectedVersion = '0.99.6')
$ErrorActionPreference = 'Stop'
$installer = New-Object -ComObject WindowsInstaller.Installer
$database = $installer.OpenDatabase([IO.Path]::GetFullPath($Path), 0)
$script:checks = 0
function Assert-Msi($condition, $message) {
    if (-not $condition) { throw "MSI: $message" }
    $script:checks++
}
function Read-Rows([string]$query, [int]$columns) {
    $view = $database.OpenView($query)
    try {
        [void]$view.Execute()
        while ($record = $view.Fetch()) {
            try { $row = @(); for ($i=1; $i -le $columns; $i++) { $row += $record.StringData($i) }; ,$row }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($record) }
        }
    } finally { [void]$view.Close(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view) }
}
try {
    $properties = @{}; Read-Rows 'SELECT `Property`, `Value` FROM `Property`' 2 | ForEach-Object { $properties[$_[0]]=$_[1] }
    Assert-Msi ($properties.ProductVersion -eq $ExpectedVersion) 'numeric RC version'
    if ($ExpectedVersion -eq '0.99.6') { Assert-Msi ($properties.ProductCode -ne '{5E6A06B4-0794-49AD-8F9B-5084BBD8C2E2}') 'new ProductCode differs from rc.2' }
    Assert-Msi (-not $properties.ContainsKey('ALLUSERS')) 'per-user installation'
    Assert-Msi ($properties.UpgradeCode -eq '{E2087C97-A6B5-4089-B8F5-86DB64EBB425}') 'stable upgrade family'
    $files = @(Read-Rows 'SELECT `FileName` FROM `File`' 1 | ForEach-Object { $_[0].Split('|')[-1] })
    Assert-Msi ($files.Count -eq 4) 'only four runtime payload files'
    foreach ($name in @('Sheaside.dll','Sheaside.dll.manifest','Sheaside.vsto','Microsoft.Office.Tools.Common.v4.0.Utilities.dll')) { Assert-Msi ($files -contains $name) "payload $name" }
    $registry = @(Read-Rows 'SELECT `Root`, `Key`, `Name`, `Value` FROM `Registry`' 4)
    Assert-Msi (@($registry | Where-Object { $_[0] -ne '1' }).Count -eq 0) 'only HKCU registration'
    Assert-Msi (@($registry | Where-Object { $_[2] -eq 'Manifest' -and $_[3] -eq 'file:///[INSTALLFOLDER]Sheaside.vsto|vstolocal' }).Count -eq 1) 'local manifest, no ClickOnce cache'
    Assert-Msi (@($registry | Where-Object { $_[2] -eq 'LoadBehavior' -and $_[3] -eq '#3' }).Count -eq 1) 'load on Excel startup'
    Assert-Msi (@(Read-Rows 'SELECT `UpgradeCode` FROM `Upgrade`' 1).Count -ge 2) 'upgrade and downgrade detection'
    $upgrades = @(Read-Rows 'SELECT `VersionMin`, `VersionMax`, `Attributes`, `ActionProperty` FROM `Upgrade`' 4)
    Assert-Msi (@($upgrades | Where-Object { $_[1] -eq $ExpectedVersion -and $_[3] -eq 'WIX_UPGRADE_DETECTED' }).Count -eq 1) 'older versions detected below current version'
    Assert-Msi (@($upgrades | Where-Object { $_[0] -eq $ExpectedVersion -and $_[3] -eq 'WIX_DOWNGRADE_DETECTED' }).Count -eq 1) 'newer versions detected for downgrade block'
    Assert-Msi (@(Read-Rows 'SELECT `Action` FROM `InstallExecuteSequence`' 1 | Where-Object { $_[0] -eq 'RemoveExistingProducts' }).Count -eq 1) 'major upgrade removes detected old product'
    Assert-Msi (@(Read-Rows 'SELECT `Condition` FROM `LaunchCondition`' 1).Count -ge 3) 'runtime and downgrade prerequisites'
    if ($ExpectedVersion -eq '0.99.6') {
        Assert-Msi ($properties.ProductCode -ne '{834E3417-08C4-4078-A73B-6A05AF410909}') 'new ProductCode differs from rc.5'
        $locators = @(Read-Rows 'SELECT `Signature_`, `Root`, `Key`, `Name`, `Type` FROM `RegLocator`' 5)
        $searches = @(Read-Rows 'SELECT `Property`, `Signature_` FROM `AppSearch`' 2)
        foreach ($entry in @(@('Net48_32','NET48_32','2'),@('Net48_64','NET48_64','18'))) {
            Assert-Msi (@($locators | Where-Object { $_[0] -eq $entry[0] -and $_[1] -eq '2' -and $_[2] -eq 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -and $_[3] -eq 'Release' -and $_[4] -eq $entry[2] }).Count -eq 1) 'correct .NET registry view/key/type'
            Assert-Msi (@($searches | Where-Object { $_[0] -eq $entry[1] -and $_[1] -eq $entry[0] }).Count -eq 1) 'registry result feeds condition property'
        }
        $conditions = @(Read-Rows 'SELECT `Condition` FROM `LaunchCondition`' 1 | ForEach-Object {$_[0]})
        Assert-Msi ($conditions -contains 'Installed OR NET48_32 >= "#528040" OR NET48_64 >= "#528040"') '.NET raw DWORD comparison'
        Assert-Msi ($conditions -contains 'Installed OR VSTO32 = "#1" OR VSTO64 = "#1"') 'VSTO raw DWORD comparison'
        foreach ($table in @('InstallUISequence','InstallExecuteSequence')) {
            $sequence = @{}; Read-Rows ('SELECT `Action`, `Sequence` FROM `' + $table + '`') 2 | ForEach-Object {$sequence[$_[0]]=[int]$_[1]}
            Assert-Msi ($sequence.ContainsKey('AppSearch') -and $sequence.ContainsKey('LaunchConditions') -and $sequence.AppSearch -lt $sequence.LaunchConditions) 'AppSearch precedes LaunchConditions'
        }
    }
    Assert-Msi (@(Read-Rows 'SELECT `Cabinet` FROM `Media`' 1 | Where-Object { $_[0].StartsWith('#') }).Count -eq 1) 'embedded cabinet for single MSI distribution'
    Write-Output "PASS: $script:checks MSI assertions (read-only; install/repair/uninstall not executed)"
} finally {
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($database)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
}
