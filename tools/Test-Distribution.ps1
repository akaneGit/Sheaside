[CmdletBinding()]
param([string]$Directory)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $Directory) { $Directory = Join-Path $root 'artifacts/rc.6' }
$payload = Join-Path $root 'Sheaside/bin/Release'
$script:checks = 0
function Assert-Package($condition, $message) {
    if (-not $condition) { throw $message }
    $script:checks++
}
$version = [Diagnostics.FileVersionInfo]::GetVersionInfo((Join-Path $payload 'Sheaside.dll'))
Assert-Package ($version.FileVersion -eq '1.0.0.6' -and $version.ProductVersion -eq '1.0.0-rc.6') 'Wrong application version'
$identity = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $payload 'Sheaside.dll'))
Assert-Package ($identity.Version.ToString() -eq '1.0.0.0' -and $identity.ProcessorArchitecture -eq 'MSIL') 'Wrong assembly identity/architecture'
$assembly = [Reflection.Assembly]::ReflectionOnlyLoadFrom((Join-Path $payload 'Sheaside.dll'))
Assert-Package ($assembly.GetManifestResourceNames() -contains 'Sheaside.SheasideRibbon.xml') 'Embedded ribbon missing'
foreach ($file in @('Sheaside.dll','Sheaside.dll.manifest','Sheaside.vsto','Microsoft.Office.Tools.Common.v4.0.Utilities.dll')) {
    $bytes = [IO.File]::ReadAllBytes((Join-Path $payload $file))
    foreach ($encoding in @([Text.Encoding]::UTF8,[Text.Encoding]::Unicode)) {
        $text = $encoding.GetString($bytes)
        Assert-Package ($text -notmatch '(?i)[A-Z]:[\\/]+Users[\\/]+') 'Developer-specific user path in distribution payload (value redacted)'
        Assert-Package ($text -notmatch '-----BEGIN (?:RSA |EC |OPENSSH |ENCRYPTED )?PRIVATE KEY-----|gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{30,}') 'Secret-like content in distribution payload (value redacted)'
    }
}
foreach ($name in @('Sheaside.vsto','Sheaside.dll.manifest')) {
    [xml]$xml = [IO.File]::ReadAllText((Join-Path $payload $name))
    foreach ($dep in $xml.SelectNodes('//*[local-name()="dependentAssembly" and @codebase]')) {
        $path = [IO.Path]::GetFullPath((Join-Path $payload $dep.codebase))
        Assert-Package ($path.StartsWith($payload + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $path)) 'Manifest reference escapes or is missing'
        $hash = $dep.SelectSingleNode('.//*[local-name()="DigestValue"]').InnerText
        $sha = [Security.Cryptography.SHA256]::Create()
        try { $actual = [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($path))) } finally { $sha.Dispose() }
        Assert-Package ($hash -eq $actual) 'Manifest dependency hash mismatch'
    }
    $node = $xml.SelectSingleNode('//*[local-name()="X509Certificate"]')
    $cert = New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (,[Convert]::FromBase64String($node.InnerText))
    Assert-Package ($cert.Thumbprint -eq 'C61660B9D07850DE0D156C9127355F4124B5093A' -and -not $cert.HasPrivateKey) 'Unexpected manifest certificate'
}
$expanded = Join-Path $root ('Sheaside/obj/zip-check-' + [guid]::NewGuid().ToString('N'))
Expand-Archive -LiteralPath (Join-Path $Directory 'Sheaside-1.0.0-rc.6.zip') -DestinationPath $expanded
$expected = @('Sheaside-1.0.0-rc.6.msi','INSTALL.md','CHECKLIST.md','MSI-SHA256.txt')
$files = @(Get-ChildItem -LiteralPath $expanded -File -Recurse)
Assert-Package ($files.Count -eq 4) 'Unexpected ZIP payload'
foreach ($name in $expected) {
    $path = Join-Path $expanded $name
    Assert-Package (Test-Path -LiteralPath $path) 'Missing ZIP file'
    Assert-Package ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath (Join-Path $Directory $name) -Algorithm SHA256).Hash) 'ZIP extracted hash mismatch'
}
foreach ($line in [IO.File]::ReadAllLines((Join-Path $Directory 'SHA256SUMS.txt'))) {
    $parts = $line -split '  ', 2
    Assert-Package ((Get-FileHash -LiteralPath (Join-Path $Directory $parts[1]) -Algorithm SHA256).Hash -eq $parts[0]) 'Published checksum mismatch'
}
Assert-Package ((Get-FileHash -LiteralPath (Join-Path $root 'artifacts/Sheaside-1.0.0-rc.2.msi')).Hash -eq '5BD7B5BB93596FBFA9E8DE0D678CAF01269A7397C038B36C2C55EF37D3F8ABFD') 'Original rc.2 MSI changed'
Assert-Package ((Get-FileHash -LiteralPath (Join-Path $root 'artifacts/Sheaside-1.0.0-rc.2.zip')).Hash -eq '7CAB5FA1B0AED9601D135D8CB783222542E16385052F251A096E27542863DC6B') 'Original rc.2 ZIP changed'
Assert-Package ((Get-FileHash -LiteralPath (Join-Path $root 'artifacts/rc.5/Sheaside-1.0.0-rc.5.msi')).Hash -eq '33942038BCA5B104A3AFF244D3EFE3C0037957720E2729FD94CD41B5C44D98F0') 'Published rc.5 MSI changed'
Assert-Package ((Get-FileHash -LiteralPath (Join-Path $root 'artifacts/rc.5/Sheaside-1.0.0-rc.5.zip')).Hash -eq '0BB09DDA87108529EC91F91732CA07A8F87ABD960B126F43731478CB7CF133CA') 'Published rc.5 ZIP changed'
Write-Output "PASS: $script:checks distribution assertions (no installation performed)"
