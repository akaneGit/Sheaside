# Creates a local-only ClickOnce code-signing certificate without changing trust stores.
# Run from Windows PowerShell 5.1 with .NET Framework 4.8 installed.
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$certificatePath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../Sheaside/Sheaside_TemporaryKey.pfx'))
if (Test-Path -LiteralPath $certificatePath) {
    throw "Certificate already exists: $certificatePath. Use the Visual Studio Signing page to replace it if necessary."
}

$rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider(2048)
$rsa.PersistKeyInCsp = $false
$certificate = $null
try {
    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        'CN=Sheaside Local Development', $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $true))
    $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
        [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature, $true))
    $usages = [System.Security.Cryptography.OidCollection]::new()
    [void]$usages.Add([System.Security.Cryptography.Oid]::new('1.3.6.1.5.5.7.3.3'))
    $request.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]::new($usages, $false))
    $certificate = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-5), [DateTimeOffset]::UtcNow.AddYears(1))
    [IO.File]::WriteAllBytes($certificatePath, $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, ''))
    $projectPath = Join-Path $PSScriptRoot '../Sheaside/Sheaside.csproj'
    $project = [System.Xml.XmlDocument]::new()
    $project.PreserveWhitespace = $true
    $project.Load($projectPath)
    $namespaces = [System.Xml.XmlNamespaceManager]::new($project.NameTable)
    $namespaces.AddNamespace('msbuild', $project.DocumentElement.NamespaceURI)
    $thumbprint = $project.SelectSingleNode('//msbuild:ManifestCertificateThumbprint', $namespaces)
    $thumbprint.InnerText = $certificate.Thumbprint
    $project.Save($projectPath)
    Write-Output "Created local development certificate: $certificatePath"
}
finally {
    if ($null -ne $certificate) { $certificate.Dispose() }
    $rsa.Dispose()
}
