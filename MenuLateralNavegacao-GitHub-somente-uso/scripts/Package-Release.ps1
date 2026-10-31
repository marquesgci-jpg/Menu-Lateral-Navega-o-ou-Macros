param(
    [string]$Version,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts')
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$deploymentRoot = Join-Path $repositoryRoot 'deployment'
$demoWorkbook = Join-Path $repositoryRoot 'demo\Preview_Funcional_MenuLateralNavegacao.xlsm'
$vbaModule = Join-Path $repositoryRoot 'demo\vba\DemoMenu.bas'
$rootManifest = Join-Path $deploymentRoot 'MenuLateralNavegacao.vsto'

foreach ($requiredPath in @($deploymentRoot, $demoWorkbook, $vbaModule, $rootManifest)) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required release file or folder is missing: $requiredPath"
    }
}

[xml]$manifest = Get-Content -LiteralPath $rootManifest -Raw
$deploymentIdentity = $manifest.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='assemblyIdentity']")
$deploymentVersion = $deploymentIdentity.GetAttribute('version')
$requestedVersion = $deploymentVersion
if (-not [string]::IsNullOrWhiteSpace($Version)) {
    $requestedVersion = $Version.TrimStart('v')
}
if ([string]::IsNullOrWhiteSpace($requestedVersion)) {
    $requestedVersion = $deploymentVersion
}
if ($requestedVersion -ne $deploymentVersion) {
    throw "Tag version $requestedVersion does not match signed deployment version $deploymentVersion. Publish a matching signed VSTO deployment first."
}

$applicationReference = $manifest.SelectSingleNode("//*[local-name()='dependentAssembly' and @dependencyType='install']")
if ($null -eq $applicationReference) {
    throw 'The VSTO deployment manifest has no application manifest reference.'
}

$applicationManifestPath = Join-Path $deploymentRoot $applicationReference.GetAttribute('codebase')
if (-not (Test-Path -LiteralPath $applicationManifestPath)) {
    throw "The referenced application manifest is missing: $applicationManifestPath"
}

$manifestBytes = [System.IO.File]::ReadAllBytes($applicationManifestPath)
if ($manifestBytes.Length -ne [int64]$applicationReference.GetAttribute('size')) {
    throw 'The application manifest size does not match the VSTO deployment manifest.'
}

$expectedDigestNode = $applicationReference.SelectSingleNode(".//*[local-name()='DigestValue']")
$sha1 = [System.Security.Cryptography.SHA1]::Create()
try {
    $actualDigest = [Convert]::ToBase64String($sha1.ComputeHash($manifestBytes))
}
finally {
    $sha1.Dispose()
}
if ($actualDigest -ne $expectedDigestNode.InnerText) {
    throw 'The application manifest SHA-1 does not match the signed VSTO deployment manifest.'
}

$applicationDirectory = Split-Path $applicationManifestPath -Parent
if (-not (Test-Path -LiteralPath (Join-Path $applicationDirectory 'MenuLateralNavegacao.dll.deploy'))) {
    throw 'The versioned deployment is missing MenuLateralNavegacao.dll.deploy.'
}

$isSourceRepository = Test-Path -LiteralPath (Join-Path $repositoryRoot 'src\MenuLateralNavegacao\MenuLateralNavegacao.csproj')
$flavor = if ($isSourceRepository) { 'com-fontes' } else { 'somente-uso' }
$archiveName = "MenuLateralNavegacao-v$requestedVersion-$flavor.zip"
$outputRoot = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null
$archivePath = Join-Path $outputRoot $archiveName
if (Test-Path -LiteralPath $archivePath) {
    throw "Release archive already exists; choose a fresh output directory: $archivePath"
}

$stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("MenuLateralRelease-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stagingRoot | Out-Null

function Copy-DirectoryContents([string]$Source, [string]$Destination) {
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $Source -File -Recurse) {
        $relativePath = $file.FullName.Substring($Source.Length).TrimStart('\')
        $destinationPath = Join-Path $Destination $relativePath
        New-Item -ItemType Directory -Force -Path (Split-Path $destinationPath -Parent) | Out-Null
        [System.IO.File]::Copy($file.FullName, $destinationPath, $true)
    }
}

try {
    Copy-DirectoryContents $deploymentRoot $stagingRoot
    [System.IO.File]::Copy($demoWorkbook, (Join-Path $stagingRoot 'Preview_Funcional_MenuLateralNavegacao.xlsm'))
    $vbaDestination = Join-Path $stagingRoot 'Fontes VBA\DemoMenu.bas'
    New-Item -ItemType Directory -Force -Path (Split-Path $vbaDestination -Parent) | Out-Null
    [System.IO.File]::Copy($vbaModule, $vbaDestination)
    [System.IO.File]::Copy((Join-Path $repositoryRoot 'README.md'), (Join-Path $stagingRoot 'LEIA-ME.md'))
    Copy-DirectoryContents (Join-Path $repositoryRoot 'docs') (Join-Path $stagingRoot 'docs')

    if ($isSourceRepository) {
        $sourceDestination = Join-Path $stagingRoot 'Fontes do Suplemento\src'
        Copy-DirectoryContents (Join-Path $repositoryRoot 'src') $sourceDestination

        $resourceRoot = Join-Path $repositoryRoot 'resources'
        if (Test-Path -LiteralPath $resourceRoot) {
            Copy-DirectoryContents $resourceRoot (Join-Path $stagingRoot 'Recursos')
        }

        $reconstructionGuide = Join-Path $repositoryRoot 'docs\RECONSTRUCAO-FONTES.md'
        if (Test-Path -LiteralPath $reconstructionGuide) {
            $sourceDocs = Join-Path $stagingRoot 'Fontes do Suplemento\docs'
            New-Item -ItemType Directory -Force -Path $sourceDocs | Out-Null
            [System.IO.File]::Copy($reconstructionGuide, (Join-Path $sourceDocs 'RECONSTRUCAO.md'))
        }
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
        $stagingRoot,
        $archivePath,
        [System.IO.Compression.CompressionLevel]::Optimal,
        $false
    )
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        Remove-Item -LiteralPath $stagingRoot -Recurse -Force
    }
}

Write-Output $archivePath
