<#
.SYNOPSIS
Creates a PowerShell Gallery package with modern license metadata.

.DESCRIPTION
Compresses the module into a NuGet package and updates its nuspec with
NuGet's file-based license metadata. The repository's LICENSE file is also
included in the package.

This is needed because Compress-PSResource otherwise generates the older
licenseUrl metadata, which NuGet has deprecated. This script replaces that
metadata with a reference to the license file.

.PARAMETER ModulePath
Path to the module directory to package.

.PARAMETER DestinationPath
Directory where the .nupkg package is saved. The directory is created if needed.

.OUTPUTS
Path to the created .nupkg package.

.EXAMPLE
$packagePath = ./tools/New-ModulePackage.ps1 `
    -ModulePath ./Office365-Scoutnet-synk `
    -DestinationPath $env:RUNNER_TEMP
#>
param (
    [Parameter(Mandatory)]
    [string]$ModulePath,

    [Parameter(Mandatory)]
    [string]$DestinationPath
)

$ErrorActionPreference = 'Stop'

Import-Module Microsoft.PowerShell.PSResourceGet -ErrorAction Stop
Add-Type -AssemblyName System.IO.Compression

$resolvedModulePath = (Resolve-Path -LiteralPath $ModulePath).Path
$resolvedDestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath)
$licenseFilePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\LICENSE')).Path
New-Item -ItemType Directory -Path $resolvedDestinationPath -Force | Out-Null

$package = Compress-PSResource -Path $resolvedModulePath -DestinationPath $resolvedDestinationPath -PassThru -ErrorAction Stop
$archive = [System.IO.Compression.ZipFile]::Open($package.FullName, [System.IO.Compression.ZipArchiveMode]::Update)

try {
    $nuspecEntries = @($archive.Entries | Where-Object FullName -Like '*.nuspec')
    if ($nuspecEntries.Count -ne 1) {
        throw "Expected one nuspec in package '$($package.FullName)', found $($nuspecEntries.Count)."
    }

    $reader = [System.IO.StreamReader]::new($nuspecEntries[0].Open())
    try {
        $nuspec = [System.Xml.XmlDocument]::new()
        $nuspec.LoadXml($reader.ReadToEnd())
    }
    finally {
        $reader.Dispose()
    }

    $metadata = $nuspec.SelectSingleNode('/*[local-name()="package"]/*[local-name()="metadata"]')
    if (-not $metadata) {
        throw "Package '$($package.FullName)' has no nuspec metadata element."
    }

    $oldLicenseUrl = $metadata.SelectSingleNode('./*[local-name()="licenseUrl"]')
    $licenseAnchor = if ($oldLicenseUrl) {
        $nextElement = $oldLicenseUrl.NextSibling
        while ($nextElement -and $nextElement.NodeType -ne [System.Xml.XmlNodeType]::Element) {
            $nextElement = $nextElement.NextSibling
        }
        $metadata.RemoveChild($oldLicenseUrl) | Out-Null
        $nextElement
    }
    else {
        $metadata.SelectSingleNode('./*[local-name()="projectUrl"]')
    }

    $license = $nuspec.CreateElement('license', $metadata.NamespaceURI)
    $license.SetAttribute('type', 'file')
    $license.InnerText = 'LICENSE'

    if ($licenseAnchor) {
        $metadata.InsertBefore($license, $licenseAnchor) | Out-Null
    }
    else {
        $metadata.AppendChild($license) | Out-Null
    }

    $nuspecEntryName = $nuspecEntries[0].FullName
    $nuspecEntries[0].Delete()
    $newNuspecEntry = $archive.CreateEntry($nuspecEntryName)
    $writer = [System.IO.StreamWriter]::new($newNuspecEntry.Open(), [System.Text.UTF8Encoding]::new($false))
    try {
        $nuspec.Save($writer)
    }
    finally {
        $writer.Dispose()
    }

    $existingLicenseEntry = $archive.GetEntry('LICENSE')
    if ($existingLicenseEntry) {
        $existingLicenseEntry.Delete()
    }

    $licenseEntry = $archive.CreateEntry('LICENSE')
    $licenseInput = [System.IO.File]::OpenRead($licenseFilePath)
    $licenseOutput = $licenseEntry.Open()
    try {
        $licenseInput.CopyTo($licenseOutput)
    }
    finally {
        $licenseOutput.Dispose()
        $licenseInput.Dispose()
    }
}
finally {
    $archive.Dispose()
}

$package.FullName
