[CmdletBinding()]
param([string]$RepositoryRoot = (Split-Path $PSScriptRoot))
$ErrorActionPreference = 'Stop'

# Read configuration only. No Visual Studio, MSI execution or registry access.
function Read-InstallerNodes([string]$Path) {
    $stack = [Collections.Generic.List[object]]::new()
    $nodes = [Collections.Generic.List[object]]::new()
    $pendingName = $null
    $lineNumber = 0
    foreach ($line in [IO.File]::ReadAllLines($Path)) {
        $lineNumber++
        $text = $line.Trim()
        if (-not $text) { continue }
        if ($text -match '^"([^"]+)"$') {
            if ($null -ne $pendingName) { throw "Missing opening brace at ${Path}:$lineNumber" }
            $pendingName = $Matches[1]
        } elseif ($text -eq '{') {
            if ($null -eq $pendingName) { throw "Unnamed block at ${Path}:$lineNumber" }
            $parent = ($stack | ForEach-Object { $_.Name }) -join '/'
            $node = [pscustomobject]@{ Name=$pendingName; Parent=$parent; Properties=@{} }
            $stack.Add($node)
            $nodes.Add($node)
            $pendingName = $null
        } elseif ($text -eq '}') {
            if ($null -ne $pendingName -or $stack.Count -eq 0) { throw "Unexpected closing brace at ${Path}:$lineNumber" }
            $stack.RemoveAt($stack.Count - 1)
        } elseif ($text -match '^"([^"]+)"\s*=\s*"(.*)"$') {
            if ($null -ne $pendingName -or $stack.Count -eq 0) { throw "Property outside a block at ${Path}:$lineNumber" }
            $stack[$stack.Count - 1].Properties[$Matches[1]] = $Matches[2]
        } else { throw "Unrecognized installer syntax at ${Path}:$lineNumber" }
    }
    if ($stack.Count -ne 0 -or $null -ne $pendingName) { throw "Unclosed block in $Path" }
    return $nodes
}

foreach ($relativePath in @('OutlookFileDragSetup/OutlookFileDragSetup.vdproj', 'OutlookFileDragSetup_x64/OutlookFileDragSetup_x64.vdproj')) {
    $nodes = @(Read-InstallerNodes (Join-Path $RepositoryRoot $relativePath))
    # These release setups use explicit File blocks only; output groups can silently
    # change the payload and are therefore rejected.
    if (@($nodes | Where-Object { $_.Parent -eq 'DeployProject/Deployable/ProjectOutput' -or $_.Properties['OutputGroupCanonicalName'] }).Count -ne 0) {
        throw "Unexpected project output group in $relativePath"
    }
    $files = @($nodes | Where-Object Parent -eq 'DeployProject/Deployable/File')
    $targets = @($files | ForEach-Object {
        if ($_.Properties['Exclude'] -ne '11:FALSE') { throw "Excluded payload in $relativePath" }
        $target = $_.Properties['TargetName'] -replace '^8:', ''
        if (-not $target) { $target = ($_.Properties['SourcePath'] -split '\\')[-1] }
        $target -replace '^8:', ''
    })
    $required = @('EasyHook.dll', 'EasyHook32.dll', 'OutlookFileDrag.dll.manifest', 'OutlookFileDrag.vsto', 'log4net.dll')
    if ($relativePath -like '*_x64/*') { $required += 'EasyHook64.dll' }
    elseif ('EasyHook64.dll' -in $targets) { throw "64-bit native payload in x86 installer" }
    foreach ($target in $required) {
        if (@($targets | Where-Object { $_ -eq $target }).Count -ne 1) { throw "Missing or duplicate $target in $relativePath" }
    }
    if (@($targets | Group-Object | Where-Object Count -gt 1).Count -ne 0) { throw "Duplicate file targets in $relativePath" }
    $common = @('EasyHook32.dll','Microsoft.VisualStudio.Tools.Applications.Runtime.dll','Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll','EasyHook.dll','Microsoft.Office.Tools.Outlook.dll','Microsoft.Office.Tools.dll','stdole.dll','License.rtf','Microsoft.Office.Tools.Common.dll','OutlookFileDrag.dll.manifest','Microsoft.Office.Tools.Common.v4.0.Utilities.dll','Microsoft.Office.Tools.v4.0.Framework.dll','log4net.dll','OutlookFileDrag.vsto','OutlookFileDrag.dll','OutlookFileDrag.dll.config')
    $isX64 = $relativePath -like '*_x64/*'
    $architecture = if ($isX64) { 'x64' } else { 'x86' }
    $expectedTargets = if ($isX64) { @($common) + @('EasyHook64.dll') } else { $common }
    if ($files.Count -ne $expectedTargets.Count -or $targets.Count -ne $expectedTargets.Count) {
        throw "Expected $($expectedTargets.Count) explicit file blocks in $relativePath; found $($files.Count)"
    }
    foreach ($file in $files) {
        if ($file.Name -notmatch '^\{1FB2D0AE-D3B9-43D4-B9DD-F88EC61E35DE\}:_[0-9A-F]{32}$' -or
            $file.Properties['IsDependency'] -ne '11:FALSE' -or
            $file.Properties['TargetName'] -notmatch '^8:(.+)$') {
            throw "Invalid explicit file metadata in $relativePath"
        }
        $target = $Matches[1]
        if ($target -notin $expectedTargets) { throw "Unexpected $target in $relativePath" }
        if ($target -eq 'License.rtf') {
            $expectedSource = '8:..\\License.rtf'
        } elseif ($target -in @('EasyHook32.dll','EasyHook64.dll')) {
            $expectedSource = '8:..\\OutlookFileDrag\\' + $target
        } elseif (-not $isX64 -and $target -eq 'stdole.dll') {
            $expectedSource = '8:..\\OutlookFileDrag\\bin\\Release\\x64\\stdole.dll'
        } else {
            $expectedSource = '8:..\\OutlookFileDrag\\bin\\Release\\' + $architecture + '\\' + $target
        }
        if ($file.Properties['SourcePath'] -cne $expectedSource) {
            throw "Wrong source path for $target in $relativePath"
        }
    }
    $product = @($nodes | Where-Object { $_.Parent -eq 'DeployProject/Deployable' -and $_.Name -eq 'Product' })
    if ($product.Count -ne 1) { throw "Missing Product block in $relativePath" }
    $expectedProductCode = if ($isX64) { '8:{FC419E59-A074-47A8-B00F-966E0AEC6819}' } else { '8:{C5621B14-2645-4D21-A8A1-526966CC1846}' }
    $expectedUpgradeCode = if ($isX64) { '8:{65870D9B-6652-4150-830B-C5199F26E62C}' } else { '8:{2F626147-8F83-4FC1-9190-32AA4F25D487}' }
    $expectedPlatform = if ($isX64) { '3:1' } else { '3:0' }
    if ($product[0].Properties['ProductCode'] -ne $expectedProductCode -or
        $product[0].Properties['UpgradeCode'] -ne $expectedUpgradeCode -or
        $product[0].Properties['ProductVersion'] -ne '8:1.0.13' -or
        $product[0].Properties['TargetPlatform'] -ne $expectedPlatform) {
        throw "Unexpected MSI identity or architecture in $relativePath"
    }
    Write-Host "PASS: $relativePath structure, $($files.Count) explicit payloads, sources and MSI identities"
}
