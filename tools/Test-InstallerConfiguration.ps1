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
    $outputs = @($nodes | Where-Object Parent -eq 'DeployProject/Deployable/ProjectOutput')
    $groups = @($outputs | ForEach-Object { $_.Properties['OutputGroupCanonicalName'] })
    if ($outputs.Count -ne 2 -or '8:Built' -notin $groups -or '8:Symbols' -notin $groups) {
        throw "Expected Built and Symbols inside ProjectOutput in $relativePath"
    }
    if (@($nodes | Where-Object { $_.Properties['OutputGroupCanonicalName'] -eq '8:ContentFiles' }).Count -ne 0) {
        throw "Duplicate content output group in $relativePath"
    }
    if (@($outputs | Where-Object { $_.Properties['Exclude'] -ne '11:FALSE' }).Count -ne 0) {
        throw "Excluded project output in $relativePath"
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
    Write-Host "PASS: $relativePath structure, output groups, EasyHook and manifest payloads"
}
