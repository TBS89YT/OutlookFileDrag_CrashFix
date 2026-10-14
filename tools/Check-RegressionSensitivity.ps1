[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$MSBuildPath)
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path $PSScriptRoot
Push-Location $taskRoot
try {
    $taskResults = @()
    foreach ($taskCase in @('AtomicExport','DescriptorBounds','ExportPaths','EnumeratorState','LargeOffsets')) {
        # Only disposable source copies are mutated. No malformed native pointer is dereferenced.
        $taskDir = Join-Path $taskRoot "artifacts/mutations/$taskCase"
        New-Item -ItemType Directory -Path $taskDir -Force | Out-Null
        foreach ($taskFile in @('DataObjectHelper','OutlookDataObject','FileUtility','NativeMethods','FormatEtcEnumerator')) {
            Copy-Item "OutlookFileDrag/$taskFile.cs" "$taskDir/$taskFile.cs" -Force
        }
        function Edit-Mutation([string]$File, [string]$Before, [string]$After) {
            $taskPath = Join-Path $taskDir $File
            $taskText = [IO.File]::ReadAllText($taskPath)
            if (-not $taskText.Contains($Before)) { throw "Mutation anchor missing: $taskCase/$File" }
            [IO.File]::WriteAllText($taskPath, $taskText.Replace($Before,$After), [Text.UTF8Encoding]::new($true))
        }
        switch ($taskCase) {
            'AtomicExport' {
                Edit-Mutation 'OutlookDataObject.cs' 'var completedFiles = new string[filenames.Length];' 'var completedFiles = new string[filenames.Length]; tempFilenames = completedFiles;'
            }
            'DescriptorBounds' {
                Edit-Mutation 'DataObjectHelper.cs' 'if (required > (ulong)bytes.Length)' 'if (false && required > (ulong)bytes.Length)'
            }
            'ExportPaths' {
                Edit-Mutation 'FileUtility.cs' 'ValidateExportName(name);' '// Mutation: filename validation removed.'
                Edit-Mutation 'FileUtility.cs' 'ValidateExportName(candidateName);' '// Mutation: candidate validation removed.'
                Edit-Mutation 'FileUtility.cs' 'if (!string.Equals(Path.GetDirectoryName(target), root, StringComparison.OrdinalIgnoreCase))' 'if (false && !string.Equals(Path.GetDirectoryName(target), root, StringComparison.OrdinalIgnoreCase))'
            }
            'EnumeratorState' {
                Edit-Mutation 'FormatEtcEnumerator.cs' 'new FormatEtcEnumerator(formats, targetDevices, index)' 'new FormatEtcEnumerator(formats, targetDevices, 0)'
                Edit-Mutation 'FormatEtcEnumerator.cs' 'int count = Math.Min(celt, formats.Length - index);' 'int count = Math.Min(celt, formats.Length - index); if (celt > formats.Length - index) return NativeMethods.S_FALSE;'
            }
            'LargeOffsets' {
                Edit-Mutation 'DataObjectHelper.cs' 'source.ReadAt(offset, buffer, requested, ptrRead);' 'source.ReadAt(unchecked((int)offset), buffer, requested, ptrRead);'
            }
        }
        [xml]$taskProject = Get-Content tests/OutlookFileDrag.RegressionTests.csproj -Raw
        foreach ($taskItem in $taskProject.Project.ItemGroup.Compile) {
            if ($taskItem.Include -eq 'RegressionTests.cs') { $taskItem.Include = "$taskRoot/tests/RegressionTests.cs" }
            else { $taskItem.Include = Join-Path $taskDir ([IO.Path]::GetFileName($taskItem.Include)) }
        }
        $taskLogReference = $taskProject.Project.ItemGroup.Reference | Where-Object Include -eq 'log4net'
        $taskLogReference.HintPath = "$taskRoot/packages/log4net.3.5.0/lib/net462/log4net.dll"
        $taskProject.Save("$taskDir/Mutation.csproj")
        & $MSBuildPath "$taskDir/Mutation.csproj" /t:Rebuild /p:Platform=x64 /p:UseSharedCompilation=false /nologo /verbosity:minimal *> "$taskDir/build.log"
        if ($LASTEXITCODE -ne 0) { throw "Mutation failed to compile: $taskCase; inspect $taskDir/build.log" }
        & "$taskDir/bin/x64/OutlookFileDrag.RegressionTests.exe" *> "$taskDir/tests.log"
        $taskExit = $LASTEXITCODE
        $taskFailed = @(Select-String -Path "$taskDir/tests.log" -Pattern '^FAIL ').Count
        if ($taskExit -eq 0 -or $taskFailed -eq 0) { throw "Regression suite did not detect mutation: $taskCase" }
        $taskResults += [pscustomobject]@{Mutation=$taskCase; Detected=$true; FailedTests=$taskFailed; ExitCode=$taskExit}
        Write-Host "$taskCase detected by $taskFailed tests."
    }
    $taskResults | Export-Csv artifacts/mutation-results.csv -NoTypeInformation
    Write-Host 'All five mutations were detected.'
    $global:LASTEXITCODE = 0
} finally { Pop-Location }
