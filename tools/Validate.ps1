[CmdletBinding()]
param(
    [string]$MSBuildPath,
    [string]$VisualStudioVersion,
    [string]$NetAnalyzersDirectory
)
$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path $PSScriptRoot
Push-Location $taskRoot
try {
    & (Join-Path $PSScriptRoot 'Test-InstallerConfiguration.ps1') -RepositoryRoot $taskRoot
    if (-not $MSBuildPath) {
        $taskMSBuild = Get-Command msbuild.exe -ErrorAction SilentlyContinue
        if ($taskMSBuild) { $MSBuildPath = $taskMSBuild.Source }
        else {
            $taskVswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
            $MSBuildPath = & $taskVswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        }
    }
    if (-not $MSBuildPath) { throw 'MSBuild not found. Install Visual Studio with Office development tools.' }
    if (-not $VisualStudioVersion) {
        $VisualStudioVersion = ([Diagnostics.FileVersionInfo]::GetVersionInfo($MSBuildPath).ProductMajorPart).ToString() + '.0'
    }
    New-Item -ItemType Directory -Path artifacts -Force | Out-Null
    $taskCommands = [Collections.Generic.List[object]]::new()
    function Invoke-ValidationCommand([string]$Executable, [string[]]$Arguments, [string]$Log) {
        $taskCommands.Add([pscustomobject]@{Executable=$Executable; Arguments=$Arguments; Log=$Log})
        $taskCommands | ConvertTo-Json -Depth 5 | Set-Content artifacts/validation-commands.json -Encoding UTF8
        & $Executable @Arguments 2>&1 | Tee-Object -FilePath $Log
        if ($LASTEXITCODE -ne 0) { throw "Command failed ($LASTEXITCODE): $Executable; see $Log" }
    }
    foreach ($taskPlatform in @('x86', 'x64')) {
        Invoke-ValidationCommand $MSBuildPath @('tests/OutlookFileDrag.RegressionTests.csproj','/t:Rebuild',"/p:Platform=$taskPlatform",'/p:UseSharedCompilation=false','/nologo','/verbosity:minimal') "artifacts/tests-build-$taskPlatform.log"
        Invoke-ValidationCommand (Join-Path $taskRoot "tests/bin/$taskPlatform/OutlookFileDrag.RegressionTests.exe") @() "artifacts/tests-$taskPlatform.log"
    }
    # Compile and copy all C# outputs. VSTO manifest/certificate/registration targets are excluded.
    # These outputs alone are not a deployable VSTO release.
    foreach ($taskVariant in @('Release-AnyCPU', 'Debug-AnyCPU', 'Release-x86', 'Release-x64')) {
        $taskConfig, $taskTarget = $taskVariant.Split('-')
        $taskArgs = @('OutlookFileDrag/OutlookFileDrag.csproj','/t:Build',"/p:Configuration=$taskConfig",'/p:Platform=AnyCPU',"/p:PlatformTarget=$taskTarget",
            '/p:SignManifests=false','/p:GenerateManifests=false','/p:PrepareForRunDependsOn=CopyFilesToOutputDirectory',
            '/p:UseSharedCompilation=false',"/p:VisualStudioVersion=$VisualStudioVersion",'/nologo','/verbosity:minimal')
        if ($taskTarget -ne 'AnyCPU') {
            $taskArgs += "/p:OutputPath=bin\Validation\$taskTarget\"
            $taskArgs += "/p:IntermediateOutputPath=obj\Validation\$taskTarget\"
        }
        Invoke-ValidationCommand $MSBuildPath $taskArgs "artifacts/managed-build-$taskVariant.log"
    }
    if ($NetAnalyzersDirectory) {
        $taskArgs = @('OutlookFileDrag/OutlookFileDrag.csproj','/t:Rebuild','/p:Configuration=Release','/p:Platform=AnyCPU',
            '/p:UseSharedCompilation=false',"/p:VisualStudioVersion=$VisualStudioVersion",'/p:SignManifests=false','/p:RunAnalyzers=true','/p:ReportAnalyzer=true',
            '/p:VSTOCleanDependsOn=','/p:PrepareForRunDependsOn=CopyFilesToOutputDirectory','/p:GenerateManifests=false',
            '/p:OutputPath=bin\Validation\Analysis\','/p:IntermediateOutputPath=obj\Validation\Analysis\',
            "/p:CustomBeforeMicrosoftCSharpTargets=$PSScriptRoot\Analyze.targets", "/p:NetAnalyzersDirectory=$NetAnalyzersDirectory",
            '/p:ErrorLog=..\artifacts\analysis.sarif','/nologo','/verbosity:normal')
        Invoke-ValidationCommand $MSBuildPath $taskArgs 'artifacts/static-analysis.log'
    } else { Write-Host 'Static analyzers not selected; pass -NetAnalyzersDirectory to enable them.' }

    $taskChecks = [Collections.Generic.List[object]]::new()
    function Assert-Consistency([string]$Name, [bool]$Condition) {
        $taskChecks.Add([pscustomobject]@{Check=$Name; Passed=$Condition})
        if (-not $Condition) { throw "Consistency check failed: $Name" }
    }
    [xml]$taskPackages = Get-Content OutlookFileDrag/packages.config -Raw
    [xml]$taskProject = Get-Content OutlookFileDrag/OutlookFileDrag.csproj -Raw
    Assert-Consistency 'log4net package 3.5.0' (@($taskPackages.packages.package | Where-Object id -eq 'log4net')[0].version -eq '3.5.0')
    Assert-Consistency 'EasyHook package 2.7.7097' (@($taskPackages.packages.package | Where-Object id -eq 'EasyHook')[0].version -eq '2.7.7097')
    Assert-Consistency 'ApplicationVersion 1.0.13.0' ($taskProject.Project.PropertyGroup.ApplicationVersion -contains '1.0.13.0')
    foreach ($taskDll in @('log4net','EasyHook')) {
        $taskVersion = if ($taskDll -eq 'log4net') {'3.5.0.0'} else {'2.7.7097.0'}
        $taskRef = @($taskProject.Project.ItemGroup.Reference | Where-Object { $_.Include -like "$taskDll,*" })[0]
        $taskPath = [IO.Path]::GetFullPath((Join-Path "$taskRoot/OutlookFileDrag" $taskRef.HintPath))
        $taskAssembly = [Reflection.AssemblyName]::GetAssemblyName($taskPath)
        Assert-Consistency "$taskDll reference identity" ($taskAssembly.Version.ToString() -eq $taskVersion -and $taskRef.Include -like "$taskDll, Version=$taskVersion,*")
        Assert-Consistency "$taskDll Release payload hash" ((Get-FileHash $taskPath).Hash -eq (Get-FileHash "OutlookFileDrag/bin/Release/$taskDll.dll").Hash)
    }
    foreach ($taskBits in @('32','64')) {
        Assert-Consistency "EasyHook$taskBits native package hash" ((Get-FileHash "OutlookFileDrag/EasyHook$taskBits.dll").Hash -eq
            (Get-FileHash "packages/EasyHook.2.7.7097/content/net40/EasyHook$taskBits.dll").Hash)
    }
    foreach ($taskInstaller in @('OutlookFileDragSetup/OutlookFileDragSetup.vdproj','OutlookFileDragSetup_x64/OutlookFileDragSetup_x64.vdproj')) {
        $taskText = Get-Content $taskInstaller -Raw
        Assert-Consistency "$taskInstaller product version" ($taskText.Contains('"ProductVersion" = "8:1.0.13"'))
        Assert-Consistency "$taskInstaller log4net identity" ($taskText.Contains('log4net, Version=3.5.0.0') -and -not $taskText.Contains('log4net, Version=2.0.8'))
        Assert-Consistency "$taskInstaller log4net source" ($taskText.Contains('..\\packages\\log4net.3.5.0\\lib\\net462\\log4net.dll'))
        Assert-Consistency "$taskInstaller .NET 4.8 requirement" ($taskText.Contains('"FrameworkVersion" = "8:.NETFramework,Version=v4.8"') -and -not $taskText.Contains('Version=v4.5') -and $taskText.Contains('https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48'))
        $taskBaseline = Join-Path $taskRoot "artifacts/baseline-20261014/$taskInstaller"
        if (Test-Path $taskBaseline) {
            $taskOriginal = Get-Content $taskBaseline -Raw
            foreach ($taskIdentity in @('ProductCode','UpgradeCode','PackageCode')) {
                $taskPattern = '"' + $taskIdentity + '" = "8:(\{[^}]+\})"'
                Assert-Consistency "$taskInstaller $taskIdentity preserved" ([regex]::Match($taskText,$taskPattern).Groups[1].Value -eq [regex]::Match($taskOriginal,$taskPattern).Groups[1].Value)
            }
        }
    }
    foreach ($taskDll in @('OutlookFileDrag/bin/Release/OutlookFileDrag.dll','OutlookFileDrag/bin/Debug/OutlookFileDrag.dll',
        'OutlookFileDrag/bin/Validation/x86/OutlookFileDrag.dll','OutlookFileDrag/bin/Validation/x64/OutlookFileDrag.dll')) {
        Assert-Consistency "$taskDll file version" ((Get-Item $taskDll).VersionInfo.FileVersion -eq '1.0.13.0')
        Assert-Consistency "$taskDll assembly version" ([Reflection.AssemblyName]::GetAssemblyName((Join-Path $taskRoot $taskDll)).Version.ToString() -eq '1.0.13.0')
    }
    foreach ($taskBits in @('x86','x64')) {
        $taskBytes = [IO.File]::ReadAllBytes((Join-Path $taskRoot "OutlookFileDrag/bin/Validation/$taskBits/OutlookFileDrag.dll"))
        $taskPeOffset = [BitConverter]::ToInt32($taskBytes, 0x3C)
        $taskMachine = [BitConverter]::ToUInt16($taskBytes, $taskPeOffset + 4)
        Assert-Consistency "$taskBits PE machine" ($taskMachine -eq $(if ($taskBits -eq 'x86') {0x14c} else {0x8664}))
    }
    $taskChecks | ConvertTo-Json -Depth 5 | Set-Content artifacts/consistency-results.json -Encoding UTF8
    Write-Host "$($taskChecks.Count) consistency checks passed. C# validation complete; VSTO/MSI/Outlook gates remain separate."
} finally { Pop-Location }
