# Building and releasing

## Development requirements

Use Windows and Visual Studio with the Office/SharePoint development workload, the .NET Framework 4.8 targeting pack, NuGet, and the Visual Studio Installer Projects extension.

Restore the `packages.config` dependencies:

```powershell
nuget restore OutlookFileDrag.sln
```

For a signed Release build, provide an approved local signing certificate through `ManifestCertificateThumbprint` in the build environment. The repository contains no signing certificate, private key, or fixed certificate thumbprint. Build the x64 add-in and setup first, followed by x86. The x86 setup references the architecture-neutral `stdole.dll` from the x64 Release output; this matches the tested 1.0.13 package.

## Continuous integration

The GitHub Actions workflow validates both installer project files, compiles managed code without manifest signing, runs regression tests in x86 and x64 processes, and checks the 1.0.13 assembly version. Its installer check parses the complete block structure and verifies the explicit 16-file x86 and 17-file x64 payload lists, source paths, metadata, architecture and MSI upgrade identities. CI does not build MSIs or sign manifests.

The explicit file references in both setup projects are needed for reproducible packaging. The 1.0.13 MSIs were built in Visual Studio and their actual payloads and signed manifests were checked separately. Before publishing future installers, inspect the actual MSIs and test installation and upgrades.

## Signing certificate

Keep signing certificates outside the repository. Supply only the local certificate thumbprint to the Release build environment; never commit the certificate, private key, password or thumbprint. Verify both generated VSTO manifests and their file hashes before packaging.

The historical PFX was removed from the current source tree but remains accessible in Git history. Do not reuse it without validating its provenance.

The current dependency set is `log4net` 3.5.0 (`net462`) and EasyHook 2.7.7097. Both installers require .NET Framework 4.8.

## Automated local validation without Outlook registration

From the repository directory in Developer PowerShell for Visual Studio:

```powershell
.\tools\Validate.ps1
```

The script finds MSBuild using PATH or `vswhere`. It builds and executes the standalone .NET Framework 4.8 regression runner in real x86 and x64 processes, then builds all managed add-in sources in Release/Debug AnyCPU and Release with x86/x64 platform targets. It verifies dependency identities, payload hashes, versions and setup metadata. No additional test-framework packages are required. Logs, exact argument lists and results are written to `artifacts/`.

The managed builds deliberately use `SignManifests=false`, `GenerateManifests=false` and `PrepareForRunDependsOn=CopyFilesToOutputDirectory`. They do not invoke the VSTO certificate, trust or registration targets. These DLLs are **not a deployable VSTO release**. In particular, ordinary VSTO `Build`/`Rebuild` targets can register/unregister an add-in and manage certificates; do not run them on a production Outlook machine as a harmless compilation check.

To run the installed Microsoft .NET analyzers, supply the analyzer directory from a local .NET SDK:

```powershell
.\tools\Validate.ps1 -NetAnalyzersDirectory 'C:\Program Files\dotnet\sdk\10.0.402\Sdks\Microsoft.NET.Sdk\analyzers'
```

Adjust the SDK version to the installed one. Reliability, security and interoperability rules are enabled through `tools/validation.globalconfig`. Analysis uses a separate output directory and suppresses VSTO cleanup/registration targets. Diagnostics are also written to `artifacts/analysis.sarif`.

Optional regression-sensitivity check:

```powershell
.\tools\Check-RegressionSensitivity.ps1 -MSBuildPath '<full path to MSBuild.exe>'
```

This compiles disposable copies under `artifacts/mutations/`, reintroducing one fault category at a time. The command succeeds only when the tests detect all five mutations. A nonzero exit from each mutated test executable is expected. Original source files remain unchanged. This check avoids native out-of-bounds reads and does not reserve gigabytes of memory.

See [the German implementation and validation report](AUDIT_FIX_REPORT_DE.md) for actual results and outstanding release gates.

See [reliability and release validation](RELIABILITY_AND_RELEASE_VALIDATION.md).
