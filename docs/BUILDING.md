# Building and releasing

## Development requirements

Use Windows, Visual Studio 2022 with the Office/SharePoint development workload, the .NET Framework 4.8 targeting pack, and NuGet. Building the MSI projects also requires the Visual Studio Installer Projects extension.

Restore the `packages.config` dependencies:

```powershell
nuget restore OutlookFileDrag.sln
```

Compile the VSTO add-in in Release configuration. When building for local tests or CI, manifest signing can be disabled, but the resulting manifests are not suitable for a public release.

## Continuous integration

The GitHub Actions workflow runs the managed compilation target (`Compile`) on a Windows runner. It does not complete the VSTO manifest generation, build MSIs, test Outlook integration, or sign deployment manifests. A passing check confirms managed-source compilation only.

The workflow also runs `tools/Test-InstallerConfiguration.ps1`. This read-only check parses both installer projects, verifies block structure and project-output nesting, rejects `ContentFiles`, and checks explicit EasyHook and VSTO manifest payloads for missing or duplicate targets. It also runs at the start of `tools/Validate.ps1`. These configuration checks do not replace building and inspecting the actual MSIs.

Before publishing an MSI:

1. Build the add-in with trusted publisher credentials and verify both signed VSTO manifests.
2. Build x86 and x64 setup projects using Visual Studio Installer Projects.
3. Inspect actual MSI ProductVersion, ProductCode, UpgradeCode, PackageCode, file versions, architectures, registry registration, and dependency payloads.
4. Test a clean install and upgrades from 1.0.11 and v1.0.12-fixed, including unattended installations.
5. Validate drag-and-drop behavior using classic Outlook 2019 and Microsoft 365, including full messages and attachments.
6. Run a staged deployment before distributing the new installers.

## Signing certificate

The historical PFX file was removed from the current source tree; it remains available in Git history. Do not use it as a release credential without verifying its provenance. Files in a public Git repository remain accessible in its history even after deletion. If any private key was published, treat that key as potentially exposed and replace/revoke it as appropriate.

Use signing certificates outside the repository and never commit private key files or passwords. Release builds must produce manifests signed with a trusted certificate. The project intentionally contains no fixed `ManifestKeyFile` or `ManifestCertificateThumbprint`; select an approved certificate in the release build environment and verify both generated VSTO manifests. A clean checkout is not configured to produce signed manifests without this step.

The current dependency set is `log4net` 3.5.0 (`net462`) and EasyHook 2.7.7097. The project and both setup projects reference these identities. Both installers require .NET Framework 4.8.

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
