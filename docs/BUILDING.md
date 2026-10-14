# Building and releasing

## Development requirements

Use Windows, Visual Studio 2022 with the Office/SharePoint development workload, the .NET Framework 4.8 targeting pack, and NuGet. Building the MSI projects also requires the Visual Studio Installer Projects extension.

Restore the `packages.config` dependencies:

```powershell
nuget restore OutlookFileDrag.sln
```

Compile the VSTO add-in in Release configuration. When building for local tests or CI, manifest signing can be disabled, but the resulting manifests are not suitable for a public release.

## Continuous integration

The GitHub Actions workflow runs the managed C# compiler target (`CoreCompile`) on a Windows runner. It does not complete the VSTO manifest generation, build MSIs, test Outlook integration, or sign deployment manifests. A passing check confirms managed-source compilation only.

Before publishing an MSI:

1. Build the add-in with trusted publisher credentials and verify both signed VSTO manifests.
2. Build x86 and x64 setup projects using Visual Studio Installer Projects.
3. Inspect actual MSI ProductVersion, ProductCode, UpgradeCode, PackageCode, file versions, architectures, registry registration, and dependency payloads.
4. Test a clean install and upgrades from 1.0.11 and v1.0.12-fixed, including unattended installations.
5. Validate drag-and-drop behavior using classic Outlook 2019 and Microsoft 365, including full messages and attachments.
6. Run a staged deployment before distributing the new installers.

## Signing certificate

Do not use the repository's historical PFX artifact as a release credential without verifying its provenance. Files in a public Git repository remain accessible in its history even after deletion. If any private key was published, treat that key as potentially exposed and replace/revoke it as appropriate.

Use signing certificates outside the repository and never commit private key files or passwords. Release builds must produce manifests signed with a trusted certificate.

The current dependency set contains `log4net` 2.0.8, which NuGet flags for published security advisories. Review a compatible supported update and run regression tests before releasing.

See [reliability and release validation](RELIABILITY_AND_RELEASE_VALIDATION.md).
