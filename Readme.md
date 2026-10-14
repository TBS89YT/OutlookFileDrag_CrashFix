# Outlook File Drag

Outlook File Drag exposes Outlook messages and attachments as temporary files for applications that accept file paths through CF_HDROP.

This fork builds on [Tony Federer's OutlookFileDrag](https://github.com/tonyfederer/OutlookFileDrag), licensed under the MIT License.

## Downloads

Latest published version: [v1.0.12-fixed](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/tag/v1.0.12-fixed).

- [64-bit Windows MSI (ZIP)](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/download/v1.0.12-fixed/OutlookFileDragSetup_x64.zip)
- [32-bit Windows MSI (ZIP)](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/download/v1.0.12-fixed/OutlookFileDragSetup.zip)

Select the installer for the Windows architecture. This add-in supports **classic Outlook for Windows**, not the new Outlook application.

The current source tree includes **unreleased 1.0.13 changes**. The published v1.0.12-fixed installers do not contain these changes; no 1.0.13 installer has been verified or released.

## How it works

Outlook exposes dragged messages through virtual file data formats. Some receiving applications require actual file paths rather than virtual streams. Outlook File Drag wraps the drag data object, writes the requested content to temporary files, and provides CF_HDROP to compatible destinations.

Applications that support Outlook's virtual file formats directly may not need the add-in. Compatibility depends on the Outlook build, Windows configuration, and destination application.

## Installation

Close classic Outlook, extract the ZIP appropriate for your Windows architecture, install the MSI, and restart Outlook.

Silent installation with a verbose log:

```powershell
msiexec.exe /i "C:\Install\OutlookFileDragSetup_x64.msi" /qn /norestart /L*v "C:\Install\OutlookFileDrag-install.log"
```

Uninstall using the matching MSI:

```powershell
msiexec.exe /x "C:\Install\OutlookFileDragSetup_x64.msi" /qn /norestart /L*v "C:\Install\OutlookFileDrag-uninstall.log"
```

These paths are examples. Product codes differ by version and installer architecture. For Intune deployments, obtain the ProductCode from the actual MSI and verify upgrade behavior on test devices before rollout.

## Troubleshooting

Check Outlook's Active, Inactive, and Disabled COM Add-ins. When logging initializes, diagnostic information is written to:

```text
%APPDATA%\OutlookFileDrag\OutlookFileDrag.log
```

Capture the Outlook build, bitness, Windows version, destination application, and whether the failed drag involved a complete email or an attachment. Redact personal and customer data from shared logs.

See [reliability and release validation](docs/RELIABILITY_AND_RELEASE_VALIDATION.md) for investigation and test guidance.

## Building

The project targets .NET Framework 4.8 and uses VSTO, EasyHook, and log4net. Consult [building and packaging](docs/BUILDING.md) for development requirements and release checks.

## License

Original project by [Tony Federer](https://github.com/tonyfederer). Licensed under the [MIT License](LICENSE). Dependencies include [EasyHook](https://easyhook.github.io/) and [log4net](https://logging.apache.org/log4net/).
