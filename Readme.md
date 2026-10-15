# Outlook File Drag

Outlook File Drag exposes Outlook messages and attachments as temporary files for applications that accept file paths through CF_HDROP.

This fork builds on [Tony Federer's OutlookFileDrag](https://github.com/tonyfederer/OutlookFileDrag), licensed under the MIT License.

## Downloads

Download the [latest published release](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/latest):

- [64-bit Windows MSI (ZIP)](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/latest/download/OutlookFileDragSetup_x64.zip)
- [32-bit Windows MSI (ZIP)](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/latest/download/OutlookFileDragSetup.zip)

Until 1.0.13 is published, these links continue to point to the previous published release. Select the installer for the Windows architecture. The add-in supports **classic Outlook for Windows** only.

The 1.0.13 installers have been packaged and tested. Their VSTO manifests use a self-signed community-build certificate; the MSI files are not signed by a publicly trusted publisher. Windows or Office may show a publisher warning.

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
