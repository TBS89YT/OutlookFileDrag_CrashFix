# Outlook File Drag: reliability audit and release gates

Audit/update date: 2026-10-14. Managed builds and 114 regression cases in each of
the x86 and x64 Windows processes pass, including native OLE helper tests.
**Classic Outlook behavior, signed VSTO deployment and MSI installation remain
untested. These changes are not in the existing v1.0.12-fixed MSI. Do not deploy yet.**

The five additional findings (partial exports, descriptor bounds, unsafe export
paths, enumeration state and large offsets) have been corrected. See the
[German implementation and validation report](AUDIT_FIX_REPORT_DE.md), including
reproducible commands, test coverage, known limits and remaining release blockers.

## Confirmed source faults

- AssemblyVersion/FileVersion 1.0.11.0 conflicted with x64 MSI 1.0.12.
  Windows Installer may skip an equal-version DLL (issue #55).
- Temporary-folder expiration compared creation time against a FUTURE
  timestamp; this removed recent files on periodic cleanup.
- IDataObject native COM interface required ref arguments, but 11 invocations
  omitted them. These are compile-time mismatches; rebuild from fresh source.
- CF_HDROP used Marshal.AllocHGlobal, whereas TYMED_HGLOBAL requires a
  GlobalAlloc(GMEM_MOVEABLE) allocation.
- ReleaseStgMedium was skipped when the medium had a release controller.
- x86 MSI project still included old EasyHook metadata and Debug object path.

## Investigate machines where add-in does nothing

Compare a failing and a working **classic Outlook** PC:
1. Capture Outlook bitness, build/channel, Windows build and destination app.
2. File > Options > Add-ins: inspect Active/Inactive/Disabled Items and Slow
   and Disabled COM Add-ins.
3. Read %APPDATA%\OutlookFileDrag\OutlookFileDrag.log for startup/hook/
   drag messages; absence of a log suggests VSTO loading/log initialization
   problems and requires checking Windows Event Viewer too.
4. Inspect HKLM/HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookFileDrag
   and WOW6432Node paths: Manifest, LoadBehavior, policy overrides.
5. Inspect the actual OutlookFileDrag.dll FileVersion, EasyHook.dll and
   native EasyHook DLLs in the install directory. Validate VSTO manifest trust.
6. Drag both one .msg and one attachment into multiple known compatible
   destinations. Inspect %TEMP%\OutlookFileDrag immediately after; compare
   produced file sizes and capture errors. Some targets support virtual file
   formats but not CF_HDROP and vice versa.
7. Inspect Windows Application log for VSTO/.NET crash information, and
   enterprise security application-control logs.
8. Capture an MSI log on a TEST DEVICE with:
   msiexec /i "<path to MSI>" /qn /norestart /L*v "<log file>"
9. Redact subjects, filenames, mail addresses and customer data from logs
   before posting them publicly.

## Release blocking validation (x86 and x64)

- [x] Managed Release/Debug compilation, native/helper regression tests in both
      process architectures, dependency/version/setup consistency checks.
- [x] Microsoft analyzers for reliability/security/interoperability; no warnings
      or errors, four informational suggestions in existing code.
- [ ] Visual Studio Release build with .NET Framework 4.8, VSTO and VS
      Installer Projects tooling, zero compilation errors.
- [ ] Inspect generated MSI and installed files for DLL file version
      1.0.13.0, correct EasyHook binaries, registration, manifests and
      trustworthy publishing signature.
- [ ] Verify distinct updated ProductCodes, preserved UpgradeCodes, unique
      PackageCodes and product version 1.0.13 in the MSI tables.
- [ ] Test clean install, 1.0.11 upgrade, 1.0.12-fixed x64 upgrade,
      silent Intune/SYSTEM install, repair, rollback and uninstall.
- [ ] Test Outlook 2019 and Microsoft 365 classic, both bitness variants,
      and at least two Office update channels across working/failing PCs.
- [ ] Test large/full/multiple messages, attachments, long and Unicode
      filenames, varied target applications and returned drop effects.
- [ ] Stress-test long Outlook uptime and drags during the cleanup window.
- [ ] Pilot a few devices; only then consider a 100+ device deployment.

Windows Installer version rules:
https://learn.microsoft.com/en-us/windows/win32/msi/file-versioning-rules

HGLOBAL COM contract:
https://learn.microsoft.com/en-us/windows/win32/api/objidl/ne-objidl-tymed

Upstream user reports:
https://github.com/tonyfederer/OutlookFileDrag/issues/55
https://github.com/tonyfederer/OutlookFileDrag/issues/56
