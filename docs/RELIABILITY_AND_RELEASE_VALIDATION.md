# Outlook File Drag: reliability audit and release gates

Audit originally performed: 2026-10-14. **Release status updated: 2026-10-15.**
Managed builds and 114 regression cases in each of the x86 and x64 Windows
processes pass, including native OLE helper tests. Since the initial audit,
the 1.0.13 x86 and x64 MSIs were built, extracted and verified, including
their architecture-specific DLL hashes, payload files, VSTO manifest
signatures and MSI metadata. The user reports successful installation tests
of both final installers. Those results supersede the original **untested
MSI/VSTO packaging** finding; the published v1.0.12-fixed assets do not
contain the 1.0.13 fixes. **Broader Outlook interoperability, upgrade,
repair, silent deployment and enterprise rollout tests are not established
by these installation tests.** The VSTO manifests use a self-signed
community-build certificate, and the MSIs are not Authenticode-signed.

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

## 1.0.13 release verification (updated 2026-10-15)

- [x] Managed Release/Debug compilation; native/helper regression tests in
      actual x86 and x64 processes (114 cases per architecture).
- [x] Microsoft analyzers: no warnings or errors, four informational findings.
- [x] Visual Studio Release packaging completed: x86 MSI has 16 files; x64
      MSI has 17 files. Extracted package DLL SHA-256 values match the
      approved architecture-specific release DLLs.
- [x] Both VSTO application and deployment manifest signatures verified for
      each architecture, including file references and hashes. Signatures
      use a self-signed community-build certificate, not a publicly trusted
      code-signing identity. The MSIs themselves are unsigned.
- [x] Inspected the actual MSI files: version 1.0.13, ProductCode/UpgradeCode,
      target platform, installation location, Outlook registration and
      included dependencies. No PDB/PFX/private-key files or identifiable
      local debug paths found in the release packages.
- [x] User confirmed successful installation tests of both final MSI variants.
- [x] Local x86/x64 ZIP archives and their SHA-256 checksum manifest verified.
      For current public availability, see the project's GitHub Releases page.

## Further compatibility and enterprise rollout checks

The items below are useful before deploying widely. They are **not claimed
as completed** by the successful packaging and basic installation tests:

- [ ] Exercise both full-message and attachment drag-and-drop in classic
      Outlook on representative 32-bit and 64-bit installations, including
      different Outlook versions, Office update channels and destination apps.
- [ ] Test upgrades from 1.0.11 and v1.0.12-fixed using the **final**
      installers, and confirm repair, rollback and uninstall behavior.
- [ ] Test unattended Intune/SYSTEM installation and application-control
      environments; document handling of self-signed publisher warnings.
- [ ] Test multiple/large messages, unusual filenames and sustained Outlook
      uptime where relevant.
- [ ] Pilot representative devices before a fleet-wide rollout.

Windows Installer version rules:
https://learn.microsoft.com/en-us/windows/win32/msi/file-versioning-rules

HGLOBAL COM contract:
https://learn.microsoft.com/en-us/windows/win32/api/objidl/ne-objidl-tymed

Upstream user reports:
https://github.com/tonyfederer/OutlookFileDrag/issues/55
https://github.com/tonyfederer/OutlookFileDrag/issues/56
