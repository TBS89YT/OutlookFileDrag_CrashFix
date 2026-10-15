# OutlookFileDrag 1.0.13

This community maintenance release improves reliability when dragging Outlook messages and attachments into applications that require real file paths.

- Fixes crashes and incomplete exports in Outlook drag-and-drop handling.
- Improves validation of exported filenames and Outlook file descriptors.
- Updates the add-in to file version 1.0.13.0 and includes log4net 3.5.0.
- Provides separate x86 and x64 Windows installers, tested with 16 and 17 packaged files respectively.

Download from the [latest published release](https://github.com/TBS89YT/OutlookFileDrag_CrashFix/releases/latest).

This fork builds on Tony Federer's original [OutlookFileDrag project](https://github.com/tonyfederer/OutlookFileDrag).

The add-in supports classic Outlook for Windows only. It does not support the new Outlook. The VSTO manifests are signed with a self-signed community-build certificate, and the MSI files are not signed by a publicly trusted publisher. Windows or Office can display a publisher warning.
