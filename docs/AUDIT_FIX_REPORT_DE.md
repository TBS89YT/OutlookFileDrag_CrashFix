# OutlookFileDrag 1.0.13 – Umsetzung und Validierung

Stand: 14.10.2026, lokales Windows-Projekt. **Alle fünf beauftragten Befunde sind korrigiert. Die vollständige C#-Quellkompilierung und die Regressionstests sind erfolgreich. Eine Releasefreigabe für Outlook/VSTO/MSI liegt noch nicht vor.**

## A. Ergebnis und Ausgangszustand

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| Beauftragte Fehler | 5 von 5 bearbeitet und mit Regressionstests abgesichert |
| Regressionstests x86 | 114 bestanden, 0 fehlgeschlagen; echter 32-Bit-Prozess |
| Regressionstests x64 | 114 bestanden, 0 fehlgeschlagen; echter 64-Bit-Prozess |
| Gesamter finaler Testlauf | 228 erfolgreiche Ausführungen derselben 114 Testfälle |
| Verwaltete Add-in-Builds | Release AnyCPU, Debug AnyCPU, Release mit x86- und x64-Plattformziel erfolgreich |
| Konsistenzprüfungen | 33 bestanden |
| Statische Analyse | Microsoft .NET Analyzers: 0 Warnungen, 0 Fehler, 4 Hinweise |
| Prüfung der Testwirksamkeit | Alle 5 absichtlich wieder eingeführten Fehlerkategorien erkannt |
| Vollständige VSTO-Manifeste/Signaturen | Nicht fertiggestellt/verifiziert |
| Neue MSI-Dateien | Nicht erstellt; keine MSI installiert |
| Klassisches Outlook / EasyHook zur Laufzeit | Nicht praktisch getestet |

Bearbeitet wurde direkt das gewünschte Verzeichnis:

```text
<lokales Projektverzeichnis>
```

Es handelt sich um einen lokalen Quellarchiv-Auszug ohne `.git`. Das mit Visual Studio gelieferte Git wurde gefunden und mit `rev-parse --show-toplevel` aufgerufen; Ergebnis: `not a git repository`. Deshalb waren Arbeitsbranch, Git-Status, Vergleich gegen Remote, Commit und Push hier nicht verfügbar. Es wurde kein neues Repository angelegt und nichts veröffentlicht.

Vor den Änderungen wurden 21 vorhandene Quell-, Konfigurations-, Lösungs-, Installer- und Dokumentationsdateien nach `artifacts/baseline-20261014/` kopiert und ihre SHA-256-Prüfsummen in `SHA256.csv` gespeichert. Zertifikatsdateien wurden dabei nicht gelesen oder kopiert. Die Sicherung ist keine vollständige Kopie aller Projektdateien; insbesondere die Workflowdatei wurde vor der Änderung gelesen, aber nicht durch diese Dateiauswahl gesichert. Die lokale Sicherung und Buildartefakte liegen im bereits ignorierten Verzeichnis `artifacts/`.

Der bereits vorhandene Stand mit log4net 3.5.0 wurde als Grundlage verwendet. `packages.config`, die Assemblyversion 1.0.13.0, der EasyHook-Callback in `DragDropHook.cs`, die `ref`-COM-Aufrufe, die bewegliche CF_HDROP-Allokation und die bereits korrigierte Ablaufzeitberechnung wurden erhalten. Es gab im Ausgangsprojekt keine Testsuite und kein Testprojekt. Die neue Suite benötigt kein zusätzliches Testframework.

## B. Die fünf Korrekturen

### 1. Unvollständiger Dateiexport

**Ursache:** `tempFilenames` wurde vor dem Export veröffentlicht. Nach einer Ausnahme galt ein teilweise gefülltes Array beim nächsten Abruf als fertiger Export.

**Umsetzung:** `ExtractFiles()` erstellt zunächst ausschließlich lokale Ergebnisse. Erst nach dem erfolgreichen Schreiben und Schließen sämtlicher Dateien wird der Cache gesetzt. Eine Sperre koordiniert parallele Abrufe; ein zusätzlicher Status weist rekursive COM-Rückrufe während eines laufenden Exports kontrolliert zurück. Fehler entfernen die in diesem Versuch erzeugten Dateien und anschließend den leeren Exportordner. Bereinigungsfehler werden protokolliert, ohne den ursprünglichen Fehler zu verdecken. Der COM-Einstieg gibt im Fehlerfall weiterhin `E_UNEXPECTED` und ein leeres `STGMEDIUM` zurück; `FilesDropped` wird erst nach erfolgreicher CF_HDROP-Erzeugung gesetzt. Nach einem Fehlschlag ist ein neuer Versuch möglich, erfolgreiche Exporte werden wiederverwendet.

**Datei:** `OutlookFileDrag/OutlookDataObject.cs`. Ein interner Konstruktor erlaubt ausschließlich für deterministische Tests die Bereitstellung eines Testordners und eines Streams mit simuliertem Schreibfehler; der produktive Konstruktor verwendet unverändert einen neuen temporären Ordner und jetzt `FileMode.CreateNew`.

**Tests:** `Export.One`, `Export.Multiple`, `Export.ReadFailureAt0AndRetry`, `Export.ReadFailureAt1AndRetry`, `Export.WriteFailureAndRetry`, `Export.SuccessCached`, `Export.ParallelCalls`, `Export.ReentrantCall`, `Export.StreamReadFailureCleansBatch`, `Export.EmptyBatchRejected`. Alle in x86/x64 erfolgreich. Die Tests prüfen den HRESULT, fehlende Ergebnisdateien bei Fehlern, Bereinigung, Dateiinhalte und Anzahl tatsächlicher Inhaltsabrufe.

### 2. Ungeprüfte native Dateideskriptoren

**Ursache:** Unzuverlässiges `cItems` steuerte Allokation und native `PtrToStructure`-Zugriffe, ohne die erforderliche Gesamtgröße gegen die gelieferten Bytes zu prüfen.

**Umsetzung:** ANSI und Unicode verwenden dieselbe begrenzte Metadatenverarbeitung. Der Parser prüft mindestens vier Headerbytes, das native Layout (332/592 Bytes, Dateiname ab Offset 72), die mit `ulong` berechnete erforderliche Gesamtgröße und die Nullterminierung innerhalb des festen Dateinamenfeldes. Erst danach wird das Ergebnisarray angelegt. Dateinamen werden aus geprüften verwalteten Bytes dekodiert; die vorherige zusätzliche unmanaged Kopie und die darauf ausgeführten nativen Strukturzugriffe entfallen. Ein begrenzter Metadatenstream weist mehr als 16 MiB kontrolliert zurück. Zusätzliche Bytes, etwa durch HGLOBAL-Allokationsrundung, sind erlaubt; die angegebene Anzahl muss vollständig vorhanden sein. Ein beschädigter Unicode-Deskriptor wird als Fehler behandelt und nicht durch stilles ANSI-Fallback verdeckt.

**Datei:** `OutlookFileDrag/DataObjectHelper.cs`.

**Tests:** Die Gruppen `Descriptor.ANSI.*` und `Descriptor.Unicode.*` prüfen gültige Daten, kurze Blöcke, Nullanzahl, überhöhte und maximale Anzahl, abgeschnittene letzte Einträge, falsches Layout, fehlenden Nullabschluss und Freigabe bei Fehlern. Zusätzlich: `Descriptor.MetadataLimit`, `Descriptor.MetadataStreamLimitAndRelease`, `Descriptor.NativeLayouts`, `Descriptor.ReleaseControllerOnFailure`. Native HGLOBAL-Freigabe und die Behandlung eines echten `pUnkForRelease`-Controllers werden geprüft. Alle in x86/x64 erfolgreich.

**Bewusste Grenze:** Die 16-MiB-Grenze betrifft nur Deskriptor-Metadaten, nicht Dateiinhalte. Extrem große Dateiauswahlen oberhalb dieser Grenze werden abgewiesen. Ein gefälschter, grundsätzlich ungültiger COM-Pointer lässt sich durch Längenprüfungen einer Anwendung nicht nachträglich vertrauenswürdig machen; die Korrektur verhindert insbesondere die durch widersprüchliche Deskriptoranzahlen verursachten Grenzüberschreitungen.

### 3. Unsichere temporäre Dateipfade

**Ursache:** Externe Dateinamen wurden direkt mit dem Exportordner kombiniert. Absolute Pfade oder Verzeichnisbestandteile konnten außerhalb des Ordners landen. Die alte Längenkürzung konnte außerdem Teile des vollständigen Pfads abschneiden.

**Umsetzung:** Exportiert werden ausschließlich einzelne, gültige Windows-Dateinamen. Absolute Pfade, Laufwerkspfade, UNC, beide Trennzeichen, Traversal, ungültige Zeichen, leere Namen, abschließende Punkte/Leerzeichen und reservierte Gerätenamen einschließlich COM/LPT-Sonderfällen werden kontrolliert abgelehnt. Der vollständige Zielpfad wird normalisiert und sein Elternordner mit dem normalisierten Exportordner verglichen. `GetUniqueFilename()` kürzt nur den Namensstamm, erhält Erweiterung und Kollisionssuffix, berücksichtigt vorhandene Verzeichnisse und vermeidet das Auftrennen eines UTF-16-Surrogatpaares. `CreateNew` verhindert Überschreiben auch bei einer Kollision nach der Namensprüfung. Gefährliche Namen führen zum Fehler des gesamten Exports, nicht zu einem Teilresultat.

**Dateien:** `OutlookFileDrag/FileUtility.cs`, `OutlookFileDrag/OutlookDataObject.cs`.

**Tests:** `Path.Valid.*`, `Path.Reject.*`, `Path.CollisionPreservesExisting`, `Path.DirectoryCollision`, `Path.LongNameCollision`, `Path.ExcessiveExtension`, `Export.DuplicateNames`, `Export.InvalidLaterNameCleansBatch`. Alle in x86/x64 erfolgreich. Gültige Unicode-Namen bleiben erhalten. Namen über 255 Zeichen und nicht sinnvoll kürzbare Pfade/Erweiterungen werden abgewiesen; normale lange Namensstämme werden innerhalb der bisherigen MAX_PATH-Grenze gekürzt.

### 4. COM-Formatenumerator

**Ursache:** `Clone()` verlor die Position; `Skip()` behandelte das exakte Ende als Fehler und ließ die Position beim Überspringen des Endes unverändert.

**Umsetzung:** Der Klon übernimmt die aktuelle Position und besitzt danach einen eigenen Positionszähler. `Skip()` verschiebt bis zur möglichen Endposition und liefert nur dann `S_OK`, wenn die verlangte Anzahl übersprungen wurde. `Next()` prüft Argumente, setzt die tatsächlich gelieferte Anzahl und liefert bei Teilabrufen/Ende `S_FALSE`. Ungültige Argumente verändern die Position nicht. `Reset()` setzt zuverlässig auf den Anfang zurück. Der ergänzende COM-Review korrigierte auch die Zielgeräte-Speicherverantwortung: optionale `FORMATETC.ptd`-Daten werden intern verwaltet gespeichert und für jeden Abruf frisch mit COM-Speicher kopiert; die vom ursprünglichen Enumerator gelieferten Blöcke werden freigegeben. Fehler des Quell-Enumerators werden nicht als erfolgreiche, unvollständige Aufzählung ausgegeben.

**Dateien:** `OutlookFileDrag/FormatEtcEnumerator.cs`, `OutlookFileDrag/OutlookDataObject.cs`.

**Tests:** `Enum.CloneAt0/1/2`, `Enum.SkipWithin`, `Enum.SkipExactEnd`, `Enum.SkipBeyondEnd`, `Enum.ResetAfterEnd`, `Enum.NextAtEndFetchedZero`, `Enum.PartialNext`, `Enum.InvalidArgumentsDoNotAdvance`, `Enum.ZeroCounts`, `Enum.Empty`, `Enum.TargetDeviceCopies`, `Enum.WrapperFormats`. Alle in x86/x64 erfolgreich. Die Aufzählung behält die bisherige Filterung von Textformaten und das zusätzliche CF_HDROP bei.

### 5. Große Offsets und native Ressourcen

**Ursache:** Der ILockBytes-Leseoffset war ein `int`; HGLOBAL-Pointerarithmetik reduzierte große Offsets auf `int`.

**Umsetzung:** ILockBytes wird mit `long`-Offsets, begrenzten Blöcken und geprüften Rückgabemengen gelesen. Ein vorzeitiges Ende gegenüber der angegebenen Storage-Länge gilt als Fehler. HGLOBAL-Größen werden als natives `SIZE_T` (`UIntPtr`) gelesen. Pointeraddition prüft vorzeichenunabhängig den Plattformadressraum und den gesamten zu lesenden Block; der Schleifenoffset wächst mit geprüfter Addition. Leere gültige HGLOBALs funktionieren. Ein x86-Prozess weist HGLOBALs über `Int32.MaxValue` konservativ zurück. Der `ILockBytes.ReadAt`-Marshallingparameter beschreibt jetzt korrekt die Pufferlänge (`SizeParamIndex=2`). Zusätzliche von `GetObjectForIUnknown` erworbene IStream-/IStorage-RCW-Referenzen werden balanciert freigegeben; Lesemengen und native Lesepuffer werden auch im Fehlerfall behandelt.

**Dateien:** `OutlookFileDrag/DataObjectHelper.cs`, `OutlookFileDrag/NativeMethods.cs`.

**Tests:** `HGlobal.Copy*`, `HGlobal.WriteFailureUnlocks`, `HGlobal.InvalidHandle`, `Pointer.*`, `LockBytes.Copy*`, `LockBytes.Over2GBWithoutLargeAllocation`, `LockBytes.PartialReads`, `LockBytes.PrematureEnd`, `LockBytes.InvalidCount`, `LockBytes.ReadFailure`, `LockBytes.WriteFailure`, `LockBytes.NegativeLength`. Zusätzlich echte OLE-Tests: `LockBytes.NativeInterop`, `IStorage.NativeCopyAndFailureRecovery`, `IStream.NativeCopyAndRelease.False/True`. Alle in x86/x64 erfolgreich. Die großen Offsettests simulieren mehr als 2 GiB mit einem 4-KiB-Puffer und einem zählenden Zielstream; sie reservieren keine Gigabytes RAM.

**Grenze:** Der bestehende IStorage-Weg serialisiert ein Compound Storage weiterhin in ein natives, HGLOBAL-basiertes ILockBytes. Diese native Kopie ist nicht durch eine neue Streamingarchitektur ersetzt worden. Das anschließende Lesen verwendet nur kleine verwaltete Puffer. Reale mehrgigabytegroße IStorage-/HGLOBAL-Übertragungen wurden nicht durchgeführt; insbesondere x86 bleibt durch seinen Adressraum begrenzt.

## C. Technische Validierung und Befehle

Verwendet: Windows, Visual Studio 18 Community, dessen MSBuild/Roslyn, installierte .NET-Framework-4.8-Referenzassemblies und lokale NuGet-Pakete. Das Programm läuft auf CLR `4.0.30319.42000`; das ist bei .NET Framework 4.8 die CLR-Versionsfamilie und kein Zielwechsel auf Framework 4.0.

Der finale reproduzierbare Hauptlauf aus dem Projektverzeichnis war:

```powershell
.\tools\Validate.ps1 `
  -MSBuildPath 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' `
  -VisualStudioVersion 18.0 `
  -NetAnalyzersDirectory 'C:\Program Files\dotnet\sdk\10.0.402\Sdks\Microsoft.NET.Sdk\analyzers'
```

Die exakten Executable-Pfade und Argumentarrays aller neun von diesem Skript gestarteten Befehle stehen in `artifacts/validation-commands.json`. Die JSON-Datei enthält auch den jeweiligen Protokollpfad. Das Skript bricht bei einem fehlerhaften Build, Test oder Konsistenzcheck ab.

Die Befehle entsprechen folgender Matrix (`MSBuild` bezeichnet den vollständigen Pfad oben):

| Ausführung | Zusätzliche Parameter / Ergebnis |
|---|---|
| `MSBuild tests/OutlookFileDrag.RegressionTests.csproj /t:Rebuild /p:Platform=x86 /p:UseSharedCompilation=false /nologo /verbosity:minimal` | Erfolgreich, Warnungen als Fehler aktiviert |
| `tests/bin/x86/OutlookFileDrag.RegressionTests.exe` | 114/114 erfolgreich |
| Derselbe Testbuild mit `/p:Platform=x64` | Erfolgreich |
| `tests/bin/x64/OutlookFileDrag.RegressionTests.exe` | 114/114 erfolgreich |
| `MSBuild OutlookFileDrag/OutlookFileDrag.csproj /t:Build` | Release AnyCPU erfolgreich |
| Derselbe Projektbuild | Debug AnyCPU erfolgreich |
| Derselbe Projektbuild | Release, `/p:PlatformTarget=x86`, separate `bin/Validation/x86` und `obj/Validation/x86` erfolgreich |
| Derselbe Projektbuild | Release, `/p:PlatformTarget=x64`, separate `bin/Validation/x64` und `obj/Validation/x64` erfolgreich |
| Statische Analyse als `/t:Rebuild` | Erfolgreich; separate `Validation/Analysis`-Ausgaben, Analyzer/SARIF eingeschaltet |

Alle vier verwalteten Projektbuilds verwenden zusätzlich `/p:Platform=AnyCPU`, die jeweilige Konfiguration und das jeweilige Plattformziel, `/p:SignManifests=false`, `/p:GenerateManifests=false`, `/p:PrepareForRunDependsOn=CopyFilesToOutputDirectory`, `/p:UseSharedCompilation=false`, `/p:VisualStudioVersion=18.0`, `/nologo` und `/verbosity:minimal`. Damit werden sämtliche im echten Add-in-Projekt eingebundenen C#-Dateien, VSTO-Referenzen und EasyHook-Verweise kompiliert und die DLL-Ausgaben kopiert. Die VSTO-Manifest-/Zertifikats-/Registrierungsphase ist ausdrücklich ausgeschlossen.

Die Analyse verwendet zusätzlich `tools/Analyze.targets`, das konfigurierte SDK-Analyzerverzeichnis, `tools/validation.globalconfig`, `/p:ErrorLog=..\artifacts\analysis.sarif`, `/p:ReportAnalyzer=true`, `/p:RunAnalyzers=true` und `/p:VSTOCleanDependsOn=`. Letzteres verhindert Registry-Bereinigung beim Analyse-Rebuild. Es gab einen tatsächlich ausgeführten Analyzerlauf; die Analyse wird nicht nur aufgrund eines inkrementell übersprungenen Compilerziels als bestanden gewertet.

**Warnungen/Hinweise:** Die finalen Builds melden keine Compilerwarnungen oder Fehler. SARIF enthält vier Hinweise: zweimal CA1861 und einmal CA1825 in unverändertem `DragDropHook.cs`, einmal CA2263 zur vorhandenen nichtgenerischen `Marshal.SizeOf`-Verwendung in `DataObjectHelper.cs`. Diese sind keine Fehlerbefunde der aktivierten Reliability-/Security-/Interop-Kategorien. Die Analyse ist kein vollständiger Sicherheitsnachweis. Ein anfänglicher Versuch mit Compiler-Server meldete einen Verbindungsfehler zum Server und kompilierte per Fallback erfolgreich; die finalen Befehle verwenden deshalb `UseSharedCompilation=false`.

**Frühe, behobene Testinfrastrukturprobleme:** Beim ersten Testprojektbuild war `STATSTG` wegen zweier Namespace-Typen mehrdeutig. Ein Alias auf `ComTypes.STATSTG` beseitigte den Fehler. Alle finalen Testbuilds sind erfolgreich.

**Standard-VSTO-Rebuild:** Vor der sicheren Buildmatrix wurde folgender vollständiger Rebuild versucht:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' `
  OutlookFileDrag/OutlookFileDrag.csproj /t:Rebuild /p:Configuration=Release `
  /p:Platform=AnyCPU /p:SignManifests=false /p:VisualStudioVersion=18.0 `
  /p:GenerateManifests=false /nologo /verbosity:minimal /fl `
  '/flp:logfile=artifacts\initial-rebuild.log;verbosity=normal'
```

Dieser Versuch scheiterte vor der Kompilierung an `RemoveOfficeAddInSecurity`/`SetInclusionListEntry`: „Der angeforderte Registrierungszugriff ist unzulässig.“ Er wurde nicht mit erweiterten Rechten wiederholt. Die installierten Office-Buildtargets enthalten außerdem Zertifikatsspeicherverwaltung und Add-in-Registrierung auch außerhalb der C#-Kompilierung. Deshalb wurde der vollständige Standard-VSTO-Build unter den Vorgaben „keine Registry-Eingriffe/keine privaten Schlüssel“ nicht weiter ausgeführt. Der nachfolgende direkte `/t:Compile`-Lauf war mit 0 Fehlern und 0 Warnungen erfolgreich; die sichere Buildmatrix ergänzt ihn um kopierte C#-Ausgaben. Die Protokolle `initial-rebuild.log` und `compile-release.log` bleiben erhalten.

### Testwirksamkeit

```powershell
.\tools\Check-RegressionSensitivity.ps1 `
  -MSBuildPath 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe'
```

Dieser Lauf baut absichtlich veränderte, wegwerfbare Quellkopien unter `artifacts/mutations/`. Das bearbeitete Projekt wird dabei nicht verändert. Insbesondere werden keine ungesicherten nativen Deskriptorzugriffe aus dem alten Code ausgeführt.

| Wieder eingeführte Fehlerkategorie | Durch fehlschlagende Tests erkannt |
|---|---:|
| Zu früh veröffentlichter Exportcache | 4 |
| Fehlende Prüfung der Deskriptorgesamtgröße | 11 |
| Entfernte Namens-/Zielpfadprüfung | 28 |
| Zurückgesetzte Klonposition / fehlerhaftes Überspringen | 5 |
| Auf `int` verengter ILockBytes-Offset | 1 |

Das Prüfskript selbst ist erfolgreich; die Exitcodes 1 der mutierten Testprogramme sind erwartete Nachweise der Testwirksamkeit. Die ursprünglichen, korrigierten Quellen bestehen dagegen alle 114 Testfälle. Dieser Nachweis ist keine Behauptung, dass die komplette alte Version mit denselben neuen internen Test-Schnittstellen gebaut worden sei.

### Paket-, Versions- und Setupkonsistenz

- log4net: Paket 3.5.0, tatsächliche Assembly 3.5.0.0, Token `669e0ddf0bb1aa2a`, `net462`-DLL. Beide Installer referenzieren jetzt diese Identität und den eindeutigen lokalen Paketpfad. Es gibt keine aktiven Projekt-/Installerreferenzen auf 2.0.8.
- EasyHook: Paket und verwaltete Assembly 2.7.7097/2.7.7097.0, Token `4b580fca19d0b0c5`. Die mitgeführten nativen 32-/64-Bit-DLLs haben Version 2.7.7097.0 und stimmen per SHA-256 mit dem Paket überein.
- Die kopierten log4net-/EasyHook-Release-DLLs stimmen per SHA-256 mit den referenzierten Paket-DLLs überein.
- OutlookFileDrag: AssemblyVersion, FileVersion und ApplicationVersion 1.0.13.0; beide MSI-Produktversionen 1.0.13. Die erzeugten C#-DLLs wurden tatsächlich auf ihre Version geprüft; die x86-/x64-Ausgaben zusätzlich auf die PE-Maschine.
- Zusätzlich korrigiert: Die Installer verlangten noch .NET 4.5. Bootstrapper und LaunchCondition verlangen jetzt passend zum Projekt .NET 4.8; auch der aktive Projekt-Bootstrapper und der Verweis auf die [offizielle .NET-4.8-Downloadseite](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48) wurden angepasst.
- ProductCode, UpgradeCode und PackageCode beider Setup-Projekte wurden unverändert erhalten und gegen die Ausgangssicherung geprüft. Die tatsächlichen Tabellen eines neuen MSI wurden nicht geprüft, weil kein neues MSI erstellt wurde. Vor einer neuen Paketfreigabe muss insbesondere ein eindeutiger PackageCode sichergestellt werden; ein geändertes MSI darf nicht mit einem identischen PackageCode wie ein anderes Paket verteilt werden. Siehe [Microsoft: Package Codes](https://learn.microsoft.com/en-us/windows/win32/msi/package-codes).

### Protokolle und Reviewartefakte

Die finalen Ergebnisse liegen in:

- `artifacts/tests-x86.log`, `artifacts/tests-x64.log` und den jeweiligen Testbuildprotokollen;
- `artifacts/managed-build-Release-AnyCPU.log`, `managed-build-Debug-AnyCPU.log`, `managed-build-Release-x86.log`, `managed-build-Release-x64.log`;
- `artifacts/static-analysis.log` und `artifacts/analysis.sarif`;
- `artifacts/consistency-results.json`, `artifacts/validation-commands.json`;
- `artifacts/mutation-results.csv`, `artifacts/mutation-run.log` und den einzelnen Mutationsprotokollen;
- `artifacts/baseline-comparison.csv`, `artifacts/source-changes.patch` und der ursprünglichen Sicherung.

Die CI-Workflowdatei führt künftig zusätzlich den x86- und x64-Testläufer aus. Diese Workflowänderung wurde lokal inhaltlich überprüft; ein entfernter GitHub-Actions-Lauf wurde nicht gestartet. Die Befunde wurden abschließend gegen den lokalen Diff auf Cacheveröffentlichung, HRESULTs, `STGMEDIUM`-Freigabe, RCW-Referenzen, Pointerbreite, Pfadkollisionen und geänderte Installeridentitäten geprüft.

## D. Offene Punkte und Release-Blocker

**Noch bestehende Fehler aus dem Auftrag:** Kein unbehandelter oder weiterhin reproduzierter Fehler der fünf beauftragten Kategorien ist bekannt. Das ersetzt keine praktische Validierung mit Outlook.

**Weitere Risiken/Grenzen:** Bereinigung kann an externen Dateisperren scheitern; solche Dateien bleiben ungültige, nicht veröffentlichte Ergebnisse und werden protokolliert. Sehr große Metadatenauswahlen werden begrenzt. Der bisherige IStorage-Serialisierungsweg benötigt nativen Speicher. Die bereits vorhandene EasyHook-Fallbacklogik, Outlook-COM-Objekte, andere Zielanwendungen und lange Laufzeiten müssen im echten Host geprüft werden. Der historische Zertifikatsbestand wurde weder verwendet noch inhaltlich geprüft.

**Fehlende technische Tests:** Vollständiger VSTO-Build mit neu erzeugten Manifesten, Prüfung vertrauenswürdiger Signaturen, x86-/x64-MSI-Erstellung und Inspektion tatsächlicher MSI-Tabellen/Payloads, Installations-/Upgrade-/Reparatur-/Rollback-/Deinstallationstests und Intune/SYSTEM-Szenarien. Installer-Projects-Komponenten sind lokal vorhanden; ein MSI-Build wurde wegen der nicht sicher erzeugten VSTO-Manifestvoraussetzungen nicht gestartet. Es wurde kein Registry-Workaround eingerichtet. Mehrgigabytegroße reale native Blöcke wurden durch Grenzwertsimulationen ersetzt.

**Notwendige manuelle Outlook-Tests:** Klassisches Outlook in beiden Bitnessvarianten; Outlook 2019 und Microsoft 365, relevante Updatekanäle; einzelne/viele vollständige Mails und Anhänge; Unicode, identische und lange Namen; verschiedene Zielanwendungen; Kopier-/Verschiebeeffekte, Abbruch, mehrfache Abrufe, Dateiinhalte und temporäre Aufräumphase. Diese Funktionalität wird ausdrücklich nicht als praktisch verifiziert bezeichnet.

## E. Anleitung für die nächsten Schritte

1. **Änderungen kontrollieren:** In Visual Studio „Datei → Öffnen → Projekt/Projektmappe“ und `OutlookFileDrag.sln` wählen. Zuerst diesen Bericht lesen. Die fünf genannten C#-Dateien und `NativeMethods.cs` mit den Kopien in `artifacts/baseline-20261014/` vergleichen; alternativ `artifacts/source-changes.patch` in einem Texteditor öffnen. `packages.config` muss weiterhin log4net 3.5.0 enthalten. Die beiden Setup-Projekte haben weiterhin ihre bisherigen ProductCodes/UpgradeCodes. Für spätere Git-Arbeit zuerst den tatsächlichen Repository-Checkout verwenden und dort einen separaten Arbeitsbranch anlegen; keine vorhandenen Änderungen überschreiben.

2. **Automatisierte aktuelle Version bauen:** In Visual Studio „Extras → Befehlszeile → Entwickler-PowerShell“ öffnen und in das Projektverzeichnis wechseln. `.\tools\Validate.ps1` ausführen. Erfolgreich sind zwei Meldungen `114 passed, 0 failed`, vier C#-Builds und die abschließenden Konsistenzprüfungen. Die neue Ausgabe unter `OutlookFileDrag/bin/Release` ist eine kompilierte DLL, noch kein installierbares Release. Bei einem frischen Checkout müssen zuvor die in `packages.config` genannten Pakete wiederhergestellt werden (`nuget restore OutlookFileDrag.sln`).

3. **VSTO und MSI auf einer Testmaschine erstellen:** Eine eigene VM oder einen entbehrlichen Entwicklungs-PC mit klassischem Outlook verwenden, weil die normalen VSTO-Buildziele Registrierungen ändern können. .NET-4.8-Targeting-Pack, Office-/SharePoint-Entwicklung und Installer-Projects-Erweiterung installieren. Release/Any CPU auswählen und den vollständigen VSTO-Build dort mit geeigneten, außerhalb des Repositorys verwalteten Signaturdaten durchführen. Beide neuen Manifeste prüfen. Danach beide Setup-Projekte per Rechtsklick „Erstellen“ bauen. ProductVersion, Bitness, neue Paketidentität, erhaltenen UpgradeCode, Upgradeverhalten und die tatsächlich eingebetteten Versionen 1.0.13.0/3.5.0.0/2.7.7097.0 kontrollieren. Auch ProductCodes gegen bereits veröffentlichte Installerversionen vergleichen. Hierzu wurde noch kein Erfolg nachgewiesen.

4. **MSI-Installation sicher testen:** Vorher einen VM-Snapshot erstellen; keine produktiven Postfächer verwenden. Outlook schließen, den zur Outlook-Bitness passenden Installer installieren und die Installation protokollieren, zum Beispiel `msiexec /i "C:\Test\OutlookFileDragSetup_x64.msi" /L*v "C:\Test\install.log"`. Zunächst eine sichtbare Installation verwenden. Saubere Installation, Upgrade von 1.0.11 und 1.0.12-fixed, Reparatur und Deinstallation jeweils mit zurücksetzbarem Zustand prüfen. Erst danach unbeaufsichtigte/SYSTEM-Installation und Rollback testen. Versionen im Installationsordner und Add-in-Ladestatus prüfen; MSI-Installation nicht mit dem erfolgreichen C#-Build gleichsetzen.

5. **Drag-and-drop in klassischem Outlook prüfen:** Outlook manuell starten. Unter „Datei → Optionen → Add-Ins“ kontrollieren, dass OutlookFileDrag aktiv ist. Mit unkritischen Testmails je eine vollständige Nachricht, einen Anhang und eine Mehrfachauswahl in den Explorer, einen kompatiblen Uploaddialog und eine weitere Zielanwendung ziehen. Exportierte MSG-Dateien wieder öffnen und Inhalte/Anhänge prüfen; Anhänge möglichst mit dem gespeicherten Original per Hash vergleichen. Unicode, gleiche Namen, lange Namen, Abbruch und wiederholte Versuche testen. Nach einem als Verschieben dargestellten Export muss die ursprüngliche Mail erhalten bleiben. `%TEMP%\OutlookFileDrag` und `%APPDATA%\OutlookFileDrag\OutlookFileDrag.log` kontrollieren. Tests in x86- und x64-Outlook sowie während der Aufräumphase wiederholen. Protokolle vor externer Weitergabe von Kundendaten bereinigen.

6. **Vor öffentlichem Release:** Beide vollständigen VSTO-/MSI-Builds und vertrauenswürdige Signaturen müssen geprüft sein; sämtliche manuellen Drag-and-drop- und Installations-/Upgradefälle müssen bestehen. Die neuen MSI-Dateien müssen die aktuellen DLLs enthalten und korrekte, eindeutige Installeridentitäten besitzen. Danach einen kleinen Pilotkreis beobachten. Erst nach dieser Freigabe veröffentlichen oder breiter verteilen. Push, Pull Request und Veröffentlichung wurden in diesem Auftrag nicht durchgeführt.

## Vertragsreferenzen

Die Implementierung wurde mit den primären Microsoft-Verträgen abgeglichen: [IEnumFORMATETC.Next](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ienumformatetc-next), [Skip](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ienumformatetc-skip), [Clone](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ienumformatetc-clone), [FILEDESCRIPTORW](https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/ns-shlobj_core-filedescriptorw), [ILockBytes.ReadAt](https://learn.microsoft.com/en-us/windows/win32/api/objidl/nf-objidl-ilockbytes-readat), [ReleaseStgMedium](https://learn.microsoft.com/en-us/windows/win32/api/ole2/nf-ole2-releasestgmedium) und [Windows-Dateinamen](https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file).

## Nachtrag: Validierung des GitHub-Transfers

Die Änderungen wurden in einem separaten frischen Clone auf Basis von `master` (`4c7b1d88ae4be6878b2ab876fb34cb5c55a41b78`) in den Branch `fix/audit-reliability-1.0.13` übertragen. Der ursprüngliche lokale Quellstand wurde nicht überschrieben. Der Vergleich mit dem lokalen Ausgangsstand zeigte nur die bereits lokale log4net-Aktualisierung und die zugehörigen Projektreferenzen als Vorababweichung gegenüber GitHub; die bestehenden GitHub-Korrekturen bleiben erhalten. Die explizite log4net-Kopieroption `Private=True` aus dem GitHub-Projekt wurde beibehalten.

Im Clone wurden erneut 114 Tests in x86 und 114 in x64, sämtliche vier verwalteten Buildvarianten, die statische Analyse und alle 33 Konsistenzprüfungen erfolgreich ausgeführt. Die Installeridentitäten wurden dazu gegen den frisch geklonten GitHub-Ausgangsstand geprüft. Auch die fünf Mutationsprüfungen wurden im Clone wiederholt. Der CI-Workflow enthält die x86-/x64-Regressionsläufe; sein tatsächlicher Status wird am Draft-Pull-Request ausgewiesen.

Übertragen werden ausschließlich die geprüften Quellen, Tests, Skripte, Paket-/Installerkonfigurationen und Dokumentationen. Pakete, native Binärdateien, Buildausgaben, Protokolle, lokale Sicherungen und Zertifikatsdateien werden nicht als Änderungen übernommen. Bestehende Zertifikatsdateien bleiben unverändert. Die Einschränkungen zu vollständigem VSTO-/MSI-Build und echten Outlook-Tests gelten weiterhin.