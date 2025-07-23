# LA1Analyzer

Ein sehr spezielles Tool, das für die meisten Benutzer nicht von Interesse ist. Nachfolgend einige Hinweise auf Deutsch.

### Wichtiger Hinweis zur Ausführung

Zurzeit sind die PowerShell-Skripte nicht signiert. Daher können sie nicht direkt per Doppelklick ausgeführt werden.

Das bedeutet, die `ps1`-Datei muss zuerst lokal gespeichert und anschließend müssen die Pfade in der Datei angepasst werden.

**Beispiel:**

```powershell
$logFilesPath = "C:\Users\top\git\LA1Analyzer\Samples\AM"
$outputStatsXlsxFile = "C:\Users\top\git\LA1Analyzer\Samples\AM\Gesamtauswertung_Statistik_Logfiles.xlsx"
```

Dann die Datei mit einem Editor öffnen den gesammten Inhalten makieren und kopieren (Strg A Strg C)

Jetzt in der Powershell einfügen (Strg V)
Warnung bestätigen.

