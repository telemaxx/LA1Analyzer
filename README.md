# LA1Analyzer

Ein sehr spezielles Tool, das für die meisten Benutzer nicht von Interesse ist. Nachfolgend einige Hinweise auf Deutsch.

### Wichtiger Hinweis zur Ausführung

Zurzeit sind die PowerShell-Skripte nicht signiert. Daher können sie nicht direkt per Doppelklick ausgeführt werden.

Das bedeutet folgenden Ablauf:

Die `ps1`-Datei muss zuerst lokal gespeichert werden

Dann die Datei mit einem Editor öffnen

Die Pfade anpassen und wieder speichern

```powershell
$logFilesPath = "C:\Users\top\git\LA1Analyzer\Samples\AM"
$outputStatsXlsxFile = "C:\Users\top\git\LA1Analyzer\Samples\AM\Gesamtauswertung_Statistik_Logfiles.xlsx"
```
Den gesammten Inhalten makieren und kopieren (Strg A Strg C)
In die Powershell wechseln und (Strg V) eingeben. Warnung bestätigen.

