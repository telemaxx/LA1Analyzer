# Gesamtes PowerShell-Skript zur Analyse von .LA1.txt Logdateien und Erstellung einer Statistik-CSV

# --- Konfiguration ---
# Speicherort der Logdateien und der Ausgabedatei festlegen
# PASSE DIESE PFADE AN DEINE UMGEBUNG AN!
$logFilesPath = "C:\Users\top\git\LA1Analyzer\Samples\AM"
$outputStatsCsvFile = "C:\Users\top\git\LA1Analyzer\Samples\AM\Gesamtauswertung_Statistik_Logfiles.csv"

Write-Host "Daten des Ordners '$logFilesPath' werden analisiert." -ForegroundColor Green


# Trennzeichen für die CSV-Datei
$csvDelimiter = ";"

# Wichtig: Kultur-Info für die korrekte Dezimaltrennzeichen-Formatierung
# Für Deutschland (Komma als Dezimaltrennzeichen):
$cultureInfo = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
# Wenn du einen Punkt als Dezimaltrennzeichen beibehalten möchtest (z.B. für englische Excel-Version):
# $cultureInfo = [System.Globalization.CultureInfo]::InvariantCulture

# Liste der erwarteten Fehlercodes für die Statistikspalten
# Diese Liste bestimmt, welche ERR_Codes in der CSV-Datei eigene Spalten erhalten.
# Aktualisiere diese Liste basierend auf den Fehlern, die du verfolgen möchtest.
$errorCodesToTrack = @(
    "ERR_00000", "ERR_00001", "ERR_00002", "ERR_00322", "ERR_00323",
    "ERR_00460", "ERR_04751", "ERR_04758", "ERR_04760", "ERR_04761",
    "ERR_04773", "ERR_04818", "ERR_04824", "ERR_05010", "ERR_05013",
    "ERR_05029", "ERR_05073", "ERR_05079", "ERR_05086", "ERR_05127",
    "ERR_05354", "ERR_05360", "ERR_05366", "ERR_05413", "ERR_05439",
    "ERR_05454", "ERR_06165", "ERR_06327", "ERR_06433", "ERR_06456",
    "ERR_06461", "ERR_06474", "ERR_06483", "ERR_06484", "ERR_06485",
    "ERR_06486", "ERR_06487", "ERR_06495", "ERR_06502", "ERR_06503",
    "ERR_06505", "ERR_06601", "ERR_07504", "ERR_07602", "ERR_07609",
    "ERR_07610", "ERR_07616", "ERR_07617", "ERR_07619", "ERR_07654",
    "ERR_07656", "ERR_07951", "ERR_08003", "ERR_08004", "ERR_08007",
    "ERR_08009", "ERR_08105", "ERR_08107", "ERR_08109", "ERR_08110",
    "ERR_08111", "ERR_08203", "ERR_08214", "ERR_08215", "ERR_08216",
    "ERR_08242", "ERR_08243", "ERR_08244", "ERR_08245", "ERR_08246",
    "ERR_08301", "ERR_08302", "ERR_08303", "ERR_08304", "ERR_08305",
    "ERR_08311", "ERR_08312", "ERR_08314", "ERR_08318", "ERR_08326",
    "ERR_08330", "ERR_08337", "ERR_10462", "ERR_10830", "ERR_10853",
    "ERR_10855", "ERR_10865", "ERR_10866", "ERR_10901", "ERR_10905",
    "ERR_10906", "ERR_10908", "ERR_10910", "ERR_10913", "ERR_10914",
    "ERR_10919", "ERR_10921", "ERR_10924", "ERR_11000", "ERR_11046"
)

# --- Dynamische Erstellung der Kopfzeile für die Statistik-CSV ---
$headerStats = "Datum;Performance_AVG;Performance_PEAK;Performance_Potential;Performance_HC;Plate_Production;Exposed_Plates;Damaged_Plates;Plate_Count"
foreach ($code in $errorCodesToTrack) {
    # Hinzufügen der Fehlercodes (ohne "ERR_") zur Kopfzeile
    $headerStats += "$($csvDelimiter)$($code.Replace('ERR_',''))"
}

# Leere Liste für alle gesammelten Statistikdaten aus den Logdateien
$allStatsEntries = @()

# --- Hauptverarbeitung: Schleife durch alle .LA1.txt-Dateien ---
Get-ChildItem -Path $logFilesPath -Filter "*.LA1.txt" | ForEach-Object {
    $currentFile = $_.FullName
    Write-Host "`nVerarbeite Statistikdatei: $($_.Name)" -ForegroundColor Green

    $fileContent = Get-Content -Path $currentFile -ErrorAction SilentlyContinue # Fängt Fehler ab, falls Datei nicht lesbar

    # Überspringe Datei, wenn Inhalt nicht gelesen werden konnte
    if (-not $fileContent) {
        Write-Warning "Konnte Inhalt der Datei '$($_.Name)' nicht lesen. Überspringe diese Datei."
        return
    }

    # --- Datum extrahieren ---
    $dateLine = $fileContent | Select-String -Pattern "LEVEL 1 : DAILY RESULTS :"

    $logFileDate = ""
    if ($dateLine) {
        $logFileDate = ($dateLine.ToString().Split(":")[-1]).Trim()
#        Write-Host "  -> Extrahiertes Datum: $logFileDate"
    } else {
#        Write-Warning "Konnte Datum in Datei '$($_.Name)' nicht finden. Überspringe diese Datei."
        return
    }

    # --- Performance-Daten extrahieren ---
    # Initialisiere Variablen für den aktuellen Datei-Durchlauf
    $avg = ""
    $peak = ""
    $potential = ""
    $hc = ""

    # Suche nach der Performance-Zeile
    $performanceLine = $fileContent | Select-String -Pattern "Performance: AVG:"

    if ($performanceLine) {
        # Angepasstes Pattern für robustere Erkennung von Leerzeichen und "No Information"
        # Es erlaubt "No Information" oder eine Zahl mit optionalem Dezimalteil.
        # HC-Teil wurde angepasst, um "zahl%" oder "zahl.zahl%" zu erkennen.
        $performancePattern = "AVG:\s*(\d+\.?\d*|No Information),\s*PEAK:\s*(\d+\.?\d*|No Information),\s*Potential:\s*(\d+\.?\d*|No Information)\s*-\s*hc\s*(\d+\.?\d*%)"
        $match = $performanceLine.ToString() | Select-String -Pattern $performancePattern

        if ($match) {
            # Überprüfen und parsen für AVG
            $rawAvgValue = $match.Matches[0].Groups[1].Value
            if ($rawAvgValue -ne "No Information" -and $rawAvgValue -notlike "*No Information*") {
                [double]$rawAvg = $rawAvgValue
                $avg = $rawAvg.ToString($cultureInfo)
            } else {
                $avg = "" # Bleibt leer, wenn "No Information"
            }

            # Überprüfen und parsen für PEAK
            $rawPeakValue = $match.Matches[0].Groups[2].Value
            if ($rawPeakValue -ne "No Information" -and $rawPeakValue -notlike "*No Information*") {
                [double]$rawPeak = $rawPeakValue
                $peak = $rawPeak.ToString($cultureInfo)
            } else {
                $peak = "" # Bleibt leer, wenn "No Information"
            }

            # Überprüfen und parsen für Potential
            $rawPotentialValue = $match.Matches[0].Groups[3].Value
            if ($rawPotentialValue -ne "No Information" -and $rawPotentialValue -notlike "*No Information*") {
                [double]$rawPotential = $rawPotentialValue
                $potential = $rawPotential.ToString($cultureInfo)
            } else {
                $potential = "" # Bleibt leer, wenn "No Information"
            }
            
            # HC ist immer ein String (Prozent), keine Notwendigkeit zu parsen
            $hc = $match.Matches[0].Groups[4].Value
            
#            Write-Host "    -> Performance gefunden und extrahiert: AVG=$avg, PEAK=$peak, Potential=$potential, HC=$hc" -ForegroundColor Cyan
        } else {
            Write-Warning "    -> Regex-Match für Performance-Werte in Zeile '$($performanceLine.ToString().Trim())' fehlgeschlagen. Werte bleiben leer. (Möglicherweise unpassendes Format)"
        }
    } else {
#        Write-Warning "    -> 'Performance: AVG:' Zeile in Datei '$($_.Name)' nicht gefunden. Performance-Werte bleiben leer."
    }

    # --- Plate Production Daten extrahieren ---
    $plateProductionLine = $fileContent | Select-String -Pattern "Plate Production:"
    $plateProduction = ""
    $exposedPlates = ""
    $damagedPlates = ""

    if ($plateProductionLine) {
        # Angepasstes Regex für robusteres Parsen von Leerzeichen
        $match = $plateProductionLine.ToString() | Select-String -Pattern "Plate Production:\s*(\d+)\s*,\s*Exposed Plates:\s*(\d+)\s*,\s*Damaged Plates:\s*(\d+)"
        if ($match) {
            $plateProduction = $match.Matches[0].Groups[1].Value
            $exposedPlates = $match.Matches[0].Groups[2].Value
            $damagedPlates = $match.Matches[0].Groups[3].Value
#            Write-Host "    -> Plate Production gefunden: Prod=$plateProduction, Exp=$exposedPlates, Dam=$damagedPlates"
        } else {
#            Write-Warning "    -> Regex-Match für Plate Production in Zeile '$($plateProductionLine.ToString().Trim())' fehlgeschlagen. (Möglicherweise unpassendes Format)"
        }
    }

    # --- Plate Count Daten extrahieren ---
    $plateCountLine = $fileContent | Select-String -Pattern "Plate Count:"
    $plateCount = ""

    if ($plateCountLine) {
        $match = $plateCountLine.ToString() | Select-String -Pattern "Plate Count:\s*(\d+)"
        if ($match) {
            $plateCount = $match.Matches[0].Groups[1].Value
#            Write-Host "    -> Plate Count gefunden: Count=$plateCount"
        } else {
#            Write-Warning "    -> Regex-Match für Plate Count in Zeile '$($plateCountLine.ToString().Trim())' fehlgeschlagen."
        }
    }

    # --- Message Statistics Daten extrahieren ---
    $messageCounts = @{}
    # Initialisiere alle zu verfolgenden Fehlercodes mit 0
    foreach ($code in $errorCodesToTrack) {
        $messageCounts[$code] = 0
    }

    $startIndex = -1
    $endIndex = -1

    # Finde den Start- und Endpunkt des "Message Statistics" Blocks
    for ($i = 0; $i -lt $fileContent.Length; $i++) {
        if ($fileContent[$i] -match "^\s*Message Statistics:") {
            $startIndex = $i
        } elseif ($startIndex -ne -1 -and $fileContent[$i] -match "^\s*--------------------------------------------------------------------------------------------------------------------------") {
            # Der zweite Trennstrich nach dem Start des Blocks markiert das Ende
            if ($i -gt $startIndex + 1) { # Stelle sicher, dass es nach dem Header-Trennstrich kommt
                $endIndex = $i
                break
            }
        }
    }

    if ($startIndex -ne -1 -and $endIndex -ne -1) {
        # Extrahiere den relevanten Block
        # +1 für den Header-Trennstrich, -1 für den End-Trennstrich, um nur die Datenzeilen zu erhalten
        $messageStatsBlock = $fileContent[($startIndex + 2)..($endIndex - 1)]
#        Write-Host "    -> Message Statistics Block gefunden. Verarbeite $($messageStatsBlock.Count) Datenzeilen."

        foreach ($line in $messageStatsBlock) {
            # Regex, um Anzahl und Fehlercode aus jeder Statistikzeile zu extrahieren
            # ^\s*(\d+)\s*\|.*?\|(ERR_\d+).*
            # (\d+) - Zähler (Gruppe 1)
            # (ERR_\d+) - Fehlercode (Gruppe 2)
            $match = $line | Select-String -Pattern "^\s*(\d+)\s*\|.*?\|(ERR_\d+).*"
            if ($match) {
                $count = [int]$match.Matches[0].Groups[1].Value # Zähler als Integer
                $errorCode = $match.Matches[0].Groups[2].Value # Fehlercode als String
                
                # Wenn der Fehlercode in unserer Liste ist, aktualisiere den Zähler
                if ($errorCodesToTrack -contains $errorCode) {
                    $messageCounts[$errorCode] = $count
#                    Write-Host "      -> Statistik: '$errorCode' = '$count'" -ForegroundColor DarkYellow
                }
            }
        }
    } else {
#        Write-Warning "    -> Message Statistics Block in Datei '$($_.Name)' nicht gefunden oder unvollständig. Fehlerzählungen bleiben 0."
    }

    # --- Daten für die aktuelle Datei in CSV-Format bringen und zur Liste hinzufügen ---
    $formattedStatsLine = "$logFileDate$csvDelimiter$avg$csvDelimiter$peak$csvDelimiter$potential$csvDelimiter$hc$csvDelimiter$plateProduction$csvDelimiter$exposedPlates$csvDelimiter$damagedPlates$csvDelimiter$plateCount"
    
    # Füge die Zählungen für die Fehlercodes hinzu
    foreach ($code in $errorCodesToTrack) {
        $formattedStatsLine += "$($csvDelimiter)$($messageCounts[$code])"
    }

    $allStatsEntries += $formattedStatsLine
#    Write-Host "  -> Hinzugefügte Gesamt-Statistik-Zeile für '$($_.Name)':`n     $formattedStatsLine" -ForegroundColor Green
}

# --- Statistik-CSV-Datei schreiben ---
#Write-Host "`n--- Schreibe Statistik-CSV-Datei ---" -ForegroundColor Yellow

# Zuerst die Kopfzeile schreiben (überschreibt die Datei, falls sie existiert)
$headerStats | Out-File -FilePath $outputStatsCsvFile -Encoding UTF8

# Dann die gesammelten Datenzeilen anhängen
$allStatsEntries | Out-File -FilePath $outputStatsCsvFile -Encoding UTF8 -Append
Write-Host "-------------------------------------" -ForegroundColor Green
Write-Host "Statistik-Verarbeitung abgeschlossen!" -ForegroundColor Green
Write-Host "Die Daten wurden in '$outputStatsCsvFile' gespeichert." -ForegroundColor Green
Write-Host "Gesamtzahl der Statistik-Einträge: $($allStatsEntries.Count)" -ForegroundColor Green
