# Gesamtes PowerShell-Skript zur Analyse von .LA1.txt Logdateien und Erstellung einer Statistik-XLSX

# --- Konfiguration ---
# Speicherort der Logdateien und der Ausgabedatei festlegen
# PASSE DIESE PFADE AN DEINE UMGEBUNG AN!
$logFilesPath = "C:\Users\top\git\LA1Analyzer\Samples\AM"
$outputStatsXlsxFile = "C:\Users\top\git\LA1Analyzer\Samples\AM\Gesamtauswertung_Statistik_Logfiles.xlsx" # Dateiendung geändert zu .xlsx

Write-Host "Daten des Ordners '$logFilesPath' werden analysiert." -ForegroundColor Green

# Trennzeichen für die CSV-Datei (wird für Export-Excel nicht direkt benötigt, aber beibehalten)
$csvDelimiter = ";"

# Wichtig: Kultur-Info für die korrekte Dezimaltrennzeichen-Formatierung
# Für Deutschland (Komma als Dezimaltrennzeichen):
$cultureInfo = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
# Wenn du einen Punkt als Dezimaltrennzeichen beibehalten möchtest (z.B. für englische Excel-Version):
# $cultureInfo = [System.Globalization.CultureInfo]::InvariantCulture

# Liste der erwarteten Fehlercodes für die Statistikspalten
# Diese Liste bestimmt, welche ERR_Codes in der XLSX-Datei eigene Spalten erhalten.
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

# --- Überprüfen und Installieren des ImportExcel Moduls ---
Write-Host "Überprüfe, ob das 'ImportExcel' Modul installiert ist..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Warning "Das 'ImportExcel' Modul ist nicht installiert. Versuche, es zu installieren."
    try {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "Das 'ImportExcel' Modul wurde erfolgreich installiert." -ForegroundColor Green
    }
    catch {
        Write-Error "Fehler beim Installieren des 'ImportExcel' Moduls: $($_.Exception.Message)"
        Write-Error "Bitte installieren Sie es manuell mit: Install-Module -Name ImportExcel -Scope CurrentUser"
        return # Skript beenden, wenn das Modul nicht installiert werden kann
    }
} else {
    Write-Host "Das 'ImportExcel' Modul ist bereits installiert." -ForegroundColor Green
}

# Importiere das Modul, um seine Funktionen nutzen zu können
Import-Module -Name ImportExcel -ErrorAction Stop

# Leere Liste für alle gesammelten Statistikobjekte aus den Logdateien
$allStatsObjects = @()

# --- Hauptverarbeitung: Schleife durch alle .LA1.txt-Dateien ---
Get-ChildItem -Path $logFilesPath -Filter "*.LA1.txt" | ForEach-Object {
    $currentFile = $_.FullName
    Write-Host "`nVerarbeite Statistikdatei: $($_.Name)" -ForegroundColor Green

    # Lese den gesamten Inhalt der Logdatei in ein Array von Zeilen
    $fileContent = Get-Content -Path $currentFile -ErrorAction SilentlyContinue # Fängt Fehler ab, falls Datei nicht lesbar

    # Überspringe Datei, wenn Inhalt nicht gelesen werden konnte
    if (-not $fileContent) {
        Write-Warning "Konnte Inhalt der Datei '$($_.Name)' nicht lesen. Überspringe diese Datei."
        return
    }

    # Initialisiere ein neues Objekt für die aktuellen Statistikdaten
    $currentStats = [PSCustomObject]@{
        Datum               = ""
        Performance_AVG     = ""
        Performance_PEAK    = ""
        Performance_Potential = ""
        Performance_HC      = ""
        Plate_Production    = ""
        Exposed_Plates      = ""
        Damaged_Plates      = ""
        Plate_Count         = ""
    }

    # Initialisiere alle zu verfolgenden Fehlercodes mit 0 im aktuellen Objekt
    foreach ($code in $errorCodesToTrack) {
        # Entferne "ERR_" und füge ein "E_" Präfix hinzu für den Eigenschaftsnamen in der XLSX
        # Dies behebt den "Add-Member" Fehler, da rein numerische Namen Probleme verursachen können
        $propertyName = "E_" + $code.Replace('ERR_','')
        Add-Member -InputObject $currentStats -NotePropertyName $propertyName -NotePropertyValue 0
    }

    # --- Datum extrahieren ---
    # Finde die Zeile mit Select-String
    $dateLineObj = $fileContent | Select-String -Pattern "^\s*LEVEL 1 : DAILY RESULTS : (.+?)\s*$"
    if ($dateLineObj) {
        # Wende den Regex auf die gefundene Zeile an, um $Matches zu füllen
        if ($dateLineObj.Line -match "^\s*LEVEL 1 : DAILY RESULTS : (.+?)\s*$") {
            $currentStats.Datum = ($Matches[1]).Trim()
        }
    } else {
        # Write-Warning "  DEBUG: Konnte Datum in Datei '$($_.Name)' nicht finden."
    }

    # --- Performance-Daten extrahieren ---
    $performanceLineObj = $fileContent | Select-String -Pattern "^\s*Performance: AVG:\s*(\d+\.?\d*|No Information),\s*PEAK:\s*(\d+\.?\d*|No Information),\s*Potential:\s*(\d+\.?\d*|No Information)\s*-\s*hc\s*(\d+\.?\d*%)\s*$"
    if ($performanceLineObj) {
        $performanceLine = $performanceLineObj.Line # Holen Sie sich die tatsächliche String-Zeile
        
        # Wende den Regex auf die gefundene Zeile an, um $Matches zu füllen
        if ($performanceLine -match "^\s*Performance: AVG:\s*(\d+\.?\d*|No Information),\s*PEAK:\s*(\d+\.?\d*|No Information),\s*Potential:\s*(\d+\.?\d*|No Information)\s*-\s*hc\s*(\d+\.?\d*%)\s*$") {
            # Überprüfen und parsen für AVG
            $rawAvgValue = $Matches[1].Trim()
            $parsedAvg = 0.0 # Standardwert
            if ($rawAvgValue -ne "No Information" -and [double]::TryParse($rawAvgValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedAvg)) {
                $currentStats.Performance_AVG = $parsedAvg.ToString($cultureInfo)
            } else {
                $currentStats.Performance_AVG = "" # Oder 0.0, je nach gewünschtem Verhalten
            }

            # Überprüfen und parsen für PEAK
            $rawPeakValue = $Matches[2].Trim()
            $parsedPeak = 0.0 # Standardwert
            if ($rawPeakValue -ne "No Information" -and [double]::TryParse($rawPeakValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedPeak)) {
                $currentStats.Performance_PEAK = $parsedPeak.ToString($cultureInfo)
            } else {
                $currentStats.Performance_PEAK = ""
            }

            # Überprüfen und parsen für Potential
            $rawPotentialValue = $Matches[3].Trim()
            $parsedPotential = 0.0 # Standardwert
            if ($rawPotentialValue -ne "No Information" -and [double]::TryParse($rawPotentialValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedPotential)) {
                $currentStats.Performance_Potential = $parsedPotential.ToString($cultureInfo)
            } else {
                $currentStats.Performance_Potential = ""
            }
            
            # HC ist immer ein String (Prozent), keine Notwendigkeit zu parsen
            $currentStats.Performance_HC = $Matches[4].Trim()
        } else {
            # Write-Warning "    DEBUG: Regex-Match für Performance-Werte in Zeile '$($performanceLine)' fehlgeschlagen. Werte bleiben leer. (Matches-Variable nicht gefüllt)"
        }
    } else {
        # Write-Warning "  DEBUG: 'Performance:' Zeile in Datei '$($_.Name)' nicht gefunden."
    }

    # --- Plate Production Daten extrahieren ---
    $plateProductionLineObj = $fileContent | Select-String -Pattern "^\s*Plate Production:\s*(\d+)\s*,\s*Exposed Plates:\s*(\d+)\s*,\s*Damaged Plates:\s*(\d+)\s*$"
    if ($plateProductionLineObj) {
        $plateProductionLine = $plateProductionLineObj.Line # Holen Sie sich die tatsächliche String-Zeile
        
        # Wende den Regex auf die gefundene Zeile an, um $Matches zu füllen
        if ($plateProductionLine -match "^\s*Plate Production:\s*(\d+)\s*,\s*Exposed Plates:\s*(\d+)\s*,\s*Damaged Plates:\s*(\d+)\s*$") {
            $parsedProd = 0 # Standardwert
            if ([int]::TryParse($Matches[1].Trim(), [ref]$parsedProd)) {
                $currentStats.Plate_Production = $parsedProd
            } else {
                $currentStats.Plate_Production = 0
            }

            $parsedExp = 0 # Standardwert
            if ([int]::TryParse($Matches[2].Trim(), [ref]$parsedExp)) {
                $currentStats.Exposed_Plates = $parsedExp
            } else {
                $currentStats.Exposed_Plates = 0
            }

            $parsedDam = 0 # Standardwert
            if ([int]::TryParse($Matches[3].Trim(), [ref]$parsedDam)) {
                $currentStats.Damaged_Plates = $parsedDam
            } else {
                $currentStats.Damaged_Plates = 0
            }
        } else {
            # Write-Warning "    DEBUG: Regex-Match für Plate Production in Zeile '$($plateProductionLine)' fehlgeschlagen. Werte bleiben 0."
        }
    } else {
        # Write-Warning "  DEBUG: 'Plate Production:' Zeile in Datei '$($_.Name)' nicht gefunden."
    }

    # --- Plate Count Daten extrahieren ---
    $plateCountLineObj = $fileContent | Select-String -Pattern "^\s*Plate Count:\s*(\d+)\s*$"
    if ($plateCountLineObj) {
        $plateCountLine = $plateCountLineObj.Line # Holen Sie sich die tatsächliche String-Zeile
        $parsedCount = 0 # Standardwert
        # Wende den Regex auf die gefundene Zeile an, um $Matches zu füllen
        if ($plateCountLine -match "^\s*Plate Count:\s*(\d+)\s*$") {
            if ([int]::TryParse($Matches[1].Trim(), [ref]$parsedCount)) {
                $currentStats.Plate_Count = $parsedCount
            } else {
                $currentStats.Plate_Count = 0
            }
        } else {
            # Write-Warning "    DEBUG: Regex-Match für Plate Count in Zeile '$($plateCountLine)' fehlgeschlagen. Wert bleibt 0."
        }
    } else {
        # Write-Warning "  DEBUG: 'Plate Count:' Zeile in Datei '$($_.Name)' nicht gefunden."
    }

    # --- Message Statistics Daten extrahieren ---
    # Finde den Start- und Endpunkt des "Message Statistics" Blocks
    $startIndex = -1
    $endIndex = -1

    for ($i = 0; $i -lt $fileContent.Length; $i++) {
        if ($fileContent[$i] -match "^\s*Message Statistics:\s*$") {
            $startIndex = $i
        } elseif ($startIndex -ne -1 -and $fileContent[$i] -match "^\s*--------------------------------------------------------------------------------------------------------------------------\s*$") {
            # Der zweite Trennstrich nach dem Start des Blocks markiert das Ende
            if ($i -gt $startIndex + 1) { # Stelle sicher, dass es nach dem Header-Trennstrich kommt
                $endIndex = $i
                break
            }
        }
    }

    if ($startIndex -ne -1 -and $endIndex -ne -1) {
        # Write-Host "  DEBUG: Message Statistics Block gefunden. Verarbeite Datenzeilen." -ForegroundColor DarkCyan
        # Extrahiere den relevanten Block (Datenzeilen zwischen den Trennstrichen)
        $messageStatsBlock = $fileContent[($startIndex + 2)..($endIndex - 1)]

        foreach ($line in $messageStatsBlock) {
            # Regex, um Anzahl und Fehlercode aus jeder Statistikzeile zu extrahieren
            # Hier verwenden wir direkt den -match Operator auf der Zeile, da $line bereits ein String ist
            if ($line -match "^\s*(\d+)\s*\|.*?\|(ERR_\d+).*$") {
                $count = 0 # Standardwert
                $errorCode = $Matches[2].Trim() # Fehlercode als String

                if ([int]::TryParse($Matches[1].Trim(), [ref]$count)) { # Zähler als Integer
                    # Wenn der Fehlercode in unserer Liste ist, aktualisiere den Zähler
                    if ($errorCodesToTrack -contains $errorCode) {
                        # Verwende den bereinigten Eigenschaftsnamen (mit "E_" Präfix)
                        $propertyName = "E_" + $errorCode.Replace('ERR_','')
                        $currentStats.$propertyName = $count
                    } else {
                        # Write-Host "      DEBUG: Fehlercode '$errorCode' nicht in 'errorCodesToTrack' Liste." -ForegroundColor DarkGray
                    }
                } else {
                    # Write-Host "      DEBUG: Konnte Zähler für Fehlercode '$errorCode' nicht parsen. Wert bleibt 0." -ForegroundColor DarkRed
                }
            } else {
                # Write-Host "    DEBUG: Zeile im Message Statistics Block passt nicht zum Muster: '$line'" -ForegroundColor DarkGray
            }
        }
    } else {
        Write-Warning "Message Statistics Block in Datei '$($_.Name)' nicht gefunden oder unvollständig. Fehlerzählungen bleiben 0."
    }

    # Füge das gesammelte Objekt zur Liste hinzu
    $allStatsObjects += $currentStats
    # Write-Host "  DEBUG: Hinzugefügtes Gesamt-Statistik-Objekt für '$($_.Name)'." -ForegroundColor Green
    # Optional: Zeigen Sie das Objekt an, um zu überprüfen, ob es Daten enthält
    # $currentStats | Format-List
}

# --- Statistik-XLSX-Datei schreiben ---
Write-Host "`n--- Schreibe Statistik-XLSX-Datei ---" -ForegroundColor Yellow

# Exportiere alle gesammelten Objekte in die XLSX-Datei
# Export-Excel kümmert sich um Kopfzeile, Formatierung etc.
$allStatsObjects | Export-Excel -Path $outputStatsXlsxFile -AutoSize -ClearSheet

Write-Host "-------------------------------------" -ForegroundColor Green
Write-Host "Statistik-Verarbeitung abgeschlossen!" -ForegroundColor Green
Write-Host "Die Daten wurden in '$outputStatsXlsxFile' gespeichert." -ForegroundColor Green
Write-Host "Gesamtzahl der Statistik-Einträge: $($allStatsObjects.Count)" -ForegroundColor Green

