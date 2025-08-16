# Gesamtes PowerShell-Skript zur Analyse von .LA1.txt Logdateien und Erstellung einer Statistik-XLSX
# Version: Nur Datenerfassung und Export ohne Diagrammerstellung.

# --- System- und Modulvoraussetzungen am Anfang laden ---
# Ben�tigt f�r GUI-Elemente (einmalig am Anfang laden)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Konfiguration ---
# Standardwert f�r "Performance_Potential", falls in der Logdatei "No Information" steht.
$defaultPotentialValue = 999 

Write-Host "Standardwert f�r 'Potential' ist auf '$defaultPotentialValue' gesetzt." -ForegroundColor Cyan

# --- Pfad f�r Logdateien festlegen (immer �ber interaktive GUI) ---
$logFilesPath = $null

do {
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    # Setze den Dialog-Titel f�r mehr Klarheit
    $browser.Description = "Bitte w�hlen Sie den Ordner mit den .LA1.txt Logdateien aus."
    
    $result = $browser.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $logFilesPath = $browser.SelectedPath.Trim()
    } else {
        # Benutzer hat den Dialog abgebrochen
        Write-Error "Ordnerauswahl abgebrochen. Skript wird beendet."
        return
    }

    if (-not (Test-Path -Path $logFilesPath -PathType Container)) {
        Write-Warning "Der angegebene Pfad '$logFilesPath' existiert nicht oder ist kein Verzeichnis. Bitte versuchen Sie es erneut."
        $logFilesPath = "" # Setzt den Pfad zur�ck, damit die Schleife wiederholt wird
    }

} while ([string]::IsNullOrEmpty($logFilesPath))

Write-Host "Verwende interaktiv ausgew�hlten Pfad: '$logFilesPath'" -ForegroundColor Green

Write-Host "Daten des Ordners '$logFilesPath' werden analysiert." -ForegroundColor Green


# Wichtig: Kultur-Info f�r die korrekte Dezimaltrennzeichen-Formatierung
# F�r Deutschland (Komma als Dezimaltrennzeichen):
$cultureInfo = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
# Wenn du einen Punkt als Dezimaltrennzeichen beibehalten m�chtest (z.B. f�r englische Excel-Version):
# $cultureInfo = [System.Globalization.CultureInfo]::InvariantCulture

# Liste der erwarteten Fehlercodes f�r die Statistikspalten
$errorCodesToTrack = @(
    "ERR_00000", "ERR_00001", "ERR_00002", "ERR_00322", "ERR_00323",
    "ERR_00460", "ERR_04751", "ERR_04758", "ERR_04760", "ERR_04761",
    "ERR_04773", "ERR_04818", "ERR_04822", "ERR_04824", "ERR_05010",
    "ERR_05013", "ERR_05029", "ERR_05070", "ERR_05073", "ERR_05079",
    "ERR_05086", "ERR_05127", "ERR_05354", "ERR_05360", "ERR_05366",
    "ERR_05413", "ERR_05433", "ERR_05439", "ERR_05454", "ERR_06165",
    "ERR_06327", "ERR_06433", "ERR_06456", "ERR_06461", "ERR_06474",
    "ERR_06483", "ERR_06484", "ERR_06485", "ERR_06486", "ERR_06487",
    "ERR_06495", "ERR_06502", "ERR_06503", "ERR_06505", "ERR_06601",
    "ERR_07504", "ERR_07602", "ERR_07609", "ERR_07610", "ERR_07616",
    "ERR_07617", "ERR_07619", "ERR_07654", "ERR_07656", "ERR_07951",
    "ERR_08003", "ERR_08004", "ERR_08007", "ERR_08009", "ERR_08105",
    "ERR_08107", "ERR_08109", "ERR_08110", "ERR_08111", "ERR_08203",
    "ERR_08214", "ERR_08215", "ERR_08216", "ERR_08242", "ERR_08243",
    "ERR_08244", "ERR_08245", "ERR_08246", "ERR_08301", "ERR_08302",
    "ERR_08303", "ERR_08304", "ERR_08305", "ERR_08311", "ERR_08312",
    "ERR_08314", "ERR_08318", "ERR_08326", "ERR_08330", "ERR_08337",
    "ERR_10462", "ERR_10830", "ERR_10853", "ERR_10855", "ERR_10865",
    "ERR_10866", "ERR_10901", "ERR_10905", "ERR_10906", "ERR_10908",
    "ERR_10910", "ERR_10913", "ERR_10914", "ERR_10919", "ERR_10921",
    "ERR_10924", "ERR_11000", "ERR_11046"
)

# --- �berpr�fen und Installieren des ImportExcel Moduls ---
# --- Hinweis:
# --- Update mit: Update-Module -Name ImportExcel
# --- Version anzeigen: Get-InstalledModule -Name ImportExcel
# --- Alle Module: einfach -Name ImportExcel weglassen.

Write-Host "�berpr�fe, ob das 'ImportExcel' Modul installiert ist..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Warning "Das 'ImportExcel' Modul ist nicht installiert. Versuche, es zu installieren."
    try {
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "Das 'ImportExcel' Modul wurde erfolgreich installiert." -ForegroundColor Green
    } catch {
        Write-Error "Fehler beim Installieren des 'ImportExcel' Moduls: $($_.Exception.Message)"
        Write-Error "Bitte installieren Sie es manuell mit: Install-Module -Name ImportExcel -Scope CurrentUser"
        return # Skript beenden, wenn das Modul nicht installiert werden kann
    }
} else {
    Write-Host "Das 'ImportExcel' Modul ist bereits installiert." -ForegroundColor Green
}

# Importiere das Modul, um seine Funktionen nutzen zu k�nnen
Import-Module -Name ImportExcel -ErrorAction Stop

# Leere Liste f�r alle gesammelten Statistikobjekte aus den Logdateien
$allStatsObjects = @()

# --- Hauptverarbeitung: Schleife durch alle .LA1.txt-Dateien ---
Get-ChildItem -Path $logFilesPath -Filter "*.LA1.txt" | ForEach-Object {
    $currentFile = $_.FullName
    Write-Host "`nVerarbeite Statistikdatei: $($_.Name)" -ForegroundColor Green

    $fileContent = Get-Content -Path $currentFile -ErrorAction SilentlyContinue

    if (-not $fileContent) {
        Write-Warning "Konnte Inhalt der Datei '$($_.Name)' nicht lesen. �berspringe diese Datei."
        return
    }

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

    foreach ($code in $errorCodesToTrack) {
        $propertyName = "E_" + $code.Replace('ERR_','')
        Add-Member -InputObject $currentStats -NotePropertyName $propertyName -NotePropertyValue ([int]0) 
    }

    # --- Datum extrahieren ---
    $dateLineObj = $fileContent | Select-String -Pattern "^\s*LEVEL 1 : DAILY RESULTS : (.+?)\s*$"
    if ($dateLineObj) {
        if ($dateLineObj.Line -match "^\s*LEVEL 1 : DAILY RESULTS : (.+?)\s*$") {
            $currentStats.Datum = ($Matches[1]).Trim()
        }
    }

    # --- Performance-Daten extrahieren ---
    $performanceLineObj = $fileContent | Select-String -Pattern "^\s*Performance: AVG:\s*(\d+\.?\d*|No Information),\s*PEAK:\s*(\d+\.?\d*|No Information),\s*Potential:\s*(\d+\.?\d*|No Information)(?:\s*-\s*hc\s*(\d+\.?\d*%))?\s*$"
    if ($performanceLineObj) {
        $performanceLine = $performanceLineObj.Line
        
        if ($performanceLine -match "^\s*Performance: AVG:\s*(\d+\.?\d*|No Information),\s*PEAK:\s*(\d+\.?\d*|No Information),\s*Potential:\s*(\d+\.?\d*|No Information)(?:\s*-\s*hc\s*(\d+\.?\d*%))?\s*$") {
            # AVG verarbeiten
            $rawAvgValue = $Matches[1].Trim()
            $parsedAvg = 0.0
            if ($rawAvgValue -ne "No Information" -and [double]::TryParse($rawAvgValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedAvg)) {
                $currentStats.Performance_AVG = $parsedAvg.ToString($cultureInfo)
            } else {
                $currentStats.Performance_AVG = ""
            }

            # PEAK verarbeiten
            $rawPeakValue = $Matches[2].Trim()
            $parsedPeak = 0.0
            if ($rawPeakValue -ne "No Information" -and [double]::TryParse($rawPeakValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedPeak)) {
                $currentStats.Performance_PEAK = $parsedPeak.ToString($cultureInfo)
            } else {
                $currentStats.Performance_PEAK = ""
            }

            # Potential verarbeiten mit Standardwert-Logik
            $rawPotentialValue = $Matches[3].Trim()
            if ($rawPotentialValue -eq "No Information") {
                $currentStats.Performance_Potential = $defaultPotentialValue.ToString($cultureInfo)
            } else {
                $parsedPotential = 0.0
                if ([double]::TryParse($rawPotentialValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedPotential)) {
                    $currentStats.Performance_Potential = $parsedPotential.ToString($cultureInfo)
                } else {
                    $currentStats.Performance_Potential = ""
                }
            }
            
            # HC verarbeiten (nur wenn vorhanden)
            if ($Matches.Count -ge 5 -and -not [string]::IsNullOrWhiteSpace($Matches[4])) {
                $currentStats.Performance_HC = $Matches[4].Trim()
            } else {
                $currentStats.Performance_HC = "" 
            }
        }
    }

    # --- Plate Production Daten extrahieren ---
    $plateProductionLineObj = $fileContent | Select-String -Pattern "^\s*Plate Production:\s*(\d+)\s*,\s*Exposed Plates:\s*(\d+)\s*,\s*Damaged Plates:\s*(\d+)\s*$"
    if ($plateProductionLineObj) {
        if ($plateProductionLineObj.Line -match "^\s*Plate Production:\s*(\d+)\s*,\s*Exposed Plates:\s*(\d+)\s*,\s*Damaged Plates:\s*(\d+)\s*$") {
            $currentStats.Plate_Production = [int]$Matches[1].Trim()
            $currentStats.Exposed_Plates = [int]$Matches[2].Trim()
            $currentStats.Damaged_Plates = [int]$Matches[3].Trim()
        }
    }

    # --- Plate Count Daten extrahieren ---
    $plateCountLineObj = $fileContent | Select-String -Pattern "^\s*Plate Count:\s*(\d+)\s*$"
    if ($plateCountLineObj) {
        if ($plateCountLineObj.Line -match "^\s*Plate Count:\s*(\d+)\s*$") {
            $currentStats.Plate_Count = [int]$Matches[1].Trim()
        }
    }

    # --- Message Statistics Daten extrahieren ---
    $startIndex = -1
    $endIndex = -1
    for ($i = 0; $i -lt $fileContent.Length; $i++) {
        if ($fileContent[$i] -match "^\s*Message Statistics:\s*$") {
            $startIndex = $i
        } elseif ($startIndex -ne -1 -and $fileContent[$i] -match "^\s*--------------------------------------------------------------------------------------------------------------------------\s*$") {
            if ($i -gt $startIndex + 1) {
                $endIndex = $i
                break
            }
        }
    }

    if ($startIndex -ne -1 -and $endIndex -ne -1) {
        $messageStatsBlock = $fileContent[($startIndex + 2)..($endIndex - 1)]
        foreach ($line in $messageStatsBlock) {
            if ($line -match "^\s*(\d+)\s*\|.*?\|(ERR_\d+).*$") {
                $count = [int]$Matches[1].Trim()
                $errorCode = $Matches[2].Trim()
                if ($errorCodesToTrack -contains $errorCode) {
                    $propertyName = "E_" + $errorCode.Replace('ERR_','')
                    $currentStats.$propertyName = ([int]$count) # Sicherstellen, dass der Wert als [int] gesetzt wird
                }
            }
        }
    } else {
        Write-Warning "Message Statistics Block in Datei '$($_.Name)' nicht gefunden oder unvollst�ndig. Fehlerz�hlungen bleiben 0."
    }

    $allStatsObjects += $currentStats
}

# --- Statistik-XLSX-Datei schreiben ---
if ($allStatsObjects.Count -gt 0) {
    Write-Host "`n--- �ffne Speicherdialog f�r die Statistik-XLSX-Datei ---" -ForegroundColor Yellow

    # Erstelle den Speichern-Dialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*"
    $saveFileDialog.Title = "Statistik-Datei speichern unter"
    $saveFileDialog.FileName = "Gesamtauswertung_Statistik_Logfiles.xlsx"
    $saveFileDialog.InitialDirectory = $logFilesPath # Setzt den Startordner auf den ausgew�hlten Log-Pfad

    $saveResult = $saveFileDialog.ShowDialog()

    if ($saveResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputStatsXlsxFile = $saveFileDialog.FileName
        Write-Host "Die Statistikdatei wird unter '$outputStatsXlsxFile' gespeichert." -ForegroundColor Green
        
        # Try-Catch Block f�r den Export
        try {
            # Export der Rohdaten auf das erste Arbeitsblatt
            $allStatsObjects | Export-Excel -Path $outputStatsXlsxFile -AutoSize -ClearSheet -AutoFilter -WorksheetName "Rohdaten" -ErrorAction Stop

            Write-Host "-------------------------------------" -ForegroundColor Green
            Write-Host "Statistik-Verarbeitung abgeschlossen!" -ForegroundColor Green
            Write-Host "Die Daten wurden in '$outputStatsXlsxFile' gespeichert." -ForegroundColor Green
            Write-Host "Gesamtzahl der Statistik-Eintr�ge: $($allStatsObjects.Count)" -ForegroundColor Green

        } catch {
            Write-Error "FEHLER beim Exportieren der Excel-Datei: $($_.Exception.Message)"
            Write-Error "M�glicherweise ist die Datei '$outputStatsXlsxFile' ge�ffnet und blockiert den Zugriff. Bitte schlie�en Sie die Datei und versuchen Sie es erneut."
        }
    } else {
        Write-Warning "Speichern abgebrochen. Die Statistik-Datei wurde nicht erstellt."
    }
} else {
    Write-Warning "`nKeine Daten zum Exportieren gefunden. Es wurde keine Excel-Datei erstellt."
}
