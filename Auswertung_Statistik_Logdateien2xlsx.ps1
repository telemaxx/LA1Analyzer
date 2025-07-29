# Gesamtes PowerShell-Skript zur Analyse von .LA1.txt Logdateien und Erstellung einer Statistik-XLSX

# --- Parameter-Definition ---
param(
    [Parameter(Mandatory=$false)]
    [string]$LogPath # Optionale Pfadangabe beim Skriptaufruf
)

# --- Konfiguration ---
# Standardwert für "Performance_Potential", falls in der Logdatei "No Information" steht.
$defaultPotentialValue = 999 

Write-Host "Standardwert für 'Potential' ist auf '$defaultPotentialValue' gesetzt." -ForegroundColor Cyan

# --- Pfad für Logdateien festlegen (Parameter > Standard > Interaktiv) ---
$defaultLogPath = "C:\Users\top\git\LA1Analyzer\Samples\AM" # Dein ursprünglicher Standardpfad

$logFilesPath = $null

# Prüfen, ob der Pfad über den Parameter übergeben wurde
if (-not [string]::IsNullOrEmpty($LogPath)) {
    $logFilesPath = $LogPath
    Write-Host "Verwende Logdateipfad aus Parameter: '$logFilesPath'" -ForegroundColor Green
}
# Wenn nicht über Parameter, versuche den Standardpfad
elseif (Test-Path -Path $defaultLogPath -PathType Container) {
    $logFilesPath = $defaultLogPath
    Write-Host "Verwende Standard-Logdateipfad: '$logFilesPath'" -ForegroundColor Green
}
# Wenn weder Parameter noch gültiger Standardpfad, interaktiv abfragen
else {
    Write-Warning "Weder ein gültiger Pfad wurde per Parameter übergeben, noch existiert der Standardpfad."
    do {
        $inputPath = Read-Host "Bitte geben Sie den Pfad zu den Logdateien ein (z.B. C:\Logs)"
        $logFilesPath = $inputPath.Trim()

        if (-not (Test-Path -Path $logFilesPath -PathType Container)) {
            Write-Warning "Der angegebene Pfad '$logFilesPath' existiert nicht oder ist kein Verzeichnis. Bitte versuchen Sie es erneut."
            $logFilesPath = "" # Setzt den Pfad zurück, damit die Schleife wiederholt wird
        }

    } while ([string]::IsNullOrEmpty($logFilesPath))
    Write-Host "Verwende interaktiv eingegebenen Pfad: '$logFilesPath'" -ForegroundColor Green
}

# Finaler Check, ob ein Pfad gefunden wurde, bevor weitergemacht wird
if ([string]::IsNullOrEmpty($logFilesPath) -or (-not (Test-Path -Path $logFilesPath -PathType Container))) {
    Write-Error "Es konnte kein gültiger Pfad zu den Logdateien ermittelt werden. Skript wird beendet."
    return
}

# Der Output-Pfad wird nun basierend auf dem gewählten Log-Pfad generiert
$outputStatsXlsxFile = Join-Path -Path $logFilesPath -ChildPath "Gesamtauswertung_Statistik_Logfiles.xlsx"

Write-Host "Daten des Ordners '$logFilesPath' werden analysiert." -ForegroundColor Green
Write-Host "Die Statistikdatei wird unter '$outputStatsXlsxFile' gespeichert." -ForegroundColor Green


# Wichtig: Kultur-Info für die korrekte Dezimaltrennzeichen-Formatierung
# Für Deutschland (Komma als Dezimaltrennzeichen):
$cultureInfo = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
# Wenn du einen Punkt als Dezimaltrennzeichen beibehalten möchtest (z.B. für englische Excel-Version):
# $cultureInfo = [System.Globalization.CultureInfo]::InvariantCulture

# Liste der erwarteten Fehlercodes für die Statistikspalten
$errorCodesToTrack = @(
    "ERR_00000", "ERR_00001", "ERR_00002", "ERR_00322", "ERR_00323",
    "ERR_00460", "ERR_04751", "ERR_04758", "ERR_04760", "ERR_04761",
    "ERR_04773", "ERR_04818", "ERR_04824", "ERR_05010", "ERR_05013",
    "ERR_05029", "ERR_05073", "ERR_05079", "ERR_05086", "ERR_05127",
    "ERR_05354", "ERR_05360", "ERR_05366", "ERR_05413", "ERR_05433",
    "ERR_05439", "ERR_05454", "ERR_06165", "ERR_06327", "ERR_06433",
    "ERR_06456", "ERR_06461", "ERR_06474", "ERR_06483", "ERR_06484",
    "ERR_06485", "ERR_06486", "ERR_06487", "ERR_06495", "ERR_06502",
    "ERR_06503", "ERR_06505", "ERR_06601", "ERR_07504", "ERR_07602",
    "ERR_07609", "ERR_07610", "ERR_07616", "ERR_07617", "ERR_07619",
    "ERR_07654", "ERR_07656", "ERR_07951", "ERR_08003", "ERR_08004",
    "ERR_08007", "ERR_08009", "ERR_08105", "ERR_08107", "ERR_08109",
    "ERR_08110", "ERR_08111", "ERR_08203", "ERR_08214", "ERR_08215",
    "ERR_08216", "ERR_08242", "ERR_08243", "ERR_08244", "ERR_08245",
    "ERR_08246", "ERR_08301", "ERR_08302", "ERR_08303", "ERR_08304",
    "ERR_08305", "ERR_08311", "ERR_08312", "ERR_08314", "ERR_08318",
    "ERR_08326", "ERR_08330", "ERR_08337", "ERR_10462", "ERR_10830",
    "ERR_10853", "ERR_10855", "ERR_10865", "ERR_10866", "ERR_10901",
    "ERR_10905", "ERR_10906", "ERR_10908", "ERR_10910", "ERR_10913",
    "ERR_10914", "ERR_10919", "ERR_10921", "ERR_10924", "ERR_11000",
    "ERR_11046"
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

    $fileContent = Get-Content -Path $currentFile -ErrorAction SilentlyContinue

    if (-not $fileContent) {
        Write-Warning "Konnte Inhalt der Datei '$($_.Name)' nicht lesen. Überspringe diese Datei."
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
        Write-Warning "Message Statistics Block in Datei '$($_.Name)' nicht gefunden oder unvollständig. Fehlerzählungen bleiben 0."
    }

    $allStatsObjects += $currentStats
}

# --- Statistik-XLSX-Datei schreiben ---
if ($allStatsObjects.Count -gt 0) {
    Write-Host "`n--- Schreibe Statistik-XLSX-Datei ---" -ForegroundColor Yellow
    
    # Try-Catch Block für den Export
    try {
        $allStatsObjects | Export-Excel -Path $outputStatsXlsxFile -AutoSize -ClearSheet -AutoFilter -ErrorAction Stop
        Write-Host "-------------------------------------" -ForegroundColor Green
        Write-Host "Statistik-Verarbeitung abgeschlossen!" -ForegroundColor Green
        Write-Host "Die Daten wurden in '$outputStatsXlsxFile' gespeichert." -ForegroundColor Green
        Write-Host "Gesamtzahl der Statistik-Einträge: $($allStatsObjects.Count)" -ForegroundColor Green

    } catch {
        Write-Error "FEHLER beim Exportieren der Excel-Datei: $($_.Exception.Message)"
        Write-Error "Möglicherweise ist die Datei '$outputStatsXlsxFile' geöffnet und blockiert den Zugriff. Bitte schließen Sie die Datei und versuchen Sie es erneut."
    }
} else {
    Write-Warning "`nKeine Daten zum Exportieren gefunden. Es wurde keine Excel-Datei erstellt."
}
