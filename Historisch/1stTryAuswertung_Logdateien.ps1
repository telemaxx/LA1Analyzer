# Speicherort der Logdateien und der Ausgabedatei festlegen
$logFilesPath = "C:\Users\top\Desktop\Analyzetool" # Passe diesen Pfad an!
$outputCsvFile = "C:\Users\top\Desktop\Analyzetool\Gesamtauswertung_Logfiles.csv" # Passe diesen Pfad an!

# Trennzeichen für die CSV-Datei (Komma ist Standard, Semikolon für Excel in DE)
$csvDelimiter = "," 

# Angepasste Kopfzeile für die CSV-Datei (Logfile_Datum entfernt)
$header = "Zeitstempel,Log_Typ,Level,Fehlercode,Nachricht"

# Leere Liste für alle gesammelten Daten
$allLogEntries = @()

# Schleife durch alle .LA1.txt-Dateien im angegebenen Pfad
Get-ChildItem -Path $logFilesPath -Filter "*.LA1.txt" | ForEach-Object {
    $currentFile = $_.FullName
    Write-Host "Verarbeite Datei: $($_.Name)"

    # Datum aus der Datei extrahieren (wird weiterhin benötigt, aber nicht mehr explizit in CSV-Spalte)
    $fileContent = Get-Content -Path $currentFile
    $dateLine = $fileContent | Select-String -Pattern "LEVEL 1 : DAILY RESULTS :"

    $logFileDate = ""
    if ($dateLine) {
        $logFileDate = ($dateLine.ToString().Split(":")[-1]).Trim()
        Write-Host "  -> Extrahiertes Datum: $logFileDate" 
    } else {
        Write-Warning "Konnte Datum in Datei $($_.Name) nicht finden. Überspringe diese Datei."
        return 
    }

    # Filtern der "Err_Line"-Einträge
    $errorLines = $fileContent | Where-Object { $_ -match "^\*?\s*\d{4}-\d{2}-\d{2} / \d{2}:\d{2}:\d{2}\.\d{3}\|Err_Line\|" }
    
    Write-Host "  -> Anzahl der gefundenen Err_Line-Zeilen: $($errorLines.Count)" 

    foreach ($line in $errorLines) {
        Write-Host "    -> Verarbeite Zeile: $line" 

        # Entferne eventuelle führende "* " oder "* | " Zeichen
        $cleanedLine = $line -replace "^\*[\s\|]*", "" 

        # Ersetze Pipe-Zeichen durch das definierte CSV-Trennzeichen
        $parts = $cleanedLine.Split("|")
        
        Write-Host "      -> Teile nach Split: $($parts.Count)" 

        if ($parts.Count -ge 5) { 
            $timestamp = $parts[0].Trim() 
            $logType = $parts[1].Trim() 
            $level = $parts[2].Trim() 
            $errorCode = $parts[3].Trim() 
            $message = $parts[4].Trim() 

            # Angepasste formatierte Zeile (Logfile_Datum am Anfang entfernt)
            $formattedLine = "$timestamp$csvDelimiter$logType$csvDelimiter$level$csvDelimiter$errorCode$csvDelimiter$message"
            $allLogEntries += $formattedLine
            Write-Host "        -> Hinzugefügte Zeile: $formattedLine" 
        } else {
            Write-Warning "      -> Zeile hatte nicht genügend Teile nach dem Split: $cleanedLine" 
        }
    }
}

# CSV-Datei schreiben
$header | Out-File -FilePath $outputCsvFile -Encoding UTF8
$allLogEntries | Out-File -FilePath $outputCsvFile -Encoding UTF8 -Append

Write-Host "Verarbeitung abgeschlossen. Die Daten wurden in '$outputCsvFile' gespeichert. Gesamtzahl der Einträge: $($allLogEntries.Count)"