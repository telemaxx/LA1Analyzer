# Gesamtes PowerShell-Skript zur Analyse von .LA1.txt Logdateien und Erstellung einer Statistik-XLSX

# --- Konfiguration ---
# Standardwert für "Performance_Potential", falls in der Logdatei "No Information" steht.
# Dieser Wert wird nun über eine GUI abgefragt.
$defaultPotentialValue = 999 

# --- GUI zur Abfrage des Standardwerts für 'Potential' ---
Write-Host "Öffne Fenster zur Eingabe des Standardwerts für 'Potential'." -ForegroundColor Cyan

try {
    # Benötigt für GUI-Elemente
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Standardwert für 'Potential' festlegen"
    $form.Size = New-Object System.Drawing.Size(400,150)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(350,20)
    $label.Text = "Standardwert für 'Performance_Potential' bei 'No Information':"

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,50)
    $textBox.Size = New-Object System.Drawing.Size(360,20)
    $textBox.Text = $defaultPotentialValue # Setze den aktuellen Standardwert als Text

    $buttonOk = New-Object System.Windows.Forms.Button
    $buttonOk.Location = New-Object System.Drawing.Point(280,80)
    $buttonOk.Size = New-Object System.Drawing.Size(90,25)
    $buttonOk.Text = "OK"
    $buttonOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.AcceptButton = $buttonOk
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($buttonOk)

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputValue = $textBox.Text.Trim()
        # **KORREKTUR:** [double] statt [int] für robustere Verarbeitung
        if ([double]::TryParse($inputValue, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedValue)) {
            $defaultPotentialValue = $parsedValue
            Write-Host "Verwende eingegebenen Standardwert: '$defaultPotentialValue'" -ForegroundColor Green
        } else {
            Write-Warning "Ungültige Eingabe. Verwende weiterhin den Standardwert von 999."
        }
    } else {
        Write-Warning "Eingabe abgebrochen. Verwende weiterhin den Standardwert von 999."
    }
} catch {
    Write-Warning "Konnte GUI nicht anzeigen. Verwende den Standardwert von 999."
}
# Informiere über den finalen Wert, der verwendet wird
Write-Host "Standardwert für 'Potential' ist auf '$defaultPotentialValue' gesetzt." -ForegroundColor Cyan

# --- Pfad für Logdateien festlegen (immer über interaktive GUI) ---
$logFilesPath = $null

do {
    # Benötigt für GUI-Elemente
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    # Setze den Dialog-Titel für mehr Klarheit
    $browser.Description = "Bitte wählen Sie den Ordner mit den .LA1.txt Logdateien aus."
    
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
        $logFilesPath = "" # Setzt den Pfad zurück, damit die Schleife wiederholt wird
    }

} while ([string]::IsNullOrEmpty($logFilesPath))

Write-Host "Verwende interaktiv ausgewählten Pfad: '$logFilesPath'" -ForegroundColor Green


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

# --- Überprüfen und Installieren des ImportExcel Moduls ---
Write-Host "Überprüfe, ob das 'ImportExcel' Modul installiert ist..." -ForegroundColor Cyan
if (-not (Get-Module -ListAvailable -Name ImportExcel))