# ---------------------------------------------------------------------
# PhinIT-easyBITKEYS (Version 0.3.5 - PowerShell 5.1)
# Autor: ANDREAS HEPP
# Letztes Update: 02.03.2025
#
# Dieses Skript liest aus dem AD alle BitLocker Keys (msFVE-RecoveryInformation)
# aus und zeigt diese in einer farbcodierten GUI an.
# Je nach Zustand:
#   - Kein Schlüssel vorhanden: LightGray
#   - Schlüssel älter als 3 Monate: LightGreen
#   - Schlüssel älter als 6 Monate: Orange
#   - Schlüssel älter als 12 Monate: LightCoral
#
# Außerdem gibt es Buttons zum Export als CSV, zum Löschen des markierten Eintrags 
# (nach Bestätigung) und zum Schließen der Anwendung.
#
# Die DataGridView zeigt 4 Spalten an:
#   "Nr" (nur so breit wie der Inhalt),
#   "ComputerName" (füllt den übrigen Platz),
#   "BitLockerKey" (füllt den übrigen Platz),
#   "Datum" (nur so breit wie der Inhalt).
#
# Der Footer (mit URL) wird als Panel unter den Buttons angezeigt und
# erhält dieselbe Hintergrundfarbe wie der Header (laut INI).
# ---------------------------------------------------------------------

# Sicherstellen, dass das Skript in PowerShell 5.1 läuft
if ($PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host "Wechsle zu PowerShell 5.1..." -ForegroundColor Yellow
    Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Erforderliche Assemblies laden
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zum Einlesen einer INI-Datei
function Get-IniContent {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if (-not (Test-Path $Path)) {
        Write-Error "INI-Datei '$Path' wurde nicht gefunden."
        exit
    }
    $ini = @{}
    $section = ""
    foreach ($line in Get-Content $Path) {
        $line = $line.Trim()
        if ($line -match '^\s*\[([^\]]+)\]') {
            $section = $matches[1]
            $ini[$section] = @{}
        }
        elseif ($line -match '^\s*([^;].*?)=(.*)$') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            if ($section) {
                $ini[$section][$key] = $value
            }
        }
    }
    return $ini
}

# INI-Datei einlesen (Pfad zur INI-Datei anpassen)
$iniFile = "easyBITkeys.ini"
$iniContent = Get-IniContent -Path $iniFile
$cfg = $iniContent["easyBITKEYSGeneral"]

# Konfigurationswerte aus der INI
$ScriptVersion   = $cfg["easyBITKEYSVersion"]
$LastUpdate      = $cfg["easyBITKEYSLastUpdate"]
$Author          = $cfg["easyBITKEYSAuthor"]
$APPName         = $cfg["easyBITKEYSAPPName"]
$ThemeColor      = $cfg["easyBITKEYSThemeColor"]
$FontFamily      = $cfg["easyBITKEYSFontFamily"]
$FontSize        = [float]$cfg["easyBITKEYSFontSize"]
$HeaderLogoPath  = $cfg["easyBITKEYSHeaderLogo"]
$HeaderLogoURL   = $cfg["easyBITSSHeaderLogoURL"]
$GUI_HeaderRaw   = $cfg["easyBITSGUI_Header"]
$FooterText      = $cfg["easyBITSFooterText"]
$TableFormat     = $cfg["easyBITSTable"]  # (Nur als Spaltenüberschrift)

# Schriftart definieren
$guiFont = New-Object System.Drawing.Font($FontFamily, $FontSize)

# Platzhalter im GUI-Header ersetzen
$GUI_Header = $GUI_HeaderRaw -replace "\{ScriptVersion\}", $ScriptVersion `
                              -replace "\{LastUpdate\}", $LastUpdate `
                              -replace "\{Author\}", $Author

# GUI erstellen
$form = New-Object System.Windows.Forms.Form
$form.Text = $APPName
$form.Size = New-Object System.Drawing.Size(915,720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# HEADER Panel (Hintergrundfarbe gemäß ThemeColor aus INI)
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Size = New-Object System.Drawing.Size(900, 65)
$headerPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml($ThemeColor)
$headerPanel.Location = New-Object System.Drawing.Point(0,0)
$form.Controls.Add($headerPanel)

# Links: Panel für den Header-Text
$headerTextPanel = New-Object System.Windows.Forms.Panel
$headerTextPanel.Size = New-Object System.Drawing.Size(630, 60)
$headerTextPanel.Location = New-Object System.Drawing.Point(10, 0)
$headerPanel.Controls.Add($headerTextPanel)

# Header-Label
$headerLabel = New-Object System.Windows.Forms.Label
$headerLabel.AutoSize = $true
$headerLabel.Location = New-Object System.Drawing.Point(10,20)
$headerTextPanel.Controls.Add($headerLabel)

# Rechts: Panel für das Logo
$headerLogoPanel = New-Object System.Windows.Forms.Panel
$headerLogoPanel.Size = New-Object System.Drawing.Size(250, 60)
$headerLogoPanel.Location = New-Object System.Drawing.Point(640, 0)
$headerPanel.Controls.Add($headerLogoPanel)

# Logo PictureBox
$logo = New-Object System.Windows.Forms.PictureBox
$logo.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$logo.Size = New-Object System.Drawing.Size(250,60)
$logo.Location = New-Object System.Drawing.Point(0,0)
if (Test-Path $HeaderLogoPath) {
    $logo.Image = [System.Drawing.Image]::FromFile($HeaderLogoPath)
}
else {
    $bmp = New-Object System.Drawing.Bitmap(250,60)
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.Clear([System.Drawing.Color]::LightGray)
    $graphics.DrawString("Logo nicht gefunden", $guiFont, [System.Drawing.Brushes]::Black, 10, 20)
    $logo.Image = $bmp
}
$headerLogoPanel.Controls.Add($logo)

# Klick-Event für Logo (öffnet URL)
$logo.Add_Click({
    Start-Process $HeaderLogoURL
})

# DataGridView zur Darstellung der Tabelle
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Size = New-Object System.Drawing.Size(880, 480)
$dataGrid.Location = New-Object System.Drawing.Point(10, 100)
# Es werden 4 Spalten definiert:
#   Spalte 0: "Nr"
#   Spalte 1: "ComputerName"
#   Spalte 2: "BitLockerKey"
#   Spalte 3: "Datum"
$dataGrid.ColumnCount = 4
$dataGrid.Columns[0].Name = "Nr"
$dataGrid.Columns[1].Name = "ComputerName"
$dataGrid.Columns[2].Name = "BitLockerKey"
$dataGrid.Columns[3].Name = "Datum"

$dataGrid.ReadOnly = $true
$dataGrid.AllowUserToAddRows = $false
$dataGrid.SelectionMode = 'FullRowSelect'

# Spalten-AutoSize Einstellungen:
# Spalte "Nr": nur so breit wie der Inhalt (AllCells)
$dataGrid.Columns[0].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
# Spalte "ComputerName": füllt den übrigen Platz
$dataGrid.Columns[1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
# Spalte "BitLockerKey": füllt den übrigen Platz
$dataGrid.Columns[2].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
# Spalte "Datum": nur so breit wie der Inhalt (AllCells)
$dataGrid.Columns[3].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells

$form.Controls.Add($dataGrid)

# Button Panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Size = New-Object System.Drawing.Size(880, 40)
$buttonPanel.Location = New-Object System.Drawing.Point(10, 600)
$form.Controls.Add($buttonPanel)

# Export Button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "EXPORT als CSV"
$exportButton.Size = New-Object System.Drawing.Size(150,30)
$exportButton.Location = New-Object System.Drawing.Point(200,5)
# Farbe: Hellgrün
$exportButton.BackColor = [System.Drawing.Color]::LightGreen
$buttonPanel.Controls.Add($exportButton)

# LÖSCHEN Button (löscht den markierten Eintrag nach Bestätigung)
$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "ENTFERNEN"
$deleteButton.Size = New-Object System.Drawing.Size(150,30)
$deleteButton.Location = New-Object System.Drawing.Point(380,5)
# Farbe: Hellorange (LightSalmon)
$deleteButton.BackColor = [System.Drawing.Color]::LightSalmon
$buttonPanel.Controls.Add($deleteButton)

$deleteButton.Add_Click({
    $selectedRows = $dataGrid.SelectedRows
    if ($selectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Bitte einen Eintrag aus der Liste markieren!", "Hinweis", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $selectedRow = $selectedRows[0]
    $confirm = [System.Windows.Forms.MessageBox]::Show("Den markierten Eintrag wirklich entfernen?", "Entfernen", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $dn = $selectedRow.Tag
        try {
            Remove-ADObject -Identity $dn -Confirm:$false
            $dataGrid.Rows.Remove($selectedRow)
            [System.Windows.Forms.MessageBox]::Show("Eintrag wurde entfernt.", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim entfernen des Eintrags: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# CLOSE Button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Text = "CLOSE"
$closeButton.Size = New-Object System.Drawing.Size(150,30)
$closeButton.Location = New-Object System.Drawing.Point(560,5)
# Farbe: Hellrot
$closeButton.BackColor = [System.Drawing.Color]::LightCoral
$buttonPanel.Controls.Add($closeButton)

$closeButton.Add_Click({
    $form.Close()
})

# Footer Panel (unter den Buttons, mit derselben Theme-Farbe wie der Header)
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Size = New-Object System.Drawing.Size(900, 35)
$footerPanel.Location = New-Object System.Drawing.Point(0, 660)
$footerPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml($ThemeColor)
$form.Controls.Add($footerPanel)

# Footer Label im Footer Panel
$footerLabel = New-Object System.Windows.Forms.Label
# Kombiniert FooterText und Header-Informationen:
$footerLabel.Text = "$FooterText  =  $GUI_Header"
$footerLabel.Font = $guiFont
$footerLabel.AutoSize = $true
$footerLabel.Location = New-Object System.Drawing.Point(10, 3)
$footerLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
$footerLabel.Add_Click({
    Start-Process $HeaderLogoURL
})
$footerPanel.Controls.Add($footerLabel)

# CSV-Export Funktion
$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Als CSV speichern"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvContent = @()
        foreach ($row in $dataGrid.Rows) {
            if ($row.IsNewRow) { continue }
            $csvContent += "{0},{1},{2},{3}" -f $row.Cells[0].Value, $row.Cells[1].Value, $row.Cells[2].Value, $row.Cells[3].Value
        }
        $csvContent | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Export abgeschlossen!", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# ActiveDirectory Modul importieren
Import-Module ActiveDirectory -ErrorAction SilentlyContinue
if (-not (Get-Module -Name ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show("ActiveDirectory Modul ist nicht verfügbar. Bitte RSAT-Tools installieren.", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# AD-Abfrage: Alle BitLocker Keys (msFVE-RecoveryInformation)
try {
    $BitLockerKeys = Get-ADObject -Filter 'objectClass -eq "msFVE-RecoveryInformation"' -Properties msFVE-RecoveryPassword, whenCreated, distinguishedName
} catch {
    [System.Windows.Forms.MessageBox]::Show("Fehler bei der AD-Abfrage: $_", "Fehler", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Daten in DataGridView einfügen mit farblicher Kennzeichnung
$counter = 1
foreach ($key in $BitLockerKeys) {
    # Extrahiere den tatsächlichen ComputerName aus dem DistinguishedName:
    # Annahme: DN: CN={RecoveryGUID},CN={ComputerName},...
    $dnParts = $key.distinguishedName.Split(',')
    if ($dnParts.Length -ge 2) {
        $computerName = $dnParts[1] -replace "^CN=", ""
    }
    else {
        $computerName = "Unbekannt"
    }
    # Korrekte Referenzierung der AD-Eigenschaft mit Bindestrich
    $bitLockerKey = $key.'msFVE-RecoveryPassword'
    $date = $key.whenCreated
    $rowIndex = $dataGrid.Rows.Add()
    $dataGrid.Rows[$rowIndex].Cells[0].Value = $counter
    $dataGrid.Rows[$rowIndex].Cells[1].Value = $computerName
    $dataGrid.Rows[$rowIndex].Cells[2].Value = $bitLockerKey
    $dataGrid.Rows[$rowIndex].Cells[3].Value = $date.ToString("dd.MM.yyyy")
    # Speichere den DistinguishedName in der Zeile (Tag-Eigenschaft) für spätere Löschaktionen
    $dataGrid.Rows[$rowIndex].Tag = $key.distinguishedName
    
    # Farbliche Kennzeichnung der gesamten Zeile
    if ([string]::IsNullOrEmpty($bitLockerKey)) {
        $dataGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
    }
    else {
        $ageDays = (Get-Date) - $date
        $ageMonths = $ageDays.Days / 30
        if ($ageMonths -ge 12) {
            $dataGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
        }
        elseif ($ageMonths -ge 6) {
            $dataGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::Orange
        }
        elseif ($ageMonths -ge 3) {
            $dataGrid.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
        }
    }
    $counter++
}

# Formular anzeigen
[void]$form.ShowDialog()
