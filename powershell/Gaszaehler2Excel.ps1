<#
.SYNOPSIS
    Liest Messwerte vom Gaszähler aus und speichert sie in einer Excel-Datei.

.DESCRIPTION
    Dieses PowerShell-Script ruft über HTTP die aktuellen Verbrauchsdaten von einem Gaszähler ab
    (z. B. Zählerstand und Timestamp) und schreibt sie in eine Excel-Datei.
    Fehlerhafte Abrufe (z. B. wegen Timeout oder Verbindungsfehler) werden ebenfalls protokolliert.

.PARAMETER ip
    (Optional) Die IP-Adresse des Gaszählers.
    Standard: 10.0.10.11

.PARAMETER excelPath
    (Optional) Der Pfad zur Excel-Datei (.xlsx), in die die Daten geschrieben werden.
    Standard: "gaszaehler.xlsx" im gleichen Verzeichnis wie das Script.

.EXAMPLE
    .\Gaszaehler2Excel.ps1

    Verwendet die Standard-IP 10.0.10.11 und legt die Datei im Scriptverzeichnis an.

.EXAMPLE
    .\Gaszaehler2Excel.ps1 -ip "10.0.10.50"

    Ruft Werte vom Gaszähler unter 10.0.10.50 ab, speichert in gaszaehler.xlsx im Scriptverzeichnis.

.EXAMPLE
    .\Gaszaehler2Excel.ps1 -ip "10.0.10.50" -excelPath "D:\Logs\gas.xlsx"

    Vollständig konfigurierter Aufruf mit abweichender IP und Excel-Datei.

.NOTES
    Autor: Heiko Huvermann (angepasst)
    Version: 1.0
    Erstellt: 2025-03-25
#>

param (
    [string]$ip = "10.0.10.11",
    [string]$excelPath = "$PSScriptRoot\gaszaehler.xlsx",
    [switch]$help
)

if ($help) {
    Write-Host @"
╔════════════════════════════════════════════════════════════════════╗
║                     Gaszaehler2Excel.ps1 - Hilfe                   ║
╚════════════════════════════════════════════════════════════════════╝

Dieses Script ruft Messwerte (Zeitstempel, Zählerstand) vom Gaszähler per HTTP ab
und speichert sie in einer Excel-Datei (.xlsx). Fehler werden ebenfalls protokolliert.

Parameter:
  -ip <string>           → IP-Adresse des Gaszählers (Standard: 10.0.10.11)
  -excelPath <string>    → Zielpfad für Excel-Datei (Standard: Scriptverzeichnis)
  -help / --help         → Zeigt diese Hilfe an

Beispiele:
  .\Gaszaehler2Excel.ps1
  .\Gaszaehler2Excel.ps1 -ip "10.0.10.50"
  .\Gaszaehler2Excel.ps1 -excelPath "D:\Logs\gas.xlsx"
  .\Gaszaehler2Excel.ps1 -ip "10.0.10.50" -excelPath "D:\Logs\gas.xlsx"
  .\Gaszaehler2Excel.ps1 -help

Erstellt von: Heiko Huvermann (angepasst)
Version: 1.0 | Erstellt: 2025-03-25
"@ -ForegroundColor Cyan
    exit
}

Write-Host "Excel-Datei wird geschrieben nach: $excelPath"
"$($executionTime): Excel-Pfad: $excelPath" | Out-File "$PSScriptRoot\log.txt" -Append


# URL zum Gaszähler (Tasmota-Gerät)
$uri = "http://$ip/cm?cmnd=Status%2010"
$executionTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$now = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
$requestSuccess = $false

# Excel vorbereiten
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $null
$worksheet = $null

# Versuch, die Excel-Datei zu öffnen oder neu zu erstellen
try {
    if (Test-Path $excelPath) {
        $workbook = $excel.Workbooks.Open($excelPath)
        $worksheet = $workbook.Worksheets.Item(1)
    } else {
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        # Spaltenüberschriften setzen
        $worksheet.Cells.Item(1,1).Value() = "Ausführung"
        $worksheet.Cells.Item(1,2).Value() = "Zeitstempel"
        $worksheet.Cells.Item(1,3).Value() = "Zählerstand"
        $worksheet.Cells.Item(1,4).Value() = "Fehler"
    }
}
catch {
    Write-Host "`n❌ Konnte die Excel-Datei nicht öffnen. Ist sie vielleicht gerade geöffnet?" -ForegroundColor Red
    Write-Host "Bitte schließen Sie die Datei und versuchen Sie es erneut." -ForegroundColor Yellow
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    exit 1
}

# Nächste freie Zeile suchen
$row = 2
while ($worksheet.Cells.Item($row,1).Value() -ne $null) {
    $row++
}

# Daten abrufen
try {
    $response = Invoke-RestMethod -Uri $uri -Method GET -TimeoutSec 5

    $time    = $response.StatusSNS.Time
    $counter = $response.StatusSNS.COUNTER.C1

    $worksheet.Cells.Item($row,1).Value() = $executionTime
    $worksheet.Cells.Item($row,2).Value() = $time
    $worksheet.Cells.Item($row,3).Value() = $counter
    $worksheet.Cells.Item($row,4).Value() = ""

    $requestSuccess = $true
}
catch {
    $worksheet.Cells.Item($row,1).Value() = $executionTime
    $worksheet.Cells.Item($row,2).Value() = $now
    $worksheet.Cells.Item($row,3).Value() = ""
    $worksheet.Cells.Item($row,4).Value() = "Verbindungsfehler oder Timeout"
}

# Sicherstellen, dass Zielordner existiert
$folderPath = Split-Path $excelPath
if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
}

# Speichern & schließen
try {
    $workbook.SaveAs($excelPath)
}
catch {
    Write-Host "`n❌ Die Excel-Datei konnte nicht gespeichert werden." -ForegroundColor Red
    Write-Host "Bitte stellen Sie sicher, dass sie nicht geöffnet ist." -ForegroundColor Yellow
}

$workbook.Close()
$excel.Quit()

# COM-Objekte bereinigen
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)  | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)     | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# Ausgabe
if ($requestSuccess) {
    Write-Host "✅ Daten wurden erfolgreich in $excelPath geschrieben." -ForegroundColor Green
} else {
    Write-Error "❌ Messdaten konnten nicht vom Gaszähler abgerufen werden. Fehler wurde in Excel protokolliert."
}
