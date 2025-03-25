param (
    [string]$ip = "10.0.10.29",
    [string]$csvPath = "$PSScriptRoot\stromverbrauch.csv"
)

# URL zusammenbauen
$uri = "http://$ip/cm?cmnd=Status%2010"

# Zeitstempel unabhängig vom Erfolg
$now = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"

try {
    # HTTP-Request mit Timeout (5 Sekunden)
    $response = Invoke-RestMethod -Uri $uri -Method GET -TimeoutSec 5

    # Daten extrahieren
    $time       = $response.StatusSNS.Time
    $totalIn    = $response.StatusSNS.SML.Total_in
    $powerCurr  = $response.StatusSNS.SML.Power_curr

    $data = [PSCustomObject]@{
        Zeitstempel = $time
        Total_in_kWh = $totalIn
        Power_Watt   = $powerCurr
        Fehler       = ""
    }
}
catch {
    # Fehlerbehandlung: Zeile mit aktuellem Zeitstempel und Fehlertext
    $data = [PSCustomObject]@{
        Zeitstempel = $now
        Total_in_kWh = ""
        Power_Watt   = ""
        Fehler       = "Verbindungsfehler oder Timeout"
    }
}

# CSV schreiben oder anhängen
if (-not (Test-Path $csvPath)) {
    $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
} else {
    $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Append
}

# Ausgabe zur Kontrolle
Write-Output $data
