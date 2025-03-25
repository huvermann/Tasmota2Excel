
# Tasmota2Excel

Dieses Repository enthält zwei PowerShell-Skripte, die zum Abrufen von Messdaten von Tasmota-Geräten entwickelt wurden. Dabei wird zwischen zwei Ausgabeformaten unterschieden:

- **Tasmota2Excel.ps1**: Ruft Verbrauchsdaten ab (wie Gesamtverbrauch in kWh und aktuelle Leistung in Watt) und speichert diese in einer Excel-Datei (.xlsx).  
- **TasmotaCsv.ps1**: Ruft dieselben Daten ab und speichert sie in einer CSV-Datei.

## Übersicht

Beide Skripte senden einen HTTP-Request an ein Tasmota-Gerät, um die aktuellen Messwerte abzurufen. Treten Verbindungsfehler oder Timeouts auf, wird dies in der Ausgabedatei protokolliert.

## Parameterübersicht

### Tasmota2Excel.ps1

- **-ip**  
  Optional. Gibt die IP-Adresse des Tasmota-Geräts an.  
  **Standardwert:** `10.0.10.29`

- **-excelPath**  
  Optional. Bestimmt den Pfad zur Excel-Datei, in der die Daten gespeichert werden.  
  **Standardwert:** `stromverbrauch.xlsx` (im gleichen Verzeichnis wie das Skript)

- **-help**  
  Schalter, der eine ausführliche Hilfemeldung mit Beispielen und Parametererklärungen anzeigt.

### TasmotaCsv.ps1

- **-ip**  
  Optional. Gibt die IP-Adresse des Tasmota-Geräts an.  
  **Standardwert:** `10.0.10.29`

- **-csvPath**  
  Optional. Bestimmt den Pfad zur CSV-Datei, in der die Daten gespeichert werden.  
  **Standardwert:** `stromverbrauch.csv` (im gleichen Verzeichnis wie das Skript)

## Beispiele

### Tasmota2Excel.ps1

```powershell
# Ausführen mit Standardparametern:
.\Tasmota2Excel.ps1

# Ausführen mit einer benutzerdefinierten IP-Adresse:
.\Tasmota2Excel.ps1 -ip "10.0.10.50"

# Ausführen mit benutzerdefiniertem Speicherpfad:
.\Tasmota2Excel.ps1 -ip "10.0.10.50" -excelPath "D:\Logs\stromverbrauch.xlsx"
```

### TasmotaCsv.ps1

```powershell
# Ausführen mit Standardparametern:
.\TasmotaCsv.ps1

# Ausführen mit benutzerdefinierter IP-Adresse und CSV-Dateipfad:
.\TasmotaCsv.ps1 -ip "10.0.10.50" -csvPath "D:\Logs\stromverbrauch.csv"
```

## Einstellung der ExecutionPolicy

Beim Ausführen der Skripte kann es zu Fehlermeldungen bezüglich der Ausführungsrichtlinie (ExecutionPolicy) kommen. Falls dies passiert, folge diesen Schritten:

1. **PowerShell als Administrator öffnen**  
   Klicke mit der rechten Maustaste auf das PowerShell-Symbol und wähle **„Als Administrator ausführen“**.

2. **ExecutionPolicy ändern**  
   Setze die Richtlinie beispielsweise auf `RemoteSigned`, um lokal erstellte Skripte ausführen zu können:

   ```powershell
   Set-ExecutionPolicy RemoteSigned
   ```

   Alternativ kannst Du auch `Bypass` wählen, um die Richtlinie temporär zu umgehen:

   ```powershell
   Set-ExecutionPolicy Bypass
   ```

3. **Bestätigung der Änderung**  
   Bestätige die Eingabeaufforderung, falls Du dazu aufgefordert wirst.

Weitere Informationen zur ExecutionPolicy findest Du in der [Microsoft-Dokumentation](https://learn.microsoft.com/powershell/module/microsoft.powershell.security/set-executionpolicy).

**Zurücksetzen der ExecutionPolicy**
Zum zurücksetzen der Richtlinie:
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Undefined
```