# easyBITLOCKERkeys

**Version:** 0.3.5 (PowerShell 5.1)  
**Author:** Andreas Hepp  
**Last Update:** 02/03/2025  

This script (`easyBITLOCKERkeys.ps1`) retrieves all BitLocker keys (`msFVE-RecoveryInformation`) from Active Directory and displays them in a color-coded GUI:

- **No key found:** LightGray  
- **Key older than 3 months:** LightGreen  
- **Key older than 6 months:** Orange  
- **Key older than 12 months:** LightCoral  

There are buttons to:
- **Export** the displayed entries to a CSV file  
- **Delete** the selected entry (after confirmation)  

## Requirements
- **PowerShell 5.1**  
  The script checks if PowerShell 5.1 is running. If not, it attempts to restart itself under PowerShell 5.1.  
- **Active Directory Module**  
  Required for `Get-ADObject` and `Remove-ADObject`. Make sure RSAT tools are installed and the Active Directory module is available.

## Usage
1. Place `easyBITLOCKERkeys.ps1` and its INI file (`easyBITkeys.ini`) in the same directory.
2. Run `easyBITLOCKERkeys.ps1` from a **PowerShell 5.1** session with sufficient privileges to query and modify AD objects.
3. The GUI will launch and display all found `msFVE-RecoveryInformation` objects, color-coded by their age.

## License
This project is released under the [MIT License](https://opensource.org/licenses/MIT). 
You are free to use, modify, and redistribute this code.

## Contributions
Contributions, suggestions, and feedback are always welcome!  
Please open an issue or submit a pull request if you have any ideas or improvements.

-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
# easyBITLOCKERkeys

**Version:** 0.3.5 (PowerShell 5.1)  
**Autor:** Andreas Hepp  
**Letztes Update:** 02.03.2025  

Dieses Skript (`easyBITLOCKERkeys.ps1`) liest alle BitLocker-Schlüssel (`msFVE-RecoveryInformation`) aus dem Active Directory aus und zeigt sie in einer farbcodierten GUI an:

- **Kein Schlüssel vorhanden:** LightGray  
- **Schlüssel älter als 3 Monate:** LightGreen  
- **Schlüssel älter als 6 Monate:** Orange  
- **Schlüssel älter als 12 Monate:** LightCoral  

Folgende Schaltflächen stehen zur Verfügung:
- **Export** der Einträge als CSV-Datei  
- **Löschen** des ausgewählten Eintrags (nach Bestätigung)  

## Voraussetzungen
- **PowerShell 5.1**  
  Das Skript überprüft, ob PowerShell 5.1 ausgeführt wird. Ist dies nicht der Fall, versucht es sich selbst in PowerShell 5.1 neu zu starten.  
- **Active Directory Modul**  
  Wird für `Get-ADObject` und `Remove-ADObject` benötigt. Stelle sicher, dass die RSAT-Tools installiert sind und das Active-Directory-Modul verfügbar ist.

## Verwendung
1. Lege `easyBITLOCKERkeys.ps1` und die INI-Datei (`easyBITkeys.ini`) im selben Verzeichnis ab.
2. Führe `easyBITLOCKERkeys.ps1` in einer **PowerShell-5.1**-Session mit ausreichenden Rechten zum Abfragen und Bearbeiten von AD-Objekten aus.
3. Die GUI startet und zeigt alle gefundenen `msFVE-RecoveryInformation`-Objekte an, farblich nach Alter der Schlüssel gekennzeichnet.

## Lizenz
Dieses Projekt wird unter der [MIT Lizenz](https://opensource.org/licenses/MIT) veröffentlicht. 
Du bist frei, den Code zu verwenden, anzupassen und weiterzuverbreiten.

## Beiträge
Beiträge, Vorschläge und Feedback sind immer willkommen!  
Erstelle gerne ein Issue oder einen Pull Request, wenn du Ideen oder Verbesserungen hast.
