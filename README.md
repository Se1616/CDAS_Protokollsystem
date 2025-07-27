# CDAS Protokoll-Erfassungssystem

Ein webbasiertes Protokoll-Erfassungssystem für den "Club der alten Säcke" (CDAS), entwickelt mit Google Apps Script.

## 📋 Übersicht

Das CDAS Protokoll-System automatisiert die komplette Protokollerstellung von der Erfassung bis zum E-Mail-Versand. Es bietet eine benutzerfreundliche Web-Oberfläche mit Auto-Resume-Funktionalität und intelligenter Zähler-Verwaltung.

## ✨ Features

### 🔄 **Auto-Resume System**
- Automatisches Laden des letzten Draft-Protokolls beim Seitenaufruf
- Nahtlose Weiterarbeit an unvollendeten Protokollen
- Erhaltung aller Daten inkl. hochgeladener Bilder

### 📊 **Intelligente Sitzungszähler**
- Automatische Berechnung der nächsten Sitzungsnummer basierend auf Historie
- Separate Zähler für ordentliche, außerordentliche und Festsitzungen
- Manuelle Anpassung möglich

### 👥 **Mitgliederverwaltung**
- Alphabetisch sortierte Mitgliederliste nach Familiennamen
- **Multi-E-Mail Support**: Mehrere E-Mail-Adressen pro Mitglied (`;` getrennt)
- Automatische Vorauswahl des Protokollführers

### 📝 **Tagesordnungspunkte**
- Dynamisches Hinzufügen/Entfernen von Tagesordnungspunkten
- KI-Textverbesserung Option
- Multiple Bild-Uploads pro Tagesordnungspunkt

### 📧 **E-Mail System mit Quota-Management**
- Automatischer Versand an alle Mitglieder-E-Mail-Adressen
- Gmail-Quota-Erkennung mit Status "PDF erstellt - E-Mail wartend"
- Nachversendung wartender E-Mails mit `sendPendingEmails()`

### 📱 **Mobile-First Design**
- Optimiert für iPhone/iPad (Breakpoint bei 1000px)
- Große Touch-Targets und Anti-Zoom für mobile Geräte
- Responsive Grid-Layout

### 📄 **Automatische Dokumentenerstellung**
- Google Docs mit formatiertem Layout
- PDF-Export mit eingebetteten Bildern
- Strukturierte Ablage in Google Drive Ordnern

## 🏗️ Architektur

### Datei-Struktur
```
├── code.gs              # Hauptdatei mit Web App API
├── index.html           # Frontend Interface
├── CoreServices.gs      # Core Services (Drive, Sheets, Utils)
├── BusinessServices.gs  # Business Logic (Counter, Member, Protocol, AutoResume)
├── DocumentServices.gs  # Document Service (Google Docs/PDF)
└── SupportServices.gs   # Support Services (Image, Email, Setup, Tests)
```

### Google Sheets Struktur
```
├── Protokolle           # Hauptprotokoll-Daten
├── Mitglieder          # Mitgliederliste mit Multi-E-Mail Support
├── Tagesordnungspunkte # Agenda-Items mit KI-Flags
└── Bilder              # Bild-Referenzen und Metadaten
```

### Google Drive Struktur
```
Club der alten Säcke/
├── Protokolle/         # Google Docs und PDFs
└── Protokoll-Bilder/   # Hochgeladene Bilder
```

## 🚀 Installation

### 1. Google Apps Script Projekt erstellen
1. Gehe zu [script.google.com](https://script.google.com)
2. Erstelle ein neues Projekt
3. Kopiere die Dateien in entsprechende Script-Dateien

### 2. Google Sheets vorbereiten
```javascript
// Einmalig ausführen (SupportServices.gs):
setupSheetsStructure()
```

### 3. Web App deployen
1. **Deploy** → **New deployment**
2. **Type**: Web app
3. **Execute as**: Me
4. **Who has access**: Anyone
5. URL notieren für Zugriff

### 4. Berechtigungen erteilen
- Google Sheets (Lesen/Schreiben)
- Google Drive (Dateien erstellen)
- Gmail (E-Mails senden)

## 📧 Multi-E-Mail Konfiguration

### Mitglieder-Sheet Format
```
ID | Name              | E-Mail                                    | Aktiv
1  | Josef Schnürer    | josef.schnuerer@gmail.com;josef@aon.at  | TRUE
2  | Gerald Faic       | gerald@work.com;gerald@private.com      | TRUE
```

### E-Mail-Verhalten
- **Versand**: An ALLE E-Mail-Adressen gleichzeitig
- **Trennung**: Semikolon (`;`) zwischen Adressen
- **Validierung**: Automatische Filterung ungültiger E-Mails
- **Statistik**: Detaillierte Protokollierung im Sheet

## 🔧 Wichtige Funktionen

### Auto-Resume
```javascript
// Wird automatisch beim Seitenaufruf ausgeführt (BusinessServices.gs)
getLastDraftProtocol()
```

### E-Mail Management
```javascript
// Wartende E-Mails nachsenden (SupportServices.gs)
sendPendingEmails()

// Gmail-Quota prüfen (SupportServices.gs)
checkGmailQuota()

// Wartende Protokolle anzeigen (code.gs)
showPendingEmails()
```

### 4. System-Wartung
```javascript
// Alle Testdaten löschen (in SupportServices.gs)
clearAllTestData()

// Status aller Protokolle anzeigen (in CoreServices.gs)
showAllProtocolStatuses()

// Bild-Referenzen reparieren (in SupportServices.gs)
repairExistingImageReferences()
```

### Test-Funktionen
```javascript
// Vollständiger Test - Draft → PDF → E-Mail (SupportServices.gs)
testCompleteProtocolToPDF()

// Multi-E-Mail System testen (SupportServices.gs)
testMultiEmailSystem()

// Protokoll-ID Logik testen (SupportServices.gs)
testProtocolIdLogic()
```

## 📱 Mobile Optimierung

### iPhone/Mobile Features
- **Breakpoint**: 1000px für frühe mobile Aktivierung
- **Anti-Zoom**: Input-Felder mit 22px+ Schriftgröße
- **Touch-Targets**: Mindestens 44px für alle interaktiven Elemente
- **Viewport**: `user-scalable=no` verhindert ungewolltes Zoomen

### Responsive Verhalten
```css
@media (max-width: 1000px) {
  /* Einspaltige Layouts */
  /* Große Schriftgrößen (22px+) */
  /* Touch-optimierte Buttons (64px+) */
}
```

## ⚠️ Gmail Quota Handling

### Problem
Google Apps Script hat ein Limit von ~500 E-Mails pro Tag.

### Lösung
```javascript
// Bei Quota-Erschöpfung:
Status: "PDF erstellt - E-Mail wartend"

// Nächsten Tag ausführen:
sendPendingEmails()
```

### Quota-Status prüfen
```javascript
checkGmailQuota()  // Verfügbarkeit testen
getQuotaStatistics()  // Tagesverbrauch anzeigen (falls implementiert)
```

## 🔒 Sicherheit & Berechtigungen

### Erforderliche OAuth Scopes
- `https://www.googleapis.com/auth/spreadsheets`
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/gmail.send`
- `https://www.googleapis.com/auth/documents`

### Datenschutz
- Alle Daten bleiben in der Google Workspace des Administrators
- Keine externen APIs außer Google Services
- Mitglieder-E-Mails werden nur für Protokoll-Versand verwendet

## 🐛 Troubleshooting

### E-Mail wird nicht versendet
```javascript
// 1. Quota prüfen (SupportServices.gs)
checkGmailQuota()

// 2. Wartende E-Mails anzeigen (code.gs)
showPendingEmails()

// 3. Nachsenden versuchen (SupportServices.gs)
sendPendingEmails()
```

### Auto-Resume funktioniert nicht
```javascript
// Debug-Informationen (code.gs)
debugServices()
manualFrontendDebugTest()
```

### Bilder werden nicht angezeigt
```javascript
// Bild-Referenzen reparieren (SupportServices.gs)
repairExistingImageReferences()
```

### System zurücksetzen
```javascript
// VORSICHT: Löscht alle Protokolle! (SupportServices.gs)
clearAllTestData()
```

## 📊 Status-Überwachung

### Protokoll-Status
- **`Draft`**: Entwurf gespeichert
- **`Final`**: Finalisiert, Dokumente werden erstellt
- **`PDF erstellt - E-Mail wartend`**: PDF fertig, E-Mail bei Quota-Problem wartend
- **`Versendet`**: Komplett abgeschlossen

### Monitoring
```javascript
// Alle Protokoll-Status anzeigen (CoreServices.gs)
showAllProtocolStatuses()

// Detaillierte Logs in Sheet Spalte "Notizen"
```

## 🔄 Workflow

1. **Seitenaufruf**: Auto-Resume lädt letzten Draft (falls vorhanden)
2. **Dateneingabe**: Sitzungstyp, Teilnehmer, Tagesordnungspunkte
3. **Zwischenspeichern**: Optional, automatisch vor Finalisierung
4. **Finalisierung**: 
   - Google Doc erstellen
   - PDF generieren
   - E-Mail versenden
   - Status aktualisieren

## 📈 Erweiterungen

### Geplante Features
- [ ] KI-Textverbesserung Integration
- [ ] Automatische Terminberechnung
- [ ] E-Mail Templates
- [ ] Backup-System
- [ ] erweiterte Statistiken

### Anpassungen
Das System ist modular aufgebaut und kann einfach erweitert werden:
- Neue Services in `SupportServices.gs`
- Frontend-Änderungen in `index.html`
- Neue API-Endpoints in `code.gs`

## 📞 Support

Bei Problemen oder Fragen:
1. Teste die eingebauten Debug-Funktionen  
2. Prüfe die Browser-Konsole auf JavaScript-Fehler
3. Kontrolliere die Google Apps Script Logs
4. Überprüfe die Sheet-Struktur mit `setupSheetsStructure()` (SupportServices.gs)

---

**Entwickelt für den Club der alten Säcke (CDAS)**  
*Version 2.1 - Juli 2025*
