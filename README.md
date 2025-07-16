# Simple Outlook Plugin

Ein einfaches Outlook Online Plugin zu Testzwecken. Dieses Plugin demonstriert grundlegende Funktionen der Outlook Add-in API.

## Features

- 📝 **Text einfügen**: Fügt einen vordefinierten Text in die aktuelle E-Mail ein
- 📧 **E-Mail-Info anzeigen**: Zeigt Informationen über die aktuelle E-Mail an
- 💬 **Nachricht anzeigen**: Zeigt eine zufällige Testnachricht an
- 👤 **Benutzer-Info**: Zeigt Informationen über den aktuellen Benutzer

## Dateien

- `manifest.xml` - Plugin-Manifest mit Konfiguration
- `taskpane.html` - Hauptinterface des Plugins
- `taskpane.css` - Styling für das Interface
- `taskpane.js` - JavaScript-Logik
- `package.json` - Projekt-Konfiguration
- `README.md` - Diese Dokumentation

## Installation & Test

### Methode 1: Lokaler HTTP-Server

1. **Server starten** (wählen Sie eine Option):
   ```bash
   # Option A: Mit Python
   python -m http.server 8080
   
   # Option B: Mit Node.js
   npx http-server -p 8080 -c-1
   
   # Option C: Mit NPM Script
   npm run serve-node
   ```

2. **Manifest-URL**: `http://localhost:8080/manifest.xml`

### Methode 2: Online-Hosting

Laden Sie alle Dateien auf einen Webserver hoch und verwenden Sie die entsprechende URL für das Manifest.

## Plugin in Outlook Online laden

1. **Gehen Sie zu Outlook Online** (https://outlook.office.com)
2. **Klicken Sie auf "Einstellungen"** (Zahnrad-Symbol)
3. **Wählen Sie "Add-Ins verwalten"**
4. **Klicken Sie auf "Benutzerdefinierte Add-Ins"**
5. **Wählen Sie "Aus URL hinzufügen"**
6. **Geben Sie die Manifest-URL ein**: `http://localhost:8080/manifest.xml`
7. **Bestätigen Sie die Installation**

## Plugin verwenden

1. **Öffnen Sie eine E-Mail** oder **erstellen Sie eine neue E-Mail**
2. **Suchen Sie nach dem "Test Plugin" Button** im Outlook-Ribbon
3. **Klicken Sie auf "Plugin öffnen"**
4. **Testen Sie die verschiedenen Funktionen**:
   - Text einfügen (nur beim Schreiben von E-Mails)
   - E-Mail-Informationen anzeigen
   - Testnachrichten anzeigen

## Entwicklung

### Struktur verstehen

- **Manifest**: Definiert Plugin-Eigenschaften und Berechtigungen
- **Taskpane**: Das Hauptinterface, das als seitliches Panel erscheint
- **Office.js**: Microsoft's JavaScript-API für Office-Interaktionen

### Anpassungen

1. **Manifest anpassen**: Ändern Sie `manifest.xml` für andere Konfigurationen
2. **Interface erweitern**: Erweitern Sie `taskpane.html` und `taskpane.css`
3. **Funktionalität hinzufügen**: Erweitern Sie `taskpane.js` mit neuen Features

### Debugging

- Verwenden Sie die Browser-Entwicklertools
- Überprüfen Sie die Konsole auf Fehler
- Testen Sie in verschiedenen Browsern

## Häufige Probleme

### Plugin wird nicht geladen
- Überprüfen Sie, ob der HTTP-Server läuft
- Stellen Sie sicher, dass alle Dateien erreichbar sind
- Überprüfen Sie die Manifest-URL

### Funktionen funktionieren nicht
- Überprüfen Sie die Browser-Konsole auf JavaScript-Fehler
- Stellen Sie sicher, dass Sie die richtige Outlook-Version verwenden
- Testen Sie mit verschiedenen E-Mail-Typen (neu/antworten/weiterleiten)

### Plugin ist nicht sichtbar
- Überprüfen Sie, ob das Plugin richtig installiert wurde
- Aktualisieren Sie die Outlook-Seite
- Überprüfen Sie die Plugin-Einstellungen

## Weiterführende Informationen

- [Office Add-ins Dokumentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Outlook Add-ins API](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/)
- [Office.js API Referenz](https://docs.microsoft.com/en-us/javascript/api/office)

## Lizenz

MIT License - Sie können diesen Code frei verwenden und anpassen. 