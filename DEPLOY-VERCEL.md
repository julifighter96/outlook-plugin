# Vercel Deployment Guide

## 🚀 Schnelle Deployment-Anleitung

### 1. Vorbereitungen
- Sie benötigen einen GitHub Account
- Sie benötigen einen Vercel Account (kostenlos)

### 2. Repository erstellen

1. **Gehen Sie zu GitHub**: https://github.com/new
2. **Erstellen Sie ein neues Repository**:
   - Name: `outlook-plugin` (oder ein anderer Name)
   - Wählen Sie "Public" 
   - Erstellen Sie das Repository

### 3. Dateien hochladen

Laden Sie diese Dateien in Ihr GitHub Repository hoch:

**Wichtige Dateien:**
- `standalone-plugin.html` - Das Plugin Interface
- `manifest-vercel.xml` - Das Manifest für Vercel
- `vercel.json` - Vercel Konfiguration
- `package.json` - Projekt Konfiguration
- `.gitignore` - Git ignore Datei

**Optionale Dateien:**
- `README.md` - Dokumentation
- `manifest.xml`, `taskpane.html`, `taskpane.css`, `taskpane.js` - Originale Dateien

### 4. Vercel Deployment

1. **Gehen Sie zu Vercel**: https://vercel.com
2. **Registrieren/Anmelden** mit Ihrem GitHub Account
3. **Klicken Sie auf "New Project"**
4. **Wählen Sie Ihr GitHub Repository** aus
5. **Deployment-Einstellungen**:
   - Framework Preset: `Other`
   - Build Command: `npm run build` (optional)
   - Output Directory: `./` (Root-Verzeichnis)
   - Install Command: `npm install` (optional)
6. **Klicken Sie auf "Deploy"**

### 5. Vercel URL erhalten

Nach dem Deployment erhalten Sie eine URL wie:
`https://your-project-name.vercel.app`

### 6. Manifest aktualisieren

1. **Öffnen Sie `manifest-vercel.xml`**
2. **Ersetzen Sie alle Instanzen von `YOUR-VERCEL-URL`** mit Ihrer echten Vercel-URL
3. **Beispiel**: 
   ```xml
   <!-- Vorher -->
   <SourceLocation DefaultValue="https://YOUR-VERCEL-URL.vercel.app/standalone-plugin.html"/>
   
   <!-- Nachher -->
   <SourceLocation DefaultValue="https://outlook-plugin-123.vercel.app/standalone-plugin.html"/>
   ```

### 7. Aktualisiertes Manifest hochladen

1. **Committen Sie die Änderungen** in GitHub
2. **Vercel deployed automatisch** die Änderungen
3. **Ihre Manifest-URL ist jetzt**: `https://ihre-url.vercel.app/manifest-vercel.xml`

## 📋 Plugin in Outlook Online installieren

1. **Gehen Sie zu Outlook Online**: https://outlook.office.com
2. **Einstellungen** ⚙️ → **"Add-Ins verwalten"**
3. **"Benutzerdefinierte Add-Ins"** → **"Aus Datei hinzufügen"**
4. **Laden Sie die aktualisierte `manifest-vercel.xml` Datei hoch**

## 🔧 Alternative: Direkter Vercel Upload

Falls Sie kein GitHub verwenden möchten:

1. **Installieren Sie Vercel CLI**:
   ```bash
   npm install -g vercel
   ```

2. **Führen Sie im Projekt-Ordner aus**:
   ```bash
   vercel login
   vercel --prod
   ```

3. **Folgen Sie den Anweisungen**

## ✅ Testen

1. **Öffnen Sie Ihre Vercel-URL** im Browser: `https://ihre-url.vercel.app/standalone-plugin.html`
2. **Prüfen Sie, ob das Plugin lädt**
3. **Installieren Sie das Plugin in Outlook Online**
4. **Testen Sie die Funktionen**

## 🔄 Updates

Für Updates:
1. **Ändern Sie Dateien** in GitHub
2. **Vercel deployed automatisch**
3. **Kein Neustart nötig**

## 🆘 Troubleshooting

### Plugin lädt nicht
- Überprüfen Sie die Vercel-URL im Browser
- Stellen Sie sicher, dass `manifest-vercel.xml` die korrekte URL hat
- Prüfen Sie die Browser-Konsole auf Fehler

### CORS-Fehler
- Die `vercel.json` Datei ist konfiguriert für CORS
- Prüfen Sie, ob die Datei richtig hochgeladen wurde

### Manifest-Fehler
- Stellen Sie sicher, dass alle URLs in `manifest-vercel.xml` korrekt sind
- Überprüfen Sie die XML-Syntax

## 📞 Support

Bei Problemen:
1. Prüfen Sie die Vercel Dashboard Logs
2. Testen Sie die URL direkt im Browser
3. Überprüfen Sie die Outlook Plugin-Logs

**Ihre finale Manifest-URL wird sein**: `https://ihre-vercel-url.vercel.app/manifest-vercel.xml` 