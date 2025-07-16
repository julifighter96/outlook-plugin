/* global Office */

// Office-Initialisierung
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById('insertTextButton').onclick = insertText;
        document.getElementById('getEmailInfoButton').onclick = getEmailInfo;
        document.getElementById('showMessageButton').onclick = showMessage;
        
        // Benutzerinformationen laden
        loadUserInfo();
        
        updateStatus('Plugin bereit');
    }
});

// Benutzerinformationen laden
function loadUserInfo() {
    try {
        Office.context.mailbox.getUserIdentityTokenAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                // Vereinfachte Benutzerinformationen
                document.getElementById('userName').textContent = Office.context.mailbox.userProfile.displayName || 'Unbekannt';
                document.getElementById('userEmail').textContent = Office.context.mailbox.userProfile.emailAddress || 'Unbekannt';
            } else {
                document.getElementById('userName').textContent = 'Fehler beim Laden';
                document.getElementById('userEmail').textContent = 'Fehler beim Laden';
            }
        });
    } catch (error) {
        // Fallback für lokale Tests
        document.getElementById('userName').textContent = 'Test User';
        document.getElementById('userEmail').textContent = 'test@example.com';
    }
}

// Text in E-Mail einfügen
function insertText() {
    updateStatus('Text wird eingefügt...');
    
    try {
        // Prüfen ob wir im Compose-Modus sind
        if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
            const textToInsert = "Hallo! Dies ist ein Test-Text vom Outlook Plugin.\n\nMit freundlichen Grüßen,\nIhr Plugin";
            
            // Text an der aktuellen Cursor-Position einfügen
            Office.context.mailbox.item.body.setSelectedDataAsync(
                textToInsert,
                { coercionType: Office.CoercionType.Text },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        updateOutput('✅ Text erfolgreich eingefügt!');
                        updateStatus('Text eingefügt');
                    } else {
                        updateOutput('❌ Fehler beim Einfügen: ' + result.error.message);
                        updateStatus('Fehler');
                    }
                }
            );
        } else {
            updateOutput('⚠️ Diese Funktion ist nur beim Schreiben von E-Mails verfügbar.');
            updateStatus('Nicht verfügbar');
        }
    } catch (error) {
        updateOutput('❌ Fehler: ' + error.message);
        updateStatus('Fehler');
    }
}

// E-Mail-Informationen anzeigen
function getEmailInfo() {
    updateStatus('E-Mail-Informationen werden geladen...');
    
    try {
        const item = Office.context.mailbox.item;
        let info = "📧 E-Mail-Informationen:\n\n";
        
        // Betreff
        info += `Betreff: ${item.subject || 'Kein Betreff'}\n`;
        
        // Absender/Empfänger
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            if (item.from) {
                info += `Von: ${item.from.displayName} (${item.from.emailAddress})\n`;
            }
            if (item.to && item.to.length > 0) {
                info += `An: ${item.to.map(recipient => recipient.displayName || recipient.emailAddress).join(', ')}\n`;
            }
        }
        
        // Datum/Zeit
        if (item.dateTimeCreated) {
            info += `Erstellt: ${item.dateTimeCreated.toLocaleString('de-DE')}\n`;
        }
        
        // Item-Typ
        info += `Typ: ${item.itemType === Office.MailboxEnums.ItemType.Message ? 'E-Mail' : 'Termin'}\n`;
        
        // Aktueller Modus
        if (item.itemType === Office.MailboxEnums.ItemType.Message) {
            const mode = Office.context.mailbox.item.itemClass.indexOf('IPM.Note') >= 0 ? 'Lese-Modus' : 'Schreib-Modus';
            info += `Modus: ${mode}\n`;
        }
        
        updateOutput(info);
        updateStatus('Informationen geladen');
    } catch (error) {
        updateOutput('❌ Fehler beim Laden der E-Mail-Informationen: ' + error.message);
        updateStatus('Fehler');
    }
}

// Nachricht anzeigen
function showMessage() {
    updateStatus('Nachricht wird angezeigt...');
    
    const messages = [
        "🎉 Herzlichen Glückwunsch! Das Plugin funktioniert!",
        "🚀 Ihr Outlook Plugin ist erfolgreich geladen!",
        "💡 Tipp: Verwenden Sie die anderen Buttons um weitere Funktionen zu testen.",
        "📝 Dieses Plugin wurde zu Testzwecken erstellt.",
        "🔧 Sie können den Code in den Dateien manifest.xml, taskpane.html, taskpane.css und taskpane.js anpassen."
    ];
    
    const randomMessage = messages[Math.floor(Math.random() * messages.length)];
    
    updateOutput(randomMessage);
    updateStatus('Nachricht angezeigt');
}

// Ausgabe aktualisieren
function updateOutput(message) {
    const outputElement = document.getElementById('output');
    const timestamp = new Date().toLocaleTimeString('de-DE');
    outputElement.textContent = `[${timestamp}] ${message}`;
    
    // Scroll nach unten
    outputElement.scrollTop = outputElement.scrollHeight;
}

// Status aktualisieren
function updateStatus(status) {
    document.getElementById('status').textContent = status;
}

// Fehlerbehandlung
window.addEventListener('error', (event) => {
    updateOutput('❌ JavaScript-Fehler: ' + event.error.message);
    updateStatus('Fehler');
});

// Debugging-Funktionen (nur für Entwicklung)
if (typeof console !== 'undefined' && console.log) {
    console.log('Outlook Plugin geladen');
    console.log('Office-Kontext:', Office.context);
} 