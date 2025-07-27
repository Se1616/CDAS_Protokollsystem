// ==================== CORE SERVICES ====================

// Drive & Ordner Service
const DriveService = {
  
  // Club-Hauptordner finden/erstellen
  getClubFolder() {
    const folders = DriveApp.getFoldersByName(CONFIG.CLUB_FOLDER_NAME);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.CLUB_FOLDER_NAME);
  },
  
  // Unterordner finden/erstellen
  getSubFolder(parentFolder, folderName) {
    const folders = parentFolder.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
  },
  
  // Protokolle-Ordner
  getProtocolsFolder() {
    const clubFolder = this.getClubFolder();
    return this.getSubFolder(clubFolder, CONFIG.PROTOCOLS_FOLDER);
  },
  
  // Bilder-Ordner
  getImagesFolder() {
    const clubFolder = this.getClubFolder();
    return this.getSubFolder(clubFolder, CONFIG.IMAGES_FOLDER);
  }
};

// Google Sheets Service
const SheetService = {
  
  // Club-Spreadsheet finden
  getClubSpreadsheet() {
    const clubFolder = DriveService.getClubFolder();
    const files = clubFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().includes('CDAS') || file.getName().includes('Protokoll')) {
        return SpreadsheetApp.open(file);
      }
    }
    
    return SpreadsheetApp.getActiveSpreadsheet();
  },
  
  // Sheet by Name holen
  getSheet(sheetName) {
    const ss = this.getClubSpreadsheet();
    return ss.getSheetByName(sheetName);
  },
  
  // Daten aus Sheet laden
  getSheetData(sheetName) {
    const sheet = this.getSheet(sheetName);
    return sheet ? sheet.getDataRange().getValues() : [];
  }
};

// ID Generator Service
const IdService = {
  
  generateProtocolId() {
    return 'PROT_' + new Date().getTime();
  },
  
  generateAgendaId() {
    return 'AGD_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 9);
  },
  
  generateImageId() {
    return 'IMG_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
  }
};

// Utility Service
const UtilService = {
  
  // Datum formatieren (deutsch)
  formatDateDE(date) {
    return new Date(date).toLocaleDateString('de-DE');
  },
  
  // Dokument-Name generieren
  generateDocumentName(formData) {
    const date = new Date(formData.sitzungsdatum);
    const dateStr = `${date.getDate().toString().padStart(2, '0')}.${(date.getMonth() + 1).toString().padStart(2, '0')}.${date.getFullYear()}`;
    const totalNumber = formData.counters.ordentlich + formData.counters.ausserordentlich + formData.counters.fest;
    return `CDAS Protokoll Nr. ${totalNumber} - ${dateStr}`;
  },
  
  // Letzten Freitag berechnen
  getLastFriday(date = new Date()) {
    const localDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const dayOfWeek = localDate.getDay();
    
    if (dayOfWeek === 5) return localDate; // Heute ist Freitag
    
    const daysBack = { 0: 2, 1: 3, 2: 4, 3: 5, 4: 6, 6: 1 }[dayOfWeek] || 1;
    const lastFriday = new Date(localDate);
    lastFriday.setDate(localDate.getDate() - daysBack);
    return lastFriday;
  },
  
  // Nächsten ersten Freitag berechnen
  getNextFirstFriday() {
    const today = new Date();
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const dayOfWeek = nextMonth.getDay();
    
    const daysToAdd = { 0: 5, 1: 4, 2: 3, 3: 2, 4: 1, 5: 0, 6: 6 }[dayOfWeek] || 0;
    nextMonth.setDate(1 + daysToAdd);
    return nextMonth;
  }
};

// Status Logging Service
const StatusService = {
  
  // Status zu Protokoll hinzufügen - KOMPAKT
  addStatusToNotes(protocolId, statusMessage) {
    try {
      const sheet = SheetService.getSheet('Protokolle');
      const rowIndex = this.findProtocolRow(protocolId);
      
      if (rowIndex > 0) {
        const currentNotes = sheet.getRange(rowIndex, 15).getValue() || '';
        const timestamp = new Date().toLocaleString('de-DE');
        const newEntry = `[${timestamp}] ${statusMessage}`;
        
        // KOMPAKT: Begrenzte Anzahl von Einträgen behalten
        const lines = currentNotes ? currentNotes.split('\n') : [];
        lines.push(newEntry);
        
        // NUR die letzten 20 Einträge behalten
        const maxLines = 20;
        if (lines.length > maxLines) {
          const keptLines = lines.slice(-maxLines);
          const updatedNotes = keptLines.join('\n');
          sheet.getRange(rowIndex, 15).setValue(updatedNotes);
        } else {
          const updatedNotes = lines.join('\n');
          sheet.getRange(rowIndex, 15).setValue(updatedNotes);
        }
        
        console.log(`✅ Status hinzugefügt: ${statusMessage}`);
      }
    } catch (error) {
      console.error('❌ Fehler beim Status-Logging:', error);
    }
  },
  
  // Protokoll-Zeile finden
  findProtocolRow(protocolId) {
    const data = SheetService.getSheetData('Protokolle');
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === protocolId) return i + 1;
    }
    return -1;
  },
  
  // Protokoll-Status aktualisieren
  updateProtocolStatus(protocolId, status, docId = '', pdfId = '') {
    const sheet = SheetService.getSheet('Protokolle');
    const rowIndex = this.findProtocolRow(protocolId);
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 3).setValue(status);
      if (docId) sheet.getRange(rowIndex, 12).setValue(docId);
      if (pdfId) sheet.getRange(rowIndex, 13).setValue(pdfId);
      if (status === 'Versendet') {
        sheet.getRange(rowIndex, 14).setValue(new Date());
      }
    }
  },
  
  // Status-Übersicht anzeigen - KOMPAKT
  showAllProtocolStatuses() {
    try {
      const data = SheetService.getSheetData('Protokolle');
      console.log('=== PROTOKOLL-STATUS ÜBERSICHT ===');
      
      for (let i = 1; i < data.length; i++) {
        const [protocolId, createdAt, status] = data[i];
        const dateStr = createdAt ? new Date(createdAt).toLocaleDateString('de-DE') : 'N/A';
        console.log(`📋 ${protocolId} (${dateStr}) - Status: ${status}`);
      }
      console.log('=== ENDE ÜBERSICHT ===');
    } catch (error) {
      console.error('Fehler beim Anzeigen der Status:', error);
    }
  }
};
