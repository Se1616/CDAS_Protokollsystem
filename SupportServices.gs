// ==================== SUPPORT SERVICES ====================

// Image Service
const ImageService = {
  
  // Multiple Bilder hochladen
  uploadMultipleImages(imagesData, agendaItemId) {
    const results = [];
    
    console.log(`Uploading ${imagesData.length} images for agenda: ${agendaItemId}`);
    
    for (let i = 0; i < imagesData.length; i++) {
      const imageData = imagesData[i];
      const result = this.uploadImage(imageData.data, imageData.name, agendaItemId);
      results.push(result);
    }
    
    const successCount = results.filter(r => r.success).length;
    const failedCount = results.filter(r => !r.success).length;
    
    // Status in Protokoll vermerken
    this.logImageUploadStatus(agendaItemId, successCount, failedCount);
    
    return {
      success: failedCount === 0,
      results: results,
      totalUploaded: successCount,
      totalFailed: failedCount,
      message: `${successCount} Bilder hochgeladen, ${failedCount} fehlgeschlagen`
    };
  },
  
  // Einzelnes Bild hochladen
  uploadImage(imageData, fileName, agendaItemId) {
    try {
      const imagesFolder = DriveService.getImagesFolder();
      
      const blob = Utilities.newBlob(
        Utilities.base64Decode(imageData.split(',')[1]),
        this.getMimeType(fileName),
        fileName
      );
      
      const compressedBlob = this.compressImage(blob);
      const file = imagesFolder.createFile(compressedBlob);
      const fileId = file.getId();
      
      const imageId = this.saveImageReference(fileId, fileName, agendaItemId, compressedBlob.getBytes().length);
      
      console.log('Bild erfolgreich hochgeladen:', fileName, 'FileID:', fileId);
      
      return {
        success: true,
        imageId: imageId,
        fileId: fileId,
        fileName: fileName,
        url: `https://drive.google.com/file/d/${fileId}/view`,
        thumbnailUrl: `https://drive.google.com/thumbnail?id=${fileId}&sz=w200-h150`
      };
      
    } catch (error) {
      console.error('Fehler beim Bild-Upload:', error);
      return {
        success: false,
        message: 'Fehler beim Upload: ' + error.toString()
      };
    }
  },
  
  // MIME-Type bestimmen
  getMimeType(fileName) {
    const extension = fileName.toLowerCase().split('.').pop();
    const mimeTypes = {
      'jpg': 'image/jpeg',
      'jpeg': 'image/jpeg',
      'png': 'image/png',
      'gif': 'image/gif',
      'webp': 'image/webp',
      'heic': 'image/heic'
    };
    return mimeTypes[extension] || 'image/jpeg';
  },
  
  // Bild komprimieren (Basic)
  compressImage(blob) {
    if (blob.getBytes().length <= CONFIG.MAX_IMAGE_SIZE) {
      return blob;
    }
    // TODO: Echte Komprimierung implementieren falls nötig
    return blob;
  },
  
  // Bild-Referenz in Sheet speichern
  saveImageReference(fileId, fileName, agendaItemId, fileSize) {
    const sheet = SheetService.getSheet('Bilder');
    
    const imageId = IdService.generateImageId();
    const existingImages = this.getImagesForAgendaItem(agendaItemId);
    const reihenfolge = existingImages.length + 1;
    
    const row = [
      imageId,
      agendaItemId,
      fileName,
      fileId,
      Math.round(fileSize / 1024),
      reihenfolge,
      new Date()
    ];
    
    sheet.appendRow(row);
    this.updateAgendaItemHasImages(agendaItemId, true);
    
    return imageId;
  },
  
  // Bilder für Tagesordnungspunkt laden
  getImagesForAgendaItem(agendaItemId) {
    const data = SheetService.getSheetData('Bilder');
    const images = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === agendaItemId) {
        images.push({
          id: data[i][0],
          agendaId: data[i][1],
          fileName: data[i][2],
          driveId: data[i][3],
          sizeKB: data[i][4],
          reihenfolge: data[i][5],
          uploadedAt: data[i][6] ? data[i][6].toISOString() : null  // ← DATE-FIX!
        });
      }
    }
    
    images.sort((a, b) => a.reihenfolge - b.reihenfolge);
    return images;
  },
  
  // Hat_Bilder Flag aktualisieren
  updateAgendaItemHasImages(agendaItemId, hasImages) {
    const sheet = SheetService.getSheet('Tagesordnungspunkte');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === agendaItemId) {
        sheet.getRange(i + 1, 9).setValue(hasImages);
        break;
      }
    }
  },
  
  // Bild löschen
  deleteImage(imageId) {
    try {
      const sheet = SheetService.getSheet('Bilder');
      const data = sheet.getDataRange().getValues();
      
      let driveId = null;
      let rowToDelete = -1;
      let agendaItemId = null;
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === imageId) {
          driveId = data[i][3];
          agendaItemId = data[i][1];
          rowToDelete = i + 1;
          break;
        }
      }
      
      if (rowToDelete > 0) {
        if (driveId) {
          try {
            DriveApp.getFileById(driveId).setTrashed(true);
          } catch (e) {
            console.log('Drive-Datei bereits gelöscht:', driveId);
          }
        }
        
        sheet.deleteRow(rowToDelete);
        
        const remainingImages = this.getImagesForAgendaItem(agendaItemId);
        if (remainingImages.length === 0) {
          this.updateAgendaItemHasImages(agendaItemId, false);
        }
        
        return { success: true, message: 'Bild erfolgreich gelöscht' };
      } else {
        return { success: false, message: 'Bild nicht gefunden' };
      }
      
    } catch (error) {
      console.error('Fehler beim Löschen:', error);
      return { success: false, message: 'Fehler beim Löschen: ' + error.toString() };
    }
  },
  
  // Bild-Referenzen korrigieren
  fixImageReferences(agendaIdMapping) {
    try {
      const sheet = SheetService.getSheet('Bilder');
      const data = sheet.getDataRange().getValues();
      
      let updatedCount = 0;
      
      console.log('📋 Verfügbare ID-Mappings:', agendaIdMapping);
      
      for (let i = 1; i < data.length; i++) {
        const currentFrontendId = data[i][1];
        
        console.log(`🔍 Prüfe Bild in Zeile ${i + 1}: Frontend-ID = "${currentFrontendId}"`);
        
        if (agendaIdMapping[currentFrontendId]) {
          const realAgendaId = agendaIdMapping[currentFrontendId];
          
          sheet.getRange(i + 1, 2).setValue(realAgendaId);
          
          console.log(`🔄 Bild-ID korrigiert: ${currentFrontendId} → ${realAgendaId}`);
          updatedCount++;
        } else {
          console.log(`⚠️ Kein Mapping gefunden für: "${currentFrontendId}"`);
        }
      }
      
      console.log(`✅ ${updatedCount} Bild-Referenzen korrigiert`);
      return { success: true, updatedCount: updatedCount };
      
    } catch (error) {
      console.error('❌ Fehler beim Korrigieren der Bild-Referenzen:', error);
      return { success: false, message: error.toString() };
    }
  },
  
  // Status für Bild-Upload protokollieren
  logImageUploadStatus(agendaItemId, successCount, failedCount) {
    try {
      const agendaData = SheetService.getSheetData('Tagesordnungspunkte');
      
      for (let i = 1; i < agendaData.length; i++) {
        if (agendaData[i][0] === agendaItemId) {
          const protocolId = agendaData[i][1];
          StatusService.addStatusToNotes(protocolId, `📷 Bilder hochgeladen: ${successCount} erfolgreich, ${failedCount} fehlgeschlagen`);
          break;
        }
      }
    } catch (error) {
      console.log('Status-Logging für Bilder übersprungen:', error.message);
    }
  }
};

const EmailService = {
  
  // VEREINFACHT: Protokoll-E-Mail versenden - ohne unnötige Retries
  sendProtocolEmail(protocolId, pdfId, formData) {
    try {
      StatusService.addStatusToNotes(protocolId, '📧 Starte E-Mail-Versand...');
      
      const pdfFile = DriveApp.getFileById(pdfId);
      const pdfBlob = pdfFile.getBlob();
      const protocolName = UtilService.generateDocumentName(formData);
      
      // Alle E-Mail-Adressen sammeln
      const activeMembers = MemberService.getActiveMembers();
      const allEmailAddresses = MemberService.getAllEmailAddresses();
      const emailStats = MemberService.getEmailStatistics();
      
      // KOMPAKTE Protokollierung
      StatusService.addStatusToNotes(protocolId, `📧 ${emailStats.totalMembers} Mitglieder, ${emailStats.totalEmailAddresses} E-Mail-Adressen`);
      
      const emailBody = this.generateEmailBody();
      
      try {
        // DIREKTER E-Mail-Versand - ohne Retries
        console.log(`📧 Versende an ${allEmailAddresses.length} E-Mail-Adressen...`);
        
        GmailApp.sendEmail(
          allEmailAddresses.join(','),
          protocolName,
          emailBody,
          {
            from: CONFIG.SENDER_EMAIL,
            attachments: [pdfBlob],
            name: CONFIG.SENDER_NAME
          }
        );
        
        // ERFOLG
        StatusService.updateProtocolStatus(protocolId, 'Versendet');
        StatusService.addStatusToNotes(protocolId, `✅ E-Mail erfolgreich versendet an ${emailStats.totalEmailAddresses} Adressen`);
        
        return {
          success: true,
          recipients: emailStats.totalMembers,
          totalEmailAddresses: emailStats.totalEmailAddresses,
          subject: protocolName,
          message: `E-Mail erfolgreich an ${emailStats.totalMembers} Mitglieder (${emailStats.totalEmailAddresses} E-Mail-Adressen) versendet`
        };
        
      } catch (emailError) {
        // QUOTA-FEHLER? → Status setzen und fertig
        if (this.isQuotaError(emailError)) {
          
          StatusService.updateProtocolStatus(protocolId, 'PDF erstellt - E-Mail wartend');
          StatusService.addStatusToNotes(protocolId, `❌ Gmail-Quota erreicht - E-Mail wird morgen versendet`);
          
          return {
            success: false,
            isQuotaError: true,
            quotaExhausted: true,
            pdfReady: true,
            message: `Gmail-Quota erreicht. PDF ist bereit, E-Mail wird morgen automatisch versendet.`,
            manualSendInstructions: `Führe morgen sendPendingEmails() aus`
          };
          
        } else {
          // ANDERER E-MAIL-FEHLER
          throw emailError;
        }
      }
      
    } catch (error) {
      console.error('❌ Fehler beim E-Mail-Versand:', error);
      StatusService.addStatusToNotes(protocolId, `❌ E-Mail FEHLER: ${error.toString()}`);
      return {
        success: false,
        message: 'Fehler beim E-Mail-Versand: ' + error.toString()
      };
    }
  },
  
  // Quota-Fehler erkennen
  isQuotaError(error) {
    const errorMsg = error.toString().toLowerCase();
    return errorMsg.includes('quota') || 
           errorMsg.includes('limit') || 
           errorMsg.includes('zu häufig') || 
           errorMsg.includes('too many') ||
           errorMsg.includes('rate limit') ||
           errorMsg.includes('dienst an einem tag zu häufig');
  },
  
  // VEREINFACHT: Wartende E-Mails versenden (für den nächsten Tag)
  sendPendingEmails() {
    try {
      console.log('🔍 Suche wartende E-Mails...');
      
      const data = SheetService.getSheetData('Protokolle');
      const pendingProtocols = [];
      
      // Wartende Protokolle finden
      for (let i = 1; i < data.length; i++) {
        const status = data[i][2]; // Spalte C: Status
        if (status === 'PDF erstellt - E-Mail wartend') {
          pendingProtocols.push({
            id: data[i][0],
            pdfId: data[i][13], // Spalte N: PDF_Drive_ID
            sitzungsdatum: data[i][3],
            ort: data[i][4],
            counters: {
              ordentlich: data[i][5],
              ausserordentlich: data[i][6],
              fest: data[i][7]
            }
          });
        }
      }
      
      if (pendingProtocols.length === 0) {
        console.log('ℹ️ Keine wartenden E-Mails gefunden');
        return {
          success: true,
          message: 'Keine wartenden E-Mails gefunden',
          processedCount: 0
        };
      }
      
      console.log(`📧 ${pendingProtocols.length} wartende E-Mails gefunden`);
      
      let successCount = 0;
      let stillFailedCount = 0;
      
      // Jedes wartende Protokoll verarbeiten
      for (let protocol of pendingProtocols) {
        try {
          console.log(`📧 Verarbeite wartende E-Mail für: ${protocol.id}`);
          
          // Vereinfachte Formular-Daten
          const formData = {
            counters: protocol.counters,
            sitzungsdatum: protocol.sitzungsdatum,
            ort: protocol.ort
          };
          
          // E-Mail-Versand versuchen
          const emailResult = this.sendProtocolEmail(protocol.id, protocol.pdfId, formData);
          
          if (emailResult.success) {
            successCount++;
            console.log(`✅ E-Mail für ${protocol.id} erfolgreich nachgesendet`);
          } else {
            stillFailedCount++;
            console.log(`❌ E-Mail für ${protocol.id} immer noch fehlgeschlagen`);
          }
          
          // Pause zwischen E-Mails (3 Sekunden)
          if (pendingProtocols.length > 1) {
            Utilities.sleep(3000);
          }
          
        } catch (error) {
          stillFailedCount++;
          console.error(`❌ Fehler bei ${protocol.id}:`, error);
        }
      }
      
      console.log(`✅ Wartende E-Mails verarbeitet: ${successCount} erfolgreich, ${stillFailedCount} immer noch wartend`);
      
      return {
        success: true,
        message: `${successCount} wartende E-Mails versendet, ${stillFailedCount} weiterhin wartend`,
        processedCount: successCount,
        stillPendingCount: stillFailedCount,
        totalFound: pendingProtocols.length
      };
      
    } catch (error) {
      console.error('❌ Fehler beim Versenden wartender E-Mails:', error);
      return {
        success: false,
        message: 'Fehler: ' + error.toString()
      };
    }
  },
  
  // E-Mail Text generieren
  generateEmailBody() {
    return `Liebe Mitglieder,

anbei das Protokoll unserer letzten Sitzung.

lg

Sepp`;
  }
};

// ==================== VEREINFACHTE QUOTA SERVICE ====================

// Einfache Quota-Prüfung
const QuotaService = {
  
  // Gmail-Quota Status prüfen
  checkGmailQuota() {
    try {
      // Test-E-Mail an sich selbst
      GmailApp.sendEmail(
        CONFIG.SENDER_EMAIL,
        'CDAS Quota Test - ' + new Date().toLocaleTimeString(),
        'Quota-Test. Diese E-Mail kann gelöscht werden.',
        {
          from: CONFIG.SENDER_EMAIL,
          name: CONFIG.SENDER_NAME
        }
      );
      
      return {
        success: true,
        status: 'available',
        message: 'Gmail-Quota verfügbar - E-Mails können versendet werden'
      };
      
    } catch (error) {
      return {
        success: false,
        status: 'exhausted',
        message: 'Gmail-Quota erschöpft - versuche es morgen wieder',
        error: error.toString()
      };
    }
  }
};

// Setup Service
const SetupService = {
  
  // Mitglieder-Sheet erstellen - ERWEITERT mit Beispiel für mehrere E-Mails
  createMembersSheet(ss) {
    let sheet = ss.getSheetByName('Mitglieder');
    if (!sheet) {
      sheet = ss.insertSheet('Mitglieder');
    } else {
      sheet.clear();
    }
    
    const headers = ['ID', 'Name', 'E-Mail (mehrere mit ; trennen)', 'Aktiv', 'Reihenfolge', 'Notizen'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ERWEITERTE Mitglieder-Daten mit Beispielen für mehrere E-Mails
    const members = [
      [2, 'Erwin Bertl', 'erwin@example.com', true, 1, ''],
      [5, 'Gerald Faic', 'gerald@example.com;gerald.business@company.com', true, 2, 'Hat 2 E-Mail-Adressen'],
      [8, 'Friedrich Fochler', 'friedrich@example.com', true, 3, ''],
      [3, 'Norbert Glaubacker', 'norbert@example.com;norbert.private@gmail.com', true, 4, 'Hat 2 E-Mail-Adressen'],
      [4, 'Gerhard John', 'gerhard@example.com', true, 5, ''],
      [9, 'Johannes Mader', 'johannes@example.com', true, 6, ''],
      [1, 'Josef Schnürer', 'josef.schnuerer@gmail.com', true, 7, 'Protokollführer'],
      [7, 'Christian Thür', 'christian@example.com;christian.work@office.com;christian.hobby@gmail.com', true, 8, 'Hat 3 E-Mail-Adressen'],
      [6, 'Otto Thür', 'otto@example.com', true, 9, '']
    ];
    
    sheet.getRange(2, 1, members.length, 6).setValues(members);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2c5282')
      .setFontColor('white')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    
    // NEUE: Spaltenbreite für E-Mail-Spalte vergrößern
    sheet.setColumnWidth(3, 300); // E-Mail-Spalte breiter machen
    
    console.log('✅ Mitglieder-Sheet erstellt - Mehrere E-Mail-Adressen unterstützt');
    console.log('📧 Beispiele für mehrere E-Mails: Gerald Faic (2), Norbert Glaubacker (2), Christian Thür (3)');
  },
  
  // Restliche Setup-Funktionen bleiben unverändert...
  setupSheetsStructure() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      this.createProtocolsSheet(ss);
      this.createMembersSheet(ss); // ERWEITERTE Version wird aufgerufen
      this.createAgendaItemsSheet(ss);
      this.createImagesSheet(ss);
      this.setupCounterProperties();
      
      console.log("✅ Sheets-Struktur erfolgreich erstellt!");
      return { success: true, message: "Sheets-Struktur mit Multi-E-Mail-Support erfolgreich erstellt!" };
    } catch (error) {
      console.error("❌ Fehler beim Erstellen der Sheets-Struktur:", error);
      return { success: false, message: error.toString() };
    }
  }
};

// Maintenance Service
const MaintenanceService = {
  
  // Alle Testdaten löschen
  clearAllTestData() {
    try {
      const ss = SheetService.getClubSpreadsheet();
      
      // Protokolle-Sheet bereinigen
      const protocolSheet = ss.getSheetByName('Protokolle');
      if (protocolSheet.getLastRow() > 1) {
        protocolSheet.deleteRows(2, protocolSheet.getLastRow() - 1);
        console.log('✅ Protokolle-Sheet bereinigt');
      }
      
      // Tagesordnungspunkte-Sheet bereinigen
      const agendaSheet = ss.getSheetByName('Tagesordnungspunkte');
      if (agendaSheet.getLastRow() > 1) {
        agendaSheet.deleteRows(2, agendaSheet.getLastRow() - 1);
        console.log('✅ Tagesordnungspunkte-Sheet bereinigt');
      }
      
      // Bilder-Sheet bereinigen
      const imagesSheet = ss.getSheetByName('Bilder');
      if (imagesSheet.getLastRow() > 1) {
        imagesSheet.deleteRows(2, imagesSheet.getLastRow() - 1);
        console.log('✅ Bilder-Sheet bereinigt');
      }
      
      // Counter zurücksetzen
      CounterService.resetCounters();
      
      console.log('✅ Alle Testdaten erfolgreich gelöscht!');
      console.log('🎯 System ist bereit für den produktiven Einsatz');
      
      return {
        success: true,
        message: 'Alle Testdaten erfolgreich gelöscht!'
      };
      
    } catch (error) {
      console.error('❌ Fehler beim Löschen der Testdaten:', error);
      return {
        success: false,
        message: 'Fehler beim Löschen der Testdaten: ' + error.toString()
      };
    }
  },
  
  // Bild-Referenzen reparieren
  repairExistingImageReferences() {
    console.log('🔧 Repariere bestehende Bild-Referenzen...');
    
    try {
      // Alle Tagesordnungspunkte laden
      const agendaData = SheetService.getSheetData('Tagesordnungspunkte');
      
      // Alle Bilder laden
      const imageData = SheetService.getSheetData('Bilder');
      
      let repairedCount = 0;
      
      // Mapping erstellen: agenda-X → echte ID
      const mapping = {};
      for (let i = 1; i < agendaData.length; i++) {
        const realId = agendaData[i][0]; // Spalte A: ID
        const nummer = agendaData[i][2];  // Spalte C: Nummer
        const frontendId = `agenda-${nummer}`;
        mapping[frontendId] = realId;
        console.log(`📋 Gefunden: ${frontendId} → ${realId}`);
      }
      
      // Bilder korrigieren
      const imageSheet = SheetService.getSheet('Bilder');
      for (let i = 1; i < imageData.length; i++) {
        const currentRef = imageData[i][1]; // Spalte B: Tagesordnungspunkt_ID
        
        if (mapping[currentRef]) {
          const correctId = mapping[currentRef];
          imageSheet.getRange(i + 1, 2).setValue(correctId);
          console.log(`🔄 Repariert: ${currentRef} → ${correctId}`);
          repairedCount++;
        }
      }
      
      console.log(`✅ ${repairedCount} Bild-Referenzen repariert!`);
      console.log('🎯 Alle Bilder sollten jetzt korrekt zugeordnet sein');
      
      return {
        success: true,
        repairedCount: repairedCount,
        message: `${repairedCount} Bild-Referenzen erfolgreich repariert`
      };
      
    } catch (error) {
      console.error('❌ Fehler bei der Reparatur:', error);
      return {
        success: false,
        message: 'Fehler bei der Reparatur: ' + error.toString()
      };
    }
  },
  
  // Status-Logs bereinigen
  cleanupOldStatusLogs(daysToKeep = 90) {
    try {
      const sheet = SheetService.getSheet('Protokolle');
      const data = sheet.getDataRange().getValues();
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
      
      let cleanedCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        const createdAt = new Date(data[i][1]);
        if (createdAt < cutoffDate) {
          // Nur Status-Logs bereinigen, nicht das gesamte Protokoll
          const currentNotes = data[i][14] || '';
          const lines = currentNotes.split('\n');
          const keepLines = lines.filter(line => {
            // Behalte wichtige Status-Meldungen
            return line.includes('✅ Protokoll vollständig') || 
                   line.includes('📧 E-Mail versendet') ||
                   line.includes('❌ FEHLER');
          });
          
          if (keepLines.length < lines.length) {
            const cleanedNotes = keepLines.join('\n');
            sheet.getRange(i + 1, 15).setValue(cleanedNotes);
            cleanedCount++;
          }
        }
      }
      
      console.log(`✅ ${cleanedCount} Status-Logs bereinigt (älter als ${daysToKeep} Tage)`);
      
      return {
        success: true,
        cleanedCount: cleanedCount,
        message: `${cleanedCount} Status-Logs bereinigt`
      };
      
    } catch (error) {
      console.error('Fehler beim Bereinigen:', error);
      return {
        success: false,
        message: 'Fehler beim Bereinigen: ' + error.toString()
      };
    }
  }
};

// Test Service
const TestService = {
    testMultiEmailSystem() {
    console.log('=== MULTI-E-MAIL SYSTEM TEST ===');
    
    try {
      console.log('📧 Schritt 1: Aktive Mitglieder laden...');
      const activeMembers = MemberService.getActiveMembers();
      
      console.log(`✅ ${activeMembers.length} aktive Mitglieder gefunden:`);
      activeMembers.forEach((member, index) => {
        console.log(`${index + 1}. ${member.name}:`);
        member.emailAddresses.forEach((email, emailIndex) => {
          const marker = emailIndex === 0 ? '(Haupt)' : '(Zusatz)';
          console.log(`   ${emailIndex + 1}. ${email} ${marker}`);
        });
      });
      
      console.log('\n📧 Schritt 2: E-Mail-Statistik...');
      const stats = MemberService.getEmailStatistics();
      console.log(`📊 Statistik:`);
      console.log(`   • Mitglieder gesamt: ${stats.totalMembers}`);
      console.log(`   • E-Mail-Adressen gesamt: ${stats.totalEmailAddresses}`);
      console.log(`   • Mit mehreren E-Mails: ${stats.membersWithMultipleEmails}`);
      console.log(`   • Mit einer E-Mail: ${stats.membersWithSingleEmail}`);
      
      console.log('\n📧 Schritt 3: Alle E-Mail-Adressen für Versand...');
      const allEmails = MemberService.getAllEmailAddresses();
      console.log(`📧 Versand-Liste (${allEmails.length} Adressen):`);
      allEmails.forEach((email, index) => {
        console.log(`${index + 1}. ${email}`);
      });
      
      console.log('\n📧 Schritt 4: E-Mail-Validierung testen...');
      const testEmails = [
        'valid@example.com',
        'also.valid+test@company.co.uk',
        'invalid-email',
        'missing@',
        '@missing.com',
        '',
        'spaces in@email.com'
      ];
      
      testEmails.forEach(email => {
        const isValid = MemberService.isValidEmail(email);
        console.log(`${isValid ? '✅' : '❌'} "${email}" → ${isValid ? 'Gültig' : 'Ungültig'}`);
      });
      
      console.log('\n✅ Multi-E-Mail System Test erfolgreich abgeschlossen!');
      
      return {
        success: true,
        message: 'Multi-E-Mail System Test erfolgreich!',
        statistics: stats,
        allEmailAddresses: allEmails,
        activeMembers: activeMembers
      };
      
    } catch (error) {
      console.error('❌ Multi-E-Mail Test fehlgeschlagen:', error);
      return {
        success: false,
        message: 'Multi-E-Mail Test fehlgeschlagen: ' + error.toString()
      };
    }
  },

  // Vollständiger Protokoll-zu-PDF Test
  testCompleteProtocolToPDF() {
    console.log('=== VOLLSTÄNDIGER PROTOKOLL-ZU-PDF TEST ===');
    
    try {
      // Test-Protokoll erstellen
      const testFormData = {
        counters: { ordentlich: 13, ausserordentlich: 6, fest: 2 },
        sitzungsdatum: '2025-07-26',
        ort: 'Clublokal',
        naechsterTermin: '2025-08-01',
        anwesende: ['1', '2', '3', '4', '5'],
        agendaItems: [
          {
            nummer: 1,
            titel: 'Test Tagesordnungspunkt für PDF',
            beschreibung: 'Dies ist eine Test-Beschreibung für den PDF-Export.',
            llmImprove: false,
            bilder: []
          }
        ]
      };
      
      console.log('📄 Schritt 1: Google Doc erstellen...');
      const docResult = DocumentService.createProtocolDocument('TEST_PDF_PROTOCOL', testFormData);
      
      if (!docResult.success) {
        console.error('❌ Doc-Erstellung fehlgeschlagen:', docResult.message);
        return { success: false, message: docResult.message };
      }
      
      console.log('✅ Google Doc erstellt:', docResult.docName);
      console.log('📄 Doc-URL:', docResult.docUrl);
      
      // PDF erstellen
      console.log('📄 Schritt 2: PDF erstellen...');
      const pdfResult = DocumentService.createProtocolPDF(docResult.docId);
      
      if (!pdfResult.success) {
        console.error('❌ PDF-Erstellung fehlgeschlagen:', pdfResult.message);
        return { success: false, message: pdfResult.message };
      }
      
      console.log('✅ PDF erstellt:', pdfResult.pdfName);
      console.log('📄 PDF-URL:', pdfResult.pdfUrl);
      console.log('⬇️ Download-URL:', pdfResult.downloadUrl);
      
      // Ordner-Struktur überprüfen
      console.log('📁 Schritt 3: Ordner-Struktur überprüfen...');
      
      const protocolsFolder = DriveService.getProtocolsFolder();
      const docFiles = protocolsFolder.getFilesByName(docResult.docName);
      const pdfFiles = protocolsFolder.getFilesByName(pdfResult.pdfName);
      
      console.log('📄 Google Doc im Ordner gefunden:', docFiles.hasNext());
      console.log('📄 PDF im Ordner gefunden:', pdfFiles.hasNext());
      
      console.log('✅ Vollständiger Test erfolgreich abgeschlossen!');
      console.log('🎯 Beide Dateien sollten jetzt im "Protokolle"-Ordner sein');
      
      return {
        success: true,
        message: 'Vollständiger Test erfolgreich abgeschlossen!',
        docResult: docResult,
        pdfResult: pdfResult
      };
      
    } catch (error) {
      console.error('❌ Test fehlgeschlagen:', error);
      return {
        success: false,
        message: 'Test fehlgeschlagen: ' + error.toString()
      };
    }
  },
  
  // Image-Mapping Test
  testCompleteImageMappingFlow() {
    console.log('=== VOLLSTÄNDIGER IMAGE-MAPPING TEST ===');
    
    try {
      // Test-Protokoll mit Bilder erstellen
      const testFormData = {
        counters: { ordentlich: 13, ausserordentlich: 6, fest: 2 },
        sitzungsdatum: '2025-07-26',
        ort: 'Clublokal',
        naechsterTermin: '2025-08-01',
        anwesende: ['1', '2', '3'],
        agendaItems: [
          {
            nummer: 1,
            titel: 'Test Agenda mit Bildern',
            beschreibung: 'Test-Beschreibung für Bild-Mapping',
            llmImprove: false,
            bilder: []
          },
          {
            nummer: 2,
            titel: 'Zweiter Test-Punkt',
            beschreibung: 'Weitere Test-Beschreibung',
            llmImprove: false,
            bilder: []
          }
        ]
      };
      
      console.log('📄 Schritt 1: Draft speichern (erstellt ID-Mapping)...');
      const draftResult = ProtocolService.saveDraft(testFormData);
      
      if (!draftResult.success) {
        console.error('❌ Draft-Erstellung fehlgeschlagen:', draftResult.message);
        return { success: false, message: draftResult.message };
      }
      
      console.log('✅ Draft gespeichert mit ID-Mapping:');
      console.log('Mapping:', draftResult.agendaIdMapping);
      
      // Agenda-Items in DB überprüfen
      console.log('📄 Schritt 2: Überprüfe Agenda-Items in DB...');
      const agendaData = SheetService.getSheetData('Tagesordnungspunkte');
      
      for (let i = 1; i < agendaData.length; i++) {
        if (agendaData[i][1] === draftResult.protocolId) {
          console.log(`✅ Agenda gefunden: ID=${agendaData[i][0]}, Nummer=${agendaData[i][2]}, Titel=${agendaData[i][3]}`);
        }
      }
      
      // Protokoll-Dokument erstellen und Bilder-Suche testen
      console.log('📄 Schritt 3: Protokoll-Dokument erstellen...');
      const docResult = DocumentService.createProtocolDocument(draftResult.protocolId, testFormData);
      
      if (docResult.success) {
        console.log('✅ Protokoll-Dokument erfolgreich erstellt');
        console.log('📄 Doc-URL:', docResult.docUrl);
      } else {
        console.error('❌ Protokoll-Dokument Erstellung fehlgeschlagen:', docResult.message);
      }
      
      console.log('✅ Vollständiger Image-Mapping Test abgeschlossen!');
      
      return {
        success: true,
        message: 'Image-Mapping Test erfolgreich abgeschlossen!',
        draftResult: draftResult,
        docResult: docResult
      };
      
    } catch (error) {
      console.error('❌ Test fehlgeschlagen:', error);
      return {
        success: false,
        message: 'Test fehlgeschlagen: ' + error.toString()
      };
    }
  },
  
  // E-Mail Test
  testEmailSending() {
    console.log('=== E-MAIL VERSAND TEST ===');
    
    try {
      // Test-Daten erstellen
      const testFormData = {
        counters: { ordentlich: 13, ausserordentlich: 6, fest: 2 },
        sitzungsdatum: '2025-07-26',
        ort: 'Clublokal',
        naechsterTermin: '2025-08-01',
        anwesende: ['1', '2', '3']
      };
      
      // Aktive Mitglieder anzeigen
      const activeMembers = MemberService.getActiveMembers();
      console.log('📧 Empfänger-Liste:');
      activeMembers.forEach((member, index) => {
        console.log(`${index + 1}. ${member.name} - ${member.email}`);
      });
      
      // Protokoll-Name testen
      const protocolName = UtilService.generateDocumentName(testFormData);
      console.log('📧 Test-Betreff:', protocolName);
      
      // E-Mail-Body testen
      const emailBody = EmailService.generateEmailBody();
      console.log('📧 E-Mail-Text:', emailBody);
      
      console.log('✅ E-Mail-Versand Test erfolgreich - bereit für echten Versand');
      
      return {
        success: true,
        message: 'E-Mail Test erfolgreich abgeschlossen!',
        recipients: activeMembers,
        subject: protocolName,
        body: emailBody
      };
      
    } catch (error) {
      console.error('❌ E-Mail Test fehlgeschlagen:', error);
      return {
        success: false,
        message: 'E-Mail Test fehlgeschlagen: ' + error.toString()
      };
    }
  },
  
  // Gmail-Berechtigung testen
  testGmailPermission() {
    try {
      console.log('Teste Gmail-Zugriff...');
      const labels = GmailApp.getUserLabels();
      console.log('✅ Gmail-Berechtigung aktiviert, Labels gefunden:', labels.length);
      return {
        success: true,
        message: 'Gmail-Berechtigung erfolgreich getestet',
        labelsFound: labels.length
      };
    } catch (error) {
      console.error('❌ Gmail-Berechtigung fehlt:', error);
      return {
        success: false,
        message: 'Gmail-Berechtigung fehlt: ' + error.toString()
      };
    }
  }
};
