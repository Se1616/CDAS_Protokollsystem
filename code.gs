// ==================== CDAS PROTOKOLL SYSTEM - HAUPTDATEI ====================
// Version: 2.1 - Mit Auto-Resume Feature
// Letzte Änderung: 27.07.2025

// ==================== WEB APP BASIS ====================

// Web App bereitstellen
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('CDAS Protokoll-Erfassung')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// HTML-Dateien einbinden
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== KONFIGURATION ====================

const CONFIG = {
  CLUB_FOLDER_NAME: 'Club der alten Säcke',
  PROTOCOLS_FOLDER: 'Protokolle',
  IMAGES_FOLDER: 'Protokoll-Bilder',
  SENDER_EMAIL: 'josef.schnuerer@gmail.com',
  SENDER_NAME: 'Josef Schnürer - CDAS Protokollführer',
  MAX_IMAGE_SIZE: 5 * 1024 * 1024, // 5MB
  IMAGE_MAX_WIDTH: 350
};

// ==================== FRONTEND API ====================

// Aktuelle Zähler für Frontend laden
function getCurrentCountersForFrontend() {
  return CounterService.getCurrentCountersForFrontend();
}

// Mitglieder für Frontend laden
function getMembersForFrontend() {
  return MemberService.getMembersForFrontend();
}

function getLastDraftProtocol() {
  try {
    const result = AutoResumeService.getLastDraftProtocol();
    
    // Date-Objekte für Frontend-Übertragung konvertieren
    if (result && result.success && result.protocolData && result.protocolData.createdAt) {
      result.protocolData.createdAt = result.protocolData.createdAt.toISOString();
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Fehler bei getLastDraftProtocol:', error);
    return {
      success: false,
      message: 'Fehler beim Laden des Draft-Protokolls'
    };
  }
}

// AUTO-RESUME: Letztes Draft-Protokoll laden
function aaaagetLastDraftProtocol() {
  try {
    const result = AutoResumeService.getLastDraftProtocol();
    
    // Date-Objekte für Frontend-Übertragung konvertieren
    if (result && result.success && result.protocolData && result.protocolData.createdAt) {
      result.protocolData.createdAt = result.protocolData.createdAt.toISOString();
    }
    
    return result;
    
  } catch (error) {
    console.error('❌ Fehler bei getLastDraftProtocol:', error);
    return {
      success: false,
      message: 'Fehler beim Laden des Draft-Protokolls'
    };
  }
}

// AUTO-RESUME: Letztes Draft-Protokoll laden
function zzzzzgetLastDraftProtocol() {
  return AutoResumeService.getLastDraftProtocol();
}
// TEMPORÄRER DEBUG - Funktion ersetzen
function zzzzzzzzzzzgetLastDraftProtocol() {
  console.log('=== FRONTEND CALL getLastDraftProtocol ===');
  
  try {
    console.log('Prüfe AutoResumeService:', typeof AutoResumeService);
    
    if (typeof AutoResumeService === 'undefined') {
      console.log('❌ AutoResumeService NICHT VERFÜGBAR bei Frontend-Call!');
      return {
        success: false,
        message: 'AutoResumeService nicht verfügbar bei Frontend-Call',
        debug: 'Service-Loading Problem'
      };
    }
    
    console.log('✅ AutoResumeService verfügbar, rufe auf...');
    const result = AutoResumeService.getLastDraftProtocol();
    console.log('Ergebnis von AutoResumeService:', result);
    
    if (!result) {
      console.log('❌ AutoResumeService gab null/undefined zurück');
      return {
        success: false,
        message: 'AutoResumeService gab null zurück',
        debug: 'Service-Result Problem'
      };
    }
    
    // DATUM-FIX: Date-Objekt zu String konvertieren
    if (result.protocolData && result.protocolData.createdAt) {
      result.protocolData.createdAt = result.protocolData.createdAt.toISOString();
    }
    
    console.log('✅ Sende an Frontend (Date-fixed):', result);
    return result;
    
  } catch (error) {
    console.error('❌ Exception in getLastDraftProtocol:', error.toString());
    console.error('Stack:', error.stack);
    return {
      success: false,
      message: 'Exception: ' + error.toString(),
      debug: 'Try-Catch Exception',
      stack: error.stack
    };
  }
}

// Entwurf speichern
function saveDraft(formData) {
  return ProtocolService.saveDraft(formData);
}

// Finales Protokoll erstellen
function createFinalProtocol(formData) {
  return ProtocolService.createFinalProtocol(formData);
}

// Bilder hochladen
function uploadMultipleImages(imagesData, agendaItemId) {
  return ImageService.uploadMultipleImages(imagesData, agendaItemId);
}

// Bild löschen
function deleteImage(imageId) {
  return ImageService.deleteImage(imageId);
}

function testMultiEmailSystem() {
  return TestService.testMultiEmailSystem();
}


// Vereinfachte Funktionen für Code.gs
function sendPendingEmails() {
  return EmailService.sendPendingEmails();
}

function checkGmailQuota() {
  return QuotaService.checkGmailQuota();
}

// NEUE Funktion: Alle wartenden Protokolle anzeigen
function showPendingEmails() {
  try {
    const data = SheetService.getSheetData('Protokolle');
    const pending = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === 'PDF erstellt - E-Mail wartend') {
        pending.push({
          id: data[i][0],
          datum: new Date(data[i][3]).toLocaleDateString('de-DE'),
          erstellt: new Date(data[i][1]).toLocaleDateString('de-DE')
        });
      }
    }
    
    if (pending.length === 0) {
      console.log('✅ Keine wartenden E-Mails');
      return { success: true, message: 'Keine wartenden E-Mails', pendingCount: 0 };
    }
    
    console.log(`📧 ${pending.length} wartende E-Mails:`);
    pending.forEach((p, index) => {
      console.log(`${index + 1}. ${p.id} - Sitzung: ${p.datum} - Erstellt: ${p.erstellt}`);
    });
    
    return {
      success: true,
      message: `${pending.length} wartende E-Mails gefunden`,
      pendingCount: pending.length,
      pendingProtocols: pending
    };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
// ==================== SETUP & WARTUNG ====================

// System einrichten
function setupSheetsStructure() {
  return SetupService.setupSheetsStructure();
}

// Testdaten löschen
function clearAllTestData() {
  return MaintenanceService.clearAllTestData();
}

// Status aller Protokolle anzeigen
function showAllProtocolStatuses() {
  return StatusService.showAllProtocolStatuses();
}

// Bild-Referenzen reparieren
function repairExistingImageReferences() {
  return MaintenanceService.repairExistingImageReferences();
}

// ==================== TEST FUNKTIONEN ====================

// Vollständiger System-Test
function testCompleteProtocolToPDF() {
  return TestService.testCompleteProtocolToPDF();
}

// Image-Mapping Test
function testCompleteImageMappingFlow() {
  return TestService.testCompleteImageMappingFlow();
}

// E-Mail Test
function testEmailSending() {
  return TestService.testEmailSending();
}

function debugServices() {
  console.log('=== SERVICE AVAILABILITY TEST ===');
  
  try {
    console.log('CounterService:', typeof CounterService);
    console.log('MemberService:', typeof MemberService);  
    console.log('AutoResumeService:', typeof AutoResumeService);
    console.log('ProtocolService:', typeof ProtocolService);
    console.log('DriveService:', typeof DriveService);
    console.log('SheetService:', typeof SheetService);
    console.log('ImageService:', typeof ImageService);
    console.log('DocumentService:', typeof DocumentService);
    console.log('EmailService:', typeof EmailService);
  } catch (error) {
    console.error('Service Error:', error.toString());
  }
  
  // Test eine einfache Service-Funktion
  try {
    const result = getCurrentCountersForFrontend();
    console.log('Counter Test Result:', result);
  } catch (error) {
    console.error('Counter Test Error:', error.toString());
  }
}

function debugServicesAfterFix() {
  console.log('=== SERVICE TEST NACH FIX ===');
  
  try {
    // Test AutoResumeService
    console.log('AutoResumeService:', typeof AutoResumeService);
    console.log('CounterService:', typeof CounterService);
    console.log('MemberService:', typeof MemberService);
    console.log('ProtocolService:', typeof ProtocolService);
    
    // Test Dependencies
    console.log('StatusService:', typeof StatusService);
    console.log('ImageService:', typeof ImageService);
    console.log('SheetService:', typeof SheetService);
    
    // Test eine Funktion
    const result = AutoResumeService.getLastDraftProtocol();
    console.log('AutoResume Test:', result ? 'Success' : 'Failed');
    
  } catch (error) {
    console.error('Service Test Error:', error.toString());
  }
}

function manualFrontendDebugTest() {
  console.log('=== FRONTEND CALL DEBUG ===');
  
  try {
    console.log('Teste getLastDraftProtocol()...');
    const result = getLastDraftProtocol();
    console.log('getLastDraftProtocol() Ergebnis:', result);
    console.log('Ergebnis-Typ:', typeof result);
    
    if (result) {
      console.log('Success-Flag:', result.success);
      console.log('Message:', result.message);
      if (result.protocolData) {
        console.log('ProtocolData ID:', result.protocolData.id);
      }
    }
    
    return result;
    
  } catch (error) {
    console.error('FEHLER in manualFrontendDebugTest:', error.toString());
    return {
      success: false,
      message: 'Debug-Fehler: ' + error.toString()
    };
  }
}

function frontendTestPing() {
  console.log('🔔 FRONTEND TEST PING aufgerufen!');
  return {
    success: true,
    message: 'Backend-Verbindung funktioniert!',
    version: 'Debug Version 2.0',
    timestamp: new Date().toISOString(),
    deployment: 'Neues Deployment aktiv'
  };
}
