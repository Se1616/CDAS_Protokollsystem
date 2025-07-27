// ==================== DOCUMENT SERVICE ====================

const DocumentService = {
  
  // Google Doc erstellen
  createProtocolDocument(protocolId, formData) {
    try {
      console.log('Erstelle Google Doc für Protokoll:', protocolId);
      
      const protocolsFolder = DriveService.getProtocolsFolder();
      const docName = UtilService.generateDocumentName(formData);
      
      // Neues Google Doc erstellen
      const doc = DocumentApp.create(docName);
      const docId = doc.getId();
      
      // Dokument in Ordner verschieben
      const file = DriveApp.getFileById(docId);
      protocolsFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
      
      // Protokoll-Inhalt generieren
      this.generateProtocolContent(doc, formData, protocolId);
      
      doc.saveAndClose();
      
      console.log('Google Doc erfolgreich erstellt:', docName, 'ID:', docId);
      
      return {
        success: true,
        docId: docId,
        docName: docName,
        docUrl: `https://docs.google.com/document/d/${docId}/edit`,
        message: 'Google Doc erfolgreich erstellt'
      };
      
    } catch (error) {
      console.error('Fehler beim Erstellen des Google Docs:', error);
      return {
        success: false,
        message: 'Fehler beim Erstellen des Docs: ' + error.toString()
      };
    }
  },
  
  // Protokoll-Inhalt generieren
  generateProtocolContent(doc, formData, protocolId) {
    const body = doc.getBody();
    body.clear();
    
    this.createDocumentHeader(body, formData);
    this.createSessionInfo(body, formData);
    this.createParticipantsInfo(body, formData);
    this.createAgendaSection(body, formData, protocolId);
    this.applyDocumentStyling(doc);
  },
  
  // Header erstellen
  createDocumentHeader(body, formData) {
    body.appendParagraph('');
    
    const clubName = body.appendParagraph('CLUB DER ALTEN SÄCKE');
    clubName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    clubName.editAsText().setFontSize(28).setBold(true).setForegroundColor('#1a365d');
    
    const protocolTitle = body.appendParagraph('Sitzungsprotokoll');
    protocolTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    protocolTitle.editAsText().setFontSize(16).setForegroundColor('#4a5568');
    
    const separator1 = body.appendParagraph('══════════════════════════════════════════════════');
    separator1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    separator1.editAsText().setForegroundColor('#2c5282').setBold(true).setFontSize(12);
    
    const separator2 = body.appendParagraph('═══════════════════════');
    separator2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    separator2.editAsText().setForegroundColor('#2c5282').setBold(true).setFontSize(10);
    
    body.appendParagraph('');
    body.appendParagraph('');
  },
  
  // Sitzungsinformationen
  createSessionInfo(body, formData) {
    const totalNumber = formData.counters.ordentlich + formData.counters.ausserordentlich + formData.counters.fest;
    
    const sessionNumber = `Sitzung Nr. ${totalNumber} (${formData.counters.ordentlich}. ordentliche, ${formData.counters.ausserordentlich}. außerordentliche, ${formData.counters.fest}. Festsitzung)`;
    const sessionPara = body.appendParagraph(sessionNumber);
    sessionPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sessionPara.editAsText().setFontSize(16).setBold(true).setForegroundColor('#1a365d');
    sessionPara.setBackgroundColor('#e2e8f0');
    
    body.appendParagraph('');
    
    const sitzungsdatum = UtilService.formatDateDE(formData.sitzungsdatum);
    const nextTermin = formData.naechsterTermin ? UtilService.formatDateDE(formData.naechsterTermin) : 'TBD';
    
    const metaInfo1 = body.appendParagraph(`Datum: ${sitzungsdatum}                                        Ort: ${formData.ort}`);
    metaInfo1.editAsText().setFontSize(12).setBold(true);
    metaInfo1.setBackgroundColor('#f7fafc');
    
    const metaInfo2 = body.appendParagraph(`Protokollführer: Josef Schnürer              Nächster Termin: ${nextTermin}`);
    metaInfo2.editAsText().setFontSize(12).setBold(true);
    metaInfo2.setBackgroundColor('#f7fafc');
    
    body.appendParagraph('');
  },
  
  // Teilnehmer-Informationen
  createParticipantsInfo(body, formData) {
    const allMembers = MemberService.getMembersForFrontend();
    const anwesendeIds = formData.anwesende.map(id => parseInt(id));
    
    const anwesende = allMembers.filter(m => anwesendeIds.includes(m.id)).map(m => m.name);
    const abwesende = allMembers.filter(m => !anwesendeIds.includes(m.id)).map(m => m.name);
    
    const participantsTitle = body.appendParagraph('TEILNEHMER');
    participantsTitle.editAsText().setFontSize(14).setBold(true).setForegroundColor('#2c5282');
    participantsTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    participantsTitle.setBackgroundColor('#e2e8f0');
    
    body.appendParagraph('');
    
    const anwesendeTitle = body.appendParagraph('✓ Anwesend:');
    anwesendeTitle.editAsText().setFontSize(13).setBold(true).setForegroundColor('#2c5282');
    
    const anwesendeText = anwesende.join(', ');
    const anwesendePara = body.appendParagraph(anwesendeText);
    anwesendePara.editAsText().setFontSize(11);
    anwesendePara.setBackgroundColor('#f0fff4');
    anwesendePara.setIndentFirstLine(20);
    
    body.appendParagraph('');
    
    if (abwesende.length > 0) {
      const abwesendeTitle = body.appendParagraph('✗ Abwesend:');
      abwesendeTitle.editAsText().setFontSize(13).setBold(true).setForegroundColor('#2c5282');
      
      const abwesendeText = abwesende.join(', ');
      const abwesendePara = body.appendParagraph(abwesendeText);
      abwesendePara.editAsText().setFontSize(11);
      abwesendePara.setBackgroundColor('#fef2f2');
      abwesendePara.setIndentFirstLine(20);
      
      body.appendParagraph('');
    }
  },
  
  // Tagesordnungssektion
  createAgendaSection(body, formData, protocolId) {
    const agendaTitle = body.appendParagraph('TAGESORDNUNG');
    agendaTitle.editAsText().setFontSize(14).setBold(true).setForegroundColor('#2c5282');
    agendaTitle.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    agendaTitle.setBackgroundColor('#e2e8f0');
    
    const separator = body.appendParagraph('════════════════════════════════════════════════════');
    separator.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    separator.editAsText().setForegroundColor('#2c5282').setFontSize(10);
    
    body.appendParagraph('');
    
    formData.agendaItems.forEach((item, index) => {
      this.createAgendaItem(body, item, index + 1, protocolId);
    });
  },
  
  // Einzelnen Tagesordnungspunkt erstellen
  createAgendaItem(body, item, number, protocolId) {
    const titleText = `${number}. ${item.titel}`;
    const titlePara = body.appendParagraph(titleText);
    titlePara.editAsText().setFontSize(13).setBold(true).setForegroundColor('#1a365d');
    titlePara.setBackgroundColor('#e6fffa');
    titlePara.setIndentFirstLine(10);
    
    if (item.beschreibung && item.beschreibung.trim()) {
      const descPara = body.appendParagraph(item.beschreibung);
      descPara.editAsText().setFontSize(11);
      descPara.setBackgroundColor('#f0f9ff');
      descPara.setIndentFirstLine(30);
      descPara.setIndentStart(15);
    } else {
      const emptyPara = body.appendParagraph('[Keine Beschreibung]');
      emptyPara.editAsText().setFontSize(11).setItalic(true).setForegroundColor('#999999');
      emptyPara.setIndentFirstLine(30);
      emptyPara.setIndentStart(15);
    }
    
    // Bilder einfügen
    const realAgendaId = ProtocolService.findRealAgendaIdByNumber(protocolId, number);
    if (realAgendaId) {
      console.log(`🔍 Suche Bilder für echte AgendaID: ${realAgendaId} (Nr. ${number})`);
      this.insertAgendaImages(body, realAgendaId);
    }
    
    body.appendParagraph('');
    
    const itemSeparator = body.appendParagraph('───────────────────────────────────────────────────────');
    itemSeparator.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    itemSeparator.editAsText().setForegroundColor('#cbd5e0').setFontSize(8);
    
    body.appendParagraph('');
  },
  
  // Bilder in Tagesordnungspunkt einfügen
  insertAgendaImages(body, realAgendaId) {
    try {
      console.log(`🔍 Suche Bilder für echte AgendaID: "${realAgendaId}"`);
      const images = ImageService.getImagesForAgendaItem(realAgendaId);
      
      console.log(`📊 Gefundene Bilder für ${realAgendaId}:`, images.length);
      
      if (images.length > 0) {
        const imageTitle = body.appendParagraph(`📷 Bilder (${images.length}):`);
        imageTitle.editAsText().setFontSize(10).setBold(true).setForegroundColor('#4a5568');
        imageTitle.setIndentFirstLine(30);
        
        images.forEach((image, index) => {
          try {
            console.log(`🖼️ Verarbeite Bild ${index + 1}:`, image.fileName, 'DriveID:', image.driveId);
            
            const file = DriveApp.getFileById(image.driveId);
            const blob = file.getBlob();
            
            const imagePara = body.appendParagraph('');
            const insertedImage = imagePara.appendInlineImage(blob);
            
            // Bild-Größe anpassen
            const originalWidth = insertedImage.getWidth();
            const originalHeight = insertedImage.getHeight();
            
            if (originalWidth > CONFIG.IMAGE_MAX_WIDTH) {
              const scaleFactor = CONFIG.IMAGE_MAX_WIDTH / originalWidth;
              insertedImage.setWidth(CONFIG.IMAGE_MAX_WIDTH);
              insertedImage.setHeight(originalHeight * scaleFactor);
            }
            
            imagePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            imagePara.setIndentStart(20);
            
            const caption = body.appendParagraph(`Bild ${index + 1}: ${image.fileName}`);
            caption.editAsText().setFontSize(9).setItalic(true).setForegroundColor('#666666');
            caption.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            caption.setIndentStart(20);
            
            body.appendParagraph('');
            
            console.log(`✅ Bild ${index + 1} erfolgreich eingefügt:`, image.fileName);
            
          } catch (imageError) {
            console.error(`❌ Fehler beim Einfügen von Bild ${image.fileName}:`, imageError);
            
            const errorPara = body.appendParagraph(`[Bild nicht verfügbar: ${image.fileName}]`);
            errorPara.editAsText().setItalic(true).setForegroundColor('#999999').setFontSize(10);
            errorPara.setIndentFirstLine(30);
          }
        });
      } else {
        console.log(`ℹ️ Keine Bilder gefunden für echte AgendaID: "${realAgendaId}"`);
      }
      
    } catch (error) {
      console.error('❌ Fehler beim Laden der Bilder für', realAgendaId, ':', error);
    }
  },
  
  // Dokument-Styling anwenden
  applyDocumentStyling(doc) {
    const body = doc.getBody();
    
    const style = {};
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    style[DocumentApp.Attribute.FONT_SIZE] = 11;
    style[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    body.setAttributes(style);
    
    body.setMarginTop(50);
    body.setMarginBottom(50);
    body.setMarginLeft(70);
    body.setMarginRight(50);
  },
  
  // PDF erstellen
  createProtocolPDF(docId) {
    try {
      console.log('Erstelle PDF aus Google Doc:', docId);
      
      const doc = DocumentApp.openById(docId);
      const docName = doc.getName();
      
      if (!docName) {
        throw new Error('Document name is null or empty');
      }
      
      console.log('Document name:', docName);
      
      const protocolsFolder = DriveService.getProtocolsFolder();
      const docFile = DriveApp.getFileById(docId);
      
      const pdfBlob = docFile.getAs('application/pdf');
      const pdfName = docName + '.pdf';
      pdfBlob.setName(pdfName);
      
      console.log('PDF blob created with name:', pdfName);
      
      const pdfFile = protocolsFolder.createFile(pdfBlob);
      const pdfId = pdfFile.getId();
      
      console.log('PDF file created with ID:', pdfId);
      
      if (pdfFile.getName() !== pdfName) {
        pdfFile.setName(pdfName);
        console.log('PDF file renamed to:', pdfName);
      }
      
      console.log('PDF erfolgreich erstellt:', pdfName, 'ID:', pdfId);
      
      return {
        success: true,
        pdfId: pdfId,
        pdfName: pdfName,
        pdfUrl: `https://drive.google.com/file/d/${pdfId}/view`,
        downloadUrl: `https://drive.google.com/uc?export=download&id=${pdfId}`,
        message: 'PDF erfolgreich erstellt'
      };
      
    } catch (error) {
      console.error('Fehler beim PDF-Export:', error);
      return {
        success: false,
        message: 'Fehler beim PDF-Export: ' + error.toString()
      };
    }
  }
};

