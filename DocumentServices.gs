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
    this.createCompactSessionInfo(body, formData);
    this.createParticipantsInfo(body, formData);
    this.createAgendaSection(body, formData, protocolId);
    this.applyDocumentStyling(doc);
  },
  
  // Header erstellen
  createDocumentHeader(body, formData) {
    body.appendParagraph('');
    
    const clubName = body.appendParagraph('CLUB DER ALTEN SÄCKE');
    clubName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    clubName.editAsText().setFontSize(30).setBold(true).setForegroundColor('#1a365d');
    
    const protocolTitle = body.appendParagraph('Sitzungsprotokoll');
    protocolTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    protocolTitle.editAsText().setFontSize(22).setForegroundColor('#4a5568');
    
    const separator1 = body.appendParagraph('══════════════════════════════════════════════════');
    separator1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    separator1.editAsText().setForegroundColor('#2c5282').setBold(true).setFontSize(12);
    
    const separator2 = body.appendParagraph('═══════════════════════');
    separator2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    separator2.editAsText().setForegroundColor('#2c5282').setBold(true).setFontSize(10);
    
    body.appendParagraph('').editAsText().setFontSize(11);
    body.appendParagraph('').editAsText().setFontSize(11);
  },

  createCompactSessionInfo(body, data) {
    const totalNumber = data.counters.ordentlich + data.counters.ausserordentlich + data.counters.fest;
    
    // Session-Nummer als einfacher Text mit leichtem Hintergrund
    const sessionHeader1 = body.appendParagraph(`Sitzung Nr. ${totalNumber}`);
    sessionHeader1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sessionHeader1.editAsText()
      .setFontSize(16)
      .setBold(true)
      .setForegroundColor('#2c5282');
    sessionHeader1.setBackgroundColor('#f0f7ff');
    sessionHeader1.setSpacingAfter(0); // Kein Abstand zwischen den Zeilen

    // Zweite Zeile
    const sessionHeader2 = body.appendParagraph(`(${data.counters.ordentlich}. ordentliche, ${data.counters.ausserordentlich}. außerordentliche, ${data.counters.fest}. Festsitzung)`);
    sessionHeader2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    sessionHeader2.editAsText()
      .setFontSize(14)  // Etwas kleiner für die Details
      .setBold(false)   // Nicht fett für weniger Gewicht
      .setForegroundColor('#2c5282');
    sessionHeader2.setBackgroundColor('#f0f7ff');
    sessionHeader2.setSpacingAfter(15);
    
    // Info als einfache Tabelle ohne viel Styling
    const infoTable = body.appendTable();
    infoTable.setBorderWidth(1);
    infoTable.setBorderColor('#eaeaea');
    
    // Erste Zeile: Datum und Ort
    const row1 = infoTable.appendTableRow();
    
    const dateCell = row1.appendTableCell();
    const datePara = dateCell.appendParagraph('Datum: ' + formatDateDE(data.sitzungsdatum));
    datePara.editAsText().setFontSize(16).setBackgroundColor('#ffffff').setFontFamily('Arial');
    dateCell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(15);
    
    const ortCell = row1.appendTableCell();
    const ortPara = ortCell.appendParagraph('Ort: ' + data.ort);
    ortPara.editAsText().setFontSize(16).setBackgroundColor('#ffffff').setFontFamily('Arial');
    ortCell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(15);
    
    // Zweite Zeile: Protokollführer und nächster Termin
    const row2 = infoTable.appendTableRow();
    
    const protocolCell = row2.appendTableCell();
    const protocolPara = protocolCell.appendParagraph('Protokollführer: Josef Schnürer');
    protocolPara.editAsText().setFontSize(16).setBackgroundColor('#f9fafb').setFontFamily('Arial');
    protocolCell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(15);
    protocolCell.setBackgroundColor('#f9fafb');
    
    const nextCell = row2.appendTableCell();
    const nextPara = nextCell.appendParagraph('Nächster Termin: ' + formatDateDE(data.naechsterTermin));
    nextPara.editAsText().setFontSize(16).setBackgroundColor('#f9fafb').setFontFamily('Arial');
    nextCell.setPaddingTop(8).setPaddingBottom(8).setPaddingLeft(15).setFontFamily('Arial');
    nextCell.setBackgroundColor('#f9fafb');
    
    body.appendParagraph('').editAsText().setFontSize(11);
  },
  
  // Teilnehmer-Informationen
  createParticipantsInfo(body, formData) {
    const allMembers = MemberService.getMembersForFrontend();
    const anwesendeIds = formData.anwesende.map(id => parseInt(id));
    
    const anwesende = allMembers.filter(m => anwesendeIds.includes(m.id)).map(m => m.name);
    const abwesende = allMembers.filter(m => !anwesendeIds.includes(m.id)).map(m => m.name);
    
    const participantsTitle = body.appendParagraph('TEILNEHMER');
    participantsTitle.editAsText().setFontSize(16).setFontFamily('Arial').setBold(true).setForegroundColor('#2c5282');
    participantsTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    participantsTitle.setBackgroundColor('#eaeaea');
    
    body.appendParagraph('').editAsText().setFontSize(11);
    
    const anwesendeTitle = body.appendParagraph('✓ Anwesend:');
    anwesendeTitle.editAsText().setFontSize(16).setBold(true).setForegroundColor('#2c5282');
    
    const anwesendeText = anwesende.join(', ');
    const anwesendePara = body.appendParagraph(anwesendeText);
    anwesendePara.editAsText().setFontSize(16);
    anwesendePara.setBackgroundColor('#f0fff4');
    anwesendePara.setIndentFirstLine(20);
    
    body.appendParagraph('').editAsText().setFontSize(11);
    
    if (abwesende.length > 0) {
      const abwesendeTitle = body.appendParagraph('✗ Abwesend:');
      abwesendeTitle.editAsText().setFontSize(16).setBold(true).setForegroundColor('#2c5282');
      
      const abwesendeText = abwesende.join(', ');
      const abwesendePara = body.appendParagraph(abwesendeText);
      abwesendePara.editAsText().setFontSize(16);
      abwesendePara.setBackgroundColor('#fef2f2');
      abwesendePara.setIndentFirstLine(20);
      
      body.appendParagraph('').editAsText().setFontSize(11);
    }
  },
  
  // Tagesordnungssektion
  createAgendaSection(body, formData, protocolId) {
    const agendaTitle = body.appendParagraph('TAGESORDNUNG');
    agendaTitle.editAsText().setFontSize(16).setBold(true).setFontFamily('Arial').setForegroundColor('#2c5282');
    agendaTitle.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    agendaTitle.setBackgroundColor('#eaeaea');
    
    const separator = body.appendParagraph('════════════════════════════════════════════════════════');
    separator.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    separator.editAsText().setForegroundColor('#2c5282').setFontSize(10);
    
    body.appendParagraph('').editAsText().setFontSize(11);
    
    formData.agendaItems.forEach((item, index) => {
      this.createAgendaItem(body, item, index + 1, protocolId);
    });
  },
  
  // Einzelnen Tagesordnungspunkt erstellen
  createAgendaItem(body, item, number, protocolId) {
    const titleText = `${number}. ${item.titel}`;
    const titlePara = body.appendParagraph(titleText);
    titlePara.editAsText().setFontSize(16).setFontFamily('Arial').setBold(true).setForegroundColor('#1a365d');
    titlePara.setBackgroundColor('#e6fffa');
    titlePara.setIndentFirstLine(10);
    
    if (item.beschreibung && item.beschreibung.trim()) {
      const descPara = body.appendParagraph(item.beschreibung);
      descPara.editAsText().setFontSize(16).setFontFamily('Arial');
      descPara.setBackgroundColor('#ffffff');
      descPara.setIndentFirstLine(30);
      descPara.setIndentStart(15);
    } else {
      const emptyPara = body.appendParagraph('[Keine Beschreibung]');
      emptyPara.editAsText().setFontSize(3).setItalic(true).setForegroundColor('#999999');
      emptyPara.setIndentFirstLine(30);
      emptyPara.setIndentStart(15);
    }
    
    // Bilder einfügen
    const realAgendaId = ProtocolService.findRealAgendaIdByNumber(protocolId, number);
    if (realAgendaId) {
      console.log(`🔍 Suche Bilder für echte AgendaID: ${realAgendaId} (Nr. ${number})`);
      this.insertAgendaImages(body, realAgendaId);
    }
    
    body.appendParagraph('').editAsText().setFontSize(11);
    
    const itemSeparator = body.appendParagraph('───────────────────────────────────────────────────────');
    itemSeparator.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    itemSeparator.editAsText().setForegroundColor('#cbd5e0').setFontSize(8);
    
    body.appendParagraph('').editAsText().setFontSize(11);
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
            
            body.appendParagraph('').editAsText().setFontSize(11);
            
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

