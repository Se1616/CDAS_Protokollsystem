// ==================== BUSINESS SERVICES MIT AUTO-RESUME ====================

// Counter Service
const CounterService = {
  
  // Aktuelle Zähler für Frontend
  getCurrentCountersForFrontend() {
    const properties = PropertiesService.getScriptProperties();
    const counters = {
      ordentlich: parseInt(properties.getProperty('counter_ordentlich') || '12'),
      ausserordentlich: parseInt(properties.getProperty('counter_ausserordentlich') || '6'),
      fest: parseInt(properties.getProperty('counter_fest') || '2')
    };
    
    const smartSuggestions = this.getSmartCounterSuggestions();
    
    return {
      success: true,
      current: counters,
      suggested: smartSuggestions || {
        ordentlich: counters.ordentlich + 1,
        ausserordentlich: counters.ausserordentlich,
        fest: counters.fest
      },
      total: counters.ordentlich + counters.ausserordentlich + counters.fest
    };
  },
  
  // Intelligente Vorschläge basierend auf Historie
  getSmartCounterSuggestions() {
    try {
      const lastProtocol = this.getLastProtocolCounters();
      
      if (lastProtocol) {
        console.log('✅ Letzte Protokoll-Zähler gefunden:', lastProtocol);
        return {
          ordentlich: lastProtocol.ordentlich + 1,
          ausserordentlich: lastProtocol.ausserordentlich,
          fest: lastProtocol.fest
        };
      } else {
        console.log('⚠️ Keine Historie gefunden, verwende Properties + 1');
        const properties = PropertiesService.getScriptProperties();
        const current = {
          ordentlich: parseInt(properties.getProperty('counter_ordentlich') || '12'),
          ausserordentlich: parseInt(properties.getProperty('counter_ausserordentlich') || '6'),
          fest: parseInt(properties.getProperty('counter_fest') || '2')
        };
        
        return {
          ordentlich: current.ordentlich + 1,
          ausserordentlich: current.ausserordentlich,
          fest: current.fest
        };
      }
    } catch (error) {
      console.error('Fehler bei intelligenten Vorschlägen:', error);
      return null;
    }
  },
  
  // Letzte Protokoll-Zähler laden
  getLastProtocolCounters() {
    try {
      const data = SheetService.getSheetData('Protokolle');
      let latestProtocol = null;
      let latestDate = new Date(0);
      
      for (let i = 1; i < data.length; i++) {
        if (data[i].length > 5 && data[i][2] === 'Final') {
          const protocolDate = new Date(data[i][3]);
          if (protocolDate > latestDate) {
            latestDate = protocolDate;
            latestProtocol = {
              ordentlich: data[i][5],
              ausserordentlich: data[i][6], 
              fest: data[i][7],
              datum: protocolDate
            };
          }
        }
      }
      
      return latestProtocol;
    } catch (error) {
      console.error('Fehler beim Laden der letzten Protokoll-Zähler:', error);
      return null;
    }
  },
  
  // Zähler aktualisieren
  updateCounters(ordentlich, ausserordentlich, fest) {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'counter_ordentlich': ordentlich.toString(),
      'counter_ausserordentlich': ausserordentlich.toString(),
      'counter_fest': fest.toString()
    });
  },
  
  // Zähler zurücksetzen
  resetCounters() {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'counter_ordentlich': '12',
      'counter_ausserordentlich': '6', 
      'counter_fest': '2'
    });
    console.log('✅ Counter zurückgesetzt auf Standardwerte');
  }
};

const MemberService = {
  
  // Mitglieder für Frontend laden (unverändert)
  getMembersForFrontend() {
    try {
      const data = SheetService.getSheetData('Mitglieder');
      const members = [];
      
      for (let i = 1; i < data.length; i++) {
        if (data[i].length > 3 && data[i][3] === true) {
          const fullName = data[i][1];
          const nameParts = fullName.trim().split(' ');
          const familienname = nameParts[nameParts.length - 1];
          
          members.push({
            id: data[i][0],
            name: fullName,
            email: data[i][2], // Anzeige der ersten/Haupt-E-Mail
            reihenfolge: data[i][4] || 999,
            familienname: familienname
          });
        }
      }
      
      // Nach Familiennamen sortieren
      members.sort((a, b) => a.familienname.localeCompare(b.familienname, 'de', { sensitivity: 'base' }));
      
      console.log(`✅ ${members.length} aktive Mitglieder nach Familiennamen sortiert geladen`);
      return members;
      
    } catch (error) {
      console.error('❌ Fehler beim Laden der Mitglieder:', error);
      return this.getFallbackMembers();
    }
  },
  
  // ERWEITERT: Aktive Mitglieder für E-Mail mit ALLEN Adressen
  getActiveMembers() {
    const data = SheetService.getSheetData('Mitglieder');
    const activeMembers = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === true && data[i][2]) { // Aktiv und hat E-Mail
        const emailField = data[i][2].toString().trim();
        
        if (emailField) {
          // E-Mail-Adressen aufteilen (Semikolon-getrennt)
          const emailAddresses = this.parseEmailAddresses(emailField);
          
          // Nur gültige E-Mail-Adressen behalten
          const validEmails = emailAddresses.filter(email => this.isValidEmail(email));
          
          if (validEmails.length > 0) {
            activeMembers.push({
              id: data[i][0],
              name: data[i][1],
              emailAddresses: validEmails, // ARRAY von E-Mail-Adressen
              primaryEmail: validEmails[0], // Erste als Haupt-E-Mail
              totalEmails: validEmails.length
            });
            
            console.log(`📧 ${data[i][1]}: ${validEmails.length} E-Mail(s) → ${validEmails.join(', ')}`);
          } else {
            console.log(`⚠️ ${data[i][1]}: Keine gültigen E-Mail-Adressen gefunden in "${emailField}"`);
          }
        }
      }
    }
    
    console.log(`📧 ${activeMembers.length} aktive Mitglieder mit E-Mail gefunden`);
    
    // Statistik ausgeben
    const totalEmailAddresses = activeMembers.reduce((sum, member) => sum + member.totalEmails, 0);
    console.log(`📧 Gesamt ${totalEmailAddresses} E-Mail-Adressen für Versand`);
    
    return activeMembers;
  },
  
  // NEUE Funktion: E-Mail-Adressen parsen
  parseEmailAddresses(emailField) {
    if (!emailField) return [];
    
    // Semikolon-getrennt aufteilen und bereinigen
    return emailField
      .split(';')
      .map(email => email.trim())
      .filter(email => email.length > 0);
  },
  
  // NEUE Funktion: E-Mail-Validierung
  isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  },
  
  // NEUE Funktion: Alle E-Mail-Adressen als flache Liste
  getAllEmailAddresses() {
    const activeMembers = this.getActiveMembers();
    const allEmails = [];
    
    activeMembers.forEach(member => {
      allEmails.push(...member.emailAddresses);
    });
    
    return allEmails;
  },
  
  // NEUE Funktion: E-Mail-Statistik
  getEmailStatistics() {
    const activeMembers = this.getActiveMembers();
    
    const stats = {
      totalMembers: activeMembers.length,
      totalEmailAddresses: 0,
      membersWithMultipleEmails: 0,
      membersWithSingleEmail: 0,
      allEmailAddresses: []
    };
    
    activeMembers.forEach(member => {
      stats.totalEmailAddresses += member.totalEmails;
      stats.allEmailAddresses.push(...member.emailAddresses);
      
      if (member.totalEmails > 1) {
        stats.membersWithMultipleEmails++;
      } else {
        stats.membersWithSingleEmail++;
      }
    });
    
    return stats;
  },

  // Fallback-Mitglieder (unverändert)
  getFallbackMembers() {
    return [
      {id: 1, name: 'Josef Schnürer', email: 'josef.schnuerer@gmail.com', familienname: 'Schnürer'},
      {id: 2, name: 'Erwin Bertl', email: 'erwin@example.com', familienname: 'Bertl'},
      {id: 3, name: 'Norbert Glaubacker', email: 'norbert@example.com', familienname: 'Glaubacker'},
      {id: 4, name: 'Gerhard John', email: 'gerhard@example.com', familienname: 'John'},
      {id: 5, name: 'Gerald Faic', email: 'gerald@example.com', familienname: 'Faic'},
      {id: 6, name: 'Otto Thür', email: 'otto@example.com', familienname: 'Thür'},
      {id: 7, name: 'Christian Thür', email: 'christian@example.com', familienname: 'Thür'},
      {id: 8, name: 'Friedrich Fochler', email: 'friedrich@example.com', familienname: 'Fochler'},
      {id: 9, name: 'Johannes Mader', email: 'johannes@example.com', familienname: 'Mader'}
    ].sort((a, b) => a.familienname.localeCompare(b.familienname, 'de'));
  }
};

// AUTO-RESUME Service - Production Ready
const AutoResumeService = {

  // Letztes Draft-Protokoll finden und laden
  getLastDraftProtocol() {
    try {
      console.log('Auto-Resume: Suche Draft-Protokoll...');

      const data = SheetService.getSheetData('Protokolle');

      // Sicherheitscheck
      if (!data || data.length <= 1) {
        console.log('ℹ️ Keine Protokoll-Daten gefunden');
        return {
          success: false,
          message: 'Keine Protokolle vorhanden'
        };
      }

      // Vom Ende her suchen (neueste zuerst)
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];

        // Sicherheitscheck für Row
        if (!row || row.length < 3) {
          console.log(`⚠️ Überspringe unvollständige Zeile: ${i}`);
          continue;
        }

        const status = row[2]; // Spalte C: Status
        const protocolId = row[0]; // Spalte A: ID

        // Wenn Status nicht "Versendet" ist, dann ist es ein Draft
        if (status !== 'Versendet' && protocolId) {
          console.log(`📋 Draft-Protokoll gefunden: ${protocolId}`);

          const protocolData = this.loadProtocolData(row);
          if (protocolData.success) {
            return protocolData;
          }
        }
      }

      console.log('ℹ️ Kein Draft-Protokoll gefunden - neue Sitzung');
      return {
        success: false,
        message: 'Kein Draft gefunden'
      };

    } catch (error) {
      console.error('❌ Fehler beim Laden des letzten Drafts:', error);
      return {
        success: false,
        message: `Fehler: ${error.toString()}`
      };
    }
  },

  // Protokoll-Daten aus Sheet-Zeile laden
  loadProtocolData(row) {
    try {
      const protocolId = row[0];

      const protocolData = {
        id: protocolId,
        status: row[2],
        counters: {
          ordentlich: row[5] || 12,
          ausserordentlich: row[6] || 6,
          fest: row[7] || 2
        },
        sitzungsdatum: this.formatDateForInput(row[3]),
        ort: row[4] || 'Clublokal',
        naechsterTermin: this.formatDateForInput(row[10]),
        anwesende: this.parseAnwesendeIds(row[9]),
        createdAt: row[1], // Date-Objekt - wird in Code.gs zu String konvertiert
        agendaItems: []
      };

      // Agenda-Items laden
      try {
        const agendaItems = this.loadAgendaItems(protocolId);
        protocolData.agendaItems = agendaItems;
      } catch (agendaError) {
        console.log(`⚠️ Agenda-Items konnten nicht geladen werden: ${agendaError}`);
        protocolData.agendaItems = [];
      }

      console.log(`✅ Protokoll-Daten geladen: ${protocolId}`);
      StatusService.addStatusToNotes(protocolId, '🔄 Protokoll automatisch geladen für Weiterarbeit');

      return {
        success: true,
        protocolData: protocolData,
        message: `Draft-Protokoll ${protocolId} geladen`
      };

    } catch (error) {
      console.error('❌ Fehler beim Laden der Protokoll-Daten:', error);
      return {
        success: false,
        message: `Fehler bei Protokoll-Daten: ${error.toString()}`
      };
    }
  },

  // Agenda-Items für Protokoll laden
  loadAgendaItems(protocolId) {
    try {
      console.log(`🔍 Lade Agenda-Items für Protokoll: ${protocolId}`);
      
      const agendaData = SheetService.getSheetData('Tagesordnungspunkte');
      const items = [];
      
      for (let i = 1; i < agendaData.length; i++) {
        if (agendaData[i][1] === protocolId) {
          const item = {
            id: agendaData[i][0],
            nummer: agendaData[i][2],
            titel: agendaData[i][3] || '',
            beschreibung: agendaData[i][4] || '',
            llmImprove: agendaData[i][5] || false,
            bilder: []
          };
          
          // Bilder für diesen Agenda-Item laden
          try {
            item.bilder = ImageService.getImagesForAgendaItem(item.id);
          } catch (imageError) {
            console.log(`⚠️ Bilder für Agenda-Item ${item.id} konnten nicht geladen werden`);
            item.bilder = [];
          }
          
          items.push(item);
        }
      }
      
      // Nach Nummer sortieren
      items.sort((a, b) => a.nummer - b.nummer);
      
      console.log(`✅ ${items.length} Agenda-Items geladen`);
      return items;
      
    } catch (error) {
      console.error('❌ Fehler beim Laden der Agenda-Items:', error);
      return [];
    }
  },

  // Anwesende IDs aus String parsen
  parseAnwesendeIds(anwesendeString) {
    if (!anwesendeString) {
      return ['1']; // Fallback: Josef Schnürer
    }
    
    try {
      return anwesendeString
        .toString()
        .replace(/'/g, '') // Entferne Anführungszeichen
        .split(',')
        .map(id => id.trim())
        .filter(id => id); // Entferne leere Strings
    } catch (error) {
      console.log(`⚠️ Fehler beim Parsen der Anwesenden-IDs: ${error}`);
      return ['1'];
    }
  },

  // Datum für HTML-Input formatieren (YYYY-MM-DD)
  formatDateForInput(dateValue) {
    if (!dateValue) return '';
    
    try {
      const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
      
      if (isNaN(date.getTime())) {
        console.log(`❌ Ungültiges Datum: ${dateValue}`);
        return '';
      }
      
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      
      const result = `${year}-${month}-${day}`;
      console.log(`✅ Datum formatiert: ${dateValue} → ${result}`);
      return result;
      
    } catch (error) {
      console.log(`❌ Fehler beim Datum-Formatieren: ${error}, Input: ${dateValue}`);
      return '';
    }
  }

};

// Protocol Service - ERWEITERT für Auto-Resume
const ProtocolService = {
  
  // ERWEITERTE saveDraft-Funktion - KEINE neuen IDs bei bestehendem Protokoll
  saveDraft(formData) {
    try {
      const protocolSheet = SheetService.getSheet('Protokolle');
      
      // WICHTIG: ID beibehalten wenn vorhanden
      const protocolId = formData.id || IdService.generateProtocolId();
      const isExistingProtocol = !!formData.id;
      
      console.log(isExistingProtocol ? 
        `🔄 Aktualisiere bestehendes Protokoll: ${protocolId}` : 
        `📝 Erstelle neues Protokoll: ${protocolId}`);
      
      const nextTermin = formData.naechsterTermin ? new Date(formData.naechsterTermin) : new Date();
      
      const row = [
        protocolId,
        isExistingProtocol ? this.getExistingCreatedDate(protocolId) : new Date(), // Erstellungsdatum beibehalten
        'Draft',
        new Date(formData.sitzungsdatum),
        formData.ort,
        formData.counters.ordentlich,
        formData.counters.ausserordentlich,
        formData.counters.fest,
        formData.counters.ordentlich + formData.counters.ausserordentlich + formData.counters.fest,
        "'" + formData.anwesende.join(','),
        UtilService.formatDateDE(nextTermin),
        '', '', '',
        isExistingProtocol ? 
          this.updateExistingNotes(protocolId, 'Entwurf aktualisiert') :
          `[${new Date().toLocaleString('de-DE')}] Entwurf gespeichert`
      ];
      
      const existingRow = StatusService.findProtocolRow(protocolId);
      if (existingRow > 0) {
        // Update bestehende Zeile
        protocolSheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
        StatusService.addStatusToNotes(protocolId, '🔄 Entwurf aktualisiert (Auto-Resume)');
      } else {
        // Neue Zeile erstellen
        protocolSheet.appendRow(row);
        StatusService.addStatusToNotes(protocolId, '📝 Neuer Entwurf erstellt');
      }
      
      // Tagesordnungspunkte speichern - ERWEITERT für bestehende IDs
      const agendaIdMapping = this.saveAgendaItemsWithExistingIds(protocolId, formData.agendaItems, isExistingProtocol);
      StatusService.addStatusToNotes(protocolId, `${formData.agendaItems.length} Tagesordnungspunkte gespeichert`);
      
      // Bild-Referenzen korrigieren nur bei neuen Mappings
      if (Object.keys(agendaIdMapping).length > 0) {
        const fixResult = ImageService.fixImageReferences(agendaIdMapping);
        
        if (fixResult.success && fixResult.updatedCount > 0) {
          StatusService.addStatusToNotes(protocolId, `${fixResult.updatedCount} Bild-Referenzen korrigiert`);
        } else {
          console.log('⚠️ Keine Bild-Referenzen korrigiert oder Fehler');
        }
      } else {
        console.log('⚠️ ID-Mapping ist leer - keine Bild-Korrektur möglich');
      }
      
      return {
        success: true,
        protocolId: protocolId,
        agendaIdMapping: agendaIdMapping,
        isExistingProtocol: isExistingProtocol,
        message: isExistingProtocol ? 'Entwurf erfolgreich aktualisiert' : 'Entwurf erfolgreich gespeichert'
      };
      
    } catch (error) {
      console.error('Fehler beim Speichern:', error);
      if (formData.id) {
        StatusService.addStatusToNotes(formData.id, `FEHLER beim Speichern: ${error.toString()}`);
      }
      return {
        success: false,
        message: 'Fehler beim Speichern: ' + error.toString()
      };
    }
  },
  
  // NEUE Hilfsfunktion: Bestehendes Erstellungsdatum holen
  getExistingCreatedDate(protocolId) {
    try {
      const data = SheetService.getSheetData('Protokolle');
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === protocolId) {
          return data[i][1]; // Spalte B: Erstellungsdatum
        }
      }
      return new Date(); // Fallback
    } catch (error) {
      return new Date();
    }
  },
  
  // NEUE Hilfsfunktion: Bestehende Notizen aktualisieren
  updateExistingNotes(protocolId, newMessage) {
    try {
      const sheet = SheetService.getSheet('Protokolle');
      const rowIndex = StatusService.findProtocolRow(protocolId);
      
      if (rowIndex > 0) {
        const currentNotes = sheet.getRange(rowIndex, 15).getValue() || '';
        const timestamp = new Date().toLocaleString('de-DE');
        const newEntry = `[${timestamp}] ${newMessage}`;
        return currentNotes ? `${currentNotes}\n${newEntry}` : newEntry;
      }
      
      return `[${new Date().toLocaleString('de-DE')}] ${newMessage}`;
    } catch (error) {
      return `[${new Date().toLocaleString('de-DE')}] ${newMessage}`;
    }
  },
  
  // ERWEITERTE Tagesordnungspunkte-Speichern - KEINE neuen IDs bei bestehenden Items
  saveAgendaItemsWithExistingIds(protocolId, agendaItems, isExistingProtocol) {
    const sheet = SheetService.getSheet('Tagesordnungspunkte');
    
    if (isExistingProtocol) {
      // Bei bestehendem Protokoll: Nur löschen was nicht mehr da ist
      this.updateExistingAgendaItems(sheet, protocolId, agendaItems);
    } else {
      // Bei neuem Protokoll: Alle alten löschen (falls vorhanden)
      this.deleteOldAgendaItems(sheet, protocolId);
    }
    
    const idMapping = {};
    
    // Agenda-Items verarbeiten
    agendaItems.forEach((item, index) => {
      let realAgendaId;
      
      if (item.id && isExistingProtocol) {
        // Bestehende ID verwenden
        realAgendaId = item.id;
        console.log(`🔄 Verwende bestehende AgendaID: ${realAgendaId}`);
        
        // Item aktualisieren
        this.updateExistingAgendaItem(sheet, realAgendaId, item, protocolId);
      } else {
        // Neue ID erstellen
        realAgendaId = IdService.generateAgendaId();
        console.log(`📝 Neue AgendaID erstellt: ${realAgendaId}`);
        
        // Neues Item erstellen
        this.createNewAgendaItem(sheet, realAgendaId, item, protocolId);
      }
      
      // Frontend-Mapping IMMER erstellen (sowohl für neue als auch bestehende Items)
      const frontendId = `agenda-${item.nummer}`;
      idMapping[frontendId] = realAgendaId;
      console.log(`🔄 Mapping erstellt: ${frontendId} → ${realAgendaId}`);
    });
    
    return idMapping;
  },
  
  // NEUE Hilfsfunktion: Bestehende Agenda-Items aktualisieren
  updateExistingAgendaItems(sheet, protocolId, newAgendaItems) {
    const data = sheet.getDataRange().getValues();
    const existingItemIds = [];
    
    // Sammle alle bestehenden Item-IDs
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === protocolId) {
        existingItemIds.push(data[i][0]);
      }
    }
    
    // Sammle alle neuen Item-IDs
    const newItemIds = newAgendaItems.filter(item => item.id).map(item => item.id);
    
    // Lösche Items die nicht mehr da sind
    const toDelete = existingItemIds.filter(id => !newItemIds.includes(id));
    toDelete.forEach(itemId => {
      this.deleteAgendaItemById(sheet, itemId);
      console.log(`🗑️ Agenda-Item gelöscht: ${itemId}`);
    });
  },
  
  // NEUE Hilfsfunktion: Alle alten Agenda-Items löschen
  deleteOldAgendaItems(sheet, protocolId) {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === protocolId) {
        sheet.deleteRow(i + 1);
      }
    }
  },
  
  // NEUE Hilfsfunktion: Bestehenden Agenda-Item aktualisieren
  updateExistingAgendaItem(sheet, agendaId, item, protocolId) {
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === agendaId) {
        const row = [
          agendaId,
          protocolId,
          item.nummer,
          item.titel,
          item.beschreibung,
          item.llmImprove || false,
          data[i][6], // Original_Text beibehalten
          data[i][7], // Verbessert_Am beibehalten
          item.bilder && item.bilder.length > 0,
          data[i][9]  // Erstellt_Am beibehalten
        ];
        
        sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
        console.log(`🔄 Agenda-Item aktualisiert: ${agendaId}`);
        return;
      }
    }
  },
  
  // NEUE Hilfsfunktion: Neuen Agenda-Item erstellen
  createNewAgendaItem(sheet, agendaId, item, protocolId) {
    const row = [
      agendaId,
      protocolId,
      item.nummer,
      item.titel,
      item.beschreibung,
      item.llmImprove || false,
      '', '',
      item.bilder && item.bilder.length > 0,
      new Date()
    ];
    
    sheet.appendRow(row);
    console.log(`📝 Neuer Agenda-Item erstellt: ${agendaId}`);
  },
  
  // NEUE Hilfsfunktion: Agenda-Item nach ID löschen
  deleteAgendaItemById(sheet, agendaId) {
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === agendaId) {
        sheet.deleteRow(i + 1);
        return;
      }
    }
  },
  
  // Finales Protokoll erstellen
  createFinalProtocol(formData) {
    let protocolId = formData.id;
    
    try {
      if (!protocolId) {
        console.log('⚠️ Keine Protokoll-ID gefunden - führe automatisches Zwischenspeichern durch...');
        StatusService.addStatusToNotes('TEMP', '🔄 Automatisches Zwischenspeichern vor Finalisierung...');
        
        const draftResult = this.saveDraft(formData);
        
        if (!draftResult.success) {
          console.error('❌ Automatisches Zwischenspeichern fehlgeschlagen:', draftResult.message);
          return {
            success: false,
            message: 'Fehler beim automatischen Zwischenspeichern: ' + draftResult.message
          };
        }
        protocolId = draftResult.protocolId;
        console.log(`✅ Automatisches Zwischenspeichern erfolgreich - ID: ${protocolId}`);
        StatusService.addStatusToNotes(protocolId, '✅ Automatisch zwischengespeichert vor Finalisierung');
      } else {
        console.log(`🔄 Verwende vorhandene Protokoll-ID: ${protocolId}`);
        
        // Bestehende Daten vor Finalisierung nochmal aktualisieren
        console.log('🔄 Aktualisiere Daten vor Finalisierung...');
        const updateResult = this.saveDraft(formData);
        
        if (!updateResult.success) {
          console.error('⚠️ Warnung: Daten-Update vor Finalisierung fehlgeschlagen:', updateResult.message);
          StatusService.addStatusToNotes(protocolId, '⚠️ Warnung: Daten-Update vor Finalisierung fehlgeschlagen');
        } else {
          StatusService.addStatusToNotes(protocolId, '🔄 Daten vor Finalisierung aktualisiert');
        }
      }

      // VALIDIERUNG: Protokoll-ID muss jetzt existieren
      if (!protocolId) {
        throw new Error('Kritischer Fehler: Keine Protokoll-ID nach Zwischenspeicherung');
      }
      
      // 1. Status auf "Final" ändern
      this.updateProtocolStatus(protocolId, 'Final');
      StatusService.addStatusToNotes(protocolId, '📝 Finales Protokoll - Start der Erstellung');
      
      // 2. Zähler aktualisieren
      CounterService.updateCounters(
        formData.counters.ordentlich,
        formData.counters.ausserordentlich,
        formData.counters.fest
      );
      StatusService.addStatusToNotes(protocolId, `🔢 Zähler aktualisiert: ${formData.counters.ordentlich}/${formData.counters.ausserordentlich}/${formData.counters.fest}`);
      
      // 3. Google Doc erstellen
      StatusService.addStatusToNotes(protocolId, '📄 Google Doc wird erstellt...');
      const docResult = DocumentService.createProtocolDocument(protocolId, formData);
      if (!docResult.success) {
        StatusService.addStatusToNotes(protocolId, `❌ FEHLER bei Google Doc: ${docResult.message}`);
        return docResult;
      }
      StatusService.addStatusToNotes(protocolId, `📄 Google Doc erstellt: ${docResult.docName}`);
      
      // 4. PDF erstellen
      StatusService.addStatusToNotes(protocolId, '📋 PDF wird erstellt...');
      const pdfResult = DocumentService.createProtocolPDF(docResult.docId);
      if (!pdfResult.success) {
        StatusService.addStatusToNotes(protocolId, `❌ FEHLER bei PDF: ${pdfResult.message}`);
        return pdfResult;
      }
      StatusService.addStatusToNotes(protocolId, `📋 PDF erstellt: ${pdfResult.pdfName}`);
      
      // 5. E-Mail versenden
      StatusService.addStatusToNotes(protocolId, '📧 E-Mail wird versendet...');
      const emailResult = EmailService.sendProtocolEmail(protocolId, pdfResult.pdfId, formData);
      if (!emailResult.success) {
        StatusService.addStatusToNotes(protocolId, `❌ FEHLER bei E-Mail: ${emailResult.message}`);
        return emailResult;
      }
      StatusService.addStatusToNotes(protocolId, `📧 E-Mail versendet an ${emailResult.recipients} Empfänger`);
      
      // 6. Status finalisieren
      StatusService.updateProtocolStatus(protocolId, 'Versendet', docResult.docId, pdfResult.pdfId);
      StatusService.addStatusToNotes(protocolId, '✅ Protokoll vollständig abgeschlossen');
      
      return {
        success: true,
        protocolId: protocolId,
        docId: docResult.docId,
        pdfId: pdfResult.pdfId,
        message: 'Protokoll erfolgreich erstellt und versendet!'
      };
      
    } catch (error) {
      if (protocolId) {
        StatusService.addStatusToNotes(protocolId, `❌ KRITISCHER FEHLER: ${error.toString()}`);
      }
      return {
        success: false,
        message: 'Fehler bei Protokollerstellung: ' + error.toString()
      };
    }
  },

  // Status eines bestehenden Protokolls ändern
  updateProtocolStatus(protocolId, newStatus) {
    try {
      const sheet = SheetService.getSheet('Protokolle');
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === protocolId) {
          sheet.getRange(i + 1, 3).setValue(newStatus);
          return;
        }
      }
    } catch (error) {
      console.error('Fehler beim Status-Update:', error);
    }
  },

  // Echte Agenda-ID anhand der Nummer finden
  findRealAgendaIdByNumber(protocolId, number) {
    try {
      const data = SheetService.getSheetData('Tagesordnungspunkte');
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === protocolId && data[i][2] === number) {
          return data[i][0]; // Echte ID zurückgeben
        }
      }
      
      return null;
      
    } catch (error) {
      console.error('Fehler beim Suchen der echten AgendaID:', error);
      return null;
    }
  }

};
