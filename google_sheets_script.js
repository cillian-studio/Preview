// ============================================================
// ANLEITUNG:
// 1. Erstelle ein neues Google Sheet (sheets.google.com → leer)
// 2. Im Sheet: Erweiterungen → Apps Script
// 3. Loesche den bestehenden Code
// 4. Kopiere DIESEN ganzen Code rein
// 5. Klick "Bereitstellen" → "Neue Bereitstellung"
// 6. Typ: "Web-App"
// 7. Ausfuehren als: "Ich"
// 8. Zugriff: "Jeder"
// 9. Klick "Bereitstellen"
// 10. Google fragt nach Berechtigung → "Zulassen"
// 11. Kopiere die URL die erscheint
// 12. Schick mir die URL → ich bau sie in die Seite ein
// ============================================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Beim allerersten Eintrag: Header-Zeile erstellen
    if (sheet.getLastRow() === 0) {
      var headers = [
        'Zeitstempel',
        'Schulung aktuell',
        'Groesste Probleme',
        'Schmerzpunkt',
        'Dokumentation',
        'Teilnehmer Anzahl',
        'Neue Teilnehmer',
        'Neue pro Jahr',
        'Schulungsfrequenz',
        'Anzahl Firmen',
        'Teilnehmertyp',
        'Themen',
        'Material',
        'Material Details',
        'Schulungsdauer',
        'Pruefung',
        'Bestehensgrenze',
        'Bei Durchfallen',
        'Zertifikat',
        'Geraete',
        'Offline',
        'Sprache',
        'Systemverwalter',
        'Land',
        'Aufbewahrungsdauer',
        'Zeitplan',
        'Software Integration',
        'Wichtigste Funktion',
        'Sonstiges'
      ];
      sheet.appendRow(headers);

      // Header fett + blauer Hintergrund
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#2563eb');
      headerRange.setFontColor('#ffffff');
      headerRange.setWrap(true);

      // Spaltenbreite setzen
      sheet.setColumnWidth(1, 160);  // Zeitstempel
      for (var i = 2; i <= headers.length; i++) {
        sheet.setColumnWidth(i, 200);
      }

      // Header fixieren
      sheet.setFrozenRows(1);
    }

    var params = e.parameter;

    // Mehrfachwerte (Checkboxen) zusammenfuegen
    var getField = function(name) {
      // Versuche Mehrfachwerte (kommasepariert von URLSearchParams)
      var val = params[name] || '';
      // Sonstiges-Feld dazuhaengen wenn vorhanden
      var sonstiges = params[name + ' Sonstiges'] || '';
      if (sonstiges) {
        val = val ? val + ', ' + sonstiges : sonstiges;
      }
      return val;
    };

    var row = [
      new Date().toLocaleString('de-AT', { timeZone: 'Europe/Vienna' }),
      getField('Schulung aktuell'),
      getField('Groesste Probleme'),
      getField('Schmerzpunkt'),
      getField('Dokumentation'),
      getField('Teilnehmer Anzahl'),
      getField('Neue Teilnehmer'),
      getField('Neue pro Jahr'),
      getField('Schulungsfrequenz'),
      getField('Anzahl Firmen'),
      getField('Teilnehmertyp'),
      getField('Themen'),
      getField('Material'),
      getField('Material Details'),
      getField('Schulungsdauer'),
      getField('Pruefung'),
      getField('Bestehensgrenze'),
      getField('Bei Durchfallen'),
      getField('Zertifikat'),
      getField('Geraete'),
      getField('Offline'),
      getField('Sprache'),
      getField('Systemverwalter'),
      getField('Land'),
      getField('Aufbewahrungsdauer'),
      getField('Zeitplan'),
      getField('Software Integration'),
      getField('Wichtigste Funktion'),
      getField('Sonstiges')
    ];

    sheet.appendRow(row);

    // Zeilenumbruch fuer die neue Zeile aktivieren
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1, 1, row.length).setWrap(true);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test-Funktion: Kannst du ausfuehren um zu pruefen ob alles geht
function testDoPost() {
  var testData = {
    parameter: {
      'Schulung aktuell': 'Praesenzschulung',
      'Groesste Probleme': 'Zu viel Zeit',
      'Teilnehmer Anzahl': '25',
      'Land': 'Oesterreich'
    }
  };
  var result = doPost(testData);
  Logger.log(result.getContent());
  Logger.log('Test-Eintrag erstellt! Schau ins Sheet.');
}
