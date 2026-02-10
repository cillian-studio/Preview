function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var isFirst = sheet.getLastRow() === 0;

    if (isFirst) {
      var headers = ['Nr.','Zeitstempel','Schulung aktuell','Groesste Probleme','Schmerzpunkt','Dokumentation','Teilnehmer Anzahl','Neue Teilnehmer','Neue pro Jahr','Schulungsfrequenz','Anzahl Firmen','Teilnehmertyp','Themen','Material','Material Details','Schulungsdauer','Pruefung','Bestehensgrenze','Bei Durchfallen','Zertifikat','Geraete','Offline','Sprache','Systemverwalter','Land','Aufbewahrungsdauer','Zeitplan','Software Integration','Wichtigste Funktion','Sonstiges'];
      sheet.appendRow(headers);
      var h = sheet.getRange(1, 1, 1, headers.length);
      h.setFontWeight('bold');
      h.setBackground('#1a1a2e');
      h.setFontColor('#ffffff');
      h.setFontSize(10);
      h.setHorizontalAlignment('center');
      h.setVerticalAlignment('middle');
      h.setWrap(true);
      sheet.setRowHeight(1, 40);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 40);
      sheet.setColumnWidth(2, 160);
      for (var i = 3; i <= headers.length; i++) {
        sheet.setColumnWidth(i, 180);
      }
      sheet.setColumnWidths(15, 1, 250);
      sheet.setColumnWidths(28, 1, 250);
      sheet.setColumnWidths(29, 1, 250);
      sheet.setColumnWidths(30, 1, 250);
    }

    var p = e.parameter;
    var g = function(n) {
      var v = p[n] || '';
      var s = p[n + ' Sonstiges'] || '';
      return s ? (v ? v + ', ' + s : s) : v;
    };

    var nr = sheet.getLastRow();
    var row = [nr, new Date().toLocaleString('de-AT'), g('Schulung aktuell'), g('Groesste Probleme'), g('Schmerzpunkt'), g('Dokumentation'), g('Teilnehmer Anzahl'), g('Neue Teilnehmer'), g('Neue pro Jahr'), g('Schulungsfrequenz'), g('Anzahl Firmen'), g('Teilnehmertyp'), g('Themen'), g('Material'), g('Material Details'), g('Schulungsdauer'), g('Pruefung'), g('Bestehensgrenze'), g('Bei Durchfallen'), g('Zertifikat'), g('Geraete'), g('Offline'), g('Sprache'), g('Systemverwalter'), g('Land'), g('Aufbewahrungsdauer'), g('Zeitplan'), g('Software Integration'), g('Wichtigste Funktion'), g('Sonstiges')];
    sheet.appendRow(row);

    var lastRow = sheet.getLastRow();
    var rowRange = sheet.getRange(lastRow, 1, 1, row.length);
    rowRange.setWrap(true);
    rowRange.setVerticalAlignment('top');
    rowRange.setFontSize(10);
    rowRange.setBorder(null, null, true, null, null, null, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    if (lastRow % 2 === 0) {
      rowRange.setBackground('#f8f9fa');
    } else {
      rowRange.setBackground('#ffffff');
    }

    sheet.getRange(lastRow, 1).setHorizontalAlignment('center').setFontColor('#999999');
    sheet.getRange(lastRow, 2).setFontColor('#666666').setFontSize(9);

    return ContentService.createTextOutput('ok');
  } catch(err) {
    return ContentService.createTextOutput('error: ' + err);
  }
}
