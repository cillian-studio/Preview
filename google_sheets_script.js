function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() === 0) {
      var headers = ['Zeitstempel','Schulung aktuell','Groesste Probleme','Schmerzpunkt','Dokumentation','Teilnehmer Anzahl','Neue Teilnehmer','Neue pro Jahr','Schulungsfrequenz','Anzahl Firmen','Teilnehmertyp','Themen','Material','Material Details','Schulungsdauer','Pruefung','Bestehensgrenze','Bei Durchfallen','Zertifikat','Geraete','Offline','Sprache','Systemverwalter','Land','Aufbewahrungsdauer','Zeitplan','Software Integration','Wichtigste Funktion','Sonstiges'];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#2563eb').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    var p = e.parameter;
    var g = function(n) {
      var v = p[n] || '';
      var s = p[n + ' Sonstiges'] || '';
      return s ? (v ? v + ', ' + s : s) : v;
    };
    var row = [new Date().toLocaleString('de-AT'),g('Schulung aktuell'),g('Groesste Probleme'),g('Schmerzpunkt'),g('Dokumentation'),g('Teilnehmer Anzahl'),g('Neue Teilnehmer'),g('Neue pro Jahr'),g('Schulungsfrequenz'),g('Anzahl Firmen'),g('Teilnehmertyp'),g('Themen'),g('Material'),g('Material Details'),g('Schulungsdauer'),g('Pruefung'),g('Bestehensgrenze'),g('Bei Durchfallen'),g('Zertifikat'),g('Geraete'),g('Offline'),g('Sprache'),g('Systemverwalter'),g('Land'),g('Aufbewahrungsdauer'),g('Zeitplan'),g('Software Integration'),g('Wichtigste Funktion'),g('Sonstiges')];
    sheet.appendRow(row);
    return ContentService.createTextOutput('ok');
  } catch(err) {
    return ContentService.createTextOutput('error: ' + err);
  }
}
