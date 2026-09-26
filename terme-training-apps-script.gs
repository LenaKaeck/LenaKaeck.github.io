/**
 * Google Apps Script backend for terme-training.html.
 *
 * Setup: see the step-by-step instructions the teacher received in chat.
 * This file lives in the repo only as a reference/backup — the code that
 * actually runs lives inside the Google Apps Script project attached to
 * the Google Sheet, and must be pasted there manually (Apps Script does
 * not read this file from GitHub).
 *
 * Expects a Google Sheet with these tabs:
 *   - "Whitelist": column A=Codename (one header row, then one erlaubter Codename pro Zeile)
 *   - "Ergebnisse": columns A=Zeitstempel, B=Codename, C=Station, D=Richtig, E=Von
 *     (header row optional — appendRow just adds below whatever is already there)
 *   - "Gesamtergebnisse": wird automatisch überschrieben — eine Zeile pro Codename
 *     mit der Summe ALLER jemals erzielten Punkte (über alle Einsendungen/Runden
 *     hinweg — kann bei mehrfachem Spielen also über 66 steigen) und dem
 *     Zeitpunkt der letzten Aktivität.
 *   - "Gast-Ergebnisse": wie "Ergebnisse", aber nur Einsendungen mit dem Codenamen
 *     "Gast" (der keine eindeutige Person identifiziert). Keine eigene
 *     Gesamtergebnisse-Zusammenfassung dafür — bewusst nur der Rohlog.
 */

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'whitelist') {
    var names = getWhitelistNames();
    return ContentService
      .createTextOutput(JSON.stringify(names))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('OK');
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput('invalid payload');
  }

  var name = (data.name || '').toString().trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // "Gast" identifiziert keine bestimmte Person — geht nur in den Rohlog,
  // wird nie gegen die Whitelist geprüft und fließt nicht in Gesamtergebnisse ein.
  if (name.toLowerCase() === 'gast') {
    var gastSheet = ss.getSheetByName('Gast-Ergebnisse') || ss.insertSheet('Gast-Ergebnisse');
    gastSheet.appendRow([new Date(), 'Gast', data.station || '', data.correct, data.total]);
    return ContentService.createTextOutput('OK (Gast)');
  }

  var whitelist = getWhitelistNames();
  if (whitelist.indexOf(name) === -1) {
    // Not on the whitelist — silently ignored, nothing is written.
    return ContentService.createTextOutput('rejected: not on whitelist');
  }

  var sheet = ss.getSheetByName('Ergebnisse');
  sheet.appendRow([
    new Date(),
    name,
    data.station || '',
    data.correct,
    data.total
  ]);
  updateSummary();
  return ContentService.createTextOutput('OK');
}

function getWhitelistNames() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Whitelist');
  var values = sheet.getDataRange().getValues();
  var names = [];
  // Row 0 is assumed to be the header row (z.B. "Codename") and wird übersprungen.
  for (var i = 1; i < values.length; i++) {
    var codename = (values[i][0] || '').toString().trim();
    if (codename) names.push(codename);
  }
  names.sort(function (a, b) { return a.localeCompare(b, 'de'); });
  return names;
}

// Baut den Tab "Gesamtergebnisse" komplett neu auf: pro Codename die Summe
// ALLER jemals erzielten Punkte (jede Einsendung zählt voll, auch bei
// mehrfach gespielten Stationen — kann also über die Zeit beliebig weiter
// wachsen), plus der Zeitpunkt der letzten Aktivität. Kann auch manuell im
// Apps-Script-Editor ausgeführt werden (Funktion "updateSummary" auswählen →
// ▶ Run), um die Übersicht ohne neue Einsendung sofort zu aktualisieren.
function updateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = ss.getSheetByName('Ergebnisse').getDataRange().getValues();
  var totals = {};   // Codename -> Summe aller "Richtig"-Werte
  var lastSeen = {}; // Codename -> letzter Zeitstempel

  for (var i = 0; i < results.length; i++) {
    var row = results[i];
    var name = (row[1] || '').toString().trim();
    var correct = Number(row[3]);
    // Kopfzeile/leere Zeilen überspringen — anhand der Punktzahl, nicht des
    // Zeitstempel-Typs (der je nach Zellformatierung mal Date, mal Text ist).
    if (!name || isNaN(correct)) continue;
    var rawTs = row[0];
    var ts = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
    totals[name] = (totals[name] || 0) + correct;
    if (!isNaN(ts.getTime()) && (!lastSeen[name] || ts > lastSeen[name])) lastSeen[name] = ts;
  }

  var summarySheet = ss.getSheetByName('Gesamtergebnisse') || ss.insertSheet('Gesamtergebnisse');
  summarySheet.clearContents();
  summarySheet.appendRow(['Codename', 'Punkte', 'Zeitstempel']);

  var names = Object.keys(totals).sort(function (a, b) { return a.localeCompare(b, 'de'); });
  names.forEach(function (name) {
    summarySheet.appendRow([name, totals[name], lastSeen[name] || '']);
  });
}
