/**
 * WebApp endpoint to append submitted form data into Google Sheets.
 * Deploy: Deploy -> New deployment -> Web app
 * Execute as: Me
 * Who has access: Anyone (or Anyone within your domain)
 */

const SPREADSHEET_ID = 'INCOLLA_QUI_ID_SHEET';
const SHEET_NAME = 'Risposte';

const COLUMNS = [
  "numero_report",
  "timestamp",
  "cliente",
  "luogo",
  "data",
  "referente_bs",
  "referente_cliente",
  "agente_bs",
  "oggetto",
  "note_sez1",
  "tipologia_servizio",
  "temperatura",
  "aggressivi_chimici_op",
  "cicli_lavaggio",
  "ciclature_termiche",
  "tempi_rientro",
  "problema_corrosione",
  "problema_vaiolature",
  "problema_erosione",
  "problema_cricche",
  "note_sez2",
  "sic_aggressivi",
  "sic_atex",
  "sic_polveri",
  "sic_quota",
  "sic_dpi",
  "sic_confinato",
  "sic_220v",
  "sic_formazione",
  "note_sez3",
  "sez4_piano_lavoro",
  "sez4_movimentazione",
  "sez4_distanza_parcheggio",
  "sez4_area_aperto",
  "sez4_area_chiuso",
  "sez4_area_riscaldata",
  "sez4_scale",
  "sez4_ascensore",
  "sez4_passi_uomo_diam",
  "sez4_illuminazione",
  "sez4_corpo",
  "sez4_braccio",
  "sez4_mano",
  "sez4_ventilazione",
  "note_sez4",
  "note_raccolta_dati"
];

function doGet() {
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ status: 'error', message: 'Missing body' });
    }

    const payload = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    const header = ensureHeader_(sheet);
    const row = header.map(k => (payload[k] !== undefined && payload[k] !== null) ? String(payload[k]) : '');

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      sheet.appendRow(row);
    } finally {
      lock.releaseLock();
    }

    return jsonOut({ status: 'ok' });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

function ensureHeader_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0) {
    sheet.appendRow(COLUMNS);
    return COLUMNS;
  }

  const existing = sheet.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0]
    .map(v => (v || '').toString().trim());

  const hasAnyHeader = existing.some(v => v);
  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
    return COLUMNS;
  }

  const existingFiltered = existing.filter(v => v);
  const missing = COLUMNS.filter(c => existingFiltered.indexOf(c) === -1);

  if (missing.length) {
    sheet.getRange(1, existingFiltered.length + 1, 1, missing.length).setValues([missing]);
    return existingFiltered.concat(missing);
  }

  return existingFiltered;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
