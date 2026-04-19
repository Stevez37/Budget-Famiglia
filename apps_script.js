// ============================================================
// BUDGET APP — Google Apps Script
// Incolla questo codice in: Extensions → Apps Script
// poi: Deploy → Manage deployments → modifica quello esistente
//   oppure New deployment → Web App
//   Execute as: Me
//   Who has access: Anyone
//
// Per creare i fogli mensili: vai su Esegui → creaFogliMensili()
// ============================================================

const SHEET_NAME_INPUT = 'Input';

const USCITE_CAT = [
  "Personali (vestiti, unghie, altro)",
  "Abbonamenti Vari",
  "Trasporti, Autostr, parcheg.",
  "Bollette",
  "Gasolio/benzina",
  "Manut. auto",
  "Mutuo",
  "Regali",
  "Ristorante/Ape/Merende",
  "Spesa alimentare",
  "Spese Mediche",
  "Spese pupino",
  "Spese Sport",
  "Varie generiche",
  "Viaggi, Vacanze",
  "Noleggio Auto",
  "Azioni per Vittoria FTSE ALL WORLD",
];

const ENTRATE_CAT = [
  "Stipendio Simo",
  "Varie Simo",
  "Stipendio Francy",
  "Fotovoltaico",
  "Regali",
  "Altri",
];

const MESI_NOMI = ['Gen','Feb','Mar','Apr','Mag','Giu','Lug','Ago','Set','Ott','Nov','Dic'];

function creaFogliMensili() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const anno = 2026;

  // Categorie esatte come arrivano dall'app
  const uscite = [
    "Personali (vestiti, unghie, altro)",
    "Abbonamenti Vari",
    "Trasporti, Autostr, parcheg.",
    "Bollette",
    "Gasolio/benzina",
    "Manut. auto",
    "Mutuo",
    "Regali",
    "Ristorante/Ape/Merende",
    "Spesa alimentare",
    "Spese Mediche",
    "Spese pupino",
    "Spese Sport",
    "Varie generiche",
    "Viaggi, Vacanze",
    "Noleggio Auto",
    "Azioni per Vittoria FTSE ALL WORLD",
  ];

  const entrate = [
    "Stipendio Simo",
    "Varie Simo",
    "Stipendio Francy",
    "Fotovoltaico",
    "Regali",
    "Altri",
  ];

  MESI_NOMI.forEach((nomeMese, idx) => {
    const meseNum = idx + 1;

    let ws = ss.getSheetByName(nomeMese);
    if (!ws) ws = ss.insertSheet(nomeMese);
    ws.clearContents();
    ws.clearFormats();

    // ── Riga 1: intestazioni affiancate ─────────────────────
    ws.getRange('A1').setValue('Uscite').setFontWeight('bold').setBackground('#f4cccc').setFontSize(12);
    ws.getRange('C1').setValue('Importo').setFontWeight('bold').setBackground('#f4cccc');
    ws.getRange('D1').setValue('Entrate').setFontWeight('bold').setBackground('#d9ead3').setFontSize(12);
    ws.getRange('F1').setValue('Importo').setFontWeight('bold').setBackground('#d9ead3');

    // ── Righe 2-18: categorie affiancate ────────────────────
    uscite.forEach((cat, i) => {
      const riga = 2 + i;
      ws.getRange(riga, 1).setValue(cat);
      const f = `=SUMPRODUCT((MONTH(Input!A$2:A$5000)=${meseNum})*(YEAR(Input!A$2:A$5000)=${anno})*(Input!D$2:D$5000="${cat}")*(Input!E$2:E$5000="Uscita")*Input!C$2:C$5000)`;
      ws.getRange(riga, 3).setFormula(f).setNumberFormat('€#,##0.00');
    });

    entrate.forEach((cat, i) => {
      const riga = 2 + i;
      ws.getRange(riga, 4).setValue(cat);
      const f = `=SUMPRODUCT((MONTH(Input!A$2:A$5000)=${meseNum})*(YEAR(Input!A$2:A$5000)=${anno})*(Input!D$2:D$5000="${cat}")*(Input!E$2:E$5000="Entrata")*Input!C$2:C$5000)`;
      ws.getRange(riga, 6).setFormula(f).setNumberFormat('€#,##0.00');
    });

    // ── Riga 19: totali ──────────────────────────────────────
    ws.getRange(19, 1).setValue('Totale Uscite').setFontWeight('bold');
    ws.getRange(19, 3).setFormula('=SUM(C2:C18)').setNumberFormat('€#,##0.00').setFontWeight('bold').setBackground('#f4cccc');
    ws.getRange(19, 4).setValue('Totale Entrate').setFontWeight('bold');
    ws.getRange(19, 6).setFormula('=SUM(F2:F18)').setNumberFormat('€#,##0.00').setFontWeight('bold').setBackground('#d9ead3');

    // ── Riga 21: netto mese ──────────────────────────────────
    ws.getRange(21, 1).setValue('Netto Mese').setFontWeight('bold').setFontSize(11);
    ws.getRange(21, 3).setFormula('=F19-C19').setNumberFormat('€#,##0.00').setFontWeight('bold').setFontSize(11);

    // ── Righe 23-24: spese extra libere ─────────────────────
    ws.getRange(23, 1).setValue('Spese Extra mese');
    ws.getRange(23, 3).setFormula(
      `=SUMPRODUCT((MONTH(Input!A$2:A$5000)=${meseNum})*(YEAR(Input!A$2:A$5000)=${anno})*(Input!E$2:E$5000="Uscita")*(Input!D$2:D$5000="Spese Extra")*Input!C$2:C$5000)`
    ).setNumberFormat('€#,##0.00');
    ws.getRange(23, 4).setValue('Entrate Extra mese');
    ws.getRange(23, 6).setFormula(
      `=SUMPRODUCT((MONTH(Input!A$2:A$5000)=${meseNum})*(YEAR(Input!A$2:A$5000)=${anno})*(Input!E$2:E$5000="Entrata")*(Input!D$2:D$5000="Entrate Extra")*Input!C$2:C$5000)`
    ).setNumberFormat('€#,##0.00');

    // ── Larghezze colonne ────────────────────────────────────
    ws.setColumnWidth(1, 260);
    ws.setColumnWidth(2, 10);
    ws.setColumnWidth(3, 110);
    ws.setColumnWidth(4, 200);
    ws.setColumnWidth(5, 10);
    ws.setColumnWidth(6, 110);

    Logger.log(`✓ ${nomeMese} creato`);
  });

  Logger.log('✅ 12 fogli mensili ricreati con la struttura corretta!');
}

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'write') {
    return scriviVoce(e);
  } else {
    return leggiStorico();
  }
}

function scriviVoce(e) {
  try {
    const payload = JSON.parse(decodeURIComponent(e.parameter.data));
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheet   = ss.getSheetByName(SHEET_NAME_INPUT) || ss.insertSheet(SHEET_NAME_INPUT);

    // Intestazioni se il foglio è vuoto
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Data', 'Descrizione', 'Importo', 'Categoria', 'Tipo']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }

    const parts   = payload.data.split('/');  // dd/mm/yyyy
    const dateObj = new Date(parts[2], parts[1]-1, parts[0]);

    sheet.appendRow([
      dateObj,
      payload.descrizione,
      parseFloat(payload.importo),
      payload.categoria,
      payload.tipo
    ]);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 1).setNumberFormat('dd/mm/yyyy');
    sheet.getRange(lastRow, 3).setNumberFormat('€#,##0.00');

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function leggiStorico() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME_INPUT);
    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService.createTextOutput('[]').setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getDataRange().getValues();
    const entries = rows.slice(1)
      .reverse()
      .slice(0, 30)
      .map(r => ({
        data:        r[0] ? Utilities.formatDate(new Date(r[0]), 'Europe/Rome', 'dd/MM/yyyy') : '',
        descrizione: r[1] || '',
        importo:     r[2] || 0,
        categoria:   r[3] || '',
        tipo:        r[4] || ''
      }));

    return ContentService
      .createTextOutput(JSON.stringify(entries))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput('[]')
      .setMimeType(ContentService.MimeType.JSON);
  }
}
