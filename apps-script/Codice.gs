/**
 * RSVP – Laurea di Sonia
 * Riceve le risposte dalla pagina dell'invito e le salva in questo foglio Google.
 * Crea da solo la scheda "Risposte" e la scheda "Riepilogo" alla prima risposta.
 * Se una persona modifica la sua risposta, la riga esistente viene aggiornata
 * (riconosciuta dal codice nella colonna F "ID").
 */

const SCHEDA = 'Risposte';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return risposta_('ok');            // anti-spam (campo nascosto)

    const nome = String(p.nome || '').trim().slice(0, 100);
    if (!nome) return risposta_('errore: nome mancante');

    const si = p.partecipa === 'si';
    const persone = si ? Math.max(1, Math.min(15, Number(p.persone) || 1)) : 0;
    const messaggio = String(p.messaggio || '').trim().slice(0, 1000);

    const id = String(p.id || '').trim().slice(0, 60);
    const riga = [new Date(), nome, si ? 'Sì' : 'No', persone, messaggio, id];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = schedaRisposte_(ss);

    let aggiornata = false;
    const ultima = sh.getLastRow();
    if (id && ultima > 1) {
      const ids = sh.getRange(2, 6, ultima - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === id) {
          sh.getRange(i + 2, 1, 1, riga.length).setValues([riga]);
          aggiornata = true;
          break;
        }
      }
    }
    if (!aggiornata) sh.appendRow(riga);
    schedaRiepilogo_(ss);

    return risposta_('ok');
  } finally {
    lock.releaseLock();
  }
}

// Aprendo il link nel browser si vede se il collegamento è attivo
function doGet() {
  return risposta_('Il modulo RSVP è attivo ✓');
}

function schedaRisposte_(ss) {
  let sh = ss.getSheetByName(SCHEDA);
  if (!sh) {
    sh = ss.insertSheet(SCHEDA, 0);
    sh.appendRow(['Data risposta', 'Nome', 'Partecipa', 'Persone', 'Messaggio']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold').setBackground('#F4E4E1');
    sh.getRange('A:A').setNumberFormat('dd/MM/yyyy HH:mm');
    sh.setColumnWidth(1, 140);
    sh.setColumnWidth(2, 200);
    sh.setColumnWidth(5, 380);
  }
  if (sh.getRange(1, 6).getValue() !== 'ID') {
    sh.getRange(1, 6).setValue('ID').setFontWeight('bold').setBackground('#F4E4E1');
    sh.getRange('F:F').setFontColor('#999999');
  }
  return sh;
}

function schedaRiepilogo_(ss) {
  if (ss.getSheetByName('Riepilogo')) return;
  const r = ss.insertSheet('Riepilogo');
  r.getRange('A1:A4').setValues([
    ['Risposte ricevute'],
    ['Ci saranno (risposte Sì)'],
    ['Persone in arrivo'],
    ['Non ci saranno'],
  ]);
  // foglio in italiano: separatore ';'
  r.getRange('B1:B4').setFormulas([
    ['=COUNTA(Risposte!B2:B)'],
    ['=COUNTIF(Risposte!C2:C;"Sì")'],
    ['=SUMIF(Risposte!C2:C;"Sì";Risposte!D2:D)'],
    ['=COUNTIF(Risposte!C2:C;"No")'],
  ]);
  r.getRange('A1:A4').setFontWeight('bold');
  r.setColumnWidth(1, 220);
}

function risposta_(testo) {
  return ContentService.createTextOutput(testo).setMimeType(ContentService.MimeType.TEXT);
}
