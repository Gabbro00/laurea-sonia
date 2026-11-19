/**
 * RSVP – Laurea di Sonia
 * Riceve le risposte dalla pagina dell'invito e le salva in questo foglio Google.
 * Crea da solo la scheda "Risposte" e la scheda "Riepilogo" alla prima risposta.
 * Se una persona modifica la sua risposta, la riga esistente viene aggiornata
 * (riconosciuta dal codice nella colonna F "ID").
 *
 * Pagina delle risposte per il telefono: aprire il link dello script
 * aggiungendo  ?vista=CHIAVE_VISTA  in fondo.
 */

const SCHEDA = 'Risposte';
const CHIAVE_VISTA = 'rose-bjk4iexca6';

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

// Senza chiave: controllo che il collegamento sia attivo.
// Con ?vista=CHIAVE_VISTA: pagina delle risposte leggibile dal telefono.
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.vista || p.vista !== CHIAVE_VISTA) return risposta_('Il modulo RSVP è attivo ✓');

  const tz = Session.getScriptTimeZone();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SCHEDA);
  const righe = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().forEach(r => {
      if (!String(r[1]).trim()) return;
      const d = r[0] instanceof Date ? r[0] : null;
      righe.push({
        quando: d ? d.getTime() : 0,
        data: d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy · HH:mm') : '',
        nome: String(r[1]),
        si: String(r[2]) === 'Sì',
        persone: Number(r[3]) || 0,
        messaggio: String(r[4] || '')
      });
    });
  }
  const ora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  const url = ScriptApp.getService().getUrl() + '?vista=' + encodeURIComponent(CHIAVE_VISTA);
  return HtmlService.createHtmlOutput(paginaRisposte_(righe, ora, url))
    .setTitle('Risposte · Laurea di Sonia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function paginaRisposte_(righe, ora, url) {
  const si = righe.filter(r => r.si).sort((a, b) => b.quando - a.quando);
  const no = righe.filter(r => !r.si).sort((a, b) => b.quando - a.quando);
  const persone = si.reduce((t, r) => t + r.persone, 0);

  const scheda = r => `
    <div class="r">
      <div class="top">
        <span class="nome">${esc_(r.nome)}</span>
        ${r.si
          ? `<span class="badge si">✓ ${r.persone} ${r.persone === 1 ? 'persona' : 'persone'}</span>`
          : `<span class="badge no">✗ non ci sarà</span>`}
      </div>
      ${r.data ? `<div class="data">${esc_(r.data)}</div>` : ''}
      ${r.messaggio ? `<p class="msg">“${esc_(r.messaggio)}”</p>` : ''}
    </div>`;

  const sezione = (titolo, lista, vuoto) => `
    <h2>${titolo} <span>${lista.length}</span></h2>
    ${lista.length ? lista.map(scheda).join('') : `<p class="vuoto">${vuoto}</p>`}`;

  return `<!DOCTYPE html><html lang="it"><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Pinyon+Script&family=Cormorant+Garamond:ital,wght@0,500;0,600;1,400&family=Montserrat:wght@500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Cormorant Garamond',Georgia,serif;color:#16110F;background:#F4E4E1;padding:22px 14px 40px}
.wrap{max-width:520px;margin:0 auto}
.head{text-align:center;margin-bottom:18px}
.k{font-family:Montserrat,sans-serif;font-size:10px;letter-spacing:.3em;text-transform:uppercase;color:#8F6B27}
h1{font-family:'Pinyon Script',cursive;font-weight:400;font-size:46px;line-height:1.1;margin-top:4px}
.stats{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:8px}
.st{background:#16110F;color:#fff;border-radius:6px;padding:14px 10px 12px;text-align:center;box-shadow:inset 0 0 0 1px rgba(198,160,82,.7)}
.st b{display:block;font-size:34px;font-weight:600;line-height:1}
.st span{font-family:Montserrat,sans-serif;font-size:9px;letter-spacing:.18em;text-transform:uppercase;color:#C6A052}
.st.light{background:#FFFCF7;color:#16110F}
h2{font-family:Montserrat,sans-serif;font-size:11px;letter-spacing:.25em;text-transform:uppercase;color:#741322;margin:26px 2px 10px;display:flex;align-items:center;gap:8px}
h2 span{background:#741322;color:#fff;border-radius:20px;padding:2px 9px;font-size:10px;letter-spacing:.05em}
.r{background:#FFFCF7;border:1px solid rgba(198,160,82,.45);border-radius:6px;padding:14px 16px;margin-bottom:10px;box-shadow:0 4px 12px rgba(110,40,50,.08)}
.top{display:flex;justify-content:space-between;align-items:center;gap:10px}
.nome{font-size:22px;font-weight:600;line-height:1.2}
.badge{flex-shrink:0;font-family:Montserrat,sans-serif;font-size:10.5px;font-weight:600;letter-spacing:.06em;border-radius:20px;padding:5px 10px;white-space:nowrap}
.badge.si{background:#16110F;color:#E7CC88}
.badge.no{background:#F3DCDB;color:#741322}
.data{font-family:Montserrat,sans-serif;font-size:10.5px;color:#8b7b77;margin-top:4px}
.msg{font-style:italic;font-size:19px;line-height:1.45;color:#3a302d;margin-top:10px;padding-left:12px;border-left:2px solid #C6A052}
.vuoto{font-style:italic;font-size:18px;color:#8b7b77;padding:4px 2px}
.foot{text-align:center;margin-top:28px}
.agg{display:inline-block;font-family:Montserrat,sans-serif;font-size:11px;font-weight:600;letter-spacing:.2em;text-transform:uppercase;color:#E7CC88;background:#16110F;border-radius:30px;padding:13px 26px;text-decoration:none}
.ora{font-family:Montserrat,sans-serif;font-size:10px;color:#8b7b77;margin-top:10px}
</style></head><body><div class="wrap">
  <div class="head"><p class="k">Laurea di Sonia</p><h1>Le risposte</h1></div>
  <div class="stats">
    <div class="st"><b>${si.length}</b><span>Ci saranno</span></div>
    <div class="st"><b>${persone}</b><span>Persone in arrivo</span></div>
    <div class="st light"><b>${no.length}</b><span>Non ci saranno</span></div>
    <div class="st light"><b>${righe.length}</b><span>Risposte totali</span></div>
  </div>
  ${sezione('Ci saranno', si, 'Ancora nessuna conferma.')}
  ${sezione('Non ci saranno', no, 'Nessuno, per ora ♡')}
  <div class="foot">
    <a class="agg" href="${esc_(url)}" target="_top">↻ Aggiorna</a>
    <p class="ora">Aggiornato il ${esc_(ora)}</p>
  </div>
</div></body></html>`;
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
