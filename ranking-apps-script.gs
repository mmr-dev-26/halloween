/**
 * Ranking da Festa de Halloween (Quiosque 3B)
 * Cole este código em: Planilha Google > Extensões > Apps Script
 * Depois: Implantar > Nova implantação > Tipo "App da Web"
 *   Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * Copie a URL que termina em /exec e mande pro Claude (ou cole em RANK_URL no index.html).
 */
var ABA = 'Rank';
var MONSTROS = ['V', 'L', 'F', 'B', 'Z', 'D'];

function aba_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.appendRow(['nome', 'tempo_segundos', 'monstro', 'data']);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var valores = aba_().getDataRange().getValues().slice(1);
  var rows = valores.map(function (r) {
    return { name: String(r[0]), time: Number(r[1]), monster: String(r[2]), at: r[3] ? new Date(r[3]).getTime() : 0 };
  }).filter(function (r) { return r.name && isFinite(r.time) && MONSTROS.indexOf(r.monster) >= 0; });
  return json_({ rows: rows });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents || '{}');
    var nome = String(d.name || '').replace(/[<>]/g, '').trim().slice(0, 20);
    var tempo = Number(d.time);
    var monstro = String(d.monster || '');
    if (!nome || !isFinite(tempo) || tempo < 20 || tempo > 7200 || MONSTROS.indexOf(monstro) < 0) {
      return json_({ ok: false, error: 'dados inválidos' });
    }
    aba_().appendRow([nome, Math.round(tempo * 10) / 10, monstro, new Date()]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
