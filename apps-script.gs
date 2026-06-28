/****************************************************************
 * POLLA MUNDIALISTA 26 — Backend FASE 2 (Google Apps Script)
 * Pegar en: Extensiones > Apps Script del MISMO Excel.
 * Desplegar: Implementar > Nueva implementación > Aplicación web
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier usuario
 * https://script.google.com/macros/s/AKfycbygP_haloKqpQA4YaR8lC2I4XUu6pt-GBnhcRhVgq8zCRCvqlw31KnA-waou2Doo2AW/exec
 *
 * SEGURIDAD: solo crea/escribe la hoja "Respuestas_F2".
 * NO toca tus pestañas de scoring, bracket ni leaderboard.
 ****************************************************************/

var HOJA = "Respuestas_F2";
var CABECERAS = [
  "timestamp","nombre","contacto","comodin",
  "16a_1","16a_2","16a_3","16a_4","16a_5","16a_6","16a_7","16a_8",
  "16a_9","16a_10","16a_11","16a_12","16a_13","16a_14","16a_15","16a_16",
  "8vos_1","8vos_2","8vos_3","8vos_4","8vos_5","8vos_6","8vos_7","8vos_8",
  "4tos_1","4tos_2","4tos_3","4tos_4",
  "semi_1","semi_2","campeon","subcampeon","tercero","cuarto"
];

function getHoja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.appendRow(CABECERAS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita choques cuando varios envían a la vez
  try {
    var d = JSON.parse(e.postData.contents);
    var sh = getHoja_();
    var fila = [
      d.ts || new Date().toISOString(), d.nombre || "", d.contacto || "", d.comodin || ""
    ]
    .concat(pad_(d.r16avos, 16))
    .concat(pad_(d.octavos, 8))
    .concat(pad_(d.cuartos, 4))
    .concat(pad_(d.semis, 2))
    .concat([d.campeon || "", d.subcampeon || "", d.tercero || "", d.cuarto || ""]);

    // UPSERT por nombre (col 2): si ya envió, sobrescribe su fila.
    var rango = sh.getDataRange().getValues();
    var idx = -1;
    for (var i = 1; i < rango.length; i++) {
      if (String(rango[i][1]).trim().toLowerCase() === String(d.nombre).trim().toLowerCase()) { idx = i + 1; break; }
    }
    if (idx === -1) sh.appendRow(fila);
    else sh.getRange(idx, 1, 1, fila.length).setValues([fila]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function pad_(arr, n) {
  arr = arr || [];
  var out = [];
  for (var i = 0; i < n; i++) out.push(arr[i] || "");
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/****************************************************************
 * doGet — STUB para la futura tabla en vivo (Fase 2).
 * Cuando construyamos el leaderboard, este endpoint devolverá
 * el ranking ya calculado (leído de tu hoja de scoring).
 ****************************************************************/
function doGet(e) {
  return json_({ ok: true, msg: "Endpoint listo. Leaderboard se conecta después." });
}
