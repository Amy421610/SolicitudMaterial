// ═══════════════════════════════════════════════════════════
// OINK — Google Apps Script v2
// Reemplaza todo el código anterior con este y vuelve a desplegar
// ═══════════════════════════════════════════════════════════

const SHEET_ID = 'TU_SHEET_ID_AQUI'; // ← el ID de tu Google Sheet

// Headers CORS — necesarios para que funcione desde GitHub Pages
function setCORSHeaders(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function doOptions(e) {
  return setCORSHeaders(ContentService.createTextOutput(''));
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if (data.action === 'guardar')    result = guardarSolicitud(data.solicitud);
    else if (data.action === 'actualizar') result = actualizarSolicitud(data.solicitud);
    else if (data.action === 'borrar') result = borrarSolicitud(data.folio);
    else if (data.action === 'borrarTodas') result = borrarTodas();
    else result = { ok: false, error: 'Acción no reconocida' };
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON)
    );
  } catch (err) {
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON)
    );
  }
}

function doGet(e) {
  try {
    const result = cargarSolicitudes();
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON)
    );
  } catch (err) {
    return setCORSHeaders(
      ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON)
    );
  }
}

function guardarSolicitud(sol) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = getOCreateHoja(ss, sol.fechaCreacion);

  // Fila de encabezado de la solicitud (resaltada)
  const filaInicio = hoja.getLastRow() + 1;
  hoja.appendRow([
    '▶ ' + sol.folio,
    sol.fechaCreacion,
    sol.fechaMod || sol.fechaCreacion,
    sol.fechaModParque || sol.fechaCreacion,
    sol.status,
    sol.obs || '',
    sol.notaGlobal || '',
    sol.items.length + ' artículo(s)','','','','','','','','','','','',''
  ]);
  const filaHeader = hoja.getRange(filaInicio, 1, 1, 20);
  filaHeader.setBackground('#fde8e2').setFontWeight('bold');
  hoja.getRange(filaInicio, 1).setFontColor('#ee4723');

  // Filas de artículos
  sol.items.forEach(it => {
    hoja.appendRow([
      sol.folio,
      sol.fechaCreacion,
      sol.fechaMod || sol.fechaCreacion,
      sol.fechaModParque || sol.fechaCreacion,
      sol.status,
      sol.obs || '',
      sol.notaGlobal || '',
      it.nombre,
      it.zona,
      it.mat,
      it.med || '',
      it.qty,
      it.motivo,
      it.obs || '',
      it.tipo,
      it.status,
      it.nota || '',
      it.checks ? (it.checks.illustrator ? 'Sí' : 'No') : 'No',
      it.checks ? (it.checks.imprenta   ? 'Sí' : 'No') : 'No',
      it.checks ? (it.checks.enviado    ? 'Sí' : 'No') : 'No',
    ]);
  });

  // Fila vacía de separación
  hoja.appendRow(['']);

  formatearHoja(hoja);
  return { ok: true, folio: sol.folio };
}

function actualizarSolicitud(sol) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    for (let i = datos.length - 1; i >= 1; i--) {
      const folioCelda = String(datos[i][0]).replace('▶ ', '');
      if (folioCelda === sol.folio) hoja.deleteRow(i + 1);
    }
  });
  return guardarSolicitud(sol);
}

function cargarSolicitudes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const corte = new Date();
  corte.setMonth(corte.getMonth() - 6);
  const solMap = {};
  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;
    datos.slice(1).forEach(row => {
      let folio = row[0];
      if (!folio) return; // fila vacía de separación
      const esHeader = String(folio).indexOf('▶') === 0;
      folio = String(folio).replace('▶ ', '');

      if (!solMap[folio]) {
        solMap[folio] = {
          folio, fechaCreacion: String(row[1]), fechaMod: String(row[2]),
          fechaModParque: String(row[3]), status: row[4],
          obs: row[5], notaGlobal: row[6], items: []
        };
      }
      if (esHeader) return; // la fila de encabezado no es un artículo

      solMap[folio].items.push({
        nombre: row[7], zona: row[8], mat: row[9], med: row[10],
        qty: row[11], motivo: row[12], obs: row[13], tipo: row[14],
        status: row[15], nota: row[16], imgData: '',
        checks: { illustrator: row[17]==='Sí', imprenta: row[18]==='Sí', enviado: row[19]==='Sí' }
      });
    });
  });
  const solicitudes = Object.values(solMap).filter(s => new Date(s.fechaCreacion) > corte);
  return { ok: true, solicitudes };
}

function getOCreateHoja(ss, fechaISO) {
  const fecha = new Date(fechaISO);
  const nombre = Utilities.formatDate(fecha, 'America/Mexico_City', 'yyyy-MM');
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    hoja.appendRow([
      'Folio','Fecha Creación','Última Mod.','Última Mod. Parque','Estado','Obs. General','Nota Diseñadora',
      'Artículo','Zona','Material','Medidas','Cantidad','Motivo',
      'Obs. Artículo','Tipo','Estado Artículo','Nota Interna',
      'Illustrator','Imprenta','Enviado'
    ]);
    hoja.getRange(1,1,1,20).setBackground('#ee4723').setFontColor('white').setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function formatearHoja(hoja) {
  try { hoja.autoResizeColumns(1, 20); } catch(e) {}
}

// ── Borrar solicitud ──
function borrarSolicitud(folio) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    for (let i = datos.length - 1; i >= 1; i--) {
      const folioCelda = String(datos[i][0]).replace('▶ ', '');
      if (folioCelda === folio) hoja.deleteRow(i + 1);
    }
  });
  return { ok: true, folio };
}

// ── Borrar todas (solo para limpiar pruebas) ──
function borrarTodas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    for (let i = datos.length - 1; i >= 1; i--) {
      hoja.deleteRow(i + 1);
    }
  });
  return { ok: true };
}
