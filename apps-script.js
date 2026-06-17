// ═══════════════════════════════════════════════════════════
// OINK — Google Apps Script
// Pega este código completo en tu Apps Script y despliégalo
// ═══════════════════════════════════════════════════════════

const SHEET_ID = 'TU_SHEET_ID_AQUI'; // ← reemplaza con el ID de tu Google Sheet

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'guardar') return guardarSolicitud(data.solicitud);
    if (action === 'actualizar') return actualizarSolicitud(data.solicitud);
    if (action === 'cargar') return cargarSolicitudes();

    return resp({ ok: false, error: 'Acción no reconocida' });
  } catch (err) {
    return resp({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  return cargarSolicitudes();
}

// ── Guardar nueva solicitud ──
function guardarSolicitud(sol) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hoja = getOCreateHoja(ss, sol.fechaCreacion);

  // Una fila por artículo
  sol.items.forEach(it => {
    hoja.appendRow([
      sol.folio,
      sol.fechaCreacion,
      sol.fechaMod,
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
      it.checks.illustrator ? 'Sí' : 'No',
      it.checks.imprenta   ? 'Sí' : 'No',
      it.checks.enviado    ? 'Sí' : 'No',
      it.imgData ? '(imagen)' : ''
    ]);
  });

  formatearHoja(hoja);
  return resp({ ok: true, folio: sol.folio });
}

// ── Actualizar solicitud existente ──
function actualizarSolicitud(sol) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  // Buscar en todas las hojas
  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    // Borrar filas del folio
    for (let i = datos.length - 1; i >= 1; i--) {
      if (datos[i][0] === sol.folio) hoja.deleteRow(i + 1);
    }
  });
  // Re-guardar actualizado
  return guardarSolicitud(sol);
}

// ── Cargar todas las solicitudes (últimos 6 meses) ──
function cargarSolicitudes() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const corte = new Date();
  corte.setMonth(corte.getMonth() - 6);

  const solMap = {};

  ss.getSheets().forEach(hoja => {
    const datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;

    datos.slice(1).forEach(row => {
      const folio = row[0];
      if (!folio) return;

      if (!solMap[folio]) {
        solMap[folio] = {
          folio:         row[0],
          fechaCreacion: row[1],
          fechaMod:      row[2],
          status:        row[3],
          obs:           row[4],
          notaGlobal:    row[5],
          items: []
        };
      }

      solMap[folio].items.push({
        nombre:   row[6],
        zona:     row[7],
        mat:      row[8],
        med:      row[9],
        qty:      row[10],
        motivo:   row[11],
        obs:      row[12],
        tipo:     row[13],
        status:   row[14],
        nota:     row[15],
        imgData:  '',
        checks: {
          illustrator: row[16] === 'Sí',
          imprenta:    row[17] === 'Sí',
          enviado:     row[18] === 'Sí'
        }
      });
    });
  });

  const solicitudes = Object.values(solMap).filter(s => {
    return new Date(s.fechaCreacion) > corte;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, solicitudes }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Obtener o crear hoja del mes ──
function getOCreateHoja(ss, fechaISO) {
  const fecha = new Date(fechaISO);
  const nombre = Utilities.formatDate(fecha, 'America/Mexico_City', 'yyyy-MM') +
    ' — ' +
    Utilities.formatDate(fecha, 'America/Mexico_City', 'MMMM yyyy');

  let hoja = ss.getSheetByName(nombre);
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    // Encabezados
    hoja.appendRow([
      'Folio', 'Fecha Creación', 'Última Modificación', 'Estado', 'Obs. General', 'Nota Diseñadora',
      'Artículo', 'Zona', 'Material', 'Medidas', 'Cantidad', 'Motivo',
      'Obs. Artículo', 'Tipo', 'Estado Artículo', 'Nota Interna',
      'Illustrator', 'Imprenta', 'Enviado', 'Imagen'
    ]);
    hoja.getRange(1, 1, 1, 20).setBackground('#ee4723').setFontColor('white').setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// ── Formatear hoja ──
function formatearHoja(hoja) {
  hoja.autoResizeColumns(1, 20);
}

// ── Helper respuesta ──
function resp(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
