// ══════════════════════════════════════════════════════════════════
// TT Audit — Plataforma BAT (Facturación y Gastos) — Apps Script v1
// ──────────────────────────────────────────────────────────────────
// Pasos para desplegar:
//   1. Crea un Google Sheet NUEVO (ej. "BAT - Facturación y Gastos")
//   2. Extensiones → Apps Script → pega TODO este código
//   3. Ejecuta initSheets() una vez (crea pestañas + catálogos)
//   4. Importa bat/data/facturacion_historico.csv en la pestaña "Facturacion"
//   5. Implementar → Nueva implementación
//        · Tipo: Aplicación web   · Ejecutar: Yo   · Acceso: Cualquiera
//   6. Copia la URL y pégala en la plataforma → Configuración
// ══════════════════════════════════════════════════════════════════

// ── Configuración ─────────────────────────────────────────────────
const SHEET_FACT     = 'Facturacion';
const SHEET_GASTOS   = 'Gastos';
const SHEET_MAESTROS = 'Maestros';

const HDR_FACT = ['ID','Cliente','Ejecutivo','Tipo de SS','Responsable de Pago',
  'Servicio / Proyecto','Mes','Importe','OS','Serie Factura','# Factura',
  'Fecha Factura','Fecha Vencimiento','Estado','Estado Detalle'];
const HDR_GASTOS   = ['ID','Cliente','Tipo de SS','Mes','Grupo','Categoria','Monto','Detalle'];
const HDR_MAESTROS = ['Tipo','Valor'];

// ── Spreadsheet (activo) ───────────────────────────────────────────
function getSS() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getOrCreate(ss, name, headers) {
  if (!ss) ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1d3b78').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Inicializar (ejecutar una sola vez) ────────────────────────────
function initSheets() {
  const ss = getSS();
  getOrCreate(ss, SHEET_FACT,     HDR_FACT);
  getOrCreate(ss, SHEET_GASTOS,   HDR_GASTOS);
  const shM = getOrCreate(ss, SHEET_MAESTROS, HDR_MAESTROS);

  // Sembrar catálogos solo si está vacío
  if (shM.getLastRow() <= 1) {
    const seed = [];
    const add = (t, arr) => arr.forEach(v => seed.push([t, v]));
    add('Cliente',   ['BAT','PALMERA','DP']);
    add('Servicio',  ['CONTRATO - FEE + PERSONAL','ENVIOS','ALMACENAJE','TRANSPORTE',
                      'ADMINISTRATIVOS','OTROS','PROMOTORIA','MERCH']);
    add('Estado',    ['Pendiente','Programado','Factura enviada','Facturado','Pagado']);
    add('Responsable', ['SOFIA MAGNI','NICOLE AGUIRRE','RICARDO','SEBASTIAN','DAVID BODERO',
                        'ANA SOFIA MOY','DIEGO LAZO','ARACELI','JACINTA','ANDREA ANDRADE',
                        'CHRIS BALVIN','FERNANDO SARRIA']);
    add('CatPersonas', ['Planilla Oficina','Planilla Campo']);
    add('CatOtros',    ['Combustible','Peajes','Envíos','Transporte','Almacenaje','Otros']);
    shM.getRange(2, 1, seed.length, 2).setValues(seed);
  }
  Logger.log('initSheets OK');
  return 'OK — pestañas y catálogos creados';
}

// ════════════════════════════════════════════════════════════════════
//  GET
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  try { return handleGet(e); }
  catch (err) { return jsonResp({ error: err.message }); }
}

function handleGet(e) {
  const ss = getSS();
  const p  = (e && e.parameter) ? e.parameter : {};
  const accion = p.accion || '';

  if (!accion) return jsonResp({ status: 'TT Audit BAT API v1 activa', sheet: ss.getName() });

  if (accion === 'getFacturacion') {
    const sh = ss.getSheetByName(SHEET_FACT);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[1]).map(r => ({
      id:        s(r[0]),
      cliente:   s(r[1]),
      ejecutivo: s(r[2]),
      tipo:      s(r[3]),
      responsable: s(r[4]),
      servicio:  s(r[5]),
      mes:       d(r[6]),
      importe:   n(r[7]),
      os:        s(r[8]),
      serie:     s(r[9]),
      factura:   s(r[10]),
      fechaFactura:  d(r[11]),
      fechaVenc:     d(r[12]),
      estado:    s(r[13]),
      estadoDetalle: s(r[14]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getGastos') {
    const sh = ss.getSheetByName(SHEET_GASTOS);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]).map(r => ({
      id:        s(r[0]),
      cliente:   s(r[1]),
      tipo:      s(r[2]),
      mes:       d(r[3]),
      grupo:     s(r[4]),
      categoria: s(r[5]),
      monto:     n(r[6]),
      detalle:   s(r[7]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getMaestros') {
    const sh = ss.getSheetByName(SHEET_MAESTROS);
    const out = {};
    if (sh) {
      sh.getDataRange().getValues().slice(1).forEach(r => {
        const t = s(r[0]), v = s(r[1]);
        if (!t || !v) return;
        (out[t] = out[t] || []).push(v);
      });
    }
    return jsonResp({ maestros: out });
  }

  return jsonResp({ status: 'BAT API v1', sheet: ss.getName() });
}

// ════════════════════════════════════════════════════════════════════
//  POST
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  try { return handlePost(e); }
  catch (err) { return jsonResp({ error: err.message }); }
}

function handlePost(e) {
  const ss = getSS();
  let body;
  try { body = JSON.parse(e && e.postData ? e.postData.contents : '{}'); }
  catch (err) { return jsonResp({ error: 'JSON inválido: ' + err.message }); }
  const accion = body.accion || '';

  // ── FACTURACIÓN ────────────────────────────────────────────────
  if (accion === 'saveFacturacion') {
    const sh = getOrCreate(ss, SHEET_FACT, HDR_FACT);
    const b = body;
    if (!b.cliente) return jsonResp({ error: 'Cliente requerido' });
    const id = b.id || ('F' + new Date().getTime());
    const row = [id, b.cliente||'', b.ejecutivo||'Claudia Camarena', b.tipo||'', b.responsable||'',
      b.servicio||'', b.mes||'', num(b.importe), b.os||'', b.serie||'', b.factura||'',
      b.fechaFactura||'', b.fechaVenc||'', b.estado||'Pendiente', b.estadoDetalle||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_FACT.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }

  if (accion === 'deleteFacturacion') {
    const sh = ss.getSheetByName(SHEET_FACT);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  // ── GASTOS ─────────────────────────────────────────────────────
  if (accion === 'saveGasto') {
    const sh = getOrCreate(ss, SHEET_GASTOS, HDR_GASTOS);
    const b = body;
    if (!b.cliente || !b.tipo || !b.mes) return jsonResp({ error: 'Cliente, servicio y mes requeridos' });
    const id = b.id || ('G' + new Date().getTime());
    const row = [id, b.cliente||'', b.tipo||'', b.mes||'', b.grupo||'Otros',
      b.categoria||'', num(b.monto), b.detalle||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_GASTOS.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }

  if (accion === 'deleteGasto') {
    const sh = ss.getSheetByName(SHEET_GASTOS);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  // ── MAESTROS ───────────────────────────────────────────────────
  if (accion === 'saveMaestro') {
    const sh = getOrCreate(ss, SHEET_MAESTROS, HDR_MAESTROS);
    const { tipo, valor } = body;
    if (!tipo || !valor) return jsonResp({ error: 'tipo y valor requeridos' });
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (s(data[i][0]) === tipo && s(data[i][1]) === valor)
        return jsonResp({ ok: true, action: 'exists' });
    }
    sh.appendRow([tipo, valor]);
    return jsonResp({ ok: true, action: 'created' });
  }

  if (accion === 'deleteMaestro') {
    const sh = ss.getSheetByName(SHEET_MAESTROS);
    const { tipo, valor } = body;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (s(data[i][0]) === tipo && s(data[i][1]) === valor) {
        sh.deleteRow(i + 1); return jsonResp({ ok: true, action: 'deleted' });
      }
    }
    return jsonResp({ ok: false, error: 'No encontrado' });
  }

  return jsonResp({ error: 'Acción no reconocida: ' + accion });
}

// ════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════
function findRow(sh, id) {
  if (!sh || !id) return -1;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (s(data[i][0]) === s(id)) return i + 1;
  }
  return -1;
}

function s(v)   { return v === null || v === undefined ? '' : v.toString().trim(); }
function n(v)   { const x = parseFloat(v); return isNaN(x) ? 0 : x; }
function num(v) { const x = parseFloat(v); return isNaN(x) ? 0 : x; }

// Fecha -> 'YYYY-MM-DD' (acepta Date o texto)
function d(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
  }
  return s(v);
}

function pad(x) { return String(x).padStart(2, '0'); }

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
