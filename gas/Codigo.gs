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
const SHEET_CUOTAS   = 'Cuotas';
const SHEET_MERCH    = 'Merch';
const SHEET_PROY     = 'Proyectos';
const SHEET_PKPI     = 'PromoKPI';
const SHEET_PMAT     = 'PromoMateriales';
const SHEET_PINC     = 'PromoIncidencias';
const SHEET_COMPRA   = 'ComprasMateriales';
const SHEET_PPTO     = 'PptoAprobado';

const HDR_FACT = ['ID','Cliente','Ejecutivo','Tipo de SS','Responsable de Pago',
  'Servicio / Proyecto','Mes','Importe','OS','Serie Factura','# Factura',
  'Fecha Factura','Fecha Vencimiento','Estado','Estado Detalle'];
const HDR_GASTOS   = ['ID','Cliente','Tipo de Servicios/Proyecto','Concepto','Mes','Grupo','Categoria','Monto'];
const HDR_MAESTROS = ['Tipo','Valor'];
const HDR_CUOTAS   = ['Ejecutivo','Anio','Meta Q1','Meta Q2','Meta Q3','Meta Q4'];
const HDR_MERCH    = ['ID','Semana','Tipo','Zona','Cadena','Programado','Por Programar','Efectivo','Obs'];
const HDR_PROY     = ['ID','Nombre','Inicio','Fin','Dias','Personal','Ciudades','Cuentas KA','Efectividad','Cobertura','Avances','Estado','Obs','Responsable'];
const HDR_PKPI     = ['Cadena','Cobertura S1','Cobertura S2','Efectividad S1','Efectividad S2','Conectados','Conectados Nota'];
const HDR_PMAT     = ['ID','Cadena','Categoria','Material','Tiendas','Avance S1','Avance S2','Obs','Inicio','Fin'];
const HDR_PINC     = ['ID','Cadena','Fecha','Incidencia','Responsable','Estado'];
const HDR_COMPRA   = ['ID','PPTO','Material','Cantidad','Precio Unitario','Total','Fecha Compra','Cuenta','Obs'];
const HDR_PPTO     = ['ID','Nombre','Mes Aprobado','Monto Aprobado','Correo','Obs'];

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
  getOrCreate(ss, SHEET_MERCH,    HDR_MERCH);
  getOrCreate(ss, SHEET_PROY,     HDR_PROY);
  getOrCreate(ss, SHEET_PKPI,     HDR_PKPI);
  getOrCreate(ss, SHEET_PMAT,     HDR_PMAT);
  getOrCreate(ss, SHEET_PINC,     HDR_PINC);
  getOrCreate(ss, SHEET_COMPRA,   HDR_COMPRA);
  getOrCreate(ss, SHEET_PPTO,     HDR_PPTO);
  getOrCreate(ss, SHEET_CUOTAS,   HDR_CUOTAS);
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
    add('CatOtros',    ['Peaje','Cochera','Envíos','Embalaje','Combustible','Transporte','Almacenaje','Otros']);
    shM.getRange(2, 1, seed.length, 2).setValues(seed);
  }
  Logger.log('initSheets OK');
  return 'OK — pestañas y catálogos creados';
}

// ── Agregar la columna "Correo" a PptoAprobado si falta (ejecutar 1 vez) ──
function agregarColumnaCorreo() {
  const ss = getSS();
  let sh = ss.getSheetByName(SHEET_PPTO);
  if (!sh) { getOrCreate(ss, SHEET_PPTO, HDR_PPTO); return 'Pestaña PptoAprobado creada con la columna Correo.'; }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const find = colFinder(headers);
  if (find(['Correo']) >= 0) return 'La columna Correo ya existe.';
  const obsCol = find(['Obs', 'Observacion', 'Observación']);   // 0-based
  let at;
  if (obsCol >= 0) { at = obsCol + 1; sh.insertColumnBefore(at); }   // antes de Obs
  else { sh.insertColumnAfter(lastCol); at = lastCol + 1; }          // al final
  sh.getRange(1, at).setValue('Correo').setFontWeight('bold')
    .setBackground('#1d3b78').setFontColor('#ffffff');
  Logger.log('Columna Correo insertada en posición ' + at);
  return 'OK — columna Correo agregada.';
}

// ── Asignar IDs a todas las filas de Facturacion (ejecutar 1 vez) ──
// Crea la columna ID al inicio si no existe y rellena las vacías.
function asignarIDsFacturacion() {
  const sh = getSS().getSheetByName(SHEET_FACT);
  if (!sh) { Logger.log('No existe la pestaña ' + SHEET_FACT); return 'No existe la pestaña ' + SHEET_FACT; }
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return 'Sin datos que numerar';

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let idCol = colFinder(headers)(['ID']);   // 0-based, -1 si no existe
  if (idCol < 0) {
    sh.insertColumnBefore(1);
    sh.getRange(1, 1).setValue('ID').setFontWeight('bold')
      .setBackground('#1d3b78').setFontColor('#ffffff');
    idCol = 0;
  }
  const col = idCol + 1;                     // 1-based
  const nData = lastRow - 1;
  const ids = sh.getRange(2, col, nData, 1).getValues();

  const existing = {};
  ids.forEach(function (r) { if (r[0] !== '' && r[0] != null) existing[s(r[0])] = true; });

  let seq = 0, count = 0;
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === '' || ids[i][0] == null) {
      let cand;
      do { seq++; cand = 'F' + ('0000' + seq).slice(-4); } while (existing[cand]);
      ids[i][0] = cand; existing[cand] = true; count++;
    }
  }
  sh.getRange(2, col, nData, 1).setValues(ids);
  Logger.log('IDs asignados: ' + count + ' (columna ' + col + ')');
  return 'OK — ' + count + ' IDs asignados';
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
    if (data.length < 2) return jsonResp({ rows: [], count: 0 });
    // Mapea por NOMBRE de encabezado (tolerante a orden, mayúsculas y acentos)
    const find = colFinder(data[0]);
    const ci = {
      id:            find(['ID']),
      cliente:       find(['Cliente']),
      ejecutivo:     find(['Ejecutivo']),
      tipo:          find(['Tipo de SS','Tipo SS','Tipo']),
      responsable:   find(['Responsable de Pago','Responsable']),
      servicio:      find(['Servicio / Proyecto','Servicio Proyecto','Servicio','Proyecto']),
      mes:           find(['Mes']),
      importe:       find(['Importe','Monto']),
      os:            find(['OS']),
      serie:         find(['Serie Factura','Serie']),
      factura:       find(['# Factura','N Factura','Nro Factura','Numero Factura','Factura']),
      fechaFactura:  find(['Fecha de Factura','Fecha Factura']),
      fechaVenc:     find(['Fecha de Vencimiento','Fecha Vencimiento','Vencimiento']),
      estado:        find(['Estado (Pagado)','Estado Pagado','Estado']),
      estadoDetalle: find(['Estado Detalle','Detalle Estado','Detalle']),
    };
    const g = (r, k) => ci[k] >= 0 ? r[ci[k]] : '';
    const cCli = ci.cliente >= 0 ? ci.cliente : 0;
    const rows = [];
    for (var i = 1; i < data.length; i++) {
      const r = data[i];
      if (!s(r[cCli])) continue;               // salta filas sin cliente
      rows.push({
        id:            (ci.id >= 0 && s(r[ci.id])) ? s(r[ci.id]) : ('F' + i),
        cliente:       s(g(r, 'cliente')),
        ejecutivo:     s(g(r, 'ejecutivo')),
        tipo:          s(g(r, 'tipo')),
        responsable:   s(g(r, 'responsable')),
        servicio:      s(g(r, 'servicio')),
        mes:           d(g(r, 'mes')),
        importe:       n(g(r, 'importe')),
        os:            s(g(r, 'os')),
        serie:         s(g(r, 'serie')),
        factura:       s(g(r, 'factura')),
        fechaFactura:  d(g(r, 'fechaFactura')),
        fechaVenc:     d(g(r, 'fechaVenc')),
        estado:        s(g(r, 'estado')),
        estadoDetalle: s(g(r, 'estadoDetalle')),
      });
    }
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
      concepto:  s(r[3]),
      mes:       d(r[4]),
      grupo:     s(r[5]),
      categoria: s(r[6]),
      monto:     n(r[7]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getMerch') {
    const sh = ss.getSheetByName(SHEET_MERCH);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]).map(r => ({
      id:           s(r[0]),
      semana:       d(r[1]),
      tipo:         s(r[2]),
      zona:         s(r[3]),
      cadena:       s(r[4]),
      programado:   n(r[5]),
      porprogramar: n(r[6]),
      efectivo:     n(r[7]),
      obs:          s(r[8]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getProyectos') {
    const sh = ss.getSheetByName(SHEET_PROY);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]).map(r => ({
      id:          s(r[0]),
      nombre:      s(r[1]),
      inicio:      d(r[2]),
      fin:         d(r[3]),
      dias:        n(r[4]),
      personal:    n(r[5]),
      ciudades:    s(r[6]),
      cuentaska:   s(r[7]),
      efectividad: n(r[8]),
      cobertura:   n(r[9]),
      avances:     s(r[10]),
      estado:      s(r[11]),
      obs:         s(r[12]),
      responsable: s(r[13]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getPromo') {
    const out = { kpi: [], materiales: [], incidencias: [] };
    const shK = ss.getSheetByName(SHEET_PKPI);
    if (shK) shK.getDataRange().getValues().slice(1).filter(r => r[0]).forEach(r => out.kpi.push({
      cadena: s(r[0]), cobS1: n(r[1]), cobS2: n(r[2]), efecS1: n(r[3]), efecS2: n(r[4]),
      conectados: n(r[5]), conectadosNota: s(r[6])
    }));
    const shM = ss.getSheetByName(SHEET_PMAT);
    if (shM) shM.getDataRange().getValues().slice(1).filter(r => r[0]).forEach(r => out.materiales.push({
      id: s(r[0]), cadena: s(r[1]), categoria: s(r[2]), material: s(r[3]),
      tiendas: n(r[4]), avanceS1: n(r[5]), avanceS2: n(r[6]), obs: s(r[7]),
      inicio: d(r[8]), fin: d(r[9])
    }));
    const shI = ss.getSheetByName(SHEET_PINC);
    if (shI) shI.getDataRange().getValues().slice(1).filter(r => r[0]).forEach(r => out.incidencias.push({
      id: s(r[0]), cadena: s(r[1]), fecha: d(r[2]), incidencia: s(r[3]),
      responsable: s(r[4]), estado: s(r[5])
    }));
    return jsonResp(out);
  }

  if (accion === 'getCompras') {
    const sh = ss.getSheetByName(SHEET_COMPRA);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]).map(r => ({
      id: s(r[0]), ppto: s(r[1]), material: s(r[2]), cantidad: n(r[3]),
      precioUnit: n(r[4]), total: n(r[5]), fechaCompra: d(r[6]), cuenta: s(r[7]), obs: s(r[8]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getPptos') {
    const sh = ss.getSheetByName(SHEET_PPTO);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]).map(r => ({
      id: s(r[0]), nombre: s(r[1]), mesAprobado: d(r[2]), montoAprobado: n(r[3]), correo: s(r[4]), obs: s(r[5]),
    }));
    return jsonResp({ rows, count: rows.length });
  }

  if (accion === 'getCuotas') {
    const sh = ss.getSheetByName(SHEET_CUOTAS);
    if (!sh) return jsonResp({ rows: [], count: 0 });
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1).filter(r => r[0]).map(r => ({
      ejecutivo: s(r[0]), anio: s(r[1]),
      metaQ1: n(r[2]), metaQ2: n(r[3]), metaQ3: n(r[4]), metaQ4: n(r[5]),
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

  // Carga masiva de facturas (una sola escritura)
  if (accion === 'saveFacturacionBulk') {
    const sh = getOrCreate(ss, SHEET_FACT, HDR_FACT);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    const base = new Date().getTime();
    const rows = items.map((b, i) => [
      b.id || ('F' + (base + i)), b.cliente||'', b.ejecutivo||'Claudia Camarena', b.tipo||'', b.responsable||'',
      b.servicio||'', b.mes||'', num(b.importe), b.os||'', b.serie||'', b.factura||'',
      b.fechaFactura||'', b.fechaVenc||'', b.estado||'Pendiente', b.estadoDetalle||''
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HDR_FACT.length).setValues(rows);
    return jsonResp({ ok: true, added: rows.length });
  }

  // ── GASTOS ─────────────────────────────────────────────────────
  if (accion === 'saveGasto') {
    const sh = getOrCreate(ss, SHEET_GASTOS, HDR_GASTOS);
    const b = body;
    if (!b.cliente || !b.tipo || !b.mes) return jsonResp({ error: 'Cliente, servicio y mes requeridos' });
    const id = b.id || ('G' + new Date().getTime());
    const row = [id, b.cliente||'', b.tipo||'', b.concepto||'', b.mes||'', b.grupo||'Otros',
      b.categoria||'', num(b.monto)];
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

  // Carga masiva de gastos (una sola escritura)
  if (accion === 'saveGastoBulk') {
    const sh = getOrCreate(ss, SHEET_GASTOS, HDR_GASTOS);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    const base = new Date().getTime();
    const rows = items.map((b, i) => [
      b.id || ('G' + (base + i)), b.cliente||'', b.tipo||'', b.concepto||'', b.mes||'', b.grupo||'Otros',
      b.categoria||'', num(b.monto)
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HDR_GASTOS.length).setValues(rows);
    return jsonResp({ ok: true, added: rows.length });
  }

  // ── MERCH ──────────────────────────────────────────────────────
  if (accion === 'saveMerch') {
    const sh = getOrCreate(ss, SHEET_MERCH, HDR_MERCH);
    const b = body;
    if (!b.cadena || !b.semana) return jsonResp({ error: 'Cadena y semana requeridos' });
    const id = b.id || ('M' + new Date().getTime());
    const row = [id, b.semana||'', b.tipo||'', b.zona||'', b.cadena||'',
      num(b.programado), num(b.porprogramar), num(b.efectivo), b.obs||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_MERCH.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }

  if (accion === 'deleteMerch') {
    const sh = ss.getSheetByName(SHEET_MERCH);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  if (accion === 'saveMerchBulk') {
    const sh = getOrCreate(ss, SHEET_MERCH, HDR_MERCH);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    const base = new Date().getTime();
    const rows = items.map((b, i) => [
      b.id || ('M' + (base + i)), b.semana||'', b.tipo||'', b.zona||'', b.cadena||'',
      num(b.programado), num(b.porprogramar), num(b.efectivo), b.obs||''
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HDR_MERCH.length).setValues(rows);
    return jsonResp({ ok: true, added: rows.length });
  }

  // ── PROYECTOS ──────────────────────────────────────────────────
  if (accion === 'saveProyecto') {
    const sh = getOrCreate(ss, SHEET_PROY, HDR_PROY);
    const b = body;
    if (!b.nombre) return jsonResp({ error: 'Nombre del proyecto requerido' });
    const id = b.id || ('P' + new Date().getTime());
    const row = [id, b.nombre||'', b.inicio||'', b.fin||'', num(b.dias), num(b.personal),
      b.ciudades||'', b.cuentaska||'', num(b.efectividad), num(b.cobertura),
      b.avances||'', b.estado||'Activo', b.obs||'', b.responsable||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_PROY.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }

  if (accion === 'deleteProyecto') {
    const sh = ss.getSheetByName(SHEET_PROY);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  // ── PROMOTORÍA ─────────────────────────────────────────────────
  if (accion === 'savePromoKPI') {
    const sh = getOrCreate(ss, SHEET_PKPI, HDR_PKPI);
    const b = body;
    if (!b.cadena) return jsonResp({ error: 'Cadena requerida' });
    const row = [b.cadena, num(b.cobS1), num(b.cobS2), num(b.efecS1), num(b.efecS2), num(b.conectados), b.conectadosNota||''];
    const found = findRow(sh, b.cadena);
    if (found > 0) sh.getRange(found, 1, 1, HDR_PKPI.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
  }

  if (accion === 'savePromoMaterial') {
    const sh = getOrCreate(ss, SHEET_PMAT, HDR_PMAT);
    const b = body;
    if (!b.cadena || !b.material) return jsonResp({ error: 'Cadena y material requeridos' });
    const id = b.id || ('PM' + new Date().getTime());
    const row = [id, b.cadena||'', b.categoria||'', b.material||'', num(b.tiendas), num(b.avanceS1), num(b.avanceS2), b.obs||'', b.inicio||'', b.fin||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_PMAT.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }
  if (accion === 'deletePromoMaterial') {
    const sh = ss.getSheetByName(SHEET_PMAT);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  if (accion === 'savePromoIncidencia') {
    const sh = getOrCreate(ss, SHEET_PINC, HDR_PINC);
    const b = body;
    if (!b.cadena) return jsonResp({ error: 'Cadena requerida' });
    const id = b.id || ('PI' + new Date().getTime());
    const row = [id, b.cadena||'', b.fecha||'', b.incidencia||'', b.responsable||'', b.estado||'Abierta'];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_PINC.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }
  if (accion === 'deletePromoIncidencia') {
    const sh = ss.getSheetByName(SHEET_PINC);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  // ── PPTO APROBADO ──────────────────────────────────────────────
  if (accion === 'savePpto') {
    const sh = getOrCreate(ss, SHEET_PPTO, HDR_PPTO);
    const b = body;
    if (!b.nombre) return jsonResp({ error: 'Nombre del PPTO requerido' });
    const id = b.id || ('A' + new Date().getTime());
    const row = [id, b.nombre||'', b.mesAprobado||'', num(b.montoAprobado), b.correo||'', b.obs||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_PPTO.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }
  if (accion === 'deletePpto') {
    const sh = ss.getSheetByName(SHEET_PPTO);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }

  // ── COMPRA DE MATERIALES ───────────────────────────────────────
  if (accion === 'saveCompra') {
    const sh = getOrCreate(ss, SHEET_COMPRA, HDR_COMPRA);
    const b = body;
    if (!b.material) return jsonResp({ error: 'Material requerido' });
    const id = b.id || ('C' + new Date().getTime());
    const total = num(b.total) || (num(b.cantidad) * num(b.precioUnit));
    const row = [id, b.ppto||'', b.material||'', num(b.cantidad), num(b.precioUnit), total, b.fechaCompra||'', b.cuenta||'', b.obs||''];
    const found = findRow(sh, id);
    if (found > 0) sh.getRange(found, 1, 1, HDR_COMPRA.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, id: id, action: found > 0 ? 'updated' : 'created' });
  }
  if (accion === 'deleteCompra') {
    const sh = ss.getSheetByName(SHEET_COMPRA);
    const found = findRow(sh, body.id);
    if (found > 0) { sh.deleteRow(found); return jsonResp({ ok: true, action: 'deleted' }); }
    return jsonResp({ ok: false, error: 'ID no encontrado' });
  }
  if (accion === 'saveCompraBulk') {
    const sh = getOrCreate(ss, SHEET_COMPRA, HDR_COMPRA);
    const items = body.rows || [];
    if (!items.length) return jsonResp({ error: 'Sin filas para cargar' });
    const base = new Date().getTime();
    const rows = items.map((b, i) => {
      const total = num(b.total) || (num(b.cantidad) * num(b.precioUnit));
      return [b.id || ('C' + (base + i)), b.ppto||'', b.material||'', num(b.cantidad), num(b.precioUnit), total, b.fechaCompra||'', b.cuenta||'', b.obs||''];
    });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HDR_COMPRA.length).setValues(rows);
    return jsonResp({ ok: true, added: rows.length });
  }

  // ── CUOTAS (metas trimestrales, manual) ────────────────────────
  if (accion === 'saveCuota') {
    const sh = getOrCreate(ss, SHEET_CUOTAS, HDR_CUOTAS);
    const b = body;
    if (!b.ejecutivo || !b.anio) return jsonResp({ error: 'Ejecutivo y año requeridos' });
    const row = [b.ejecutivo, ''+b.anio, num(b.metaQ1), num(b.metaQ2), num(b.metaQ3), num(b.metaQ4)];
    const data = sh.getDataRange().getValues();
    let found = -1;
    for (var i = 1; i < data.length; i++) {
      if (s(data[i][0]) === s(b.ejecutivo) && s(data[i][1]) === s(b.anio)) { found = i + 1; break; }
    }
    if (found > 0) sh.getRange(found, 1, 1, HDR_CUOTAS.length).setValues([row]);
    else sh.appendRow(row);
    return jsonResp({ ok: true, action: found > 0 ? 'updated' : 'created' });
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

// Número robusto: acepta number real o texto (incl. formato PE "1.250,50" y "S/")
function n(v) {
  if (typeof v === 'number') return v;
  var str = s(v).replace(/\s/g, '').replace(/[^0-9.,\-]/g, '');
  if (!str) return 0;
  var hasC = str.indexOf(',') >= 0, hasD = str.indexOf('.') >= 0;
  if (hasC && hasD) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) str = str.replace(/\./g, '').replace(',', '.');
    else str = str.replace(/,/g, '');
  } else if (hasC) {
    str = str.replace(',', '.');
  }
  var x = parseFloat(str);
  return isNaN(x) ? 0 : x;
}
function num(v) { return n(v); }

// Localiza columnas por nombre de encabezado (normaliza mayúsculas/acentos/símbolos)
function colFinder(headerRow) {
  var norm = function (x) {
    return s(x).toUpperCase()
      .replace(/[ÁÀÄÂ]/g, 'A').replace(/[ÉÈÊ]/g, 'E').replace(/[ÍÌÎ]/g, 'I')
      .replace(/[ÓÒÔ]/g, 'O').replace(/[ÚÙÛ]/g, 'U').replace(/Ñ/g, 'N')
      .replace(/[^A-Z0-9]/g, '');
  };
  var idx = {};
  headerRow.forEach(function (h, i) { var k = norm(h); if (k && !(k in idx)) idx[k] = i; });
  return function (aliases) {
    for (var j = 0; j < aliases.length; j++) { var k = norm(aliases[j]); if (k in idx) return idx[k]; }
    return -1;
  };
}

// Zona horaria del sheet (cacheada) — evita el corrimiento de 1 día
var _TZ;
function tz() {
  if (!_TZ) {
    try { _TZ = getSS().getSpreadsheetTimeZone() || 'America/Lima'; }
    catch (e) { _TZ = 'America/Lima'; }
  }
  return _TZ;
}

// Fecha -> 'YYYY-MM-DD' (acepta Date o texto). Formatea en la zona horaria del sheet.
function d(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, tz(), 'yyyy-MM-dd');
  }
  const t = s(v);
  if (!t) return '';
  // Normaliza textos DD/MM/YYYY -> YYYY-MM-DD
  const m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return m[3] + '-' + pad(m[2]) + '-' + pad(m[1]);
  return t;
}

function pad(x) { return String(x).padStart(2, '0'); }

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
