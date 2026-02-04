// ====== CONFIG ======
const SPREADSHEET_ID = '1mrBkcP3Wz04KfBxmNXP0tn6GI645lKctT095uW43ezw';
const SHEETS = { Empaquetado: 'EMPAQUETADO', Merma: 'MERMA' };

const PRODUCT_SHEET = 'PRODUCTOS';
const PRODUCT_HEADERS = ['CODIGOS', 'DESCRIPCION', 'Unidad_Primaria'];

const ADMIN_KEY = 'PASANTIAS90';
const TZ = 'America/Caracas';
const NONCE_TTL_SECONDS = 3600;
const PRODUCT_CACHE_SECONDS = 60;

// ====== HTTP ======
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.ping) return respond({ ok: true, pong: String(p.ping), version: '2025-11-13' });
  if (String(p.action || '') === 'recent') return recentEndpoint(String(p.sheet||'').trim(), clampLimit(p.limit));
  if (String(p.action || '') === 'productos') return getProductosEndpoint();
  return respond({ ok: true, version: '2025-11-13' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(5000);
    const params = e && e.parameter ? e.parameter : {};
    const action = (params.action || '').trim();

    if (action === 'addProduct') return addProductEndpoint(params);
    if (action === 'deleteProduct') return deleteProductEndpoint(params);

    const sheetKey = String(params.sheet || '').trim();
    const sheetName = SHEETS[sheetKey];
    if (!sheetName) return respond({ ok:false, error:'Parametro sheet inválido (usa Empaquetado o Merma)' });

    const nonce = (params.nonce || '').trim();
    if (nonce && isDuplicateNonce(nonce)) {
      return respond({ ok:true, duplicate:true, nonce });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return respond({ ok:false, error:'No existe la pestaña: '+sheetName });
    // Mantener tablas pero asegurar encabezados válidos
    normalizeTables(sh, null);

    const productos = parseProductos(params.productos_json);
    const marcaTemporal = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  let rows = [];
  let entradasRows = [];
  var writeCols = null; // will hold the number of columns to write (ensures Merma uses colCount)

      if (sheetKey === 'Empaquetado') {
      const desiredHeader = [
        'Marca temporal',
        'DIRECCION',
        'FECHA',
        'PRODUCTO',
        'CANTIDAD',
        'ENTREGADO A',
        'NUMERO REGISTRO',
        'RESPONSABLE',
        'SEDE'
      ];
  const colCount = ensureHeaderFull(sh, desiredHeader);
  normalizeTables(sh, desiredHeader);
  writeCols = colCount;

      const fecha = params.fecha || '';
      const entregado = params.entregado || '';
      const registro = params.registro || '';
      const responsable = params.responsable || '';
      const sede = params.sede || params.empresa || '';
      const loteGlobal = (params.lote || '').trim();
      const direccionValor = '.'; // punto para evitar celda vacía

      if (productos.length) {
        productos.forEach(it => {
          const loteItem = (it && (it as any).lote) ? String((it as any).lote).trim() : loteGlobal;
          let r = [
            marcaTemporal,
            direccionValor,
            fecha,
            it.descripcion || it.codigo || '',
            toNumber(it.cantidad),
            entregado,
            registro,
            responsable,
            sede
          ];
          if (r.length < colCount) {
            while (r.length < colCount) r.push('.');
          }
          rows.push(r);
        });
      } else {
        let r = [marcaTemporal, direccionValor, fecha, '', 0, entregado, registro, responsable, sede];
        while (r.length < colCount) r.push('.');
        rows.push(r);
      }

      // Construir Entradas09 desde las filas ya armadas (mismo dato que EMPAQUETADO)
      entradasRows = rows.map(r => {
        const loteVal = r[11] || '';
        const prodVal = r[3] || '';
        const cantVal = toNumber(r[4]);
        const fechaVal = r[2] || '';
        return [loteVal, prodVal, cantVal, fechaVal];
      });

    } else if (sheetKey === 'Merma') {
      const desiredHeader = [
        'Marca Temporal',
        'FECHA',
        'PRODUCTO',
        'UNIDAD DE MEDIDA',
        'SEDE',
        'MOTIVO DE MERMA',
        'CANTIDAD DEL MOTIVO DE MERMA',
        'NUMERO DE LOTE',
        'RESPONSABLE'
      ];
  const colCount = ensureHeaderFull(sh, desiredHeader);
  writeCols = colCount;

      const fecha = params.fecha || '';
      const sede = params.sede || params.empresa || '';
      const responsable = params.responsable || '';
      const motivoGlobal = (params.motivo || '').trim();
      const loteGlobal = (params.lote || '').trim();

      // debug: show incoming productos and global motivo/lote
      try { console.log('MERMA incoming productos:', JSON.stringify(productos)); } catch(_) {}
      try { console.log('MERMA motivoGlobal:', motivoGlobal, 'loteGlobal:', loteGlobal); } catch(_) {}

      if (productos.length) {
        productos.forEach(function(it) {
          var motivoItem = (it && it.motivo) ? String(it.motivo).trim() : '';
          if (!motivoItem) motivoItem = motivoGlobal;
          var loteItem = (it && it.lote) ? String(it.lote).trim() : '';
          if (!loteItem) loteItem = loteGlobal;

          var descripcionVal = it && (it.descripcion || it.codigo) ? (it.descripcion || it.codigo) : '';
          var unidadVal = it && it.unidad ? it.unidad : '';
          var cantidadVal = toNumber(it && it.cantidad);

          var r = [
            marcaTemporal,
            fecha,
            descripcionVal,
            unidadVal,
            sede,
            motivoItem,
            cantidadVal,
            loteItem,
            responsable
          ];
          while (r.length < colCount) r.push('.');
          rows.push(r);
        });
      } else {
        var r = [marcaTemporal, fecha, '', '', sede, motivoGlobal, 0, loteGlobal, responsable];
        while (r.length < colCount) r.push('.');
        rows.push(r);
      }
    }

    if (rows.length) {
      var colsToWrite = writeCols || (rows[0] ? rows[0].length : 1);
  try { console.log('Writing rows count:', rows.length, 'cols:', colsToWrite); } catch(_) {}
  try { console.log('MERMA rows to write:', JSON.stringify(rows)); } catch(_) {}
  writeRowsSafe(sh, sh.getLastRow() + 1, 1, rows, colsToWrite, 'write_main_'+sheetKey);
      if (sheetKey === 'Empaquetado' && entradasRows.length) {
        let entradasSheet = ss.getSheetByName('Entradas09');
        if (!entradasSheet) {
          entradasSheet = ss.insertSheet('Entradas09');
        }
        normalizeTables(entradasSheet, [
          'NUMERO DE LOTE',
          'PRODUCTO',
          'CANTIDAD EMPAQUETADO',
          'FECHA EMPAQUETADO',
          'CANTIDAD ALMACEN',
          'FECHA ENTRADA'
        ]);
        ensureHeaderPrefixFull(entradasSheet, [
          'NUMERO DE LOTE',
          'PRODUCTO',
          'CANTIDAD EMPAQUETADO',
          'FECHA EMPAQUETADO'
        ]);
        writeRowsSafe(entradasSheet, entradasSheet.getLastRow() + 1, 1, entradasRows, 4, 'write_entradas09');
      }
      if (nonce) storeNonce(nonce);
    }

    return respond({ ok:true, inserted: rows.length, nonce });
  } catch(err) {
    const msg = String(err && (err as any).message || err);
    if (msg.indexOf('encabezado') !== -1) {
      try {
        const ssDbg = SpreadsheetApp.openById(SPREADSHEET_ID);
        const shEmp = ssDbg.getSheetByName('EMPAQUETADO');
        const shEnt = ssDbg.getSheetByName('Entradas09');
        return respond({
          ok:false,
          error: msg,
          debug: {
            empaquetado: shEmp ? getSheetDebug_(shEmp) : null,
            entradas09: shEnt ? getSheetDebug_(shEnt) : null
          }
        });
      } catch(_){ }
    }
    return respond({ ok:false, error: msg });
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// ====== Catálogo PRODUCTOS ======
function getProductosEndpoint() {
  const cache = CacheService.getScriptCache();
  const key = 'productos_cache_v1';
  const cached = cache.get(key);
  if (cached) {
    try { return respond(JSON.parse(cached)); } catch(_) {}
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) return respond({ ok:false, error:'Hoja PRODUCTOS no encontrada' });

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    const resp = { ok:true, products:[] };
    cache.put(key, JSON.stringify(resp), PRODUCT_CACHE_SECONDS);
    return respond(resp);
  }

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idxCodigo = headers.findIndex(h => String(h).toUpperCase() === 'CODIGOS');
  const idxDesc   = headers.findIndex(h => String(h).toUpperCase() === 'DESCRIPCION');
  const idxUnidad = headers.findIndex(h => String(h).toUpperCase() === 'UNIDAD_PRIMARIA');

  const values = sh.getRange(2,1,lastRow-1,sh.getLastColumn()).getValues();
  const products = values.map(r => ({
    codigo: String(idxCodigo >=0 ? r[idxCodigo] : '').trim(),
    descripcion: String(idxDesc >=0 ? r[idxDesc] : '').trim(),
    unidad: String(idxUnidad >=0 ? r[idxUnidad] : '').trim()
  })).filter(p => p.codigo || p.descripcion);

  const resp = { ok:true, products };
  cache.put(key, JSON.stringify(resp), PRODUCT_CACHE_SECONDS);
  return respond(resp);
}

function addProductEndpoint(params) {
  const codigo = String(params.codigo || '').trim();
  const descripcion = String(params.descripcion || '').trim();
  const unidad = String(params.unidad || '').trim() || 'PAQ';
  const adminKey = String(params.adminKey || '').trim();

  if (!codigo || !descripcion) return respond({ ok:false, error:'Faltan codigo o descripcion' });
  if (adminKey !== ADMIN_KEY) return respond({ ok:false, error:'adminKey inválido' });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(PRODUCT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PRODUCT_SHEET);
    sh.getRange(1,1,1,PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]);
  }
  ensureHeaderFull(sh, PRODUCT_HEADERS);

  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const codigos = sh.getRange(2,1,lastRow-1,1).getValues().flat().map(v => String(v).trim().toLowerCase());
    if (codigos.includes(codigo.toLowerCase())) {
      return respond({ ok:false, error:'Código ya existe' });
    }
  }
  sh.appendRow([codigo, descripcion, unidad]);
  CacheService.getScriptCache().remove('productos_cache_v1');
  return respond({ ok:true, inserted:1 });
}

function deleteProductEndpoint(params) {
  const codigo = String(params.codigo || '').trim();
  const adminKey = String(params.adminKey || '').trim();
  if (!codigo) return respond({ ok:false, error:'Falta codigo' });
  if (adminKey !== ADMIN_KEY) return respond({ ok:false, error:'adminKey inválido' });

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(PRODUCT_SHEET);
    if (!sh) return respond({ ok:false, error:'Hoja PRODUCTOS no encontrada' });

    ensureHeaderFull(sh, PRODUCT_HEADERS);
    const last = sh.getLastRow();
    if (last < 2) return respond({ ok:false, notFound:true });

    const codigos = sh.getRange(2,1,last-1,1).getValues();
    let rowToDelete = -1;
    for (let i=0;i<codigos.length;i++){
      if (String(codigos[i][0]).trim().toLowerCase() === codigo.toLowerCase()) {
        rowToDelete = i+2;
        break;
      }
    }
    if (rowToDelete === -1) return respond({ ok:false, notFound:true });
    sh.deleteRow(rowToDelete);
    CacheService.getScriptCache().remove('productos_cache_v1');
    return respond({ ok:true, deleted:1, codigo });
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

// ====== Recent ======
function clampLimit(raw){
  let limit = parseInt(raw,10);
  if (!isFinite(limit) || limit < 1) limit = 20;
  if (limit > 200) limit = 200;
  return limit;
}

function recentEndpoint(sheetKey, limit){
  const sheetName = SHEETS[sheetKey];
  if (!sheetName) return respond({ ok:false, error:'sheet inválido' });
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return respond({ ok:false, error:'Hoja no encontrada: '+sheetName });

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return respond({ ok:true, headers:[], rows:[], total:0, sheet:sheetName });

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;
  const values = sh.getRange(startRow,1,numRows,lastCol).getValues();
  values.reverse();
  const rows = values.map(r => {
    const obj = {};
    for (let i=0;i<headers.length;i++) obj[headers[i]] = r[i];
    return obj;
  });
  return respond({ ok:true, headers, rows, total:numRows, sheet:sheetName });
}

// ====== Nonce ======
function cacheKey(nonce){ return 'nonce:'+nonce; }
function isDuplicateNonce(nonce){
  if (!nonce) return false;
  const cache = CacheService.getScriptCache();
  return cache.get(cacheKey(nonce)) === '1';
}
function storeNonce(nonce){
  if (!nonce) return;
  CacheService.getScriptCache().put(cacheKey(nonce),'1', NONCE_TTL_SECONDS);
}

// ====== Utilidades ======
function parseProductos(jsonStr){
  try { const arr = JSON.parse(jsonStr || '[]'); return Array.isArray(arr)?arr:[]; }
  catch(_){ return []; }
}
function toNumber(n){
  const v = Number(n);
  return isFinite(v) && v > 0 ? v : 0;
}

// Encabezado robusto: rellena vacíos y extiende hasta el total de columnas usadas por la hoja.
function ensureHeaderFull(sh, desired){
  const colCount = Math.max(sh.getLastColumn(), desired.length);
  const existing = colCount ? sh.getRange(1,1,1,colCount).getValues()[0] : [];
  const out = [];
  for (let i=0;i<colCount;i++){
    let val = i < desired.length ? desired[i] : existing[i];
    if (!val || String(val).trim()==='') val = 'Col_'+(i+1);
    out.push(val);
  }
  try {
    sh.getRange(1,1,1,colCount).setValues([out]);
  } catch (err) {
    const msg = String(err && (err as any).message || err);
    if (msg && msg.indexOf('encabezado') !== -1) {
      tryFixTables(sh);
      sh.getRange(1,1,1,colCount).setValues([out]);
    } else {
      throw err;
    }
  }
  ensureTableHeaders(sh);
  return colCount;
}

function respond(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetDebug_(sh){
  const info:any = {
    name: sh.getName(),
    lastRow: sh.getLastRow(),
    lastCol: sh.getLastColumn(),
    maxCols: sh.getMaxColumns(),
    tables: []
  };
  try {
    if (typeof (sh as any).getTables === 'function') {
      const tables = (sh as any).getTables();
      tables.forEach(t => {
        let hr = (typeof (t as any).getHeaderRange === 'function')
          ? (t as any).getHeaderRange()
          : (typeof (t as any).getHeaderRowRange === 'function' ? (t as any).getHeaderRowRange() : null);
        if (!hr && typeof (t as any).getRange === 'function') {
          const tr = (t as any).getRange();
          hr = sh.getRange(tr.getRow(), tr.getColumn(), 1, tr.getNumColumns());
        }
        const rowVals = hr ? hr.getValues()[0] : null;
        info.tables.push({
          headerRow: hr ? hr.getRow() : null,
          headerCol: hr ? hr.getColumn() : null,
          headerNumCols: hr ? hr.getNumColumns() : null,
          headerValues: rowVals
        });
      });
    } else {
      info.tables = null;
    }
  } catch (e) {
    info.tablesError = String(e && (e as any).message || e);
  }
  return info;
}

// Adjustments to handle dynamic motivo and lote per product in MERMA
function postMerma_(payload) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    if (isDuplicateNonce(payload.nonce)) {
      return respond({ ok: true, duplicate: true, nonce: payload.nonce });
    }
    const sheet = getSheet_(SHEETS.Merma);
    const { colCount } = ensureHeaderFull(sheet, [
      'Marca Temporal',
      'FECHA',
      'PRODUCTO',
      'UNIDAD DE MEDIDA',
      'SEDE',
      'MOTIVO DE MERMA',
      'CANTIDAD DEL MOTIVO DE MERMA',
      'NUMERO DE LOTE',
      'RESPONSABLE'
    ]);

    const fecha = payload.fecha || '';
    const sede = payload.sede || payload.empresa || '';
    const responsable = payload.responsable || '';
    const motivoGlobal = (payload.motivo || '').trim();
    const loteGlobal = (payload.lote || '').trim();

    const productos = Array.isArray(payload.productos) ? payload.productos : [];
    if (!productos.length) {
      throw new Error('Sin productos para registrar en MERMA');
    }

    const rows = productos.map(it => {
      const motivoItem = String(it.motivo || '').trim() || motivoGlobal;
      const loteItem = String(it.lote || '').trim() || loteGlobal;
      return [
        Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
        fecha,
        it.descripcion || it.codigo || '',
        it.unidad || '',
        sede,
        motivoItem,
        toNumber(it.cantidad),
        loteItem,
        responsable
      ];
    });

    rows.forEach(row => {
      while (row.length < colCount) row.push('.');
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, colCount).setValues(rows);
    if (payload.nonce) storeNonce(payload.nonce);

    return respond({ ok: true, inserted: rows.length, nonce: payload.nonce });
  } catch (err) {
    return respond({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// Encabezado completo: fija A:D y rellena el resto para evitar celdas vacías en tablas
function ensureHeaderPrefixFull(sh, headers){
  const colCount = Math.max(sh.getLastColumn(), sh.getMaxColumns(), headers.length, 1);
  const existing = sh.getRange(1,1,1,colCount).getValues()[0];
  const out = [];
  for (let i=0;i<colCount;i++){
    let val = i < headers.length ? headers[i] : existing[i];
    if (!val || String(val).trim()==='') val = 'Col_'+(i+1);
    out.push(val);
  }
  try {
    sh.getRange(1,1,1,colCount).setValues([out]);
  } catch (err) {
    const msg = String(err && (err as any).message || err);
    if (msg && msg.indexOf('encabezado') !== -1) {
      tryFixTables(sh);
      sh.getRange(1,1,1,colCount).setValues([out]);
    } else {
      throw err;
    }
  }
  ensureTableHeaders(sh);
}

// Rellena encabezados vacíos dentro de tablas (si existen)
function ensureTableHeaders(sh){
  try {
    if (typeof sh.getTables !== 'function') return;
    const tables = sh.getTables();
    if (!tables || !tables.length) return;
    tables.forEach(t => {
      try {
        let headerRange = (typeof t.getHeaderRange === 'function')
          ? t.getHeaderRange()
          : (typeof t.getHeaderRowRange === 'function' ? t.getHeaderRowRange() : null);
        if (!headerRange && typeof t.getRange === 'function') {
          const tr = t.getRange();
          headerRange = sh.getRange(tr.getRow(), tr.getColumn(), 1, tr.getNumColumns());
        }
        if (!headerRange) return;
        const values = headerRange.getValues();
        const startCol = headerRange.getColumn();
        let changed = false;
        for (let r=0;r<values.length;r++){
          for (let c=0;c<values[r].length;c++){
            if (!values[r][c] || String(values[r][c]).trim()==='') {
              values[r][c] = 'Col_'+(startCol + c);
              changed = true;
            }
          }
        }
        if (changed) headerRange.setValues(values);
      } catch(_){ }
    });
  } catch(_){ }
}

function writeRowsSafe(sh, startRow, startCol, rows, colCount, label){
  try {
    sh.getRange(startRow, startCol, rows.length, colCount).setValues(rows);
  } catch (err) {
    const msg = String(err && err.message || err);
    if (msg && msg.indexOf('encabezado') !== -1) {
      normalizeTables(sh, null);
      sh.getRange(startRow, startCol, rows.length, colCount).setValues(rows);
      return;
    }
    throw new Error(label + ': ' + msg);
  }
}

function tryFixTables(sh){
  try { normalizeTables(sh, null); } catch(_){ }
}

function normalizeTables(sh, preferredHeaders){
  try {
    if (typeof (sh as any).getTables !== 'function') return;
    const tables = (sh as any).getTables();
    if (!tables || !tables.length) return;
    tables.forEach(t => {
      try {
        const tr = (t as any).getRange();
        const headerRow = tr.getRow();
        const headerCol = tr.getColumn();
        const numCols = tr.getNumColumns();
        const headerRange = sh.getRange(headerRow, headerCol, 1, numCols);
        const values = headerRange.getValues()[0];
        for (let c=0;c<numCols;c++){
          let val = (preferredHeaders && c < preferredHeaders.length) ? preferredHeaders[c] : values[c];
          if (!val || String(val).trim()==='') val = 'Col_'+(headerCol + c);
          values[c] = val;
        }
        headerRange.setValues([values]);
      } catch(_){ }
    });
  } catch(_){ }
}