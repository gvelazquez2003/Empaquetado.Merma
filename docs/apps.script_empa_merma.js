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

    const productos = hydrateProductos(parseProductos(params.productos_json), params);
    const marcaTemporal = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  let rows = [];
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
        'SEDE',
        'FORMULA 1',
        'FORMULA 2',
        'NUMERO DE LOTE'
      ];
  const colCount = ensureHeaderFull(sh, desiredHeader);
  writeCols = desiredHeader.length;

      const fecha = params.fecha || '';
      const entregado = params.entregado || '';
      const registro = params.registro || '';
      const responsable = params.responsable || '';
      const sede = params.sede || params.empresa || '';
      const direccionValor = '.'; // punto para evitar celda vacía

      if (productos.length) {
        productos.forEach(it => {
          let r = [
            marcaTemporal,
            direccionValor,
            fecha,
            it.descripcion || it.codigo || '',
            toNumber(it.cantidad),
            entregado,
            registro,
            responsable,
            sede,
            '',
            '',
            (it && it.lote) ? String(it.lote).trim() : ''
          ];
          rows.push(fitRow(r, writeCols));
        });
      } else {
        let r = [marcaTemporal, direccionValor, fecha, '', 0, entregado, registro, responsable, sede, '', '', ''];
        rows.push(fitRow(r, writeCols));
      }

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
  writeCols = desiredHeader.length;

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
          rows.push(fitRow(r, writeCols));
        });
      } else {
        var r = [marcaTemporal, fecha, '', '', sede, motivoGlobal, 0, loteGlobal, responsable];
        rows.push(fitRow(r, writeCols));
      }
    }

    if (rows.length) {
      var colsToWrite = writeCols || (rows[0] ? rows[0].length : 1);
  try { console.log('Writing rows count:', rows.length, 'cols:', colsToWrite); } catch(_) {}
  try { console.log('MERMA rows to write:', JSON.stringify(rows)); } catch(_) {}
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, colsToWrite).setValues(rows);
      if (nonce) storeNonce(nonce);
    }

    return respond({ ok:true, inserted: rows.length, nonce });
  } catch(err) {
    return respond({ ok:false, error:String(err && err.message || err) });
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
function fitRow(row, targetLen){
  const out = Array.isArray(row) ? row.slice(0, targetLen) : [];
  while (out.length < targetLen) out.push('.');
  if (out.length > targetLen) out.length = targetLen;
  return out;
}
function hydrateProductos(productos, params){
  const arr = Array.isArray(productos) ? productos.slice() : [];
  const countRaw = params && params.productos_count;
  const count = parseInt(countRaw, 10);
  if (!isFinite(count) || count <= 0) return arr;
  for (let i = 0; i < count; i++){
    const motivo = params['motivo_'+i] ? String(params['motivo_'+i]).trim() : '';
    const lote = params['lote_'+i] ? String(params['lote_'+i]).trim() : '';
    const codigo = params['prodCodigo_'+i] ? String(params['prodCodigo_'+i]).trim() : '';
    if (!arr[i]) arr[i] = {};
    if (motivo && !arr[i].motivo) arr[i].motivo = motivo;
    if (lote && !arr[i].lote) arr[i].lote = lote;
    if (codigo && !arr[i].codigo) arr[i].codigo = codigo;
  }
  return arr;
}

// Encabezado robusto: rellena vacíos y extiende hasta el total de columnas usadas por la hoja.
function ensureHeaderFull(sh, desired){
  const colCount = Math.max(sh.getLastColumn(), desired.length, 1);
  const existing = colCount ? sh.getRange(1,1,1,colCount).getValues()[0] : [];
  const out = [];
  for (let i=0;i<colCount;i++){
    let val = i < desired.length ? desired[i] : existing[i];
    if (!val || String(val).trim()==='') val = 'Col_'+(i+1);
    out.push(val);
  }
  try {
    sh.getRange(1,1,1,colCount).setValues([out]);
    return colCount;
  } catch (err) {
    // Fallback: escribir encabezados hasta el máximo de columnas de la hoja
    const maxCols = Math.max(sh.getMaxColumns(), colCount, desired.length, 1);
    const existingAll = maxCols ? sh.getRange(1,1,1,maxCols).getValues()[0] : [];
    const outAll = [];
    for (let i=0;i<maxCols;i++){
      let val = i < desired.length ? desired[i] : existingAll[i];
      if (!val || String(val).trim()==='') val = 'Col_'+(i+1);
      outAll.push(val);
    }
    sh.getRange(1,1,1,maxCols).setValues([outAll]);
    return maxCols;
  }
}

function respond(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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