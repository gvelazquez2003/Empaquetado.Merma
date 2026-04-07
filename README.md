# Formularios Empaquetados y Merma

Frontend estático para registrar Empaquetados y Merma, conectado a un Google Apps Script que escribe en Google Sheets.

## Estructura del repo

```
/docs/                 # Carpeta publicada por GitHub Pages (única fuente)
   index.html           # App principal con ambos formularios
   menu.html            # Pantalla de menú para elegir formulario
   styles.css           # Estilos
   script.js            # Envío hacia Apps Script + lógica de formularios
   BRAND.svg            # Logo de ejemplo (opcional)
   CODIGOS DESCRIPCION Unidad_Primaria.tsv  # Referencia de productos (opcional)
   productos.html       # Gestión del catálogo (listar y agregar productos)
```

GitHub Pages publicará el contenido de `docs/`. Se eliminaron duplicados para mantener una sola fuente.

## Configurar la URL del backend (Apps Script)

En `docs/script.js` está esta constante:

```js
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzNw7ko-872o3NPxXxN8pnsEoviYdG2JhWJNh5tK1k4KTHv-1m8qDUm9pjgQVGfEJyACg/exec";
```

Sustitúyela por tu URL de despliegue si cambia (termina en `/exec`). El resto se arma solo.

## Correr localmente (opciones)

- Opción rápida: abre `docs/menu.html` con doble clic. Debido a CORS abiertos en Apps Script, suele funcionar.
- Recomendado: usa una extensión tipo "Live Server" en VS Code y abre `docs/`.

## Publicar en GitHub Pages

1. Crea el repo en GitHub (ej: `usuario/Formularios.Empa.Merma`).
2. En tu carpeta del proyecto:

```bat
:: Inicializa el repo local y primer commit
git init
git add .
git commit -m "Inicializa front + docs para GitHub Pages"

:: Crea rama principal
git branch -M main

:: Conecta con tu repo remoto (copia tu URL HTTPS)
:: Ejemplo: git remote add origin https://github.com/usuario/Formularios.Empa.Merma.git

:: Sube los cambios
git push -u origin main
```

3. En GitHub > Settings > Pages:
   - Source: Deploy from a branch
   - Branch: `main` y carpeta `/docs`

4. Tu sitio quedará disponible en: `https://empaquetado-merma.vercel.app/`.
   - Entra directo por `https://empaquetado-merma.vercel.app/`.

## Verificación rápida

- Abre `docs/menu.html` y entra a cada formulario.
- Completa datos mínimos.
- Agrega un par de productos y cantidades.
- Envía y confirma que aparecen filas en las pestañas Empaquetado/Merma de tu hoja.

## Gestión de Productos (Catálogo)

Ahora cuentas con la página `docs/productos.html`:

- Lista todos los productos leyendo la hoja `PRODUCTOS` vía endpoint `?action=productos`.
- Permite agregar un nuevo producto (Código, Descripción, Unidad) enviando `?action=addProduct`.
- Usa la misma URL base (`WEB_APP_URL_DYNAMIC` si está definida en el menú de Ajustes).

### Estructura esperada de la hoja `PRODUCTOS`

Fila 1 (encabezados exactos):

```
Codigo    Descripcion    Unidad
```

Cada fila posterior representa un producto. El backend devolverá objetos:

```jsonc
{ "codigo": "PTEM0001", "descripcion": "PAN DE HAMBURGUESA 85 GR 6 UND", "unidad": "PAQ" }
```

### Código Apps Script (versión completa con formularios y catálogo)

Tu script puede verse así (resumen adaptado). Incluye:

- Inserción de formularios Empaquetado / Merma con prevención de duplicados via `nonce`.
- Catálogo de productos (`?action=productos` y `?action=addProduct`).
- Endpoint de registros recientes (`?action=recent&sheet=Empaquetado&limit=20`).

Puntos a revisar si lo adaptas:
1. Asegura que `SHEETS` mapee las claves enviadas por el frontend (`Empaquetado`, `Merma`) a los nombres EXACTOS de tus pestañas (en mayúsculas si así están). 
2. Añade validación de `adminKey` en `addProductEndpoint` si quieres restringir quién agrega productos. El HTML ya envía `adminKey`.
3. Encabezados de la hoja `PRODUCTOS` usados: `CODIGOS | DESCRIPCION | Unidad_Primaria`.
4. El frontend usa `productos_json` (array de objetos con `codigo, descripcion, unidad, cantidad`).

```javascript
// ===== CONFIG =====
const SPREADSHEET_ID = 'TU_ID_AQUI';
const SHEETS = { Empaquetado: 'EMPAQUETADO', Merma: 'MERMA' };
const PRODUCT_SHEET = 'PRODUCTOS';
const PRODUCT_HEADERS = ['CODIGOS','DESCRIPCION','Unidad_Primaria'];
const ADMIN_KEY = 'PASANTIAS90'; // Cambia esto y el valor en frontend si quieres proteger alta productos
const TZ = 'America/Caracas';
const NONCE_TTL_SECONDS = 3600; // 1h anti duplicados

function doGet(e){
   const p = (e && e.parameter) || {};
   if(p.ping) return respond({ ok:true, pong:p.ping });
   if(p.action==='recent') return recentEndpoint(String(p.sheet||''), clampLimit(p.limit));
   if(p.action==='productos') return getProductosEndpoint();
   return respond({ ok:true, version:'catalogo-v1' });
}

function doPost(e){
   const params = e && e.parameter ? e.parameter : {};
   const action = (params.action||'').trim();
   if(action==='addProduct') return addProductEndpoint(params); // alta catálogo

   // Inserción de formulario
   const sheetKey = String(params.sheet||'').trim();
   const sheetName = SHEETS[sheetKey];
   if(!sheetName) return respond({ ok:false, error:'Parametro sheet inválido' });

   // Anti-duplicado
   const nonce = (params.nonce||'').trim();
   if(nonce){ if(isDuplicateNonce(nonce)) return respond({ ok:true, duplicate:true, nonce }); storeNonce(nonce); }

   const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
   const sh = ss.getSheetByName(sheetName);
   if(!sh) return respond({ ok:false, error:'Hoja no existe: '+sheetName });

   const productos = parseProductos(params.productos_json);
   const marca = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
   let rows=[];
   if(sheetKey==='Empaquetado'){
       const header=['Marca temporal','FECHA','PRODUCTO','CANTIDAD','ENTREGADO A','NUMERO REGISTRO','RESPONSABLE','SEDE'];
       ensureHeader(sh, header);
       const fecha=params.fecha||'', entregado=params.entregado||'', registro=params.registro||'', responsable=params.responsable||'', sede=params.sede||'';
       if(productos.length){ productos.forEach(p=> rows.push([marca, fecha, p.descripcion||p.codigo||'', toNumber(p.cantidad), entregado, registro, responsable, sede])); }
       else rows.push([marca, fecha,'',0,entregado,registro,responsable,sede]);
   } else if(sheetKey==='Merma') {
       const header=['Marca Temporal','FECHA','PRODUCTO','UNIDAD DE MEDIDA','SEDE','MOTIVO DE MERMA','CANTIDAD DEL MOTIVO DE MERMA','NUMERO DE LOTE','RESPONSABLE'];
       ensureHeader(sh, header);
       const fecha=params.fecha||'', sede=params.sede||'', motivo=params.motivo||'', lote=params.lote||'', responsable=params.responsable||'';
       if(productos.length){ productos.forEach(p=> rows.push([marca, fecha, p.descripcion||p.codigo||'', p.unidad||'', sede, motivo, toNumber(p.cantidad), lote, responsable])); }
       else rows.push([marca, fecha,'','',sede,motivo,0,lote,responsable]);
   }
   if(rows.length){ sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
   return respond({ ok:true, inserted: rows.length, nonce });
}

// ===== Catálogo =====
function getProductosEndpoint(){
   const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const sh=ss.getSheetByName(PRODUCT_SHEET);
   if(!sh) return respond({ ok:false, error:'Hoja PRODUCTOS no encontrada' });
   const last=sh.getLastRow(); if(last<2) return respond({ ok:true, products:[] });
   const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
   const idxCodigo=headers.findIndex(h=>String(h).toUpperCase()==='CODIGOS');
   const idxDesc=headers.findIndex(h=>String(h).toUpperCase()==='DESCRIPCION');
   const idxUnidad=headers.findIndex(h=>String(h).toUpperCase()==='UNIDAD_PRIMARIA');
   const values=sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
   const products=values.map(r=>({ codigo:String(r[idxCodigo]||'').trim(), descripcion:String(r[idxDesc]||'').trim(), unidad:String(r[idxUnidad]||'').trim()})).filter(p=>p.codigo||p.descripcion);
   return respond({ ok:true, products });
}

function addProductEndpoint(params){
   const codigo=String(params.codigo||'').trim(); const descripcion=String(params.descripcion||'').trim(); const unidad=(String(params.unidad||'').trim())||'PAQ';
   if(!codigo||!descripcion) return respond({ ok:false, error:'Faltan codigo o descripcion' });
   // Seguridad simple opcional
   if((params.adminKey||'')!==ADMIN_KEY) return respond({ ok:false, error:'adminKey inválido' });
   const ss=SpreadsheetApp.openById(SPREADSHEET_ID); let sh=ss.getSheetByName(PRODUCT_SHEET);
   if(!sh){ sh=ss.insertSheet(PRODUCT_SHEET); sh.getRange(1,1,1,PRODUCT_HEADERS.length).setValues([PRODUCT_HEADERS]); }
   // Duplicado
   const colCod=sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),1).getValues().flat().map(v=>String(v).trim().toLowerCase());
   if(colCod.includes(codigo.toLowerCase())) return respond({ ok:false, error:'Código ya existe' });
   // Asegurar cabecera
   ensureHeader(sh, PRODUCT_HEADERS);
   sh.appendRow([codigo, descripcion, unidad]);
   return respond({ ok:true, inserted:1 });
}

// ===== Recent =====
function recentEndpoint(sheetKey, limit){ /* igual a tu versión extendida */ }
function clampLimit(raw){ var n=parseInt(raw,10); if(!isFinite(n)||n<1)n=20; if(n>200)n=200; return n; }

// ===== Utilidades =====
function parseProductos(str){ try{ const a=JSON.parse(str||'[]'); return Array.isArray(a)?a:[]; }catch(_){ return []; } }
function toNumber(n){ const v=Number(n); return isFinite(v)&&v>0?v:0; }
function ensureHeader(sheet, header){ sheet.getRange(1,1,1,header.length).setValues([header]); }
function respond(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
// Nonce cache igual a tu implementación original (CacheService)
```

### Despliegue del Web App

1. En Apps Script: Deploy > New Deployment > Web App.
2. Description: "API Productos".
3. Execute as: Me (tu cuenta).
4. Who has access: Anyone (o Anyone with the link).
5. Copia la URL que termina en `/exec` y guárdala en Ajustes (menú principal).

### Endpoints usados por el frontend

| Acción    | Método | Query/Body                                                     | Respuesta                        |
|-----------|--------|----------------------------------------------------------------|----------------------------------|
| Listar    | GET    | `?action=productos`                                            | `{ ok:true, products:[...] }`    |
| Ping      | GET    | `?ping=test`                                                   | `{ ok:true, pong:"test" }`       |
| Recent    | GET    | `?action=recent&sheet=Empaquetado&limit=20`                    | `{ ok:true, rows:[...], ... }`   |
| Agregar   | POST   | `action=addProduct`, `codigo`, `descripcion`, `unidad`, `adminKey` | `{ ok:true, inserted:1 }`        |
| Formulario| POST   | Campos + `sheet=Empaquetado|Merma` + `productos_json` + `nonce` | `{ ok:true, inserted:n }`        |

Nota: `adminKey` es comprobada en el backend; cámbiala si lo deseas en ambos lados.

### Flujo para agregar un producto

1. Ve a `productos.html` > "Agregar nuevo".
2. Escribe Código, Descripción y selecciona Unidad.
3. Click en Guardar.
4. Si todo sale bien verás "Guardado" y la tabla se refresca.
5. En los formularios de Empaquetado/Merma ya podrás buscarlo y asignar cantidades.

### Buenas prácticas adicionales (opcional)

- Implementar control de permisos real (OAuth / Sheets service account) si escala más allá del uso interno.
- Añadir columna "Activo" para permitir ocultar productos sin borrarlos.
- Añadir endpoint `updateProduct` para correcciones.


## Notas

- No necesitas Vercel para usar Google Apps Script como backend.
- Si actualizas el código del Apps Script, recuerda desplegar una nueva versión del Web App y actualizar `WEB_APP_URL` si cambia.
- Si prefieres no usar GitHub Pages, puedes publicar `docs/` en cualquier hosting estático.
