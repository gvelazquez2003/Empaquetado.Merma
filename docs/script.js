// Configura aquí la URL de tu Apps Script Web App (deployment URL que termina en /exec)
// Ejemplo: const WEB_APP_URL = "https://script.google.com/macros/s/AKfycby.../exec";
// Si hay URL guardada en ajustes, úsala; si no, fallback a la fija:
const DEFAULT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxm5G6Xq-LU3o-IOUtrWpGO0a4a6832UPC0AcBTFAAwmlEh84goMwVnfs95SRzMG4Vu9A/exec";
const LEGACY_WEB_APP_URLS = new Set([
    "https://script.google.com/macros/s/AKfycbxn-tf1gi_efLVnSgletElQBSMgvn1Ma-SCET5sy48G2QGDpUs93gX5lsRjXA8vsP_9Sg/exec",
    "https://script.google.com/macros/s/AKfycbyNV-0aAlvp3TTfnfiBhGvBeuzlMkSZKl0dOWRkKR8-jBLcmaPs2bnNuF4lYu9k2Yneuw/exec"
]);
function resolveWebAppUrl_(){
    if (typeof localStorage === 'undefined') return DEFAULT_WEB_APP_URL;
    const saved = (localStorage.getItem('WEB_APP_URL_DYNAMIC') || '').trim();
    if (!saved) return DEFAULT_WEB_APP_URL;
    if (LEGACY_WEB_APP_URLS.has(saved)) {
        localStorage.setItem('WEB_APP_URL_DYNAMIC', DEFAULT_WEB_APP_URL);
        return DEFAULT_WEB_APP_URL;
    }
    return saved;
}
const WEB_APP_URL = resolveWebAppUrl_();

// Endpoints por hoja (el Apps Script espera ?sheet=Empaquetado | ?sheet=Merma)
const APPS_SCRIPT_URL_EMPAQUETADOS = WEB_APP_URL ? WEB_APP_URL + "?sheet=Empaquetado" : "";
const APPS_SCRIPT_URL_MERMA = WEB_APP_URL ? WEB_APP_URL + "?sheet=Merma" : "";

const BACKEND_URL = (typeof localStorage !== 'undefined' && localStorage.getItem('BACKEND_URL'))
    ? localStorage.getItem('BACKEND_URL')
    : 'https://almac-n-09.onrender.com';

function isBackendApiUrl(url) {
    const value = String(url || '').trim();
    return /^https?:\/\//i.test(value) && !/script\.google\.com/i.test(value);
}

async function registrarLoteBackend(seleccionados, codigoLote) {
    const productos = (seleccionados || []).map(item => ({
        codigo: item.codigo,
        descripcion: item.descripcion || "",
        cantidad: item.cantidad,
        paquetes: item.paquetes || "",
        sobre_piso: item.sobre_piso || item.sobrePiso || "",
        lote: item.lote || ""
    }));

    if (!productos.length) return { ok: false };

    const payload = { productos };
    if (codigoLote) payload.codigo_lote = codigoLote;

    const response = await fetch(`${BACKEND_URL}/nuevo-lote`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });

    if (response.status === 409) {
        return { ok: true, duplicate: true };
    }

    if (!response.ok) {
        const text = await response.text();
        throw new Error(text || `HTTP ${response.status}`);
    }

    return { ok: true, duplicate: false };
}

function generarNonce() {
    try {
        if (window.crypto && window.crypto.getRandomValues) {
            const arr = new Uint32Array(4);
            window.crypto.getRandomValues(arr);
            return Array.from(arr).map(n => n.toString(16)).join('');
        }
    } catch (_) {}
    return Date.now().toString(36) + Math.random().toString(36).slice(2);
}

function escapeHtml(value) {
    return String(value || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function obtenerResumenConfirmacion(formId, loteGlobal) {
    const read = (id) => {
        const el = document.getElementById(id);
        return el ? String(el.value || '').trim() : '';
    };

    if (formId === 'empaquetados-form') {
        return [
            { label: 'Fecha', value: read('empa-fecha') },
            { label: 'Hora', value: read('empa-hora') },
            { label: 'Maquina', value: read('empa-maquina') },
            { label: 'Lote global', value: loteGlobal || '' },
            { label: 'Entregado a', value: read('empa-entregado') },
            { label: 'Registro', value: read('empa-registro') },
            { label: 'Responsable', value: read('empa-responsable') },
            { label: 'Sede', value: read('empa-sede') }
        ].filter(item => item.value);
    }

    return [
        { label: 'Fecha', value: read('merma-fecha') },
        { label: 'Hora', value: read('merma-hora') },
        { label: 'Responsable', value: read('merma-responsable') },
        { label: 'Sede', value: read('merma-sede') }
    ].filter(item => item.value);
}

function mostrarConfirmacionEnvio(formId, resumen, seleccionados) {
    const overlay = document.getElementById('confirm-overlay');
    const title = document.getElementById('confirm-title');
    const resumenEl = document.getElementById('confirm-resumen');
    const productosEl = document.getElementById('confirm-productos');
    const checkEl = document.getElementById('confirm-check');
    const btnCancel = document.getElementById('confirm-btn-cancel');
    const btnSend = document.getElementById('confirm-btn-send');

    if (!overlay || !title || !resumenEl || !productosEl || !checkEl || !btnCancel || !btnSend) {
        return Promise.resolve(true);
    }

    title.textContent = formId === 'empaquetados-form'
        ? 'Confirmar envio de empaquetados'
        : 'Confirmar envio de merma';

    resumenEl.innerHTML = (resumen || []).map(item => (
        `<div class="confirm-item"><span class="confirm-item-label">${escapeHtml(item.label)}</span><span class="confirm-item-value">${escapeHtml(item.value)}</span></div>`
    )).join('');

    productosEl.innerHTML = (seleccionados || []).map(item => {
        const detalle = formId === 'merma-form'
            ? (item.motivo ? item.motivo : '-')
            : ((item.lote && String(item.lote).trim()) ? item.lote : '-');
        const detalleLabel = formId === 'merma-form' ? 'Motivo' : 'Lote';
        return `<div class="confirm-prod-row">
            <span><strong>${escapeHtml(item.codigo || '')}</strong> - ${escapeHtml(item.descripcion || '')}</span>
            <span>Cant: ${escapeHtml(item.cantidad || '')}</span>
            <span>${detalleLabel}: ${escapeHtml(detalle)}</span>
            <span>Lote: ${escapeHtml(item.lote || '-')}</span>
        </div>`;
    }).join('');

    if (!productosEl.innerHTML) {
        productosEl.innerHTML = '<div class="confirm-prod-row"><span>No hay productos seleccionados.</span></div>';
    }

    checkEl.checked = false;
    btnSend.disabled = true;
    overlay.style.display = 'flex';

    return new Promise((resolve) => {
        const onCheck = () => {
            btnSend.disabled = !checkEl.checked;
        };
        const close = (result) => {
            overlay.style.display = 'none';
            checkEl.removeEventListener('change', onCheck);
            btnCancel.removeEventListener('click', onCancel);
            btnSend.removeEventListener('click', onSend);
            resolve(result);
        };
        const onCancel = () => close(false);
        const onSend = () => {
            if (!checkEl.checked) return;
            close(true);
        };

        checkEl.addEventListener('change', onCheck);
        btnCancel.addEventListener('click', onCancel);
        btnSend.addEventListener('click', onSend);
    });
}

function enviarFormulario(formId, url) {
    const form = document.getElementById(formId);
    form.addEventListener("submit", async function(e) {
        e.preventDefault();
        if (!url) {
            document.getElementById("mensaje").textContent = "Configura la URL del Apps Script (WEB_APP_URL)";
            return;
        }
        // Evitar envíos dobles (doble click, redoble toque)
        if (form.dataset.submitting === "1") {
            return; // ya se está enviando
        }
        form.dataset.submitting = "1";
        const submitBtn = form.querySelector('button[type="submit"]');
        if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = "Enviando..."; }
        const msgEl = document.getElementById("mensaje");
        if (msgEl) msgEl.textContent = "Enviando...";
        const datos = new FormData(form);
        // Lote global para Empaquetado (respaldo si el lote por producto está vacío)
        let loteGlobal = '';
        if (formId === "empaquetados-form") {
            try {
                const preview = document.getElementById('empa-lote-preview');
                if (preview && preview.value) {
                    loteGlobal = String(preview.value).replace(/^Lote:\s*/i, '').trim();
                }
                if (!loteGlobal) {
                    const fechaInput = document.getElementById('empa-fecha');
                    const maqInput = document.getElementById('empa-maquina');
                    const raw = fechaInput ? (fechaInput.value||'').trim() : '';
                    const maq = maqInput ? (maqInput.value||'').trim() : '';
                    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
                        const [y,m,d] = raw.split('-');
                        loteGlobal = `BC${d}${m}${y.slice(2)}${maq}`;
                    }
                }
                if (loteGlobal) datos.append('lote', loteGlobal);
            } catch(_) { /* no-op */ }
        }
        const qtyInputs = form.querySelectorAll('.prod-qty');
        let seleccionados = [];
        // Formatear fecha a dd-mm-aaaa si viene como yyyy-mm-dd
        try {
            const fechaInput = form.querySelector('input[name="fecha"]');
            const raw = fechaInput ? (fechaInput.value||'').trim() : '';
            if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
                const [y,m,d] = raw.split('-');
                datos.set('fecha', `${d}-${m}-${y}`);
            }
        } catch(_) { /* no-op */ }
        // Idempotencia: token anti-duplicado (reutiliza el mismo nonce en reintentos)
        let nonce = form.dataset.nonce || localStorage.getItem(`nonce_${formId}`) || '';
        if (!nonce) {
            nonce = generarNonce();
            form.dataset.nonce = nonce;
            try { localStorage.setItem(`nonce_${formId}`, nonce); } catch(_) {}
        }
        datos.append('nonce', nonce);
        // Agregar solo los productos con cantidad > 0 como JSON
        try {
            const seleccionadosTmp = [];
            qtyInputs.forEach(inp => {
                const val = parseInt(inp.value, 10);
                if (!isNaN(val) && val > 0) {
                    const row = inp.closest('.producto-line');
                    const motivoEl = row ? row.querySelector('.merma-motivo') : null;
                    const loteEl = row ? row.querySelector('.merma-lote, .empa-lote') : null;
                    // read motivo and lote robustly: prefer value, fallback to selected option text
                    var motivoVal = '';
                    if (motivoEl) {
                        try {
                            motivoVal = (motivoEl.value || '').toString().trim();
                        } catch(_) { motivoVal = ''; }
                        try {
                            if (!motivoVal && typeof motivoEl.selectedIndex === 'number' && motivoEl.selectedIndex >= 0) {
                                var opt = motivoEl.options[motivoEl.selectedIndex];
                                motivoVal = (opt && (opt.value || opt.text) || '').toString().trim();
                            }
                        } catch(_) {}
                    }
                    var loteVal = '';
                    if (loteEl) {
                        try { loteVal = (loteEl.value || '').toString().trim(); } catch(_) { loteVal = ''; }
                    }
                    if (!loteVal && loteGlobal) loteVal = loteGlobal;
                    seleccionadosTmp.push({
                        codigo: inp.dataset.codigo,
                        descripcion: inp.dataset.desc || '',
                        unidad: inp.dataset.unidad || '',
                        paquetes: inp.dataset.paquetes || '',
                        cestas: inp.dataset.cestas || '',
                        sobre_piso: inp.dataset.sobrePiso || '',
                        cantidad: val,
                        motivo: motivoVal,
                        lote: loteVal
                    });
                }
            });
            seleccionados = seleccionadosTmp;
            // Validar motivo y lote en Merma
            if (formId === "merma-form") {
                const falta = seleccionados.find(it => !String(it.motivo || '').trim() || !String(it.lote || '').trim());
                if (falta) {
                    if (msgEl) msgEl.textContent = "Completa el motivo y el número de lote en todos los productos.";
                    form.dataset.submitting = "0";
                    if (submitBtn) {
                        submitBtn.disabled = false;
                        submitBtn.textContent = "Enviar";
                    }
                    return;
                }
            }
            // Evitar duplicados de producto + lote (en Merma permite repetir lote si el motivo es distinto)
            const dupMap = new Set();
            let hasDup = false;
            const isMermaForm = formId === "merma-form";
            seleccionados.forEach(item => {
                const codigo = (item.codigo || '').trim().toLowerCase();
                const lote = (item.lote || '').trim().toLowerCase();
                const motivo = (item.motivo || '').trim().toLowerCase();
                if (!codigo) return;
                const key = isMermaForm ? (codigo + '|' + lote + '|' + motivo) : (codigo + '|' + lote);
                if (dupMap.has(key)) hasDup = true;
                else dupMap.add(key);
            });
            if (hasDup) {
                if (msgEl) msgEl.textContent = isMermaForm
                    ? "No se permite el mismo producto con el mismo lote y motivo."
                    : "No se permite el mismo producto con el mismo número de lote.";
                form.dataset.submitting = "0";
                if (submitBtn) {
                    submitBtn.disabled = false;
                    submitBtn.textContent = "Enviar";
                }
                return;
            }
            if (seleccionados.length) {
                datos.append('productos_json', JSON.stringify(seleccionados));
                if (formId === "merma-form") {
                    // Captura redundante para Apps Script en caso de que el JSON se recorte en tránsito
                    datos.append('productos_count', String(seleccionados.length));
                    seleccionados.forEach((item, idx) => {
                        datos.append(`prodCodigo_${idx}`, item.codigo || '');
                        datos.append(`motivo_${idx}`, item.motivo || '');
                        datos.append(`lote_${idx}`, item.lote || '');
                    });
                }
            }
            // Identificar a qué hoja va (para depuración opcional en backend)
            if (url.includes('Empaquetado')) datos.append('sheet', 'Empaquetado');
            if (url.includes('Merma')) datos.append('sheet', 'Merma');
        } catch(_) { /* no-op */ }
        if (!seleccionados.length) {
            if (msgEl) msgEl.textContent = "Agrega al menos un producto con cantidad.";
            form.dataset.submitting = "0";
            if (submitBtn) {
                submitBtn.disabled = false;
                submitBtn.textContent = "Enviar";
            }
            return;
        }

        if (form.dataset.confirmedSubmission !== '1') {
            const resumen = obtenerResumenConfirmacion(formId, loteGlobal);
            form.dataset.submitting = "0";
            if (submitBtn) {
                submitBtn.disabled = false;
                submitBtn.textContent = "Enviar";
            }
            if (msgEl) msgEl.textContent = "";
            const confirmado = await mostrarConfirmacionEnvio(formId, resumen, seleccionados);
            if (!confirmado) {
                return;
            }
            form.dataset.confirmedSubmission = '1';
            form.requestSubmit();
            return;
        }

        delete form.dataset.confirmedSubmission;
        if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = "Enviando..."; }
        if (msgEl) msgEl.textContent = "Enviando...";

        let backendSyncStatus = 'not-applicable';

        fetch(url, {
            method: "POST",
            body: datos
        })
        .then(async (response) => {
            let txt;
            try { txt = await response.text(); } catch(_) { txt = ''; }
            let ok = response.ok;
            let duplicate = false;
            let errorMsg = '';
            let parsed = null;
            try {
                parsed = JSON.parse(txt);
                if (parsed.ok !== undefined) ok = parsed.ok;
                duplicate = !!parsed.duplicate;
                if (!ok && parsed.error) errorMsg = parsed.error;
            } catch(e) {
                // texto no JSON; mantener valores por defecto
            }
            // Log detallado para depuración
            try { console.log('[ENVIAR_FORM]', formId, 'status:', response.status, 'okFlag:', ok, 'duplicate:', duplicate, 'raw:', txt); } catch(_) {}
            // Consideramos éxito también si la respuesta es no legible pero status 200 (opaque redirect no-cors)
            if (ok || response.status === 0) {
                if (formId === "empaquetados-form" && isBackendApiUrl(BACKEND_URL)) {
                    try {
                        const backendResult = await registrarLoteBackend(seleccionados, loteGlobal);
                        backendSyncStatus = backendResult && backendResult.ok ? 'ok' : 'failed';
                    } catch (backendError) {
                        backendSyncStatus = 'failed';
                        try { console.error('[BACKEND_SYNC_ERROR]', backendError); } catch(_) {}
                    }
                }

                if (msgEl) {
                    if (duplicate) {
                        msgEl.textContent = "Registro ya existente (deduplicado).";
                    } else if (formId === "empaquetados-form") {
                        msgEl.textContent = backendSyncStatus === 'ok'
                            ? "¡Formulario enviado! Registro disponible para validación en Almacen09."
                            : "¡Formulario enviado! Verifica sincronización con Almacen09.";
                    } else {
                        msgEl.textContent = "¡Formulario enviado correctamente!";
                    }
                }
                // Disparar evento para página de registros
                try {
                    const insertedCount = Array.from(form.querySelectorAll('.prod-qty')).filter(inp => parseInt(inp.value,10)>0).length;
                    const hoja = url.includes('Empaquetado') ? 'Empaquetado' : (url.includes('Merma') ? 'Merma' : '');
                    window.dispatchEvent(new CustomEvent('registroInsertado',{ detail:{ sheet:hoja, productos:insertedCount, nonce: form.dataset.nonce || '' }}));
                } catch(_) {}
                form.reset();
                const qtyInputs = form.querySelectorAll('.prod-qty');
                qtyInputs.forEach(i => i.value = "");
                const contenedores = form.querySelectorAll('.seleccionados');
                contenedores.forEach(c => c.innerHTML = "");
                delete form.dataset.nonce;
                try { localStorage.removeItem(`nonce_${formId}`); } catch(_) {}
                setTimeout(() => { if (msgEl) msgEl.textContent = ""; }, 3000);
            } else {
                // Mostrar mensaje específico si lo tenemos
                if (!errorMsg && parsed && !parsed.ok && !parsed.error) {
                    errorMsg = 'Error desconocido (respuesta JSON sin ok=true).';
                }
                if (!errorMsg && !parsed && response.status !== 200) {
                    errorMsg = 'HTTP '+response.status+' sin detalle del servidor.';
                }
                let debugMsg = '';
                try {
                    if (parsed && parsed.debug) {
                        const d = parsed.debug;
                        const ent = d.entradas09;
                        const emp = d.empaquetado;
                        const entInfo = ent ? `Entradas09: lastCol=${ent.lastCol}, maxCols=${ent.maxCols}, tablas=${Array.isArray(ent.tables)?ent.tables.length:'?'}; ` : '';
                        const empInfo = emp ? `Empaquetado: lastCol=${emp.lastCol}, maxCols=${emp.maxCols}, tablas=${Array.isArray(emp.tables)?emp.tables.length:'?'}; ` : '';
                        debugMsg = entInfo || empInfo ? (` Diagnóstico: ${entInfo}${empInfo}`) : '';
                    }
                } catch(_){ }
                if (msgEl) msgEl.textContent = "Error al enviar el formulario. " + (errorMsg ? ("Detalle: "+ errorMsg) : "Puedes reintentar.") + debugMsg;
            }
        })
        .catch(error => {
            // Fallback: asumimos que puede haber sido un bloqueo de lectura pero el backend insertó la fila.
            if (msgEl) msgEl.textContent = "Posible envío exitoso (respuesta no legible). Verifica en la hoja. Si falta, reintenta.";
            try { console.error('[ENVIAR_FORM][ERROR]', formId, error); } catch(_) {}
            // No limpiamos por si realmente no llegó; conservamos nonce para reintentar.
        })
        .finally(() => {
            // Pequeño enfriamiento para evitar reenvío inmediato
            setTimeout(() => {
                form.dataset.submitting = "0";
                const btn = form.querySelector('button[type="submit"]');
                if (btn) {
                    btn.disabled = false;
                    // Si hay nonce activo, ofrecer reintento; si no, volver a "Enviar"
                    btn.textContent = (form.dataset.nonce || localStorage.getItem(`nonce_${formId}`)) ? "Reintentar" : "Enviar";
                }
            }, 800);
        });
    });
}

// Limpieza manual de formulario
function clearForm(formId){
    const form = document.getElementById(formId);
    if(!form) return;
    form.reset();
    // Limpiar cantidades y contenedores de productos seleccionados
    const qtyInputs = form.querySelectorAll('.prod-qty');
    qtyInputs.forEach(i => i.value = "");
    const contenedores = form.querySelectorAll('.seleccionados');
    contenedores.forEach(c => c.innerHTML = "");
    // Limpiar nonce para permitir nuevo envío independiente
    delete form.dataset.nonce;
    try { localStorage.removeItem(`nonce_${formId}`); } catch(_) {}
    const msgEl = document.getElementById('mensaje');
    if (msgEl) {
        msgEl.textContent = 'Formulario limpiado.';
        setTimeout(()=>{ if(msgEl.textContent==='Formulario limpiado.') msgEl.textContent=''; },2000);
    }
    // Restaurar texto del botón si estaba en otro estado
    const btn = form.querySelector('button[type="submit"]');
    if(btn) btn.textContent = 'Enviar';
}

enviarFormulario("empaquetados-form", APPS_SCRIPT_URL_EMPAQUETADOS);
enviarFormulario("merma-form", APPS_SCRIPT_URL_MERMA);
