// Configura aquí la URL de tu Apps Script Web App (deployment URL que termina en /exec)
// Ejemplo: const WEB_APP_URL = "https://script.google.com/macros/s/AKfycby.../exec";
// Si hay URL guardada en ajustes, úsala; si no, fallback a la fija:
const WEB_APP_URL = (typeof localStorage !== 'undefined' && localStorage.getItem('WEB_APP_URL_DYNAMIC'))
    ? localStorage.getItem('WEB_APP_URL_DYNAMIC')
    : "https://script.google.com/macros/s/AKfycby-3yjw0XHJNRseljo3o4UaLajKem3vk32fSSYqwe-m9zs6jTEKLp0qvI9g60YvKDGavg/exec"; // URL por defecto (deployment actual)

// Endpoints por hoja (el Apps Script espera ?sheet=Empaquetado | ?sheet=Merma)
const APPS_SCRIPT_URL_EMPAQUETADOS = WEB_APP_URL ? WEB_APP_URL + "?sheet=Empaquetado" : "";
const APPS_SCRIPT_URL_MERMA = WEB_APP_URL ? WEB_APP_URL + "?sheet=Merma" : "";

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

function enviarFormulario(formId, url) {
    const form = document.getElementById(formId);
    form.addEventListener("submit", function(e) {
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
                    seleccionadosTmp.push({
                        codigo: inp.dataset.codigo,
                        descripcion: inp.dataset.desc || '',
                        unidad: inp.dataset.unidad || '',
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
                if (msgEl) msgEl.textContent = duplicate ? "Registro ya existente (deduplicado)." : "¡Formulario enviado correctamente!";
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
