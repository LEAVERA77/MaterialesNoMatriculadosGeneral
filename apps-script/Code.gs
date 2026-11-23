/**
 * ENERSA · Gestión de Materiales — Google Apps Script
 * BACKEND PROPIO de LEAVERA77/MaterialesNoMatriculadosGeneral
 * ------------------------------------------------------------------
 * IMPORTANTE — Este archivo es el backend de ESTE repositorio.
 * NO modifica ni depende de LEAVERA77/MaterialesNoMatriculados (producción).
 *
 * Usá un proyecto de Apps Script NUEVO (no el de producción), así producción
 * y este entorno quedan totalmente separados. Cada uno tiene su propia URL /exec.
 *
 * Reemplaza el contenido de tu proyecto de Apps Script por este archivo.
 *
 * Qué hace:
 *   1) Recibe el POST de la aplicación (index.html).
 *   2) Si viene un movimiento (retiro / devolución), dispara la GitHub Action
 *      que actualiza listado_materiales_ENERSA.xlsx, enviando también el
 *      LEGAJO y los datos del DISPOSITIVO (ID + huella SHA-256).
 *   3) Si viene {accion:'pin_establecer'} (alta del PIN en el primer ingreso)
 *      o {accion:'pin_resetear'} (restablecimiento por un administrador),
 *      dispara la misma Action para que se escriba en la hoja "Empleados".
 *      El PIN nunca viaja en claro: se envía SAL + HASH SHA-256.
 *   4) Si viene {accion:'login'}, valida legajo + PIN contra la hoja
 *      "Empleados" y responde si las credenciales son válidas. Esto permite
 *      activar la validación del lado del servidor en la web
 *      (const USAR_VALIDACION_SERVIDOR = true).
 *
 * CONFIGURACIÓN — completá las 4 constantes de abajo y volvé a implementar
 * (Implementar > Administrar implementaciones > editar > Nueva versión).
 */

const GH_OWNER = 'LEAVERA77';
const GH_REPO = 'MaterialesNoMatriculadosGeneral';  // backend propio (producción NO se toca)
const GH_TOKEN = 'PEGAR_AQUI_EL_TOKEN';           // fine-grained PAT con permiso "Contents: Read and write"
const SPREADSHEET_ID = '';                        // VACÍO = sin respaldo en Sheets. NO poner el de producción.

// ------------------------------------------------------------------
// MODO SEGURO — red de protección contra tocar producción.
// Si GH_REPO apunta al repositorio de producción, el script NO ejecuta
// ninguna acción y devuelve un error explicando el motivo.
// ------------------------------------------------------------------
const REPO_PRODUCCION = 'MaterialesNoMatriculados';   // sin el sufijo "General"

// Acciones de gestión de legajos (las valida el script de la Action).
const ACCIONES_EMPLEADO = ['empleado_alta', 'empleado_baja', 'empleado_activar', 'empleado_eliminar'];

// Sello de versión: permite saber QUÉ código está corriendo el /exec.
// Si no coincide con el de apps-script/Code.gs del repo, el deploy está viejo.
const BUILD = 'general-2026-11-23-r5';

/**
 * Estado del token: UNA sola verificación para todo el archivo (así el
 * diagnóstico no puede contradecirse) + saneado de lo pegado a mano
 * (comillas, espacios, saltos de línea, BOM) que es la causa más común
 * de "falta_token" con el token ya cargado.
 */
function estadoToken_() {
  var bruto = (GH_TOKEN === null || GH_TOKEN === undefined) ? '' : String(GH_TOKEN);
  var limpio = bruto
    .replace(/^\uFEFF/, '')        // BOM invisible al inicio
    .replace(/\s+/g, '')           // espacios y saltos de línea
    .replace(/["']/g, '');          // comillas pegadas por error

  var esPlaceholder = /^PEGAR_AQUI/.test(limpio);
  // Un PAT fine-grained ronda los 93 caracteres; uno clásico (ghp_) los 40.
  // Los largos cortos delatan un pegado incompleto, el error más común.
  var formatoOk = /^github_pat_[A-Za-z0-9_]{50,}$/.test(limpio) ||
                  /^ghp_[A-Za-z0-9]{36}$/.test(limpio);

  return {
    limpio: limpio,
    ok: !!limpio && !esPlaceholder && formatoOk,
    largo: limpio.length,
    pista: limpio ? (limpio.slice(0, 12) + '…' + limpio.slice(-4)) : '',
    formato: esPlaceholder ? 'placeholder' : (formatoOk ? 'valido' : (limpio ? 'formato_inesperado' : 'vacio')),
    fue_saneado: (limpio !== bruto)
  };
}

/** Devuelve {ok, motivo, detalle} si la configuración es insegura o incompleta. */
function estadoConfiguracion_() {
  if (!GH_REPO || GH_REPO === REPO_PRODUCCION) {
    return {
      ok: false,
      motivo: 'modo_seguro',
      detalle: 'GH_REPO apunta al repositorio de PRODUCCIÓN ("' + GH_REPO + '"). ' +
               'Este script es para MaterialesNoMatriculadosGeneral. No se envió nada.'
    };
  }

  var tok = estadoToken_();
  if (!tok.ok) {
    return {
      ok: false,
      motivo: 'falta_token',
      detalle: 'El token no está usable: estado "' + tok.formato + '", largo ' + tok.largo +
               ' (un PAT fine-grained ronda los 90 caracteres y empieza con github_pat_).'
    };
  }

  return { ok: true };
}

/** Diagnóstico completo, sin exponer el token. Lo usan doGet y verificar(). */
function diagnostico_() {
  var config = estadoConfiguracion_();
  var tok = estadoToken_();
  return {
    ok: config.ok,
    build: BUILD,
    servicio: 'ENERSA materiales',
    version: '2.0',
    entorno: 'backend propio',
    repo_destino: GH_OWNER + '/' + GH_REPO,
    modo_seguro: (GH_REPO === REPO_PRODUCCION) ? 'BLOQUEADO (es producción)' : 'ok',
    token_cargado: tok.ok,
    token_detalle: { largo: tok.largo, formato: tok.formato, pista: tok.pista, fue_saneado: tok.fue_saneado },
    planilla_configurada: !!SPREADSHEET_ID,
    motivo: config.ok ? '' : config.motivo,
    detalle: config.ok ? '' : config.detalle
  };
}

/**
 * GET de diagnóstico. Abrí la URL /exec en el navegador para ver a qué
 * repositorio apunta este script y qué versión de código está corriendo.
 */
function doGet() {
  return json_(diagnostico_());
}

/**
 * Ejecutable desde el editor (Implementar no: ▶ Ejecutar con "verificar"
 * seleccionado). Deja el informe en el registro de ejecuciones.
 */
function verificar() {
  var informe = JSON.stringify(diagnostico_(), null, 2);
  Logger.log(informe);
  return informe;
}

/**
 * POST de la aplicación.
 */
function doPost(e) {
  try {
    var datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // --- 0. Modo seguro: nunca disparar el repo de producción ---
    var config = estadoConfiguracion_();
    if (!config.ok) {
      Logger.log('BLOQUEADO: ' + config.detalle);
      return json_({ ok: false, motivo: config.motivo, detalle: config.detalle });
    }

    // --- 1. Validación de identidad (opcional, del lado del servidor) ---
    if (datos.accion === 'login') {
      return json_(validarLogin_(datos.legajo, datos.pin));
    }

    // --- 1b. Alta / restablecimiento de PIN (el Excel lo escribe la Action) ---
    if (datos.accion === 'pin_establecer' || datos.accion === 'pin_resetear') {
      if (!datos.legajo) return json_({ ok: false, motivo: 'falta_legajo' });
      var okPin = dispararGitHub_(datos);
      return json_({ ok: okPin, accion: datos.accion, legajo: datos.legajo });
    }

    // --- 1c. Gestión de legajos autorizados (solo administradores) ---
    if (ACCIONES_EMPLEADO.indexOf(datos.accion) >= 0) {
      if (!datos.legajo) return json_({ ok: false, motivo: 'falta_legajo' });
      if (!datos.admin_legajo) return json_({ ok: false, motivo: 'falta_admin' });
      var okEmp = dispararGitHub_(datos);
      return json_({ ok: okEmp, accion: datos.accion, legajo: datos.legajo });
    }

    // --- 2. Movimiento de materiales ---
    if (!datos.items || !datos.operacion) {
      return json_({ ok: false, motivo: 'payload_invalido' });
    }

    // Copia de respaldo en la planilla de Google (si está configurada)
    guardarEnPlanilla_(datos);

    // Disparo de la GitHub Action (es la que modifica el Excel del repo)
    var ok = dispararGitHub_(datos);

    return json_({ ok: ok, registrado: true, operacion: datos.operacion });
  } catch (err) {
    return json_({ ok: false, motivo: 'error', detalle: String(err) });
  }
}


/**
 * Valida LEGAJO + PIN contra la hoja "Empleados" (col A legajo, B nombre, C pin o hash, D activo).
 * Devuelve {ok:true, nombre} o {ok:false, motivo:'credenciales'}.
 */
function validarLogin_(legajo, pin) {
  if (!SPREADSHEET_ID) return { ok: false, motivo: 'metodo_no_soportado' };
  var legajoNorm = String(legajo || '').replace(/[^0-9A-Za-z]/g, '').toUpperCase();
  var pinNorm = String(pin || '').trim();
  if (!legajoNorm || !pinNorm) return { ok: false, motivo: 'credenciales' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hoja = ss.getSheetByName('Empleados');
  if (!hoja) return { ok: false, motivo: 'metodo_no_soportado' };

  var filas = hoja.getDataRange().getValues();
  for (var i = 1; i < filas.length; i++) {
    var f = filas[i];
    if (String(f[0] || '').replace(/[^0-9A-Za-z]/g, '').toUpperCase() !== legajoNorm) continue;

    var activo = String(f[3] === undefined || f[3] === '' ? 'SI' : f[3]).trim().toUpperCase();
    if (activo === 'NO' || activo === 'FALSE' || activo === '0') {
      return { ok: false, motivo: 'inactivo' };
    }

    var guardado = String(f[2] === undefined || f[2] === null ? '' : f[2]).trim();
    var valido = false;
    if (guardado.indexOf('$') > 0) {
      // Formato actual: "SAL$HASH_SHA256" (hash de SAL|LEGAJO|PIN)
      var partes = guardado.split('$');
      valido = (hashSha256_(partes[0] + '|' + legajoNorm + '|' + pinNorm) === String(partes[1]).toLowerCase());
    } else if (/^[a-f0-9]{64}$/i.test(guardado)) {
      valido = (hashSha256_(pinNorm) === guardado.toLowerCase());
    } else if (guardado) {
      valido = (guardado === pinNorm) || (Number(guardado) === Number(pinNorm));
    }
    return valido ? { ok: true, nombre: String(f[1] || '').trim() } : { ok: false, motivo: 'credenciales' };
  }
  return { ok: false, motivo: 'credenciales' };
}

function hashSha256_(texto) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/** Respaldo del movimiento en la planilla (no afecta el Excel del repo). */
function guardarEnPlanilla_(datos) {
  if (!SPREADSHEET_ID) return;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hoja = ss.getSheetByName('Historial') || ss.insertSheet('Historial');
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(['Fecha', 'Usuario', 'Operacion', 'Material', 'Cantidad', 'Uso', 'Legajo', 'Dispositivo', 'Huella']);
    }
    (datos.items || []).forEach(function (item) {
      hoja.appendRow([
        datos.fecha || new Date(), datos.usuario || '', String(datos.operacion || '').toUpperCase(),
        item.material || '', item.cantidad || 0, datos.uso || '',
        datos.legajo || '', datos.dispositivo_id || '', datos.dispositivo_huella || ''
      ]);
    });
  } catch (err) {
    Logger.log('No se pudo guardar el respaldo: ' + err);
  }
}

/** Dispara la GitHub Action "Actualizar Materiales". */
function dispararGitHub_(datos) {
  var url = 'https://api.github.com/repos/' + GH_OWNER + '/' + GH_REPO + '/dispatches';
  // IMPORTANTE: GitHub acepta como máximo 10 propiedades en client_payload.
  // Por eso los datos del equipo van agrupados en "dispositivo" (y no planos).
  var payload = {
    event_type: 'actualizar_stock',
    client_payload: {
      fecha: datos.fecha,                 // 1
      usuario: datos.usuario,             // 2
      legajo: datos.legajo || '',         // 3
      operacion: datos.operacion,         // 4
      uso: datos.uso,                     // 5
      items: datos.items,                 // 6
      dispositivo: (function () {
        // La app ya manda el objeto anidado; si viene plano (clientes viejos)
        // se arma igual para que el Excel no pierda trazabilidad.
        var d = datos.dispositivo || {};
        return {
          id: d.id || datos.dispositivo_id || '',
          huella: d.huella || datos.dispositivo_huella || '',
          resumen: d.resumen || datos.dispositivo_resumen || '',
          ua: d.ua || datos.dispositivo_ua || '',
          pantalla: d.pantalla || datos.dispositivo_pantalla || '',
          zona: d.zona || datos.zona_horaria || '',
          timestamp: d.timestamp || datos.timestamp_cliente || ''
        };
      })(),                               // 7
      origen: datos.origen || ''          // 8
    }
  };

  // Gestión de legajos: claves propias (8, dentro del límite de 10).
  if (ACCIONES_EMPLEADO.indexOf(datos.accion) >= 0) {
    payload.client_payload = {
      accion: datos.accion,                          // 1
      admin_legajo: datos.admin_legajo || '',        // 2
      legajo: datos.legajo || '',                    // 3
      nombre: datos.nombre || '',                    // 4
      usuario: datos.usuario || '',                   // 5
      operacion: datos.operacion || '',               // 6
      fecha: datos.fecha || new Date().toString(),    // 7
      dispositivo: (function () {
        var d = datos.dispositivo || {};
        return {
          id: d.id || '', huella: d.huella || '', resumen: d.resumen || '',
          ua: d.ua || '', pantalla: d.pantalla || '', zona: d.zona || '',
          timestamp: d.timestamp || ''
        };
      })()                                            // 8
    };
  }

  // Las acciones de PIN usan su propio conjunto de claves (siempre <= 10).
  if (datos.accion === 'pin_establecer' || datos.accion === 'pin_resetear') {
    payload.client_payload = {
      accion: datos.accion,                         // 1
      legajo: datos.legajo || '',                   // 2
      sal: datos.sal || '',                         // 3
      pin_hash: datos.pin_hash || '',               // 4
      admin_legajo: datos.admin_legajo || '',       // 5
      usuario: datos.usuario || '',                 // 6
      operacion: datos.operacion || '',             // 7
      fecha: datos.fecha || new Date().toString(),  // 8
      dispositivo: (function () {
        var d = datos.dispositivo || {};
        return {
          id: d.id || '', huella: d.huella || '', resumen: d.resumen || '',
          ua: d.ua || '', pantalla: d.pantalla || '', zona: d.zona || '',
          timestamp: d.timestamp || ''
        };
      })()                                          // 9
    };
  }

  // Token saneado: si se pegó con saltos de línea o comillas, el header
  // Authorization falla con 401 aunque el token sea correcto.
  var tokenHeader = estadoToken_().limpio;

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + tokenHeader,
      Accept: 'application/vnd.github+json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  if (code !== 204 && code !== 200) {
    Logger.log('GitHub respondió ' + code + ': ' + res.getContentText());
    return false;
  }
  return true;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
