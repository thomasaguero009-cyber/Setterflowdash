/**
 * SetterFlow · Factor Studios
 * Recibe los datos del formulario de Métricas y los escribe en la hoja
 * individual de cada setter (THOMI, FLOR, VALERIA, FRANCO).
 * También maneja registro/login de usuarios individuales para Métricas
 * (hoja "Usuarios": Nombre, Password — protección básica, no es seguridad
 * de nivel bancario, es solo para que cada persona tenga su propio acceso).
 *
 * Estructura real de cada hoja de setter:
 * B: Fecha (dd/mm/aaaa)  C: Seguimientos  D: Cash collected
 * E: Horas trabajadas    F: Agendas       G: Tasa de conversión
 * H: Leads
 * (fila 5 = TOTALES/PROMEDIOS, los datos empiezan en la fila 6)
 *
 * OJO: la columna H (Leads) es nueva — hay que agregarle el título "Leads"
 * a mano en la fila de encabezados de CADA una de las 4 hojas (THOMI (1),
 * FLOR (1), VALERIA, FRANCO), igual que ya dice "Seguimientos"/"Cash"/etc
 * en las columnas C-G. El dashboard busca la columna por el texto del
 * encabezado (no por letra fija), así que sin ese título no la va a
 * encontrar.
 *
 * INSTALACIÓN:
 * 1. Pegá este código en Extensiones > Apps Script (dentro de tu Google Sheet real).
 * 2. Guardá y Deployá como Web App (ver instrucciones en el chat).
 */

const SHEET_NAMES = {
  thomi:   "THOMI (1)",
  flor:    "FLOR (1)",
  valeria: "VALERIA",
  franco:  "FRANCO",
  paula:   "PAULA"
};

const COLS = {
  fecha: 2,        // B
  seguimientos: 3, // C
  cash: 4,         // D
  horas: 5,        // E
  agendas: 6,      // F
  conversion: 7,   // G
  leads: 8         // H
};

const FIRST_DATA_ROW = 6; // fila 5 es TOTALES/PROMEDIOS
const USERS_SHEET_NAME = "Usuarios";

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // Aviso (webhook) de Hyros: formato distinto al de nuestra propia
    // extensión/dashboard — trae "type" y "subscriptionId", nunca
    // "action". Se lo detecta acá arriba de todo, antes del switch normal.
    if (body.type && body.subscriptionId) {
      console.log('doPost — webhook de Hyros: ' + body.type);
      return handleHyrosWebhook(body);
    }

    const action = body.action || 'save';
    // Registro liviano de qué está pidiendo cada llamada — para poder ver
    // en "Ejecuciones" (script.google.com → ícono de lista con flecha de
    // play) qué se estuvo llamando, sin tener que adivinar. No guarda
    // nada sensible, solo el nombre de la acción.
    console.log('doPost — action: ' + action + (action === 'fetchhyros' ? ' (fecha: ' + body.fecha + (body.fechaHasta ? ' a ' + body.fechaHasta : '') + ')' : ''));

    if (action === 'register') return handleRegister(body);
    if (action === 'login') return handleLogin(body);
    if (action === 'changepassword') return handleChangePassword(body);
    if (action === 'updatecorreo') return handleUpdateCorreo(body);
    if (action === 'updateprofile') return handleUpdateProfile(body);
    if (action === 'updatemeta') return handleUpdateMeta(body);
    if (action === 'updateavatar') return handleUpdateAvatar(body);
    if (action === 'updatecover') return handleUpdateCover(body);
    if (action === 'listperfilespublicos') return handleListPerfilesPublicos(body);
    if (action === 'listusuarios') return handleListUsuarios(body);
    if (action === 'setestado') return handleSetEstado(body);
    if (action === 'setrol') return handleSetRol(body);
    if (action === 'getdiagram') return handleGetDiagram(body);
    if (action === 'savediagram') return handleSaveDiagram(body);
    if (action === 'getsharedcontent') return handleGetSharedContent(body);
    if (action === 'savesharedcontent') return handleSaveSharedContent(body);
    if (action === 'incrementseguimiento') return handleIncrementSeguimiento(body);
    if (action === 'incrementhoras') return handleIncrementHoras(body);
    if (action === 'reportaractividadgeneral') return handleReportarActividadGeneral(body);
    if (action === 'fetchhyros') return handleFetchHyros(body);
    if (action === 'gettagcolors') return handleGetTagColors(body);
    if (action === 'savetagcolors') return handleSaveTagColors(body);

    // acción por defecto: guardar métricas (compatibilidad con versiones anteriores)
    return handleSaveMetrics(body);

  } catch (err) {
    console.log('ERROR en doPost:', err.message, err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getUsersSheetWithCorreo(ss) {
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 11).setValues([['Nombre', 'Password', 'Correo', 'Estado', 'Rol', 'Apodo', 'Telefono', 'Avatar', 'Cover', 'Pensamiento', 'NombreMostrar']]);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
  } else if (sheet.getLastColumn() < 11) {
    if (sheet.getLastColumn() < 3) { sheet.getRange(1, 3).setValue('Correo'); }
    if (sheet.getLastColumn() < 4) { sheet.getRange(1, 4).setValue('Estado'); }
    if (sheet.getLastColumn() < 5) { sheet.getRange(1, 5).setValue('Rol'); }
    if (sheet.getLastColumn() < 6) { sheet.getRange(1, 6).setValue('Apodo'); }
    if (sheet.getLastColumn() < 7) { sheet.getRange(1, 7).setValue('Telefono'); }
    if (sheet.getLastColumn() < 8) { sheet.getRange(1, 8).setValue('Avatar'); }
    if (sheet.getLastColumn() < 9) { sheet.getRange(1, 9).setValue('Cover'); }
    if (sheet.getLastColumn() < 10) { sheet.getRange(1, 10).setValue('Pensamiento'); }
    sheet.getRange(1, 11).setValue('NombreMostrar');
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    // A cualquier usuario viejo sin Estado/Rol, lo dejamos Aprobado + Setter para no romper accesos existentes
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const datos = sheet.getRange(2, 4, lastRow - 1, 2).getValues();
      for (let i = 0; i < datos.length; i++) {
        if (!datos[i][0]) sheet.getRange(i + 2, 4).setValue('Aprobado');
        if (!datos[i][1]) sheet.getRange(i + 2, 5).setValue('Setter');
      }
    }
  }
  return sheet;
}

function handleUpdateCover(body) {
  const nombre = (body.nombre || '').toString().trim();
  const cover = (body.cover || '').toString();
  if (!nombre) return jsonOut({ status: "error", message: "Falta nombre" });
  if (cover.length > 45000) return jsonOut({ status: "error", message: "La imagen es demasiado grande para sincronizar" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      sheet.getRange(idx + 2, 9).setValue(cover);
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleUpdateAvatar(body) {
  const nombre = (body.nombre || '').toString().trim();
  const avatar = (body.avatar || '').toString();
  if (!nombre) return jsonOut({ status: "error", message: "Falta nombre" });
  if (avatar.length > 45000) return jsonOut({ status: "error", message: "La imagen es demasiado grande para sincronizar" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      sheet.getRange(idx + 2, 8).setValue(avatar);
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleListPerfilesPublicos(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  let rows = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    rows = data
      .filter(r => (r[3] || 'Aprobado') === 'Aprobado')
      .map(r => ({ nombre: r[0], apodo: r[5] || '', avatar: r[7] || '', cover: r[8] || '', pensamiento: r[9] || '', nombreMostrar: r[10] || '' }));
  }
  return jsonOut({ status: "ok", rows });
}

// Meta de facturación PERSONAL — cada setter la pone a su gusto en "Mi
// Perfil" (no tiene nada que ver con la meta compartida del equipo de
// $50.000 que se usa en el dashboard general). Columna L (12), libre hasta
// ahora en la hoja "Usuarios".
function handleUpdateMeta(body) {
  const nombre = (body.nombre || '').toString().trim();
  const metaFacturacion = Number(body.metaFacturacion);
  if (!nombre || !metaFacturacion || metaFacturacion <= 0) {
    return jsonOut({ status: "error", message: "Falta nombre o el monto no es válido" });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      sheet.getRange(idx + 2, 12).setValue(metaFacturacion);
      return jsonOut({ status: "ok", metaFacturacion });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleUpdateProfile(body) {
  const nombre = (body.nombre || '').toString().trim();
  const correo = (body.correo || '').toString().trim();
  const apodo = (body.apodo || '').toString().trim();
  const telefono = (body.telefono || '').toString().trim();
  const pensamiento = body.pensamiento !== undefined ? (body.pensamiento || '').toString().trim() : null;
  const nombreMostrar = body.nombreMostrar !== undefined ? (body.nombreMostrar || '').toString().trim() : null;
  if (!nombre) return jsonOut({ status: "error", message: "Falta nombre" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      const rowNum = idx + 2;
      sheet.getRange(rowNum, 3).setValue(correo);
      sheet.getRange(rowNum, 6).setValue(apodo);
      sheet.getRange(rowNum, 7).setValue(telefono);
      if (pensamiento !== null) sheet.getRange(rowNum, 10).setValue(pensamiento);
      if (nombreMostrar !== null) sheet.getRange(rowNum, 11).setValue(nombreMostrar);
      return jsonOut({ status: "ok", correo, apodo, telefono, pensamiento, nombreMostrar });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleRegister(body) {
  const nombre = (body.nombre || '').toString().trim();
  const password = (body.password || '').toString();
  const correo = (body.correo || '').toString().trim();
  if (!nombre || !password) {
    return jsonOut({ status: "error", message: "Falta nombre o contraseña" });
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    if (nombres.includes(nombre.toLowerCase())) {
      return jsonOut({ status: "error", message: "Ese nombre ya tiene una cuenta creada. Iniciá sesión en vez de crear una nueva." });
    }
  }
  sheet.appendRow([nombre, password, correo, 'Pendiente', 'Setter']);
  return jsonOut({ status: "pending", nombre, correo, message: "Cuenta creada. Un administrador tiene que aprobarla antes de que puedas ingresar." });
}

function handleLogin(body) {
  const nombre = (body.nombre || '').toString().trim();
  const password = (body.password || '').toString();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        if (String(rows[i][1]) !== password) {
          return jsonOut({ status: "error", message: "Contraseña incorrecta" });
        }
        const estado = rows[i][3] || 'Aprobado';
        if (estado === 'Pendiente') {
          return jsonOut({ status: "error", message: "Tu cuenta todavía está pendiente de aprobación por un administrador." });
        }
        if (estado === 'Rechazado') {
          return jsonOut({ status: "error", message: "Tu acceso fue rechazado. Contactá al administrador." });
        }
        return jsonOut({ status: "ok", nombre: rows[i][0], correo: rows[i][2] || '', rol: rows[i][4] || 'Setter', apodo: rows[i][5] || '', telefono: rows[i][6] || '', avatar: rows[i][7] || '', cover: rows[i][8] || '', pensamiento: rows[i][9] || '', nombreMostrar: rows[i][10] || '', metaFacturacion: rows[i][11] || 0 });
      }
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre. Creá una primero." });
}

function handleListUsuarios(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  let rows = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    rows = data.map(r => ({ nombre: r[0], password: String(r[1] || ''), correo: r[2] || '', estado: r[3] || 'Aprobado', rol: r[4] || 'Setter', apodo: r[5] || '', telefono: r[6] || '', avatar: r[7] || '', cover: r[8] || '', pensamiento: r[9] || '', nombreMostrar: r[10] || '', metaFacturacion: r[11] || 0 }));
  }
  return jsonOut({ status: "ok", rows });
}

function handleSetEstado(body) {
  const nombre = (body.nombre || '').toString().trim();
  const estado = (body.estado || '').toString().trim(); // 'Aprobado' | 'Rechazado' | 'Pendiente'
  if (!nombre || !estado) return jsonOut({ status: "error", message: "Falta nombre o estado" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      sheet.getRange(idx + 2, 4).setValue(estado);
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleSetRol(body) {
  const nombre = (body.nombre || '').toString().trim();
  const rol = (body.rol || '').toString().trim(); // 'Lider' | 'Setter' | 'Invitado'
  if (!nombre || !rol) return jsonOut({ status: "error", message: "Falta nombre o rol" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      sheet.getRange(idx + 2, 5).setValue(rol);
      return jsonOut({ status: "ok" });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleChangePassword(body) {
  const nombre = (body.nombre || '').toString().trim();
  const passwordActual = (body.passwordActual || '').toString();
  const passwordNueva = (body.passwordNueva || '').toString();
  const correo = body.correo !== undefined ? (body.correo || '').toString().trim() : null;

  if (!nombre) return jsonOut({ status: "error", message: "Falta nombre" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === nombre.toLowerCase()) {
        if (String(rows[i][1]) !== passwordActual) {
          return jsonOut({ status: "error", message: "La contraseña actual no es correcta" });
        }
        const rowNum = i + 2;
        if (passwordNueva) sheet.getRange(rowNum, 2).setValue(passwordNueva);
        if (correo !== null) sheet.getRange(rowNum, 3).setValue(correo);
        return jsonOut({ status: "ok" });
      }
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleUpdateCorreo(body) {
  const nombre = (body.nombre || '').toString().trim();
  const correo = (body.correo || '').toString().trim();
  if (!nombre) return jsonOut({ status: "error", message: "Falta nombre" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getUsersSheetWithCorreo(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const nombres = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(n => String(n).trim().toLowerCase());
    const idx = nombres.indexOf(nombre.toLowerCase());
    if (idx > -1) {
      sheet.getRange(idx + 2, 3).setValue(correo);
      return jsonOut({ status: "ok", correo });
    }
  }
  return jsonOut({ status: "error", message: "No existe una cuenta con ese nombre" });
}

function handleSaveMetrics(body) {
  const fecha = body.fecha; // formato "dd/mm/aaaa"
  const entries = body.entries || {}; // { thomi: {seguimientos, cash, horas, agendas}, flor: {...}, ... }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const written = [];
  const skipped = [];
  const debug = [];

  Object.keys(SHEET_NAMES).forEach(setter => {
    const sheetName = SHEET_NAMES[setter];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { skipped.push(setter + ' (hoja "' + sheetName + '" no encontrada)'); return; }
    const data = entries[setter];
    if (!data) { skipped.push(setter + ' (sin datos enviados)'); return; }

    const rowIndex = findOrCreateDateRow(sheet, fecha);

    const seg = data.seguimientos || 0;
    const cash = data.cash || 0;
    const horas = data.horas || 0;
    const agendas = data.agendas || 0;
    const conversion = seg > 0 ? `=F${rowIndex}/C${rowIndex}` : 0;

    // Un solo llamado a la hoja para las 5 columnas (Seguimientos..Conversión, que son
    // contiguas: C-G) en vez de 5 llamados separados — mucho más rápido.
    sheet.getRange(rowIndex, COLS.seguimientos, 1, 5).setValues([[seg, cash, horas, agendas, conversion]]);

    // Leads (columna H) es opcional y aparte: si el formulario no mandó
    // nada (undefined), no se toca la celda — evita pisar un valor que ya
    // estaba ahí con un 0 sin querer.
    if (data.leads !== undefined) {
      sheet.getRange(rowIndex, COLS.leads).setValue(data.leads || 0);
    }

    written.push(setter);
    debug.push({ setter, sheetName, fila: rowIndex, fechaGuardada: fecha });
  });

  return jsonOut({ status: "ok", written, skipped, debug });
}

// Usada por la extensión de Chrome/Brave (ManyChat Shortcuts) para sumar
// seguimientos automáticamente (mensajes manuales + automatizaciones
// enviadas) sin pisar Cash/Horas/Agendas, que se siguen cargando a mano
// desde el dashboard. Es aditiva: lee lo que ya hay en la columna
// Seguimientos de HOY para ese setter, le suma "cantidad" (1 por defecto),
// y escribe solo esa celda — nunca toca las columnas D, E, F.
function handleIncrementSeguimiento(body) {
  const nombre = (body.nombre || '').toString().trim().toLowerCase();
  const fecha = (body.fecha || '').toString().trim();
  const cantidad = Number(body.cantidad) || 1;

  const sheetName = SHEET_NAMES[nombre];
  if (!sheetName) return jsonOut({ status: "error", message: "Setter desconocido: " + nombre });
  if (!fecha) return jsonOut({ status: "error", message: "Falta fecha" });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonOut({ status: "error", message: 'Hoja "' + sheetName + '" no encontrada' });

  const rowIndex = findOrCreateDateRow(sheet, fecha);

  const seguimientosCell = sheet.getRange(rowIndex, COLS.seguimientos);
  const actual = Number(seguimientosCell.getValue()) || 0;
  const nuevoTotal = actual + cantidad;
  seguimientosCell.setValue(nuevoTotal);

  // Si la fila es nueva (recién creada por findOrCreateDateRow), todavía no
  // tiene la fórmula de Conversión — se la agrega para que quede igual que
  // una fila cargada a mano. Si ya existía, no se toca (evita pisar algo
  // que alguien haya puesto manual ahí).
  const conversionCell = sheet.getRange(rowIndex, COLS.conversion);
  if (!conversionCell.getFormula()) {
    conversionCell.setValue(`=F${rowIndex}/C${rowIndex}`);
  }

  return jsonOut({ status: "ok", nombre, fecha, seguimientos: nuevoTotal });
}

// Une una lista de tramos [inicioMs, finMs] — puede venir desordenada y
// con solapamientos — y devuelve la versión fusionada: ordenada, sin
// superposiciones. Es la base de cómo se evita duplicar horas cuando dos
// personas usan la misma cuenta de setter desde computadoras distintas al
// mismo tiempo: en vez de sumar cada tramo por separado, se calcula la
// UNIÓN real del tiempo cubierto.
function fusionarTramos(tramos) {
  const validos = (tramos || []).filter((t) => Array.isArray(t) && t.length === 2 && t[1] > t[0]);
  if (!validos.length) return [];
  const ordenados = validos.slice().sort((a, b) => a[0] - b[0]);
  const resultado = [ordenados[0].slice()];
  for (let i = 1; i < ordenados.length; i++) {
    const actual = ordenados[i];
    const ultimo = resultado[resultado.length - 1];
    if (actual[0] <= ultimo[1]) {
      ultimo[1] = Math.max(ultimo[1], actual[1]);
    } else {
      resultado.push(actual.slice());
    }
  }
  return resultado;
}

function segundosDeTramos(tramos) {
  return tramos.reduce((s, t) => s + (t[1] - t[0]) / 1000, 0);
}

// ===== ACTIVIDAD GENERAL: horas activas de CUALQUIER persona que use la
// extensión (no solo los 4 setters de SetterFlow) — se guarda en su
// propia hoja, agrupada por Fecha + Página de ManyChat + Nombre. Usa el
// mismo mecanismo de tramos + fusión que Horas (ver arriba) para que dos
// sesiones de la misma persona no dupliquen tiempo solapado. =====
const ACTIVIDAD_SHEET_NAME = 'Actividad Extensión';

function getOrCreateActividadSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ACTIVIDAD_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ACTIVIDAD_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['Fecha', 'Página', 'Nombre', 'Horas']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Pone cada palabra en mayúscula inicial — así "juan", "JUAN" y "Juan"
// (según cómo cada uno tenga puesto su nombre en ManyChat) terminan en la
// MISMA fila en vez de crear una fila nueva por cada variante de mayúsculas.
function tituloCap(str) {
  return String(str || '').trim().toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
}

function findOrCreateActividadRow(sheet, fecha, pageId, nombre) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const datos = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
    for (let i = 0; i < datos.length; i++) {
      if (datos[i][0] === fecha && datos[i][1] === pageId && datos[i][2] === nombre) {
        return 2 + i;
      }
    }
  }
  const newRow = lastRow + 1;
  sheet.getRange(newRow, 1, 1, 3).setValues([[fecha, pageId, nombre]]);
  return newRow;
}

function handleReportarActividadGeneral(body) {
  const nombre = tituloCap(body.nombre);
  const pageId = (body.pageId || '').toString().trim();
  const fecha = (body.fecha || '').toString().trim();
  const tramosNuevos = Array.isArray(body.tramos) ? body.tramos : [];
  if (!nombre || !fecha || !tramosNuevos.length) return jsonOut({ status: "ok", ignored: true });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    return jsonOut({ status: "error", message: "No se pudo tomar el lock, probá de nuevo." });
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const propKey = 'actividad_tramos_' + nombre.toLowerCase() + '_' + (pageId || 'sinpagina') + '_' + fecha.replace(/\//g, '-');
    const guardados = JSON.parse(props.getProperty(propKey) || '[]');
    const combinados = fusionarTramos(guardados.concat(tramosNuevos));
    props.setProperty(propKey, JSON.stringify(combinados));

    const nuevoTotal = Math.round((segundosDeTramos(combinados) / 3600) * 10000) / 10000;

    const sheet = getOrCreateActividadSheet();
    const rowIndex = findOrCreateActividadRow(sheet, fecha, pageId, nombre);
    sheet.getRange(rowIndex, 4).setValue(nuevoTotal);

    return jsonOut({ status: "ok", nombre, pageId, fecha, horas: nuevoTotal });
  } finally {
    lock.releaseLock();
  }
}

// Usada por la extensión de Chrome/Brave para sumar "Horas" automáticamente
// — el tiempo real que el setter pasa con la pestaña de ManyChat visible
// (lo mide el content script), en vez de que cada quien cargue esa columna
// a mano en el dashboard.
//
// A diferencia de Seguimientos (un simple contador que suma), acá se
// guardan los TRAMOS de tiempo (desde–hasta) de cada reporte, fusionados
// con los que ya había guardados para ese setter y ese día (en
// Propiedades del Script) — así, si dos personas distintas usan la misma
// cuenta desde computadoras diferentes al mismo tiempo, sus tramos
// superpuestos se cuentan UNA sola vez, no dos. La celda de Horas siempre
// queda con el total real de la unión, no con una simple suma.
function handleIncrementHoras(body) {
  const nombre = (body.nombre || '').toString().trim().toLowerCase();
  const fecha = (body.fecha || '').toString().trim();
  const tramosNuevos = Array.isArray(body.tramos) ? body.tramos : [];
  if (!tramosNuevos.length) return jsonOut({ status: "ok", ignored: true });

  const sheetName = SHEET_NAMES[nombre];
  if (!sheetName) return jsonOut({ status: "error", message: "Setter desconocido: " + nombre });
  if (!fecha) return jsonOut({ status: "error", message: "Falta fecha" });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    return jsonOut({ status: "error", message: "No se pudo tomar el lock, probá de nuevo." });
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const propKey = 'horas_tramos_' + nombre + '_' + fecha.replace(/\//g, '-');
    const guardados = JSON.parse(props.getProperty(propKey) || '[]');
    const combinados = fusionarTramos(guardados.concat(tramosNuevos));
    props.setProperty(propKey, JSON.stringify(combinados));

    // Redondeado a 4 decimales para no acumular ruido de punto flotante —
    // de sobra para algo que en el dashboard se muestra con 1 decimal
    // (ej. "13.0h").
    const nuevoTotal = Math.round((segundosDeTramos(combinados) / 3600) * 10000) / 10000;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonOut({ status: "error", message: 'Hoja "' + sheetName + '" no encontrada' });

    const rowIndex = findOrCreateDateRow(sheet, fecha);
    sheet.getRange(rowIndex, COLS.horas).setValue(nuevoTotal);

    const conversionCell = sheet.getRange(rowIndex, COLS.conversion);
    if (!conversionCell.getFormula()) {
      conversionCell.setValue(`=F${rowIndex}/C${rowIndex}`);
    }

    return jsonOut({ status: "ok", nombre, fecha, horas: nuevoTotal });
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function findOrCreateDateRow(sheet, fecha) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= FIRST_DATA_ROW) {
    // Usamos getDisplayValues() (el texto tal cual se ve en la celda) en vez de
    // getValues() (que reconstruye un objeto Date interno y puede confundir
    // día/mes según el idioma/zona horaria configurados). Así comparamos
    // exactamente lo que aparece escrito en tu hoja, sin ambigüedad.
    const dateColDisplay = sheet.getRange(FIRST_DATA_ROW, COLS.fecha, lastRow - FIRST_DATA_ROW + 1, 1).getDisplayValues();
    for (let i = 0; i < dateColDisplay.length; i++) {
      const normalized = normalizeDate(dateColDisplay[i][0]);
      if (normalized === fecha) {
        return FIRST_DATA_ROW + i;
      }
    }
  }
  const newRow = Math.max(lastRow + 1, FIRST_DATA_ROW);
  sheet.getRange(newRow, COLS.fecha).setValue(fecha);
  return newRow;
}

function normalizeDate(val) {
  if (val instanceof Date) {
    const dd = String(val.getDate()).padStart(2, '0');
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    const yyyy = val.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
  }
  const str = String(val).trim();

  // Formato "aaaa-mm-dd" o "aaaa/mm/dd" (el año va primero, con 4 dígitos)
  let match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (match) {
    const yyyy = match[1];
    const mm = match[2].padStart(2, '0');
    const dd = match[3].padStart(2, '0');
    return `${dd}/${mm}/${yyyy}`;
  }

  // Formato "dd/mm/aaaa", "dd-mm-aaaa" (con o sin ceros adelante)
  match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    const dd = match[1].padStart(2, '0');
    const mm = match[2].padStart(2, '0');
    let yyyy = match[3];
    if (yyyy.length === 2) yyyy = '20' + yyyy;
    return `${dd}/${mm}/${yyyy}`;
  }

  return str;
}

// ===== DIAGRAMA DEL EQUIPO (compartido, líder edita / setters solo ven) =====
// Se guarda como un solo bloque de JSON en las Propiedades del Script — no
// necesita una hoja nueva, y es instantáneo de leer/escribir.
function handleGetDiagram(body) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('diagram_data');
  const data = raw ? JSON.parse(raw) : { boxes: [], connections: [] };
  return jsonOut({ status: "ok", data });
}

function handleSaveDiagram(body) {
  const data = body.data || { boxes: [], connections: [] };
  PropertiesService.getScriptProperties().setProperty('diagram_data', JSON.stringify(data));
  return jsonOut({ status: "ok" });
}

// ===== LINKS/TAREAS QUE EL LÍDER MANDA A TODO EL EQUIPO (Configuración → Herramientas) =====
function handleGetSharedContent(body) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('shared_content');
  const data = raw ? JSON.parse(raw) : { links: [], todos: [] };
  return jsonOut({ status: "ok", data });
}

function handleSaveSharedContent(body) {
  const data = body.data || { links: [], todos: [] };
  PropertiesService.getScriptProperties().setProperty('shared_content', JSON.stringify(data));
  return jsonOut({ status: "ok" });
}

// ===== COLORES DE ETIQUETAS (extensión ManyChat Shortcuts) =====
// Uno por página de ManyChat (pageId), compartido entre TODOS los que
// manejan esa misma cuenta — así si una persona ya le puso color a una
// etiqueta, otra persona distinta que la abre ve ese mismo color, sin
// tener que elegirlo de nuevo. Mismo patrón que el diagrama de arriba.
function handleGetTagColors(body) {
  const pageId = String(body.pageId || '').trim();
  if (!pageId) return jsonOut({ status: "error", message: "Falta pageId" });
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('tagcolors_' + pageId);
  const tagColors = raw ? JSON.parse(raw) : [];
  return jsonOut({ status: "ok", tagColors });
}

function handleSaveTagColors(body) {
  const pageId = String(body.pageId || '').trim();
  if (!pageId) return jsonOut({ status: "error", message: "Falta pageId" });
  const tagColors = Array.isArray(body.tagColors) ? body.tagColors : [];
  PropertiesService.getScriptProperties().setProperty('tagcolors_' + pageId, JSON.stringify(tagColors));
  return jsonOut({ status: "ok" });
}

// ===== HYROS: traer facturación, leads y agendas del día por setter =====
// La clave de API NO va escrita acá en el código — se lee de las
// "Propiedades del Script" (Configuración del proyecto ⚙️ → Propiedades
// del script, en el editor de Apps Script). Guardala ahí una sola vez con
// el nombre HYROS_API_KEY. Así nunca queda en texto plano en un archivo
// que se pueda compartir/copiar por error, y solo la ve quien tenga
// acceso de edición a este script.
const HYROS_BASE_URL = "https://api.hyros.com/v1/api/v1.0/";

// Los "códigos de setter" (ST01/ST05/ST06/ST07) NO aparecen como tags
// exactos en Hyros — aparecen como parte de tags más largos por campaña
// (ej. "@ct15_ig_st05", "@ct03_fb_st06"), y aparecen tags de campaña
// nuevos todo el tiempo. Por eso NO se filtra por tag exacto contra la
// API (se quedaría afuera de cualquier campaña nueva que no esté en una
// lista fija) — se trae TODO lo del día y se clasifica acá mirando si
// algún tag CONTIENE el código del setter.
const HYROS_SETTER_CODES = { thomi: 'st01', flor: 'st05', valeria: 'st07', franco: 'st06', paula: 'st02' };

// Glosario de cuentas/canales de origen (para el panel "Por fuente" del
// dashboard — de qué cuenta/canal vienen los leads/agendas/ventas, no de
// qué setter). Los tags de Hyros siguen el patrón "ctXX_canal_stYY" (ej.
// "ct03_ig_st05" = cuenta ct03, canal Instagram, setter 05).
const HYROS_ACCOUNT_NAMES = {
  ct03: 'tino.mossu',
  ct15: 'teotinivelli',
  ct16: 'teotinivelliprime',
  ct07: 'tinosinfiltro',
  ct12: 'tinohub',
  ct11: 'tinolifestylee',
  ct23: 'gisenriquez.m'
};
const HYROS_CHANNEL_NAMES = {
  ig: 'Instagram',
  tk: 'TikTok',
  zoom: 'Zoom',
  fb: 'Facebook'
};

// Busca en los tags el primero que matchee una cuenta conocida (ctXX) y,
// si lo encuentra, el canal (ig/tk/zoom/fb) que venga en el MISMO tag.
//
// OJO: antes, si ningún tag tenía una cuenta conocida (tráfico orgánico,
// campaña vieja sin ese formato, etc.), esto devolvía null y esa
// agenda/lead/venta se descartaba ENTERA del ranking "Fuentes" — sin
// aparecer en ningún lado. Eso hacía que "Métricas" (que cuenta por tag
// de SETTER, sin importar si tiene cuenta reconocida) diera más total que
// "Fuentes" (Ej: 651 vs 521), aunque las dos deberían representar lo
// mismo. Ahora, si no matchea ninguna cuenta conocida, cae en un cajón
// "Otra cuenta" en vez de desaparecer — así los dos paneles siempre
// suman el mismo total.
function hyrosClasificarFuente(tags) {
  if (tags && tags.length) {
    for (let i = 0; i < tags.length; i++) {
      const partes = String(tags[i]).toLowerCase().replace(/^[@!$]/, '').split(/[_\-]/);
      let cuenta = null;
      let canal = null;
      partes.forEach((p) => {
        if (HYROS_ACCOUNT_NAMES[p]) cuenta = HYROS_ACCOUNT_NAMES[p];
        if (HYROS_CHANNEL_NAMES[p]) canal = HYROS_CHANNEL_NAMES[p];
      });
      if (cuenta) {
        const canalFinal = canal || 'Otro canal';
        return { cuenta: cuenta, canal: canalFinal, key: cuenta + ' · ' + canalFinal };
      }
    }
  }
  return { cuenta: 'Otra cuenta', canal: 'Otro canal', key: 'Otra cuenta · Otro canal' };
}

function hyrosSumarFuente(fuentes, tags, campo, delta) {
  const clasif = hyrosClasificarFuente(tags);
  if (!clasif) return;
  if (!fuentes[clasif.key]) {
    fuentes[clasif.key] = { cuenta: clasif.cuenta, canal: clasif.canal, leads: 0, agendas: 0, cash: 0 };
  }
  fuentes[clasif.key][campo] += delta;
}

// Caché COMPARTIDA entre todos los que tengan el dashboard abierto (a
// diferencia de la caché del navegador, que es una por persona). Sin
// esto, si 4 personas tienen el dashboard abierto a la vez con
// auto-refresh, son 4 consultas a Hyros en paralelo por cada actualización
// — eso fue justamente lo que siguió agotando la cuota diaria de UrlFetch
// incluso después de haber espaciado el sync automático de fondo. Con
// CacheService (compartido a nivel de todo el proyecto, no por sesión de
// navegador), la primera persona que pide un rango de fechas paga el
// costo real, y todas las demás que pidan ESE MISMO rango en los
// siguientes 2 minutos reciben la respuesta guardada, sin pegarle a
// Hyros de nuevo.
function handleFetchHyros(body) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) {
    return jsonOut({ status: "error", message: "Falta configurar HYROS_API_KEY en Propiedades del Script (⚙️ en el editor de Apps Script)." });
  }

  const fecha = (body.fecha || '').toString().trim(); // dd/mm/aaaa
  if (!fecha) return jsonOut({ status: "error", message: "Falta fecha" });
  const fechaHasta = (body.fechaHasta || '').toString().trim() || null; // opcional, dd/mm/aaaa

  const cache = CacheService.getScriptCache();
  const cacheKey = 'hyros_' + fecha + '_' + (fechaHasta || fecha);
  const cached = cache.get(cacheKey);
  if (cached) {
    console.log('handleFetchHyros — servido desde caché (' + cacheKey + '), no se llamó a Hyros.');
    return jsonOut(JSON.parse(cached));
  }
  console.log('handleFetchHyros — SIN caché, pegándole a Hyros de verdad (' + cacheKey + ').');

  let resultado;
  try {
    resultado = hyrosFetchDataForDate(apiKey, fecha, fechaHasta);
  } catch (err) {
    console.log('handleFetchHyros — error real de Hyros: ' + err.message);
    return jsonOut({ status: "error", message: "Error consultando Hyros: " + err.message });
  }

  const payload = { status: "ok", fecha, fechaHasta: fechaHasta || fecha, data: resultado.data, fuentes: resultado.fuentes, fuentesPorSetter: resultado.fuentesPorSetter };
  try { cache.put(cacheKey, JSON.stringify(payload), 120); } catch (e) { /* si el payload pasa los 100KB de límite de CacheService, seguimos igual sin cachear */ }
  return jsonOut(payload);
}

// Lógica compartida entre el pedido manual (botón "Traer de Hyros", vía
// handleFetchHyros) y el automático (autoSyncHyrosHoy, disparado solo por
// un trigger de tiempo). Separado en su propia función para no duplicar
// código entre los dos caminos.
// "fecha" es el día de INICIO del rango (dd/mm/aaaa). "fechaHasta" es
// opcional — si no se manda, es el mismo día (un solo día, como antes).
// Si se manda, se trae TODO el rango desde "fecha" 00:00 hasta
// "fechaHasta" 23:59, ambos incluidos.
function hyrosFetchDataForDate(apiKey, fecha, fechaHasta) {
  const partes = fecha.split('/');
  if (partes.length !== 3) throw new Error('Formato de fecha inválido, esperado dd/mm/aaaa');

  const partesHasta = (fechaHasta || fecha).split('/');
  if (partesHasta.length !== 3) throw new Error('Formato de fechaHasta inválido, esperado dd/mm/aaaa');

  // Arma la lista de días entre "fecha" y "fechaHasta" (si no hay
  // fechaHasta, es un solo día) — cada uno se le va a pedir a Hyros POR
  // SEPARADO más abajo, en vez de un solo pedido con el rango entero.
  // OJO: esto es a propósito. hyrosForEachLead/Sale/Call tienen un tope de
  // seguridad de 10.000 registros por pedido (ver esas funciones) — con el
  // volumen de tráfico de esta cuenta, un rango de varios días junto se
  // pasa de ese tope fácil y corta el conteo antes de terminar, dando un
  // total más bajo del real (un solo día individual nunca se acerca a esa
  // cantidad). Pedir día por día evita el problema sin tocar el tope.
  const desde = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  const hasta = new Date(Number(partesHasta[2]), Number(partesHasta[1]) - 1, Number(partesHasta[0]));
  const dias = [];
  for (let d = new Date(desde); d <= hasta; d.setDate(d.getDate() + 1)) {
    dias.push({ dd: String(d.getDate()).padStart(2, '0'), mm: String(d.getMonth() + 1).padStart(2, '0'), yyyy: d.getFullYear() });
  }

  const data = {};
  Object.keys(SHEET_NAMES).forEach(function (setter) {
    data[setter] = { cash: 0, leads: 0, agendas: 0 };
  });
  // "fuentes": lo mismo que "data" pero agrupado por cuenta/canal de
  // origen (tino.mossu, teotinivelli, etc. · Instagram/TikTok/...) en vez
  // de por setter — para el panel "Por fuente" del dashboard. Se calcula
  // en el mismo recorrido de leads/ventas/llamadas de acá abajo, sin
  // pedidos extra a Hyros.
  const fuentes = {};
  // "fuentesPorSetter": lo mismo pero UNA COPIA separada por cada setter —
  // para el panel "Fuentes" que se ve dentro de "Ver perfil" de cada
  // persona (de qué cuenta/canal vienen SUS leads/agendas/ventas puntual).
  const fuentesPorSetter = {};
  Object.keys(SHEET_NAMES).forEach(function (setter) { fuentesPorSetter[setter] = {}; });

  // Se probó cambiar esto a lastSource (el campo "Source" que se ve en el
  // panel de Hyros) pensando que iba a coincidir mejor, pero probado
  // contra los números reales del CRM de Hyros, ESTE método (tags) da
  // igual o más preciso — el problema real nunca fue el método de
  // conteo, era el huso horario de arriba. No cambiar esto sin volver a
  // probar contra el CRM real primero.
  // OJO: "fuentes" (el ranking global de cuenta/canal) solo debe sumar lo
  // que efectivamente vino de un setter (matcheó stXX) — antes se sumaba
  // acá abajo SIEMPRE, tenga o no tenga setter asignado, así que traía
  // TODO el tráfico de la cuenta de Hyros (incluido lo que nunca tocó a
  // un setter) y el total no cerraba contra "Métricas". Ahora se suma
  // adentro del mismo forEach de setters matcheados, igual que
  // fuentesPorSetter — si un lead/venta/llamada no matchea ningún
  // setter, no entra a "Fuentes" en absoluto.
  // OJO: la cuenta de Hyros de Factor Studios está en huso horario
  // -06:00 (confirmado contra /api/v1.0/user-info) — pedir el rango del
  // día en "Z" (UTC) corre la ventana 6 horas, cuela leads de la noche
  // anterior (hora local) y se pierde el final del día de hoy. Con el
  // offset puesto directo acá, la API ya interpreta "todo el día" en el
  // huso horario correcto sin tener que calcular nada a mano.
  dias.forEach((f) => {
    const fromDate = `${f.yyyy}-${f.mm}-${f.dd}T00:00:00-06:00`;
    const toDate = `${f.yyyy}-${f.mm}-${f.dd}T23:59:59-06:00`;

    hyrosForEachLead(apiKey, fromDate, toDate, (lead) => {
      const setters = hyrosSetterForTags(lead.tags);
      setters.forEach((setter) => {
        data[setter].leads += 1;
        hyrosSumarFuente(fuentesPorSetter[setter], lead.tags, 'leads', 1);
        hyrosSumarFuente(fuentes, lead.tags, 'leads', 1);
      });
    });
    hyrosForEachSale(apiKey, fromDate, toDate, (sale) => {
      const leadTags = sale.lead && sale.lead.tags;
      const price = sale.usdPrice || sale.price || {};
      if (hyrosPrecioExcluido(price.price)) return; // ventas de $1 y $17 (T.I.N.O.) no cuentan como facturación
      const neto = (Number(price.price) || 0) - (Number(price.refunded) || 0);
      const setters = hyrosSetterForTags(leadTags);
      setters.forEach((setter) => {
        data[setter].cash += neto;
        hyrosSumarFuente(fuentesPorSetter[setter], leadTags, 'cash', neto);
        hyrosSumarFuente(fuentes, leadTags, 'cash', neto);
      });
    });
    // "Agendas" = llamadas ("calls") agendadas en Hyros. Igual que ventas,
    // el tag del setter vive en lead.tags, no en la llamada directamente.
    hyrosForEachCall(apiKey, fromDate, toDate, (call) => {
      const leadTags = call.lead && call.lead.tags;
      const setters = hyrosSetterForTags(leadTags);
      setters.forEach((setter) => {
        data[setter].agendas += 1;
        hyrosSumarFuente(fuentesPorSetter[setter], leadTags, 'agendas', 1);
        hyrosSumarFuente(fuentes, leadTags, 'agendas', 1);
      });
    });
  });

  return { data: data, fuentes: fuentes, fuentesPorSetter: fuentesPorSetter };
}

// ===== AUTO-SYNC: el trigger en sí corre cada 1 minuto (es el mínimo que
// permite Apps Script — no existe un disparador de "cada 30 segundos").
// Escribe DIRECTO el Cash/Agendas/Leads de HOY para los 4 setters, sin
// que nadie tenga que abrir el dashboard ni tocar "Traer de Hyros"/
// "Guardar".
//
// Esto YA NO es el mecanismo principal de actualización — desde que están
// los webhooks (más abajo), Hyros nos avisa solo, al instante, cada vez
// que pasa algo. Esta función quedó como una RED DE SEGURIDAD: si algún
// aviso se pierde (ej. Hyros reintenta y falla, o pasa algo que no está
// suscripto como un reembolso), este sync recalcula el día entero desde
// cero y corrige cualquier diferencia. Por eso no hace falta que corra
// seguido — antes corría 2 veces por minuto (~2880/día) y eso fue lo que
// terminó agotando la cuota diaria de UrlFetch de Google Apps Script
// ("Service invoked too many times for one day: urlfetch"). Ahora corre
// una sola vez por disparo (ver instalarAutoSyncHyros más abajo).
//
// OJO — efecto secundario a tener en cuenta: como pisa esas 3 columnas de
// HOY (y de los días anteriores, ver MC_RESYNC_DIAS_ATRAS abajo),
// cualquier corrección manual que alguien escriba a mano en
// Cash/Agendas/Leads de esos días se va a volver a pisar con el valor de
// Hyros en la próxima corrida. Seguimientos y Horas NO se tocan acá
// (siguen siendo 100% manuales/de la extensión).
//
// Además de HOY, también re-sincroniza los últimos días atrás (ver
// HYROS_RESYNC_DIAS_ATRAS): en Hyros, un lead puede seguir ganando tags
// DESPUÉS de haberse creado (ej. un lead que entró el lunes sin tag de
// setter, y el miércoles una campaña de retargeting lo retoca y ahí sí
// queda tageado) — sin este re-sync, el Sheet quedaba con una foto
// congelada del momento exacto en que pasó "hoy" para ese día, y nunca
// se enteraba de esos tags que se agregan más tarde. Re-sincronizar los
// últimos días (no TODA la historia, para no disparar la cantidad de
// pedidos a Hyros) corrige eso solo, sin tener que tocar nada a mano.
const HYROS_RESYNC_DIAS_ATRAS = 3; // hoy + los 2 días anteriores

function autoSyncHyrosHoy() {
  const tz = Session.getScriptTimeZone();
  for (let i = 0; i < HYROS_RESYNC_DIAS_ATRAS; i++) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const fecha = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
    syncHyrosUnaVez(fecha);
  }
}

function syncHyrosUnaVez(fecha) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) return; // corre sola, no hay a quién avisarle un error acá

  if (!fecha) {
    const tz = Session.getScriptTimeZone();
    fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  }

  let resultado;
  try {
    resultado = hyrosFetchDataForDate(apiKey, fecha);
  } catch (err) {
    console.log('autoSyncHyrosHoy — error consultando Hyros (' + fecha + '): ' + err.message);
    return;
  }
  const data = resultado.data;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_NAMES).forEach((setter) => {
    const sheet = ss.getSheetByName(SHEET_NAMES[setter]);
    if (!sheet) return;
    const rowIndex = findOrCreateDateRow(sheet, fecha);
    const info = data[setter];
    sheet.getRange(rowIndex, COLS.cash).setValue(info.cash);
    sheet.getRange(rowIndex, COLS.agendas).setValue(info.agendas);
    sheet.getRange(rowIndex, COLS.leads).setValue(info.leads);
    const conversionCell = sheet.getRange(rowIndex, COLS.conversion);
    if (!conversionCell.getFormula()) {
      conversionCell.setValue(`=F${rowIndex}/C${rowIndex}`);
    }
  });

  guardarFuentesDelDia(fecha, resultado.fuentes);
}

// ===== "Fuentes" guardado en el Sheet (por cuenta/canal, Leads/Agendas/
// Cash) — antes este desglose solo vivía en vivo (el panel "Fuentes" del
// dashboard le pegaba a Hyros cada vez que alguien lo abría, con una
// caché de 2 minutos nomás). Ahora se guarda acá mismo, gratis, con los
// mismos datos que ya se calculan arriba para Cash/Agendas/Leads — sin
// ningún pedido extra a Hyros.
const FUENTES_SHEET_NAME = 'Fuentes (diario)';
const FUENTES_COLS = { fecha: 1, cuenta: 2, canal: 3, leads: 4, agendas: 5, cash: 6 };

function getOrCreateFuentesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FUENTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FUENTES_SHEET_NAME);
    sheet.getRange(1, 1, 1, 6).setValues([['Fecha', 'Cuenta', 'Canal', 'Leads', 'Agendas', 'Cash']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Guarda (reemplazando lo que hubiera) el desglose de un día puntual — así
// cada re-sync (automático cada 15 min, o el resyncMesActual manual) deja
// esa fecha al día sin ir acumulando filas duplicadas.
function guardarFuentesDelDia(fecha, fuentes) {
  const sheet = getOrCreateFuentesSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const fechasCol = sheet.getRange(2, FUENTES_COLS.fecha, lastRow - 1, 1).getDisplayValues();
    for (let i = fechasCol.length - 1; i >= 0; i--) {
      if (fechasCol[i][0] === fecha) sheet.deleteRow(2 + i);
    }
  }

  const filas = Object.keys(fuentes || {}).map((key) => {
    const f = fuentes[key];
    return [fecha, f.cuenta, f.canal, f.leads || 0, f.agendas || 0, f.cash || 0];
  });
  if (filas.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, filas.length, 6).setValues(filas);
  }
}

// Corré ESTA función UNA SOLA VEZ (▷ Ejecutar, con "resyncMesActual"
// elegida en el desplegable) para poner al día, con la lógica ACTUAL del
// script, los días de ESTE MES que el auto-sync de cada 15 minutos ya no
// toca (ese solo cubre hoy + los 2 días anteriores) — cualquier día más
// viejo de este mes se quedó con lo que se escribió la última vez que
// SÍ estuvo dentro de esa ventana de 3 días, posiblemente con una versión
// más vieja del script. No toca nada de meses anteriores a propósito.
//
// Si algún día falla (ej. se corta la cuota de Hyros a mitad de camino),
// sigue con el resto y lo deja anotado en el log — volvés a correr esta
// misma función otro día para completar los que hayan quedado pendientes
// (no rompe nada si un día ya estaba bien, lo vuelve a escribir igual).
function resyncMesActual() {
  const tz = Session.getScriptTimeZone();
  const hoy = new Date();
  const primerDiaDelMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const ultimoDiaAResincronizar = new Date(hoy);
  ultimoDiaAResincronizar.setDate(ultimoDiaAResincronizar.getDate() - HYROS_RESYNC_DIAS_ATRAS);

  if (primerDiaDelMes > ultimoDiaAResincronizar) {
    Logger.log('No hay días de este mes para re-sincronizar aparte de los que ya cubre el auto-sync.');
    return;
  }

  const dias = [];
  for (let d = new Date(primerDiaDelMes); d <= ultimoDiaAResincronizar; d.setDate(d.getDate() + 1)) {
    dias.push(Utilities.formatDate(d, tz, 'dd/MM/yyyy'));
  }

  Logger.log(`Re-sincronizando ${dias.length} día(s): ${dias.join(', ')}`);
  dias.forEach((fecha) => {
    Logger.log('--- ' + fecha + ' ---');
    syncHyrosUnaVez(fecha);
  });
  Logger.log('Listo.');
}

// ===== WEBHOOK: Hyros nos avisa apenas pasa algo (lead nuevo, venta,
// llamada) — instantáneo. Suma/resta SOLO lo puntual de ese aviso (no
// recalcula el día entero cada vez, eso era lo que generaba el delay
// cuando entraban varios avisos seguidos) — mucho más rápido.
//
// Los avisos de venta ("sale.attributed") y llamada ("call.attributed")
// ya traen los tags del lead adentro del mismo aviso — no hace falta
// pedirle nada más a Hyros. El de lead nuevo ("lead.opted.in") NO trae
// tags, así que para ESE caso puntual sí hace falta un pedido chico
// (1 lead) para poder clasificarlo.
//
// Por las dudas algo se calcule de más/de menos acá (ej. reintentos de
// Hyros, algún aviso raro), el sync de cada 30 segundos sigue corriendo
// igual y recalcula el total real desde cero — así nunca queda
// desactualizado por mucho tiempo, pase lo que pase acá.
function handleHyrosWebhook(payload) {
  try {
    const type = payload.type;
    const body = payload.body || {};
    if (type === 'lead.opted.in' || type === 'lead.opted.in.first.time') {
      handleWebhookLead(body);
    } else if (type === 'lead.tag.added') {
      handleWebhookLeadTagAdded(body, payload.eventId);
    } else if (type === 'sale.attributed') {
      handleWebhookSale(body);
    } else if (type === 'call.attributed') {
      handleWebhookCall(body);
    }
    // Otros tipos (sale.refunded, subscription.*, lead.origin.assigned,
    // lead.stage.changed, lead.tag.removed) no están suscriptos y no
    // deberían llegar acá, pero si llegan se ignoran sin romper nada.
  } catch (err) {
    console.log('handleHyrosWebhook — error: ' + err.message);
  }
  // Hyros solo necesita un 200 OK para no reintentar de más — no hace
  // falta devolverle nada específico.
  return jsonOut({ status: "ok" });
}

// lead.opted.in NO trae todavía la tag del setter en ese instante — la
// automatización de ManyChat recién se la pone DESPUÉS del opt-in inicial,
// así que aunque el aviso trae un array de tags, para clasificar por
// setter acá haría falta pedirle el lead de nuevo a Hyros más tarde. Ese
// pedido extra por cada lead nuevo (en una cuenta con muchísimo tráfico,
// no solo el de los 4 setters) fue lo que agotó la cuota diaria de
// UrlFetch de Apps Script ("Service invoked too many times for one day:
// urlfetch") — y una vez agotada, se rompía TODO lo demás (el resync de
// 15 min, el botón de "Traer de Hyros" del dashboard, etc.), no solo los
// Leads. Por eso este aviso puntual sigue sin pedirle nada a Hyros — ver
// handleWebhookLeadTagAdded más abajo para el conteo real en vivo.
function handleWebhookLead(body) {
  // No-op a propósito — ver comentario de arriba.
}

// Conteo de Leads EN VIVO, sin el problema de cuota de arriba: en vez de
// contar en el opt-in (cuando todavía no sabemos de quién es), se cuenta
// apenas Hyros avisa que se le agregó la tag del setter — el aviso
// "lead.tag.added" ya trae la lista de tags agregadas adentro del mismo
// paquete (igual que sale.attributed/call.attributed), así que no hace
// falta pedirle nada más a Hyros. Sigue siendo prácticamente instantáneo
// (segundos, no los 15 minutos del resync), sin gastar cuota.
//
// body.tags = SOLO las tags que se acaban de agregar en este aviso puntual
// (no la lista completa del lead) — así que si más tarde se le agrega una
// tag sin relación (ej. "@seguimiento"), este aviso no vuelve a sumar de
// más, porque esa tag nueva no matchea ningún código de setter.
//
// Dedup: Hyros puede reintentar el mismo aviso si no le contestamos rápido
// — se usa el eventId (único por aviso, no por lead) en CacheService para
// no sumar el mismo aviso dos veces. 6 horas de margen es de sobra para
// cualquier reintento real; no cubre el caso raro de que a un lead se le
// saque y se le vuelva a poner la MISMA tag de setter (eso sí generaría
// otro eventId y sumaría de nuevo) — un caso tan infrecuente que no vale
// la pena complicar esto por él, y el resync de 15 min lo termina
// corrigiendo solo si llegara a pasar.
function handleWebhookLeadTagAdded(body, eventId) {
  const setters = hyrosSetterForTags(body.tags);
  if (!setters.length) return;
  if (eventId) {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'leadtag_' + eventId;
    if (cache.get(cacheKey)) return; // ya procesado (reintento de Hyros)
    cache.put(cacheKey, '1', 21600); // 6h
  }
  setters.forEach((setter) => incrementarCeldaHoy(setter, 'leads', 1));
}

function handleWebhookSale(body) {
  const tags = body.lead && body.lead.tags;
  const price = (body.product && (body.product.USDPrice || body.product.usdPrice || body.product.price)) || {};
  if (hyrosPrecioExcluido(price.price)) return; // ventas de $1 y $17 (T.I.N.O.) no cuentan como facturación
  const neto = (Number(price.price) || 0) - (Number(price.refunded) || 0);
  hyrosSetterForTags(tags).forEach((setter) => incrementarCeldaHoy(setter, 'cash', neto));
}

// Precios que NO cuentan como facturación real — son el mismo producto
// ("UDC Closer Elite") vendido en pasos bajos del funnel (a $1 de entrada,
// y a $17 en la oferta "Sistema T.I.N.O."), no una venta de verdad. No hay
// un producto/tag separado para distinguirlos en Hyros — la única forma
// de identificarlos es por el precio exacto. Con margen chico (0.01) por
// si el precio viene con algún redondeo de centavos.
function hyrosPrecioExcluido(precio) {
  const p = Number(precio);
  if (isNaN(p)) return false;
  return Math.abs(p - 1) < 0.01 || Math.abs(p - 17) < 0.01;
}

function handleWebhookCall(body) {
  const tags = body.lead && body.lead.tags;
  hyrosSetterForTags(tags).forEach((setter) => incrementarCeldaHoy(setter, 'agendas', 1));
}

// Suma "delta" a la celda de HOY de un setter (columna según colKey:
// 'leads'/'cash'/'agendas'). Usa un lock corto para que dos avisos que
// lleguen casi al mismo tiempo no se pisen entre sí y se pierda un
// incremento (si no llega a tomar el lock a tiempo, se lo salta sin
// romper nada — el sync de 30 segundos lo termina corrigiendo solo).
function incrementarCeldaHoy(setter, colKey, delta) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    console.log('incrementarCeldaHoy — no se pudo tomar el lock a tiempo, se salta (el sync de 30s lo corrige).');
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES[setter]);
    if (!sheet) return;
    const tz = Session.getScriptTimeZone();
    const fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    const rowIndex = findOrCreateDateRow(sheet, fecha);
    const cell = sheet.getRange(rowIndex, COLS[colKey]);
    const actual = Number(cell.getValue()) || 0;
    cell.setValue(actual + delta);
    const conversionCell = sheet.getRange(rowIndex, COLS.conversion);
    if (!conversionCell.getFormula()) {
      conversionCell.setValue(`=F${rowIndex}/C${rowIndex}`);
    }
  } finally {
    lock.releaseLock();
  }
}

// Diagnóstico de solo lectura — NO borra ni crea nada, solo lista en el
// Logger (Ver → Registros, o Ctrl+Enter después de ejecutar) todas las
// suscripciones de webhook que Hyros tiene activas apuntando a nuestra URL.
// Correr esto UNA VEZ desde el editor (▷ Ejecutar, "listarWebhooksHyros")
// para confirmar si hay duplicadas — si aparece más de una fila, cada
// evento real de Hyros nos está llegando esa misma cantidad de veces.
function listarWebhooksHyros() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) {
    Logger.log('Falta configurar HYROS_API_KEY en Propiedades del Script primero.');
    return;
  }
  const targetUrl = "https://script.google.com/macros/s/AKfycbxX_y6KKd75EjQ5GelM47lK0BJVuGmiKITo0vqnXWNF4i3jSonTEZy22L_9rA0R4SJFnw/exec";
  const listResp = UrlFetchApp.fetch('https://api.hyros.com/v1/api/v1.0/webhook-subscriptions', {
    method: 'get',
    headers: { 'API-Key': apiKey },
    muteHttpExceptions: true
  });
  const listData = JSON.parse(listResp.getContentText());
  const propias = (listData.result || []).filter((sub) => sub.targetUrl === targetUrl);
  Logger.log('Suscripciones totales en la cuenta de Hyros: ' + ((listData.result || []).length));
  Logger.log('Suscripciones apuntando a NUESTRA url: ' + propias.length);
  propias.forEach((sub, i) => {
    Logger.log('#' + (i + 1) + ' — id: ' + sub.externalId + ' | eventos: ' + JSON.stringify(sub.eventTypes) + ' | activa: ' + sub.active);
  });
}

// Correr ESTA función UNA SOLA VEZ desde el editor (▷ Ejecutar, con
// "instalarWebhookHyros" elegida en el desplegable de arriba) para que
// Hyros empiece a avisarnos en tiempo real. Necesita HYROS_API_KEY ya
// configurada en Propiedades del Script (el mismo paso de siempre).
//
// OJO: si se corre esta función más de una vez sin borrar antes las
// suscripciones viejas, Hyros las va acumulando — cada evento real llega
// UNA VEZ POR CADA suscripción activa apuntando a nuestra URL, así que 3
// suscripciones duplicadas significan 3 avisos por evento, pisándose
// entre sí y generando demoras. Por eso, antes de crear una nueva, esta
// función borra cualquier suscripción vieja que ya apunte a nuestra
// misma URL — así nunca quedan duplicados, sin importar cuántas veces se
// la corra.
function instalarWebhookHyros() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) {
    Logger.log('Falta configurar HYROS_API_KEY en Propiedades del Script primero.');
    return;
  }
  // La misma URL fija que ya usan la extensión y el dashboard — no se usa
  // ScriptApp.getService().getUrl() porque corrida a mano desde el editor
  // (no como parte de un pedido web real) puede no devolver la URL
  // pública correcta.
  const targetUrl = "https://script.google.com/macros/s/AKfycbxX_y6KKd75EjQ5GelM47lK0BJVuGmiKITo0vqnXWNF4i3jSonTEZy22L_9rA0R4SJFnw/exec";

  const listResp = UrlFetchApp.fetch('https://api.hyros.com/v1/api/v1.0/webhook-subscriptions', {
    method: 'get',
    headers: { 'API-Key': apiKey },
    muteHttpExceptions: true
  });
  const listData = JSON.parse(listResp.getContentText());
  (listData.result || []).forEach((sub) => {
    if (sub.targetUrl === targetUrl) {
      UrlFetchApp.fetch('https://api.hyros.com/v1/api/v1.0/webhook-subscriptions/' + sub.externalId, {
        method: 'delete',
        headers: { 'API-Key': apiKey },
        muteHttpExceptions: true
      });
      Logger.log('Borrada suscripción vieja: ' + sub.externalId);
    }
  });

  const payload = {
    name: 'SetterFlow — sync instantáneo',
    targetUrl: targetUrl,
    eventTypes: ['lead.opted.in', 'sale.attributed', 'call.attributed', 'lead.tag.added']
  };
  const response = UrlFetchApp.fetch('https://api.hyros.com/v1/api/v1.0/webhook-subscriptions', {
    method: 'post',
    headers: { 'API-Key': apiKey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('Respuesta de Hyros (' + response.getResponseCode() + '): ' + response.getContentText());
}

// Correr ESTA función UNA SOLA VEZ desde el editor (▷ Ejecutar, con
// "instalarAutoSyncHyros" elegida en el desplegable de arriba) para
// activar el auto-sync. No hace falta correrla de nuevo después — el
// trigger queda instalado permanentemente hasta que lo borres a mano
// (⏰ ícono de relojito en el menú de la izquierda del editor →
// "Disparadores" → basurero al lado de "autoSyncHyrosHoy").
//
// Corre cada 15 minutos nomás — no cada 1 minuto como antes. Ahora que
// los webhooks avisan al instante, este trigger es solo la red de
// seguridad (ver comentario en autoSyncHyrosHoy más arriba), así que no
// necesita correr seguido. Si YA tenías el trigger viejo instalado (cada
// 1 minuto), hace falta volver a correr esta función una vez para que
// reemplace ese trigger por este más espaciado — si no, sigue corriendo
// cada 1 minuto y se va a volver a agotar la cuota.
function instalarAutoSyncHyros() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'autoSyncHyrosHoy') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSyncHyrosHoy')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Listo — autoSyncHyrosHoy va a correr sola cada 15 minutos de ahora en más.');
}

// Devuelve las claves de setter (thomi/flor/valeria/franco) cuyos tags
// aparecen en este lead/venta/llamada. Normalmente 1 sola — puede ser 0
// (sin setter asignado todavía) o más de 1 si el lead quedó con tags de
// más de un setter (ej. reasignado entre agentes).
//
// Se probó cambiar esto a mirar solo firstSource/lastSource (pensando
// que iba a coincidir mejor con el filtro de "Sources" de Hyros), pero
// probado en vivo contra los números reales del CRM de Hyros dio PEOR
// resultado, no mejor — el problema real era el huso horario del rango
// de fechas (ver hyrosFetchDataForDate), no el método de conteo. Con el
// huso horario corregido, este método de tags coincide casi exacto.
function hyrosSetterForTags(tags) {
  if (!tags || !tags.length) return [];
  const tagsMin = tags.map((t) => String(t).toLowerCase());
  const matches = [];
  Object.keys(HYROS_SETTER_CODES).forEach((setter) => {
    const codigo = HYROS_SETTER_CODES[setter];
    if (tagsMin.some((t) => t.indexOf(codigo) !== -1)) matches.push(setter);
  });
  return matches;
}

function hyrosApiGet(apiKey, path, params) {
  const query = Object.keys(params)
    .map((k) => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
    .join('&');
  const url = HYROS_BASE_URL + path + '?' + query;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'API-Key': apiKey },
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Hyros ' + path + ' devolvió ' + code + ': ' + response.getContentText());
  }
  return JSON.parse(response.getContentText());
}

// Tope de 40 páginas (40 * 250 = 10.000 registros) como red de seguridad
// contra un loop infinito si Hyros devolviera un nextPageId inválido — un
// solo día de leads/ventas/llamadas de esta cuenta no debería acercarse a eso.
function hyrosForEachLead(apiKey, fromDate, toDate, cb) {
  let pageId = null;
  let guard = 0;
  do {
    const params = { fromDate: fromDate, toDate: toDate, pageSize: 250 };
    if (pageId) params.pageId = pageId;
    const resp = hyrosApiGet(apiKey, 'leads', params);
    (resp.result || []).forEach(cb);
    pageId = resp.nextPageId || null;
    guard++;
  } while (pageId && guard < 40);
}

function hyrosForEachSale(apiKey, fromDate, toDate, cb) {
  let pageId = null;
  let guard = 0;
  do {
    const params = { fromDate: fromDate, toDate: toDate, pageSize: 250 };
    if (pageId) params.pageId = pageId;
    const resp = hyrosApiGet(apiKey, 'sales', params);
    (resp.result || []).forEach(cb);
    pageId = resp.nextPageId || null;
    guard++;
  } while (pageId && guard < 40);
}

function hyrosForEachCall(apiKey, fromDate, toDate, cb) {
  let pageId = null;
  let guard = 0;
  do {
    const params = { fromDate: fromDate, toDate: toDate, pageSize: 250 };
    if (pageId) params.pageId = pageId;
    const resp = hyrosApiGet(apiKey, 'calls', params);
    (resp.result || []).forEach(cb);
    pageId = resp.nextPageId || null;
    guard++;
  } while (pageId && guard < 40);
}

// Función SOLO para autorizar el permiso de llamadas externas (UrlFetchApp)
// la primera vez — Apps Script no lo pide solo con "Guardar", hay que
// ejecutar manualmente cualquier función que use UrlFetchApp una vez desde
// el editor para que aparezca el cartel de autorización. Corré ESTA
// función (▷ Ejecutar, con "autorizarHyros" elegida en el desplegable de
// arriba), aceptá los permisos, y listo — no hace falta correrla de nuevo.
function autorizarHyros() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) {
    Logger.log('Falta configurar HYROS_API_KEY en Propiedades del Script primero.');
    return;
  }
  const resp = hyrosApiGet(apiKey, 'leads', { pageSize: 1 });
  Logger.log('OK, la conexión con Hyros funciona: ' + JSON.stringify(resp).slice(0, 200));
}

// Corré ESTA función (▷ Ejecutar, con "diagnosticoUnDia" elegida en el
// desplegable de arriba) para comparar, para UN día puntual, el total
// "de verdad" que devuelve Hyros (contando cada lead/venta/llamada UNA
// sola vez, sin importar el setter) contra lo que suma nuestro sistema
// por tags. Sirve para separar dos causas posibles de un número inflado:
//   - Webhooks duplicados (el mismo evento contado 2 o 3 veces) — si es
//      esto, "reasignados" abajo va a dar 0 o casi 0, y el total de
//      Hyros ya de por sí no va a coincidir con lo esperado.
//   - Leads/ventas/llamadas con tags de MÁS DE UN setter (ej. un lead
//      que pasó por la campaña de dos personas distintas, típico en
//      retargeting) — cada uno de esos se cuenta una vez POR CADA setter
//      que matchea, así que la suma de los 4 setters puede superar al
//      total real de Hyros sin que haya ningún duplicado de por medio.
//      Si es esto, "reasignados" va a explicar buena parte de la
//      diferencia entre el total de Hyros y la suma de los 4 setters.
// Editá FECHA_A_REVISAR acá abajo con el día que quieras chequear (por
// defecto, ayer) y mirá el resultado en "Registro de ejecución".
function diagnosticoUnDia() {
  const FECHA_A_REVISAR = null; // ej. '01/08/2026' — si se deja null, usa el día de ayer

  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) { Logger.log('Falta configurar HYROS_API_KEY.'); return; }

  let fecha = FECHA_A_REVISAR;
  if (!fecha) {
    const ayer = new Date();
    ayer.setDate(ayer.getDate() - 1);
    const tz = Session.getScriptTimeZone();
    fecha = Utilities.formatDate(ayer, tz, 'dd/MM/yyyy');
  }

  const partes = fecha.split('/');
  const fromDate = `${partes[2]}-${partes[1]}-${partes[0]}T00:00:00-06:00`;
  const toDate = `${partes[2]}-${partes[1]}-${partes[0]}T23:59:59-06:00`;

  function analizar(nombre, forEachFn) {
    let total = 0;
    let reasignados = 0; // matchean con MÁS de un setter a la vez
    let sinSetter = 0;   // no matchean con ningún setter
    const porSetter = {};
    Object.keys(SHEET_NAMES).forEach(function (setter) { porSetter[setter] = 0; });
    forEachFn(apiKey, fromDate, toDate, (item) => {
      total++;
      const tags = nombre === 'Leads' ? item.tags : (item.lead && item.lead.tags);
      const setters = hyrosSetterForTags(tags);
      if (setters.length > 1) reasignados++;
      if (setters.length === 0) sinSetter++;
      setters.forEach((s) => { porSetter[s]++; });
    });
    const sumaSetters = Object.values(porSetter).reduce((a, b) => a + b, 0);
    Logger.log(`--- ${nombre} (${fecha}) ---`);
    Logger.log(`Total real de Hyros (cada uno contado 1 vez): ${total}`);
    Logger.log(`Suma de los 4 setters (con reasignados contados de más): ${sumaSetters}`);
    Logger.log(`De esos, matchean con más de 1 setter (reasignados): ${reasignados}`);
    Logger.log(`De esos, no matchean con ningún setter: ${sinSetter}`);
    Logger.log(`Por setter: ${JSON.stringify(porSetter)}`);
  }

  analizar('Leads', hyrosForEachLead);
  analizar('Calls (Agendas)', hyrosForEachCall);
  analizar('Sales (Cash)', hyrosForEachSale);
}

// Complementa a diagnosticoUnDia: en vez de solo el total, lista CADA venta
// de un setter puntual para un día puntual, con su precio y sus tags
// completos — para comparar a mano contra lo que muestra la pantalla
// "Sales" de Hyros filtrada por Sources. Si acá aparece una venta de más
// (o con un tag que no está en la lista de Sources guardada en Hyros), esa
// es la diferencia entre lo que muestra el dashboard y lo que muestra esa
// vista filtrada de Hyros.
function diagnosticoVentasSetter() {
  const FECHA_A_REVISAR = '01/08/2026';
  const SETTER_A_REVISAR = 'thomi'; // thomi | flor | valeria | franco

  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) { Logger.log('Falta configurar HYROS_API_KEY.'); return; }

  const partes = FECHA_A_REVISAR.split('/');
  const fromDate = `${partes[2]}-${partes[1]}-${partes[0]}T00:00:00-06:00`;
  const toDate = `${partes[2]}-${partes[1]}-${partes[0]}T23:59:59-06:00`;

  let total = 0;
  let cantidad = 0;
  Logger.log(`--- Ventas de ${SETTER_A_REVISAR} el ${FECHA_A_REVISAR} (según el mismo método que usa el dashboard) ---`);
  hyrosForEachSale(apiKey, fromDate, toDate, (sale) => {
    const tags = sale.lead && sale.lead.tags;
    const setters = hyrosSetterForTags(tags);
    if (setters.indexOf(SETTER_A_REVISAR) === -1) return;
    const price = sale.usdPrice || sale.price || {};
    if (hyrosPrecioExcluido(price.price)) {
      Logger.log(`(excluida por precio $${price.price}, T.I.N.O.) — tags: ${JSON.stringify(tags)}`);
      return;
    }
    const neto = (Number(price.price) || 0) - (Number(price.refunded) || 0);
    cantidad++;
    total += neto;
    Logger.log(`#${cantidad} — $${neto} — tags: ${JSON.stringify(tags)}`);
  });
  Logger.log(`--- Total: ${cantidad} ventas, $${total} ---`);
}

// Lista de Sources que se ve GUARDADA en el filtro de la pantalla "Leads"
// de Hyros (la del panel de la izquierda) — se usa acá solo para marcar,
// tag por tag, cuáles de los que matchea el script NO están en esa lista.
// Si en Hyros agregás/sacás alguna Source de esa lista, actualizala acá
// también para que la comparación siga siendo justa.
const HYROS_SOURCES_GUARDADAS_EN_UI = [
  '_st01', '_st01_', '_yt_st01', '_zoom_st01',
  'ct01_ig_st01', 'ct02_tk_st01', 'ct03_fb_st01', 'ct03_ig_st01', 'ct03_tk_st01',
  'ct07_ig_st01', 'ct07_ig_st011', 'ct07_ig_st01v', 'ct07_tk_st01', 'ct09_ig_st01',
  'ct11_ig_st01', 'ct12_ig_st01', 'ct12_tk_st01', 'ct15_ig_st01', 'ct15_tk_st01',
  'ct17_ig_st01', 'ct18_ig_st01', 'ct19_ig_st01', 'ct21_ig_st01', 'ct22_ig_st01',
  't01_bus_st01', 'v1_video_reel_0307_ct15_st01', 'yo_palabra_clave_dm_ct11_st01',
  'zoom_st01_'
];

// Corré ESTA función (▷ Ejecutar, con "diagnosticoLeadsRango" elegida en el
// desplegable) para comparar, en un RANGO de fechas, el total de leads que
// el script le atribuye a un setter contra lo que muestra la pantalla
// "Leads" de Hyros filtrada por Sources. En vez de listar cada lead (acá
// pueden ser miles), agrupa por TAG único y marca cuáles de esos tags NO
// están en la lista de arriba — esos son justamente los que el filtro
// guardado en Hyros se está perdiendo, y explican la diferencia de total.
//
// OJO: pide los leads DÍA POR DÍA (no el rango entero de una sola vez),
// igual que hace el sistema real (autoSyncHyrosHoy/resyncMesActual) —
// hyrosForEachLead tiene un tope de seguridad de 10.000 registros por
// pedido, y con el volumen de esta cuenta, un rango de varios días junto
// se pasa de ese tope y corta el conteo antes de terminar (dando un total
// mucho más bajo del real). Un día individual nunca se acerca a ese tope.
function diagnosticoLeadsRango() {
  const FECHA_DESDE = '01/08/2026';
  const FECHA_HASTA = '07/08/2026';
  const SETTER_A_REVISAR = 'thomi'; // thomi | flor | valeria | franco

  const apiKey = PropertiesService.getScriptProperties().getProperty('HYROS_API_KEY');
  if (!apiKey) { Logger.log('Falta configurar HYROS_API_KEY.'); return; }

  const tz = Session.getScriptTimeZone();
  const [d1, m1, y1] = FECHA_DESDE.split('/').map(Number);
  const [d2, m2, y2] = FECHA_HASTA.split('/').map(Number);
  const desde = new Date(y1, m1 - 1, d1);
  const hasta = new Date(y2, m2 - 1, d2);

  let total = 0;
  const tagsUnicos = {}; // tag original -> cantidad de leads con ese tag

  for (let d = new Date(desde); d <= hasta; d.setDate(d.getDate() + 1)) {
    const fechaStr = Utilities.formatDate(d, tz, 'dd/MM/yyyy');
    const [dd, mm, yyyy] = fechaStr.split('/');
    const fromDate = `${yyyy}-${mm}-${dd}T00:00:00-06:00`;
    const toDate = `${yyyy}-${mm}-${dd}T23:59:59-06:00`;

    hyrosForEachLead(apiKey, fromDate, toDate, (lead) => {
      const setters = hyrosSetterForTags(lead.tags);
      if (setters.indexOf(SETTER_A_REVISAR) === -1) return;
      total++;
      (lead.tags || []).forEach((t) => {
        const tMin = String(t).toLowerCase();
        if (tMin.indexOf('st01') === -1) return; // solo interesan los tags que matchearon
        tagsUnicos[t] = (tagsUnicos[t] || 0) + 1;
      });
    });
  }

  Logger.log(`--- Leads de ${SETTER_A_REVISAR} entre ${FECHA_DESDE} y ${FECHA_HASTA} (día por día): ${total} ---`);
  Logger.log('Tags que matchearon (✗ = NO está en la lista de Sources guardada en Hyros):');
  Object.keys(tagsUnicos).sort().forEach((tag) => {
    const tagNormalizado = String(tag).toLowerCase().replace(/^[@!$]/, '');
    const enLista = HYROS_SOURCES_GUARDADAS_EN_UI.indexOf(tagNormalizado) !== -1;
    Logger.log(`${enLista ? '✓' : '✗ NO en la lista'} — "${tag}" (${tagsUnicos[tag]} leads)`);
  });
}

