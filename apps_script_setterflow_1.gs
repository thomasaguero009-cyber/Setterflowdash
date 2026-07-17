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
 * (fila 5 = TOTALES/PROMEDIOS, los datos empiezan en la fila 6)
 *
 * INSTALACIÓN:
 * 1. Pegá este código en Extensiones > Apps Script (dentro de tu Google Sheet real).
 * 2. Guardá y Deployá como Web App (ver instrucciones en el chat).
 */

const SHEET_NAMES = {
  thomi:   "THOMI (1)",
  flor:    "FLOR (1)",
  valeria: "VALERIA",
  franco:  "FRANCO"
};

const COLS = {
  fecha: 2,        // B
  seguimientos: 3, // C
  cash: 4,         // D
  horas: 5,        // E
  agendas: 6,      // F
  conversion: 7    // G
};

const FIRST_DATA_ROW = 6; // fila 5 es TOTALES/PROMEDIOS
const USERS_SHEET_NAME = "Usuarios";

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'save';

    if (action === 'register') return handleRegister(body);
    if (action === 'login') return handleLogin(body);
    if (action === 'changepassword') return handleChangePassword(body);
    if (action === 'updatecorreo') return handleUpdateCorreo(body);
    if (action === 'updateprofile') return handleUpdateProfile(body);
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
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    rows = data
      .filter(r => (r[3] || 'Aprobado') === 'Aprobado')
      .map(r => ({ nombre: r[0], apodo: r[5] || '', avatar: r[7] || '', cover: r[8] || '', pensamiento: r[9] || '', nombreMostrar: r[10] || '' }));
  }
  return jsonOut({ status: "ok", rows });
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
    const rows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
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
        return jsonOut({ status: "ok", nombre: rows[i][0], correo: rows[i][2] || '', rol: rows[i][4] || 'Setter', apodo: rows[i][5] || '', telefono: rows[i][6] || '', avatar: rows[i][7] || '', cover: rows[i][8] || '', pensamiento: rows[i][9] || '', nombreMostrar: rows[i][10] || '' });
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
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    rows = data.map(r => ({ nombre: r[0], password: String(r[1] || ''), correo: r[2] || '', estado: r[3] || 'Aprobado', rol: r[4] || 'Setter', apodo: r[5] || '', telefono: r[6] || '', avatar: r[7] || '', cover: r[8] || '', pensamiento: r[9] || '', nombreMostrar: r[10] || '' }));
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
    written.push(setter);
    debug.push({ setter, sheetName, fila: rowIndex, fechaGuardada: fecha });
  });

  return jsonOut({ status: "ok", written, skipped, debug });
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

