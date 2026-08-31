const { quoteSheetName } = require("./sheets");
const {
  SPREADSHEET_ID,
  SHEET_NAMES,
  COL_LETTERS,
  HEADER_ROW,
  FIRST_DATA_ROW,
  FUENTES_SHEET_NAME,
  VENTAS_SHEET_NAME,
  VENTAS_FIRST_DATA_ROW,
  VENTAS_COL_LETTERS
} = require("./constants");

// Mismo parser de fechas que normalizeDate() en el .gs — soporta
// dd/mm/aaaa, dd-mm-aaaa y aaaa-mm-dd, con o sin ceros adelante.
function normalizeDate(val) {
  const str = String(val).trim();

  let match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (match) {
    const yyyy = match[1];
    const mm = match[2].padStart(2, "0");
    const dd = match[3].padStart(2, "0");
    return `${dd}/${mm}/${yyyy}`;
  }

  match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    const dd = match[1].padStart(2, "0");
    const mm = match[2].padStart(2, "0");
    let yyyy = match[3];
    if (yyyy.length === 2) yyyy = "20" + yyyy;
    return `${dd}/${mm}/${yyyy}`;
  }

  return str;
}

// "Hoy" en formato dd/mm/aaaa, calculado en el huso horario que se le
// indique (ver SHEET_TIMEZONE en el README) — importante para no
// escribir en la fila del día equivocado cerca de la medianoche.
function hoyDDMMYYYY(timeZone) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());
  const map = {};
  parts.forEach((p) => { map[p.type] = p.value; });
  return `${map.day}/${map.month}/${map.year}`;
}

// Mismo comportamiento que findOrCreateDateRow() del .gs: busca la fila
// cuya columna Fecha (B) matchea, y si no existe, la crea justo después
// de la última fila con datos (no rellena huecos intermedios).
async function findOrCreateDateRow(sheets, sheetName, fecha, opts) {
  const firstRow = (opts && opts.firstRow) || FIRST_DATA_ROW;
  const fechaCol = (opts && opts.fechaCol) || "B";
  const q = quoteSheetName(sheetName);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${q}!${fechaCol}${firstRow}:${fechaCol}2000`,
    valueRenderOption: "FORMATTED_VALUE"
  });
  const rows = res.data.values || [];
  for (let i = 0; i < rows.length; i++) {
    const cell = rows[i] && rows[i][0];
    if (!cell) continue;
    if (normalizeDate(cell) === fecha) return firstRow + i;
  }
  const newRow = firstRow + rows.length;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${q}!${fechaCol}${newRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[fecha]] }
  });
  return newRow;
}

// Suma "delta" a UNA celda del día de hoy (leads/cash/agendas) — usado por
// el webhook para el conteo en vivo. A diferencia del .gs no hay
// LockService acá (no existe un equivalente directo en Netlify Functions,
// que son sin estado entre invocaciones); el riesgo de dos avisos
// pisándose entre sí en el mismo instante es bajo, y el resync periódico
// (cada 15 min) recalcula el día entero desde cero y corrige cualquier
// diferencia igual que hacía antes.
async function incrementarCeldaHoy(sheets, setter, colKey, delta, timeZone) {
  const sheetName = SHEET_NAMES[setter];
  const q = quoteSheetName(sheetName);
  const fecha = hoyDDMMYYYY(timeZone);
  const rowIndex = await findOrCreateDateRow(sheets, sheetName, fecha);
  const colLetter = COL_LETTERS[colKey];
  const cellRange = `${q}!${colLetter}${rowIndex}`;

  const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: cellRange });
  const actual = Number((cur.data.values && cur.data.values[0] && cur.data.values[0][0]) || 0) || 0;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: cellRange,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[actual + delta]] }
  });

  await asegurarFormulaConversion(sheets, q, rowIndex);
}

async function asegurarFormulaConversion(sheets, q, rowIndex) {
  const convRange = `${q}!${COL_LETTERS.conversion}${rowIndex}`;
  const cur = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: convRange,
    valueRenderOption: "FORMULA"
  });
  const already = cur.data.values && cur.data.values[0] && cur.data.values[0][0];
  if (!already) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: convRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[`=${COL_LETTERS.agendas}${rowIndex}/${COL_LETTERS.seguimientos}${rowIndex}`]] }
    });
  }
}

// Encabezados de I/J/K (Cash orgánico / Cash mixto / Ventas) — no existían
// en la hoja original de cada setter, se agregan solos la primera vez que
// hace falta, sin pisar nada si ya están puestos (mismo criterio que
// asegurarFormulaConversion). El resync de "todo el mes" llama a esto una
// vez POR DÍA para las mismas 4 hojas de setter — sin este caché, terminaba
// releyendo el mismo encabezado ya confirmado 25+ veces por corrida,
// gastando cuota de lectura para nada. El caché vive en memoria y se
// reinicia solo (cada corrida de GitHub Actions es un proceso nuevo).
const headersVentasYaConfirmados = new Set();

async function asegurarHeaderVentasIndividual(sheets, q) {
  if (headersVentasYaConfirmados.has(q)) return;
  const headerRange = `${q}!${COL_LETTERS.cashOrganico}${HEADER_ROW}:${COL_LETTERS.ventas}${HEADER_ROW}`;
  const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: headerRange });
  const already = cur.data.values && cur.data.values[0] && cur.data.values[0][0];
  if (!already) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: headerRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["CASH ORGÁNICO", "CASH MIXTO", "VENTAS"]] }
    });
  }
  headersVentasYaConfirmados.add(q);
}

// Sobrescribe Cash/Agendas/Leads (+ el desglose orgánico/mixto/ventas) del
// día ENTERO (no suma, reemplaza) con el total real recalculado desde
// Hyros — usado por el resync periódico, igual que syncHyrosUnaVez() en el
// .gs. No toca Seguimientos (C) ni Horas (E), esas siguen siendo
// manuales/de la extensión.
async function escribirTotalesDelDia(sheets, setter, fecha, info) {
  const sheetName = SHEET_NAMES[setter];
  const q = quoteSheetName(sheetName);
  const rowIndex = await findOrCreateDateRow(sheets, sheetName, fecha);

  const data = [
    { range: `${q}!${COL_LETTERS.cash}${rowIndex}`, values: [[info.cash]] },
    { range: `${q}!${COL_LETTERS.agendas}${rowIndex}`, values: [[info.agendas]] },
    { range: `${q}!${COL_LETTERS.leads}${rowIndex}`, values: [[info.leads]] },
    { range: `${q}!${COL_LETTERS.cashOrganico}${rowIndex}`, values: [[info.cashOrganico || 0]] },
    { range: `${q}!${COL_LETTERS.cashMixto}${rowIndex}`, values: [[info.cashMixto || 0]] },
    { range: `${q}!${COL_LETTERS.ventas}${rowIndex}`, values: [[info.ventas || 0]] }
  ];
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: "USER_ENTERED", data }
  });

  await asegurarFormulaConversion(sheets, q, rowIndex);
  await asegurarHeaderVentasIndividual(sheets, q);
}

// Caché por corrida — igual motivo que headersVentasYaConfirmados: sin
// esto, cada día del resync de "todo el mes" volvía a pedir los metadatos
// completos del spreadsheet solo para confirmar que "Fuentes (diario)" y
// "Ventas (diario)" ya existen (algo que no cambia en el medio de UNA
// corrida).
const sheetIdCache = new Map();

async function getSheetIdByName(sheets, sheetName) {
  if (sheetIdCache.has(sheetName)) return sheetIdCache.get(sheetName);
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: "sheets.properties" });
  (meta.data.sheets || []).forEach((s) => sheetIdCache.set(s.properties.title, s.properties.sheetId));
  return sheetIdCache.has(sheetName) ? sheetIdCache.get(sheetName) : null;
}

async function ensureFuentesSheet(sheets) {
  const existingId = await getSheetIdByName(sheets, FUENTES_SHEET_NAME);
  if (existingId !== null) return existingId;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: FUENTES_SHEET_NAME } } }] }
  });
  const newSheetId = res.data.replies[0].addSheet.properties.sheetId;
  sheetIdCache.set(FUENTES_SHEET_NAME, newSheetId);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(FUENTES_SHEET_NAME)}!A1:G1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [["Fecha", "Cuenta", "Canal", "Leads", "Agendas", "Cash", "Setter"]] }
  });
  return newSheetId;
}

// "Fuentes (diario)" ya existía con 6 columnas (A-F) antes de que se
// agregara el desglose por setter — esto le agrega el título de la G solo
// si todavía no está (mismo criterio que asegurarHeaderVentasIndividual,
// mismo caché por corrida para no releer lo mismo un día tras otro).
let headerSetterFuentesYaConfirmado = false;

async function asegurarHeaderSetterFuentes(sheets, q) {
  if (headerSetterFuentesYaConfirmado) return;
  const headerRange = `${q}!G1`;
  const cur = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: headerRange });
  const already = cur.data.values && cur.data.values[0] && cur.data.values[0][0];
  if (!already) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: headerRange,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [["Setter"]] }
    });
  }
  headerSetterFuentesYaConfirmado = true;
}

// Mismo patrón "borrar filas de esa fecha, después agregar de nuevo" que
// guardarFuentesDelDia() en el .gs — así cada resync deja la fecha al día
// sin ir acumulando filas duplicadas. "fuentesPorSetter" es opcional
// ({ thomi: {...}, flor: {...}, ... }, mismo shape que "fuentes") — si se
// pasa, además de las filas agregadas de siempre (columna Setter vacía) se
// agregan filas por setter (columna Setter con su nombre), para que "Ver
// perfil" pueda filtrar Fuentes por UNA sola persona.
async function guardarFuentesDelDia(sheets, fecha, fuentes, fuentesPorSetter) {
  const sheetId = await ensureFuentesSheet(sheets);
  const q = quoteSheetName(FUENTES_SHEET_NAME);
  await asegurarHeaderSetterFuentes(sheets, q);

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${q}!A2:A10000`,
    valueRenderOption: "FORMATTED_VALUE"
  });
  const fechas = res.data.values || [];
  const filasABorrar = [];
  fechas.forEach((row, i) => { if (row[0] === fecha) filasABorrar.push(2 + i); });

  if (filasABorrar.length) {
    const requests = filasABorrar
      .sort((a, b) => b - a) // de abajo hacia arriba para no correr los índices de las que faltan borrar
      .map((row) => ({ deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: row - 1, endIndex: row } } }));
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests } });
  }

  const filas = Object.keys(fuentes || {}).map((key) => {
    const f = fuentes[key];
    return [fecha, f.cuenta, f.canal, f.leads || 0, f.agendas || 0, f.cash || 0, ""];
  });
  Object.keys(fuentesPorSetter || {}).forEach((setter) => {
    Object.keys(fuentesPorSetter[setter] || {}).forEach((key) => {
      const f = fuentesPorSetter[setter][key];
      filas.push([fecha, f.cuenta, f.canal, f.leads || 0, f.agendas || 0, f.cash || 0, setter]);
    });
  });
  if (filas.length) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${q}!A1`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: filas }
    });
  }
}

async function ensureVentasSheet(sheets) {
  const existingId = await getSheetIdByName(sheets, VENTAS_SHEET_NAME);
  if (existingId !== null) return existingId;
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests: [{ addSheet: { properties: { title: VENTAS_SHEET_NAME } } }] }
  });
  const newSheetId = res.data.replies[0].addSheet.properties.sheetId;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${quoteSheetName(VENTAS_SHEET_NAME)}!A1:D1`,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [["Fecha", "Cash orgánico", "Cash mixto", "Ventas"]] }
  });
  return newSheetId;
}

// Un solo renglón agregado (suma de los 4 setters) por día — sobrescribe,
// no suma, igual que escribirTotalesDelDia. "info" = { cashOrganico,
// cashMixto, ventas }.
async function escribirVentasDelDia(sheets, fecha, info) {
  await ensureVentasSheet(sheets);
  const q = quoteSheetName(VENTAS_SHEET_NAME);
  const rowIndex = await findOrCreateDateRow(sheets, VENTAS_SHEET_NAME, fecha, {
    firstRow: VENTAS_FIRST_DATA_ROW,
    fechaCol: VENTAS_COL_LETTERS.fecha
  });

  const data = [
    { range: `${q}!${VENTAS_COL_LETTERS.cashOrganico}${rowIndex}`, values: [[info.cashOrganico || 0]] },
    { range: `${q}!${VENTAS_COL_LETTERS.cashMixto}${rowIndex}`, values: [[info.cashMixto || 0]] },
    { range: `${q}!${VENTAS_COL_LETTERS.ventas}${rowIndex}`, values: [[info.ventas || 0]] }
  ];
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { valueInputOption: "USER_ENTERED", data }
  });
}

module.exports = {
  hoyDDMMYYYY,
  findOrCreateDateRow,
  incrementarCeldaHoy,
  escribirTotalesDelDia,
  guardarFuentesDelDia,
  escribirVentasDelDia
};
