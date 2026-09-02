// Mismos valores que apps_script_setterflow.gs — portados tal cual, no
// inventar estructura nueva. Si cambia algo del Sheet (nombres de pestaña,
// columnas), hay que actualizarlo ACÁ y en el .gs por igual mientras
// convivan los dos.

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEET_NAMES = {
  thomi: "THOMI (1)",
  flor: "FLOR (1)",
  valeria: "VALERIA",
  franco: "FRANCO",
  paula: "PAULA"
};

// B: Fecha  C: Seguimientos  D: Cash  E: Horas  F: Agendas  G: Conversión  H: Leads
// I: Cash orgánico  J: Cash mixto  K: Ventas — agregadas hoy, para poder
// mostrar el desglose de "Ventas" INDIVIDUAL por setter en "Ver perfil"
// (antes solo existía el agregado de los 4 en la hoja "Ventas (diario)").
// L-P: Agendas por cantidad de estrellas (1 a 5) — Q-U: lo mismo para
// Ventas. Hyros/HighLevel clasifica cada llamada/venta con un tag
// "udc-N-estrellas" (calidad del lead); se desglosa acá para el gráfico
// de "Matemática del éxito" (ver hyrosExtraerEstrellas en lib/hyros.js).
const COL_LETTERS = {
  fecha: "B",
  seguimientos: "C",
  cash: "D",
  horas: "E",
  agendas: "F",
  conversion: "G",
  leads: "H",
  cashOrganico: "I",
  cashMixto: "J",
  ventas: "K",
  agendas1: "L",
  agendas2: "M",
  agendas3: "N",
  agendas4: "O",
  agendas5: "P",
  ventas1: "Q",
  ventas2: "R",
  ventas3: "S",
  ventas4: "T",
  ventas5: "U"
};

const HEADER_ROW = 4; // fila con los títulos de columna (FECHA, SEGUIMIENTOS, ...)
const FIRST_DATA_ROW = 6; // fila 5 es TOTALES/PROMEDIOS

const FUENTES_SHEET_NAME = "Fuentes (diario)";
// G: Setter — nueva. Las filas AGREGADAS (las de siempre) quedan con esta
// celda vacía; además de esas se agregan filas por setter con la misma
// cuenta/canal pero el nombre del setter acá, para poder filtrar "Fuentes"
// por UN setter en "Ver perfil" (antes esa vista quedaba vacía a propósito
// porque este desglose nunca se guardaba, solo se calculaba y se tiraba).
const FUENTES_COLS = { fecha: 1, cuenta: 2, canal: 3, leads: 4, agendas: 5, cash: 6, setter: 7 };

// Hoja nueva — un solo renglón por día con el AGREGADO de los 4 setters
// (no por setter, "cash colected" ya es la suma de los 4). Mismo patrón de
// find-or-create-by-date que las hojas de setter, pero con encabezado en
// la fila 1 y datos desde la fila 2 (no hay bloque de TOTALES/PROMEDIOS
// como en esas).
const VENTAS_SHEET_NAME = "Ventas (diario)";
const VENTAS_FIRST_DATA_ROW = 2;
const VENTAS_COL_LETTERS = { fecha: "A", cashOrganico: "B", cashMixto: "C", ventas: "D" };

module.exports = {
  SPREADSHEET_ID,
  SHEET_NAMES,
  COL_LETTERS,
  HEADER_ROW,
  FIRST_DATA_ROW,
  FUENTES_SHEET_NAME,
  FUENTES_COLS,
  VENTAS_SHEET_NAME,
  VENTAS_FIRST_DATA_ROW,
  VENTAS_COL_LETTERS
};
