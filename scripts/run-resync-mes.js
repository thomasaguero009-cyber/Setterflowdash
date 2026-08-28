// Reemplazo de netlify/functions/hyros-resync-mes-background.js en GitHub
// Actions. Recorre todo el mes actual hasta "hoy - HYROS_RESYNC_DIAS_ATRAS"
// (los días recientes ya los cubre run-resync.js) para corregir Cash/Agendas
// de días viejos si algo quedó desalineado (ej. un reembolso tardío).
const { getSheetsClient } = require("../lib/sheets");
const { syncHyrosUnaVez } = require("../lib/sync");

const TIMEZONE = process.env.SHEET_TIMEZONE || "America/Bogota";
const DIAS_ATRAS = Number(process.env.HYROS_RESYNC_DIAS_ATRAS || 3);

async function main() {
  const apiKey = process.env.HYROS_API_KEY;
  if (!apiKey) throw new Error("Falta HYROS_API_KEY");

  const dias = diasDelMesAResincronizar();
  if (!dias.length) {
    console.log("No hay días de este mes para re-sincronizar aparte de los que ya cubre run-resync.js.");
    return;
  }

  const sheets = await getSheetsClient();
  console.log(`Re-sincronizando ${dias.length} día(s): ${dias.join(", ")}`);

  for (let i = 0; i < dias.length; i++) {
    if (i > 0) await sleep(1500); // evita el rate-limit de lecturas/min de Sheets API
    const fecha = dias[i];
    try {
      await syncHyrosUnaVez(sheets, apiKey, fecha);
      console.log("OK " + fecha);
    } catch (err) {
      console.error("ERROR (" + fecha + "): " + err.message);
      process.exitCode = 1;
    }
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function hoyEnTZ() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());
  const map = {};
  parts.forEach((p) => { map[p.type] = p.value; });
  return new Date(Number(map.year), Number(map.month) - 1, Number(map.day));
}

function diasDelMesAResincronizar() {
  const hoy = hoyEnTZ();
  const primerDiaDelMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const ultimoDia = new Date(hoy);
  ultimoDia.setDate(ultimoDia.getDate() - DIAS_ATRAS);

  if (primerDiaDelMes > ultimoDia) return [];

  const dias = [];
  for (let d = new Date(primerDiaDelMes); d <= ultimoDia; d.setDate(d.getDate() + 1)) {
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    dias.push(`${dd}/${mm}/${d.getFullYear()}`);
  }
  return dias;
}

main().catch((err) => {
  console.error("FALLO:", err);
  process.exit(1);
});
