// Reemplazo de netlify/functions/hyros-resync.js, corriendo en GitHub Actions
// en vez de Netlify — Netlify quedó pausado por límite de crédito del team,
// así que el sync automático (cada 15 min) se movió acá, que no depende de
// esa cuenta ni de su facturación.
const { getSheetsClient } = require("../lib/sheets");
const { syncHyrosUnaVez } = require("../lib/sync");

const TIMEZONE = process.env.SHEET_TIMEZONE || "America/Bogota";
const DIAS_ATRAS = Number(process.env.HYROS_RESYNC_DIAS_ATRAS || 3);

async function main() {
  const apiKey = process.env.HYROS_API_KEY;
  if (!apiKey) throw new Error("Falta HYROS_API_KEY");

  const sheets = await getSheetsClient();

  for (let i = 0; i < DIAS_ATRAS; i++) {
    const fecha = fechaHaceNDias(i);
    try {
      await syncHyrosUnaVez(sheets, apiKey, fecha);
      console.log("OK " + fecha);
    } catch (err) {
      console.error("ERROR (" + fecha + "): " + err.message);
      process.exitCode = 1;
    }
  }
}

function fechaHaceNDias(n) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());
  const map = {};
  parts.forEach((p) => { map[p.type] = p.value; });
  const d = new Date(Number(map.year), Number(map.month) - 1, Number(map.day));
  d.setDate(d.getDate() - n);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

main().catch((err) => {
  console.error("FALLO:", err);
  process.exit(1);
});
