const { SHEET_NAMES } = require("./constants");
const { hyrosFetchDataForDate } = require("./hyros");
const { escribirTotalesDelDia, guardarFuentesDelDia, escribirVentasDelDia } = require("./sheetWrite");

// Compartido entre hyros-resync (últimos días) y hyros-resync-mes-background
// (todo el mes) — recalcula Cash/Agendas/Leads de UN día desde cero contra
// Hyros y los sobrescribe, igual que syncHyrosUnaVez() en el .gs.
async function syncHyrosUnaVez(sheets, apiKey, fecha) {
  const resultado = await hyrosFetchDataForDate(apiKey, fecha);
  const data = resultado.data;

  for (const setter of Object.keys(SHEET_NAMES)) {
    await escribirTotalesDelDia(sheets, setter, fecha, data[setter]);
  }

  // YouTube (mc02) y Zoom (zoom_st01) se guardan en la MISMA hoja "Fuentes
  // (diario)", como filas extra ("YouTube · ...", "Zoom · General") — no
  // pisan nada de lo de los setters porque son claves distintas a las que
  // ya arma hyrosClasificarFuente (cuenta · canal).
  const fuentesConExtras = Object.assign({}, resultado.fuentes, resultado.fuentesYoutube, resultado.fuentesZoom);
  await guardarFuentesDelDia(sheets, fecha, fuentesConExtras, resultado.fuentesPorSetter);

  // "Ventas (diario)" — agregado de los 4 setters (orgánico puro vs.
  // mixto + cantidad de ventas), para el panel nuevo del dashboard.
  const ventasAgregado = Object.keys(SHEET_NAMES).reduce(
    (acc, setter) => {
      const d = data[setter];
      acc.cashOrganico += d.cashOrganico || 0;
      acc.cashMixto += d.cashMixto || 0;
      acc.ventas += d.ventas || 0;
      return acc;
    },
    { cashOrganico: 0, cashMixto: 0, ventas: 0 }
  );
  await escribirVentasDelDia(sheets, fecha, ventasAgregado);
}

module.exports = { syncHyrosUnaVez };
