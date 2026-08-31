// Puerto 1:1 de la lógica de Hyros de apps_script_setterflow.gs — mismos
// nombres, mismos comentarios explicativos donde aplica, para que se
// pueda comparar línea a línea contra el .gs si hace falta.

const HYROS_BASE_URL = "https://api.hyros.com/v1/api/v1.0/";

// Los "códigos de setter" (ST01/ST05/ST06/ST07) NO aparecen como tags
// exactos en Hyros — aparecen como parte de tags más largos por campaña
// (ej. "@ct15_ig_st05"). Por eso se clasifica mirando si algún tag
// CONTIENE el código, no por match exacto.
const HYROS_SETTER_CODES = { thomi: "st01", flor: "st05", valeria: "st07", franco: "st06" };

const HYROS_ACCOUNT_NAMES = {
  ct03: "tino.mossu",
  ct15: "teotinivelli",
  ct16: "teotinivelliprime",
  ct07: "tinosinfiltro",
  ct12: "tinohub",
  ct11: "tinolifestylee",
  ct23: "gisenriquez.m"
};
const HYROS_CHANNEL_NAMES = {
  ig: "Instagram",
  tk: "TikTok",
  zoom: "Zoom",
  fb: "Facebook"
};

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

function hyrosClasificarFuente(tags) {
  if (tags && tags.length) {
    for (let i = 0; i < tags.length; i++) {
      const partes = String(tags[i]).toLowerCase().replace(/^[@!$]/, "").split(/[_-]/);
      let cuenta = null;
      let canal = null;
      partes.forEach((p) => {
        if (HYROS_ACCOUNT_NAMES[p]) cuenta = HYROS_ACCOUNT_NAMES[p];
        if (HYROS_CHANNEL_NAMES[p]) canal = HYROS_CHANNEL_NAMES[p];
      });
      if (cuenta) {
        const canalFinal = canal || "Otro canal";
        return { cuenta, canal: canalFinal, key: cuenta + " · " + canalFinal };
      }
    }
  }
  return { cuenta: "Otra cuenta", canal: "Otro canal", key: "Otra cuenta · Otro canal" };
}

function hyrosSumarFuente(fuentes, tags, campo, delta) {
  const clasif = hyrosClasificarFuente(tags);
  if (!clasif) return;
  if (!fuentes[clasif.key]) {
    fuentes[clasif.key] = { cuenta: clasif.cuenta, canal: clasif.canal, leads: 0, agendas: 0, cash: 0 };
  }
  fuentes[clasif.key][campo] += delta;
}

// ===== YouTube (tags "mc02") — SOLO para el panel Fuentes =====
// A diferencia de todo lo de arriba (que solo suma lo que ya matcheó a
// alguno de los 4 setters, vía hyrosSetterForTags), esto se fija en
// CUALQUIER lead/venta/llamada que tenga un tag "mc02", tenga o no tenga
// además un setter asignado — es un eje de clasificación aparte (de dónde
// vino el tráfico), no reemplaza ni se resta de la facturación de los
// setters en ningún lado. "mc02_colab_v01" es un caso particular (la
// colaboración) que se separa del resto de mc02 ("YouTube general") a
// propósito, para no mezclar los dos números.
function hyrosEsMc02Colab(tags) {
  if (!tags || !tags.length) return false;
  return tags.some((t) => String(t).toLowerCase().includes("mc02_colab_v01"));
}
function hyrosEsMc02(tags) {
  if (!tags || !tags.length) return false;
  return tags.some((t) => String(t).toLowerCase().includes("mc02"));
}

function hyrosSumarYoutube(fuentesYoutube, tags, campo, delta) {
  if (!hyrosEsMc02(tags)) return; // no tiene ningún tag mc02 — no es de YouTube, no se toca
  const canal = hyrosEsMc02Colab(tags) ? "Colaboración" : "General";
  const key = "YouTube · " + canal;
  if (!fuentesYoutube[key]) {
    fuentesYoutube[key] = { cuenta: "YouTube", canal, leads: 0, agendas: 0, cash: 0 };
  }
  fuentesYoutube[key][campo] += delta;
}

// ===== Zoom (tag "zoom_st01") — SOLO para el panel Fuentes =====
// Mismo patrón que YouTube/mc02 de arriba: se suma en Fuentes sin importar
// si matcheó a un setter o no. La diferencia (y el motivo de que exista
// aparte, en vez de dejar que hyrosClasificarFuente lo agarre solo) es que
// "zoom_st01" CONTIENE el código de setter de Thomi ("st01") — por eso, a
// pedido explícito, esta plata NO debe sumarse al Cash Collected de los
// setters aunque el tag matchee ese código (ver el "esZoom" en el loop de
// ventas más abajo). Leads/Agendas de estos tags sí siguen yendo al setter
// normalmente — solo el cash está excluido.
function hyrosEsZoom(tags) {
  if (!tags || !tags.length) return false;
  return tags.some((t) => String(t).toLowerCase().includes("zoom_st01"));
}

function hyrosSumarZoom(fuentesZoom, tags, campo, delta) {
  if (!hyrosEsZoom(tags)) return;
  const key = "Zoom · General";
  if (!fuentesZoom[key]) {
    fuentesZoom[key] = { cuenta: "Zoom", canal: "General", leads: 0, agendas: 0, cash: 0 };
  }
  fuentesZoom[key][campo] += delta;
}

// ===== Orgánico puro vs. mixto =====
// Hyros ya trae, por venta, "firstSource"/"lastSource" (primer y último
// touch de atribución), cada uno con un booleano "organic" propio. Es
// orgánico puro solo si AMBOS touches fueron orgánicos — si cualquiera de
// los dos fue pago (o falta el dato), se clasifica como mixto. Esto
// reemplaza con datos reales de Hyros el criterio de "mirar el tag a
// mano" (ct03_ig_st01 vs. "ARG - OPEN") que describió el cliente: el
// booleano "organic" de Hyros ya captura exactamente esa misma distinción.
function hyrosEsOrganicoPuro(sale) {
  const fs = sale && sale.firstSource;
  const ls = sale && sale.lastSource;
  return !!(fs && fs.organic) && !!(ls && ls.organic);
}

// Precios que NO cuentan como facturación real (UDC Closer Elite vendido a
// $1/$17 en pasos bajos del funnel) — mismo criterio que el .gs.
function hyrosPrecioExcluido(precio) {
  const p = Number(precio);
  if (isNaN(p)) return false;
  return Math.abs(p - 1) < 0.01 || Math.abs(p - 17) < 0.01;
}

async function hyrosApiGet(apiKey, path, params) {
  const query = Object.keys(params)
    .map((k) => encodeURIComponent(k) + "=" + encodeURIComponent(params[k]))
    .join("&");
  const url = HYROS_BASE_URL + path + "?" + query;
  const res = await fetch(url, { method: "GET", headers: { "API-Key": apiKey } });
  if (!res.ok) {
    const text = await res.text();
    throw new Error("Hyros " + path + " devolvió " + res.status + ": " + text);
  }
  return res.json();
}

// Mismo tope de seguridad de 40 páginas (10.000 registros) que el .gs.
async function hyrosForEachPage(apiKey, path, fromDate, toDate, cb) {
  let pageId = null;
  let guard = 0;
  do {
    const params = { fromDate, toDate, pageSize: 250 };
    if (pageId) params.pageId = pageId;
    const resp = await hyrosApiGet(apiKey, path, params);
    (resp.result || []).forEach(cb);
    pageId = resp.nextPageId || null;
    guard++;
  } while (pageId && guard < 40);
}

const hyrosForEachLead = (apiKey, fromDate, toDate, cb) => hyrosForEachPage(apiKey, "leads", fromDate, toDate, cb);
const hyrosForEachSale = (apiKey, fromDate, toDate, cb) => hyrosForEachPage(apiKey, "sales", fromDate, toDate, cb);
const hyrosForEachCall = (apiKey, fromDate, toDate, cb) => hyrosForEachPage(apiKey, "calls", fromDate, toDate, cb);

// "fecha"/"fechaHasta" en formato dd/mm/aaaa. Pide Hyros DÍA POR DÍA (no un
// solo rango ancho) por el mismo motivo que el .gs: el tope de 10.000
// registros por pedido se pasa fácil en un rango de varios días con el
// volumen de esta cuenta.
//
// OJO — huso horario: la cuenta de Hyros de Factor Studios está en -06:00
// (confirmado contra /api/v1.0/user-info en su momento). Pedir el rango en
// UTC corre la ventana 6 horas y desalinea los días.
async function hyrosFetchDataForDate(apiKey, fecha, fechaHasta) {
  const partes = fecha.split("/");
  if (partes.length !== 3) throw new Error("Formato de fecha inválido, esperado dd/mm/aaaa");
  const partesHasta = (fechaHasta || fecha).split("/");
  if (partesHasta.length !== 3) throw new Error("Formato de fechaHasta inválido, esperado dd/mm/aaaa");

  const desde = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
  const hasta = new Date(Number(partesHasta[2]), Number(partesHasta[1]) - 1, Number(partesHasta[0]));
  const dias = [];
  for (let d = new Date(desde); d <= hasta; d.setDate(d.getDate() + 1)) {
    dias.push({ dd: String(d.getDate()).padStart(2, "0"), mm: String(d.getMonth() + 1).padStart(2, "0"), yyyy: d.getFullYear() });
  }

  const data = {
    thomi: { cash: 0, leads: 0, agendas: 0, ventas: 0, cashOrganico: 0, cashMixto: 0 },
    flor: { cash: 0, leads: 0, agendas: 0, ventas: 0, cashOrganico: 0, cashMixto: 0 },
    valeria: { cash: 0, leads: 0, agendas: 0, ventas: 0, cashOrganico: 0, cashMixto: 0 },
    franco: { cash: 0, leads: 0, agendas: 0, ventas: 0, cashOrganico: 0, cashMixto: 0 }
  };
  const fuentes = {};
  const fuentesPorSetter = { thomi: {}, flor: {}, valeria: {}, franco: {} };
  // Aparte de todo lo de arriba (que solo cuenta setters) — YouTube (mc02)
  // y Zoom (zoom_st01) se suman sin importar si matcheó a un setter o no.
  // Ver hyrosSumarYoutube / hyrosSumarZoom.
  const fuentesYoutube = {};
  const fuentesZoom = {};

  for (const f of dias) {
    const fromDate = `${f.yyyy}-${f.mm}-${f.dd}T00:00:00-06:00`;
    const toDate = `${f.yyyy}-${f.mm}-${f.dd}T23:59:59-06:00`;

    await hyrosForEachLead(apiKey, fromDate, toDate, (lead) => {
      const setters = hyrosSetterForTags(lead.tags);
      setters.forEach((setter) => {
        data[setter].leads += 1;
        hyrosSumarFuente(fuentesPorSetter[setter], lead.tags, "leads", 1);
        hyrosSumarFuente(fuentes, lead.tags, "leads", 1);
      });
      hyrosSumarYoutube(fuentesYoutube, lead.tags, "leads", 1);
      hyrosSumarZoom(fuentesZoom, lead.tags, "leads", 1);
    });

    await hyrosForEachSale(apiKey, fromDate, toDate, (sale) => {
      const leadTags = sale.lead && sale.lead.tags;
      const price = sale.usdPrice || sale.price || {};
      if (hyrosPrecioExcluido(price.price)) return;
      const neto = (Number(price.price) || 0) - (Number(price.refunded) || 0);
      // "zoom_st01" contiene el código de Thomi ("st01") y esa venta SÍ es
      // de él — cuenta normal para su Cash Collected (y por lo tanto para
      // la facturación total). Zoom solo se trackea aparte en Fuentes (ver
      // hyrosSumarZoom más abajo), no se excluye de ningún setter.
      const setters = hyrosSetterForTags(leadTags);
      const esOrganicoPuro = hyrosEsOrganicoPuro(sale);
      setters.forEach((setter) => {
        data[setter].cash += neto;
        data[setter].ventas += 1;
        if (esOrganicoPuro) data[setter].cashOrganico += neto;
        else data[setter].cashMixto += neto;
        hyrosSumarFuente(fuentesPorSetter[setter], leadTags, "cash", neto);
        hyrosSumarFuente(fuentes, leadTags, "cash", neto);
      });
      hyrosSumarYoutube(fuentesYoutube, leadTags, "cash", neto);
      hyrosSumarZoom(fuentesZoom, leadTags, "cash", neto);
    });

    // "Agendas" = llamadas ("calls") agendadas en Hyros.
    await hyrosForEachCall(apiKey, fromDate, toDate, (call) => {
      const leadTags = call.lead && call.lead.tags;
      const setters = hyrosSetterForTags(leadTags);
      setters.forEach((setter) => {
        data[setter].agendas += 1;
        hyrosSumarFuente(fuentesPorSetter[setter], leadTags, "agendas", 1);
        hyrosSumarFuente(fuentes, leadTags, "agendas", 1);
      });
      hyrosSumarYoutube(fuentesYoutube, leadTags, "agendas", 1);
      hyrosSumarZoom(fuentesZoom, leadTags, "agendas", 1);
    });
  }

  return { data, fuentes, fuentesPorSetter, fuentesYoutube, fuentesZoom };
}

module.exports = {
  HYROS_SETTER_CODES,
  hyrosSetterForTags,
  hyrosClasificarFuente,
  hyrosSumarFuente,
  hyrosEsZoom,
  hyrosEsOrganicoPuro,
  hyrosPrecioExcluido,
  hyrosApiGet,
  hyrosFetchDataForDate
};
