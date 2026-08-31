// Reintento compartido entre run-resync.js y run-resync-mes.js. Antes cada
// uno reintentaba UNA sola vez tras 10s — insuficiente cuando los dos
// workflows quedaban compitiendo por la misma cuota de lecturas/min de
// Sheets API por varios minutos seguidos. Con el concurrency group de los
// workflows (ver .github/workflows/*.yml) ya no deberían pisarse, pero
// esto queda como red de seguridad extra ante cualquier pico de cuota.
async function conReintentoSiCuota(fn, { intentos = 4, esperaBaseMs = 5000 } = {}) {
  let ultimoError;
  for (let intento = 1; intento <= intentos; intento++) {
    try {
      return await fn();
    } catch (err) {
      ultimoError = err;
      if (!/Quota exceeded/i.test(err.message) || intento === intentos) throw err;
      const espera = esperaBaseMs * intento; // 5s, 10s, 15s...
      await sleep(espera);
    }
  }
  throw ultimoError;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

module.exports = { conReintentoSiCuota, sleep };
