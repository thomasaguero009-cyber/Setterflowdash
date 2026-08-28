const { google } = require("googleapis");

let cachedClient = null;

// Autenticación por cuenta de servicio — el Sheet tiene que estar
// compartido con el email de esa cuenta (con permiso de Editor), igual que
// se comparte con una persona más. Las credenciales viven en variables de
// entorno de Netlify, nunca en el código.
async function getSheetsClient() {
  if (cachedClient) return cachedClient;

  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const rawKey = process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY;
  if (!email || !rawKey) {
    throw new Error("Faltan GOOGLE_SERVICE_ACCOUNT_EMAIL / GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY en las variables de entorno de Netlify.");
  }
  // Netlify guarda saltos de línea de una variable de entorno como "\n"
  // literal (texto), no como salto real — hay que reconvertirlos o la
  // clave privada no es válida.
  const privateKey = rawKey.replace(/\\n/g, "\n");

  const auth = new google.auth.JWT(email, null, privateKey, ["https://www.googleapis.com/auth/spreadsheets"]);
  await auth.authorize();
  cachedClient = google.sheets({ version: "v4", auth });
  return cachedClient;
}

// Los nombres de pestaña tienen espacios/paréntesis ("THOMI (1)") — en
// notación A1 hace falta encerrarlos entre comillas simples para que no
// se rompa el rango.
function quoteSheetName(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

module.exports = { getSheetsClient, quoteSheetName };
