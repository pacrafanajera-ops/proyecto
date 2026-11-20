/**
 * Backend del formulario Slendytubbies (Zeoworks)
 * - Guarda cada respuesta como una fila en esta hoja de Google Sheets (base de datos)
 * - Envía un correo a tu Gmail con las respuestas
 */

// ======================= CONFIGURACIÓN =======================
const CORREO_DESTINO = "pacrafanajera@gmail.com"; // <- donde quieres recibir las respuestas
const NOMBRE_HOJA    = "Respuestas";
const GUARDAR_CLAVE  = false; // false = la contraseña NO se guarda ni se envía por correo
// =============================================================

// [nombre del campo en el HTML, título de la columna]
const CAMPOS = [
  ["nombre_usuario",     "Nombre de usuario"],
  ["clave_acceso",       "Clave de acceso"],
  ["juego_favorito",     "Entrega favorita"],
  ["dificultad",         "Dificultad"],
  ["monstruo",           "Antagonistas que asustan"],
  ["modo",               "Modos probados"],
  ["mapa_favorito",      "Mapa favorito"],
  ["personaje_jugador",  "Personaje/Skin"],
  ["plataformas",        "Plataformas"],
  ["idiomas",            "Idioma"],
  ["experiencia_terror", "Momento más terrorífico"],
  ["feedback_zeoworks",  "Sugerencias"],
  ["terminos",           "Sobrevivió a la noche"],
  ["boletin",            "Quiere boletín"],
  ["codigo_fan",         "Código comunidad"],
  ["form_version",       "Versión del formulario"],
  ["engine",             "Engine"]
];

function camposActivos() {
  return CAMPOS.filter(c => GUARDAR_CLAVE || c[0] !== "clave_acceso");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // evita que dos envíos simultáneos se pisen
    const p = (e && e.parameters) || {};

    // Anti-spam: si el campo trampa viene relleno, se ignora en silencio
    if (valor(p, "website")) return respuesta("ok");

    // Validación mínima en el servidor
    if (!valor(p, "nombre_usuario")) return respuesta("falta_nombre");

    const campos = camposActivos();
    const datos = campos.map(c => valor(p, c[0]).slice(0, 2000));

    // 1) Guardar en la base de datos (Google Sheets)
    const hoja = obtenerHoja(campos);
    hoja.appendRow([new Date()].concat(datos.map(protegerFormula)));

    // 2) Avisar por Gmail (si falla, la respuesta ya quedó guardada)
    try {
      enviarCorreo(campos, datos);
    } catch (errCorreo) {
      console.error("Error enviando correo: " + errCorreo);
    }

    return respuesta("ok");
  } catch (err) {
    console.error(err);
    return respuesta("error");
  } finally {
    lock.releaseLock();
  }
}

// Abrir la URL /exec en el navegador sirve para comprobar que está activo
function doGet() {
  return respuesta("activo");
}

// Une los valores de campos repetidos (checkbox / listas múltiples)
function valor(p, nombre) {
  return p[nombre] ? p[nombre].join(", ").trim() : "";
}

// Evita que un texto como "=HYPERLINK(...)" se ejecute como fórmula en Sheets
function protegerFormula(texto) {
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}

function obtenerHoja(campos) {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) hoja = libro.insertSheet(NOMBRE_HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha"].concat(campos.map(c => c[1])));
    hoja.getRange(1, 1, 1, campos.length + 1).setFontWeight("bold");
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function enviarCorreo(campos, datos) {
  const zona  = Session.getScriptTimeZone();
  const fecha = Utilities.formatDate(new Date(), zona, "dd/MM/yyyy HH:mm");

  let filas = "";
  campos.forEach((c, i) => {
    filas += "<tr>" +
      "<td style='padding:6px 10px;border:1px solid #ddd;background:#f5f5f5;vertical-align:top'><b>" + escapar(c[1]) + "</b></td>" +
      "<td style='padding:6px 10px;border:1px solid #ddd;white-space:pre-wrap'>" + (escapar(datos[i]) || "<i>(sin respuesta)</i>") + "</td>" +
      "</tr>";
  });

  const html =
    "<h2>Nueva respuesta del formulario Slendytubbies</h2>" +
    "<p>Recibida el " + fecha + "</p>" +
    "<table style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px'>" + filas + "</table>";

  MailApp.sendEmail({
    to: CORREO_DESTINO,
    subject: "Nueva respuesta: " + datos[0],
    htmlBody: html,
    name: "Formulario Zeoworks"
  });
}

function escapar(t) {
  return String(t)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function respuesta(estado) {
  return ContentService.createTextOutput(estado);
}

// Ejecútala UNA vez a mano desde el editor para autorizar los permisos de Gmail/Sheets
function autorizar() {
  MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getActiveSpreadsheet();
}
