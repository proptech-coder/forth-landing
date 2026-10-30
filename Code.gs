/**
 * Forth - Puente entre la landing y Google Sheets.
 *
 * Recibe los envíos del formulario de index.html (POST), los guarda como una
 * fila nueva en la pestaña "Leads" y envía un aviso por correo.
 *
 * Uso:
 *   1. Abre el Sheet > Extensiones > Apps Script, crea un archivo "Code" y pega esto.
 *   2. Ejecuta iniciar() una vez y acepta los permisos.
 *   3. Implementar > Nueva implementación > Aplicación web
 *      (Ejecutar como: Yo | Acceso: Cualquier persona) y copia la URL /exec.
 */

const CONFIG = {
  HOJA: 'Leads',
  CORREO_AVISO: '',              // Vacío = llega a la cuenta que publica el script
  COLUMNAS: ['Fecha', 'Nombre', 'Empresa', 'Email', 'Teléfono', 'Interés', 'Mensaje', 'Origen', 'Estado'],
  ESTADO_INICIAL: 'Nuevo',
  ZONA_HORARIA: 'America/Bogota',
  COLOR_ENCABEZADO: '#0E0F2C',
  FORMATO_FECHA: 'dd/mm/yyyy',
  FILAS_MAX: 1000,               // Filas cubiertas por formatos y listas desplegables
  SHEET_ID: '',                  // Solo si el script NO está vinculado al Sheet (por ejemplo, con clasp)
  INTERESES: ['Automatizaciones', 'Agentes de IA', 'Formación ejecutiva', 'Consultoría estratégica'],
  ORIGENES: ['Sitio web', 'WhatsApp', 'Referido', 'Evento', 'Otro'],
  ESTADOS: ['Nuevo', 'Contactado', 'Propuesta enviada']
};

// Índices de columna (1 = A)
const COL = { FECHA: 1, TELEFONO: 5, INTERES: 6, MENSAJE: 7, ORIGEN: 8, ESTADO: 9 };

/* ------------------------------------------------------------------ */
/* Configuración inicial                                               */
/* ------------------------------------------------------------------ */

/** Paso único: vincula el Sheet, prepara la hoja y envía un correo de prueba. */
function iniciar() {
  const activo = SpreadsheetApp.getActiveSpreadsheet();
  if (!activo && !CONFIG.SHEET_ID) {
    throw new Error('Abre Apps Script desde el Sheet (Extensiones > Apps Script) o completa CONFIG.SHEET_ID.');
  }
  const ss = activo || SpreadsheetApp.openById(CONFIG.SHEET_ID);
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  setup();

  MailApp.sendEmail({
    to: destinatario_(),
    subject: 'Forth leads: conexión lista',
    name: 'Forth leads',
    htmlBody: '<p>El script quedó vinculado a <a href="' + ss.getUrl() + '">' + escapar_(ss.getName()) + '</a>.</p>' +
              '<p>Así se verán los avisos de cada lead nuevo.</p>'
  });

  try {
    SpreadsheetApp.getUi().alert(
      'Vinculado al Sheet',
      'Listo. Te enviamos un correo de prueba.\n\n' +
      'Ahora publica el script: Implementar > Nueva implementación > Aplicación web.\n' +
      'Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.\n' +
      'Copia la URL que termina en /exec y pégala en SCRIPT_URL del index.html.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (err) {
    // Sin interfaz disponible (por ejemplo, ejecutado con clasp): basta con el registro.
    Logger.log('Vinculado a ' + ss.getUrl());
  }
}

/** Crea o actualiza la pestaña Leads: encabezados, formatos y listas desplegables. */
function setup() {
  const ss = hojaDeCalculo_();
  ss.setSpreadsheetTimeZone(CONFIG.ZONA_HORARIA);

  let hoja = ss.getSheetByName(CONFIG.HOJA);
  if (!hoja) hoja = ss.insertSheet(CONFIG.HOJA, 0);

  const n = CONFIG.COLUMNAS.length;
  hoja.getRange(1, 1, 1, n)
    .setValues([CONFIG.COLUMNAS])
    .setFontWeight('bold')
    .setBackground(CONFIG.COLOR_ENCABEZADO)
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 36);

  if (hoja.getMaxRows() < CONFIG.FILAS_MAX) {
    hoja.insertRowsAfter(hoja.getMaxRows(), CONFIG.FILAS_MAX - hoja.getMaxRows());
  }
  const filas = CONFIG.FILAS_MAX - 1;
  hoja.getRange(2, COL.FECHA, filas, 1).setNumberFormat(CONFIG.FORMATO_FECHA);
  hoja.getRange(2, COL.TELEFONO, filas, 1).setNumberFormat('@');
  hoja.getRange(2, COL.MENSAJE, filas, 1).setWrap(true);

  // Interés y Estado: solo valores de la lista. Origen: avisa, pero acepta otros.
  const lista = (valores, estricta) => SpreadsheetApp.newDataValidation()
    .requireValueInList(valores, true)
    .setAllowInvalid(!estricta)
    .build();
  hoja.getRange(2, COL.INTERES, filas, 1).setDataValidation(lista(CONFIG.INTERESES, true));
  hoja.getRange(2, COL.ORIGEN, filas, 1).setDataValidation(lista(CONFIG.ORIGENES, false));
  hoja.getRange(2, COL.ESTADO, filas, 1).setDataValidation(lista(CONFIG.ESTADOS, true));

  [110, 190, 230, 290, 150, 190, 420, 120, 150].forEach((ancho, i) => hoja.setColumnWidth(i + 1, ancho));
  return hoja;
}

/* ------------------------------------------------------------------ */
/* Aplicación web                                                       */
/* ------------------------------------------------------------------ */

/** Recibe el formulario de la landing. */
function doPost(e) {
  const p = (e && e.parameter) || {};
  const lead = {
    nombre: limpiar_(p.nombre),
    empresa: limpiar_(p.empresa),
    email: limpiar_(p.email).toLowerCase(),
    telefono: limpiar_(p.telefono),
    interes: normalizar_(p.interes, CONFIG.INTERESES),
    mensaje: limpiar_(p.mensaje, 2000),
    origen: normalizarOrigen_(p.origen)
  };
  if (!lead.nombre && !lead.email) {
    return json_({ ok: false, error: 'Falta nombre o email' });
  }

  let hoja;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    hoja = hojaDeCalculo_().getSheetByName(CONFIG.HOJA);
    if (!hoja) throw new Error('No existe la pestaña "' + CONFIG.HOJA + '". Ejecuta iniciar().');

    const fila = hoja.getLastRow() + 1;
    hoja.getRange(fila, COL.FECHA).setNumberFormat(CONFIG.FORMATO_FECHA);
    hoja.getRange(fila, COL.TELEFONO).setNumberFormat('@'); // conserva el +57 aunque la fila pase de FILAS_MAX
    hoja.getRange(fila, 1, 1, CONFIG.COLUMNAS.length).setValues([[
      new Date(),
      seguro_(lead.nombre),
      seguro_(lead.empresa),
      seguro_(lead.email),
      lead.telefono,                  // la celda ya es texto plano
      lead.interes,
      seguro_(lead.mensaje),
      seguro_(lead.origen),
      CONFIG.ESTADO_INICIAL
    ]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // El correo va fuera del bloqueo para no frenar otros envíos.
  try {
    avisarLead_(lead, hoja);
  } catch (err) {
    console.error('No se pudo enviar el aviso: ' + err);
  }
  return json_({ ok: true });
}

/** Comprobación rápida: abrir la URL /exec en el navegador. */
function doGet() {
  return ContentService.createTextOutput('Forth leads: activo');
}

/** Prueba el flujo completo desde el editor, sin pasar por la landing. */
function probarLead() {
  const respuesta = doPost({ parameter: {
    nombre: 'Prueba Forth',
    empresa: 'Empresa Demo',
    email: 'prueba@forth.test',
    telefono: '+57 300 000 0000',
    interes: 'Agentes de IA',
    mensaje: 'Lead de prueba enviado desde el editor de Apps Script.',
    origen: 'Sitio web'
  }});
  Logger.log(respuesta.getContent());
}

/* ------------------------------------------------------------------ */
/* Auxiliares                                                           */
/* ------------------------------------------------------------------ */

function hojaDeCalculo_() {
  const id = CONFIG.SHEET_ID || PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const activo = SpreadsheetApp.getActiveSpreadsheet();
  if (!activo) throw new Error('Proyecto no vinculado: ejecuta iniciar() desde el Sheet o completa CONFIG.SHEET_ID.');
  return activo;
}

function destinatario_() {
  return CONFIG.CORREO_AVISO || Session.getEffectiveUser().getEmail();
}

function avisarLead_(lead, hoja) {
  const ss = hoja.getParent();
  const enlace = ss.getUrl() + '#gid=' + hoja.getSheetId();
  const fila = (etiqueta, valor) =>
    '<tr><td style="padding:6px 16px 6px 0;color:#4A4C68;vertical-align:top">' + etiqueta + '</td>' +
    '<td style="padding:6px 0;color:#0E0F2C"><strong>' + (escapar_(valor) || '-') + '</strong></td></tr>';

  MailApp.sendEmail({
    to: destinatario_(),
    replyTo: lead.email || undefined,
    subject: 'Nuevo lead: ' + (lead.nombre || lead.email) + ' - ' + (lead.empresa || 'sin empresa'),
    name: 'Forth leads',
    htmlBody:
      '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.5">' +
      '<h2 style="margin:0 0 12px;color:#0E0F2C">Nuevo lead desde la landing</h2>' +
      '<table style="border-collapse:collapse">' +
      fila('Nombre', lead.nombre) + fila('Empresa', lead.empresa) + fila('Email', lead.email) +
      fila('Teléfono', lead.telefono) + fila('Interés', lead.interes) + fila('Mensaje', lead.mensaje) +
      fila('Origen', lead.origen) +
      '</table>' +
      '<p style="margin-top:20px"><a href="' + enlace + '" style="background:#F26B3A;color:#fff;padding:10px 18px;border-radius:999px;text-decoration:none;font-weight:bold">Abrir el Sheet</a></p>' +
      '</div>'
  });
}

/** Recorta espacios y largo; devuelve '' si no hay valor. */
function limpiar_(valor, max) {
  return String(valor == null ? '' : valor).replace(/\s+/g, ' ').trim().slice(0, max || 300);
}

/** Evita que un texto que empieza con = + - @ se interprete como fórmula. */
function seguro_(texto) {
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}

/** Busca el valor en la lista sin importar mayúsculas ni tildes. */
function normalizar_(valor, lista) {
  const clave = s => limpiar_(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const buscado = clave(valor);
  return lista.find(v => clave(v) === buscado) || '';
}

/** La landing envía su dominio como origen: se guarda como "Sitio web". */
function normalizarOrigen_(valor) {
  const v = limpiar_(valor);
  if (!v) return 'Sitio web';
  const conocido = normalizar_(v, CONFIG.ORIGENES);
  if (conocido) return conocido;
  if (v === 'local' || /^[a-z0-9-]+(\.[a-z0-9-]+)+$/i.test(v)) return 'Sitio web';
  return v;
}

function escapar_(texto) {
  return String(texto == null ? '' : texto)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
