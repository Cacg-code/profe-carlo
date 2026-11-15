/**
 * Profe Carlo · guarda cada test terminado en esta hoja.
 * Se pega en la hoja: Extensiones > Apps Script. Luego Implementar > Nueva implementación > Aplicación web
 * (Ejecutar como: yo · Quién tiene acceso: cualquier usuario) y se copia la URL que termina en /exec.
 */
const HOJA = 'Tests';
const AVISAR_POR_CORREO = true; // te llega un correo por cada test terminado

const COLUMNAS = ['Fecha', 'Alumno', 'Grado', 'Zona', 'WhatsApp', 'Correctas', 'Nota %',
  'Temas a reforzar', 'Temas que domina', 'Recomendación', 'Preguntas falladas', 'Estado'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fila = [
      new Date(),
      limpiar(d.alumno, 60),
      limpiar(d.grado, 10),
      limpiar(d.zona, 40),
      limpiar(String(d.whatsapp || '').replace(/\D/g, ''), 9),
      limpiar(d.correctas, 10),
      Math.max(0, Math.min(100, Number(d.nota) || 0)),
      limpiar(d.reforzar, 200),
      limpiar(d.domina, 200),
      limpiar(d.recomendacion, 60),
      limpiar(d.falladas, 200),
      'Nuevo'
    ];
    if (!fila[1] || !fila[2]) return salida({ ok: false });
    hoja().appendRow(fila);
    if (AVISAR_POR_CORREO) avisar(fila);
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return salida({ ok: true, servicio: 'tests Profe Carlo' });
}

function hoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let h = libro.getSheetByName(HOJA);
  if (!h) {
    h = libro.insertSheet(HOJA);
    h.appendRow(COLUMNAS);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold').setBackground('#0a0a0a').setFontColor('#ffffff');
    h.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
    h.getRange('L2:L').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Nuevo', 'Contactado', 'Clase separada', 'No interesado'], true).build());
    h.autoResizeColumns(1, COLUMNAS.length);
  }
  return h;
}

// Corta el largo y evita que un texto empiece con = + - @ (se tomaría como fórmula).
function limpiar(v, max) {
  let t = String(v == null ? '' : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function avisar(f) {
  const yo = Session.getEffectiveUser().getEmail();
  if (!yo) return;
  const wa = f[4] ? 'https://wa.me/51' + f[4] : 'no dejó número';
  MailApp.sendEmail(yo, 'Nuevo test: ' + f[1] + ' (' + f[2] + ' sec.) ' + f[6] + ' %',
    'Alumno: ' + f[1] + '\nGrado: ' + f[2] + ' sec.\nZona: ' + f[3] + '\nWhatsApp: ' + wa +
    '\nResultado: ' + f[5] + ' (' + f[6] + ' %)\nA reforzar: ' + (f[7] || 'nada') +
    '\nRecomendación: ' + f[9] + '\n\nAbre la hoja para ver todos los tests.');
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Ejecútala una vez desde el editor para crear la hoja y dar los permisos.
function prueba() {
  doPost({ postData: { contents: JSON.stringify({
    alumno: 'Prueba', grado: '3.º', zona: 'Villa Club (Carabayllo)', whatsapp: '',
    correctas: '8/12', nota: 67, reforzar: 'Sistemas, Funciones', domina: 'Potencias',
    recomendacion: 'Pack examen', falladas: '5, 6, 9, 10' }) } });
}
