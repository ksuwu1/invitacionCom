// ════════════════════════════════════════════════════════════════════
//  PASO A PASO: Conectar la invitación con Google Sheets
// ════════════════════════════════════════════════════════════════════
//
//  1. Ve a Google Sheets: https://sheets.google.com
//     Crea una hoja nueva (nombre sugerido: "Confirmaciones Primera Comunión")
//
//  2. En el menú superior elige:
//     Extensiones  →  Apps Script
//
//  3. Borra todo el código que aparece por defecto y pega
//     TODO el contenido de este archivo.
//
//  4. Guarda el proyecto (Ctrl + S).
//     Ponle cualquier nombre (ej. "RSVP Primera Comunión").
//
//  5. Haz clic en  Implementar  →  Nueva implementación
//        · Tipo:              Aplicación web
//        · Descripción:       (cualquier texto)
//        · Ejecutar como:     Yo (tu cuenta de Google)
//        · Quién tiene acceso: Cualquier usuario
//     → Haz clic en "Implementar"
//
//  6. Google pedirá que autorices la app.
//     Haz clic en "Revisar permisos" → elige tu cuenta → "Permitir".
//
//  7. Copia la URL que aparece (termina en /exec). Ejemplo:
//     https://script.google.com/macros/s/AKfycb.../exec
//
//  8. Abre el archivo  index.html  y localiza la línea:
//        const APPS_SCRIPT_URL = 'TU_URL_AQUI';
//     Reemplaza  TU_URL_AQUI  con la URL que copiaste.
//
//  ¡Listo! Cada vez que alguien confirme asistencia, los datos
//  aparecerán automáticamente en tu Google Sheet.
// ════════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Crear fila de encabezados la primera vez
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Fecha y Hora',
        'Nombre del Invitado',
        'Adultos',
        'Niños',
        'Total Personas'
      ]);

      const headers = sheet.getRange(1, 1, 1, 5);
      headers.setFontWeight('bold');
      headers.setBackground('#8B6914');
      headers.setFontColor('#FFFFFF');
      headers.setHorizontalAlignment('center');
      headers.setFontSize(11);
    }

    // Leer los datos enviados desde la invitación
    const nombre  = (e.parameter.nombre  || '').trim();
    const adultos = parseInt(e.parameter.adultos) || 1;
    const ninos   = parseInt(e.parameter.ninos)   || 0;
    const total   = adultos + ninos;

    // Timestamp en zona horaria de Aguascalientes
    const timestamp = Utilities.formatDate(
      new Date(),
      'America/Mexico_City',
      'dd/MM/yyyy HH:mm:ss'
    );

    // Agregar la fila con los datos
    sheet.appendRow([timestamp, nombre, adultos, ninos, total]);

    // Ajustar ancho de columnas automáticamente
    sheet.autoResizeColumns(1, 5);

    // Centrar columnas numéricas
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3, 1, 3).setHorizontalAlignment('center');

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permite verificar que el script está activo (visita la URL en el navegador)
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', mensaje: 'Script activo ✓' }))
    .setMimeType(ContentService.MimeType.JSON);
}
