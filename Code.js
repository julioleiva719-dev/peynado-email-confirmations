// ======= CONFIGURACIÓN GENERAL =======
const NOMBRE_HOJA = 'Emails Confirmation'; // <-- CAMBIA esto por el nombre real de tu pestaña
const NOMBRE_NEGOCIO = 'Peynado Law Firm';
const TELEFONO_NEGOCIO = '(404) 829-9989';
const DIRECCION_NEGOCIO = '1670 McKendree Church Road, Building #600, Lawrenceville, GA 30043';
const WEBSITE_NEGOCIO = 'https://peynadolawfirm.com';
const URL_LOGO = 'https://peynadolawfirm.com/images/logo/logo.svg';

const URL_YOUTUBE = 'https://www.youtube.com/@pamelapeynado';
const URL_FACEBOOK = 'https://www.facebook.com/pamelapeynado';
const URL_INSTAGRAM = 'https://www.instagram.com/atlantaimmigrationlawyer';
const URL_TIKTOK = 'https://www.tiktok.com/@atlimmigrationlawyer';
const URL_WHATSAPP = 'https://wa.me/14048299989';

// ======= AUTOMÁTICO: se dispara al escribir un correo en la columna M =======
function alEditarHoja(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== NOMBRE_HOJA) return;

  const filaInicio = e.range.getRow();
  const filaFin = e.range.getLastRow();
  const colInicio = e.range.getColumn();
  const colFin = e.range.getLastColumn();

  // Columna M = 13. Si la edición no toca esa columna, ignorar.
  if (colFin < 13 || colInicio > 13) return;

  for (let fila = Math.max(filaInicio, 2); fila <= filaFin; fila++) {
    procesarFila(sheet, fila);
  }
}

// Revisa una fila específica y manda la confirmación si corresponde
function procesarFila(sheet, filaEditada) {
  const fila = sheet.getRange(filaEditada, 1, 1, 14).getValues()[0]; // columnas A-N
  const cliente = fila[3];             // D
  const fechaCita = fila[6];           // G
  const horaCita = fila[7];            // H
  const tipoConsulta = fila[8];        // I
  const nombreAbogado = fila[11];      // L
  const email = fila[12];              // M
  const statusActual = String(fila[13] || '').trim(); // N

  if (statusActual === 'Sent') return; // ya se mandó, no repetir
  if (!email) return; // no hay correo, no se manda nada (el usuario marca "Not Sent" a mano)

  try {
    enviarConfirmacion(email, cliente, fechaCita, horaCita, tipoConsulta, nombreAbogado);
    sheet.getRange(filaEditada, 14).setValue('Sent'); // columna N
    Logger.log(`Confirmación enviada a: ${cliente} (${email}) - fila ${filaEditada}`);
  } catch (err) {
    Logger.log(`Error enviando a ${cliente} (fila ${filaEditada}): ${err}`);
  }
}

// ======= PRUEBA: lee la fila 2 y manda la confirmación (para probar) =======
function probarPrimeraCita() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const fila = sheet.getRange(2, 1, 1, 14).getValues()[0];

  const cliente = fila[3];
  const fechaCita = fila[6];
  const horaCita = fila[7];
  const tipoConsulta = fila[8];
  const nombreAbogado = fila[11];
  const email = fila[12];

  if (!email) {
    Logger.log('No hay email en la columna M de la fila 2.');
    return;
  }

  enviarConfirmacion(email, cliente, fechaCita, horaCita, tipoConsulta, nombreAbogado);
  Logger.log(`Confirmación de prueba enviada a: ${email}`);
}

// ======= ENVIAR A TODAS LAS FILAS PENDIENTES (por si agregas varias de golpe) =======
function enviarATodasLasPendientes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA);
  const data = sheet.getDataRange().getValues();
  let enviados = 0;

  for (let i = 1; i < data.length; i++) {
    const fila = i + 1;
    const cliente = data[i][3];
    const fechaCita = data[i][6];
    const horaCita = data[i][7];
    const tipoConsulta = data[i][8];
    const nombreAbogado = data[i][11];
    const email = data[i][12];
    const status = String(data[i][13] || '').trim();

    if (!email || status === 'Sent') continue;

    try {
      enviarConfirmacion(email, cliente, fechaCita, horaCita, tipoConsulta, nombreAbogado);
      sheet.getRange(fila, 14).setValue('Sent');
      enviados++;
      Logger.log(`Enviado a: ${cliente} (${email})`);
    } catch (err) {
      Logger.log(`Error con ${cliente}: ${err}`);
    }
  }

  Logger.log(`Listo. Confirmaciones enviadas en esta corrida: ${enviados}`);
}

// ======= FORMATEAR FECHA Y HORA PARA QUE SE VEAN BONITAS =======
function formatearFecha(fecha) {
  if (!fecha) return '';
  if (Object.prototype.toString.call(fecha) === '[object Date]') {
    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
  }
  return String(fecha);
}

function formatearHora(hora) {
  if (!hora) return '';
  if (Object.prototype.toString.call(hora) === '[object Date]') {
    return Utilities.formatDate(hora, Session.getScriptTimeZone(), 'h:mm a');
  }
  return String(hora);
}

// ======= ENVÍO: EMAIL DE CONFIRMACIÓN =======
function enviarConfirmacion(email, cliente, fechaCita, horaCita, tipoConsulta, nombreAbogado) {
  const fechaTexto = formatearFecha(fechaCita);
  const horaTexto = formatearHora(horaCita);

  const asunto = `Confirmación de su cita - ${NOMBRE_NEGOCIO}`;

  const cuerpoTexto =
    `Estimado(a) ${cliente || ''},\n\n` +
    `Le confirmamos su cita con ${NOMBRE_NEGOCIO}:\n\n` +
    `Fecha: ${fechaTexto}\n` +
    `Hora: ${horaTexto}\n` +
    `Tipo de consulta: ${tipoConsulta || ''}\n` +
    (nombreAbogado ? `Atendido por: ${nombreAbogado}\n` : '') +
    `\nSi necesita reprogramar o tiene alguna pregunta, no dude en contactarnos al ${TELEFONO_NEGOCIO}.\n\n` +
    `Gracias por confiar en nosotros.\n${NOMBRE_NEGOCIO}`;

  const cuerpoHtml = `
    <div style="font-family: 'Georgia', 'Times New Roman', serif; max-width: 520px; margin: 0 auto; background-color: #ffffff;">

      <!-- Encabezado con logo -->
      <div style="background-color: #0b1f3a; text-align: center; padding: 32px 20px;">
        <img src="${URL_LOGO}" alt="${NOMBRE_NEGOCIO}" style="max-width: 220px; height: auto;">
      </div>

      <!-- Cuerpo -->
      <div style="padding: 32px 28px; color: #222;">
        <h2 style="color: #0b1f3a; font-size: 20px; font-weight: normal; margin-top: 0;">
          Estimado(a) ${cliente || ''},
        </h2>
        <p style="font-size: 15px; line-height: 1.7; color: #333;">
          Le confirmamos su cita con <strong>${NOMBRE_NEGOCIO}</strong>. A continuación los detalles:
        </p>

        <!-- Tarjeta con detalles de la cita -->
        <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%; background-color: #f9f7f0; border-left: 4px solid #c9a227; margin: 24px 0;">
          <tr>
            <td style="padding: 20px 24px;">
              <table role="presentation" cellpadding="0" cellspacing="0" style="width: 100%;">
                <tr>
                  <td style="padding: 6px 0; font-size: 14px; color: #888; width: 40%;">Fecha:</td>
                  <td style="padding: 6px 0; font-size: 15px; color: #0b1f3a; font-weight: bold;">${fechaTexto}</td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; font-size: 14px; color: #888;">Hora:</td>
                  <td style="padding: 6px 0; font-size: 15px; color: #0b1f3a; font-weight: bold;">${horaTexto}</td>
                </tr>
                <tr>
                  <td style="padding: 6px 0; font-size: 14px; color: #888;">Tipo de consulta:</td>
                  <td style="padding: 6px 0; font-size: 15px; color: #0b1f3a;">${tipoConsulta || ''}</td>
                </tr>
                ${nombreAbogado ? `
                <tr>
                  <td style="padding: 6px 0; font-size: 14px; color: #888;">Atendido por:</td>
                  <td style="padding: 6px 0; font-size: 15px; color: #0b1f3a;">${nombreAbogado}</td>
                </tr>` : ''}
              </table>
            </td>
          </tr>
        </table>

        <p style="font-size: 14px; line-height: 1.7; color: #333;">
          Si necesita reprogramar o tiene alguna pregunta, no dude en contactarnos al
          <strong>${TELEFONO_NEGOCIO}</strong>.
        </p>

        <!-- Redes sociales -->
        <div style="text-align: center; margin: 28px 0 8px 0;">
          <p style="font-size: 12px; color: #888; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 16px;">
            Síganos en nuestras redes
          </p>
          <table role="presentation" align="center" cellpadding="0" cellspacing="0" style="margin: 0 auto;">
            <tr>
              <td style="padding: 0 5px;">
                <a href="${URL_INSTAGRAM}">
                  <table role="presentation" cellpadding="0" cellspacing="0" style="width: 42px; height: 42px; background-color: #0b1f3a; border-radius: 50%;">
                    <tr><td align="center" valign="middle">
                      <img src="https://img.icons8.com/ios-filled/24/ffffff/instagram-new.png" width="20" height="20" alt="Instagram" style="display: block;">
                    </td></tr>
                  </table>
                </a>
              </td>
              <td style="padding: 0 5px;">
                <a href="${URL_FACEBOOK}">
                  <table role="presentation" cellpadding="0" cellspacing="0" style="width: 42px; height: 42px; background-color: #0b1f3a; border-radius: 50%;">
                    <tr><td align="center" valign="middle">
                      <img src="https://img.icons8.com/ios-filled/24/ffffff/facebook-new.png" width="20" height="20" alt="Facebook" style="display: block;">
                    </td></tr>
                  </table>
                </a>
              </td>
              <td style="padding: 0 5px;">
                <a href="${URL_YOUTUBE}">
                  <table role="presentation" cellpadding="0" cellspacing="0" style="width: 42px; height: 42px; background-color: #0b1f3a; border-radius: 50%;">
                    <tr><td align="center" valign="middle">
                      <img src="https://img.icons8.com/ios-filled/24/ffffff/youtube-play.png" width="20" height="20" alt="YouTube" style="display: block;">
                    </td></tr>
                  </table>
                </a>
              </td>
              <td style="padding: 0 5px;">
                <a href="${URL_TIKTOK}">
                  <table role="presentation" cellpadding="0" cellspacing="0" style="width: 42px; height: 42px; background-color: #0b1f3a; border-radius: 50%;">
                    <tr><td align="center" valign="middle">
                      <img src="https://img.icons8.com/ios-filled/24/ffffff/tiktok.png" width="20" height="20" alt="TikTok" style="display: block;">
                    </td></tr>
                  </table>
                </a>
              </td>
              <td style="padding: 0 5px;">
                <a href="${URL_WHATSAPP}">
                  <table role="presentation" cellpadding="0" cellspacing="0" style="width: 42px; height: 42px; background-color: #0b1f3a; border-radius: 50%;">
                    <tr><td align="center" valign="middle">
                      <img src="https://img.icons8.com/ios-filled/24/ffffff/whatsapp.png" width="20" height="20" alt="WhatsApp" style="display: block;">
                    </td></tr>
                  </table>
                </a>
              </td>
            </tr>
          </table>
        </div>

        <p style="font-size: 14px; color: #555; margin-top: 32px; line-height: 1.6;">
          Gracias por confiar en nosotros.<br>
          <strong style="color: #0b1f3a;">${NOMBRE_NEGOCIO}</strong>
        </p>
      </div>

      <!-- Pie con datos de contacto -->
      <div style="background-color: #f4f4f4; padding: 20px 28px; text-align: center;
                  border-top: 3px solid #c9a227;">
        <p style="font-size: 12px; color: #666; margin: 4px 0;">
          ${DIRECCION_NEGOCIO}
        </p>
        <p style="font-size: 12px; color: #666; margin: 4px 0;">
          Tel: ${TELEFONO_NEGOCIO} &nbsp;|&nbsp;
          <a href="${WEBSITE_NEGOCIO}" style="color: #0b1f3a;">${WEBSITE_NEGOCIO}</a>
        </p>
      </div>

    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: asunto,
    body: cuerpoTexto,
    htmlBody: cuerpoHtml
  });
}