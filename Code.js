// Archivo: Code.gs / Code.js

const ADMIN_EMAIL = 'capaza@buk.pe';

/**
 * Disparador automático que se ejecuta al enviar el formulario de Google.
 * Asigna número de folio/código correlativo y envía correos de confirmación y alerta.
 */
function notificarOnFormSubmit(e) {
  try {
    var valores = e ? e.values : [];
    var marca = valores[0] || new Date();
    var nombre = valores[1] || 'Usuario';
    var email = valores[2] || '';
    var cargo = valores[3] || 'N/A';
    var area = valores[4] || 'General';
    var celular = valores[5] || 'N/A';
    var detalle = valores[6] || 'Sin descripción';
    var evidencia = valores[7] || '';

    // 1. Asigna el código único correlativo (C-XXXX)
    var registro = agregarFolioRegistro();

    // 2. Correo de confirmación al Solicitante
    if (email && email.includes('@')) {
      enviarConfirmacionSolicitante(email, nombre, registro);
    }

    // 3. Correo de alerta al Administrador de TI (capaza@buk.pe)
    enviarAlertaAdmin({
      registro: registro,
      nombre: nombre,
      email: email,
      cargo: cargo,
      area: area,
      celular: celular,
      detalle: detalle,
      evidencia: evidencia,
      marca: marca
    });

  } catch (error) {
    Logger.log('Error en notificarOnFormSubmit: ' + error.message);
  }
}

/**
 * Envía correo HTML de confirmación al usuario que creó el ticket.
 */
function enviarConfirmacionSolicitante(email, nombre, registro) {
  var asunto = `Soporte TI: Ticket creado con éxito [#${registro}]`;

  var emailPlano = `Hola ${nombre},\n\nTu solicitud de asistencia de Soporte TI fue registrada con el número: ${registro}.\nNos comunicaremos contigo a la brevedad.\n\nAtentamente,\nSoporte TI - Cristian Apaza`;

  var emailHtml = `
    <div style="font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: auto; padding: 25px; border: 1px solid #e2e8f0; border-radius: 12px; background-color: #ffffff; color: #1e293b;">
      <div style="text-align: center; margin-bottom: 20px;">
        <span style="background-color: #0f172a; color: #ffffff; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 700; letter-spacing: 1px;">SOPORTE TI</span>
      </div>
      <h2 style="color: #0f172a; margin-top: 0; font-size: 20px; text-align: center;">¡Hola, <strong>${nombre}</strong>!</h2>
      <p style="font-size: 14px; line-height: 1.6; color: #475569; text-align: center;">Tu solicitud de asistencia técnica ha sido recibida y asignada a la cola de atención.</p>
      
      <div style="background-color: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 10px; padding: 18px; text-align: center; margin: 25px 0;">
        <p style="margin: 0; font-size: 12px; text-transform: uppercase; color: #64748b; font-weight: 600;">Número de Ticket</p>
        <p style="margin: 6px 0 0 0; font-size: 28px; font-weight: 800; color: #2563eb; letter-spacing: 1px;">${registro}</p>
      </div>

      <p style="font-size: 14px; color: #475569; text-align: center;">Nos pondremos en contacto contigo a través de este correo o teléfono registrado.</p>
      
      <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 25px 0;" />
      <p style="font-size: 12px; color: #94a3b8; text-align: center; margin: 0;">Plataforma de Soporte TI &bull; Buk Perú</p>
    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: asunto,
    body: emailPlano,
    htmlBody: emailHtml,
    name: "Soporte TI - Buk"
  });
}

/**
 * Envía correo HTML detallado de alerta al administrador (capaza@buk.pe).
 */
function enviarAlertaAdmin(datos) {
  var asunto = `🚨 [NUEVO TICKET ${datos.registro}] ${datos.area} - ${datos.nombre}`;

  var linkEvidencia = datos.evidencia 
    ? `<a href="${datos.evidencia}" target="_blank" style="color: #2563eb; font-weight: 600; text-decoration: underline;">Ver archivo adjunto en Google Drive</a>`
    : `<span style="color: #94a3b8; font-style: italic;">Sin archivo adjunto</span>`;

  var emailHtml = `
    <div style="font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 650px; margin: auto; padding: 25px; border: 1px solid #e2e8f0; border-radius: 12px; background-color: #ffffff; color: #1e293b;">
      <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #f1f5f9; padding-bottom: 15px; margin-bottom: 20px;">
        <h2 style="margin: 0; color: #0f172a; font-size: 20px;">🚨 Nuevo Ticket Ingresado</h2>
        <span style="background-color: #dbeafe; color: #1e40af; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 14px;">${datos.registro}</span>
      </div>

      <table style="width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 20px;">
        <tr style="border-bottom: 1px solid #f1f5f9;">
          <td style="padding: 10px 0; color: #64748b; font-weight: 600; width: 140px;">Solicitante:</td>
          <td style="padding: 10px 0; color: #0f172a; font-weight: 700;">${datos.nombre}</td>
        </tr>
        <tr style="border-bottom: 1px solid #f1f5f9;">
          <td style="padding: 10px 0; color: #64748b; font-weight: 600;">Área / Cargo:</td>
          <td style="padding: 10px 0; color: #334155;">${datos.area} &bull; ${datos.cargo}</td>
        </tr>
        <tr style="border-bottom: 1px solid #f1f5f9;">
          <td style="padding: 10px 0; color: #64748b; font-weight: 600;">Contacto:</td>
          <td style="padding: 10px 0; color: #334155;">
            <a href="mailto:${datos.email}" style="color: #2563eb; text-decoration: none;">${datos.email}</a> | Tel: ${datos.celular}
          </td>
        </tr>
        <tr style="border-bottom: 1px solid #f1f5f9;">
          <td style="padding: 10px 0; color: #64748b; font-weight: 600;">Evidencia / Foto:</td>
          <td style="padding: 10px 0;">${linkEvidencia}</td>
        </tr>
      </table>

      <div style="background-color: #f8fafc; border-left: 4px solid #3b82f6; padding: 14px; border-radius: 6px; margin-bottom: 20px;">
        <p style="margin: 0 0 6px 0; font-size: 12px; font-weight: 700; color: #475569; text-transform: uppercase;">Descripción del Requerimiento:</p>
        <p style="margin: 0; font-size: 14px; color: #1e293b; line-height: 1.5; white-space: pre-wrap;">${datos.detalle}</p>
      </div>

      <div style="text-align: center; margin-top: 25px;">
        <span style="font-size: 12px; color: #94a3b8;">Notificación generada automáticamente por la Plataforma de Soporte TI.</span>
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: asunto,
    body: `Nuevo ticket ${datos.registro} de ${datos.nombre} (${datos.area}):\n\n${datos.detalle}\n\nContacto: ${datos.email} | ${datos.celular}`,
    htmlBody: emailHtml,
    name: "Ticketera Bot TI"
  });
}

/**
 * Busca y asigna el correlativo C-XXXX más alto en la columna "Codigo".
 */
function agregarFolioRegistro() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');
  if (!hoja) return 'C-0001';

  var renglon = hoja.getLastRow();
  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var columnaCodigo = headers.findIndex(h => String(h).toLowerCase().includes('codigo')) + 1;

  if (columnaCodigo === 0) columnaCodigo = 10; 

  var codigos = hoja.getRange(2, columnaCodigo, renglon > 1 ? renglon - 1 : 1, 1).getValues();
  var maximoActual = 0;

  for (var i = 0; i < codigos.length; i++) {
    var valor = String(codigos[i][0]).trim();
    if (valor.startsWith('C-')) {
      var numero = parseInt(valor.replace('C-', ''), 10);
      if (!isNaN(numero) && numero > maximoActual) {
        maximoActual = numero;
      }
    }
  }

  var nuevoNumero = maximoActual + 1;
  var folioRegistro = `C-${Utilities.formatString("%04d", nuevoNumero)}`;

  hoja.getRange(renglon, columnaCodigo).setValue(folioRegistro);
  return folioRegistro;
}
