function notificarOnFormSubmit(e) {

    var marca = e.values[0];

    var nombre = e.values[1];

    var email = e.values[2];

    var cargo = e.values[3];

    var area = e.values[4];

    var celular = e.values[5];

    var detalle = e.values[6];

    var evidencia = e.values[7];



    // Llama la función que obtiene el número de registro

    var registro = agregarFolioRegistro();
    
var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');


    var asunto = `Registro número ${registro}`;



    // email_plano es para cuando el correo del receptor no soporta HTML

    var email_plano = `

    Hola ${nombre}, tu solicitud de asistencia de Soporte TI fue registrada.\n\n

    Tu número de registro es: ${registro}\n\n

    Gracias por confiar en nosotros.\n\n

    Atentamente,\n

    Cristian Apaza

  `;



    var email_html = `

<p>&iexcl;Hola, <strong>${nombre}</strong>!</p>

<p>Se ha generado tu ticket de solicitud de asistencia de Soporte TI.</p>

<p>Tu n&uacute;mero de ticket es:&nbsp;<span style="color: #ff0000;"><strong>${registro}</strong></span></p>

<p><span style="color: #000000;">Nos comunicaremos contigo a la brevedad.</span></p>

<p>&iexcl;Mil gracias!</p>

<p>&nbsp;</p>

<table border="0" width="420" cellspacing="3" cellpadding="0">

<tbody>



</tbody>

</table>

  `;



    var optsAvanzadas = {name: "Soporte TI - Cristian Apaza", htmlBody: email_html};

    MailApp.sendEmail(email, asunto, email_plano, optsAvanzadas);

}



function agregarFolioRegistro() {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');
    var renglon = hoja.getLastRow();

    // 1. Buscamos automáticamente en qué columna está el encabezado "Codigo"
    var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    var columnaCodigo = headers.findIndex(h => String(h).toLowerCase().includes('codigo')) + 1;

    // Si por alguna razón no encuentra la palabra "Codigo", usa la columna 10 por defecto
    if (columnaCodigo === 0) columnaCodigo = 10; 

    // 2. Traemos todos los códigos que ya existen en esa columna
    var codigos = hoja.getRange(2, columnaCodigo, renglon > 1 ? renglon - 1 : 1, 1).getValues();
    var maximoActual = 0;

    // 3. Revisamos uno por uno para encontrar el número más alto
    for (var i = 0; i < codigos.length; i++) {
        var valor = String(codigos[i][0]).trim();
        if (valor.startsWith('C-')) {
            var numero = parseInt(valor.replace('C-', ''), 10);
            if (!isNaN(numero) && numero > maximoActual) {
                maximoActual = numero;
            }
        }
    }

    // 4. Creamos el nuevo código sumando 1 al más alto que encontró
    var nuevoNumero = maximoActual + 1;
    var folioRegistro = `C-${Utilities.formatString("%04d", nuevoNumero)}`;

    // 5. Lo guardamos en la celda correspondiente
    hoja.getRange(renglon, columnaCodigo).setValue(folioRegistro);

    return folioRegistro;
}
