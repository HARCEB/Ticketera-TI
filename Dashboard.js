// Archivo: Dashboard.gs

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Soporte TI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!sheet) return { error: 'Hoja no encontrada' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const getIdx = (name) => headers.findIndex(h => h.toString().toLowerCase().includes(name.toLowerCase()));

    // NUEVOS CAMPOS: Agregamos el nombre y la solución
    const idx = {
      fecha: getIdx('Marca temporal'), 
      categoria: getIdx('Categoria'), 
      estado: getIdx('Estado'),       
      id: getIdx('Codigo'),         
      detalle: getIdx('Detalle'),
      tipo: getIdx('Tipo'),
      nombre: getIdx('Nombre'),     // <-- Busca columna de Nombre
      solucion: getIdx('Solución'),  // <-- Busca columna de Solución (asegúrate de que tenga tilde si tu sheet la tiene)
      cambioEquipo: getIdx('Cambio de equipo'), // <-- Busca columna de Cambio de equipo
      equipoAveriado: getIdx('Equipo averiado') // <-- Busca columna de Equipo averiado
    };

    let tickets = rows.map(row => {
      let fechaDoc = row[idx.fecha];
      return {
        fecha: fechaDoc instanceof Date ? fechaDoc.getTime() : new Date().getTime(),
        id: row[idx.id] || '-',
        cat: String(row[idx.categoria] || 'General').trim(),
        tipo: String(row[idx.tipo] || 'N/A').trim(),
        estado: String(row[idx.estado] || 'Abierto').trim(),
        descCorta: String(row[idx.detalle] || '').substring(0, 45) + '...', 
        descCompleta: String(row[idx.detalle] || ''),
        // Nuevos datos que enviamos al HTML:
        nombre: String(row[idx.nombre] || 'Desconocido').trim(),
        solucion: String(row[idx.solucion] || '').trim(),
        cambioEquipo: row[idx.cambioEquipo] === true || String(row[idx.cambioEquipo] || '').trim().toLowerCase() === 'sí' || String(row[idx.cambioEquipo] || '').trim().toLowerCase() === 'si',
        equipoAveriado: String(row[idx.equipoAveriado] || '').trim()
      };
    });

    tickets = tickets.filter(t => t.id !== '-' && t.id !== '');
    return { tickets: tickets };
  } catch (e) {
    return { error: e.message };
  }
}

// ACTUALIZADO: Ahora recibe también el texto de la solución y cambio de equipo
function updateTicketStatus(ticketId, nuevoEstado, textoSolucion, cambioEquipoVal, equipoAveriadoVal) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.findIndex(h => h.toString().toLowerCase().includes('codigo'));
    const estadoCol = headers.findIndex(h => h.toString().toLowerCase().includes('estado'));
    const solucionCol = headers.findIndex(h => h.toString().toLowerCase().includes('solución')); // Asegúrate de la tilde
    const cambioEquipoCol = headers.findIndex(h => h.toString().toLowerCase().includes('cambio de equipo'));
    const equipoAveriadoCol = headers.findIndex(h => h.toString().toLowerCase().includes('equipo averiado'));

    if (idCol === -1 || estadoCol === -1) return { error: 'No se encontraron columnas necesarias' };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] == ticketId) {
        // Guarda el estado
        sheet.getRange(i + 1, estadoCol + 1).setValue(nuevoEstado);
        
        // Si hay una columna de solución y mandamos texto, lo guarda
        if(solucionCol !== -1 && textoSolucion !== undefined) {
           sheet.getRange(i + 1, solucionCol + 1).setValue(textoSolucion);
        }
        
        // Guarda el cambio de equipo
        if(cambioEquipoCol !== -1 && cambioEquipoVal !== undefined) {
           sheet.getRange(i + 1, cambioEquipoCol + 1).setValue(cambioEquipoVal ? 'Sí' : 'No');
        }
        
        // Guarda el equipo averiado
        if(equipoAveriadoCol !== -1 && equipoAveriadoVal !== undefined) {
           sheet.getRange(i + 1, equipoAveriadoCol + 1).setValue(equipoAveriadoVal);
        }
        
        return { success: true };
      }
    }
    return { error: 'Ticket no encontrado en la base de datos' };
  } catch (e) {
    return { error: e.message };
  }
}