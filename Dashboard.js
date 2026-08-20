// Archivo: Dashboard.gs / Dashboard.js

/**
 * Sirve la aplicación web HTML (Index.html).
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard Soporte TI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Obtiene y estructura todos los datos de tickets de la hoja de cálculo,
 * incluyendo evidencias de Google Drive, comentarios e historial, y niveles de caso.
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!sheet) return { error: 'Hoja "Respuestas de formulario 1" no encontrada' };

    const data = sheet.getDataRange().getValues();
    if (!data || data.length === 0) return { tickets: [] };

    const headers = data[0] || [];
    const rows = data.slice(1);

    // Función segura para buscar índices en encabezados evitando errores con celdas vacías o nulas
    const getIdx = (name) => headers.findIndex(h => String(h || '').toLowerCase().includes(String(name).toLowerCase()));

    // Detección robusta de Cambio de equipo (Columna O / índice 14 aprox)
    let idxCambioEquipo = getIdx('Cambio de equipo');
    if (idxCambioEquipo === -1) idxCambioEquipo = getIdx('Cambio de máquina');
    if (idxCambioEquipo === -1) idxCambioEquipo = getIdx('Cambio de maquina');
    if (idxCambioEquipo === -1 && headers.length > 14) idxCambioEquipo = 14;

    // Detección de Nivel de caso (Columna P / índice 15 aprox)
    let idxNivelCaso = getIdx('Nivel de caso');
    if (idxNivelCaso === -1) idxNivelCaso = getIdx('Nivel');
    if (idxNivelCaso === -1 && headers.length > 15) idxNivelCaso = 15;

    // Detección de Evidencia / Captura / Drive (Columna H / índice 7 aprox)
    let idxEvidencia = getIdx('Evidencia');
    if (idxEvidencia === -1) idxEvidencia = getIdx('Adjunto');
    if (idxEvidencia === -1) idxEvidencia = getIdx('Archivo');
    if (idxEvidencia === -1) idxEvidencia = getIdx('Foto');
    if (idxEvidencia === -1 && headers.length > 7) idxEvidencia = 7;

    // Detección de Comentarios / Historial
    let idxComentarios = getIdx('Comentarios');
    if (idxComentarios === -1) idxComentarios = getIdx('Historial');
    if (idxComentarios === -1) idxComentarios = getIdx('Notas');

    const idx = {
      fecha: getIdx('Marca temporal'), 
      categoria: getIdx('Categoria'), 
      estado: getIdx('Estado'),       
      id: getIdx('Codigo'),         
      detalle: getIdx('Detalle'),
      tipo: getIdx('Tipo'),
      nombre: getIdx('Nombre'),     
      email: getIdx('Puntuación') === -1 ? getIdx('correo') : getIdx('email'),
      solucion: getIdx('Solución'), 
      cambioEquipo: idxCambioEquipo,
      equipoAveriado: getIdx('Equipo averiado'),
      nivelCaso: idxNivelCaso,
      evidencia: idxEvidencia,
      comentarios: idxComentarios
    };

    let tickets = rows.map(row => {
      let fechaDoc = row[idx.fecha];
      let fechaTimestamp = fechaDoc instanceof Date ? fechaDoc.getTime() : new Date().getTime();
      let rawEvidencia = idx.evidencia !== -1 ? String(row[idx.evidencia] || '').trim() : '';
      let rawComentarios = idx.comentarios !== -1 ? String(row[idx.comentarios] || '').trim() : '';
      let rawNivel = idx.nivelCaso !== -1 ? String(row[idx.nivelCaso] || 'Bajo').trim() : 'Bajo';
      if (!rawNivel) rawNivel = 'Bajo';

      return {
        fecha: fechaTimestamp,
        id: String(row[idx.id] || '-').trim(),
        cat: String(row[idx.categoria] || 'General').trim(),
        tipo: String(row[idx.tipo] || 'N/A').trim(),
        estado: String(row[idx.estado] || 'Pendiente').trim(),
        descCorta: String(row[idx.detalle] || '').substring(0, 45) + '...', 
        descCompleta: String(row[idx.detalle] || ''),
        nombre: String(row[idx.nombre] || 'Desconocido').trim(),
        solucion: String(row[idx.solucion] || '').trim(),
        cambioEquipo: idx.cambioEquipo !== -1 && (
          row[idx.cambioEquipo] === true || 
          String(row[idx.cambioEquipo] || '').trim().toLowerCase() === 'sí' || 
          String(row[idx.cambioEquipo] || '').trim().toLowerCase() === 'si'
        ),
        equipoAveriado: String(idx.equipoAveriado !== -1 ? row[idx.equipoAveriado] : '').trim(),
        nivelCaso: rawNivel,
        evidencias: parsearEvidenciasDrive(rawEvidencia),
        comentarios: parsearComentarios(rawComentarios)
      };
    });

    // Filtra filas vacías o sin código de ticket
    tickets = tickets.filter(t => t.id !== '-' && t.id !== '');
    return { tickets: tickets };
  } catch (e) {
    Logger.log('Error en getDashboardData: ' + e.message);
    return { error: e.message };
  }
}

/**
 * Convierte URLs de Google Drive en objetos estructurados con IDs, enlaces de visualización y miniaturas.
 */
function parsearEvidenciasDrive(rawEvidencia) {
  if (!rawEvidencia) return [];

  const urls = rawEvidencia.split(/[\n,]+/).map(u => u.trim()).filter(u => u.length > 0);
  
  return urls.map(url => {
    let fileId = '';
    
    // Extrae ID de formatos: ?id=XXXX o /d/XXXX/ o /file/d/XXXX
    const matchId = url.match(/id=([a-zA-Z0-9_-]+)/);
    const matchD = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    
    if (matchId && matchId[1]) {
      fileId = matchId[1];
    } else if (matchD && matchD[1]) {
      fileId = matchD[1];
    }

    return {
      originalUrl: url,
      fileId: fileId,
      // URL para visor embebido nativo de Google Drive
      previewUrl: fileId ? `https://drive.google.com/file/d/${fileId}/preview` : url,
      // URL para abrir en pestaña nueva de Google Drive
      viewUrl: fileId ? `https://drive.google.com/file/d/${fileId}/view?usp=sharing` : url,
      // Miniatura directa
      thumbnailUrl: fileId ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w800` : url
    };
  });
}

/**
 * Parsea el contenido de la celda de comentarios (soporta JSON estructurado o texto libre acumulado).
 */
function parsearComentarios(rawComentarios) {
  if (!rawComentarios) return [];

  try {
    const parsed = JSON.parse(rawComentarios);
    if (Array.isArray(parsed)) return parsed;
  } catch (e) {
    const lineas = rawComentarios.split('\n').filter(l => l.trim().length > 0);
    return lineas.map(linea => ({
      fecha: '',
      autor: 'Soporte TI',
      texto: linea
    }));
  }

  return [];
}

/**
 * Actualiza el estado, solución, cambio de equipo, nivel de caso y agrega nuevos comentarios.
 */
function updateTicketStatus(ticketId, nuevoEstado, textoSolucion, cambioEquipoVal, equipoAveriadoVal, nivelCasoVal, nuevoComentario, autorComentario) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Respuestas de formulario 1');
    if (!sheet) return { error: 'Hoja no encontrada' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    
    const getIdxSafe = (name) => headers.findIndex(h => String(h || '').toLowerCase().includes(String(name).toLowerCase()));

    const idCol = getIdxSafe('codigo');
    const estadoCol = getIdxSafe('estado');
    const solucionCol = getIdxSafe('solución'); 
    
    let cambioEquipoCol = getIdxSafe('cambio de equipo');
    if (cambioEquipoCol === -1) cambioEquipoCol = getIdxSafe('cambio de máquina');
    if (cambioEquipoCol === -1) cambioEquipoCol = getIdxSafe('cambio de maquina');
    if (cambioEquipoCol === -1 && headers.length > 14) cambioEquipoCol = 14;

    const equipoAveriadoCol = getIdxSafe('equipo averiado');

    let nivelCasoCol = getIdxSafe('nivel de caso');
    if (nivelCasoCol === -1) nivelCasoCol = getIdxSafe('nivel');
    if (nivelCasoCol === -1 && headers.length > 15) nivelCasoCol = 15;

    // Buscar o crear la columna de comentarios si no existe
    let comentariosCol = getIdxSafe('comentarios');
    if (comentariosCol === -1) {
      comentariosCol = headers.length;
      sheet.getRange(1, comentariosCol + 1).setValue('Comentarios');
    }

    if (idCol === -1 || estadoCol === -1) return { error: 'No se encontraron las columnas de Código o Estado en el Sheet' };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(ticketId).trim()) {
        const rowNum = i + 1;

        // 1. Guardar Estado
        sheet.getRange(rowNum, estadoCol + 1).setValue(nuevoEstado);
        
        // 2. Guardar Solución
        if (solucionCol !== -1 && textoSolucion !== undefined) {
           sheet.getRange(rowNum, solucionCol + 1).setValue(textoSolucion);
        }
        
        // 3. Guardar Cambio de Equipo
        if (cambioEquipoCol !== -1 && cambioEquipoVal !== undefined) {
           sheet.getRange(rowNum, cambioEquipoCol + 1).setValue(cambioEquipoVal ? 'Sí' : 'No');
        }
        
        // 4. Guardar Equipo Averiado
        if (equipoAveriadoCol !== -1 && equipoAveriadoVal !== undefined) {
           sheet.getRange(rowNum, equipoAveriadoCol + 1).setValue(equipoAveriadoVal);
        }

        // 5. Guardar Nivel de Caso
        if (nivelCasoCol !== -1 && nivelCasoVal !== undefined) {
           sheet.getRange(rowNum, nivelCasoCol + 1).setValue(nivelCasoVal);
        }

        // 6. Guardar nuevo comentario en la bitácora si se ingresó
        if (nuevoComentario && String(nuevoComentario).trim().length > 0) {
          const rawActual = String(data[i][comentariosCol] || '').trim();
          let listaComentarios = parsearComentarios(rawActual);

          const fechaActualStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Lima', 'dd/MM/yyyy HH:mm');
          listaComentarios.push({
            fecha: fechaActualStr,
            autor: autorComentario || 'Soporte TI',
            texto: String(nuevoComentario).trim()
          });

          sheet.getRange(rowNum, comentariosCol + 1).setValue(JSON.stringify(listaComentarios));
        }
        
        return { success: true };
      }
    }
    return { error: 'Ticket no encontrado en la base de datos' };
  } catch (e) {
    Logger.log('Error en updateTicketStatus: ' + e.message);
    return { error: e.message };
  }
}