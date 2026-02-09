// ==========================================
// GOOGLE APPS SCRIPT - RECEPCIÓN DE TICKETS
// ==========================================
// Este script debe ser pegado en Google Apps Script
// (Extensiones > Apps Script en tu Google Sheet)

function doPost(e) {
  try {
    // Parsear el JSON recibido
    const data = JSON.parse(e.postData.contents);
    
    // Conectar con la hoja activa
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Tickets');
    
    // Si no existe la hoja "Tickets", crearla
    if (!sheet) {
      sheet = ss.insertSheet('Tickets');
      
      // Crear encabezados
      const headers = [
        'Fecha Solicitud',
        'Nombre',
        'Prioridad',
        'Pedido',
        'Comentario',
        'Tipo Servicio',
        'Urgencia Declarada',
        'Proceso Actual',
        'Resultado Esperado',
        'Impacto',
        'Horas Ahorradas',
        'Fecha Límite',
        'Estado',
        'Ticket Number'
      ];
      
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Formatear encabezados
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#fbbf24')
        .setFontColor('#000000')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      // Congelar primera fila
      sheet.setFrozenRows(1);
    }
    
    // Preparar la fila de datos en el orden correcto
    const rowData = [
      data.fecha_solicitud || '',
      data.nombre || '',
      data.prioridad || '',
      data.pedido || '',
      data.comentario || '',
      data.tipo_servicio || '',
      data.urgencia_declarada || '',
      data.proceso_actual || '',
      data.resultado_esperado || '',
      data.impacto || '',
      data.horas_ahorradas || '',
      data.fecha_limite || '',
      data.estado || '',
      data.ticket_number || ''
    ];
    
    // Insertar en la primera fila disponible (después del header)
    const lastRow = sheet.getLastRow();
    sheet.insertRowAfter(1); // Insertar después del header
    const targetRow = 2; // Siempre insertar en la fila 2 (los más nuevos arriba)
    
    // Escribir los datos
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Aplicar formato según prioridad
    const prioridadCell = sheet.getRange(targetRow, 3); // Columna "Prioridad"
    const prioridad = data.prioridad;
    
    if (prioridad === 'CRÍTICA') {
      prioridadCell.setBackground('#ef4444').setFontColor('#ffffff');
    } else if (prioridad === 'ALTA') {
      prioridadCell.setBackground('#f97316').setFontColor('#ffffff');
    } else if (prioridad === 'MEDIA') {
      prioridadCell.setBackground('#84cc16').setFontColor('#000000');
    } else if (prioridad === 'BAJA') {
      prioridadCell.setBackground('#3b82f6').setFontColor('#ffffff');
    } else {
      prioridadCell.setBackground('#a1a1aa').setFontColor('#000000');
    }
    
    // Aplicar bordes
    sheet.getRange(targetRow, 1, 1, rowData.length)
      .setBorder(true, true, true, true, true, true);
    
    // Auto-ajustar columnas (solo la primera vez)
    if (lastRow === 1) {
      sheet.autoResizeColumns(1, rowData.length);
    }
    
    // Respuesta exitosa
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'success',
        message: 'Ticket guardado correctamente',
        ticket: data.ticket_number
      })
    ).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Respuesta de error
    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'error',
        message: error.toString()
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Función de prueba (opcional)
function testDoPost() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        fecha_solicitud: '9 de febrero de 2026',
        nombre: 'Maria Laura Colque',
        prioridad: 'ALTA',
        pedido: 'Automatizar proceso de facturación',
        comentario: 'Necesitamos integrar el sistema actual',
        tipo_servicio: 'Automatización',
        urgencia_declarada: '4',
        proceso_actual: 'Actualmente se hace manual',
        resultado_esperado: 'Sistema automatizado',
        impacto: '16-50 personas',
        horas_ahorradas: '10',
        fecha_limite: '15 de marzo de 2026',
        estado: 'Backlog',
        ticket_number: 'TKT-20260209-1234'
      })
    }
  };
  
  const result = doPost(testData);
  Logger.log(result.getContent());
}

// ==========================================
// INSTRUCCIONES DE IMPLEMENTACIÓN
// ==========================================
/*
1. Abre tu Google Sheet
2. Ve a Extensiones > Apps Script
3. Pega este código completo
4. Guarda el proyecto (Ctrl+S)
5. Haz clic en "Implementar" > "Nueva implementación"
6. Tipo: "Aplicación web"
7. Ejecutar como: "Yo"
8. Quién tiene acceso: "Cualquier persona"
9. Copia la URL que te da
10. Pega esa URL en la variable GOOGLE_SHEETS_URL del HTML

IMPORTANTE: 
- Después de implementar, autoriza el script
- Copia la URL de la Web App
- Reemplaza 'TU_URL_DE_GOOGLE_APPS_SCRIPT_AQUI' en el HTML
*/