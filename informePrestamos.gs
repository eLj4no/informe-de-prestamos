/**
 * Script vinculado a BD_PRESTAMOS para generar informes y notificaciones.
 */

// 1. MENÚ PERSONALIZADO
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestión Sindicato')
      .addItem('📥 Generar Informe Semanal (Manual)', 'generarInformeConMonto')
      .addSeparator()
      .addItem('⚙️ Configurar Automatización (Lunes 16:30)', 'configurarDisparadorAutomatico')
      .addToUi();
}

// 2. FUNCIÓN PRINCIPAL DE GENERACIÓN DE INFORME
function generarInformeConMonto() {
  try {
    const ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
    const sheetOrigen = ssOrigen.getSheetByName("BD_PRESTAMOS");
    const data = sheetOrigen.getDataRange().getValues();
    
    // Índices de columna en BD_PRESTAMOS (A=0)
    const COL_ESTADO = 9; 
    
    // Filtrar solo los "Solicitado"
    const pendientes = data.filter((row, i) => i > 0 && String(row[COL_ESTADO]) === "Solicitado");
    
    // Si es ejecución manual, avisamos si no hay datos. Si es automática (Trigger), no hacemos nada.
    if (pendientes.length === 0) {
      if (Session.getActiveUser().getEmail()) { 
        SpreadsheetApp.getUi().alert('Estado del Reporte', 'No hay solicitudes pendientes con estado "Solicitado" para procesar.', SpreadsheetApp.getUi().ButtonSet.OK);
      }
      return;
    }

    // Separar por Tipo de Préstamo
    const listEmergencia = pendientes.filter(r => String(r[5]).includes("Emergencia"));
    const listVacaciones = pendientes.filter(r => String(r[5]).includes("Vacaciones"));

    // Crear Hoja Temporal con Nombre ISO
    const now = new Date();
    const fechaNombre = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
    const fechaLegible = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const nombreArchivo = `Informe_Prestamos_${fechaNombre}`;
    
    const ssTemp = SpreadsheetApp.create(nombreArchivo);
    const idTemp = ssTemp.getId();

    const headers = ["ID", "RUT", "NOMBRE", "TIPO PRÉSTAMO", "MONTO", "CUOTAS", "MEDIO PAGO"];

    // --- FUNCIÓN HELPER PARA LLENAR DATOS ---
    const llenarHoja = (sheet, datos) => {
      if (datos.length === 0) return;
      
      sheet.appendRow(headers);
      
      // Estilo Cabecera: Azul Corporativo Oscuro
      sheet.getRange(1, 1, 1, headers.length)
           .setFontWeight("bold")
           .setBackground("#1e3a8a") // Azul oscuro profesional
           .setFontColor("white")
           .setHorizontalAlignment("center");
      
      // Forzar formato texto en MONTO
      sheet.getRange(2, 5, datos.length, 1).setNumberFormat("@");

      datos.forEach(row => {
        let rutLimpio = String(row[2]).replace(/[^0-9kK]/g, '').toUpperCase();
        
        let tipoPrestamo = String(row[5]);
        let montoTexto = "$0";
        if (tipoPrestamo.includes("Emergencia")) montoTexto = "$200.000";
        else if (tipoPrestamo.includes("Vacaciones")) montoTexto = "$150.000";
        else montoTexto = String(row[6]).replace(/'/g, '');
        
        let montoParaExcel = "'" + montoTexto; 

        sheet.appendRow([ row[0], rutLimpio, row[3], row[5], montoParaExcel, row[7], row[8] ]);
      });
      
      sheet.autoResizeColumns(1, headers.length);
      sheet.hideColumns(1); // Ocultar ID
    };

    // Configurar Pestañas
    const sheetEmergencia = ssTemp.getSheets()[0];
    sheetEmergencia.setName("Emergencia");
    if (listEmergencia.length > 0) llenarHoja(sheetEmergencia, listEmergencia);
    else sheetEmergencia.appendRow(["No hay solicitudes de emergencia pendientes."]);

    const sheetVacaciones = ssTemp.insertSheet("Vacaciones");
    if (listVacaciones.length > 0) llenarHoja(sheetVacaciones, listVacaciones);
    else sheetVacaciones.appendRow(["No hay solicitudes de vacaciones pendientes."]);

    SpreadsheetApp.flush();

    // Exportar Blob
    const url = "https://docs.google.com/spreadsheets/d/" + idTemp + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    const blob = response.getBlob().setName(`${nombreArchivo}.xlsx`);

    // --- DISEÑO DE CORREO SOFISTICADO ---
    const destinatarios = "secretario@sindicatoslim3.com,slim3comunicaciones@gmail.com";
    const asunto = `📑 Informe Semanal de Préstamos - ${fechaLegible}`;
    
    const htmlBody = `
    <!DOCTYPE html>
    <html>
    <body style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background-color: #f4f4f5; margin: 0; padding: 0;">
      <div style="max-width: 600px; margin: 30px auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: 1px solid #e5e7eb;">
        
        <div style="background-color: #1e293b; padding: 30px; text-align: center;">
          <h2 style="color: #ffffff; margin: 0; font-size: 22px; font-weight: 600; letter-spacing: 0.5px;">Reporte de Gestión</h2>
          <p style="color: #94a3b8; margin: 5px 0 0; font-size: 13px; text-transform: uppercase; letter-spacing: 1px;">Sindicato SLIM N°3</p>
        </div>

        <div style="padding: 40px 30px;">
          <p style="color: #334155; font-size: 16px; line-height: 1.6; margin-top: 0;">
            Estimados Administradores,
          </p>
          <p style="color: #475569; font-size: 15px; line-height: 1.6;">
            Se ha generado exitosamente el informe consolidado de solicitudes de préstamos correspondientes a la semana en curso (<strong>${fechaLegible}</strong>).
          </p>
          
          <div style="background-color: #f8fafc; border-left: 4px solid #3b82f6; padding: 15px; margin: 25px 0; border-radius: 4px;">
            <p style="margin: 0; color: #1e293b; font-size: 14px;"><strong>Contenido del Archivo:</strong></p>
            <ul style="margin: 10px 0 0 20px; color: #475569; font-size: 14px;">
              <li>Pestaña 1: Solicitudes de Emergencia</li>
              <li>Pestaña 2: Solicitudes de Vacaciones</li>
            </ul>
          </div>

          <p style="color: #475569; font-size: 14px;">
            El archivo Excel se encuentra adjunto a este correo para su revisión y envío a la empresa.
          </p>
        </div>

        <div style="background-color: #f1f5f9; padding: 20px; text-align: center; border-top: 1px solid #e2e8f0;">
          <p style="color: #64748b; font-size: 11px; margin: 0;">
            © ${now.getFullYear()} Plataforma de Gestión Sindicato SLIM N°3<br>
            Este es un mensaje automático del sistema.
          </p>
        </div>
      </div>
    </body>
    </html>
    `;

    MailApp.sendEmail({
      to: destinatarios,
      subject: asunto,
      htmlBody: htmlBody,
      attachments: [blob]
    });

    // Limpieza
    DriveApp.getFileById(idTemp).setTrashed(true);

    // Alerta en Pantalla (Solo si se ejecuta manualmente)
    if (Session.getActiveUser().getEmail()) {
      SpreadsheetApp.getUi().alert(
        'Informe Generado Exitosamente',
        `Se ha procesado la base de datos y el informe ha sido enviado a:\n\n${destinatarios.replace(',', '\n')}\n\nPor favor, verifica tu bandeja de entrada.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

  } catch (e) {
    if (Session.getActiveUser().getEmail()) {
      SpreadsheetApp.getUi().alert('Error en el Proceso', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      console.error(e); // Log para ejecución automática
    }
  }
}

// 3. CONFIGURACIÓN DEL TRIGGER AUTOMÁTICO (SE EJECUTA UNA SOLA VEZ)
function configurarDisparadorAutomatico() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  
  // Verificar si ya existe para no duplicar
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generarInformeConMonto') {
      ui.alert('Configuración Existente', 'La automatización ya se encuentra activa.', ui.ButtonSet.OK);
      return;
    }
  }

  // Crear Trigger: Todos los Lunes cerca de las 16:30
  ScriptApp.newTrigger('generarInformeConMonto')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(16)
      .nearMinute(30)
      .create();

  ui.alert('✅ Automatización Activada', 'El informe se generará y enviará automáticamente todos los días Lunes entre las 16:30 y 17:00 hrs.', ui.ButtonSet.OK);
}
