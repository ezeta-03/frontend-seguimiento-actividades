/**
 * Sistema de Seguimiento de Actividades Multi-Rol
 * Backend API con Google Apps Script
 */

// ============================================
// CONFIGURACIÓN GLOBAL
// ============================================

function testVencimiento() {
  // Cambia este ID por el ID real de una actividad que ya pasó su fecha fin
  const resultado = actualizarActividad('ACT-XXXXX', { estado: 'Completada' });
  Logger.log(JSON.stringify(resultado));
}


const CONFIG = {
  SPREADSHEET_ID: '1UfMW9IwFavcxoUAVGKDQeacall_Pfn-O63IuQZZTK7s',
  CALENDAR_ID: 'primary',
  ROLES: {
    GERENCIA: 'Gerencia',
    JEFATURA: 'Jefatura',
    COLABORADOR: 'Colaborador'
  },
  ESTADOS: {
    PENDIENTE: 'Pendiente',
    EN_PROCESO: 'En Proceso',
    REVISION: 'En Revisión',
    COMPLETADA: 'Completada',
    VENCIDA: 'Vencida',
    COMPLETADA_CON_ATRASO: 'Completada con Atraso',
    CANCELADA: 'Cancelada',
    SUSPENDIDA: 'Suspendida'
  },
  PRIORIDADES: {
    ALTA: 'Alta',
    MEDIA: 'Media',
    BAJA: 'Baja'
  },
  // ← NUEVO: Tipos de actividad
  TIPOS_ACTIVIDAD: {
    RECURRENTE: 'Recurrente',
    EMERGENTE: 'Emergente'
  },
  ESTADOS_EMERGENTE: {
    PENDIENTE_ACEPTACION: 'Pendiente de Aceptación',
    ACEPTADA: 'Aceptada',
    RECHAZADA: 'Rechazada',
    EN_EJECUCION: 'En Ejecución',
    COMPLETADA: 'Completada'
  }
};

// ============================================
// FUNCIONES PRINCIPALES DE LA API
// ============================================

// Lógica:
// 1. Si el usuario tiene sesión activa en Google → sirve Dashboard
// 2. Si no → sirve Login
// 3. Ya no se necesita ?page=dashboard
// ══════════════════════════════════════════════════════════

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('Sistema de Actividades')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Solo se usa para cerrar sesión
function getLoginHtml() {
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .getContent();
}

// Esta función ya la tienes en Code.gs — verifica que exista
// Si no existe, agrégala:
function obtenerUsuarioPorEmail(email) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Usuarios');
    if (!sheet) return null;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) {
        return {
          email:         data[i][2],
          nombre:        data[i][1],
          rol:           data[i][3],
          estado:        data[i][4],
          fechaRegistro: data[i][5] instanceof Date ? data[i][5].toISOString() : ''
        };
      }
    }
    return null;
  } catch(e) {
    Logger.log('obtenerUsuarioPorEmail error: ' + e);
    return null;
  }
}

function obtenerUsuarioPorEmail(email) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Usuarios');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === email) {  // columna C = email
        return {
          email:         data[i][2],
          nombre:        data[i][1],
          rol:           data[i][3],
          estado:        data[i][4],
          fechaRegistro: data[i][5] instanceof Date ? data[i][5].toISOString() : ''
        };
      }
    }
    return null;
  } catch(e) {
    Logger.log('obtenerUsuarioPorEmail error: ' + e);
    return null;
  }
}

// ── INCLUDE ───────────────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── URL DEL APP ───────────────────────────────────────────
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ── DASHBOARD HTML ────────────────────────────────────────
function getDashboardHtml() {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .getContent();
}

function getLoginHtml() {
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .getContent();
}

function obtenerSesionActiva() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return null;

    const usuario = obtenerUsuarioPorEmail(email);
    if (!usuario || usuario.estado !== 'Activo') return null;

    return usuario;
  } catch(e) {
    Logger.log('obtenerSesionActiva error: ' + e);
    return null;
  }
}

/**
 * Incluye archivos HTML parciales
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// GESTIÓN DE USUARIOS Y AUTENTICACIÓN
// ============================================

/**
 * Obtiene el usuario actual - VERSIÓN FINAL CON FECHAS ISO
 */
function obtenerUsuarioActual() {
  try {
    Logger.log('=== INICIO obtenerUsuarioActual ===');
    
    const email = Session.getActiveUser().getEmail();
    Logger.log('1. Email obtenido: ' + email);
    
    if (!email) {
      Logger.log('ERROR: No hay email');
      throw new Error('No se pudo obtener el email del usuario');
    }
    
    Logger.log('2. Intentando abrir spreadsheet: ' + CONFIG.SPREADSHEET_ID);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    if (!ss) {
      Logger.log('ERROR: No se pudo abrir el spreadsheet');
      throw new Error('No se pudo acceder a la hoja de cálculo');
    }
    
    Logger.log('2. Spreadsheet abierto: ' + ss.getName());
    
    Logger.log('3. Buscando hoja Usuarios...');
    let sheet = ss.getSheetByName('Usuarios');
    
    if (!sheet) {
      Logger.log('Hoja Usuarios no existe, creando...');
      sheet = ss.insertSheet('Usuarios');
      sheet.appendRow(['Email', 'Nombre', 'Rol', 'Estado', 'FechaRegistro']);
      sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
      Logger.log('Hoja Usuarios creada, registrando nuevo usuario...');
      return registrarNuevoUsuario(email);
    }
    
    Logger.log('3. Hoja Usuarios encontrada');
    
    Logger.log('4. Obteniendo datos de la hoja...');
    const lastRow = sheet.getLastRow();
    Logger.log('4. Última fila con datos: ' + lastRow);
    
    if (lastRow < 2) {
      Logger.log('4. No hay usuarios registrados, creando primero...');
      return registrarNuevoUsuario(email);
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('4. Datos obtenidos: ' + data.length + ' filas');
    
    Logger.log('5. Buscando usuario con email: ' + email);
    
    for (let i = 1; i < data.length; i++) {
      Logger.log('5. Comparando con fila ' + i + ': ' + data[i][0]);
      
      if (data[i][0] === email) {
        const usuario = {
          email: String(data[i][0]),
          nombre: String(data[i][1]),
          rol: String(data[i][2]),
          estado: String(data[i][3]),
          fechaRegistro: data[i][4] ? new Date(data[i][4]).toISOString() : null
        };
        
        Logger.log('5. ✅ Usuario encontrado: ' + JSON.stringify(usuario));
        Logger.log('=== FIN obtenerUsuarioActual ===');
        return usuario;
      }
    }
    
    Logger.log('6. Usuario no encontrado, registrando nuevo usuario...');
    const nuevoUsuario = registrarNuevoUsuario(email);
    Logger.log('6. ✅ Nuevo usuario creado: ' + JSON.stringify(nuevoUsuario));
    Logger.log('=== FIN obtenerUsuarioActual ===');
    
    return nuevoUsuario;
    
  } catch (error) {
    Logger.log('❌ ERROR en obtenerUsuarioActual: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    Logger.log('=== FIN obtenerUsuarioActual (CON ERROR) ===');
    throw error;
  }
}

/**
 * Registra un nuevo usuario - VERSIÓN FINAL
 */
function registrarNuevoUsuario(email) {
  try {
    Logger.log('=== INICIO registrarNuevoUsuario ===');
    Logger.log('Email: ' + email);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Usuarios');
    
    if (!sheet) {
      Logger.log('Creando hoja Usuarios...');
      sheet = ss.insertSheet('Usuarios');
      sheet.appendRow(['Email', 'Nombre', 'Rol', 'Estado', 'FechaRegistro']);
      sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
      Logger.log('Hoja Usuarios creada');
    }
    
    const nombre = email.split('@')[0];
    const ahora = new Date();
    
    const nuevoUsuario = {
      email: email,
      nombre: nombre,
      rol: CONFIG.ROLES.COLABORADOR,
      estado: 'Activo',
      fechaRegistro: ahora.toISOString()
    };
    
    Logger.log('Nuevo usuario: ' + JSON.stringify(nuevoUsuario));
    
    sheet.appendRow([
      nuevoUsuario.email,
      nuevoUsuario.nombre,
      nuevoUsuario.rol,
      nuevoUsuario.estado,
      ahora
    ]);
    
    Logger.log('✅ Usuario registrado en la hoja');
    Logger.log('=== FIN registrarNuevoUsuario ===');
    
    return nuevoUsuario;
    
  } catch (error) {
    Logger.log('❌ ERROR en registrarNuevoUsuario: ' + error.toString());
    throw error;
  }
}

/**
 * Obtiene todos los usuarios por rol
 */
function obtenerUsuariosPorRol(rol) {
  try {
    Logger.log('=== INICIO obtenerUsuariosPorRol ===');
    Logger.log('Rol solicitado: ' + rol);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Usuarios');
    
    if (!sheet) {
      Logger.log('ERROR: No existe la hoja Usuarios');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('Total de filas en Usuarios: ' + data.length);
    
    const usuarios = [];
    
    for (let i = 1; i < data.length; i++) {
      const rolUsuario = data[i][2];
      const estadoUsuario = data[i][3];
      
      if (rolUsuario === rol && estadoUsuario === 'Activo') {
        const usuario = {
          email: String(data[i][0]),
          nombre: String(data[i][1]),
          rol: String(data[i][2])
        };
        
        usuarios.push(usuario);
        Logger.log('  ✅ Usuario agregado: ' + usuario.nombre);
      }
    }
    
    Logger.log('Total usuarios con rol "' + rol + '": ' + usuarios.length);
    Logger.log('=== FIN obtenerUsuariosPorRol ===');
    
    return usuarios;
    
  } catch (error) {
    Logger.log('❌ ERROR en obtenerUsuariosPorRol: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return [];
  }
}

/**
 * Función de prueba para obtenerUsuariosPorRol
 */
function probarObtenerUsuariosPorRol() {
  Logger.log('=== PRUEBA obtenerUsuariosPorRol ===');
  
  const colaboradores = obtenerUsuariosPorRol('Colaborador');
  Logger.log('Colaboradores encontrados: ' + colaboradores.length);
  colaboradores.forEach(u => Logger.log('  - ' + u.nombre + ' (' + u.email + ')'));
  
  const jefes = obtenerUsuariosPorRol('Jefatura');
  Logger.log('Jefes encontrados: ' + jefes.length);
  jefes.forEach(u => Logger.log('  - ' + u.nombre + ' (' + u.email + ')'));
  
  const gerentes = obtenerUsuariosPorRol('Gerencia');
  Logger.log('Gerentes encontrados: ' + gerentes.length);
  gerentes.forEach(u => Logger.log('  - ' + u.nombre + ' (' + u.email + ')'));
  
  return {
    colaboradores: colaboradores,
    jefes: jefes,
    gerentes: gerentes
  };
}

/**
 * Actualiza el rol de un usuario (solo Gerencia)
 */
function actualizarRolUsuario(email, nuevoRol) {
  const usuario = obtenerUsuarioActual();
  
  if (usuario.rol !== CONFIG.ROLES.GERENCIA) {
    throw new Error('No tienes permisos para cambiar roles');
  }
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Usuarios');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      sheet.getRange(i + 1, 3).setValue(nuevoRol);
      return { success: true, message: 'Rol actualizado correctamente' };
    }
  }
  
  throw new Error('Usuario no encontrado');
}

// ============================================
// GESTIÓN DE ACTIVIDADES
// ============================================

/**
 * Reglas por rol:
 * - Gerencia:    auto-asignada, sin personas adicionales
 * - Jefatura:    auto-asignada, puede agregar jefaturas Y colaboradores
 * - Colaborador: auto-asignado, puede agregar otros colaboradores
 */
function crearActividad(datos) {
  try {
    Logger.log('=== CREANDO ACTIVIDAD ===');
    const usuario = obtenerUsuarioActual();
    Logger.log('Usuario: ' + usuario.email + ' (' + usuario.rol + ')');

    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    if (!sheet) return { success: false, message: 'No se encontró la hoja Actividades' };

    const id = 'ACT-' + new Date().getTime();

    // Normalizar adicionales
    let jefaturasAdicionales     = Array.isArray(datos.jefaturasAdicionales)
      ? datos.jefaturasAdicionales.filter(Boolean) : [];
    let colaboradoresAdicionales = Array.isArray(datos.colaboradoresAdicionales)
      ? datos.colaboradoresAdicionales.filter(Boolean) : [];

    // Asignación y restricciones según rol
    if (usuario.rol === CONFIG.ROLES.GERENCIA) {
      datos.responsable            = usuario.email;
      datos.jefe                   = '';
      datos.gerente                = '';
      jefaturasAdicionales         = [];
      colaboradoresAdicionales     = [];

    } else if (usuario.rol === CONFIG.ROLES.JEFATURA) {
      datos.responsable = usuario.email;
      datos.jefe        = usuario.email;
      datos.gerente     = datos.gerente || '';
      // jefaturasAdicionales y colaboradoresAdicionales vienen del form

    } else if (usuario.rol === CONFIG.ROLES.COLABORADOR) {
      datos.responsable        = usuario.email;
      datos.jefe               = datos.jefe   || '';
      datos.gerente            = datos.gerente || '';
      jefaturasAdicionales     = []; // Colaborador no agrega jefaturas
    }

    Logger.log('Responsable: '             + datos.responsable);
    Logger.log('JefaturasAdicionales: '    + jefaturasAdicionales.join(','));
    Logger.log('ColaboradoresAdicionales: '+ colaboradoresAdicionales.join(','));

    const ahora = new Date();
if (usuario.rol === "Jefatura" || usuario.rol === "Colaborador") {
  const ini = datos.fechaInicio ? new Date(datos.fechaInicio) : null;
  const fin = datos.fechaFin ? new Date(datos.fechaFin) : null;
  if (ini && fin) {
    const cruce = verificarCruceYSugerencias(datos.responsable, ini, fin, null);
    if (cruce.cruce)
      return {
        success: false,
        cruce: true,
        actividad: cruce.actividad,
        horasLibres: cruce.horasLibres,
      };
  }
}
    sheet.appendRow([
      id,                                            // A  ID
      datos.proyectoId  || '',                       // B  ProyectoID
      datos.titulo      || '',                       // C  Título
      datos.descripcion || '',                       // D  Descripción
      datos.responsable || '',                       // E  Responsable
      datos.jefe        || '',                       // F  Jefe
      datos.gerente     || '',                       // G  Gerente
      CONFIG.ESTADOS.PENDIENTE,                      // H  Estado
      datos.prioridad   || CONFIG.PRIORIDADES.MEDIA, // I  Prioridad
      datos.fechaInicio ? new Date(datos.fechaInicio) : '', // J FechaInicio
      datos.fechaFin    ? new Date(datos.fechaFin)    : '', // K FechaFin
      ahora,                                         // L  FechaCreacion
      ahora,                                         // M  UltimaActualizacion
      0,                                             // N  Avance
      'Recurrente',                                  // O  Tipo
      '',                                            // P  FechaAceptacion
      0,                                             // Q  TiempoEjecucion
      '',                                            // R  EstadoAnterior
      '',                                            // S  FechaSuspension
      jefaturasAdicionales.join(','),                // T  JefaturasAdicionales
      colaboradoresAdicionales.join(',')             // U  ColaboradoresAdicionales
    ]);

    try { registrarHistorial(id, usuario.email, 'Creación', 'Actividad creada'); } catch(e){ Logger.log('Error historial: ' + e); }
    try { enviarNotificacionNuevaActividad({ ...datos, id }); }                    catch(e){ Logger.log('Error notif: '     + e); }

    return {
      success: true,
      message: 'Actividad creada correctamente',
      actividad: {
        id,
        proyectoId:               datos.proyectoId  || '',
        titulo:                   datos.titulo      || '',
        descripcion:              datos.descripcion || '',
        responsable:              datos.responsable || '',
        jefe:                     datos.jefe        || '',
        gerente:                  datos.gerente     || '',
        estado:                   CONFIG.ESTADOS.PENDIENTE,
        prioridad:                datos.prioridad   || CONFIG.PRIORIDADES.MEDIA,
        fechaInicio:              datos.fechaInicio || '',
        fechaFin:                 datos.fechaFin    || '',
        fechaCreacion:            ahora.toISOString(),
        ultimaActualizacion:      ahora.toISOString(),
        avance:                   0,
        jefaturasAdicionales,
        colaboradoresAdicionales
      }
    };

  } catch (error) {
    Logger.log('Error en crearActividad: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


/**
 * Visibilidad:
 * - Gerencia:    ve todas
 * - Jefatura:    ve las suyas + donde aparece en jefe o JefaturasAdicionales
 * - Colaborador: ve las suyas + donde aparece en ColaboradoresAdicionales
 */
function obtenerActividades(filtros = {}) {
  try {
    Logger.log('=== INICIO obtenerActividades ===');
    const usuario = obtenerUsuarioActual();
    Logger.log('Usuario: ' + usuario.email + ' (' + usuario.rol + ')');

    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];

    const actividades = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const jefaturasAdicionales     = data[i][19] ? String(data[i][19]).split(',').map(e => e.trim()).filter(Boolean) : [];
      const colaboradoresAdicionales = data[i][20] ? String(data[i][20]).split(',').map(e => e.trim()).filter(Boolean) : [];

      const actividad = {
        id:                      data[i][0]  || '',
        proyectoId:              data[i][1]  || '',
        titulo:                  data[i][2]  || '',
        descripcion:             data[i][3]  || '',
        responsable:             data[i][4]  || '',
        jefe:                    data[i][5]  || '',
        gerente:                 data[i][6]  || '',
        estado:                  data[i][7]  || CONFIG.ESTADOS.PENDIENTE,
        prioridad:               data[i][8]  || CONFIG.PRIORIDADES.MEDIA,
        fechaInicio:             data[i][9]  instanceof Date ? data[i][9].toISOString()  : (data[i][9]  || ''),
        fechaFin:                data[i][10] instanceof Date ? data[i][10].toISOString() : (data[i][10] || ''),
        fechaCreacion:           data[i][11] instanceof Date ? data[i][11].toISOString() : (data[i][11] || ''),
        ultimaActualizacion:     data[i][12] instanceof Date ? data[i][12].toISOString() : (data[i][12] || ''),
        avance:                  data[i][13] || 0,
        tipo:                    data[i][14] || 'Recurrente',
        fechaAceptacion:         data[i][15] ? new Date(data[i][15]).toISOString() : null,
        tiempoEjecucion:         data[i][16] || 0,
        estadoAnterior:          data[i][17] || '',
        fechaSuspension:         data[i][18] ? new Date(data[i][18]).toISOString() : null,
        jefaturasAdicionales,
        colaboradoresAdicionales
      };

      let incluir = false;

      switch (usuario.rol) {
        case CONFIG.ROLES.GERENCIA:
          incluir = true;
          break;

        case CONFIG.ROLES.JEFATURA:
          incluir = actividad.responsable === usuario.email ||
                    actividad.jefe        === usuario.email ||
                    actividad.gerente     === usuario.email ||
                    jefaturasAdicionales.includes(usuario.email);
          break;

        case CONFIG.ROLES.COLABORADOR:
          incluir = actividad.responsable === usuario.email ||
                    colaboradoresAdicionales.includes(usuario.email);
          break;
      }

      if (incluir) actividades.push(actividad);
    }

    Logger.log('Actividades encontradas: ' + actividades.length);
    return actividades;

  } catch (error) {
    Logger.log('ERROR en obtenerActividades: ' + error.toString());
    return [];
  }
}

/**
 * Ejecutada por el trigger diario a las 8am.
 * Marca como 'Vencida' toda actividad con fechaFin < hoy
 * cuyo estado sea Pendiente, En Proceso o En Revisión.
 * Envía email al responsable y al jefe.
 * Registra en la hoja 'Historial de Vencimientos'.
 */
function marcarActividadesVencidas() {
  try {
    Logger.log('=== TRIGGER: marcarActividadesVencidas ===');

    const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet   = ss.getSheetByName('Actividades');
    if (!sheet) { Logger.log('Sin hoja Actividades'); return; }

    // Obtener o crear hoja de historial
    let histSheet = ss.getSheetByName('Historial de Vencimientos');
    if (!histSheet) {
      histSheet = ss.insertSheet('Historial de Vencimientos');
      histSheet.getRange(1, 1, 1, 6).setValues([[
        'Fecha Marcado', 'ID Actividad', 'Título', 'Responsable', 'Jefe', 'Fecha Fin Original'
      ]]).setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
    }

    const data    = sheet.getDataRange().getValues();
    const ahora   = new Date();
    const hoy     = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
    const estados = ['Pendiente', 'En Proceso', 'En Revisión'];
    let   marcadas = 0;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const estadoActual = data[i][7]  || '';
      const fechaFin     = data[i][10] ? new Date(data[i][10]) : null;

      if (!fechaFin || !estados.includes(estadoActual)) continue;

      const fechaFinSoloFecha = new Date(
        fechaFin.getFullYear(), fechaFin.getMonth(), fechaFin.getDate()
      );

      if (fechaFinSoloFecha >= hoy) continue; // No vencida aún

      const filaSheet = i + 1;
      const id          = data[i][0];
      const titulo      = data[i][2] || '';
      const responsable = data[i][4] || '';
      const jefe        = data[i][5] || '';

      // Marcar como Vencida
      sheet.getRange(filaSheet, 8).setValue('Vencida');
      sheet.getRange(filaSheet, 13).setValue(ahora);

      // Registrar en historial
      histSheet.appendRow([
        ahora, id, titulo, responsable, jefe,
        fechaFin instanceof Date ? fechaFin.toLocaleDateString('es-CL') : fechaFin
      ]);

      // Enviar emails
      enviarEmailVencimiento({ id, titulo, responsable, jefe, fechaFin });

      marcadas++;
      Logger.log('Marcada como Vencida: ' + id + ' - ' + titulo);
    }

    Logger.log('Total marcadas: ' + marcadas);

  } catch(error) {
    Logger.log('ERROR en marcarActividadesVencidas: ' + error.toString());
  }
}



function actualizarActividad(id, cambios) {
  try {
    Logger.log('=== ACTUALIZANDO ACTIVIDAD: ' + id + ' ===');
    const usuario = obtenerUsuarioActual();

    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    if (!sheet) return { success: false, message: 'Hoja no encontrada' };

    const data     = sheet.getDataRange().getValues();
    let   filaIdx  = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) { filaIdx = i + 1; break; }
    }

    if (filaIdx === -1) return { success: false, message: 'Actividad no encontrada' };

    const fila         = sheet.getRange(filaIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
    const estadoActual = fila[7]  || '';
    const fechaFin     = fila[10] ? new Date(fila[10]) : null;
    const ahora        = new Date();

if (usuario.rol === "Jefatura" || usuario.rol === "Colaborador") {
  const ini = cambios.fechaInicio
    ? new Date(cambios.fechaInicio)
    : fechaInicio
      ? new Date(fechaInicio)
      : null;
  const fin = cambios.fechaFin
    ? new Date(cambios.fechaFin)
    : fechaFin
      ? new Date(fechaFin)
      : null;
  const resp = cambios.responsable || fila[4];
  if (ini && fin) {
    const cruce = verificarCruceYSugerencias(resp, ini, fin, id);
    if (cruce.cruce)
      return {
        success: false,
        cruce: true,
        actividad: cruce.actividad,
        horasLibres: cruce.horasLibres,
      };
  }
}

    // ── Reactivar Vencida si se extiende la fecha ────────
    if (estadoActual === 'Vencida' && cambios.fechaFin) {
      const nuevaFechaFin = new Date(cambios.fechaFin);
      if (nuevaFechaFin > ahora) {
        // Forzar En Proceso independientemente de lo que traiga el select
        cambios.estado = 'En Proceso';
        Logger.log('Vencida reactivada a En Proceso por nueva fechaFin');
      }
    }

    // ── Completar con atraso ─────────────────────────────
    if (cambios.estado === 'Completada') {
      const fechaFinEfectiva = cambios.fechaFin ? new Date(cambios.fechaFin) : fechaFin;
      const estaVencida      = estadoActual === 'Vencida';
      const yaVencioFecha    = fechaFinEfectiva && fechaFinEfectiva < ahora;
      if (estaVencida || yaVencioFecha) {
        cambios.estado = 'Completada con Atraso';
        Logger.log('Completada con atraso');
      }
    }

    // ── Estado → Avance automático ───────────────────────
    const estadoFinal = cambios.estado !== undefined ? cambios.estado : estadoActual;
    if (estadoFinal === 'Completada' || estadoFinal === 'Completada con Atraso') {
      cambios.avance = 100;
    } else if (estadoFinal === 'Pendiente') {
      if (cambios.avance === undefined) cambios.avance = 0;
    }

    // ── Avance → Estado automático ───────────────────────
    if (cambios.avance !== undefined && cambios.estado === undefined) {
      const avance = Number(cambios.avance);
      if (avance === 100) {
        const yaVencioFecha = fechaFin && fechaFin < ahora;
        cambios.estado = (estadoActual === 'Vencida' || yaVencioFecha)
          ? 'Completada con Atraso'
          : 'Completada';
      } else if (avance === 0 && estadoActual === 'Completada') {
        cambios.estado = 'Pendiente';
      }
    }

    // ── Aplicar cambios celda por celda ──────────────────
    const mapa = {
      proyectoId:   2,   // B
      titulo:       3,   // C
      descripcion:  4,   // D
      responsable:  5,   // E
      jefe:         6,   // F
      gerente:      7,   // G
      estado:       8,   // H
      prioridad:    9,   // I
      fechaInicio:  10,  // J
      fechaFin:     11,  // K
      avance:       14   // N
    };

    Object.keys(cambios).forEach(function(campo) {
      if (mapa[campo] === undefined) return;
      const col  = mapa[campo];
      let   valor = cambios[campo];

      if ((campo === 'fechaInicio' || campo === 'fechaFin') && valor) {
        valor = new Date(valor);
      }
      if (campo === 'avance') valor = Number(valor) || 0;

      sheet.getRange(filaIdx, col).setValue(valor);
    });

    // Actualizar ultimaActualizacion (col M = 13)
    sheet.getRange(filaIdx, 13).setValue(ahora);

    try {
      registrarHistorial(
        id,
        usuario.email,
        'Actualización',
        'Estado: ' + (cambios.estado || estadoActual) +
        (cambios.avance !== undefined ? ' | Avance: ' + cambios.avance + '%' : '')
      );
    } catch(e) { Logger.log('Error historial: ' + e); }

    // ── Email de completado ───────────────────────────────
    const estadoGuardado = cambios.estado || estadoActual;
    if (estadoGuardado === 'Completada' || estadoGuardado === 'Completada con Atraso') {
      try {
        const tituloAct     = cambios.titulo      || fila[2] || '';
        const responsableAct = cambios.responsable || fila[4] || '';
        const jefeAct        = cambios.jefe        || fila[5] || '';
        const fechaFinAct    = cambios.fechaFin    ? new Date(cambios.fechaFin) : fechaFin;

        enviarEmailCompletado({
          id,
          titulo:      tituloAct,
          responsable: responsableAct,
          jefe:        jefeAct,
          fechaFin:    fechaFinAct,
          estado:      estadoGuardado,
          completadoPor: usuario.email
        });
      } catch(e) { Logger.log('Error email completado: ' + e); }
    }

    return { success: true, message: 'Actividad actualizada' };

  } catch(error) {
    Logger.log('Error en actualizarActividad: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Envía email de confirmación al responsable y jefe
 * cuando una actividad es completada (a tiempo o con atraso).
 */
/**
 * Envía email de confirmación al responsable y jefe
 * cuando una actividad es completada (a tiempo o con atraso).
 */
/**
 * Envía email de confirmación al responsable y jefe
 * cuando una actividad es completada (a tiempo o con atraso).
 */
function enviarEmailCompletado(actividad) {
  try {
    const conAtraso = actividad.estado === 'Completada con Atraso';
    const asunto    = conAtraso
      ? '[Atraso] Actividad completada: ' + actividad.titulo
      : '[Completada] Actividad finalizada: ' + actividad.titulo;

    const fechaFinStr = actividad.fechaFin instanceof Date
      ? actividad.fechaFin.toLocaleDateString('es-CL')
      : (actividad.fechaFin || 'No definida');

    const cuerpo = `
Hola,

La siguiente actividad ha sido marcada como "${actividad.estado}":

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 Actividad:    ${actividad.titulo}
🆔 ID:           ${actividad.id}
📅 Fecha límite: ${fechaFinStr}
👤 Completada por: ${actividad.completadoPor}
📊 Estado final: ${actividad.estado}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${conAtraso ? '\n⚠️  Esta actividad fue completada fuera del plazo original.\n' : ''}
— Sistema de Actividades | ZaazMago
    `.trim();

    // Email al responsable
    if (actividad.responsable) {
      try {
        MailApp.sendEmail({ to: actividad.responsable, subject: asunto, body: cuerpo });
        Logger.log('Email completado enviado a: ' + actividad.responsable);
      } catch(e) { Logger.log('Error email responsable: ' + e); }
    }

    // Email al jefe (si es diferente al responsable)
    if (actividad.jefe && actividad.jefe !== actividad.responsable) {
      try {
        MailApp.sendEmail({ to: actividad.jefe, subject: asunto, body: cuerpo });
        Logger.log('Email completado enviado a jefe: ' + actividad.jefe);
      } catch(e) { Logger.log('Error email jefe: ' + e); }
    }

  } catch(error) {
    Logger.log('Error en enviarEmailCompletado: ' + error.toString());
  }
}

/**
 * Ejecutada por el trigger diario a las 8am.
 * Marca como 'Vencida' toda actividad con fechaFin < hoy
 * cuyo estado sea Pendiente, En Proceso o En Revisión.
 * Envía email al responsable y al jefe.
 * Registra en la hoja 'Historial de Vencimientos'.
 */
function marcarActividadesVencidas() {
  try {
    Logger.log('=== TRIGGER: marcarActividadesVencidas ===');

    const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet   = ss.getSheetByName('Actividades');
    if (!sheet) { Logger.log('Sin hoja Actividades'); return; }

    // Obtener o crear hoja de historial
    let histSheet = ss.getSheetByName('Historial de Vencimientos');
    if (!histSheet) {
      histSheet = ss.insertSheet('Historial de Vencimientos');
      histSheet.getRange(1, 1, 1, 6).setValues([[
        'Fecha Marcado', 'ID Actividad', 'Título', 'Responsable', 'Jefe', 'Fecha Fin Original'
      ]]).setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
    }

    const data    = sheet.getDataRange().getValues();
    const ahora   = new Date();
    const hoy     = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate());
    const estados = ['Pendiente', 'En Proceso', 'En Revisión'];
    let   marcadas = 0;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const estadoActual = data[i][7]  || '';
      const fechaFin     = data[i][10] ? new Date(data[i][10]) : null;

      if (!fechaFin || !estados.includes(estadoActual)) continue;

      const fechaFinSoloFecha = new Date(
        fechaFin.getFullYear(), fechaFin.getMonth(), fechaFin.getDate()
      );

      if (fechaFinSoloFecha >= hoy) continue; // No vencida aún

      const filaSheet = i + 1;
      const id          = data[i][0];
      const titulo      = data[i][2] || '';
      const responsable = data[i][4] || '';
      const jefe        = data[i][5] || '';

      // Marcar como Vencida
      sheet.getRange(filaSheet, 8).setValue('Vencida');
      sheet.getRange(filaSheet, 13).setValue(ahora);

      // Registrar en historial
      histSheet.appendRow([
        ahora, id, titulo, responsable, jefe,
        fechaFin instanceof Date ? fechaFin.toLocaleDateString('es-CL') : fechaFin
      ]);

      // Enviar emails
      enviarEmailVencimiento({ id, titulo, responsable, jefe, fechaFin });

      marcadas++;
      Logger.log('Marcada como Vencida: ' + id + ' - ' + titulo);
    }

    Logger.log('Total marcadas: ' + marcadas);

  } catch(error) {
    Logger.log('ERROR en marcarActividadesVencidas: ' + error.toString());
  }
}


/**
 * Envía email de alerta a responsable y jefe cuando una actividad vence.
 */
function enviarEmailVencimiento(actividad) {
  try {
    const asunto = '[Vencida] Actividad sin completar: ' + actividad.titulo;

    const cuerpo = `
Hola,

La siguiente actividad ha vencido sin ser completada:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 Actividad: ${actividad.titulo}
🆔 ID: ${actividad.id}
📅 Fecha límite: ${actividad.fechaFin instanceof Date
  ? actividad.fechaFin.toLocaleDateString('es-CL')
  : actividad.fechaFin}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Su estado ha sido actualizado automáticamente a "Vencida".

Puedes acceder al Sistema de Actividades para tomar acción.

— Sistema de Actividades | ZaazMago
    `.trim();

    // Email al responsable
    if (actividad.responsable) {
      try {
        MailApp.sendEmail({ to: actividad.responsable, subject: asunto, body: cuerpo });
        Logger.log('Email enviado a responsable: ' + actividad.responsable);
      } catch(e) { Logger.log('Error email responsable: ' + e); }
    }

    // Email al jefe (si existe y es diferente al responsable)
    if (actividad.jefe && actividad.jefe !== actividad.responsable) {
      try {
        MailApp.sendEmail({ to: actividad.jefe, subject: asunto, body: cuerpo });
        Logger.log('Email enviado a jefe: ' + actividad.jefe);
      } catch(e) { Logger.log('Error email jefe: ' + e); }
    }

  } catch(error) {
    Logger.log('Error en enviarEmailVencimiento: ' + error.toString());
  }
}

/**
 * Instala el trigger que ejecuta marcarActividadesVencidas()
 * todos los días a las 8am.
 * EJECUTAR UNA SOLA VEZ manualmente desde el editor.
 */
function crearTriggerDiario() {
  // Eliminar triggers existentes del mismo nombre para no duplicar
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'marcarActividadesVencidas') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Trigger anterior eliminado');
    }
  });

  // Crear nuevo trigger diario a las 8am
  ScriptApp.newTrigger('marcarActividadesVencidas')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('✅ Trigger diario creado: marcarActividadesVencidas() a las 8am');
  return 'Trigger creado correctamente';
}

/**
 * Valida permisos de edición
 */
function validarPermisosEdicion(usuario, actividad) {
  switch (usuario.rol) {
    case CONFIG.ROLES.GERENCIA:
      return true;
    case CONFIG.ROLES.JEFATURA:
      return usuario.email === actividad.jefe || usuario.email === actividad.responsable;
    case CONFIG.ROLES.COLABORADOR:
      return usuario.email === actividad.responsable;
    default:
      return false;
  }
}

/**
 * Elimina una actividad (solo Gerencia y Jefatura)
 */
function eliminarActividad(id) {
  const usuario = obtenerUsuarioActual();

  if (usuario.rol === CONFIG.ROLES.COLABORADOR) {
    throw new Error('No tienes permisos para eliminar actividades');
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Actividades');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      registrarHistorial(id, usuario.email, 'Eliminación', 'Actividad eliminada');
      return { success: true, message: 'Actividad eliminada' };
    }
  }

  throw new Error('Actividad no encontrada');
}

// ============================================
// GESTIÓN DE COMENTARIOS
// ============================================

/**
 * Agrega un comentario a una actividad
 */
function agregarComentario(actividadId, comentario) {
  const usuario = obtenerUsuarioActual();
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Comentarios');

  const id = 'COM-' + new Date().getTime();

  sheet.appendRow([
    id,
    actividadId,
    usuario.email,
    comentario,
    new Date()
  ]);

  // Notificar a los involucrados
  enviarNotificacionComentario(actividadId, usuario.nombre, comentario);

  return { success: true, id: id };
}

/**
 * Obtiene comentarios de una actividad
 */
function obtenerComentarios(actividadId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Comentarios');
  const data = sheet.getDataRange().getValues();

  const comentarios = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === actividadId) {
      comentarios.push({
        id: data[i][0],
        actividadId: data[i][1],
        usuario: data[i][2],
        comentario: data[i][3],
        fecha: data[i][4]
      });
    }
  }

  return comentarios;
}

// ============================================
// CALENDARIO
// ============================================

/**
 * Crea evento en Google Calendar
 */
function crearEventoCalendario(actividad) {
  const calendario = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);

  const evento = calendario.createEvent(
    actividad.titulo,
    actividad.fechaInicio,
    actividad.fechaFin,
    {
      description: `${actividad.descripcion}\n\nResponsable: ${actividad.responsable}\nEstado: ${actividad.estado}\nPrioridad: ${actividad.prioridad}\n\nID: ${actividad.id}`,
      location: ''
    }
  );

  // Agregar recordatorios
  evento.addEmailReminder(60); // 1 hora antes
  evento.addPopupReminder(30); // 30 minutos antes

  return evento.getId();
}

/**
 * Exporta UNA actividad al Google Calendar del usuario activo.
 */
function exportarActividadAlCalendario(actividadId) {
  try {
    Logger.log('=== exportarActividadAlCalendario: ' + actividadId + ' ===');

    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    if (!sheet) throw new Error('No existe la hoja Actividades');

    const data = sheet.getDataRange().getValues();
    let actividad = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === actividadId) {
        actividad = {
          id:          data[i][0],
          titulo:      data[i][1],
          descripcion: data[i][2],
          responsable: data[i][3],
          jefe:        data[i][4],
          gerente:     data[i][5],
          estado:      data[i][6],
          prioridad:   data[i][7],
          fechaInicio: data[i][8] ? new Date(data[i][8]) : null,
          fechaFin:    data[i][9] ? new Date(data[i][9]) : null
        };
        break;
      }
    }

    if (!actividad) throw new Error('Actividad no encontrada: ' + actividadId);
    if (!actividad.fechaInicio || !actividad.fechaFin)
      throw new Error('La actividad no tiene fechas válidas');

    const calendario = CalendarApp.getDefaultCalendar();
    const colorMap = {
      'Alta':  CalendarApp.EventColor.TOMATO,
      'Media': CalendarApp.EventColor.BANANA,
      'Baja':  CalendarApp.EventColor.SAGE
    };

    const evento = calendario.createEvent(
      actividad.titulo,
      actividad.fechaInicio,
      actividad.fechaFin,
      {
        description: [
          'Estado: '      + actividad.estado,
          'Prioridad: '   + actividad.prioridad,
          'Responsable: ' + actividad.responsable,
          actividad.jefe     ? 'Jefe: '    + actividad.jefe    : '',
          actividad.gerente  ? 'Gerente: ' + actividad.gerente : '',
          '',
          actividad.descripcion || 'Sin descripción',
          '',
          'ID Sistema: ' + actividad.id
        ].filter(Boolean).join('\n')
      }
    );

    if (colorMap[actividad.prioridad]) evento.setColor(colorMap[actividad.prioridad]);
    evento.addPopupReminder(30);
    evento.addEmailReminder(60);

    Logger.log('✅ Evento creado: ' + evento.getId());
    return { success: true, message: 'Evento agregado a tu calendario' };

  } catch (error) {
    Logger.log('❌ exportarActividadAlCalendario: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Exporta TODAS las actividades visibles del usuario al calendario.
 */
function exportarTodasAlCalendario() {
  try {
    Logger.log('=== exportarTodasAlCalendario ===');

    const actividades = obtenerActividades();
    if (!actividades || actividades.length === 0) {
      return { success: false, message: 'No hay actividades para exportar' };
    }

    const calendario = CalendarApp.getDefaultCalendar();
    const colorMap = {
      'Alta':  CalendarApp.EventColor.TOMATO,
      'Media': CalendarApp.EventColor.BANANA,
      'Baja':  CalendarApp.EventColor.SAGE
    };
    let creados = 0;
    let errores = 0;

    actividades.forEach(function(act) {
      try {
        const inicio = act.fechaInicio ? new Date(act.fechaInicio) : null;
        const fin    = act.fechaFin    ? new Date(act.fechaFin)    : null;
        if (!inicio || !fin) { errores++; return; }

        const evento = calendario.createEvent(
          act.titulo, inicio, fin,
          {
            description: [
              'Estado: '      + act.estado,
              'Prioridad: '   + act.prioridad,
              'Responsable: ' + act.responsable,
              act.jefe     ? 'Jefe: '    + act.jefe    : '',
              act.gerente  ? 'Gerente: ' + act.gerente : '',
              '',
              act.descripcion || 'Sin descripción',
              '',
              'ID Sistema: ' + act.id
            ].filter(Boolean).join('\n')
          }
        );

        if (colorMap[act.prioridad]) evento.setColor(colorMap[act.prioridad]);
        evento.addPopupReminder(30);
        evento.addEmailReminder(60);
        creados++;
      } catch (e) {
        Logger.log('  ⚠️ Error con actividad ' + act.id + ': ' + e.toString());
        errores++;
      }
    });

    Logger.log('✅ Exportación completa. Creados: ' + creados + ' | Errores: ' + errores);
    return {
      success: true,
      message: 'Se exportaron ' + creados + ' actividad(es) a tu calendario.' +
               (errores > 0 ? ' (' + errores + ' sin fechas válidas)' : ''),
      creados: creados,
      errores: errores
    };

  } catch (error) {
    Logger.log('❌ exportarTodasAlCalendario: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ============================================
// NOTIFICACIONES
// ============================================

/**
 * Envía notificación de nueva actividad
 */
function enviarNotificacionNuevaActividad(actividad) {
  const destinatarios = [actividad.responsable];

  if (actividad.jefe) destinatarios.push(actividad.jefe);
  if (actividad.gerente) destinatarios.push(actividad.gerente);

  const asunto = `Nueva actividad asignada: ${actividad.titulo}`;
  const cuerpo = `
    <h2>Nueva Actividad Asignada</h2>
    <p><strong>Título:</strong> ${actividad.titulo}</p>
    <p><strong>Descripción:</strong> ${actividad.descripcion}</p>
    <p><strong>Responsable:</strong> ${actividad.responsable}</p>
    <p><strong>Prioridad:</strong> ${actividad.prioridad}</p>
    <p><strong>Fecha inicio:</strong> ${Utilities.formatDate(actividad.fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}</p>
    <p><strong>Fecha fin:</strong> ${Utilities.formatDate(actividad.fechaFin, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}</p>
    <p><strong>Estado:</strong> ${actividad.estado}</p>
  `;

  destinatarios.forEach(email => {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpo
    });
  });
}

/**
 * Envía notificación de actualización
 */
function enviarNotificacionActualizacion(actividadId, cambios) {
  // Implementación similar a enviarNotificacionNuevaActividad
  // pero destacando los cambios realizados
}

/**
 * Envía notificación de nuevo comentario
 */
function enviarNotificacionComentario(actividadId, usuario, comentario) {
  // Implementación para notificar nuevos comentarios
}

// ============================================
// HISTORIAL Y AUDITORÍA
// ============================================

/**
 * Registra acción en el historial
 */
function registrarHistorial(actividadId, usuario, accion, detalle) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Historial');

  const id = 'HIST-' + new Date().getTime();

  sheet.appendRow([
    id,
    actividadId,
    usuario,
    accion,
    detalle,
    new Date()
  ]);
}

/**
 * Obtiene historial de una actividad
 */
function obtenerHistorial(actividadId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Historial');
  const data = sheet.getDataRange().getValues();

  const historial = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === actividadId) {
      historial.push({
        id: data[i][0],
        actividadId: data[i][1],
        usuario: data[i][2],
        accion: data[i][3],
        detalle: data[i][4],
        fecha: data[i][5]
      });
    }
  }

  return historial;
}

// ============================================
// REPORTES Y ESTADÍSTICAS
// ============================================

/**
 * Genera reporte de actividades
 */
function generarReporte(filtros) {
  const usuario = obtenerUsuarioActual();

  if (usuario.rol === CONFIG.ROLES.COLABORADOR) {
    throw new Error('No tienes permisos para generar reportes');
  }

  const actividades = obtenerActividades(filtros);

  const estadisticas = {
    total: actividades.length,
    porEstado: {},
    porPrioridad: {},
    porResponsable: {}
  };

  actividades.forEach(act => {
    // Por estado
    estadisticas.porEstado[act.estado] = (estadisticas.porEstado[act.estado] || 0) + 1;

    // Por prioridad
    estadisticas.porPrioridad[act.prioridad] = (estadisticas.porPrioridad[act.prioridad] || 0) + 1;

    // Por responsable
    estadisticas.porResponsable[act.responsable] = (estadisticas.porResponsable[act.responsable] || 0) + 1;
  });

  return {
    actividades: actividades,
    estadisticas: estadisticas
  };
}

/**
 * Función para inicializar el sistema
 * Ejecuta esta función manualmente la primera vez
 */
function inicializarSistema() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // Crear hoja Usuarios
    let sheet = ss.getSheetByName('Usuarios');
    if (!sheet) {
      sheet = ss.insertSheet('Usuarios');
      sheet.appendRow(['Email', 'Nombre', 'Rol', 'Estado', 'FechaRegistro']);
      sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
    }

    // Crear hoja Actividades
    sheet = ss.getSheetByName('Actividades');
    if (!sheet) {
      sheet = ss.insertSheet('Actividades');
      sheet.appendRow([
        'ID', 'ProyectoID', 'Título', 'Descripción', 'Responsable', 'Jefe', 'Gerente',
        'Estado', 'Prioridad', 'FechaInicio', 'FechaFin', 'FechaCreacion', 'UltimaActualizacion'
      ]);
      sheet.getRange('A1:M1').setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Crear hoja Comentarios
    sheet = ss.getSheetByName('Comentarios');
    if (!sheet) {
      sheet = ss.insertSheet('Comentarios');
      sheet.appendRow(['ID', 'ActividadID', 'Usuario', 'Comentario', 'Fecha']);
      sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
    }

    // Crear hoja Historial
    sheet = ss.getSheetByName('Historial');
    if (!sheet) {
      sheet = ss.insertSheet('Historial');
      sheet.appendRow(['ID', 'ActividadID', 'Usuario', 'Accion', 'Detalle', 'Fecha']);
      sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#2c3e50').setFontColor('#ffffff');
    }

    // Eliminar hoja por defecto si existe
    const defaultSheet = ss.getSheetByName('Hoja 1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }

    Logger.log('✅ Sistema inicializado correctamente');
    return { success: true, message: 'Sistema inicializado correctamente' };

  } catch (error) {
    Logger.log('❌ Error al inicializar: ' + error.toString());
    throw new Error('Error al inicializar: ' + error.toString());
  }
}

/**
 * Función de prueba para verificar la conexión
 */
function probarConexion() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('✅ Conexión exitosa con la hoja: ' + ss.getName());

    const sheets = ss.getSheets();
    Logger.log('📋 Pestañas encontradas:');
    sheets.forEach(sheet => {
      Logger.log('  - ' + sheet.getName());
    });

    return {
      success: true,
      nombreHoja: ss.getName(),
      pestanas: sheets.map(s => s.getName())
    };

  } catch (error) {
    Logger.log('❌ Error de conexión: ' + error.toString());
    throw new Error('Error de conexión: ' + error.toString());
  }
}

/**
 * Crea un usuario de prueba con rol Gerencia
 */
function crearUsuarioGerencia() {
  try {
    const email = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Usuarios');

    // Verificar si ya existe
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        // Actualizar a Gerencia
        sheet.getRange(i + 1, 3).setValue('Gerencia');
        Logger.log('✅ Usuario actualizado a Gerencia: ' + email);
        return { success: true, message: 'Usuario actualizado a Gerencia' };
      }
    }

    // Si no existe, crear nuevo
    sheet.appendRow([
      email,
      email.split('@')[0],
      'Gerencia',
      'Activo',
      new Date()
    ]);

    Logger.log('✅ Usuario Gerencia creado: ' + email);
    return { success: true, message: 'Usuario Gerencia creado' };

  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    throw new Error('Error: ' + error.toString());
  }
}

/**
 * Función para verificar permisos
 */
function verificarPermisos() {
  try {
    // Probar acceso a Spreadsheet
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('✅ Acceso a Spreadsheet OK');

    // Probar acceso a Calendar
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    Logger.log('✅ Acceso a Calendar OK');

    // Probar acceso a Email
    const email = Session.getActiveUser().getEmail();
    Logger.log('✅ Email del usuario: ' + email);

    // Probar acceso a Mail
    MailApp.getRemainingDailyQuota();
    Logger.log('✅ Acceso a Mail OK');

    return { success: true, message: 'Todos los permisos OK' };

  } catch (error) {
    Logger.log('❌ Error de permisos: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Función de prueba directa
 */
function probarObtenerUsuario() {
  try {
    Logger.log('╔════════════════════════════════════════╗');
    Logger.log('║   PRUEBA DE OBTENER USUARIO ACTUAL    ║');
    Logger.log('╚════════════════════════════════════════╝');

    const usuario = obtenerUsuarioActual();

    Logger.log('\n📊 RESULTADO:');
    Logger.log('Email: ' + usuario.email);
    Logger.log('Nombre: ' + usuario.nombre);
    Logger.log('Rol: ' + usuario.rol);
    Logger.log('Estado: ' + usuario.estado);
    Logger.log('Fecha Registro: ' + usuario.fechaRegistro);

    Logger.log('\n✅ PRUEBA EXITOSA');

    return usuario;

  } catch (error) {
    Logger.log('\n❌ PRUEBA FALLIDA');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);

    return null;
  }
}

/**
 * Lista todos los archivos del proyecto
 */
function listarArchivosProyecto() {
  try {
    Logger.log('=== ARCHIVOS DEL PROYECTO ===');

    // Obtener el proyecto actual
    const project = ScriptApp.getProjectId();
    Logger.log('ID del proyecto: ' + project);

    // Listar archivos de código
    Logger.log('\nArchivos .gs encontrados:');
    // No hay API directa para listar, pero podemos intentar cargar cada uno

    Logger.log('\nIntentando cargar archivos HTML:');

    try {
      const login = HtmlService.createHtmlOutputFromFile('login');
      Logger.log('✅ login.html - EXISTE');
    } catch (e) {
      Logger.log('❌ login.html - NO EXISTE o tiene errores');
      Logger.log('   Error: ' + e.message);
    }

    try {
      const dashboard = HtmlService.createHtmlOutputFromFile('dashboard');
      Logger.log('✅ dashboard.html - EXISTE');
    } catch (e) {
      Logger.log('❌ dashboard.html - NO EXISTE o tiene errores');
      Logger.log('   Error: ' + e.message);
    }

    Logger.log('\n=== FIN LISTA DE ARCHIVOS ===');

  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
  }
}

/**
 * Diagnóstico completo del sistema
 */
function diagnosticoCompleto() {
  Logger.log('╔════════════════════════════════════════════════════╗');
  Logger.log('║        DIAGNÓSTICO COMPLETO DEL SISTEMA           ║');
  Logger.log('╚════════════════════════════════════════════════════╝');

  let errores = [];
  let advertencias = [];

  // 1. Verificar configuración
  Logger.log('\n1️⃣ VERIFICANDO CONFIGURACIÓN...');
  Logger.log('SPREADSHEET_ID: ' + CONFIG.SPREADSHEET_ID);
  Logger.log('CALENDAR_ID: ' + CONFIG.CALENDAR_ID);

  // 2. Verificar acceso a Spreadsheet
  Logger.log('\n2️⃣ VERIFICANDO ACCESO A SPREADSHEET...');
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    Logger.log('✅ Spreadsheet accesible: ' + ss.getName());
    Logger.log('   URL: ' + ss.getUrl());

    // Verificar hojas
    Logger.log('\n   Verificando hojas:');
    const hojasRequeridas = ['Usuarios', 'Actividades', 'Comentarios', 'Historial'];
    const hojasExistentes = ss.getSheets().map(s => s.getName());

    hojasRequeridas.forEach(nombre => {
      if (hojasExistentes.includes(nombre)) {
        const hoja = ss.getSheetByName(nombre);
        const filas = hoja.getLastRow();
        Logger.log('   ✅ ' + nombre + ' - ' + filas + ' filas');
      } else {
        Logger.log('   ❌ ' + nombre + ' - NO EXISTE');
        errores.push('Falta la hoja: ' + nombre);
      }
    });

  } catch (error) {
    Logger.log('❌ Error al acceder al Spreadsheet');
    Logger.log('   ' + error.toString());
    errores.push('No se puede acceder al Spreadsheet: ' + error.message);
  }

  // 3. Verificar usuario activo
  Logger.log('\n3️⃣ VERIFICANDO USUARIO ACTIVO...');
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      Logger.log('✅ Email del usuario: ' + email);
    } else {
      Logger.log('❌ No se pudo obtener el email del usuario');
      errores.push('No se puede obtener el email del usuario');
    }
  } catch (error) {
    Logger.log('❌ Error al obtener usuario activo');
    Logger.log('   ' + error.toString());
    errores.push('Error al obtener usuario: ' + error.message);
  }

  // 4. Verificar acceso a Calendar
  Logger.log('\n4️⃣ VERIFICANDO ACCESO A CALENDAR...');
  try {
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    Logger.log('✅ Calendario accesible: ' + cal.getName());
  } catch (error) {
    Logger.log('⚠️ Advertencia con Calendar');
    Logger.log('   ' + error.toString());
    advertencias.push('Problema con Calendar: ' + error.message);
  }

  // 5. Verificar archivos HTML
  Logger.log('\n5️⃣ VERIFICANDO ARCHIVOS HTML...');
  try {
    HtmlService.createHtmlOutputFromFile('login');
    Logger.log('✅ login.html existe');
  } catch (error) {
    Logger.log('❌ login.html no existe o tiene errores');
    Logger.log('   ' + error.toString());
    errores.push('Problema con login.html: ' + error.message);
  }

  try {
    HtmlService.createHtmlOutputFromFile('dashboard');
    Logger.log('✅ dashboard.html existe');
  } catch (error) {
    Logger.log('❌ dashboard.html no existe o tiene errores');
    Logger.log('   ' + error.toString());
    errores.push('Problema con dashboard.html: ' + error.message);
  }

  // 6. Probar obtenerUsuarioActual
  Logger.log('\n6️⃣ PROBANDO obtenerUsuarioActual()...');
  try {
    const usuario = obtenerUsuarioActual();

    if (usuario) {
      Logger.log('✅ Usuario obtenido correctamente:');
      Logger.log('   Email: ' + usuario.email);
      Logger.log('   Nombre: ' + usuario.nombre);
      Logger.log('   Rol: ' + usuario.rol);
      Logger.log('   Estado: ' + usuario.estado);
    } else {
      Logger.log('❌ obtenerUsuarioActual() devolvió null');
      errores.push('obtenerUsuarioActual() devuelve null');
    }
  } catch (error) {
    Logger.log('❌ Error en obtenerUsuarioActual()');
    Logger.log('   ' + error.toString());
    Logger.log('   Stack: ' + error.stack);
    errores.push('Error en obtenerUsuarioActual(): ' + error.message);
  }

  // RESUMEN
  Logger.log('\n╔════════════════════════════════════════════════════╗');
  Logger.log('║                    RESUMEN                         ║');
  Logger.log('╚════════════════════════════════════════════════════╝');

  if (errores.length === 0 && advertencias.length === 0) {
    Logger.log('✅ TODO FUNCIONA CORRECTAMENTE');
  } else {
    if (errores.length > 0) {
      Logger.log('\n❌ ERRORES ENCONTRADOS:');
      errores.forEach((error, i) => {
        Logger.log('   ' + (i + 1) + '. ' + error);
      });
    }

    if (advertencias.length > 0) {
      Logger.log('\n⚠️ ADVERTENCIAS:');
      advertencias.forEach((adv, i) => {
        Logger.log('   ' + (i + 1) + '. ' + adv);
      });
    }
  }

  Logger.log('\n════════════════════════════════════════════════════');

  return {
    errores: errores,
    advertencias: advertencias,
    exito: errores.length === 0
  };
}

/**
 * Wrapper específico para el login - VERSIÓN CORREGIDA
 * Esta función está diseñada para ser llamada desde el frontend
 */
function loginUsuario() {
  try {
    Logger.log('=== INICIO loginUsuario ===');
    
    const usuario = obtenerUsuarioActual();
    
    if (!usuario) {
      throw new Error('No se pudo obtener el usuario');
    }
    
    const usuarioSerializable = {
      email: String(usuario.email || ''),
      nombre: String(usuario.nombre || ''),
      rol: String(usuario.rol || ''),
      estado: String(usuario.estado || ''),
      fechaRegistro: usuario.fechaRegistro ? new Date(usuario.fechaRegistro).toISOString() : null
    };
    
    Logger.log('Retornando: ' + JSON.stringify(usuarioSerializable));
    Logger.log('=== FIN loginUsuario ===');
    
    return usuarioSerializable;
    
  } catch (error) {
    Logger.log('ERROR en loginUsuario: ' + error.toString());
    throw error;
  }
}



/**
 * Función de test simple para verificar comunicación frontend-backend
 */
function testConexion() {
  return {
    success: true,
    mensaje: "Conexión exitosa",
    timestamp: new Date().toString(),
    email: Session.getActiveUser().getEmail()
  };
}

function getAppUrl() { return ScriptApp.getService().getUrl(); }

function getDashboardHtml() { 
  return HtmlService.createTemplateFromFile('Dashboard').evaluate().getContent(); 
}

function obtenerActividadesDashboard() {
  try {
    const actividades = obtenerActividades(); // tu función ya existe
    return actividades || [];
  } catch (error) {
    Logger.log("Error en obtenerActividadesDashboard: " + error);
    return [];
  }
}

/**
 * Crear usuarios de prueba para todos los roles
 */
function crearUsuariosPrueba() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Usuarios');
  
  const usuariosPrueba = [
    ['colaborador1@test.com', 'Juan Pérez', 'Colaborador', 'Activo', new Date()],
    ['colaborador2@test.com', 'María García', 'Colaborador', 'Activo', new Date()],
    ['jefe1@test.com', 'Carlos López', 'Jefatura', 'Activo', new Date()],
    ['jefe2@test.com', 'Ana Martínez', 'Jefatura', 'Activo', new Date()],
    ['gerente1@test.com', 'Roberto Sánchez', 'Gerencia', 'Activo', new Date()]
  ];
  
  usuariosPrueba.forEach(usuario => {
    sheet.appendRow(usuario);
  });
  
  Logger.log('✅ Usuarios de prueba creados');
  
  return { success: true, message: 'Se crearon ' + usuariosPrueba.length + ' usuarios de prueba' };
}

function verTodosLosUsuarios() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Usuarios');
  
  if (!sheet) {
    Logger.log('❌ No existe la hoja Usuarios');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  Logger.log('╔════════════════════════════════════════════════════╗');
  Logger.log('║           USUARIOS EN LA HOJA                      ║');
  Logger.log('╚════════════════════════════════════════════════════╝');
  Logger.log('');
  Logger.log('Total de filas: ' + data.length);
  Logger.log('');
  
  // Mostrar encabezados
  Logger.log('ENCABEZADOS:');
  Logger.log(data[0].join(' | '));
  Logger.log('─'.repeat(80));
  
  // Mostrar datos
  for (let i = 1; i < data.length; i++) {
    Logger.log('Fila ' + i + ':');
    Logger.log('  Email: ' + data[i][0]);
    Logger.log('  Nombre: ' + data[i][1]);
    Logger.log('  Rol: ' + data[i][2]);
    Logger.log('  Estado: ' + data[i][3]);
    Logger.log('  Fecha: ' + data[i][4]);
    Logger.log('');
  }
  
  // Contar por rol
  const conteoRoles = {};
  for (let i = 1; i < data.length; i++) {
    const rol = data[i][2];
    conteoRoles[rol] = (conteoRoles[rol] || 0) + 1;
  }
  
  Logger.log('════════════════════════════════════════════════════');
  Logger.log('RESUMEN POR ROL:');
  Object.keys(conteoRoles).forEach(rol => {
    Logger.log('  ' + rol + ': ' + conteoRoles[rol] + ' usuario(s)');
  });
  Logger.log('════════════════════════════════════════════════════');
}

// ============================================
// GESTIÓN DE PROYECTOS
// ============================================

/**
 * Inicializa la hoja de Proyectos
 */
function inicializarHojaProyectos() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Proyectos');
    
    if (!sheet) {
      Logger.log('Creando hoja Proyectos...');
      sheet = ss.insertSheet('Proyectos');
      
      // Encabezados
      sheet.appendRow([
        'ID',
        'Nombre',
        'Descripción',
        'GerenteResponsable',
        'JefeResponsable',
        'Colaboradores',
        'Estado',
        'Prioridad',
        'FechaInicio',
        'FechaFin',
        'FechaCreacion',
        'UltimaActualizacion',
        'PresupuestoEstimado',
        'AvanceGeneral'
      ]);
      
      // Formato de encabezados
      sheet.getRange('A1:N1')
        .setFontWeight('bold')
        .setBackground('#2c3e50')
        .setFontColor('#ffffff');
      
      sheet.setFrozenRows(1);
      
      Logger.log('✅ Hoja Proyectos creada');
    } else {
      Logger.log('Hoja Proyectos ya existe');
    }
    
    return { success: true, message: 'Hoja Proyectos lista' };
    
  } catch (error) {
    Logger.log('❌ Error al inicializar hoja Proyectos: ' + error.toString());
    throw error;
  }
}

/**
 * Actualiza la hoja Actividades para agregar columna ProyectoID
 */
function actualizarHojaActividadesParaProyectos() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    
    if (!sheet) {
      throw new Error('La hoja Actividades no existe');
    }
    
    // Verificar si ya existe la columna ProyectoID
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    if (headers.includes('ProyectoID')) {
      Logger.log('La columna ProyectoID ya existe en Actividades');
      return { success: true, message: 'Columna ProyectoID ya existe' };
    }
    
    // Insertar nueva columna después de ID (columna B)
    sheet.insertColumnAfter(1);
    sheet.getRange(1, 2).setValue('ProyectoID');
    
    // Aplicar formato al encabezado
    sheet.getRange(1, 2)
      .setFontWeight('bold')
      .setBackground('#2c3e50')
      .setFontColor('#ffffff');
    
    Logger.log('✅ Columna ProyectoID agregada a Actividades');
    
    return { success: true, message: 'Columna ProyectoID agregada' };
    
  } catch (error) {
    Logger.log('❌ Error al actualizar hoja Actividades: ' + error.toString());
    throw error;
  }
}

/**
 * Configuración de estados y prioridades para proyectos
 */
const CONFIG_PROYECTOS = {
  ESTADOS: {
    PLANIFICACION: 'Planificación',
    EN_CURSO: 'En Curso',
    PAUSADO: 'Pausado',
    COMPLETADO: 'Completado',
    CANCELADO: 'Cancelado'
  },
  PRIORIDADES: {
    CRITICA: 'Crítica',
    ALTA: 'Alta',
    MEDIA: 'Media',
    BAJA: 'Baja'
  }
};

/**
 * Crea un nuevo proyecto
 */
function crearProyecto(datos) {
  try {
    const usuario = obtenerUsuarioActual();
    
    // Validar permisos (solo Gerencia y Jefatura)
    if (usuario.rol === CONFIG.ROLES.COLABORADOR) {
      throw new Error('No tienes permisos para crear proyectos');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Proyectos');
    
    if (!sheet) {
      throw new Error('La hoja Proyectos no existe. Ejecuta inicializarHojaProyectos()');
    }
    
    // Generar ID único
    const id = 'PROY-' + new Date().getTime();
    
    // Preparar datos
    const ahora = new Date();
    const colaboradoresStr = Array.isArray(datos.colaboradores) 
      ? datos.colaboradores.join(',') 
      : datos.colaboradores || '';
    
    const proyecto = {
      id: id,
      nombre: datos.nombre,
      descripcion: datos.descripcion || '',
      gerenteResponsable: datos.gerenteResponsable || '',
      jefeResponsable: datos.jefeResponsable || '',
      colaboradores: colaboradoresStr,
      estado: CONFIG_PROYECTOS.ESTADOS.PLANIFICACION,
      prioridad: datos.prioridad || CONFIG_PROYECTOS.PRIORIDADES.MEDIA,
      fechaInicio: new Date(datos.fechaInicio),
      fechaFin: new Date(datos.fechaFin),
      fechaCreacion: ahora,
      ultimaActualizacion: ahora,
      presupuestoEstimado: datos.presupuestoEstimado || 0,
      avanceGeneral: 0
    };
    
    // Guardar en Sheets
    sheet.appendRow([
      proyecto.id,
      proyecto.nombre,
      proyecto.descripcion,
      proyecto.gerenteResponsable,
      proyecto.jefeResponsable,
      proyecto.colaboradores,
      proyecto.estado,
      proyecto.prioridad,
      proyecto.fechaInicio,
      proyecto.fechaFin,
      proyecto.fechaCreacion,
      proyecto.ultimaActualizacion,
      proyecto.presupuestoEstimado,
      proyecto.avanceGeneral
    ]);
    
    // Registrar en historial
    registrarHistorial(proyecto.id, usuario.email, 'Creación Proyecto', 'Proyecto creado: ' + proyecto.nombre);
    
    // Enviar notificaciones
    enviarNotificacionNuevoProyecto(proyecto);
    
    Logger.log('✅ Proyecto creado: ' + proyecto.id);
    
    // Serializar fechas: google.script.run no puede enviar objetos Date nativos.
    const proyectoSerializable = {
      id: proyecto.id,
      nombre: proyecto.nombre,
      descripcion: proyecto.descripcion,
      gerenteResponsable: proyecto.gerenteResponsable,
      jefeResponsable: proyecto.jefeResponsable,
      colaboradores: proyecto.colaboradores,
      estado: proyecto.estado,
      prioridad: proyecto.prioridad,
      fechaInicio: proyecto.fechaInicio instanceof Date ? proyecto.fechaInicio.toISOString() : (proyecto.fechaInicio || ''),
      fechaFin: proyecto.fechaFin instanceof Date ? proyecto.fechaFin.toISOString() : (proyecto.fechaFin || ''),
      fechaCreacion: proyecto.fechaCreacion instanceof Date ? proyecto.fechaCreacion.toISOString() : (proyecto.fechaCreacion || ''),
      ultimaActualizacion: proyecto.ultimaActualizacion instanceof Date ? proyecto.ultimaActualizacion.toISOString() : (proyecto.ultimaActualizacion || ''),
      presupuestoEstimado: proyecto.presupuestoEstimado,
      avanceGeneral: proyecto.avanceGeneral
    };
    
    return { success: true, proyecto: proyectoSerializable };
    
  } catch (error) {
    Logger.log('❌ Error al crear proyecto: ' + error.toString());
    throw error;
  }
}

/**
 * Visibilidad:
 * - Gerencia:    ve todos
 * - Jefatura:    ve los suyos (jefeResponsable o gerenteResponsable) o donde es colaborador
 * - Colaborador: ve los proyectos donde está asignado
 */
function obtenerProyectos(filtros = {}) {
  try {
    const usuario = obtenerUsuarioActual();
    const ss      = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet   = ss.getSheetByName('Proyectos');
    if (!sheet) return [];

    const data      = sheet.getDataRange().getValues();
    const proyectos = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      const colaboradores = data[i][5]
        ? String(data[i][5]).split(',').map(e => e.trim()).filter(Boolean)
        : [];

      const proyecto = {
        id:                  data[i][0],
        nombre:              data[i][1],
        descripcion:         data[i][2],
        gerenteResponsable:  data[i][3],
        jefeResponsable:     data[i][4],
        colaboradores,
        estado:              data[i][6],
        prioridad:           data[i][7],
        fechaInicio:         data[i][8]  instanceof Date ? data[i][8].toISOString()  : (data[i][8]  || ''),
        fechaFin:            data[i][9]  instanceof Date ? data[i][9].toISOString()  : (data[i][9]  || ''),
        fechaCreacion:       data[i][10] instanceof Date ? data[i][10].toISOString() : (data[i][10] || ''),
        ultimaActualizacion: data[i][11] instanceof Date ? data[i][11].toISOString() : (data[i][11] || ''),
        presupuestoEstimado: data[i][12],
        avanceGeneral:       data[i][13]
      };

      let incluir = false;

      switch (usuario.rol) {
        case CONFIG.ROLES.GERENCIA:
          incluir = true;
          break;
        case CONFIG.ROLES.JEFATURA:
          incluir = proyecto.jefeResponsable    === usuario.email ||
                    proyecto.gerenteResponsable === usuario.email ||
                    colaboradores.includes(usuario.email);
          break;
        case CONFIG.ROLES.COLABORADOR:
          incluir = colaboradores.includes(usuario.email);
          break;
      }

      if (incluir && filtros.estado    && proyecto.estado    !== filtros.estado)    incluir = false;
      if (incluir && filtros.prioridad && proyecto.prioridad !== filtros.prioridad) incluir = false;

      if (incluir) proyectos.push(proyecto);
    }

    Logger.log('Proyectos encontrados: ' + proyectos.length);
    return proyectos;

  } catch (error) {
    Logger.log('ERROR en obtenerProyectos: ' + error.toString());
    return [];
  }
}

/**
 * Actualiza un proyecto
 */
function actualizarProyecto(id, cambios) {
  try {
    const usuario = obtenerUsuarioActual();
    
    // Validar permisos
    if (usuario.rol === CONFIG.ROLES.COLABORADOR) {
      throw new Error('No tienes permisos para actualizar proyectos');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Proyectos');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        const fila = i + 1;
        
        // Actualizar campos según lo que se haya cambiado
        if (cambios.nombre !== undefined) sheet.getRange(fila, 2).setValue(cambios.nombre);
        if (cambios.descripcion !== undefined) sheet.getRange(fila, 3).setValue(cambios.descripcion);
        if (cambios.gerenteResponsable !== undefined) sheet.getRange(fila, 4).setValue(cambios.gerenteResponsable);
        if (cambios.jefeResponsable !== undefined) sheet.getRange(fila, 5).setValue(cambios.jefeResponsable);
        if (cambios.colaboradores !== undefined) {
          const colaboradoresStr = Array.isArray(cambios.colaboradores) 
            ? cambios.colaboradores.join(',') 
            : cambios.colaboradores;
          sheet.getRange(fila, 6).setValue(colaboradoresStr);
        }
        if (cambios.estado !== undefined) sheet.getRange(fila, 7).setValue(cambios.estado);
        if (cambios.prioridad !== undefined) sheet.getRange(fila, 8).setValue(cambios.prioridad);
        if (cambios.fechaInicio !== undefined) sheet.getRange(fila, 9).setValue(new Date(cambios.fechaInicio));
        if (cambios.fechaFin !== undefined) sheet.getRange(fila, 10).setValue(new Date(cambios.fechaFin));
        if (cambios.presupuestoEstimado !== undefined) sheet.getRange(fila, 13).setValue(cambios.presupuestoEstimado);
        if (cambios.avanceGeneral !== undefined) sheet.getRange(fila, 14).setValue(cambios.avanceGeneral);
        
        // Actualizar timestamp
        sheet.getRange(fila, 12).setValue(new Date());
        
        // Registrar en historial
        const detalles = Object.keys(cambios).map(key => `${key}: ${cambios[key]}`).join(', ');
        registrarHistorial(id, usuario.email, 'Actualización Proyecto', detalles);
        
        Logger.log('✅ Proyecto actualizado: ' + id);
        
        return { success: true, message: 'Proyecto actualizado' };
      }
    }
    
    throw new Error('Proyecto no encontrado');
    
  } catch (error) {
    Logger.log('❌ Error al actualizar proyecto: ' + error.toString());
    throw error;
  }
}

/**
 * Elimina un proyecto (solo Gerencia)
 */
function eliminarProyecto(id) {
  try {
    const usuario = obtenerUsuarioActual();
    
    if (usuario.rol !== CONFIG.ROLES.GERENCIA) {
      throw new Error('Solo Gerencia puede eliminar proyectos');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Proyectos');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        registrarHistorial(id, usuario.email, 'Eliminación Proyecto', 'Proyecto eliminado');
        
        Logger.log('✅ Proyecto eliminado: ' + id);
        
        return { success: true, message: 'Proyecto eliminado' };
      }
    }
    
    throw new Error('Proyecto no encontrado');
    
  } catch (error) {
    Logger.log('❌ Error al eliminar proyecto: ' + error.toString());
    throw error;
  }
}

/**
 * Obtiene actividades de un proyecto específico
 */
function obtenerActividadesDeProyecto(proyectoId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    const data = sheet.getDataRange().getValues();
    
    const actividades = [];
    
    // Buscar la columna ProyectoID
    const headers = data[0];
    const proyectoIdIndex = headers.indexOf('ProyectoID');
    
    if (proyectoIdIndex === -1) {
      Logger.log('La columna ProyectoID no existe en Actividades');
      return [];
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][proyectoIdIndex] === proyectoId) {
        // Construir objeto actividad según tu estructura actual
        const actividad = {
          id: data[i][0],
          proyectoId: data[i][proyectoIdIndex],
          titulo: data[i][2],
          descripcion: data[i][3],
          responsable: data[i][4],
          estado: data[i][7],
          prioridad: data[i][8]
        };
        
        actividades.push(actividad);
      }
    }
    
    Logger.log('Actividades del proyecto ' + proyectoId + ': ' + actividades.length);
    return actividades;
    
  } catch (error) {
    Logger.log('❌ Error al obtener actividades del proyecto: ' + error.toString());
    return [];
  }
}

/**
 * Calcula la carga de trabajo de un colaborador
 */
function calcularCargaTrabajo(emailColaborador) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // Contar proyectos
    const sheetProyectos = ss.getSheetByName('Proyectos');
    const dataProyectos = sheetProyectos.getDataRange().getValues();
    
    let proyectosActivos = 0;
    const proyectosIds = [];
    
    for (let i = 1; i < dataProyectos.length; i++) {
      const colaboradores = dataProyectos[i][5] ? dataProyectos[i][5].split(',') : [];
      const estado = dataProyectos[i][6];
      
      if (colaboradores.includes(emailColaborador) && estado !== 'Completado' && estado !== 'Cancelado') {
        proyectosActivos++;
        proyectosIds.push(dataProyectos[i][0]);
      }
    }
    
    // Contar actividades
    const sheetActividades = ss.getSheetByName('Actividades');
    const dataActividades = sheetActividades.getDataRange().getValues();
    
    let actividadesPendientes = 0;
    let actividadesEnProceso = 0;
    
    for (let i = 1; i < dataActividades.length; i++) {
      const responsable = dataActividades[i][4]; // Ajusta el índice según tu estructura
      const estado = dataActividades[i][7]; // Ajusta el índice según tu estructura
      
      if (responsable === emailColaborador) {
        if (estado === 'Pendiente') actividadesPendientes++;
        if (estado === 'En Proceso') actividadesEnProceso++;
      }
    }
    
    const cargaTrabajo = {
      colaborador: emailColaborador,
      proyectosActivos: proyectosActivos,
      proyectosIds: proyectosIds,
      actividadesPendientes: actividadesPendientes,
      actividadesEnProceso: actividadesEnProceso,
      totalActividades: actividadesPendientes + actividadesEnProceso,
      nivelCarga: calcularNivelCarga(proyectosActivos, actividadesPendientes + actividadesEnProceso)
    };
    
    Logger.log('Carga de trabajo de ' + emailColaborador + ': ' + JSON.stringify(cargaTrabajo));
    
    return cargaTrabajo;
    
  } catch (error) {
    Logger.log('❌ Error al calcular carga de trabajo: ' + error.toString());
    throw error;
  }
}

/**
 * Calcula el nivel de carga (Baja, Media, Alta, Crítica)
 */
function calcularNivelCarga(proyectos, actividades) {
  const puntaje = (proyectos * 10) + (actividades * 2);
  
  if (puntaje >= 50) return 'Crítica';
  if (puntaje >= 30) return 'Alta';
  if (puntaje >= 15) return 'Media';
  return 'Baja';
}

/**
 * Obtiene la carga de trabajo de todos los colaboradores
 */
function obtenerCargaTrabajoEquipo() {
  try {
    const usuario = obtenerUsuarioActual();
    
    // Solo Gerencia y Jefatura pueden ver la carga del equipo
    if (usuario.rol === CONFIG.ROLES.COLABORADOR) {
      throw new Error('No tienes permisos para ver la carga del equipo');
    }
    
    const colaboradores = obtenerUsuariosPorRol('Colaborador');
    const cargasTrabajo = [];
    
    colaboradores.forEach(colab => {
      const carga = calcularCargaTrabajo(colab.email);
      carga.nombre = colab.nombre;
      cargasTrabajo.push(carga);
    });
    
    // Ordenar por nivel de carga (crítica primero)
    const ordenCarga = { 'Crítica': 4, 'Alta': 3, 'Media': 2, 'Baja': 1 };
    cargasTrabajo.sort((a, b) => ordenCarga[b.nivelCarga] - ordenCarga[a.nivelCarga]);
    
    Logger.log('Carga de trabajo del equipo calculada');
    
    return cargasTrabajo;
    
  } catch (error) {
    Logger.log('❌ Error al obtener carga del equipo: ' + error.toString());
    throw error;
  }
}

/**
 * Actualiza el avance general de un proyecto basado en sus actividades
 */
function actualizarAvanceProyecto(proyectoId) {
  try {
    const actividades = obtenerActividadesDeProyecto(proyectoId);
    
    if (actividades.length === 0) {
      return { success: true, avance: 0 };
    }
    
    const completadas = actividades.filter(a => a.estado === 'Completada').length;
    const avance = Math.round((completadas / actividades.length) * 100);
    
    // Actualizar en la hoja
    actualizarProyecto(proyectoId, { avanceGeneral: avance });
    
    Logger.log('✅ Avance del proyecto ' + proyectoId + ' actualizado: ' + avance + '%');
    
    return { success: true, avance: avance };
    
  } catch (error) {
    Logger.log('❌ Error al actualizar avance del proyecto: ' + error.toString());
    throw error;
  }
}

/**
 * Envía notificación de nuevo proyecto
 */
function enviarNotificacionNuevoProyecto(proyecto) {
  try {
    const destinatarios = [];
    
    if (proyecto.gerenteResponsable) destinatarios.push(proyecto.gerenteResponsable);
    if (proyecto.jefeResponsable) destinatarios.push(proyecto.jefeResponsable);
    if (proyecto.colaboradores) {
      proyecto.colaboradores.split(',').forEach(email => {
        if (email.trim()) destinatarios.push(email.trim());
      });
    }
    
    const asunto = `Nuevo proyecto: ${proyecto.nombre}`;
    const cuerpo = `
      <h2>Nuevo Proyecto Creado</h2>
      <p><strong>Nombre:</strong> ${proyecto.nombre}</p>
      <p><strong>Descripción:</strong> ${proyecto.descripcion}</p>
      <p><strong>Gerente:</strong> ${proyecto.gerenteResponsable || 'No asignado'}</p>
      <p><strong>Jefe:</strong> ${proyecto.jefeResponsable || 'No asignado'}</p>
      <p><strong>Prioridad:</strong> ${proyecto.prioridad}</p>
      <p><strong>Fecha inicio:</strong> ${Utilities.formatDate(proyecto.fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy')}</p>
      <p><strong>Fecha fin:</strong> ${Utilities.formatDate(proyecto.fechaFin, Session.getScriptTimeZone(), 'dd/MM/yyyy')}</p>
      <p><strong>Estado:</strong> ${proyecto.estado}</p>
    `;
    
    destinatarios.forEach(email => {
      if (email) {
        MailApp.sendEmail({
          to: email,
          subject: asunto,
          htmlBody: cuerpo
        });
      }
    });
    
    Logger.log('Notificaciones enviadas para proyecto: ' + proyecto.id);
    
  } catch (error) {
    Logger.log('Error al enviar notificaciones: ' + error.toString());
  }
}
/**
 * Obtiene los proyectos disponibles para asignar actividades según el rol del usuario
 */
function obtenerProyectosDisponibles() {
  try {
    Logger.log('=== INICIO obtenerProyectosDisponibles ===');
    
    const usuario = obtenerUsuarioActual();
    Logger.log('Usuario: ' + usuario.email + ' (' + usuario.rol + ')');
    
    const proyectos = obtenerProyectos();
    Logger.log('Total proyectos obtenidos: ' + proyectos.length);
    
    // CAMBIO AQUÍ: Filtrar excluyendo solo completados y cancelados
    const proyectosActivos = proyectos.filter(p => {
      const estado = p.estado || '';
      return estado !== 'Completado' && estado !== 'Cancelado';
    });
    
    Logger.log('Proyectos disponibles: ' + proyectosActivos.length);
    Logger.log('=== FIN obtenerProyectosDisponibles ===');
    
    return proyectosActivos;
    
  } catch (error) {
    Logger.log('❌ ERROR en obtenerProyectosDisponibles: ' + error.toString());
    return [];
  }
}

function calcularAvanceProyecto(proyectoId) {
  try {
    Logger.log('');
    Logger.log('═══ CALCULANDO AVANCE DEL PROYECTO: ' + proyectoId + ' ═══');
    
    if (!proyectoId || proyectoId === '' || proyectoId === null) {
      Logger.log('❌ proyectoId vacío, retornando 0');
      return 0;
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    
    if (!sheet) {
      Logger.log('❌ Hoja Actividades no existe');
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('Total de filas: ' + data.length);
    
    if (data.length < 2) {
      Logger.log('❌ No hay actividades');
      return 0;
    }
    
    // Encontrar actividades del proyecto
    const actividadesDelProyecto = [];
    
    for (let i = 1; i < data.length; i++) {
      const actProyectoId = data[i][1]; // Columna B = ProyectoID
      const actTitulo = data[i][2]; // Columna C = Título
      const actAvance = data[i][13]; // Columna N = Avance (índice 13)
      
      // Comparación estricta como strings
      if (String(actProyectoId) === String(proyectoId)) {
        const avanceNum = parseInt(actAvance) || 0;
        actividadesDelProyecto.push({
          titulo: actTitulo,
          avance: avanceNum
        });
        Logger.log('  ✅ "' + actTitulo + '" - Avance: ' + avanceNum + '%');
      }
    }
    
    Logger.log('Total actividades encontradas: ' + actividadesDelProyecto.length);
    
    if (actividadesDelProyecto.length === 0) {
      Logger.log('⚠️ No hay actividades para este proyecto');
      return 0;
    }
    
    // Calcular promedio
    const sumaAvances = actividadesDelProyecto.reduce((suma, act) => suma + act.avance, 0);
    const avancePromedio = Math.round(sumaAvances / actividadesDelProyecto.length);
    
    Logger.log('CÁLCULO: ' + sumaAvances + ' / ' + actividadesDelProyecto.length + ' = ' + avancePromedio + '%');
    Logger.log('═══════════════════════════════════════════');
    
    return avancePromedio;
    
  } catch (error) {
    Logger.log('❌ ERROR en calcularAvanceProyecto: ' + error.toString());
    return 0;
  }
}

function calcularAvanceProyecto(proyectoId) {
  try {
    Logger.log('=== Calculando avance del proyecto: ' + proyectoId);
    
    if (!proyectoId) {
      Logger.log('No hay proyecto ID, retornando 0');
      return 0;
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    
    if (!sheet) {
      Logger.log('Hoja Actividades no existe');
      return 0;
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      Logger.log('No hay actividades');
      return 0;
    }
    
    // Encontrar actividades del proyecto
    const actividadesDelProyecto = [];
    
    for (let i = 1; i < data.length; i++) {
      const actProyectoId = data[i][1]; // Columna B = ProyectoID
      const actAvance = data[i][13] || 0; // Columna N = Avance
      
      if (actProyectoId === proyectoId) {
        actividadesDelProyecto.push({
          avance: parseInt(actAvance) || 0
        });
      }
    }
    
    Logger.log('Actividades del proyecto encontradas: ' + actividadesDelProyecto.length);
    
    if (actividadesDelProyecto.length === 0) {
      Logger.log('No hay actividades para este proyecto, retornando 0');
      return 0;
    }
    
    // Calcular promedio de avance
    const sumaAvances = actividadesDelProyecto.reduce((suma, act) => suma + act.avance, 0);
    const avancePromedio = Math.round(sumaAvances / actividadesDelProyecto.length);
    
    Logger.log('Avance calculado: ' + avancePromedio + '%');
    Logger.log('(Suma: ' + sumaAvances + ' / Cantidad: ' + actividadesDelProyecto.length + ')');
    
    return avancePromedio;
    
  } catch (error) {
    Logger.log('❌ Error al calcular avance del proyecto: ' + error.toString());
    return 0;
  }
}

function actualizarAvanceEnProyecto(proyectoId) {
  try {
    Logger.log('');
    Logger.log('🔄 ACTUALIZANDO HOJA PROYECTOS para: ' + proyectoId);
    
    if (!proyectoId || proyectoId === '' || proyectoId === null) {
      Logger.log('❌ proyectoId vacío, abortando');
      return;
    }
    
    // Calcular el nuevo avance
    const avance = calcularAvanceProyecto(proyectoId);
    Logger.log('Avance calculado: ' + avance + '%');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Proyectos');
    
    if (!sheet) {
      Logger.log('❌ Hoja Proyectos no existe');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Buscar el proyecto
    let encontrado = false;
    
    for (let i = 1; i < data.length; i++) {
      const proyectoIdEnHoja = data[i][0]; // Columna A = ID
      
      if (String(proyectoIdEnHoja) === String(proyectoId)) {
        const fila = i + 1;
        const nombreProyecto = data[i][1]; // Columna B = Nombre
        
        Logger.log('✅ Proyecto encontrado: ' + nombreProyecto + ' (fila ' + fila + ')');
        
        // Actualizar columna N (índice 13, pero getRange usa 1-based = 14)
        sheet.getRange(fila, 14).setValue(avance);
        Logger.log('💾 Avance guardado en N' + fila + ': ' + avance + '%');
        
        // Actualizar timestamp (columna L = 12)
        sheet.getRange(fila, 12).setValue(new Date());
        Logger.log('🕐 Timestamp actualizado');
        
        Logger.log('✅ PROYECTO ACTUALIZADO: ' + nombreProyecto + ' → ' + avance + '%');
        
        encontrado = true;
        break;
      }
    }
    
    if (!encontrado) {
      Logger.log('⚠️ Proyecto NO encontrado en la hoja');
      Logger.log('ID buscado: "' + proyectoId + '"');
    }
    
  } catch (error) {
    Logger.log('❌ ERROR en actualizarAvanceEnProyecto: ' + error.toString());
  }
}

function recalcularTodosLosProyectos() {
  try {
    Logger.log('');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('RECALCULANDO TODOS LOS PROYECTOS');
    Logger.log('═══════════════════════════════════════════');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Proyectos');
    
    if (!sheet) {
      Logger.log('❌ Hoja Proyectos no existe');
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log('Total de proyectos: ' + (data.length - 1));
    
    let actualizados = 0;
    
    for (let i = 1; i < data.length; i++) {
      const proyectoId = data[i][0]; // Columna A = ID
      const nombreProyecto = data[i][1]; // Columna B = Nombre
      
      if (proyectoId && proyectoId !== '') {
        Logger.log('');
        Logger.log('─────────────────────────────────────────');
        Logger.log((i) + '. Proyecto: ' + nombreProyecto);
        Logger.log('   ID: ' + proyectoId);
        
        actualizarAvanceEnProyecto(proyectoId);
        actualizados++;
      }
    }
    
    Logger.log('');
    Logger.log('═══════════════════════════════════════════');
    Logger.log('✅ FINALIZADO: ' + actualizados + ' proyectos actualizados');
    Logger.log('═══════════════════════════════════════════');
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
  }
}

function debugEstructuraDatos() {
  Logger.log('═══════════════════════════════════════════');
  Logger.log('DEBUG: ESTRUCTURA DE DATOS');
  Logger.log('═══════════════════════════════════════════');
  
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // Ver Actividades
  Logger.log('\n📋 HOJA ACTIVIDADES:');
  const sheetAct = ss.getSheetByName('Actividades');
  const dataAct = sheetAct.getRange(1, 1, 2, 15).getValues(); // Primeras 2 filas
  Logger.log('Headers: ' + dataAct[0].join(' | '));
  if (dataAct.length > 1) {
    Logger.log('Ejemplo fila 2: ' + dataAct[1].join(' | '));
    Logger.log('ProyectoID (columna B): "' + dataAct[1][1] + '" (tipo: ' + typeof dataAct[1][1] + ')');
    Logger.log('Avance (columna N): "' + dataAct[1][13] + '" (tipo: ' + typeof dataAct[1][13] + ')');
  }
  
  // Ver Proyectos
  Logger.log('\n📁 HOJA PROYECTOS:');
  const sheetProy = ss.getSheetByName('Proyectos');
  const dataProy = sheetProy.getRange(1, 1, 2, 15).getValues(); // Primeras 2 filas
  Logger.log('Headers: ' + dataProy[0].join(' | '));
  if (dataProy.length > 1) {
    Logger.log('Ejemplo fila 2: ' + dataProy[1].join(' | '));
    Logger.log('ID (columna A): "' + dataProy[1][0] + '" (tipo: ' + typeof dataProy[1][0] + ')');
    Logger.log('AvanceGeneral (columna N): "' + dataProy[1][13] + '" (tipo: ' + typeof dataProy[1][13] + ')');
  }
  
  Logger.log('═══════════════════════════════════════════');
}

// ===========================================================
// NUEVA FUNCIÓN: Calcular Carga de Trabajo Detallada
// ===========================================================

function calcularCargaTrabajoDetallada(emailColaborador) {
  try {
    Logger.log('Calculando carga para: ' + emailColaborador);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    const data = sheet.getDataRange().getValues();
    
    let actividadesPendientes = 0;
    let actividadesEnProceso = 0;
    let actividadesEmergentes = 0;
    let horasEstimadas = 0;
    
    for (let i = 1; i < data.length; i++) {
      const responsable = data[i][4]; // Columna E
      const estado = data[i][7]; // Columna H
      const tipo = data[i][14] || 'Recurrente'; // Columna O (nueva)
      const fechaFin = data[i][10]; // Columna K
      
      if (responsable === emailColaborador) {
        // Contar por estado
        if (estado === 'Pendiente' || estado === 'Pendiente de Aceptación') {
          actividadesPendientes++;
        }
        if (estado === 'En Proceso' || estado === 'En Ejecución' || estado === 'Aceptada') {
          actividadesEnProceso++;
        }
        if (tipo === 'Emergente' && estado !== 'Completada') {
          actividadesEmergentes++;
        }
        
        // Calcular horas pendientes (aproximado)
        if (estado !== 'Completada' && fechaFin) {
          const ahora = new Date();
          const fin = new Date(fechaFin);
          const horasPendientes = Math.max(0, (fin - ahora) / (1000 * 60 * 60));
          horasEstimadas += horasPendientes;
        }
      }
    }
    
    // Calcular puntuación de carga (0-100)
    // Fórmula: Emergentes pesan más, luego en proceso, luego pendientes
    const puntajeCarga = (actividadesEmergentes * 30) + 
                         (actividadesEnProceso * 15) + 
                         (actividadesPendientes * 5);
    
    // Determinar nivel y color
    let nivel, color, disponibilidad;
    if (puntajeCarga <= 30) {
      nivel = 'Baja';
      color = 'verde';
      disponibilidad = 'Disponible';
    } else if (puntajeCarga <= 60) {
      nivel = 'Media';
      color = 'amarillo';
      disponibilidad = 'Ocupado';
    } else {
      nivel = 'Alta';
      color = 'rojo';
      disponibilidad = 'Sobrecargado';
    }
    
    return {
      colaborador: emailColaborador,
      actividadesPendientes: actividadesPendientes,
      actividadesEnProceso: actividadesEnProceso,
      actividadesEmergentes: actividadesEmergentes,
      totalActividades: actividadesPendientes + actividadesEnProceso,
      horasEstimadas: Math.round(horasEstimadas),
      puntajeCarga: puntajeCarga,
      nivel: nivel,
      color: color,
      disponibilidad: disponibilidad
    };
    
  } catch (error) {
    Logger.log('Error calculando carga: ' + error.toString());
    return {
      colaborador: emailColaborador,
      actividadesPendientes: 0,
      actividadesEnProceso: 0,
      actividadesEmergentes: 0,
      totalActividades: 0,
      horasEstimadas: 0,
      puntajeCarga: 0,
      nivel: 'Desconocida',
      color: 'gris',
      disponibilidad: 'Desconocido'
    };
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Obtener Colaboradores con Carga
// ===========================================================

function obtenerColaboradoresConCarga() {
  try {
    Logger.log('=== Obteniendo colaboradores con análisis de carga ===');
    
    const colaboradores = obtenerUsuariosPorRol('Colaborador');
    const colaboradoresConCarga = [];
    
    colaboradores.forEach(colab => {
      const carga = calcularCargaTrabajoDetallada(colab.email);
      colaboradoresConCarga.push({
        email: colab.email,
        nombre: colab.nombre,
        carga: carga
      });
    });
    
    // Ordenar por puntaje de carga (menor primero = más disponible)
    colaboradoresConCarga.sort((a, b) => a.carga.puntajeCarga - b.carga.puntajeCarga);
    
    Logger.log('Colaboradores analizados: ' + colaboradoresConCarga.length);
    
    return colaboradoresConCarga;
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return [];
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Crear Actividad Emergente
// ===========================================================

function crearActividadEmergente(datos) {
  try {
    Logger.log('=== Creando actividad emergente ===');
    Logger.log('Datos: ' + JSON.stringify(datos));
    
    const usuario = obtenerUsuarioActual();
    
    // Solo Gerencia puede crear emergentes
    if (usuario.rol !== CONFIG.ROLES.GERENCIA) {
      throw new Error('Solo Gerencia puede crear actividades emergentes');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Actividades');
    
    if (!sheet) {
      sheet = ss.insertSheet('Actividades');
      // Headers con columnas nuevas
      sheet.appendRow([
        'ID', 'ProyectoID', 'Título', 'Descripción', 'Responsable', 
        'Jefe', 'Gerente', 'Estado', 'Prioridad', 'FechaInicio', 
        'FechaFin', 'FechaCreacion', 'UltimaActualizacion', 'Avance',
        'Tipo', 'FechaAceptacion', 'TiempoEjecucion' // ← NUEVAS COLUMNAS
      ]);
    }
    
    const id = 'EMERG-' + new Date().getTime();
    const ahora = new Date();
    
    const nuevaActividad = [
      id,                                    // A: ID
      '',                                    // B: ProyectoID (vacío para emergentes)
      datos.titulo,                          // C: Título
      datos.descripcion,                     // D: Descripción
      datos.responsable,                     // E: Responsable
      datos.jefe || '',                      // F: Jefe
      usuario.email,                         // G: Gerente (quien la crea)
      'Pendiente de Aceptación',            // H: Estado
      'Emergente',                          // I: Prioridad (todas emergentes son prioridad emergente)
      new Date(datos.fechaInicio),          // J: FechaInicio
      new Date(datos.fechaFin),             // K: FechaFin
      ahora,                                // L: FechaCreacion
      ahora,                                // M: UltimaActualizacion
      0,                                    // N: Avance
      'Emergente',                          // O: Tipo ← NUEVO
      '',                                   // P: FechaAceptacion ← NUEVO
      0                                     // Q: TiempoEjecucion (segundos) ← NUEVO
    ];
    
    sheet.appendRow(nuevaActividad);
    
    Logger.log('✅ Actividad emergente creada: ' + id);
    
    // Enviar notificación al colaborador
    try {
      enviarNotificacionEmergente(id, datos);
    } catch (e) {
      Logger.log('⚠️ No se pudo enviar notificación: ' + e.toString());
    }
    
    return {
      success: true,
      message: 'Actividad emergente creada y notificada',
      actividadId: id
    };
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    throw error;
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Obtener Actividades Emergentes Pendientes
// ===========================================================

function obtenerActividadesEmergentesPendientes(emailColaborador) {
  try {
    Logger.log('Obteniendo emergentes para: ' + emailColaborador);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const emergentes = [];
    
    for (let i = 1; i < data.length; i++) {
      const tipo = data[i][14] || 'Recurrente'; // Columna O
      const responsable = data[i][4]; // Columna E
      const estado = data[i][7]; // Columna H
      
      if (tipo === 'Emergente' && 
          responsable === emailColaborador && 
          estado === 'Pendiente de Aceptación') {
        
        emergentes.push({
          id: data[i][0],
          titulo: data[i][2],
          descripcion: data[i][3],
          gerente: data[i][6],
          fechaFin: data[i][10] instanceof Date ? data[i][10].toISOString() : data[i][10],
          fechaCreacion: data[i][11] instanceof Date ? data[i][11].toISOString() : data[i][11]
        });
      }
    }
    
    Logger.log('Emergentes pendientes: ' + emergentes.length);
    return emergentes;
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return [];
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Aceptar Actividad Emergente
// ===========================================================

function aceptarActividadEmergente(actividadId, actividadesAPausar) {
  try {
    Logger.log('=== Aceptando actividad emergente: ' + actividadId);
    Logger.log('Actividades a pausar: ' + JSON.stringify(actividadesAPausar));
    
    const usuario = obtenerUsuarioActual();
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === actividadId) {
        const responsable = data[i][4];
        
        // Verificar que sea el responsable quien acepta
        if (responsable !== usuario.email) {
          throw new Error('No eres el responsable de esta actividad');
        }
        
        const fila = i + 1;
        const ahora = new Date();
        
        // Actualizar estado de la emergente
        sheet.getRange(fila, 8).setValue('En Ejecución'); // Estado
        sheet.getRange(fila, 16).setValue(ahora); // FechaAceptacion (columna P)
        sheet.getRange(fila, 13).setValue(ahora); // UltimaActualizacion
        
        Logger.log('✅ Emergente aceptada');
        
        // Pausar actividades seleccionadas
        let resultadoPausa = { pausadas: 0 };
        if (actividadesAPausar && actividadesAPausar.length > 0) {
          resultadoPausa = pausarActividades(actividadesAPausar);
        }
        
        return {
          success: true,
          message: 'Actividad aceptada. Cronómetro iniciado.' + 
                   (resultadoPausa.pausadas > 0 ? ' ' + resultadoPausa.pausadas + ' actividad(es) pausada(s).' : ''),
          fechaInicio: ahora.toISOString(),
          actividadesPausadas: resultadoPausa.pausadas
        };
      }
    }
    
    throw new Error('Actividad no encontrada');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    throw error;
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Completar Actividad Emergente
// ===========================================================

function completarActividadEmergente(actividadId) {
  try {
    Logger.log('=== Completando actividad emergente: ' + actividadId);
    
    const usuario = obtenerUsuarioActual();
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === actividadId) {
        const responsable = data[i][4];
        const fechaAceptacion = data[i][15]; // Columna P
        
        if (responsable !== usuario.email) {
          throw new Error('No eres el responsable de esta actividad');
        }
        
        const fila = i + 1;
        const ahora = new Date();
        
        // Calcular tiempo de ejecución
        let tiempoEjecucion = 0;
        if (fechaAceptacion) {
          const inicio = new Date(fechaAceptacion);
          tiempoEjecucion = Math.round((ahora - inicio) / 1000); // segundos
        }
        
        // Actualizar actividad emergente
        sheet.getRange(fila, 8).setValue('Completada'); // Estado
        sheet.getRange(fila, 14).setValue(100); // Avance
        sheet.getRange(fila, 17).setValue(tiempoEjecucion); // TiempoEjecucion (columna Q)
        sheet.getRange(fila, 13).setValue(ahora); // UltimaActualizacion
        
        // Formatear tiempo
        const horas = Math.floor(tiempoEjecucion / 3600);
        const minutos = Math.floor((tiempoEjecucion % 3600) / 60);
        const segundos = tiempoEjecucion % 60;
        const tiempoFormateado = horas + 'h ' + minutos + 'm ' + segundos + 's';
        
        Logger.log('✅ Emergente completada en: ' + tiempoFormateado);
        
        // Reanudar actividades suspendidas
        const resultadoReanudar = reanudarActividadesSuspendidas(usuario.email);
        
        Logger.log('✅ Actividades reanudadas: ' + resultadoReanudar.reanudadas);
        
        return {
          success: true,
          message: 'Actividad completada en ' + tiempoFormateado + '. ' + 
                   resultadoReanudar.reanudadas + ' actividad(es) reanudada(s).',
          tiempoEjecucion: tiempoEjecucion,
          tiempoFormateado: tiempoFormateado,
          actividadesReanudadas: resultadoReanudar.reanudadas
        };
      }
    }

    throw new Error('Actividad no encontrada');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    throw error;
  }
}

// ===========================================================
// FUNCIÓN AUXILIAR: Enviar Notificación
// ===========================================================

function enviarNotificacionEmergente(actividadId, datos) {
  try {
    const asunto = '🚨 ACTIVIDAD EMERGENTE ASIGNADA';
    const cuerpo = `
      <h2 style="color: #e74c3c;">🚨 Nueva Actividad Emergente</h2>
      <p><strong>Título:</strong> ${datos.titulo}</p>
      <p><strong>Descripción:</strong> ${datos.descripcion}</p>
      <p><strong>Fecha límite:</strong> ${new Date(datos.fechaFin).toLocaleString('es-ES')}</p>
      <p style="color: #e74c3c;"><strong>⚠️ Esta actividad requiere tu atención inmediata.</strong></p>
      <p>Por favor, ingresa al sistema para aceptarla y comenzar a trabajar.</p>
    `;
    
    MailApp.sendEmail({
      to: datos.responsable,
      subject: asunto,
      htmlBody: cuerpo
    });
    
    Logger.log('Notificación enviada a: ' + datos.responsable);
    
  } catch (error) {
    Logger.log('Error enviando email: ' + error.toString());
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Obtener Actividades Activas del Colaborador
// ===========================================================

function obtenerActividadesActivasColaborador(emailColaborador) {
  try {
    Logger.log('=== Obteniendo actividades activas de: ' + emailColaborador);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const actividadesActivas = [];
    
    for (let i = 1; i < data.length; i++) {
      const responsable = data[i][4]; // Columna E
      const estado = data[i][7]; // Columna H
      const tipo = data[i][14] || 'Recurrente'; // Columna O
      
      // Solo actividades del colaborador que están en proceso y son recurrentes
      if (responsable === emailColaborador && 
          estado === 'En Proceso' && 
          tipo === 'Recurrente') {
        
        actividadesActivas.push({
          id: data[i][0],
          titulo: data[i][2],
          descripcion: data[i][3],
          proyectoId: data[i][1],
          prioridad: data[i][8],
          fechaFin: data[i][10] instanceof Date ? data[i][10].toISOString() : data[i][10],
          avance: data[i][13] || 0
        });
      }
    }
    
    Logger.log('Actividades en proceso encontradas: ' + actividadesActivas.length);
    return actividadesActivas;
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    return [];
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Pausar Actividades
// ===========================================================

function pausarActividades(actividadIds) {
  try {
    Logger.log('=== Pausando actividades: ' + JSON.stringify(actividadIds));
    
    if (!actividadIds || actividadIds.length === 0) {
      Logger.log('No hay actividades para pausar');
      return { success: true, message: 'Sin actividades a pausar', pausadas: 0 };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    const data = sheet.getDataRange().getValues();
    
    let pausadas = 0;
    const ahora = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const actividadId = data[i][0];
      
      if (actividadIds.includes(actividadId)) {
        const fila = i + 1;
        const estadoActual = data[i][7]; // Columna H
        
        // Guardar estado anterior y marcar como suspendida
        sheet.getRange(fila, 18).setValue(estadoActual); // Columna R: EstadoAnterior
        sheet.getRange(fila, 19).setValue(ahora); // Columna S: FechaSuspension
        sheet.getRange(fila, 8).setValue('Suspendida'); // Columna H: Estado
        sheet.getRange(fila, 13).setValue(ahora); // Columna M: UltimaActualizacion
        
        Logger.log('  ✅ Pausada: ' + data[i][2] + ' (estado anterior: ' + estadoActual + ')');
        pausadas++;
      }
    }
    
    Logger.log('✅ Total pausadas: ' + pausadas);
    
    return {
      success: true,
      message: pausadas + ' actividad(es) pausada(s)',
      pausadas: pausadas
    };
    
  } catch (error) {
    Logger.log('❌ Error pausando actividades: ' + error.toString());
    throw error;
  }
}

// ===========================================================
// NUEVA FUNCIÓN: Reanudar Actividades Suspendidas
// ===========================================================

function reanudarActividadesSuspendidas(emailColaborador) {
  try {
    Logger.log('=== Reanudando actividades de: ' + emailColaborador);
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    const data = sheet.getDataRange().getValues();
    
    let reanudadas = 0;
    const ahora = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const responsable = data[i][4]; // Columna E
      const estado = data[i][7]; // Columna H
      const estadoAnterior = data[i][17]; // Columna R
      
      if (responsable === emailColaborador && estado === 'Suspendida') {
        const fila = i + 1;
        const tituloActividad = data[i][2];
        
        // Restaurar estado anterior
        const estadoARestaurar = estadoAnterior || 'En Proceso';
        sheet.getRange(fila, 8).setValue(estadoARestaurar); // Estado
        sheet.getRange(fila, 13).setValue(ahora); // UltimaActualizacion
        
        // Limpiar campos de suspensión
        sheet.getRange(fila, 18).setValue(''); // EstadoAnterior
        sheet.getRange(fila, 19).setValue(''); // FechaSuspension
        
        Logger.log('  ✅ Reanudada: ' + tituloActividad + ' → ' + estadoARestaurar);
        reanudadas++;
      }
    }
    
    Logger.log('✅ Total reanudadas: ' + reanudadas);
    
    return {
      success: true,
      message: reanudadas + ' actividad(es) reanudada(s)',
      reanudadas: reanudadas
    };
    
  } catch (error) {
    Logger.log('❌ Error reanudando actividades: ' + error.toString());
    throw error;
  }
}

function verificarCruceYSugerencias(responsableEmail, nuevaInicio, nuevaFin, excluirId) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Actividades');
    if (!sheet) return { cruce: false };

    const data    = sheet.getDataRange().getValues();
    const activas = ['Pendiente', 'En Proceso', 'En Revisión'];
    const duracionMs = nuevaFin - nuevaInicio;

    // Recopilar todas las actividades del responsable que estén activas
    const bloques = [];
    let actividadCruce = null;

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][0] === excluirId) continue;
      if (data[i][4] !== responsableEmail) continue; // col E = responsable
      if (!activas.includes(data[i][7])) continue;   // col H = estado

      const ini = data[i][9]  ? new Date(data[i][9])  : null; // col J
      const fin = data[i][10] ? new Date(data[i][10]) : null; // col K
      if (!ini || !fin) continue;

      bloques.push({ ini, fin, titulo: data[i][2] || 'Sin título' });

      // Detectar cruce: nuevaInicio < fin existente Y nuevaFin > ini existente
      if (nuevaInicio < fin && nuevaFin > ini) {
        actividadCruce = data[i][2] || 'Sin título';
      }
    }

    if (!actividadCruce) return { cruce: false };

    // ── Calcular slots libres en próximos 3 días con disponibilidad ──
    const horasLibres = [];
    let diaActual    = new Date(nuevaInicio);
    diaActual.setHours(0, 0, 0, 0);
    let diasConSlots = 0;
    let intentos     = 0;

    while (diasConSlots < 3 && intentos < 30) {
      intentos++;

      // Bloques ocupados ese día
      const bloquesDelDia = bloques
        .filter(function(b) {
          return b.ini.toDateString() === diaActual.toDateString() ||
                 b.fin.toDateString() === diaActual.toDateString() ||
                 (b.ini <= diaActual && b.fin >= new Date(diaActual.getTime() + 86400000));
        })
        .map(function(b) { return { ini: b.ini, fin: b.fin }; })
        .sort(function(a, b) { return a.ini - b.ini; });

      // Generar slots libres del día (00:00 a 23:59)
      const inicioDia = new Date(diaActual); inicioDia.setHours(0, 0, 0, 0);
      const finDia    = new Date(diaActual); finDia.setHours(23, 59, 59, 0);

      const slotsDelDia = [];
      let cursor = new Date(inicioDia);

      bloquesDelDia.forEach(function(b) {
        const bIni = b.ini < inicioDia ? inicioDia : b.ini;
        const bFin = b.fin > finDia    ? finDia    : b.fin;

        if (cursor < bIni) {
          const gapMs = bIni - cursor;
          if (gapMs >= duracionMs) {
            const slotFin = new Date(cursor.getTime() + duracionMs);
            slotsDelDia.push(formatSlot(diaActual, cursor, slotFin));
          }
        }
        if (bFin > cursor) cursor = new Date(bFin);
      });

      // Gap al final del día
      if (cursor < finDia) {
        const gapMs = finDia - cursor;
        if (gapMs >= duracionMs) {
          const slotFin = new Date(cursor.getTime() + duracionMs);
          slotsDelDia.push(formatSlot(diaActual, cursor, slotFin));
        }
      }

      // Si no hay bloques ese día, todo el día es libre
      if (bloquesDelDia.length === 0) {
        const slotFin = new Date(inicioDia.getTime() + duracionMs);
        slotsDelDia.push(formatSlot(diaActual, inicioDia, slotFin));
      }

      if (slotsDelDia.length > 0) {
        // Tomar máximo 2 slots por día
        slotsDelDia.slice(0, 2).forEach(function(s) { horasLibres.push(s); });
        diasConSlots++;
      }

      // Avanzar al día siguiente
      diaActual = new Date(diaActual.getTime() + 86400000);
    }

    return {
      cruce:       true,
      actividad:   actividadCruce,
      horasLibres: horasLibres
    };

  } catch(e) {
    Logger.log('verificarCruceYSugerencias error: ' + e);
    return { cruce: false };
  }
}

function formatSlot(dia, ini, fin) {
  const fecha = Utilities.formatDate(dia, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const hIni  = Utilities.formatDate(ini, Session.getScriptTimeZone(), 'HH:mm');
  const hFin  = Utilities.formatDate(fin, Session.getScriptTimeZone(), 'HH:mm');
  return {
    label:      fecha + '  ' + hIni + ' – ' + hFin,
    fechaInicio: ini.toISOString(),
    fechaFin:    fin.toISOString()
  };
}

