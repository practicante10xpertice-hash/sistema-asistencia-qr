// ============================================================
// SISTEMA DE CONTROL DE ASISTENCIA QR - Code.gs CORREGIDO
// ============================================================

var CONFIG = {
  SHEETS: {
    USUARIOS:    'Usuarios',
    ASISTENCIAS: 'Asistencias',
    PERSONAL:    'Personal',
    CONFIG_HOJA: 'Config'
  },
  TOLERANCIA_MINUTOS: 15
};

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Sistema de Asistencia QR')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function hashPassword(password) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(function(b){ return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function generateId(prefix) {
  return prefix + Date.now().toString(36).toUpperCase() + Math.random().toString(36).substr(2,4).toUpperCase();
}

function getSheet(name) {
  var sh = getSS().getSheetByName(name);
  if (!sh) { initSheets(); sh = getSS().getSheetByName(name); }
  return sh;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return [];
  var headers = data[0].map(function(h){ return h.toString().trim(); });
  var tz = Session.getScriptTimeZone();
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var val = row[i];
      if (val instanceof Date) {
        obj[h] = Utilities.formatDate(val, tz, 'HH:mm');
      } else {
        obj[h] = (val !== null && val !== undefined) ? val.toString() : '';
      }
    });
    return obj;
  });
}

function getConfigValue(key) {
  try {
    var sheet = getSheet(CONFIG.SHEETS.CONFIG_HOJA);
    var data  = sheetToObjects(sheet);
    var row   = data.filter(function(r){ return r['Clave'] === key; })[0];
    return row ? row['Valor'] : null;
  } catch(e) { return null; }
}

function initSheets() {
  try {
    var ss = getSS();

    var shUsuarios = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
    if (!shUsuarios) {
      shUsuarios = ss.insertSheet(CONFIG.SHEETS.USUARIOS);
      shUsuarios.appendRow(['ID','Nombre','Email','Password','Rol','Activo','FechaCreacion']);
      shUsuarios.appendRow(['USR001','Administrador','admin@empresa.com', hashPassword('admin123'),'Administrador','true', new Date().toISOString()]);
      shUsuarios.appendRow(['USR002','Usuario Demo','usuario@empresa.com', hashPassword('usuario123'),'Usuario','true', new Date().toISOString()]);
      shUsuarios.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    }

    var shPersonal = ss.getSheetByName(CONFIG.SHEETS.PERSONAL);
    if (!shPersonal) {
      shPersonal = ss.insertSheet(CONFIG.SHEETS.PERSONAL);
      shPersonal.appendRow(['ID','Nombre','DNI','Cargo','Area','DiasLaborales','HoraEntrada','HoraSalida','SalidaAlmuerzo','EntradaTarde','SalidaTarde','DobleTurno','Activo','FechaCreacion']);
      shPersonal.appendRow(['EMP001','Juan Pérez García','12345678','Desarrollador','TI','Lunes,Martes,Miércoles,Jueves,Viernes','08:00','17:00','13:00','14:00','18:00','false','true', new Date().toISOString()]);
      shPersonal.appendRow(['EMP002','María López Torres','87654321','Diseñadora','Marketing','Lunes,Martes,Miércoles,Jueves,Viernes','09:00','18:00','13:00','14:00','19:00','false','true', new Date().toISOString()]);
      shPersonal.appendRow(['EMP003','Carlos Ruiz Mendoza','11223344','Analista','Finanzas','Lunes,Martes,Miércoles,Jueves,Viernes','07:00','16:00','12:00','13:00','17:00','true','true', new Date().toISOString()]);
      shPersonal.getRange(1,1,1,14).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    }

    var shAsistencias = ss.getSheetByName(CONFIG.SHEETS.ASISTENCIAS);
    if (!shAsistencias) {
      shAsistencias = ss.insertSheet(CONFIG.SHEETS.ASISTENCIAS);
      shAsistencias.appendRow(['ID','EmpleadoID','NombreEmpleado','DNI','Area','Cargo','Fecha','Hora','Turno','Estado','IP','Observacion']);
      shAsistencias.getRange(1,1,1,12).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    }

    var shConfig = ss.getSheetByName(CONFIG.SHEETS.CONFIG_HOJA);
    if (!shConfig) {
      shConfig = ss.insertSheet(CONFIG.SHEETS.CONFIG_HOJA);
      shConfig.appendRow(['Clave','Valor','Descripcion']);
      shConfig.appendRow(['empresa_nombre','Mi Empresa S.A.','Nombre de la empresa']);
      shConfig.appendRow(['empresa_logo','','URL del logo']);
      shConfig.appendRow(['tolerancia_minutos','15','Minutos de tolerancia']);
      shConfig.appendRow(['qr_token', Utilities.getUuid(),'Token del QR general']);
      shConfig.getRange(1,1,1,3).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    }

    return { ok: true, message: 'Hojas inicializadas correctamente' };
  } catch(e) {
    return { ok: false, message: 'Error initSheets: ' + e.message };
  }
}

function login(email, password) {
  try {
    var sheet    = getSheet(CONFIG.SHEETS.USUARIOS);
    var usuarios = sheetToObjects(sheet);
    var hash     = hashPassword(password);
    var usuario  = null;
    for (var i = 0; i < usuarios.length; i++) {
      var u = usuarios[i];
      if (u['Email'].toLowerCase() === email.toLowerCase() && u['Password'] === hash && u['Activo'] === 'true') {
        usuario = u; break;
      }
    }
    if (!usuario) return { ok: false, message: 'Credenciales incorrectas' };
    return { ok: true, token: 'TOKEN_' + Date.now(), usuario: { id: usuario['ID'], nombre: usuario['Nombre'], email: usuario['Email'], rol: usuario['Rol'] }};
  } catch(e) {
    return { ok: false, message: 'Error login: ' + e.message };
  }
}

function getPersonal() {
  try {
    var sheet = getSS().getSheetByName(CONFIG.SHEETS.PERSONAL);
    if (!sheet) return { ok: false, message: 'Hoja Personal no encontrada', data: [] };
    var data = sheet.getDataRange().getValues();
    if (!data || data.length < 2) return { ok: true, data: [] };
    var headers = data[0].map(function(h){ return h.toString().trim(); });
    var tz = Session.getScriptTimeZone();
    var personal = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = row[j];
        if (val instanceof Date) {
          obj[headers[j]] = Utilities.formatDate(val, tz, 'HH:mm');
        } else {
          obj[headers[j]] = (val !== null && val !== undefined) ? val.toString() : '';
        }
      }
      if (obj['Activo'] === 'true' || obj['Activo'] === 'TRUE') personal.push(obj);
    }
    return { ok: true, data: personal };
  } catch(e) {
    return { ok: false, message: 'Error getPersonal: ' + e.message, data: [] };
  }
}

function crearEmpleado(datos) {
  try {
    var sheet = getSheet(CONFIG.SHEETS.PERSONAL);
    var id    = generateId('EMP');
    var dias  = Array.isArray(datos.diasLaborales) ? datos.diasLaborales.join(',') : (datos.diasLaborales || '');
    sheet.appendRow([id, datos.nombre||'', datos.dni||'', datos.cargo||'', datos.area||'', dias, datos.horaEntrada||'08:00', datos.horaSalida||'17:00', datos.salidaAlmuerzo||'', datos.entradaTarde||'', datos.salidaTarde||'', datos.dobleTurno?'true':'false', 'true', new Date().toISOString()]);
    return { ok: true, message: 'Empleado creado exitosamente', id: id };
  } catch(e) {
    return { ok: false, message: 'Error crearEmpleado: ' + e.message };
  }
}

function actualizarEmpleado(id, datos) {
  try {
    var sheet   = getSheet(CONFIG.SHEETS.PERSONAL);
    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var idIdx   = headers.indexOf('ID');
    var rowNum  = -1;
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][idIdx].toString() === id.toString()) { rowNum = i + 1; break; }
    }
    if (rowNum === -1) return { ok: false, message: 'Empleado no encontrado' };
    var dias = Array.isArray(datos.diasLaborales) ? datos.diasLaborales.join(',') : (datos.diasLaborales || '');
    var campos = { 'Nombre': datos.nombre||'', 'DNI': datos.dni||'', 'Cargo': datos.cargo||'', 'Area': datos.area||'', 'DiasLaborales': dias, 'HoraEntrada': datos.horaEntrada||'', 'HoraSalida': datos.horaSalida||'', 'SalidaAlmuerzo': datos.salidaAlmuerzo||'', 'EntradaTarde': datos.entradaTarde||'', 'SalidaTarde': datos.salidaTarde||'', 'DobleTurno': datos.dobleTurno?'true':'false' };
    var keys = Object.keys(campos);
    for (var k = 0; k < keys.length; k++) {
      var ci = headers.indexOf(keys[k]);
      if (ci !== -1) sheet.getRange(rowNum, ci + 1).setValue(campos[keys[k]]);
    }
    return { ok: true, message: 'Empleado actualizado exitosamente' };
  } catch(e) {
    return { ok: false, message: 'Error actualizarEmpleado: ' + e.message };
  }
}

function eliminarEmpleado(id) {
  try {
    var sheet     = getSheet(CONFIG.SHEETS.PERSONAL);
    var allData   = sheet.getDataRange().getValues();
    var headers   = allData[0];
    var idIdx     = headers.indexOf('ID');
    var activoIdx = headers.indexOf('Activo');
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][idIdx].toString() === id.toString()) {
        sheet.getRange(i + 1, activoIdx + 1).setValue('false');
        return { ok: true, message: 'Empleado eliminado' };
      }
    }
    return { ok: false, message: 'Empleado no encontrado' };
  } catch(e) {
    return { ok: false, message: 'Error eliminarEmpleado: ' + e.message };
  }
}

function registrarAsistenciaQR(token, empleadoId) {
  try {
    var qrToken = getConfigValue('qr_token');
    if (token !== qrToken && token !== 'DEMO_QR_TOKEN') return { ok: false, message: 'QR inválido o expirado' };
    var empleados = sheetToObjects(getSheet(CONFIG.SHEETS.PERSONAL)).filter(function(p){ return p['Activo'] === 'true'; });
    var empleado  = null;
    for (var i = 0; i < empleados.length; i++) {
      if (empleados[i]['ID'] === empleadoId || empleados[i]['DNI'] === empleadoId) { empleado = empleados[i]; break; }
    }
    if (!empleado) return { ok: false, message: 'Empleado no encontrado' };
    var tz        = Session.getScriptTimeZone();
    var ahora     = new Date();
    var fechaStr  = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
    var horaStr   = Utilities.formatDate(ahora, tz, 'HH:mm:ss');
    var dias      = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
    var diaNombre = dias[ahora.getDay()];
    var diasLab   = empleado['DiasLaborales'].split(',').map(function(d){ return d.trim(); });
    if (diasLab.indexOf(diaNombre) === -1) return { ok: false, message: 'Hoy (' + diaNombre + ') no es día laboral' };
    var asistencias   = sheetToObjects(getSheet(CONFIG.SHEETS.ASISTENCIAS));
    var yaRegistrado  = asistencias.filter(function(a){ return a['EmpleadoID'] === empleado['ID'] && a['Fecha'] === fechaStr; })[0];
    var turno = 'Entrada';
    if (yaRegistrado) {
      if (empleado['DobleTurno'] === 'true' && empleado['EntradaTarde']) {
        var yaEntradaTarde = asistencias.filter(function(a){ return a['EmpleadoID'] === empleado['ID'] && a['Fecha'] === fechaStr && a['Turno'] === 'Entrada Tarde'; })[0];
        if (yaEntradaTarde) return { ok: false, message: 'Ya registraste asistencia para el turno tarde hoy' };
        turno = 'Entrada Tarde';
      } else {
        return { ok: false, message: 'Ya registraste asistencia hoy' };
      }
    }
    var horaEsperada = (turno === 'Entrada Tarde') ? empleado['EntradaTarde'] : empleado['HoraEntrada'];
    var estado = 'Puntual';
    if (horaEsperada) {
      var partsEsp = horaEsperada.split(':');
      var minEsp   = parseInt(partsEsp[0]) * 60 + parseInt(partsEsp[1]);
      var partsAct = horaStr.split(':');
      var minAct   = parseInt(partsAct[0]) * 60 + parseInt(partsAct[1]);
      if (minAct > minEsp + 15) estado = 'Tarde';
    }
    getSheet(CONFIG.SHEETS.ASISTENCIAS).appendRow([generateId('AST'), empleado['ID'], empleado['Nombre'], empleado['DNI'], empleado['Area'], empleado['Cargo'], fechaStr, horaStr, turno, estado, '', '']);
    return { ok: true, message: 'Asistencia: ' + estado, data: { nombre: empleado['Nombre'], hora: horaStr, fecha: fechaStr, estado: estado, turno: turno }};
  } catch(e) {
    return { ok: false, message: 'Error registrar: ' + e.message };
  }
}

function getAsistencias(filtros) {
  try {
    var sheet = getSheet(CONFIG.SHEETS.ASISTENCIAS);
    var data  = sheetToObjects(sheet);
    filtros = filtros || {};
    if (filtros.fechaInicio) data = data.filter(function(a){ return a['Fecha'] >= filtros.fechaInicio; });
    if (filtros.fechaFin)    data = data.filter(function(a){ return a['Fecha'] <= filtros.fechaFin; });
    if (filtros.empleadoId)  data = data.filter(function(a){ return a['EmpleadoID'] === filtros.empleadoId; });
    if (filtros.area)        data = data.filter(function(a){ return a['Area'] === filtros.area; });
    data.sort(function(a,b){ return ((b['Fecha']||'')+(b['Hora']||'')).localeCompare((a['Fecha']||'')+(a['Hora']||'')); });
    return { ok: true, data: data };
  } catch(e) {
    return { ok: false, message: 'Error getAsistencias: ' + e.message, data: [] };
  }
}

function getDashboardData() {
  try {
    var tz   = Session.getScriptTimeZone();
    var hoy  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var personal    = sheetToObjects(getSheet(CONFIG.SHEETS.PERSONAL)).filter(function(p){ return p['Activo'] === 'true'; });
    var asistencias = sheetToObjects(getSheet(CONFIG.SHEETS.ASISTENCIAS));
    var hoyAsist    = asistencias.filter(function(a){ return a['Fecha'] === hoy; });
    var graf = [];
    for (var i = 6; i >= 0; i--) {
      var d = new Date(); d.setDate(d.getDate() - i);
      var dStr   = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      var dLabel = Utilities.formatDate(d, tz, 'dd/MM');
      graf.push({ fecha: dLabel, cantidad: asistencias.filter(function(a){ return a['Fecha'] === dStr; }).length });
    }
    var porArea = {};
    hoyAsist.forEach(function(a){ var ar = a['Area']||'Sin área'; porArea[ar] = (porArea[ar]||0)+1; });
    var presentesIds = [];
    hoyAsist.forEach(function(a){ if (presentesIds.indexOf(a['EmpleadoID']) === -1) presentesIds.push(a['EmpleadoID']); });
    return { ok: true, data: { totalEmpleados: personal.length, asistenciasHoy: hoyAsist.length, tardanzasHoy: hoyAsist.filter(function(a){ return a['Estado']==='Tarde'; }).length, ausentesHoy: personal.filter(function(p){ return presentesIds.indexOf(p['ID'])===-1; }).length, graficaSemanal: graf, porArea: porArea, ultimasAsistencias: hoyAsist.slice(0,8) }};
  } catch(e) {
    return { ok: false, message: 'Error dashboard: ' + e.message, data: null };
  }
}

function getQRToken() {
  try {
    var token = getConfigValue('qr_token');
    if (!token) {
      token = Utilities.getUuid();
      var sheet = getSheet(CONFIG.SHEETS.CONFIG_HOJA);
      var allData = sheet.getDataRange().getValues();
      var claveIdx = allData[0].indexOf('Clave'), valorIdx = allData[0].indexOf('Valor');
      for (var i = 1; i < allData.length; i++) {
        if (allData[i][claveIdx] === 'qr_token') { sheet.getRange(i+1, valorIdx+1).setValue(token); break; }
      }
    }
    var url = ScriptApp.getService().getUrl();
    return { ok: true, token: token, url: url + '?qr=' + token };
  } catch(e) {
    return { ok: false, message: 'Error getQRToken: ' + e.message };
  }
}

function regenerarQR() {
  try {
    var nuevoToken = Utilities.getUuid();
    var sheet = getSheet(CONFIG.SHEETS.CONFIG_HOJA);
    var allData = sheet.getDataRange().getValues();
    var claveIdx = allData[0].indexOf('Clave'), valorIdx = allData[0].indexOf('Valor');
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][claveIdx] === 'qr_token') { sheet.getRange(i+1, valorIdx+1).setValue(nuevoToken); break; }
    }
    return { ok: true, token: nuevoToken, url: ScriptApp.getService().getUrl() + '?qr=' + nuevoToken };
  } catch(e) {
    return { ok: false, message: 'Error regenerarQR: ' + e.message };
  }
}

function generarReportePDF(filtros) {
  try {
    var empresa     = getConfigValue('empresa_nombre') || 'Mi Empresa S.A.';
    var asistencias = getAsistencias(filtros);
    if (!asistencias.ok) return asistencias;
    var ss   = getSS();
    var temp = ss.insertSheet('RPT_' + Date.now());
    temp.getRange(1,1,1,8).merge().setValue(empresa + ' - REPORTE DE ASISTENCIAS').setFontSize(14).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff').setHorizontalAlignment('center');
    temp.getRange(2,1,1,8).merge().setValue('Período: ' + (filtros.fechaInicio||'Inicio') + ' al ' + (filtros.fechaFin||'Hoy')).setFontSize(10).setHorizontalAlignment('center').setBackground('#e8f0fe');
    temp.getRange(3,1,1,8).merge().setValue('Generado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')).setFontSize(9).setHorizontalAlignment('right').setFontColor('#666');
    temp.getRange(5,1,1,8).setValues([['#','Empleado','DNI','Área','Cargo','Fecha','Hora','Estado']]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff').setHorizontalAlignment('center');
    var rows = asistencias.data.map(function(a,i){ return [i+1, a['NombreEmpleado'], a['DNI'], a['Area'], a['Cargo'], a['Fecha'], a['Hora'], a['Estado']]; });
    if (rows.length > 0) {
      temp.getRange(6,1,rows.length,8).setValues(rows);
      rows.forEach(function(r,i){ if(r[7]==='Tarde') temp.getRange(6+i,8).setBackground('#fce8e6').setFontColor('#c62828'); else temp.getRange(6+i,8).setBackground('#e6f4ea').setFontColor('#2e7d32'); });
    }
    temp.autoResizeColumns(1,8);
    var pdfUrl = 'https://docs.google.com/spreadsheets/d/'+ss.getId()+'/export?format=pdf&gid='+temp.getSheetId()+'&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false';
    var response = UrlFetchApp.fetch(pdfUrl, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }});
    ss.deleteSheet(temp);
    return { ok: true, pdf: Utilities.base64Encode(response.getContent()), filename: 'Reporte_Asistencias_' + (filtros.fechaInicio||'') + '_' + (filtros.fechaFin||'') + '.pdf' };
  } catch(e) {
    return { ok: false, message: 'Error PDF: ' + e.message };
  }
}

function dispatch(accion, payload) {
  var result;
  try {
    if (payload && typeof payload === 'object') {
      payload = JSON.parse(JSON.stringify(payload));
    }
    payload = payload || {};
    switch(accion) {
      case 'init':               result = initSheets(); break;
      case 'login':              result = login(payload.email, payload.password); break;
      case 'getPersonal':        result = getPersonal(); break;
      case 'crearEmpleado':      result = crearEmpleado(payload); break;
      case 'actualizarEmpleado': result = actualizarEmpleado(payload.id, payload.datos); break;
      case 'eliminarEmpleado':   result = eliminarEmpleado(payload.id); break;
      case 'getAsistencias':     result = getAsistencias(payload); break;
      case 'getDashboard':       result = getDashboardData(); break;
      case 'registrarQR':        result = registrarAsistenciaQR(payload.token, payload.empleadoId); break;
      case 'getQRToken':         result = getQRToken(); break;
      case 'regenerarQR':        result = regenerarQR(); break;
      case 'generarPDF':         result = generarReportePDF(payload); break;
      default:                   result = { ok: false, message: 'Acción no reconocida: ' + accion };
    }
  } catch(e) {
    result = { ok: false, message: 'Error en dispatch [' + accion + ']: ' + e.message };
  }
  if (!result || typeof result !== 'object') result = { ok: false, message: 'Respuesta inválida del servidor' };
  if (result.ok === undefined) result.ok = false;
  return result;
}
