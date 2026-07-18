/* ============================================================
   CÓDIGO DEL SERVIDOR - App de Recibos FOTHESALUD / SISOLE'S
   Pega este archivo completo en Apps Script como "Código.gs"
   ============================================================ */

function ensureSheets_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let configSheet = ss.getSheetByName('Config');
  if(!configSheet){
    configSheet = ss.insertSheet('Config');
    configSheet.appendRow(['Empresa','Nombre','NIT','Direccion','Email','Celular','ContadorEmail','SiguienteNumero']);
    configSheet.appendRow(['sisoles',"SISOLE'S S.A.S.",'901.877.147-6','Medellín, Antioquia','','','',1]);
    configSheet.appendRow(['fothesalud','FOTHESALUD S.A.S.','900.729.153-3','Carrera 65 8B 91, Medellín - Antioquia','','','',1]);
  }

  let recibosSheet = ss.getSheetByName('Recibos');
  if(!recibosSheet){
    recibosSheet = ss.insertSheet('Recibos');
    recibosSheet.appendRow(['ID','Empresa','Numero','Fecha','Nombre','Cedula','Celular','ItemsJSON','Total','Estado','GeneradoPor','CreadoEn']);
  }

  let usuariosSheet = ss.getSheetByName('Usuarios');
  if(!usuariosSheet){
    usuariosSheet = ss.insertSheet('Usuarios');
    usuariosSheet.appendRow(['Nombre','Email']);
  }

  return { configSheet: configSheet, recibosSheet: recibosSheet, usuariosSheet: usuariosSheet };
}

/* Sirve la aplicación cuando se abre el link del despliegue.
   El ícono ya NO se sirve desde aquí: ahora vive incrustado en base64
   directo en Index.html, y además está la página "cascarón" (hospedada
   fuera de Google) que es la que de verdad controla el ícono en la
   pantalla de inicio del iPhone, porque Apps Script siempre entrega su
   contenido dentro de un iframe interno y Safari no puede leer el <head>
   de una página que está anidada así. El setXFrameOptionsMode(ALLOWALL)
   de abajo es justamente lo que permite que esa página cascarón pueda
   incrustar esta app en su propio iframe. */
function doGet(e){
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Recibos - FOTHESALUD / SISOLE'S")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function formatDateForClient_(d){
  if(!d) return '';
  if(Object.prototype.toString.call(d) === '[object Date]'){
    const y = d.getFullYear(), m = ('0'+(d.getMonth()+1)).slice(-2), day = ('0'+d.getDate()).slice(-2);
    return y+'-'+m+'-'+day;
  }
  return String(d);
}

/* Igual que formatDateForClient_ pero conservando hora:minuto:segundo, para
   campos de auditoría (CreadoEn). Nunca se debe devolver un objeto Date "crudo"
   al cliente: Apps Script puede fallar en silencio al empaquetar la respuesta. */
function formatDateTimeForClient_(d){
  if(!d) return '';
  if(Object.prototype.toString.call(d) === '[object Date]'){
    return d.toISOString();
  }
  return String(d);
}

/* Devuelve toda la información que necesita la app al abrirse:
   configuración de cada empresa, historial completo de recibos y usuarios registrados. */
function cargarTodo(){
  const s = ensureSheets_();

  const configRows = s.configSheet.getDataRange().getValues();
  const configHeaders = configRows.shift();
  const companies = {};
  configRows.forEach(function(row){
    const obj = {};
    configHeaders.forEach(function(h,i){ obj[h] = row[i]; });
    companies[obj.Empresa] = {
      nombre: obj.Nombre, nit: obj.NIT, direccion: obj.Direccion,
      email: obj.Email || '', celular: obj.Celular || '',
      contadorEmail: obj.ContadorEmail || '', nextNumero: obj.SiguienteNumero || 1
    };
  });

  const recRows = s.recibosSheet.getDataRange().getValues();
  const recHeaders = recRows.shift();
  const recibos = recRows.filter(function(row){ return row[0]; }).map(function(row){
    const obj = {};
    recHeaders.forEach(function(h,i){ obj[h] = row[i]; });
    let items = [];
    try{ items = JSON.parse(obj.ItemsJSON || '[]'); }catch(e){ items = []; }
    return {
      id: obj.ID, empresa: obj.Empresa, numero: obj.Numero,
      fecha: formatDateForClient_(obj.Fecha), nombre: obj.Nombre, cedula: obj.Cedula,
      celular: obj.Celular, items: items, total: obj.Total, estado: obj.Estado,
      generadoPor: obj.GeneradoPor, createdAt: formatDateTimeForClient_(obj.CreadoEn)
    };
  }).reverse();

  const usrRows = s.usuariosSheet.getDataRange().getValues();
  usrRows.shift();
  const usuarios = {};
  usrRows.forEach(function(row){
    if(row[0]) usuarios[row[0]] = { email: row[1] || '' };
  });

  return { companies: companies, recibos: recibos, usuarios: usuarios };
}

/* Crea un recibo nuevo, asignando el consecutivo de forma segura
   (con bloqueo) para que dos personas guardando al mismo tiempo
   nunca reciban el mismo número. */
function crearRecibo(payload){
  const lock = LockService.getScriptLock();
  let tieneLock = false;
  try{
    lock.waitLock(20000);
    tieneLock = true;
    const s = ensureSheets_();
    const configRows = s.configSheet.getDataRange().getValues();
    const headers = configRows[0];
    const empresaColIdx = headers.indexOf('Empresa');
    const siguienteColIdx = headers.indexOf('SiguienteNumero');

    let filaEmpresa = -1;
    for(let i=1;i<configRows.length;i++){
      if(configRows[i][empresaColIdx] === payload.empresa){ filaEmpresa = i+1; break; }
    }
    if(filaEmpresa === -1) throw new Error('Empresa no encontrada en Config: '+payload.empresa);

    const numeroAsignado = s.configSheet.getRange(filaEmpresa, siguienteColIdx+1).getValue() || 1;
    s.configSheet.getRange(filaEmpresa, siguienteColIdx+1).setValue(Number(numeroAsignado)+1);

    const id = 'r_' + new Date().getTime() + '_' + Math.floor(Math.random()*1000);
    const creadoEn = new Date();

    s.recibosSheet.appendRow([
      id, payload.empresa, numeroAsignado, payload.fecha, payload.nombre, payload.cedula || '',
      payload.celular || '', JSON.stringify(payload.items || []), payload.total || 0,
      payload.estado || 'Pendiente', payload.generadoPor || '', creadoEn
    ]);

    return {
      id: id, empresa: payload.empresa, numero: numeroAsignado, fecha: payload.fecha,
      nombre: payload.nombre, cedula: payload.cedula || '', celular: payload.celular || '',
      items: payload.items || [], total: payload.total || 0, estado: payload.estado || 'Pendiente',
      generadoPor: payload.generadoPor || '', createdAt: creadoEn.toISOString()
    };
  } finally {
    if(tieneLock) lock.releaseLock();
  }
}

/* Guarda cambios en el perfil de una empresa (NIT, dirección, correo, celular,
   correo del contador y, si se ajusta manualmente, el siguiente número). */
function guardarConfigEmpresa(empresa, campos){
  const s = ensureSheets_();
  const rows = s.configSheet.getDataRange().getValues();
  const headers = rows[0];
  const empresaColIdx = headers.indexOf('Empresa');
  let fila = -1;
  for(let i=1;i<rows.length;i++){
    if(rows[i][empresaColIdx] === empresa){ fila = i+1; break; }
  }
  if(fila === -1) throw new Error('Empresa no encontrada: '+empresa);

  function setCol(nombreCol, valor){
    const idx = headers.indexOf(nombreCol);
    if(idx>=0) s.configSheet.getRange(fila, idx+1).setValue(valor);
  }
  if(campos.nit !== undefined) setCol('NIT', campos.nit);
  if(campos.direccion !== undefined) setCol('Direccion', campos.direccion);
  if(campos.email !== undefined) setCol('Email', campos.email);
  if(campos.celular !== undefined) setCol('Celular', campos.celular);
  if(campos.contadorEmail !== undefined) setCol('ContadorEmail', campos.contadorEmail);
  if(campos.nextNumero !== undefined) setCol('SiguienteNumero', campos.nextNumero);
  return true;
}

/* Guarda o actualiza el correo remitente de un usuario. */
function guardarUsuario(nombre, email){
  const s = ensureSheets_();
  const rows = s.usuariosSheet.getDataRange().getValues();
  let fila = -1;
  for(let i=1;i<rows.length;i++){
    if(rows[i][0] === nombre){ fila = i+1; break; }
  }
  if(fila === -1){
    s.usuariosSheet.appendRow([nombre, email]);
  } else {
    s.usuariosSheet.getRange(fila, 2).setValue(email);
  }
  return true;
}
