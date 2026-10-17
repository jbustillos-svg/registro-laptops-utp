/*
 * Coordinador central de contingencia v1.4.1.
 * Debe instalarse como script VINCULADO al Spreadsheet Control de Laptops.
 * El secreto se guarda exclusivamente en Script Properties: UTP_SYNC_API_KEY.
 */

const HOJA_ALUMNOS = 'Alumnos';
const HOJA_REGISTROS = 'Registros';
const HOJA_EVENTOS = 'Eventos_Sincronizacion';
const HOJA_CONFIGURACION = 'Configuracion_Sistema';
const CABECERAS_EVENTOS = [
  'ID_EVENTO', 'TIPO', 'MATRICULA', 'LAPTOP_ID', 'FECHA_HORA_EVENTO',
  'ESTADO', 'FECHA_PREPARACION', 'FECHA_APLICACION',
  'REFERENCIA_REGISTRO', 'DATOS_CONTROL', 'EQUIPO_ORIGEN'
];
const CABECERAS_CONFIGURACION = ['CLAVE', 'VALOR'];

function doPost(e) {
  let evento = null;
  let hojaEventos = null;
  let filaEvento = null;
  let bloqueo = null;
  try {
    evento = leerSolicitud_(e);
    if (!claveValida_(evento.api_key)) {
      return responder_({estado: 'ERROR', codigo: 'NO_AUTORIZADO'});
    }
    validarEvento_(evento);
    bloqueo = LockService.getScriptLock();
    if (!bloqueo.tryLock(25000)) {
      return responder_({estado: 'ERROR', codigo: 'COORDINADOR_OCUPADO'});
    }

    const libro = SpreadsheetApp.getActiveSpreadsheet();
    hojaEventos = asegurarHojaEventos_(libro);
    const encontrado = buscarEvento_(hojaEventos, evento.id_evento);
    if (encontrado && encontrado.estado === 'APLICADO') {
      return responder_({
        estado: 'YA_APLICADO', referencia_registro: encontrado.referencia || ''
      });
    }
    if (encontrado) {
      filaEvento = encontrado.fila;
      evento = eventoGuardado_(encontrado, evento);
    } else {
      filaEvento = registrarPreparando_(hojaEventos, evento);
    }

    const referencia = aplicarEvento_(libro, evento);
    actualizarEvento_(hojaEventos, filaEvento, 'APLICADO', referencia, '');
    return responder_({estado: 'APLICADO', referencia_registro: referencia});
  } catch (error) {
    if (hojaEventos && filaEvento) {
      actualizarEvento_(hojaEventos, filaEvento, 'ERROR', '', codigoError_(error));
    }
    return responder_({estado: 'ERROR', codigo: codigoError_(error)});
  } finally {
    if (bloqueo && bloqueo.hasLock()) {
      bloqueo.releaseLock();
    }
  }
}

function prepararInfraestructuraContingencia() {
  /*
   * Ejecutar una vez, después de desplegar el Web App y definir
  * UTP_SYNC_API_KEY en Project Settings > Script properties. Habilitar además
  * el servicio avanzado Google Sheets para que NO_ENTREGA sea atómico.
   */
  const apiKey = PropertiesService.getScriptProperties().getProperty('UTP_SYNC_API_KEY');
  if (!apiKey) throw new Error('CONFIGURAR_UTP_SYNC_API_KEY');
  const urlDetectada = ScriptApp.getService().getUrl();

  const libro = SpreadsheetApp.getActiveSpreadsheet();
  asegurarHojaEventos_(libro);
  const configuracion = asegurarHojaConfiguracion_(libro);
  const urlExistente = valorConfiguracion_(configuracion, 'URL_APPS_SCRIPT');
  const url = esUrlWebAppProduccion_(urlExistente)
    ? urlExistente
    : (esUrlWebAppProduccion_(urlDetectada) ? urlDetectada : '');
  if (!url) {
    throw new Error('CONFIGURAR_URL_APPS_SCRIPT_EXEC');
  }
  configuracion.getRange(1, 1, 1, 2).setValues([CABECERAS_CONFIGURACION]);
  configuracion.getRange(2, 1, 3, 2).setValues([
    ['URL_APPS_SCRIPT', url],
    ['API_KEY', apiKey],
    ['VERSION_CONFIG', '1']
  ]);
  configuracion.setFrozenRows(1);
  try { configuracion.hideSheet(); } catch (ignored) {}
  try {
    const proteccion = configuracion.protect();
    proteccion.setWarningOnly(true);
  } catch (ignored) {}
}

function esUrlWebAppProduccion_(url) {
  return /^https:\/\/script\.google\.com(?:\/a\/[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+)?\/macros\/s\/[A-Za-z0-9_-]+\/exec$/.test(String(url || ''));
}

function valorConfiguracion_(hoja, clave) {
  const filas = hoja.getDataRange().getValues();
  for (let indice = 1; indice < filas.length; indice++) {
    if (String(filas[indice][0] || '').trim().toUpperCase() === clave) {
      return String(filas[indice][1] || '').trim();
    }
  }
  return '';
}

function leerSolicitud_(e) {
  if (!e || !e.postData || !e.postData.contents) throw new Error('SOLICITUD_INVALIDA');
  try {
    const valor = JSON.parse(e.postData.contents);
    if (!valor || typeof valor !== 'object') throw new Error();
    return valor;
  } catch (ignored) {
    throw new Error('JSON_INVALIDO');
  }
}

function claveValida_(recibida) {
  const esperada = PropertiesService.getScriptProperties().getProperty('UTP_SYNC_API_KEY');
  return typeof recibida === 'string' && !!esperada && recibida === esperada;
}

function validarEvento_(evento) {
  const tipos = ['PRESTAMO', 'DEVOLUCION', 'NO_ENTREGA', 'CIERRE_AUTOMATICO'];
  if (!evento.id_evento || !/^[a-zA-Z0-9-]{8,100}$/.test(String(evento.id_evento))) {
    throw new Error('ID_EVENTO_INVALIDO');
  }
  if (tipos.indexOf(evento.tipo_evento) === -1) throw new Error('TIPO_EVENTO_INVALIDO');
  if (!evento.matricula || !evento.laptop_id) throw new Error('EVENTO_INCOMPLETO');
  if (!evento.datos || typeof evento.datos !== 'object') evento.datos = {};
}

function asegurarHojaEventos_(libro) {
  let hoja = libro.getSheetByName(HOJA_EVENTOS);
  if (!hoja) hoja = libro.insertSheet(HOJA_EVENTOS);
  if (hoja.getMaxColumns() < CABECERAS_EVENTOS.length) {
    hoja.insertColumnsAfter(hoja.getMaxColumns(), CABECERAS_EVENTOS.length - hoja.getMaxColumns());
  }
  const existentes = hoja.getRange(1, 1, 1, CABECERAS_EVENTOS.length).getValues()[0];
  if (existentes.join('|') !== CABECERAS_EVENTOS.join('|')) {
    hoja.getRange(1, 1, 1, CABECERAS_EVENTOS.length).setValues([CABECERAS_EVENTOS]);
  }
  hoja.setFrozenRows(1);
  return hoja;
}

function asegurarHojaConfiguracion_(libro) {
  let hoja = libro.getSheetByName(HOJA_CONFIGURACION);
  if (!hoja) hoja = libro.insertSheet(HOJA_CONFIGURACION);
  return hoja;
}

function buscarEvento_(hoja, idEvento) {
  const ultima = hoja.getLastRow();
  if (ultima < 2) return null;
  const filas = hoja.getRange(2, 1, ultima - 1, CABECERAS_EVENTOS.length).getValues();
  for (let indice = 0; indice < filas.length; indice++) {
    if (String(filas[indice][0]) === String(idEvento)) {
      return {
        fila: indice + 2, valores: filas[indice], estado: String(filas[indice][5]),
        referencia: String(filas[indice][8] || '')
      };
    }
  }
  return null;
}

function registrarPreparando_(hoja, evento) {
  const fecha = new Date().toISOString();
  const datosControl = JSON.stringify({
    nombre: String(evento.nombre || ''), datos: evento.datos || {},
    fecha_creacion: String(evento.fecha_creacion || '')
  });
  hoja.appendRow([
    evento.id_evento, evento.tipo_evento, evento.matricula, evento.laptop_id,
    fechaHoraEvento_(evento), 'PREPARANDO', fecha, '', '', datosControl,
    String(evento.equipo_origen || '')
  ]);
  return hoja.getLastRow();
}

function eventoGuardado_(encontrado, recibido) {
  let control = {};
  try { control = JSON.parse(encontrado.valores[9] || '{}'); } catch (ignored) {}
  return {
    id_evento: String(encontrado.valores[0]), tipo_evento: String(encontrado.valores[1]),
    matricula: String(encontrado.valores[2]), laptop_id: String(encontrado.valores[3]),
    fecha_evento: String(encontrado.valores[4]).substring(0, 10),
    hora_evento: String(encontrado.valores[4]).substring(11),
    nombre: String(control.nombre || recibido.nombre || ''),
    datos: control.datos || recibido.datos || {},
    equipo_origen: String(encontrado.valores[10] || '')
  };
}

function actualizarEvento_(hoja, fila, estado, referencia, codigo) {
  hoja.getRange(fila, 6, 1, 4).setValues([[
    estado, hoja.getRange(fila, 7).getValue() || new Date().toISOString(),
    estado === 'APLICADO' ? new Date().toISOString() : '', referencia || ''
  ]]);
  if (codigo) {
    let control = {};
    try { control = JSON.parse(hoja.getRange(fila, 10).getValue() || '{}'); } catch (ignored) {}
    control.ultimo_error = codigo;
    hoja.getRange(fila, 10).setValue(JSON.stringify(control));
  }
}

function aplicarEvento_(libro, evento) {
  const registros = libro.getSheetByName(HOJA_REGISTROS);
  const alumnos = libro.getSheetByName(HOJA_ALUMNOS);
  if (!registros || !alumnos) throw new Error('HOJA_OPERATIVA_NO_ENCONTRADA');
  const datos = evento.datos || {};
  if (evento.tipo_evento === 'PRESTAMO') {
    let fila = buscarRegistro_(registros, evento, datos);
    if (!fila) fila = insertarPrestamoCronologico_(registros, evento, datos);
    verificarPrestamo_(registros, fila, evento, datos);
    return 'Registros!' + fila;
  }
  const fila = buscarRegistro_(registros, evento, datos);
  if (!fila) throw new Error('DEPENDENCIA_PRESTAMO_NO_ENCONTRADA');
  if (evento.tipo_evento === 'DEVOLUCION') {
    registros.getRange(fila, 5).setValue(String(datos.hora_salida || evento.hora_evento));
    if (String(registros.getRange(fila, 5).getValue()) !== String(datos.hora_salida || evento.hora_evento)) {
      throw new Error('VERIFICACION_DEVOLUCION_FALLIDA');
    }
  } else if (evento.tipo_evento === 'NO_ENTREGA') {
    aplicarNoEntrega_(alumnos, registros, fila, evento);
  } else if (evento.tipo_evento === 'CIERRE_AUTOMATICO') {
    registros.getRange(fila, 5, 1, 3).setValues([[
      String(datos.hora_salida || evento.hora_evento), '', 'CIERRE_AUTOMATICO_POR_NUEVA_SESION'
    ]]);
    if (String(registros.getRange(fila, 7).getValue()) !== 'CIERRE_AUTOMATICO_POR_NUEVA_SESION') {
      throw new Error('VERIFICACION_CIERRE_FALLIDA');
    }
  }
  return 'Registros!' + fila;
}

function buscarRegistro_(hoja, evento, datos) {
  const filas = hoja.getDataRange().getValues();
  const fecha = String(datos.fecha || evento.fecha_evento || '');
  const hora = String(datos.hora_ingreso || '');
  const nombre = String(datos.nombre || evento.nombre || '');
  for (let indice = 1; indice < filas.length; indice++) {
    const fila = filas[indice];
    if (String(fila[0]) === String(evento.matricula) && String(fila[5]) === String(evento.laptop_id) &&
        String(fila[2]) === fecha && String(fila[3]) === hora && (!nombre || String(fila[1]) === nombre)) {
      return indice + 1;
    }
  }
  return null;
}

function insertarPrestamoCronologico_(hoja, evento, datos) {
  const fecha = String(datos.fecha || evento.fecha_evento || '');
  const hora = String(datos.hora_ingreso || evento.hora_evento || '');
  const nuevaClave = fecha + 'T' + hora;
  const filas = hoja.getDataRange().getValues();
  let destino = hoja.getLastRow() + 1;
  for (let indice = 1; indice < filas.length; indice++) {
    const clave = String(filas[indice][2] || '') + 'T' + String(filas[indice][3] || '');
    if (clave > nuevaClave) { destino = indice + 1; break; }
  }
  if (destino <= hoja.getLastRow()) hoja.insertRowBefore(destino);
  hoja.getRange(destino, 1, 1, 7).setValues([[
    evento.matricula, String(datos.nombre || evento.nombre || ''), fecha, hora, '',
    evento.laptop_id, 'S/N'
  ]]);
  return destino;
}

function verificarPrestamo_(hoja, fila, evento, datos) {
  const registro = hoja.getRange(fila, 1, 1, 7).getValues()[0];
  if (String(registro[0]) !== String(evento.matricula) || String(registro[5]) !== String(evento.laptop_id) ||
      String(registro[2]) !== String(datos.fecha || evento.fecha_evento || '') ||
      String(registro[3]) !== String(datos.hora_ingreso || evento.hora_evento || '')) {
    throw new Error('VERIFICACION_PRESTAMO_FALLIDA');
  }
}

function aplicarNoEntrega_(alumnos, registros, filaRegistro, evento) {
  const observacion = String(registros.getRange(filaRegistro, 7).getValue() || '');
  if (observacion === 'NO_ENTREGA_CONTADA') return;
  const matriculas = alumnos.getRange(1, 1, alumnos.getLastRow(), 1).getValues();
  let filaAlumno = null;
  for (let indice = 0; indice < matriculas.length; indice++) {
    if (String(matriculas[indice][0]) === String(evento.matricula)) { filaAlumno = indice + 1; break; }
  }
  if (!filaAlumno) throw new Error('ALUMNO_NO_ENCONTRADO');
  const actual = Number(alumnos.getRange(filaAlumno, 5).getValue() || 0);
  if (!isFinite(actual)) throw new Error('NO_ENTREGAS_INVALIDO');
  // Un solo request atómico para ambas hojas. Requiere el servicio avanzado
  // Google Sheets (identificador global `Sheets`) habilitado en Apps Script.
  Sheets.Spreadsheets.Values.batchUpdate({
    valueInputOption: 'RAW',
    data: [
      {
        range: HOJA_REGISTROS + '!E' + filaRegistro + ':G' + filaRegistro,
        values: [['-----', '', 'NO_ENTREGA_CONTADA']]
      },
      {
        range: HOJA_ALUMNOS + '!E' + filaAlumno,
        values: [[actual + 1]]
      }
    ]
  }, SpreadsheetApp.getActiveSpreadsheet().getId());
  if (String(registros.getRange(filaRegistro, 7).getValue()) !== 'NO_ENTREGA_CONTADA' ||
      Number(alumnos.getRange(filaAlumno, 5).getValue()) !== actual + 1) {
    throw new Error('VERIFICACION_NO_ENTREGA_FALLIDA');
  }
}

function fechaHoraEvento_(evento) {
  return String(evento.fecha_evento || '') + ' ' + String(evento.hora_evento || '');
}

function codigoError_(error) {
  const codigo = String(error && error.message || 'ERROR_INTERNO');
  return /^[A-Z0-9_:-]{1,120}$/.test(codigo) ? codigo : 'ERROR_INTERNO';
}

function responder_(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
