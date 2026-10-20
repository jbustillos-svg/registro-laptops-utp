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
  asegurarFormatoRegistros_(libro.getSheetByName(HOJA_REGISTROS));
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
  // La bitácora conserva la misma identidad literal que Registros; un reintento
  // nunca debe reconstruir una matrícula numérica sin sus ceros iniciales.
  hoja.getRange(1, 3, Math.max(hoja.getLastRow(), 1), 1).setNumberFormat('@');
  hoja.setFrozenRows(1);
  return hoja;
}

function asegurarHojaConfiguracion_(libro) {
  let hoja = libro.getSheetByName(HOJA_CONFIGURACION);
  if (!hoja) hoja = libro.insertSheet(HOJA_CONFIGURACION);
  return hoja;
}

function asegurarFormatoRegistros_(hoja) {
  // A:G se presenta a la izquierda. Sólo A se fuerza como texto: fecha y horas
  // conservan sus tipos y formatos existentes.
  if (!hoja) return;
  const filas = Math.max(hoja.getLastRow(), 1);
  hoja.getRange(1, 1, filas, 7).setHorizontalAlignment('left');
  hoja.getRange(1, 1, filas, 1).setNumberFormat('@');
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
  // La bitácora también conserva el identificador literal para que un reintento
  // reconstruya exactamente el mismo evento, aun si Registros aún no existe.
  hoja.getRange(hoja.getLastRow() + 1, 3).setNumberFormat('@');
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
    matricula: String(control.datos.matricula || encontrado.valores[2]),
    laptop_id: String(control.datos.laptop_id || encontrado.valores[3]),
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
  asegurarFormatoRegistros_(registros);
  const datos = evento.datos || {};
  if (evento.tipo_evento === 'PRESTAMO') {
    // PREPARANDO y ERROR son estados ambiguos: antes de insertar se reconcilian
    // contra Registros. Esto hace idempotente un reintento tras escritura aplicada.
    let coincidencias = buscarRegistrosPrestamoCompatibles_(registros, evento, datos);
    if (!coincidencias.length) {
      insertarPrestamoCronologico_(registros, evento, datos);
      // La inserción cronológica puede cambiar números de fila; se busca de nuevo.
      coincidencias = buscarRegistrosPrestamoCompatibles_(registros, evento, datos);
    }
    if (!coincidencias.length) {
      verificarPrestamo_(registros, null, evento, datos);
    }
    if (coincidencias.length > 1) {
      registrarDiagnosticoPrestamo_(
        'DUPLICADOS_PREEXISTENTES', evento, coincidencias[0].fila, coincidencias.length
      );
    }
    const fila = verificarPrestamo_(registros, coincidencias[0].fila, evento, datos);
    return 'Registros!' + fila;
  }
  const fila = buscarRegistro_(libro, registros, evento, datos);
  if (!fila) throw new Error('DEPENDENCIA_PRESTAMO_NO_ENCONTRADA');
  if (evento.tipo_evento === 'DEVOLUCION') {
    const horaSalidaOriginal = String(datos.hora_salida || evento.hora_evento);
    const celdaSalida = registros.getRange(fila, 5);
    // Si una escritura anterior alcanzó Sheets pero falló la lectura/verificación,
    // se reconoce como aplicada sin reemplazar la hora original de la entrega.
    if (!horaSalidaCoincide_(celdaSalida.getValue(), horaSalidaOriginal)) {
      celdaSalida.setValue(horaSalidaOriginal);
    }
    if (!horaSalidaCoincide_(celdaSalida.getValue(), horaSalidaOriginal)) {
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

function buscarRegistro_(libro, hoja, evento, datos) {
  const coincidencias = buscarRegistrosPrestamoCompatibles_(hoja, evento, datos);
  if (coincidencias.length) return coincidencias[0].fila;

  // Las sesiones v1.4.1 llevan id_sesion. Cuando existe la referencia de su
  // PRESTAMO, se usa sólo si sigue señalando una fila compatible.
  const filaPorSesion = buscarRegistroPorSesion_(libro, hoja, evento, datos);
  if (filaPorSesion) return filaPorSesion;

  // Recupera filas antiguas en las que Sheets quitó exclusivamente ceros
  // iniciales de una matrícula numérica. Se exige el resto de la identidad del
  // préstamo y una única coincidencia para no tocar otro préstamo.
  const recuperadas = buscarRegistrosPrestamoCompatibles_(hoja, evento, datos, true);
  if (recuperadas.length === 1) {
    Logger.log(
      'MATRICULA_CERO_INICIAL_RECUPERADA id=' + String(evento.id_evento || '') +
      ' referencia=Registros!' + recuperadas[0].fila
    );
    return recuperadas[0].fila;
  }
  if (recuperadas.length > 1) {
    Logger.log(
      'DEPENDENCIA_PRESTAMO_AMBIGUA id=' + String(evento.id_evento || '') +
      ' coincidencias=' + recuperadas.length
    );
  }
  return null;
}

function identidadPrestamo_(evento, datos) {
  return {
    matricula: normalizarTexto_(evento.matricula),
    nombre: normalizarTexto_(datos.nombre || evento.nombre || ''),
    fecha: normalizarFecha_(datos.fecha || evento.fecha_evento || ''),
    hora: normalizarHora_(datos.hora_ingreso || evento.hora_evento || ''),
    laptopId: normalizarTexto_(evento.laptop_id)
  };
}

function buscarRegistrosPrestamoCompatibles_(hoja, evento, datos, tolerarCeroInicial) {
  const identidad = identidadPrestamo_(evento, datos);
  const filas = hoja.getDataRange().getValues();
  const coincidencias = [];
  for (let indice = 1; indice < filas.length; indice++) {
    const fila = filas[indice];
    if (registroPrestamoCompatible_(fila, identidad, tolerarCeroInicial)) {
      coincidencias.push({fila: indice + 1, valores: fila});
    }
  }
  return coincidencias;
}

function registroPrestamoCompatible_(registro, identidad, tolerarCeroInicial) {
  const matriculaRegistro = normalizarTexto_(registro[0]);
  const matriculaCompatible = matriculaRegistro === identidad.matricula ||
    (tolerarCeroInicial && matriculaEquivalentePorCeroInicial_(
      matriculaRegistro, identidad.matricula
    ));
  return matriculaCompatible &&
    normalizarTexto_(registro[5]) === identidad.laptopId &&
    normalizarFecha_(registro[2]) === identidad.fecha &&
    normalizarHora_(registro[3]) === identidad.hora &&
    (!identidad.nombre || normalizarTexto_(registro[1]) === identidad.nombre);
}

function insertarPrestamoCronologico_(hoja, evento, datos) {
  const identidad = identidadPrestamo_(evento, datos);
  const nuevaClave = identidad.fecha + 'T' + identidad.hora;
  const filas = hoja.getDataRange().getValues();
  let destino = hoja.getLastRow() + 1;
  for (let indice = 1; indice < filas.length; indice++) {
    const clave = normalizarFecha_(filas[indice][2]) + 'T' + normalizarHora_(filas[indice][3]);
    if (clave > nuevaClave) { destino = indice + 1; break; }
  }
  if (destino <= hoja.getLastRow()) hoja.insertRowBefore(destino);
  hoja.getRange(destino, 1, 1, 7).setHorizontalAlignment('left');
  hoja.getRange(destino, 1).setNumberFormat('@');
  hoja.getRange(destino, 1, 1, 7).setValues([[
    identidad.matricula, identidad.nombre, identidad.fecha, identidad.hora, '',
    identidad.laptopId, 'S/N'
  ]]);
  return destino;
}

function verificarPrestamo_(hoja, fila, evento, datos) {
  const identidad = identidadPrestamo_(evento, datos);
  let registro = null;
  if (fila) {
    registro = hoja.getRange(fila, 1, 1, 7).getValues()[0];
    if (registroPrestamoCompatible_(registro, identidad)) return fila;
  }
  // Releer toda la hoja permite recuperarse de un desplazamiento de fila o de
  // una representación Date/número que cambió entre la escritura y la lectura.
  const coincidencias = buscarRegistrosPrestamoCompatibles_(hoja, evento, datos);
  if (coincidencias.length) return coincidencias[0].fila;
  registrarDiagnosticoPrestamo_('VERIFICACION_PRESTAMO_FALLIDA', evento, fila, 0, registro);
  throw new Error('VERIFICACION_PRESTAMO_FALLIDA');
}

function normalizarTexto_(valor) {
  return String(valor === null || valor === undefined ? '' : valor).trim();
}

function matriculaEquivalentePorCeroInicial_(izquierda, derecha) {
  const valorIzquierdo = normalizarTexto_(izquierda);
  const valorDerecho = normalizarTexto_(derecha);
  if (!/^\d+$/.test(valorIzquierdo) || !/^\d+$/.test(valorDerecho)) return false;
  if (valorIzquierdo.charAt(0) !== '0' && valorDerecho.charAt(0) !== '0') return false;
  return valorIzquierdo.replace(/^0+(?=\d)/, '') ===
    valorDerecho.replace(/^0+(?=\d)/, '');
}

function horaSalidaCoincide_(valorLeido, horaOriginal) {
  return normalizarHora_(valorLeido) === normalizarHora_(horaOriginal);
}

function buscarRegistroPorSesion_(libro, registros, evento, datos) {
  const idSesion = normalizarTexto_(datos.id_sesion);
  const idPrestamo = normalizarTexto_(datos.prestamo_evento_id);
  if (!idSesion && !idPrestamo) return null;
  const hojaEventos = libro.getSheetByName(HOJA_EVENTOS);
  if (!hojaEventos || hojaEventos.getLastRow() < 2) return null;
  const filas = hojaEventos.getDataRange().getValues();
  const identidad = identidadPrestamo_(evento, datos);
  for (let indice = 1; indice < filas.length; indice++) {
    const filaEvento = filas[indice];
    if (String(filaEvento[1]) !== 'PRESTAMO') continue;
    if (idPrestamo && normalizarTexto_(filaEvento[0]) !== idPrestamo) continue;
    let control = {};
    try { control = JSON.parse(filaEvento[9] || '{}'); } catch (ignored) {}
    if (idSesion && normalizarTexto_(control.datos && control.datos.id_sesion) !== idSesion) continue;
    const referencia = String(filaEvento[8] || '');
    const coincidencia = referencia.match(/^Registros!(\d+)$/);
    if (!coincidencia) continue;
    const filaRegistro = Number(coincidencia[1]);
    if (filaRegistro < 2 || filaRegistro > registros.getLastRow()) continue;
    const registro = registros.getRange(filaRegistro, 1, 1, 7).getValues()[0];
    if (registroPrestamoCompatible_(registro, identidad, true)) return filaRegistro;
  }
  return null;
}

function esFecha_(valor) {
  return Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime());
}

function normalizarFecha_(valor) {
  if (esFecha_(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  // getValues() puede entregar una fecha de Sheets como serial. Se convierte
  // sólo para comparar identidades; la celda original no se reescribe ni se
  // convierte a texto.
  if (typeof valor === 'number' && isFinite(valor) && valor > 0) {
    const fechaSerial = new Date(Date.UTC(1899, 11, 30) + Math.floor(valor) * 86400000);
    return Utilities.formatDate(fechaSerial, 'Etc/UTC', 'yyyy-MM-dd');
  }
  const texto = normalizarTexto_(valor);
  const coincidencia = texto.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (coincidencia) {
    return coincidencia[1] + '-' + ('0' + coincidencia[2]).slice(-2) + '-' + ('0' + coincidencia[3]).slice(-2);
  }
  return texto;
}

function normalizarHora_(valor) {
  if (esFecha_(valor)) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'HH:mm:ss');
  }
  if (typeof valor === 'number' && isFinite(valor) && valor >= 0 && valor < 1) {
    const segundos = Math.round(valor * 86400) % 86400;
    return ('0' + Math.floor(segundos / 3600)).slice(-2) + ':' +
      ('0' + Math.floor((segundos % 3600) / 60)).slice(-2) + ':' +
      ('0' + (segundos % 60)).slice(-2);
  }
  const texto = normalizarTexto_(valor);
  const coincidencia = texto.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (coincidencia) {
    return ('0' + coincidencia[1]).slice(-2) + ':' + coincidencia[2] + ':' + (coincidencia[3] || '00');
  }
  return texto;
}

function registrarDiagnosticoPrestamo_(codigo, evento, fila, coincidencias, registro) {
  const identidad = identidadPrestamo_(evento, evento.datos || {});
  const esperado = {
    matricula: identidad.matricula, fecha: identidad.fecha,
    hora: identidad.hora, laptop_id: identidad.laptopId
  };
  const leido = registro ? {
    matricula: normalizarTexto_(registro[0]), fecha: normalizarFecha_(registro[2]),
    hora: normalizarHora_(registro[3]), laptop_id: normalizarTexto_(registro[5])
  } : {};
  Logger.log(codigo + ' id=' + String(evento.id_evento || '') +
    ' referencia=' + (fila ? 'Registros!' + fila : 'SIN_REFERENCIA') +
    ' coincidencias=' + Number(coincidencias || 0) +
    ' esperado=' + JSON.stringify(esperado) + ' leido=' + JSON.stringify(leido));
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
