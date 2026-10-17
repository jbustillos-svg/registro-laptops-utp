/* Pruebas con fake de Sheets: no usa Google ni modifica datos reales. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const fuente = fs.readFileSync('apps_script_sincronizacion.gs', 'utf8');
const CABECERA = ['MATRICULA', 'NOMBRE', 'FECHA', 'HORA', 'SALIDA', 'LAPTOP', 'OBSERVACION'];

function celdaAColumna(letra) {
  return letra.charCodeAt(0) - 64;
}

class HojaFalsa {
  constructor(filas, opciones = {}) {
    this.filas = filas.map((fila) => fila.slice());
    this.fallosLecturaTrasEscritura = opciones.fallosLecturaTrasEscritura || 0;
  }

  getDataRange() {
    return {getValues: () => {
      if (this.fallosLecturaTrasEscritura > 0) {
        this.fallosLecturaTrasEscritura -= 1;
        return [this.filas[0].slice()];
      }
      return this.filas.map((fila) => fila.slice());
    }};
  }

  getLastRow() { return this.filas.length; }

  insertRowBefore(fila) {
    this.filas.splice(fila - 1, 0, Array(7).fill(''));
  }

  getRange(fila, columna, alto = 1, ancho = 1) {
    return {
      getValues: () => Array.from({length: alto}, (_, y) => Array.from(
        {length: ancho}, (_, x) => this.filas[fila - 1 + y]?.[columna - 1 + x] ?? ''
      )),
      setValues: (valores) => valores.forEach((renglon, y) => renglon.forEach((valor, x) => {
        while (this.filas.length < fila + y) this.filas.push(Array(7).fill(''));
        this.filas[fila - 1 + y][columna - 1 + x] = valor;
      })),
      getValue: () => this.filas[fila - 1]?.[columna - 1] ?? '',
      setValue: (valor) => { this.filas[fila - 1][columna - 1] = valor; },
    };
  }
}

function crearEntorno(registrosFilas = [CABECERA], opciones = {}) {
  const registros = new HojaFalsa(registrosFilas, opciones);
  const alumnos = new HojaFalsa([['MATRICULA', '', '', '', 'NO_ENTREGAS'], ['824118001', '', '', '', 0]]);
  const libro = {
    getSheetByName: (nombre) => ({Registros: registros, Alumnos: alumnos}[nombre] || null),
    getId: () => 'libro-falso',
  };
  const logs = [];
  const contexto = {
    Logger: {log: (mensaje) => logs.push(mensaje)},
    Utilities: {formatDate: (fecha, _zona, patron) => {
      const iso = fecha.toISOString();
      return patron === 'yyyy-MM-dd' ? iso.slice(0, 10) : iso.slice(11, 19);
    }},
    Session: {getScriptTimeZone: () => 'UTC'},
    SpreadsheetApp: {getActiveSpreadsheet: () => libro},
    Sheets: {Spreadsheets: {Values: {batchUpdate: (peticion) => {
      peticion.data.forEach((dato) => {
        const coincidencia = dato.range.match(/^(Registros|Alumnos)!([A-Z])(\d+)(?::([A-Z])\d+)?$/);
        const hoja = coincidencia[1] === 'Registros' ? registros : alumnos;
        const inicio = celdaAColumna(coincidencia[2]);
        hoja.getRange(Number(coincidencia[3]), inicio, 1, dato.values[0].length).setValues(dato.values);
      });
    }}}},
    console,
  };
  vm.createContext(contexto);
  vm.runInContext(fuente, contexto);
  return {contexto, libro, registros, logs, aplicar: (evento) => contexto.aplicarEvento_(libro, evento)};
}

function prestamo(datos = {}) {
  return {
    id_evento: '949c0931-d4d3-4289-add7-0dc82e2f16e8', tipo_evento: 'PRESTAMO',
    matricula: '824118001', nombre: 'Lizeth Bustillos Morales',
    fecha_evento: '2026-10-16', hora_evento: '18:12:37', laptop_id: 'AulaMovil24-B',
    datos: {fecha: '2026-10-16', hora_ingreso: '18:12:37', nombre: 'Lizeth Bustillos Morales', ...datos},
  };
}

function filaPrestamo(matricula = '824118001', fecha = '2026-10-16', hora = '18:12:37') {
  return [matricula, 'Lizeth Bustillos Morales', fecha, hora, '', 'AulaMovil24-B', 'S/N'];
}

// 1: préstamo normal; 5: cinco reintentos no agregan filas.
let caso = crearEntorno();
assert.equal(caso.aplicar(prestamo()), 'Registros!2');
for (let intento = 0; intento < 5; intento++) assert.equal(caso.aplicar(prestamo()), 'Registros!2');
assert.equal(caso.registros.filas.length, 2, 'un UUID reintentado conserva una sola fila');

// 2 y 11: escritura aplicada pero dos lecturas fallan; el reintento reconcilia.
caso = crearEntorno([CABECERA], {fallosLecturaTrasEscritura: 0});
const originalSetValues = caso.registros.getRange.bind(caso.registros);
let activarFallo = true;
caso.registros.getRange = (...args) => {
  const rango = originalSetValues(...args);
  const original = rango.setValues;
  rango.setValues = (valores) => { original(valores); if (activarFallo) caso.registros.fallosLecturaTrasEscritura = 2; };
  return rango;
};
assert.throws(() => caso.aplicar(prestamo()), /VERIFICACION_PRESTAMO_FALLIDA/);
activarFallo = false;
assert.equal(caso.aplicar(prestamo()), 'Registros!2');
assert.equal(caso.registros.filas.length, 2, 'un timeout/error posterior no inserta segunda fila');

// 3, 4 y 10: PREPARANDO/ERROR se reconcilian con la fila y quedan listos para APLICADO.
for (const estado of ['PREPARANDO', 'ERROR']) {
  caso = crearEntorno([CABECERA, filaPrestamo()]);
  const evento = prestamo({estado_remoto: estado});
  assert.equal(caso.aplicar(evento), 'Registros!2', `${estado} debe reconciliar`);
  assert.equal(caso.registros.filas.length, 2);
}

// 10: el flujo real de doPost convierte el PRESTAMO reconciliado en APLICADO.
caso = crearEntorno([CABECERA, filaPrestamo()]);
let actualizacion = null;
caso.contexto.LockService = {getScriptLock: () => ({tryLock: () => true, hasLock: () => true, releaseLock: () => {}})};
caso.contexto.leerSolicitud_ = () => prestamo();
caso.contexto.claveValida_ = () => true;
caso.contexto.asegurarHojaEventos_ = () => ({});
caso.contexto.buscarEvento_ = () => ({estado: 'ERROR', fila: 2, valores: []});
caso.contexto.eventoGuardado_ = () => prestamo();
caso.contexto.actualizarEvento_ = (_hoja, _fila, estado, referencia) => { actualizacion = {estado, referencia}; };
caso.contexto.responder_ = (respuesta) => respuesta;
const respuestaReconciliada = caso.contexto.doPost({});
assert.equal(respuestaReconciliada.estado, 'APLICADO');
assert.deepEqual(actualizacion, {estado: 'APLICADO', referencia: 'Registros!2'});

// 6 y 7: número/texto, Date y fracción horaria de Sheets son compatibles.
caso = crearEntorno([CABECERA, filaPrestamo(824118001, new Date('2026-10-16T00:00:00Z'), (18 * 3600 + 12 * 60 + 37) / 86400)]);
assert.equal(caso.aplicar(prestamo()), 'Registros!2');
assert.equal(caso.registros.filas.length, 2);

// 8: insertar cronológicamente desplaza filas y se vuelve a localizar la creada.
caso = crearEntorno([CABECERA, filaPrestamo('999', '2026-10-17', '08:00:00')]);
assert.equal(caso.aplicar(prestamo()), 'Registros!2');
assert.equal(caso.registros.filas.length, 3);
assert.equal(caso.registros.filas[1][0], '824118001');

// 9: duplicados históricos no se borran ni provocan un sexto inserto; se registra diagnóstico.
caso = crearEntorno([CABECERA, filaPrestamo(), filaPrestamo()]);
assert.equal(caso.aplicar(prestamo()), 'Registros!2');
assert.equal(caso.registros.filas.length, 3);
assert.ok(caso.logs.some((log) => log.includes('DUPLICADOS_PREEXISTENTES')));

// 12: siguen funcionando DEVOLUCION, CIERRE_AUTOMATICO y NO_ENTREGA.
caso = crearEntorno([CABECERA, filaPrestamo()]);
const base = prestamo();
assert.equal(caso.aplicar({...base, tipo_evento: 'DEVOLUCION', datos: {...base.datos, hora_salida: '19:00:00'}}), 'Registros!2');
assert.equal(caso.registros.filas[1][4], '19:00:00');
assert.equal(caso.aplicar({...base, tipo_evento: 'CIERRE_AUTOMATICO'}), 'Registros!2');
assert.equal(caso.registros.filas[1][6], 'CIERRE_AUTOMATICO_POR_NUEVA_SESION');
caso = crearEntorno([CABECERA, filaPrestamo()]);
assert.equal(caso.aplicar({...base, tipo_evento: 'NO_ENTREGA'}), 'Registros!2');
assert.equal(caso.registros.filas[1][6], 'NO_ENTREGA_CONTADA');

console.log('Pruebas de idempotencia Apps Script: OK');
