/* Prueba local de prepararInfraestructuraContingencia; no llama Google real. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const fuente = fs.readFileSync('apps_script_sincronizacion.gs', 'utf8');

function ejecutarPreparacion(urlDetectada, filasIniciales) {
  const escrituras = [];
  const hoja = {
    getDataRange: () => ({getValues: () => filasIniciales}),
    getRange: () => ({setValues: (valores) => escrituras.push(valores)}),
    setFrozenRows: () => {},
    hideSheet: () => {},
    protect: () => ({setWarningOnly: () => {}}),
  };
  const contexto = {
    PropertiesService: {getScriptProperties: () => ({getProperty: () => 'clave-de-prueba'})},
    ScriptApp: {getService: () => ({getUrl: () => urlDetectada})},
    SpreadsheetApp: {getActiveSpreadsheet: () => ({})},
    console,
  };
  vm.createContext(contexto);
  vm.runInContext(fuente, contexto);
  contexto.asegurarHojaEventos_ = () => {};
  contexto.asegurarHojaConfiguracion_ = () => hoja;
  return {ejecutar: () => contexto.prepararInfraestructuraContingencia(), escrituras};
}

const existenteExec = 'https://script.google.com/a/utparral.edu.mx/macros/s/existente_123/exec';
let caso = ejecutarPreparacion(
  'https://script.google.com/macros/s/detectada/dev',
  [['CLAVE', 'VALOR'], ['URL_APPS_SCRIPT', existenteExec]],
);
caso.ejecutar();
assert.equal(caso.escrituras[1][0][1], existenteExec, 'debe conservar /exec existente');

caso = ejecutarPreparacion(
  'https://script.google.com/macros/s/detectada/dev',
  [['CLAVE', 'VALOR']],
);
assert.throws(caso.ejecutar, /CONFIGURAR_URL_APPS_SCRIPT_EXEC/);
assert.equal(caso.escrituras.length, 0, 'nunca debe escribir /dev');

const detectadaExec = 'https://script.google.com/macros/s/detectada_123/exec';
caso = ejecutarPreparacion(detectadaExec, [['CLAVE', 'VALOR']]);
caso.ejecutar();
assert.equal(caso.escrituras[1][0][1], detectadaExec, 'debe usar /exec detectada');

console.log('Pruebas de preparación Apps Script: OK');
