/**
 * Parametro global que se utiliza para contar filas, toma el valor de la primera fila que se encuentre vacia en la celda que se quiere escribir.
 */
let fila = 2


const VACIO = "";
const CELDAS_TELEFONOS = "AG"
const CELDAS_MEMBERS = "AC"
const CELDAS_NAME = "A"
const CELDAS_GIVEN_NAME = "B"
const MEMBERS_VALUE = "* myContacts"

const SELECCION_TELEFONOS = "T11:T150";
const SELECCION_PERSONAS = "C11:C150";
const GUION = "-";
const ESPACIO = " ";
const PARENTESIS_IZQ = "(";
const PARENTESIS_DER = ")";
const STR_54 = "54";
const STR_261 = "261";
const TYPE_NUMBER = "number"
const EXCLAMATION = "!"
const PLUS_SIGN = "+"



/**
 * URL del documento a analizar e ir leyendo hoja por hoja para normalizar los telefonos
 */
const URL_DOCUMENTO_ANALIZAR = "https://docs.google.com/spreadsheets/d/1Mqf3kyq1bPOqPk2CVSGxAlw9KZuH_dG4zFoe4eW1AUo/edit#gid=228148124"

let primera;
let segunda;
let tercera;
let telefono;
let persona;
let nuevoTelefono;

/**
 * Url del documento en el que se va a escribir el csv para importar a contactos de google.
 */
const URL_CSV_SALIDA = "https://docs.google.com/spreadsheets/d/1ieVLgszeTa9UZ1g9QsYg3DSHNSWC3e7aSDS5SuB1Eyw/edit#gid=1964355463"
/**
 * Documento en el que se va a escribir el csv para importar a contactos de google.
 */
const documentoContactos = SpreadsheetApp.openByUrl(URL_CSV_SALIDA);

/**
 * Analiza todas las hojas del documento a analizar
 */
function myFunction() {

  let doc1 = SpreadsheetApp.openByUrl(URL_DOCUMENTO_ANALIZAR);


  const HOJAS = doc1.getSheets().length;

  for (var i = 1; i < HOJAS + 1; i++) {

    // Obtener valores de columna de telefonos en la hoja i
    let telefonos = doc1.getSheets()[i].getRange(SELECCION_TELEFONOS).getValues();
    // Obtener valores de columna personas para poder iterarlos en la hoja i
    let personas = doc1.getSheets()[i].getRange(SELECCION_PERSONAS).getValues();

    let paresDeDatos = [];

    paresDeDatos = armadorDePares(telefonos, personas);

    telefonos = distribuidorTelefonos(paresDeDatos)
    personas = distribuidorPersonas(paresDeDatos)


    procesarHoja(telefonos, personas)

  }

  /* 
  * Puntos siguientes
    Expandir para que tome todas las hojas del Google Sheet -- actualmente me choco con el limite de ejecucion de 6 mins
    Google Workspace accounts, because you are paying a monthly fee to Google per user, the timeout limit is more generous at 30 minutes.
    Expandir para que lea todos los Sheets de una carpeta -- no es dificil pero sin superar el limite de ejecucion de 6 mins we re fucked
   */

}

/**
 * Analiza por el nombre de hoja provisto del documento a analizar
 */
function myFunction2_hojaPorHoja() {

  //url del documento a analizar
  let doc1 = SpreadsheetApp.openByUrl(URL_DOCUMENTO_ANALIZAR);

  // Nombre de la Hoja a escanear del documento -> doc1
  let HOJA = "1"

  // Te deje esto aca por si queres pasar hoja a hoja
  let telefonos = doc1.getRange(HOJA + EXCLAMATION + SELECCION_TELEFONOS).getValues();

  // Te deje esto aca por si queres pasar hoja a hoja
  let personas = doc1.getRange(HOJA + EXCLAMATION + SELECCION_PERSONAS).getValues();

  let paresDeDatos = [];

  paresDeDatos = armadorDePares(telefonos, personas);

  telefonos = distribuidorTelefonos(paresDeDatos)
  personas = distribuidorPersonas(paresDeDatos)

  procesarHoja(telefonos, personas)

}

/**
 *  Concatena 1 a 1 valores de telefono y persona cuando ambos no sean vacios
 *  Devuelve el array resultante
 */
function armadorDePares(telefonos, personas) {

  let paresDeDatos = [];

  telefonos = telefonos.flat();
  personas = personas.flat();


  for (var i = 0; i < telefonos.length; i++) {
    if (telefonos[i] !== VACIO && personas[i] !== VACIO) {
      paresDeDatos = paresDeDatos.concat([personas[i], telefonos[i]])
    }
  }

  return paresDeDatos;
}
/**
 * Divide el array [Pares de Datos] por todos sus indices impares
 *  Devuelve un array de todos los indices impares -> telefonos
 */
function distribuidorTelefonos(paresDeDatos) {
  let odd = [];

  for (var i = 0; i < paresDeDatos.length; ++i) {
    if ((i % 2) === 0) {}else {
      odd.push(paresDeDatos[i]);
    }
  };
  return odd;


}

/**
 *  Divide el array [Pares de Datos] por todos sus indices pares
 *  Devuelve un array de todos los indices pares -> personas
 */
function distribuidorPersonas(paresDeDatos) {

  let even = [];

  for (var i = 0; i < paresDeDatos.length; ++i) {
    if ((i % 2) === 0) {
      even.push(paresDeDatos[i]);
    }
  };
  return even;

}


/**
 * Returns a flattend and clean array
 * Validates if the entry value is a number or not
 */
function limpiarArray(arr) {

  arr = arr.flat();

  arr = arr.filter(entry => {

    if (typeof (entry) === TYPE_NUMBER) {
      entry = entry.toString()
    }
    return entry.trim() != VACIO;
  });

  return arr

}

/**
 * Obtiene una celda activa de la primera hoja del documentoContactos
 * @celdas toma el valor de una columna para escribir 
 * Utiliza el parametro global fila para determinar la posicion en la columna a escribir. 
 * Selecciona esa celda con el cursor activo y la devuelve como objeto.
 */
function getCurrentCellComplete(celdas) {

  currentCell = documentoContactos.getSheets()[0].getRange(celdas + fila)

  documentoContactos.getActiveSheet().setCurrentCell(currentCell)

  selection = documentoContactos.getSelection();

  currentCell = selection.getCurrentCell();

  return currentCell;

}

/**
 * Escribe valores en el excel para la celda designada de los telefonos
 */
function escribirNumerosTelefonoExcel(nuevoTelefono) {


  let cellValue = getCurrentCellComplete(CELDAS_TELEFONOS).getValue();

  if (cellValue == VACIO) {

    getCurrentCellComplete(CELDAS_TELEFONOS).setValue(nuevoTelefono);

  } else if (cellValue !== VACIO) {

    fila = fila + 1

    getCurrentCellComplete(CELDAS_TELEFONOS).setValue(nuevoTelefono);

  }
}

/**
 * Escribe valores en el excel. 
 * @celda para la celda en que se va a escribir.
 * Utiliza el parametro global fila para determinar la ultima posicion vacia en la columna a escribir. 
 * @valor para el valor a insertarse en la celda 
 */
function escribirValoresExcel(celda, valor) {

  let cellValue = getCurrentCellComplete(celda).getValue();

  if (cellValue == VACIO) {

    getCurrentCellComplete(celda).setValue(valor);

  } else if (cellValue !== VACIO) {

    /**
     * Calcular la siguiente fila para una celda distinta de vacia 
     * */
    while(cellValue !== VACIO) {
      fila = fila + 1
      cellValue = getCurrentCellComplete(celda).getValue();
    }

    getCurrentCellComplete(celda).setValue(valor);

  }

}

function normalizarTelefonos(telefono) {

  primera = VACIO;
  segunda = VACIO;
  tercera = VACIO;

  /**
    * Validacion para [+54 2611234567]
    */
  if (telefono.length === 15 && telefono.substring(0, 1) === PLUS_SIGN) {
    telefono = telefono.substring(4)

    primera = telefono.substring(-1, 4).trim();
    segunda = telefono.substring(4, 7).trim();
    tercera = telefono.substring(7, telefono.length).trim();

    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;
  }
  /**
   * Validacion para parentesis con guion [(261) 123-4567]
   */
  if (telefono.length === 14 && telefono.substring(0, 1) === PARENTESIS_IZQ) {


    // borrar primer parentesis 
    telefono = telefono.replace(PARENTESIS_IZQ, VACIO);
    // borrar segundo parentesis ()
    telefono = telefono.replace(PARENTESIS_DER, VACIO);

    nuevoTelefono = telefono;

  }

  /**
   * Validacion para  [+54 2612762107]
   */
  if (telefono.length === 14 && telefono.substring(0, 1) === PLUS_SIGN) {

    telefono = telefono.replace(PLUS_SIGN, VACIO);
    telefono = telefono.replace(ESPACIO, VACIO);

  }

  /**
   * Validacion para parentesis sin guion [(261) 1234567]
   */
  if (telefono.length === 13 && telefono.substring(2, 3) === ESPACIO) {

    telefono = telefono.substring(3);
    primera = telefono.substring(0, 3)
    segunda = telefono.substring(3, 6)
    tercera = telefono.substring(6, telefono.length)

    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;

  }
  /**
   *  Validacion para numero con pais sin el mas [54 2611234567]
   */
  else if (telefono.length === 13 && telefono.substring(0, 1) === PARENTESIS_IZQ) {

    // borrar primer parentesis 
    telefono = telefono.replace(PARENTESIS_IZQ, VACIO);
    // borrar segundo parentesis ()
    telefono = telefono.replace(PARENTESIS_DER, VACIO).trim();

    primera = telefono.substring(0, 3)
    segunda = telefono.substring(4, 7)
    tercera = telefono.substring(7, telefono.length)


    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;


  }

  /**
   * Validacion para el formato correcto [261 123-4567]
   */
  if (telefono.length === 12 && telefono.substring(7, 8) === GUION) {

    nuevoTelefono = telefono

  }
  /**
   *  Validacion para formato todo junto sin espacios [542611234567]
   */
  if (telefono.length === 12 && telefono.substring(0, 2) == STR_54) {

    telefono = telefono.replace(STR_54, VACIO)

    primera = telefono.substring(0, 3)
    segunda = telefono.substring(3, 6)
    tercera = telefono.substring(6, telefono.length)

    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;

  }
  /**
   *  Validacion para [2611 99-9999]
   */

  if (telefono.length === 12
    && telefono.substring(4, 5) === ESPACIO
    && telefono.replace(ESPACIO, VACIO).length == 11
    && telefono.replace(ESPACIO, VACIO).substring(6, 7) === GUION) {

    telefono = telefono.replace(ESPACIO, VACIO);
    primera = telefono.substring(0, 3)
    segunda = telefono.substring(3, 6)
    tercera = telefono.substring(6, telefono.length)

    nuevoTelefono = primera + ESPACIO + segunda + tercera;
  }

  /**                 
   * Validacion para [2611234567]
   */
  if (telefono.length === 10) {

    primera = telefono.substring(0, 3)
    segunda = telefono.substring(3, 6)
    tercera = telefono.substring(6, telefono.length)

    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;

  }

  /**
   *  Telefonos fijos con guion [412-3456]
   */
  if (telefono.length === 8) {



    nuevoTelefono = STR_261 + ESPACIO + telefono

  }

  /**
   *  Telefonos fijos sin guion [4123456]
   */
  if (telefono.length === 7) {

    segunda = telefono.substring(0, 3)
    tercera = telefono.substring(4, telefono.length)

    nuevoTelefono = STR_261 + ESPACIO + segunda + GUION + tercera

  }

  return nuevoTelefono;

}

/**
 * 
 */
function procesarHoja(telefonos, personas) {

  for (var i = 0; i < telefonos.length; i++) {

    telefono = telefonos[i];
    persona = personas[i];
    nuevoTelefono = VACIO;


    /**
     * Convertir numbers a string en caso de que el valor de la celda se recupere como number
     */
    if (telefono.length === undefined || typeof (telefono) === TYPE_NUMBER) {
      telefono = telefono.toString()
      console.log("Telefono convertido a string: " + telefono + " Longitud: " + telefono.length)
    } else {
      console.log(telefono.length)
    }


    telefono = normalizarTelefonos(telefono)
    // Modificar hoja de la que se tomaron los datos con los valores normalizados. 
    escribirValoresExcel(CELDAS_NAME, persona);
    escribirValoresExcel(CELDAS_GIVEN_NAME, persona);
    escribirValoresExcel(CELDAS_MEMBERS, MEMBERS_VALUE);
    escribirValoresExcel(CELDAS_TELEFONOS,telefono)
    //escribirNumerosTelefonoExcel(nuevoTelefono, telefono);

  }

}
