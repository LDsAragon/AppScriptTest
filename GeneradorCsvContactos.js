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

const SELECCION_TELEFONOS = "T11:T50";
const SELECCION_PERSONAS = "C11:C50";
const GUION = "-";
const ESPACIO = " ";

let primera;
let segunda;
let tercera;
let telefono;
let persona;
let nuevoTelefono;


/**
 * Documento en el que se va a escribir el csv para importar a contactos de google.
 */
const documentoContactos = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ieVLgszeTa9UZ1g9QsYg3DSHNSWC3e7aSDS5SuB1Eyw/edit#gid=1964355463");

function myFunction() {

  let doc1 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1Mqf3kyq1bPOqPk2CVSGxAlw9KZuH_dG4zFoe4eW1AUo/edit#gid=1174061784");


  const HOJAS = doc1.getSheets().length;

  for (var i = 1; i < HOJAS + 1; i++) {

    // Obtener valores de columna de telefonos en la hoja i
    let telefonos = doc1.getSheets()[i].getRange(SELECCION_TELEFONOS).getValues();
    // Obtener valores de columna personas para poder iterarlos en la hoja i
    let personas = doc1.getSheets()[i].getRange(SELECCION_PERSONAS).getValues();

    // Limpiar array de telefonos para poder iterarlos
    telefonos = limpiarArray(telefonos)
    // Limpiar array de personas para poder iterarlos 
    personas = limpiarArray(personas);

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
 * Returns a flattend and clean array
 * Validates if the entry value is a number or not
 */
function limpiarArray(arr) {

  arr = arr.flat();

  arr = arr.filter(entry => {

    if (typeof (entry) === "number") {
      entry = entry.toString()
    }
    return entry.trim() != VACIO;
  });

  return arr

}

/**
 * Obtiene una celda activa de la primera hoja del documentoContactos
 * @celdas toma el valor de una columna oara escribir 
 * Utiliza el parametro global fila para determinar la posicion en la columna a escribir. 
 * Selecciona esa celda con el cursos activo y la devuelve como objeto.
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
function escribirNumerosTelefonoExcel(nuevoTelefono, telefono) {


  let cellValue = getCurrentCellComplete(CELDAS_TELEFONOS).getValue();

  if (cellValue == VACIO) {

    getCurrentCellComplete(CELDAS_TELEFONOS).setValue(nuevoTelefono);

  } else if (cellValue !== VACIO) {

    fila = fila + 1

    if (nuevoTelefono === undefined) {
      nuevoTelefono = telefono;
    }

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

    fila = fila + 1
    getCurrentCellComplete(celda).setValue(valor);

  }
}

function normalizarTelefonos(telefono) {

  primera = "";
  segunda = "";
  tercera = "";


  /**
    * Validacion para [+54 2611234567]
    */
  if (telefono.length === 15 && telefono.substring(0, 1) === "+") {
    telefono = telefono.substring(4)

    primera = telefono.substring(-1, 4).trim();
    segunda = telefono.substring(4, 7).trim();
    tercera = telefono.substring(7, telefono.length).trim();

    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;
  }
  /**
   * Validacion para parentesis con guion [(261) 123-4567]
   */
  if (telefono.length === 14 && telefono.substring(0, 1) === "(") {


    // borrar primer parentesis 
    telefono = telefono.replace("(", "");
    // borrar segundo parentesis ()
    telefono = telefono.replace(")", "");

    nuevoTelefono = telefono;

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
  else if (telefono.length === 13 && telefono.substring(0, 1) === "(") {

    // borrar primer parentesis 
    telefono = telefono.replace("(", "");
    // borrar segundo parentesis ()
    telefono = telefono.replace(")", "").trim();

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
  if (telefono.length === 12 && telefono.substring(0, 2) == "54") {

    telefono = telefono.replace("54", VACIO)

    primera = telefono.substring(0, 3)
    segunda = telefono.substring(3, 6)
    tercera = telefono.substring(6, telefono.length)

    nuevoTelefono = primera + ESPACIO + segunda + GUION + tercera;

  }
  /**
   *  Validacion para [6211 99-9999]
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



    nuevoTelefono = "261" + ESPACIO + telefono

  }

  /**
   *  Telefonos fijos sin guion [4123456]
   */
  if (telefono.length === 7) {

    segunda = telefono.substring(0, 3)
    tercera = telefono.substring(4, telefono.length)

    nuevoTelefono = "261" + ESPACIO + segunda + GUION + tercera

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
    if (telefono.length === undefined || typeof (telefono) === "number") {
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
    escribirNumerosTelefonoExcel(nuevoTelefono, telefono);

  }

}
