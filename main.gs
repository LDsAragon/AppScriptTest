 
function obtenerValoresEnXML() {

  const url = "https://docs.google.com/spreadsheets/d/1LX2Dnz0ufE6ko57kcp8iavdh1sSSs_g7cscsotKCDsY/edit?usp=sharing";
  const url_con_espacios = "https://docs.google.com/spreadsheets/d/1HDuZjzPK7SyD7nfrs2nPpCsk_wrBJFtneH6H2LlRQMg/edit#gid=0";
  const sheetName = "Datos" ;


  var SPREADSHEET_URL = url_con_espacios;
  // Name of the specific sheet in the spreadsheet.
  var SHEET_NAME = sheetName;

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  // This represents ALL the data.
  var range = sheet.getDataRange();
  var values = range.getValues();

  const filteredData = sanearValores(values) // ayuda a evitar error por config del xml como columnas vacias entre los datos

  values = filteredData 

  const columnas = values[0] ; 
  values.splice(0, 1) // remover elemento columnas del array de valores
  const datos =  values

// Control 
   console.log(columnas)
   console.log(datos) 

  //let xml = createXML(columnas, datos) // feo 
  //let xml = createXML_2(columnas, datos) // bonito 
  let xml = createXML_3(columnas, datos) //bonito ++
  imprimirXmlBonito(xml);
  guardarXmlEnRaizDeDrive(xml)

}
/**
 * 
 * Limpiar array de espacios vacios
 */
function sanearValores(values){

  // console.log(values) // valores sin sanear

  let itemToReplace 

    values.forEach((item, index) => {   
        item.forEach(() => {
          itemToReplace = item.filter(element => {
            if(element !== null 
            && element !== undefined 
            && element !== "" 
            && element !== [])
            {   return true;  }
            })
        });

       values.splice(index, 1 , itemToReplace)
    });
    
    // console.log(values) // valores saneados
    return values ;
}

function createXML(columnas, datos){
    let data = XmlService.createElement("data")

    for (var i = 0; i < datos.length; i++) { 
      data.addContent(XmlService.createElement(columnas[0]).setText(datos[i][0]))
      data.addContent(XmlService.createElement(columnas[1]).setText(datos[i][1]))
      data.addContent(XmlService.createElement(columnas[2]).setText(datos[i][2]))
      data.addContent(XmlService.createElement(columnas[3]).setText(datos[i][3]))
    }
    
    return data ;
} 

function createXML_2(columnas, datos){
    let data = XmlService.createElement("data")

    for (var i = 0; i < datos.length; i++) {
      for (var j = 0; j < columnas.length; j++) {
        data.addContent(XmlService.createElement(columnas[j]).setText(datos[i][j]))
      }
    }
    
    return data ;
}

function createXML_3(columnas, datos){
     let data = XmlService.createElement("data")

    datos.forEach((itemDatos, indexDatos) => {    // por cada fila en la hoja de calculo
        columnas.forEach((itemCol, index) => {    // por cada columna en esa fila 
          data.addContent(XmlService.createElement(columnas[index]).setText(datos[indexDatos][index])) //crear elemento en el xml
        });
    });
    
    return data ;
}

function imprimirXmlBonito(xml) {
  var document = XmlService.createDocument(xml);
  var result = XmlService.getPrettyFormat().format(document);
  console.log(result)
  return result;
}

function guardarXmlEnRaizDeDrive(xml){

  const nombreArchivo = "Salida Xml" ; 
  const urlArchivo = "https://docs.google.com/document/d/1eiU5kgSNYVNDPL4TNrUaARt2ZnU0jKc2vrSlRK2izEw/edit?usp=sharing" ; 
  let doc ;

  try {
    doc = DocumentApp.openByUrl(urlArchivo);
    console.log("Recuperando el archivo de drive con nombre : " +nombreArchivo)
  } catch (err) {
    console.log("Creando el archivo: " +nombreArchivo)
    doc = DocumentApp.create(nombreArchivo);
  }
  
  var body = doc.getBody() 

  var result = XmlService.getPrettyFormat().format(xml);
  body.appendParagraph(result);
  
  doc.saveAndClose

  console.log("Guardado en drive en Archivo con nombre: " + nombreArchivo )
}



