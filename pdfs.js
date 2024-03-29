/**
 * Esta funcion recibe como parametro una URL y le se le setea valores 
 * para luego devolver un archivo blob de tipo PDF
 **/
function crearPdf(url) {

  let exportarUrl = url.replace(/\/edit.*$/, '')
    + '/export?exportFormat=pdf'
    + '&format=pdf'
    + '&size=A4'
    + '&portrait=true'
    + '&fitw=true'
    //+ '&scale=1' + //Ajustes {1=100%,2=Ancho,3=Alto,4=Página}       
    + '&top_margin=0.5'
    + '&bottom_margin=0.5'
    + '&left_margin=0.2'
    + '&right_margin=0.2'
    + '&sheetnames=false'
    + '&printtitle=false'
    // + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
    + '&gridlines=true'
    + '&fzr=FALSE'

  var respuesta = UrlFetchApp.fetch(exportarUrl, {
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  })
  return respuesta
}
//---------------------------------------------------------------------------------
/**
 * crea reporte PDF de cada UF y links, luego extrae y pega el id del PDF y el link del mismo para ubicarlos 
 * en la columna y fila correspondiente en la HojaMails para su posterior uso.
 **/
function crearPdfsyLinksMasivos (libro, carpeta, hojaMails, cantUF, celdaNombre, celdaUF, rangoCol, mes) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = libro.getSheetByName("DEUDORES Y PRORRATEO")

  let rangoUF = hojaProrrateo.getRange(6, 1, cantUF).getValues()
  let mesActual = hojaProrrateo.getRange(mes).getValue()
  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()
  let carpetaMes = carpeta.createFolder(mesActual)

  let url = libro.getUrl()
  for (let i = 0; i < cantUF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(celdaUF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(celdaNombre).getValue()
    // --------------------------------------------------------------------------------------
    let blob = crearPdf(url)
    Utilities.sleep(3000)

    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    // -------------------------------------------------------------------------------------------
    // Logger.log(archivo.getName()+ rangoUF[i])
    // Logger.log("archivo")
    // Logger.log(archivo)
    // Logger.log(archivo.getDownloadUrl())
    // Logger.log(archivo.getId())

    hojaMails.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // hojaMails.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    hojaMails.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(libro, rangoCol)
}

//-------------------------------------------------------------------------------------------------------------------------------------------
/**
 * Reinicia los reportes PDF desde una UF especifica justo con la creacion de links, luego extrae y pega el id del PDF y el link 
 * del mismo para ubicarlos en la columna y fila correspondiente en la HojaMails para su posterior uso
 **/
function reiniciarPdfsyLinksMasivos(libro, carpeta, hojaMails, cantUF, celdaNombre, celdaUF, rangoCol, mes, uf) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = libro.getSheetByName("DEUDORES Y PRORRATEO")

  let rangoUF = hojaProrrateo.getRange(6, 1, cantUF).getValues()
  let mesActual = hojaProrrateo.getRange(mes).getValue()
  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()
  let carpetaMes = carpeta.createFolder(mesActual)

  Logger.log("Mostrando uf:")
  Logger.log(uf)

  for (let i = uf; i < cantUF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(celdaUF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(celdaNombre).getValue()
    let url = libro.getUrl()
    let blob = crearPdf(url)
    Utilities.sleep(3000)

    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    Logger.log(archivo.getName() + rangoUF[i])
    Logger.log("archivo")
    Logger.log(archivo)
    Logger.log(archivo.getDownloadUrl())
    Logger.log(archivo.getId())

    hojaMails.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // hojaMails.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    hojaMails.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(libro, rangoCol)
}

//---------------------------------------------------------------------------------
/**Crea un reporte PDF con el nombre del archivo y la unidad funcional concatenada 
 * eliminando antes de crearlo el reporte personalizado
 **/
function crearReportePdf2(libro, carpeta, rangoCol) {


  let nombrelibro = libro.getName()

  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()

  let url = libro.getUrl()
  let blob = crearPdf(url)
  carpeta.createFile(blob).setName(nombrelibro)

  mostrarHojasyColumnasAH(libro, rangoCol)
}

//---------------------------------------------------------------------------------
/** Crea un reporte PDF con el nombre del archivo y la unidad funcional concatenada **/
function crearReportePdf(libro, carpeta, celdaNombre, celdaUF, rangoCol) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let nombreUf = hojaDetalle.getRange(celdaNombre).getValue()
  let numeroUF = hojaDetalle.getRange(celdaUF).getValue()

  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()

  let url = libro.getUrl()
  let blob = crearPdf(url)
  carpeta.createFile(blob).setName("UF" + numeroUF + espacio + nombrelibro + espacio + nombreUf)

  mostrarHojasyColumnasAH(libro, rangoCol)
}
//----------------------------------------------------------------------------------------------------------------------------------------------
/**
 * Crea un reporte PDF del detalle personalizado solamente
 **/
function crearDetallePdfsyLinksMasivos(libro, carpeta, hojaMails, cantUF, celdaNombre, celdaUF, rangoCol, mes) {

  let nombrelibro = libro.getName()
  let espacio = " "
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  let hojaProrrateo = libro.getSheetByName("DEUDORES Y PRORRATEO")

  let rangoUF = hojaProrrateo.getRange(6, 1, cantUF).getValues()
  let mesActual = hojaProrrateo.getRange(mes).getValue()
  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()
  let carpetaMes = carpeta.createFolder(mesActual)

  let url = libro.getUrl()
  for (let i = 0; i < cantUF; i++) {

    Logger.log("ESTOY MOSTRANDO NUM UF: " + rangoUF[i])

    Utilities.sleep(3000)
    hojaDetalle.getRange(celdaUF).setValue(rangoUF[i])
    SpreadsheetApp.flush()
    Utilities.sleep(2000)
    let nombreUF = hojaDetalle.getRange(celdaNombre).getValue()
    // --------------------------------------------------------------------------------------
    let blob = crearPdf(url)
    Utilities.sleep(3000)

    let archivo = carpetaMes.createFile(blob).setName("UF" + rangoUF[i] + espacio + nombrelibro + espacio + nombreUF)

    // -------------------------------------------------------------------------------------------
    // Logger.log(archivo.getName()+ rangoUF[i])
    // Logger.log("archivo")
    // Logger.log(archivo)
    // Logger.log(archivo.getDownloadUrl())
    // Logger.log(archivo.getId())

    hojaMails.getRange(i + 2, 5).setValue(archivo.getDownloadUrl())
    // hojaMails.getRange(i+2,5).setValue(archivo.getUrl())  OTRA MANERA DE CREAR LINK
    hojaMails.getRange(i + 2, 6).setValue(archivo.getId())

  }
  mostrarHojasyColumnasAH(libro, rangoCol)
}

//----------------------------------------------------------------------------------------------------------------------------------------------
/**
 * Crea un reporte PDF generico sin el apartado personalizado expresado en el intervalo 
 * que comienza en la fila r1 hasta r2 filas hacia abajo.
 **/
function crearPdfSinPersonalizar(libro, carpeta, rangoCol, r1, r2) {

  let nombrelibro = libro.getName()
  let hojaDetalle = libro.getSheetByName("DETALLE DE GASTOS")
  hojaDetalle.hideRows(r1, r2)//---------------- Esta linea no me permite reutilizar la funcion para los edificios que no tienen personalizacion
  SpreadsheetApp.flush()

  ocultarHojasyColumnasAH(libro, rangoCol)
  SpreadsheetApp.flush()

  let url = libro.getUrl()
  let blob = crearPdf(url)
  carpeta.createFile(blob).setName(nombrelibro)

  mostrarHojasyColumnasAH(libro, rangoCol)
  rango = hojaDetalle.getRange(r1, 1, r2)
  hojaDetalle.unhideRow(rango)
  Browser.msgBox("Se ah creado un PDF SIN personalizar ", Browser.Buttons.OK)

}
//---------------------------------------------------------------------------------------------------------------------------------