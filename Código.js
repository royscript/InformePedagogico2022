function doGet(request) {//Esta función permite que se ejecute primero la páfina index.html
  return HtmlService.createTemplateFromFile('index')
      .evaluate();
}

function include(filename) {//Esta función activa la importacion de elementos al html
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
const idExcel = "1zPgEhN6BOzLwjlBPExyMRxTfOdQOCN1n3pvsltoT7EY";
function obtenerDatosDeUnExcel(nombreHoja,cantidadColumnas, idPlanilla = null){
  var ss = "";
  var ws = "";
  var data = [];
  try {
      if(idPlanilla != null) {
        ss = SpreadsheetApp.openById(idPlanilla)
      }else{
        ss = SpreadsheetApp.openById(idExcel)
      }
      ws = ss.getSheetByName(nombreHoja);
      data = ws.getRange(1, 1, ws.getLastRow(),cantidadColumnas).getValues();
      return data;
  } catch (error) {
      Logger.log(error);
      return [];
  }
}
function transformarNota (nota) {
   var notaString = String(nota);
   if(notaString.length==1){
     notaString = notaString+'.0';
   }
   return notaString;
}
function eMedia (curso){
	//var curso = '1NMA';
	let contador = 0;
	const data = obtenerDatosDeUnExcel(curso,87);
    var datos = [];
    for(var x=2;x<data.length;x++){
        var dato = data[x];
        contador++;
        if(dato[1]!='' && dato[7].length>3 && dato[7] != "NOMBRES"){
          datos.push({
           numero : contador,
           curso : curso,
           apellidoPaterno : dato[5],
           apellidoMaterno : dato[6],
           nombres : dato[7],
           rut : dato[4],

           promedioLenguaje : transformarNota(dato[17]),
           conceptoLenguaje : dato[18],
                         
           promedioESociales : transformarNota(dato[28]),
           conceptoESociales : dato[29],
                         
           promedioMatematica : transformarNota(dato[39]),
           conceptoMatematica : dato[40],
                         
           promedioIngles : transformarNota(dato[50]),
           conceptoIngles : dato[51],
                         
           promedioCNaturales : transformarNota(dato[61]),
           conceptoCNaturales : dato[62],
                         
           promedioTallerUno : transformarNota(dato[72]),
           conceptoTallerUno : dato[73],
           
           promedioTallerDos : transformarNota(dato[83]),
           conceptoTallerDos : dato[84],
                         
           promedioFinal : transformarNota(dato[85]),
           conceptoFinal : dato[86]
          });
        }
        
    }
    Logger.log(datos[1]);
	/*var datos = data.map(dato=>{
      contador++;
      return {
           numero : contador,
           curso : '',
           apellidoPaterno : dato[2],
           apellidoMaterno : dato[3],
           nombres : dato[4],
           rut : dato[1],

           promedioLenguaje : dato[24],
           conceptoLenguaje : dato[25],
                         
           promedioESociales : dato[33],
           conceptoESociales : dato[34],
                         
           promedioMatematica : dato[42],
           conceptoMatematica : dato[43],
                         
           promedioIngles : dato[51],
           conceptoIngles : dato[52],
                         
           promedioCNaturales : dato[60],
           conceptoCNaturales : dato[61],
                         
           promedioTallerUno : dato[65],
           conceptoTallerUno : dato[66],
           
           promedioTallerDos : dato[70],
           conceptoTallerDos : dato[71],
                         
           promedioFinal : dato[72],
           conceptoFinal : dato[73]
      };
    });*/
  
  return datos;
}

function eBasica (){
	//var curso = '1NMA';
	let contador = 0;
	const data = obtenerDatosDeUnExcel('3NBA',80);
    var datos = [];
    for(var x=2;x<data.length;x++){
        var dato = data[x];
        contador++;
        if(dato[1]!='' && dato[7].length>3 && dato[7] != "NOMBRES"){
          datos.push({
           numero : contador,
           curso : '3NBA',
           apellidoPaterno : dato[5],
           apellidoMaterno : dato[6],
           nombres : dato[7],
           rut : dato[4],

           promedioLenguaje : transformarNota(dato[17]),
           conceptoLenguaje : dato[18],
                         
           promedioESociales : transformarNota(dato[28]),
           conceptoESociales : dato[29],
                         
           promedioMatematica : transformarNota(dato[39]),
           conceptoMatematica : dato[40],
                         
           promedioTallerUno : transformarNota(dato[72]),
           conceptoTallerUno : dato[73],
                         
           promedioCNaturales : transformarNota(dato[61]),
           conceptoCNaturales : dato[62],
                         
           promedioFinal : transformarNota(dato[74]),
           conceptoFinal : dato[75]
          });
        }
        
    }
    Logger.log(datos[1]);
  return datos;
}

function listarUsuariosDeUnDirectorio(nombreDirectorio) {
  var pageToken;
  var page;
  do {
    page = AdminDirectory.Users.list({
      domain: 'institutodetierrasblancas.cl',
      query: "orgUnitPath:'"+nombreDirectorio+"'",
      orderBy : 'familyName',
      pageToken: pageToken
    });
    var users = page.users;
    if (users) {
      return users;
    } else {
      return null;
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
}
/**
 * Sends emails with data from the current spreadsheet.
 */
var alumnoActual = [];
function getAlumnoActual(){
  return alumnoActual;
}

const idPlanillaComentarios = '1gzbrjJMM4wpYJ5qyoq3jMj4rx2DAjmxOzr-sM05a1Wk';
function obtenerDatos(nombreHoja,cantidadColumnas,cabeceras,camposBuscar,consultaBase = null){
  let registros = obtenerDatosDeUnExcel(nombreHoja,cantidadColumnas,idPlanillaComentarios);
  if(consultaBase!=null){
      //Ejemplo
      //Select * FROM ? WHERE [1] LIKE '%admin%' AND [0] = 1
      //[1] hace referencia que es la segunda casilla de la planilla excel
      
      registros = alasql(consultaBase,[registros]);
  }
  if(camposBuscar.length>0){
      let sql = 'SELECT * FROM ? WHERE ';
      camposBuscar.forEach((campo, index)=>{
          index == 0 ? sql += `[${campo.posicion}] LIKE '%${campo.valorCampo}%' ` : sql += `AND [${campo.posicion}] LIKE '%${campo.valorCampo}%' `
          
      });
      registros = alasql(sql,[registros]);
  }
  registros = registros.reverse();//Ponemos los datos desde el ultimo hasta el primero ingresado

  if(registros.length>0){
      let arrayDatos = [];
      registros.forEach((rangoDatos,index)=>{
          //let datos = registros[rangoDatos.numero];
          let datos = rangoDatos;
          //------Cada vez que se ejecute un ciclo se creará esta variable y al final se almacenará en el vector
          let union = [];
          let bandera = true;
          cabeceras.forEach(regCabecera=>{
              if(bandera==true){
                  bandera = false;
                  //registros[regCabecera.posicion]
                  union = {[regCabecera.nombreDato] : datos[regCabecera.posicion]};
              }else{
                  union = Object.assign(union,{[regCabecera.nombreDato] :datos[regCabecera.posicion]});
              }
          });
          arrayDatos.push(union);
          //---------------Fin almacenaje------------------------------------
      });
      return {
              registros : arrayDatos,
              totalRegistros : registros.length
      };
  }
  
  return [];
}