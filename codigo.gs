// Función para unir varios sheets iguales en un mismo sheet. 
// Copiar y pegar este código en Extentiones/ Apps Scripts en Google Sheets, cambiando las variables donde se indique

function unificar_planillas() {
  
// CAMBIAR
// Definir los nombres de las columnas de la hoja donde se van a unificar los sheets. 
// Cada nombre va entre comilla, separados por comas. Tienen que estar en el mismo orden que las hojas que se van a copiar.

  var header = [["Columna_1", "Columna_2", "Columna_3", "Columna_4"]]
  
  // CAMBIAR
  // Número de columna (posición) que contiene los ids de sheets que voy a utilizar.
  // Las columnas se cuentan comenzando desde 0, o sea que la columna "A" es 0, "B" es 1, etc.
  var id_col = 1; 
  
  // CAMBIAR
  // Nombre de la pestaña que tiene los datos que queremos copiar, DEBE SER EL MISMO NOMBRE EN TODOS LOS SHEETS
  var data_tab = 'Hoja 1';
  
  
// CAMBIAR
// ID de la hoja de origen, es un sheet que contiene una columna con un identificador (nombre del cliente por ejemplo) y el id del sheet que queremos copiar.
  var ss = SpreadsheetApp.openById ("XXXXXXXXXXXXXXXXX"); 
  
// CAMBIAR
// Nombre de la pestaña (tab) que tiene los datos que precisamos, por defecto es 'Hoja 1'
  var hoja_ids = ss.getSheetByName("Hoja 1");
  
  // CAMBIAR
  // De qué columna a qué columna voy a tomar los datos. Si tengo solo dos columnas sin encabezado, es 'A1:B'
  var rango_ids = hoja_ids.getRange('A1:B').getValues(); 
  
  // CAMBIAR
  // ID del sheet de destino, debería ser el sheet donde se está pegando este código
  // Si la pestaña tiene otro nombre, reemplazar también el "Hoja 1"
  var hoja_destino = SpreadsheetApp.openById ("XXXXXXXXXXXX").getSheetByName('Hoja 1');
  
 // PRECAUCIÓN: esta línea borra todo el contenido de la hoja de destino
  hoja_destino.clear(); 
  
  hoja_destino.getRange(1,1,1,header[0].length).setValues(header); // Escribir header
  
  var array_ids = rango_ids.filter(function(x) { // para filtrar los datos y que no haya nada en blanco al final
  return (x.join('').length !== 0);
});

for(i=0;i<array_ids.length; i++){ // loop por los datos
try{ // para no detener la ejecucion si hay un error

  if (array_ids[i][id_col].lenght !=0){ //chequea que haya ID

  var hoja_origen = SpreadsheetApp.openById (array_ids[i][id_col]); // abre el sheet por ID
  
  // CAMBIAR
  // El rango de datos que va a tomar desde la hoja de origen, si tengo 10 columnas con encabezados el rango es 'A2:J'
  // (Empieza en la posición A2 y toma datos hasta la letra J)
  var rango_datos = "A2:J"
  var datos_origen = hoja_origen.getSheetByName(data_tab).getRange(rango_datos).getValues();
  
  datos_origen = datos_origen.filter(function(x) {
    return (x.join('').length !== 0);}); // para limpiar las celdas vacias
  // Agrego nombre del cliente o id de identificación en una nueva columna
  for(k=0;k<datos_origen.length; k++){
    datos_origen[k].push(array_ids[i][0])
  
  }
  
// escribo a la hoja unificada: numero de row, numero de columna, cantidad de rows, cantidad de columnas
hoja_destino.getRange(hoja_destino.getLastRow()+1,1, datos_origen.length, datos_origen[0].length).setValues(datos_origen); 

  }
}
catch(err){Logger.log('Error en '+array_ids[i][0]+ ' ' + err)} // log del error
}


}
