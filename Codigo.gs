function main() {
  var plantilla = "Plantilla";
  var data = 'Respuestas de formulario 1'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetData = ss.getSheetByName(data);
  var lastRow = sheetData.getLastRow();
  date = sheetData.getRange(lastRow,1).getValue();
  name = Utilities.formatDate(date, "GMT-5", "yyyy-MM-dd HH:mm:ss");
  console.log(name)
  cloneSheet(ss,plantilla,name);
  copyData(ss,name, lastRow, sheetData, date)
  copyImages(ss,name,lastRow,sheetData)
}

function cloneSheet(ss,plantilla,name){
  var sheet = ss.getSheetByName(plantilla).copyTo(ss);
  SpreadsheetApp.flush(); // Utilities.sleep(2000);
  sheet.setName(name);
  sheet.showSheet();
}

function copyData(ss, name, lastRow, sheetData, date){
  var sheet = ss.getSheetByName(name)
  coordenadas = sheetData.getRange(lastRow,8).getValue()
  indxLatitud = coordenadas.indexOf("4.")
  indxLongitud = coordenadas.indexOf("74.")
  indxComma = coordenadas.indexOf(',')
  latitud = coordenadas.substring(indxLatitud,indxComma)+"N"
  longitud = coordenadas.substring(indxLongitud,50)+"W"

  sheet.getRange(5,6).setValue("DELTEC") // Empresa Colaboradora:
  sheet.getRange(5,19).setValue(sheetData.getRange(lastRow,4).getValue()) // Responsable Ejecución:				
  sheet.getRange(6,4).setValue(8400135558) // N° de Contrato:
  sheet.getRange(6,12).setValue(Utilities.formatDate(date, "GMT-5", "dd")) // Fecha:	DIA
  sheet.getRange(6,13).setValue(Utilities.formatDate(date, "GMT-5", "MM")) // Fecha:	MES
  sheet.getRange(6,14).setValue(Utilities.formatDate(date, "GMT-5", "yy")) // Fecha:	AÑO	
  sheet.getRange(7,22).setValue(sheetData.getRange(lastRow,3).getValue()) // CD / DME:	ok
  sheet.getRange(8,5).setValue("ZONA CENTRO") // Zona y Subzona:		
  sheet.getRange(8,18).setValue("BOGOTÁ") // Municipio:		OK
  sheet.getRange(9,4).setValue(sheetData.getRange(lastRow,5).getValue()) // Dirección:	OK
  sheet.getRange(9,18).setValue(sheetData.getRange(lastRow,6).getValue()) // Barrio ok
  sheet.getRange(13,2).setValue(sheetData.getRange(lastRow,9).getValue()) // P. Físico	ok

  sheet.getRange(13,5).setValue(latitud) // Latitud				
  sheet.getRange(13,10).setValue(longitud) // Longitud			
  sheet.getRange(14,14).setValue(sheetData.getRange(lastRow,26).getValue()) // Medidas1								
  sheet.getRange(14,20).setValue("NA") // Medidas2
  sheet.getRange(17,18).setValue(sheetData.getRange(lastRow,27).getValue()) // Medidas3
  sheet.getRange(21,18).setValue(sheetData.getRange(lastRow,28).getValue()) // Estado

}

function copyImages(ss,name,lastRow,sheetData){
  console.log('1')
  var sheet = ss.getSheetByName(name)
  urls = sheetData.getRange(lastRow, 29,1,4).getValues().flat()
  cnt = 1
  console.log(urls)
    for (var url in urls){
      console.log(urls[url])
      indx = urls[url].indexOf('=')
      id = urls[url].substring(indx+1,80)
      console.log(id)
      if (cnt == 1){
        var res = ImgApp.doResize(id, 270);
        sheet.insertImage(res.blob, 2, 26);
      }else if (cnt == 2){
        var res = ImgApp.doResize(id, 270);
        sheet.insertImage(res.blob, 8, 26);
      }else if (cnt == 3){
        var res = ImgApp.doResize(id, 265);
        sheet.insertImage(res.blob, 14, 26);
      }else if (cnt == 4){
        var res = ImgApp.doResize(id, 265);
        sheet.insertImage(res.blob, 20, 26);
      }
      cnt+=1
    }
}

function descargarHoja(){
  // Get current spreadsheet's ID, place in download URL
  var ssId = SpreadsheetApp.getActive().getId();
  var spId = SpreadsheetApp.getActive().getSheetId();
  var URL = 'https://docs.google.com/spreadsheets/d/'+ssId+'/export?format=xlsx&gid='+spId;

  // Display a modal dialog box with download link.
  var htmlOutput = HtmlService
                  .createHtmlOutput('<a href="'+URL+'">Click to download</a>')
                  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
                  .setWidth(80)
                  .setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Descargar XLSX');
}

