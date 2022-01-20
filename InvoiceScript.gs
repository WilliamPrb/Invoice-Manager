 /*

  * @license MIT

  *A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70 by reverse engineering the export feature
  *Credit to xfanatical.com, Jason Huang by him export / conversion to PDF code. 


  *Full complete spreadsheet system for issue, manage and store invoices on Google drive.
  *This tool is developed for this spreadsheet base model https://docs.google.com/spreadsheets/d/1I5Am92HMgCaUU15wOqFfIrTxJEW693Dc9MXVzUpVSCs/edit?usp=sharing

  *You should change the folder ID to store your invoices before start issue invoices. 
  
 */
 
 
 
 function newCostumer() {

  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const generate = ss.getSheetByName("Generate");
  const clients = ss.getSheetByName("Costumers")

  const taxId = generate.getRange("B4").getValue();
  const company = generate.getRange("C4").getValue();
  const country = generate.getRange("D4").getValue();
  const street = generate.getRange("E4").getValue();
  const city = generate.getRange("F4").getValue();
  const state = generate.getRange("G4").getValue();
  const zip = generate.getRange("H4").getValue();

  const lastRow = clients.getLastRow() + 1;

  clients.getRange(lastRow,1).setValue(company);
  clients.getRange(lastRow,2).setValue(taxId);
  clients.getRange(lastRow,3).setValue(country);
  clients.getRange(lastRow,4).setValue(street);
  clients.getRange(lastRow,5).setValue(city);
  clients.getRange(lastRow,6).setValue(state);
  clients.getRange(lastRow,7).setValue(zip);


  generate.getRange("B4:H4").clearContent();


}

function newInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoice = ss.getSheetByName("Generate"); 
  const print = ss.getSheetByName("Print");
  const log = ss.getSheetByName("Invoice Log");

  /* Getting Data */

  const costumerName = invoice.getRange("B11").getValue();
  const dueDate = invoice.getRange("D11").getValue();
  const productType = invoice.getRange("F11").getValue();
  const productName = invoice.getRange("C15").getValue();
  const productUnity = invoice.getRange("D15").getValue();
  const productValue = invoice.getRange("E15").getValue();
  const productCurrency = invoice.getRange("F15").getValue();
  const notes = invoice.getRange("E11").getValue();

 



  /* Generate Invoice Number */
  const random = Math.floor(Math.random() * (10 - 1 + 1)) + 1; 
  const invoiceNumber = log.getRange(log.getLastRow(), 1).getValue() + random;


  /* Setting Data */

  print.getRange("B12:C12").setValue(costumerName);
  print.getRange("F15:G15").setValue(dueDate);
  print.getRange("D12:E12").setValue(productType);
  print.getRange("B19:D19").setValue(productName);
  print.getRange("E19").setValue(productUnity);
  print.getRange("F19").setValue(productValue + "€");
  print.getRange("B23:E23").setValue(notes);
  print.getRange("D15").setValue(productCurrency);
  print.getRange("F12:G12").setValue(invoiceNumber);


  clearGenerate ();
}

function clearGenerate () {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoice = ss.getSheetByName("Generate"); 
  
  invoice.getRange("B11:F11").clearContent();
  invoice.getRange("B15:F15").clearContent();
  invoice.getRange("B4:H4").clearContent();
}

function printInvoice() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invoice = ss.getSheetByName("Generate"); 
  const print = ss.getSheetByName("Print");
  const log = ss.getSheetByName("Invoice Log");



    /* Getting Data */

  const invoiceNumber = print.getRange("F12:G12").getValue();
  const costumerName = print.getRange("B12:C12").getValue();
  const dueDate = print.getRange("F15:G15").getValue();
  const issuedDate = print.getRange("B9:C9").getValue();
  const productType = print.getRange("D12:E12").getValue();
  const productUnity = print.getRange("E19").getValue();
  const productValue = print.getRange("F19").getValue();
  const notes = print.getRange("B23:E23").getValue();


  /* Setting Data */

  const lastRow = log.getLastRow() + 1;

  log.getRange(lastRow,1).setValue(invoiceNumber);
  log.getRange(lastRow,2).setValue(costumerName);
  log.getRange(lastRow,3).setValue(issuedDate);
  log.getRange(lastRow,4).setValue(dueDate);
  log.getRange(lastRow,5).setValue(productType); 
  log.getRange(lastRow,6).setValue(productUnity);
  log.getRange(lastRow,7).setValue(productValue);
  log.getRange(lastRow,8).setValue(notes);

  exportNamedRangesAsPDF()

  print.getRange("F12:G12").clearContent();
  print.getRange("B12:C12").clearContent();
  print.getRange("F15:G15").clearContent();
  print.getRange("B9:C9").clearContent();
  print.getRange("D12:E12").clearContent();
  print.getRange("D15").clearContent();
  print.getRange("E19").clearContent();
  print.getRange("F19").clearContent();
  print.getRange("B23:E23").clearContent();

}

//Export pdf by Jason Huang
var ss = SpreadsheetApp.getActiveSpreadsheet();
var settings = ss.getSheetByName("Settings");

var saveToRootFolder = false

function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Export all sheets', 'exportAsPDF')
    .addItem('Export all sheets as separate files', 'exportAllSheetsAsSeparatePDFs')
    .addItem('Export current sheet', 'exportCurrentSheetAsPDF')
    .addItem('Export selected area', 'exportPartAsPDF')
    .addItem('Export predefined area', 'exportNamedRangesAsPDF')
}
function _exportBlob(blob, fileName, spreadsheet) {
  blob = blob.setName(fileName)
  var folder = saveToRootFolder ? DriveApp : DriveApp.getFolderById('')
  var pdfFile = folder.createFile(blob)
  
  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(80)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Invoice Issued')
}
function exportAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var blob = _getAsBlob(spreadsheet.getUrl())
  _exportBlob(blob, spreadsheet.getName(), spreadsheet)
}
function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true'       
      + '&top_margin=0.75'              
      + '&bottom_margin=0.75'          
      + '&left_margin=0.7'             
      + '&right_margin=0.7'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam
      
  Logger.log('exportUrl=' + exportUrl)
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}
function exportAllSheetsAsSeparatePDFs() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var settings = spreadsheet.getRange(2,2).getValue();
  var files = []
  var folder = saveToRootFolder ? DriveApp : DriveApp.getFolderById('')
  spreadsheet.getSheets().forEach(function (sheet) {
    spreadsheet.setActiveSheet(sheet)
    
    var blob = _getAsBlob(spreadsheet.getUrl(), sheet)
    var fileName = sheet.getName()
    blob = blob.setName(fileName)
    var pdfFile = folder.createFile(blob)
    
    files.push({
      url: pdfFile.getUrl(),
      name: fileName,
    })
  })
  
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open PDF files</p>'
      + '<ul>'
      + files.reduce(function (prev, file) {
        prev += '<li><a href="' + file.url + '" target="_blank">' + file.name + '</a></li>'
        return prev
      }, '')
      + '</ul>')
    .setWidth(300)
    .setHeight(150)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
}
function exportCurrentSheetAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = SpreadsheetApp.getActiveSheet()
  
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet)
  _exportBlob(blob, currentSheet.getName(), spreadsheet)
}
function exportPartAsPDF(predefinedRanges) {
  var ui = SpreadsheetApp.getUi()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var print = spreadsheet.getSheetByName("Print");


  var selectedRanges
  var fileSuffix
  if (predefinedRanges) {
    selectedRanges = predefinedRanges
    fileSuffix = print.getRange("F12:G12").getValue()
  } else {
    var activeRangeList = spreadsheet.getActiveRangeList()
    if (!activeRangeList) {
      ui.alert('Please select at least one range to export')
      return
    }
    selectedRanges = activeRangeList.getRanges()
    fileSuffix = '-selected'
  }
  
  if (selectedRanges.length === 1) {
    // special export with formatting
    var currentSheet = selectedRanges[0].getSheet()
    var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet, selectedRanges[0])
    
    var fileName = "Invoice n° " + fileSuffix
    _exportBlob(blob, fileName, spreadsheet)
    return
  }
  
  var tempSpreadsheet = SpreadsheetApp.create(spreadsheet.getName() + fileSuffix)
  if (!saveToRootFolder) {
    DriveApp.getFileById(tempSpreadsheet.getId()).moveTo(DriveApp.getFileById(spreadsheet.getId()).getParents().next())
  }
  var tempSheets = tempSpreadsheet.getSheets()
  var sheet1 = tempSheets.length > 0 ? tempSheets[0] : undefined
  SpreadsheetApp.setActiveSpreadsheet(tempSpreadsheet)
  tempSpreadsheet.setSpreadsheetTimeZone(spreadsheet.getSpreadsheetTimeZone())
  tempSpreadsheet.setSpreadsheetLocale(spreadsheet.getSpreadsheetLocale())
  
  for (var i = 0; i < selectedRanges.length; i++) {
    var selectedRange = selectedRanges[i]
    var originalSheet = selectedRange.getSheet()
    var originalSheetName = originalSheet.getName()
    
    var destSheet = tempSpreadsheet.getSheetByName(originalSheetName)
    if (!destSheet) {
      destSheet = tempSpreadsheet.insertSheet(originalSheetName)
    }
    
    Logger.log('a1notation=' + selectedRange.getA1Notation())
    var destRange = destSheet.getRange(selectedRange.getA1Notation())
    destRange.setValues(selectedRange.getValues())
    destRange.setTextStyles(selectedRange.getTextStyles())
    destRange.setBackgrounds(selectedRange.getBackgrounds())
    destRange.setFontColors(selectedRange.getFontColors())
    destRange.setFontFamilies(selectedRange.getFontFamilies())
    destRange.setFontLines(selectedRange.getFontLines())
    destRange.setFontStyles(selectedRange.getFontStyles())
    destRange.setFontWeights(selectedRange.getFontWeights())
    destRange.setHorizontalAlignments(selectedRange.getHorizontalAlignments())
    destRange.setNumberFormats(selectedRange.getNumberFormats())
    destRange.setTextDirections(selectedRange.getTextDirections())
    destRange.setTextRotations(selectedRange.getTextRotations())
    destRange.setVerticalAlignments(selectedRange.getVerticalAlignments())
    destRange.setWrapStrategies(selectedRange.getWrapStrategies())
  }
  
  // remove empty Sheet1
  if (sheet1) {
    Logger.log('lastcol = ' + sheet1.getLastColumn() + ',lastrow=' + sheet1.getLastRow())
    if (sheet1 && sheet1.getLastColumn() === 0 && sheet1.getLastRow() === 0) {
      tempSpreadsheet.deleteSheet(sheet1)
    }
  }
  
  exportAsPDF()
  SpreadsheetApp.setActiveSpreadsheet(spreadsheet)
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true)
}

function exportNamedRangesAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var allNamedRanges = spreadsheet.getNamedRanges()
  var toPrintNamedRanges = []
  for (var i = 0; i < allNamedRanges.length; i++) {
    var namedRange = allNamedRanges[i]
    if (/^print_area_.*$/.test(namedRange.getName())) {
      Logger.log('found named range ' + namedRange.getName())
      toPrintNamedRanges.push(namedRange.getRange())
    }
  }
  if (toPrintNamedRanges.length === 0) {
    SpreadsheetApp.getUi().alert('No print areas found. Please add at least one \'print_area_1\' named range in the menu Data > Named ranges.')
    return
  } else {
    toPrintNamedRanges.sort(function (a, b) {
      return a.getSheet().getIndex() - b.getSheet().getIndex()
    })
    exportPartAsPDF(toPrintNamedRanges)
    
  }
}




