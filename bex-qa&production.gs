var sheetNameInformation = 'Información';
var sheetNameSupplies = 'Insumos';

function importFilteredChecklist() {
  var filters = getComponentFilters();
//Armamos los filtros dinamicos para el qry
  var qryValues = '             Col2 contains \'DEFAULT\' \n';
  if (filters!=null && filters.length>0) {
    for (var i=0;i<filters.length;i++) {
      qryValues+= '          OR Col2 contains \'' + filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "") + '\' \n';
    }
  }
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSupplies);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.getRange('B2').activate();
//  spreadsheet.getCurrentCell().setFormula('QUERY(IMPORTRANGE(VLOOKUP("idChecklist",\'Configuración\'!C8:D,2,0),VLOOKUP("rangoCheklist",\'Configuración\'!C8:D,2,0)), "select * where Col1 contains \'"&\'Información\'!D27&"\' and (Col2 contains \'DEFAULT\' OR Col2 contains \'HOST\' OR Col2 contains \'APX\') order by Col3, Col2")');
  spreadsheet.getCurrentCell().setFormula(
    'QUERY(IMPORTRANGE(VLOOKUP("idChecklist",\'Configuración\'!C8:D,2,0),VLOOKUP("rangoCheklist",\'Configuración\'!C8:D,2,0)), \n'+
    '" select * \n' +
    '   where Col1 contains \'"&\'Información\'!D27&"\' \n' +
    '     and ( \n' +
               qryValues +
    '         ) \n'+
    'order by Col3, Col2",0)'
  );
}

function getComponentFilters() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameInformation);
  var filterComponents = sheet.getRange("D33:F35").getValue();
  if (filterComponents!=null && filterComponents.length>0) {
    return filterComponents.split(",")
  } else {
    return null;
  }
}

function closeChecklist() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSupplies);
  spreadsheet.setActiveSheet(sheet);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()-1).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}