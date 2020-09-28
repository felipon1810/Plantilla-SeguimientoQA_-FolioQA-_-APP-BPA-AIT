var sheetNameInformation = 'Información';
var sheetNameSuppliesPrev = 'Insumos (Previo Instalación)';
var sheetNameSuppliesPost = 'Insumos (Post Instalación)';

function importFilteredChecklist() {
  var isBEx = isBuldingBlockBEx();
  Logger.log('isBEx: ' + isBEx);
  
  var filters = getComponentFilters();
  
  importFilteredChecklistPost(isBEx, filters);
  importFilteredChecklistPrev(isBEx, filters);
}

function importFilteredChecklistPrev(isBEx, filters) {
//Armamos los filtros dinamicos para el qry
  var qryValues = '             Col2 contains \'DEFAULT\' \n';
  if (filters!=null && filters.length>0) {
    for (var i=0;i<filters.length;i++) {
      qryValues+= '          OR Col2 contains \'' + filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "") + '\' \n';
    }
  }  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSuppliesPrev);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula(
    'QUERY(IMPORTRANGE(VLOOKUP("idChecklist",\'Configuración\'!C8:D,2,0),VLOOKUP("rangoCheklist",\'Configuración\'!C8:D,2,0)), \n'+
    '" select Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col10, Col11 \n' +
    '   where Col1 contains \'"&\'Información\'!D34&"\' \n' +
    (isBEx ? '     and not Col13 contains true \n' : '') + 
    '     and not Col3 contains \'5. POST PRODUCCIÓN\' \n' +
    '     and ( \n' +
               qryValues +
    '         ) \n'+
    'order by Col3, Col2, Col4, Col1",0)'
  );
}

function importFilteredChecklistPost(isBEx, filters) {
//Armamos los filtros dinamicos para el qry
  var qryValues = '             Col2 contains \'DEFAULT\' \n';
  if (filters!=null && filters.length>0) {
    for (var i=0;i<filters.length;i++) {
      qryValues+= '          OR Col2 contains \'' + filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "") + '\' \n';
    }
  }
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSuppliesPost);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula(
    'QUERY(IMPORTRANGE(VLOOKUP("idChecklist",\'Configuración\'!C8:D,2,0),VLOOKUP("rangoCheklist",\'Configuración\'!C8:D,2,0)), \n'+
    '" select Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col10, Col11 \n' +
    '   where Col1 contains \'"&\'Información\'!D34&"\' \n' +
    (isBEx ? '     and not Col13 contains true \n' : '') + 
    '     and Col3 contains \'5. POST PRODUCCIÓN\' \n' +
    '     and ( \n' +
               qryValues +
    '         ) \n'+
    'order by Col3, Col2, Col4, Col1",0)'
  );
}

function isBuldingBlockBEx() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameInformation);
  var filterBuldingBlock = sheet.getRange("D23:F23").getValue();
  Logger.log('Valor de buldingBlock: ' + filterBuldingBlock);
  if (filterBuldingBlock!=null && filterBuldingBlock.length>0) {
      if (filterBuldingBlock.replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='BEx')
        return true;
  }
  return false;
}

function getComponentFilters() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameInformation);
  var filterComponents = sheet.getRange("D40:F42").getValue();
  Logger.log('Valor de Componentes: ' + filterComponents);
  if (filterComponents!=null && filterComponents.length>0) {
    return filterComponents.split(",")
  } else {
    return null;
  }
}

function closeChecklist() {
  closeChecklistPost();
  closeChecklistPrev();
}

function closeChecklistPrev() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSuppliesPrev);
  spreadsheet.setActiveSheet(sheet);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()-1).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}

function closeChecklistPost() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSuppliesPost);
  spreadsheet.setActiveSheet(sheet);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()-1).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}