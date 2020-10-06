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

function importFilteredChecklistV2() {
  var isBEx = isBuldingBlockBEx();
  Logger.log('isBEx: ' + isBEx);
  var exchangeRate = getExchangeRate();
  var components = getComponentsFilters();
  importFilteredChecklistPostV2(isBEx, exchangeRate, components);
  importFilteredChecklistPrevV2(isBEx, exchangeRate, components);
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
    (isBEx ? '     and Col13 contains true \n' : '') + 
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
    (isBEx ? '     and Col13 contains true \n' : '') + 
    '     and Col3 contains \'5. POST PRODUCCIÓN\' \n' +
    '     and ( \n' +
               qryValues +
    '         ) \n'+
    'order by Col3, Col2, Col4, Col1",0)'
  );
}

function importFilteredChecklistPrevV2(isBEx, exchangeRate, components) {
//Armamos los filtros dinamicos para el qry
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSuppliesPrev);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula(
    'QUERY(IMPORTRANGE(VLOOKUP("idChecklist",\'Configuración\'!C8:D,2,0),VLOOKUP("rangoCheklistV2",\'Configuración\'!C8:D,2,0)), \n'+
    '" select Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col14, Col15 \n' +
    '   where Col1 is not null \n' + 
    '     and not Col1 contains \'7. INSTALACIÓN\' \n' +
    '     and not Col1 contains \'8. POST INSTALACIÓN\' \n' +
    (!isBEx ? '     and Col12 != \'NO\' \n' : '     and Col13 != \'NO\' \n') + 
    (exchangeRate=='NORMAL'    ? '     and Col9 = TRUE \n'  :
    (exchangeRate=='EMERGENTE' ? '     and Col10 = TRUE \n' :
    (exchangeRate=='URGENTE'   ? '     and Col11 = TRUE \n' : ''))) +
    '     and ( \n' +
               components +
    '         ) \n'+
    'order by Col1, Col2",0)'
  );
}

function importFilteredChecklistPostV2(isBEx, exchangeRate, components) {
//Armamos los filtros dinamicos para el qry
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameSuppliesPost);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula(
    'QUERY(IMPORTRANGE(VLOOKUP("idChecklist",\'Configuración\'!C8:D,2,0),VLOOKUP("rangoCheklistV2",\'Configuración\'!C8:D,2,0)), \n'+
    '" select Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col14, Col15 \n' +
    '   where Col1 is not null \n' + 
    '     and (   Col1 contains \'7. INSTALACIÓN\' \n' +
    '          or Col1 contains \'8. POST INSTALACIÓN\') \n' +
    (!isBEx ? '     and Col12 != \'NO\' \n' : '     and Col13 != \'NO\' \n') + 
    (exchangeRate=='NORMAL'    ? '     and Col9 = TRUE \n'  :
    (exchangeRate=='EMERGENTE' ? '     and Col10 = TRUE \n' :
    (exchangeRate=='URGENTE'   ? '     and Col11 = TRUE \n' : ''))) +
    '     and ( \n' +
               components +
    '         ) \n'+
    'order by Col1, Col2",0)'
  );
}

function isBuldingBlockBEx() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameInformation);
  var filter = sheet.getRange("D23:F23").getValue();
  Logger.log('Valor de buldingBlock: ' + filter);
  if (filter!=null && filter.length>0) {
      if (filter.replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='BEx')
        return true;
  }
  return false;
}

function getExchangeRate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetNameInformation);
  var filter = sheet.getRange("D34:F34").getValue();
  Logger.log('Valor de tipo de cambio: ' + filter);
  if (filter!=null && filter.length>0) {
      return filter.replace(/^\s+|\s+$|\s+(?=\s)/g, "");
  }
  return null;
}

function getComponentsFilters() {
  var filters = getComponentFilters();
  var qryValues = '             Col16 = TRUE \n'; // DEFAULT
  if (filters!=null && filters.length>0) {
    for (var i=0;i<filters.length;i++) {
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='AMELIA (Flujos)') {
        qryValues+= '          OR Col17 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='AMELIA (Entrenamiento)') {
        qryValues+= '          OR Col18 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='APX (Batch)') {
        qryValues+= '          OR Col19 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='APX (Online)') {
        qryValues+= '          OR Col20 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='ASO (Consumo)') {
        qryValues+= '          OR Col21 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='ASO (Create)') {
        qryValues+= '          OR Col22 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='BACKEND') {
        qryValues+= '          OR Col23 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='CONTROL-M') {
        qryValues+= '          OR Col24 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='DB-QUERYS') {
        qryValues+= '          OR Col25 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='DB-MODEL') {
        qryValues+= '          OR Col26 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='FRONTEND') {
        qryValues+= '          OR Col27 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='HOST') {
        qryValues+= '          OR Col28 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='ROBOTICS') {
        qryValues+= '          OR Col29 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='WM-BPM (Modelado)') {
        qryValues+= '          OR Col30 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='WM-IS (Servicios)') {
        qryValues+= '          OR Col31 = TRUE \n';
      }
      if (filters[i].replace(/^\s+|\s+$|\s+(?=\s)/g, "")=='WM-MWS (Conf)') {
        qryValues+= '          OR Col32 = TRUE \n';
      }      
    }
  }
  return qryValues;
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