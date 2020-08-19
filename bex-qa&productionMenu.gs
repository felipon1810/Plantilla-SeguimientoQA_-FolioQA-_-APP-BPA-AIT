function onOpen(e) {
  SpreadsheetApp.getUi()
     .createMenu('BEx QA & Production')
     .addItem('1. Import Filtered Cheklist', 'importFilteredChecklist')
     .addItem('2. Close Cheklist', 'closeChecklist')
     .addToUi();
  SpreadsheetApp.getActiveSpreadsheet().toast('Menú de BEx QA & Production en ejecución', 'AVISO', 5);
}