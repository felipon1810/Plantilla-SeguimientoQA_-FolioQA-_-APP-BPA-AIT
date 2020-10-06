function onOpen(e) {
  SpreadsheetApp.getUi()
     .createMenu('BEx QA & Production')
     .addItem('Import Filtered Cheklist V1', 'importFilteredChecklist')
     .addItem('Import Filtered Cheklist V2', 'importFilteredChecklistV2')
     .addSeparator()
     .addItem('Close Cheklist', 'closeChecklist')
     .addToUi();
  SpreadsheetApp.getActiveSpreadsheet().toast('Menú de BEx QA & Production en ejecución', 'AVISO', 5);
}