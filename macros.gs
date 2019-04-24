function cambios() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getCurrentCell().offset(2, 8, 2, 2).activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=sum(R[4]C[6]:R[10]C[6];R[11]C[7];R[13]C[6]:R[17]C[6])/13');
  spreadsheet.getCurrentCell().offset(11, -1, 2, 1).activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=if(AND(R[0]C[-1]>=7;R[1]C[-1]>=7);MAX(R[0]C[-1]:R[1]C[-1]);if(and(R[0]C[-1]<7;R[1]C[-1]<7);MAX(R[0]C[-1]:R[1]C[-1]);min(R[0]C[-1]:R[1]C[-1])))');
  spreadsheet.getCurrentCell().offset(4, -1, 1, 2).activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=\'PRIMERO A\'!R19C[75]');
  spreadsheet.getCurrentCell().offset(0, 2, 1, 2).activate();
  spreadsheet.getCurrentCell().setFormulaR1C1('=\'PRIMERO A\'!R19C[74]');
  spreadsheet.getCurrentCell().offset(1, 0, 1, 2).activate();
};