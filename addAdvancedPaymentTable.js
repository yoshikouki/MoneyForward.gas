function addAdvancedPaymentTable() {
  const ss = SpreadsheetApp.getActive();
  const cfTable = ss.getSheetByName('cfTable')
  const advancedPaymentTable = ss.getSheetByName('advancedPaymentTable')

  // advancedPaymentTableの最終行・列番号を取得
  const aptLastRow = advancedPaymentTable.getLastRow()
  const aptLastCol = advancedPaymentTable.getLastColumn()
  // advancedPaymentTable."立替"の列番号
  const aptAdvancedPaymentCol = 1
  // advancedPaymentTable."家計"の列番号
  const aptHouseholdAccountCol = 2
  // advancedPaymentTable."内容"の列番号
  const aptTitleCol = 3

  const value = ss.getActiveCell().getValue()

  advancedPaymentTable.getRange(aptLastRow + 1, 1, 1, aptLastCol).setValues([
    [1, 1, value]
  ])
};