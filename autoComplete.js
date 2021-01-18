function autoComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const cfTable = ss.getSheetByName('cfTable')
  const advancedPaymentTable = ss.getSheetByName('advancedPaymentTable')

  // 列番号・行番号を定義
  const firstRow = 2
  const firstCol = 1
  // 最終行・列番号を取得
  const lastRow = cfTable.getDataRange().getLastRow()
  const lastCol = cfTable.getDataRange().getLastColumn()
  // "立替"の列番号
  const advancedPaymentCol = 1
  // "家計"の列番号
  const householdAccountCol = 2
  // "内容"の列番号(csvの場合は-1)
  const titleCol = 6
  // "計算対象"の列番号(csvの場合は-1)
  const isValidCol = 11
  // advancedPaymentTableの最終行・列番号を取得
  const aptLastRow = advancedPaymentTable.getLastRow()
  const aptLastCol = advancedPaymentTable.getLastColumn()
  // advancedPaymentTable."立替"の列番号
  const aptAdvancedPaymentCol = 1
  // advancedPaymentTable."家計"の列番号
  const aptHouseholdAccountCol = 2
  // advancedPaymentTable."内容"の列番号
  const aptTitleCol = 3

  const valuesRange = cfTable.getRange(firstRow, firstCol, lastRow - 1, lastCol)
  const values = valuesRange.getValues()

  const patterns = advancedPaymentTable.getRange(2, 1, aptLastRow -1, aptLastCol).getValues()

  const completedValues = values.map((value) => {
    const isValidValue = value[isValidCol - 1] === 1
    // 計算対象の場合はそのまま返す
    if (!isValidValue) return value

    const valueTitle = String(value[titleCol - 1])
    patterns.forEach((pattern) => {
      // 内容が立替項目一覧に該当した場合は立替・家計を変更する
      if (valueTitle.includes(String(pattern[aptTitleCol - 1]))) {
        value[advancedPaymentCol - 1]   = pattern[aptAdvancedPaymentCol - 1]
        value[householdAccountCol - 1]  = pattern[aptHouseholdAccountCol - 1]
      }
    })
    return value
  })

  valuesRange.setValues(completedValues);
};
