/** @OnlyCurrentDoc */
const dataSheetName = 'cashFlow'
const dirPath = 'Money/CashFlow/'
const filePath = '収入・支出詳細_2020-04-01_2020-04-30.csv'
const mailAddress = 'yoshikouki@gmail.com'

function importCashFlowFromCsv() {
  // スプレッドシートとデータシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(dataSheetName);
  
  try {
    // csv を取得・パースして配列化
    Logger.log('filePath: ', filePath)
    const file = DriveApp.getFilesByName(filePath).next()
    Logger.log('file: ', file)
    const data = file.getBlob().getDataAsString("Shift_JIS")
    Logger.log('data: ', data)
    const csv = Utilities.parseCsv(data)
    Logger.log('csv: ', csv)

    const newRow = dataSheet.getLastRow() + 1
    dataSheet.getRange(newRow, 1).setValue(csv[1][0])
  } catch(err) {
    MailApp.sendEmail(
      mailAddress, 
      'CSV のインポートに失敗しました', 
      err.message
    )
  }
};
