/** @OnlyCurrentDoc */
const dataSheetName = 'cashFlow'
const fileManagedSheetName = 'imortedFiles'
const dirPath = 'Money/CashFlow/'
const filePath = '収入・支出詳細_2020-04-01_2020-04-30.csv'
const mailAddress = 'yoshikouki@gmail.com'

function importCashFlowFromCsv() {
  // スプレッドシートとデータシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(dataSheetName);
  const fileManagementSheet = ss.getSheetByName(fileManagedSheetName)
  
  try {
    // csv を取得・パースして配列化
    const file = DriveApp.getFilesByName(filePath).next()
    const data = file.getBlob().getDataAsString("Shift_JIS")
    const csv = Utilities.parseCsv(data)
    Logger.log('csv.firstRow: ', csv[0])

    // データがなかった場合は最新を2行目に設定
    const latestDataRow = (dataSheet.getLastRow() > 1) ? dataSheet.getLastRow() : 2
    let importedFiles = []
    // インポート先のテーブルを初期化
    dataSheet.deleteRows(2, latestDataRow - 1)
    dataSheet.insertRowsAfter(2, latestDataRow - 1)

    // 収支データのインポート処理
    for (let r = 1; r < csv.length; r ++) {
      for (let c = 0; c < csv[0].length; c ++) {
        // Google スプレッドシートの列番号は 1 から始まるので CSV と差がある
        dataSheet.getRange(r + 1, c + 1).setValue(csv[r][c])
      }
    }
  
    // インポート日時を取得
    const timeDifference = (new Date().getTimezoneOffset() + (9 * 60)) * 60 * 1000
    const nowDate = new Date(Date.now() + timeDifference)
    Logger.log('nowDate: ', nowDate)
    const importedAt = [
      [nowDate.getFullYear(), nowDate.getMonth() + 1, nowDate.getDate()].join('/'),
      // タイムゾーン設定ができないのでハードコーディング
      [nowDate.getHours(), nowDate.getMinutes(), nowDate.getSeconds()].join(':')
    ].join(' ')
    
    // インポート履歴への入力データ
    importedFiles.push([
      importedAt, 
      file.getName(), 
      file.getDateCreated(), 
      file.getLastUpdated(), 
      file.getSize()
    ])

    // インポートファイル管理表への記入
    const latestFileManagementRow = fileManagementSheet.getLastRow()
    const fmRow = fileManagementSheet.getRange(latestFileManagementRow + 1, 1, 1, importedFiles[0].length)
    fmRow.setValues(importedFiles)
  } catch(err) {
    MailApp.sendEmail(
      mailAddress, 
      'CSV のインポートに失敗しました', 
      err.message
    )
  }
};
