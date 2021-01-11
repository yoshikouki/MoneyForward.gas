/** @OnlyCurrentDoc */
const dataSheetName = 'cashFlow'
const fileManagedSheetName = 'imortedFiles'
const dirPath = 'Money/CashFlow/'
const filePath = '収入・支出詳細_2020-04-01_2020-04-30.csv'
const fileSearchQuery = 'title contains "収入・支出詳細_"'
const mailAddress = 'yoshikouki@gmail.com'

function importCashFlowFromCsv() {
  // スプレッドシートとデータシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(dataSheetName);
  const fileManagementSheet = ss.getSheetByName(fileManagedSheetName)
  
  try {
    // csv を取得・パースして配列化
    const csvFiles = DriveApp.searchFiles(fileSearchQuery)
    // インポート日時を取得
    const importedAt = getImportedTime()
    // 初期化
    let csv = []
    let importedFiles = []    

    while (csvFiles.hasNext()) {
      let file = csvFiles.next()
      let dataString = file.getBlob().getDataAsString('Shift_JIS')
      let data = Utilities.parseCsv(dataString)
      data.shift()
      csv.push(...data)

      Logger.log('parse: ', file.getName())

      // インポート履歴への入力データ
      importedFiles.push([
        importedAt, 
        file.getName(), 
        file.getDateCreated(), 
        file.getLastUpdated(), 
        file.getSize()
      ])
    }
    csv.sort(function (a, b) {
      let date1 = new Date(a[1]).getTime()
      let date2 = new Date(b[1]).getTime()
      return(date1 - date2)
    })

    // データがなかった場合は最新を2行目に設定
    const latestDataRow = (dataSheet.getLastRow() > 1) ? dataSheet.getLastRow() : 2
    // インポート先のテーブルを初期化
    dataSheet.deleteRows(2, latestDataRow - 1)
    dataSheet.insertRowsAfter(2, latestDataRow - 1)
    // 収支データをテーブルへ入力
    dataSheet.getRange(2, 1, csv.length, csv[0].length).setValues(csv)
  
    // インポートファイル管理表への入力
    const latestFileManagementRow = fileManagementSheet.getLastRow()
    const fmRow = fileManagementSheet.getRange(
      latestFileManagementRow + 1, 
      1, 
      importedFiles.length, 
      importedFiles[0].length
    )
    fmRow.setValues(importedFiles)
  } catch(err) {
    MailApp.sendEmail(
      mailAddress, 
      'CSV のインポートに失敗しました', 
      err.message
    )
  }
};

function getImportedTime() {
  const timeDifference = (new Date().getTimezoneOffset() + (9 * 60)) * 60 * 1000
  const nowDate = new Date(Date.now() + timeDifference)
  Logger.log('nowDate: ', nowDate)
  return [
    [nowDate.getFullYear(), nowDate.getMonth() + 1, nowDate.getDate()].join('/'),
    // タイムゾーン設定ができないのでハードコーディング
    [nowDate.getHours(), nowDate.getMinutes(), nowDate.getSeconds()].join(':')
  ].join(' ')
}
