/** @OnlyCurrentDoc */
const dataSheetName = 'cashFlow'
const cfTableName = 'cfTable'
const fileManagedSheetName = 'imortedFiles'
const filePath = '収入・支出詳細_2020-04-01_2020-04-30.csv'
const fileSearchQuery = 'title contains "収入・支出詳細_"'
const mailAddress = 'yoshikouki@gmail.com'

// MoneyForward csv の列番号を定義
const calculatedCol = 0
const dateCol = 1
const titleCol = 2
const amountCol = 3
const finantialInstitutionCol = 4
const majorClassificationCol = 5
const middleClassificationCol = 6

function importCashFlowFromCsv() {
  // スプレッドシートとデータシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(dataSheetName);
  const cfTable = ss.getSheetByName(cfTableName);
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
      // csv のヘッダーを削除
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

    // CSV を日付順にソート
    csv.sort(function (a, b) {
      let date1 = new Date(a[1]).getTime()
      let date2 = new Date(b[1]).getTime()
      return(date1 - date2)
    })

    // データがなかった場合は最新を2行目に設定
    const latestDataRow = (dataSheet.getLastRow() > 1) ? dataSheet.getLastRow() : 2
    // インポート先のテーブルを初期化
    dataSheet.getRange(2, latestDataRow).clear
    // 収支データをテーブルへ入力
    dataSheet.getRange(2, 1, csv.length, csv[0].length).setValues(csv)

    // csv を cfTable 用に入れ替えして入力
    const cfTableValues = csv.map((row) => {
      let date = new Date(row[dateCol])
      return [
        date.getFullYear(),
        date.getMonth() + 1,
        date.getDate(),
        row[titleCol],
        row[amountCol],
        row[finantialInstitutionCol],
        row[majorClassificationCol],
        row[middleClassificationCol],
        row[calculatedCol],
      ]
    })
    // すでに立替・家計が入力されているので3列目から入力する
    cfTable.getRange(2, 3, cfTableValues.length, cfTableValues[0].length).setValues(cfTableValues)
  
    // インポート履歴をファイル名順でソート
    importedFiles.sort(function(a,b) {
      let fileNameA = a[1]
      let fileNameB = b[1]
      if (fileNameA < fileNameB) {
        return -1
      } else if (fileNameA > fileNameB) {
        return 1
      } else {
        return 0
      }
    })
    // インポート履歴への入力
    fileManagementSheet.getRange(
      fileManagementSheet.getLastRow() + 1, 
      1, 
      importedFiles.length, 
      importedFiles[0].length
    ).setValues(importedFiles)
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
