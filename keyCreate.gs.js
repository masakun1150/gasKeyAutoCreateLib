function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('日本語自動CamelCase変換', [
      { name: '変換', functionName: 'handleKeyCreate' },
    ]);
}

function handleKeyCreate() {
  const sheet = handleGetActiveSheet()

  const lastRow = sheet.getLastRow()

  const lastColumn = sheet.getLastColumn()

  for (let a = 1; a <= lastRow; a++) {
    for (let b = 1; b <= lastColumn; b++) {
      let setValue = handleTranslateWords(sheet.getRange(a, b).getValue())

      if (setValue) {
        setValue = handleTransCamel(setValue, false)

        if (setValue.length >= 30) {
          console.log(setValue)
          setValue = handleGetNumeronym(setValue)
          console.log(setValue)
        }
      } else {
        setValue = ''
      }

      handleSetValueSpredSheet(setValue, a, b)
    }
  }
}

function isNotEmptyItem(element) {
  if ( element === "" || element === undefined ) {
    return false;
  }
  return true;
}

function handleGetNumeronym(str) {
  str = str.replace(/\([^\)]*\)/g,'')

  str = str.split(/(^[a-z]+)|([A-Z][a-z]+)/).filter( isNotEmptyItem )

  console.log(str)

  return str.slice(0, 1) + (str.length - 2) + str.slice(-1)
}

function handleGetActiveSheet() {
  //1. 現在のスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //2. 現在のシートを取得
  const sheet = spreadsheet.getActiveSheet();

  return sheet
}

function handleTranslateWords(value) {

  if (value === '') {
    return
  }

  const transFrom = 'ja'
  const transTo = 'en'

  let transVal = LanguageApp.translate(value, transFrom, transTo)

  return transVal
}

function handleTransCamel(str, upper) {

  str = str.charAt(0).toLowerCase() + str.substring(1)

  str = str
    .replace(/^[\-_ ]/g, "")
    .replace(/[\-_ ]./g, function (match) {
      return match.charAt(1).toUpperCase();
    })

  return upper === true ?
    str.replace(/^[a-z]/g, function (match) {
      return match.toUpperCase();
    }) : str
}

function handleSetValueSpredSheet(value, row, column) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const setSheet = ss.getSheets()[1]
  setSheet.getRange(row, column).setValue(value)
}


