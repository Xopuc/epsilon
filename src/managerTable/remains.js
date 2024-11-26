/**
 * Функция на кнопке для создания файла Excel из остатков
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Стандартный объект ивента Edit
 */
function createExcelFromRemainsTrigger(e) {
  const editedRange = e.range
  const editedSheet = editedRange.getSheet()
  const editedSheetId = editedSheet.getSheetId()
  if (editedSheetId != REMAINS_SHEET_ID) {
    return
  }
  const editedA1Notation = editedRange.getA1Notation()
  if (editedA1Notation != 'D2') {
    return
  }
  const editedValue = e.value
  if (editedValue != 'TRUE') {
    return
  }
  const lock = LockService.getScriptLock()
  const success = lock.tryLock(1);
  if (!success) {
    closeForm('Скрипт уже запущен')
    lock.releaseLock()
  } else {
    gif('Загрузка')
    createExcelFromRemains()
    editedRange.setValue(false)
    closeForm('Готово')
    lock.releaseLock()
  }
}

/**
 * Функция для создания Excel файла из остатков и запись данных о файле в реестр
 */
function createExcelFromRemains() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  const remainsSheet = sheets.find(a => a.properties.sheetId == REMAINS_SHEET_ID)
  copyRemainsDataToTempSheet(remainsSheet)

  const name = Utilities.formatDate(new Date(), 'GMT+3', 'dd.MM.yyyy HH:mm')
  const fileUrl = exportSheetAsXLStoDrive(REMAINS_TECH_SS_ID, REMAINS_TECH_SHEET_ID, name, EXCEL_REMAINS_FOLDER_ID)

  const registrySheet = sheets.find(a => a.properties.sheetId == EXCEL_REGISTRY_SHEET_ID)
  writeExcelFileDataToRegistry(name, fileUrl, registrySheet)
}

/**
 * Функция для копирования данных с листа "Остатки" на лист "Остатки (для выгрузки)"
 * @param {GoogleAppsScript.Sheets.Schema.Sheet} remainsSheet Лист "Остатки"
 */
function copyRemainsDataToTempSheet(remainsSheet) {
  const remainsSheetName = remainsSheet.properties.title
  const remainsSheetLastCol = remainsSheet.properties.gridProperties.columnCount
  const remainsSheetLastRow = remainsSheet.properties.gridProperties.rowCount

  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(REMAINS_TECH_SS_ID, { fields })

  const techRemainsSheet = sheets.find(a => a.properties.sheetId == REMAINS_TECH_SHEET_ID)
  const techRemainsSheetName = techRemainsSheet.properties.title
  const techRemainsSheetLastCol = techRemainsSheet.properties.gridProperties.columnCount
  const techRemainsSheetLastRow = techRemainsSheet.properties.gridProperties.rowCount

  const remainsData = Sheets.Spreadsheets.Values.get(
    MAIN_SS_ID,
    `${remainsSheetName}!R${REMAINS_DATA_FIRST_ROW}C${REMAINS_DATA_FIRST_COL}:R${remainsSheetLastRow}C${REMAINS_DATA_LAST_COL}`,
    {
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).values

  if (remainsData) {
    const remainsDataFixed = remainsData.map(a => {
      for (let j = a.length; j < REMAINS_TECH_DATA_LAST_COL; j++) {
        a.push('')
      }
      return a
    })
    for (let i = remainsDataFixed.length + REMAINS_TECH_DATA_FIRST_ROW; i <= techRemainsSheetLastRow; i++) {
      const tmp = []
      for (let j = REMAINS_TECH_DATA_FIRST_COL; j <= techRemainsSheetLastCol; j++) {
        tmp.push('')
      }
      remainsDataFixed.push(tmp)
    }
    Sheets.Spreadsheets.Values.update(
      {
        majorDimension: 'ROWS',
        values: remainsDataFixed
      },
      REMAINS_TECH_SS_ID,
      `${remainsSheetName}!R${REMAINS_TECH_DATA_FIRST_ROW}C${REMAINS_TECH_DATA_FIRST_COL}:R${remainsDataFixed.length + REMAINS_TECH_DATA_FIRST_ROW - 1}C${REMAINS_TECH_DATA_LAST_COL}`,
      {
        valueInputOption: 'USER_ENTERED'
      }
    )
  }
}

/**
 * Функция для создания XLSX файла из листа таблицы
 * @param {string} spreadsheetId ID таблицы из которой импортировать лист
 * @param {string} sheetId ID листа в таблице, который нужно импортировать
 * @param {string} name Название конечного файла
 * @param {string} destinationFolderId ID папки, в которую нужно сохранить файл
 * @returns {string} Ссылка на созданный XLSX файл
 */
function exportSheetAsXLStoDrive(spreadsheetId, sheetId, name, destinationFolderId) {
  let response = UrlFetchApp.fetch(
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?gid=${sheetId}&format=xlsx`,
    {
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      },
    })
  const blob = response.getBlob()
  const file = Drive.Files.insert(
    {
      kind: 'drive#file',
      title: name,
      parents: [{ id: destinationFolderId }],
    },
    blob
  )
  return `https://docs.google.com/spreadsheets/d/${file.id}`
}

/**
 * Функция для записи данных о созданном файле в реестр
 * @param {string} fileName Название файла Excel
 * @param {string} fileUrl URL файла Excel
 * @param {GoogleAppsScript.Sheets.Schema.Sheet} registrySheet Лист "Реестр файлов"
 */
function writeExcelFileDataToRegistry(fileName, fileUrl, registrySheet) {
  const registrySheetName = registrySheet.properties.title
  const registrySheetLastCol = registrySheet.properties.gridProperties.columnCount
  const registrySheetLastRow = registrySheet.properties.gridProperties.rowCount

  const existingData = Sheets.Spreadsheets.Values.get(
    MAIN_SS_ID,
    `${registrySheetName}!R${EXCEL_REGISTRY_DATA_FIRST_ROW}C${EXCEL_REGISTRY_DATA_FIRST_COL}:R${registrySheetLastRow}C${registrySheetLastCol}`,
    {
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).values

  if (!existingData) {
    const newExtMass = [[fileName, fileUrl]]
    Sheets.Spreadsheets.Values.update(
      {
        majorDimension: 'ROWS',
        values: newExtMass
      },
      MAIN_SS_ID,
      `${registrySheetName}!R${EXCEL_REGISTRY_DATA_FIRST_ROW}C${EXCEL_REGISTRY_DATA_FIRST_COL}:R${newExtMass.length + EXCEL_REGISTRY_DATA_FIRST_ROW - 1}C${newExtMass[0].length + EXCEL_REGISTRY_DATA_FIRST_COL - 1}`,
      {
        valueInputOption: 'USER_ENTERED'
      }
    )
    return
  }

  existingData.push([
    fileName,
    fileUrl
  ])
  Sheets.Spreadsheets.Values.update(
    {
      majorDimension: 'ROWS',
      values: existingData
    },
    MAIN_SS_ID,
    `${registrySheetName}!R${EXCEL_REGISTRY_DATA_FIRST_ROW}C${EXCEL_REGISTRY_DATA_FIRST_COL}:R${existingData.length + EXCEL_REGISTRY_DATA_FIRST_ROW - 1}C${existingData[0].length + EXCEL_REGISTRY_DATA_FIRST_COL - 1}`,
    {
      valueInputOption: 'USER_ENTERED'
    }
  )
}