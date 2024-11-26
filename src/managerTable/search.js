/**
 * Функция на кнопке для поиска остатков
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e Стандартный объект ивента Edit
 */
function searchRefreshTrigger(e) {
  const editedRange = e.range
  const editedSheet = editedRange.getSheet()
  const editedSheetId = editedSheet.getSheetId()
  if (editedSheetId != SEARCH_SHEET_ID) {
    return
  }
  const editedA1Notation = editedRange.getA1Notation()
  if (editedA1Notation != 'G3') {
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
    if (searchRefresh()) {
      editedRange.setValue(false)
      closeForm('Готово')
    } else {
      editedRange.setValue(false)
    }
    lock.releaseLock()
  }
}

/**
 * Функция для поиска остатков
 * @returns {boolean} Возвращает false, если не выбран фильтр
 */
function searchRefresh() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  const searchSheet = sheets.find(a => a.properties.sheetId == SEARCH_SHEET_ID)
  const searchSheetName = searchSheet.properties.title
  const searchSheetLastCol = searchSheet.properties.gridProperties.columnCount
  const searchSheetLastRow = searchSheet.properties.gridProperties.rowCount

  const remainsSheet = sheets.find(a => a.properties.sheetId == REMAINS_SHEET_ID)
  const remainsSheetName = remainsSheet.properties.title
  const remainsSheetLastCol = remainsSheet.properties.gridProperties.columnCount
  const remainsSheetLastRow = remainsSheet.properties.gridProperties.rowCount

  const batchData = Sheets.Spreadsheets.Values.batchGet(
    MAIN_SS_ID,
    {
      ranges: [
        `${searchSheetName}!R${SEARCH_FILTER_ROW}C${SEARCH_FILTER_FIRST_COL}:R${SEARCH_FILTER_ROW}C${SEARCH_FILTER_LAST_COL}`,
        `${remainsSheetName}!R${REMAINS_DATA_FIRST_ROW}C${REMAINS_DATA_FIRST_COL}:R${remainsSheetLastRow}C${REMAINS_DATA_LAST_COL}`,
      ],
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).valueRanges

  const searchFilterData = batchData[0].values ? batchData[0].values : []
  const remainsData = batchData[1].values ? batchData[1].values : []

  if (!searchFilterData.length) {
    alertDialog('Ни один из фильтров не выбран!')
    return false
  }
  const filter = {
    prefix: searchFilterData[0][0],
    name: searchFilterData[0][1],
    suffixFirst: searchFilterData[0][2],
    suffixSecond: searchFilterData[0][3],
    maker: searchFilterData[0][4]
  }

  const remainsDataFixed = remainsData.map(a => {
    for (let k = a.length; k < REMAINS_DATA_LAST_COL; k++) {
      a.push('')
    }
    return [
      a[0],
      a[1],
      a[2],
      a[3],
      a[4],
      a[5],
      a[6],
      a[8],
    ]
  })

  const filteredRemains = remainsDataFixed.filter(item => {
    if (item[6] > 0) {
      return Object.keys(filter).every((key, i) => {
        return !filter[key] || item[i] === filter[key]
      })
    }
  })

  for (let i = filteredRemains.length; i < searchSheetLastRow - SEARCH_DATA_FIRST_ROW + 1; i++) {
    const tmp = []
    for (let j = 0; j < searchSheetLastCol; j++) {
      tmp.push('')
    }
    filteredRemains.push(tmp)
  }

  if (filteredRemains.length) {
    Sheets.Spreadsheets.Values.update(
      {
        majorDimension: 'ROWS',
        values: filteredRemains
      },
      MAIN_SS_ID,
      `${searchSheetName}!R${SEARCH_DATA_FIRST_ROW}C${SEARCH_DATA_FIRST_COL}:R${filteredRemains.length + SEARCH_DATA_FIRST_ROW - 1}C${filteredRemains[0].length + SEARCH_DATA_FIRST_COL - 1}`,
      {
        valueInputOption: 'USER_ENTERED'
      }
    )
  }
  return true
}