/**
 * Функция на кнопке для проведения размещения
 */
function performPlacementButton() {
  gif('Загрузка')
  performPlacement()
  fifoMethod()
  closeForm('Готово')
}

/**
 * Функция для проведения размещения
 */
function performPlacement() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  const unplacedRemainsSheet = sheets.find(a => a.properties.sheetId == UNPLACED_REMAINS_SHEET_ID)
  const unplacedRemainsSheetName = unplacedRemainsSheet.properties.title
  const unplacedRemainsSheetLastCol = unplacedRemainsSheet.properties.gridProperties.columnCount
  const unplacedRemainsSheetLastRow = unplacedRemainsSheet.properties.gridProperties.rowCount

  const placementsSheet = sheets.find(a => a.properties.sheetId == PLACEMENTS_SHEET_ID)
  const placementsSheetName = placementsSheet.properties.title
  const placementsSheetLastCol = placementsSheet.properties.gridProperties.columnCount
  const placementsSheetLastRow = placementsSheet.properties.gridProperties.rowCount


  const batchData = Sheets.Spreadsheets.Values.batchGet(
    MAIN_SS_ID,
    {
      ranges: [
        `${unplacedRemainsSheetName}!R${UNPLACED_REMAINS_DATA_FIRST_ROW}C${UNPLACED_REMAINS_DATA_FIRST_COL}:R${unplacedRemainsSheetLastRow}C${UNPLACED_REMAINS_DATA_LAST_COL}`,
        `${placementsSheetName}!R${PLACEMENTS_DATA_FIRST_ROW}C${PLACEMENTS_DATA_FIRST_COL}:R${placementsSheetLastRow}C${PLACEMENTS_DATA_LAST_COL}`,
      ],
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).valueRanges

  const unplacedRemainsData = batchData[0].values ? batchData[0].values : []
  const placementsData = batchData[1].values ? batchData[1].values : []

  let isAmountOver = false
  unplacedRemainsData.forEach(a => {
    if (a[6] && a[5] - a[6] < 0) {
      isAmountOver = true
      return
    }
  })
  if (isAmountOver) {
    alertDialog('У одной из позиций количество для размещения больше, чем есть на складе!')
    return false
  }

  const placementsDataFixed = placementsData.map(a => {
    for (let k = a.length; k < PLACEMENTS_DATA_LAST_COL; k++) {
      a.push('')
    }
    return a
  })

  const todayString = Utilities.formatDate(new Date(), 'GMT+3', 'dd.MM.yyyy')

  const unplacedRemainsAdded = []
  for (let i = 0; i < unplacedRemainsData.length; i++) {
    if (!unplacedRemainsData[i][7] || unplacedRemainsData[i][5] - unplacedRemainsData[i][6] < 0) {
      continue
    }
    unplacedRemainsAdded.push([
      todayString,
      unplacedRemainsData[i][0],
      unplacedRemainsData[i][1],
      unplacedRemainsData[i][2],
      unplacedRemainsData[i][3],
      unplacedRemainsData[i][4],
      unplacedRemainsData[i][6],
      unplacedRemainsData[i][7],
      ''
    ])
  }

  const placedData = placementsDataFixed.concat(unplacedRemainsAdded)

  const emptyMass = []
  for (let i = UNPLACED_REMAINS_DATA_FIRST_ROW; i <= unplacedRemainsSheetLastRow; i++) {
    const tmp = []
    for (let j = UNPLACED_REMAINS_DATA_EMPTY_FIRST_COL; j <= UNPLACED_REMAINS_DATA_LAST_COL; j++) {
      tmp.push('')
    }
    emptyMass.push(tmp)
  }

  Sheets.Spreadsheets.Values.batchUpdate(
    {
      valueInputOption: 'USER_ENTERED',
      data: [
        {
          range: `${placementsSheetName}!R${PLACEMENTS_DATA_FIRST_ROW}C${PLACEMENTS_DATA_FIRST_COL}:R${placedData.length + PLACEMENTS_DATA_FIRST_ROW - 1}C${PLACEMENTS_DATA_LAST_COL}`,
          majorDimension: 'ROWS',
          values: placedData
        },  // Запись массива размещений
        {
          range: `${unplacedRemainsSheetName}!R${UNPLACED_REMAINS_DATA_FIRST_ROW}C${UNPLACED_REMAINS_DATA_EMPTY_FIRST_COL}:R${emptyMass.length + UNPLACED_REMAINS_DATA_FIRST_ROW - 1}C${emptyMass[0].length + UNPLACED_REMAINS_DATA_EMPTY_FIRST_COL - 1}`,
          majorDimension: 'ROWS',
          values: emptyMass
        }   // Запись массива пустых ячеек в СВХ
      ]
    },
    MAIN_SS_ID
  )
}