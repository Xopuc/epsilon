/**
 * Функция запуска скрипта обновления себестоимостей на кнопке
 */
function refreshProfitButton() {
  gif("Загрузка...");

  fifoMethod();

  closeForm("Готово");
}

/**
 * Функция расчёта себестоимости по методу FIFO
 */
function fifoMethod() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))';
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  // Лист "Отгрузки"
  const shippingSheet = sheets.find(a => a.properties.sheetId == SHIPPING_SHEET_ID)
  const shippingSheetName = shippingSheet.properties.title
  const shippingSheetLastCol = shippingSheet.properties.gridProperties.columnCount
  const shippingSheetLastRow = shippingSheet.properties.gridProperties.rowCount

  // Лист "Закупки"
  const purchasingSheet = sheets.find(a => a.properties.sheetId == PURCHASING_SHEET_ID)
  const purchasingSheetName = purchasingSheet.properties.title
  const purchasingSheetLastCol = purchasingSheet.properties.gridProperties.columnCount
  const purchasingSheetLastRow = purchasingSheet.properties.gridProperties.rowCount

  // Лист "Размещения"
  const placementsSheet = sheets.find(a => a.properties.sheetId == PLACEMENTS_SHEET_ID)
  const placementsSheetName = placementsSheet.properties.title
  const placementsSheetLastCol = placementsSheet.properties.gridProperties.columnCount
  const placementsSheetLastRow = placementsSheet.properties.gridProperties.rowCount


  const batchData = Sheets.Spreadsheets.Values.batchGet(
    MAIN_SS_ID,
    {
      ranges: [
        `${purchasingSheetName}!R${PURCHASING_DATA_FIRST_ROW}C${PURCHASING_DATA_FIRST_COL}:R${purchasingSheetLastRow}C${purchasingSheetLastCol - PURCHASING_DATA_LAST_COL_OFFSET}`,
        `${shippingSheetName}!R${SHIPPING_DATA_FIRST_ROW}C${SHIPPING_DATA_FIRST_COL}:R${shippingSheetLastRow}C${shippingSheetLastCol - SHIPPING_DATA_LAST_COL_OFFSET}`,
        `${placementsSheetName}!R${PLACEMENTS_DATA_FIRST_ROW}C${PLACEMENTS_DATA_FIRST_COL}:R${placementsSheetLastRow}C${placementsSheetLastCol - PLACEMENTS_DATA_LAST_COL_OFFSET}`
      ],
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).valueRanges
  const purchasingData = batchData[0].values ? batchData[0].values : [] // Данные с листа "Закупки"
  const shippingData = batchData[1].values ? batchData[1].values : []   // Данные с листа "Отгрузки"
  const placementsData = batchData[2].values ? batchData[2].values : [] // Данные с листа "Размещения"

  const extMass = []
  const shippingDataWithIndexes = shippingData.map((a, i) => {
    extMass.push(['', '', '', ''])  // Создание пустого массива себестоимостей
    for (let c = a.length; c < SHIPPING_DATA_PROFIT_LAST_COL; c++) {
      a.push('')  // Выравнивание массива по столбцам
    }
    a.push(i) // Добавление индекса строкам
    return a
  })

  const placementSebesMass = []
  const placementsDataWithIndexes = placementsData.map((a, i) => {
    placementSebesMass.push(['', ''])
    for (let c = a.length; c < placementsSheetLastCol - PLACEMENTS_DATA_LAST_COL_OFFSET; c++) {
      a.push('')
    }
    a.push(i) // добавляем положение строки в изначальном массиве
    return a
  })

  // Фильтр операций с датой
  const purchasingDataFilteredByDateCol = purchasingData.filter(row => row[0])
  const placementsDataFilteredByDateCol = placementsDataWithIndexes.filter(row => row[0])
  const shippingDataFilteredByDateCol = shippingDataWithIndexes.filter(row => row[0])

  const funcAscendingSortByDate = function (a, b) {
    const aParts = a[0].split('.')
    const bParts = b[0].split('.')
    a = new Date(aParts[2], aParts[1] - 1, aParts[0]).getTime();
    b = new Date(bParts[2], bParts[1] - 1, bParts[0]).getTime();

    if (a > b)
      return 1;
    else if (a < b)
      return -1;
    else
      return 0;
  }

  // Сортировка операций по дате по возрастанию
  const purchasingDataSorted = purchasingDataFilteredByDateCol.sort(funcAscendingSortByDate)
  const placementsDataSorted = placementsDataFilteredByDateCol.sort(funcAscendingSortByDate)
  const shippingDataSorted = shippingDataFilteredByDateCol.sort(funcAscendingSortByDate)

  const placeExtMass = []

  for (let i = 0; i < placementsDataSorted.length; i++) {
    const date = placementsDataSorted[i][0]
    const name = getStrValue(placementsDataSorted[i][11])
    let quantity = getNumValue(placementsDataSorted[i][6])
    const baseQuantity = placementsDataSorted[i][6]
    const placement = placementsDataSorted[i][7]
    const index = placementsDataSorted[i][15]

    let priceCost = 0;  // Себестоимость товара за ед.
    let priceSum = 0;   // Сумма себестоимости всего количества товара

    for (let j = 0; j < purchasingDataSorted.length; j++) {
      const PURCHASE_NAME = 15
      const PURCHASE_PRICE = 6
      const PURCHASE_QUANTITY = 7

      purchasingDataSorted[j][PURCHASE_PRICE] = getNumValue(purchasingDataSorted[j][PURCHASE_PRICE]);
      purchasingDataSorted[j][PURCHASE_QUANTITY] = getNumValue(purchasingDataSorted[j][PURCHASE_QUANTITY]);

      if (name != purchasingDataSorted[j][PURCHASE_NAME])
        continue;

      let quantityTaked = purchasingDataSorted[j][PURCHASE_QUANTITY] - quantity;
      if (quantityTaked < 0) {
        quantity -= purchasingDataSorted[j][PURCHASE_QUANTITY];
        quantityTaked = purchasingDataSorted[j][PURCHASE_QUANTITY];
        purchasingDataSorted[j][PURCHASE_QUANTITY] = 0;
      } else if (quantityTaked >= 0) {
        purchasingDataSorted[j][PURCHASE_QUANTITY] -= quantity;
        quantityTaked = quantity;
        quantity = 0;
      }

      priceSum += quantityTaked * purchasingDataSorted[j][PURCHASE_PRICE];

      if (quantity == 0)
        break;
    }

    if (quantity > 0) {
      priceCost = "Для расчёта";
      priceSum = "не хватает товара";
    } else {
      quantity = getNumValue(baseQuantity);
      priceCost = quantity == 0 ? 0 : priceSum / quantity;
    }

    placeExtMass.push(  // Массив для расчета себестоимости на листе "Отгрузки"
      [
        date,
        name,
        baseQuantity,
        priceCost,
        priceSum,
        placement,
        `${name}::${placement}`
      ]
    )

    placementSebesMass[index] = [ // Массив себестоимостей на листе "Размещения"
      priceCost,
      priceSum,
    ]
  }

  const purchasingExtMass = []

  for (let i = 0; i < purchasingData.length; i++) {
    purchasingExtMass.push(
      [
        purchasingData[i][0],         // Дата
        purchasingData[i][15],        // Наименование
        purchasingData[i][7],         // Количество
        purchasingData[i][6],         // Себестоимость за единицу
        purchasingData[i][8],         // Сумма себестоимости
        '',                           // Место (пустое, т.к. это закупки)
        `${purchasingData[i][15]}::`  // Наименование::Место
      ]
    )
  }

  const purchasingPlaceMass = purchasingExtMass.concat(placeExtMass)  // Объединение массива закупок и размещений
  const purchasingPlaceSortedData = purchasingPlaceMass.sort(funcAscendingSortByDate)

  for (let i = 0; i < shippingDataSorted.length; i++) {
    const name = getStrValue(shippingDataSorted[i][20])
    let quantity = getNumValue(shippingDataSorted[i][10])
    const baseQuantity = getNumValue(shippingDataSorted[i][10])
    const saleSum = getNumValue(shippingDataSorted[i][11])
    const index = shippingDataSorted[i][22]
    const placement = getStrValue(shippingDataSorted[i][6])
    const nameAndPlacement = `${name}::${placement}`

    let priceCost = 0;      // Себестоимость товара за ед.
    let priceSum = 0;       // Сумма себестоимости всего количества товара
    let profit = 0;         // Прибыль
    let profitability = 0;  // Рентабельность

    for (let j = 0; j < purchasingPlaceSortedData.length; j++) {
      const PURCHASE_NAME = 1
      const PURCHASE_QUANTITY = 2
      const PURCHASE_PRICE = 3
      const PURCHASE_SUM = 4
      const PURCHASE_PLACEMENT = 5
      const PURCHASE_NAME_AND_PLACEMENT = 6

      purchasingPlaceSortedData[j][PURCHASE_PRICE] = getNumValue(purchasingPlaceSortedData[j][PURCHASE_PRICE]);
      purchasingPlaceSortedData[j][PURCHASE_QUANTITY] = getNumValue(purchasingPlaceSortedData[j][PURCHASE_QUANTITY]);

      if (nameAndPlacement != purchasingPlaceSortedData[j][PURCHASE_NAME_AND_PLACEMENT])
        continue;

      let quantityTaked = purchasingPlaceSortedData[j][PURCHASE_QUANTITY] - quantity;
      if (quantityTaked < 0) {
        quantity -= purchasingPlaceSortedData[j][PURCHASE_QUANTITY];
        quantityTaked = purchasingPlaceSortedData[j][PURCHASE_QUANTITY];
        purchasingPlaceSortedData[j][PURCHASE_QUANTITY] = 0;
      } else if (quantityTaked >= 0) {
        purchasingPlaceSortedData[j][PURCHASE_QUANTITY] -= quantity;
        quantityTaked = quantity;
        quantity = 0;
      }

      priceSum += quantityTaked * purchasingPlaceSortedData[j][PURCHASE_PRICE];

      if (quantity == 0)
        break;
    }

    if (quantity > 0) {
      priceCost = "Для расчёта";
      priceSum = "не хватает товара";
      profit = "на листе \"Закупки\"";
      profitability = "";
    } else {
      quantity = getNumValue(baseQuantity);
      priceCost = quantity == 0 ? 0 : priceSum / quantity;
      profit = saleSum - priceSum;
      profitability = priceSum == 0 ? 0 : profit / priceSum;
    }

    extMass[index] = [
      priceCost,
      priceSum,
      profit,
      profitability
    ]
  }

  if (!extMass.length) {
    extMass.push([''])
  }
  if (!placementSebesMass.length) {
    placementSebesMass.push([''])
  }

  Sheets.Spreadsheets.Values.batchUpdate(
    {
      valueInputOption: 'USER_ENTERED',
      data: [
        {
          range: `${placementsSheetName}!R${PLACEMENTS_DATA_FIRST_ROW}C${PLACEMENTS_DATA_SEBES_COL}:R${placementSebesMass.length + PLACEMENTS_DATA_FIRST_ROW - 1}C${placementSebesMass[0].length + PLACEMENTS_DATA_SEBES_COL - 1}`,
          majorDimension: 'ROWS',
          values: placementSebesMass
        },  // Запись массива себестоимостей размещений
        {
          range: `${shippingSheetName}!R${SHIPPING_DATA_FIRST_ROW}C${SHIPPING_DATA_PROFIT_COL}:R${extMass.length + SHIPPING_DATA_FIRST_ROW - 1}C${extMass[0].length + SHIPPING_DATA_PROFIT_COL - 1}`,
          majorDimension: 'ROWS',
          values: extMass
        }   // Запись массива себестоимостей отгрузок
      ]
    },
    MAIN_SS_ID
  )
}