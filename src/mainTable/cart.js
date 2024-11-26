/**
 * Функция на кнопке для переноса данных с листа "Поиск" в "Корзину"
 */
function addToCartButton() {
  gif('Загрузка')
  if (addToCart()) {
    closeForm('Готово')
  }
}

/**
 * Функция для переноса данных с листа "Поиск" в "Корзину"
 * @returns {boolean} False, если данных на листе "Поиск" нет, либо если ничего не выбрано
 */
function addToCart() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  const searchSheet = sheets.find(a => a.properties.sheetId == SEARCH_SHEET_ID)
  const searchSheetName = searchSheet.properties.title
  const searchSheetLastCol = searchSheet.properties.gridProperties.columnCount
  const searchSheetLastRow = searchSheet.properties.gridProperties.rowCount

  const cartSheet = sheets.find(a => a.properties.sheetId == CART_SHEET_ID)
  const cartSheetName = cartSheet.properties.title
  const cartSheetLastCol = cartSheet.properties.gridProperties.columnCount
  const cartSheetLastRow = cartSheet.properties.gridProperties.rowCount

  const batchData = Sheets.Spreadsheets.Values.batchGet(
    MAIN_SS_ID,
    {
      ranges: [
        `${searchSheetName}!R${SEARCH_DATA_FIRST_ROW}C${SEARCH_DATA_FIRST_COL}:R${searchSheetLastRow}C${searchSheetLastCol}`,
        `${cartSheetName}!R${CART_DATA_FIRST_ROW}C${CART_DATA_FIRST_COL}:R${cartSheetLastRow}C${cartSheetLastCol}`,
      ],
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).valueRanges

  const searchData = batchData[0].values ? batchData[0].values : []
  const cartData = batchData[1].values ? batchData[1].values : []

  if (!searchData) {
    alertDialog('Нечего переносить в корзину!')
    return false
  }

  const cartDataFixed = cartData.map(a => {
    for (let i = a.length; i < cartSheetLastCol; i++) {
      a.push('')
    }
    a[12] = ''
    return a
  })

  const searchDataFixed = searchData.map(a => {
    for (let i = a.length; i < searchSheetLastCol; i++) {
      a.push('')
    }
    return a
  })

  const searchDataCheckboxFiltered = searchDataFixed.filter(a => a[a.length - 1])

  if (!searchDataCheckboxFiltered.length) {
    alertDialog('Не выбрана ни одна позиция')
    return false
  }

  const currDateString = Utilities.formatDate(new Date(), 'GMT+3', 'dd.MM.yyyy')
  const searchDataToCart = searchDataCheckboxFiltered.map(a => {
    a.unshift(currDateString)
    a.unshift('')
    const tmpReccomendCost = a[a.length - 2]
    const tmpAmount = a[a.length - 3]
    a[a.length - 1] = tmpAmount
    a[a.length - 2] = ''
    a[a.length - 3] = tmpReccomendCost
    a.push('', '', '', '', '')
    return a
  })

  const cartAndSearchData = cartDataFixed.concat(searchDataToCart)

  for (let i = 0; i < cartSheetLastRow - CART_DATA_FIRST_ROW + 1; i++) {
    const tmp = []
    if (cartAndSearchData[i]) {
      for (let j = cartAndSearchData[i].length; j < cartSheetLastCol; j++) {
        cartAndSearchData[i].push('')
      }
      continue
    }
    for (let j = 0; j < cartSheetLastCol; j++) {
      tmp.push('')
    }
    cartAndSearchData.push(tmp)
  }

  const emptyCheckboxMass = []
  for (let i = SEARCH_DATA_FIRST_ROW; i <= searchSheetLastRow; i++) {
    emptyCheckboxMass.push([''])
  }

  if (cartAndSearchData.length) {
    Sheets.Spreadsheets.Values.batchUpdate(
      {
        valueInputOption: 'USER_ENTERED',
        data: [
          {
            range: `${searchSheetName}!R${SEARCH_DATA_FIRST_ROW}C${searchSheetLastCol}:R${emptyCheckboxMass.length + SEARCH_DATA_FIRST_ROW - 1}C${searchSheetLastCol}`,
            majorDimension: 'ROWS',
            values: emptyCheckboxMass
          },  // Запись массива пустых чекбоксов на лист поиска
          {
            range: `${cartSheetName}!R${CART_DATA_FIRST_ROW}C${CART_DATA_FIRST_COL}:R${cartAndSearchData.length + CART_DATA_FIRST_ROW - 1}C${cartAndSearchData[0].length + CART_DATA_FIRST_COL - 1}`,
            majorDimension: 'ROWS',
            values: cartAndSearchData
          }   // Запись массива корзины
        ]
      },
      MAIN_SS_ID
    )
  }
  return true
}

/**
 * Функция на кнопке для сохранения "Корзины" в "Отгрузки"
 */
function saveCartToShippingButton() {
  gif('Загрузка')
  if (saveCartToShipping()) {
    fifoMethod()
    closeForm('Готово')
  } else {
    closeForm()
  }
}

/**
 * Функция для переноса корзины на лист "Отгрузки"
 * @returns {boolean} false - если корзина пуста
 */
function saveCartToShipping() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  const cartSheet = sheets.find(a => a.properties.sheetId == CART_SHEET_ID)
  const cartSheetName = cartSheet.properties.title
  const cartSheetLastCol = cartSheet.properties.gridProperties.columnCount
  const cartSheetLastRow = cartSheet.properties.gridProperties.rowCount

  const shippingSheet = sheets.find(a => a.properties.sheetId == SHIPPING_SHEET_ID)
  const shippingSheetName = shippingSheet.properties.title
  const shippingSheetLastCol = shippingSheet.properties.gridProperties.columnCount
  const shippingSheetLastRow = shippingSheet.properties.gridProperties.rowCount

  const batchData = Sheets.Spreadsheets.Values.batchGet(
    MAIN_SS_ID,
    {
      ranges: [
        `${cartSheetName}!R${CART_DATA_FIRST_ROW}C${CART_DATA_FIRST_COL}:R${cartSheetLastRow}C${cartSheetLastCol}`,
        `${shippingSheetName}!R${SHIPPING_DATA_FIRST_ROW}C${SHIPPING_DATA_FIRST_COL}:R${shippingSheetLastRow}C${SHIPPING_DATA_PROFIT_COL - 1}`,
      ],
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).valueRanges

  const cartData = batchData[0].values ? batchData[0].values : []
  const shippingData = batchData[1].values ? batchData[1].values : []

  if (!cartData.length) {
    alertDialog('Нечего переносить в отгрузки!')
    return false
  }

  const isPlacementEmpty = cartData.every(a => a[7])
  if (!isPlacementEmpty) {
    alertDialog('У одной из позиций не заполнено место!')
    return false
  }

  const isAmountOver = cartData.every(a => a[10] - a[11] >= 0)
  if (!isAmountOver) {
    alertDialog('У одной из позиций количество отгрузки больше, чем есть на складе!')
    return false
  }

  const cartDataFixed = cartData.map(a => {
    a.shift()
    return a
  })

  const shippingAndCartData = shippingData.concat(cartDataFixed)

  const shippingAndCartDataFixed = shippingAndCartData.map(a => {
    for (let j = a.length; j < SHIPPING_DATA_PROFIT_COL - 1; j++) {
      a.push('')
    }
    a[7] = ''
    a[9] = ''
    a[11] = ''
    return a
  })

  const emptyCartMass = []
  for (let i = 0; i < cartSheetLastRow; i++) {
    const tmp = []
    for (let j = 0; j < cartSheetLastCol; j++) {
      tmp.push('')
    }
    emptyCartMass.push(tmp)
  }

  Sheets.Spreadsheets.Values.batchUpdate(
    {
      valueInputOption: 'USER_ENTERED',
      data: [
        {
          range: `${cartSheetName}!R${CART_DATA_FIRST_ROW}C${CART_DATA_FIRST_COL}:R${emptyCartMass.length + SEARCH_DATA_FIRST_ROW - 1}C${cartSheetLastCol}`,
          majorDimension: 'ROWS',
          values: emptyCartMass
        },  // Запись массива пустых значений на лист "Корзина"
        {
          range: `${shippingSheetName}!R${SHIPPING_DATA_FIRST_ROW}C${SHIPPING_DATA_FIRST_COL}:R${shippingAndCartDataFixed.length + SHIPPING_DATA_FIRST_ROW - 1}C${shippingAndCartDataFixed[0].length + SHIPPING_DATA_FIRST_COL - 1}`,
          majorDimension: 'ROWS',
          values: shippingAndCartDataFixed
        }   // Запись массива отгрузки
      ]
    },
    MAIN_SS_ID
  )
  return true
}

function deleteItemsFromCartButton() {
  gif('Загрузка')
  deleteItemsFromCart()
  closeForm('Готово')
}

function deleteItemsFromCart() {
  const fields = 'sheets(properties(sheetId,title,gridProperties(columnCount,rowCount)))'
  const { sheets } = Sheets.Spreadsheets.get(MAIN_SS_ID, { fields })

  const cartSheet = sheets.find(a => a.properties.sheetId == CART_SHEET_ID)
  const cartSheetName = cartSheet.properties.title
  const cartSheetLastCol = cartSheet.properties.gridProperties.columnCount
  const cartSheetLastRow = cartSheet.properties.gridProperties.rowCount

  const cartData = Sheets.Spreadsheets.Values.get(
    MAIN_SS_ID,
    `${cartSheetName}!R${CART_DATA_FIRST_ROW}C${CART_DATA_FIRST_COL}:R${cartSheetLastRow}C${cartSheetLastCol}`,
    {
      majorDimension: 'ROWS',
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    }
  ).values

  const isNotSelected = cartData.every(a => !a[0])
  if (isNotSelected) {
    alertDialog('Ничего не выбрано для удаления!')
    return
  }

  const deletedCartData = []
  for (let i = 0; i < cartData.length; i++) {
    if (cartData[i][0]) {
      continue
    }
    cartData[i][12] = ''
    deletedCartData.push(cartData[i])
  }

  for (let i = 0; i < cartSheetLastRow - CART_DATA_FIRST_ROW + 1; i++) {
    const tmp = []
    if (deletedCartData[i]) {
      for (let j = deletedCartData[i].length; j < cartSheetLastCol; j++) {
        deletedCartData[i].push('')
      }
      continue
    }
    for (let j = 0; j < cartSheetLastCol; j++) {
      tmp.push('')
    }
    deletedCartData.push(tmp)
  }

  Sheets.Spreadsheets.Values.update(
    {
      majorDimension: 'ROWS',
      values: deletedCartData
    },
    MAIN_SS_ID,
    `${cartSheetName}!R${CART_DATA_FIRST_ROW}C${CART_DATA_FIRST_COL}:R${deletedCartData.length + CART_DATA_FIRST_ROW - 1}C${deletedCartData.length + CART_DATA_FIRST_COL - 1}`,
    {
      valueInputOption: 'USER_ENTERED'
    }
  )
}