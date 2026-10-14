const PD = {
  SHEET_NAME:      'Продукты',
  HOSTESS_SHEET:   'food_hostess',
  WAREHOUSE_SHEET: 'Склад_Продукты',
  DELIVERY_SHEET:  'postavka_food',

  DATE_ROW:       1,
  HEADER_ROW:     2,
  DATA_START_ROW: 3,

  COLOR_HEADER:   '#1c4587',
  COLOR_TOTAL:    '#4a86e8',
  COLOR_ISSUED:   '#34A853',
  COLOR_DELIVERY: '#FF6D00',
  COLOR_STOCK:    '#6a329f',
  COLOR_ROW_ODD:  '#f8f9fa',
  COLOR_ROW_EVEN: '#ffffff'
};

function getOrCreateDashSheet(ss) {
  let sheet = SH_(ss, PD.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(L(PD.SHEET_NAME));
    Logger.log('✓ Создан лист "Продукты"');
  }
  return sheet;
}

function drawEmptyDashboard(sheet, coffeePoints, todayStr) {
  sheet.clear();
  const totalCols = 1 + coffeePoints.length + 4;

  sheet.getRange(PD.DATE_ROW, 1, 1, totalCols).merge();
  sheet.getRange(PD.DATE_ROW, 1)
    .setValue(todayStr)
    .setNumberFormat('@STRING@')
    .setBackground(PD.COLOR_HEADER)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(PD.DATE_ROW, 32);

  const headerValues = [L('Кофе-поинт:')];
  coffeePoints.forEach(cp => headerValues.push(cp.name));
  headerValues.push(L('ИТОГО'), L('ВЫДАНО'), L('Поставка'), L('Остаток на складе'));

  const headerRange = sheet.getRange(PD.HEADER_ROW, 1, 1, totalCols);
  headerRange.setValues([headerValues])
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(PD.COLOR_HEADER)
    .setBorder(true, true, true, true, true, true, '#ffffff', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(PD.HEADER_ROW, 30);

  const cpCount = coffeePoints.length;
  sheet.getRange(PD.HEADER_ROW, 2 + cpCount).setBackground(PD.COLOR_TOTAL);
  sheet.getRange(PD.HEADER_ROW, 3 + cpCount).setBackground(PD.COLOR_ISSUED);
  sheet.getRange(PD.HEADER_ROW, 4 + cpCount).setBackground(PD.COLOR_DELIVERY);
  sheet.getRange(PD.HEADER_ROW, 5 + cpCount).setBackground(PD.COLOR_STOCK);

  sheet.getRange(PD.DATA_START_ROW, 1, 1, totalCols).merge();
  sheet.getRange(PD.DATA_START_ROW, 1)
    .setValue(L('Данные появятся после первых заказов'))
    .setFontColor('#999999')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  formatColumns(sheet, cpCount, totalCols);
  sheet.setFrozenRows(2);

  Logger.log('✓ Пустой дашборд нарисован: ' + todayStr);
}

function refreshProductsDashboard() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss             = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet      = getOrCreateDashSheet(ss);
    const hostessSheet   = SH_(ss, PD.HOSTESS_SHEET);
    const warehouseSheet = SH_(ss, PD.WAREHOUSE_SHEET);

    const todayStr     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const coffeePoints = BaseInfo.getActiveCoffeePoints();

    const isFreshToday = _isDashboardForToday(dashSheet, todayStr);
    const issuedMap     = isFreshToday ? readCurrentIssued(dashSheet) : {};

    if (coffeePoints.length === 0) {
      Logger.log('⚠️ Нет активных кофе-поинтов');
      drawEmptyDashboard(dashSheet, [], todayStr);
      return;
    }

    if (!hostessSheet) {
      Logger.log('⚠️ Лист food_hostess не найден — рисуем пустой дашборд');
      drawEmptyDashboard(dashSheet, coffeePoints, todayStr);
      return;
    }

    const hostessData       = readTodayHostessData(hostessSheet, todayStr);
    const deliveryMap       = readTodayDelivery(ss, todayStr);
    const warehouseStockMap = warehouseSheet ? readStockFromWarehouse(warehouseSheet) : {};

    buildDashboard(
      dashSheet, coffeePoints, hostessData,
      warehouseStockMap, deliveryMap, issuedMap, todayStr
    );

    Logger.log('✅ Дашборд "Продукты" обновлён: ' + todayStr);
  } finally {
    lock.releaseLock();
  }
}

function _isDashboardForToday(dashSheet, todayStr) {
  try {
    if (dashSheet.getLastRow() < PD.DATE_ROW) return false;
    const currentDateVal = dashSheet.getRange(PD.DATE_ROW, 1).getValue();
    return currentDateVal && currentDateVal.toString().trim() === todayStr;
  } catch(e) {
    return false;
  }
}

function readTodayHostessData(sheet, todayStr) {
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const result  = {};

  for (let i = 1; i < data.length; i++) {
    const row       = data[i];
    const timestamp = row[0];
    if (!timestamp) continue;

    let rowDate;
    try {
      rowDate = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    } catch(e) { continue; }
    if (rowDate !== todayStr) continue;

    const cpName = row[1] ? row[1].toString().trim() : '';
    if (!cpName) continue;
    if (!result[cpName]) result[cpName] = {};

    for (let col = 2; col < headers.length; col++) {
      const productName = headers[col] ? headers[col].toString().trim() : '';
      const qty         = row[col];
      if (!productName || qty === '' || qty === null || qty === undefined) continue;
      const numQty = parseFloat(qty);
      if (isNaN(numQty) || numQty === 0) continue;
      result[cpName][productName] = (result[cpName][productName] || 0) + numQty;
    }
  }
  return result;
}

function readStockFromWarehouse(sheet) {
  const data   = sheet.getDataRange().getValues();
  const result = {};
  if (data.length < 2) return result;

  const headers = data[1];
  let stockColIdx = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c] ? hdrRu_(headers[c]).toLowerCase() : '';
    if (h === 'количество') { stockColIdx = c; break; }
  }
  if (stockColIdx === -1) stockColIdx = 1;

  for (let i = 2; i < data.length; i++) {
    const name  = data[i][0] ? data[i][0].toString().trim() : '';
    const stock = data[i][stockColIdx];
    if (!name) continue;
    if (stock === '' || stock === null || stock === undefined) continue;
    const num = parseFloat(stock);
    if (!isNaN(num)) result[name] = num;
  }
  return result;
}

function readTodayDelivery(ss, todayStr) {
  const result = {};

  const deliverySheet = SH_(ss, PD.DELIVERY_SHEET);
  if (!deliverySheet) return result;

  const data    = deliverySheet.getDataRange().getValues();
  const headers = data[0] || [];

  for (let i = 1; i < data.length; i++) {
    const timestamp = data[i][0];
    if (!timestamp) continue;
    let rowDate;
    try {
      rowDate = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    } catch(e) { continue; }
    if (rowDate !== todayStr) continue;

    for (let col = 1; col < headers.length; col++) {
      const productName = headers[col] ? headers[col].toString().trim() : '';
      const qty         = data[i][col];
      if (!productName || qty === '' || qty === null || qty === undefined) continue;
      const numQty = parseFloat(qty);
      if (!isNaN(numQty) && numQty > 0) {
        result[productName] = (result[productName] || 0) + numQty;
      }
    }
  }

  return result;
}

function readCurrentIssued(dashSheet) {
  const result = {};
  const data   = dashSheet.getDataRange().getValues();
  if (data.length < PD.HEADER_ROW) return result;

  const headers    = data[PD.HEADER_ROW - 1];
  let issuedColIdx = -1;
  for (let c = 0; c < headers.length; c++) {
    if (headers[c] && K(headers[c]).toUpperCase() === 'ВЫДАНО') {
      issuedColIdx = c; break;
    }
  }
  if (issuedColIdx === -1) return result;

  for (let i = PD.DATA_START_ROW - 1; i < data.length; i++) {
    const name   = data[i][0] ? data[i][0].toString().trim() : '';
    const issued = data[i][issuedColIdx];
    if (!name) continue;
    if (issued !== '' && issued !== null && issued !== undefined) {
      const num = parseFloat(issued);
      if (!isNaN(num) && num > 0) result[name] = num;
    }
  }
  return result;
}

function buildDashboard(sheet, coffeePoints, hostessData, warehouseStockMap, deliveryMap, issuedMap, todayStr) {
  const cpCount   = coffeePoints.length;
  const totalCols = 1 + cpCount + 4;

  const TOTAL_COL    = 1 + cpCount;
  const ISSUED_COL   = TOTAL_COL + 1;
  const DELIVERY_COL = ISSUED_COL + 1;
  const STOCK_COL    = DELIVERY_COL + 1;

  const orderedProducts  = collectAllProducts(hostessData);
  const deliveryProducts = Object.keys(deliveryMap);
  const allProducts      = [...new Set([...orderedProducts, ...deliveryProducts])].sort();

  drawEmptyDashboard(sheet, coffeePoints, todayStr);
  if (allProducts.length === 0) return;

  const stubRange = sheet.getRange(PD.DATA_START_ROW, 1, 1, totalCols);
  try { stubRange.breakApart(); } catch(e) {}
  stubRange
    .clearContent()
    .setFontColor('#000000')
    .setFontStyle('normal')
    .setFontWeight('normal')
    .setHorizontalAlignment('left')
    .setBackground(PD.COLOR_ROW_EVEN);

  const rowsData = [];

  allProducts.forEach(productName => {
    const rowValues = [productName];
    let totalQty    = 0;

    coffeePoints.forEach(cp => {
      const qty = (hostessData[cp.name] || {})[productName];
      if (qty !== undefined && qty !== null && qty !== '') {
        rowValues.push(qty);
        totalQty += parseFloat(qty) || 0;
      } else {
        rowValues.push('');
      }
    });

    rowValues.push(totalQty > 0 ? totalQty : '');
    rowValues.push(issuedMap[productName]   !== undefined ? issuedMap[productName]   : '');
    rowValues.push(deliveryMap[productName] !== undefined ? deliveryMap[productName] : '');

    const stockToShow = warehouseStockMap[productName] !== undefined ? warehouseStockMap[productName] : '';
    rowValues.push(stockToShow);

    rowsData.push(rowValues);
  });

  sheet.getRange(PD.DATA_START_ROW, 1, rowsData.length, totalCols).setValues(rowsData);

  sheet.getRange(PD.DATA_START_ROW, 1, rowsData.length, 1)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setFontWeight('bold');

  for (let i = 0; i < rowsData.length; i++) {
    const row   = PD.DATA_START_ROW + i;
    const rowBg = i % 2 === 0 ? PD.COLOR_ROW_EVEN : PD.COLOR_ROW_ODD;

    sheet.getRange(row, 1, 1, totalCols)
      .setBackground(rowBg)
      .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(row, 2, 1, cpCount).setHorizontalAlignment('center');

    const totalCell = sheet.getRange(row, TOTAL_COL + 1);
    totalCell.setHorizontalAlignment('center').setFontWeight('bold');
    if (rowsData[i][TOTAL_COL] !== '') totalCell.setFontColor('#1c4587');

    const issuedCell = sheet.getRange(row, ISSUED_COL + 1);
    issuedCell.setHorizontalAlignment('center');
    if (rowsData[i][ISSUED_COL] !== '') issuedCell.setFontColor('#0a7a0a').setFontWeight('bold');

    const deliveryCell = sheet.getRange(row, DELIVERY_COL + 1);
    deliveryCell.setHorizontalAlignment('center');
    if (rowsData[i][DELIVERY_COL] !== '') deliveryCell.setFontColor('#e65100').setFontWeight('bold');

    const stockCell = sheet.getRange(row, STOCK_COL + 1);
    stockCell.setHorizontalAlignment('center');
    if (rowsData[i][STOCK_COL] !== '') stockCell.setFontColor('#4a148c');

    sheet.getRange(row, 1, 1, totalCols)
      .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  formatColumns(sheet, cpCount, totalCols);
  sheet.setFrozenRows(2);

  Logger.log('✅ Дашборд построен: ' + rowsData.length + ' позиций, ' + cpCount + ' КП');
}

function formatColumns(sheet, cpCount, totalCols) {
  sheet.setColumnWidth(1, 320);
  for (let c = 2; c <= 1 + cpCount; c++) sheet.setColumnWidth(c, 75);
  sheet.setColumnWidth(2 + cpCount, 70);
  sheet.setColumnWidth(3 + cpCount, 80);
  sheet.setColumnWidth(4 + cpCount, 90);
  sheet.setColumnWidth(5 + cpCount, 130);
}

function collectAllProducts(hostessData) {
  const products = new Set();
  Object.values(hostessData).forEach(cpOrders => {
    Object.keys(cpOrders).forEach(p => products.add(p));
  });
  return [...products].sort();
}
