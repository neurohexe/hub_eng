// DailyArchive.gs

function archiveAllWarehouses() {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  Logger.log('=== АРХИВАЦИЯ ' + today + ' ===');
  Object.keys(WAREHOUSE_CONFIG).forEach(function(warehouseType) {
    try {
      _archiveWarehouse(warehouseType, today);
      Logger.log('✅ ' + warehouseType + ' архивирован');
    } catch(e) {
      Logger.log('❌ ' + warehouseType + ': ' + e.toString());
    }
  });
}

function _archiveWarehouse(warehouseType, dateStr) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = getWarehouseConfigByType(warehouseType);

  var warehouseSheet = SH_(ss, cfg.warehouseSheet);
  if (!warehouseSheet) throw new Error(L('Лист "') + L(cfg.warehouseSheet) + L('" не найден'));

  var coffeePoints = BaseInfo.getActiveCoffeePoints();
  var headers      = _buildHeaders(cfg, coffeePoints);

  _ensureHeaders(cfg, headers);

  var todayData = _collectDayData(cfg, warehouseSheet, coffeePoints, dateStr);
  if (!todayData || todayData.length === 0) {
    Logger.log('⚠️ ' + warehouseType + ': нет движений за ' + dateStr);
    return;
  }

  _appendRows(cfg.dbSheet, todayData, dateStr, headers.length, headers);
  Logger.log('✅ ' + warehouseType + ': записано ' + todayData.length + ' строк');
}

function _collectDayData(cfg, warehouseSheet, coffeePoints, dateStr) {
  var products = _readWarehouseProducts(warehouseSheet, cfg.excludeCategories);
  if (products.length === 0) return [];

  var hostessData  = _readDayOrders(cfg.hostessSheet, dateStr, true);
  var cleaningData = cfg.hasCleaning ? _readDayOrders(cfg.cleaningSheet, dateStr, false) : {};
  var deliveryData = _readDayDelivery(cfg.logistSheet, dateStr);
  var issuedData   = _readIssuedFromDash(cfg.dashSheet);

  var hasMovement = Object.keys(hostessData).length > 0 ||
                    Object.keys(cleaningData).length > 0 ||
                    Object.keys(deliveryData).length > 0 ||
                    Object.keys(issuedData).length > 0;
  if (!hasMovement) return [];

  return products.map(function(product) {
    var row = { date: dateStr, name: product.name, category: product.category, stock: product.qty };
    var totalOrders = 0;

    coffeePoints.forEach(function(cp) {
      var qty = (hostessData[cp.name] || {})[product.name] || 0;
      row['cp_' + cp.id] = qty > 0 ? qty : '';
      totalOrders += qty;
    });

    if (cfg.hasCleaning) {
      var cleanQty = 0;
      Object.values(cleaningData).forEach(function(d) { cleanQty += parseFloat(d[product.name]) || 0; });
      row.cleaning = cleanQty > 0 ? cleanQty : '';
      totalOrders += cleanQty;
    }

    var issued   = parseFloat(issuedData[product.name])   || 0;
    var delivery = parseFloat(deliveryData[product.name]) || 0;

    row.totalOrders = totalOrders > 0 ? totalOrders : '';
    row.issued      = issued   > 0 ? issued   : '';
    row.delivery    = delivery > 0 ? delivery : '';
    return row;
  });
}

function _readWarehouseProducts(warehouseSheet, excludeCategories) {
  var data = warehouseSheet.getDataRange().getValues();
  var products = [], currentCategory = '', skipCategory = false;

  for (var i = 2; i < data.length; i++) {
    var name  = data[i][0] ? data[i][0].toString().trim() : '';
    var qty   = data[i][1], unit = data[i][2], active = data[i][3];
    var isCat = (qty === '' || qty === null) && (unit === '' || unit === null) &&
                (active === '' || active === false || active === null);
    if (isCat && name) {
      currentCategory = name;
      skipCategory = excludeCategories.indexOf(K(name)) !== -1;
      continue;
    }
    if (skipCategory || !name || active !== true) continue;
    products.push({ name: name, category: currentCategory, qty: parseFloat(qty) || 0 });
  }
  return products;
}

function _readDayOrders(sheetName, dateStr, isHostess) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {}, sheet = SH_(ss, sheetName);
  if (!sheet) return result;

  var data = sheet.getDataRange().getValues(), headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var ts = data[i][0]; if (!ts) continue;
    var rowDate;
    try { rowDate = Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy'); } catch(e) { continue; }
    if (rowDate !== dateStr) continue;

    var groupName = data[i][1] ? data[i][1].toString().trim() : (isHostess ? L('КП') : L('Группа'));
    if (!result[groupName]) result[groupName] = {};

    for (var col = 2; col < headers.length; col++) {
      var product = headers[col] ? headers[col].toString().trim() : '';
      var qty     = data[i][col];
      if (!product || qty === '' || qty === null) continue;
      var num = parseFloat(qty);
      if (!isNaN(num) && num > 0) result[groupName][product] = (result[groupName][product] || 0) + num;
    }
  }
  return result;
}

function _readDayDelivery(sheetName, dateStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {}, sheet = SH_(ss, sheetName);
  if (!sheet) return result;

  var data = sheet.getDataRange().getValues(), headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var ts = data[i][0]; if (!ts) continue;
    var rowDate;
    try { rowDate = Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy'); } catch(e) { continue; }
    if (rowDate !== dateStr) continue;

    for (var col = 1; col < headers.length; col++) {
      var product = headers[col] ? headers[col].toString().trim() : '';
      var qty     = data[i][col];
      if (!product || qty === '' || qty === null) continue;
      var num = parseFloat(qty);
      if (!isNaN(num) && num > 0) result[product] = (result[product] || 0) + num;
    }
  }
  return result;
}

function _readIssuedFromDash(dashSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {}, sheet = SH_(ss, dashSheetName);
  if (!sheet) return result;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  var headers = data[1], idx = -1;
  for (var c = 0; c < headers.length; c++) {
    if (headers[c] && K(headers[c]).toUpperCase() === 'ВЫДАНО') { idx = c; break; }
  }
  if (idx === -1) return result;

  for (var i = 2; i < data.length; i++) {
    var name = data[i][0] ? data[i][0].toString().trim() : '';
    var val  = data[i][idx];
    if (!name || val === '' || val === null) continue;
    var num = parseFloat(val);
    if (!isNaN(num) && num > 0) result[name] = num;
  }
  return result;
}

function _ensureHeaders(cfg, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = SH_(ss, cfg.dbSheet);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(L(cfg.dbSheet));
    Logger.log('✓ Создан "' + cfg.dbSheet + '"');
  }
  if (dbSheet.getLastRow() > 0) return;
  dbSheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(cfg.color).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  dbSheet.setFrozenRows(1);
  dbSheet.setColumnWidth(1, 110); dbSheet.setColumnWidth(2, 240); dbSheet.setColumnWidth(3, 160);
  for (var c = 4; c <= headers.length; c++) dbSheet.setColumnWidth(c, 85);
}

function _buildHeaders(cfg, coffeePoints) {
  var headers = [L('Дата'), L('Товар'), L('Категория')];
  coffeePoints.forEach(function(cp) { headers.push(cp.name); });
  if (cfg.hasCleaning) headers.push(L('🧹 Клининг'));
  headers.push(L('ИТОГО заказов'), L('Выдано'), L('Поставка'), L('Остаток конец дня'));
  return headers;
}

function _appendRows(dbSheetName, todayData, dateStr, totalCols, headers) {
  if (todayData.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = SH_(ss, dbSheetName);
  if (!dbSheet) {
    Logger.log('❌ Лист "' + dbSheetName + '" не найден');
    return;
  }

  var startRow = dbSheet.getLastRow() + 1;
  
  dbSheet.getRange(startRow, 1)
    .setValue('📅 ' + dateStr + L('  (товаров: ') + todayData.length + ')')
    .setBackground('#e8eaf6').setFontWeight('bold').setFontSize(11).setFontColor('#1a237e')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  
  if (totalCols > 1) {
    dbSheet.getRange(startRow, 2, 1, totalCols - 1).setBackground('#e8eaf6');
  }
  dbSheet.setRowHeight(startRow, 28);

  var colHeaderRow = startRow + 1;
  dbSheet.getRange(colHeaderRow, 1, 1, totalCols).setValues([headers])
    .setBackground('#c5cae9').setFontWeight('bold').setFontColor('#1a237e').setHorizontalAlignment('center');
  dbSheet.setRowHeight(colHeaderRow, 24);

  var dataFirstRow = colHeaderRow + 1;
  var rowsToWrite  = [];

  for (var i = 0; i < todayData.length; i++) {
    var row   = todayData[i];
    var cells = [row.date, row.name, row.category];

    Object.keys(row).forEach(function(key) {
      if (key.startsWith('cp_')) cells.push(row[key] === '' ? '' : row[key]);
    });
    if (row.cleaning !== undefined) cells.push(row.cleaning === '' ? '' : row.cleaning);
    cells.push(
      row.totalOrders === '' ? '' : row.totalOrders,
      row.issued      === '' ? '' : row.issued,
      row.delivery    === '' ? '' : row.delivery,
      row.stock !== undefined ? row.stock : 0
    );
    rowsToWrite.push(cells);
  }

  dbSheet.getRange(dataFirstRow, 1, rowsToWrite.length, totalCols).setValues(rowsToWrite);

  for (var j = 0; j < rowsToWrite.length; j++) {
    var r  = dataFirstRow + j;
    var bg = (j % 2 === 0) ? '#ffffff' : '#f3f4f6';
    dbSheet.getRange(r, 1, 1, totalCols)
      .setBackground(bg)
      .setBorder(null, null, null, null, null, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
    dbSheet.setRowHeight(r, 22);
  }

  var lastRow = dataFirstRow + rowsToWrite.length - 1;
  dbSheet.getRange(lastRow, 1, 1, totalCols)
    .setBorder(null, null, true, null, null, null, '#9fa8da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  
  try {
    dbSheet.getRange(colHeaderRow, 1, rowsToWrite.length + 1, 1).shiftRowGroupDepth(1);
    Logger.log('✓ Группировка OK');
  } catch(e) {
    Logger.log('⚠️ Группировка не удалась: ' + e.toString());
  }

  Logger.log('✓ Записано ' + rowsToWrite.length + ' строк за ' + dateStr);
}

function endOfDayAll() {
  Logger.log('=== КОНЕЦ ДНЯ ===');
  try { archiveAllWarehouses(); } catch(e) { Logger.log('❌ archiveAllWarehouses: ' + e.toString()); }
  Logger.log('✅ Конец дня завершён');
}

function resetAllDashboards() {
  Logger.log('=== СБРОС ДАШБОРДОВ ===');
  try { refreshProductsDashboard();       Logger.log('✅ Продукты');    } catch(e) { Logger.log('❌ ' + e); }
  try { HouseholdDashboard.refresh();     Logger.log('✅ Хозтовары');   } catch(e) { Logger.log('❌ ' + e); }
  try { swRefreshDashboard('STATIONERY'); Logger.log('✅ Канцелярия');  } catch(e) { Logger.log('❌ ' + e); }
  try { swRefreshDashboard('MEDICINE');   Logger.log('✅ Медикаменты'); } catch(e) { Logger.log('❌ ' + e); }
}


function archiveAllWarehousesManual() {
  var ui    = SpreadsheetApp.getUi();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  if (ui.alert(L('💾 Архивация'), L('Записать данные за ') + today + '?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  archiveAllWarehouses();
  ui.alert(L('✅ Архивация завершена!'));
}

function endOfDayManual() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(L('🌙 Конец дня'), L('Запустить архивацию и сброс дашбордов?'), ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  endOfDayAll();
  resetAllDashboards();
  ui.alert(L('✅ Конец дня выполнен!'));
}


function findArchiveError() {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var cfg   = getWarehouseConfigByType('PRODUCTS');

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('1. Таблица: ' + ss.getName());
    var warehouseSheet = SH_(ss, cfg.warehouseSheet);
    Logger.log('2. warehouseSheet: ' + (warehouseSheet ? 'OK' : 'НЕ НАЙДЕН'));
    var coffeePoints = BaseInfo.getActiveCoffeePoints();
    Logger.log('3. coffeePoints: ' + coffeePoints.length);
    var headers = _buildHeaders(cfg, coffeePoints);
    Logger.log('4. headers: ' + headers.length);
    _ensureHeaders(cfg, headers);
    Logger.log('5. _ensureHeaders: OK');
    var products = _readWarehouseProducts(warehouseSheet, cfg.excludeCategories);
    Logger.log('6. products: ' + products.length);
    var todayData = _collectDayData(cfg, warehouseSheet, coffeePoints, today);
    Logger.log('7. todayData: ' + todayData.length + ' строк');
    if (todayData.length > 0) {
      _appendRows(cfg.dbSheet, todayData, today, headers.length, headers);
      Logger.log('8. _appendRows: OK');
    }
    Logger.log('✅ Всё прошло успешно!');
  } catch(e) {
    Logger.log('❌ УПАЛО: ' + e.toString());
    Logger.log('❌ Stack: ' + e.stack);
  }
}
