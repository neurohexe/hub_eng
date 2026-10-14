// SetupWarehouse.gs

function manageUnits() {
  const ui      = SpreadsheetApp.getUi();
  const current = getUnits();
  const listText = current.map(function(u, i) { return (i + 1) + '. ' + u; }).join('\n');
  const choice = ui.alert(
    L('📐 Единицы измерения'),
    L('Текущий список:\n\n') + listText + '\n\n' +
    L('YES — Добавить новую единицу\n') +
    L('NO  — Удалить единицу из списка'),
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (choice === ui.Button.CANCEL) return;

  if (choice === ui.Button.YES) {
    const res = ui.prompt(
      L('Добавить единицу'),
      L('Введите название новой единицы измерения:'),
      ui.ButtonSet.OK_CANCEL
    );
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const newUnit = res.getResponseText().trim().toLowerCase();
    if (!newUnit) { ui.alert(L('Название не может быть пустым')); return; }
    if (current.includes(newUnit)) { ui.alert(L('Единица "') + newUnit + L('" уже есть в списке')); return; }

    const saved  = PropertiesService.getScriptProperties().getProperty(PROP_UNITS);
    const custom = saved ? JSON.parse(saved) : [];
    custom.push(newUnit);
    PropertiesService.getScriptProperties().setProperty(PROP_UNITS, JSON.stringify(custom));

    ui.alert(
      L('✅ Единица "') + newUnit + L('" добавлена!\n\n') +
      L('Нажмите: Настройки → Обновить единицы на складах')
    );
  } else {
    const saved  = PropertiesService.getScriptProperties().getProperty(PROP_UNITS);
    const custom = saved ? JSON.parse(saved) : [];
    if (custom.length === 0) {
      ui.alert(L('Пользовательских единиц нет.\n\nБазовые: ') + DEFAULT_UNITS.map(L).join(', '));
      return;
    }
    const customList = custom.map(function(u, i) { return (i + 1) + '. ' + u; }).join('\n');
    const res = ui.prompt(
      L('Удалить единицу'),
      L('Пользовательские единицы:\n\n') + customList + L('\n\nВведите номер для удаления:'),
      ui.ButtonSet.OK_CANCEL
    );
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const idx = parseInt(res.getResponseText().trim()) - 1;
    if (isNaN(idx) || idx < 0 || idx >= custom.length) {
      ui.alert(L('Неверный номер'));
      return;
    }
    const removed = custom.splice(idx, 1)[0];
    PropertiesService.getScriptProperties().setProperty(PROP_UNITS, JSON.stringify(custom));
    ui.alert(L('✅ Единица "') + removed + L('" удалена.'));
  }
}

function updateUnitsOnWarehouses() {
  const ui       = SpreadsheetApp.getUi();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const unitRule = getUnitsValidationRule();
  const units    = getUnits();
  let   updated  = 0;

  Object.values(WAREHOUSE_SHEET_NAMES).forEach(sheetName => {
    const sheet = SH_(ss, sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const qty    = data[i][1];
      const unit   = data[i][2];
      const active = data[i][3];
      if (qty === '' && unit === '' && active === '') continue;
      sheet.getRange(i + 1, 3).setDataValidation(unitRule);
      updated++;
    }
  });

  ui.alert(L('✅ Единицы обновлены!\n\nСписок: ') + units.join(', ') + L('\nСтрок: ') + updated);
}

function createWarehouseStructure() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    L('Создание структуры склада'),
    L('Для какого склада создать структуру?\n\n') +
    L('1 - Продукты\n2 - Хозтовары\n3 - Канцелярия\n4 - Медикаменты'),
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const choice = response.getResponseText().trim();
  switch (choice) {
    case '1': createWarehouseSheet('PRODUCTS',  getProductCategories());  break;
    case '2': HouseholdWarehouse.createStructure();                        break;
    case '3': _createStationerySheet();                                    break;
    case '4': createWarehouseSheet('MEDICINE',  getMedicineCategories());  break;
    default:  ui.alert(L('❌ Неверный выбор'));
  }
}

function getProductCategories() {
  return getProductCategoriesList();
}

function getMedicineCategories() {
  return [
    L('Первая помощь'), L('Обезболивающие'), L('Антисептики'),
    L('Перевязочные материалы'), L('Прочее')
  ];
}

function getStationeryCategories() {
  return [L('Бумага'), L('Канцелярия'), L('Папки и файлы'), L('Блокноты'), L('Прочее')];
}

function _createStationerySheet() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const cfg = WAREHOUSE_TYPES['STATIONERY'];
  const sheetName  = L('Склад_Канцелярия');
  const categories = getStationeryCategories();

  let sheet = SH_(ss, sheetName);
  if (sheet) {
    const resp = ui.alert(
      L('Пересоздать склад канцелярии?'),
      L('Лист уже существует. Данные будут удалены.'),
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(L(sheetName));
  const today     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const suppliers = getSuppliersFromBaseInfo();

  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 100);
  sheet.setColumnWidth(9, 50);

  sheet.getRange(1, 1, 1, 9).merge();
  sheet.getRange('A1')
    .setValue(L('Данные канцелярия склада на ') + today)
    .setBackground(cfg.color).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  sheet.getRange(2, 1, 1, 9).setValues([[
    L('Название продукта'), L('Количество'), L('Единица измерения'),
    L('Активность'), L('Артикул'), L('Поставщик'), L('Хозяюшки'), L('Выдаётся'), ''
  ]])
    .setBackground(cfg.color).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(2, 35);

  const unitRule = getUnitsValidationRule();
  let supplierRule = null;
  if (suppliers.length > 0) {
    supplierRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(suppliers, true).setAllowInvalid(true).build();
  }

  let row = 3;
  categories.forEach(category => {
    sheet.getRange(row, 1, 1, 8).merge();
    sheet.getRange(row, 1)
      .setValue(category)
      .setBackground('#d9d9d9').setFontWeight('bold')
      .setFontSize(12).setVerticalAlignment('middle');
    sheet.getRange(row, 9)
      .setValue('➕').setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#4CAF50').setFontColor('#FFFFFF');
    sheet.setRowHeight(row, 30);
    row++;

    sheet.getRange(row, 2).setValue(0);
    sheet.getRange(row, 3).setValue(L('штука')).setDataValidation(unitRule);
    sheet.getRange(row, 4).insertCheckboxes().check();
    sheet.getRange(row, 7).insertCheckboxes().check();
    sheet.getRange(row, 8).insertCheckboxes().uncheck();
    if (supplierRule) sheet.getRange(row, 6).setDataValidation(supplierRule);

    sheet.getRange(row, 1, 1, 8)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle')
      .setFontColor('#000000')
      .setBackground('#FFFFFF').setWrap(true);
    sheet.setRowHeight(row, 25);
    row += 2;
  });

  sheet.setFrozenRows(2);
  SpreadsheetApp.flush(); 
  updateAllWarehouseDates();
  ui.alert(L('✅ Склад канцелярии создан!'));
  sheet.activate();
}

function createWarehouseSheet(warehouseType, categories) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const config = WAREHOUSE_TYPES[warehouseType];
  const sheetName = WAREHOUSE_SHEET_NAMES[warehouseType];

  let sheet = SH_(ss, sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(L(sheetName));
  } else {
    sheet.clear();
  }

  const today     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const suppliers = getSuppliersFromBaseInfo();

  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 50);

  sheet.getRange(1, 1, 1, 7).merge();
  sheet.getRange('A1')
    .setValue(L('Данные ') + config.name.toLowerCase() + L(' склада на ') + today)
    .setBackground('#1c4587').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  sheet.getRange(2, 1, 1, 7).setValues([[
    L('Название продукта'), L('Количество'), L('Единица измерения'),
    L('Активность'), L('Артикул'), L('Поставщик'), ''
  ]])
    .setBackground('#1c4587').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(2, 35);

  const unitRule = getUnitsValidationRule();
  let supplierRule = null;
  if (suppliers.length > 0) {
    supplierRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(suppliers, true)
      .setAllowInvalid(true)
      .build();
  }

  let currentRow = 3;
  categories.forEach(category => {
    sheet.getRange(currentRow, 1, 1, 6).merge();
    sheet.getRange(currentRow, 1)
      .setValue(category)
      .setBackground('#d9d9d9').setFontWeight('bold')
      .setFontSize(12).setVerticalAlignment('middle');
    sheet.getRange(currentRow, 7)
      .setValue('➕').setFontSize(16)
      .setHorizontalAlignment('center')
      .setBackground('#4CAF50').setFontColor('#FFFFFF');
    sheet.setRowHeight(currentRow, 30);
    currentRow++;

    sheet.getRange(currentRow, 1).setValue('');
    sheet.getRange(currentRow, 2).setValue(0);
    sheet.getRange(currentRow, 3).setValue(L('штука')).setDataValidation(unitRule);
    sheet.getRange(currentRow, 4).insertCheckboxes().check();
    sheet.getRange(currentRow, 5).setValue('');
    if (supplierRule) sheet.getRange(currentRow, 6).setDataValidation(supplierRule);

 
    sheet.getRange(currentRow, 1, 1, 6)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
      .setVerticalAlignment('middle')
      .setFontColor('#000000')
      .setBackground('#FFFFFF');
    sheet.setRowHeight(currentRow, 25);
    currentRow++;
    currentRow++;
  });

  sheet.setFrozenRows(2);
  SpreadsheetApp.flush(); 
  updateAllWarehouseDates();

  const supplierNote = suppliers.length > 0
    ? L('\n✅ Загружено поставщиков: ') + suppliers.length
    : L('\n⚠️ Поставщики не найдены — добавьте их в "Базовую информацию"');

  SpreadsheetApp.getUi().alert(
    L('✅ Структура склада "') + config.name + L('" создана!') +
    supplierNote
  );
  sheet.activate();
}

function getSuppliersFromBaseInfo() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const baseSheet = SH_(ss, L('Базовая информация'));
    if (!baseSheet) return [];

    const data       = baseSheet.getDataRange().getValues();
    const suppliers  = [];
    let inSection    = false;
    let headerPassed = false;

    for (let i = 0; i < data.length; i++) {
      const cell = data[i][0] ? data[i][0].toString().trim() : '';

      if (K(cell).toUpperCase() === 'ПОСТАВЩИКИ') {
        inSection    = true;
        headerPassed = false;
        continue;
      }
      if (!inSection) continue;

      if (!headerPassed && K(cell) === 'Название поставщика') {
        headerPassed = true;
        continue;
      }

      const KNOWN_SECTIONS_EXACT = [
        'ИДЕНТИФИКАЦИЯ ОФИСА', 'КОФЕ-ПОИНТЫ',
        'ПАРАМЕТРЫ РАБОТЫ ОФИСА', 'ПРАВИЛА СКЛАДА', 'AI ПРОФИЛЬ ОФИСА',
        'LIVE АНАЛИТИКА (АВТОМАТИЧЕСКИ)'
      ];
      const cellUpper = K(cell).toUpperCase();
      if (headerPassed &&
          (cellUpper.startsWith('TELEGRAM') || KNOWN_SECTIONS_EXACT.includes(cellUpper))) {
        break;
      }

      if (headerPassed && cell) {
        const activeVal = data[i][4];
        const isActive  = activeVal === true || activeVal === '' ||
                          activeVal === null || activeVal === undefined;
        if (isActive) suppliers.push(cell);
      }
    }

    Logger.log('✅ Загружено поставщиков: ' + suppliers.length);
    return suppliers;

  } catch (e) {
    Logger.log('⚠️ getSuppliersFromBaseInfo: ' + e.toString());
    return [];
  }
}

function getSupplierChatMap() {
  const map = {};
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const baseSheet = SH_(ss, L('Базовая информация'));
    if (!baseSheet) return map;

    const data       = baseSheet.getDataRange().getValues();
    let inSection    = false;
    let headerPassed = false;

    const KNOWN_SECTIONS_EXACT = [
      'ИДЕНТИФИКАЦИЯ ОФИСА', 'КОФЕ-ПОИНТЫ',
      'ПАРАМЕТРЫ РАБОТЫ ОФИСА', 'ПРАВИЛА СКЛАДА', 'AI ПРОФИЛЬ ОФИСА',
      'LIVE АНАЛИТИКА (АВТОМАТИЧЕСКИ)'
    ];

    for (let i = 0; i < data.length; i++) {
      const cell = data[i][0] ? data[i][0].toString().trim() : '';

      if (K(cell).toUpperCase() === 'ПОСТАВЩИКИ') {
        inSection    = true;
        headerPassed = false;
        continue;
      }
      if (!inSection) continue;

      if (!headerPassed && K(cell) === 'Название поставщика') {
        headerPassed = true;
        continue;
      }

      const cellUpper = K(cell).toUpperCase();
      if (headerPassed &&
          (cellUpper.startsWith('TELEGRAM') || KNOWN_SECTIONS_EXACT.includes(cellUpper))) {
        break;
      }

      if (headerPassed && cell) {
        const chatId = data[i][3] ? data[i][3].toString().trim() : '';
        if (chatId) map[cell] = chatId;
      }
    }
  } catch(e) {
    Logger.log('⚠️ getSupplierChatMap: ' + e.toString());
  }
  return map;
}

function updateSupplierDropdowns() {
  const ui        = SpreadsheetApp.getUi();
  const suppliers = getSuppliersFromBaseInfo();

  if (suppliers.length === 0) {
    ui.alert(
      L('❌ Поставщики не найдены в "Базовой информации".\n\n') +
      L('Добавьте раздел ПОСТАВЩИКИ в базовую информацию.')
    );
    return;
  }

  const updatedCount = _applySupplierDropdowns(suppliers);

  ui.alert(
    L('✅ Список поставщиков обновлён!\n\n') +
    L('Поставщиков: ') + suppliers.length + '\n' +
    L('Обновлено строк: ') + updatedCount
  );
}


function _updateSupplierDropdownsSilent() {
  try {
    const suppliers = getSuppliersFromBaseInfo();
    if (suppliers.length === 0) return 0;
    return _applySupplierDropdowns(suppliers);
  } catch(e) {
    Logger.log('⚠️ _updateSupplierDropdownsSilent: ' + e.toString());
    return 0;
  }
}

function _applySupplierDropdowns(suppliers) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  let   updatedCount = 0;
  const supplierRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(suppliers, true)
    .setAllowInvalid(true)
    .build();

  Object.values(WAREHOUSE_SHEET_NAMES).forEach(name => {
    const sheet = SH_(ss, name);
    if (!sheet) return;

    const data    = sheet.getDataRange().getValues();
    const headers = data[1] || [];
    let supplierColIdx = -1;
    for (let c = 0; c < headers.length; c++) {
      if (headers[c] && hdrRu_(headers[c]) === 'Поставщик') {
        supplierColIdx = c;
        break;
      }
    }
    if (supplierColIdx === -1) return;

    for (let i = 2; i < data.length; i++) {
      const qty    = data[i][1];
      const unit   = data[i][2];
      const active = data[i][3];
      if (qty === '' && unit === '' && active === '') continue;
      sheet.getRange(i + 1, supplierColIdx + 1).setDataValidation(supplierRule);
      updatedCount++;
    }
  });

  return updatedCount;
}
