const ORDER_SHEET_NAME = 'Заказ продуктов';

const ORDER_SHEET_NAMES = {
  PRODUCTS:   'Заказ продуктов',
  HOUSEHOLD:  'Заказ хозтоваров',
  STATIONERY: 'Заказ канцелярии',
  MEDICINE:   'Заказ медикаментов'
};

function getOrCreateOrderSheet(ss) {
  let sheet = SH_(ss, ORDER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(L(ORDER_SHEET_NAME));
    Logger.log('✓ Создан лист "' + ORDER_SHEET_NAME + '"');
  }
  return sheet;
}

function _getLatestMenuSheet(ss) {
  const todayMenu = MenuManager.findTodaySheet();
  if (todayMenu) return todayMenu.sheet;
  return SH_(ss, L('Меню'));
}

function updateOrderSheet() {
  try {
    const ss             = SpreadsheetApp.getActiveSpreadsheet();
    const warehouseSheet = SH_(ss, L('Склад_Продукты'));

    if (!warehouseSheet) {
      SpreadsheetApp.getUi().alert(
        L('Лист "Склад_Продукты" не найден.\n\n') +
        L('Сначала создайте структуру склада:\n') +
        L('Настройки → Создать структуру склада')
      );
      return;
    }

    const orderSheet = getOrCreateOrderSheet(ss);
    orderSheet.clear();

    const menuSheet          = _getLatestMenuSheet(ss);
    const config             = loadMenuConfig();
    const rotatingCategories = (config.rotatingCategories || []).map(rc => rc.category);

    const weekPeriod = menuSheet ? getWeekPeriodFromMenu(menuSheet) : '';
    const today      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

    orderSheet.getRange(1, 1, 1, 5).merge();
    orderSheet.getRange('A1')
      .setValue(L('Заказ продуктов ') + today + (weekPeriod ? ' (' + weekPeriod + ')' : ''))
      .setBackground('#1c4587').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center');
    orderSheet.setRowHeight(1, 35);

    orderSheet.getRange(2, 1, 1, 5)
      .setValues([[L('Позиция'), L('Количество'), L('Единица измерения'), L('Артикул'), L('Поставщик')]])
      .setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center');
    orderSheet.setRowHeight(2, 28);
    orderSheet.setFrozenRows(2);

    const unitMap = getUnitMapFromWarehouse(warehouseSheet);
    let   row     = 3;

    addSectionHeader(orderSheet, row, L('📦 Стандартный ассортимент'));
    row++;
    const standardProducts = getStandardProducts(warehouseSheet, unitMap, rotatingCategories);
    standardProducts.forEach(p => {
      writeOrderRow(orderSheet, row, p.name, p.unit, p.articul, p.supplier);
      row++;
    });

    const selectedAssortment = getSelectedAssortmentFromMenu(menuSheet, warehouseSheet);
    if (Object.keys(selectedAssortment).length > 0) {
      row++;
      addSectionHeader(orderSheet, row, L('🗓 По выбору меню (сменные категории)'));
      row++;
      Object.values(selectedAssortment).forEach(product => {
        if (!product || isMenuFiller_(product)) return;
        const unit  = unitMap[product] || '';
        const extra = getArticulAndSupplier(warehouseSheet, product);
        writeOrderRow(orderSheet, row, product, unit, extra.articul, extra.supplier);
        row++;
      });
    }

    orderSheet.setColumnWidth(1, 300);
    orderSheet.setColumnWidth(2, 100);
    orderSheet.setColumnWidth(3, 150);
    orderSheet.setColumnWidth(4, 120);
    orderSheet.setColumnWidth(5, 180);
    if (row > 3) orderSheet.getRange(3, 1, row - 3, 1).setWrap(true);
    formatOrderRows(orderSheet, 3, row - 1);
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(
      L('Лист "') + L(ORDER_SHEET_NAME) + L('" обновлён!\n\n') +
      L('Стандарт: ') + standardProducts.length + L(' позиций\n') +
      L('По выбору: ') + Object.keys(selectedAssortment).length + L(' позиций\n\n') +
      L('Заполните количество и нажмите:\n') +
      L('Заказы → Продукты → Отправить заказ (Telegram)')
    );

  } catch(error) {
    Logger.log('❌ updateOrderSheet: ' + error.toString());
    SpreadsheetApp.getUi().alert(L('Ошибка: ') + error.toString());
  }
}

function sendOrderToTelegramAndUpdateForm() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = SH_(ss, ORDER_SHEET_NAME);

    if (!orderSheet) {
      SpreadsheetApp.getUi().alert(
        L('Лист "') + L(ORDER_SHEET_NAME) + L('" не найден.\n\n') +
        L('Сначала сформируйте заказ:\n') +
        L('Заказы → Продукты → Обновить лист заказа')
      );
      return;
    }

    const data       = orderSheet.getDataRange().getValues();
    const orderItems = [];

    for (let i = 2; i < data.length; i++) {
      const name     = data[i][0] ? data[i][0].toString().trim() : '';
      const qty      = data[i][1];
      const unit     = data[i][2] ? data[i][2].toString().trim() : '';
      const articul  = data[i][3] ? data[i][3].toString().trim() : '';
      const supplier = data[i][4] ? data[i][4].toString().trim() : '';
      if (!name || name.startsWith('📦') || name.startsWith('🗓')) continue;
      if (!qty || qty === '' || parseFloat(qty) === 0) continue;
      orderItems.push({ name, qty, unit, articul, supplier });
    }

    if (orderItems.length === 0) {
      SpreadsheetApp.getUi().alert(L('Нет позиций с заполненным количеством. Сначала укажите количество.'));
      return;
    }

    const headerValue = data[0][0] ? data[0][0].toString() : '';
    const periodMatch = headerValue.match(/\(([^)]+)\)/);
    const weekPeriod  = periodMatch ? periodMatch[1] : '';

    let officeName = L('Офис');
    try {
      const config = BaseInfo.getOfficeConfig();
      if (config && config.officeName) officeName = config.officeName;
    } catch(e) { Logger.log('⚠️ getOfficeConfig: ' + e.toString()); }

    let telegramOk = false;
    try {
      TelegramManager.sendOrder(orderItems, weekPeriod, officeName);
      telegramOk = true;
    } catch(telegramError) {
      Logger.log('⚠️ Ошибка Telegram: ' + telegramError.toString());
    }

    let formResult = null;
    try {
      formResult = updateProductsLogistForm(orderItems, weekPeriod);
      Logger.log('✅ Форма логиста продуктов обновлена');
    } catch(formError) {
      Logger.log('⚠️ Форма логиста не обновилась: ' + formError.toString());
    }

    SpreadsheetApp.getUi().alert(
      (telegramOk ? L('✅ Заказ отправлен в Telegram!\n\n') : L('⚠️ Telegram не отправлен.\n\n')) +
      L('Форма логиста: ') + (formResult ? formResult.short : L('не обновилась')) + '\n\n' +
      L('Форму хозяюшек обновите вручную:\n') +
      L('Заказы → Продукты → Обновить форму хозяюшек')
    );

  } catch(error) {
    Logger.log('❌ sendOrderToTelegramAndUpdateForm: ' + error.toString());
    SpreadsheetApp.getUi().alert(L('Ошибка:\n\n') + error.toString());
  }
}

function updateOrderSheetForWarehouse(warehouseType) {
  try {
    const ss             = SpreadsheetApp.getActiveSpreadsheet();
    const warehouseSheet = getWarehouseSheetByType(warehouseType);

    if (!warehouseSheet) {
      SpreadsheetApp.getUi().alert(
        L('Склад ') + warehouseType + L(' не найден.\n\n') +
        L('Сначала создайте структуру склада:\n') +
        L('Настройки → Создать структуру склада')
      );
      return;
    }

    const orderSheetName = ORDER_SHEET_NAMES[warehouseType] || (L('Заказ_') + warehouseType);
    let   orderSheet     = SH_(ss, orderSheetName);
    if (orderSheet) {
      orderSheet.clear();
    } else {
      orderSheet = ss.insertSheet(L(orderSheetName));
      Logger.log('✓ Создан лист "' + orderSheetName + '"');
    }

    const unitMap       = getUnitMapFromWarehouse(warehouseSheet);
    const today         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const warehouseName = getWarehouseDisplayName(warehouseType);

    orderSheet.getRange(1, 1, 1, 5).merge();
    orderSheet.getRange('A1')
      .setValue(L('Заказ ') + warehouseName + ' ' + today)
      .setBackground('#1c4587').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
    orderSheet.setRowHeight(1, 35);

    orderSheet.getRange(2, 1, 1, 5)
      .setValues([[L('Позиция'), L('Количество'), L('Единица измерения'), L('Артикул'), L('Поставщик')]])
      .setBackground('#4a86e8').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
    orderSheet.setFrozenRows(2);

    const products = collectAllActiveProducts(warehouseSheet, unitMap);
    let   row      = 3;
    products.forEach(p => {
      writeOrderRow(orderSheet, row, p.name, p.unit, p.articul, p.supplier);
      row++;
    });

    orderSheet.setColumnWidth(1, 300);
    orderSheet.setColumnWidth(2, 100);
    orderSheet.setColumnWidth(3, 150);
    orderSheet.setColumnWidth(4, 120);
    orderSheet.setColumnWidth(5, 180);
    if (row > 3) orderSheet.getRange(3, 1, row - 3, 1).setWrap(true);
    formatOrderRows(orderSheet, 3, row - 1);
    SpreadsheetApp.flush();

    SpreadsheetApp.getUi().alert(L('Лист "') + L(orderSheetName) + L('" обновлён (') + products.length + L(' позиций).'));

  } catch(error) {
    Logger.log('❌ updateOrderSheetForWarehouse: ' + error.toString());
    SpreadsheetApp.getUi().alert(L('Ошибка: ') + error.toString());
  }
}

function sendOrderToTelegramForWarehouse(warehouseType) {
  try {
    const orderSheetName = ORDER_SHEET_NAMES[warehouseType] || (L('Заказ_') + warehouseType);
    const orderSheet     = SH_(SpreadsheetApp.getActiveSpreadsheet(), orderSheetName);

    if (!orderSheet) {
      SpreadsheetApp.getUi().alert(
        L('Лист "') + L(orderSheetName) + L('" не найден.\n\n') +
        L('Сначала создайте заказ:\n') +
        L('Заказы → Обновить лист заказа')
      );
      return;
    }

    const data       = orderSheet.getDataRange().getValues();
    const orderItems = [];

    for (let i = 2; i < data.length; i++) {
      const name     = data[i][0] ? data[i][0].toString().trim() : '';
      const qty      = data[i][1];
      const unit     = data[i][2] ? data[i][2].toString().trim() : '';
      const articul  = data[i][3] ? data[i][3].toString().trim() : '';
      const supplier = data[i][4] ? data[i][4].toString().trim() : '';
      if (!name || isSectionHeader(name)) continue;
      if (!qty || qty === '' || parseFloat(qty) === 0) continue;
      orderItems.push({ name, qty, unit, articul, supplier });
    }

    if (orderItems.length === 0) {
      SpreadsheetApp.getUi().alert(L('Нет позиций с заполненным количеством.'));
      return;
    }

    const warehouseName = getWarehouseDisplayName(warehouseType);
    const today         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

    const grouped = {};
    orderItems.forEach(item => {
      const sup = (item.supplier || '').trim() || L('Без поставщика');
      if (!grouped[sup]) grouped[sup] = [];
      grouped[sup].push(item);
    });

    let telegramOk = false;
    try {
      const defaultChatId   = TelegramManager.getChatId();
      const defaultThreadId = TelegramManager.getDefaultThreadId();
      const supplierChats   = getSupplierChatMap();
      let anyOk = false;

      Object.entries(grouped).forEach(([supplier, items]) => {
        let message = L('<b>ЗАКАЗ ') + warehouseName.toUpperCase() + '</b>\n' + today + '\n\n';
        message += '<b>' + supplier + '</b>\n';
        items.forEach((item, idx) => {
          const art = item.articul ? ' [' + item.articul + ']' : '';
          message += (idx + 1) + '. ' + item.name + art + ' — ' + item.qty + ' ' + (item.unit || L('шт')) + '\n';
        });

        const hasOwnChat = !!supplierChats[supplier];
        const chatId      = hasOwnChat ? supplierChats[supplier] : defaultChatId;
        const threadId    = hasOwnChat ? null : defaultThreadId;
        if (!hasOwnChat) message += L('\n📋 <i>Нет своего Telegram — скопируйте и перешлите вручную</i>');

        const result = TelegramManager.sendMessage(chatId, message, threadId);
        if (result.ok) anyOk = true;
        else Logger.log('⚠️ Telegram (' + supplier + '): ' + result.description);
      });

      telegramOk = anyOk;
      if (anyOk) Logger.log('✅ Заказ ' + warehouseName + ' отправлен по поставщикам в Telegram');
    } catch(telegramError) {
      Logger.log('⚠️ Telegram: ' + telegramError.toString());
    }

    let formResult = null;
    try {
      if (warehouseType === 'HOUSEHOLD') {
        formResult = updateHouseholdLogistForm(orderItems, today);
      } else if (warehouseType === 'STATIONERY') {
        formResult = updateStationeryLogistForm(orderItems, today);
      } else if (warehouseType === 'MEDICINE') {
        formResult = updateMedicineLogistForm(orderItems, today);
      }
      Logger.log('✅ Форма логиста ' + warehouseType + ' обновлена');
    } catch(formError) {
      Logger.log('⚠️ Форма логиста не обновилась: ' + formError.toString());
    }

    const hostessHint = warehouseType === 'HOUSEHOLD'
      ? L('Заказы → Хозтовары → Обновить форму хозяюшек\nЗаказы → Хозтовары → Обновить форму клининга')
      : warehouseType === 'STATIONERY'
      ? L('Заказы → Канцелярия → Обновить форму хозяюшек')
      : L('Заказы → Медикаменты → Обновить форму хозяюшек');

    SpreadsheetApp.getUi().alert(
      (telegramOk ? L('✅ Заказ ') + warehouseName + L(' отправлен!\n\n') : L('⚠️ Telegram не отправлен.\n\n')) +
      L('Форма логиста: ') + (formResult ? formResult.short : L('не обновилась')) + '\n\n' +
      L('Формы хозяюшек обновите вручную:\n') + hostessHint
    );

  } catch(error) {
    Logger.log('❌ sendOrderToTelegramForWarehouse: ' + error.toString());
    SpreadsheetApp.getUi().alert(L('Ошибка: ') + error.toString());
  }
}

function getStandardProducts(warehouseSheet, unitMap, rotatingCategories) {
  const allCategories     = getAllWarehouseCategories(warehouseSheet);
  const excludeCategories = ['Алкоголь', ...(rotatingCategories || []).map(K)];
  const standardCats      = allCategories.filter(cat => !excludeCategories.includes(K(cat)));
  const products          = [];
  for (const cat of standardCats) {
    for (const name of getActiveProductsByCategory(warehouseSheet, cat)) {
      const extra = getArticulAndSupplier(warehouseSheet, name);
      products.push({ name, unit: unitMap[name] || '', articul: extra.articul, supplier: extra.supplier });
    }
  }
  return products;
}

function getSelectedAssortmentFromMenu(menuSheet, warehouseSheet) {
  if (!menuSheet) return {};

  const data        = menuSheet.getDataRange().getValues();
  const allProducts = getAllWarehouseProducts(warehouseSheet);
  const seen        = new Set();
  const ordered     = [];
  let   inAssortment = false;

  for (let i = 0; i < data.length; i++) {
    const labelRaw = data[i][0] ? data[i][0].toString().trim() : '';
    if (isMenuAssortLabel_(labelRaw)) { inAssortment = true; continue; }
    if (inAssortment) {
      if (isMenuAlwaysLabel_(labelRaw) || isMenuTimeLabel_(labelRaw)) break;
    }
    if (!inAssortment || !labelRaw) continue;

    const rowValues = [];
    for (let col = 1; col < data[i].length; col++) {
      const v = data[i][col] ? data[i][col].toString().trim() : '';
      if (v && !isMenuFiller_(v)) rowValues.push(v);
    }
    rowValues.forEach(product => {
      if (!allProducts.includes(product) || seen.has(product)) return;
      seen.add(product); ordered.push(product);
    });
  }

  const result = {};
  ordered.forEach((p, idx) => { result[idx] = p; });
  return result;
}

function getAllWarehouseCategories(warehouseSheet) {
  const data = warehouseSheet.getDataRange().getValues(), cats = [];
  for (let i = 2; i < data.length; i++) {
    const name    = data[i][0] ? data[i][0].toString().trim() : '';
    const hasPlus = data[i][6] && data[i][6].toString().includes('➕');
    if (hasPlus && name) cats.push(name);
  }
  return cats;
}

function getAllWarehouseProducts(warehouseSheet) {
  const data = warehouseSheet.getDataRange().getValues(), products = [];
  for (let i = 2; i < data.length; i++) {
    const name       = data[i][0] ? data[i][0].toString().trim() : '';
    const isCategory = data[i][6] && data[i][6].toString().includes('➕');
    if (name && !isCategory && data[i][1] !== '' && data[i][1] !== null) products.push(name);
  }
  return products;
}

function getActiveProductsByCategory(warehouseSheet, categoryName) {
  const data = warehouseSheet.getDataRange().getValues(), products = [];
  let inCategory = false;
  for (let i = 2; i < data.length; i++) {
    const name   = data[i][0] ? data[i][0].toString().trim() : '';
    const qty    = data[i][1], unit = data[i][2], active = data[i][3];
    const isCat  = (qty===''||qty===null)&&(unit===''||unit===null)&&(active===''||active===false||active===null);
    if (isCat && name) { if (name===categoryName) inCategory=true; else if (inCategory) break; continue; }
    if (inCategory && name && active===true) products.push(name);
  }
  return products;
}

function getArticulAndSupplier(warehouseSheet, productName) {
  const data    = warehouseSheet.getDataRange().getValues();
  const headers = data[1] || [];
  let articulIdx = -1, supplierIdx = -1;
  for (let c = 0; c < headers.length; c++) {
    const h = headers[c] ? hdrRu_(headers[c]) : '';
    if (h === 'Артикул')   articulIdx  = c;
    if (h === 'Поставщик') supplierIdx = c;
  }
  for (let i = 2; i < data.length; i++) {
    if ((data[i][0]||'').toString().trim() !== productName) continue;
    return {
      articul:  articulIdx  !== -1 ? (data[i][articulIdx]  || '') : '',
      supplier: supplierIdx !== -1 ? (data[i][supplierIdx] || '') : ''
    };
  }
  return { articul: '', supplier: '' };
}

function getUnitMapFromWarehouse(warehouseSheet) {
  const data = warehouseSheet.getDataRange().getValues(), unitMap = {};
  for (let i = 2; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    const unit = data[i][2] ? data[i][2].toString().trim() : '';
    if (name && unit) unitMap[name] = unit;
  }
  return unitMap;
}

function addSectionHeader(sheet, row, title) {
  sheet.getRange(row, 1, 1, 5).merge();
  sheet.getRange(row, 1).setValue(title)
    .setBackground('#d9ead3').setFontWeight('bold').setFontSize(11).setFontColor('#1c4587');
  sheet.setRowHeight(row, 26);
}

function writeOrderRow(sheet, row, name, unit, articul, supplier) {
  sheet.getRange(row, 1).setValue(name).setWrap(true);
  sheet.getRange(row, 2).setValue('');
  sheet.getRange(row, 3).setValue(unit     || '');
  sheet.getRange(row, 4).setValue(articul  || '');
  sheet.getRange(row, 5).setValue(supplier || '');
  sheet.setRowHeight(row, 21);
}

function formatOrderRows(sheet, startRow, endRow) {
  const colors = ['#ffffff', '#f8f9fa'];
  let colorIdx = 0;
  for (let r = startRow; r <= endRow; r++) {
    const val = sheet.getRange(r, 1).getValue();
    if (!val || isSectionHeader(val.toString()) || val.toString().trim() === '') continue;
    sheet.getRange(r, 1, 1, 5).setBackground(colors[colorIdx % 2]).setWrap(true);
    colorIdx++;
  }
  if (endRow >= startRow) sheet.autoResizeRows(startRow, endRow - startRow + 1);
}

function isSectionHeader(text) {
  return text.startsWith('📦') || text.startsWith('🗓');
}

function collectAllActiveProducts(warehouseSheet, unitMap) {
  const data = warehouseSheet.getDataRange().getValues(), products = [];
  let inCategory = false;
  for (let i = 2; i < data.length; i++) {
    const name  = data[i][0] ? data[i][0].toString().trim() : '';
    const qty   = data[i][1], unit = data[i][2], active = data[i][3];
    const isCat = (qty===''||qty===null)&&(unit===''||unit===null)&&(active===''||active===false||active===null);
    if (isCat && name) { inCategory = true; continue; }
    if (inCategory && name && active === true) {
      const extra = getArticulAndSupplier(warehouseSheet, name);
      products.push({ name, unit: unit || unitMap[name] || '', articul: extra.articul, supplier: extra.supplier });
    }
  }
  return products;
}

function getWarehouseSheetByType(warehouseType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return WAREHOUSE_SHEET_NAMES[warehouseType] ? SH_(ss, WAREHOUSE_SHEET_NAMES[warehouseType]) : null;
}

function getWarehouseDisplayName(warehouseType) {
  const names = { 'HOUSEHOLD': L('хозтоваров'), 'STATIONERY': L('канцелярии'), 'MEDICINE': L('медикаментов') };
  return names[warehouseType] || warehouseType;
}
