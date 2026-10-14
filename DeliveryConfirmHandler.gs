// DeliveryConfirmHandler.gs

function showDeliveryConfirmDialog(warehouseType) {
  try {
    const wh = WAREHOUSE_TYPES[warehouseType];
    if (!wh) {
      SpreadsheetApp.getUi().alert(L('❌ Неизвестный тип склада: ') + warehouseType);
      return;
    }

    const deliveries = _getDeliveryDataFromDashboard(warehouseType);

    if (deliveries.length === 0) {
      SpreadsheetApp.getUi().alert(
        L('📭 Нет данных о поставке\n\n') +
        L('Заполните форму логиста — данные появятся в дашборде\n') +
        L('в колонке "Поставка".\n\n') +
        L('Если поставка уже была подтверждена сегодня — новых\n') +
        L('непроверенных позиций нет.')
      );
      return;
    }

    const totalUnits = deliveries.reduce(function(sum, d) { return sum + d.fromForm; }, 0);

    const template           = htmlTemplate_('DeliveryConfirmDialog');
    template.warehouseType   = warehouseType;
    template.deliveries      = deliveries;
    template.totalUnits      = totalUnits;
    template.config          = { name: wh.name, icon: wh.icon, color: wh.color };

    SpreadsheetApp.getUi().showModalDialog(
      template.evaluate().setWidth(580).setHeight(640),
      wh.icon + L(' Подтверждение поставки — ') + wh.name
    );

  } catch(e) {
    SpreadsheetApp.getUi().alert(L('❌ Ошибка открытия диалога:\n') + e.toString());
  }
}

function _getDeliveryDataFromDashboard(warehouseType) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const cfg       = getWarehouseConfigByType(warehouseType);
  const dashSheet = SH_(ss, cfg.dashSheet);

  if (!dashSheet) return [];

  const data = dashSheet.getDataRange().getValues();
  if (data.length < 3) return [];

  const headers     = data[1] || [];
  let   postavkaIdx = -1;

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c] ? K(headers[c]) : '';
    if (h === 'Поставка') { postavkaIdx = c; break; }
  }

  if (postavkaIdx === -1) return [];

  const unitMap      = _getUnitMapForWarehouseType(ss, cfg);
  const todayStr     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const confirmedMap = _getConfirmedDeliveryMap(warehouseType, todayStr);

  const deliveries = [];

  for (let i = 2; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    if (!name) continue;

    const qty = data[i][postavkaIdx];
    if (qty === '' || qty === null || qty === undefined) continue;

    const raw = parseFloat(qty);
    if (isNaN(raw) || raw <= 0) continue;

    const alreadyConfirmed = confirmedMap[name] || 0;
    const remaining        = raw - alreadyConfirmed;
    if (remaining <= 0) continue; 

    deliveries.push({
      name:     name,
      fromForm: remaining,
      unit:     unitMap[name] || ''
    });
  }

  return deliveries;
}

function _getUnitMapForWarehouseType(ss, cfg) {
  const map = {};
  const warehouseSheet = SH_(ss, cfg.warehouseSheet);
  if (!warehouseSheet) return map;
  const data = warehouseSheet.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    const unit = data[i][2] ? data[i][2].toString().trim() : '';
    if (name && unit) map[name] = unit;
  }
  return map;
}


function _confirmedDeliveryPropKey(warehouseType, dateStr) {
  return 'CONFIRMED_DELIVERY_' + warehouseType + '_' + dateStr;
}

function _getConfirmedDeliveryMap(warehouseType, dateStr) {
  try {
    const key  = _confirmedDeliveryPropKey(warehouseType, dateStr);
    const json = PropertiesService.getScriptProperties().getProperty(key);
    return json ? JSON.parse(json) : {};
  } catch(e) {
    Logger.log('⚠️ _getConfirmedDeliveryMap: ' + e.toString());
    return {};
  }
}

function _addConfirmedDelivery(warehouseType, dateStr, formData) {
  try {
    const key = _confirmedDeliveryPropKey(warehouseType, dateStr);
    const map = _getConfirmedDeliveryMap(warehouseType, dateStr);
    for (const [name, item] of Object.entries(formData)) {
      const qty = parseFloat(item.qty) || 0;
      if (qty <= 0) continue;
      map[name] = (map[name] || 0) + qty;
    }
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(map));
  } catch(e) {
    Logger.log('⚠️ _addConfirmedDelivery: ' + e.toString());
  }
}

function processDeliveryConfirmationNew(warehouseType, formData) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = getWarehouseConfigByType(warehouseType);

    const confirmed = _addToWarehouseStock(ss, cfg, formData);

    if (confirmed.length === 0) {
      return { success: false, message: L('❌ Не удалось обновить остатки. Проверьте названия товаров на складе.') };
    }

    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    _addConfirmedDelivery(warehouseType, todayStr, formData);

    _updateDashboardDelivery(ss, cfg, formData);

    try {
      switch(warehouseType) {
        case 'PRODUCTS':   refreshProductsDashboard();       break;
        case 'HOUSEHOLD':  HouseholdDashboard.refresh();     break;
        case 'STATIONERY': swRefreshDashboard('STATIONERY'); break;
        case 'MEDICINE':   swRefreshDashboard('MEDICINE');   break;
      }
    } catch(e) { Logger.log('⚠️ refresh после поставки: ' + e.toString()); }

    try {
      _sendDeliveryConfirmToTelegram(warehouseType, confirmed, cfg);
    } catch(e) {
      Logger.log('⚠️ Telegram уведомление: ' + e.toString());
    }

    Logger.log('✅ Поставка ' + warehouseType + ' подтверждена: ' + confirmed.length + ' позиций');

    return {
      success: true,
      message: L('✅ Поставка подтверждена!\n\n') + confirmed.length + L(' позиций добавлено к остаткам.\nУведомления отправлены поставщикам.')
    };

  } catch(e) {
    Logger.log('❌ processDeliveryConfirmationNew: ' + e.toString());
    return { success: false, message: L('❌ Ошибка: ') + e.toString() };
  }
}

function _addToWarehouseStock(ss, cfg, formData) {
  const warehouseSheet = SH_(ss, cfg.warehouseSheet);
  if (!warehouseSheet) throw new Error(L('Склад "') + L(cfg.warehouseSheet) + L('" не найден'));

  const data     = warehouseSheet.getDataRange().getValues();
  const confirmed = [];

  for (const [productName, item] of Object.entries(formData)) {
    const qty = parseFloat(item.qty) || 0;
    if (qty <= 0) continue;

    let found = false;
    for (let i = 2; i < data.length; i++) {
      const name = data[i][0] ? data[i][0].toString().trim() : '';
      if (name !== productName) continue;

      const currentQty = parseFloat(data[i][1]) || 0;
      const newQty     = currentQty + qty;

      warehouseSheet.getRange(i + 1, 2).setValue(newQty);

      const supplier = data[i][5] ? data[i][5].toString().trim() : '';

      Logger.log('✓ Склад ' + productName + ': ' + currentQty + ' + ' + qty + ' = ' + newQty);
      confirmed.push({
        name:     productName,
        qty:      qty,
        unit:     item.unit || '',
        newStock: newQty,
        supplier: supplier
      });
      found = true;
      break;
    }

    if (!found) {
      Logger.log('⚠️ Товар не найден на складе: "' + productName + '"');
    }
  }

  SpreadsheetApp.flush();
  return confirmed;
}

function _updateDashboardDelivery(ss, cfg, formData) {
  const dashSheet = SH_(ss, cfg.dashSheet);
  if (!dashSheet) return;

  const data    = dashSheet.getDataRange().getValues();
  const headers = data[1] || [];

  let postavkaIdx = -1;
  for (let c = 0; c < headers.length; c++) {
    if (headers[c] && K(headers[c]) === 'Поставка') { postavkaIdx = c; break; }
  }
  if (postavkaIdx === -1) return;

  for (const [productName, item] of Object.entries(formData)) {
    const qty = parseFloat(item.qty) || 0;
    for (let i = 2; i < data.length; i++) {
      const name = data[i][0] ? data[i][0].toString().trim() : '';
      if (name !== productName) continue;
      dashSheet.getRange(i + 1, postavkaIdx + 1).setValue(qty);
      break;
    }
  }
}

function _sendDeliveryConfirmToTelegram(warehouseType, confirmed, cfg) {
  const defaultChatId   = TelegramManager.getChatId();
  const defaultThreadId = TelegramManager.getDefaultThreadId();
  const supplierChats   = getSupplierChatMap();
  const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  const wh     = WAREHOUSE_TYPES[warehouseType];

  const grouped = {};
  confirmed.forEach(function(item) {
    const supplier = item.supplier || L('Без поставщика');
    if (!grouped[supplier]) grouped[supplier] = [];
    grouped[supplier].push(item);
  });

  Object.entries(grouped).forEach(function([supplier, items]) {
    let text = cfg.icon + L(' <b>ПОСТАВКА ПОДТВЕРЖДЕНА</b>\n');
    text += '🏪 ' + wh.name + '\n';
    text += L('📦 Поставщик: <b>') + supplier + '</b>\n';
    text += '📅 ' + today + '\n';
    text += '─'.repeat(28) + '\n';

    items.forEach(function(item, idx) {
      const unit = item.unit ? ' ' + item.unit : L(' шт');
      text += '\n' + (idx + 1) + '. <b>' + item.name + '</b> — ' + item.qty + unit;
      text += L('  (склад: ') + item.newStock + unit + ')';
    });

    text += '\n\n' + '─'.repeat(28) + '\n';
    text += L('📦 Позиций принято: <b>') + items.length + '</b>\n';
    text += L('✅ Остатки обновлены');

    const hasOwnChat = !!supplierChats[supplier];
    const chatId      = hasOwnChat ? supplierChats[supplier] : defaultChatId;
    const threadId    = hasOwnChat ? null : defaultThreadId;

    const result = TelegramManager.sendMessage(chatId, text, threadId);
    if (!result.ok) {
      Logger.log('⚠️ Telegram sendMessage (' + supplier + '): ' + JSON.stringify(result));
    } else {
      Logger.log('✅ Уведомление о поставке (' + supplier + ') отправлено в Telegram');
    }
  });
}

function showDeliveryConfirmDialog_PRODUCTS()   { showDeliveryConfirmDialog('PRODUCTS');   }
function showDeliveryConfirmDialog_HOUSEHOLD()  { showDeliveryConfirmDialog('HOUSEHOLD');  }
function showDeliveryConfirmDialog_STATIONERY() { showDeliveryConfirmDialog('STATIONERY'); }
function showDeliveryConfirmDialog_MEDICINE()   { showDeliveryConfirmDialog('MEDICINE');   }

function cleanupOldConfirmedDeliveryProps() {
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  const props    = PropertiesService.getScriptProperties();
  const all      = props.getProperties();
  let removed = 0;
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('CONFIRMED_DELIVERY_') !== 0) return;
    if (key.indexOf(todayStr) !== -1) return; 
    props.deleteProperty(key);
    removed++;
  });
  Logger.log('✅ cleanupOldConfirmedDeliveryProps: удалено ' + removed);
}
