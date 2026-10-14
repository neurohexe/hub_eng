function onOpen() {
  updateAllWarehouseDates();
  updateAllDashboardDates();

  const ui  = SpreadsheetApp.getUi();
  const hub = isHubMode();

  if (!hub) {
    const warehouseMenu = ui.createMenu(L('Управление складами'));
    Object.keys(WAREHOUSE_TYPES).forEach(key => {
      const wh = WAREHOUSE_TYPES[key];
      warehouseMenu.addSubMenu(
        ui.createMenu(L(wh.name))
          .addItem(L('Выдать со склада'),         'showIssueDialog_'          + key)
          .addItem(L('Подтвердить поставку'),      'showDeliveryConfirmDialog_' + key)
          .addSeparator()
          .addItem(L('Добавить категорию'),        'addNewCategoryToWarehouse')
          .addItem(L('Добавление строк'),          'quickAddMultipleProductsInCurrentSheet')
          .addItem(L('Починить строки склада'),    'repairWarehouseRowsInCurrentSheet')
          .addSeparator()
          .addItem(L('Ссылки на формы'),           'showFormLinks_' + key)
      );
    });
    warehouseMenu.addToUi();

    ui.createMenu(L('Заказы'))
      .addSubMenu(ui.createMenu(L('Продукты'))
        .addItem(L('Обновить форму хозяюшек'),   'updateFormForHousekeepers')
        .addItem(L('Обновить лист заказа'),       'updateOrderSheetWrapper')
        .addItem(L('Отправить заказ (Telegram)'), 'sendOrderToTelegramWrapper')
        .addSeparator()
        .addItem(L('Обновить дашборд'),           'refreshProductsDashboardWrapper'))
      .addSubMenu(ui.createMenu(L('Хозтовары'))
        .addItem(L('Обновить форму хозяюшек'),   'updateHouseholdHostessFormManual')
        .addItem(L('Обновить форму клининга'),    'updateHouseholdCleaningFormManual')
        .addItem(L('Обновить лист заказа'),       'updateOrderSheetHousehold')
        .addItem(L('Отправить заказ (Telegram)'), 'sendOrderTelegramHousehold')
        .addSeparator()
        .addItem(L('Обновить дашборд'),           'refreshHouseholdDashboardWrapper'))
      .addSubMenu(ui.createMenu(L('Канцелярия'))
        .addItem(L('Обновить форму хозяюшек'),   'updateStationeryHostessFormManual')
        .addItem(L('Обновить лист заказа'),       'updateOrderSheetStationery')
        .addItem(L('Отправить заказ (Telegram)'), 'sendOrderTelegramStationery')
        .addSeparator()
        .addItem(L('Обновить дашборд'),           'refreshStationeryDashboardWrapper'))
      .addSubMenu(ui.createMenu(L('Медикаменты'))
        .addItem(L('Обновить форму хозяюшек'),   'updateMedicineHostessFormManual')
        .addItem(L('Обновить лист заказа'),       'updateOrderSheetMedicine')
        .addItem(L('Отправить заказ (Telegram)'), 'sendOrderTelegramMedicine')
        .addSeparator()
        .addItem(L('Обновить дашборд'),           'refreshMedicineDashboardWrapper'))
      .addToUi();
  } else {
    ui.createMenu(L('Хаб — остатки и заказы'))
      .addItem(L('Обновить форму остатков'),              'updateHubStockFormManual')
      .addItem(L('Менеджер для подтверждения заказов'),   'setResponsibleChatIdRow')
      .addItem(L('Массово добавить ассортимент'),          'bulkImportAssortmentManual')
      .addSeparator()
      .addItem(L('Добавить категорию (на текущем складе)'),         'addNewCategoryToWarehouse')
      .addItem(L('Добавление строк (на текущем складе)'),           'quickAddMultipleProductsInCurrentSheet')
      .addSeparator()
      .addItem(L('Скрыть неиспользуемые листы (дашборды/БД)'),      'hideHubUnusedSheetsManual')
      .addToUi();
  }

  ui.createMenu(L('Меню'))
    .addItem(L('Создать меню на следующую неделю'), 'createNextWeekMenuWrapper')
    .addItem(L('Конструктор меню'),                 'openMenuConstructorWrapper')
    .addSeparator()
    .addItem(L('Сохранить меню для планшетов'),     'forceSaveMenuDataFromSheets')
    .addItem(L('Обновить ассортимент в меню'),      'refreshMenuAssortment')
    .addSeparator()
    .addItem(L('🎲 Заполнить ассортимент по остаткам'), 'autoFillMenuByStock')
    .addToUi();

  let settingsMenu = ui.createMenu(L('Настройки'))
    .addSubMenu(ui.createMenu(L('Первоначальная настройка'))
      .addItem(L('Создать "Базовую информацию"'),    'createBaseInfoSheet')
      .addItem(L('Создать структуру склада'),         'createWarehouseStructure')
      .addItem(L('Массово добавить ассортимент'),     'bulkImportAssortmentManual')
      .addItem(L('Установить все триггеры'),          'setupAllTriggersWrapper')
      .addItem(L('Тип объекта (Полный / Хаб)'),       'chooseOfficeModeDialog'))
    .addSeparator()
    .addSubMenu(ui.createMenu(L('Поставщики'))
      .addItem(L('Добавить поставщика'),             'addSupplierRowWrapper')
      .addItem(L('Обновить поставщиков на складах'), 'updateSupplierDropdownsManual'))
    .addSubMenu(ui.createMenu(L('Единицы измерения'))
      .addItem(L('Управление единицами'),        'manageUnits')
      .addItem(L('Обновить единицы на складах'), 'updateUnitsOnWarehouses'))
    .addSubMenu(ui.createMenu(L('Telegram'))
      .addItem(L('Настроить Telegram'),          'connectTelegramChatWrapper')
      .addItem(L('Настроить тему (топик) чата'), 'setDefaultThreadIdWrapper')
      .addItem(L('Статус Telegram'),             'showTelegramStatusWrapper')
      .addItem(L('Настроить Webhook'),           'setupTelegramWebhookWrapper'))
    .addSubMenu(ui.createMenu(L('AI'))
      .addItem(L('Подключить AI-классификатор'), 'connectAIClassifierKey'));

  if (!hub) {
    settingsMenu = settingsMenu
      .addSeparator()
      .addItem(L('Архивировать все склады сейчас'), 'archiveAllWarehousesManual')
      .addItem(L('Конец дня (архив + сброс)'),       'endOfDayManual');
  }
  settingsMenu.addToUi();

  addLanguageMenu_(ui);
}

function setupAllTriggersWrapper() {
  try {
    const ui  = SpreadsheetApp.getUi();
    const hub = isHubMode();

    const dailyCycleLines = hub
      ? L('\n(режим "Хаб" — ежедневный архив и сброс дашбордов НЕ устанавливаются:\n') +
        L('у хаба нет ежедневного цикла выдач/поставок, только еженедельные остатки)')
      : L('• 23:30 — конец дня (архив)\n• 23:59 — сброс дашбордов\n');

    if (ui.alert(
      L('Установить триггеры'),
      L('Будут установлены:\n\n') +
      L('• onFormSubmit — автообновление') + (hub ? L(' (форма остатков хаба)') : L(' дашборда')) + '\n' +
      L('• Пн 00:05 — активация меню следующей недели\n') +
      L('• Пн 00:15 — удаление старых листов меню\n') +
      L('• 00:05 — обновление дат\n') +
      dailyCycleLines + L('\nПродолжить?'),
      ui.ButtonSet.YES_NO
    ) !== ui.Button.YES) return;

    const managed = [
      'onAnyFormSubmit', 'onHostessFormSubmit', 'onFormSubmitFOOD',
      'autoUpdateMenu', 'autoUpdateWarehouseDates', 'autoUpdateAllDates',
      'setupAllDailyTriggers', 'triggerDailyCleanup', 'resetProductsDashboardDaily',
      'resetAllDashboards', 'startOfDayProducts', 'endOfDayAll',
      'autoSaveDailyDataAllWarehouses', 'activateNextWeekMenu', 'cleanupOldMenuSheets',
      'promoteNextWeekMenu'
    ];
    ScriptApp.getProjectTriggers().forEach(t => {
      if (managed.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    ScriptApp.newTrigger('onAnyFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
    ScriptApp.newTrigger('autoUpdateAllDates').timeBased().everyDays(1).atHour(0).nearMinute(5).create();
    ScriptApp.newTrigger('activateNextWeekMenu').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0).nearMinute(5).create();
    ScriptApp.newTrigger('cleanupOldMenuSheets').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0).nearMinute(15).create();

    // Ежедневный архив и сброс дашбордов — только для полного склада (логист, ежедневный цикл).
    // У хаба этого цикла нет: остатки актуализируются раз в неделю через форму хозяюшки.
    if (!hub) {
      ScriptApp.newTrigger('endOfDayAll').timeBased().everyDays(1).atHour(23).nearMinute(30).create();
      ScriptApp.newTrigger('resetAllDashboards').timeBased().everyDays(1).atHour(23).nearMinute(59).create();
    }

    ui.alert(L('✅ Триггеры установлены!'));
  } catch(err) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + err.toString()); }
}

function onAnyFormSubmit(e) {
  try {
    Utilities.sleep(2000);
    const sheetName = (e && e.range) ? e.range.getSheet().getName() : '';
    Logger.log('onAnyFormSubmit: лист = "' + sheetName + '"');

    // ── НОВОЕ: ответы формы остатков хаба обрабатываются отдельно от дашбордов ──
    if (sheetName === HUB_STOCK_RESPONSE_SHEET) {
      onHubStockFormSubmit(e);
      return;
    }

    const warehouseType = FORM_SHEET_TO_DASHBOARD[sheetName];
    if (!warehouseType) { Logger.log('⚠️ Лист "' + sheetName + '" не в маппинге'); return; }
    switch (warehouseType) {
      case 'PRODUCTS':   refreshProductsDashboard();       break;
      case 'HOUSEHOLD':  HouseholdDashboard.refresh();     break;
      case 'STATIONERY': swRefreshDashboard('STATIONERY'); break;
      case 'MEDICINE':   swRefreshDashboard('MEDICINE');   break;
    }
    Logger.log('✅ Дашборд ' + warehouseType + ' обновлён');
  } catch(err) { Logger.log('❌ onAnyFormSubmit: ' + err.toString()); }
}

function updateAllWarehouseDates() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    Object.keys(WAREHOUSE_TYPES).forEach(key => {
      const sheet = SH_(ss, WAREHOUSE_SHEET_NAMES[key]);
      if (!sheet) return;
      const newTitle = warehouseTitle_(key, today);
      const cell = sheet.getRange(1, 1);
      if (cell.getValue().toString() !== newTitle) cell.setValue(newTitle);
    });
    applyLanguageToWarehouseHeaders_();
  } catch(err) { Logger.log('⚠️ updateAllWarehouseDates: ' + err.toString()); }
}

function autoUpdateAllDates() {
  updateAllWarehouseDates();
  updateAllDashboardDates();
}

function updateAllDashboardDates() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    [L('Продукты'), L('Хозтовары'), L('Канцелярия'), L('Медикаменты')].forEach(name => {
      const sheet = SH_(ss, name);
      if (!sheet || sheet.getLastColumn() === 0) return;
      sheet.getRange(1, 1).setValue(today).setNumberFormat('@STRING@');
    });
  } catch(err) { Logger.log('⚠️ updateAllDashboardDates: ' + err.toString()); }
}

function showFormLinks_PRODUCTS() {
  try {
    const logistId  = PropertiesService.getScriptProperties().getProperty('DELIVERY_FORM_ID');
    const hostessId = PropertiesService.getScriptProperties().getProperty('PRODUCTS_FORM_ID');
    SpreadsheetApp.getUi().alert(
      L('Формы Продукты:\n\nЛогист:\n') + (logistId  ? _getFormUrl(logistId)  : L('не создана')) +
      L('\n\nХозяюшки:\n')              + (hostessId ? _getFormUrl(hostessId) : L('не создана'))
    );
  } catch(err) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + err.toString()); }
}

function showFormLinks_HOUSEHOLD() {
  try {
    const l = HouseholdFormManager.getFormLinks();
    SpreadsheetApp.getUi().alert(
      L('Формы Хозтовары:\n\nЛогист:\n')   + (l.logist   ? l.logist.short   : L('не создана')) +
      L('\n\nХозяюшки:\n') + (l.hostess  ? l.hostess.short  : L('не создана')) +
      L('\n\nКлининг:\n')  + (l.cleaning ? l.cleaning.short : L('не создана'))
    );
  } catch(err) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + err.toString()); }
}

function showFormLinks_STATIONERY() {
  try {
    const l = swGetFormLinks('STATIONERY');
    SpreadsheetApp.getUi().alert(
      L('Формы Канцелярия:\n\nЛогист:\n')   + (l.logist  ? l.logist.short  : L('не создана')) +
      L('\n\nХозяюшки:\n') + (l.hostess ? l.hostess.short : L('не создана'))
    );
  } catch(err) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + err.toString()); }
}

function showFormLinks_MEDICINE() {
  try {
    const l = swGetFormLinks('MEDICINE');
    SpreadsheetApp.getUi().alert(
      L('Формы Медикаменты:\n\nЛогист:\n')  + (l.logist  ? l.logist.short  : L('не создана')) +
      L('\n\nХозяюшки:\n') + (l.hostess ? l.hostess.short : L('не создана'))
    );
  } catch(err) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + err.toString()); }
}

function _getFormUrl(formId) {
  try { const f = FormApp.openById(formId); return f.shortenFormUrl(f.getPublishedUrl()); }
  catch(e) { return L('недоступна'); }
}

function showIssueDialog_PRODUCTS()   { showIssueDialog('PRODUCTS');   }
function showIssueDialog_HOUSEHOLD()  { showIssueDialog('HOUSEHOLD');  }
function showIssueDialog_STATIONERY() { showIssueDialog('STATIONERY'); }
function showIssueDialog_MEDICINE()   { showIssueDialog('MEDICINE');   }

function showIssueDialog(warehouseType) {
  try {
    const wh = WAREHOUSE_TYPES[warehouseType];
    if (!wh) { SpreadsheetApp.getUi().alert(L('Неизвестный тип: ') + warehouseType); return; }
    const { products, totalOrdered } = _getIssueProducts(warehouseType);
    const template = htmlTemplate_('IssueDialog');
    template.warehouseType = warehouseType;
    template.products      = products;
    template.totalOrdered  = totalOrdered;
    template.config        = { name: wh.name, icon: wh.icon, color: wh.color };
    SpreadsheetApp.getUi().showModalDialog(
      template.evaluate().setWidth(560).setHeight(620),
      L('Выдача со склада — ') + wh.name
    );
  } catch(e) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + e.toString()); }
}

function _getIssueProducts(warehouseType) {
  const { products } = ProductWarehouse.getProductsToIssue(warehouseType);
  let totalOrdered = 0;
  const enriched = products.map(p => {
    const alreadyIssued = p.issued || 0;
    totalOrdered += p.ordered;
    return { name: p.name, ordered: p.ordered, alreadyIssued, newToIssue: Math.max(0, p.ordered - alreadyIssued) };
  });
  return { products: enriched, totalOrdered };
}

function updateOrderSheetWrapper()     { updateOrderSheet(); }
function sendOrderToTelegramWrapper()  { sendOrderToTelegramAndUpdateForm(); }
function updateOrderSheetHousehold()   { updateOrderSheetForWarehouse('HOUSEHOLD');  }
function updateOrderSheetStationery()  { updateOrderSheetForWarehouse('STATIONERY'); }
function updateOrderSheetMedicine()    { updateOrderSheetForWarehouse('MEDICINE');   }
function sendOrderTelegramHousehold()  { sendOrderToTelegramForWarehouse('HOUSEHOLD');  }
function sendOrderTelegramStationery() { sendOrderToTelegramForWarehouse('STATIONERY'); }
function sendOrderTelegramMedicine()   { sendOrderToTelegramForWarehouse('MEDICINE');   }

function updateFormForHousekeepers() { updateForm(); }

function createNextWeekMenuWrapper()  { MenuManager.createNextWeekMenu(); }
function openMenuConstructorWrapper() { openMenuConstructor(); }

function connectTelegramChatWrapper() { TelegramManager.connectChat(); }
function setDefaultThreadIdWrapper()  { TelegramManager.setDefaultThreadId(); }
function showTelegramStatusWrapper()  { TelegramManager.showStatus();  }
function setupTelegramWebhookWrapper(){ TelegramManager.setupWebhook(); }

function refreshProductsDashboardWrapper()  { refreshProductsDashboard();       SpreadsheetApp.getUi().alert(L('Дашборд Продукты обновлён.'));     }
function refreshHouseholdDashboardWrapper() { HouseholdDashboard.refresh();     SpreadsheetApp.getUi().alert(L('Дашборд Хозтовары обновлён.'));    }
function refreshStationeryDashboardWrapper(){ swRefreshDashboard('STATIONERY'); SpreadsheetApp.getUi().alert(L('Дашборд Канцелярия обновлён.'));   }
function refreshMedicineDashboardWrapper()  { swRefreshDashboard('MEDICINE');   SpreadsheetApp.getUi().alert(L('Дашборд Медикаменты обновлён.'));  }
function refreshHouseholdDashboard()  { HouseholdDashboard.refresh(); }
function refreshStationeryDashboard() { swRefreshDashboard('STATIONERY'); }
function refreshMedicineDashboard()   { swRefreshDashboard('MEDICINE'); }

function updateSupplierDropdownsManual() { updateSupplierDropdowns(); }

function processIssueConfirmation(warehouseType, formData) {
  try {
    Logger.log('processIssueConfirmation: тип=' + warehouseType + ' user=' + Session.getActiveUser().getEmail());
    const result = ProductWarehouse.issueProducts(warehouseType, formData);
    Logger.log('processIssueConfirmation: ' + JSON.stringify(result));
    return result;
  } catch(err) {
    Logger.log('❌ processIssueConfirmation: ' + err.toString());
    return { success: false, message: '❌ ' + err.toString() };
  }
}

function hideAllResponseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ['food_hostess','postavka_food','household_hostess','household_cleaning',
    'postavka_household','stationery_hostess','postavka_stationery',
    'medicine_hostess','postavka_medicine'];
  let hidden = 0;
  names.forEach(name => {
    const sheet = SH_(ss, name);
    if (sheet && !sheet.isSheetHidden()) { sheet.hideSheet(); hidden++; }
  });
  SpreadsheetApp.getUi().alert(L('Скрыто листов: ') + hidden);
}

function getWarehouseConfig(warehouseType) {
  const wh = WAREHOUSE_TYPES[warehouseType];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return { type: warehouseType, name: wh.name, icon: wh.icon, color: wh.color,
    warehouseSheet: SH_(ss, WAREHOUSE_SHEET_NAMES[warehouseType]) };
}

function escapeHtml(str) {
  return String(str).replace(/&/g,'&amp;').replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
