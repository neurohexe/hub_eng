function attachFormToSheet(form, targetName, headerColor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('🔁 [' + targetName + '] Начинаем пересоздание листа ответов...');

  try {
    form.removeDestination();
    SpreadsheetApp.flush();
    Logger.log('✅ [' + targetName + '] Шаг 1: форма отвязана');
  } catch(e) {
    Logger.log('ℹ️ [' + targetName + '] Шаг 1: форма не была привязана');
  }

  var old = SH_(ss, targetName);
  if (old) {
    var removed = false;
    for (var t = 0; t < 5 && !removed; t++) {
      try {
        ss.deleteSheet(old);
        removed = true;
        Logger.log('✅ [' + targetName + '] Шаг 2: старый лист удалён');
      } catch(e) {
        Utilities.sleep(2000);
        SpreadsheetApp.flush();
      }
    }
    if (!removed) {
      try {
        old.setName('_tmp_' + targetName + '_' + old.getSheetId());
        old.hideSheet();
        Logger.log('⚠️ [' + targetName + '] Шаг 2: старый лист не удалился, уведён во временный');
      } catch(e2) {
        throw new Error(L('Не удалось освободить лист "') + targetName + '": ' + e2.message);
      }
    }
  }

  var idsBefore = ss.getSheets().map(function(s) { return s.getSheetId(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Logger.log('✅ [' + targetName + '] Шаг 3: форма привязана к таблице');

  var ok = _waitRenameHide(idsBefore, targetName, headerColor);

  ss.getSheets().forEach(function(s) {
    if (s.getName().indexOf('_tmp_' + targetName + '_') === 0) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  if (!ok) throw new Error(L('Новый лист ответов "') + targetName + L('" не появился вовремя'));
  Logger.log('✅ [' + targetName + '] Готово');
  return ok;
}

function _waitRenameHide(idsBefore, targetName, headerColor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var patterns = [/^ответы на форму/i, /^form responses/i, /^responses/i]; // имена листов, которые даёт Google (зависят от языка аккаунта)

  for (var attempt = 1; attempt <= 15; attempt++) {
    Utilities.sleep(2000);
    SpreadsheetApp.flush();
    var sheets = ss.getSheets();
    var newSheet = null;

    for (var i = 0; i < sheets.length; i++) {
      if (idsBefore.indexOf(sheets[i].getSheetId()) === -1) {
        newSheet = sheets[i];
        Logger.log('✅ [' + targetName + '] найден по ID: "' + newSheet.getName() + '"');
        break;
      }
    }

    if (!newSheet) {
      for (var j = 0; j < sheets.length; j++) {
        if (patterns.some(function(p) { return p.test(sheets[j].getName()); })) {
          newSheet = sheets[j];
          Logger.log('✅ [' + targetName + '] найден по паттерну: "' + newSheet.getName() + '"');
          break;
        }
      }
    }

    if (!newSheet) {
      if (attempt % 5 === 0) Logger.log('⏳ [' + targetName + '] ждём новый лист... попытка ' + attempt + '/15');
      continue;
    }

    var clash = SH_(ss, targetName);
    if (clash && clash.getSheetId() !== newSheet.getSheetId()) {
      try { ss.deleteSheet(clash); SpreadsheetApp.flush(); Logger.log('🗑️ [' + targetName + '] удалён дублирующий лист'); } catch(e) {}
    }

    try {
      newSheet.setName(targetName);
      newSheet.hideSheet();
      _applySheetHeader(newSheet, headerColor);
      Logger.log('✅ [' + targetName + '] лист ответов готов');
      return true;
    } catch(e) {
      Logger.log('⚠️ [' + targetName + '] попытка ' + attempt + ': ' + e.message);
    }
  }

  Logger.log('❌ [' + targetName + '] новый лист не появился');
  return false;
}

function _applySheetHeader(sheet, color) {
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    sheet.getRange(1, 1, 1, lastCol)
      .setBackground(color || '#1c4587')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  } catch(e) {
    Logger.log('⚠️ _applySheetHeader: ' + e.toString());
  }
}

function waitAndRenameSheet(idsBefore, targetName, headerColor) {
  return _waitRenameHide(idsBefore, targetName, headerColor);
}

function _getOrCreateForm(propKey, defaultTitle) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(propKey);

  if (id) {
    try {
      var form = FormApp.openById(id);
      Logger.log('✅ [' + propKey + '] используем существующую: ' + id);
      return form;
    } catch(e) {
      throw new Error(
        L('Форма "') + propKey + '" (ID: ' + id + L(') не открывается: ') + e.message +
        L('. Новая форма НЕ создана, чтобы не менять постоянную ссылку. ') +
        L('Проверьте, не удалена ли форма в корзину.')
      );
    }
  }

  var newForm = FormApp.create(defaultTitle);
  props.setProperty(propKey, newForm.getId());
  Logger.log('🆕 [' + propKey + '] создана форма (единожды): ' + newForm.getId());
  return newForm;
}

function _saveFormLink(propKey, form) {
  try {
    var url = form.getPublishedUrl();
    var short;
    try {
      short = form.shortenFormUrl(url);
    } catch(e) {
      short = url;
    }
    PropertiesService.getScriptProperties().setProperty(propKey, form.getId());
    PropertiesService.getScriptProperties().setProperty(propKey + '_URL', url);
    PropertiesService.getScriptProperties().setProperty(propKey + '_SHORT', short);
    Logger.log('✓ [' + propKey + '] ссылка сохранена: ' + short);
  } catch(e) {
    Logger.log('⚠️ _saveFormLink [' + propKey + ']: ' + e.toString());
  }
}

function _saveFormId(propKey, form) {
  _saveFormLink(propKey, form);
}

function _formResult(form) {
  var url = form.getPublishedUrl();
  return { url: url, short: form.shortenFormUrl(url) };
}

function updateProductsHostessForm(ss, weekPeriod, menuSheet) {
  var coffeePoints = BaseInfo.getActiveCoffeePoints();
  if (!coffeePoints || coffeePoints.length === 0) throw new Error(L('Нет активных кофе-поинтов'));

  var standardProducts = _getStandardProductsForForm(ss);
  if (standardProducts.length === 0) throw new Error(L('Нет товаров для страницы 1.'));

  var menuProducts = _getMenuProductsFromMenuSheet(menuSheet, ss);
  if (menuProducts.length === 0) throw new Error(
    L('Нет товаров для страницы 2.\nВыберите ассортимент в листе меню "') + (menuSheet ? menuSheet.getName() : '') + '"'
  );

  var form       = _getOrCreateForm('PRODUCTS_FORM_ID', L('Заказ продуктов — форма хозяюшек'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Заказ продуктов — хозяюшки'));
  form.setDescription(L('Период: ') + weekPeriod + L('\n\nСтр. 1 — стандарт | Стр. 2 — по выбору меню'));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setProgressBar(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addMultipleChoiceItem()
    .setTitle(L('Выберите ваш кофе-поинт'))
    .setChoiceValues(coffeePoints.map(function(cp) { return cp.name; }))
    .setRequired(true);
  standardProducts.forEach(function(p) {
    form.addTextItem().setTitle(p.name).setHelpText(p.unit || '').setRequired(false);
  });
  form.addPageBreakItem().setTitle(L('Ассортимент по выбору')).setHelpText(L('Период: ') + weekPeriod);
  menuProducts.forEach(function(p) {
    form.addTextItem().setTitle(p.name).setHelpText(p.unit || '').setRequired(false);
  });

  _saveFormId('PRODUCTS_FORM_ID', form);
  attachFormToSheet(form, 'food_hostess', '#1c4587');

  Logger.log('✅ Форма продукты хозяюшки обновлена: ' + standardProducts.length + ' + ' + menuProducts.length + ' позиций');
  return _formResult(form);
}

function updateProductsLogistForm(orderItems, weekPeriod) {
  var form       = _getOrCreateForm('DELIVERY_FORM_ID', L('Поставка продуктов — форма логиста'));
  var officeName = getOfficeName();
  var today      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

  form.setTitle((officeName ? officeName + '. ' : '') + L('Поставка продуктов — логист'));
  form.setDescription(L('Подтвердите фактически полученные товары.\n\nДата: ') + today + L('\nПериод: ') + (weekPeriod || ''));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setProgressBar(true);
  form.setAcceptingResponses(true); form.setShuffleQuestions(false);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addDateItem().setTitle(L('Дата поставки')).setRequired(true);
  form.addTextItem().setTitle(L('Поставщик / водитель (ФИО или название)')).setRequired(false);

  var grouped = _groupBySupplier(orderItems);
  Object.entries(grouped).forEach(function(entry) {
    var supplier = entry[0], items = entry[1];
    form.addPageBreakItem().setTitle('📦 ' + supplier)
      .setHelpText(L('Укажите фактически полученное количество. Если товар не поставлен — оставьте пустым.'));
    items.forEach(function(item) {
      var art = (item.articul && item.articul.toString().trim()) ? L('  |  Артикул: ') + item.articul : '';
      form.addTextItem().setTitle(item.name)
        .setHelpText(L('Заказано: ') + item.qty + ' ' + (item.unit || L('шт')) + art)
        .setRequired(false);
    });
  });
  form.addPageBreakItem().setTitle(L('💬 Комментарий к поставке'));
  form.addParagraphTextItem().setTitle(L('Комментарий'))
    .setHelpText(L('Недостачи, замены, повреждения, замечания по качеству')).setRequired(false);

  _saveFormId('DELIVERY_FORM_ID', form);
  attachFormToSheet(form, 'postavka_food', '#1c4587');

  Logger.log('✅ Форма продукты логист обновлена');
  return _formResult(form);
}

function updateHouseholdHostessForm(coffeePoints) {
  var products = HouseholdWarehouse.getProducts('hostess');
  if (products.length === 0) return null;

  var form       = _getOrCreateForm('HOUSEHOLD_HOSTESS_FORM_ID', L('Заказ хозтоваров — хозяюшки'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Заказ хозтоваров — хозяюшки'));
  form.setDescription(L('Укажите необходимое количество.'));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setProgressBar(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addMultipleChoiceItem()
    .setTitle(L('Выберите ваш кофе-поинт'))
    .setChoiceValues(coffeePoints.map(function(cp) { return cp.name; }))
    .setRequired(true);
  products.forEach(function(p) {
    form.addTextItem().setTitle(p.name).setHelpText(p.unit || '').setRequired(false);
  });

  _saveFormId('HOUSEHOLD_HOSTESS_FORM_ID', form);
  attachFormToSheet(form, 'household_hostess', '#2196F3');

  Logger.log('✅ Форма хозтовары хозяюшки обновлена');
  return _formResult(form);
}

function updateHouseholdCleaningForm() {
  var products = HouseholdWarehouse.getProducts('cleaning');
  if (products.length === 0) return null;

  var form       = _getOrCreateForm('HOUSEHOLD_CLEANING_FORM_ID', L('Заказ хозтоваров — клининг'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Заказ хозтоваров — клининг'));
  form.setDescription(L('Укажите необходимое количество.'));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setProgressBar(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addTextItem().setTitle(L('Ваше имя или подразделение')).setRequired(false);
  products.forEach(function(p) {
    form.addTextItem().setTitle(p.name).setHelpText(p.unit || '').setRequired(false);
  });

  _saveFormId('HOUSEHOLD_CLEANING_FORM_ID', form);
  attachFormToSheet(form, 'household_cleaning', '#2196F3');

  Logger.log('✅ Форма хозтовары клининг обновлена');
  return _formResult(form);
}

function updateHouseholdLogistForm(orderItems, weekPeriod) {
  var form       = _getOrCreateForm('HOUSEHOLD_LOGIST_FORM_ID', L('Поставка хозтоваров — форма логиста'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Поставка хозтоваров — логист'));
  form.setDescription(L('Подтвердите полученные товары. Период: ') + (weekPeriod || ''));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addDateItem().setTitle(L('Дата поставки')).setRequired(true);
  form.addTextItem().setTitle(L('Поставщик / водитель')).setRequired(false);

  var grouped = _groupBySupplier(orderItems);
  Object.entries(grouped).forEach(function(entry) {
    var supplier = entry[0], items = entry[1];
    form.addPageBreakItem().setTitle(supplier).setHelpText(L('Оставьте поле пустым, если не поставлено.'));
    items.forEach(function(item) {
      form.addTextItem().setTitle(item.name)
        .setHelpText(L('Заказано: ') + item.qty + ' ' + (item.unit || L('шт')) + (item.articul ? '  |  ' + item.articul : ''))
        .setRequired(false);
    });
  });
  form.addPageBreakItem().setTitle(L('Комментарий'));
  form.addParagraphTextItem().setTitle(L('Комментарий')).setRequired(false);

  _saveFormId('HOUSEHOLD_LOGIST_FORM_ID', form);
  attachFormToSheet(form, 'postavka_household', '#2196F3');

  Logger.log('✅ Форма хозтовары логист обновлена');
  return _formResult(form);
}

function updateStationeryHostessForm() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var warehouseSheet = SH_(ss, L('Склад_Канцелярия'));
  if (!warehouseSheet) throw new Error(L('Лист "Склад_Канцелярия" не найден'));
  var coffeePoints = BaseInfo.getActiveCoffeePoints();
  if (coffeePoints.length === 0) throw new Error(L('Нет активных кофе-поинтов'));
  var products = _getSimpleActiveProducts(warehouseSheet, 'hostess');
  if (products.length === 0) return null;

  var form       = _getOrCreateForm('STATIONERY_HOSTESS_FORM_ID', L('Заказ канцелярии — хозяюшки'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Заказ канцелярии — хозяюшки'));
  form.setDescription(L('✏️ Укажите необходимое количество.'));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setProgressBar(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addMultipleChoiceItem()
    .setTitle(L('Выберите ваш кофе-поинт'))
    .setChoiceValues(coffeePoints.map(function(cp) { return cp.name; }))
    .setRequired(true);
  products.forEach(function(p) {
    form.addTextItem().setTitle(p.name).setHelpText(p.unit || '').setRequired(false);
  });

  _saveFormId('STATIONERY_HOSTESS_FORM_ID', form);
  attachFormToSheet(form, 'stationery_hostess', '#FF9800');

  Logger.log('✅ Форма канцелярия хозяюшки обновлена');
  return _formResult(form);
}

function updateStationeryLogistForm(orderItems, weekPeriod) {
  var form       = _getOrCreateForm('STATIONERY_LOGIST_FORM_ID', L('Поставка канцелярии — форма логиста'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Поставка канцелярии — логист'));
  form.setDescription(L('✏️ Подтвердите полученные товары.\n\nПериод: ') + (weekPeriod || ''));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addDateItem().setTitle(L('Дата поставки')).setRequired(true);
  form.addTextItem().setTitle(L('Поставщик / водитель')).setRequired(false);

  var grouped = _groupBySupplier(orderItems);
  Object.entries(grouped).forEach(function(entry) {
    var supplier = entry[0], items = entry[1];
    form.addPageBreakItem().setTitle(supplier).setHelpText(L('Оставьте поле пустым, если не поставлено.'));
    items.forEach(function(item) {
      form.addTextItem().setTitle(item.name)
        .setHelpText(L('Заказано: ') + item.qty + ' ' + (item.unit || L('шт')) + (item.articul ? '  |  ' + item.articul : ''))
        .setRequired(false);
    });
  });
  form.addPageBreakItem().setTitle(L('Комментарий'));
  form.addParagraphTextItem().setTitle(L('Комментарий')).setRequired(false);

  _saveFormId('STATIONERY_LOGIST_FORM_ID', form);
  attachFormToSheet(form, 'postavka_stationery', '#FF9800');

  Logger.log('✅ Форма канцелярия логист обновлена');
  return _formResult(form);
}

function updateMedicineHostessForm() {
  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var warehouseSheet = SH_(ss, L('Склад_Медикаменты'));
  if (!warehouseSheet) throw new Error(L('Лист "Склад_Медикаменты" не найден'));
  var coffeePoints = BaseInfo.getActiveCoffeePoints();
  if (coffeePoints.length === 0) throw new Error(L('Нет активных кофе-поинтов'));
  var products = _getSimpleActiveProducts(warehouseSheet, 'hostess');
  if (products.length === 0) return null;

  var form       = _getOrCreateForm('MEDICINE_HOSTESS_FORM_ID', L('Заказ медикаментов — хозяюшки'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Заказ медикаментов — хозяюшки'));
  form.setDescription(L('💊 Укажите необходимое количество.'));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setProgressBar(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addMultipleChoiceItem()
    .setTitle(L('Выберите ваш кофе-поинт'))
    .setChoiceValues(coffeePoints.map(function(cp) { return cp.name; }))
    .setRequired(true);
  products.forEach(function(p) {
    form.addTextItem().setTitle(p.name).setHelpText(p.unit || '').setRequired(false);
  });

  _saveFormId('MEDICINE_HOSTESS_FORM_ID', form);
  attachFormToSheet(form, 'medicine_hostess', '#E91E63');

  Logger.log('✅ Форма медикаменты хозяюшки обновлена');
  return _formResult(form);
}

function updateMedicineLogistForm(orderItems, weekPeriod) {
  var form       = _getOrCreateForm('MEDICINE_LOGIST_FORM_ID', L('Поставка медикаментов — форма логиста'));
  var officeName = getOfficeName();

  form.setTitle((officeName ? officeName + '. ' : '') + L('Поставка медикаментов — логист'));
  form.setDescription(L('💊 Подтвердите полученные товары.\n\nПериод: ') + (weekPeriod || ''));
  form.setCollectEmail(false); form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true); form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(function(i) { form.deleteItem(i); });
  form.addDateItem().setTitle(L('Дата поставки')).setRequired(true);
  form.addTextItem().setTitle(L('Поставщик / водитель')).setRequired(false);

  var grouped = _groupBySupplier(orderItems);
  Object.entries(grouped).forEach(function(entry) {
    var supplier = entry[0], items = entry[1];
    form.addPageBreakItem().setTitle(supplier).setHelpText(L('Оставьте поле пустым, если не поставлено.'));
    items.forEach(function(item) {
      form.addTextItem().setTitle(item.name)
        .setHelpText(L('Заказано: ') + item.qty + ' ' + (item.unit || L('шт')) + (item.articul ? '  |  ' + item.articul : ''))
        .setRequired(false);
    });
  });
  form.addPageBreakItem().setTitle(L('Комментарий'));
  form.addParagraphTextItem().setTitle(L('Комментарий')).setRequired(false);

  _saveFormId('MEDICINE_LOGIST_FORM_ID', form);
  attachFormToSheet(form, 'postavka_medicine', '#E91E63');

  Logger.log('✅ Форма медикаменты логист обновлена');
  return _formResult(form);
}

function updateAllHouseholdForms(orderItems, weekPeriod) {
  var results = {};
  try { results.logist   = updateHouseholdLogistForm(orderItems, weekPeriod);              Logger.log('✅ Логист HH'); }   catch(e) { Logger.log('❌ logist HH: '+e);   results.logistError   = e.toString(); }
  try { results.hostess  = updateHouseholdHostessForm(BaseInfo.getActiveCoffeePoints());   Logger.log('✅ Хозяюшки HH'); } catch(e) { Logger.log('❌ hostess HH: '+e);  results.hostessError  = e.toString(); }
  try { results.cleaning = updateHouseholdCleaningForm();                                  Logger.log('✅ Клининг HH'); }  catch(e) { Logger.log('❌ cleaning HH: '+e); results.cleaningError = e.toString(); }
  return results;
}

function updateAllStationeryForms(orderItems, weekPeriod) {
  var results = {};
  try { results.logist  = updateStationeryLogistForm(orderItems, weekPeriod); Logger.log('✅ Логист ST'); }  catch(e) { Logger.log('❌ logist ST: '+e);  results.logistError  = e.toString(); }
  try { results.hostess = updateStationeryHostessForm();                       Logger.log('✅ Хозяюшки ST'); } catch(e) { Logger.log('❌ hostess ST: '+e); results.hostessError = e.toString(); }
  return results;
}

function updateAllMedicineForms(orderItems, weekPeriod) {
  var results = {};
  try { results.logist  = updateMedicineLogistForm(orderItems, weekPeriod); Logger.log('✅ Логист MD'); }  catch(e) { Logger.log('❌ logist MD: '+e);  results.logistError  = e.toString(); }
  try { results.hostess = updateMedicineHostessForm();                       Logger.log('✅ Хозяюшки MD'); } catch(e) { Logger.log('❌ hostess MD: '+e); results.hostessError = e.toString(); }
  return results;
}

function getFormLink(propKey) {
  var cachedUrl   = PropertiesService.getScriptProperties().getProperty(propKey + '_URL');
  var cachedShort = PropertiesService.getScriptProperties().getProperty(propKey + '_SHORT');

  if (cachedUrl && cachedShort) {
    return { url: cachedUrl, short: cachedShort };
  }

  var id = PropertiesService.getScriptProperties().getProperty(propKey);
  if (!id) return null;

  try {
    var f   = FormApp.openById(id);
    var url = f.getPublishedUrl();
    var short;
    try {
      short = f.shortenFormUrl(url);
    } catch(e) {
      short = url;
    }
    PropertiesService.getScriptProperties().setProperty(propKey + '_URL',   url);
    PropertiesService.getScriptProperties().setProperty(propKey + '_SHORT', short);
    return { url, short };
  } catch(e) {
    Logger.log('⚠️ getFormLink [' + propKey + ']: ' + e.toString());
    var directUrl = 'https://docs.google.com/forms/d/' + id + '/viewform';
    return { url: directUrl, short: directUrl };
  }
}

function getAllFormLinks() {
  return {
    products: {
      hostess: getFormLink('PRODUCTS_FORM_ID'),
      logist:  getFormLink('DELIVERY_FORM_ID')
    },
    household: {
      hostess:  getFormLink('HOUSEHOLD_HOSTESS_FORM_ID'),
      cleaning: getFormLink('HOUSEHOLD_CLEANING_FORM_ID'),
      logist:   getFormLink('HOUSEHOLD_LOGIST_FORM_ID')
    },
    stationery: {
      hostess: getFormLink('STATIONERY_HOSTESS_FORM_ID'),
      logist:  getFormLink('STATIONERY_LOGIST_FORM_ID')
    },
    medicine: {
      hostess: getFormLink('MEDICINE_HOSTESS_FORM_ID'),
      logist:  getFormLink('MEDICINE_LOGIST_FORM_ID')
    }
  };
}

function updateForm() {
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var menuSheet = _getLatestMenuSheetForForm(ss);
    if (!menuSheet) {
      SpreadsheetApp.getUi().alert(L('Сначала создайте меню!\nМеню → Создать меню на следующую неделю'));
      return;
    }
    var weekPeriod = getWeekPeriodFromMenu(menuSheet);
    var result     = updateProductsHostessForm(ss, weekPeriod, menuSheet);
    _showFormUpdateDialog(result, L('Ссылка для хозяюшек (продукты)'));
  } catch(e) {
    Logger.log('❌ updateForm: ' + e.toString());
    SpreadsheetApp.getUi().alert(L('Ошибка: ') + e.toString());
  }
}

function updateHouseholdHostessFormManual() {
  try {
    var r = updateHouseholdHostessForm(BaseInfo.getActiveCoffeePoints());
    SpreadsheetApp.getUi().alert(r ? L('Форма хозяюшек обновлена!\n\n') + r.short : L('Нет товаров с отметкой "Хозяюшки".'));
  } catch(e) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + e.toString()); }
}

function updateHouseholdCleaningFormManual() {
  try {
    var r = updateHouseholdCleaningForm();
    SpreadsheetApp.getUi().alert(r ? L('Форма клининга обновлена!\n\n') + r.short : L('Нет товаров с отметкой "Клининг".'));
  } catch(e) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + e.toString()); }
}

function updateStationeryHostessFormManual() {
  try {
    var r = updateStationeryHostessForm();
    SpreadsheetApp.getUi().alert(r ? L('Форма хозяюшек (канцелярия) обновлена!\n\n') + r.short : L('Нет активных товаров.'));
  } catch(e) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + e.toString()); }
}

function updateMedicineHostessFormManual() {
  try {
    var r = updateMedicineHostessForm();
    SpreadsheetApp.getUi().alert(r ? L('Форма хозяюшек (медикаменты) обновлена!\n\n') + r.short : L('Нет активных товаров.'));
  } catch(e) { SpreadsheetApp.getUi().alert(L('Ошибка: ') + e.toString()); }
}

function updateDeliveryForm(orderItems, weekPeriod) {
  return updateProductsLogistForm(orderItems, weekPeriod);
}

function updateHouseholdFormsAfterOrder(orderItems, weekPeriod) {
  return updateAllHouseholdForms(orderItems, weekPeriod);
}

function swUpdateAllForms(warehouseType, orderItems, weekPeriod) {
  if (warehouseType === 'STATIONERY') return updateAllStationeryForms(orderItems, weekPeriod);
  if (warehouseType === 'MEDICINE')   return updateAllMedicineForms(orderItems, weekPeriod);
  throw new Error(L('Неизвестный тип: ') + warehouseType);
}

function swGetFormLinks(warehouseType) {
  var links = getAllFormLinks();
  if (warehouseType === 'STATIONERY') return links.stationery;
  if (warehouseType === 'MEDICINE')   return links.medicine;
  return {};
}

var HouseholdFormManager = {
  updateAllForms:    function(o, w) { return updateAllHouseholdForms(o, w); },
  updateHostessForm: function(cp)   { return updateHouseholdHostessForm(cp); },
  updateCleaningForm: function()    { return updateHouseholdCleaningForm(); },
  updateLogistForm:  function(o, w) { return updateHouseholdLogistForm(o, w); },
  getFormLinks: function() {
    var l = getAllFormLinks().household;
    return { logist: l.logist, hostess: l.hostess, cleaning: l.cleaning };
  },
  RESPONSE_SHEET_NAMES: {
    logist:   'postavka_household',
    hostess:  'household_hostess',
    cleaning: 'household_cleaning'
  }
};

function _groupBySupplier(orderItems) {
  var grouped = {};
  orderItems.forEach(function(item) {
    var sup = (item.supplier && item.supplier.toString().trim()) ? item.supplier.toString().trim() : L('Без поставщика');
    if (!grouped[sup]) grouped[sup] = [];
    grouped[sup].push(item);
  });
  return grouped;
}

function _getSimpleActiveProducts(warehouseSheet, audience) {
  var data          = warehouseSheet.getDataRange().getValues();
  var headers       = data[1] || [];
  var products      = [];
  var currentCat    = '';
  var hostessColIdx = -1;

  for (var c = 0; c < headers.length; c++) {
    if (headers[c] && hdrRu_(headers[c]) === 'Хозяюшки') { hostessColIdx = c; break; }
  }
  for (var i = 2; i < data.length; i++) {
    var name   = data[i][0] ? data[i][0].toString().trim() : '';
    var qty    = data[i][1], unit = data[i][2], active = data[i][3];
    var isCat  = (qty === '' || qty === null) && (unit === '' || unit === null) &&
                 (active === '' || active === false || active === null);
    if (isCat && name) { currentCat = name; continue; }
    if (!name || active !== true) continue;
    if (audience === 'hostess' && hostessColIdx !== -1 && data[i][hostessColIdx] !== true) continue;
    products.push({ name: name, category: currentCat, unit: unit ? unit.toString().trim() : '' });
  }
  return products;
}

function _getStandardProductsForForm(ss) {
  var products = [];

  var orderSheet = SH_(ss, L('Заказ продуктов')) || SH_(ss, L('Заказ на неделю'));
  if (orderSheet) {
    var data = orderSheet.getDataRange().getValues();
    var inSection = false;
    for (var i = 2; i < data.length; i++) {
      var name = data[i][0] ? data[i][0].toString().trim() : '';
      if (!name) continue;
      if (name.startsWith('📦') || name.startsWith('🗓')) { inSection = name.startsWith('📦'); continue; }
      if (!inSection) continue;
      products.push({ name: name, unit: data[i][2] ? data[i][2].toString().trim() : '' });
    }
  }

  if (products.length === 0) {
    var wSheet = SH_(ss, L('Склад_Продукты'));
    if (wSheet) {
      var config             = loadMenuConfig();
      var rotatingCats       = (config.rotatingCategories || []).map(function(rc) { return rc.category; });
      var excludeCategories  = ['Алкоголь'].concat(rotatingCats.map(K));
      var allCategories      = getAllWarehouseCategories(wSheet);
      var standardCats       = allCategories.filter(function(cat) {
        return excludeCategories.indexOf(K(cat)) === -1;
      });
      standardCats.forEach(function(cat) {
        products = products.concat(ProductMenu.getProductsByCategory(wSheet, cat));
      });
    }
  }
  return products;
}

function _getMenuProductsFromMenuSheet(menuSheet, ss) {
  if (!menuSheet) return [];
  var wSheet  = ss ? SH_(ss, L('Склад_Продукты')) : null;
  var unitMap = wSheet ? buildUnitMap(wSheet) : {};
  var data    = menuSheet.getDataRange().getValues();

  var startRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] && isMenuAssortLabel_(data[i][0])) {
      startRow = i + 1; break;
    }
  }
  if (startRow === -1) return [];

  var endRow = data.length;
  for (var i = startRow; i < data.length; i++) {
    var v = data[i][0] ? data[i][0].toString() : '';
    if (isMenuAlwaysLabel_(v) || isMenuTimeLabel_(v)) { endRow = i; break; }
  }

  var seen = {}, products = [];
  for (var i = startRow; i < endRow; i++) {
    var label = data[i][0] ? data[i][0].toString().trim() : '';
    if (!label || isMenuAssortLabel_(label)) continue;
    for (var col = 1; col < data[i].length; col++) {
      var val = data[i][col] ? data[i][col].toString().trim() : '';
      if (!val || isMenuFiller_(val) || seen[val]) continue;
      seen[val] = true;
      products.push({ name: val, unit: unitMap[val] || '' });
    }
  }
  return products;
}

function _getLatestMenuSheetForForm(ss) {
  var todayMenu = MenuManager.findTodaySheet();
  if (todayMenu) return todayMenu.sheet;
  return SH_(ss, L('Меню'));
}

function _showFormUpdateDialog(result, label) {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:20px;background:#f5f5f5}' +
    '.h{background:#4CAF50;color:white;padding:15px;margin:-20px -20px 20px;text-align:center}' +
    '.box{background:white;padding:12px;border-radius:8px;margin-bottom:10px;border:1px solid #e0e0e0}' +
    '.note{background:#fff8e1;padding:10px;border-radius:6px;font-size:12px;color:#5d4037;margin-bottom:10px}' +
    'a{color:#1a73e8;font-size:12px;word-break:break-all}</style>' +
    L('<div class="h"><h2>Форма обновлена!</h2></div>') +
    L('<div class="note">✅ Ссылка постоянная — не меняется!</div>') +
    '<div class="box"><b>' + label + ':</b><br><a href="' + result.url + '">' + result.short + '</a></div>'
  ).setWidth(480).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, L('Форма обновлена'));
}

function cleanupDeadTriggers() {
  var deadList = [
    'onDeliveryFormSubmit', 'renameDeliverySheetTrigger', 'renameFoodHostessSheet',
    'renameHouseholdSheet_postavka_household', 'renameHouseholdSheet_household_hostess',
    'renameHouseholdSheet_household_cleaning', 'swRenameSheet_stationery_hostess',
    'swRenameSheet_postavka_stationery', 'swRenameSheet_medicine_hostess',
    'swRenameSheet_postavka_medicine', 'onFormSubmitFOOD', 'onHostessFormSubmit'
  ];
  var removed = [];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (deadList.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      removed.push(t.getHandlerFunction());
    }
  });
  SpreadsheetApp.getUi().alert(
    removed.length > 0
      ? L('✅ Удалено триггеров: ') + removed.length + '\n\n' + removed.join('\n')
      : L('✅ Мёртвых триггеров не найдено')
  );
}

function buildUnitMap(wSheet) {
  var map = {};
  var data = wSheet.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    var name = data[i][0] ? data[i][0].toString().trim() : '';
    var unit = data[i][2] ? data[i][2].toString().trim() : '';
    if (name) map[name] = unit;
  }
  return map;
}

function resetAllFormIds() {
  const props = PropertiesService.getScriptProperties();
  const keys = [
    'PRODUCTS_FORM_ID', 'DELIVERY_FORM_ID',
    'HOUSEHOLD_HOSTESS_FORM_ID', 'HOUSEHOLD_CLEANING_FORM_ID', 'HOUSEHOLD_LOGIST_FORM_ID',
    'STATIONERY_HOSTESS_FORM_ID', 'STATIONERY_LOGIST_FORM_ID',
    'MEDICINE_HOSTESS_FORM_ID', 'MEDICINE_LOGIST_FORM_ID'
  ];
  keys.forEach(k => {
    props.deleteProperty(k);
    props.deleteProperty(k + '_URL');
    props.deleteProperty(k + '_SHORT');
  });
  Logger.log('✅ ID форм сброшены. Теперь заново прогоните обновление форм по каждому складу.');
}
