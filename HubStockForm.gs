// HubStockForm.gs

const HUB_STOCK_RESPONSE_SHEET = 'hub_stock_form';

function updateHubStockForm() {
  const officeName   = getOfficeName();
  const allProducts  = getAllHubProducts();

  const form = _getOrCreateForm(HUB_STOCK_FORM_ID_KEY, L('Остатки на складе — форма хозяюшки'));
  form.setTitle((officeName ? officeName + '. ' : '') + L('Остатки на складе — форма хозяюшки'));
  form.setDescription(
    L('Раз в неделю проверьте физический остаток по каждой позиции и впишите число.\n') +
    L('Если товара нет совсем — впишите 0.')
  );
  form.setCollectEmail(false);
  form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true);
  form.setProgressBar(true);
  form.setAcceptingResponses(true);

  form.getItems().reverse().forEach(i => form.deleteItem(i));

  const numberValidation = FormApp.createTextValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setHelpText(L('Введите число (0, если товара нет)'))
    .build();

  let isFirstPage = true;
  Object.keys(WAREHOUSE_CONFIG).forEach(warehouseType => {
    const cfg      = WAREHOUSE_CONFIG[warehouseType];
    const products = allProducts.filter(p => p.warehouseType === warehouseType);
    if (products.length === 0) return;

    if (!isFirstPage) form.addPageBreakItem().setTitle(cfg.icon + ' ' + L(cfg.name));
    isFirstPage = false;

    products.forEach(p => {
      form.addTextItem()
        .setTitle(p.name)
        .setHelpText(L('Текущий остаток, ') + (p.unit || L('шт')))
        .setValidation(numberValidation)
        .setRequired(false);
    });
  });

  if (isFirstPage) throw new Error(L('Нет ни одного активного товара ни на одном складе.'));

  _saveFormId(HUB_STOCK_FORM_ID_KEY, form);
  attachFormToSheet(form, HUB_STOCK_RESPONSE_SHEET, '#6a329f');

  Logger.log('✅ Форма остатков (хаб) обновлена');
  return _formResult(form);
}

function updateHubStockFormManual() {
  try {
    const r = updateHubStockForm();
    SpreadsheetApp.getUi().alert(L('✅ Форма остатков обновлена!\n\n') + r.short);
  } catch(e) {
    SpreadsheetApp.getUi().alert(L('❌ Ошибка: ') + e.toString());
  }
}

function onHubStockFormSubmit(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SH_(ss, HUB_STOCK_RESPONSE_SHEET);
    if (!sheet) return;

    const row      = e && e.range ? e.range.getRow() : sheet.getLastRow();
    const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values   = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

    const historySheet = _getOrCreateHubHistorySheet();
    const historyRows  = [];
    const touched      = {}; // название (в нижнем регистре) → true, чтобы не задвоить ниже
    let changedCount    = 0;

    for (let col = 1; col < headers.length; col++) {
      const productName = headers[col] ? headers[col].toString().trim() : '';
      const rawValue     = values[col];
      if (!productName || rawValue === '' || rawValue === null || rawValue === undefined) continue;

      const newStock = parseFloat(rawValue);
      if (isNaN(newStock)) continue;

      const found = findProductAcrossWarehouses(productName);
      if (!found) { Logger.log('⚠️ Товар из формы не найден на складе: ' + productName); continue; }

      const prevStock = parseFloat(found.sheet.getRange(found.row, 2).getValue()) || 0;
      found.sheet.getRange(found.row, 2).setValue(newStock);

      const consumed = Math.max(0, prevStock - newStock);
      historyRows.push([productName, found.category, prevStock, newStock, consumed, found.unit]);
      touched[productName.toLowerCase()] = true;
      changedCount++;
    }

    getAllHubProducts().forEach(p => {
      if (touched[p.name.toLowerCase()]) return;
      historyRows.push([p.name, p.category, p.stock, p.stock, 0, p.unit]);
    });

    if (historyRows.length > 0) _appendHubHistoryDay(historySheet, todayStr, historyRows);

    SpreadsheetApp.flush();
    Logger.log('✅ onHubStockFormSubmit: обновлено остатков — ' + changedCount + ', позиций в отчёте — ' + historyRows.length);
  } catch(err) {
    Logger.log('❌ onHubStockFormSubmit: ' + err.toString());
  }
}

function _getOrCreateHubHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = SH_(ss, HUB_HISTORY_SHEET);

  if (sheet) {
    const firstHeader = K(sheet.getRange(1, 1).getValue());
    if (firstHeader === 'Дата') {
      // старый формат листа (дата отдельной колонкой на каждой строке) — пересоздаём
      // под новый, сгруппированный по дням вид; данных там пока минимум (лист новый)
      sheet.clear();
      sheet.clearFormats();
      try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).shiftRowGroupDepth(-1); } catch(e) {}
    } else {
      return sheet;
    }
  } else {
    sheet = ss.insertSheet(L(HUB_HISTORY_SHEET));
  }

  sheet.getRange(1, 1, 1, 6).setValues([[
    L('Товар'), L('Категория'), L('Остаток был'), L('Остаток стал'), L('Расход за неделю'), L('Единица')
  ]]).setBackground('#4a148c').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  return sheet;
}

function _appendHubHistoryDay(historySheet, dateLabel, rows) {
  if (!rows || rows.length === 0) return;

  const headerRow = historySheet.getLastRow() + 1;

  historySheet.getRange(headerRow, 1, 1, 6).merge()
    .setValue('📅 ' + dateLabel)
    .setBackground('#7b1fa2')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  historySheet.setRowHeight(headerRow, 26);

  const dataStartRow = headerRow + 1;
  historySheet.getRange(dataStartRow, 1, rows.length, 6).setValues(rows);

  try {
    historySheet.getRange(dataStartRow, 1, rows.length, 1).shiftRowGroupDepth(1);
  } catch(e) {
    Logger.log('⚠️ _appendHubHistoryDay: не удалось сгруппировать строки — ' + e.toString());
  }
}
