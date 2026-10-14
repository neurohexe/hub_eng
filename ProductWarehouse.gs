// ProductWarehouse.gs

const ProductWarehouse = (function() {

  function findHeaderRowIndex(data) {
    for (let i = 0; i < Math.min(6, data.length); i++) {
      for (let j = 0; j < data[i].length; j++) {
        const val = data[i][j] ? K(data[i][j]).toUpperCase() : '';
        if (val === 'ИТОГО') return i;
      }
    }
    return 1;
  }

  function findColIndex(headers, name) {
    const nameLower = name.trim().toLowerCase();
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && hdrRu_(headers[i]).toLowerCase() === nameLower) return i;
    }
    return -1;
  }

  function buildRowMap(data, startIdx, nameColIdx) {
    const map = {};
    for (let i = startIdx; i < data.length; i++) {
      const name = data[i][nameColIdx] ? data[i][nameColIdx].toString().trim() : '';
      if (name) map[name] = i;
    }
    return map;
  }

  function startOfDay() {
    try {
      const ss             = SpreadsheetApp.getActiveSpreadsheet();
      const productsSheet  = SH_(ss, L('Продукты'));
      const warehouseSheet = SH_(ss, L('Склад_Продукты'));

      if (!productsSheet)  return { success: false, message: L('Лист "Продукты" не найден') };
      if (!warehouseSheet) return { success: false, message: L('Лист "Склад_Продукты" не найден') };

      const prodData     = productsSheet.getDataRange().getValues();
      const headerRowIdx = findHeaderRowIndex(prodData);
      const dataStartIdx = headerRowIdx + 1;
      const prodHeaders  = prodData[headerRowIdx];
      const ostatokIdx   = findColIndex(prodHeaders, 'Остаток на складе');

      if (ostatokIdx === -1) return { success: false, message: L('Колонка "Остаток на складе" не найдена') };

      const prodRowMap = buildRowMap(prodData, dataStartIdx, 0);
      const whData     = warehouseSheet.getDataRange().getValues();
      const whHeaders  = whData[1];
      const kolIdx     = findColIndex(whHeaders, 'Количество');

      if (kolIdx === -1) return { success: false, message: L('Колонка "Количество" не найдена в Склад_Продукты') };

      let updated = 0;
      for (let i = 2; i < whData.length; i++) {
        const name = whData[i][0] ? whData[i][0].toString().trim() : '';
        if (!name) continue;
        const qty = whData[i][kolIdx];
        if (qty === '' || qty === null || qty === undefined) continue;
        if (prodRowMap[name] !== undefined) {
          productsSheet.getRange(prodRowMap[name] + 1, ostatokIdx + 1).setValue(Number(qty) || 0);
          updated++;
        }
      }

      return { success: true, message: L('Начало дня: перенесено ') + updated + L(' остатков из склада') };

    } catch(error) {
      Logger.log('❌ startOfDay: ' + error.toString());
      return { success: false, message: L('Ошибка: ') + error.toString() };
    }
  }

  function getProductsToIssue(warehouseType) {
    try {
      const cfg = WAREHOUSE_CONFIG[warehouseType];
      if (!cfg) return { products: [] };

      const ss         = SpreadsheetApp.getActiveSpreadsheet();
      const dashSheet  = SH_(ss, cfg.dashSheet);
      if (!dashSheet) return { products: [] };

      const data = dashSheet.getDataRange().getValues();
      if (data.length < 3) return { products: [] };

      const headerRowIdx = findHeaderRowIndex(data);
      const dataStartIdx = headerRowIdx + 1;
      const headers      = data[headerRowIdx];
      const itogoIdx     = findColIndex(headers, 'ИТОГО');
      const vydanoIdx    = findColIndex(headers, 'ВЫДАНО');

      if (itogoIdx === -1) return { products: [] };

      const products = [];
      for (let i = dataStartIdx; i < data.length; i++) {
        const name   = data[i][0] ? data[i][0].toString().trim() : '';
        const itogo  = data[i][itogoIdx];
        const vydano = vydanoIdx !== -1 ? data[i][vydanoIdx] : 0;
        if (!name || !itogo || Number(itogo) <= 0) continue;
        products.push({
          name,
          ordered: Number(itogo)  || 0,
          issued:  Number(vydano) || 0
        });
      }

      return { products };

    } catch(error) {
      Logger.log('❌ getProductsToIssue: ' + error.toString());
      return { products: [] };
    }
  }

  function issueProducts(warehouseType, formData) {
    try {
      const cfg = WAREHOUSE_CONFIG[warehouseType];
      if (!cfg) return { success: false, message: L('Неизвестный тип склада: ') + warehouseType };

      const ss             = SpreadsheetApp.getActiveSpreadsheet();
      const dashSheet      = SH_(ss, cfg.dashSheet);
      const warehouseSheet = SH_(ss, cfg.warehouseSheet);

      if (!dashSheet)      return { success: false, message: L('Лист "') + L(cfg.dashSheet) + L('" не найден') };
      if (!warehouseSheet) return { success: false, message: L('Лист "') + L(cfg.warehouseSheet) + L('" не найден') };

      const dashData     = dashSheet.getDataRange().getValues();
      const headerRowIdx = findHeaderRowIndex(dashData);
      const dataStartIdx = headerRowIdx + 1;
      const dashHeaders  = dashData[headerRowIdx];

      const vydanoIdx  = findColIndex(dashHeaders, 'ВЫДАНО');
      const ostatokIdx = findColIndex(dashHeaders, 'Остаток на складе');

      if (vydanoIdx  === -1) return { success: false, message: L('Колонка "ВЫДАНО" не найдена в ') + cfg.dashSheet };
      if (ostatokIdx === -1) return { success: false, message: L('Колонка "Остаток на складе" не найдена в ') + cfg.dashSheet };

      const dashRowMap = buildRowMap(dashData, dataStartIdx, 0);

      const whData   = warehouseSheet.getDataRange().getValues();
      const whRowMap = {};
      for (let i = 2; i < whData.length; i++) {
        const name = whData[i][0] ? whData[i][0].toString().trim() : '';
        if (name) whRowMap[name] = i;
      }

      let updated = 0;

      for (const [name, issuedQty] of Object.entries(formData)) {
        const issued = parseFloat(issuedQty) || 0;
        if (issued <= 0) continue;

        const dashRowIdx = dashRowMap[name];
        if (dashRowIdx !== undefined) {
          const sheetRow = dashRowIdx + 1;

          dashSheet.getRange(sheetRow, vydanoIdx + 1).setValue(issued);

          const ostatokCell    = dashSheet.getRange(sheetRow, ostatokIdx + 1);
          const currentOstatok = parseFloat(ostatokCell.getValue()) || 0;
          const newOstatok     = Math.max(0, currentOstatok - issued);
          ostatokCell.setValue(newOstatok);

          Logger.log('✓ Выдача "' + name + '": ' + currentOstatok + ' - ' + issued + ' = ' + newOstatok);
        }

        const whRowIdx = whRowMap[name];
        if (whRowIdx !== undefined) {
          const currentWhQty = parseFloat(whData[whRowIdx][1]) || 0;
          const newWhQty     = Math.max(0, currentWhQty - issued);
          warehouseSheet.getRange(whRowIdx + 1, 2).setValue(newWhQty);
        }

        updated++;
      }

      SpreadsheetApp.flush();

      try {
        switch(warehouseType) {
          case 'PRODUCTS':   refreshProductsDashboard();       break;
          case 'HOUSEHOLD':  HouseholdDashboard.refresh();     break;
          case 'STATIONERY': swRefreshDashboard('STATIONERY'); break;
          case 'MEDICINE':   swRefreshDashboard('MEDICINE');   break;
        }
      } catch(e) {
        Logger.log('⚠️ refresh после выдачи: ' + e.toString());
      }

      return { success: true, message: L('Выдача подтверждена: ') + updated + L(' позиций') };

    } catch(error) {
      Logger.log('❌ issueProducts: ' + error.toString());
      return { success: false, message: L('Ошибка: ') + error.toString() };
    }
  }

  function saveDailyData(warehouseType) {
    try {
      if (warehouseType !== 'PRODUCTS') {
        return { success: true, message: L('Для этого склада сохранение не настроено') };
      }

      Logger.log('⚠️ ProductWarehouse.saveDailyData: функция устарела и не вызывается ' +
                 'автоматически (см. комментарий в коде). Ручной вызов ничего не делает.');
      return {
        success: true,
        message: L('Функция устарела: дашборд Продуктов теперь всегда синхронен со ') +
                 L('складом, отдельный перенос остатков в конец дня не требуется.')
      };

    } catch(error) {
      Logger.log('❌ saveDailyData: ' + error.toString());
      return { success: false, message: L('Ошибка: ') + error.toString() };
    }
  }

  function debugProductsSheet() {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const msg = [];

    const ps = SH_(ss, L('Продукты'));
    if (!ps) { SpreadsheetApp.getUi().alert(L('Лист "Продукты" не найден!')); return; }

    const pd           = ps.getDataRange().getValues();
    const headerRowIdx = findHeaderRowIndex(pd);
    const headers      = pd[headerRowIdx] || [];

    msg.push(L('=== Лист "Продукты" ==='));
    msg.push(L('Строк: ')             + pd.length);
    msg.push(L('Строка заголовков: ') + (headerRowIdx + 1));
    msg.push(L('ИТОГО:             col ') + (findColIndex(headers, 'ИТОГО')             + 1));
    msg.push(L('ВЫДАНО:            col ') + (findColIndex(headers, 'ВЫДАНО')            + 1));
    msg.push(L('Поставка:          col ') + (findColIndex(headers, 'Поставка')          + 1));
    msg.push(L('Остаток на складе: col ') + (findColIndex(headers, 'Остаток на складе') + 1));

    const ws = SH_(ss, L('Склад_Продукты'));
    msg.push('');
    msg.push(L('=== Лист "Склад_Продукты" ==='));
    if (!ws) {
      msg.push(L('НЕ НАЙДЕН'));
    } else {
      const wd = ws.getDataRange().getValues();
      const wh = wd[1] || [];
      msg.push(L('Количество: col ') + (findColIndex(wh, 'Количество') + 1));
    }

    SpreadsheetApp.getUi().alert(msg.join('\n'));
  }

  return {
    startOfDay,
    getProductsToIssue,
    issueProducts,
    saveDailyData,
    debugProductsSheet
  };

})();

function debugProductsSheetWrapper() { ProductWarehouse.debugProductsSheet(); }

function startOfDayProducts() {
  const result = ProductWarehouse.startOfDay();
  Logger.log('startOfDayProducts: ' + result.message);
}

function startOfDayProductsManual() {
  const result = ProductWarehouse.startOfDay();
  SpreadsheetApp.getUi().alert(result.message);
}
