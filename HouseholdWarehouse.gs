const HouseholdWarehouse = (function() {
  const SHEET_NAME = 'Склад_Хозтовары';

  function getSheet() {
    const sheet = SH_(SpreadsheetApp.getActiveSpreadsheet(), SHEET_NAME);
    if (!sheet) throw new Error(L('Лист "') + L(SHEET_NAME) + L('" не найден'));
    return sheet;
  }

  function findCol(headers, name) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && hdrRu_(headers[i]) === name) return i;
    }
    return -1;
  }

  function createStructure() {
    const ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
    let sheet = SH_(ss, SHEET_NAME);
    if (sheet) {
      if (ui.alert(L('Пересоздать?'), L('Данные будут удалены.'), ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet(L(SHEET_NAME));
    _buildSheet(sheet);
    updateAllWarehouseDates();
    ui.alert(L('Склад хозтоваров создан!\n\nG — Хозяюшки\nH — Клининг'));
    sheet.activate();
  }

  function _buildSheet(sheet) {
    const today     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const suppliers = getSuppliersFromBaseInfo();
    const categories = [L('Моющие средства'),L('Чистящие средства'),L('Бумажная продукция'),L('Хозяйственные принадлежности'),L('Инвентарь')];
    sheet.setColumnWidth(1,280); sheet.setColumnWidth(2,100); sheet.setColumnWidth(3,130);
    sheet.setColumnWidth(4,90);  sheet.setColumnWidth(5,120); sheet.setColumnWidth(6,160);
    sheet.setColumnWidth(7,90);  sheet.setColumnWidth(8,90);  sheet.setColumnWidth(9,50);
    sheet.getRange(1,1,1,9).merge();
    sheet.getRange('A1').setValue(L('Данные хозтовары склада на ') + today)
      .setBackground('#2196F3').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1,40);
    sheet.getRange(2,1,1,9).setValues([[L('Название продукта'),L('Количество'),L('Единица измерения'),L('Активность'),L('Артикул'),L('Поставщик'),L('Хозяюшки'),L('Клининг'),'']])
      .setBackground('#2196F3').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setRowHeight(2,35);
    const unitRule = getUnitsValidationRule();
    const supRule  = suppliers.length > 0
      ? SpreadsheetApp.newDataValidation().requireValueInList(suppliers,true).setAllowInvalid(true).build()
      : null;
    let row = 3;
    categories.forEach(cat => {
      sheet.getRange(row,1,1,8).merge();
      sheet.getRange(row,1).setValue(cat).setBackground('#d9d9d9').setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
      sheet.getRange(row,9).setValue('➕').setFontSize(14).setHorizontalAlignment('center').setBackground('#4CAF50').setFontColor('#FFFFFF');
      sheet.setRowHeight(row,30); row++;
      sheet.getRange(row,2).setValue(0);
      sheet.getRange(row,3).setValue(L('штука')).setDataValidation(unitRule);
      sheet.getRange(row,4).insertCheckboxes().check();
      sheet.getRange(row,7).insertCheckboxes().check();
      sheet.getRange(row,8).insertCheckboxes().uncheck();
      if (supRule) sheet.getRange(row,6).setDataValidation(supRule);
      sheet.getRange(row,1,1,8)
        .setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment('middle')
        .setFontColor('#000000')
        .setBackground('#FFFFFF');
      sheet.setRowHeight(row,25); row += 2;
    });
    sheet.setFrozenRows(2);
    sheet.getRange(1,1,1,9).protect().setDescription(L('Заголовок')).setWarningOnly(true);
    sheet.getRange(2,1,1,9).protect().setDescription(L('Заголовки')).setWarningOnly(true);
    SpreadsheetApp.flush();
  }

  function getProducts(audience) {
    const sheet      = getSheet();
    const data       = sheet.getDataRange().getValues();
    const headers    = data[1] || [];
    const articulIdx  = findCol(headers, 'Артикул');
    const supplierIdx = findCol(headers, 'Поставщик');
    const hostessIdx  = findCol(headers, 'Хозяюшки');
    const cleaningIdx = findCol(headers, 'Клининг');
    const products = [];
    let currentCategory = '';

    for (let i = 2; i < data.length; i++) {
      const row    = data[i];
      const name   = row[0] ? row[0].toString().trim() : '';
      const qty    = row[1], unit = row[2], active = row[3];
      const isCat  = (qty===''||qty===null) && (unit===''||unit===null) && (active===''||active===false||active===null);
      if (isCat && name) { currentCategory = name; continue; }
      if (!name || active !== true) continue;
      const forH    = hostessIdx  !== -1 ? row[hostessIdx]  === true : true;
      const forC    = cleaningIdx !== -1 ? row[cleaningIdx] === true : false;
      const include = audience === 'all'
        || (audience === 'hostess'  && forH)
        || (audience === 'cleaning' && forC);
      if (!include) continue;
      products.push({
        name, category: currentCategory,
        qty:      parseFloat(qty) || 0,
        unit:     unit     ? unit.toString().trim()               : '',
        articul:  articulIdx  !== -1 ? (row[articulIdx]  || '') : '',
        supplier: supplierIdx !== -1 ? (row[supplierIdx] || '') : '',
        forHostess: forH, forCleaning: forC
      });
    }
    return products;
  }

  function updateStock(productName, delta) {
    const sheet = getSheet(), data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim() !== productName) continue;
      const newQty = Math.max(0, (parseFloat(data[i][1]) || 0) + delta);
      sheet.getRange(i + 1, 2).setValue(newQty);
      return newQty;
    }
    return null;
  }

  return { createStructure, getProducts, updateStock, SHEET_NAME };
})();

const HH_SHEETS = {
  hostess:  'household_hostess',
  cleaning: 'household_cleaning',
  logist:   'postavka_household'
};

const HouseholdDashboard = (function() {
  const DASH_SHEET = 'Хозтовары';
  const COLORS = {
    header:   '#2196F3',
    total:    '#4a86e8',
    issued:   '#34A853',
    delivery: '#FF6D00',
    stock:    '#6a329f',
    cleaning: '#00ACC1',
    rowOdd:   '#f8f9fa',
    rowEven:  '#ffffff'
  };

  function getOrCreateSheet(ss) {
    let s = SH_(ss, DASH_SHEET);
    if (!s) s = ss.insertSheet(L(DASH_SHEET));
    return s;
  }

  function refresh() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const ss           = SpreadsheetApp.getActiveSpreadsheet();
      const dash         = getOrCreateSheet(ss);
      const coffeePoints = BaseInfo.getActiveCoffeePoints();
      const today        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

      const hostessData  = _readResponseSheet(ss, HH_SHEETS.hostess,  today, true);
      const cleaningData = _readResponseSheet(ss, HH_SHEETS.cleaning, today, false);
      const deliveryData = _readDeliverySheet(ss, today);
      const isFreshToday = _isDashboardForToday(dash, today);
      const issuedMap     = isFreshToday ? _readIssuedFromDash(dash) : {};
      const stockMap      = _readStockFromWarehouse();

      _buildDashboard(dash, coffeePoints, hostessData, cleaningData, deliveryData, issuedMap, stockMap, today);
      Logger.log('✅ Дашборд Хозтовары обновлён');
    } finally {
      lock.releaseLock();
    }
  }

  function _readResponseSheet(ss, sheetName, todayStr, isHostess) {
    const result = {}, sheet = SH_(ss, sheetName);
    if (!sheet) return result;
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      const ts = data[i][0]; if (!ts) continue;
      let rowDate;
      try { rowDate = Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy'); } catch(e) { continue; }
      if (rowDate !== todayStr) continue;
      const gn = data[i][1] ? data[i][1].toString().trim() : (isHostess ? L('КП') : L('Клининг'));
      if (!result[gn]) result[gn] = {};
      for (let col = 2; col < headers.length; col++) {
        const p = headers[col] ? headers[col].toString().trim() : '';
        const q = data[i][col];
        if (!p || q === '' || q === null) continue;
        const n = parseFloat(q);
        if (!isNaN(n) && n > 0) result[gn][p] = (result[gn][p] || 0) + n;
      }
    }
    return result;
  }

  function _readDeliverySheet(ss, todayStr) {
    const result = {}, sheet = SH_(ss, HH_SHEETS.logist);
    if (!sheet) return result;
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      const ts = data[i][0]; if (!ts) continue;
      let d;
      try { d = Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy'); } catch(e) { continue; }
      if (d !== todayStr) continue;
      for (let col = 1; col < headers.length; col++) {
        const p = headers[col] ? headers[col].toString().trim() : '';
        const q = data[i][col];
        if (!p || q === '' || q === null) continue;
        const n = parseFloat(q);
        if (!isNaN(n) && n > 0) result[p] = (result[p] || 0) + n;
      }
    }
    return result;
  }

  function _readIssuedFromDash(dash) {
    const result = {}, data = dash.getDataRange().getValues();
    if (data.length < 2) return result;
    const headers = data[1];
    let idx = -1;
    for (let c = 0; c < headers.length; c++) {
      if (headers[c] && K(headers[c]).toUpperCase() === 'ВЫДАНО') { idx = c; break; }
    }
    if (idx === -1) return result;
    for (let i = 2; i < data.length; i++) {
      const name = data[i][0] ? data[i][0].toString().trim() : '';
      const val  = data[i][idx];
      if (name && val !== '' && val !== null) {
        const n = parseFloat(val);
        if (!isNaN(n) && n > 0) result[name] = n;
      }
    }
    return result;
  }

  function _readStockFromWarehouse() {
    const result = {};
    try {
      const sheet = SH_(SpreadsheetApp.getActiveSpreadsheet(), HouseholdWarehouse.SHEET_NAME);
      if (!sheet) return result;
      const data = sheet.getDataRange().getValues();
      for (let i = 2; i < data.length; i++) {
        const name = data[i][0] ? data[i][0].toString().trim() : '';
        const qty  = data[i][1];
        if (name && qty !== '' && qty !== null) {
          const n = parseFloat(qty);
          if (!isNaN(n)) result[name] = n;
        }
      }
    } catch(e) { Logger.log('⚠️ _readStockHH: ' + e); }
    return result;
  }

  function _buildDashboard(dash, coffeePoints, hostessData, cleaningData, deliveryData, issuedMap, stockMap, todayStr) {
    dash.clear();
    const cpCount   = coffeePoints.length;
    const totalCols = 1 + cpCount + 1 + 4;

    dash.getRange(1,1,1,totalCols).merge();
    dash.getRange(1,1).setValue(todayStr).setNumberFormat('@STRING@')
      .setBackground(COLORS.header).setFontColor('#ffffff').setFontWeight('bold').setFontSize(13)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.setRowHeight(1,32);

    const hv = [L('Товар')];
    coffeePoints.forEach(cp => hv.push(cp.name));
    hv.push(L('Клининг'),L('ИТОГО'),L('ВЫДАНО'),L('Поставка'),L('Остаток на складе'));
    dash.getRange(2,1,1,totalCols).setValues([hv])
      .setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setBackground(COLORS.header);
    dash.getRange(2, 2 + cpCount).setBackground(COLORS.cleaning);
    dash.getRange(2, 3 + cpCount).setBackground(COLORS.total);
    dash.getRange(2, 4 + cpCount).setBackground(COLORS.issued);
    dash.getRange(2, 5 + cpCount).setBackground(COLORS.delivery);
    dash.getRange(2, 6 + cpCount).setBackground(COLORS.stock);
    dash.setRowHeight(2,30);

    const allP = new Set();
    Object.values(hostessData).forEach(d  => Object.keys(d).forEach(p => allP.add(p)));
    Object.values(cleaningData).forEach(d => Object.keys(d).forEach(p => allP.add(p)));
    Object.keys(deliveryData).forEach(p   => allP.add(p));
    const products = [...allP].sort();

    if (products.length === 0) {
      dash.getRange(3,1,1,totalCols).merge();
      dash.getRange(3,1).setValue(L('Заказов за сегодня ещё нет'))
        .setFontColor('#999').setFontStyle('italic').setHorizontalAlignment('center');
    } else {
      const rd = products.map(product => {
        const row = [product];
        let total = 0;
        coffeePoints.forEach(cp => {
          const q = (hostessData[cp.name] || {})[product];
          row.push(q !== undefined ? q : '');
          total += parseFloat(q) || 0;
        });
        let cq = 0;
        Object.values(cleaningData).forEach(d => { cq += parseFloat(d[product]) || 0; });
        row.push(cq > 0 ? cq : '');
        total += cq;
        row.push(total > 0 ? total : '');
        row.push(issuedMap[product]   !== undefined ? issuedMap[product]   : '');
        row.push(deliveryData[product] !== undefined ? deliveryData[product] : '');
        row.push(stockMap[product]    !== undefined ? stockMap[product]    : '');
        return row;
      });

      dash.getRange(3, 1, rd.length, totalCols).setValues(rd);
      dash.getRange(3, 1, rd.length, 1)
        .setWrap(true).setVerticalAlignment('middle').setFontWeight('bold');

      rd.forEach((row, idx) => {
        const r  = 3 + idx;
        const bg = idx % 2 === 0 ? COLORS.rowEven : COLORS.rowOdd;
        dash.getRange(r,1,1,totalCols).setBackground(bg)
          .setBorder(true,true,true,true,true,true,'#e0e0e0',SpreadsheetApp.BorderStyle.SOLID);
        dash.getRange(r, 2, 1, cpCount + 1).setHorizontalAlignment('center');
        if (row[1+cpCount+1] !== '') dash.getRange(r, 3+cpCount).setFontColor('#1c4587').setFontWeight('bold');
        if (row[1+cpCount+2] !== '') dash.getRange(r, 4+cpCount).setFontColor('#0a7a0a').setFontWeight('bold');
        if (row[1+cpCount+3] !== '') dash.getRange(r, 5+cpCount).setFontColor('#e65100').setFontWeight('bold');
        if (row[1+cpCount+4] !== '') dash.getRange(r, 6+cpCount).setFontColor('#4a148c');
        [3+cpCount, 4+cpCount, 5+cpCount, 6+cpCount].forEach(c => dash.getRange(r, c).setHorizontalAlignment('center'));
      });
    }

    dash.setColumnWidth(1, 320);
    for (let c = 2; c <= 2 + cpCount; c++) dash.setColumnWidth(c, 80);
    dash.setColumnWidth(3+cpCount, 70);
    dash.setColumnWidth(4+cpCount, 80);
    dash.setColumnWidth(5+cpCount, 90);
    dash.setColumnWidth(6+cpCount, 130);
    dash.setFrozenRows(2);
  }

  return { refresh, DASH_SHEET };
})();
