function swRefreshDashboard(warehouseType) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cfg = getWarehouseConfigByType(warehouseType);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const coffeePoints = BaseInfo.getActiveCoffeePoints();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    let dash = SH_(ss, cfg.dashSheet);
    if (!dash) dash = ss.insertSheet(L(cfg.dashSheet));
    const hostessData  = _swReadResponseSheet(ss, cfg.hostessSheet, today);
    const deliveryData = _swReadDeliverySheet(ss, cfg.logistSheet, today);
    const isFreshToday = _isDashboardForToday(dash, today);
    const issuedMap     = isFreshToday ? _swReadIssuedFromDash(dash) : {};
    const stockMap      = _swReadStockFromWarehouse(ss, cfg.warehouseSheet);
    _swBuildDashboard(dash, cfg, coffeePoints, hostessData, deliveryData, issuedMap, stockMap, today);
    Logger.log('✅ Дашборд "' + cfg.dashSheet + '" обновлён');
  } finally {
    lock.releaseLock();
  }
}

function _swReadResponseSheet(ss, sheetName, todayStr) {
  const result = {}, sheet = SH_(ss, sheetName);
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues(), headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const ts = data[i][0]; if (!ts) continue;
    let d;
    try { d = Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy'); } catch(e) { continue; }
    if (d !== todayStr) continue;
    const cpName = data[i][1] ? data[i][1].toString().trim() : L('КП');
    if (!result[cpName]) result[cpName] = {};
    for (let col = 2; col < headers.length; col++) {
      const p = headers[col] ? headers[col].toString().trim() : '';
      const q = data[i][col];
      if (!p || q === '' || q === null) continue;
      const n = parseFloat(q);
      if (!isNaN(n) && n > 0) result[cpName][p] = (result[cpName][p] || 0) + n;
    }
  }
  return result;
}

function _swReadDeliverySheet(ss, sheetName, todayStr) {
  const result = {}, sheet = SH_(ss, sheetName);
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues(), headers = data[0];
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

function _swReadIssuedFromDash(dash) {
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

function _swReadStockFromWarehouse(ss, sheetName) {
  const result = {}, sheet = SH_(ss, sheetName);
  if (!sheet) return result;
  const data = sheet.getDataRange().getValues();
  for (let i = 2; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    const qty  = data[i][1];
    if (!name || qty === '' || qty === null) continue;
    const n = parseFloat(qty);
    if (!isNaN(n)) result[name] = n;
  }
  return result;
}

function _swBuildDashboard(dash, cfg, coffeePoints, hostessData, deliveryData, issuedMap, stockMap, todayStr) {
  dash.clear();
  const cpCount   = coffeePoints.length;
  const totalCols = 1 + cpCount + 4;
  const C = {
    header:  cfg.color,
    total:   '#4a86e8',
    issued:  '#34A853',
    delivery:'#FF6D00',
    stock:   '#6a329f',
    rowOdd:  '#f8f9fa',
    rowEven: '#ffffff'
  };

  dash.getRange(1, 1, 1, totalCols).merge();
  dash.getRange(1, 1)
    .setValue(todayStr).setNumberFormat('@STRING@')
    .setBackground(C.header).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(1, 32);

  const hv = [L('Товар')];
  coffeePoints.forEach(cp => hv.push(cp.name));
  hv.push(L('ИТОГО'), L('ВЫДАНО'), L('Поставка'), L('Остаток на складе'));
  dash.getRange(2, 1, 1, totalCols)
    .setValues([hv])
    .setFontWeight('bold').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setBackground(C.header);
  dash.getRange(2, 2 + cpCount).setBackground(C.total);
  dash.getRange(2, 3 + cpCount).setBackground(C.issued);
  dash.getRange(2, 4 + cpCount).setBackground(C.delivery);
  dash.getRange(2, 5 + cpCount).setBackground(C.stock);
  dash.setRowHeight(2, 30);

  const allP = new Set();
  Object.values(hostessData).forEach(d => Object.keys(d).forEach(p => allP.add(p)));
  Object.keys(deliveryData).forEach(p => allP.add(p));
  const products = [...allP].sort();

  if (products.length === 0) {
    dash.getRange(3, 1, 1, totalCols).merge();
    dash.getRange(3, 1)
      .setValue(L('Заказов за сегодня ещё нет'))
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
      const bg = idx % 2 === 0 ? C.rowEven : C.rowOdd;
      dash.getRange(r, 1, 1, totalCols)
        .setBackground(bg)
        .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
      dash.getRange(r, 2, 1, cpCount).setHorizontalAlignment('center');
      if (row[1 + cpCount] !== '') dash.getRange(r, 2 + cpCount).setFontColor('#1c4587').setFontWeight('bold').setHorizontalAlignment('center');
      if (row[2 + cpCount] !== '') dash.getRange(r, 3 + cpCount).setFontColor('#0a7a0a').setFontWeight('bold').setHorizontalAlignment('center');
      if (row[3 + cpCount] !== '') dash.getRange(r, 4 + cpCount).setFontColor('#e65100').setFontWeight('bold').setHorizontalAlignment('center');
      if (row[4 + cpCount] !== '') dash.getRange(r, 5 + cpCount).setFontColor('#4a148c').setHorizontalAlignment('center');
    });
  }

  dash.setColumnWidth(1, 320);
  for (let c = 2; c <= 1 + cpCount; c++) dash.setColumnWidth(c, 80);
  dash.setColumnWidth(2 + cpCount, 70);
  dash.setColumnWidth(3 + cpCount, 80);
  dash.setColumnWidth(4 + cpCount, 90);
  dash.setColumnWidth(5 + cpCount, 130);
  dash.setFrozenRows(2);
}
