const MenuManager = (function() {

  function findMenuSheets() {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const result = [];
    sheets.forEach(function(sheet) {
      const name = sheet.getName();
      const m = name.match(/^(?:Меню|Menu)\s+(\d{2}\.\d{2})-(\d{2}\.\d{2})$/);
      if (!m) return;
      const monday = _parseDateDDMM(m[1]);
      const friday = _parseDateDDMM(m[2]);
      if (!monday || !friday) return;
      result.push({ sheet, monday, friday, name });
    });
    result.sort(function(a, b) { return a.monday - b.monday; });
    return result;
  }

  function _parseDateDDMM(str) {
    try {
      const parts = str.split('.');
      const day   = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      const now   = new Date();
      let year    = now.getFullYear();
      if (month < now.getMonth() - 1) year++;
      const d = new Date(year, month, day);
      return isNaN(d.getTime()) ? null : d;
    } catch(e) { return null; }
  }

  function findTodaySheet() {
    const menus = findMenuSheets();
    if (menus.length === 0) return null;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    for (const menu of menus) {
      const mon = new Date(menu.monday); mon.setHours(0, 0, 0, 0);
      const fri = new Date(menu.friday); fri.setHours(23, 59, 59, 999);
      if (today >= mon && today <= fri) return menu;
    }
    if (today < menus[0].monday) {
      for (let i = menus.length - 1; i >= 0; i--) {
        if (menus[i].monday <= today) return menus[i];
      }
      return menus[0];
    }
    return menus[menus.length - 1];
  }

  function _getProductsFromWarehouse(categoryName) {
    const ss             = SpreadsheetApp.getActiveSpreadsheet();
    const warehouseSheet = SH_(ss, L('Склад_Продукты'));
    if (!warehouseSheet) return [L('Остатки'), L('Отсутствует')];
    const data     = warehouseSheet.getDataRange().getValues();
    const products = [];
    let   inCat    = false;
    for (let i = 2; i < data.length; i++) {
      const name    = data[i][0] ? data[i][0].toString().trim() : '';
      const active  = data[i][3];
      const hasPlus = data[i][6] && data[i][6].toString().includes('➕');
      if (hasPlus && name) {
        if (name === categoryName) inCat = true;
        else if (inCat) break;
        continue;
      }
      if (inCat && name && active === true) products.push(name);
    }
    return [...products, L('Остатки'), L('Отсутствует')];
  }

  function _setDropdown(sheet, row, col, categoryName) {
    const values = _getProductsFromWarehouse(categoryName);
    if (values.length === 0) return;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(row, col).setDataValidation(rule);
  }

  function _calcMaxPositions(config) {
    let max = 1;

    (config.timings || []).forEach(timing => {
      let total = 0;
      (timing.items || []).forEach(item => {
        if (item.type === 'category') {
          const rc = (config.rotatingCategories || []).find(r => r.category === item.value);
          if (rc && (rc.mode === 'daily_multi' || rc.mode === 'weekly_multi')) {
            total += (rc.countPerDay || 1);
          } else {
            total += 1;
          }
        } else {
          total += 1;
        }
      });
      if (total > max) max = total;
    });

    (config.rotatingCategories || []).forEach(rc => {
      if (rc.mode === 'daily_multi' || rc.mode === 'weekly_multi') {
        const count = rc.countPerDay || 1;
        if (count > max) max = count;
      }
    });

    return max;
  }

  function _firstAvailableDayRef(refs) {
    if (!refs) return null;
    const dayKeys = Object.keys(refs)
      .filter(k => !isNaN(Number(k)))
      .map(Number)
      .sort((a, b) => a - b);
    if (dayKeys.length === 0) return null;
    return refs[dayKeys[0]];
  }

  function createNextWeekMenu() {
    try {
      const config = loadMenuConfig();
      if (!config || (!config.rotatingCategories.length && !config.alwaysOn.length &&
          (!config.timings || !config.timings.length))) {
        SpreadsheetApp.getUi().alert(L('⚠️ Конструктор меню не настроен.'));
        return;
      }
      const sheet = buildMenu(config, true);
      if (sheet) saveMenuDataToStorage(sheet, true);
      SpreadsheetApp.getUi().alert(
        L('✅ Меню на следующую неделю создано!\n\n') +
        L('• Выберите ассортимент в новом листе\n') +
        L('• После выбора нажмите: Меню → Сохранить меню для планшетов\n') +
        L('• Старое меню остаётся до конца текущей недели')
      );
    } catch(error) {
      SpreadsheetApp.getUi().alert(L('❌ Ошибка: ') + error.toString());
    }
  }

  function createWeeklyMenu() {
    try {
      const config = loadMenuConfig();
      if (!config || (!config.rotatingCategories.length && !config.alwaysOn.length &&
          (!config.timings || !config.timings.length))) {
        SpreadsheetApp.getUi().alert(L('⚠️ Конструктор меню не настроен.'));
        return;
      }
      const sheet = buildMenu(config, false);
      if (sheet) saveMenuDataToStorage(sheet, false);
      SpreadsheetApp.getUi().alert(L('✅ Меню на текущую неделю создано!'));
    } catch(error) {
      SpreadsheetApp.getUi().alert(L('❌ Ошибка: ') + error.toString());
    }
  }

  function buildMenu(config, nextWeek) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!SH_(ss, L('Склад_Продукты'))) {
      throw new Error(L('Лист Склад_Продукты не найден'));
    }

    const today  = new Date();
    let monday   = _getMonday(today);
    if (nextWeek) monday.setDate(monday.getDate() + 7);
    const friday = new Date(monday);
    friday.setDate(friday.getDate() + 4);

    const mondayStr = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'dd.MM');
    const fridayStr = Utilities.formatDate(friday, Session.getScriptTimeZone(), 'dd.MM');
    const sheetName = L('Меню ') + mondayStr + '-' + fridayStr;

    const existing = SH_(ss, sheetName);
    if (existing) {
      const resp = SpreadsheetApp.getUi().alert(
        L('Лист "') + sheetName + L('" уже существует'),
        L('Пересоздать? Выбранный ассортимент будет сброшен.'),
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (resp !== SpreadsheetApp.getUi().Button.YES) return null;
      ss.deleteSheet(existing);
    }

    const menuSheet = ss.insertSheet(L(sheetName));

    const maxPositions = _calcMaxPositions(config);
    const totalCols    = 1 + maxPositions;

    menuSheet.setColumnWidth(1, 180);
    for (let i = 2; i <= totalCols; i++) menuSheet.setColumnWidth(i, 220);
    if (totalCols < 10) {
      try { menuSheet.hideColumns(totalCols + 1, 20); } catch(e) {}
    }

    menuSheet.setFrozenRows(1);

    let row = 1;

    menuSheet.getRange(row, 1, 1, totalCols).merge();
    menuSheet.getRange(row, 1)
      .setValue(L('Меню на неделю с ') + mondayStr + L(' по ') + fridayStr)
      .setHorizontalAlignment('center')
      .setFontSize(16).setFontWeight('bold')
      .setBackground('#1c4587').setFontColor('white')
      .setWrap(true);
    menuSheet.setRowHeight(row, 40);
    row++;

    const assortmentRefs = {};

    if (config.rotatingCategories.length) {
      menuSheet.getRange(row, 1, 1, totalCols).merge();
      menuSheet.getRange(row, 1)
        .setValue(L('ВЫБОР АССОРТИМЕНТА НА НЕДЕЛЮ:'))
        .setFontWeight('bold').setBackground('#FFE599')
        .setFontSize(12);
      menuSheet.setRowHeight(row, 30);
      row++;

      for (const rc of config.rotatingCategories) {
        if (!rc.category) continue;
        const days        = (rc.days || [0,1,2,3,4]).slice().sort((a, b) => a - b);
        const mode        = rc.mode || 'weekly';
        const countPerDay = rc.countPerDay || 1;
        if (!assortmentRefs[rc.category]) assortmentRefs[rc.category] = {};

        if (mode === 'weekly') {
          menuSheet.getRange(row, 1).setValue(rc.category + L(' на неделю:')).setFontWeight('bold');
          menuSheet.getRange(row, 2, 1, totalCols - 1).merge();
          _setDropdown(menuSheet, row, 2, rc.category);
          assortmentRefs[rc.category].weekly = row;
          menuSheet.getRange(row, 1, 1, totalCols).setBackground('#FAFAFA').setWrap(true);
          menuSheet.setRowHeight(row, 25);
          row++;

        } else if (mode === 'weekly_multi') {
          for (let s = 0; s < countPerDay; s++) {
            menuSheet.getRange(row, 1).setValue(rc.category + L(' вид ') + (s + 1) + ':').setFontWeight('bold');
            menuSheet.getRange(row, 2, 1, totalCols - 1).merge();
            _setDropdown(menuSheet, row, 2, rc.category);
            if (!assortmentRefs[rc.category].weeklyMulti) assortmentRefs[rc.category].weeklyMulti = [];
            assortmentRefs[rc.category].weeklyMulti.push(row);
            menuSheet.getRange(row, 1, 1, totalCols).setBackground('#FAFAFA').setWrap(true);
            menuSheet.setRowHeight(row, 25);
            row++;
          }

        } else if (mode === 'daily_same') {
          for (const dayIdx of days) {
            const date    = new Date(monday);
            date.setDate(date.getDate() + dayIdx);
            const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM');
            const dayName = [L('Пн'),L('Вт'),L('Ср'),L('Чт'),L('Пт'),L('Сб'),L('Вс')][dayIdx];
            menuSheet.getRange(row, 1).setValue(rc.category + ' ' + dayName + ' (' + dateStr + '):').setFontWeight('bold');
            menuSheet.getRange(row, 2, 1, totalCols - 1).merge();
            _setDropdown(menuSheet, row, 2, rc.category);
            assortmentRefs[rc.category][dayIdx] = { row: row, cols: 1 };
            menuSheet.getRange(row, 1, 1, totalCols).setBackground('#FAFAFA').setWrap(true);
            menuSheet.setRowHeight(row, 25);
            row++;
          }

        } else if (mode === 'daily_multi') {
          for (const dayIdx of days) {
            const date    = new Date(monday);
            date.setDate(date.getDate() + dayIdx);
            const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM');
            const dayName = [L('Пн'),L('Вт'),L('Ср'),L('Чт'),L('Пт'),L('Сб'),L('Вс')][dayIdx];
            menuSheet.getRange(row, 1).setValue(rc.category + ' ' + dayName + ' (' + dateStr + '):').setFontWeight('bold');
            for (let s = 0; s < Math.min(countPerDay, maxPositions); s++) {
              _setDropdown(menuSheet, row, 2 + s, rc.category);
            }
            assortmentRefs[rc.category][dayIdx] = { row: row, cols: countPerDay };
            menuSheet.getRange(row, 1, 1, totalCols).setBackground('#FAFAFA').setWrap(true);
            menuSheet.setRowHeight(row, 25);
            row++;
          }
        }
        row++;
      }
      row++;
    }

    if (config.alwaysOn && config.alwaysOn.length) {
      menuSheet.getRange(row, 1, 1, totalCols).merge();
      menuSheet.getRange(row, 1)
        .setValue(L('Круглосуточно стоят:'))
        .setFontWeight('bold').setFontSize(12)
        .setBackground('#B4D7A8');
      menuSheet.setRowHeight(row, 30);
      row++;

      const daysShort = [L('пн'),L('вт'),L('ср'),L('чт'),L('пт'),L('сб'),L('вс')];

      for (const item of config.alwaysOn) {
        if (!item.value || item.value.toString().trim() === '') continue;

        const period = (typeof item.dayFrom === 'number' && typeof item.dayTo === 'number')
          ? L('С ') + daysShort[item.dayFrom] + L(' по ') + daysShort[item.dayTo]
          : (typeof item.dayFrom === 'number' ? L('С ') + daysShort[item.dayFrom] : '');

        menuSheet.getRange(row, 1).setValue(period).setFontColor('#555555').setFontSize(10);
        menuSheet.getRange(row, 2, 1, totalCols - 1).merge();

        if (item.type === 'category') {
          const refs = assortmentRefs[item.value];
          if (refs && refs.weekly) {
            menuSheet.getRange(row, 2).setValue('=B' + refs.weekly);
          } else if (refs && refs.weeklyMulti && refs.weeklyMulti.length) {
            refs.weeklyMulti.forEach((refRow, idx) => {
              if (idx > 0) {
                menuSheet.getRange(row + idx, 1).setValue('');
                menuSheet.getRange(row + idx, 2, 1, totalCols - 1).merge();
              }
              menuSheet.getRange(row + idx, 2).setValue('=B' + refRow);
            });
            row += refs.weeklyMulti.length - 1;
          } else if (refs) {
            const firstRef = _firstAvailableDayRef(refs);
            if (firstRef) {
              menuSheet.getRange(row, 2).setValue('=B' + firstRef.row);
            } else {
              menuSheet.getRange(row, 2).setValue('[' + item.value + L(' не настроен]'));
            }
          } else {
            menuSheet.getRange(row, 2).setValue('[' + item.value + L(' не настроен]'));
          }
        } else {
          menuSheet.getRange(row, 2).setValue(item.value.toString().trim());
        }

        menuSheet.getRange(row, 1, 1, totalCols)
          .setBackground('#F0FAF0')
          .setWrap(true)
          .setVerticalAlignment('middle');
        menuSheet.setRowHeight(row, 25);
        row++;
      }
      row++;
    }

    const daysFull = [L('понедельник'),L('вторник'),L('среда'),L('четверг'),L('пятница')];

    menuSheet.getRange(row, 1).setValue(L('Время')).setFontWeight('bold').setBackground('#E8EAED').setFontSize(11);
    if (totalCols > 1) {
      menuSheet.getRange(row, 2, 1, totalCols - 1).merge();
      menuSheet.getRange(row, 2).setValue(L('Что выкладываем')).setFontWeight('bold').setBackground('#E8EAED').setFontSize(11);
    }
    menuSheet.getRange(row, 1, 1, totalCols).setBackground('#E8EAED');
    menuSheet.setRowHeight(row, 30);
    row++;

    for (let dayIdx = 0; dayIdx < 5; dayIdx++) {
      const date    = new Date(monday);
      date.setDate(date.getDate() + dayIdx);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy');

      const dayTimings = (config.timings || []).filter(t =>
        t.days && t.days.includes(dayIdx) && t.items && t.items.length
      );
      if (!dayTimings.length) continue;

      menuSheet.getRange(row, 1, 1, totalCols).merge();
      menuSheet.getRange(row, 1)
        .setValue(daysFull[dayIdx] + '  ' + dateStr)
        .setBackground('#D9EAD3')
        .setFontWeight('bold')
        .setFontSize(11)
        .setHorizontalAlignment('center')
        .setWrap(false);
      menuSheet.setRowHeight(row, 28);
      row++;

      const timeGroups = {};
      for (const timing of dayTimings) {
        const key = timing.timeFrom + '|' + (timing.timeTo || '');
        if (!timeGroups[key]) timeGroups[key] = [];
        for (const item of timing.items) {
          if (!item || !item.value) continue;

          if (item.type === 'category') {
            const refs = assortmentRefs[item.value];
            if (refs) {
              if (refs.weekly) {
                timeGroups[key].push('=B' + refs.weekly);

              } else if (refs.weeklyMulti && refs.weeklyMulti.length) {
                refs.weeklyMulti.forEach(function(refRow) {
                  timeGroups[key].push('=B' + refRow);
                });

              } else if (refs[dayIdx]) {
                const ref = refs[dayIdx];
                if (ref.cols && ref.cols > 1) {
                  const colLetters = ['B','C','D','E','F','G','H'];
                  for (let c = 0; c < ref.cols; c++) {
                    timeGroups[key].push('=' + (colLetters[c] || 'B') + ref.row);
                  }
                } else {
                  timeGroups[key].push('=B' + ref.row);
                }
              } else {
                timeGroups[key].push('[' + item.value + L(' не настроен]'));
              }
            } else {
              timeGroups[key].push('[' + item.value + L(' не найдена]'));
            }
          } else {
            if (item.value) timeGroups[key].push(item.value);
          }
        }
      }

      const sortedKeys = Object.keys(timeGroups).sort((a, b) =>
        a.split('|')[0].localeCompare(b.split('|')[0])
      );

      let rowColorIdx = 0;
      const rowColors = ['#FFFFFF', '#F8F9FA'];

      for (const key of sortedKeys) {
        const [timeFrom, timeTo] = key.split('|');
        const timeLabel = timeTo
          ? timeFrom + ' - ' + timeTo
          : timeFrom + L(' - пока не съедят');
        const items = timeGroups[key];

        const rowBg = rowColors[rowColorIdx % 2];
        rowColorIdx++;

        menuSheet.getRange(row, 1)
          .setValue(timeLabel)
          .setFontWeight('bold')
          .setFontSize(11)
          .setBackground(rowBg)
          .setWrap(true)
          .setVerticalAlignment('middle');

        for (let c = 0; c < Math.min(items.length, maxPositions); c++) {
          menuSheet.getRange(row, 2 + c)
            .setValue(items[c])
            .setBackground(rowBg)
            .setFontSize(11)
            .setWrap(true)
            .setVerticalAlignment('middle');
        }

        for (let c = items.length; c < maxPositions; c++) {
          menuSheet.getRange(row, 2 + c)
            .setValue('')
            .setBackground(rowBg);
        }

        menuSheet.setRowHeight(row, 40);
        row++;
      }
    }

    Logger.log('✅ Меню создано: ' + sheetName + ' (' + totalCols + ' колонок, макс. ' + maxPositions + ' позиций)');
    return menuSheet;
  }

  function _getMonday(date) {
    const d    = new Date(date);
    const day  = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1);
    return new Date(d.setDate(diff));
  }

  return { createWeeklyMenu, createNextWeekMenu, findMenuSheets, findTodaySheet };

})();

function refreshMenuAssortment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const warehouseSheet = SH_(ss, L('Склад_Продукты'));
  if (!warehouseSheet) return;

  const menuSheets = MenuManager.findMenuSheets();
  if (menuSheets.length === 0) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const relevantMenus = menuSheets.filter(menu => {
    const fri = new Date(menu.friday);
    fri.setHours(23, 59, 59, 999);
    return fri >= today;
  });

  const targets = relevantMenus.length > 0 ? relevantMenus : [menuSheets[menuSheets.length - 1]];

  targets.forEach(menu => _refreshAssortmentOnSheet(menu.sheet, warehouseSheet));

  SpreadsheetApp.getUi().alert(
    L('✅ Ассортимент обновлён на листах:\n\n') + targets.map(m => m.name).join('\n')
  );
}

function _refreshAssortmentOnSheet(menuSheet, warehouseSheet) {
  const data = menuSheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    const label = data[i][0] ? data[i][0].toString().trim() : '';
    if (!label) continue;

    const categoryName = _extractCategoryName(label, warehouseSheet);
    if (!categoryName) continue;

    const values = _getProductsFromWarehouseForDropdown(warehouseSheet, categoryName);
    if (values.length === 0) continue;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(true)
      .build();

    const sheetRow = i + 1;
    const validatedCols = _getValidatedColumnsInRow(menuSheet, sheetRow);

    if (validatedCols.length > 0) {
      validatedCols.forEach(col => {
        menuSheet.getRange(sheetRow, col).setDataValidation(rule);
      });
    } else {
      menuSheet.getRange(sheetRow, 2).setDataValidation(rule);
    }
  }
}

function _extractCategoryName(label, warehouseSheet) {
  const suffixes = [
    / (?:на неделю|for the week):?$/i,
    / (?:вид|type) \d+:?$/i,
    / (?:пн|вт|ср|чт|пт|сб|вс|mon|tue|wed|thu|fri|sat|sun)\s*\([\d.]+\):?$/i
  ];

  let candidate = label.replace(/:$/, '').trim();

  for (const pattern of suffixes) {
    candidate = candidate.replace(pattern, '').trim();
  }

  if (!candidate) return null;

  const categories = _getAllWarehouseCategoriesFromSheet(warehouseSheet);
  if (categories.includes(candidate)) return candidate;

  return null;
}

function _getAllWarehouseCategoriesFromSheet(warehouseSheet) {
  const data = warehouseSheet.getDataRange().getValues();
  const cats = [];
  for (let i = 2; i < data.length; i++) {
    const name    = data[i][0] ? data[i][0].toString().trim() : '';
    const hasPlus = data[i][6] && data[i][6].toString().includes('➕');
    if (hasPlus && name) cats.push(name);
  }
  return cats;
}

function _getProductsFromWarehouseForDropdown(warehouseSheet, categoryName) {
  const data     = warehouseSheet.getDataRange().getValues();
  const products = [];
  let inCategory = false;

  for (let i = 2; i < data.length; i++) {
    const name    = data[i][0] ? data[i][0].toString().trim() : '';
    const hasPlus = data[i][6] && data[i][6].toString().includes('➕');
    const active  = data[i][3];

    if (hasPlus && name) {
      if (name === categoryName) inCategory = true;
      else if (inCategory)       break;
      continue;
    }
    if (inCategory && name && active === true) products.push(name);
  }

  return [...products, L('Остатки'), L('Отсутствует')];
}

function cleanupOldMenuSheets() {
  const menus = MenuManager.findMenuSheets();
  if (menus.length <= 1) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let deleted = 0;
  menus.forEach(function(menu) {
    const fri = new Date(menu.friday); fri.setHours(23, 59, 59, 999);
    if (fri < today && menus.length - deleted > 1) {
      try { ss.deleteSheet(menu.sheet); deleted++; Logger.log('✅ Удалён: ' + menu.name); }
      catch(e) { menu.sheet.hideSheet(); Logger.log('⚠️ Скрыт: ' + menu.name); }
    }
  });
  Logger.log('cleanupOldMenuSheets: удалено ' + deleted + ' листов');
}

function autoFillMenuByStock() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const menuSheet = ss.getActiveSheet();

  if (!/^(?:Меню|Menu)\s+\d{2}\.\d{2}-\d{2}\.\d{2}$/.test(menuSheet.getName())) {
    ui.alert(
      L('❌ Откройте лист меню\n\n') +
      L('Сначала перейдите на вкладку листа меню ') +
      L('(вида "Меню 01.07-05.07"), затем повторите команду.')
    );
    return;
  }

  const warehouseSheet = SH_(ss, L('Склад_Продукты'));
  if (!warehouseSheet) {
    ui.alert(L('❌ Лист "Склад_Продукты" не найден'));
    return;
  }

  const data = menuSheet.getDataRange().getValues();
  let startRow = -1, endRow = data.length;

  for (let i = 0; i < data.length; i++) {
    const label = data[i][0] ? data[i][0].toString() : '';
    if (startRow === -1 && isMenuAssortLabel_(label)) { startRow = i + 1; continue; }
    if (startRow !== -1 &&
        (isMenuAlwaysLabel_(label) || isMenuTimeLabel_(label))) {
      endRow = i;
      break;
    }
  }

  if (startRow === -1) {
    ui.alert(
      L('⚠️ В этом меню нет раздела "Выбор ассортимента на неделю"\n\n') +
      L('Рандомайзер работает только со сменными категориями,\n') +
      L('настроенными через Конструктор меню.')
    );
    return;
  }

  const categoryRows = {};
  let skippedNoCategory = 0;

  for (let idx = startRow; idx < endRow; idx++) {
    const label = data[idx][0] ? data[idx][0].toString().trim() : '';
    if (!label) continue;

    const category = _extractCategoryName(label, warehouseSheet);
    if (!category) { skippedNoCategory++; continue; }

    const sheetRow = idx + 1;
    const cols = _getValidatedColumnsInRow(menuSheet, sheetRow);
    if (cols.length === 0) continue;

    if (!categoryRows[category]) categoryRows[category] = [];
    categoryRows[category].push({ row: sheetRow, cols: cols });
  }

  let filledRows = 0, filledSlots = 0;

  Object.keys(categoryRows).forEach(category => {
    const rowsForCategory = categoryRows[category];
    const totalSlots = rowsForCategory.reduce((sum, r) => sum + r.cols.length, 0);

    const products = _getCategoryProductsWithStock(warehouseSheet, category);
    const chosen   = _pickProductsByStock(products, totalSlots);

    let cursor = 0;
    rowsForCategory.forEach(r => {
      r.cols.forEach(col => {
        menuSheet.getRange(r.row, col).setValue(chosen[cursor]);
        cursor++;
      });
      filledRows++;
      filledSlots += r.cols.length;
    });
  });

  if (filledRows === 0) {
    ui.alert(L('⚠️ Не найдено ни одной заполняемой позиции ассортимента.'));
    return;
  }

  ui.alert(
    L('✅ Рандомайзер заполнил ассортимент!\n\n') +
    L('Категорий: ') + Object.keys(categoryRows).length + '\n' +
    L('Строк: ') + filledRows + '\n' +
    L('Позиций: ') + filledSlots +
    (skippedNoCategory > 0 ? L('\n\nПропущено строк (категория не распознана): ') + skippedNoCategory : '')
  );
}

function _getValidatedColumnsInRow(sheet, row) {
  const MAX_SCAN = 20;
  const rng = sheet.getRange(row, 2, 1, MAX_SCAN);
  const validations = rng.getDataValidations()[0];
  const cols = [];
  for (let i = 0; i < validations.length; i++) {
    if (validations[i]) cols.push(2 + i);
  }
  return cols;
}

function _getCategoryProductsWithStock(warehouseSheet, categoryName) {
  const data = warehouseSheet.getDataRange().getValues();
  const products = [];
  let inCategory = false;

  for (let i = 2; i < data.length; i++) {
    const name    = data[i][0] ? data[i][0].toString().trim() : '';
    const qty     = data[i][1];
    const active  = data[i][3];
    const hasPlus = data[i][6] && data[i][6].toString().includes('➕');

    if (hasPlus && name) {
      if (name === categoryName) inCategory = true;
      else if (inCategory) break;
      continue;
    }
    if (inCategory && name && active === true) {
      products.push({ name: name, qty: parseFloat(qty) || 0 });
    }
  }
  return products;
}

function _pickProductsByStock(products, count) {
  if (!products || products.length === 0) {
    return new Array(count).fill(L('Отсутствует'));
  }

  const shuffled = _shuffleArray(products);
  const ordered  = shuffled.slice().sort((a, b) => b.qty - a.qty);

  const chosen = ordered.slice(0, count).map(p => p.name);
  while (chosen.length < count) chosen.push(L('Остатки'));
  return chosen;
}

function _shuffleArray(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const tmp = a[i]; a[i] = a[j]; a[j] = tmp;
  }
  return a;
}

function autoUpdateMenu()         { MenuManager.createNextWeekMenu(); }
function generateMenuFromConfig() { MenuManager.createWeeklyMenu();   }
function createNextWeekMenu()     { MenuManager.createNextWeekMenu(); }
