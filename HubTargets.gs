// HubTargets.gs

function _getOrCreateHubTargetsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = SH_(ss, HUB_TARGETS_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(L(HUB_TARGETS_SHEET));
  sheet.getRange(1, 1, 1, 6).setValues([[
    L('Товар'), L('Категория'), L('Целевой остаток'), L('Предложение (авто)'), L('Источник'), L('Обновлено')
  ]]).setBackground('#6a329f').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260); sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 110); sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 100); sheet.setColumnWidth(6, 100);
  return sheet;
}

// Синхронизирует лист целевых остатков со списком активных товаров на складах:
// добавляет новые строки (целевой остаток = 0, заполняется вручную), ничего не удаляет
// и не трогает уже заполненные значения.
function syncHubTargetsSheet() {
  const sheet    = _getOrCreateHubTargetsSheet();
  const products = getAllHubProducts();
  const data     = sheet.getDataRange().getValues();
  const existing = new Set(data.slice(1).map(r => (r[0] || '').toString().trim()));

  let row = sheet.getLastRow() + 1, added = 0;
  products.forEach(p => {
    if (existing.has(p.name)) return;
    sheet.getRange(row, 1).setValue(p.name);
    sheet.getRange(row, 2).setValue(p.category);
    sheet.getRange(row, 3).setValue(0);
    row++; added++;
  });

  SpreadsheetApp.getUi().alert(L('✅ Целевые остатки синхронизированы.\nДобавлено новых товаров: ') + added);
  return added;
}

// Карта "товар → целевой остаток" — только заполненные вручную значения > 0.
// Товары без заданного target в авто-заказ не попадают.
function getHubTargetsMap() {
  const sheet = _getOrCreateHubTargetsSheet();
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < data.length; i++) {
    const name   = data[i][0] ? data[i][0].toString().trim() : '';
    const target = parseFloat(data[i][2]);
    if (name && !isNaN(target) && target > 0) map[name] = target;
  }
  return map;
}

// Простая эвристика: средний расход за последние N записей истории × запас.
// Пишет результат в колонку "Предложение (авто)" — саму колонку "Целевой остаток" не трогает.
function suggestHubTargets(weeksWindow, safetyFactor) {
  weeksWindow  = weeksWindow  || 6;
  safetyFactor = safetyFactor || 1.2;

  const history      = _readHubHistory();
  const targetsSheet = _getOrCreateHubTargetsSheet();
  const data         = targetsSheet.getDataRange().getValues();

  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    if (!name) continue;
    const consumptions = (history[name] || []).slice(-weeksWindow);
    if (consumptions.length === 0) continue;

    const avg        = consumptions.reduce((s, v) => s + v, 0) / consumptions.length;
    const suggestion = Math.ceil(avg * safetyFactor);
    targetsSheet.getRange(i + 1, 4).setValue(suggestion);
    targetsSheet.getRange(i + 1, 6).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy')
    );
    updated++;
  }

  SpreadsheetApp.getUi().alert(
    L('Предложения обновлены для ') + updated + L(' товаров.\n\n') +
    L('Это только рекомендация — чтобы применить, скопируйте значение\n') +
    L('из колонки "Предложение (авто)" в колонку "Целевой остаток".')
  );
  return updated;
}

// Читает лист истории остатков и возвращает { товар: [расход1, расход2, ...] } по датам.
function _readHubHistory() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = SH_(ss, HUB_HISTORY_SHEET);
  const result = {};
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const name     = data[i][1] ? data[i][1].toString().trim() : '';
    const consumed = parseFloat(data[i][5]);
    if (!name || isNaN(consumed)) continue;
    if (!result[name]) result[name] = [];
    result[name].push(consumed);
  }
  return result;
}

function suggestHubTargetsMenu()  { suggestHubTargets();  }
function syncHubTargetsSheetMenu() { syncHubTargetsSheet(); }
