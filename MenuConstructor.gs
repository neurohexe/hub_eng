// MenuConstructor.gs

function openMenuConstructor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const warehouseSheet = SH_(ss, L('Склад_Продукты'));
  if (!warehouseSheet) {
    SpreadsheetApp.getUi().alert(L('❌ Сначала создайте склад продуктов!'));
    return;
  }
  const categories = getWarehouseCategories(warehouseSheet);
  if (categories.length === 0) {
    SpreadsheetApp.getUi().alert(L('⚠️ На складе нет категорий. Пересоздайте склад.'));
    return;
  }
  const html = htmlTemplate_('ConstructorDialog')
    .evaluate()
    .setWidth(950)
    .setHeight(750)
    .setTitle(L('Конструктор меню'));
  SpreadsheetApp.getUi().showModalDialog(html, L('Конструктор меню'));
}

function getWarehouseCategories(warehouseSheet) {
  const data = warehouseSheet.getDataRange().getValues();
  const categories = [];
  for (let i = 2; i < data.length; i++) {
    const name = data[i][0] ? data[i][0].toString().trim() : '';
    const hasPlus = data[i][6] && data[i][6].toString().includes('➕');
    if (hasPlus && name) categories.push(name);
  }
  return categories;
}

function getCategoryProducts(categoryName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const warehouseSheet = SH_(ss, L('Склад_Продукты'));
    if (!warehouseSheet) return [];
    const data = warehouseSheet.getDataRange().getValues();
    const products = [];
    let inCategory = false;
    for (let i = 2; i < data.length; i++) {
      const name = data[i][0] ? data[i][0].toString().trim() : '';
      const hasPlus = data[i][6] && data[i][6].toString().includes('➕');
  
      if (hasPlus && name) {
        if (name === categoryName) inCategory = true;
        else if (inCategory) break;
        continue;
      }

      if (inCategory && name && data[i][3] === true) {
        products.push(name);
      }
    }
    return products;
  } catch(e) { return []; }
}

function getMenuConstructorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const warehouseSheet = SH_(ss, L('Склад_Продукты'));
  const categories = warehouseSheet ? getWarehouseCategories(warehouseSheet) : [];
  const config = loadMenuConfig();
  return { categories: categories, config: config };
}

function saveMenuConfig(configJson) {
  try {
    const config = JSON.parse(configJson);
    PropertiesService.getScriptProperties().setProperty('MENU_CONFIG', JSON.stringify(config));
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function loadMenuConfig() {
  try {
    const saved = PropertiesService.getScriptProperties().getProperty('MENU_CONFIG');
    if (saved) return JSON.parse(saved);
  } catch(e) {}
  return { rotatingCategories: [], timings: [], alwaysOn: [] };
}


function fixMenuConfig() {
  const config = loadMenuConfig();

  config.rotatingCategories = config.rotatingCategories.map(rc => ({
    category:   rc.category   || '',
    mode:       rc.mode       || 'weekly',
    days:       (rc.days      || [0,1,2,3,4]).slice().sort((a,b) => a - b),
    countPerDay: rc.countPerDay || 1
  }));

  config.alwaysOn = config.alwaysOn.map(item => ({
    ...item,
    type:    item.type    || 'text',
    value:   (item.value  || '').trim(),
    dayFrom: item.dayFrom !== undefined ? item.dayFrom : null,
    dayTo:   item.dayTo   !== undefined ? item.dayTo   : null
  }));

  delete config.schedule;


  config.timings = config.timings.map(t => ({
    ...t,
    days: (t.days || []).slice().sort((a,b) => a - b)
  }));

  PropertiesService.getScriptProperties()
    .setProperty('MENU_CONFIG', JSON.stringify(config));

  Logger.log('✅ Конфиг почищен');
  Logger.log('rotatingCategories: ' + JSON.stringify(config.rotatingCategories, null, 2));
  Logger.log('alwaysOn: ' + JSON.stringify(config.alwaysOn, null, 2));

  SpreadsheetApp.getUi().alert(
    L('✅ Конфиг исправлен!\n\n') +
    L('Категорий: ') + config.rotatingCategories.length + '\n' +
    L('Таймингов: ') + config.timings.length + '\n' +
    L('Круглосуточно: ') + config.alwaysOn.length + '\n\n' +
    L('Откройте конструктор заново.')
  );
}
