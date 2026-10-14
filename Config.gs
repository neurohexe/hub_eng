// Config.gs

const WAREHOUSE_CONFIG = {
  PRODUCTS: {
    name:              'Продукты',
    icon:              '🥗',
    color:             '#1c4587',
    warehouseSheet:    'Склад_Продукты',
    dashSheet:         'Продукты',
    hostessSheet:      'food_hostess',
    logistSheet:       'postavka_food',
    cleaningSheet:     null,
    dbSheet:           'Products_DB',
    excludeCategories: ['Алкоголь'],
    hasCleaning:       false
  },
  HOUSEHOLD: {
    name:              'Хозтовары',
    icon:              '🧹',
    color:             '#2196F3',
    warehouseSheet:    'Склад_Хозтовары',
    dashSheet:         'Хозтовары',
    hostessSheet:      'household_hostess',
    logistSheet:       'postavka_household',
    cleaningSheet:     'household_cleaning',
    dbSheet:           'Household_DB',
    excludeCategories: [],
    hasCleaning:       true
  },
  STATIONERY: {
    name:              'Канцелярия',
    icon:              '✏️',
    color:             '#FF9800',
    warehouseSheet:    'Склад_Канцелярия',
    dashSheet:         'Канцелярия',
    hostessSheet:      'stationery_hostess',
    logistSheet:       'postavka_stationery',
    cleaningSheet:     null,
    dbSheet:           'Stationery_DB',
    excludeCategories: [],
    hasCleaning:       false
  },
  MEDICINE: {
    name:              'Медикаменты',
    icon:              '💊',
    color:             '#E91E63',
    warehouseSheet:    'Склад_Медикаменты',
    dashSheet:         'Медикаменты',
    hostessSheet:      'medicine_hostess',
    logistSheet:       'postavka_medicine',
    cleaningSheet:     null,
    dbSheet:           'Medicine_DB',
    excludeCategories: [],
    hasCleaning:       false
  }
};


// name — геттер: переводится в момент обращения (L нельзя вызывать при загрузке файла)
const WAREHOUSE_TYPES = Object.keys(WAREHOUSE_CONFIG).reduce((acc, key) => {
  const c = WAREHOUSE_CONFIG[key];
  acc[key] = { get name() { return L(c.name); }, color: c.color, icon: c.icon };
  return acc;
}, {});

const WAREHOUSE_SHEET_NAMES = Object.keys(WAREHOUSE_CONFIG).reduce((acc, key) => {
  acc[key] = WAREHOUSE_CONFIG[key].warehouseSheet;
  return acc;
}, {});


const FORM_SHEET_TO_DASHBOARD = (function() {
  const map = {};
  Object.keys(WAREHOUSE_CONFIG).forEach(key => {
    const c = WAREHOUSE_CONFIG[key];
    if (c.hostessSheet)  map[c.hostessSheet]  = key;
    if (c.logistSheet)   map[c.logistSheet]   = key;
    if (c.cleaningSheet) map[c.cleaningSheet] = key;
  });
  return map;
})();

function getWarehouseConfigByType(warehouseType) {
  const cfg = WAREHOUSE_CONFIG[warehouseType];
  if (!cfg) throw new Error(L('Неизвестный тип склада: ') + warehouseType);
  return cfg;
}


const PRODUCT_CATEGORIES_LIST = [
  'Чай', 'Молоко', 'Молоко растительное', 'Кофе для кофемашин',
  'Напитки растворимые', 'Мед/Джем', 'Сладкое', 'Сухие завтраки',
  'Крекеры', 'Снэки', 'Бакалея', 'Масло свежее (холодильник)',
  'Сиропы', 'Орех/Сухофрукт', 'Шоколад', 'Алкоголь'
];

function getProductCategoriesList() {
  return PRODUCT_CATEGORIES_LIST.map(L);
}

const DEFAULT_UNITS = ['кг', 'штука', 'упаковка', 'блок', 'л', 'мл', 'г', 'пакет'];
const PROP_UNITS    = 'CUSTOM_UNITS';

function getUnits() {
  try {
    const saved  = PropertiesService.getScriptProperties().getProperty(PROP_UNITS);
    const custom = saved ? JSON.parse(saved) : [];
    const all    = DEFAULT_UNITS.map(L);
    custom.forEach(u => { if (u && !all.includes(u)) all.push(u); });
    return all;
  } catch(e) {
    Logger.log('⚠️ getUnits: ' + e.toString());
    return DEFAULT_UNITS.map(L);
  }
}

function getUnitsValidationRule() {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(getUnits(), true)
    .setAllowInvalid(true)
    .build();
}

function getOfficeName() {
  try {
    const cfg = BaseInfo.getOfficeConfig();
    return (cfg && cfg.officeName) ? cfg.officeName.toString() : '';
  } catch(e) {
    return '';
  }
}

function getWeekPeriodFromMenu(menuSheet) {
  try {
    if (menuSheet) {
      var nm = menuSheet.getName().match(/(\d{2}\.\d{2})-(\d{2}\.\d{2})/);
      if (nm) return nm[1] + ' - ' + nm[2];
    }
  } catch(e) {}
  try {
    if (menuSheet) {
      var v = menuSheet.getRange('A1').getValue().toString();
      var m = v.match(/(?:с|from)\s+(\d{2}\.\d{2})\s+(?:по|to)\s+(\d{2}\.\d{2})/i);
      if (m) return m[1] + ' - ' + m[2];
    }
  } catch(e) {}
  var mon = _getMonday(new Date()), fri = new Date(mon);
  fri.setDate(fri.getDate() + 4);
  return _fmtDate(mon) + ' - ' + _fmtDate(fri);
}

function _getMonday(date) {
  var d = new Date(date), day = d.getDay();
  d.setDate(d.getDate() - day + (day === 0 ? -6 : 1));
  return d;
}

function _fmtDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM');
}
