// Language.gs
// Переключатель языка RU / EN для всей системы склада.
//
// Как устроено:
//   L('русский текст')  — текст для человека (меню, окна, формы, Telegram, заголовки листов).
//                         В режиме EN возвращает английский вариант, иначе — как есть.
//   K(значение ячейки)  — для сравнений в коде: приводит и русский, и английский вариант
//                         к одному русскому ключу. Поэтому логика работает с листами
//                         на любом языке, в том числе со старыми.
//   SH_(ss, имя листа)  — ищет лист под русским и под английским именем.
//
// Важно: L/K/SH_ нельзя вызывать на верхнем уровне других файлов (при загрузке
// скрипта) — Apps Script загружает файлы по очереди, и словарь может быть ещё не готов.
// Поэтому константы в коде хранятся русскими ключами и переводятся в момент использования.
//
// Товары, категории и единицы, которые уже внесены в таблицу, не переводятся —
// это данные пользователя. Новые склады создаются с категориями на выбранном языке.

const LANG_PROP = 'UI_LANG';
var LANG_CACHE_ = null;
var I18N_MAPS_ = null;

function getLang() {
  if (LANG_CACHE_) return LANG_CACHE_;
  var v = null;
  try { v = PropertiesService.getScriptProperties().getProperty(LANG_PROP); } catch (e) {}
  if (!v) {
    // первая версия переключателя хранила язык в свойствах документа
    try { var dp = PropertiesService.getDocumentProperties(); v = dp && dp.getProperty(LANG_PROP); } catch (e) {}
  }
  LANG_CACHE_ = v === 'en' ? 'en' : 'ru';
  return LANG_CACHE_;
}

function setLang_(lang) {
  PropertiesService.getScriptProperties().setProperty(LANG_PROP, lang);
  try { var dp = PropertiesService.getDocumentProperties(); if (dp) dp.deleteProperty(LANG_PROP); } catch (e) {}
  LANG_CACHE_ = lang;
}

function i18nMaps_() {
  if (I18N_MAPS_) return I18N_MAPS_;
  var en = {}, revExact = {}, revLower = {};
  [I18N_UI_, I18N_DATA_].forEach(function(dict) {
    Object.keys(dict).forEach(function(ru) {
      en[ru] = dict[ru];
      var e = dict[ru];
      if (!Object.prototype.hasOwnProperty.call(revExact, e)) revExact[e] = ru;
      var lo = e.toLowerCase();
      if (!Object.prototype.hasOwnProperty.call(revLower, lo)) revLower[lo] = ru;
    });
  });
  I18N_MAPS_ = { en: en, revExact: revExact, revLower: revLower };
  return I18N_MAPS_;
}

function hasOwn_(o, k) { return Object.prototype.hasOwnProperty.call(o, k); }

/** Английский вариант русского текста (независимо от текущего языка); null, если перевода нет. */
function enOf_(ru) {
  var s = String(ru == null ? '' : ru), m = i18nMaps_().en;
  if (hasOwn_(m, s)) return m[s];
  var p = s.match(/^(\s*)([\s\S]*?)(\s*)$/);
  if (p[2] && hasOwn_(m, p[2])) return p[1] + m[p[2]] + p[3];
  return null;
}

/** Текст для человека на текущем языке. */
function L(ru) {
  if (ru === null || ru === undefined) return ru;
  if (getLang() !== 'en') return ru;
  var e = enOf_(ru);
  return e === null ? ru : e;
}

/** Ключ для сравнения: английский системный текст → русский; всё остальное — как есть (обрезанное). */
function K(v) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim(), m = i18nMaps_();
  if (hasOwn_(m.revExact, s)) return m.revExact[s];
  var lo = s.toLowerCase();
  if (hasOwn_(m.revLower, lo)) return m.revLower[lo];
  return s;
}

/** Совместимость со старым кодом. */
function hdrRu_(h) { return K(h); }

/** Лист по имени: текущий язык → как передано → русский ключ → английский вариант. */
function SH_(ss, name) {
  if (!ss || name === null || name === undefined) return null;
  var n = String(name);
  var tries = [L(n), n, K(n), enOf_(n)];
  for (var i = 0; i < tries.length; i++) {
    if (!tries[i]) continue;
    var sh = ss.getSheetByName(tries[i]);
    if (sh) return sh;
  }
  return null;
}

/** Лист склада (Склад_* / Warehouse_*). */
function isWarehouseSheet_(sheet) {
  var name = sheet.getName(), key = K(name);
  return key.indexOf('Склад_') === 0 || name.indexOf('Warehouse_') === 0;
}

// ── Метки листа меню (понимают оба языка) ──
function isMenuAssortLabel_(s) {
  var u = String(s == null ? '' : s).toUpperCase();
  return u.indexOf('ВЫБОР АССОРТИМЕНТА') !== -1 || u.indexOf('ASSORTMENT FOR THE WEEK') !== -1;
}
function isMenuAlwaysLabel_(s) {
  var u = String(s == null ? '' : s).toUpperCase();
  return u.indexOf('КРУГЛОСУТОЧНО') !== -1 || u.indexOf('ALWAYS AVAILABLE') !== -1;
}
function isMenuTimeLabel_(s) {
  var u = String(s == null ? '' : s).trim().toUpperCase();
  return u.indexOf('ВРЕМЯ ВЫКЛАДКИ') !== -1 || u === 'ВРЕМЯ' || u === 'TIME';
}
/** «Остатки» / «Отсутствует» в выпадающих списках меню — не товары. */
function isMenuFiller_(s) {
  var k = K(s);
  return k === 'Остатки' || k === 'Отсутствует';
}

/** Заголовок A1 листа склада на текущем языке. */
function warehouseTitle_(key, today) {
  var name = WAREHOUSE_CONFIG[key].name;
  return getLang() === 'en'
    ? L(name) + ' warehouse data as of ' + today
    : 'Данные ' + name.toLowerCase() + ' склада на ' + today;
}

// ── HTML-окна ──
/** Шаблон HTML-файла с текстами на текущем языке (вызывать вместо createTemplateFromFile). */
function htmlTemplate_(file) {
  var raw = HtmlService.createTemplateFromFile(file).getRawContent();
  return HtmlService.createTemplate(localizeHtml_(raw));
}

function localizeHtml_(raw) {
  if (getLang() !== 'en') return raw;
  var out = raw.replace('<html lang="ru">', '<html lang="en">');
  var keys = Object.keys(I18N_HTML_).sort(function(a, b) { return b.length - a.length; });
  keys.forEach(function(ru) { out = out.split(ru).join(I18N_HTML_[ru]); });
  return out;
}

// ── Переключение ──
/** Системные листы: русское имя → английское берётся из словаря. */
const I18N_SHEETS_ = [
  'Склад_Продукты', 'Склад_Хозтовары', 'Склад_Канцелярия', 'Склад_Медикаменты',
  'Продукты', 'Хозтовары', 'Канцелярия', 'Медикаменты',
  'Базовая информация', 'Меню',
  'Заказ продуктов', 'Заказ на неделю', 'Заказ хозтоваров', 'Заказ канцелярии', 'Заказ медикаментов',
  'Хаб_Целевые_остатки', 'Хаб_История_движения_остатков'
];

/** Кнопка переключения языка (пункт меню 🌐). */
function toggleLanguage() {
  var next = getLang() === 'en' ? 'ru' : 'en';
  setLang_(next);
  I18N_MAPS_ = null;

  var steps = [
    function() { renameSystemSheets_(); },
    function() { relabelSystemSheets_(); },
    function() { updateAllWarehouseDates(); },
    function() { resaveMenuStorage_(); }
  ];
  steps.forEach(function(fn) { try { fn(); } catch (e) { Logger.log('⚠️ toggleLanguage: ' + e); } });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    next === 'en'
      ? 'Sheets switched to English. Reload the page (F5) to switch the menus. Google Forms update the next time you refresh them.'
      : 'Листы переключены на русский. Обновите страницу (F5), чтобы сменилось меню. Google Формы обновятся при следующем обновлении форм.',
    next === 'en' ? '🌐 English' : '🌐 Русский', 12);
}

function addLanguageMenu_(ui) {
  var en = getLang() === 'en';
  ui.createMenu(en ? '🌐 EN | ru' : '🌐 RU | en')
    .addItem(en ? 'Переключить на русский' : 'Switch to English', 'toggleLanguage')
    .addToUi();
}

function renameSystemSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var en = getLang() === 'en';
  var map = {};
  I18N_SHEETS_.forEach(function(ru) {
    var e = enOf_(ru);
    if (e) { if (en) map[ru] = e; else map[e] = ru; }
  });
  ss.getSheets().forEach(function(sh) {
    var name = sh.getName(), target = map[name] || null;
    var m = name.match(/^(Меню|Menu)(\s+\d{2}\.\d{2}-\d{2}\.\d{2})$/);
    if (!target && m) target = (en ? 'Menu' : 'Меню') + m[2];
    if (!target || target === name || ss.getSheetByName(target)) return;
    sh.setName(target);
  });
}

/** Переводит системные подписи на листах. Товары, категории и единицы не трогает. */
function relabelSystemSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var base = SH_(ss, 'Базовая информация');
  if (base) relabelRange_(base.getDataRange());

  Object.keys(WAREHOUSE_CONFIG).forEach(function(key) {
    var cfg = WAREHOUSE_CONFIG[key];
    var wh = SH_(ss, cfg.warehouseSheet);
    if (wh && wh.getLastRow() >= 2 && wh.getLastColumn() > 0) relabelRange_(wh.getRange(2, 1, 1, wh.getLastColumn()));
    var dash = SH_(ss, cfg.dashSheet);
    if (dash && dash.getLastRow() >= 2 && dash.getLastColumn() > 0) {
      relabelRange_(dash.getRange(2, 1, Math.min(2, dash.getLastRow() - 1), dash.getLastColumn()));
    }
  });

  Object.keys(ORDER_SHEET_NAMES).forEach(function(k) {
    var sh = SH_(ss, ORDER_SHEET_NAMES[k]);
    if (!sh || sh.getLastRow() < 2) return;
    relabelRange_(sh.getRange(2, 1, 1, sh.getLastColumn()));
    relabelRange_(sh.getRange(3, 1, Math.max(1, sh.getLastRow() - 2), 1));
  });

  [HUB_TARGETS_SHEET, HUB_HISTORY_SHEET].forEach(function(n) {
    var sh = SH_(ss, n);
    if (sh && sh.getLastColumn() > 0) relabelRange_(sh.getRange(1, 1, 1, sh.getLastColumn()));
  });
}

/** Переводит только ячейки-подписи из словаря интерфейса; формулы и данные не трогает. */
function relabelRange_(range) {
  var values = range.getValues(), formulas = range.getFormulas();
  var en = getLang() === 'en', m = i18nMaps_();
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var v = values[r][c];
      if (typeof v !== 'string' || !v.trim() || formulas[r][c]) continue;
      var s = v.trim(), target = null;
      var ru = hasOwn_(I18N_UI_, s) ? s : (hasOwn_(m.revExact, s) && hasOwn_(I18N_UI_, m.revExact[s]) ? m.revExact[s] : null);
      if (ru) target = en ? I18N_UI_[ru] : ru;
      if (!target) {
        // «БАЗОВАЯ ИНФОРМАЦИЯ ОФИСА — обновлено 01.01.2026»
        var bm = s.match(/^(БАЗОВАЯ ИНФОРМАЦИЯ ОФИСА — обновлено|OFFICE BASE INFO — updated)\s+(.*)$/);
        if (bm) target = (en ? 'OFFICE BASE INFO — updated' : 'БАЗОВАЯ ИНФОРМАЦИЯ ОФИСА — обновлено') + ' ' + bm[2];
      }
      if (target && target !== v) range.getCell(r + 1, c + 1).setValue(target);
    }
  }
}

/** Пересохраняет данные меню для планшетов, чтобы дни и месяцы были на новом языке. */
function resaveMenuStorage_() {
  var today = new Date(); today.setHours(0, 0, 0, 0);
  MenuManager.findMenuSheets().forEach(function(menu) {
    var mon = new Date(menu.monday); mon.setHours(0, 0, 0, 0);
    saveMenuDataToStorage(menu.sheet, mon > today);
  });
}

/** Шапка складов на текущем языке (вызывается при открытии таблицы). */
function applyLanguageToWarehouseHeaders_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(WAREHOUSE_CONFIG).forEach(function(key) {
    var wh = SH_(ss, WAREHOUSE_CONFIG[key].warehouseSheet);
    if (wh && wh.getLastRow() >= 2 && wh.getLastColumn() > 0) relabelRange_(wh.getRange(2, 1, 1, wh.getLastColumn()));
  });
}

// ── Словари (сгенерированы; русский → английский) ──
const I18N_UI_ = {
 "\" в конец склада?": "\" at the end of the warehouse?",
 "\" добавлена!": "\" added!",
 "\" ещё не создан": "\" has not been created yet",
 "\" не найден": "\" not found",
 "\" не найден.": "\" not found.",
 "\" не появился вовремя": "\" did not appear in time",
 "\" обновлён (": "\" updated (",
 "\" обновлён!": "\" updated!",
 "\" создана!": "\" created!",
 "\" удалена.": "\" removed.",
 "\" уже есть в списке": "\" is already in the list",
 "\" уже существует": "\" already exists",
 "\"message_thread_id\": <число>": "\"message_thread_id\": <number>",
 "(а НЕ отдельная группа/чат) — нажмите YES.": "(NOT a separate group/chat), press YES.",
 "(бесплатно, карта не нужна — есть бесплатный лимит запросов в день)": "(free, no card needed: there is a free daily request limit)",
 "(вида \"Меню 01.07-05.07\"), затем повторите команду.": "(like \"Menu 01.07-05.07\"), then run the command again.",
 "(дневные дашборды и архивы *_DB — хабу они не нужны)": "(daily dashboards and *_DB archives: a hub does not need them)",
 "(ед. изм.:": "(unit:",
 "(не в \"Общие\", а именно в теме с логистом)": "(not in \"General\", but in the topic with the logistician)",
 "(пока ни одной)": "(none yet)",
 "(поставщик:": "(supplier:",
 "(режим \"Хаб\" — ежедневный архив и сброс дашбордов НЕ устанавливаются:": "(\"Hub\" mode: daily archive and dashboard reset are NOT installed:",
 "(склад:": "(stock:",
 "(товаров:": "(items:",
 "(форма остатков хаба)": "(hub stock form)",
 "(➕ Добавить категорию / Добавление строк).": "(➕ Add category / Add rows).",
 ") не открывается:": ") cannot be opened:",
 "- category: если товар подходит под уже существующую категорию этого склада —": "- category: if the item fits an existing category of this warehouse,",
 "- isNewCategory: true, если категории с таким названием ещё нет в списке существующих для этого склада.": "- isNewCategory: true if a category with this name is not yet in the list of existing ones for this warehouse.",
 "- warehouseType: один из [": "- warehouseType: one of [",
 "- пока не съедят": "- until it runs out",
 ". Новая форма НЕ создана, чтобы не менять постоянную ссылку.": ". A new form was NOT created, so the permanent link stays the same.",
 "1 - Продукты\n2 - Хозтовары\n3 - Канцелярия\n4 - Медикаменты": "1 - Food\n2 - Household\n3 - Stationery\n4 - Medicine",
 "1. Напишите /start боту в чате": "1. Send /start to the bot in the chat",
 "1. Напишите любое сообщение боту ВНУТРИ нужной темы": "1. Send any message to the bot INSIDE the topic you need",
 "1. Развернуть → Новое развёртывание": "1. Deploy → New deployment",
 "1. Скопируйте ID выше (нажмите на него)": "1. Copy the ID above (tap it)",
 "2. Откройте в браузере:": "2. Open in your browser:",
 "2. Откройте таблицу Google Sheets": "2. Open the Google Sheets spreadsheet",
 "2. Тип: Веб-приложение": "2. Type: Web app",
 "3. Доступ: Все": "3. Access: Anyone",
 "3. Лист \"Базовая информация\" → раздел <b>ПОСТАВЩИКИ</b>": "3. Sheet \"Base info\" → section <b>SUPPLIERS</b>",
 "3. Найдите \"chat\":{\"id\": и скопируйте число:": "3. Find \"chat\":{\"id\": and copy the number:",
 "3. Найдите это сообщение, рядом будет поле": "3. Find that message; next to it there is a field",
 "4. Введите это число:": "4. Enter that number:",
 "4. Найдите строку вашей компании-поставщика": "4. Find the row of your supplier company",
 "4. Скопируйте URL и повторите эту команду": "4. Copy the URL and run this command again",
 "5. Вставьте скопированный ID в колонку <b>\"Telegram Chat ID\"</b>": "5. Paste the copied ID into the <b>\"Telegram Chat ID\"</b> column",
 "<b>ЗАКАЗ": "<b>ORDER",
 "<b>ПОСТАВКА ПОДТВЕРЖДЕНА</b>": "<b>DELIVERY CONFIRMED</b>",
 "<div class=\"h\"><h2>Форма обновлена!</h2></div>": "<div class=\"h\"><h2>Form updated!</h2></div>",
 "<div class=\"note\">✅ Ссылка постоянная — не меняется!</div>": "<div class=\"note\">✅ The link is permanent and does not change!</div>",
 "=== Лист \"Продукты\" ===": "=== Sheet \"Food\" ===",
 "=== Лист \"Склад_Продукты\" ===": "=== Sheet \"Warehouse_Food\" ===",
 "AI вернул не массив": "AI did not return an array",
 "AI вернул неожиданный ответ (finishReason=": "AI returned an unexpected response (finishReason=",
 "AI вернул нераспознаваемый ответ (код": "AI returned an unreadable response (code",
 "AI не смог классифицировать": "AI could not classify it",
 "AI-классификатор не настроен": "AI classifier is not set up",
 "Chat ID логистики:": "Logistics chat ID:",
 "Chat ID ответственного (хаб):": "Hub manager chat ID:",
 "ID офиса:": "Office ID:",
 "NO  — Удалить единицу из списка": "NO  — Remove a unit from the list",
 "NO  — Хаб (только хозяюшка, еженедельный пересчёт остатков)": "NO  — Hub (housekeeper only, weekly stock count)",
 "TELEGRAM — ЧАТ ЛОГИСТИКИ": "TELEGRAM — LOGISTICS CHAT",
 "Telegram чат не подключён!": "Telegram chat is not connected!",
 "YES — Добавить новую единицу": "YES — Add a new unit",
 "YES — Полный склад (менеджер/логист, ежедневный цикл)": "YES — Full warehouse (manager/logistician, daily cycle)",
 "[склад ещё не создан]": "[warehouse not created yet]",
 "api.telegram.org/bot<ВАШ_ТОКЕН>/getUpdates": "api.telegram.org/bot<YOUR_TOKEN>/getUpdates",
 "api.telegram.org/bot<ТОКЕН>/getUpdates": "api.telegram.org/bot<TOKEN>/getUpdates",
 "|  Артикул:": "|  SKU:",
 "Активен": "Enabled",
 "Активность": "Active",
 "Активных кофе-поинтов:": "Active coffee points:",
 "Артикул": "SKU",
 "Архивировать все склады сейчас": "Archive all warehouses now",
 "Ассортимент по выбору": "Assortment of choice",
 "БАЗОВАЯ ИНФОРМАЦИЯ ОФИСА — обновлено": "OFFICE BASE INFO — updated",
 "Базовая информация": "Base info",
 "Без поставщика": "No supplier",
 "Бот: @": "Bot: @",
 "Будут установлены:": "The following will be installed:",
 "ВЫБОР АССОРТИМЕНТА НА НЕДЕЛЮ:": "ASSORTMENT FOR THE WEEK:",
 "ВЫДАНО": "ISSUED",
 "ВЫДАНО:            col": "ISSUED:            col",
 "Ваше имя или подразделение": "Your name or department",
 "Введите Chat ID общего чата (или супергруппы, если у вас темы)": "Enter the Chat ID of the shared chat (or of the supergroup, if you use topics)",
 "Введите название новой единицы измерения:": "Enter the name of the new unit:",
 "Введите номер\n(пусто = в самый конец):": "Enter the number\n(empty = at the very end):",
 "Введите номер для удаления:": "Enter the number to remove:",
 "Введите токен вашего бота\n(получите у @BotFather в Telegram):": "Enter your bot token\n(get it from @BotFather in Telegram):",
 "Введите число (0, если товара нет)": "Enter a number (0 if out of stock)",
 "Введите число от 1 до": "Enter a number from 1 to",
 "Верни РОВНО": "Return EXACTLY",
 "Воскресенье": "Sunday",
 "Время": "Time",
 "Вс": "Sun",
 "Вставить ПЕРЕД какой категорией?": "Insert BEFORE which category?",
 "Вставьте ключ Google Gemini API (начинается с \"AIza...\").": "Paste your Google Gemini API key (starts with \"AIza...\").",
 "Вт": "Tue",
 "Вторник": "Tuesday",
 "Выберите ваш кофе-поинт": "Choose your coffee point",
 "Выдано": "Issued",
 "Выдать со склада": "Issue from stock",
 "Выдача подтверждена:": "Issue confirmed:",
 "Выдача со склада —": "Issue from stock —",
 "Выдаётся": "Issued to staff",
 "Город:": "City:",
 "Группа": "Group",
 "Данные": "Data:",
 "Данные будут удалены.": "The data will be deleted.",
 "Данные канцелярия склада на": "Stationery warehouse data as of",
 "Данные появятся после первых заказов": "Data will appear after the first orders",
 "Данные хозтовары склада на": "Household warehouse data as of",
 "Дата": "Date",
 "Дата запуска:": "Launch date:",
 "Дата поставки": "Delivery date",
 "Дашборд Канцелярия обновлён.": "Stationery dashboard updated.",
 "Дашборд Медикаменты обновлён.": "Medicine dashboard updated.",
 "Дашборд Продукты обновлён.": "Food dashboard updated.",
 "Дашборд Хозтовары обновлён.": "Household dashboard updated.",
 "День": "Day",
 "Для каждого товара определи:": "For each item determine:",
 "Для какого склада создать структуру?": "Which warehouse should the structure be created for?",
 "Для подключения Telegram:": "To connect Telegram:",
 "Для этого склада сохранение не настроено": "Saving is not set up for this warehouse",
 "Добавить \"": "Add \"",
 "Добавить единицу": "Add unit",
 "Добавить категорию": "Add category",
 "Добавить категорию (на текущем складе)": "Add category (current warehouse)",
 "Добавить поставщика": "Add supplier",
 "Добавление строк": "Add rows",
 "Добавление строк (на текущем складе)": "Add rows (current warehouse)",
 "Добавлено на склад:": "Added to the warehouse:",
 "Добавьте раздел ПОСТАВЩИКИ в базовую информацию.": "Add a SUPPLIERS section to the base info.",
 "Дозаполнено значений по умолчанию:": "Default values filled in:",
 "Единица": "Unit",
 "Единица \"": "Unit \"",
 "Единица измерения": "Unit",
 "Единицы измерения": "Units",
 "Если ваш общий чат логиста — это отдельная ТЕМА (топик)": "If your shared logistics chat is a separate TOPIC",
 "Если общий чат — это ТЕМА внутри супергруппы,": "If the shared chat is a TOPIC inside a supergroup,",
 "Если поставка уже была подтверждена сегодня — новых": "If the delivery was already confirmed today, there are no new",
 "Если товара нет совсем — впишите 0.": "If an item is completely out, enter 0.",
 "Если это обычная отдельная группа — нажмите NO.": "If it is a regular separate group, press NO.",
 "Если это тема внутри супергруппы — не забудьте:": "If this is a topic inside a supergroup, do not forget:",
 "Заголовки": "Headers",
 "Заголовок": "Title",
 "Заказ": "Order",
 "Заказ канцелярии": "Stationery order",
 "Заказ канцелярии — хозяюшки": "Stationery order — housekeepers",
 "Заказ медикаментов": "Medicine order",
 "Заказ медикаментов — хозяюшки": "Medicine order — housekeepers",
 "Заказ на неделю": "Weekly order",
 "Заказ продуктов": "Food order",
 "Заказ продуктов — форма хозяюшек": "Food order — housekeepers form",
 "Заказ продуктов — хозяюшки": "Food order — housekeepers",
 "Заказ хозтоваров": "Household order",
 "Заказ хозтоваров — клининг": "Household order — cleaning",
 "Заказ хозтоваров — хозяюшки": "Household order — housekeepers",
 "Заказ_": "Order_",
 "Заказано:": "Ordered:",
 "Заказов за сегодня ещё нет": "No orders today yet",
 "Заказы": "Orders",
 "Заказы поставщиков без своего Telegram будут приходить сюда 🛒": "Orders for suppliers without their own Telegram will arrive here 🛒",
 "Заказы → Канцелярия → Обновить форму хозяюшек": "Orders → Stationery → Update housekeepers form",
 "Заказы → Медикаменты → Обновить форму хозяюшек": "Orders → Medicine → Update housekeepers form",
 "Заказы → Обновить лист заказа": "Orders → Update order sheet",
 "Заказы → Продукты → Обновить лист заказа": "Orders → Food → Update order sheet",
 "Заказы → Продукты → Обновить форму хозяюшек": "Orders → Food → Update housekeepers form",
 "Заказы → Продукты → Отправить заказ (Telegram)": "Orders → Food → Send order (Telegram)",
 "Заказы → Хозтовары → Обновить форму хозяюшек\nЗаказы → Хозтовары → Обновить форму клининга": "Orders → Household → Update housekeepers form\nOrders → Household → Update cleaning form",
 "Записать данные за": "Record the data for",
 "Заполните данные офиса.": "Fill in the office details.",
 "Заполните количество и нажмите:": "Fill in the quantities and choose:",
 "Заполните форму логиста — данные появятся в дашборде": "Fill in the logistician form; the data will appear on the dashboard",
 "Запустить архивацию и сброс дашбордов?": "Run the archive and reset the dashboards?",
 "ИДЕНТИФИКАЦИЯ ОФИСА": "OFFICE IDENTIFICATION",
 "ИТОГО": "TOTAL",
 "ИТОГО заказов": "Total orders",
 "ИТОГО:             col": "TOTAL:             col",
 "Источник": "Source",
 "КОФЕ-ПОИНТЫ": "COFFEE POINTS",
 "КП": "CP",
 "Каждую неделю, когда хозяюшка сдаёт остатки, черновик заказа": "Every week, when the housekeeper submits the stock count, the draft order",
 "Как подключить: добавьте бота в чат поставщика →": "How to connect: add the bot to the supplier chat →",
 "Как получить Chat ID:": "How to get the Chat ID:",
 "Как получить:": "How to get it:",
 "Как узнать chat id: напишите /start вашему боту в личку,": "How to find the chat id: send /start to your bot in a private chat,",
 "Канцелярия": "Stationery",
 "Категорий:": "Categories:",
 "Категория": "Category",
 "Клининг": "Cleaning",
 "Клининг:": "Cleaning:",
 "Ключ используется только для распределения товаров": "The key is only used to sort items",
 "Количество": "Quantity",
 "Количество: col": "Quantity: col",
 "Колонка \"ВЫДАНО\" не найдена в": "Column \"ISSUED\" not found in",
 "Колонка \"Количество\" не найдена в Склад_Продукты": "Column \"Quantity\" not found in Warehouse_Food",
 "Колонка \"Остаток на складе\" не найдена": "Column \"In stock\" not found",
 "Колонка \"Остаток на складе\" не найдена в": "Column \"In stock\" not found in",
 "Комментарий": "Comment",
 "Конец дня (архив + сброс)": "End of day (archive + reset)",
 "Конструктор меню": "Menu builder",
 "Контакт": "Contact",
 "Кофе-поинт:": "Coffee point:",
 "Круглосуточно стоят:": "Always available:",
 "Круглосуточно:": "Always available:",
 "Лист \"": "Sheet \"",
 "Лист \"Базовая информация\" не найден!": "Sheet \"Base info\" not found!",
 "Лист \"Базовая информация\" уже существует": "Sheet \"Base info\" already exists",
 "Лист \"Продукты\" не найден": "Sheet \"Food\" not found",
 "Лист \"Продукты\" не найден!": "Sheet \"Food\" not found!",
 "Лист \"Склад_Канцелярия\" не найден": "Sheet \"Warehouse_Stationery\" not found",
 "Лист \"Склад_Медикаменты\" не найден": "Sheet \"Warehouse_Medicine\" not found",
 "Лист \"Склад_Продукты\" не найден": "Sheet \"Warehouse_Food\" not found",
 "Лист \"Склад_Продукты\" не найден.": "Sheet \"Warehouse_Food\" not found.",
 "Лист Склад_Продукты не найден": "Sheet Warehouse_Food not found",
 "Лист уже существует. Данные будут удалены.": "The sheet already exists. Its data will be deleted.",
 "Массово добавить ассортимент": "Bulk add assortment",
 "Медикаменты": "Medicine",
 "Менеджер для подтверждения заказов": "Order approval manager",
 "Меню": "Menu",
 "Меню на неделю с": "Menu for the week from",
 "Меню не найдено": "Menu not found",
 "Меню сегодня": "Today's menu",
 "НЕ НАЙДЕН": "NOT FOUND",
 "Нажмите: Настройки → Обновить единицы на складах": "Choose: Settings → Update units on warehouses",
 "Название категории:": "Category name:",
 "Название кофе-поинта": "Coffee point name",
 "Название не может быть пустым": "The name cannot be empty",
 "Название офиса:": "Office name:",
 "Название поставщика": "Supplier name",
 "Название продукта": "Product name",
 "Название чата:": "Chat name:",
 "Настроить Telegram": "Set up Telegram",
 "Настроить Webhook": "Set up webhook",
 "Настроить тему (топик) чата": "Set chat topic",
 "Настройки": "Settings",
 "Настройки → Настроить Telegram → вставьте ID": "Settings → Set up Telegram → paste the ID",
 "Настройки → Создать структуру склада": "Settings → Create warehouse structure",
 "Начало дня: перенесено": "Start of day: carried over",
 "Не найдена таблица кофе-поинтов.": "Coffee points table not found.",
 "Не настроен ключ AI (Gemini). ⚙️ Настройки → AI → Подключить AI-классификатор": "AI key (Gemini) is not set. ⚙️ Settings → AI → Connect AI classifier",
 "Не подключён": "Not connected",
 "Не удалось освободить лист \"": "Could not free the sheet \"",
 "Не удалось разобрать JSON от AI:": "Could not parse JSON from AI:",
 "Неверный номер": "Invalid number",
 "Недостачи, замены, повреждения, замечания по качеству": "Shortages, substitutions, damage, quality notes",
 "Неизвестный режим офиса:": "Unknown office mode:",
 "Неизвестный тип склада:": "Unknown warehouse type:",
 "Неизвестный тип:": "Unknown type:",
 "Нет активных кофе-поинтов": "No active coffee points",
 "Нет активных товаров.": "No active items.",
 "Нет ни одного активного товара ни на одном складе.": "There are no active items in any warehouse.",
 "Нет позиций с заполненным количеством.": "No items have a quantity filled in.",
 "Нет позиций с заполненным количеством. Сначала укажите количество.": "No items have a quantity filled in. Enter the quantities first.",
 "Нет товаров для страницы 1.": "No items for page 1.",
 "Нет товаров для страницы 2.\nВыберите ассортимент в листе меню \"": "No items for page 2.\nChoose the assortment in the menu sheet \"",
 "Нет товаров с отметкой \"Клининг\".": "No items marked \"Cleaning\".",
 "Нет товаров с отметкой \"Хозяюшки\".": "No items marked \"Housekeepers\".",
 "Новый лист ответов \"": "New response sheet \"",
 "Обновить ассортимент в меню": "Refresh menu assortment",
 "Обновить дашборд": "Refresh dashboard",
 "Обновить единицы на складах": "Update units on warehouses",
 "Обновить лист заказа": "Update order sheet",
 "Обновить поставщиков на складах": "Update suppliers on warehouses",
 "Обновить форму клининга": "Update cleaning form",
 "Обновить форму остатков": "Update stock form",
 "Обновить форму хозяюшек": "Update housekeepers form",
 "Обновлено": "Updated",
 "Обновлено строк:": "Rows updated:",
 "Обработано товаров:": "Items processed:",
 "Оставьте поле пустым, если не поставлено.": "Leave the field empty if it was not delivered.",
 "Остатки": "Leftovers",
 "Остатки на складе — форма хозяюшки": "Stock count — housekeeper form",
 "Остаток был": "Stock before",
 "Остаток конец дня": "Stock at end of day",
 "Остаток на складе": "In stock",
 "Остаток на складе: col": "In stock: col",
 "Остаток стал": "Stock after",
 "Ответь ТОЛЬКО JSON-массивом, без пояснений и без markdown-разметки (без ```):": "Reply ONLY with a JSON array, no explanations and no markdown (no ```):",
 "Откройте конструктор заново.": "Open the builder again.",
 "Отправить заказ (Telegram)": "Send order (Telegram)",
 "Отсутствует": "Not available",
 "Офис": "Office",
 "Ошибка Telegram: сообщения не отправились ни одному поставщику": "Telegram error: messages were not sent to any supplier",
 "Ошибка:": "Error:",
 "ПОЛНЫЙ (менеджер + логист)": "FULL (manager + logistician)",
 "ПОСТАВЩИКИ": "SUPPLIERS",
 "Первоначальная настройка": "Initial setup",
 "Перейдите: ⚙️ Настройки → 🔗 Настроить Telegram": "Go to: ⚙️ Settings → 🔗 Set up Telegram",
 "Пересоздайте лист: Настройки → Создать \"Базовую информацию\"": "Recreate the sheet: Settings → Create \"Base info\"",
 "Пересоздать склад канцелярии?": "Recreate the stationery warehouse?",
 "Пересоздать?": "Recreate?",
 "Пересоздать? Все данные будут удалены.": "Recreate? All data will be deleted.",
 "Пересоздать? Выбранный ассортимент будет сброшен.": "Recreate? The selected assortment will be reset.",
 "Период:": "Period:",
 "Пн": "Mon",
 "По выбору:": "Of choice:",
 "Подключить AI-классификатор": "Connect AI classifier",
 "Подтвердите полученные товары. Период:": "Confirm the items received. Period:",
 "Подтвердите фактически полученные товары.\n\nДата:": "Confirm the items actually received.\n\nDate:",
 "Подтвердить поставку": "Confirm delivery",
 "Подтверждение поставки —": "Delivery confirmation —",
 "Позиций:": "Items:",
 "Позиция": "Item",
 "Получить ключ: aistudio.google.com → Get API key": "Get a key: aistudio.google.com → Get API key",
 "Пользовательские единицы:": "Custom units:",
 "Пользовательских единиц нет.\n\nБазовые:": "There are no custom units.\n\nBuilt-in:",
 "Понедельник": "Monday",
 "После этого заказы для вас будут приходить прямо в этот чат! 🛒": "After that, your orders will arrive right in this chat! 🛒",
 "Поставка": "Delivery",
 "Поставка канцелярии — логист": "Stationery delivery — logistician",
 "Поставка канцелярии — форма логиста": "Stationery delivery — logistician form",
 "Поставка медикаментов — логист": "Medicine delivery — logistician",
 "Поставка медикаментов — форма логиста": "Medicine delivery — logistician form",
 "Поставка продуктов — логист": "Food delivery — logistician",
 "Поставка продуктов — форма логиста": "Food delivery — logistician form",
 "Поставка хозтоваров — логист": "Household delivery — logistician",
 "Поставка хозтоваров — форма логиста": "Household delivery — logistician form",
 "Поставка:          col": "Delivery:          col",
 "Поставщик": "Supplier",
 "Поставщик / водитель": "Supplier / driver",
 "Поставщик / водитель (ФИО или название)": "Supplier / driver (name or company)",
 "Поставщикам заказ уходит только после этого.": "The order goes to suppliers only after that.",
 "Поставщики": "Suppliers",
 "Поставщиков:": "Suppliers:",
 "Починить строки склада": "Repair warehouse rows",
 "Предложение (авто)": "Suggestion (auto)",
 "Предложения обновлены для": "Suggestions updated for",
 "Примечание": "Note",
 "Проверено строк-товаров:": "Item rows checked:",
 "Проверьте Thread ID и что бот состоит в супергруппе.": "Check the Thread ID and that the bot is a member of the supergroup.",
 "Проверьте склады и при необходимости поправьте вручную": "Check the warehouses and fix manually if needed",
 "Проверьте токен": "Check the token",
 "Проверьте что бот добавлен в чат и Chat ID верный.": "Check that the bot is added to the chat and the Chat ID is correct.",
 "Проверьте, не удалена ли форма в корзину.": "Check whether the form was moved to the trash.",
 "Продолжить?": "Continue?",
 "Продукты": "Food",
 "Пропущено строк (категория не распознана):": "Rows skipped (category not recognized):",
 "Пт": "Fri",
 "Пятница": "Friday",
 "Раз в неделю проверьте физический остаток по каждой позиции и впишите число.": "Once a week, check the physical stock of each item and enter the number.",
 "Разбейте на части по": "Split it into parts of",
 "Рандомайзер работает только со сменными категориями,": "The randomizer only works with rotating categories",
 "Расход за неделю": "Used this week",
 "С": "From",
 "С указанным остатком:": "With stock specified:",
 "СКЛАДЫ И ИХ СУЩЕСТВУЮЩИЕ КАТЕГОРИИ:": "WAREHOUSES AND THEIR EXISTING CATEGORIES:",
 "Сб": "Sat",
 "Склад": "Warehouse",
 "Склад \"": "Warehouse \"",
 "Склад хозтоваров создан!\n\nG — Хозяюшки\nH — Клининг": "Household warehouse created!\n\nG — Housekeepers\nH — Cleaning",
 "Склад_Канцелярия": "Warehouse_Stationery",
 "Склад_Медикаменты": "Warehouse_Medicine",
 "Склад_Продукты": "Warehouse_Food",
 "Склад_Хозтовары": "Warehouse_Household",
 "Сколько строк добавить? (1–50)": "How many rows to add? (1–50)",
 "Скрыто листов:": "Sheets hidden:",
 "Скрыто неиспользуемых листов (дашборды/БД):": "Unused sheets hidden (dashboards/DB):",
 "Скрыть неиспользуемые листы (дашборды/БД)": "Hide unused sheets (dashboards/DB)",
 "Слишком много позиций за один раз (": "Too many items at once (",
 "Сначала опубликуйте скрипт:": "Deploy the script first:",
 "Сначала перейдите на вкладку листа меню": "First switch to the menu sheet tab",
 "Сначала создайте заказ:": "Create the order first:",
 "Сначала создайте меню!\nМеню → Создать меню на следующую неделю": "Create a menu first!\nMenu → Create next week menu",
 "Сначала создайте структуру склада:": "Create the warehouse structure first:",
 "Сначала сформируйте заказ:": "Build the order first:",
 "Снято чередующихся цветов (banding):": "Alternating colors removed (banding):",
 "Создайте его: Настройки → Создать \"Базовую информацию\"": "Create it: Settings → Create \"Base info\"",
 "Создание структуры склада": "Create warehouse structure",
 "Создать \"Базовую информацию\"": "Create \"Base info\"",
 "Создать меню на следующую неделю": "Create next week menu",
 "Создать структуру склада": "Create warehouse structure",
 "Сохранить меню для планшетов": "Save menu for tablets",
 "Список пуст — не удалось распознать ни одной строки.": "The list is empty: no rows could be recognized.",
 "Ср": "Wed",
 "Среда": "Wednesday",
 "Ссылка для хозяюшек (продукты)": "Link for housekeepers (food)",
 "Ссылки на формы": "Form links",
 "Стандарт:": "Standard:",
 "Статус Telegram": "Telegram status",
 "Статус:": "Status:",
 "Стр. 1 — стандарт | Стр. 2 — по выбору меню": "Page 1 — standard | Page 2 — menu choice",
 "Строк:": "Rows:",
 "Строка заголовков:": "Header row:",
 "Суббота": "Saturday",
 "ТОВАРЫ ДЛЯ РАСПРЕДЕЛЕНИЯ (": "ITEMS TO SORT (",
 "Таймингов:": "Timings:",
 "Текущий остаток,": "Current stock,",
 "Текущий режим:": "Current mode:",
 "Текущий список:": "Current list:",
 "Теперь бот будет автоматически присылать Chat ID": "Now the bot will automatically send the Chat ID",
 "Тестовое сообщение отправлено в неё.": "A test message was sent to it.",
 "Тип объекта": "Site type",
 "Тип объекта (Полный / Хаб)": "Site type (Full / Hub)",
 "Товар": "Product",
 "Токен Telegram-бота не настроен!": "Telegram bot token is not set!",
 "Токен не может быть пустым": "The token cannot be empty",
 "Ты помогаешь разложить список товаров склада по 4 складам и категориям внутри каждого склада.": "You help sort a list of warehouse items into 4 warehouses and the categories inside each warehouse.",
 "Удалить единицу": "Remove unit",
 "Укажите необходимое количество.": "Enter the quantity you need.",
 "Укажите фактически полученное количество. Если товар не поставлен — оставьте пустым.": "Enter the quantity actually received. If an item was not delivered, leave it empty.",
 "Управление единицами": "Manage units",
 "Управление складами": "Warehouses",
 "Установить все триггеры": "Install all triggers",
 "Установить триггеры": "Install triggers",
 "Форма \"": "Form \"",
 "Форма клининга обновлена!": "Cleaning form updated!",
 "Форма логиста:": "Logistician form:",
 "Форма обновлена": "Form updated",
 "Форма хозяюшек (канцелярия) обновлена!": "Housekeepers form (stationery) updated!",
 "Форма хозяюшек (медикаменты) обновлена!": "Housekeepers form (medicine) updated!",
 "Форма хозяюшек обновлена!": "Housekeepers form updated!",
 "Форму хозяюшек обновите вручную:": "Update the housekeepers form manually:",
 "Формы Канцелярия:\n\nЛогист:": "Stationery forms:\n\nLogistician:",
 "Формы Медикаменты:\n\nЛогист:": "Medicine forms:\n\nLogistician:",
 "Формы Продукты:\n\nЛогист:": "Food forms:\n\nLogistician:",
 "Формы Хозтовары:\n\nЛогист:": "Household forms:\n\nLogistician:",
 "Формы хозяюшек обновите вручную:": "Update the housekeepers forms manually:",
 "Функция устарела: дашборд Продуктов теперь всегда синхронен со": "This function is obsolete: the Food dashboard is now always in sync with the",
 "ХАБ (упрощённый)": "HUB (simplified)",
 "Хаб — остатки и заказы": "Hub — stock & orders",
 "Хаб_История_движения_остатков": "Hub_Stock_history",
 "Хаб_Целевые_остатки": "Hub_Target_stock",
 "Хозтовары": "Household",
 "Хозяюшки": "Housekeepers",
 "Хозяюшки:": "Housekeepers:",
 "Целевой остаток": "Target stock",
 "Чекбоксы, единицы измерения и список поставщиков": "Checkboxes, units and the supplier list",
 "Четверг": "Thursday",
 "Чт": "Thu",
 "Что выкладываем": "What to put out",
 "Чтобы автоматически распределять товары по складам и категориям,": "To sort items into warehouses and categories automatically,",
 "Этаж/Зона": "Floor/Zone",
 "Это только рекомендация — чтобы применить, скопируйте значение": "This is only a suggestion; to apply it, copy the value",
 "августа": "August",
 "апреля": "April",
 "бот напишет Chat ID → скопируйте ID →": "the bot will post the Chat ID → copy the ID →",
 "в колонке \"Поставка\".": "in the \"Delivery\" column.",
 "вид": "type",
 "внутри большой супергруппы с несколькими темами": "inside a large supergroup with several topics",
 "вс": "Sun",
 "вт": "Tue",
 "вторник": "Tuesday",
 "даже если не до конца уверен — в этом случае выбери наиболее вероятный вариант.": "even if you are not fully sure; in that case pick the most likely option.",
 "дашборда": "dashboard",
 "декабря": "December",
 "добавлен!\n\nВведите название в выделенную ячейку.": "added!\n\nEnter the name in the selected cell.",
 "затем откройте api.telegram.org/bot<ТОКЕН>/getUpdates": "then open api.telegram.org/bot<TOKEN>/getUpdates",
 "из колонки \"Предложение (авто)\" в колонку \"Целевой остаток\".": "from the \"Suggestion (auto)\" column to the \"Target stock\" column.",
 "используй её название ТОЧНО как в списке выше (посимвольно). Если ни одна не подходит —": "use its name EXACTLY as in the list above (character for character). If none fits,",
 "июля": "July",
 "июня": "June",
 "канцелярии": "stationery",
 "когда его добавляют в чат.": "when it is added to a chat.",
 "марта": "March",
 "мая": "May",
 "медикаментов": "medicine",
 "на неделю:": "for the week:",
 "настроенными через Конструктор меню.": "set up in the Menu builder.",
 "не используется": "not used",
 "не найден": "not found",
 "не найден.": "not found.",
 "не найдена]": "not found]",
 "не настроен]": "not set up]",
 "не обновилась": "not updated",
 "не создана": "not created",
 "недоступна": "unavailable",
 "непроверенных позиций нет.": "unchecked items.",
 "ноября": "November",
 "нужен ключ Google Gemini API.\n\nНастроить сейчас?": "a Google Gemini API key is needed.\n\nSet it up now?",
 "объектов, СТРОГО в том же порядке, что и список товаров выше,": "objects, STRICTLY in the same order as the item list above,",
 "октября": "October",
 "остатков из склада": "stock values from the warehouse",
 "отправлен!": "sent!",
 "пересозданы для всех строк на этом складе.": "were recreated for all rows in this warehouse.",
 "пн": "Mon",
 "по": "to",
 "по складам и категориям при массовой вставке ассортимента.": "into warehouses and categories during bulk assortment import.",
 "позиций": "items",
 "позиций добавлено к остаткам.\nУведомления отправлены поставщикам.": "items added to stock.\nNotifications sent to suppliers.",
 "позиций).": "items).",
 "понедельник": "Monday",
 "после этого шага зайдите в:": "after this step go to:",
 "придумай короткое точное название новой категории на русском (в том же стиле, что и существующие).": "make up a short, precise name for a new category in English (in the same style as the existing ones).",
 "пт": "Fri",
 "пятница": "Friday",
 "с": "from",
 "сб": "Sat",
 "сентября": "September",
 "склад \"": "warehouse \"",
 "склада на": "warehouse as of",
 "складом, отдельный перенос остатков в конец дня не требуется.": "warehouse; a separate end-of-day stock transfer is not needed.",
 "сначала уходит СЮДА в Telegram — с кнопками:": "first goes HERE in Telegram, with buttons:",
 "ср": "Wed",
 "сразу и автоматически, без подтверждения.": "immediately and automatically, without approval.",
 "среда": "Wednesday",
 "строк в категорию \"": "rows to category \"",
 "товаров.": "items.",
 "у хаба нет ежедневного цикла выдач/поставок, только еженедельные остатки)": "a hub has no daily issue/delivery cycle, only weekly stock counts)",
 "февраля": "February",
 "хозтоваров": "household",
 "четверг": "Thursday",
 "чт": "Thu",
 "шт": "pcs",
 "шт.):": "pcs):",
 "штук и вставляйте по очереди.": "items and paste them one by one.",
 "этот чат": "this chat",
 "января": "January",
 "— существующие категории:": "— existing categories:",
 "• 00:05 — обновление дат": "• 00:05 — date update",
 "• 23:30 — конец дня (архив)\n• 23:59 — сброс дашбордов": "• 23:30 — end of day (archive)\n• 23:59 — dashboard reset",
 "• onFormSubmit — автообновление": "• onFormSubmit — auto-refresh of",
 "• Выберите ассортимент в новом листе": "• Choose the assortment in the new sheet",
 "• Пн 00:05 — активация меню следующей недели": "• Mon 00:05 — activate next week's menu",
 "• Пн 00:15 — удаление старых листов меню": "• Mon 00:15 — delete old menu sheets",
 "• После выбора нажмите: Меню → Сохранить меню для планшетов": "• After choosing, select: Menu → Save menu for tablets",
 "• Старое меню остаётся до конца текущей недели": "• The old menu stays until the end of the current week",
 "… и ещё": "… and",
 "ℹ️ Скрывать нечего — такие листы уже скрыты или их нет.": "ℹ️ Nothing to hide: these sheets are already hidden or do not exist.",
 "⚙️ Настройки → Telegram → Настроить тему (топик) чата": "⚙️ Settings → Telegram → Set chat topic",
 "⚙️ Настройки → 🔗 Настроить Telegram": "⚙️ Settings → 🔗 Set up Telegram",
 "⚠️ Chat ID сохранён, но тестовое сообщение не отправилось.": "⚠️ Chat ID saved, but the test message was not sent.",
 "⚠️ Telegram не отправлен.": "⚠️ Telegram message not sent.",
 "⚠️ В этом меню нет раздела \"Выбор ассортимента на неделю\"": "⚠️ This menu has no \"Assortment for the week\" section",
 "⚠️ Если не настроить — заказ будет уходить поставщикам": "⚠️ If this is not set up, the order goes to suppliers",
 "⚠️ Конструктор меню не настроен.": "⚠️ The menu builder is not set up.",
 "⚠️ На складе нет категорий. Пересоздайте склад.": "⚠️ The warehouse has no categories. Recreate the warehouse.",
 "⚠️ Не найден блок TELEGRAM в \"Базовой информации\" — добавьте строку вручную.": "⚠️ TELEGRAM block not found in \"Base info\"; add the row manually.",
 "⚠️ Не найдено ни одной заполняемой позиции ассортимента.": "⚠️ No fillable assortment slots found.",
 "⚠️ Не удалось распределить (": "⚠️ Could not sort (",
 "⚠️ Поставщики не найдены — добавьте их в \"Базовую информацию\"": "⚠️ No suppliers found; add them to \"Base info\"",
 "⚠️ Тестовое сообщение не отправилось:": "⚠️ The test message was not sent:",
 "✅ <b>Telegram подключён!</b>": "✅ <b>Telegram connected!</b>",
 "✅ <b>Бот успешно добавлен!</b>": "✅ <b>Bot added successfully!</b>",
 "✅ Telegram успешно подключён!": "✅ Telegram connected successfully!",
 "✅ Webhook настроен!": "✅ Webhook set up!",
 "✅ Архивация завершена!": "✅ Archive complete!",
 "✅ Ассортимент обновлён на листах:": "✅ Assortment refreshed on sheets:",
 "✅ Бот активен: @": "✅ Bot active: @",
 "✅ Готово!": "✅ Done!",
 "✅ Данные меню сохранены!\n\nПланшеты обновятся при следующем открытии.": "✅ Menu data saved!\n\nTablets will update next time they open.",
 "✅ Добавлено": "✅ Added",
 "✅ Единица \"": "✅ Unit \"",
 "✅ Единицы обновлены!\n\nСписок:": "✅ Units updated!\n\nList:",
 "✅ Загружено поставщиков:": "✅ Suppliers loaded:",
 "✅ Заказ": "✅ Order",
 "✅ Заказ отправлен в Telegram!": "✅ Order sent to Telegram!",
 "✅ Категория \"": "✅ Category \"",
 "✅ Ключ сохранён.": "✅ Key saved.",
 "✅ Кнопки созданы для": "✅ Buttons created for",
 "✅ Конец дня выполнен!": "✅ End of day complete!",
 "✅ Конфиг исправлен!": "✅ Config fixed!",
 "✅ Кофе-поинт #": "✅ Coffee point #",
 "✅ Лист \"Базовая информация\" создан!": "✅ Sheet \"Base info\" created!",
 "✅ Меню на следующую неделю создано!": "✅ Next week menu created!",
 "✅ Меню на текущую неделю создано!": "✅ Current week menu created!",
 "✅ Мёртвых триггеров не найдено": "✅ No dead triggers found",
 "✅ Остатки обновлены": "✅ Stock updated",
 "✅ Отправить · ✏️ Скорректировать · ❌ Отменить.": "✅ Send · ✏️ Adjust · ❌ Cancel.",
 "✅ Подключён (Chat ID:": "✅ Connected (Chat ID:",
 "✅ Поставка подтверждена!": "✅ Delivery confirmed!",
 "✅ Привязка к теме отключена.\n\nСообщения будут уходить в общий чат как обычно.": "✅ Topic binding turned off.\n\nMessages will go to the shared chat as usual.",
 "✅ Рандомайзер заполнил ассортимент!": "✅ The randomizer filled in the assortment!",
 "✅ Режим: Полный склад": "✅ Mode: Full warehouse",
 "✅ Режим: Хаб": "✅ Mode: Hub",
 "✅ Склад канцелярии создан!": "✅ Stationery warehouse created!",
 "✅ Скрыто листов:": "✅ Sheets hidden:",
 "✅ Сохранено:": "✅ Saved:",
 "✅ Список поставщиков обновлён!": "✅ Supplier list updated!",
 "✅ Строка добавлена в \"": "✅ Row added to \"",
 "✅ Структура склада \"": "✅ Warehouse structure \"",
 "✅ Тема подключена!\n\nThread ID:": "✅ Topic connected!\n\nThread ID:",
 "✅ Тема подключена!\n\nЗаказы без своего чата поставщика будут приходить сюда.": "✅ Topic connected!\n\nOrders for suppliers without their own chat will arrive here.",
 "✅ Триггеры установлены!": "✅ Triggers installed!",
 "✅ Удалено триггеров:": "✅ Triggers removed:",
 "✅ Форма остатков обновлена!": "✅ Stock form updated!",
 "✅ Целевые остатки синхронизированы.\nДобавлено новых товаров:": "✅ Target stock synced.\nNew items added:",
 "✏️ Подтвердите полученные товары.\n\nПериод:": "✏️ Confirm the items received.\n\nPeriod:",
 "✏️ Укажите необходимое количество.": "✏️ Enter the quantity you need.",
 "❌ Chat ID не может быть пустым": "❌ Chat ID cannot be empty",
 "❌ Telegram не настроен:": "❌ Telegram is not set up:",
 "❌ Бот недоступен:": "❌ Bot unavailable:",
 "❌ Введите число (message_thread_id)": "❌ Enter a number (message_thread_id)",
 "❌ Введите число от 1 до 50": "❌ Enter a number from 1 to 50",
 "❌ Веб-приложение не опубликовано!": "❌ The web app is not deployed!",
 "❌ Выберите строку внутри категории (или на строке категории)": "❌ Select a row inside a category (or the category row itself)",
 "❌ Категория \"": "❌ Category \"",
 "❌ Ключ не может быть пустым": "❌ The key cannot be empty",
 "❌ Лист": "❌ Sheet",
 "❌ Лист \"Базовая информация\" не найден": "❌ Sheet \"Base info\" not found",
 "❌ Лист \"Базовая информация\" не найден.": "❌ Sheet \"Base info\" not found.",
 "❌ Лист \"Склад_Продукты\" не найден": "❌ Sheet \"Warehouse_Food\" not found",
 "❌ Листы меню не найдены": "❌ No menu sheets found",
 "❌ Название не может быть пустым": "❌ The name cannot be empty",
 "❌ Не найдена таблица кофе-поинтов.": "❌ Coffee points table not found.",
 "❌ Не удалось обновить остатки. Проверьте названия товаров на складе.": "❌ Could not update stock. Check the item names in the warehouse.",
 "❌ Не удалось определить категорию": "❌ Could not determine the category",
 "❌ Не удалось определить категорию.\n\nВыберите строку внутри нужной категории.": "❌ Could not determine the category.\n\nSelect a row inside the category you need.",
 "❌ Неверный выбор": "❌ Invalid choice",
 "❌ Неверный номер.": "❌ Invalid number.",
 "❌ Неверный токен:": "❌ Invalid token:",
 "❌ Неизвестный тип склада:": "❌ Unknown warehouse type:",
 "❌ Откройте лист меню": "❌ Open a menu sheet",
 "❌ Ошибка открытия диалога:": "❌ Error opening the dialog:",
 "❌ Ошибка проверки токена:": "❌ Token check error:",
 "❌ Ошибка регистрации webhook:": "❌ Webhook registration error:",
 "❌ Ошибка:": "❌ Error:",
 "❌ Поставщики не найдены в \"Базовой информации\".": "❌ No suppliers found in \"Base info\".",
 "❌ Раздел ПОСТАВЩИКИ не найден": "❌ SUPPLIERS section not found",
 "❌ Сначала настройте общий чат:\n⚙️ Настройки → Настроить Telegram": "❌ Set up the shared chat first:\n⚙️ Settings → Set up Telegram",
 "❌ Сначала создайте \"Базовую информацию\"": "❌ Create \"Base info\" first",
 "❌ Сначала создайте склад продуктов!": "❌ Create the food warehouse first!",
 "❌ Токен не может быть пустым": "❌ The token cannot be empty",
 "❌ Только для листов склада": "❌ Only for warehouse sheets",
 "❌ Только для складов": "❌ Only for warehouses",
 "❓ Без поставщика": "❓ No supplier",
 "🆔 <b>Ваш Chat ID:</b>": "🆔 <b>Your Chat ID:</b>",
 "🆕 Созданы новые категории (": "🆕 New categories created (",
 "🌙 Конец дня": "🌙 End of day",
 "🎲 Заполнить ассортимент по остаткам": "🎲 Fill assortment from stock",
 "👤 Менеджер хаба (подтверждение заказов)": "👤 Hub manager (order approval)",
 "💊 Подтвердите полученные товары.\n\nПериод:": "💊 Confirm the items received.\n\nPeriod:",
 "💊 Укажите необходимое количество.": "💊 Enter the quantity you need.",
 "💬 Комментарий к поставке": "💬 Delivery comment",
 "💬 Общий Chat ID:": "💬 Shared Chat ID:",
 "💾 Архивация": "💾 Archive",
 "📁 Новая категория": "📁 New category",
 "📊 Позиций у поставщика: <b>": "📊 Items for this supplier: <b>",
 "📋 <b>Инструкция:</b>": "📋 <b>Instructions:</b>",
 "📋 <i>Нет своего Telegram — скопируйте и перешлите вручную</i>": "📋 <i>No own Telegram: copy and forward manually</i>",
 "📋 Массовое добавление ассортимента": "📋 Bulk assortment import",
 "📍 Выберите позицию": "📍 Choose a position",
 "📍 Позиция категории": "📍 Category position",
 "📐 Единицы измерения": "📐 Units",
 "📡 Статус Telegram": "📡 Telegram status",
 "📦 Добавление строк": "📦 Add rows",
 "📦 Позиций принято: <b>": "📦 Items received: <b>",
 "📦 Поставщик: <b>": "📦 Supplier: <b>",
 "📦 Поставщиков со своим чатом:": "📦 Suppliers with their own chat:",
 "📦 Стандартный ассортимент": "📦 Standard assortment",
 "📭 Нет данных о поставке": "📭 No delivery data",
 "📱 Чат:": "📱 Chat:",
 "🔗 Настройка Telegram — шаг 1/2": "🔗 Telegram setup — step 1/2",
 "🔗 Настройка Telegram — шаг 2/2": "🔗 Telegram setup — step 2/2",
 "🗓 По выбору меню (сменные категории)": "🗓 Menu choice (rotating categories)",
 "🛒 <b>ЗАКАЗ ПРОДУКТОВ</b>": "🛒 <b>FOOD ORDER</b>",
 "🤖 Подключение AI-классификатора": "🤖 Connect AI classifier",
 "🧵 Thread ID темы": "🧵 Topic Thread ID",
 "🧵 Общий чат — это тема внутри супергруппы?": "🧵 Is the shared chat a topic inside a supergroup?",
 "🧵 Тема (topic):": "🧵 Topic:",
 "🧹 Клининг": "🧹 Cleaning",
 "Как подключить: добавьте бота в чат поставщика → бот напишет Chat ID → скопируйте ID → Настройки → Настроить Telegram → вставьте ID": "How to connect: add the bot to the supplier chat → the bot will post the Chat ID → copy the ID → Settings → Set up Telegram → paste the ID",
 "Переключить язык": "Switch language"
};
const I18N_DATA_ = {
 "Чай": "Tea",
 "Молоко": "Milk",
 "Молоко растительное": "Plant-based milk",
 "Кофе для кофемашин": "Coffee for machines",
 "Напитки растворимые": "Instant drinks",
 "Мед/Джем": "Honey/Jam",
 "Сладкое": "Sweets",
 "Сухие завтраки": "Cereals",
 "Крекеры": "Crackers",
 "Снэки": "Snacks",
 "Бакалея": "Groceries",
 "Масло свежее (холодильник)": "Fresh butter (fridge)",
 "Сиропы": "Syrups",
 "Орех/Сухофрукт": "Nuts/Dried fruit",
 "Шоколад": "Chocolate",
 "Алкоголь": "Alcohol",
 "Первая помощь": "First aid",
 "Обезболивающие": "Painkillers",
 "Антисептики": "Antiseptics",
 "Перевязочные материалы": "Bandages and dressings",
 "Прочее": "Other",
 "Бумага": "Paper",
 "Папки и файлы": "Folders and sleeves",
 "Блокноты": "Notepads",
 "Моющие средства": "Detergents",
 "Чистящие средства": "Cleaning agents",
 "Бумажная продукция": "Paper products",
 "Хозяйственные принадлежности": "Household supplies",
 "Инвентарь": "Equipment",
 "кг": "kg",
 "штука": "piece",
 "упаковка": "pack",
 "блок": "block",
 "л": "l",
 "мл": "ml",
 "г": "g",
 "пакет": "bag"
};
const I18N_HTML_ = {
 "Выдача со склада": "Issue from stock",
 "Нет заказов для выдачи": "No orders to issue",
 "Хозяюшки ещё не подали заявки": "Housekeepers have not submitted requests yet",
 "Позиций:": "Items:",
 "Всего заказано:": "Total ordered:",
 "✓ Выдать всё как заказано": "✓ Issue everything as ordered",
 "✕ Обнулить всё": "✕ Reset all to zero",
 "заказано:": "ordered:",
 "выдано ранее:": "issued before:",
 "новое:": "new:",
 "Выдать сейчас:": "Issue now:",
 "Подтверждение выдачи...": "Confirming issue...",
 "Отмена": "Cancel",
 "✓ Подтвердить выдачу": "✓ Confirm issue",
 "❌ Укажите количество хотя бы для одной позиции": "❌ Enter a quantity for at least one item",
 "Ошибка: ": "Error: ",
 "Подтверждение поставки": "Delivery confirmation",
 "— логист подтверждает получение": "— the logistician confirms receipt",
 "Нет данных о поставке": "No delivery data",
 "Заполните форму логиста — данные появятся в дашборде в колонке \"Поставка\"": "Fill in the logistician form; the data will appear on the dashboard in the \"Delivery\" column",
 "💡 Проверьте и при необходимости скорректируйте количество.": "💡 Check and adjust the quantities if needed.",
 "После подтверждения данные прибавятся к остаткам на складе и уйдут уведомление поставщику в Telegram.": "After confirmation the quantities are added to stock and the supplier gets a Telegram notification.",
 "Итого единиц:": "Total units:",
 "✓ Подтвердить все как есть": "✓ Confirm all as is",
 "из формы:": "from form:",
 "'шт'": "'pcs'",
 "Фактически получено:": "Actually received:",
 "Подтверждение поставки...": "Confirming delivery...",
 "✓ Подтвердить поставку": "✓ Confirm delivery",
 "📋 Массовое добавление ассортимента": "📋 Bulk assortment import",
 "Вставьте товары, по одному на строку.": "Paste items, one per line.",
 "Можно просто название, а можно через \" / \" добавить единицу, остаток на складе": "You can enter just the name, or add the unit, stock quantity",
 "и поставщика — порядок частей неважен, число всегда распознаётся как остаток": "and supplier separated by \" / \". The order does not matter; a number is always read as stock",
 "(запятую внутри названия — например \"2,5%\" — можно не бояться):": "(a comma inside a name, like \"2,5%\", is fine):",
 "<code>Ножницы</code> — только название": "<code>Scissors</code> — name only",
 "<code>Ножницы / штука</code> — название + единица": "<code>Scissors / piece</code> — name + unit",
 "<code>Ножницы / 14 / штук</code> — название + остаток (14) + единица": "<code>Scissors / 14 / pcs</code> — name + stock (14) + unit",
 "<code>Молоко 2,5% / л / ООО Молоко</code> — название + единица + поставщик": "<code>Milk 2.5% / l / Dairy Ltd</code> — name + unit + supplier",
 "AI сам определит склад (Продукты / Хозтовары / Канцелярия / Медикаменты) и категорию,": "AI will pick the warehouse (Food / Household / Stationery / Medicine) and the category,",
 "при необходимости создаст новую категорию на нужном складе.": "creating a new category in the right warehouse if needed.",
 "Молоко 2,5%&#10;Стиральный порошок / кг&#10;Ножницы / 14 / штук&#10;Ручки шариковые / упаковка&#10;Бинт стерильный": "Milk 2.5%&#10;Laundry powder / kg&#10;Scissors / 14 / pcs&#10;Ballpoint pens / pack&#10;Sterile bandage",
 "Закрыть": "Close",
 "Распределить": "Sort",
 "Вставьте хотя бы один товар": "Paste at least one item",
 "Обрабатываю через AI, это может занять до минуты…": "Processing with AI, this may take up to a minute…",
 "'Готово'": "'Done'",
 "<title>Меню</title>": "<title>Menu</title>",
 "📅 Сегодня": "📅 Today",
 "🗓 Вся неделя": "🗓 Whole week",
 "Загружаем меню...": "Loading menu...",
 "Загружаем...": "Loading...",
 "Обновление...": "Updating...",
 "Обновляем...": "Refreshing...",
 "'Обновлено в '": "'Updated at '",
 "На сегодня меню не задано": "No menu set for today",
 "Создайте меню через конструктор": "Create a menu in the builder",
 "Расписание выкладки": "Serving schedule",
 "<th>Время</th><th>Что выкладываем</th>": "<th>Time</th><th>What to put out</th>",
 "<span class=\"now-pill\">Сейчас</span>": "<span class=\"now-pill\">Now</span>",
 "☕ Круглосуточно стоит": "☕ Always available",
 "☕ Круглосуточно": "☕ Always available",
 "Меню на неделю не найдено": "No menu for the week",
 "<span class=\"today-badge\">Сегодня</span>": "<span class=\"today-badge\">Today</span>",
 "Нет данных": "No data",
 "'января'": "'January'",
 "'февраля'": "'February'",
 "'марта'": "'March'",
 "'апреля'": "'April'",
 "'мая'": "'May'",
 "'июня'": "'June'",
 "'июля'": "'July'",
 "'августа'": "'August'",
 "'сентября'": "'September'",
 "'октября'": "'October'",
 "'ноября'": "'November'",
 "'декабря'": "'December'",
 "'Воскресенье'": "'Sunday'",
 "'Понедельник'": "'Monday'",
 "'Вторник'": "'Tuesday'",
 "'Среда'": "'Wednesday'",
 "'Четверг'": "'Thursday'",
 "'Пятница'": "'Friday'",
 "'Суббота'": "'Saturday'",
 "⏳ Загрузка...": "⏳ Loading...",
 "1️⃣ Сменные категории": "1️⃣ Rotating categories",
 "2️⃣ Тайминги и позиции": "2️⃣ Timings and items",
 "3️⃣ Круглосуточно": "3️⃣ Always available",
 "+ Добавить категорию по выбору": "+ Add a rotating category",
 "+ Добавить новый тайминг": "+ Add a new timing",
 "+ Добавить позицию": "+ Add item",
 "💾 Сохранить": "💾 Save",
 "✅ Сохранить и создать меню": "✅ Save and create menu",
 "'Пн','Вт','Ср','Чт','Пт','Сб','Вс'": "'Mon','Tue','Wed','Thu','Fri','Sat','Sun'",
 "'Понедельник','Вторник','Среда','Четверг','Пятница','Суббота','Воскресенье'": "'Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'",
 "Сохраняем...": "Saving...",
 "Черновик сохранён ✓": "Draft saved ✓",
 "Ошибка сохранения: ": "Save error: ",
 "Ошибка сохранения": "Save error",
 "Ошибка загрузки: ": "Load error: ",
 "Данные загружены": "Data loaded",
 "Нет сменных категорий. Нажмите \"+\"": "No rotating categories. Press \"+\"",
 "<b>Категория ${idx+1}</b>": "<b>Category ${idx+1}</b>",
 ">Удалить</button>": ">Delete</button>",
 "Удалить тайминг": "Delete timing",
 "Категория:": "Category:",
 "— выберите —": "— choose —",
 "-- выберите --": "-- choose --",
 "Дни:": "Days:",
 "Режим выбора:": "Selection mode:",
 "Один на всю неделю": "One for the whole week",
 "Несколько видов на всю неделю": "Several types for the whole week",
 "Каждый день разное (1 вид)": "Different every day (1 type)",
 "Каждый день несколько видов": "Several types every day",
 "Количество видов:": "Number of types:",
 "Нет таймингов. Нажмите \"+\"": "No timings. Press \"+\"",
 ">Текст<": ">Text<",
 ">Категория<": ">Category<",
 "Введите название": "Enter a name",
 "<b>Тайминг ${idx+1}</b>": "<b>Timing ${idx+1}</b>",
 "Время:": "Time:",
 "(если \"до\" не указано = \"пока не съедят\")": "(if \"to\" is empty = \"until it runs out\")",
 "<strong>Позиции:</strong>": "<strong>Items:</strong>",
 "Нет позиций. Нажмите \"+\"": "No items. Press \"+\"",
 "Нет позиций": "No items",
 "<b>Позиция ${idx+1}</b>": "<b>Item ${idx+1}</b>",
 "С дня:": "From day:",
 "По день:": "To day:",
 "не ограничено": "not limited",
 "Категория ${i+1}: не выбрана": "Category ${i+1}: not selected",
 "Категория ${i+1}: нет дней": "Category ${i+1}: no days",
 "Ошибки:": "Errors:"
};
