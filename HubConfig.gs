// HubConfig.gs

const HUB_PROP_MODE                = 'OFFICE_MODE';           // 'FULL' | 'HUB'
const HUB_TARGETS_SHEET            = 'Хаб_Целевые_остатки';
const HUB_HISTORY_SHEET            = 'Хаб_История_движения_остатков';
const HUB_STOCK_FORM_ID_KEY        = 'HUB_STOCK_FORM_ID';
const HUB_PENDING_ORDERS_KEY       = 'HUB_PENDING_ORDER_';    // + draftId, в Script Properties
const HUB_RESPONSIBLE_CHATID_LABEL = 'Chat ID ответственного (хаб):'; // менеджер: подтверждает/правит/отменяет заказ хаба

function isHubMode() {
  return PropertiesService.getScriptProperties().getProperty(HUB_PROP_MODE) === 'HUB';
}

function setOfficeMode(mode) {
  if (mode !== 'FULL' && mode !== 'HUB') throw new Error(L('Неизвестный режим офиса: ') + mode);
  PropertiesService.getScriptProperties().setProperty(HUB_PROP_MODE, mode);
}

function chooseOfficeModeDialog() {
  const ui = SpreadsheetApp.getUi();
  const current = isHubMode() ? L('ХАБ (упрощённый)') : L('ПОЛНЫЙ (менеджер + логист)');
  const res = ui.alert(
    L('Тип объекта'),
    L('Текущий режим: ') + current + '\n\n' +
    L('YES — Полный склад (менеджер/логист, ежедневный цикл)\n') +
    L('NO  — Хаб (только хозяюшка, еженедельный пересчёт остатков)'),
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (res === ui.Button.YES) { setOfficeMode('FULL'); ui.alert(L('✅ Режим: Полный склад')); onOpen(); }
  else if (res === ui.Button.NO) {
    setOfficeMode('HUB');
    let hidden = 0;
    try { hidden = hideHubUnusedSheets(); } catch(e) { Logger.log('⚠️ hideHubUnusedSheets: ' + e.toString()); }
    ui.alert(L('✅ Режим: Хаб') + (hidden > 0 ? L('\n\nСкрыто неиспользуемых листов (дашборды/БД): ') + hidden : ''));
    onOpen();
  }
}

// У хаба нет ежедневного цикла выдач/поставок, поэтому дневные дашборды
// (Продукты/Хозтовары/Канцелярия/Медикаменты) и листы-архивы (*_DB) ему не нужны —
// они всё равно создаются при первоначальной настройке структуры склада (общий код
// с полным офисом), поэтому просто скрываем их, чтобы не мешались в списке листов.
// Движения по хабу вместо этого ведутся в HUB_HISTORY_SHEET (см. HubStockForm.js).
function hideHubUnusedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hidden = 0;
  Object.keys(WAREHOUSE_CONFIG).forEach(key => {
    const cfg = WAREHOUSE_CONFIG[key];
    [cfg.dashSheet, cfg.dbSheet].forEach(name => {
      if (!name) return;
      const sheet = SH_(ss, name);
      if (sheet && !sheet.isSheetHidden()) { sheet.hideSheet(); hidden++; }
    });
  });
  return hidden;
}

function hideHubUnusedSheetsManual() {
  const ui = SpreadsheetApp.getUi();
  try {
    const hidden = hideHubUnusedSheets();
    ui.alert(hidden > 0
      ? L('✅ Скрыто листов: ') + hidden + L('\n\n(дневные дашборды и архивы *_DB — хабу они не нужны)')
      : L('ℹ️ Скрывать нечего — такие листы уже скрыты или их нет.'));
  } catch(e) {
    ui.alert(L('❌ Ошибка: ') + e.toString());
  }
}

// Ищет товар по всем 4 складам (Продукты/Хозтовары/Канцелярия/Медикаменты).
// Возвращает { sheet, row (1-based), unit, category, warehouseType } или null.
function findProductAcrossWarehouses(productName) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = (productName || '').toString().trim();
  if (!name) return null;

  for (const warehouseType of Object.keys(WAREHOUSE_CONFIG)) {
    const cfg   = WAREHOUSE_CONFIG[warehouseType];
    const sheet = SH_(ss, cfg.warehouseSheet);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    let currentCategory = '';
    for (let i = 2; i < data.length; i++) {
      const rowName = data[i][0] ? data[i][0].toString().trim() : '';
      const qty     = data[i][1], unit = data[i][2], active = data[i][3];
      const isCat   = (qty === '' || qty === null) && (unit === '' || unit === null) &&
                      (active === '' || active === false || active === null);
      if (isCat && rowName) { currentCategory = rowName; continue; }
      if (rowName === name) {
        return {
          sheet, row: i + 1,
          unit: unit ? unit.toString().trim() : '',
          category: currentCategory,
          warehouseType
        };
      }
    }
  }
  return null;
}

// Все активные товары по всем 4 складам: остаток, категория, тип склада, поставщик, артикул.
function getAllHubProducts() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  Object.keys(WAREHOUSE_CONFIG).forEach(warehouseType => {
    const cfg   = WAREHOUSE_CONFIG[warehouseType];
    const sheet = SH_(ss, cfg.warehouseSheet);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    let currentCategory = '';
    for (let i = 2; i < data.length; i++) {
      const name  = data[i][0] ? data[i][0].toString().trim() : '';
      const qty   = data[i][1], unit = data[i][2], active = data[i][3];
      const isCat = (qty === '' || qty === null) && (unit === '' || unit === null) &&
                    (active === '' || active === false || active === null);
      if (isCat && name) { currentCategory = name; continue; }
      if (!name || active !== true) continue;

      const extra = getArticulAndSupplier(sheet, name);
      result.push({
        name, warehouseType, category: currentCategory,
        stock: parseFloat(qty) || 0,
        unit: unit ? unit.toString().trim() : '',
        articul: extra.articul, supplier: extra.supplier
      });
    }
  });

  return result;
}

function getResponsibleChatId() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SH_(ss, L('Базовая информация'));
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && K(data[i][0]) === HUB_RESPONSIBLE_CHATID_LABEL) {
      const v = data[i][1] ? data[i][1].toString().trim() : '';
      return v && K(v) !== 'Не подключён' ? v : null;
    }
  }
  return null;
}

function setResponsibleChatIdRow() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SH_(ss, L('Базовая информация'));
  if (!sheet) { ui.alert(L('❌ Сначала создайте "Базовую информацию"')); return; }

  const res = ui.prompt(
    L('👤 Менеджер хаба (подтверждение заказов)'),
    L('Каждую неделю, когда хозяюшка сдаёт остатки, черновик заказа\n') +
    L('сначала уходит СЮДА в Telegram — с кнопками:\n') +
    L('✅ Отправить · ✏️ Скорректировать · ❌ Отменить.\n') +
    L('Поставщикам заказ уходит только после этого.\n\n') +
    L('⚠️ Если не настроить — заказ будет уходить поставщикам\n') +
    L('сразу и автоматически, без подтверждения.\n\n') +
    L('Как узнать chat id: напишите /start вашему боту в личку,\n') +
    L('затем откройте api.telegram.org/bot<ТОКЕН>/getUpdates'),
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const chatId = res.getResponseText().trim();
  if (!chatId) { ui.alert(L('❌ Chat ID не может быть пустым')); return; }

  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && K(data[i][0]) === HUB_RESPONSIBLE_CHATID_LABEL) {
      sheet.getRange(i + 1, 2).setValue(chatId);
      ui.alert(L('✅ Сохранено: ') + chatId);
      return;
    }
  }
  // строки ещё нет — добавляем сразу под блоком TELEGRAM, третьей строкой после "Статус:"
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && K(data[i][0]) === 'Статус:') {
      sheet.insertRowAfter(i + 1);
      sheet.getRange(i + 2, 1).setValue(L(HUB_RESPONSIBLE_CHATID_LABEL))
        .setFontWeight('bold').setBackground('#f3f3f3');
      sheet.getRange(i + 2, 2).setValue(chatId).setBackground('#ffffff');
      ui.alert(L('✅ Сохранено: ') + chatId);
      return;
    }
  }
  ui.alert(L('⚠️ Не найден блок TELEGRAM в "Базовой информации" — добавьте строку вручную.'));
}
