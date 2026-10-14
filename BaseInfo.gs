// BaseInfo.gs

const BaseInfo = (function() {

  const SHEET_NAME = 'Базовая информация';

  const CP_HEADER_ROW  = 9;
  const CP_DATA_START  = 10;
  const CP_DATA_COUNT  = 10;
  const CP_COUNTER_ROW = CP_DATA_START + CP_DATA_COUNT; // = 20

  function _getSS() {
    try {
      return SpreadsheetApp.getActiveSpreadsheet();
    } catch(e) {
      const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      return SpreadsheetApp.openById(ssId);
    }
  }

  function createSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = SH_(ss, SHEET_NAME);

    if (sheet) {
      const response = SpreadsheetApp.getUi().alert(
        L('Лист "Базовая информация" уже существует'),
        L('Пересоздать? Все данные будут удалены.'),
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      if (response === SpreadsheetApp.getUi().Button.YES) {
        ss.deleteSheet(sheet);
      } else {
        return;
      }
    }

    sheet = ss.insertSheet(L(SHEET_NAME));
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 140); 

    buildStructure(sheet);

    SpreadsheetApp.getUi().alert(
      L('✅ Лист "Базовая информация" создан!\n\n') +
      L('Заполните данные офиса.\n\n') +
      L('Для подключения Telegram:\n') +
      L('⚙️ Настройки → 🔗 Настроить Telegram')
    );
  }

  function buildStructure(sheet) {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');

    sheet.getRange('A1:D1').merge();
    sheet.getRange('A1')
      .setValue(L('БАЗОВАЯ ИНФОРМАЦИЯ ОФИСА — обновлено ') + today)
      .setBackground('#1c4587').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(14)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 40);

    addSectionTitle(sheet, 3, L('ИДЕНТИФИКАЦИЯ ОФИСА'));
    addDataBlock(sheet, 4, [
      [L('Название офиса:'), ''],
      [L('ID офиса:'), ''],
      [L('Город:'), ''],
      [L('Дата запуска:'), today]
    ]);

    addSectionTitle(sheet, 8, L('КОФЕ-ПОИНТЫ'));

    sheet.getRange(CP_HEADER_ROW, 1, 1, 4)
      .setValues([['ID', L('Название кофе-поинта'), L('Этаж/Зона'), L('Активен')]]);
    sheet.getRange(CP_HEADER_ROW, 1, 1, 4)
      .setFontWeight('bold').setBackground('#6fa8dc').setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#4a86e8', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(CP_HEADER_ROW, 26);

    for (let i = 0; i < CP_DATA_COUNT; i++) {
      const r = CP_DATA_START + i;
      sheet.getRange(r, 1).setValue(i + 1).setHorizontalAlignment('right').setBackground('#ffffff');
      sheet.getRange(r, 2).setValue('').setBackground('#ffffff');
      sheet.getRange(r, 3).setValue('').setBackground('#ffffff');
      sheet.getRange(r, 4).setBackground('#ffffff').insertCheckboxes();
      sheet.getRange(r, 1, 1, 4)
        .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(r, 24);
    }

    sheet.getRange(CP_DATA_START + CP_DATA_COUNT - 1, 1, 1, 4)
      .setBorder(null, null, true, null, null, null, '#4a86e8', SpreadsheetApp.BorderStyle.MEDIUM);

    sheet.getRange(CP_COUNTER_ROW, 1)
      .setValue(L('Активных кофе-поинтов:')).setFontWeight('bold').setBackground('#f3f3f3');

    sheet.getRange(CP_COUNTER_ROW, 2)
      .setFormula('=COUNTIF(D10:D19;TRUE)').setFontWeight('bold')
      .setFontColor('#0000FF').setBackground('#f3f3f3');
    sheet.getRange(CP_COUNTER_ROW, 1, 1, 2)
      .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(CP_COUNTER_ROW, 24);

    const TG_START = CP_COUNTER_ROW + 2;
    sheet.getRange(TG_START, 1, 1, 4).merge();
    sheet.getRange(TG_START, 1)
      .setValue(L('TELEGRAM — ЧАТ ЛОГИСТИКИ'))
      .setFontWeight('bold').setFontSize(12)
      .setBackground('#2196F3').setFontColor('#ffffff').setVerticalAlignment('middle');
    sheet.setRowHeight(TG_START, 30);

    const tgData = [
      [L('Chat ID логистики:'), L('Не подключён')],
      [L('Название чата:'),      L('Не подключён')],
      [L('Статус:'),             L('Не подключён')]
    ];
    for (let i = 0; i < tgData.length; i++) {
      const r = TG_START + 1 + i;
      sheet.getRange(r, 1).setValue(tgData[i][0]).setFontWeight('bold').setBackground('#f3f3f3').setWrap(true);
      sheet.getRange(r, 2).setValue(tgData[i][1]).setBackground('#ffffff').setWrap(true);
      sheet.getRange(r, 1, 1, 2)
        .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment('middle');
      sheet.setRowHeight(r, 24);
    }
    sheet.autoResizeRows(TG_START + 1, tgData.length);

    const TG_INSTR = TG_START + 4;
    sheet.getRange(TG_INSTR, 1, 1, 4).merge();
    sheet.getRange(TG_INSTR, 1)
      .setValue(
        L('Как подключить: добавьте бота в чат поставщика → ') +
        L('бот напишет Chat ID → скопируйте ID → ') +
        L('Настройки → Настроить Telegram → вставьте ID')
      )
      .setFontColor('#666666').setFontStyle('italic').setFontSize(10).setWrap(true);
    sheet.setRowHeight(TG_INSTR, 40);

    const SUPPLIERS_START = TG_INSTR + 2;
    _buildSuppliersSection(sheet, SUPPLIERS_START);
  }

  function _buildSuppliersSection(sheet, startRow) {
    const ROWS = 10, COLS = 5; 

    sheet.getRange(startRow, 1, 1, COLS).merge();
    sheet.getRange(startRow, 1)
      .setValue(L('ПОСТАВЩИКИ'))
      .setBackground('#1c4587').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(13)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(startRow, 1, 1, COLS)
      .setBorder(true, true, true, true, null, null, '#1c4587', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(startRow, 32);

    const hdrRow = startRow + 1;
    sheet.getRange(hdrRow, 1, 1, COLS)
      .setValues([[L('Название поставщика'), L('Контакт'), L('Примечание'), 'Telegram Chat ID', L('Активен')]])
      .setBackground('#4a86e8').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#1c4587', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(hdrRow, 26);

    const dataStart = hdrRow + 1;
    for (let i = 0; i < ROWS; i++) {
      const r  = dataStart + i;
      const bg = i % 2 === 0 ? '#ffffff' : '#f0f4ff';
      sheet.getRange(r, 1).setValue('').setBackground(bg).setFontSize(10).setWrap(true);
      sheet.getRange(r, 2).setValue('').setBackground(bg).setFontSize(10).setWrap(true);
      sheet.getRange(r, 3).setValue('').setBackground(bg).setFontSize(10).setWrap(true);
      sheet.getRange(r, 4).setValue('').setBackground(bg).setFontSize(10).setWrap(true);
      sheet.getRange(r, 5).setBackground(bg).insertCheckboxes().check();
      sheet.getRange(r, 1, 1, COLS)
        .setBorder(true, true, true, true, true, true, '#d0d8e4', SpreadsheetApp.BorderStyle.SOLID);
      sheet.setRowHeight(r, 24);
    }

    sheet.getRange(dataStart + ROWS - 1, 1, 1, COLS)
      .setBorder(null, null, true, null, null, null, '#1c4587', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  function getActiveCoffeePoints() {
    const ss    = _getSS();
    const sheet = SH_(ss, SHEET_NAME);

    if (!sheet) throw new Error(
      L('Лист "Базовая информация" не найден!\n') +
      L('Создайте его: Настройки → Создать "Базовую информацию"')
    );

    const data = sheet.getDataRange().getValues();
    let cpHeaderRow = -1, cpCounterRow = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === 'ID' &&
          data[i][1] && K(data[i][1]).includes('Название')) {
        cpHeaderRow = i + 1;
      }
      if (data[i][0] && K(data[i][0]) === 'Активных кофе-поинтов:') {
        cpCounterRow = i + 1; break;
      }
    }

    if (cpHeaderRow === -1 || cpCounterRow === -1) throw new Error(
      L('Не найдена таблица кофе-поинтов.\n') +
      L('Пересоздайте лист: Настройки → Создать "Базовую информацию"')
    );

    const coffeePoints = [];
    for (let i = cpHeaderRow; i < cpCounterRow - 1; i++) {
      const id = data[i][0], name = data[i][1], zone = data[i][2], active = data[i][3];
      if (!id || id === '') continue;
      if (active === true && name && name.toString().trim() !== '') {
        coffeePoints.push({ id, name: name.toString().trim(), zone: zone ? zone.toString().trim() : '' });
      }
    }
    return coffeePoints;
  }

  function getOfficeConfig() {
    const ss    = _getSS();
    const sheet = SH_(ss, SHEET_NAME);
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const config = {};
    const fieldMap = {
      'Название офиса:':     'officeName',
      'ID офиса:':           'officeId',
      'Город:':              'city',
      'Дата запуска:':       'launchDate',
      'Chat ID логистики:': 'telegramChatId',
      'Название чата:':      'telegramChatTitle'
    };
    for (let i = 0; i < data.length; i++) {
      const label = data[i][0] ? K(data[i][0]) : '';
      if (fieldMap[label] !== undefined) config[fieldMap[label]] = data[i][1];
    }
    return config;
  }

  function updateTelegramStatus(status) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SH_(ss, SHEET_NAME);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && K(data[i][0]) === 'Статус:') {
        sheet.getRange(i + 1, 2)
          .setValue(status)
          .setFontColor(status.includes('✅') ? '#0a7a0a' : '#cc0000')
          .setFontWeight('bold');
        break;
      }
    }
  }


  function addSupplierRow() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SH_(ss, SHEET_NAME);
    if (!sheet) { SpreadsheetApp.getUi().alert(L('❌ Лист "Базовая информация" не найден')); return; }

    const data = sheet.getDataRange().getValues();
    let suppColRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && K(data[i][0]).toUpperCase() === 'ПОСТАВЩИКИ') {
        suppColRow = i + 1; break;
      }
    }
    if (suppColRow === -1) { SpreadsheetApp.getUi().alert(L('❌ Раздел ПОСТАВЩИКИ не найден')); return; }

    const STOP_WORDS = ['ИДЕНТИФИКАЦИЯ','КОФЕ-ПОИНТЫ','TELEGRAM',
                        'ПАРАМЕТРЫ','ПРАВИЛА','AI ПРОФИЛЬ','LIVE'];
    let rowCount = 0;
    for (let i = suppColRow + 1; i < data.length; i++) {
      const v = data[i][0] ? K(data[i][0]).toUpperCase() : '';
      if (STOP_WORDS.some(h => v.startsWith(h))) break;
      rowCount++;
    }

    const lastDataNum = suppColRow + 1 + rowCount;
    const newRowNum   = lastDataNum + 1;
    sheet.insertRowAfter(lastDataNum);

    const bg = (rowCount + 1) % 2 === 1 ? '#ffffff' : '#f0f4ff';
    sheet.getRange(newRowNum, 1).setValue('').setBackground(bg).setFontSize(10);
    sheet.getRange(newRowNum, 2).setValue('').setBackground(bg).setFontSize(10);
    sheet.getRange(newRowNum, 3).setValue('').setBackground(bg).setFontSize(10);
    sheet.getRange(newRowNum, 4).setValue('').setBackground(bg).setFontSize(10);
    sheet.getRange(newRowNum, 5).setBackground(bg).insertCheckboxes().check();
    sheet.getRange(newRowNum, 1, 1, 5)
      .setBorder(true, true, true, true, true, true, '#d0d8e4', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(newRowNum, 24);
    sheet.getRange(newRowNum, 1).activate();
  }

  function fixSuppliersColors() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SH_(ss, SHEET_NAME);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let suppColRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && K(data[i][0]).toUpperCase() === 'ПОСТАВЩИКИ') {
        suppColRow = i + 1; break;
      }
    }
    if (suppColRow === -1) return;
    // 'ПОЛЬЗОВАТЕЛИ' убрано из списка стоп-заголовков.
    const STOP = ['ИДЕНТИФИКАЦИЯ','КОФЕ-ПОИНТЫ','TELEGRAM',
                  'ПАРАМЕТРЫ','ПРАВИЛА','AI ПРОФИЛЬ','LIVE'];
    let pos = 0;
    for (let i = suppColRow + 1; i < data.length; i++) {
      const v = data[i][0] ? K(data[i][0]).toUpperCase() : '';
      if (STOP.some(h => v.startsWith(h))) break;
      pos++;
      sheet.getRange(i + 1, 1, 1, 5).setBackground(pos % 2 === 1 ? '#ffffff' : '#f0f4ff');
    }
  }


  function addSectionTitle(sheet, row, title) {
    sheet.getRange(row, 1, 1, 4).merge();
    sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(12)
      .setBackground('#FFD966').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 28);
  }

  function addDataBlock(sheet, startRow, data, isReadOnly = false) {
    for (let i = 0; i < data.length; i++) {
      const r = startRow + i;
      sheet.getRange(r, 1).setValue(data[i][0]).setFontWeight('bold').setBackground('#f3f3f3').setWrap(true);
      sheet.getRange(r, 2).setValue(data[i][1]).setBackground(isReadOnly ? '#f8f8f8' : '#ffffff').setWrap(true);
      if (isReadOnly) sheet.getRange(r, 2).setFontColor('#0000FF');
      sheet.getRange(r, 1, 1, 2)
        .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment('middle');
      sheet.setRowHeight(r, 24);
    }
  }

  return {
    createSheet,
    getActiveCoffeePoints,
    getOfficeConfig,
    updateTelegramStatus,
    addSupplierRow,
    fixSuppliersColors,
    CP_DATA_START,
    CP_COUNTER_ROW
  };

})();

function addCoffeePointRow() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = SH_(ss, L('Базовая информация'));
  if (!sheet) { SpreadsheetApp.getUi().alert(L('❌ Лист "Базовая информация" не найден.')); return; }

  const data = sheet.getDataRange().getValues();
  let cpHeaderRow = -1, cpCounterRow = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === 'ID' &&
        data[i][1] && K(data[i][1]).includes('Название')) {
      cpHeaderRow = i + 1;
    }
    if (data[i][0] && K(data[i][0]) === 'Активных кофе-поинтов:') {
      cpCounterRow = i + 1; break;
    }
  }

  if (cpHeaderRow === -1 || cpCounterRow === -1) {
    SpreadsheetApp.getUi().alert(L('❌ Не найдена таблица кофе-поинтов.'));
    return;
  }

  const cpDataStart = cpHeaderRow + 1;

  const nextId = cpCounterRow - cpDataStart + 1;

  sheet.insertRowBefore(cpCounterRow);
  const newRow = cpCounterRow;

  sheet.getRange(newRow, 1).setValue(nextId).setHorizontalAlignment('right').setBackground('#ffffff');
  sheet.getRange(newRow, 2).setValue('').setBackground('#ffffff');
  sheet.getRange(newRow, 3).setValue('').setBackground('#ffffff');
  sheet.getRange(newRow, 4).setBackground('#ffffff').insertCheckboxes();
  sheet.getRange(newRow, 1, 1, 4)
    .setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(newRow, 24);

  const newCounterRow = cpCounterRow + 1;

  sheet.getRange(newCounterRow, 2)
    .setFormula('=COUNTIF(D' + cpDataStart + ':D' + (newCounterRow - 1) + ';TRUE)')
    .setFontWeight('bold').setFontColor('#0000FF');

  sheet.getRange(newRow, 2).activate();
  SpreadsheetApp.getUi().alert(L('✅ Кофе-поинт #') + nextId + L(' добавлен!\n\nВведите название в выделенную ячейку.'));
}

function createBaseInfoSheet()       { BaseInfo.createSheet();        }
function addSupplierRowWrapper()     { BaseInfo.addSupplierRow();     }
function fixSuppliersColorsWrapper() { BaseInfo.fixSuppliersColors(); }
function addCoffeePointRowWrapper()  { addCoffeePointRow();           }

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (K(sheet.getName()) !== 'Базовая информация') return;
    sheet.autoResizeRows(e.range.getRow(), e.range.getNumRows());
  } catch(err) {
    Logger.log('⚠️ onEdit autoResize: ' + err.toString());
  }
}
