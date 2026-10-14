// AddProductButton.gs

const AddProductButton = (function() {

  function isExtendedSheet(sheet) {
    const name = K(sheet.getName());
    return name === 'Склад_Хозтовары' || name === 'Склад_Канцелярия';
  }

  function getDataCols(sheet) {
    return isExtendedSheet(sheet) ? 9 : 7;
  }


  function getCategories(sheet) {
    const data   = sheet.getDataRange().getValues();
    const btnCol = getDataCols(sheet);
    const cats   = [];

    for (let i = 2; i < data.length; i++) {
      const name    = data[i][0];
      const hasPlus = data[i][btnCol - 1] &&
                      data[i][btnCol - 1].toString().includes('➕');

      if (hasPlus && name) {
        let end = data.length + 1;
        for (let j = i + 1; j < data.length; j++) {
          const nextHasPlus = data[j][btnCol - 1] &&
                              data[j][btnCol - 1].toString().includes('➕');
          if (nextHasPlus && data[j][0]) { end = j + 1; break; }
        }
        cats.push({ name, categoryStart: i + 1, categoryEnd: end });
      }
    }
    return cats;
  }


  function findCategoryForRow(sheet, row) {
    const cats     = getCategories(sheet);
    const exactCat = cats.find(c => c.categoryStart === row);
    if (exactCat) return exactCat;
    return cats.find(c => row > c.categoryStart && row < c.categoryEnd) || null;
  }

 
  function getUnitsValidation() {
    return getUnitsValidationRule();
  }

  function getSuppliersValidation() {
    const s = getSuppliersFromBaseInfo();
    if (!s || s.length === 0) return null;
    return SpreadsheetApp.newDataValidation()
      .requireValueInList(s, true).setAllowInvalid(true).build();
  }


  function _drawCategoryRow(sheet, row, catName, dataCols) {
    sheet.getRange(row, 1, 1, dataCols).clearContent().clearFormat();

  
    const mergeRange = sheet.getRange(row, 1, 1, dataCols - 1);
    mergeRange.merge();

    sheet.getRange(row, 1)
      .setValue(catName)
      .setBackground('#d9d9d9')
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setFontSize(12)
      .setFontFamily('Arial')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');

  
    mergeRange.setBorder(
      true, true, true, true, null, null,
      '#000000', SpreadsheetApp.BorderStyle.SOLID
    );

    // Кнопка ➕
    sheet.getRange(row, dataCols)
      .clearFormat()
      .setValue('➕')
      .setFontSize(14)
      .setFontColor('#FFFFFF')
      .setBackground('#4CAF50')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontWeight('bold');

    sheet.setRowHeight(row, 30);
  }


  function _drawProductRow(sheet, row, dataCols, extended) {
    sheet.getRange(row, 1, 1, dataCols).clearContent().clearFormat();

    sheet.getRange(row, 2).setValue(0);
    sheet.getRange(row, 3).setValue(L('штука')).setDataValidation(getUnitsValidation());
    sheet.getRange(row, 4).insertCheckboxes().check();

    const suppRule = getSuppliersValidation();
    if (suppRule) sheet.getRange(row, 6).setDataValidation(suppRule);

    if (extended) {
      sheet.getRange(row, 7).insertCheckboxes().check();
      sheet.getRange(row, 8).insertCheckboxes().uncheck();
    }

    sheet.getRange(row, 1, 1, dataCols - 1)
      .setBackground('#ffffff')
      .setFontColor('#000000')
      .setFontWeight('normal')
      .setFontSize(10)
      .setFontFamily('Arial')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setBorder(
        true, true, true, true, true, true,
        '#000000', SpreadsheetApp.BorderStyle.SOLID
      );

    sheet.getRange(row, dataCols)
      .clearContent()
      .clearFormat()
      .setBackground(null);

    sheet.setRowHeight(row, 25);
  }

  function addProductToCategory(sheet, categoryInfo) {
    const dataCols = getDataCols(sheet);
    const extended = isExtendedSheet(sheet);

    const insertAfter = categoryInfo.categoryEnd - 2;
    sheet.insertRowAfter(insertAfter);
    const newRow = insertAfter + 1;

    _drawProductRow(sheet, newRow, dataCols, extended);
    SpreadsheetApp.flush();

    categoryInfo.categoryEnd++;
    return newRow;
  }

  function addNewCategory() {
    const ui    = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();

    if (!isWarehouseSheet_(sheet)) {
      ui.alert(L('❌ Только для листов склада'));
      return;
    }

    const extended = isExtendedSheet(sheet);
    const dataCols = getDataCols(sheet);


    const nameRes = ui.prompt(
      L('📁 Новая категория'),
      L('Название категории:'),
      ui.ButtonSet.OK_CANCEL
    );
    if (nameRes.getSelectedButton() !== ui.Button.OK) return;
    const catName = nameRes.getResponseText().trim();
    if (!catName) { ui.alert(L('❌ Название не может быть пустым')); return; }

    const categories = getCategories(sheet);
    if (categories.some(c => c.name.toLowerCase() === catName.toLowerCase())) {
      ui.alert(L('❌ Категория "') + catName + L('" уже существует'));
      return;
    }
   
    const posRes = ui.alert(
      L('📍 Позиция категории'),
      L('Добавить "') + catName + L('" в конец склада?'),
      ui.ButtonSet.YES_NO
    );

    let insertRow;

    if (posRes === ui.Button.YES) {

      insertRow = sheet.getLastRow() + 1;
      sheet.insertRowsBefore(insertRow, 2);

    } else {
  
      if (categories.length === 0) {

        insertRow = 3;
        sheet.insertRowsBefore(insertRow, 2);

      } else {
    
        const catList = categories.map((c, i) => `${i + 1}. ${c.name}`).join('\n');
        const choiceRes = ui.prompt(
          L('📍 Выберите позицию'),
          L('Вставить ПЕРЕД какой категорией?\n\n') +
          catList + '\n\n' +
          L('Введите номер\n(пусто = в самый конец):'),
          ui.ButtonSet.OK_CANCEL
        );
        if (choiceRes.getSelectedButton() !== ui.Button.OK) return;

        const choiceText = choiceRes.getResponseText().trim();

        if (choiceText === '') {
        
          insertRow = sheet.getLastRow() + 1;
          sheet.insertRowsBefore(insertRow, 2);
        } else {
          const choiceIdx = parseInt(choiceText) - 1;
          if (isNaN(choiceIdx) || choiceIdx < 0 || choiceIdx >= categories.length) {
            ui.alert(
              L('❌ Неверный номер.\n\n') +
              L('Введите число от 1 до ') + categories.length
            );
            return;
          }
        
          insertRow = categories[choiceIdx].categoryStart;
          sheet.insertRowsBefore(insertRow, 2);
        }
      }
    }

    _drawCategoryRow(sheet, insertRow,     catName,  dataCols);
    _drawProductRow (sheet, insertRow + 1, dataCols, extended);

    ui.alert(L('✅ Категория "') + catName + L('" добавлена!'));
    sheet.getRange(insertRow + 1, 1).activate();
  }

  function quickAddMultipleProducts() {
    const ui    = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();

    if (!isWarehouseSheet_(sheet)) {
      ui.alert(L('❌ Только для складов'));
      return;
    }

    const row = sheet.getActiveCell().getRow();
    if (row < 3) {
      ui.alert(L('❌ Выберите строку внутри категории (или на строке категории)'));
      return;
    }

    const cntRes = ui.prompt(
      L('📦 Добавление строк'),
      L('Сколько строк добавить? (1–50)'),
      ui.ButtonSet.OK_CANCEL
    );
    if (cntRes.getSelectedButton() !== ui.Button.OK) return;

    const count = parseInt(cntRes.getResponseText());
    if (isNaN(count) || count < 1 || count > 50) {
      ui.alert(L('❌ Введите число от 1 до 50'));
      return;
    }

    const category = findCategoryForRow(sheet, row);
    if (!category) {
      ui.alert(L('❌ Не удалось определить категорию.\n\nВыберите строку внутри нужной категории.'));
      return;
    }

    for (let i = 0; i < count; i++) {
      addProductToCategory(sheet, category);
    }

    ui.alert(L('✅ Добавлено ') + count + L(' строк в категорию "') + category.name + '"');
  }

  function repairWarehouseRows() {
    const ui    = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();

    if (!isWarehouseSheet_(sheet)) {
      ui.alert(L('❌ Только для листов склада'));
      return;
    }

    const bandings = sheet.getBandings();
    let bandingsRemoved = 0;
    bandings.forEach(b => { b.remove(); bandingsRemoved++; });

    const extended = isExtendedSheet(sheet);
    const dataCols = getDataCols(sheet);

    const data        = sheet.getDataRange().getValues();
    const unitRule     = getUnitsValidationRule();
    const suppliers     = getSuppliersFromBaseInfo();
    const supplierRule  = suppliers.length > 0
      ? SpreadsheetApp.newDataValidation().requireValueInList(suppliers, true).setAllowInvalid(true).build()
      : null;

    let fixedCount = 0, checkedRows = 0;

    for (let i = 2; i < data.length; i++) {
      const rowNum  = i + 1;
      const name    = data[i][0] ? data[i][0].toString().trim() : '';
      const hasPlus = data[i][dataCols - 1] && data[i][dataCols - 1].toString().includes('➕');

      if (hasPlus) continue;   
      if (!name) continue;     

      checkedRows++;
      let changed = false;

   
      if (data[i][1] === '' || data[i][1] === null) {
        sheet.getRange(rowNum, 2).setValue(0);
        changed = true;
      }


      const unitCell = sheet.getRange(rowNum, 3);
      unitCell.setDataValidation(unitRule);
      if (data[i][2] === '' || data[i][2] === null) {
        unitCell.setValue(L('штука'));
        changed = true;
      }


      const activeCell = sheet.getRange(rowNum, 4);
      const activeVal  = data[i][3];
      activeCell.insertCheckboxes();
      if (activeVal === false) {
        activeCell.uncheck();
      } else {
        activeCell.check();
        if (activeVal !== true) changed = true;
      }

    
      if (supplierRule) sheet.getRange(rowNum, 6).setDataValidation(supplierRule);

      
      if (extended) {
        const col7 = sheet.getRange(rowNum, 7);
        col7.insertCheckboxes();
        if (data[i][6] === false) col7.uncheck();
        else { col7.check(); if (data[i][6] !== true) changed = true; }

        const col8 = sheet.getRange(rowNum, 8);
        col8.insertCheckboxes();
        if (data[i][7] === true) col8.check();
        else col8.uncheck();
      }

    
      sheet.getRange(rowNum, 1, 1, dataCols - 1)
        .setBackground('#ffffff')
        .setFontColor('#000000')
        .setFontWeight('normal')
        .setFontSize(10)
        .setFontFamily('Arial')
        .setWrap(true)
        .setVerticalAlignment('middle')
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

      if (changed) fixedCount++;
    }

    SpreadsheetApp.flush();

    ui.alert(
      L('✅ Готово!\n\n') +
      L('Проверено строк-товаров: ') + checkedRows + '\n' +
      L('Дозаполнено значений по умолчанию: ') + fixedCount + '\n' +
      L('Снято чередующихся цветов (banding): ') + bandingsRemoved + '\n\n' +
      L('Чекбоксы, единицы измерения и список поставщиков\n') +
      L('пересозданы для всех строк на этом складе.')
    );
  }


  function createButtons(warehouseType) {
    const sheetName = WAREHOUSE_SHEET_NAMES[warehouseType];
    if (!sheetName) {
      SpreadsheetApp.getUi().alert(L('❌ Неизвестный тип склада: ') + warehouseType);
      return;
    }

    const sheet = SH_(SpreadsheetApp.getActiveSpreadsheet(), sheetName);
    if (!sheet) {
      SpreadsheetApp.getUi().alert(L('❌ Лист ') + L(sheetName) + L(' не найден'));
      return;
    }

    const cats     = getCategories(sheet);
    const dataCols = getDataCols(sheet);

    cats.forEach(cat => {
      sheet.getRange(cat.categoryStart, dataCols)
        .setValue('➕')
        .setBackground('#34A853')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold')
        .setFontSize(10)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    });

    SpreadsheetApp.getUi().alert(L('✅ Кнопки созданы для ') + WAREHOUSE_TYPES[warehouseType].name);
  }

  function handleButtonClick(row) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const cat   = findCategoryForRow(sheet, row);
    if (!cat) {
      SpreadsheetApp.getUi().alert(L('❌ Не удалось определить категорию'));
      return;
    }
    addProductToCategory(sheet, cat);
    SpreadsheetApp.getUi().alert(L('✅ Строка добавлена в "') + cat.name + '"');
  }

  return {
    quickAddMultipleProducts,
    createButtons,
    addNewCategory,
    handleButtonClick,
    repairWarehouseRows,
    // ── НОВОЕ: для массового импорта ассортимента (BulkAssortmentImport.js) ──
    getCategories,
    addProductToCategory,
    isExtendedSheet,
    getDataCols,
    drawCategoryRow: _drawCategoryRow,
    drawProductRow:  _drawProductRow
  };

})();

function addNewCategoryToWarehouse()              { AddProductButton.addNewCategory();           }
function quickAddMultipleProductsInCurrentSheet() { AddProductButton.quickAddMultipleProducts(); }
function repairWarehouseRowsInCurrentSheet()      { AddProductButton.repairWarehouseRows();      }
function createAddButtons_PRODUCTS()   { AddProductButton.createButtons('PRODUCTS');   }
function createAddButtons_HOUSEHOLD()  { AddProductButton.createButtons('HOUSEHOLD');  }
function createAddButtons_STATIONERY() { AddProductButton.createButtons('STATIONERY'); }
function createAddButtons_MEDICINE()   { AddProductButton.createButtons('MEDICINE');   }
function onButtonClick() {
  AddProductButton.handleButtonClick(SpreadsheetApp.getActiveSheet().getActiveCell().getRow());
}
