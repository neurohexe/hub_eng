// BulkAssortmentImport.gs

const BULK_IMPORT_API_KEY_PROP  = 'GEMINI_API_KEY';
const BULK_IMPORT_MODEL_PROP    = 'GEMINI_MODEL';
const BULK_IMPORT_DEFAULT_MODEL = 'gemini-3.5-flash-lite'; // самая дешёвая/быстрая актуальная модель Gemini
const BULK_IMPORT_MAX_ITEMS     = 150;

function bulkImportAssortmentManual() {
  openBulkImportDialog();
}

function openBulkImportDialog() {
  const ui = SpreadsheetApp.getUi();

  if (!PropertiesService.getScriptProperties().getProperty(BULK_IMPORT_API_KEY_PROP)) {
    const res = ui.alert(
      L('AI-классификатор не настроен'),
      L('Чтобы автоматически распределять товары по складам и категориям,\n') +
      L('нужен ключ Google Gemini API.\n\nНастроить сейчас?'),
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
    connectAIClassifierKey();
    if (!PropertiesService.getScriptProperties().getProperty(BULK_IMPORT_API_KEY_PROP)) return;
  }

  const html = htmlTemplate_('BulkImportDialog').evaluate()
    .setWidth(560).setHeight(600);
  ui.showModalDialog(html, L('📋 Массовое добавление ассортимента'));
}

function connectAIClassifierKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    L('🤖 Подключение AI-классификатора'),
    L('Вставьте ключ Google Gemini API (начинается с "AIza...").\n\n') +
    L('Получить ключ: aistudio.google.com → Get API key\n') +
    L('(бесплатно, карта не нужна — есть бесплатный лимит запросов в день)\n\n') +
    L('Ключ используется только для распределения товаров\n') +
    L('по складам и категориям при массовой вставке ассортимента.'),
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!key) { ui.alert(L('❌ Ключ не может быть пустым')); return; }
  PropertiesService.getScriptProperties().setProperty(BULK_IMPORT_API_KEY_PROP, key);
  ui.alert(L('✅ Ключ сохранён.'));
}

// Вызывается из BulkImportDialog.html через google.script.run.
// Принимает текст (по строке на товар), возвращает готовый текстовый отчёт.
function processBulkAssortmentImport(rawText) {
  const parsed = _parseBulkAssortmentText(rawText);
  if (parsed.length === 0) {
    throw new Error(L('Список пуст — не удалось распознать ни одной строки.'));
  }
  if (parsed.length > BULK_IMPORT_MAX_ITEMS) {
    throw new Error(
      L('Слишком много позиций за один раз (') + parsed.length + ').\n' +
      L('Разбейте на части по ') + BULK_IMPORT_MAX_ITEMS + L(' штук и вставляйте по очереди.')
    );
  }

  const snapshot   = _getWarehouseCategoriesSnapshot();
  const classified = _classifyProductsWithAI(parsed, snapshot);
  const merged      = _mergeClassification(parsed, classified);
  const report       = _insertClassifiedProducts(merged);

  SpreadsheetApp.flush();
  return _formatBulkImportReport(report, parsed.length);
}

function _parseBulkAssortmentText(rawText) {
  const numberPattern = /^\d+(?:[.,]\d+)?$/;
  return (rawText || '')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean)
    .map(line => {
      const parts = line.split(/\s*\/\s*/).map(p => p.trim()).filter(Boolean);
      const name  = parts[0] || '';
      let unit = '', supplier = '', qty = '';

      for (let i = 1; i < parts.length; i++) {
        const part = parts[i];
        if (qty === '' && numberPattern.test(part)) {
          qty = part.replace(',', '.');
        } else if (unit === '') {
          unit = part;
        } else if (supplier === '') {
          supplier = part;
        }
      }

      return { name, unit, supplier, qty };
    })
    .filter(p => p.name);
}

function _getWarehouseCategoriesSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const snapshot = {};
  Object.keys(WAREHOUSE_CONFIG).forEach(key => {
    const cfg   = WAREHOUSE_CONFIG[key];
    const sheet = SH_(ss, cfg.warehouseSheet);
    snapshot[key] = {
      name:       L(cfg.name),
      exists:     !!sheet,
      categories: sheet ? AddProductButton.getCategories(sheet).map(c => c.name) : []
    };
  });
  return snapshot;
}

function _classifyProductsWithAI(products, snapshot) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(BULK_IMPORT_API_KEY_PROP);
  if (!apiKey) {
    throw new Error(L('Не настроен ключ AI (Gemini). ⚙️ Настройки → AI → Подключить AI-классификатор'));
  }
  const model = PropertiesService.getScriptProperties().getProperty(BULK_IMPORT_MODEL_PROP)
    || BULK_IMPORT_DEFAULT_MODEL;

  const warehouseTypes = Object.keys(snapshot);
  const warehouseDescriptions = warehouseTypes.map(key => {
    const w = snapshot[key];
    return key + ' ("' + w.name + '")' + (w.exists ? '' : L(' [склад ещё не создан]')) +
      L(' — существующие категории: ') + (w.categories.length ? w.categories.join('; ') : L('(пока ни одной)'));
  }).join('\n');

  const productLines = products.map((p, i) => (i + 1) + '. ' + p.name +
    (p.unit ? L(' (ед. изм.: ') + p.unit + ')' : '') +
    (p.supplier ? L(' (поставщик: ') + p.supplier + ')' : '')
  ).join('\n');

  const prompt =
    L('Ты помогаешь разложить список товаров склада по 4 складам и категориям внутри каждого склада.\n\n') +
    L('СКЛАДЫ И ИХ СУЩЕСТВУЮЩИЕ КАТЕГОРИИ:\n') + warehouseDescriptions + '\n\n' +
    L('ТОВАРЫ ДЛЯ РАСПРЕДЕЛЕНИЯ (') + products.length + L(' шт.):\n') + productLines + '\n\n' +
    L('Для каждого товара определи:\n') +
    L('- warehouseType: один из [') + warehouseTypes.join(', ') + ']\n' +
    L('- category: если товар подходит под уже существующую категорию этого склада — ') +
    L('используй её название ТОЧНО как в списке выше (посимвольно). Если ни одна не подходит — ') +
    L('придумай короткое точное название новой категории на русском (в том же стиле, что и существующие).\n') +
    L('- isNewCategory: true, если категории с таким названием ещё нет в списке существующих для этого склада.\n\n') +
    L('Верни РОВНО ') + products.length + L(' объектов, СТРОГО в том же порядке, что и список товаров выше, ') +
    L('даже если не до конца уверен — в этом случае выбери наиболее вероятный вариант.\n\n') +
    L('Ответь ТОЛЬКО JSON-массивом, без пояснений и без markdown-разметки (без \`\`\`):\n') +
    '[{"name":"...", "warehouseType":"...", "category":"...", "isNewCategory":true}, ...]';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model +
    ':generateContent?key=' + encodeURIComponent(apiKey);
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      maxOutputTokens:  8192,
      responseMimeType: 'application/json' // просим модель вернуть чистый JSON, без разметки вокруг
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  let result;
  try { result = JSON.parse(response.getContentText()); }
  catch(e) { throw new Error(L('AI вернул нераспознаваемый ответ (код ') + response.getResponseCode() + ')'); }

  if (result.error) throw new Error('AI: ' + (result.error.message || JSON.stringify(result.error)));

  const candidate = result.candidates && result.candidates[0];
  const parts     = candidate && candidate.content && candidate.content.parts;
  const text      = (parts && parts.map(p => p.text || '').join('')) || '';
  const jsonMatch = text.match(/\[[\s\S]*\]/);
  if (!jsonMatch) {
    const finishReason = (candidate && candidate.finishReason) || '?';
    throw new Error(
      L('AI вернул неожиданный ответ (finishReason=') + finishReason + '): ' + text.substring(0, 300)
    );
  }

  let parsedJson;
  try { parsedJson = JSON.parse(jsonMatch[0]); }
  catch(e) { throw new Error(L('Не удалось разобрать JSON от AI: ') + e.toString()); }

  if (!Array.isArray(parsedJson)) throw new Error(L('AI вернул не массив'));
  return parsedJson;
}

function _mergeClassification(products, classified) {
  const byName = {};
  classified.forEach(c => { if (c && c.name) byName[c.name.toString().toLowerCase().trim()] = c; });

  const sameLength = classified.length === products.length;

  return products.map((p, i) => {
    let c = sameLength ? classified[i] : null;
    if (!c || !c.warehouseType || !c.category) {
      c = byName[p.name.toLowerCase().trim()];
    }
    if (!c || !c.warehouseType || !c.category) {
      return Object.assign({}, p, { warehouseType: null, category: null, isNewCategory: false, matched: false });
    }
    return Object.assign({}, p, {
      warehouseType: c.warehouseType,
      category:      c.category.toString().trim(),
      isNewCategory: !!c.isNewCategory,
      matched:       true
    });
  });
}

function _insertClassifiedProducts(items) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const report = { added: [], newCategories: [], skipped: [], withStock: 0 };

  const byWarehouse = {};
  items.forEach(item => {
    if (!item.matched || !WAREHOUSE_CONFIG[item.warehouseType]) {
      report.skipped.push({ name: item.name, reason: L('AI не смог классифицировать') });
      return;
    }
    if (!byWarehouse[item.warehouseType]) byWarehouse[item.warehouseType] = [];
    byWarehouse[item.warehouseType].push(item);
  });

  Object.keys(byWarehouse).forEach(warehouseType => {
    const cfg   = WAREHOUSE_CONFIG[warehouseType];
    const sheet = SH_(ss, cfg.warehouseSheet);

    if (!sheet) {
      byWarehouse[warehouseType].forEach(item =>
        report.skipped.push({ name: item.name, reason: L('склад "') + L(cfg.name) + L('" ещё не создан') })
      );
      return;
    }

    const categories = AddProductButton.getCategories(sheet);

    byWarehouse[warehouseType].forEach(item => {
      let cat = categories.find(c => c.name.toLowerCase().trim() === item.category.toLowerCase().trim());

      if (!cat) {
        cat = _createCategoryAtEnd(sheet, item.category);
        categories.push(cat);
        report.newCategories.push(L(cfg.name) + ' → ' + item.category);
      }

      let row;
      if (cat._firstRowFree) {
        row = cat.categoryStart + 1;
        cat._firstRowFree = false;
      } else {
        row = AddProductButton.addProductToCategory(sheet, cat);
        categories.forEach(c => {
          if (c !== cat && c.categoryStart > cat.categoryStart) { c.categoryStart++; c.categoryEnd++; }
        });
      }

      sheet.getRange(row, 1).setValue(item.name);
      if (item.unit)     sheet.getRange(row, 3).setValue(item.unit);
      if (item.supplier) sheet.getRange(row, 6).setValue(item.supplier);

      if (item.qty !== '' && item.qty !== undefined && item.qty !== null) {
        const qtyNum = parseFloat(item.qty);
        if (!isNaN(qtyNum)) {
          sheet.getRange(row, 2).setValue(qtyNum);
          report.withStock++;
        }
      }

      report.added.push({ warehouse: L(cfg.name), category: item.category, name: item.name });
    });
  });

  try { updateSupplierDropdowns(); } catch(e) { Logger.log('⚠️ updateSupplierDropdowns: ' + e.toString()); }
  try { updateUnitsOnWarehouses();  } catch(e) { Logger.log('⚠️ updateUnitsOnWarehouses: '  + e.toString()); }

  return report;
}

// Создаёт новую категорию в конце листа склада (header + одна готовая пустая строка
// товара — как при ручном добавлении категории), возвращает объект, совместимый
// с тем, что использует AddProductButton.addProductToCategory.
function _createCategoryAtEnd(sheet, catName) {
  const dataCols  = AddProductButton.getDataCols(sheet);
  const extended  = AddProductButton.isExtendedSheet(sheet);
  const insertRow = sheet.getLastRow() + 1;
  sheet.insertRowsBefore(insertRow, 2);

  AddProductButton.drawCategoryRow(sheet, insertRow, catName, dataCols);
  AddProductButton.drawProductRow(sheet, insertRow + 1, dataCols, extended);

  return {
    name:          catName,
    categoryStart: insertRow,
    categoryEnd:   insertRow + 2,
    _firstRowFree: true   // ещё не занятая пустая строка товара, только что нарисованная выше
  };
}

function _formatBulkImportReport(report, totalRequested) {
  let msg = L('✅ Готово!\n\n');
  msg += L('Обработано товаров: ') + totalRequested + '\n';
  msg += L('Добавлено на склад: ') + report.added.length;
  if (report.withStock > 0) {
    msg += L('\nС указанным остатком: ') + report.withStock;
  }

  if (report.newCategories.length > 0) {
    const uniq = [...new Set(report.newCategories)];
    msg += L('\n\n🆕 Созданы новые категории (') + uniq.length + '):\n' + uniq.map(s => '• ' + s).join('\n');
  }

  if (report.skipped.length > 0) {
    msg += L('\n\n⚠️ Не удалось распределить (') + report.skipped.length + '):\n' +
      report.skipped.slice(0, 20).map(s => '• ' + s.name + ' — ' + s.reason).join('\n');
    if (report.skipped.length > 20) msg += L('\n… и ещё ') + (report.skipped.length - 20);
  }

  msg += L('\n\nПроверьте склады и при необходимости поправьте вручную ') +
         L('(➕ Добавить категорию / Добавление строк).');
  return msg;
}
