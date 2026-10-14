const TelegramManager = (function() {

  const PROP_BOT_TOKEN = 'TELEGRAM_BOT_TOKEN';
  const PROP_CHAT_ID   = 'TELEGRAM_CHAT_ID';
  const PROP_THREAD_ID = 'TELEGRAM_THREAD_ID';

  function getBotToken() {
    const token = PropertiesService.getScriptProperties().getProperty(PROP_BOT_TOKEN);
    if (!token) {
      throw new Error(
        L('Токен Telegram-бота не настроен!\n\n') +
        L('Перейдите: ⚙️ Настройки → 🔗 Настроить Telegram')
      );
    }
    return token;
  }

  function saveBotToken(token) {
    if (!token || token.trim() === '') throw new Error(L('Токен не может быть пустым'));
    PropertiesService.getScriptProperties().setProperty(PROP_BOT_TOKEN, token.trim());
    Logger.log('✅ Токен бота сохранён');
  }

  function getChatId() {
    const saved = PropertiesService.getScriptProperties().getProperty(PROP_CHAT_ID);
    if (saved && K(saved) !== 'Не подключён') return saved;

    try {
      const ss    = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = SH_(ss, L('Базовая информация'));
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = 0; i < data.length; i++) {
          if (data[i][0] && K(data[i][0]) === 'Chat ID логистики:') {
            const chatId = data[i][1] ? data[i][1].toString().trim() : '';
            if (chatId && K(chatId) !== 'Не подключён') {
              PropertiesService.getScriptProperties().setProperty(PROP_CHAT_ID, chatId);
              return chatId;
            }
          }
        }
      }
    } catch(e) {
      Logger.log('⚠️ getChatId (из листа): ' + e.toString());
    }

    throw new Error(
      L('Telegram чат не подключён!\n\n') +
      L('Перейдите: ⚙️ Настройки → 🔗 Настроить Telegram')
    );
  }

  function getDefaultThreadId() {
    const v = PropertiesService.getScriptProperties().getProperty(PROP_THREAD_ID);
    return v ? v : null;
  }

  function setDefaultThreadId() {
    const ui = SpreadsheetApp.getUi();

    const res = ui.alert(
      L('🧵 Общий чат — это тема внутри супергруппы?'),
      L('Если ваш общий чат логиста — это отдельная ТЕМА (топик)\n') +
      L('внутри большой супергруппы с несколькими темами\n') +
      L('(а НЕ отдельная группа/чат) — нажмите YES.\n\n') +
      L('Если это обычная отдельная группа — нажмите NO.'),
      ui.ButtonSet.YES_NO
    );

    if (res !== ui.Button.YES) {
      PropertiesService.getScriptProperties().deleteProperty(PROP_THREAD_ID);
      ui.alert(L('✅ Привязка к теме отключена.\n\nСообщения будут уходить в общий чат как обычно.'));
      return;
    }

    const idRes = ui.prompt(
      L('🧵 Thread ID темы'),
      L('Как получить:\n\n') +
      L('1. Напишите любое сообщение боту ВНУТРИ нужной темы\n') +
      L('   (не в "Общие", а именно в теме с логистом)\n') +
      L('2. Откройте в браузере:\n') +
      L('   api.telegram.org/bot<ВАШ_ТОКЕН>/getUpdates\n') +
      L('3. Найдите это сообщение, рядом будет поле\n') +
      L('   "message_thread_id": <число>\n') +
      L('4. Введите это число:'),
      ui.ButtonSet.OK_CANCEL
    );
    if (idRes.getSelectedButton() !== ui.Button.OK) return;

    const threadId = idRes.getResponseText().trim();
    if (!threadId || isNaN(Number(threadId))) {
      ui.alert(L('❌ Введите число (message_thread_id)'));
      return;
    }

    let chatId;
    try {
      chatId = getChatId();
    } catch(e) {
      ui.alert(L('❌ Сначала настройте общий чат:\n⚙️ Настройки → Настроить Telegram'));
      return;
    }

    const result = sendMessage(chatId, L('✅ Тема подключена!\n\nЗаказы без своего чата поставщика будут приходить сюда.'), threadId);
    if (!result.ok) {
      ui.alert(
        L('⚠️ Тестовое сообщение не отправилось:\n') + result.description +
        L('\n\nПроверьте Thread ID и что бот состоит в супергруппе.')
      );
      return;
    }

    PropertiesService.getScriptProperties().setProperty(PROP_THREAD_ID, threadId);
    ui.alert(L('✅ Тема подключена!\n\nThread ID: ') + threadId + L('\nТестовое сообщение отправлено в неё.'));
  }


  function sendMessage(chatId, text, threadId) {
    const token = getBotToken();
    const url   = 'https://api.telegram.org/bot' + token + '/sendMessage';
    try {
      const payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
      if (threadId) payload.message_thread_id = Number(threadId);

      const response = UrlFetchApp.fetch(url, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true
      });
      return JSON.parse(response.getContentText());
    } catch(e) {
      return { ok: false, description: e.toString() };
    }
  }

  // ── НОВОЕ: сообщение с inline-кнопками (например, подтверждение заказа хаба) ──
  function sendMessageWithButtons(chatId, text, inlineKeyboard, threadId) {
    const token = getBotToken();
    const url   = 'https://api.telegram.org/bot' + token + '/sendMessage';
    try {
      const payload = {
        chat_id: chatId, text: text, parse_mode: 'HTML',
        reply_markup: { inline_keyboard: inlineKeyboard }
      };
      if (threadId) payload.message_thread_id = Number(threadId);

      const response = UrlFetchApp.fetch(url, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true
      });
      return JSON.parse(response.getContentText());
    } catch(e) {
      return { ok: false, description: e.toString() };
    }
  }

  // ── НОВОЕ: правка уже отправленного сообщения (после нажатия кнопки) ──
  // inlineKeyboard необязателен — передайте его, чтобы заодно обновить и кнопки
  // под сообщением (например, актуальные количества после ➖/➕).
  function editMessage(chatId, messageId, text, inlineKeyboard) {
    const token = getBotToken();
    const url   = 'https://api.telegram.org/bot' + token + '/editMessageText';
    try {
      const payload = { chat_id: chatId, message_id: messageId, text: text, parse_mode: 'HTML' };
      if (inlineKeyboard) payload.reply_markup = { inline_keyboard: inlineKeyboard };
      const response = UrlFetchApp.fetch(url, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify(payload),
        muteHttpExceptions: true
      });
      return JSON.parse(response.getContentText());
    } catch(e) {
      return { ok: false, description: e.toString() };
    }
  }

  // ── НОВОЕ: гасит "часики" на нажатой кнопке в Telegram-клиенте ──
  function answerCallbackQuery(callbackQueryId, text) {
    const token = getBotToken();
    const url   = 'https://api.telegram.org/bot' + token + '/answerCallbackQuery';
    try {
      UrlFetchApp.fetch(url, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({ callback_query_id: callbackQueryId, text: text || '' }),
        muteHttpExceptions: true
      });
    } catch(e) {
      Logger.log('⚠️ answerCallbackQuery: ' + e.toString());
    }
  }

  function sendOrder(orderItems, weekPeriod, officeName) {
    const defaultChatId   = getChatId();
    const defaultThreadId = getDefaultThreadId();
    const supplierChats   = getSupplierChatMap();
    const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const office = officeName || L('Офис');

    const grouped = new Map();
    for (const item of orderItems) {
      const supplier = (item.supplier && item.supplier.toString().trim() !== '')
        ? item.supplier.toString().trim() : L('❓ Без поставщика');
      if (!grouped.has(supplier)) grouped.set(supplier, []);
      grouped.get(supplier).push(item);
    }

    const results = [];

    for (const [supplier, items] of grouped.entries()) {
      let text = L('🛒 <b>ЗАКАЗ ПРОДУКТОВ</b>\n');
      text += '🏢 ' + office + '\n📅 ' + today;
      if (weekPeriod) text += ' (' + weekPeriod + ')';
      text += '\n' + '─'.repeat(30) + '\n';
      text += '\n📦 <b>' + supplier + '</b>\n';

      for (let i = 0; i < items.length; i++) {
        const item        = items[i];
        const articul     = (item.articul || item.sku || '').toString().trim();
        const articulPart = articul !== '' ? ' [' + articul + ']' : '';
        const unit        = (item.unit || L('шт')).toString().trim();
        text += '\n' + (i + 1) + '. <b>' + item.name + '</b>' + articulPart;
        text += '\n   ' + (item.qty || '') + ' ' + unit;
      }
      text += L('\n\n📊 Позиций у поставщика: <b>') + items.length + '</b>';

      const hasOwnChat = !!supplierChats[supplier];
      const chatId      = hasOwnChat ? supplierChats[supplier] : defaultChatId;
      const threadId    = hasOwnChat ? null : defaultThreadId; // тема только для общего чата
      if (!hasOwnChat) text += L('\n\n📋 <i>Нет своего Telegram — скопируйте и перешлите вручную</i>');

      const result = sendMessage(chatId, text, threadId);
      if (!result.ok) Logger.log('⚠️ Telegram (' + supplier + '): ' + result.description);
      results.push({ supplier, chatId, ok: result.ok });
    }

    const allFailed = results.length > 0 && results.every(r => !r.ok);
    if (allFailed) throw new Error(L('Ошибка Telegram: сообщения не отправились ни одному поставщику'));

    return { ok: true, results };
  }


  function connectChat() {
    const ui = SpreadsheetApp.getUi();

    const tokenResult = ui.prompt(
      L('🔗 Настройка Telegram — шаг 1/2'),
      L('Введите токен вашего бота\n(получите у @BotFather в Telegram):'),
      ui.ButtonSet.OK_CANCEL
    );
    if (tokenResult.getSelectedButton() !== ui.Button.OK) return;

    const token = tokenResult.getResponseText().trim();
    if (!token) { ui.alert(L('❌ Токен не может быть пустым')); return; }

    try {
      const testUrl  = 'https://api.telegram.org/bot' + token + '/getMe';
      const response = UrlFetchApp.fetch(testUrl, { muteHttpExceptions: true });
      const result   = JSON.parse(response.getContentText());
      if (!result.ok) {
        ui.alert(L('❌ Неверный токен:\n') + (result.description || L('Проверьте токен')));
        return;
      }
      saveBotToken(token);
      Logger.log('✅ Токен проверен: @' + result.result.username);
    } catch(e) {
      ui.alert(L('❌ Ошибка проверки токена:\n') + e.toString());
      return;
    }

    const chatResult = ui.prompt(
      L('🔗 Настройка Telegram — шаг 2/2'),
      L('Введите Chat ID общего чата (или супергруппы, если у вас темы)\n\n') +
      L('Как получить Chat ID:\n') +
      L('1. Напишите /start боту в чате\n') +
      L('2. Откройте в браузере:\n') +
      L('api.telegram.org/bot<ТОКЕН>/getUpdates\n') +
      L('3. Найдите "chat":{"id": и скопируйте число:\n\n') +
      L('Если общий чат — это ТЕМА внутри супергруппы,\n') +
      L('после этого шага зайдите в:\n') +
      L('⚙️ Настройки → Telegram → Настроить тему (топик) чата'),
      ui.ButtonSet.OK_CANCEL
    );
    if (chatResult.getSelectedButton() !== ui.Button.OK) return;

    const chatId = chatResult.getResponseText().trim();
    if (!chatId) { ui.alert(L('❌ Chat ID не может быть пустым')); return; }

    PropertiesService.getScriptProperties().setProperty(PROP_CHAT_ID, chatId);
    Logger.log('✅ Chat ID сохранён в PropertiesService: ' + chatId);

    try {
      _saveChatIdToSheet(chatId);
    } catch(e) {
      Logger.log('⚠️ Не удалось записать Chat ID в лист (не критично): ' + e.toString());
    }

    try {
      const testResult = sendMessage(chatId,
        L('✅ <b>Telegram подключён!</b>\n\n') +
        L('Заказы поставщиков без своего Telegram будут приходить сюда 🛒')
      );
      if (!testResult.ok) throw new Error(testResult.description);
      Logger.log('✅ Тестовое сообщение отправлено');
    } catch(e) {
      ui.alert(
        L('⚠️ Chat ID сохранён, но тестовое сообщение не отправилось.\n\n') +
        L('Проверьте что бот добавлен в чат и Chat ID верный.\n\n') +
        'Chat ID: ' + chatId + L('\nОшибка: ') + e.toString()
      );
      return;
    }

    try { BaseInfo.updateTelegramStatus(L('✅ Подключён (Chat ID: ') + chatId + ')'); }
    catch(e) { Logger.log('⚠️ updateTelegramStatus: ' + e.toString()); }

    ui.alert(
      L('✅ Telegram успешно подключён!\n\n') +
      L('Бот: @') + _getBotUsername(token) + '\n' +
      'Chat ID: ' + chatId + '\n\n' +
      L('Если это тема внутри супергруппы — не забудьте:\n') +
      L('⚙️ Настройки → Telegram → Настроить тему (топик) чата')
    );
  }

  function showStatus() {
    const ui = SpreadsheetApp.getUi();
    try {
      const token    = getBotToken();
      const chatId   = getChatId();
      const threadId = getDefaultThreadId();

      const url      = 'https://api.telegram.org/bot' + token + '/getMe';
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const result   = JSON.parse(response.getContentText());

      const supplierChats = getSupplierChatMap();
      const supplierCount = Object.keys(supplierChats).length;

      if (result.ok) {
        ui.alert(
          L('📡 Статус Telegram\n\n') +
          L('✅ Бот активен: @') + result.result.username + '\n' +
          L('💬 Общий Chat ID: ') + chatId + '\n' +
          L('🧵 Тема (topic): ') + (threadId ? threadId : L('не используется')) + '\n' +
          L('📦 Поставщиков со своим чатом: ') + supplierCount
        );
      } else {
        ui.alert(L('❌ Бот недоступен:\n') + result.description);
      }
    } catch(e) {
      ui.alert(L('❌ Telegram не настроен:\n\n') + e.toString());
    }
  }


  function setupWebhook() {
    const ui = SpreadsheetApp.getUi();
    try {
      const token = getBotToken();

      const webhookUrl = ScriptApp.getService().getUrl();
      if (!webhookUrl) {
        ui.alert(
          L('❌ Веб-приложение не опубликовано!\n\n') +
          L('Сначала опубликуйте скрипт:\n') +
          L('1. Развернуть → Новое развёртывание\n') +
          L('2. Тип: Веб-приложение\n') +
          L('3. Доступ: Все\n') +
          L('4. Скопируйте URL и повторите эту команду')
        );
        return;
      }

      const url      = 'https://api.telegram.org/bot' + token + '/setWebhook';
      const response = UrlFetchApp.fetch(url, {
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({ url: webhookUrl }),
        muteHttpExceptions: true
      });
      const result = JSON.parse(response.getContentText());

      if (result.ok) {
        ui.alert(
          L('✅ Webhook настроен!\n\n') +
          L('Теперь бот будет автоматически присылать Chat ID\n') +
          L('когда его добавляют в чат.\n\n') +
          'URL: ' + webhookUrl
        );
        Logger.log('✅ Webhook зарегистрирован: ' + webhookUrl);
      } else {
        ui.alert(L('❌ Ошибка регистрации webhook:\n') + result.description);
      }
    } catch(e) {
      ui.alert(L('❌ Ошибка: ') + e.toString());
    }
  }


  function _saveChatIdToSheet(chatId) {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SH_(ss, L('Базовая информация'));
    if (!sheet) {
      Logger.log('⚠️ Лист "Базовая информация" не найден — Chat ID сохранён только в PropertiesService');
      return;
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] && K(data[i][0]) === 'Chat ID логистики:') {
        sheet.getRange(i + 1, 2).setValue(chatId);
        Logger.log('✅ Chat ID записан в Базовую информацию: ' + chatId);
        return;
      }
    }
    Logger.log('⚠️ Строка "Chat ID логистики:" не найдена в Базовой информации');
  }

  function _getBotUsername(token) {
    try {
      const response = UrlFetchApp.fetch(
        'https://api.telegram.org/bot' + token + '/getMe',
        { muteHttpExceptions: true }
      );
      const result = JSON.parse(response.getContentText());
      return result.ok ? result.result.username : '?';
    } catch(e) { return '?'; }
  }


  return {
    getBotToken,
    saveBotToken,
    getChatId,
    getDefaultThreadId,
    setDefaultThreadId,
    sendMessage,
    sendMessageWithButtons,
    editMessage,
    answerCallbackQuery,
    sendOrder,
    connectChat,
    showStatus,
    setupWebhook
  };

})();
