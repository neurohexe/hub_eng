// TelegramWebhook.gs

function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    Logger.log('📨 Webhook получил: ' + JSON.stringify(update, null, 2));
    handleUpdate(update);
  } catch (error) {
    Logger.log('❌ doPost error: ' + error.toString());
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleUpdate(update) {
  const BOT_TOKEN = TelegramManager.getBotToken();

  if (update.my_chat_member) {
    const member = update.my_chat_member;
    const chat = member.chat;
    const newStatus = member.new_chat_member && member.new_chat_member.status;
    if (newStatus === 'member' || newStatus === 'administrator') {
      Logger.log(`✓ Бот добавлен в чат: ${chat.title} (${chat.id})`);
      sendWelcomeMessage(BOT_TOKEN, chat);
    }
    if (newStatus === 'kicked' || newStatus === 'left') {
      Logger.log(`ℹ️ Бот удалён из чата: ${chat.title} (${chat.id})`);
    }
    return;
  }

  if (update.message) {
    const msg = update.message;
    const chat = msg.chat;
    const text = msg.text || '';

    if (text.startsWith('/start') || text.startsWith('/chatid')) {
      Logger.log(`✓ Команда ${text} в чате: ${chat.title || chat.id}`);
      sendWelcomeMessage(BOT_TOKEN, chat);
    }
  }
}

function sendWelcomeMessage(botToken, chat) {
  const chatId = chat.id;
  const chatTitle = chat.title || chat.first_name || L('этот чат');
  const text =
    L('✅ <b>Бот успешно добавлен!</b>') + '\n\n' +
    L('📱 Чат: ') + '<b>' + chatTitle + '</b>\n\n' +
    L('🆔 <b>Ваш Chat ID:</b>') + '\n' +
    '<code>' + chatId + '</code>\n\n' +
    L('📋 <b>Инструкция:</b>') + '\n' +
    L('1. Скопируйте ID выше (нажмите на него)') + '\n' +
    L('2. Откройте таблицу Google Sheets') + '\n' +
    L('3. Лист "Базовая информация" → раздел <b>ПОСТАВЩИКИ</b>') + '\n' +
    L('4. Найдите строку вашей компании-поставщика') + '\n' +
    L('5. Вставьте скопированный ID в колонку <b>"Telegram Chat ID"</b>') + '\n\n' +
    L('После этого заказы для вас будут приходить прямо в этот чат! 🛒');

  const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
  const payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'HTML'
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (result.ok) {
      Logger.log(`✅ Приветственное сообщение отправлено в ${chatId}`);
    } else {
      Logger.log(`❌ Ошибка отправки: ${JSON.stringify(result)}`);
    }
  } catch (error) {
    Logger.log('❌ sendWelcomeMessage error: ' + error.toString());
  }
}
