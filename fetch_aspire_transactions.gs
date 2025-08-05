function fetchAspireTransactionsToSheet() {
  const accountId = '9d6d2e6b-4833-4b73-bdeb-a80718916cb3'; // Правильный Account ID
  const startDate = '2024-04-01T00:00:00Z';

  // URL прокси на Render
  const proxyUrl = `https://aspire-proxy-render.onrender.com/aspire?account_id=${encodeURIComponent(accountId)}&start_date=${encodeURIComponent(startDate)}`;

  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      'User-Agent': 'GoogleAppsScript/1.0'
    }
  };

  try {
    Logger.log("📡 Отправляем запрос на прокси...");
    Logger.log("🔗 URL: " + proxyUrl);
    
    const response = UrlFetchApp.fetch(proxyUrl, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    Logger.log("📥 Код ответа: " + code);
    Logger.log("📥 Размер ответа: " + body.length + " символов");

    if (code !== 200) {
      throw new Error(`Не удалось получить транзакции. Код ответа: ${code}, Ответ: ${body}`);
    }

    const data = JSON.parse(body);
    const transactions = data.data || []; // Правильная структура ответа

    Logger.log(`✅ Получено транзакций: ${transactions.length}`);

    // Записываем в Google Sheets
    const sheetName = 'Aspire_USD';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log("📋 Создан новый лист: " + sheetName);
    } else {
      sheet.clear();
      Logger.log("📋 Очищен существующий лист: " + sheetName);
    }

    // Улучшенные заголовки
    const headers = [
      'Date', 'Amount', 'Currency', 'Type', 'Status', 'Reference', 
      'Description', 'Counterparty', 'Balance', 'Category', 'Card Number'
    ];
    sheet.appendRow(headers);

    // Записываем транзакции с улучшенной обработкой
    let processedCount = 0;
    transactions.forEach((tx, index) => {
      try {
        const row = [
          tx.datetime || '',
          tx.amount || 0,
          tx.currency_code || '',
          tx.type || '',
          tx.status || '',
          tx.reference || '',
          tx.counterparty_name || '',
          tx.counterparty_name || '',
          tx.balance || 0,
          tx.additional_info?.spend_category || '',
          tx.additional_info?.card_number || ''
        ];
        
        sheet.appendRow(row);
        processedCount++;
        
        if (index % 10 === 0) {
          Logger.log(`📊 Обработано транзакций: ${index + 1}/${transactions.length}`);
        }
      } catch (rowError) {
        Logger.log(`❌ Ошибка при обработке транзакции ${index}: ${rowError.message}`);
      }
    });

    // Добавляем информацию о метаданных
    if (data.metadata) {
      const metaSheet = ss.getSheetByName('Metadata') || ss.insertSheet('Metadata');
      metaSheet.clear();
      metaSheet.appendRow(['Параметр', 'Значение']);
      metaSheet.appendRow(['Всего транзакций', data.metadata.total]);
      metaSheet.appendRow(['Текущая страница', data.metadata.current_page]);
      metaSheet.appendRow(['Транзакций на странице', data.metadata.per_page]);
      metaSheet.appendRow(['Запрос ID', data.metadata['aspire-request-id']]);
    }

    Logger.log(`✅ Успешно обработано транзакций: ${processedCount}/${transactions.length}`);
    Logger.log(`📊 Всего транзакций в системе: ${data.metadata?.total || 'неизвестно'}`);
    
    // Добавляем информацию о запросе
    sheet.appendRow([]);
    sheet.appendRow(['Запрос выполнен:', new Date().toISOString()]);
    sheet.appendRow(['Account ID:', accountId]);
    sheet.appendRow(['Start Date:', startDate]);
    sheet.appendRow(['Транзакций получено:', transactions.length]);
    sheet.appendRow(['Транзакций обработано:', processedCount]);

  } catch (e) {
    Logger.log("❌ Ошибка: " + e.message);
    Logger.log("❌ Stack trace: " + e.stack);
    
    // Создаем лист с ошибкой
    const errorSheetName = 'Aspire_Error';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName(errorSheetName);
    
    if (!errorSheet) {
      errorSheet = ss.insertSheet(errorSheetName);
    } else {
      errorSheet.clear();
    }
    
    errorSheet.appendRow(['Ошибка при получении транзакций']);
    errorSheet.appendRow(['Время:', new Date().toISOString()]);
    errorSheet.appendRow(['Сообщение:', e.message]);
    errorSheet.appendRow(['URL:', proxyUrl]);
    
    throw e; // Перебрасываем ошибку дальше
  }
}

// Функция для тестирования прокси
function testProxyConnection() {
  const testUrl = 'https://aspire-proxy-render.onrender.com/ping';
  
  try {
    Logger.log("🧪 Тестируем подключение к прокси...");
    
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    Logger.log("📥 Код ответа: " + code);
    Logger.log("📥 Тело ответа: " + body);
    
    if (code === 200) {
      Logger.log("✅ Прокси работает!");
    } else {
      Logger.log("❌ Прокси недоступен");
    }
    
  } catch (e) {
    Logger.log("❌ Ошибка подключения к прокси: " + e.message);
  }
}

// Функция для получения токена (для отладки)
function testTokenEndpoint() {
  const tokenUrl = 'https://aspire-proxy-render.onrender.com/token';
  
  try {
    Logger.log("🔑 Тестируем получение токена...");
    
    const response = UrlFetchApp.fetch(tokenUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    Logger.log("📥 Код ответа: " + code);
    Logger.log("📥 Токен получен: " + (code === 200 ? "Да" : "Нет"));
    
    if (code === 200) {
      const data = JSON.parse(body);
      Logger.log("🔑 Access token: " + data.access_token.substring(0, 50) + "...");
    }
    
  } catch (e) {
    Logger.log("❌ Ошибка получения токена: " + e.message);
  }
}

// Функция для получения всех транзакций (без фильтрации по account_id)
function fetchAllTransactions() {
  const startDate = '2024-04-01T00:00:00Z';
  const proxyUrl = `https://aspire-proxy-render.onrender.com/aspire?start_date=${encodeURIComponent(startDate)}`;

  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'Content-Type': 'application/json',
      'User-Agent': 'GoogleAppsScript/1.0'
    }
  };

  try {
    Logger.log("📡 Получаем все транзакции...");
    
    const response = UrlFetchApp.fetch(proxyUrl, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    Logger.log("📥 Код ответа: " + code);

    if (code !== 200) {
      throw new Error(`Не удалось получить транзакции. Код ответа: ${code}, Ответ: ${body}`);
    }

    const data = JSON.parse(body);
    const transactions = data.data || [];

    Logger.log(`✅ Получено транзакций: ${transactions.length}`);
    Logger.log(`📊 Всего транзакций в системе: ${data.metadata?.total || 'неизвестно'}`);

    // Создаем лист для всех транзакций
    const sheetName = 'All_Transactions';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    else sheet.clear();

    const headers = [
      'Account ID', 'Date', 'Amount', 'Currency', 'Type', 'Status', 
      'Reference', 'Description', 'Counterparty', 'Balance', 'Category'
    ];
    sheet.appendRow(headers);

    transactions.forEach((tx, index) => {
      try {
        const row = [
          tx.account_id || '',
          tx.datetime || '',
          tx.amount || 0,
          tx.currency_code || '',
          tx.type || '',
          tx.status || '',
          tx.reference || '',
          tx.counterparty_name || '',
          tx.counterparty_name || '',
          tx.balance || 0,
          tx.additional_info?.spend_category || ''
        ];
        
        sheet.appendRow(row);
      } catch (txError) {
        Logger.log(`❌ Ошибка обработки транзакции ${index}: ${txError.message}`);
      }
    });

    Logger.log(`✅ Загружено всех транзакций: ${transactions.length}`);
    
  } catch (e) {
    Logger.log("❌ Ошибка: " + e.message);
  }
} 