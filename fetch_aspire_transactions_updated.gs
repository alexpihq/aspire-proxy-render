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
    const transactions = data.data || [];

    Logger.log(`✅ Получено транзакций: ${transactions.length}`);

    const sheetName = 'Aspire_USD';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    else sheet.clear();

    // Улучшенные заголовки с дополнительными полями
    const headers = [
      'Date', 'Amount (USD)', 'Currency', 'Type', 'Status', 'Reference', 
      'Description', 'Counterparty', 'Balance (USD)', 'Category', 'Card Number'
    ];
    sheet.appendRow(headers);

    // Обработка каждой транзакции
    transactions.forEach((tx, index) => {
      try {
        const row = [
          tx.datetime || '',
          (tx.amount || 0) / 100, // Делим на 100 для конвертации из центов в доллары
          tx.currency_code || '',
          tx.type || '',
          tx.status || '',
          tx.reference || '',
          tx.counterparty_name || '',
          tx.counterparty_name || '',
          (tx.balance || 0) / 100, // Также делим баланс на 100
          tx.additional_info?.spend_category || '',
          tx.additional_info?.card_number || ''
        ];
        
        sheet.appendRow(row);
        
        if (index % 10 === 0) {
          Logger.log(`📊 Обработано транзакций: ${index + 1}/${transactions.length}`);
        }
      } catch (txError) {
        Logger.log(`❌ Ошибка обработки транзакции ${index}: ${txError.message}`);
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

    Logger.log(`✅ Загружено транзакций: ${transactions.length}`);
    Logger.log(`📊 Всего транзакций в системе: ${data.metadata?.total || 'неизвестно'}`);
    
  } catch (e) {
    Logger.log("❌ Ошибка: " + e.message);
    Logger.log("❌ Stack trace: " + e.stack);
  }
}

// Функция для тестирования подключения к прокси
function testProxyConnection() {
  try {
    Logger.log("🧪 Тестируем подключение к прокси...");
    
    const response = UrlFetchApp.fetch('https://aspire-proxy-render.onrender.com/ping', {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    Logger.log(`📡 Код ответа: ${code}`);
    Logger.log(`📡 Ответ: ${body}`);
    
    if (code === 200) {
      Logger.log("✅ Прокси доступен!");
    } else {
      Logger.log("❌ Прокси недоступен");
    }
    
  } catch (e) {
    Logger.log("❌ Ошибка подключения: " + e.message);
  }
}

// Функция для получения токена (для отладки)
function testTokenEndpoint() {
  try {
    Logger.log("🔑 Тестируем получение токена...");
    
    const response = UrlFetchApp.fetch('https://aspire-proxy-render.onrender.com/token', {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    const body = response.getContentText();
    
    Logger.log(`📡 Код ответа: ${code}`);
    
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
      'Account ID', 'Date', 'Amount (USD)', 'Currency', 'Type', 'Status', 
      'Reference', 'Description', 'Counterparty', 'Balance (USD)', 'Category'
    ];
    sheet.appendRow(headers);

    transactions.forEach((tx, index) => {
      try {
        const row = [
          tx.account_id || '',
          tx.datetime || '',
          (tx.amount || 0) / 100, // Делим на 100 для конвертации из центов в доллары
          tx.currency_code || '',
          tx.type || '',
          tx.status || '',
          tx.reference || '',
          tx.counterparty_name || '',
          tx.counterparty_name || '',
          (tx.balance || 0) / 100, // Также делим баланс на 100
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