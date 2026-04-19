/**
 * ТЕСТ: Запись уведомления всеми способами для проверки видимости
 */
function TestAllStorageMethods(ss) {
  const testMsg = "Проверка связи от " + new Date().toLocaleTimeString();
  
  // 1. Свойства Пользователя (User)
  PropertiesService.getUserProperties().setProperty('TEST_USER_PROP', "USER: " + testMsg);
  
  // 2. Свойства Документа (Document)
  PropertiesService.getDocumentProperties().setProperty('TEST_DOC_PROP', "DOC: " + testMsg);
  
  // 3. Метаданные (Developer Metadata)
  // Предварительно удалим старые, если были
  ss.getDeveloperMetadata().filter(m => m.getKey() === 'TEST_META_PROP').forEach(m => m.remove());
  ss.addDeveloperMetadata('TEST_META_PROP', "META: " + testMsg, SpreadsheetApp.DeveloperMetadataVisibility.PROJECT);
  
  console.log("✅ Библиотека: Записала данные всеми тремя способами.");
}

function simulateLibraryMessage() {
  SpreadsheetApp.getActiveSpreadsheet().addDeveloperMetadata(
    'PENDING_UPDATE_MSG', 
    'ТЕСТ: Найдено 5 новых источников!', 
    SpreadsheetApp.DeveloperMetadataVisibility.PROJECT
  );
}

/**
 * ТЕСТ: Выгружает последние 50 звонков в файл JSON на Google Drive.
 */
function debugSaveCallsToFile() {
  const response = CallBitrix('voximplant.statistic.get', {
    SORT: "CALL_START_DATE",
    ORDER: "DESC"
  });

  const rawData = response.result || [];
  
  if (rawData.length > 0) {
    const fileName = `Bitrix_Calls_Debug_${new Date().getTime()}.json`;
    // Создаем файл на Google Drive
    const file = DriveApp.createFile(fileName, JSON.stringify(rawData, null, 2), MimeType.PLAIN_TEXT);
    
    console.log(`✅ Файл создан: ${file.getName()}`);
    console.log(`🔗 Ссылка: ${file.getUrl()}`);
    
    return file.getUrl();
  } else {
    console.warn("Звонки не найдены.");
    return null;
  }
}

/**
 * ТЕСТ: Проверка фильтрации по прямому примеру из документации.
 */
function debugTestExactDocFilter() {
  // Пробуем максимально "чистый" фильтр
  const testParams = {
    FILTER: {
       ">=CALL_START_DATE": "2026-04-01T00:00:00+03:00"
    },
    SORT: "ID",
    ORDER: "ASC"
  };

  console.log("📡 Отправка тестового запроса...");
  const response = CallBitrix('voximplant.statistic.get', testParams);
  
  const result = response.result || [];
  
  if (result.length > 0) {
    console.log(`✅ Получено звонков: ${result.length}`);
    console.log(`📅 Дата ПЕРВОГО звонка в ответе: ${result[0].CALL_START_DATE}`);
    console.log(`🆔 ID ПЕРВОГО звонка: ${result[0].ID}`);
    
    if (result[0].CALL_START_DATE.includes("2026-04-01")) {
      console.log("🎯 ФОРМАТ ИЗ ДОКУМЕНТАЦИИ СРАБОТАЛ!");
    } else {
      console.log("❌ ФИЛЬТР ПРОИГНОРИРОВАН (снова вернул старые данные)");
    }
  } else {
    console.warn("📭 Звонков не найдено.");
  }
}


