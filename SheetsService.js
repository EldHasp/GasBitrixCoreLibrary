
/**
 * УНИВЕРСАЛЬНЫЙ МЕТОД ЗАПИСИ ДАННЫХ (Оптимизированный для больших объемов)
 * 
 * @param {Sheet} sheet - Объект листа.
 * @param {Object[]} dataItems - Массив лидов/сделок из Битрикс.
 * @param {string[]} headers - Заголовки таблицы (uniqueTableHeaders).
 * @param {Object} liveMap - Карта полей из API { "Имя": {id, type} }.
 * @param {Object} maps - Объект справочников { "TECH_ID": {id: value} }.
 * @param {number[]} dateCols - Индексы колонок с датами (1-based).
 * @param {string} entityType - Тип сущности ('lead' или 'deal').
 */
/**
 * УНИВЕРСАЛЬНЫЙ МЕТОД ЗАПИСИ ДАННЫХ (Рефакторинг)
 */
function BtxWriteDataToSheet(sheet, dataItems, headers, liveMap, maps, dateCols, entityType = 'lead') {
  if (!dataItems || !dataItems.length) return;
  const startTime = new Date();

  // 1. Определяем базовые параметры
  const baseUrl = entityType === 'lead' ? BITRIX_URL_LEAD : (entityType === 'deal' ? BITRIX_URL_DEAL : "");
  const titleIdx = headers.indexOf("Название лида") !== -1 ? headers.indexOf("Название лида") : headers.indexOf("Название");

  // 2. Трансформация данных (через _transformItemToRow)
  console.time("⏱ Трансформация");
  const rows = dataItems.map(item => _transformItemToRow(item, headers, liveMap, maps, dateCols));
  console.timeEnd("⏱ Трансформация");

  // 3. Запись значений
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);

  // 4. Наложение RichText ссылок
  if (titleIdx !== -1 && baseUrl) {
    console.time("⏱ Наложение ссылок");
    const richTextRows = _prepareRichTextLinks(dataItems, baseUrl);
    sheet.getRange(startRow, titleIdx + 1, richTextRows.length, 1).setRichTextValues(richTextRows);
    console.timeEnd("⏱ Наложение ссылок");
  }

  // 5. Финальное форматирование
  console.time("⏱ Форматирование");
  _applyTableFormatting(sheet, startRow, rows.length, headers, dateCols);
  console.timeEnd("⏱ Форматирование");

  console.log(`📊 Запись ${dataItems.length} строк завершена за ${((new Date() - startTime) / 1000).toFixed(1)} сек.`);
}



/**
 * Подготовка листа: жесткая очистка всего, что правее нужных колонок
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Объект листа
 * @param {string[]} headers - Массив заголовков
 */
function PrepareLeadsSheet(sheet, headers) {
  // 1. Полная очистка содержимого и форматов
  sheet.clear(); 

  // 2. Если в таблице больше колонок, чем в нашем списке — удаляем лишние
  const maxCols = sheet.getMaxColumns();
  if (maxCols > headers.length) {
    // Удаляем всё, что идет после последней нужной колонки
    sheet.deleteColumns(headers.length + 1, maxCols - headers.length);
  } else if (maxCols < headers.length) {
    // Если колонок не хватает (новый лист) — добавляем их
    sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
  }

  // 3. Записываем свежую шапку
  const range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers])
       .setFontWeight("bold")
       .setBackground("#cfe2f3")
       .setBorder(true, true, true, true, true, true);
  
  // 1. Закрепляем первую строку
  sheet.setFrozenRows(1);
  
  // 2. Включаем фильтр (сначала удаляем старый, если был, чтобы не было конфликтов)
  if (sheet.getFilter()) sheet.getFilter().remove();
  range.createFilter();

  console.log(`✅ Лист "${sheet.getName()}" подготовлен (закрепление и фильтры ок).`);
}

/**
 * Автоматическая синхронизация справочника источников.
 * Добавляет новые элементы в диапазон, сохраняя рамку, и настраивает уведомления.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Range} sourcesRange - Именованный диапазон "ИсточникиПоГруппам".
 * @param {string[]} currentSources - Массив названий источников из текущей выгрузки.
 * @returns {GoogleAppsScript.Spreadsheet.Range|null} Диапазон новых строк или null.
 */
function SyncSourcesReference(sourcesRange, currentSources) {
  if (!sourcesRange) return null;

  const sheet = sourcesRange.getSheet();
  const startRow = sourcesRange.getRow();
  const startCol = sourcesRange.getColumn();
  const numRows = sourcesRange.getNumRows();
  
  // 1. Извлекаем ТОЛЬКО первую колонку и очищаем её
  const existingValues = sourcesRange.getValues().map(row => 
    row[0] ? row[0].toString().trim() : ""
  );
  
  // 2. Оставляем только те источники, которых ТОЧНО нет в первой колонке
  const newItems = [...new Set(currentSources)]
    .map(s => s ? s.toString().trim() : "")
    .filter(s => s !== "" && !existingValues.includes(s));

  if (newItems.length === 0) return null;

  // 3. Вставка перед предпоследней строкой (сохранение рамки)
  const insertIndex = startRow + numRows - 1; 
  sheet.insertRowsAfter(insertIndex - 1, newItems.length);
  
  const targetRange = sheet.getRange(insertIndex, startCol, newItems.length, 1);
  targetRange.setValues(newItems.map(s => [s]))
             .setBackground("#f4cccc") 
             .setFontStyle("italic");

  return targetRange;
}

/**
 * @typedef {Object} PreparedCall
 * @property {string} ID - Уникальный идентификатор записи звонка.
 * @property {string} DATE - Дата и время начала звонка (ISO).
 * @property {string} USER - Имя сотрудника (мапированное).
 * @property {string} PHONE - Номер телефона клиента.
 * @property {string} TYPE - Направление (📥 Входящий, 📤 Исходящий).
 * @property {string} STATUS - Результат вызова (✅ Успешно, ⏳ Пропущен).
 * @property {string} DURATION - Длительность в формате ММ:СС.
 * @property {string} LINE - Название линии (например, verba-ats).
 * @property {number} COST - Стоимость звонка.
 * @property {string} CRM - Ссылка или ID сущности CRM.
 */

/**
 * ПИСАТЕЛЬ: Записывает массив подготовленных объектов звонков на лист.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист для записи.
 * @param {PreparedCall[]} callsData - Массив объектов, подготовленных в оркестраторе.
 */
function BtxWriteCallsToSheet(sheet, callsData) {
  if (!callsData || callsData.length === 0) return;

  // 1. Заголовки (соответствуют полям объекта PreparedCall)
  const headers = ["ID", "Дата вызова", "Сотрудник", "Номер", "Тип", "Статус", "Длительность", "Линия", "Стоимость", "CRM"];
  
  // 2. Преобразование массива объектов PreparedCall в плоский массив для таблицы
  // Теперь здесь нет маппинга сотрудников — он уже сделан в ExportCallsToSheet
  const rows = callsData.map(c => [
    c.ID, 
    c.DATE, 
    c.USER, 
    c.PHONE, 
    c.TYPE, 
    c.STATUS, 
    c.DURATION, 
    c.LINE, 
    c.COST, 
    c.CRM
  ]);

  // 3. Запись заголовков
  // Подготовка листа (аналогично лидам)
  sheet.clear();
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
             .setFontWeight("bold")
             .setBackground("#d9ead3"); // Зеленоватый фон для звонков
  
  // Закрепление и фильтры
  sheet.setFrozenRows(1);
  if (sheet.getFilter()) sheet.getFilter().remove();
  headerRange.createFilter();

  // 4. Запись новых данных
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  // Форматирование дат для звонков
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat("dd.mm.yyyy hh:mm");

  updateStatus("✅ Лист звонков оформлен: закрепление и фильтры включены.");
}



