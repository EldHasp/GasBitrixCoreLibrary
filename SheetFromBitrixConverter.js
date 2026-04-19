/**
 * ТРАНСФОРМАТОР: Преобразует один объект Битрикс в массив значений для строки.
 * 
 * @param {Object} item - Сырой объект лида/сделки.
 * @param {string[]} headers - Заголовки.
 * @param {Object} liveMap - Карта полей.
 * @param {Object} maps - Справочники + Карта звонков.
 * @param {number[]} dateCols - Индексы колонок с датами.
 * @returns {any[]} Массив значений для строки.
 */
function _transformItemToRow(item, headers, liveMap, maps, dateCols) {
  return headers.map((header, colIdx) => {
    // 1. Виртуальные поля аналитики
    if (header === "Дата первого звонка") {
      const callDate = maps.FIRST_CALL_DATE ? maps.FIRST_CALL_DATE[item.ID] : null;
      return callDate ? new Date(callDate) : "";
    }

    if (header === "Скорость реакции (сек)") {
      const callDate = maps.FIRST_CALL_DATE ? maps.FIRST_CALL_DATE[item.ID] : null;
      if (!callDate || !item.DATE_CREATE) return "";
      return Math.round((new Date(callDate).getTime() - new Date(item.DATE_CREATE).getTime()) / 1000);
    }

    // 2. Стандартные поля
    const fieldData = liveMap[header];
    const techId = fieldData ? fieldData.id : null;
    if (!techId) return ""; 

    let value = item[techId];
    if (value === null || value === undefined || value === "") return "";

    // 3. Маппинг справочников (Сотрудники, Источники)
    if (maps[techId]) {
      if (Array.isArray(value)) {
        value = value.map(id => maps[techId][id.toString()] || id).join("; ");
      } else {
        value = maps[techId][value.toString()] || value;
      }
    }

    // 4. Типизация (Даты, Деньги, Мультиполя)
    const currentVisualIndex = colIdx + 1;
    if (dateCols.includes(currentVisualIndex)) {
      const d = new Date(value);
      return isNaN(d.getTime()) ? value : d;
    }

    if (fieldData.type === "crm_multifield" && Array.isArray(value)) {
      return value.map(v => v.VALUE).filter(v => v).join("; ");
    }
    
    if (fieldData.type === "money" && typeof value === 'string' && value.includes('|')) {
      return parseFloat(value.split('|')[0]);
    }

    // 5. Защита от формул
    if (typeof value === 'string' && ['=', '+', '-', '@'].includes(value[0])) {
      return " " + value;
    }

    return value;
  });
}

/**
 * Подготавливает массив RichTextValue для создания ссылок на сущности CRM.
 * 
 * @param {Object[]} dataItems - Массив лидов/сделок.
 * @param {string} baseUrl - Базовый URL портала (Лиды или Сделки).
 * @returns {Array<GoogleAppsScript.Spreadsheet.RichTextValue[]>}
 * @private
 */
function _prepareRichTextLinks(dataItems, baseUrl) {
  return dataItems.map(item => {
    const title = item.TITLE || item.NAME || `ID: ${item.ID}`;
    const url = baseUrl ? `${baseUrl}${item.ID}/` : "";
    let displayTitle = title.toString();
    
    // Защита от автозапуска формул Excel/Sheets
    if (['+', '-', '=', '@'].includes(displayTitle[0])) {
      displayTitle = " " + displayTitle;
    }
    
    const rt = SpreadsheetApp.newRichTextValue().setText(displayTitle);
    if (url) rt.setLinkUrl(url);
    return [rt.build()];
  });
}

/**
 * Накладывает форматирование на вставленный диапазон.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист.
 * @param {number} startRow - Строка начала вставки.
 * @param {number} numRows - Количество вставленных строк.
 * @param {string[]} headers - Заголовки.
 * @param {number[]} dateCols - Базовые индексы колонок с датами.
 * @private
 */
function _applyTableFormatting(sheet, startRow, numRows, headers, dateCols) {
  // 1. Форматирование всех дат (включая виртуальную колонку звонка)
  const callDateIdx = headers.indexOf("Дата первого звонка");
  const extendedDateCols = [...dateCols];
  if (callDateIdx !== -1) extendedDateCols.push(callDateIdx + 1);

  if (extendedDateCols.length > 0) {
    extendedDateCols.forEach(colIndex => {
      sheet.getRange(startRow, colIndex, numRows, 1).setNumberFormat("dd.mm.yyyy hh:mm");
    });
  }

  // 2. Форматирование ID (число без десятичных знаков)
  sheet.getRange(startRow, 1, numRows, 1).setNumberFormat("0");

  // 3. Форматирование "Скорости реакции" (если есть)
  const speedIdx = headers.indexOf("Скорость реакции (сек)");
  if (speedIdx !== -1) {
    sheet.getRange(startRow, speedIdx + 1, numRows, 1).setNumberFormat("#,##0");
  }
}
