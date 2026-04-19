/**
 * Преобразование даты в формат ISO 8601 для фильтров Битрикс24.
 * 
 * @param {Date|string} dateInput - Объект даты или строка.
 * @param {string} timeStr - Время в формате "HH:mm:ss" (например, "00:00:00").
 * @returns {string} Дата в формате "YYYY-MM-DDTHH:mm:ss+03:00"
 */
function ParseDateToIso(dateInput, timeStr) {
  // 1. Приведение к типу Date (аналог DateTime.Parse)
  let dateObj = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
  
  if (isNaN(dateObj.getTime())) {
    throw new Error(`Критическая ошибка: Значение "${dateInput}" не является корректной датой.`);
  }
  
  const year = dateObj.getFullYear();
  const month = ("0" + (dateObj.getMonth() + 1)).slice(-2); // Формат MM
  const day = ("0" + dateObj.getDate()).slice(-2);         // Формат DD
  
  // Битрикс требует жесткий формат ISO с часовым поясом (+03:00 — Москва)
  return `${year}-${month}-${day}T${timeStr}+03:00`;
}

/**
 * Форматирует имя пользователя.
 * Реализует логику Fallback: Name -> Email -> Login -> ID.
 * 
 * @param {{NAME?: string, LAST_NAME?: string, EMAIL?: string, LOGIN?: string, ID: number|string}} u 
 * - Объект, содержащий данные о пользователе (из API, из кэша или созданный вручную).
 * @returns {string}
 */
function FormatUserName(u) {
  if (!u) return "Неизвестно";
  let name = `${u.NAME || ""} ${u.LAST_NAME || ""}`.trim();
  return name || u.EMAIL || u.LOGIN || `ID ${u.ID}`;
}

/**
 * УНИВЕРСАЛЬНЫЙ КОНВЕРТЕР (Lazy Loading)
 * 
 * @param {string} fieldId - Технический ID поля (например, 'SOURCE_ID' или 'UF_CRM_123')
 * @param {string} fieldType - Тип поля (из crm.lead.fields)
 * @param {any} value - Сырое значение из API
 * @param {Object} fieldMaps - Глобальный объект со справочниками (хранится в Main)
 * @returns {any} Читаемое значение
 */
function GetReadableValue(fieldId, fieldType, value, fieldMaps) {
  // 1. Быстрый выход для пустых данных
  if (value === null || value === undefined || value === "") return "";

  // 2. Логика справочников (Lazy Loading)
  const typesRequiringMap = ['enumeration', 'crm_status', 'iblock_element'];
  
  if (typesRequiringMap.includes(fieldType) || fieldId === 'ASSIGNED_BY_ID' || fieldId === 'CREATED_BY_ID') {
    
    // Если словаря для этого поля еще нет в коллекции — создаем его
    if (!fieldMaps[fieldId]) {
      console.log(`🚀 Lazy Loading справочника для: ${fieldId} (${fieldType})`);
      
      if (fieldId.includes('BY_ID')) {
        // Спец-обработка для сотрудников (через твой GetUsersMap)
        fieldMaps[fieldId] = GetUsersMap(); 
      } else {
        // Универсальная обработка через твой умный GetDynamicMap
        fieldMaps[fieldId] = GetDynamicMap(fieldType, fieldId);
      }
    }

    const currentMap = fieldMaps[fieldId];

    // Возвращаем расшифровку (поддержка массивов для множественных полей)
    if (Array.isArray(value)) {
      return value.map(id => currentMap[id.toString()] || id).join("; ");
    }
    return currentMap[value.toString()] || value;
  }

  // 3. Стандартные типы (преобразование форматов)
  switch (fieldType) {
    case "date":
    case "datetime":
      const d = new Date(value);
      return isNaN(d.getTime()) ? value : d;

    case "crm_multifield": // Телефоны/Email
      return Array.isArray(value) ? value.map(v => v.VALUE).join("; ") : value;

    case "money":
      return (typeof value === 'string' && value.includes('|')) ? parseFloat(value.split('|')[0]) : value;

    case "boolean":
      return (value === "Y" || value === true) ? "✅" : "❌";

    case "integer":
    case "double":
      return Number(value);

    default:
      // Все остальное (string, char, text)
      return value;
  }
}

