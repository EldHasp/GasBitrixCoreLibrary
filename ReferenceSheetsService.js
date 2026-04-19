/**
 * Справочник: Выгрузка всех источников лидов.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Целевой лист из config.sourcesSheet
 */
function GetSourcesReference(sheet) {
  if (!sheet) return;

  console.log(`📢 Запрос источников (SOURCE) для листа: "${sheet.getName()}"...`);

  // Получаем список из crm.status, фильтруя только сущности типа SOURCE
  const response = CallBitrix('crm.status.list', { 
    "filter": { "ENTITY_ID": "SOURCE" } 
  });

  if (!response || !response.result) {
    console.error("❌ Не удалось получить справочник источников.");
    return;
  }

  // Заголовки (Колонки DTO)
  const header = ["STATUS_ID (Код)", "NAME (Название)", "ID (Системный)"];
  
  // Трансформация (LINQ .Select)
  const rows = response.result.map(s => [
    s.STATUS_ID, // Техническое имя (часто используется как значение поля)
    s.NAME,      // Название для людей
    s.ID         // Числовой ID записи
  ]);

  // Пакетная запись
  sheet.clear();
  
  // Шапка (светло-желтый цвет)
  sheet.getRange(1, 1, 1, header.length)
       .setValues([header])
       .setFontWeight("bold")
       .setBackground("#fff2cc");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // UI оформление
  sheet.autoResizeColumns(1, header.length);
  sheet.setFrozenRows(1);
  
  console.log(`✅ Справочник источников обновлен. Записей: ${rows.length}`);
}

/**
 * ВЫГРУЗКА ЭЛЕМЕНТОВ ИНФОБЛОКА НА ЛИСТ-СПРАВОЧНИК
 * 
 * Данный метод решает проблему получения названий для полей типа "iblock_element" (Город/Офис).
 * 
 * 📔 ИСТОРИЯ ТРУДНОСТЕЙ И РЕШЕНИЙ:
 * 
 * 1. Ограничение метода crm.lead.fields:
 *    - Проблема: Стандартный запрос полей лида возвращает только структуру поля. 
 *      В отличие от простых списков (enumeration), варианты значений для iblock_element 
 *      не приходят в общем ответе, так как они хранятся в отдельном модуле "Списки".
 *    - Решение: Использование специализированного метода 'lists.element.get'.
 * 
 * 2. Поиск технического адреса (IBLOCK_ID):
 *    - Проблема: Для получения данных скрипту нужно знать ID инфоблока и его тип. 
 *    - Решение: В ходе отладки через crm.lead.userfield.get был найден точный адрес: 
 *      IBLOCK_ID = 128 и IBLOCK_TYPE_ID = "bitrix_processes".
 * 
 * 3. Отсутствие пагинации в простых вызовах:
 *    - Нюанс: Метод 'lists.element.get' по умолчанию возвращает все элементы (если их не тысячи), 
 *      что удобно для небольших справочников офисов. Для очень больших списков потребуется 
 *      добавление параметра 'start'.
 * 
 * 4. Контроль точности данных:
 *    - Решение: Создание физического листа-справочника CITIES_REFERENCE_SHEET ("CITIES REFERENCE"). 
 *      Это позволяет визуально сопоставить ID элемента (например, 109734) с его 
 *      названием ("Офис Москва") перед внедрением в основной код парсинга.
 */
/**
 * Справочник: Выгрузка элементов инфоблока (Офисы/Города).
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Целевой лист из config.officesSheet
 */
function GetOfficesReference(sheet) {
  if (!sheet) return;

  // Используем глобальные константы из Config.gs (128 и "bitrix_processes")
  console.log(`🏙 Запрос элементов инфоблока ${IBLOCK_OFFICES_ID} для листа: "${sheet.getName()}"...`);

  // Запрос к модулю "Списки" (Инфоблоки)
  const response = CallBitrix('lists.element.get', { 
    "IBLOCK_TYPE_ID": IBLOCK_OFFICES_TYPE, 
    "IBLOCK_ID": IBLOCK_OFFICES_ID 
  });

  // Проверка результата (аналог if (response?.result == null))
  if (!response || !response.result) {
    console.error(`❌ Офисы не получены. Проверьте ID (${IBLOCK_OFFICES_ID}) и права вебхука.`);
    return;
  }

  const header = ["ID Офиса (Элемента)", "Название офиса (Город)"];
  
  // Трансформация: извлекаем только нужные поля (LINQ .Select)
  const rows = response.result.map(item => [
    item.ID,     // Технический ID (string)
    item.NAME    // Название города/офиса
  ]);

  // Очистка и запись данных
  sheet.clear();
  
  // Рисуем шапку
  sheet.getRange(1, 1, 1, header.length)
       .setValues([header])
       .setFontWeight("bold")
       .setBackground("#cfe2f3");
       
  // Записываем строки
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // Настройка UI
  sheet.autoResizeColumns(1, header.length);
  sheet.setFrozenRows(1);
  
  console.log(`✅ Справочник офисов обновлен. Найдено элементов: ${rows.length}`);
}

/**
 * СБОР ПОЛНОЙ ИНФОРМАЦИИ О ПОЛЯХ ЛИДА И СОЗДАНИЕ ИНТЕРАКТИВНОЙ ШПАРГАЛКИ
 * 
 * 📔 ИСТОРИЯ ТРУДНОСТЕЙ И РЕШЕНИЙ:
 * 
 * 1. Выбор метода API (crm.lead.userfield.list vs crm.lead.fields):
 *    - Проблема: Метод 'userfield.list' возвращает ТОЛЬКО пользовательские поля.
 *    - Решение: Использование 'crm.lead.fields'. Он возвращает ПОЛНЫЙ список (системные + пользовательские).
 * 
 * 2. Проблема визуализации ссылок (RichText vs Formula):
 *    - Проблема: Формула =HYPERLINK() неудобна для копирования. RichText часто записывался некорректно.
 *    - Решение: Двухэтапная запись: сначала .setValues(), затем наложение .setRichTextValues() 
 *      с принудительным синим стилем (setForegroundColor) и подчеркиванием.
 * 
 * 3. Потеря данных для парсинга (Списки и Множественность):
 *    - Проблема: При выгрузке 17 000 лидов нужно знать структуру поля (массив или строка).
 *    - Решение: Добавление колонок с флагами (isMultiple) и содержимым списков (items).
 * 
 * 4. Унификация констант:
 *    - Решение: Вынос имен листов и базовых URL в Config.gs (FIELDS_INFO_SHEET).
 * 
 * 5. Конфликт названий в пользовательских полях (Ловушка 'title'):
 *    - Проблема: У полей UF_CRM_ в свойстве 'title' часто хранится технический ID, 
 *      а русское название спрятано во вложенных метках. Обычный парсинг выдавал ID вместо имени.
 *    - Решение: Реализована логика приоритета: для полей UF_... принудительно берем 
 *      названия из 'listLabel' или 'formLabel', игнорируя 'title'.
 */
/**
 * Справочник: Выгрузка структуры всех полей Лида.
 * Помогает понять, какие системные и кастомные (UF_*) поля доступны.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист из config.leadFieldsSheet
 */
function GetFieldsInfo(sheet) {
  if (!sheet) return;

  console.log(`📋 Запрос структуры полей Лида для листа: "${sheet.getName()}"...`);

  // Запрос всех доступных полей сущности LEAD
  const response = CallBitrix('crm.lead.fields', {});

  if (!response || !response.result) {
    console.error("❌ Не удалось получить структуру полей из Битрикс24.");
    return;
  }

  const fields = response.result;
  const header = ["Техническое ID", "Название (RU)", "Тип", "Только чтение", "Обязательное"];
  
  // Трансформируем объект-словарь в массив строк для таблицы
  const rows = Object.keys(fields).map(key => {
    const f = fields[key];
    return [
      key,                         // Например: UF_CRM_1712123456
      f.title || f.formLabel || "", // Ищем читаемое название
      f.type,                      // Тип данных (string, enumeration, etc.)
      f.isReadOnly ? "🔒" : "",     // Помечаем системные поля
      f.isRequired ? "⭐" : ""       // Помечаем обязательные поля
    ];
  });

  // Запись данных
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length)
       .setValues([header])
       .setFontWeight("bold")
       .setBackground("#ead1dc"); // Нежно-розовый для полей

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // Оформление
  sheet.autoResizeColumns(1, header.length);
  sheet.setFrozenRows(1);
  
  console.log(`✅ Справочник полей Лида обновлен. Всего: ${rows.length}`);
}


/**
 * Выгрузка пользователей в справочник.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Целевой лист из конфига.
 */
function GetUsersReference(sheet) {
  if (!sheet) return;

  let allUsers = [];
  let start = 0;
  let hasMore = true;

  console.log("👥 Запрос пользователей из Битрикс24...");

  // 1. Сбор данных (Пагинация)
  while (hasMore) {
    const response = CallBitrix('user.get', { 
      "SORT": "ID", "ORDER": "ASC", "start": start 
    });

    if (response && response.result) {
      allUsers = allUsers.concat(response.result);
      start = response.next || null;
      hasMore = !!response.next; // Двойное отрицание превращает число в true, а null в false
    } else {
      hasMore = false;
    }
  }

  // 2. Преобразование данных (LINQ Select)
  const header = ["ID", "Имя", "Фамилия", "Полное имя", "Должность", "E-mail", "Активен"];
  
  // .map() создает новый массив строк для таблицы
  const rows = allUsers.map(u => [
    u.ID,
    u.NAME || "",
    u.LAST_NAME || "",
    `${u.NAME || ""} ${u.LAST_NAME || ""}`.trim() || u.LOGIN || `ID ${u.ID}`,
    u.WORK_POSITION || "",
    u.EMAIL || "",
    u.ACTIVE ? "✅" : "❌"
  ]);

  // 3. Вывод в таблицу (Пакетная запись)
  sheet.clear(); // Полная очистка листа (содержимое + форматирование)
  
  // Записываем шапку
  sheet.getRange(1, 1, 1, header.length)
       .setValues([header])
       .setFontWeight("bold")
       .setBackground("#d9ead3");

  // Записываем данные одним махом (как транзакция в БД)
  if (rows.length > 0) {
    // getRange(строка, столбец, количество_строк, количество_столбцов)
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // 4. Оформление (UI)
  sheet.autoResizeColumns(1, header.length);
  sheet.setFrozenRows(1); // Закрепить первую строку (Freeze Panes)
  
  console.log(`✅ Справочник пользователей обновлен: ${rows.length}`);
}
