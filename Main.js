/**
 * ГЛАВНЫЙ МЕТОД ВЫГРУЗКИ: Автономный экспорт лидов с динамическим определением типов.
 * 
 * 📔 ИСТОРИЯ РЕШЕНИЙ И ОПТИМИЗАЦИЙ:
 * 
 * 1. Отказ от промежуточных листов:
 *    - Проблема: Использование листа "Все Поля Лидов" требовало его предварительного обновления.
 *    - Решение: Внедрение метода getLiveFieldsMap(). Теперь скрипт запрашивает структуру 
 *      напрямую из crm.lead.fields в оперативную память перед каждым запуском.
 * 
 * 2. Решение проблемы "Ложного форматирования":
 *    - Проблема: Применение формата даты ко всей таблице превращало ID (например, 142204) 
 *      в бессмысленные даты (30.09.2170).
 *    - Решение: Динамическое вычисление индексов колонок (dateColumnIndexes) на основе 
 *      типов данных из API (datetime/date). Формат накладывается точечно.
 * 
 * 3. Умное объединение полей:
 *    - Решение: Списки REQUIRED_HEADERS и FILTERABLE_HEADERS сливаются в один массив 
 *      уникальных ID через Set, что исключает дублирование колонок и лишние данные в SELECT.
 * 
 * 4. Защита от ошибок формул:
 *    - Решение: Все строковые значения проверяются на наличие математических операторов 
 *      (+, -, =, *, /) в начале строки. При обнаружении добавляется апостроф-пробел.
 */
/**
 * ГЛАВНЫЙ МЕТОД: Оркестратор экспорта лидов и звонков из Bitrix24 в Google Таблицы.
 * Выполняет инициализацию, сбор данных с дедупликацией, маппинг ID в читаемые названия,
 * запись в таблицу и синхронизацию справочников.
 * 
 * Управляет последовательностью вызовов модулей и отображением прогресса.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Объект активной таблицы клиента.
 * @param {boolean} [isManual=false] - Флаг способа запуска. 
 *    true — ручной (показывает диалог прогресса и алерты), 
 *    false — автоматический (триггер, тихий режим).
 */

function ExportLeadsToSheet(ss, isManual = false) {
  const startTime = new Date();

  CacheService.getUserCache().remove('status');
  SpreadsheetApp.flush();
  
  if (isManual) {
    updateStatus("⚙️ Инициализация..."); 
    const html = HtmlService.createHtmlOutputFromFile('Progress').setWidth(350).setHeight(500);
    SpreadsheetApp.getUi().showModelessDialog(html, '🚀 Синхронизация с Bitrix24');
  }

  try {
    const config = _initializeConfig(ss);
    if (!config) throw new Error("Не удалось загрузить конфигурацию.");

    const period = _prepareExportPeriod(config);
    const liveMap = GetLiveFieldsMap();

    // 1. Сбор лидов
    updateStatus("🛰 Загрузка лидов из Битрикс24...");
    const allLeads = GetLeadsByFiltersMap(period, config.requiredHeaders, liveMap, config.filterableHeaders, updateStatus);
    if (!allLeads || allLeads.length === 0)
    {
       updateStatus("⚠️ Лиды за период не найдены.");
       return;
    }

    // 2. Обработка звонков (Модуль 2)
    const callsData = _processCalls(config, period, updateStatus);

    // 3. Запись лидов с аналитикой (Модуль 3)
    const maps = _processLeads(config, period, allLeads, callsData, liveMap, updateStatus);

    // 4. Справочники
    updateStatus("🔄 Синхронизация справочников...");
    SyncSystemReferences(config);
    
    const currentSourcesNames = allLeads.map(l => maps.SOURCE_ID[l.SOURCE_ID] || l.SOURCE_ID);
    const updatedRange = SyncSourcesReference(config.sourcesRange, currentSourcesNames);

    // 5. Финальные уведомления (Модуль 4)
    _handleMetadataNotification(ss, updatedRange, startTime, allLeads.length, updateStatus);


    SpreadsheetApp.flush();
    // Добавляем 🏆 в сообщение успеха
    const duration = ((new Date() - startTime) / 1000).toFixed(1);
    updateStatus(`🏆 Готово! Обработано за ${duration} сек.`); 

  } catch (e) {
    updateStatus(`❌ Ошибка: ${e.message}`);
    console.error(e);
  }
}

/**
 * Глобальный метод логирования состояния процесса для UI-интерфейса.
 * 
 * ФУНКЦИОНАЛ:
 * 1. Накапливает сообщения в массив, формируя историю (лог) событий.
 * 2. Использует CacheService для передачи данных между серверным скриптом и HTML-окном.
 * 3. Автоматически ограничивает длину лога (последние 10 записей) для соблюдения лимитов кэша (100 КБ).
 * 4. Дублирует сообщения в системный журнал Execution Log для отладки.
 * 
 * @param {string} msg - Текст сообщения (например: "🛰 Загрузка лидов...", "🏆 Готово!").
 * 
 * Накапливает сообщения с временной меткой в одну строку через перенос строки \n.
 * @example
 * updateStatus("📞 Загрузка статистики звонков...");
 * 
 * @returns {void}
 * 
 * @param {string} msg - Текст сообщения.
 */
function updateStatus(msg) {
  const cache = CacheService.getUserCache();
  const logKey = 'status'; 
  
  // 1. Формируем временную метку [14:30:05]
  const now = new Date();
  const timeStr = "[" + now.getHours().toString().padStart(2, '0') + ":" + 
                  now.getMinutes().toString().padStart(2, '0') + ":" + 
                  now.getSeconds().toString().padStart(2, '0') + "] ";

  // 2. Читаем текущий лог
  let currentLog = cache.get(logKey) || "";
  
  // 3. Склеиваем: Старый лог + Новая строка с временем
  const newEntry = timeStr + msg;
  
  // Проверка на дубликат (чтобы не писать одно и то же время/статус дважды в секунду)
  if (!currentLog.endsWith(msg)) {
    currentLog = currentLog ? currentLog + "\n" + newEntry : newEntry;
  }
  
  // 4. Ограничиваем размер (кэш не резиновый, оставляем ~15-20 последних строк)
  if (currentLog.length > 2500) {
    currentLog = "..." + currentLog.slice(-2500);
  }
  
  // 5. Записываем в кэш
  cache.put(logKey, currentLog, 120);
  
  // Дублируем в консоль для отладки
  console.log(`[ST] ${newEntry}`);
}


/**
 * Текущий лог прогресса для UI (кэш-ключ совпадает с {@link updateStatus}).
 * В скрипте таблицы обычно вызывают через «мост», например:
 * `google.script.run.getProcessingStatus()` → `BitrixCore_Library.getLibraryStatus()`,
 * как в Progress.html (опрос каждые 500 мс).
 *
 * @returns {string} Многострочный лог или «Подготовка...», пока экспорт не писал статус.
 */
function getLibraryStatus() {
  try {
    const status = CacheService.getUserCache().get('status');
    return status ? status : "Подготовка...";
  } catch (e) {
    console.error("❌ Ошибка получения статуса из кэша: " + e.message);
    return "⚠️ Ошибка кэша";
  }
}

function initClient(ss) {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Создаем меню
  ui.createMenu('🚀 Битрикс24')
    .addItem('📥 Обновить данные', 'runLeadExportManual')
    .addSeparator() // Визуальный разделитель
    .addItem('⏰ Включить ночную выгрузку', 'runSetupTrigger') // Имя функции из ТАБЛИЦЫ
    .addToUi();
/**
  // 2. Проверяем кэш уведомлений
  const cachedMsg = CacheService.getUserCache().get('PENDING_UPDATE_MSG');
  if (cachedMsg) {
    ui.alert("🌙 Обновление справочников", cachedMsg, ui.ButtonSet.OK);
    CacheService.getUserCache().remove('PENDING_UPDATE_MSG');
  }
*/
}


/**
 * Возвращает значение именованного диапазона или null, если диапазон не найден.
 * Пишет диагностический лог для быстрой проверки конфигурации.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} rangeName
 * @returns {*|null}
 * @private
 */
function _getNamedRangeValue(ss, rangeName) {
  const namedRange = ss.getRangeByName(rangeName);
  const value = namedRange ? namedRange.getValue() : null;
  console.log(`📡 Диапазон [${rangeName}]: ${namedRange ? '✅ Найдено' : '❌ НЕ НАЙДЕНО'}. Значение: ${value}`);
  return value;
}

/**
 * Возвращает корректную дату из именованного диапазона или null.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} rangeName
 * @returns {Date|null}
 * @private
 */
function _getNamedRangeDateValue(ss, rangeName) {
  const value = _getNamedRangeValue(ss, rangeName);
  if (!value) return null;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * Положительное целое из одной ячейки (часы, размер окна выгрузки звонков и т.п.).
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} rangeName
 * @returns {number|null}
 * @private
 */
function _getNamedRangePositiveInteger_(ss, rangeName) {
  const value = _getNamedRangeValue(ss, rangeName);
  if (value === null || value === undefined || value === "") return null;
  const n = typeof value === "number" ? Math.floor(value) : parseInt(String(value).trim(), 10);
  if (isNaN(n) || n <= 0) return null;
  return n;
}

/**
 * Возвращает уникальный список непустых значений из именованного диапазона.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} rangeName
 * @returns {string[]}
 * @private
 */
function _getNamedRangeUniqueList(ss, rangeName) {
  const namedRange = ss.getRangeByName(rangeName);
  if (!namedRange) return [];
  const values = namedRange.getValues().flat().filter(String);
  return [...new Set(values)];
}

/**
 * Сборка объекта конфигурации из именованных диапазонов таблицы.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Объект активной таблицы.
 * @returns {{
 *   ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
 *   leadsSheet: string|null,
 *   callsSheet: string|null,
 *   sourcesSheet: string|null,
 *   officesSheet: string|null,
 *   leadFieldsSheet: string|null,
 *   usersSheet: string|null,
 *   firstDay: Date|null,
 *   lastDay: Date|null,
 *   baseHeaders: string[],
 *   filterableHeaders: string[],
 *   requiredHeaders: string[],
 *   callsBatchWindowHours: number|null,
 *   sourcesRange: GoogleAppsScript.Spreadsheet.Range|null
 * }|null} Объект конфигурации или null при критической ошибке.
 */
function _initializeConfig(ss) {
  try {
    // Присваиваем уточненные имена согласно сущностям
    const leadFieldsSheet = _getNamedRangeValue(ss, "ЛистПолейЛида");
    const firstDay = _getNamedRangeDateValue(ss, "ПервыйДень");
    const lastDay = _getNamedRangeDateValue(ss, "ПоследнийДень");

    // Валидация: если заданы поля лида, нужны даты
    if (leadFieldsSheet && (!firstDay || !lastDay)) {
      throw new Error("Для работы с данными CRM необходимы корректные даты (ПервыйДень/ПоследнийДень).");
    }

    const baseHeaders = _getNamedRangeUniqueList(ss, "ПоляВывода");
    const filterableHeaders = _getNamedRangeUniqueList(ss, "ПоляДатФильтра");

    const config = {
      ss: ss,
      leadsSheet:      _getNamedRangeValue(ss, "ЛистВыгрузкиЛидов"),
      callsSheet:      _getNamedRangeValue(ss, "ЛистВыгрузкиЗвонков"),
      sourcesSheet:    _getNamedRangeValue(ss, "ЛистИсточников"),
      officesSheet:    _getNamedRangeValue(ss, "ЛистОфисов"),
      leadFieldsSheet: leadFieldsSheet,             
      usersSheet:      _getNamedRangeValue(ss, "ЛистСотрудников"),
      firstDay:        firstDay,
      lastDay:         lastDay,
      baseHeaders:     baseHeaders,
      filterableHeaders:   filterableHeaders,
      requiredHeaders: [...new Set(['ID', 'Название лида', 'Стадия', ...baseHeaders, ...filterableHeaders])],
      // Именованный диапазон «РазмерОкнаВыгрузкиЗвонков», часы (например 72). Пусто — fallback в BitrixService.
      callsBatchWindowHours: _getNamedRangePositiveInteger_(ss, "РазмерОкнаВыгрузкиЗвонков"),
      sourcesRange: ss.getRangeByName("ИсточникиПоГруппам")
    };

    console.log("🛠 CONFIG CHECK:", {
        leads: config.leadsSheet,
        calls: config.callsSheet,
        dates: !!config.firstDay,
        callsWindowHours: config.callsBatchWindowHours
      });

    return config;

  } catch (e) {
    console.error(`❌ Ошибка конфигурации: ${e.message}`);
    return null;
  }
}




/**
 * СЛУЖЕБНАЯ: Автоматическое создание и обновление справочников
 * 
 * @param {{
 *   ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
 *   leadsSheet: string|null,
 *   sourcesSheet: string|null,
 *   officesSheet: string|null,
 *   leadFieldsSheet: string|null,
 *   usersSheet: string|null,
 *   firstDay: Date|null,
 *   lastDay: Date|null,
 *   baseHeaders: string[],
 *   filterHeaders: string[],
 *   requiredHeaders: string[]
 * }} config - Объект конфигурации (результат _initializeConfig)
 */
function SyncSystemReferences(config) {
  if (!config) return;

  console.log("📂 Проверка имен листов из конфига:");
  console.log("- Сотрудники:", config.usersSheet);
  console.log("- Офисы:", config.officesSheet);
  console.log("- Источники:", config.sourcesSheet);
  console.log("- Поля:", config.leadFieldsSheet);

  // Описываем маппинг: какое поле конфига какой функции соответствует
  const tasks = [
    { name: config.usersSheet,      action: GetUsersReference,   label: "Сотрудники" },
    { name: config.officesSheet,    action: GetOfficesReference, label: "Офисы" },
    { name: config.sourcesSheet,    action: GetSourcesReference, label: "Источники" },
    { name: config.leadFieldsSheet, action: GetFieldsInfo,       label: "Поля лидов" }
  ];

 tasks.forEach(task => {
    try {
      if (!task.name) {
        console.log(`⏩ Пропуск "${task.label}": имя листа не задано в конфиге.`);
        return;
      }

      let targetSheet = config.ss.getSheetByName(task.name);
      if (!targetSheet) {
        console.log(`🆕 Создаю лист для сущности "${task.label}": "${task.name}"`);
        targetSheet = config.ss.insertSheet(task.name);
      }

      console.log(`📡 Запуск обновления справочника: ${task.label}...`);
      
      // ⚡️ ВАЖНО: вызываем функцию и проверяем, не упала ли она внутри
      task.action(targetSheet); 
      
      console.log(`✅ Справочник "${task.label}" обработан. Проверьте лист "${task.name}".`);
      
    } catch (e) {
      console.warn(`⚠️ Ошибка обновления справочника "${task.label}": ${e.message}`);
    }
  });
}

/**
 * Автоматически создает ежедневный триггер для функции runLeadExport в таблице.
 */
function setupDailySync() {
  const functionName = 'runLeadExport'; // Имя функции в скрипте ТАБЛИЦЫ
  
  // 1. Очищаем старые триггеры на эту функцию (чтобы не дублировать)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(t);
  });

  // 2. Создаем новый триггер: каждый день между 2 и 3 часами ночи
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(2) // Запуск в 02:00
    .create();

  return "🏆 Ночная синхронизация успешно настроена на 02:00!";
}

/**
 * Вычисляет временной интервал для выгрузки в формате [start, end[.
 * Конечная дата сдвигается на +1 день, чтобы захватить весь последний день выбранного периода.
 * 
 * @param {Object} config - Объект конфигурации из _initializeConfig.
 * @param {Date} config.firstDay - Дата начала из таблицы.
 * @param {Date} config.lastDay - Дата конца из таблицы.
 * 
 * @returns {{start: string, end: string}} Объект с ISO-строками для фильтрации в API.
 * @private
 */
function _prepareExportPeriod(config) {
  const dateEndLimit = new Date(config.lastDay);
  dateEndLimit.setDate(dateEndLimit.getDate() + 1);
  return {
    start: ParseDateToIso(config.firstDay, "00:00:00"),
    end: ParseDateToIso(dateEndLimit, "00:00:00")
  };
}

/**
 * Управляет полным циклом обработки звонков: запрос, маппинг и запись на лист.
 * Если лист для звонков не указан в конфигурации, процесс пропускается.
 * 
 * @param {Object} config - Объект конфигурации.
 * @param {{start: string, end: string}} period - Интервал дат.
 * @param {Function} updateStatus - Функция для обновления UI-статуса и логов.
 * 
 * @returns {PreparedCall[]} Массив подготовленных объектов звонков (может быть пустым).
 * @private
 */
/**
 * Управляет полным циклом обработки звонков.
 * Ошибка в этом модуле не прерывает основной цикл, но выводится пользователю.
 */
function _processCalls(config, period, updateStatus) {
  if (!config.callsSheet) return [];
  
  try {
    updateStatus("📞 Загрузка статистики звонков...");
    const callsData = ExportCallsToSheet(config, period, updateStatus); 
    
    if (callsData && callsData.length > 0) {
      updateStatus(`✍️ Запись ${callsData.length} звонков...`);
      let targetCallsSheet = config.ss.getSheetByName(config.callsSheet) || config.ss.insertSheet(config.callsSheet);
      BtxWriteCallsToSheet(targetCallsSheet, callsData);
    }
    return callsData;
    
  } catch (e) {
    const errorMsg = "⚠️ Ошибка звонков: " + e.message;
    updateStatus(errorMsg); // Пользователь увидит это в окне прогресса
    console.error(errorMsg);
    
    // Помечаем в кэше, что была ошибка, чтобы основной метод знал об этом
    CacheService.getUserCache().put('CALLS_ERROR', e.message, 60);
    return []; 
  }
}


/**
 * Выполняет обогащение данных лидов аналитикой звонков и производит запись на лист.
 * Рассчитывает связи Лид-Контакт и находит первый звонок для каждого лида.
 * 
 * @param {Object} config - Объект конфигурации.
 * @param {{start: string, end: string}} period - Интервал дат.
 * @param {Object[]} allLeads - Массив сырых лидов из API.
 * @param {PreparedCall[]} callsData - Массив ранее выгруженных звонков.
 * @param {Object} liveMap - Структура полей Битрикса.
 * @param {Function} updateStatus - Функция статуса.
 * 
 * @returns {Object} Объект со всеми мапами (maps), использованными при записи.
 * @private
 */
function _processLeads(config, period, allLeads, callsData, liveMap, updateStatus) {
  // 1. Аналитика связей
  const contactToLeadsMap = GetContactToLeadsMap(allLeads);
  
  const contactsFound = Object.keys(contactToLeadsMap).length;
  updateStatus(`🔗 Контактов в базе лидов: ${contactsFound}`);
  
  // 2. Строим индекс первых звонков
  const callsIndex = _buildLeadFirstCallsMap(callsData, contactToLeadsMap);

  // 3. Сопоставляем звонки ТОЛЬКО для лидов, созданных в период отчета
  const firstCallsMap = {};
  const reportStartMs = new Date(period.start).getTime();

  allLeads.forEach(lead => {
    const leadCreatedMs = new Date(lead.DATE_CREATE).getTime();

    // ПРОВЕРКА: Проставляем аналитику только если лид "свежий"
    if (leadCreatedMs >= reportStartMs) {
      const leadId = String(lead.ID);
      if (callsIndex[leadId]) {
        firstCallsMap[leadId] = callsIndex[leadId];
      }
    }
  });

  // 4. Подготовка заголовков и маппингов
  const finalHeaders = [...config.requiredHeaders, "Дата первого звонка", "Скорость реакции (сек)"];
  
  // ИСПРАВЛЕНО: Используем let, чтобы можно было создать лист, если его нет
  let targetSheet = config.ss.getSheetByName(config.leadsSheet);
  if (!targetSheet) {
    updateStatus(`⚠️ Лист "${config.leadsSheet}" не найден. Создаю новый...`);
    targetSheet = config.ss.insertSheet(config.leadsSheet);
  }

  const maps = {
    "ASSIGNED_BY_ID": GetUsersMap(allLeads.map(l => l.ASSIGNED_BY_ID)),
    "SOURCE_ID": GetSourcesMap(allLeads.map(l => l.SOURCE_ID)),
    "STATUS_ID": GetLeadStagesMap(allLeads.map(l => l.STATUS_ID)),
    "UF_CRM_1675850186": GetOfficesMap(allLeads.map(l => l.UF_CRM_1675850186)),
    "FIRST_CALL_DATE": firstCallsMap
  };

  // 5. Расчет индексов колонок с датами
  const dateCols = finalHeaders.map((h, i) => 
    (liveMap[h]?.type === 'datetime' || liveMap[h]?.type === 'date' || h === "Дата первого звонка") ? i + 1 : null
  ).filter(i => i);

  // 6. Физическая запись
  updateStatus(`✍️ Запись ${allLeads.length} лидов...`);
  updateStatus(`📈 Лидов ${allLeads.length}; Звонков ${callsData.length}`);
  
  const connectionsFound = Object.keys(firstCallsMap).length;
  updateStatus(`🔍 Связей Лид-Звонок в памяти: ${connectionsFound}`);

  PrepareLeadsSheet(targetSheet, finalHeaders);
  BtxWriteDataToSheet(targetSheet, allLeads, finalHeaders, liveMap, maps, dateCols, 'lead', updateStatus);
  
  return maps; 
}



/**
 * Формирует уведомление о завершении и записывает метаданные для всплывающего окна onOpen.
 * Теперь учитывает ошибки, произошедшие в фоновых модулях (например, звонках).
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Таблица.
 * @param {GoogleAppsScript.Spreadsheet.Range|null} updatedRange - Диапазон новых источников.
 * @param {Date} startTime - Время начала процесса.
 * @param {number} leadsCount - Кол-во выгруженных лидов.
 * @param {Function} updateStatus - Функция статуса.
 * @private
 */
function _handleMetadataNotification(ss, updatedRange, startTime, leadsCount, updateStatus) {
  const duration = ((new Date() - startTime) / 1000).toFixed(1);
  
  // 1. Проверяем наличие ошибок из модуля звонков (через CacheService)
  const callsError = CacheService.getUserCache().get('CALLS_ERROR');
  
  // 2. Формируем базовый текст
  let msg = updatedRange 
    ? `⚠️ Найдено новых источников: ${updatedRange.getNumRows()}.` 
    : `🏆 Выгружено лидов: ${leadsCount}.`;
  
  // 3. Если была ошибка звонков — дописываем её в сообщение
  if (callsError) {
    msg += `\n❌ Ошибка звонков: ${callsError}`;
    CacheService.getUserCache().remove('CALLS_ERROR'); // Очищаем кэш
  }

  msg += ` (Время: ${duration} сек.)`;

  // 4. Запись в метаданные (для всплывающего окна при открытии таблицы)
  ss.getDeveloperMetadata()
    .filter(m => m.getKey() === 'PENDING_UPDATE_MSG')
    .forEach(m => m.remove());
    
  ss.addDeveloperMetadata('PENDING_UPDATE_MSG', msg, SpreadsheetApp.DeveloperMetadataVisibility.PROJECT);
  
  // 5. Обновляем статус в текущем окне прогресса
  updateStatus(msg);
}

/**
 * Создает финальную карту первых звонков для ЛИДОВ с фильтрацией мусора.
 * 
 * @param {PreparedCall[]} callsData - Массив звонков.
 * @param {Object.<string, string[]>} contactToLeadsMap - Карта Контакт -> [Лиды].
 * @returns {Object.<string, string>} Карта { "ID_Лида": "ISO_дата_звонка" }
 */
function _buildLeadFirstCallsMap(callsData, contactToLeadsMap) {
  const leadCallsIndex = {};
  let skippedCount = 0;

  const updateLead = (leadId, callDate) => {
    const sId = String(leadId);
    const newTime = new Date(callDate).getTime();
    if (!leadCallsIndex[sId] || newTime < new Date(leadCallsIndex[sId]).getTime()) {
      leadCallsIndex[sId] = callDate;
    }
  };

  callsData.forEach(call => {
    if (!call.DATE) return;

    const entityType = call.CRM_ENTITY_TYPE ? String(call.CRM_ENTITY_TYPE).toUpperCase() : "";
    const entityIdRaw = call.CRM_ENTITY_ID;
    const entityId = entityIdRaw != null && entityIdRaw !== "" ? String(entityIdRaw).trim() : "";
    if (!entityId || entityId === "0") return;

    const statusCode = String(call.RAW_STATUS);

    // --- 1. Общий мусор (роботы, инфозвонки) ---
    const userLabel = call.USER != null ? String(call.USER) : "";
    if (userLabel.indexOf("Система") !== -1 || call.TYPE === "информационный") {
      skippedCount++;
      return;
    }

    // Входящие: в зачёт первого звонка только успешные (200). Исходящие / обратные — любые коды.
    if (call.TYPE === "входящий" && statusCode !== "200") {
      skippedCount++;
      return;
    }

    // --- 2. РАСПРЕДЕЛЕНИЕ — поля как в ответе voximplant (CRM_ENTITY_TYPE / CRM_ENTITY_ID), не строка CRM с листа
    if (entityType === "LEAD") {
      updateLead(entityId, call.DATE);
    } else if (entityType === "CONTACT") {
      const relatedLeads = contactToLeadsMap[entityId];
      if (relatedLeads) {
        relatedLeads.forEach(lId => updateLead(lId, call.DATE));
      }
    }
  });


  updateStatus(`🧹 Очистка: Пропущено ${skippedCount} звонков (робот/инфо или неуспешный входящий).`);
  return leadCallsIndex;
}

