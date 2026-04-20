
/** Вспомогательный метод для единообразного форматирования имени */
function FormatUserName(u) {
    let name = `${u.NAME || ""} ${u.LAST_NAME || ""}`.trim();
    return name || u.EMAIL || u.LOGIN || `ID ${u.ID}`;
}

/**
 * УНИВЕРСАЛЬНЫЙ СБОРЩИК СТРОКИ КОМАНДЫ ДЛЯ BATCH
 * Реализует правильное кодирование операторов и вложенных массивов (Triple Encoding).
 *
 * @param {string} method - Метод API (например, 'crm.lead.list')
 * @param {Object} params - Объект с параметрами (filter, select, order и т.д.)
 * @return {string} - Полностью закодированная строка для поля cmd
 */
function BuildBatchCmd(method, params = {}) {
    const parts = [];

    /**
     * Рекурсивная функция для превращения объекта в Query String формата Bitrix
     */
    function buildString(obj, prefix) {
        for (const key in obj) {
            if (obj.hasOwnProperty(key)) {
                const k = prefix ? `${prefix}[${key}]` : key;
                const v = obj[key];

                if (v !== null && typeof v === 'object') {
                    buildString(v, k);
                } else {
                    // Критически важное кодирование: encodeURIComponent делает ту самую
                    // работу, чтобы [ ] > < = дошли до ядра Bitrix в целости.
                    parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(v)}`);
                }
            }
        }
    }

    buildString(params);
    return `${method}?${parts.join('&')}`;
}


/**
 * 1. УНИВЕРСАЛЬНЫЙ ВЫЗОВ (REST 2.0)
 * Самый стабильный метод для Лидов, Сделок и Batch-запросов.
 */
function CallBitrix(method, params = {}) {
    const url = BITRIX_URL + method;
    return SendRequest(url, params);
}

/**
 * 2. ВЫЗОВ REST 3.0 (Для Задач и будущего)
 * Используется, когда нужно работать через эндпоинт /api/
 */
function CallBitrix3(method, params = {}) {
    const url = BITRIX_URL_REST_3_0 + method;
    return SendRequest(url, params);
}

/**
 * 3. ВНУТРЕННЯЯ СИСТЕМНАЯ ФУНКЦИЯ
 * Чтобы не дублировать настройки UrlFetchApp
 */
function SendRequest(url, params) {
    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(params),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        return JSON.parse(response.getContentText());
    } catch (e) {
        Logger.log(`⚠ Ошибка запроса к ${url}: ${e.message}`);
        return { error: 'FETCH_ERROR', error_description: e.message };
    }
}

/**
 * УНИВЕРСАЛЬНЫЙ ПАРСЕР ДАТЫ (Улучшенная версия)
 */
function ParseDateToIso(dateInput, timeStr) {
  // Пытаемся создать объект даты, если пришла строка
  let dateObj = (dateInput instanceof Date) ? dateInput : new Date(dateInput);
  
  if (isNaN(dateObj.getTime())) {
    throw new Error(`Критическая ошибка: Значение "${dateInput}" не является датой.`);
  }
  
  const year = dateObj.getFullYear();
  const month = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  const day = ("0" + dateObj.getDate()).slice(-2);
  
  return `${year}-${month}-${day}T${timeStr}+03:00`;
}


/**
 * Преобразует числовой код завершения звонка (CALL_FAILED_CODE) в текстовый статус.
 * Базируется на системных кодах Bitrix24 и анализе реальных JSON-ответов.
 * 
 * @param {string|number} code - Код ошибки или завершения звонка (например, "200", "304").
 * @returns {string} Эмодзи и расшифровка статуса (например, "✅ Успешно", "⏳ Пропущен").
 * 
 * @private
 */
function _getCallStatusText(code) {
  const codes = {
    '200': '✅ Успешный звонок',
    '304': '⏳ Пропущенный звонок',
    '400': '🔍 Ошибка запроса (400)',
    '402': '💸 Недостаточно средств',
    '403': '🚫 Запрещено',
    '404': '🔢 Неверный номер',
    '408': '⏲ Истекло время (408)',
    '423': '🔒 Заблокировано (423)',
    '480': '📵 Временно недоступен',
    '484': '📡 Направление недоступно',
    '486': '🚱 Занято',
    '500': '🖥 Ошибка сервера (500)',
    '503': '⚠️ Линия перегружена (503)',
    '504': '🔌 Нет ответа сервера (504)',
    '603': '📵 Отклонено (603)',
    '603-S': '❌ Отменено (603-S)',
    'OTHER': '❓ Не определено'
  };
  
    const sCode = String(code || "").trim();
    if (!sCode || sCode === "0") return "❓ Неизвестно";

    // Если в ключе лежит готовый перевод - возвращаем его
    if (codes[sCode]) return codes[sCode];

    // Если это просто текст (не цифры), возвращаем как есть с иконкой
    if (isNaN(parseInt(sCode.charAt(0)))) return `💬 ${sCode}`;

    // В остальных случаях пишем "Код: ..."
    return `🆔 Код: ${sCode}`;
}


/**
 * Преобразует числовой тип звонка (CALL_TYPE) в текстовое описание направления.
 * Помогает отличить обычные звонки от инфозвонков и перенаправлений.
 * 
 * @param {string|number} type - Идентификатор типа звонка (например, "1", "2", "5").
 * @returns {string} Эмодзи и название направления (например, "📥 Входящий", "🤖 Инфозвонок").
 * 
 * @private
 */
function _getCallTypeText(type) {
  const types = {
    '1': 'исходящий',              // Первый в вашем списке
    '2': 'входящий',               // Второй
    '3': 'входящий',               // "входящий с перенаправлением" (для нас это входящий)
    '4': 'обратный',               // Четвертый
    '5': 'информационный'          // Пятый (тот, что мы исключаем)
  };
  return types[String(type)] || `код: ${type}`;
}

/**
 * @typedef {Object} PreparedCall
 * @property {string} ID - Уникальный идентификатор записи звонка в Битрикс24.
 * @property {string} DATE - Дата и время начала звонка в формате ISO 8601.
 * @property {string} USER - Имя сотрудника (мапированное) или его ID.
 * @property {string} PHONE - Номер телефона клиента.
 * @property {string} TYPE - Направление (📥 Входящий, 📤 Исходящий и т.д.).
 * @property {string} STATUS - Результат вызова (✅ Успешно, ⏳ Пропущен и т.д.).
 * @property {string} DURATION - Длительность в формате ММ:СС.
 * @property {string} LINE - Название линии или номер портала (например, verba-ats).
 * @property {number} COST - Стоимость звонка (если передана Битриксом).
 * @property {string} CRM - Ссылка на сущность CRM (Лид: 123, Контакт: 456).
 */

/**
 * ОРКЕСТРАТОР СБОРА ЗВОНКОВ
 * Собирает данные, мапит сотрудников и возвращает массив подготовленных объектов.
 * 
 * @param {Object} period - Объект {start, end} в формате ISO.
 * @returns {PreparedCall[]} Массив подготовленных объектов звонков для записи.
 */
function ExportCallsToSheet(config, period) {
  updateStatus("📞 Телефония: Начинаем сбор данных...");
  
  // 1. Сбор сырых данных
  const rawCalls = GetCallsData(period);
  updateStatus(`📞 Телефония: Получено ${rawCalls.length} записей из API`);

  if (rawCalls.length === 0) return [];

  // 2. Подготовка сотрудников
  updateStatus("📞 Телефония: Маппинг менеджеров...");
  const userIds = rawCalls.map(c => c.PORTAL_USER_ID).filter(id => id);
  const usersMap = GetUsersMap(userIds);

  // 3. Формирование объектов
  updateStatus("📞 Телефония: Формирование отчета...");
  return rawCalls.map(c => {
    const durationSec = parseInt(c.CALL_DURATION) || 0;
    return {
      ID: String(c.ID),
      DATE: c.CALL_START_DATE,
      USER: c.PORTAL_USER_ID ? (usersMap[c.PORTAL_USER_ID] || c.PORTAL_USER_ID) : "🤖 Система",
      PHONE: c.PHONE_NUMBER,
      TYPE: _getCallTypeText(c.CALL_TYPE),
      STATUS: _getCallStatusText(c.CALL_FAILED_CODE), // Красивый текст для таблицы
      RAW_STATUS: String(c.CALL_FAILED_CODE),         // Чистый код для логики (304, 200 и т.д.)
      DURATION: Math.floor(durationSec / 60) + ":" + ("0" + (durationSec % 60)).slice(-2),
      LINE: c.PORTAL_NUMBER || "Внешняя линия",
      COST: parseFloat(c.COST) || 0,
      CRM: c.CRM_ENTITY_TYPE ? `${c.CRM_ENTITY_TYPE}: ${c.CRM_ENTITY_ID}` : ""
    };
  });
}




/**
 * Потоковое получение статистики звонков из Bitrix24 (метод voximplant.statistic.get).
 * Использует фильтрацию по времени начала звонка для обхода ограничений пагинации.
 * Реализована защита от пропуска данных при наличии нескольких звонков в одну секунду.
 * 
 * @param {Object} period - Объект с границами временного периода.
 * @param {string} period.start - Дата начала в формате ISO 8601 (например, "2024-01-01T00:00:00+03:00").
 * @param {string} period.end - Дата конца в формате ISO 8601 (например, "2024-01-31T23:59:59+03:00").
 * 
 * @returns {Object[]} Массив "сырых" объектов звонков из API Bitrix24.
 * 
 * @example
 * const period = { start: "2024-10-01T00:00:00+03:00", end: "2024-10-01T23:59:59+03:00" };
 * const calls = GetCallsData(period);
 */
function GetCallsData(period) {
  const allCalls = [];
  const processedIds = new Set();
  
  // Используем чистый ISO формат из периода
  let lastCallTime = period.start; 
  let hasMore = true;

  console.log(`📞 СТАРТ СБОРА ЗВОНКОВ: с [${period.start}] по [${period.end}]`);

  while (hasMore) {
    const response = CallBitrix('voximplant.statistic.get', {
      FILTER: {
        ">=CALL_START_DATE": lastCallTime,
        "<CALL_START_DATE": period.end
      },
      SORT: "CALL_START_DATE",
      ORDER: "ASC"
    });

    const calls = response.result || [];

    if (calls.length > 0) {
      calls.forEach(call => {
        // Дедупликация на случай нахлеста секунд
        if (!processedIds.has(call.ID)) {
          allCalls.push(call);
          processedIds.add(call.ID);
        }
      });

      // Обновляем отметку времени для следующей страницы
      lastCallTime = calls[calls.length - 1].CALL_START_DATE;
      
      // Если пришло меньше 50 — мы дошли до конца TO_DATE
      if (calls.length < 50) hasMore = false;
      
      // Очистка кэша ID (защита памяти)
      if (processedIds.size > 500) {
        const idsArray = Array.from(processedIds);
        processedIds.clear();
        idsArray.slice(-100).forEach(id => processedIds.add(id));
      }
    } else {
      hasMore = false;
    }
  }

  console.log(`✅ Сбор завершен. Получено звонков: ${allCalls.length}`);
  return allCalls;
}

