/**
 * ПОЛУЧЕНИЕ КОЛЛЕКЦИИ ИМЕН ПОЛЬЗОВАТЕЛЕЙ ПО МАССИВУ ИХ ID (С ПОДДЕРЖКОЙ BATCH И BuildBatchCmd)
 *
 * Данный метод превращает массив "безликих" ID из лидов в человекочитаемые имена.
 *
 * 📔 ИСТОРИЯ ТРУДНОСТЕЙ И РЕШЕНИЙ:
 *
 * 1. Проблема лимитов (50 записей):
 * - Метод 'user.get' возвращает максимум 50 объектов. Если в выгрузке 17 000 лидов
 * участвуют 51+ сотрудник, обычный запрос потеряет данные.
 * - Решение: Использование BATCH-пакетов. Метод разбивает входной массив ID на пачки
 * по 50 и запрашивает их одним пакетом. Лимит расширен до 2500 уникальных юзеров.
 *
 * 2. Тройное кодирование параметров (ID[]):
 * - Проблема: При передаче массива ID через строковые команды BATCH (cmd),
 * Bitrix игнорирует фильтр, если скобки [] не экранированы должным образом.
 * - Решение: Использование вспомогательной функции BuildBatchCmd. Она автоматически
 * выполняет тройное кодирование (Triple Encoding), превращая массив JavaScript
 * в понятный для PHP-ядра Битрикса формат ID[0]=123&ID[1]=456...
 *
 * 3. Обработка "Анонимных" пользователей (Fallback логика):
 * - Проблема: У приглашенных или внешних сотрудников часто не заполнены поля NAME
 * и LAST_NAME. В таблице вместо имени появлялись пустые ячейки.
 * - Решение: Многоуровневая проверка. Если Имени нет, скрипт последовательно
 * пытается подставить EMAIL, затем LOGIN, и только в крайнем случае — "ID ###".
 *
 * 4. Скорость и Кэширование:
 * - Решение: Метод вызывается ОДИН РАЗ за всю сессию выгрузки 17 000 лидов.
 * Полученная коллекция (Map) хранится в оперативной памяти скрипта, обеспечивая
 * мгновенный доступ к именам без повторных обращений к API.
 */
function GetUsersMap(userIds = []) {
    const usersMap = {};
    const queryParams = {
        filter: { "ACTIVE": "Y" } // По умолчанию только работающие
    };

    // Режим 1: Точечный запрос через Batch (если ID переданы)
    if (userIds && userIds.length > 0) {
        const uniqueIds = [...new Set(userIds)].filter(id => id);
        const cmd = {};
        const chunks = Math.ceil(uniqueIds.length / 50);

        for (let i = 0; i < chunks; i++) {
            const chunk = uniqueIds.slice(i * 50, (i + 1) * 50);
            cmd[`u_${i}`] = BuildBatchCmd('user.get', { ID: chunk });
            if (i >= 49) break;
        }

        const batchRes = CallBitrix('batch', { halt: 0, cmd: cmd });
        const results = batchRes.result?.result || {};

        for (let key in results) {
            if (Array.isArray(results[key])) {
                results[key].forEach(u => {
                    usersMap[u.ID.toString()] = FormatUserName(u);
                });
            }
        }
        return usersMap;
    }

    // Режим 2: Загрузка всех сотрудников (если userIds не передан)
    const response = CallBitrix('user.get', queryParams);
    const allUsers = response.result || [];
    allUsers.forEach(u => {
        usersMap[u.ID.toString()] = FormatUserName(u);
    });

    return usersMap;
}

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
 * ПОЛУЧЕНИЕ КОЛЛЕКЦИИ ОФИСОВ (Инфоблок 128)
 * Возвращает объект { "109734": "Офис Москва", ... }
 * Проблемы и пути решения в комментарии метода GetOfficesReference
 */
function GetOfficesMap(officeIds = []) {
    const params = { IBLOCK_TYPE_ID: 'lists', IBLOCK_ID: 128 };

    if (officeIds && officeIds.length > 0) {
        params['filter'] = { "ID": [...new Set(officeIds)].filter(id => id) };
    }

    const response = CallBitrix('lists.element.get', params);
    const elements = response.result || [];

    return elements.reduce((map, el) => {
        map[el.ID.toString()] = el.NAME;
        return map;
    }, {});
}

/**
 * ПОЛУЧЕНИЕ КОЛЛЕКЦИИ ИСТОЧНИКОВ
 * @return {Object} - { "WEB": "Сайт", "PHONE": "Звонок" }
 */
function GetSourcesMap(sourceIds = []) {
    const response = CallBitrix('crm.status.list', {
        filter: { "ENTITY_ID": "SOURCE" }
    });

    const statuses = response.result || [];
    const filterIds = (sourceIds && sourceIds.length > 0) ? sourceIds.map(String) : null;

    return statuses.reduce((map, st) => {
        if (filterIds && !filterIds.includes(st.STATUS_ID.toString())) return map;
        map[st.STATUS_ID.toString()] = st.NAME;
        return map;
    }, {});
}

/**
 * Получает маппинг полей напрямую из API Битрикс24 (без использования листа).
 * @return {Object} - { "Название RU": {id: "TECH_ID", type: "type"}, ... }
 */
function GetLiveFieldsMap() {
    console.log("📡 Запрос структуры полей напрямую из CRM...");
    const response = CallBitrix('crm.lead.fields', {});
    const allFields = response.result;

    if (!allFields) throw new Error("❌ Не удалось получить структуру полей из API.");

    const liveMap = {};
    Object.keys(allFields).forEach(key => {
        const f = allFields[key];
        // Используем ту же логику приоритетов, что мы отладили для GetFieldsInfo
        let russianName = "";
        if (key.indexOf('UF_CRM_') === 0) {
            russianName = f.listLabel || f.formLabel || f.filterLabel || key;
        } else {
            russianName = f.title || key;
        }

        // Создаем расширенный объект данных
        liveMap[russianName.toString()] = {
            id: key,
            type: f.type
        };
    });

    return liveMap;
}

/**
 * Универсальный загрузчик справочников из Bitrix24
 */
function GetDynamicMap(fieldType, fieldId) {
    let items = [];

    try {
        if (fieldType === 'enumeration') {
            // Для пользовательских полей (UF_...) запрашиваем настройки поля
            const resp = CallBitrix('crm.userfield.get', { id: fieldId.replace('UF_CRM_', '') });
            items = resp.result ? resp.result.LIST : [];
        }
        else if (fieldType === 'crm_status') {
            // Для системных статусов (стадии, источники и т.д.)
            const resp = CallBitrix('crm.status.list', { filter: { ENTITY_ID: fieldId } });
            items = resp.result || [];
        }
    } catch (e) {
        console.error(`❌ Ошибка загрузки справочника для ${fieldId}: ${e.message}`);
    }

    // Превращаем массив объектов в мапу { ID: VALUE }
    return items.reduce((map, item) => {
        map[item.ID.toString()] = item.VALUE || item.NAME;
        return map;
    }, {});
}

/**
 * КОМБИНИРОВАННЫЙ СБОР ЛИДОВ ПО НЕСКОЛЬКИМ ФИЛЬТРАМ (С ПОДДЕРЖКОЙ OR-ЛОГИКИ И BATCH)
 *
 * Метод реализует сбор уникальных лидов, попадающих под условия разных полей дат
 * (например, "Дата создания" ИЛИ "Дата квалификации"), исключая дублирование записей.
 *
 * 📔 ИСТОРИЯ РЕШЕНИЙ И ОПТИМИЗАЦИЙ:
 *
 * 1. Реализация логики "ИЛИ" (OR) в REST API:
 * - Проблема: Стандартный фильтр Битрикс24 работает по логике "И" (AND). Нельзя
 * одним запросом найти лиды, созданные ИЛИ квалифицированные в период.
 * - Решение: Последовательный перебор полей из FILTERABLE_HEADERS. Скрипт делает
 * отдельные проходы для каждого поля и объединяет результаты.
 *
 * 2. Использование структуры Set для дедупликации:
 * - Проблема: Метод [].includes() при проверке 17 000+ ID начинает существенно
 * замедлять выполнение (сложность O(n^2)).
 * - Решение: Использование объекта new Set(). Проверка наличия ID (метод .has)
 * выполняется за константное время O(1), что критично для больших выгрузок.
 *
 * 3. Агрегированное логирование:
 * - Решение: Внедрение компактного лога "Всего | Новых | Дублей". Это позволяет
 * визуально оценить пересечение выборок и эффективность каждого фильтра.
 *
 * 4. Инкапсуляция Batch-логики:
 * - Решение: Метод берет на себя всю работу с циклами while и пакетами по 50 команд,
 * возвращая в Main.gs уже "чистый" и расшифрованный массив объектов.
 *
 * @param {Object} period - Объект {start: "ISO", end: "ISO"}.
 * @param {string[]} headers - Массив уникальных заголовков таблицы.
 * @param {Object} liveMap - Живая карта полей из API.
 * @return {Object[]} Массив уникальных объектов лидов.
 */
function GetLeadsByFiltersMap(period, headers, liveMap, filterableHeaders) {
  const allLeads = [];
  const collectedIds = new Set();
  
  // 1. ПРИНУДИТЕЛЬНО ДОБАВЛЯЕМ ТЕХНИЧЕСКИЕ ПОЛЯ В SELECT
  // Нам нужны ID, CONTACT_ID, CONTACT_IDS и DATE_CREATE (для скорости реакции)
  const essentialIds = ["ID", "CONTACT_ID", "CONTACT_IDS", "DATE_CREATE"];
  const selectFields = headers.map(h => liveMap[h]?.id).filter(id => id);
  
  // Объединяем визуальные поля с техническими, убирая дубли
  const finalSelect = [...new Set([...selectFields, ...essentialIds])];

  filterableHeaders.forEach(headerName => {
    const fieldInfo = liveMap[headerName];
    if (!fieldInfo) return;

    const filterFieldId = fieldInfo.id;
    const filter = {
      [`>=${filterFieldId}`]: period.start,
      [`<${filterFieldId}`]: period.end
    };

    let currentStart = 0;
    const probe = CallBitrix('crm.lead.list', { filter: filter, select: ["ID"], limit: 1 });
    const totalForFilter = probe.total || 0;

    if (totalForFilter === 0) return;

    while (currentStart < totalForFilter) {
      const cmd = {};
      let remaining = totalForFilter - currentStart;
      let commandsCount = Math.min(50, Math.ceil(remaining / 50));

      for (let i = 0; i < commandsCount; i++) {
        cmd[`b_${i}`] = BuildBatchCmd('crm.lead.list', {
          filter: filter,
          order: { "ID": "ASC" },
          start: currentStart + (i * 50),
          select: finalSelect // ✅ Используем расширенный список полей
        });
      }

      const response = CallBitrix('batch', { halt: 0, cmd: cmd });
      const results = response.result.result;

      for (let key in results) {
        if (results[key]) {
          const batchLeads = results[key];

          batchLeads.forEach(lead => {
            if (!collectedIds.has(lead.ID)) {
              
              // 2. ФОРМИРУЕМ ВИРТУАЛЬНОЕ ПОЛЕ ALL_CONTACT_IDS
              const contactSet = new Set();
              if (lead.CONTACT_ID && lead.CONTACT_ID !== "0") contactSet.add(String(lead.CONTACT_ID));
              if (Array.isArray(lead.CONTACT_IDS)) {
                lead.CONTACT_IDS.forEach(id => { if (id && id !== "0") contactSet.add(String(id)); });
              }
              
              // Добавляем массив всех контактов прямо в объект лида
              lead.ALL_CONTACT_IDS = Array.from(contactSet);

              allLeads.push(lead);
              collectedIds.add(lead.ID);
            }
          });
        }
      }
      currentStart += (commandsCount * 50);
    }
  });

  return allLeads;
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
    '200': '✅ Успешно',
    '304': '⏳ Пропущен',
    '400': '🔍 Ошибка запроса (400)',
    '403': '🚫 Запрещено',
    '404': '🔢 Неверный номер',
    '408': '⏲ Истекло время (408)',
    '423': '🔒 Заблокировано (423)',
    '480': '📵 Вне зоны',
    '486': '🚱 Занято',
    '500': '🖥 Ошибка сервера (500)',
    '503': '⚠️ Перегрузка',
    '504': '🔌 Нет ответа сервера (504)',
    '603': '📵 Отклонено (603)',
    '603-S': '❌ Отменено (603-S)'
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
      STATUS: _getCallStatusText(c.CALL_FAILED_CODE),
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




/**
 * Создает карту связей Лид <-> Контакт
 * @param {Array} allLeads - Массив сырых лидов.
 * @returns {Object} { "CONTACT_ID": ["LEAD_ID_1", "LEAD_ID_2"] }
 */
function GetContactToLeadsMap(allLeads) {
  const contactToLeads = {};
  let leadsWithContacts = 0;

  allLeads.forEach(lead => {
    const contactIds = lead.ALL_CONTACT_IDS;
    if (contactIds && contactIds.length > 0) {
      leadsWithContacts++;
      contactIds.forEach(cId => {
        if (!contactToLeads[cId]) contactToLeads[cId] = [];
        contactToLeads[cId].push(String(lead.ID));
      });
    }
  });

  updateStatus(`🔗 Аналитика: Подготовлено ${leadsWithContacts} лидов для связи со звонками.`);
  return contactToLeads;
}



/**
 * Анализирует массив звонков и сопоставляет их с лидами (напрямую или через связанные контакты).
 * Находит дату самого раннего звонка для каждого лида.
 * 
 * @param {Object[]} rawCalls - Массив сырых объектов звонков из Bitrix24 (voximplant.statistic.get).
 * @param {Object.<string, string[]>} contactToLeadsMap - Карта связей Контакт -> [Лиды].
 *        Пример: { "123": ["5501", "5502"], "124": ["5503"] }
 * 
 * @returns {Object.<string, string>} Карта первых звонков для лидов.
 *          Ключ — ID лида, значение — ISO дата самого раннего звонка.
 *          Пример: { "5501": "2024-10-25T14:30:05+03:00" }
 */
function GetLeadFirstCallMap(rawCalls, contactToLeadsMap) {
  const leadFirstCall = {};

  const updateLeadDate = (leadId, callDate) => {
    if (!leadId) return;
    const sId = String(leadId);
    const newTime = new Date(callDate).getTime();
    if (!leadFirstCall[sId] || newTime < new Date(leadFirstCall[sId]).getTime()) {
      leadFirstCall[sId] = callDate;
    }
  };

  rawCalls.forEach(call => {
    // Берем данные напрямую из свойств, которые мы видим на скриншоте
    const entityType = call.CRM_ENTITY_TYPE; // "LEAD" или "CONTACT"
    const entityId = call.CRM_ENTITY_ID;     // "144838" (строка)
    const callDate = call.CALL_START_DATE;

    if (!entityId || entityId === "0" || !callDate) return;

    if (entityType === 'LEAD') {
      updateLeadDate(entityId, callDate);
    } 
    else if (entityType === 'CONTACT') {
      const relatedLeads = contactToLeadsMap[entityId];
      if (relatedLeads && Array.isArray(relatedLeads)) {
        relatedLeads.forEach(lId => updateLeadDate(lId, callDate));
      }
    }
  });
  return leadFirstCall;
}

