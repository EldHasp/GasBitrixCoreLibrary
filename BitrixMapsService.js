/**
 * УНИВЕРСАЛЬНЫЙ ГЕНЕРАТОР КАРТ (Dictionary Builder)
 * Теперь сам находит нужные инфоблоки для полей типа iblock_element.
 * 
 * @param {string} fieldType - Тип из crm.lead.fields (enumeration, iblock_element, crm_status).
 * @param {string} fieldId - ID поля (UF_CRM_... или SOURCE_ID).
 * @returns {Object<string, string>} Словарь { "ID": "Название" }.
 */
function GetDynamicMap(fieldType, fieldId) {
    let items = [];

    try {
        // 1. ПОЛЬЗОВАТЕЛЬСКИЕ СПИСКИ (Простые Enum)
        if (fieldType === 'enumeration') {
            const resp = CallBitrix('crm.userfield.get', { id: fieldId.replace('UF_CRM_', '') });
            items = (resp && resp.result) ? resp.result.LIST : [];
        }
        
        // 2. СИСТЕМНЫЕ СТАТУСЫ (Источники, Стадии)
        else if (fieldType === 'crm_status') {
            const resp = CallBitrix('crm.status.list', { filter: { ENTITY_ID: fieldId } });
            items = (resp && resp.result) ? resp.result : [];
        }

        // 3. ПРИВЯЗКА К ЭЛЕМЕНТАМ (Инфоблоки/Процессы) — ТВОЙ ГОРОД/ОФИС
        else if (fieldType === 'iblock_element') {
            // ШАГ А: Узнаем, к какому списку привязано поле (Reflection в C#)
            const fieldInfo = CallBitrix('crm.userfield.get', { id: fieldId.replace('UF_CRM_', '') });
            
            if (fieldInfo && fieldInfo.result && fieldInfo.result.SETTINGS) {
                const settings = fieldInfo.result.SETTINGS;
                const iblockId = settings.IBLOCK_ID;
                const iblockTypeId = settings.IBLOCK_TYPE_ID || 'bitrix_processes';

                // ШАГ Б: Загружаем элементы этого списка
                const resp = CallBitrix('lists.element.get', { 
                    "IBLOCK_TYPE_ID": iblockTypeId, 
                    "IBLOCK_ID": iblockId 
                });
                items = (resp && resp.result) ? resp.result : [];
            }
        }
    } catch (e) {
        console.error(`❌ Ошибка GetDynamicMap для ${fieldId}: ${e.message}`);
    }

    // Превращаем массив в Dictionary
    return items.reduce((map, item) => {
        const id = item.ID ? item.ID.toString() : null;
        if (id) {
            // Для списков берем VALUE, для инфоблоков NAME
            map[id] = item.VALUE || item.NAME || id;
        }
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
/**
 * Справочник: Формирование словаря имен пользователей (Dictionary).
 * 
 * @param {string[]|number[]} [userIds=[]] - Список ID для точечной загрузки.
 * @returns {Object<string, string>} Словарь { "ID": "Имя Фамилия" }.
 */
function GetUsersMap(userIds = []) {
    const usersMap = {};
    
    // Вспомогательная функция для форматирования (аналог метода расширения или лямбды)
    const formatName = (u) => {
        let name = `${u.NAME || ""} ${u.LAST_NAME || ""}`.trim();
        return name || u.EMAIL || u.LOGIN || `ID ${u.ID}`;
    };

    // РЕЖИМ 1: Пакетная загрузка по конкретным ID
    if (userIds && userIds.length > 0) {
        const uniqueIds = [...new Set(userIds)].filter(id => id);
        const cmd = {};
        
        // Разбиваем на чанки по 50 (LINQ: .Chunk(50))
        for (let i = 0; i < Math.ceil(uniqueIds.length / 50); i++) {
            const chunk = uniqueIds.slice(i * 50, (i + 1) * 50);
            // Используем наш BuildBatchCmd для кодирования
            cmd[`batch_users_${i}`] = BuildBatchCmd('user.get', { ID: chunk });
            if (i >= 49) break; // Лимит Batch — 50 команд по 50 юзеров (итого 2500 чел)
        }

        const batchRes = CallBitrix('batch', { halt: 0, cmd: cmd });
        const results = (batchRes && batchRes.result && batchRes.result.result) ? batchRes.result.result : {};

        // Собираем результаты из всех пакетов
        Object.keys(results).forEach(key => {
            if (Array.isArray(results[key])) {
                results[key].forEach(u => {
                    usersMap[u.ID.toString()] = formatName(u);
                });
            }
        });
        return usersMap;
    }

    // РЕЖИМ 2: Загрузка всех активных сотрудников (Fallback)
    const response = CallBitrix('user.get', { "filter": { "ACTIVE": "Y" } });
    const allUsers = response.result || [];
    allUsers.forEach(u => {
        usersMap[u.ID.toString()] = formatName(u);
    });

    return usersMap;
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
  updateStatus(`Загрузка периода ${period}`);
  const allLeads = [];
  const collectedIds = new Set();
  
  // 1. ПРИНУДИТЕЛЬНО ДОБАВЛЯЕМ ТЕХНИЧЕСКИЕ ПОЛЯ В SELECT
  // Нам нужны ID, CONTACT_ID, CONTACT_IDS и DATE_CREATE (для скорости реакции)
  const essentialIds = ["ID", "CONTACT_ID", "CONTACT_IDS", "DATE_CREATE"];
  const selectFields = headers.map(h => liveMap[h]?.id).filter(id => id);
  
  // Объединяем визуальные поля с техническими, убирая дубли
  const finalSelect = [...new Set([...selectFields, ...essentialIds])];

  filterableHeaders.forEach(headerName => {
    updateStatus(`Загрузка ${headerName}`);
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
      updateStatus(`Загружено лидов: ${currentStart.length}`);
    }
  });

  return allLeads;
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
  const selectFields = headers.map(h => liveMap[h]?.id).filter(id => id);

  // 1. ПЕРВЫЙ ЭТАП: Забираем "Дата создания" (DATE_CREATE) как базу
  // Предполагаем, что DATE_CREATE — это первый элемент или находим его
  const mainHeader = filterableHeaders.find(h => liveMap[h].id === 'DATE_CREATE') || filterableHeaders[0];
  const mainFieldId = liveMap[mainHeader].id;

  console.log(`🚀 Начинаем сбор по основному полю: ${mainHeader}`);
  
  let start = 0;
  let total = 1;
  while (start < total) {
    const res = CallBitrix('crm.lead.list', {
      filter: { [`>=${mainFieldId}`]: period.start, [`<${mainFieldId}`]: period.end },
      select: selectFields,
      start: start
    });
    
    if (res.result) {
      res.result.forEach(lead => {
        allLeads.push(lead);
        collectedIds.add(lead.ID);
      });
    }
    total = res.total || 0;
    start += 50;
  }
  console.log(`✅ Основной сбор завершен: ${allLeads.length} лидов.`);

  // 2. ВТОРОЙ ЭТАП: Проверяем остальные фильтры только на наличие НОВЫХ ID
  const secondaryHeaders = filterableHeaders.filter(h => h !== mainHeader);
  const idsToFetch = new Set();

  secondaryHeaders.forEach(headerName => {
    const fieldId = liveMap[headerName].id;
    let sStart = 0;
    let sTotal = 1;
    
    while (sStart < sTotal) {
      const res = CallBitrix('crm.lead.list', {
        filter: { [`>=${fieldId}`]: period.start, [`<${fieldId}`]: period.end },
        select: ["ID"], // ТЯНЕМ ТОЛЬКО ID (это быстро!)
        start: sStart
      });
      
      if (res.result) {
        res.result.forEach(l => {
          if (!collectedIds.has(l.ID)) {
            idsToFetch.add(l.ID); // Только те, кого еще нет в базе
          }
        });
      }
      sTotal = res.total || 0;
      sStart += 50;
    }
  });

  // 3. ТРЕТИЙ ЭТАП: Дозагрузка недостающих данных (те самые 7%)
  if (idsToFetch.size > 0) {
    console.log(`📡 Догружаем уникальные лиды из других фильтров: ${idsToFetch.size} шт.`);
    const extraIds = Array.from(idsToFetch);
    
    // Используем Batch по 50 ID, как обсуждали ранее
    let i = 0;
    while (i < extraIds.length) {
      const chunk = extraIds.slice(i, i + 50);
      const res = CallBitrix('crm.lead.list', {
        filter: { "@ID": chunk },
        select: selectFields
      });
      if (res.result) allLeads.push(...res.result);
      i += 50;
    }
  }

  return allLeads;
}