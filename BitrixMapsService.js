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
 * ПОЛУЧЕНИЕ КОЛЛЕКЦИИ СТАДИЙ ЛИДА (STATUS)
 * @param {string[]|number[]} [statusIds=[]]
 * @returns {Object} - { "NEW": "Новый", "IN_PROCESS": "В работе", ... }
 */
function GetLeadStagesMap(statusIds = []) {
    const response = CallBitrix('crm.status.list', {
        filter: { "ENTITY_ID": "STATUS" }
    });

    const statuses = response.result || [];
    const filterIds = (statusIds && statusIds.length > 0) ? statusIds.map(String) : null;

    return statuses.reduce((map, st) => {
        const key = st.STATUS_ID ? st.STATUS_ID.toString() : null;
        if (!key) return map;
        if (filterIds && !filterIds.includes(key)) return map;
        map[key] = st.NAME || key;
        return map;
    }, {});
}


/**
 * Виртуальное поле ALL_CONTACT_IDS: объединение CONTACT_ID и CONTACT_IDS без дублей.
 * В REST из Битрикса приходит только CONTACT_ID / CONTACT_IDS; ALL_CONTACT_IDS в select не запрашивается.
 * Заполняется один раз после сбора лидов — дальше связь со звонками читает только ALL_CONTACT_IDS.
 *
 * @param {Object} lead — объект лида из crm.lead.list (мутируется).
 */
function fillVirtualAllContactIds(lead) {
  const out = [];
  const add = (v) => {
    if (v === null || v === undefined || v === '' || v === '0') return;
    const s = String(v).trim();
    if (!s) return;
    if (out.indexOf(s) === -1) out.push(s);
  };

  if (lead.CONTACT_ID != null && lead.CONTACT_ID !== '' && lead.CONTACT_ID !== '0') {
    add(lead.CONTACT_ID);
  }

  const multi = lead.CONTACT_IDS;
  if (multi != null) {
    if (Array.isArray(multi)) {
      multi.forEach(add);
    } else if (typeof multi === 'string') {
      multi.split(/[,\s;]+/).forEach(function (part) { add(part); });
    }
  }

  lead.ALL_CONTACT_IDS = out;
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
    if (lead.ALL_CONTACT_IDS === undefined) {
      fillVirtualAllContactIds(lead);
    }
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
 * 5. Вторичные поля дат (OR) без длинного !@ID:
 * - Покрытие: (основное поле M в [start, end)) ∪ (S1 в периоде) ∪ … =
 *   M ∪ (S1\\M) ∪ (S2\\M) … — снимаем пересечение с основной пачкой двумя фильтрами
 *   (M<start, M>=end) вместо полного обхода «S в периоде».
 *
 * @param {Object} period - Объект {start: "ISO", end: "ISO"}.
 * @param {string[]} headers - Массив уникальных заголовков таблицы.
 * @param {Object} liveMap - Живая карта полей из API.
 * @return {Object[]} Массив уникальных объектов лидов.
 */
function _collectLeadIdsByCursor(filter, progressCb, progressPrefix) {
  const emit = typeof progressCb === 'function' ? progressCb : null;
  const ids = [];
  const seen = new Set();
  const pagesPerBatch = 20; // 20 страниц * 50 = до 1000 ID за один batch-вызов
  let lastId = 0;
  let pageNo = 0;   // сквозной счетчик страниц
  let batchNo = 0;

  while (true) {
    const cmd = {};
    for (let i = 0; i < pagesPerBatch; i++) {
      const key = `p${i}`;
      if (i === 0) {
        const reqFilter = Object.assign({}, filter);
        if (lastId > 0) reqFilter[">ID"] = lastId;
        cmd[key] = BuildBatchCmd("crm.lead.list", {
          filter: reqFilter,
          order: { ID: "ASC" },
          select: ["ID"],
          start: 0
        });
      } else {
        cmd[key] = BuildBatchCmd("crm.lead.list", {
          filter: Object.assign({}, filter, { ">ID": `$result[p${i - 1}][49][ID]` }),
          order: { ID: "ASC" },
          select: ["ID"],
          start: 0
        });
      }
    }

    batchNo++;
    const batchRes = CallBitrix("batch", { halt: 1, cmd: cmd });
    const result = batchRes && batchRes.result ? batchRes.result : {};
    const resultMap = result.result || {};
    const errors = result.result_error || {};

    if (Object.keys(errors).length > 0) {
      console.warn(`⚠️ Ошибки batch при сборе ID: ${JSON.stringify(errors)}`);
      break;
    }

    let anyRows = false;
    let shouldStop = false;
    let lastTailIdInBatch = lastId;
    let addedInBatch = 0;
    let rowsInBatch = 0;
    let pagesProcessedInBatch = 0;

    for (let i = 0; i < pagesPerBatch; i++) {
      const rows = Array.isArray(resultMap[`p${i}`]) ? resultMap[`p${i}`] : [];
      if (rows.length > 0) anyRows = true;
      pageNo++;
      pagesProcessedInBatch++;
      rowsInBatch += rows.length;

      rows.forEach(row => {
        const id = row && row.ID != null ? String(row.ID) : "";
        if (!id || seen.has(id)) return;
        seen.add(id);
        ids.push(id);
        addedInBatch++;
      });

      if (rows.length > 0) {
        const tailRaw = rows[rows.length - 1] && rows[rows.length - 1].ID;
        const tailId = Number(tailRaw);
        if (isFinite(tailId) && tailId > lastTailIdInBatch) {
          lastTailIdInBatch = tailId;
        }
      }

      if (rows.length < 50) {
        shouldStop = true;
        break;
      }
    }

    if (emit) {
      emit(`${progressPrefix} — batch ${batchNo}: страниц ${pagesProcessedInBatch}, получено ${rowsInBatch}, новых ${addedInBatch}, всего ${ids.length}`);
    }

    if (!anyRows) break;
    if (lastTailIdInBatch <= lastId) {
      console.warn(`⚠️ Останов курсора: lastId=${lastId}, lastTailIdInBatch=${lastTailIdInBatch}`);
      break;
    }
    lastId = lastTailIdInBatch;
    if (shouldStop) break;
  }

  return { ids: ids, total: ids.length, pages: pageNo };
}

/**
 * ID по полю secondField в [period.start, period.end), у которых mainField **вне** того же
 * полуинтервала. Эквивалентно (S в периоде) \ (M в периоде) при OR-логике, если M — основной фильтр.
 * Два компактных фильтра (M &lt; start / M &gt;= end) вместо полного обхода пересечения.
 *
 * @param {string} secondaryFieldId
 * @param {string} mainFieldId
 * @param {{start:string,end:string}} period
 * @param {Function|null} progressCb
 * @param {string} progressPrefix
 * @returns {{ids: string[], total: number, pages: number, scanned: number}}
 */
function _collectLeadIdsSecondaryExcludingMainWindow(secondaryFieldId, mainFieldId, period, progressCb, progressPrefix) {
  if (secondaryFieldId === mainFieldId) {
    return _collectLeadIdsByCursor(
      { [`>=${secondaryFieldId}`]: period.start, [`<${secondaryFieldId}`]: period.end },
      progressCb,
      progressPrefix
    );
  }
  const fBase = {
    [`>=${secondaryFieldId}`]: period.start,
    [`<${secondaryFieldId}`]: period.end
  };
  const rBefore = _collectLeadIdsByCursor(
    Object.assign({}, fBase, { [`<${mainFieldId}`]: period.start }),
    progressCb,
    progressPrefix
  );
  const rAfter = _collectLeadIdsByCursor(
    Object.assign({}, fBase, { [`>=${mainFieldId}`]: period.end }),
    null,
    ""
  );
  const seen = new Set();
  const out = [];
  rBefore.ids.forEach(function (id) {
    if (seen.has(id)) return;
    seen.add(id);
    out.push(id);
  });
  rAfter.ids.forEach(function (id) {
    if (seen.has(id)) return;
    seen.add(id);
    out.push(id);
  });
  return {
    ids: out,
    total: out.length,
    pages: (rBefore.pages || 0) + (rAfter.pages || 0),
    scanned: rBefore.total + rAfter.total
  };
}

function _fetchLeadsByIdsInBatch(leadIds, selectFields, progressCb, progressPrefix) {
  const emit = typeof progressCb === 'function' ? progressCb : null;
  const rowsOut = [];
  const seen = new Set();
  let i = 0;
  let batchNo = 0;

  while (i < leadIds.length) {
    const cmd = {};
    let cmdCount = 0;

    while (cmdCount < 50 && i < leadIds.length) {
      const chunk = leadIds.slice(i, i + 50);
      cmd[`lead_by_id_${cmdCount}`] = BuildBatchCmd('crm.lead.list', {
        filter: { "@ID": chunk },
        select: selectFields
      });
      i += 50;
      cmdCount++;
    }

    batchNo++;
    const batchRes = CallBitrix('batch', { halt: 0, cmd: cmd });
    const resultMap = (batchRes && batchRes.result && batchRes.result.result) ? batchRes.result.result : {};

    Object.keys(resultMap).forEach(key => {
      const rows = resultMap[key];
      if (!Array.isArray(rows) || rows.length === 0) return;
      rows.forEach(lead => {
        const leadId = lead && lead.ID != null ? String(lead.ID) : "";
        if (!leadId || seen.has(leadId)) return;
        seen.add(leadId);
        rowsOut.push(lead);
      });
    });

    if (emit) {
      emit(`${progressPrefix} — батч ${batchNo} (~${cmdCount * 50}): ${Math.min(i, leadIds.length)}/${leadIds.length}`);
    }
  }

  return rowsOut;
}

function GetLeadsByFiltersMap(period, headers, liveMap, filterableHeaders, progressCb) {
  const emit = typeof progressCb === 'function' ? progressCb : null;
  const collectedIds = new Set();

  // Добавляем обязательные поля для аналитики
  const essentialIds = ["ID", "CONTACT_ID", "CONTACT_IDS", "DATE_CREATE"];
  const selectFields = [...new Set([...headers.map(h => liveMap[h]?.id).filter(id => id), ...essentialIds])];
  if (!filterableHeaders || filterableHeaders.length === 0) return [];

  // 1. ПЕРВЫЙ ЭТАП: собираем только ID через стабильный курсор >ID
  const mainHeader = filterableHeaders.find(h => liveMap[h].id === 'DATE_CREATE') || filterableHeaders[0];
  const mainFieldId = liveMap[mainHeader].id;

  const mainResult = _collectLeadIdsByCursor(
    { [`>=${mainFieldId}`]: period.start, [`<${mainFieldId}`]: period.end },
    emit,
    `🛰 Лиды: основной фильтр "${mainHeader}"`
  );
  mainResult.ids.forEach(id => collectedIds.add(id));
  console.log(`✅ Основной сбор ID завершен: ${mainResult.total} уникальных ID.`);

  // 2. ВТОРОЙ ЭТАП: остальные фильтры (OR) — только «новые» к основной выборке:
  // вторичное поле в периоде, но основное (main) поле вне [start, end) — без длинного !@ID.
  const secondaryHeaders = filterableHeaders.filter(h => h !== mainHeader);
  secondaryHeaders.forEach(headerName => {
    const fieldId = liveMap[headerName].id;
    const before = collectedIds.size;
    const secResult = _collectLeadIdsSecondaryExcludingMainWindow(
      fieldId,
      mainFieldId,
      period,
      null,
      ""
    );
    secResult.ids.forEach(id => collectedIds.add(id));
    const added = collectedIds.size - before;
    if (emit) {
      emit(
        `🔎 Лиды: фильтр "${headerName}" (вне «${mainHeader}») — уник. ${secResult.total}, новых: ${added}, всего: ${collectedIds.size}`
      );
    }
  });

  // 3. ТРЕТИЙ ЭТАП: Дозагрузка полных данных только по собранным ID
  const allIds = Array.from(collectedIds);
  if (allIds.length === 0) return [];
  if (emit) emit(`📡 Лиды: дозагрузка карточек по ID (${allIds.length})...`);
  const allLeads = _fetchLeadsByIdsInBatch(allIds, selectFields, emit, `📡 Лиды: загрузка по @ID`);

  allLeads.forEach(fillVirtualAllContactIds);
  if (emit) emit(`✅ Лиды: собрано ${allLeads.length} уникальных записей.`);

  return allLeads;
}