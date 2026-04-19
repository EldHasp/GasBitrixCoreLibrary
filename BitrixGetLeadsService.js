




/**
 * ПОЛУЧЕНИЕ КОЛЛЕКЦИИ ОФИСОВ (Инфоблок 128)
 * Возвращает объект { "109734": "Офис Москва", ... }
 * Проблемы и пути решения в комментарии метода GetOfficesReference
 *//**
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
 */
/**
 * ПОЛУЧЕНИЕ КОЛЛЕКЦИИ ИСТОЧНИКОВ
 * @return {Object} - { "WEB": "Сайт", "PHONE": "Звонок" }
 *//**
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
 */
/**
 * Получает маппинг полей напрямую из API Битрикс24 (без использования листа).
 * @return {Object} - { "Название RU": {id: "TECH_ID", type: "type"}, ... }
 *//**
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
*/

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



