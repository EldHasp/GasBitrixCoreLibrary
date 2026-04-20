




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





