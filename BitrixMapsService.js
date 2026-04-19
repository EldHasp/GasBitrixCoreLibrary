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
 * Получает маппинг полей напрямую из API.
 * @returns {Object<string, {id: string, type: string}>}
 */
function GetLiveFieldsMap() {
    const response = CallBitrix('crm.lead.fields', {});
    const allFields = response.result;
    if (!allFields) throw new Error("❌ Ошибка API: структура полей не получена.");

    return Object.keys(allFields).reduce((liveMap, key) => {
        const f = allFields[key];
        // Логика именования: приоритет на человеческое название
        const russianName = f.title || f.formLabel || f.listLabel || key;
        
        liveMap[russianName.toString()] = {
            id: key,
            type: f.type
        };
        return liveMap;
    }, {});
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
