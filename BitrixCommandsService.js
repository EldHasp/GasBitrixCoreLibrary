/**
 * 1. УНИВЕРСАЛЬНЫЙ ВЫЗОВ (REST 2.0)
 * @param {string} method - Имя метода (например, 'crm.lead.list')
 * @param {Object} [params={}] - Параметры запроса
 * @returns {{result: any, next: number|null, total: number, error: string|null, error_description: string|null}}
 */
function CallBitrix(method, params = {}) {
    if (!BITRIX_URL) return { error: 'CONFIG_ERROR', error_description: 'BITRIX_URL не инициализирован' };
    
    const url = BITRIX_URL.endsWith('/') ? BITRIX_URL + method : BITRIX_URL + '/' + method;
    return SendRequest(url, params);
}

/**
 * 2. ВЫЗОВ REST 3.0
 * @param {string} method - Имя метода
 * @param {Object} [params={}] - Параметры запроса
 * @returns {{result: any, next: string|null, error: string|null}}
 */
function CallBitrix3(method, params = {}) {
    if (!BITRIX_URL_REST_3_0) return { error: 'CONFIG_ERROR', error_description: 'BITRIX_URL_REST_3_0 не инициализирован' };
    
    const url = BITRIX_URL_REST_3_0.endsWith('/') ? BITRIX_URL_REST_3_0 + method : BITRIX_URL_REST_3_0 + '/' + method;
    return SendRequest(url, params);
}

/**
 * 3. ВНУТРЕННЯЯ ФУНКЦИЯ ОТПРАВКИ
 * @param {string} url - Полный URL эндпоинта
 * @param {Object} params - Объект с данными
 * @returns {Object} Распарсенный JSON ответа или объект ошибки
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
        const content = response.getContentText();
        const json = JSON.parse(content);

        if (json.error) {
            console.error(`❌ Bitrix API Error: ${json.error} - ${json.error_description}`);
        }

        return json;
    } catch (e) {
        console.error(`⚠️ Network Error [${url}]: ${e.message}`);
        return { error: 'FETCH_ERROR', error_description: e.message };
    }
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
