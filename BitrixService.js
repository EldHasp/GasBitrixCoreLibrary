/** Вспомогательный метод для единообразного форматирования имени */
function FormatUserName(u) {
    let name = `${u.NAME || ""} ${u.LAST_NAME || ""}`.trim();
    return name || u.EMAIL || u.LOGIN || `ID ${u.ID}`;
}

// CallBitrix, CallBitrix3, SendRequest, BuildBatchCmd — см. BitrixCommandsService.js

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
 * @property {string} [CRM_ENTITY_TYPE] - Как в API: "LEAD" | "CONTACT" | "".
 * @property {string} [CRM_ENTITY_ID] - ID сущности в CRM (строка), как в API.
 * @property {string} CRM - Только для листа: человекочитаемая строка «LEAD: …» / «CONTACT: …».
 */

/**
 * ОРКЕСТРАТОР СБОРА ЗВОНКОВ
 * Собирает данные, мапит сотрудников и возвращает массив подготовленных объектов.
 * 
 * @param {Object} period - Объект {start, end} в формате ISO.
 * @returns {PreparedCall[]} Массив подготовленных объектов звонков для записи.
 */
function ExportCallsToSheet(config, period, statusCb) {
  const emit = typeof statusCb === 'function' ? statusCb : updateStatus;
  emit("📞 Телефония: Начинаем сбор данных...");
  
  // 1. Сбор сырых данных
  const callOpts = {};
  if (config && config.callsBatchWindowHours != null) {
    const co = typeof config.callsBatchWindowHours === "number" && isFinite(config.callsBatchWindowHours)
      ? config.callsBatchWindowHours
      : _parseCallsBatchWindowHoursString_(String(config.callsBatchWindowHours));
    if (co != null) callOpts.windowHours = co;
  }
  callOpts.onProgress = function(summaryMsg) {
    if (typeof updateStatusSummary === "function") {
      updateStatusSummary(summaryMsg);
    } else if (emit) {
      emit(summaryMsg);
    }
  };
  const rawCalls = GetCallsData(period, emit, callOpts);
  emit(`📞 Телефония: Получено ${rawCalls.length} записей из API`);

  if (rawCalls.length === 0) return [];

  // 2. Подготовка сотрудников
  emit("📞 Телефония: Маппинг менеджеров...");
  const userIds = rawCalls.map(c => c.PORTAL_USER_ID).filter(id => id);
  const usersMap = GetUsersMap(userIds);

  // 3. Формирование объектов
  emit("📞 Телефония: Формирование отчета...");
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
      CRM_ENTITY_TYPE: c.CRM_ENTITY_TYPE ? String(c.CRM_ENTITY_TYPE).toUpperCase() : "",
      CRM_ENTITY_ID: c.CRM_ENTITY_ID != null && c.CRM_ENTITY_ID !== "" ? String(c.CRM_ENTITY_ID) : "",
      CRM: c.CRM_ENTITY_TYPE ? `${c.CRM_ENTITY_TYPE}: ${c.CRM_ENTITY_ID}` : ""
    };
  });
}




/**
 * Форматирует дату в ISO с фиксированным смещением +03:00 (МСК).
 * Используется для стабильного оконного сбора статистики звонков.
 *
 * @param {Date} dateObj
 * @returns {string}
 * @private
 */
function _formatIsoMsk(dateObj) {
  const shifted = new Date(dateObj.getTime() + 3 * 60 * 60 * 1000);
  const year = shifted.getUTCFullYear();
  const month = ("0" + (shifted.getUTCMonth() + 1)).slice(-2);
  const day = ("0" + shifted.getUTCDate()).slice(-2);
  const hours = ("0" + shifted.getUTCHours()).slice(-2);
  const minutes = ("0" + shifted.getUTCMinutes()).slice(-2);
  const seconds = ("0" + shifted.getUTCSeconds()).slice(-2);
  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}+03:00`;
}

/**
 * Старт текущих суток по МСК в миллисекундах UTC-эпохи.
 * @returns {number}
 * @private
 */
function _getTodayStartMsMsk_() {
  const now = new Date();
  const mskShiftMs = 3 * 60 * 60 * 1000;
  const mskNow = new Date(now.getTime() + mskShiftMs);
  const y = mskNow.getUTCFullYear();
  const m = mskNow.getUTCMonth();
  const d = mskNow.getUTCDate();
  const todayStartMskUtc = Date.UTC(y, m, d, 0, 0, 0) - mskShiftMs;
  return todayStartMskUtc;
}

/**
 * Короткий запрос первого/последнего звонка в диапазоне.
 * Возвращает первую строку (по CALL_START_DATE ASC/DESC) или null.
 *
 * @param {string} fromIso
 * @param {string} toIso
 * @param {"ASC"|"DESC"} order
 * @returns {Object|null}
 * @private
 */
function _fetchBoundaryCall_(fromIso, toIso, order) {
  const res = CallBitrix("voximplant.statistic.get", {
    FILTER: {
      ">=CALL_START_DATE": fromIso,
      "<CALL_START_DATE": toIso
    },
    SORT: "CALL_START_DATE",
    ORDER: order,
    start: 0
  });
  if (!res || !res.result || !Array.isArray(res.result) || res.result.length === 0) {
    return null;
  }
  return res.result[0] || null;
}

/**
 * Грубая оценка объема звонков по разнице ID:
 * - первый звонок: от периода start (ASC),
 * - последний: до периода end (DESC),
 * - если end попадает в "сегодня/будущее" и звонков за сегодня нет, ищем вчера (DESC).
 *
 * @param {{start: string, end: string}} period
 * @returns {{estimated: number|null, firstId: number|null, lastId: number|null}}
 * @private
 */
function _estimateCallsByIdRange_(period) {
  let first = null;
  let last = null;
  try {
    first = _fetchBoundaryCall_(period.start, period.end, "ASC");
    if (!first || first.ID == null) {
      return { estimated: 0, firstId: null, lastId: null };
    }

    const endMs = new Date(period.end).getTime();
    const todayStartMs = _getTodayStartMsMsk_();

    // Базово берём последний звонок по всему диапазону.
    last = _fetchBoundaryCall_(period.start, period.end, "DESC");

    // Если конец периода сегодня/позже и "хвост" пуст, пробуем вчера.
    if ((!last || last.ID == null) && endMs >= todayStartMs) {
      const yStartMs = todayStartMs - 24 * 60 * 60 * 1000;
      const yStartIso = _formatIsoMsk(new Date(yStartMs));
      const yEndIso = _formatIsoMsk(new Date(todayStartMs));
      last = _fetchBoundaryCall_(yStartIso, yEndIso, "DESC");
    }

    const firstId = Number(first.ID);
    const lastId = last && last.ID != null ? Number(last.ID) : firstId;
    if (isNaN(firstId) || isNaN(lastId)) {
      return { estimated: null, firstId: null, lastId: null };
    }
    // В voximplant ID звонков в этом портале идут чётным шагом (2),
    // поэтому оценка по диапазону должна учитывать шаг.
    const step = (firstId % 2 === 0 && lastId % 2 === 0) ? 2 : 1;
    const estimated = Math.max(0, Math.floor((lastId - firstId) / step) + 1);
    return { estimated: estimated, firstId: firstId, lastId: lastId };
  } catch (e) {
    console.warn(`⚠ Не удалось оценить объём звонков по ID: ${e.message}`);
    return { estimated: null, firstId: null, lastId: null };
  }
}


/**
 * @param {string|number} raw
 * @returns {number|null} положительное целое или null
 * @private
 */
function _parseCallsBatchWindowHoursString_(raw) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim();
  if (s === "") return null;
  const n = parseInt(s, 10);
  if (isNaN(n) || n <= 0) return null;
  return n;
}

/**
 * Размер временного окна для batch-загрузки звонков.
 *
 * Приоритет: `optionalOverride` (из `config.callsBatchWindowHours` — обычно ячейка
 * `РазмерОкнаВыгрузкиЗвонков` в «Справочнике») → **Свойства скрипта** GAS (редко) →
 * **Свойства документа** таблицы → 24. Основной путь: именованный диапазон в таблице.
 *
 * @param {string|number|undefined|null} [optionalOverride]
 * @returns {{hours: number, ms: number}}
 * @private
 */
function _resolveCallsWindowSize(optionalOverride) {
  const o = (typeof optionalOverride === "number" && optionalOverride > 0 && isFinite(optionalOverride))
    ? Math.floor(optionalOverride)
    : (optionalOverride != null && optionalOverride !== ""
      ? _parseCallsBatchWindowHoursString_(String(optionalOverride))
      : null);
  if (o != null) {
    return { hours: o, ms: o * 60 * 60 * 1000 };
  }

  let fromScript = null;
  try {
    fromScript = _parseCallsBatchWindowHoursString_(
      PropertiesService.getScriptProperties().getProperty("CALLS_BATCH_WINDOW_HOURS")
    );
  } catch (e) {
    console.warn(`⚠ Скрипт: CALLS_BATCH_WINDOW_HOURS: ${e.message}`);
  }
  if (fromScript != null) {
    return { hours: fromScript, ms: fromScript * 60 * 60 * 1000 };
  }

  let fromDocument = null;
  try {
    fromDocument = _parseCallsBatchWindowHoursString_(
      PropertiesService.getDocumentProperties().getProperty("CALLS_BATCH_WINDOW_HOURS")
    );
  } catch (e) {
    // Нет привязанного документа (редко при вызове не из таблицы)
  }
  if (fromDocument != null) {
    return { hours: fromDocument, ms: fromDocument * 60 * 60 * 1000 };
  }

  return { hours: 24, ms: 24 * 60 * 60 * 1000 };
}

/**
 * Загружает звонки для одного временного окна через batch:
 * 50 команд × 50 записей = до 2500 звонков за один batch-вызов.
 *
 * @param {string} windowStartIso
 * @param {string} windowEndIso
 * @param {Function|null} emit
 * @param {number} windowNo
 * @param {number} windowsTotal
 * @returns {{rows: Object[], total: number, batches: number}}
 * @private
 */
function _fetchCallsWindowBatch(windowStartIso, windowEndIso, emit, windowNo, windowsTotal) {
  const collected = [];
  let start = 0;
  let total = 1;
  let batchNo = 0;

  while (start < total) {
    const cmd = {};
    for (let i = 0; i < 50; i++) {
      const offset = start + i * 50;
      if (total !== 1 && offset >= total) break;
      cmd[`calls_${i}`] = BuildBatchCmd('voximplant.statistic.get', {
        FILTER: {
          ">=CALL_START_DATE": windowStartIso,
          "<CALL_START_DATE": windowEndIso
        },
        SORT: "CALL_START_DATE",
        ORDER: "ASC",
        start: offset
      });
    }
    const cmdCount = Object.keys(cmd).length;
    if (cmdCount === 0) break;

    const batchRes = CallBitrix('batch', { halt: 0, cmd: cmd });
    const resultMap = (batchRes && batchRes.result && batchRes.result.result) ? batchRes.result.result : {};
    const totalMap = (batchRes && batchRes.result && batchRes.result.result_total) ? batchRes.result.result_total : {};

    Object.keys(resultMap).forEach(key => {
      const rows = resultMap[key];
      if (Array.isArray(rows) && rows.length > 0) collected.push(...rows);
    });

    const anyTotal = Object.values(totalMap)[0];
    if (anyTotal !== undefined && anyTotal !== null && anyTotal !== "") {
      const parsedTotal = Number(anyTotal);
      if (!isNaN(parsedTotal) && parsedTotal >= 0) total = parsedTotal;
    } else if (total === 1 && collected.length === 0) {
      total = 0;
    } else if (total === 1) {
      total = start + cmdCount * 50 + 1;
    }

    start += cmdCount * 50;
    batchNo++;

    if (cmdCount < 50 && Object.keys(resultMap).length < 50) break;
  }

  return { rows: collected, total: total, batches: batchNo };
}

/**
 * Оконная загрузка статистики звонков из Bitrix24 (voximplant.statistic.get).
 * Для скорости использует batch-пакеты до 2500 записей за итерацию,
 * для устойчивости — дедупликацию по ID и сбор по непересекающимся временным окнам.
 * 
 * @param {Object} period - Объект с границами временного периода.
 * @param {string} period.start - Дата начала в формате ISO 8601 (например, "2024-01-01T00:00:00+03:00").
 * @param {string} period.end - Дата конца в формате ISO 8601 (например, "2024-01-31T23:59:59+03:00").
 * 
 * @returns {Object[]} Массив "сырых" объектов звонков из API Bitrix24.
 * 
 * @param {{ windowHours?: number }|null} [options] — например, `{ windowHours: 72 }` из `config`.
 * @example
 * const period = { start: "2024-10-01T00:00:00+03:00", end: "2024-10-01T23:59:59+03:00" };
 * const calls = GetCallsData(period);
 */
function GetCallsData(period, statusCb, options) {
  const emit = typeof statusCb === "function" ? statusCb : null;
  const emitProgress =
    options && typeof options.onProgress === "function" ? options.onProgress : null;
  const allCalls = [];
  const seenCallIds = new Set();
  const winArg =
    options && (options.windowHours != null)
      ? options.windowHours
      : undefined;
  const windowSize = _resolveCallsWindowSize(winArg);
  const WINDOW_MS = windowSize.ms;

  console.log(`📞 СТАРТ СБОРА ЗВОНКОВ: с [${period.start}] по [${period.end}]`);
  console.log(`📞 Размер окна batch: ${windowSize.hours} ч.`);
  const startMs = new Date(period.start).getTime();
  const endMs = new Date(period.end).getTime();
  if (isNaN(startMs) || isNaN(endMs) || startMs >= endMs) {
    console.warn("⚠ Некорректный период для звонков, возвращаю пустой массив.");
    return [];
  }

  const windowsTotal = Math.max(1, Math.ceil((endMs - startMs) / WINDOW_MS));
  let cursor = startMs;
  let windowNo = 0;
  let discoveredTotal = 0;
  const estimate = _estimateCallsByIdRange_(period);
  const estimatedAdjusted = estimate && estimate.estimated != null ? Number(estimate.estimated) : null;
  if (emitProgress) {
    if (estimatedAdjusted != null) {
      emitProgress(`📞 Телефония: оценка ~${estimatedAdjusted} звонков (по диапазону ID)`);
    } else {
      emitProgress(`📞 Телефония: оценка объёма недоступна, начинаю загрузку...`);
    }
  }

  while (cursor < endMs) {
    windowNo++;
    const windowEnd = Math.min(cursor + WINDOW_MS, endMs);
    const windowStartIso = _formatIsoMsk(new Date(cursor));
    const windowEndIso = _formatIsoMsk(new Date(windowEnd));

    const result = _fetchCallsWindowBatch(windowStartIso, windowEndIso, emit, windowNo, windowsTotal);
    const windowTotal = Math.max(result.total || 0, result.rows.length || 0);
    discoveredTotal += windowTotal;
    result.rows.forEach(call => {
      const callId = call && call.ID != null ? String(call.ID) : "";
      if (!callId) return;
      if (!seenCallIds.has(callId)) {
        seenCallIds.add(callId);
        allCalls.push(call);
      }
    });

    if (emitProgress) {
      const estimatedTotal = estimatedAdjusted;
      const liveTotal = discoveredTotal > 0 ? discoveredTotal : allCalls.length;
      const totalLabel = estimatedTotal != null
        ? `~${Math.max(estimatedTotal, liveTotal)}`
        : String(liveTotal);
      let percentLabel = "";
      if (estimatedTotal != null && estimatedTotal > 0) {
        const denominator = Math.max(estimatedTotal, liveTotal);
        const pct = Math.max(0, Math.min(100, (allCalls.length / denominator) * 100));
        percentLabel = ` (${pct.toFixed(1)}%)`;
      }
      emitProgress(`📞 Телефония: загружено ${allCalls.length}/${totalLabel}${percentLabel}, окно ${windowNo}/${windowsTotal} (${windowSize.hours}ч)`);
    }

    cursor = windowEnd;
  }

  allCalls.sort((a, b) => {
    const ta = new Date(a.CALL_START_DATE).getTime();
    const tb = new Date(b.CALL_START_DATE).getTime();
    if (ta !== tb) return ta - tb;
    return String(a.ID).localeCompare(String(b.ID), 'en', { numeric: true });
  });

  console.log(`✅ Сбор завершен. Получено звонков: ${allCalls.length}`);
  return allCalls;
}

