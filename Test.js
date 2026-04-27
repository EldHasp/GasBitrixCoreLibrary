/**
 * ТЕСТ: Запись уведомления всеми способами для проверки видимости
 */
function TestAllStorageMethods(ss) {
  const testMsg = "Проверка связи от " + new Date().toLocaleTimeString();
  
  // 1. Свойства Пользователя (User)
  PropertiesService.getUserProperties().setProperty('TEST_USER_PROP', "USER: " + testMsg);
  
  // 2. Свойства Документа (Document)
  PropertiesService.getDocumentProperties().setProperty('TEST_DOC_PROP', "DOC: " + testMsg);
  
  // 3. Метаданные (Developer Metadata)
  // Предварительно удалим старые, если были
  ss.getDeveloperMetadata().filter(m => m.getKey() === 'TEST_META_PROP').forEach(m => m.remove());
  ss.addDeveloperMetadata('TEST_META_PROP', "META: " + testMsg, SpreadsheetApp.DeveloperMetadataVisibility.PROJECT);
  
  console.log("✅ Библиотека: Записала данные всеми тремя способами.");
}

function simulateLibraryMessage() {
  SpreadsheetApp.getActiveSpreadsheet().addDeveloperMetadata(
    'PENDING_UPDATE_MSG', 
    'ТЕСТ: Найдено 5 новых источников!', 
    SpreadsheetApp.DeveloperMetadataVisibility.PROJECT
  );
}

/**
 * ТЕСТ: Выгружает последние 50 звонков в файл JSON на Google Drive.
 */
function debugSaveCallsToFile() {
  const response = CallBitrix('voximplant.statistic.get', {
    SORT: "CALL_START_DATE",
    ORDER: "DESC"
  });

  const rawData = response.result || [];
  
  if (rawData.length > 0) {
    const fileName = `Bitrix_Calls_Debug_${new Date().getTime()}.json`;
    // Создаем файл на Google Drive
    const file = DriveApp.createFile(fileName, JSON.stringify(rawData, null, 2), MimeType.PLAIN_TEXT);
    
    console.log(`✅ Файл создан: ${file.getName()}`);
    console.log(`🔗 Ссылка: ${file.getUrl()}`);
    
    return file.getUrl();
  } else {
    console.warn("Звонки не найдены.");
    return null;
  }
}

/**
 * ТЕСТ: Проверка фильтрации по прямому примеру из документации.
 */
function debugTestExactDocFilter() {
  // Пробуем максимально "чистый" фильтр
  const testParams = {
    FILTER: {
       ">=CALL_START_DATE": "2026-04-01T00:00:00+03:00"
    },
    SORT: "ID",
    ORDER: "ASC"
  };

  console.log("📡 Отправка тестового запроса...");
  const response = CallBitrix('voximplant.statistic.get', testParams);
  
  const result = response.result || [];
  
  if (result.length > 0) {
    console.log(`✅ Получено звонков: ${result.length}`);
    console.log(`📅 Дата ПЕРВОГО звонка в ответе: ${result[0].CALL_START_DATE}`);
    console.log(`🆔 ID ПЕРВОГО звонка: ${result[0].ID}`);
    
    if (result[0].CALL_START_DATE.includes("2026-04-01")) {
      console.log("🎯 ФОРМАТ ИЗ ДОКУМЕНТАЦИИ СРАБОТАЛ!");
    } else {
      console.log("❌ ФИЛЬТР ПРОИГНОРИРОВАН (снова вернул старые данные)");
    }
  } else {
    console.warn("📭 Звонков не найдено.");
  }
}


/**
 * ТЕСТ: Диагностика дублей лидов на листе по колонке ID.
 *
 * @param {string} [sheetName='Лиды'] - Имя листа с выгрузкой лидов.
 * @param {string} [idHeader='ID'] - Заголовок колонки с ID лида.
 */
function debugLeadDuplicates(sheetName = 'Лиды', idHeader = 'ID') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Лист "${sheetName}" не найден`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    console.log('Нет данных для анализа дублей.');
    return;
  }

  const headers = values[0].map(h => String(h || '').trim());
  const idCol = headers.indexOf(idHeader);
  if (idCol === -1) throw new Error(`Колонка "${idHeader}" не найдена`);

  const counts = {};
  let emptyId = 0;

  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][idCol] || '').trim();
    if (!id) {
      emptyId++;
      continue;
    }
    counts[id] = (counts[id] || 0) + 1;
  }

  const entries = Object.entries(counts);
  const duplicates = entries
    .filter(([, count]) => count > 1)
    .sort((a, b) => b[1] - a[1]);

  console.log(`Лист: ${sheetName}`);
  console.log(`Всего строк данных: ${values.length - 1}`);
  console.log(`Уникальных ID: ${entries.length}`);
  console.log(`Пустых ID: ${emptyId}`);
  console.log(`ID с дублями: ${duplicates.length}`);
  console.log(`Лишних строк из-за дублей: ${duplicates.reduce((acc, [, c]) => acc + (c - 1), 0)}`);

  const top = duplicates.slice(0, 30);
  if (top.length === 0) {
    console.log('Дубли не найдены.');
    return;
  }

  console.log('TOP дублей (ID => count):');
  top.forEach(([id, count]) => console.log(`${id} => ${count}`));
}

/**
 * ТЕСТ: Проверка зависимых команд внутри batch через $result[...].
 *
 * Цепочка:
 * 1) lead_first_page  - берём первые 2 лида по ID ASC
 * 2) lead_from_first  - получаем карточку первого лида из шага 1
 * 3) lead_after_first - берём следующий лид с ID > первого из шага 1
 *
 * Метод нужен для быстрой проверки, что подстановка результатов
 * предыдущих команд внутри одного batch работает корректно.
 */
function debugBatchLeadDependencies() {
  const cmd = {
    lead_first_page: BuildBatchCmd('crm.lead.list', {
      order: { ID: "ASC" },
      select: ["ID", "TITLE"],
      start: 0
    }),
    lead_from_first: BuildBatchCmd('crm.lead.get', {
      id: "$result[lead_first_page][0][ID]"
    }),
    lead_after_first: BuildBatchCmd('crm.lead.list', {
      order: { ID: "ASC" },
      filter: { ">ID": "$result[lead_first_page][0][ID]" },
      select: ["ID", "TITLE"],
      start: 0
    })
  };

  const batchRes = CallBitrix('batch', { halt: 1, cmd: cmd });
  const result = batchRes && batchRes.result ? batchRes.result : {};
  const resultMap = result.result || {};
  const errors = result.result_error || {};

  const firstPage = Array.isArray(resultMap.lead_first_page) ? resultMap.lead_first_page : [];
  const firstLead = resultMap.lead_from_first || null;
  const nextPage = Array.isArray(resultMap.lead_after_first) ? resultMap.lead_after_first : [];

  console.log("🧪 Batch dependency test started");
  console.log(`Ошибки: ${Object.keys(errors).length}`);
  if (Object.keys(errors).length > 0) {
    console.log(`result_error: ${JSON.stringify(errors)}`);
  }

  console.log(`lead_first_page: ${firstPage.length} строк`);
  if (firstPage[0]) {
    console.log(`first ID from lead_first_page: ${firstPage[0].ID}`);
  }

  if (firstLead && firstLead.ID) {
    console.log(`lead_from_first ID: ${firstLead.ID}`);
  } else {
    console.log("lead_from_first: пусто");
  }

  console.log(`lead_after_first: ${nextPage.length} строк`);
  if (nextPage[0]) {
    console.log(`first ID from lead_after_first: ${nextPage[0].ID}`);
  }

  return {
    errors: errors,
    firstPageCount: firstPage.length,
    firstLeadId: firstPage[0] ? String(firstPage[0].ID) : null,
    leadFromFirstId: firstLead && firstLead.ID ? String(firstLead.ID) : null,
    nextPageCount: nextPage.length,
    nextFirstLeadId: nextPage[0] ? String(nextPage[0].ID) : null
  };
}

/**
 * ТЕСТ: Глубокая проверка зависимостей в batch-цепочке по лидам.
 *
 * Идея:
 * - В одном batch строим несколько страниц по 50 лидов.
 * - Каждая следующая страница зависит от ID последнего лида предыдущей:
 *   filter[>ID] = $result[prev_page][49][ID]
 * - Дополнительно проверяем "сцепку" через crm.lead.get по первому ID каждой страницы.
 *
 * @param {number} [pages=6] - Кол-во страниц в цепочке (1..20).
 * @returns {{ok:boolean, pages:number, failures:string[]}}
 */
function debugBatchLeadDependenciesDeep(pages = 6) {
  const pagesCount = Math.max(1, Math.min(20, parseInt(pages, 10) || 6));
  const cmd = {};

  // Базовая команда
  cmd["page_0"] = BuildBatchCmd("crm.lead.list", {
    order: { ID: "ASC" },
    select: ["ID", "TITLE"],
    start: 0
  });
  cmd["check_0"] = BuildBatchCmd("crm.lead.get", {
    id: "$result[page_0][0][ID]"
  });

  // Зависимые страницы
  for (let i = 1; i < pagesCount; i++) {
    cmd[`page_${i}`] = BuildBatchCmd("crm.lead.list", {
      order: { ID: "ASC" },
      filter: { ">ID": `$result[page_${i - 1}][49][ID]` },
      select: ["ID", "TITLE"],
      start: 0
    });
    cmd[`check_${i}`] = BuildBatchCmd("crm.lead.get", {
      id: `$result[page_${i}][0][ID]`
    });
  }

  const batchRes = CallBitrix("batch", { halt: 1, cmd: cmd });
  const result = batchRes && batchRes.result ? batchRes.result : {};
  const resultMap = result.result || {};
  const errors = result.result_error || {};
  const failures = [];

  const toId = (v) => {
    const n = Number(v);
    return isNaN(n) ? null : n;
  };

  console.log(`🧪 Deep batch dependency test started (pages=${pagesCount})`);
  console.log(`Ошибки batch: ${Object.keys(errors).length}`);
  if (Object.keys(errors).length > 0) {
    console.log(`result_error: ${JSON.stringify(errors)}`);
    failures.push("Есть ошибки в result_error");
  }

  let prevLastId = null;
  for (let i = 0; i < pagesCount; i++) {
    const page = Array.isArray(resultMap[`page_${i}`]) ? resultMap[`page_${i}`] : [];
    const check = resultMap[`check_${i}`] || null;

    const firstId = page[0] ? toId(page[0].ID) : null;
    const lastId = page.length > 0 ? toId(page[page.length - 1].ID) : null;
    const checkId = check && check.ID != null ? toId(check.ID) : null;

    console.log(`page_${i}: rows=${page.length}, first=${firstId}, last=${lastId}, check=${checkId}`);

    if (page.length === 0) {
      failures.push(`page_${i}: пустая страница`);
      break;
    }

    if (checkId === null || firstId === null || checkId !== firstId) {
      failures.push(`page_${i}: check_${i} не совпал с первым ID страницы`);
    }

    if (prevLastId !== null && firstId !== null && firstId <= prevLastId) {
      failures.push(`page_${i}: firstId (${firstId}) <= prevLastId (${prevLastId})`);
    }

    prevLastId = lastId;
  }

  if (failures.length === 0) {
    console.log("✅ Deep batch test passed: зависимости и курсор по ID работают корректно.");
  } else {
    console.log("❌ Deep batch test failed:");
    failures.forEach(f => console.log(` - ${f}`));
  }

  return {
    ok: failures.length === 0,
    pages: pagesCount,
    failures: failures
  };
}

/**
 * ТЕСТ: Сравнивает сбор ID лидов двумя способами:
 * 1) Последовательный курсор >ID (как в текущем коде)
 * 2) Batch-цепочка страниц по 50 через зависимости $result
 *
 * @param {string} startIso - Начало периода (включительно), например "2025-01-01T00:00:00+03:00"
 * @param {string} endIso - Конец периода (исключительно), например "2026-01-01T00:00:00+03:00"
 * @param {number} [batchPagesPerCall=10] - Кол-во страниц в одной batch-цепочке (1..20)
 * @returns {{ok:boolean, sequentialCount:number, batchCount:number, onlySequential:number, onlyBatch:number}}
 */
function debugCompareLeadIdCollection(startIso, endIso, batchPagesPerCall = 10) {
  if (!startIso || !endIso) {
    throw new Error("Нужно передать startIso и endIso");
  }

  const pagesPerCall = Math.max(1, Math.min(20, parseInt(batchPagesPerCall, 10) || 10));
  const baseFilter = {
    ">=DATE_CREATE": startIso,
    "<DATE_CREATE": endIso
  };

  const toId = (v) => {
    const n = Number(v);
    return isNaN(n) ? null : n;
  };

  const collectSequential = () => {
    const ids = [];
    const seen = new Set();
    let lastId = 0;
    let pageNo = 0;

    while (true) {
      const filter = Object.assign({}, baseFilter);
      if (lastId > 0) filter[">ID"] = lastId;

      const res = CallBitrix("crm.lead.list", {
        filter: filter,
        order: { ID: "ASC" },
        select: ["ID"],
        start: 0
      });
      const rows = res && Array.isArray(res.result) ? res.result : [];
      pageNo++;

      rows.forEach(row => {
        const id = row && row.ID != null ? String(row.ID) : "";
        if (!id || seen.has(id)) return;
        seen.add(id);
        ids.push(id);
      });

      console.log(`SEQ page ${pageNo}: rows=${rows.length}, total=${ids.length}`);
      if (rows.length === 0) break;

      const tailId = toId(rows[rows.length - 1] && rows[rows.length - 1].ID);
      if (!isFinite(tailId) || tailId <= lastId) break;
      lastId = tailId;
      if (rows.length < 50) break;
    }

    return ids;
  };

  const collectBatch = () => {
    const ids = [];
    const seen = new Set();
    let lastId = 0;
    let batchNo = 0;

    while (true) {
      const cmd = {};
      for (let i = 0; i < pagesPerCall; i++) {
        const key = `p${i}`;
        if (i === 0) {
          const f = Object.assign({}, baseFilter);
          if (lastId > 0) f[">ID"] = lastId;
          cmd[key] = BuildBatchCmd("crm.lead.list", {
            filter: f,
            order: { ID: "ASC" },
            select: ["ID"],
            start: 0
          });
        } else {
          cmd[key] = BuildBatchCmd("crm.lead.list", {
            filter: Object.assign({}, baseFilter, { ">ID": `$result[p${i - 1}][49][ID]` }),
            order: { ID: "ASC" },
            select: ["ID"],
            start: 0
          });
        }
      }

      batchNo++;
      const res = CallBitrix("batch", { halt: 1, cmd: cmd });
      const result = res && res.result ? res.result : {};
      const map = result.result || {};
      const errors = result.result_error || {};

      if (Object.keys(errors).length > 0) {
        console.warn(`BATCH ${batchNo}: errors=${JSON.stringify(errors)}`);
        break;
      }

      let anyRows = false;
      let shouldStop = false;
      let lastTailIdInBatch = lastId;

      for (let i = 0; i < pagesPerCall; i++) {
        const rows = Array.isArray(map[`p${i}`]) ? map[`p${i}`] : [];
        if (rows.length > 0) anyRows = true;

        rows.forEach(row => {
          const id = row && row.ID != null ? String(row.ID) : "";
          if (!id || seen.has(id)) return;
          seen.add(id);
          ids.push(id);
        });

        if (rows.length > 0) {
          const tailId = toId(rows[rows.length - 1] && rows[rows.length - 1].ID);
          if (isFinite(tailId) && tailId > lastTailIdInBatch) {
            lastTailIdInBatch = tailId;
          }
        }

        // Любая неполная страница значит хвост выборки достигнут.
        if (rows.length < 50) {
          shouldStop = true;
          break;
        }
      }

      console.log(`BATCH call ${batchNo}: total=${ids.length}, lastTail=${lastTailIdInBatch}`);

      if (!anyRows) break;
      if (lastTailIdInBatch <= lastId) break;
      lastId = lastTailIdInBatch;
      if (shouldStop) break;
    }

    return ids;
  };

  console.log(`🧪 Compare started: ${startIso} -> ${endIso}, pagesPerBatch=${pagesPerCall}`);
  const seqIds = collectSequential();
  const batchIds = collectBatch();

  const seqSet = new Set(seqIds);
  const batchSet = new Set(batchIds);

  const onlySeq = seqIds.filter(id => !batchSet.has(id));
  const onlyBatch = batchIds.filter(id => !seqSet.has(id));

  console.log(`SEQ unique: ${seqSet.size}`);
  console.log(`BATCH unique: ${batchSet.size}`);
  console.log(`Only in SEQ: ${onlySeq.length}`);
  console.log(`Only in BATCH: ${onlyBatch.length}`);

  if (onlySeq.length > 0) {
    console.log(`Only SEQ sample: ${onlySeq.slice(0, 20).join(", ")}`);
  }
  if (onlyBatch.length > 0) {
    console.log(`Only BATCH sample: ${onlyBatch.slice(0, 20).join(", ")}`);
  }

  const ok = onlySeq.length === 0 && onlyBatch.length === 0;
  console.log(ok ? "✅ Compare passed: наборы ID совпали." : "❌ Compare failed: наборы ID различаются.");

  return {
    ok: ok,
    sequentialCount: seqSet.size,
    batchCount: batchSet.size,
    onlySequential: onlySeq.length,
    onlyBatch: onlyBatch.length
  };
}

function debugCompareLeadIdCollectionHard() {
   debugCompareLeadIdCollection("2025-01-01T00:00:00+03:00", "2026-01-01T00:00:00+03:00", 20);
}

/**
 * ТЕСТ: ветвление OR по датам.
 *
 * 1) Сбор ID по «Дата создания» в полуинтервале [start, end) — как в выгрузке.
 * 2) Сбор ID по дате квалификации в [start, end), но с исключением окна по DATE_CREATE.
 *    Дополнение к [start, end) по дате создания: DATE_CREATE < start  ИЛИ  DATE_CREATE >= end.
 *    В REST — два запроса с AND (квалиф. в период + одна ветка по созданию) и объединение Set.
 *
 * 3) Контроль: (все ID с квалиф. в периоде) \\ (ID с датой создания в [start, end)) должны совпасть с (2).
 *
 * @param {string} startIso - Начало [start, end) для полей (как в export)
 * @param {string} endIso - Конец [start, end) (исключая границу)
 * @param {string} qualifiedFieldId - Тех. ID поля даты квалификации, например "UF_CRM_..."
 * @returns {Object} Сводка счётчиков и флаг совпадения контроля
 */
function debugTestQualifiedExcludingCreationPeriod(startIso, endIso, qualifiedFieldId) {
  if (!startIso || !endIso || !qualifiedFieldId) {
    throw new Error("Нужны startIso, endIso и qualifiedFieldId (например UF_CRM_...).");
  }

  const range = (fieldId) => ({
    [`>=${fieldId}`]: startIso,
    [`<${fieldId}`]: endIso
  });

  console.log("🧪 Тест: DATE_CREATE + квалиф. с исключением окна по созданию");
  console.log(`Период [start, end): ${startIso} … ${endIso}`);

  const rCreation = _collectLeadIdsByCursor(
    range("DATE_CREATE"),
    null,
    "[1] DATE_CREATE"
  );
  const setCreation = new Set(rCreation.ids);
  console.log(`[1] Только дата создания в периоде: ${setCreation.size} ID`);

  const fQual = range(qualifiedFieldId);
  const rQBefore = _collectLeadIdsByCursor(
    Object.assign({}, fQual, { ["<DATE_CREATE"]: startIso }),
    null,
    "[2a] квалиф. в периоде, DATE_CREATE < start"
  );
  const rQAfter = _collectLeadIdsByCursor(
    Object.assign({}, fQual, { [">=DATE_CREATE"]: endIso }),
    null,
    "[2b] квалиф. в периоде, DATE_CREATE >= end"
  );
  const setExcl = new Set();
  rQBefore.ids.forEach(function (id) { setExcl.add(id); });
  rQAfter.ids.forEach(function (id) { setExcl.add(id); });
  console.log(`[2a] новых (до start): ${rQBefore.total}, [2b] (с end): ${rQAfter.total}, объединение: ${setExcl.size}`);

  const rQualAll = _collectLeadIdsByCursor(
    fQual,
    null,
    "[ref] квалиф. в периоде (все по полю)"
  );
  const setQualAll = new Set(rQualAll.ids);
  console.log(`[ref] Квалиф. в периоде (без отсечки по созданию): ${setQualAll.size} ID`);

  const setFromDiff = new Set();
  rQualAll.ids.forEach(function (id) {
    if (!setCreation.has(id)) setFromDiff.add(id);
  });

  let onlyExcl = 0;
  setExcl.forEach(function (id) {
    if (!setFromDiff.has(id)) onlyExcl++;
  });
  let onlyDiff = 0;
  setFromDiff.forEach(function (id) {
    if (!setExcl.has(id)) onlyDiff++;
  });

  const ok = onlyExcl === 0 && onlyDiff === 0;
  if (ok) {
    console.log("✅ Контроль: (квалиф. в периоде) \\ (создан в периоде) = ветка с исключением по DATE_CREATE");
  } else {
    console.log(`❌ Расхождение: only в (2) не в diff: ${onlyExcl}, only в diff не в (2): ${onlyDiff}`);
  }

  return {
    countCreation: setCreation.size,
    countQualInPeriod: setQualAll.size,
    countExclusionBranch: setExcl.size,
    countQualMinusCreation: setFromDiff.size,
    mismatchOnlyInExcl: onlyExcl,
    mismatchOnlyInDiff: onlyDiff,
    controlOk: ok
  };
}

/**
 * Подставляет UF поля «квалификация» для теста (по приоритету):
 * 1) Script property `TEST_QUALIFIED_UF` = "UF_CRM_..."
 * 2) константа `DEBUG_APRIL2026_QUAL_UF` в этом файле
 * 3) первое поле из GetLiveFieldsMap(), у которого в русском названии есть "квалиф"
 *
 * @returns {string}
 * @private
 */
function _resolveQualifiedUfForTest_() {
  try {
    const p = PropertiesService.getScriptProperties().getProperty("TEST_QUALIFIED_UF");
    if (p && String(p).trim()) return String(p).trim();
  } catch (e) {
    console.warn("Script properties: " + e.message);
  }
  if (typeof DEBUG_APRIL2026_QUAL_UF === "string" && /^UF_CRM_\d+/.test(DEBUG_APRIL2026_QUAL_UF.trim())) {
    return DEBUG_APRIL2026_QUAL_UF.trim();
  }
  const live = GetLiveFieldsMap();
  for (const name in live) {
    if (!/квалиф/i.test(name)) continue;
    const id = live[name] && live[name].id;
    if (id && String(id).indexOf("UF_") === 0) {
      console.log(`Авто: взято поле квалификации по подписи «${name}» → ${id}`);
      return id;
    }
  }
  throw new Error(
    "Укажите поле квалификации: Script property TEST_QUALIFIED_UF, либо DEBUG_APRIL2026_QUAL_UF в Test.js, либо дайте полю RU-название с «квалиф» в crm.lead.fields"
  );
}

/**
 * Подставьте UF поля квалификации, если нет property и нет подходящего поля в CRM по слову «квалиф».
 * Пример: "UF_CRM_1234567890"
 */
var DEBUG_APRIL2026_QUAL_UF = "";

/**
 * Быстрый запуск: апрель 2026, интервал [2026-04-01, 2026-05-01) в МСК.
 * В редакторе выбери `debugTestQualifiedExcludingCreationPeriod_April2026` → Run.
 *
 * @returns {ReturnType<typeof debugTestQualifiedExcludingCreationPeriod>}
 */
function debugTestQualifiedExcludingCreationPeriod_April2026() {
  return debugTestQualifiedExcludingCreationPeriod(
    "2026-04-01T00:00:00+03:00",
    "2026-05-01T00:00:00+03:00",
    _resolveQualifiedUfForTest_()
  );
}

/**
 * Возвращает total по окну максимально "дешево":
 * - сначала пытается взять total из первого запроса (start=0),
 * - если total в ответе отсутствует, досчитывает пагинацией по 50.
 *
 * @param {string} startIso
 * @param {string} endIso
 * @returns {number}
 * @private
 */
function _debugCountCallsInWindow_(startIso, endIso) {
  const baseParams = {
    FILTER: {
      ">=CALL_START_DATE": startIso,
      "<CALL_START_DATE": endIso
    },
    SORT: "CALL_START_DATE",
    ORDER: "ASC",
    start: 0
  };

  const first = CallBitrix("voximplant.statistic.get", baseParams) || {};
  const rows = Array.isArray(first.result) ? first.result : [];
  const totalRaw = (first.total != null) ? first.total : first.result_total;
  const totalNum = Number(totalRaw);
  if (!isNaN(totalNum) && totalNum >= 0) {
    return totalNum;
  }

  // fallback: считаем страницами по 50
  let count = rows.length;
  if (rows.length < 50) return count;

  let start = 50;
  while (true) {
    const page = CallBitrix("voximplant.statistic.get", Object.assign({}, baseParams, { start: start })) || {};
    const pRows = Array.isArray(page.result) ? page.result : [];
    count += pRows.length;
    if (pRows.length < 50) break;
    start += 50;
  }
  return count;
}

/**
 * Быстрый подсчёт количества звонков за период по окнам.
 *
 * @param {{start:string, end:string}} period
 * @param {number} windowHours
 * @returns {number}
 * @private
 */
function _debugCountCallsByPeriodWindows_(period, windowHours) {
  const startMs = new Date(period.start).getTime();
  const endMs = new Date(period.end).getTime();
  if (isNaN(startMs) || isNaN(endMs) || startMs >= endMs) return 0;

  const stepMs = Math.max(1, parseInt(windowHours, 10) || 72) * 60 * 60 * 1000;
  let cursor = startMs;
  let total = 0;
  let windowNo = 0;
  const windowsTotal = Math.ceil((endMs - startMs) / stepMs);

  while (cursor < endMs) {
    windowNo++;
    const wEnd = Math.min(cursor + stepMs, endMs);
    const wStartIso = _formatIsoMsk(new Date(cursor));
    const wEndIso = _formatIsoMsk(new Date(wEnd));
    const c = _debugCountCallsInWindow_(wStartIso, wEndIso);
    total += c;
    if (windowNo === 1 || windowNo === windowsTotal || windowNo % 10 === 0) {
      console.log(`COUNT window ${windowNo}/${windowsTotal}: +${c}, total=${total}`);
    }
    cursor = wEnd;
  }

  return total;
}

/**
 * ТЕСТ: сравнить фактический count звонков и ID-оценку по годам 2023..2026.
 * Пишет всё в Execution log.
 *
 * @param {number} [fromYear=2023]
 * @param {number} [toYear=2026]
 * @param {number} [windowHours=72]
 */
function debugCallsEstimateByYears(fromYear = 2023, toYear = 2026, windowHours = 72) {
  const y1 = parseInt(fromYear, 10) || 2023;
  const y2 = parseInt(toYear, 10) || 2026;
  const startYear = Math.min(y1, y2);
  const endYear = Math.max(y1, y2);

  const now = new Date();
  const nowMs = now.getTime();
  const rows = [];

  console.log(`🧪 Calls estimate test: years ${startYear}..${endYear}, windowHours=${windowHours}`);

  for (let y = startYear; y <= endYear; y++) {
    const periodStart = `${y}-01-01T00:00:00+03:00`;
    const nextYearStartIso = `${y + 1}-01-01T00:00:00+03:00`;
    const nextYearStartMs = new Date(nextYearStartIso).getTime();
    const periodEnd = (nextYearStartMs > nowMs) ? _formatIsoMsk(now) : nextYearStartIso;

    if (new Date(periodStart).getTime() >= new Date(periodEnd).getTime()) {
      console.log(`- ${y}: период в будущем, пропуск`);
      continue;
    }

    const period = { start: periodStart, end: periodEnd };
    console.log(`\n📅 YEAR ${y}: ${period.start} -> ${period.end}`);

    const estimate = _estimateCallsByIdRange_(period);
    const estimated = estimate && estimate.estimated != null ? Number(estimate.estimated) : null;
    const actual = _debugCountCallsByPeriodWindows_(period, windowHours);

    const ratio = (estimated && estimated > 0) ? (actual / estimated) : null;
    const diffAbs = (estimated != null) ? (actual - estimated) : null;
    const diffPct = (estimated && estimated > 0) ? ((diffAbs / estimated) * 100) : null;

    rows.push({
      year: y,
      actual: actual,
      estimated: estimated,
      ratio: ratio,
      diffAbs: diffAbs,
      diffPct: diffPct
    });

    console.log(
      `YEAR ${y}: actual=${actual}, estimate=${estimated}, ratio=${ratio != null ? ratio.toFixed(4) : "n/a"}, diff=${diffAbs != null ? diffAbs : "n/a"} (${diffPct != null ? diffPct.toFixed(2) + "%" : "n/a"})`
    );
  }

  if (rows.length === 0) {
    console.log("Нет данных для отчёта.");
    return;
  }

  const valid = rows.filter(r => r.ratio != null && isFinite(r.ratio));
  const avgRatio = valid.length ? valid.reduce((s, r) => s + r.ratio, 0) / valid.length : null;
  const medianRatio = valid.length
    ? valid
        .map(r => r.ratio)
        .sort((a, b) => a - b)[Math.floor(valid.length / 2)]
    : null;

  console.log("\n=== SUMMARY ===");
  rows.forEach(r => {
    console.log(
      `${r.year}: actual=${r.actual}, estimate=${r.estimated}, ratio=${r.ratio != null ? r.ratio.toFixed(4) : "n/a"}, diff=${r.diffAbs != null ? r.diffAbs : "n/a"} (${r.diffPct != null ? r.diffPct.toFixed(2) + "%" : "n/a"})`
    );
  });
  console.log(`avg_ratio(actual/estimate)=${avgRatio != null ? avgRatio.toFixed(4) : "n/a"}`);
  console.log(`median_ratio(actual/estimate)=${medianRatio != null ? medianRatio.toFixed(4) : "n/a"}`);
  if (avgRatio != null) {
    console.log(`recommended_coef_for_estimate ≈ ${avgRatio.toFixed(4)} (estimate * coef)`);
  }
}

/**
 * ТЕСТ ТОЧНОСТИ по "Дата создания":
 * 1) База: ID только по основному фильтру DATE_CREATE в [start, end)
 * 2) Финал: ID после полного OR-сбора
 * 3) Подсчёт в финале тех, у кого DATE_CREATE реально в [start, end)
 * 4) Подсчёт строк на листе выгрузки с DATE_CREATE в [start, end)
 *
 * Нужен для диагностики расхождения вида "в логе 843, на листе 850".
 *
 * @param {string|null} [periodStartIso=null] - Явное начало периода, иначе берётся из named ranges.
 * @param {string|null} [periodEndIso=null] - Явный конец периода, иначе из named ranges.
 * @returns {Object}
 */
function debugAuditDateCreatePrecision(periodStartIso = null, periodEndIso = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = _initializeConfig(ss);
  if (!config) throw new Error("Не удалось инициализировать config.");

  const period = (periodStartIso && periodEndIso)
    ? { start: String(periodStartIso), end: String(periodEndIso) }
    : _prepareExportPeriod(config);

  const startMs = new Date(period.start).getTime();
  const endMs = new Date(period.end).getTime();
  if (!isFinite(startMs) || !isFinite(endMs) || startMs >= endMs) {
    throw new Error(`Некорректный период: ${period.start} .. ${period.end}`);
  }

  const liveMap = GetLiveFieldsMap();
  const mainFieldId = (liveMap["Дата создания"] && liveMap["Дата создания"].id) ? liveMap["Дата создания"].id : "DATE_CREATE";

  console.log("🧪 debugAuditDateCreatePrecision");
  console.log(`Период [start, end): ${period.start} .. ${period.end}`);
  console.log(`Основное поле: ${mainFieldId}`);

  // 1) Только основной фильтр
  const baseRes = _collectLeadIdsByCursor(
    { [`>=${mainFieldId}`]: period.start, [`<${mainFieldId}`]: period.end },
    null,
    ""
  );
  const baseSet = new Set(baseRes.ids.map(String));
  console.log(`[BASE] IDs по DATE_CREATE в периоде: ${baseSet.size}`);

  // 2) Полный OR-сбор (как в рабочей выгрузке)
  const allLeads = GetLeadsByFiltersMap(period, config.requiredHeaders, liveMap, config.filterableHeaders, null);
  const finalSet = new Set(allLeads.map(l => String(l.ID)));
  console.log(`[FINAL] IDs после OR/дедуп: ${finalSet.size}`);

  // 3) Из финала берём только тех, у кого DATE_CREATE в периоде
  const finalInRange = [];
  const finalOutRange = [];
  allLeads.forEach(l => {
    const d = l && l.DATE_CREATE ? new Date(l.DATE_CREATE).getTime() : NaN;
    if (!isFinite(d)) return;
    if (d >= startMs && d < endMs) finalInRange.push(String(l.ID));
    else finalOutRange.push(String(l.ID));
  });
  const finalInRangeSet = new Set(finalInRange);
  console.log(`[FINAL in range] IDs с DATE_CREATE в периоде: ${finalInRangeSet.size}`);
  console.log(`[FINAL out range] IDs вне периода: ${finalOutRange.length}`);

  // 4) Сравнение BASE vs FINAL(in range)
  const onlyBase = [];
  baseSet.forEach(id => { if (!finalInRangeSet.has(id)) onlyBase.push(id); });
  const onlyFinalInRange = [];
  finalInRangeSet.forEach(id => { if (!baseSet.has(id)) onlyFinalInRange.push(id); });

  console.log(`Δ only BASE (есть в base, нет в finalInRange): ${onlyBase.length}`);
  console.log(`Δ only FINAL(in range) (есть в finalInRange, нет в base): ${onlyFinalInRange.length}`);
  if (onlyBase.length) {
    console.log(`sample only BASE: ${onlyBase.slice(0, 20).join(", ")}`);
  }
  if (onlyFinalInRange.length) {
    console.log(`sample only FINAL(in range): ${onlyFinalInRange.slice(0, 20).join(", ")}`);
  }

  // 5) Проверка строк на листе (что реально видит пользователь)
  let sheetCountInRange = null;
  if (config.leadsSheet) {
    const sh = ss.getSheetByName(config.leadsSheet);
    if (sh) {
      const values = sh.getDataRange().getValues();
      if (values.length > 1) {
        const headers = values[0].map(h => String(h || "").trim());
        const idIdx = headers.indexOf("ID");
        const dcIdx = headers.indexOf("Дата создания");
        if (idIdx !== -1 && dcIdx !== -1) {
          let cnt = 0;
          const idsInRangeOnSheet = [];
          for (let r = 1; r < values.length; r++) {
            const rawDate = values[r][dcIdx];
            const ms = (rawDate instanceof Date) ? rawDate.getTime() : new Date(rawDate).getTime();
            if (!isFinite(ms)) continue;
            if (ms >= startMs && ms < endMs) {
              cnt++;
              idsInRangeOnSheet.push(String(values[r][idIdx] || "").trim());
            }
          }
          sheetCountInRange = cnt;
          const sheetSet = new Set(idsInRangeOnSheet.filter(Boolean));
          console.log(`[SHEET] строк с "Дата создания" в периоде: ${sheetCountInRange}`);

          const onlySheet = [];
          sheetSet.forEach(id => { if (!finalInRangeSet.has(id)) onlySheet.push(id); });
          const missingOnSheet = [];
          finalInRangeSet.forEach(id => { if (!sheetSet.has(id)) missingOnSheet.push(id); });

          console.log(`Δ sheet-only vs finalInRange: ${onlySheet.length}`);
          console.log(`Δ missing on sheet vs finalInRange: ${missingOnSheet.length}`);
          if (onlySheet.length) console.log(`sample sheet-only: ${onlySheet.slice(0, 20).join(", ")}`);
          if (missingOnSheet.length) console.log(`sample missing-on-sheet: ${missingOnSheet.slice(0, 20).join(", ")}`);
        } else {
          console.log(`⚠ На листе "${config.leadsSheet}" не найдены колонки ID/Дата создания`);
        }
      }
    }
  }

  const result = {
    period: period,
    baseCount: baseSet.size,
    finalCount: finalSet.size,
    finalInRangeCount: finalInRangeSet.size,
    finalOutRangeCount: finalOutRange.length,
    onlyBaseCount: onlyBase.length,
    onlyFinalInRangeCount: onlyFinalInRange.length,
    sheetInRangeCount: sheetCountInRange
  };

  console.log("=== RESULT ===");
  console.log(JSON.stringify(result));
  return result;
}