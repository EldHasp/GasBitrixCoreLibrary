/**
 * Подробный лог синхронизации (отладка): накопление в памяти, временный файл в Drive,
 * кнопка «Сохранить» в Progress.html, удаление временного при закрытии окна (best effort)
 * и уборка «хвостов» при следующем запуске.
 */

var BTX_DEBUG_CACHE_TEMP_ID = "BTX_DEBUG_TEMP_FILE_ID";
var BTX_DEBUG_TEMP_TITLE_PREFIX = "Bitrix24_sync_debug_temp_";
var BTX_DEBUG_PERM_TITLE_PREFIX = "Bitrix24_sync_log_";

/** Состояние одного выполнения ExportLeadsToSheet (сбрасывается в конце). */
var BTX_DEBUG_EXEC_ = { verbose: false, lines: [] };

/**
 * @param {boolean} verbose — из именованного диапазона «ПодробныйЛогСинхронизации».
 */
function BtxDebugLogOnSyncStart_(verbose) {
  BTX_DEBUG_EXEC_.verbose = !!verbose;
  BTX_DEBUG_EXEC_.lines = [];
  BtxDebugTrashTempFileByCacheId_();
  BtxDebugCleanupStaleTempFiles_();
}

function BtxDebugAppendLine_(lineWithTime) {
  if (!BTX_DEBUG_EXEC_.verbose || !lineWithTime) return;
  BTX_DEBUG_EXEC_.lines.push(lineWithTime);
}

function BtxDebugResetExec_() {
  BTX_DEBUG_EXEC_.verbose = false;
  BTX_DEBUG_EXEC_.lines = [];
}

/**
 * После завершения экспорта: при ручном режиме и включённой отладке — пишем временный файл и кладём id в кэш.
 * @param {boolean} isManual
 */
function BtxDebugFinalizeTempAfterSync_(isManual) {
  try {
    if (!isManual || !BTX_DEBUG_EXEC_.verbose || BTX_DEBUG_EXEC_.lines.length === 0) {
      BtxDebugResetExec_();
      return;
    }
    const body = BTX_DEBUG_EXEC_.lines.join("\n");
    const title =
      BTX_DEBUG_TEMP_TITLE_PREFIX +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") +
      ".txt";
    const file = DriveApp.createFile(title, body, MimeType.PLAIN_TEXT);
    CacheService.getUserCache().put(BTX_DEBUG_CACHE_TEMP_ID, file.getId(), 21600);
  } catch (e) {
    console.warn("BtxDebugFinalizeTempAfterSync_: " + e.message);
  } finally {
    BtxDebugResetExec_();
  }
}

function BtxDebugTrashTempFileByCacheId_() {
  const cache = CacheService.getUserCache();
  const id = cache.get(BTX_DEBUG_CACHE_TEMP_ID);
  if (!id) return;
  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (e) {
    /* уже удалён */
  }
  cache.remove(BTX_DEBUG_CACHE_TEMP_ID);
}

/**
 * Временные файлы старше 48 ч — в корзину (если поиск доступен).
 */
function BtxDebugCleanupStaleTempFiles_() {
  try {
    const q = 'title contains "' + BTX_DEBUG_TEMP_TITLE_PREFIX + '" and trashed = false';
    const it = DriveApp.searchFiles(q);
    const cutoff = Date.now() - 48 * 60 * 60 * 1000;
    while (it.hasNext()) {
      const f = it.next();
      if (f.getDateCreated().getTime() < cutoff) {
        try {
          f.setTrashed(true);
        } catch (e) {
          /* ignore */
        }
      }
    }
  } catch (e) {
    console.warn("BtxDebugCleanupStaleTempFiles_: " + e.message);
  }
}

/**
 * Для Progress.html: есть ли несохранённый временный лог.
 * @returns {{ canSave: boolean, fileId: string|null }}
 */
function getLibraryDebugSaveInfo() {
  const id = CacheService.getUserCache().get(BTX_DEBUG_CACHE_TEMP_ID);
  return { canSave: !!id, fileId: id || null };
}

/**
 * Копия в постоянный файл с новым именем; временный — в корзину.
 * @returns {{ ok: boolean, url?: string, name?: string, message?: string }}
 */
function BtxSaveDebugSyncLogToDrive() {
  const cache = CacheService.getUserCache();
  const fileId = cache.get(BTX_DEBUG_CACHE_TEMP_ID);
  if (!fileId) {
    return { ok: false, message: "Нет временного лога для сохранения." };
  }
  try {
    const temp = DriveApp.getFileById(fileId);
    const content = temp.getBlob().getDataAsString();
    const name =
      BTX_DEBUG_PERM_TITLE_PREFIX +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmmss") +
      ".txt";
    const perm = DriveApp.createFile(name, content, MimeType.PLAIN_TEXT);
    temp.setTrashed(true);
    cache.remove(BTX_DEBUG_CACHE_TEMP_ID);
    return { ok: true, url: perm.getUrl(), name: name };
  } catch (e) {
    return { ok: false, message: e.message || String(e) };
  }
}

/**
 * Закрытие окна прогресса без сохранения — удалить временный файл (best effort).
 */
function BtxDiscardDebugSyncTempSilent() {
  try {
    BtxDebugTrashTempFileByCacheId_();
  } catch (e) {
    /* ignore */
  }
}
