/**
 * Временный отладочный лог синхронизации: файл на Drive + кнопка «Сохранить» в Progress.html.
 * Сохраняется только по явному нажатию; без сохранения временный файл удаляется
 * при старте следующего экспорта или (best effort) при закрытии окна прогресса.
 */
var _SYNC_DEBUG_KEY_TEMP_ = "SYNC_DEBUG_TEMP_FILE";
var _SYNC_DEBUG_KEY_PENDING_ = "SYNC_DEBUG_SAVE_PENDING";
var _SYNC_DEBUG_CACHE_TTL_ = 21600;

/**
 * @returns {boolean}
 * @protected (из Main / Progress)
 */
function syncDebugLogHasPendingSave_() {
  try {
    return !!CacheService.getUserCache().get(_SYNC_DEBUG_KEY_PENDING_);
  } catch (e) {
    return false;
  }
}

/**
 * Перед новым отладочным прогоном: удаляет незавершённую сессию (лог с прошлого раза).
 */
function syncDebugLogDiscardAllPending_() {
  const cache = CacheService.getUserCache();
  const raw = cache.get(_SYNC_DEBUG_KEY_PENDING_) || cache.get(_SYNC_DEBUG_KEY_TEMP_);
  if (!raw) {
    cache.remove(_SYNC_DEBUG_KEY_TEMP_);
    cache.remove(_SYNC_DEBUG_KEY_PENDING_);
    return;
  }
  let o;
  try {
    o = JSON.parse(raw);
  } catch (e) {
    cache.remove(_SYNC_DEBUG_KEY_TEMP_);
    cache.remove(_SYNC_DEBUG_KEY_PENDING_);
    return;
  }
  if (o && o.fileId) {
    try {
      DriveApp.getFileById(o.fileId).setTrashed(true);
    } catch (e) {
      /* file уже удалён вручную / нет прав */
    }
  }
  cache.remove(_SYNC_DEBUG_KEY_TEMP_);
  cache.remove(_SYNC_DEBUG_KEY_PENDING_);
}

/**
 * @param {boolean} enabled — им. диапазон «ПодробныйЛогСинхронизации»
 */
function syncDebugLogStart_(enabled) {
  syncDebugLogDiscardAllPending_();
  if (!enabled) {
    return;
  }
  const cache = CacheService.getUserCache();
  const tz = Session.getScriptTimeZone() || "Europe/Moscow";
  const stamp = Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");
  const suffix = Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  const name = "Bitrix24_sync_debug_" + stamp + "_" + suffix + ".tmp.txt";
  const file = DriveApp.createFile(name, "", MimeType.PLAIN_TEXT);
  const payload = { fileId: file.getId(), created: Date.now() };
  cache.put(_SYNC_DEBUG_KEY_TEMP_, JSON.stringify(payload), _SYNC_DEBUG_CACHE_TTL_);
}

/**
 * @param {string} line — уже с временем, если из updateStatus; или сырой текст.
 */
function syncDebugLogAppend_(line) {
  if (!line && line !== "") {
    return;
  }
  const cache = CacheService.getUserCache();
  const raw = cache.get(_SYNC_DEBUG_KEY_TEMP_);
  if (!raw) {
    return;
  }
  let o;
  try {
    o = JSON.parse(raw);
  } catch (e) {
    return;
  }
  if (!o || !o.fileId) {
    return;
  }
  try {
    const file = DriveApp.getFileById(o.fileId);
    const prev = file.getBlob().getDataAsString();
    file.setContent(prev + String(line) + "\n");
  } catch (e) {
    console.warn("syncDebugLogAppend_:", e.message);
  }
}

/**
 * В конце экспорта: даём UI показать кнопку «Сохранить».
 */
function syncDebugLogFinishRun_() {
  const cache = CacheService.getUserCache();
  const raw = cache.get(_SYNC_DEBUG_KEY_TEMP_);
  if (!raw) {
    return;
  }
  cache.put(_SYNC_DEBUG_KEY_PENDING_, raw, _SYNC_DEBUG_CACHE_TTL_);
}

/**
 * Копирует .tmp в постоянный .txt (новый файл, имя с timestamp), .tmp в корзину.
 * @returns {{ok: boolean, name?: string, url?: string, id?: string, message?: string}}
 * @protected
 */
function syncDebugLogSavePermanently_() {
  const cache = CacheService.getUserCache();
  const raw = cache.get(_SYNC_DEBUG_KEY_PENDING_) || cache.get(_SYNC_DEBUG_KEY_TEMP_);
  if (!raw) {
    return { ok: false, message: "Нет отладочного лога (возможно, сессия истекла или режим не включён)." };
  }
  let o;
  try {
    o = JSON.parse(raw);
  } catch (e) {
    return { ok: false, message: "Некорректные данные сессии в кэше." };
  }
  if (!o || !o.fileId) {
    return { ok: false, message: "Нет файла для сохранения." };
  }
  let tmp;
  let content;
  try {
    tmp = DriveApp.getFileById(o.fileId);
    content = tmp.getBlob().getDataAsString();
  } catch (e) {
    return { ok: false, message: "Временный файл не найден: " + (e && e.message ? e.message : e) };
  }
  const tz = Session.getScriptTimeZone() || "Europe/Moscow";
  const stamp = Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");
  const name = "Bitrix24_sync_log_" + stamp + ".txt";
  const perm = DriveApp.createFile(name, content, MimeType.PLAIN_TEXT);
  try {
    tmp.setTrashed(true);
  } catch (e) {
    /* */ }
  cache.remove(_SYNC_DEBUG_KEY_TEMP_);
  cache.remove(_SYNC_DEBUG_KEY_PENDING_);
  return { ok: true, name: name, url: perm.getUrl(), id: perm.getId() };
}

/**
 * Удалить временный файл, очистить кэш (закрытие окна без сохранения).
 * @returns {{ok: boolean}}
 * @protected
 */
function syncDebugLogDiscardAfterClose_() {
  syncDebugLogDiscardAllPending_();
  return { ok: true };
}
