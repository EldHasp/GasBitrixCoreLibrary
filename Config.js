/**
 * Глобальные настройки подключения к Битрикс24
 */
const BITRIX_URL_BASE = 'https://braincon.bitrix24.ru/';
// const BITRIX_URL = BITRIX_URL_BASE+'rest/218/3ttpxdkg6fgbr85y/';
// const BITRIX_URL_REST_3_0 = BITRIX_URL_BASE+'rest/api/218/3ttpxdkg6fgbr85y/';
const BITRIX_URL_LEAD = BITRIX_URL_BASE+'crm/lead/details/';
const BITRIX_URL_Deal = BITRIX_URL_BASE+'crm/deal/details/';
const BITRIX_URL_LEADS_FIELDS = BITRIX_URL_BASE+'crm/configs/fields/CRM_LEAD/edit/';

/**
 * Глобальный URL, который увидят методы библиотеки.
 * Сама строка вебхука здесь НЕ написана, она подтянется из "сейфа".
 */
const BITRIX_URL = (function() {
  const secret = PropertiesService.getScriptProperties().getProperty('BITRIX_SECRET_PART');
  
  if (!secret) {
    // Если кто-то скопирует твой код, но у него нет доступа к ScriptProperties — он получит ошибку
    console.error("❌ Критическая ошибка: Секретный ключ библиотеки не найден!");
    return null;
  }
  
  // Возвращаем https://bitrix24.rurest/218/3ttpxdkg6fgbr85y/
  return `${BITRIX_URL_BASE}rest/${secret}`;
})();

/**
 * Глобальный URL для REST 3.0, который увидят методы библиотеки.
 * Сама строка вебхука здесь НЕ написана, она подтянется из "сейфа".
 */
const BITRIX_URL_REST_3_0 = (function() {
  const secret = PropertiesService.getScriptProperties().getProperty('BITRIX_SECRET_PART');
  
  if (!secret) {
    // Если кто-то скопирует твой код, но у него нет доступа к ScriptProperties — он получит ошибку
    console.error("❌ Критическая ошибка: Секретный ключ библиотеки не найден!");
    return null;
  }
  
  // Возвращаем https://bitrix24.rurest/218/3ttpxdkg6fgbr85y/
  return `${BITRIX_URL_BASE}rest/api/${secret}`;
})();

/**
 * Настройки справочников (Инфоблоки/Списки)
 */
const IBLOCK_OFFICES_ID = 128;
const IBLOCK_OFFICES_TYPE = "bitrix_processes";

/**
 * СЛУЖЕБНАЯ: Запустить один раз для сохранения ключа в "железо" библиотеки
 */
/* function _DEV_saveWebhookInSecretStorage() {
  // Сохраняем только чувствительную часть: "/ID_сотрудника/Токен/"
  const secretPart = "218/__webhook__code__/"; 
  
  // ScriptProperties — это аналог Environment Variables уровня проекта
  PropertiesService.getScriptProperties().setProperty('BITRIX_SECRET_PART', secretPart);
  
  console.log("✅ Секрет успешно сохранен в ScriptProperties библиотеки.");
} */