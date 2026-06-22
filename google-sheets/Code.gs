/**
 * Google Apps Script — приём заявок с сайта прямо в Google-таблицу.
 *
 * Сайт (index.html) отправляет POST с полями: name, email, phone, tg, magnet.
 * Скрипт дописывает их новой строкой в таблицу. Никакой Google Формы не нужно.
 *
 * Установка — см. README.md рядом с этим файлом.
 */

// Имя листа, куда писать заявки (если нет — берётся первый лист).
var SHEET_NAME = 'Заявки';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // Заголовки при первом запуске
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Материал', 'Имя', 'Почта', 'Телефон', 'Telegram']);
    }

    var p = (e && e.parameter) ? e.parameter : {};
    sheet.appendRow([
      new Date(),
      p.magnet || '',
      p.name || '',
      p.email || '',
      p.phone || '',
      p.tg || '',
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Чтобы открытие URL в браузере не выдавало ошибку.
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, info: 'Sestram leads endpoint' }))
    .setMimeType(ContentService.MimeType.JSON);
}
