/**
 * Google Apps Script — приём заявок с сайта прямо в Google-таблицу.
 *
 * Сайт (index.html) отправляет POST с полями: name, email, phone, tg, magnet.
 * Скрипт дописывает их новой строкой в таблицу. Никакой Google Формы не нужно.
 *
 * Защита от дублей: если в таблице уже есть строка с такой же ПОЧТОЙ
 * (а если почты нет — с таким же ТЕЛЕФОНОМ), новая строка не создаётся.
 * Вместо этого обновляется дата и в колонку «Материал» дописывается то,
 * что человек запросил (например: «Гайды по планированию, Справочник постящейся»).
 *
 * Установка и обновление — см. README.md рядом с этим файлом.
 */

// Имя листа, куда писать заявки (если нет — берётся первый лист).
var SHEET_NAME = 'Заявки';

// ID вашей Google-таблицы (из ссылки .../spreadsheets/d/ЭТОТ_ID/edit).
// Благодаря ему скрипт пишет в нужную таблицу, даже если создан отдельно от неё.
var SHEET_ID = '118uNb4a4SLIwU92P3o0ROB0nLqvDTguK6C2PfZy3SCg';

// Колонки: 1=Дата 2=Материал 3=Имя 4=Почта 5=Телефон 6=Telegram
var COL = { date: 1, magnet: 2, name: 3, email: 4, phone: 5, tg: 6 };

function getSpreadsheet_() {
  if (SHEET_ID && SHEET_ID.indexOf('ВСТАВИТЬ') === -1) {
    return SpreadsheetApp.openById(SHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    // Заголовки при первом запуске
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Материал', 'Имя', 'Почта', 'Телефон', 'Telegram']);
    }

    var p = (e && e.parameter) ? e.parameter : {};
    var name = String(p.name || '').trim();
    var email = String(p.email || '').trim();
    var phone = String(p.phone || '').trim();
    var tg = String(p.tg || '').trim();
    var magnet = String(p.magnet || '').trim();

    var emailKey = email.toLowerCase();

    // Ищем уже существующую строку по почте (или по телефону, если почты нет)
    var lastRow = sheet.getLastRow();
    var foundRow = -1;
    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      for (var i = 0; i < values.length; i++) {
        var rowEmail = String(values[i][COL.email - 1] || '').trim().toLowerCase();
        var rowPhone = String(values[i][COL.phone - 1] || '').trim();
        if (emailKey && rowEmail === emailKey) { foundRow = i + 2; break; }
        if (!emailKey && phone && rowPhone === phone) { foundRow = i + 2; break; }
      }
    }

    if (foundRow > 0) {
      // Человек уже есть — обновляем дату и дописываем материал
      sheet.getRange(foundRow, COL.date).setValue(new Date());

      var current = String(sheet.getRange(foundRow, COL.magnet).getValue() || '');
      var mats = current ? current.split(',').map(function (s) { return s.trim(); }).filter(String) : [];
      if (magnet && mats.indexOf(magnet) === -1) mats.push(magnet);
      sheet.getRange(foundRow, COL.magnet).setValue(mats.join(', '));

      // Дозаполняем пустые поля (имя/телефон/telegram), не затирая существующие
      if (name && !String(sheet.getRange(foundRow, COL.name).getValue()).trim()) sheet.getRange(foundRow, COL.name).setValue(name);
      if (phone && !String(sheet.getRange(foundRow, COL.phone).getValue()).trim()) sheet.getRange(foundRow, COL.phone).setValue(phone);
      if (tg && !String(sheet.getRange(foundRow, COL.tg).getValue()).trim()) sheet.getRange(foundRow, COL.tg).setValue(tg);
    } else {
      // Новый человек — добавляем строку
      sheet.appendRow([new Date(), magnet, name, email, phone, tg]);
    }

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
