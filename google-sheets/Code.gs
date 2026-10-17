/**
 * Google Apps Script — приём заявок с сайта прямо в Google-таблицу.
 *
 * Сайт (index.html) отправляет POST с полями: name, email, phone, tg, magnet.
 * Скрипт дописывает их новой строкой в таблицу. Никакой Google Формы не нужно.
 *
 * Дополнительно скрипт ОТПРАВЛЯЕТ ЧЕЛОВЕКУ ПИСЬМО со ссылками на материалы —
 * через Resend, с адреса info@urokisestram.com. Если отправить не удалось,
 * заявка всё равно записывается в таблицу, а в колонку «Письмо» попадает причина.
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

// Колонки: 1=Дата 2=Материал 3=Имя 4=Почта 5=Телефон 6=Telegram 7=Письмо
var COL = { date: 1, magnet: 2, name: 3, email: 4, phone: 5, tg: 6, mail: 7 };

var HEADERS = ['Дата', 'Материал', 'Имя', 'Почта', 'Телефон', 'Telegram', 'Письмо'];

// ───────── Письмо с материалами ─────────

// Имя и адрес отправителя, которые увидит получатель.
var MAIL_FROM_NAME = 'Уроки Сёстрам';
var MAIL_FROM = 'info@urokisestram.com';

// Письма уходят через Resend (resend.com) — так адрес отправителя принадлежит
// домену, а личная почта владельца нигде не показывается.
// Ключ хранится в настройках скрипта, НЕ в коде:
// Apps Script → ⚙️ Настройки проекта → Свойства скрипта → RESEND_API_KEY.
var RESEND_ENDPOINT = 'https://api.resend.com/emails';

// Адрес сайта (для ссылок на гайды).
var SITE_URL = 'https://urokisestram.com';

// Гайды по планированию — порядок как на сайте.
var GUIDE_FILES = [
  { label: '5 причин, по которым ты ничего не успеваешь', href: SITE_URL + '/downloads/guide-1.pdf' },
  { label: 'Как выполнять запланированное',               href: SITE_URL + '/downloads/guide-2.pdf' },
  { label: 'Как победить прокрастинацию',                 href: SITE_URL + '/downloads/guide-3.pdf' },
  { label: 'Как вставать рано',                           href: SITE_URL + '/downloads/guide-4.pdf' },
  { label: 'Мои секреты продуктивности',                  href: SITE_URL + '/downloads/guide-5.pdf' },
];

// Справочник постящейся — лежит на Google Drive.
var BOOK_FILE = {
  label: 'Справочник постящейся (PDF)',
  href: 'https://drive.google.com/uc?export=download&id=11ygUPBreZKZW65kqDWVd9o9uk4B7GbwN',
};

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
      sheet.appendRow(HEADERS);
    } else if (!String(sheet.getRange(1, COL.mail).getValue()).trim()) {
      // Таблица создана прежней версией скрипта — дописываем заголовок «Письмо».
      sheet.getRange(1, COL.mail).setValue(HEADERS[COL.mail - 1]);
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
      var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
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
      sheet.appendRow([new Date(), magnet, name, email, phone, tg, '']);
      foundRow = sheet.getLastRow();
    }

    // Письмо с материалами — заявка остаётся в таблице, даже если письмо не ушло.
    var mailStatus = sendMagnetEmail_(email, name, magnet);
    if (foundRow > 0) sheet.getRange(foundRow, COL.mail).setValue(mailStatus);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, mail: mailStatus }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ───────── Отправка письма ─────────

/**
 * Отправляет человеку письмо со ссылками на материалы.
 * Возвращает короткий статус для колонки «Письмо»: 'отправлено' или причину отказа.
 */
function sendMagnetEmail_(email, name, magnet) {
  if (!email) return 'нет почты';

  var isBook = magnet.indexOf('Справочник') !== -1;
  var files = isBook ? [BOOK_FILE] : GUIDE_FILES;
  var subject = isBook ? 'Справочник постящейся' : 'Ваши гайды по планированию';

  var key = PropertiesService.getScriptProperties().getProperty('RESEND_API_KEY');
  if (!key) return 'нет ключа RESEND_API_KEY';

  try {
    var payload = {
      from: MAIL_FROM_NAME + ' <' + MAIL_FROM + '>',
      to: [email],
      subject: subject,
      html: buildMagnetHtml_(name, isBook, files),
      text: buildMagnetText_(name, isBook, files),
    };

    var res = UrlFetchApp.fetch(RESEND_ENDPOINT, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    var code = res.getResponseCode();
    if (code >= 200 && code < 300) return 'отправлено';
    return 'ошибка ' + code + ': ' + String(res.getContentText()).slice(0, 200);
  } catch (err) {
    return 'ошибка: ' + String(err);
  }
}

/** Обращение по имени, если имя есть. */
function greeting_(name) {
  var n = String(name || '').trim();
  return n ? ('Ассаляму алейкум ва рахматуллах, ' + n + '!') : 'Ассаляму алейкум ва рахматуллах!';
}

/** Текстовая версия письма (для почтовых программ без оформления). */
function buildMagnetText_(name, isBook, files) {
  var lines = [greeting_(name), ''];
  lines.push(isBook
    ? 'Отправляю вам «Справочник постящейся». Скачать можно по ссылке:'
    : 'Отправляю вам гайды по планированию. Скачать можно по ссылкам ниже:');
  lines.push('');
  for (var i = 0; i < files.length; i++) {
    lines.push('— ' + files[i].label);
    lines.push('  ' + files[i].href);
  }
  lines.push('');
  lines.push('Ссылки не истекают — можно вернуться к ним в любое время.');
  lines.push('Если что-то не открылось, просто ответьте на это письмо.');
  lines.push('');
  lines.push('Баракаллаху фикум,');
  lines.push(MAIL_FROM_NAME);
  lines.push(SITE_URL);
  return lines.join('\n');
}

/** Оформленная версия письма. */
function buildMagnetHtml_(name, isBook, files) {
  var intro = isBook
    ? 'Отправляю вам «Справочник постящейся». Скачать можно по ссылке ниже.'
    : 'Отправляю вам гайды по планированию. Скачать можно по ссылкам ниже.';

  var items = '';
  for (var i = 0; i < files.length; i++) {
    items +=
      '<tr><td style="padding:0 0 14px;">' +
        '<a href="' + files[i].href + '" style="display:inline-block;font-size:15px;color:#1C1A19;text-decoration:none;border-bottom:1px solid #E0CBCF;padding-bottom:2px;">' +
          escapeHtml_(files[i].label) +
        '</a>' +
      '</td></tr>';
  }

  return '' +
    '<div style="margin:0;padding:28px 16px;background:#FBF6F3;font-family:Arial,Helvetica,sans-serif;color:#1C1A19;">' +
      '<table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" width="100%" style="max-width:520px;margin:0 auto;background:#FFFFFF;border:1px solid #ECE3DF;border-radius:20px;">' +
        '<tr><td style="padding:34px 30px;">' +
          '<p style="margin:0 0 18px;font-size:16px;line-height:1.6;">' + escapeHtml_(greeting_(name)) + '</p>' +
          '<p style="margin:0 0 24px;font-size:15px;line-height:1.65;color:#5C5653;">' + escapeHtml_(intro) + '</p>' +
          '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">' + items + '</table>' +
          '<p style="margin:22px 0 0;font-size:14px;line-height:1.65;color:#5C5653;">Ссылки не истекают — можно вернуться к ним в любое время. Если что-то не открылось, просто ответьте на это письмо.</p>' +
          '<p style="margin:26px 0 0;font-size:14px;line-height:1.7;color:#5C5653;">Баракаллаху фикум,<br>' +
            '<span style="color:#1C1A19;">' + escapeHtml_(MAIL_FROM_NAME) + '</span><br>' +
            '<a href="' + SITE_URL + '" style="color:#B26E7A;text-decoration:none;">' + SITE_URL.replace('https://', '') + '</a><br>' +
            '<a href="mailto:' + MAIL_FROM + '" style="color:#B26E7A;text-decoration:none;">' + escapeHtml_(MAIL_FROM) + '</a>' +
          '</p>' +
        '</td></tr>' +
      '</table>' +
    '</div>';
}

/** Экранирование текста, который подставляется в HTML письма. */
function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Проверка письма без заполнения формы.
 * Выберите эту функцию в редакторе Apps Script и нажмите «Выполнить» —
 * письмо придёт на вашу же почту.
 */
function testMail() {
  var me = Session.getEffectiveUser().getEmail();
  Logger.log('Гайды: ' + sendMagnetEmail_(me, 'Проверка', 'Гайды по планированию'));
  Logger.log('Книга: ' + sendMagnetEmail_(me, 'Проверка', 'Справочник постящейся'));
}

// Чтобы открытие URL в браузере не выдавало ошибку.
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, info: 'Sestram leads endpoint', version: 'v2-dedup' }))
    .setMimeType(ContentService.MimeType.JSON);
}
