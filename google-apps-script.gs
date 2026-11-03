/**
 * Сервер записей для сайта «Кагами».
 * Хранит записи в Google Таблице, в которой установлен этот скрипт.
 *
 * Как установить — README.md, раздел «Подключение Google Таблицы».
 *
 * ВАЖНО: замените пароль ниже на свой. Его вводят на странице admin.html.
 */
const ADMIN_PASSWORD = 'ЗАМЕНИТЕ-НА-СВОЙ-ПАРОЛЬ';

const SHEET_NAME = 'Записи';
const HEADERS = ['id', 'ts', 'createdAt', 'name', 'phone', 'services', 'master', 'date', 'time', 'total', 'status'];
const STATUSES = ['new', 'confirmed', 'done', 'cancelled'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:K').setNumberFormat('@'); // всё храним как текст
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Защита от формул в ячейках и обрезка слишком длинного текста
function clean_(value, max) {
  let s = String(value == null ? '' : value).slice(0, max || 300);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function doGet() {
  return json_({ ok: true, service: 'kagami-bookings' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();

    // Новая запись с сайта — пароль не нужен
    if (req.action === 'add') {
      const b = req.booking || {};
      const services = Array.isArray(b.services) ? b.services.join(', ') : b.services;
      if (!b.name || !b.phone) return json_({ ok: false, error: 'missing_fields' });
      sh.appendRow([
        clean_(b.id, 40), clean_(b.ts, 20), clean_(b.createdAt, 40), clean_(b.name, 100), clean_(b.phone, 30),
        clean_(services, 1000), clean_(b.master, 60), clean_(b.date, 20), clean_(b.time, 10),
        clean_(b.total, 12), 'new'
      ]);
      return json_({ ok: true });
    }

    // Всё остальное — только с паролем администратора
    if (req.password !== ADMIN_PASSWORD) return json_({ ok: false, error: 'auth' });

    if (req.action === 'login') return json_({ ok: true });

    if (req.action === 'list') {
      const last = sh.getLastRow();
      const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, HEADERS.length).getValues().map(r => {
        const o = {};
        HEADERS.forEach((h, i) => { o[h] = String(r[i]).replace(/^'/, ''); });
        return o;
      });
      return json_({ ok: true, rows });
    }

    if (req.action === 'status') {
      if (STATUSES.indexOf(req.status) === -1) return json_({ ok: false, error: 'bad_status' });
      const row = findRow_(sh, req.id);
      if (row < 0) return json_({ ok: false, error: 'not_found' });
      sh.getRange(row, HEADERS.indexOf('status') + 1).setValue(req.status);
      return json_({ ok: true });
    }

    if (req.action === 'delete') {
      const row = findRow_(sh, req.id);
      if (row < 0) return json_({ ok: false, error: 'not_found' });
      sh.deleteRow(row);
      return json_({ ok: true });
    }

    if (req.action === 'clear') {
      const last = sh.getLastRow();
      if (last > 1) sh.deleteRows(2, last - 1);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
