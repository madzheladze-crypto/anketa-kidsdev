/**
 * Запись на консультации МГУ.
 * Слоты читаются из таблицы «Профтест» (лист «Расписание консультаций») — только чтение.
 * Брони пишутся в ЭТУ таблицу (к которой привязан скрипт), лист «Брони».
 * Слот свободен, если в «Профтест» не заполнено «ФИ ребенка» и в «Бронях» нет записи на это время.
 *
 * Установка: в таблице для броней «Расширения → Apps Script», вставить этот код,
 * «Начать развертывание → Новое развертывание → Веб-приложение»,
 * запуск от имени «Я», доступ «Все». Адрес /exec вставить в zapis-mgu.html (SCRIPT_URL).
 */
const SOURCE_ID = '1k1NkxGE4ASwA06DNS5eL3G9g3--nQi61pF-2jfyqkn0'; // «Профтест»
const SOURCE_SHEET = 'Расписание консультаций';
const BOOK_SHEET = 'Брони';
const HEAD = ['Когда записались', 'Консультант', 'Дата', 'Время (мск)', 'ФИ ребёнка', 'Родитель', 'Контакт', 'ID слота'];

function readSlots_() {
  const sh = SpreadsheetApp.openById(SOURCE_ID).getSheetByName(SOURCE_SHEET);
  if (!sh) throw new Error('Нет листа «' + SOURCE_SHEET + '»');
  const v = sh.getDataRange().getDisplayValues();
  const slots = [];
  let lastText = '', cons = '', cols = null, date = '';
  for (let r = 0; r < v.length; r++) {
    const row = v[r].map(c => String(c).trim());
    const isHeader = row.some(c => /^Дата/i.test(c)) && row.some(c => /^Время/i.test(c));
    if (isHeader) {
      cols = {
        d: row.findIndex(c => /^Дата/i.test(c)),
        t: row.findIndex(c => /^Время/i.test(c)),
        fi: row.findIndex(c => /^ФИ/i.test(c))
      };
      cons = lastText.replace(/[\w.+-]+@[\w-]+\.[\w.]+/g, '').replace(/^Консультант\s*\d*\s*[:.]?\s*/i, '').replace(/[,;\s]+$/, '').trim();
      date = '';
      continue;
    }
    const time = cols ? row[cols.t] : '';
    if (cols && /^\d{1,2}[:.]\d{2}$/.test(time)) {
      if (row[cols.d]) date = row[cols.d];
      if (!date) continue;
      slots.push({ id: cons + '|' + date + '|' + time, consultant: cons, date: date, time: time,
                   free: cols.fi < 0 || !row[cols.fi] });
      continue;
    }
    if (cols && /^\d{1,2}\.\d{1,2}\.\d{2,4}/.test(row[cols.d])) { date = row[cols.d]; continue; }
    const text = row.filter(Boolean).join(', ');
    if (text) { lastText = text; cols = null; }
  }
  return slots;
}

function bookSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(BOOK_SHEET);
  if (!sh) {
    const first = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && first.getLastRow() === 0) { sh = first; sh.setName(BOOK_SHEET); }
    else sh = ss.insertSheet(BOOK_SHEET);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function bookedIds_() {
  const sh = bookSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return new Set();
  return new Set(sh.getRange(2, HEAD.length, n, 1).getValues().map(r => String(r[0])).filter(Boolean));
}

function slotsWithBookings_() {
  const booked = bookedIds_();
  return readSlots_().map(s => Object.assign(s, { free: s.free && !booked.has(s.id) }));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    return json_({ ok: true, slots: slotsWithBookings_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const p = e.parameter;
    const child = (p.child || '').trim(), parent = (p.parent || '').trim(), contact = (p.contact || '').trim();
    if (!child || !contact || !p.slot) return json_({ ok: false, error: 'fields' });
    const slot = slotsWithBookings_().find(s => s.id === p.slot);
    if (!slot) return json_({ ok: false, error: 'notfound' });
    if (!slot.free) return json_({ ok: false, error: 'taken' });
    bookSheet_().appendRow([new Date(), slot.consultant, slot.date, slot.time, child, parent, contact, slot.id]);
    return json_({ ok: true, consultant: slot.consultant, date: slot.date, time: slot.time });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
