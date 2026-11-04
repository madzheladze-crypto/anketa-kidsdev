/**
 * Запись на консультации МГУ — скрипт для таблицы «Профтест».
 * Установка: отдельный проект на script.new (или «Расширения → Apps Script» в таблице),
 * «Начать развертывание → Новое развертывание → Веб-приложение»,
 * запуск от имени «Я», доступ «Все». Адрес /exec вставить в zapis-mgu.html (SCRIPT_URL).
 *
 * Скрипт читает лист «Расписание консультаций»: строка с ФИО консультанта,
 * под ней шапка «Дата | Время (мск) | ФИ ребенка | Контакты родителя | …»,
 * дальше строки слотов (дата пишется в первой строке дня).
 * Слот свободен, если ячейка «ФИ ребенка» пустая.
 */
const SPREADSHEET_ID = '1k1NkxGE4ASwA06DNS5eL3G9g3--nQi61pF-2jfyqkn0'; // таблица «Профтест»
const SHEET_NAME = 'Расписание консультаций';
const book_ = () => SpreadsheetApp.openById(SPREADSHEET_ID);
const LOG_SHEET = 'Записи (журнал)';

function readSlots_() {
  const sh = book_().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Нет листа «' + SHEET_NAME + '»');
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
        fi: row.findIndex(c => /^ФИ/i.test(c)),
        ct: row.findIndex(c => /^Контакт/i.test(c))
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
                   free: !row[cols.fi], row: r + 1, fi: cols.fi + 1, ct: cols.ct + 1 });
      continue;
    }
    if (cols && /^\d{1,2}\.\d{1,2}\.\d{2,4}/.test(row[cols.d])) { date = row[cols.d]; continue; }
    const text = row.filter(Boolean).join(', ');
    if (text) { lastText = text; cols = null; }
  }
  return slots;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try {
    const slots = readSlots_().map(s => ({ id: s.id, consultant: s.consultant, date: s.date, time: s.time, free: s.free }));
    return json_({ ok: true, slots: slots });
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
    const slot = readSlots_().find(s => s.id === p.slot);
    if (!slot) return json_({ ok: false, error: 'notfound' });
    if (!slot.free) return json_({ ok: false, error: 'taken' });
    const sh = book_().getSheetByName(SHEET_NAME);
    sh.getRange(slot.row, slot.fi).setValue(child);
    if (slot.ct > 0) sh.getRange(slot.row, slot.ct).setValue([parent, contact].filter(Boolean).join(', '));
    const ss = book_();
    const log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);
    if (log.getLastRow() === 0) {
      log.appendRow(['Когда записались', 'Консультант', 'Дата', 'Время (мск)', 'ФИ ребёнка', 'Родитель', 'Контакт']);
      log.getRange(1, 1, 1, 7).setFontWeight('bold');
      log.setFrozenRows(1);
    }
    log.appendRow([new Date(), slot.consultant, slot.date, slot.time, child, parent, contact]);
    return json_({ ok: true, consultant: slot.consultant, date: slot.date, time: slot.time });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
