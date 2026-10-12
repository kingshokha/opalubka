/**
 * Учёт аренды опалубки — хранилище данных в этой Google Таблице.
 *
 * Установка (подробно — в README репозитория и в приложении: «Ещё» → «Google Таблица»):
 *  1. Сохраните (Ctrl+S).
 *  2. «Развернуть» (Deploy) → «Новое развертывание» → тип «Веб-приложение»:
 *     «Запуск от имени» — Я, «У кого есть доступ» — Все. Разрешите доступ.
 *  3. Скопируйте URL веб-приложения (заканчивается на /exec). Его вписывают в код
 *     приложения (BUILTIN_SHEET в index.html) — тогда все устройства подключаются сами.
 *
 * Листы «Каталог», «Сделки», «Возвраты», «История» перезаписываются при каждом
 * сохранении: правьте данные в приложении, а не здесь. Сами данные лежат
 * в скрытом листе «_data» — не трогайте его.
 *
 * Пароля нет: читать и менять данные может любой, у кого есть адрес веб-приложения.
 */

const DATA_SHEET = '_data';
const CHUNK = 40000;     // в ячейке не больше 50 000 символов
const REPORT_SHEETS = ['Каталог', 'Сделки', 'Возвраты', 'История'];
const MAX_COL_WIDTH = 420;

/* Запустите один раз из редактора (выберите authorize → «Выполнить») и разрешите доступ —
   без этого веб-приложение не сможет работать с таблицей. */
function authorize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Скрипт должен быть создан из таблицы: «Расширения» → «Apps Script».');
  Logger.log('Доступ к таблице «' + ss.getName() + '» разрешён.');
}

function doGet() {
  return ContentService.createTextOutput(
    'Скрипт учёта аренды опалубки работает.');
}

/* Приложение шлёт POST с телом-JSON в формате text/plain: так браузер не делает
   предварительный CORS-запрос, который Apps Script не умеет обрабатывать. */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return reply({ ok: false, error: 'bad-request' }); }
  if (!req) return reply({ ok: false, error: 'bad-request' });

  // одновременно данные меняет только один запрос
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return reply({ ok: false, error: 'busy' });
  try {
    const props = PropertiesService.getScriptProperties();
    const version = Number(props.getProperty('version') || 0);

    if (req.action === 'get') {
      if (req.since === version) return reply({ ok: true, version: version, unchanged: true });
      return reply({ ok: true, version: version, state: readState(props) });
    }

    if (req.action === 'save') {
      const s = req.state;
      if (!s || !Array.isArray(s.items) || !Array.isArray(s.rentals)) return reply({ ok: false, error: 'bad-state' });
      // кто-то сохранил раньше нас — отдаём свежие данные, приложение сольёт и повторит
      if (Number(req.baseVersion) !== version)
        return reply({ ok: false, error: 'conflict', version: version, state: readState(props) });

      const w = writeState(props, s);
      const next = version + 1;
      const meta = { version: String(next), col: String(w.col), updatedAt: new Date().toISOString(),
                     updatedBy: String(req.by || '') };
      meta['n' + w.col] = String(w.n);
      props.setProperties(meta);   // переключение на новые данные — одной записью

      let reportError = null;
      try { writeReport(props, req.report || {}); } catch (err) { reportError = String(err); }
      return reply({ ok: true, version: next, reportError: reportError });
    }

    return reply({ ok: false, error: 'bad-action' });
  } finally {
    lock.releaseLock();
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- данные: JSON по кускам в скрытом листе ----------
   Две колонки по очереди: новая версия пишется в ту, что сейчас не используется,
   и только потом на неё переключается свойство col. Если запись оборвётся на
   середине, рабочей остаётся прежняя версия. */

function dataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    try { sh.hideSheet(); } catch (err) { /* единственный видимый лист скрыть нельзя — не страшно */ }
  }
  return sh;
}

function readState(props) {
  const col = Number(props.getProperty('col') || 0);
  const n = Number(props.getProperty('n' + col) || 0);
  if (!col || !n) return null;
  const values = dataSheet().getRange(1, col, n, 1).getValues();
  // у каждого куска впереди «~», чтобы таблица не приняла его за число, дату или формулу
  return JSON.parse(values.map(function (r) { return String(r[0]).slice(1); }).join(''));
}

function writeState(props, state) {
  const json = JSON.stringify(state);
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK) chunks.push(['~' + json.slice(i, i + CHUNK)]);
  if (!chunks.length) chunks.push(['~']);

  const cur = Number(props.getProperty('col') || 0);
  const col = cur === 1 ? 2 : 1;
  const sh = dataSheet();
  if (sh.getMaxRows() < chunks.length) sh.insertRowsAfter(sh.getMaxRows(), chunks.length - sh.getMaxRows());
  sh.getRange(1, col, sh.getMaxRows(), 1).clearContent();
  const range = sh.getRange(1, col, chunks.length, 1);
  range.setNumberFormat('@');
  range.setValues(chunks);
  return { col: col, n: chunks.length };
}

/* ---------- листы для просмотра ----------
   Строки целиком готовит приложение (у него вся логика расчётов), здесь — только запись.
   Неизменившиеся листы не трогаем: сравниваем отпечаток содержимого. */

function writeReport(props, report) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  REPORT_SHEETS.forEach(function (name) {
    const rows = report[name];
    if (!Array.isArray(rows) || !rows.length) return;
    const hash = hashOf(JSON.stringify(rows));
    let sh = ss.getSheetByName(name);
    if (sh && props.getProperty('h_' + name) === hash) return;
    if (!sh) sh = ss.insertSheet(name);

    const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    const dateCols = {}, timeCols = {};
    const values = rows.map(function (r, i) {
      const out = [];
      for (let c = 0; c < width; c++) out.push(toCell(r[c], i, c, dateCols, timeCols));
      return out;
    });

    sh.clear();
    sh.getRange(1, 1, values.length, width).setValues(values);
    sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#eef0f3');
    sh.setFrozenRows(1);
    if (values.length > 1) {
      const body = values.length - 1;
      Object.keys(dateCols).forEach(function (c) { sh.getRange(2, Number(c) + 1, body, 1).setNumberFormat('dd.mm.yyyy'); });
      Object.keys(timeCols).forEach(function (c) { sh.getRange(2, Number(c) + 1, body, 1).setNumberFormat('dd.mm.yyyy hh:mm'); });
      sh.getRange(2, 1, body, width).setVerticalAlignment('top');
    }
    sh.autoResizeColumns(1, width);
    for (let c = 1; c <= width; c++) {
      if (sh.getColumnWidth(c) > MAX_COL_WIDTH) sh.setColumnWidth(c, MAX_COL_WIDTH);
    }
    if (values.length > 1) sh.getRange(2, 1, values.length - 1, width).setWrap(true);
    props.setProperty('h_' + name, hash);
  });

  // пустой лист, созданный вместе с таблицей, больше не нужен
  ['Лист1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}

/* «2026-10-12» → дата, «2026-10-12 14:30» → дата и время, остальной текст — как есть:
   апостроф впереди не даёт таблице сделать из «+992…» число, а из «=…» формулу. */
function toCell(v, row, col, dateCols, timeCols) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  const s = String(v);
  if (s === '') return '';
  if (row > 0) {
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) { dateCols[col] = 1; return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])); }
    m = s.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/);
    if (m) { timeCols[col] = 1; return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5])); }
  }
  return "'" + s;
}

function hashOf(s) {   // FNV-1a, 32 бита — только чтобы заметить изменения
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
  return h.toString(16) + ':' + s.length;
}
