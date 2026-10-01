/** =====================================================================
 *  OZON MANAGER — ВЕСЬ КОД В ОДНОМ ФАЙЛЕ
 *  В проекте Apps Script должен остаться ТОЛЬКО этот файл.
 * ===================================================================== */



/* ================== 01_core.gs ================== */
/** =====================================================================
 *  CORE: ключи, настройки, клиент Ozon API, утилиты листов, лог
 *  Ключи API хранятся в свойствах скрипта, в коде их нет.
 * ===================================================================== */
var OZON_HOST = 'https://api-seller.ozon.ru';
var SHEETS = {
  MAIN: 'Ozon', BOG: 'Бог акций', SETTINGS: 'Настройки', SOURCES: 'Источники закупа',
  TARIFFS: 'Тарифы', DISCOUNTS: 'Заявки на скидку', LOG: 'Лог'
};
var MAIN_HDR_ROW = 2, MAIN_FIRST = 3;          // лист Ozon: строка 1 — группы колонок, 2 — заголовки
var NEW_BLOCK = '🆕 Новые товары';

/* ---------- Ключи ---------- */
function getCreds_() {
  const p = PropertiesService.getScriptProperties();
  const clientId = p.getProperty('OZON_CLIENT_ID'), apiKey = p.getProperty('OZON_API_KEY');
  if (!clientId || !apiKey) throw new Error('Не заданы ключи API: ⚙ НАСТРОЙКИ → Ключи API');
  return { clientId, apiKey };
}
function setupCredentials() {
  const ui = SpreadsheetApp.getUi();
  const c = ui.prompt('Ozon: Client-Id', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  const k = ui.prompt('Ozon: Api-Key (новый, после перевыпуска!)', ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperties({
    OZON_CLIENT_ID: c.getResponseText().trim(), OZON_API_KEY: k.getResponseText().trim()
  });
  ui.alert('Ключи сохранены в свойствах скрипта.');
}

/* ---------- Настройки ---------- */
let SETTINGS_CACHE_ = null;
function cfg_(key, fallback) {
  if (!SETTINGS_CACHE_) {
    SETTINGS_CACHE_ = {};
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
    if (sh && sh.getLastRow() > 1)
      sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, v]) => { if (k) SETTINGS_CACHE_[String(k).trim()] = v; });
  }
  const v = SETTINGS_CACHE_[key];
  return (v === '' || v === undefined || v === null) ? fallback : v;
}
function setCfg_(key, value) {
  const sh = sheet_(SHEETS.SETTINGS);
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim());
  const i = keys.indexOf(key);
  if (i < 0) throw new Error('В «Настройках» нет ключа ' + key);
  sh.getRange(i + 2, 2).setValue(value);
  SETTINGS_CACHE_ = null;
}
function isDryRun_() { return cfg_('DRY_RUN', true) !== false; }
/** Доля из «Настроек» для подписей и сообщений: 0.12 → «12%» */
function pctCfg_(key, fallback) {
  let v = Number(fallback);
  try { const x = Number(cfg_(key, fallback)); if (isFinite(x)) v = x; } catch (e) {}
  return Math.round(v * 1000) / 10 + '%';
}

/* ---------- Ozon API ---------- */
function ozon_(path, body, method) {
  const { clientId, apiKey } = getCreds_();
  const m = (method || 'post').toLowerCase();
  const opts = { method: m, contentType: 'application/json', muteHttpExceptions: true,
                 headers: { 'Client-Id': String(clientId), 'Api-Key': apiKey } };
  if (m !== 'get' && body !== undefined) opts.payload = JSON.stringify(body);
  for (let attempt = 1; attempt <= 5; attempt++) {
    let resp;
    try { resp = UrlFetchApp.fetch(OZON_HOST + path, opts); }
    catch (e) {                                   // сбой сети («Адрес недоступен») — повторяем, как при 5xx
      if (attempt === 5) throw e;
      Utilities.sleep(1000 * attempt * attempt); continue;
    }
    const code = resp.getResponseCode(), text = resp.getContentText();
    if (code === 200) return JSON.parse(text);
    if (code === 429 || code >= 500) { Utilities.sleep(1000 * attempt * attempt); continue; }
    throw new Error(`Ozon ${path} → HTTP ${code}: ${text.slice(0, 500)}`);
  }
  throw new Error(`Ozon ${path}: не удалось после 5 попыток`);
}
/** Все страницы: pick(resp) → {items, next}; tokenKey — 'cursor' или 'last_id' */
function ozonAll_(path, body, pick, tokenKey) {
  const out = []; let token = '';
  for (let page = 0; page < 1000; page++) {
    const b = Object.assign({}, body); if (token) b[tokenKey] = token;
    const { items, next } = pick(ozon_(path, b));
    if (!items || !items.length) break;
    for (const it of items) out.push(it);
    if (!next || next === token) break;
    token = next;
  }
  return out;
}

/* ---------- Листы (заголовок в строке 1) ---------- */
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Нет листа: ' + name);
  return sh;
}
function headersAt_(sh, row) { return sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim()); }
function col_(h, name) { const i = h.indexOf(name); if (i < 0) throw new Error(`Нет колонки «${name}»`); return i; }

function readTable_(name) {
  const sh = sheet_(name), h = headersAt_(sh, 1), last = sh.getLastRow();
  if (last < 2) return { sh, h, rows: [] };
  const rows = sh.getRange(2, 1, last - 1, h.length).getValues()
    .map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; })
    .filter(o => o[h[0]] !== '' && o[h[0]] !== null);
  return { sh, h, rows };
}
/** Формулы-шаблон строки (R1C1), кешируются в свойствах документа */
function templateRow_(sh, h, row, key) {
  const props = PropertiesService.getDocumentProperties();
  if (row) {
    const f = sh.getRange(row, 1, 1, h.length).getFormulasR1C1()[0];
    if (f.some(Boolean)) { props.setProperty(key, JSON.stringify(f)); return f; }
  }
  const saved = props.getProperty(key);
  return saved ? JSON.parse(saved) : h.map(() => '');
}
/** Перезаписывает лист (заголовок в строке 1). Колонки-формулы заполняются формулой из строки 2. */
function writeTable_(name, objs) {
  const sh = sheet_(name), h = headersAt_(sh, 1);
  const tpl = templateRow_(sh, h, sh.getLastRow() >= 2 ? 2 : 0, 'TPL_' + name);
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, h.length).clearContent();
  if (!objs.length) return;
  h.forEach((key, j) => {
    const range = sh.getRange(2, j + 1, objs.length, 1);
    if (tpl[j]) range.setFormulaR1C1(tpl[j]);
    else range.setValues(objs.map(o => [o[key] === undefined || o[key] === null ? '' : o[key]]));
  });
}
/** Точечная запись в одну колонку: patch = {номерСтроки: значение}. Только для колонок без формул! */
function patchColumn_(sh, h, name, patch, firstRow) {
  if (!Object.keys(patch).length) return;
  firstRow = firstRow || 2;
  const j = col_(h, name) + 1, last = sh.getLastRow();
  if (last < firstRow) return;
  const range = sh.getRange(firstRow, j, last - firstRow + 1, 1), v = range.getValues();
  Object.keys(patch).forEach(row => { v[Number(row) - firstRow][0] = patch[row]; });
  range.setValues(v);
}

/* ---------- Формулы: разделитель аргументов зависит от локали таблицы ---------- */
var FSEP_ = null;
/** Пишет тестовую формулу и смотрит, какой разделитель понимает таблица */
function fsep_() {
  if (FSEP_) return FSEP_;
  const cached = PropertiesService.getDocumentProperties().getProperty('FSEP');
  if (cached) { FSEP_ = cached; return FSEP_; }
  const sh = sheet_(SHEETS.SETTINGS);
  const cell = sh.getRange(1, 9);                       // служебная ячейка I1 на «Настройках»
  let sep = ',';
  try {
    cell.setFormula('=IF(1=1,"OK","NO")');
    SpreadsheetApp.flush();
    if (String(cell.getDisplayValue()).indexOf('OK') < 0) {
      cell.setFormula('=IF(1=1;"OK";"NO")');
      SpreadsheetApp.flush();
      if (String(cell.getDisplayValue()).indexOf('OK') >= 0) sep = ';';
    }
  } catch (e) {} finally { try { cell.clearContent(); } catch (e) {} }
  PropertiesService.getDocumentProperties().setProperty('FSEP', sep);
  log_('Формулы', 'INFO', 'Разделитель аргументов: «' + sep + '»');
  FSEP_ = sep;
  return sep;
}
/** Приводит формулу к разделителю таблицы (текст в кавычках не трогаем) */
function fx_(formula) {
  if (fsep_() === ',') return formula;
  let out = '', inStr = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') inStr = !inStr;
    out += (!inStr && ch === ',') ? ';' : ch;
  }
  return out;
}

/* ---------- Мелочи ---------- */
function chunk_(arr, n) { const o = []; for (let i = 0; i < arr.length; i += n) o.push(arr.slice(i, i + n)); return o; }
function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[\s\u00a0₽]/g, '').replace(',', '.'));
  return isFinite(n) ? n : '';
}
function key_(v) { return String(v === null || v === undefined ? '' : v).trim().replace(/\.0+$/, '').toLowerCase(); }
function date_(v) { if (!v) return ''; const d = new Date(v); return isNaN(d) ? '' : d; }
function toast_(msg, sec) { try { SpreadsheetApp.getActive().toast(msg, 'Ozon', sec || 5); } catch (e) {} }

/* ---------- Лог и запуск ---------- */
function log_(op, level, msg) {
  console.log(`[${level}] ${op}: ${msg}`);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.LOG); if (!sh) return;
    sh.insertRowAfter(1);
    sh.getRange(2, 1, 1, 4).setValues([[new Date(), op, level, String(msg).slice(0, 45000)]]);
    if (sh.getLastRow() > 3000) sh.deleteRows(2001, sh.getLastRow() - 2000);
  } catch (e) {}
}
function run_(title, fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) { toast_('Уже выполняется другая операция — подождите'); return; }
  const t0 = Date.now();
  try {
    toast_(title + '…', 3);
    const msg = fn() || 'готово';
    log_(title, 'OK', msg);
    toast_(`${title}: ${msg} (${Math.round((Date.now() - t0) / 1000)} с)`, 10);
  } catch (e) {
    log_(title, 'ERROR', e.stack || e);
    toast_(`${title}: ОШИБКА — ${e.message}`, 15);
    throw e;
  } finally { lock.releaseLock(); }
}


/* ================== 02_main.gs ================== */
/** =====================================================================
 *  ГЛАВНЫЙ ЛИСТ «Ozon»: блоки по категориям, добавление товаров,
 *  тарифы, остатки, закуп, заказы
 * ===================================================================== */
function addMissingProducts() { run_('Добавить новые товары', addMissingProducts_); }
function regroupByCategory()  { run_('Разложить по блокам', () => regroup_([])); }
function syncTariffs()        { run_('Тарифы и цены', syncTariffs_); }
function syncStocks()         { run_('Остатки FBS', syncStocks_); }
function importCosts()        { run_('Закуп', importCosts_); }
function syncOrders60()       { run_('Заказы', syncOrders60_); }
function updateUsdRate()      { run_('Курс USD', updateUsdRate_); }

var isBlockRow_ = v => String(v || '').startsWith('▌');

/** Товарные строки листа Ozon (строки-заголовки блоков пропускаются) */
function readMain_() {
  const sh = sheet_(SHEETS.MAIN), h = headersAt_(sh, MAIN_HDR_ROW), last = sh.getLastRow(), rows = [];
  if (last >= MAIN_FIRST) {
    sh.getRange(MAIN_FIRST, 1, last - MAIN_FIRST + 1, h.length).getValues().forEach((r, i) => {
      if (!r[0] || isBlockRow_(r[0])) return;
      const o = { _row: i + MAIN_FIRST }; h.forEach((k, j) => o[k] = r[j]); rows.push(o);
    });
  }
  return { sh, h, rows };
}
function mainTemplate_(m) { return templateRow_(m.sh, m.h, m.rows.length ? m.rows[0]._row : 0, 'TPL_MAIN'); }
function mainPatch_(m, name, patch) { patchColumn_(m.sh, m.h, name, patch, MAIN_FIRST); }

/**
 * Пересобирает лист по блокам-категориям. extra — новые товары (объекты по заголовкам).
 * Ручные значения (цена продажи, категория, коды) сохраняются, формулы ставятся заново.
 */
function regroup_(extra) {
  const m = readMain_(), sh = m.sh, h = m.h, W = h.length;
  const tpl = mainTemplate_(m);
  if (!tpl.some(Boolean)) throw new Error('Не найдены формулы-шаблон на листе Ozon');
  const fmtRow = m.rows.length ? m.rows[0]._row : null;
  const all = m.rows.concat(extra || []);

  const groups = {};
  all.forEach(r => { const c = String(r['Категория'] || '').trim() || NEW_BLOCK; (groups[c] = groups[c] || []).push(r); });
  const order = Object.keys(groups).filter(c => c !== NEW_BLOCK)
    .sort((a, b) => groups[b].length - groups[a].length || a.localeCompare(b, 'ru'));
  if (groups[NEW_BLOCK]) order.push(NEW_BLOCK);          // новые товары — в самом низу

  const grid = [], blocks = [];
  let r = MAIN_FIRST;
  order.forEach(cat => {
    const items = groups[cat].sort((a, b) => String(a['Название'] || a['Артикул']).localeCompare(String(b['Название'] || b['Артикул']), 'ru'));
    blocks.push({ cat, hr: r, first: r + 1, last: r + items.length, n: items.length });
    grid.push(null); items.forEach(it => grid.push(it));
    r += items.length + 1;
  });
  const n = grid.length, lastRow = MAIN_FIRST + n - 1;

  // очистка старого содержимого и групп
  const oldLast = sh.getLastRow();
  if (oldLast >= MAIN_FIRST) {
    const rg = sh.getRange(MAIN_FIRST, 1, oldLast - MAIN_FIRST + 1, W);
    for (let k = 0; k < 3; k++) { try { rg.shiftRowGroupDepth(-1); } catch (e) { break; } }
    rg.clearContent();
  }
  if (sh.getMaxRows() < lastRow) sh.insertRowsAfter(sh.getMaxRows(), lastRow - sh.getMaxRows());

  const c = name => col_(h, name) + 1;
  // в строке категории — только её название; средняя маржа проставится значением после расчёта
  const hdr = {};
  blocks.forEach(b => { hdr[b.hr] = { 'Артикул': `▌ ${b.cat}` }; });

  h.forEach((name, j) => {
    const colVals = grid.map((it, i) => {
      const row = MAIN_FIRST + i;
      if (!it) { const v = (hdr[row] || {})[name]; return [v === undefined ? '' : v]; }
      if (tpl[j]) return [tpl[j]];
      const v = it[name]; return [v === undefined || v === null ? '' : v];
    });
    const rg = sh.getRange(MAIN_FIRST, j + 1, n, 1);
    if (!tpl[j]) { rg.setValues(colVals); return; }
    rg.setFormulasR1C1(colVals);
    // в строках категорий формул быть не должно
    blocks.forEach(b => sh.getRange(b.hr, j + 1).clearContent());
  });

  // оформление: товарные строки — как прежняя товарная строка, заголовки блоков — тёмные
  if (fmtRow) sh.getRange(fmtRow, 1, 1, W).copyFormatToRange(sh, 1, W, MAIN_FIRST, lastRow);
  const checkCols = ['Отправить', 'Цена вручную'].filter(x => h.indexOf(x) >= 0).map(x => c(x));
  checkCols.forEach(cc => sh.getRange(MAIN_FIRST, cc, n, 1).clearDataValidations());
  blocks.forEach(b => {
    sh.getRange(b.hr, 1, 1, W).setBackground('#C7D7EA').setFontColor('#0B2545').setFontWeight('bold');
    sh.getRange(b.first, 1, b.n, 1).shiftRowGroupDepth(1);
    checkCols.forEach(cc => sh.getRange(b.first, cc, b.n, 1).insertCheckboxes());
  });
  sh.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  if (oldLast > lastRow) sh.getRange(lastRow + 1, 1, oldLast - lastRow, W).clear();

  // средняя маржа категории — значением, без формул
  SpreadsheetApp.flush();
  const mgCol = c('Маржа, %'), stCol = c('Остаток FBS');
  blocks.forEach(b => {
    const mg = sh.getRange(b.first, mgCol, b.n, 1).getValues();
    const st = sh.getRange(b.first, stCol, b.n, 1).getValues();
    const vals = mg.map((r, i) => (Number(st[i][0]) > 0 ? r[0] : ''))    // без остатка в среднюю маржу не берём
      .filter(v => typeof v === 'number' && isFinite(v));
    const avg = vals.length ? vals.reduce((x, v) => x + v, 0) / vals.length : '';
    sh.getRange(b.hr, mgCol).setValue(avg).setNumberFormat('0.0%');
  });

  // условное форматирование и фильтр — на весь новый диапазон
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().map(rule => rule.copy().setRanges(
    rule.getRanges().map(x => sh.getRange(MAIN_FIRST, x.getColumn(), n, x.getNumColumns()))).build()));
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(MAIN_HDR_ROW, 1, lastRow - MAIN_HDR_ROW + 1, W).createFilter();
  return `блоков: ${blocks.length}, товаров: ${all.length}`;
}

/* ---------- Добавить товары, которых нет в таблице — одна кнопка ---------- */
function addMissingProducts_() {
  const m = readMain_();
  const haveArt = new Set(m.rows.map(r => key_(r['Артикул']))), havePid = new Set(m.rows.map(r => key_(r['Product ID'])));
  const list = ozonAll_('/v3/product/list', { filter: { visibility: 'ALL' }, limit: 1000 },
    r => ({ items: r.result.items, next: r.result.last_id }), 'last_id');
  const fresh = list.filter(x => !x.archived && !haveArt.has(key_(x.offer_id)) && !havePid.has(key_(x.product_id)));
  if (!fresh.length) return 'новых товаров нет — таблица совпадает с кабинетом';

  const info = {};
  chunk_(fresh.map(x => x.product_id), 1000).forEach(part =>
    (ozon_('/v3/product/info/list', { product_id: part }).items || []).forEach(it => info[it.id] = it));
  const extra = fresh.map(x => {
    const it = info[x.product_id] || {};
    return { 'Артикул': String(x.offer_id), 'Product ID': x.product_id, 'Название': it.name || '',
      'SKU': it.sku || (it.sources && it.sources[0] && it.sources[0].sku) || '', 'Категория': '', 'Отправить': false };
  });
  regroup_(extra);
  syncTariffs_(); syncStocks_(); importCosts_();
  SpreadsheetApp.flush();

  // цена = «Мин. цена» (маржа MIN_MARGIN), отметка «Отправить»
  const m2 = readMain_(), added = new Set(extra.map(e => key_(e['Артикул'])));
  const price = {}, send = {}, res = {};
  m2.rows.forEach(r => {
    if (!added.has(key_(r['Артикул']))) return;
    const mp = Number(r['Мин. цена, ₽']);
    if (mp > 0) { price[r._row] = mp; send[r._row] = true; res[r._row] = '🆕 цена по мин. марже — проверьте и отправьте'; }
    else res[r._row] = '🆕 нет закупа: укажите «Код в прайсе» и «Источник закупа»';
  });
  mainPatch_(m2, 'Цена продажи, ₽', price); mainPatch_(m2, 'Отправить', send); mainPatch_(m2, 'Результат', res);
  return `добавлено: ${fresh.length}, цена назначена: ${Object.keys(price).length}, без закупа: ${fresh.length - Object.keys(price).length}`;
}

/* ---------- Тарифы, текущие цены, конкуренты ---------- */
function syncTariffs_() {
  const items = ozonAll_('/v5/product/info/prices', { filter: { visibility: 'ALL' }, limit: 1000 },
    r => ({ items: r.items, next: r.cursor }), 'cursor');
  const minOf = d => d ? num_(d.minimal_price !== undefined ? d.minimal_price : d.min_price) : '';
  const now = new Date();
  writeTable_(SHEETS.TARIFFS, items.map(o => {
    const p = o.price || {}, c = o.commissions || {}, ix = o.price_indexes || {};
    return {
      'product_id': o.product_id, 'Артикул': String(o.offer_id),
      'Цена': num_(p.price), 'Цена до скидки': num_(p.old_price), 'Мин. цена': num_(p.min_price),
      'Комиссия FBS, %': c.sales_percent_fbs,
      'Логистика FBS мин, ₽': c.fbs_direct_flow_trans_min_amount, 'Логистика FBS макс, ₽': c.fbs_direct_flow_trans_max_amount,
      'Обработка FBS, ₽': c.fbs_first_mile_max_amount, 'Последняя миля FBS, ₽': c.fbs_deliv_to_customer_amount,
      'Мин. цена конкурента на Ozon': minOf(ix.ozon_index_data),
      'Мин. цена на других площадках': minOf(ix.external_index_data),
      'Индекс цены': ix.color_index || ix.price_index || '', 'Обновлено': now,
      'Обработка возврата FBS, ₽': c.fbs_return_flow_amount,
      'Обратная логистика FBS, ₽': c.fbs_return_flow_trans_max_amount
    };
  }));
  return `позиций: ${items.length}`;
}

/* ---------- Остатки FBS (склад из настроек) ---------- */
function syncStocks_() {
  const m = readMain_();
  const wh = String(cfg_('FBS_WAREHOUSE', '')).trim().toLowerCase();
  const bySku = {};
  if (wh) {
    const skus = [...new Set(m.rows.map(r => Number(r['SKU'])).filter(Boolean))];
    chunk_(skus, 500).forEach(part => {
      ozonAll_('/v2/product/info/stocks-by-warehouse/fbs', { sku: part, limit: 1000 },
        r => ({ items: r.products, next: r.has_next ? r.cursor : '' }), 'cursor')
        .forEach(p => { if (String(p.warehouse_name || '').trim().toLowerCase() === wh)
          bySku[key_(p.sku)] = (bySku[key_(p.sku)] || 0) + (Number(p.present) || 0) - (Number(p.reserved) || 0); });
    });
  } else {
    ozonAll_('/v4/product/info/stocks', { filter: { visibility: 'ALL' }, limit: 1000 }, r => ({ items: r.items, next: r.cursor }), 'cursor')
      .forEach(it => (it.stocks || []).filter(s => s.type === 'fbs').forEach(s => { if (s.sku) bySku[key_(s.sku)] = (bySku[key_(s.sku)] || 0) + (s.present || 0) - (s.reserved || 0); }));
  }
  const patch = {};
  m.rows.forEach(r => { patch[r._row] = bySku[key_(r['SKU'])] || 0; });
  mainPatch_(m, 'Остаток FBS', patch);
  return `товаров: ${m.rows.length}, склад: ${wh || 'все FBS'}`;
}

/* ---------- Заказы за N дней (для «Хватит, дней») ---------- */
function syncOrders60_() {
  const days = Number(cfg_('ORDERS_DAYS', 60));
  const to = new Date(), from = new Date(to.getTime() - days * 864e5);
  const qty = {};
  for (let offset = 0; offset < 200000; offset += 1000) {
    const r = ozon_('/v3/posting/fbs/list', { dir: 'ASC', filter: { since: from.toISOString(), to: to.toISOString() }, limit: 1000, offset });
    const part = (r.result && r.result.postings) || [];
    part.forEach(p => { if (p.status === 'cancelled') return;
      (p.products || []).forEach(pr => qty[key_(pr.offer_id)] = (qty[key_(pr.offer_id)] || 0) + (Number(pr.quantity) || 0)); });
    if (!r.result || !r.result.has_next || !part.length) break;
  }
  const m = readMain_();
  if (m.h.indexOf('Заказы 60 дн') < 0) return 'колонка «Заказы 60 дн» удалена — шаг пропущен';
  const patch = {};
  m.rows.forEach(r => patch[r._row] = qty[key_(r['Артикул'])] || 0);
  mainPatch_(m, 'Заказы 60 дн', patch);
  return `за ${days} дн., артикулов с заказами: ${Object.keys(qty).length}`;
}

/* ---------- Закуп: MAX(прайс поставщика; себестоимость 1С на сегодня) ----------
 * Заменяет формулы IMPORTRANGE в колонке «Закупочная цена». Источники — лист «Источники закупа».
 * Код ищется только в прайсе своего источника; источник «1С» — только себестоимость из 1С.
 */
function importCosts_() {
  const colIdx = L => L.toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const usd = Number(cfg_('USD_RATE', 0));
  const sources = {}, noAccess = [];
  readTable_(SHEETS.SOURCES).rows.forEach(src => {
    const name = String(src['Источник']).trim();
    // прайс может быть в долларах: тогда цену умножаем на курс USD_RATE
    const cur = String(src['Валюта'] || 'RUB').trim().toUpperCase();
    const rate = cur === 'USD' ? usd : 1;
    if (cur === 'USD' && !(usd > 0)) log_('Закуп', 'WARN', `Источник «${name}» в USD, но курс USD_RATE не задан`);
    const map = {};
    // в колонке «Лист» можно перечислить несколько листов через запятую или точку с запятой
    const bookId = String(src['ID таблицы']).trim();
    const book = (() => { try { return SpreadsheetApp.openById(bookId); } catch (e) { return null; } })();
    if (!book) {
      noAccess.push(`${name}: https://docs.google.com/spreadsheets/d/${bookId}`);
      sources[name] = {};
      return;
    }
    String(src['Лист']).split(/[;,]/).map(x => x.trim()).filter(Boolean).forEach(sheetName => {
      try {
        // названия листов сравниваем без учёта пробелов по краям («Прайс » у CAS)
        const sh = book.getSheets().find(x => x.getName().trim().toLowerCase() === sheetName.toLowerCase());
        if (!sh) { log_('Закуп', 'WARN', `Источник «${name}»: нет листа «${sheetName}»`); return; }
        const rng = sh.getDataRange();
        const data = rng.getValues(), disp = rng.getDisplayValues();   // по виду ячейки узнаём валюту: $ или р.
        const kc = colIdx(String(src['Колонка кода'])), pc = colIdx(String(src['Колонка закупа']));
        const rc = src['Колонка РРЦ'] ? colIdx(String(src['Колонка РРЦ'])) : -1;
        // валюта ячейки: знак $ — доллары, «р.»/«₽»/«руб» — рубли, без пометки — валюта источника
        const cellRate = txt => {
          const t = String(txt || '');
          if (/\$|usd/i.test(t)) return usd;
          if (/р\.|₽|руб|rub/i.test(t)) return 1;
          return rate;
        };
        const parse = (v, txt) => {
          let n = num_(v);
          if (n === '') n = num_(String(txt || '').replace(/[^\d,.\-]/g, ''));   // «3 400р.» как текст
          return n;
        };
        data.forEach((row, i) => {
          const k = key_(row[kc]);
          if (!k || map[k] !== undefined) return;
          const cost = parse(row[pc], disp[i][pc]);
          const rrc = rc >= 0 ? parse(row[rc], disp[i][rc]) : '';
          if (cost !== '' && cellRate(disp[i][pc]) === usd && !(usd > 0)) return;       // курс не задан — не гадаем
          map[k] = {
            cost: (cost === '' ? '' : Math.round(cost * cellRate(disp[i][pc]))),
            rrc: (rrc === '' ? '' : Math.round(rrc * cellRate(rc >= 0 ? disp[i][rc] : '')))
          };
        });
      } catch (e) { log_('Закуп', 'WARN', `Источник «${name}», лист «${sheetName}»: ${e.message}`); }
    });
    sources[name] = map;
  });

  // 1С: строки — product_id, колонки — даты. Берём сегодняшнюю, иначе последнюю доступную.
  const c1 = {}, has1c = {};
  const id1c = String(cfg_('COST_1C_SHEET_ID', '')).trim();
  if (id1c) {
    let data = [];
    try {
      data = SpreadsheetApp.openById(id1c).getSheetByName(String(cfg_('COST_1C_SHEET', 'Prices'))).getDataRange().getValues();
    } catch (e) {
      noAccess.push(`Таблица 1С: https://docs.google.com/spreadsheets/d/${id1c}`);
    }
    const tz = Session.getScriptTimeZone(), today = Utilities.formatDate(new Date(), tz, 'yyyy.MM.dd');
    const hdr = (data[0] || []).map(v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy.MM.dd') : String(v).trim());
    let col = data.length ? hdr.indexOf(today) : -1;
    if (col < 0) { const past = hdr.map((v, i) => [v, i]).filter(([v]) => /^\d{4}\.\d{2}\.\d{2}$/.test(v) && v <= today).sort(); if (past.length) col = past[past.length - 1][1]; }
    if (col >= 0) data.slice(1).forEach(row => {
      const k = key_(row[0]);
      if (k) has1c[k] = true;
      const v = num_(row[col]);
      if (v !== '' && k) c1[k] = v;
    });
    else if (data.length) log_('Закуп', 'WARN', 'В таблице 1С не найдена колонка с датой');
  }

  const m = readMain_(), cost = {}, rrc = {}, miss = [], notInPrice = [];
  m.rows.forEach(r => {
    const src = String(r['Источник закупа'] || '').trim();
    if (src.toLowerCase() === 'вручную') return;
    const code = key_(r['Код в прайсе']);
    // цену ищем только в своём прайсе: коды у поставщиков пересекаются, в чужом прайсе под тем же кодом — другой товар.
    // Источник «1С» — только себестоимость из 1С
    const s = sources[src] && sources[src][code];
    if (sources[src] && code && !s) notInPrice.push(`${r['Артикул']} (${src}, код ${code})`);
    const vals = [s && s.cost, c1[key_(r['Product ID'])]].filter(v => typeof v === 'number' && v > 0);
    cost[r._row] = vals.length ? Math.max(...vals) : '';
    if (s && s.rrc !== '') rrc[r._row] = s.rrc;
    if (!vals.length) {
      const pid = key_(r['Product ID']);
      const why = !code ? 'нет «Кода в прайсе»'
        : !sources[src] ? `источник «${src || '—'}» не настроен`
        : `кода нет в прайсе «${src}»`;
      const why1c = !id1c ? 'таблица 1С не указана' : !has1c[pid] ? 'в 1С нет строки по product_id' : 'в 1С пусто на выбранную дату';
      miss.push(`${r['Артикул']} (${why}; ${why1c})`);
    }
  });
  mainPatch_(m, 'Закуп, ₽', cost); mainPatch_(m, 'РРЦ, ₽', rrc);
  if (noAccess.length) log_('Закуп', 'WARN', 'НЕТ ДОСТУПА к таблицам (откройте ссылку и запросите доступ для своего аккаунта):\n' + noAccess.join('\n'));
  if (notInPrice.length) log_('Закуп', 'WARN', `Кода нет в своём прайсе, взят только закуп из 1С (${notInPrice.length}) — проверьте «Код в прайсе» и «Источник закупа»: ${notInPrice.join(', ')}`);
  if (miss.length) log_('Закуп', 'WARN', `Нет закупа (${miss.length}): ${miss.join(', ')}`);
  return `обновлено: ${Object.keys(cost).length - miss.length}, без закупа: ${miss.length}` +
    (noAccess.length ? ` | нет доступа к ${noAccess.length} табл. — ссылки в «Логе»` : '');
}

/* ---------- Источник закупа по формулам старой таблицы (разовая операция) ----------
 * В старом «Командном пункте» формула закупа у каждого товара ссылалась на свой прайс.
 * По ID таблиц в формуле находим источник на листе «Источники закупа» и ставим его в «Источник закупа».
 * Формула только с таблицей 1С → «1С». Строки «вручную» и формулы с несколькими прайсами не трогаем.
 */
var OLD_TABLE_ID_ = '1Rdh8-EQEt6UiPl8ta032kKbQ2JsFDHauBmPDY9SBqX0';   // старый «Командный пункт»
function sourcesFromOldTable() { run_('Источники закупа из старой таблицы', sourcesFromOldTable_); }

function sourcesFromOldTable_() {
  let oldSh;
  try { oldSh = SpreadsheetApp.openById(OLD_TABLE_ID_).getSheetByName('Ozon'); }
  catch (e) { return `нет доступа к старой таблице: https://docs.google.com/spreadsheets/d/${OLD_TABLE_ID_}`; }
  if (!oldSh) throw new Error('В старой таблице нет листа Ozon');
  const oh = headersAt_(oldSh, 2).map(x => x.toLowerCase());
  const find = re => oh.findIndex(x => re.test(x));
  const iArt = find(/^артикул/), iPid = find(/^product id/), iCost = find(/закуп/), iCode = find(/^прайс/);
  if (iArt < 0 || iPid < 0 || iCost < 0) throw new Error('В старой таблице нет колонок «Артикул», «Product ID» или «Закупочная цена»');
  const n = oldSh.getLastRow() - 2;
  if (n < 1) return 'старая таблица пустая';
  const vals = oldSh.getRange(3, 1, n, oh.length).getValues();
  const fx = oldSh.getRange(3, iCost + 1, n, 1).getFormulas();

  const names = {};                                       // ID таблицы → название источника
  readTable_(SHEETS.SOURCES).rows.forEach(s => { const id = String(s['ID таблицы']).trim(); if (id) names[id] = String(s['Источник']).trim(); });
  const id1c = String(cfg_('COST_1C_SHEET_ID', '')).trim();

  const byPid = {}, byArt = {};
  vals.forEach((row, i) => {
    const f = String(fx[i][0] || ''); if (!f) return;
    const found = Array.from(new Set(Object.keys(names).filter(id => f.indexOf(id) >= 0).map(id => names[id])));
    const d = found.length === 1 ? { src: found[0] }
      : !found.length && id1c && f.indexOf(id1c) >= 0 ? { src: '1С' } : { several: found };
    d.code = iCode >= 0 ? row[iCode] : '';
    if (/^\d+$/.test(key_(row[iPid]))) byPid[key_(row[iPid])] = d;
    if (key_(row[iArt])) byArt[key_(row[iArt])] = d;
  });

  const m = readMain_(), src = {}, code = {}, changes = [], skipped = [];
  let same = 0;
  m.rows.forEach(r => {
    const d = byPid[key_(r['Product ID'])] || byArt[key_(r['Артикул'])];
    const cur = String(r['Источник закупа'] || '').trim();
    if (!d || cur.toLowerCase() === 'вручную') return;
    if (!d.src) { skipped.push(`${r['Артикул']} (${d.several.join(' + ') || 'прайс не найден'})`); return; }
    if (cur === d.src) same++;
    else { src[r._row] = d.src; changes.push(`${r['Артикул']}: ${cur || '—'} → ${d.src}`); }
    if (!key_(r['Код в прайсе']) && key_(d.code) && d.src !== '1С') code[r._row] = d.code;
  });
  const skippedNote = skipped.length ? `, не тронуты (в старой формуле несколько прайсов): ${skipped.length}` : '';
  if (!changes.length && !Object.keys(code).length) return `менять нечего: совпадает ${same}${skippedNote}`;

  const ui = SpreadsheetApp.getUi();
  const a = ui.alert('Источники закупа из старой таблицы',
    `Поменяется «Источник закупа» у ${changes.length} товаров, «Код в прайсе» заполнится у ${Object.keys(code).length}. ` +
    'Список будет в «Логе». Применить?', ui.ButtonSet.YES_NO);
  if (a !== ui.Button.YES) return 'отменено';
  mainPatch_(m, 'Источник закупа', src);
  mainPatch_(m, 'Код в прайсе', code);
  log_('Источники закупа', 'INFO', `Поменяли (${changes.length}):\n${changes.join('\n')}` +
    (skipped.length ? `\n\nНе тронуты — в старой формуле несколько прайсов, выберите вручную (${skipped.length}):\n${skipped.join('\n')}` : ''));
  return `источник поменян у ${changes.length}, совпадал у ${same}${skippedNote}. Теперь нажмите «Обновить закуп»`;
}

function updateUsdRate_() {
  const d = JSON.parse(UrlFetchApp.fetch('https://www.cbr-xml-daily.ru/daily_json.js').getContentText());
  const usd = d.Valute && d.Valute.USD; if (!usd) throw new Error('USD не найден');
  setCfg_('USD_RATE', usd.Value / usd.Nominal);
  return 'USD = ' + (usd.Value / usd.Nominal).toFixed(4);
}


/* ================== 03_actions.gs ================== */
/** =====================================================================
 *  «БОГ АКЦИЙ»: товар × акция, прибыль по цене акции, массовые кнопки
 *  Порог = колонка «Порог акций, ₽» листа Ozon (маржа PROMO_MARGIN).
 *  Строка 1 — панель: выбранная акция, порог, счётчики, прибыль и потери.
 * ===================================================================== */
function refreshActions()         { run_('Обновить «Бог акций»', refreshActions_); }
function addEligibleToSelected()  { run_('Добавить подходящие в выбранную акцию', () => addEligible_('selected')); }
function addEligibleToAll()       { run_('Добавить подходящие во все акции', () => addEligible_('all')); }
function removeIneligible()       { run_('Убрать неподходящие из акций', removeIneligible_); }
function applySelectedActions()   { run_('Применить действия по выбранной акции', () => applySelectedActions_()); }
function removeMarkedFromAction() { run_('Удалить из акции выделенные и отмеченные', removeMarked_); }
function applySelectedMin()       { run_('Применить отметки с маржой от ' + pctCfg_('MIN_MARGIN', 0.10), () => applySelectedActions_('min')); }

/** Выделяет ключевые колонки: прибыль и маржа в акции */
function highlightBogKeyCols_(sh, n) {
  const h = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const cols = ['Прибыль в акции, ₽', 'Маржа в акции, %'].map(x => h.indexOf(x) + 1).filter(x => x > 0);
  cols.forEach(c => {
    sh.getRange(2, c).setBackground('#1A3B5C').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.getRange(BOG_FIRST, c, n, 1).setFontWeight('bold').setFontSize(11)
      .setBorder(null, true, null, true, false, false, '#5B7DA6', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  // закрепляем всё до «Порога акций» включительно, чтобы при прокрутке вправо был виден товар и его порог
  const freeze = h.indexOf('Порог акций, ₽') + 1;
  if (freeze > 0) { try { sh.setFrozenColumns(freeze); } catch (e) {} }
}

/** Пишет причины в колонку «Результат» листа «Бог акций» */
function patchBogResult_(sh, res, last) {
  const h = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const c = h.indexOf('Результат') >= 0 ? h.indexOf('Результат') + 1 : BOG.RESULT;
  const rng = sh.getRange(BOG_FIRST, c, last - BOG_FIRST + 1, 1), v = rng.getValues();
  Object.keys(res).forEach(row => { v[Number(row) - BOG_FIRST][0] = res[row]; });
  rng.setValues(v);
}

function distributeBestActions()  { run_('Распределить по лучшим акциям', distributeBestActions_); }

var BOG_FIRST = 3, BOG_ACT_COL = 25;   // данные с 3-й строки, акции с колонки Y
var MAXSHEET = 'Акции_макс';           // скрытый лист: макс. цены участия по каждой акции
var BOG = { PID: 1, PRICE: 7, FLOOR: 9, IN_ACTION: 10, MAX: 11, PROFIT: 13, MARGIN: 14,
              STATUS: 16, PASS: 17, ACTION: 18, RESULT: 19 };

function fetchActionProducts_(id) {
  return ozonAll_('/v1/actions/products', { action_id: Number(id), limit: 1000 },
    r => ({ items: r.result && r.result.products, next: r.result && r.result.last_id }), 'last_id');
}
function fetchActionCandidates_(id) {
  return ozonAll_('/v1/actions/candidates', { action_id: Number(id), limit: 1000 },
    r => ({ items: r.result && r.result.products, next: r.result && r.result.last_id }), 'last_id');
}
function listActions_() { return ozon_('/v1/actions', undefined, 'get').result || []; }
function excludeRe_() { const s = String(cfg_('ACTIONS_EXCLUDE', '')).trim(); return s ? new RegExp(s, 'i') : /$^/; }
function selectedTitle_() { return String(sheet_(SHEETS.BOG).getRange('B1').getValue()).trim(); }

/** product_id → {floor, min, stock, art} с листа Ozon: floor — «Порог акций» (PROMO_MARGIN), min — «Мин. цена» (MIN_MARGIN) */
function thresholds_() {
  const out = {};
  readMain_().rows.forEach(r => {
    const pid = key_(r['Product ID']); if (!pid) return;
    out[pid] = { floor: Number(r['Порог акций, ₽']) || 0, min: Number(r['Мин. цена, ₽']) || 0,
                 stock: Number(r['Остаток FBS']) || 0, art: r['Артикул'] };
  });
  return out;
}

/* ---------- Товары, вручную добавленные в акцию с маржой от MIN_MARGIN ----------
 * Кнопка «Применить отметки с маржой от 10%» запоминает пары «акция:товар», которые
 * вошли в акцию ниже «Порога акций». Для них ночная чистка, аудит и распределение
 * сверяют цену акции с «Мин. ценой» (MIN_MARGIN), а не с «Порогом акций» (PROMO_MARGIN).
 * Ниже «Мин. цены» товар убирается из акции всегда.
 */
var MANUAL_PROMO_PROP = 'PROMO_MIN_OK';
function manualKey_(actionId, pid) { return key_(actionId) + ':' + key_(pid); }
function manualPromo_() {
  try { return new Set(JSON.parse(PropertiesService.getDocumentProperties().getProperty(MANUAL_PROMO_PROP) || '[]')); }
  catch (e) { return new Set(); }
}
function saveManualPromo_(set) {
  try { PropertiesService.getDocumentProperties().setProperty(MANUAL_PROMO_PROP, JSON.stringify(Array.from(set))); }
  catch (e) { log_('Акции: ручные решения', 'WARN', 'не удалось сохранить список: ' + e.message); }
}
/** Порог для товара в конкретной акции: ручное решение — «Мин. цена», иначе «Порог акций» */
function floorIn_(t, manual, actionId, pid) {
  return manual.has(manualKey_(actionId, pid)) && t.min > 0 ? t.min : t.floor;
}

/* ---------- Перестроить матрицу ---------- */
function refreshActions_() {
  const sh = sheet_(SHEETS.BOG), acts = listActions_(), main = readMain_();
  // колонку «Ozon-карта» больше не ведём — удаляем со старых листов один раз
  const oldCard = headersAt_(sh, 2).indexOf('Ozon-карта');
  if (oldCard >= 0 && oldCard < BOG_ACT_COL) sh.deleteColumn(oldCard + 1);
  const prods = main.rows.filter(r => Number(r['Product ID']) > 0);
  const sel = acts.find(a => String(a.title).trim() === selectedTitle_()) || acts[0];

  // тянем участие и максимальные цены сразу по ВСЕМ акциям — переключение акции потом мгновенное
  const price = {}, maxAll = {};
  acts.forEach(a => {
    price[a.id] = {}; maxAll[a.id] = {};
    fetchActionProducts_(a.id).forEach(p => {
      price[a.id][key_(p.id)] = p.action_price;
      if (p.max_action_price) maxAll[a.id][key_(p.id)] = p.max_action_price;
    });
    fetchActionCandidates_(a.id).forEach(p => {
      if (maxAll[a.id][key_(p.id)] === undefined) maxAll[a.id][key_(p.id)] = p.max_action_price;
    });
  });

  const na = acts.length, n = prods.length, lastAct = BOG_ACT_COL + Math.max(na, 1) - 1;
  if (sh.getMaxColumns() < lastAct) sh.insertColumnsAfter(sh.getMaxColumns(), lastAct - sh.getMaxColumns());
  const oldLast = sh.getLastRow();
  sh.getRange(1, BOG_ACT_COL, sh.getMaxRows(), sh.getMaxColumns() - BOG_ACT_COL + 1).clearContent();
  if (oldLast >= BOG_FIRST) sh.getRange(BOG_FIRST, 1, oldLast - BOG_FIRST + 1, BOG_ACT_COL - 1).clearContent();
  if (sh.getMaxRows() < BOG_FIRST + n) sh.insertRowsAfter(sh.getMaxRows(), BOG_FIRST + n - sh.getMaxRows());

  const excl = excludeRe_();
  if (na) {
    sh.getRange(1, BOG_ACT_COL, 1, na).setValues([acts.map(a => a.id)]);
    sh.getRange(2, BOG_ACT_COL, 1, na).setValues([acts.map(a => a.title)])
      .setBackgrounds([acts.map(a => excl.test(a.title || '') ? '#A6A6A6' : '#7030A0')])
      .setFontColor('#FFFFFF').setWrap(true);
  }
  if (!n) return 'нет товаров';

  // значения: id товара и матрица цен участия
  sh.getRange(BOG_FIRST, BOG.PID, n, 1).setValues(prods.map(r => [r['Product ID']]));
  if (na) sh.getRange(BOG_FIRST, BOG_ACT_COL, n, na).setValues(prods.map(r =>
    acts.map(a => { const v = price[a.id][key_(r['Product ID'])]; return v === undefined ? '' : v; })));
  writeMaxSheet_(acts, prods, maxAll);

  // формулы: пишем в первую строку в нотации A1, затем копируем вниз (ссылки сдвинутся сами)
  const L = i => { let t = '', x = i; while (x > 0) { const m = (x - 1) % 26; t = String.fromCharCode(65 + m) + t; x = (x - m - 1) / 26; } return t; };
  const mc = name => '$' + L(col_(main.h, name) + 1);
  const R = BOG_FIRST;
  const idx = name => `INDEX(Ozon!${mc(name)}:${mc(name)},MATCH($A${R},Ozon!${mc('Product ID')}:${mc('Product ID')},0))`;
  const txt = name => `=IFERROR(${idx(name)}&"","")`;
  const nm = name => `=IFERROR(IF(${idx(name)}="","",${idx(name)}),"")`;
  const actFrom = L(BOG_ACT_COL), actTo = L(lastAct);
  const selCol = `MATCH($B$1,$${actFrom}$2:$${actTo}$2,0)`;
  const inAct = `INDEX($${actFrom}${R}:$${actTo}${R},1,${selCol})`;
  const maxVal = `INDEX(${MAXSHEET}!$B:$${L(1 + Math.max(na, 1))},MATCH($A${R},${MAXSHEET}!$A:$A,0),MATCH($B$1,${MAXSHEET}!$B$2:$${L(1 + Math.max(na, 1))}$2,0))`;
  const P = `IF(ISNUMBER($J${R}),$J${R},$K${R})`;
  const F = {
    2: txt('Артикул'), 3: txt('Название'), 4: txt('Категория'), 5: nm('Остаток FBS'),
    6: main.h.indexOf('Заказы 60 дн') >= 0 ? nm('Заказы 60 дн') : '',
    7: `=IFERROR(IF(${idx('Цена продажи, ₽')}="",${idx('Цена на Ozon, ₽')},${idx('Цена продажи, ₽')}),"")`,
    8: nm('Мин. цена, ₽'), 9: nm('Порог акций, ₽'),
    10: `=IFERROR(IF(${inAct}="","",${inAct}),"")`,
    11: `=IFERROR(IF(${maxVal}="","",${maxVal}),"")`,
    12: `=IF(OR(NOT(ISNUMBER(${P})),NOT(ISNUMBER($G${R}))),"",1-${P}/$G${R})`,
    13: `=IF(OR(NOT(ISNUMBER(${P})),$W${R}=""),"",ROUND(${P}-${P}*$V${R}-$W${R},0))`,
    14: `=IF(OR($M${R}="",NOT(ISNUMBER(${P}))),"",$M${R}/${P})`,
    15: `=IF(OR($M${R}="",$X${R}=""),"",$X${R}-$M${R})`,
    16: `=IF(ISNUMBER($J${R}),"Участвует",IF(ISNUMBER($K${R}),"Кандидат","Нет в акции"))`,
    17: `=IF(NOT(ISNUMBER($I${R})),"❔ нет закупа",IF(NOT(ISNUMBER(${P})),"—",IF(${P}>=$I${R},"✓ проходит","✗ ниже порога")))`,
    20: `=COUNT($${actFrom}${R}:$${actTo}${R})`,
    21: `=IF(NOT(ISNUMBER($I${R})),"",COUNTIF($${actFrom}${R}:$${actTo}${R},"<"&$I${R}))`,
    22: `=IFERROR(${idx('Комиссия, %')}+${idx('Эквайринг, %')},"")`,
    23: nm('Затраты фикс., ₽'),
    24: `=IF(OR(NOT(ISNUMBER($G${R})),$W${R}=""),"",ROUND($G${R}-$G${R}*$V${R}-$W${R},0))`
  };
  Object.keys(F).forEach(col => {
    if (!F[col]) return;
    const cell = sh.getRange(BOG_FIRST, Number(col));
    cell.setFormula(fx_(F[col]));
    if (n > 1) cell.copyTo(sh.getRange(BOG_FIRST + 1, Number(col), n - 1, 1));
  });

  // ввод: выпадающий список действий и список акций в B1
  sh.getRange(BOG_FIRST, BOG.ACTION, n, 1).clearDataValidations().setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Добавить', 'Удалить'], true).build());
  if (na) {
    sh.getRange('B1').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange(2, BOG_ACT_COL, 1, na), true).build());
    if (sel) sh.getRange('B1').setValue(sel.title);
  }
  paintBog_(sh, n, na, lastAct);
  highlightBogKeyCols_(sh, n);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(2, 1, n + 1, lastAct).createFilter();
  if (na > 60) log_('Бог акций', 'WARN', 'Акций больше 60: в формуле «Мин. цена в акциях» листа Ozon учтены только первые 60');
  return `акций: ${na}, товаров: ${n}, выбрана: ${sel ? sel.title : '—'}`;
}

/** Скрытый лист с максимальными ценами участия по всем акциям */
function writeMaxSheet_(acts, prods, maxAll) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MAXSHEET) || ss.insertSheet(MAXSHEET);
  sh.clear();
  const na = acts.length, n = prods.length;
  if (sh.getMaxColumns() < na + 1) sh.insertColumnsAfter(sh.getMaxColumns(), na + 1 - sh.getMaxColumns());
  sh.getRange(1, 1).setValue('id акции ▶');
  sh.getRange(2, 1).setValue('product_id');
  if (na) {
    sh.getRange(1, 2, 1, na).setValues([acts.map(a => a.id)]);
    sh.getRange(2, 2, 1, na).setValues([acts.map(a => a.title)]);
  }
  if (n) {
    sh.getRange(3, 1, n, 1).setValues(prods.map(r => [r['Product ID']]));
    if (na) sh.getRange(3, 2, n, na).setValues(prods.map(r =>
      acts.map(a => { const v = maxAll[a.id][key_(r['Product ID'])]; return v === undefined ? '' : v; })));
  }
  sh.hideSheet();
}

/** Подсветка: цена акции ниже порога — красная, проходит — зелёная */
function paintBog_(sh, n, na, lastAct) {
  const grid = sh.getRange(BOG_FIRST, BOG_ACT_COL, n, Math.max(na, 1));
  const tl = grid.getA1Notation().split(':')[0].replace(/\d+/, '');
  const rule = (formula, bg, fc, ranges, bold) => {
    // разделитель аргументов зависит от локали таблицы, иначе правило молча не срабатывает
    let b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(fx_(formula)).setBackground(bg).setRanges(ranges);
    if (fc) b = b.setFontColor(fc);
    if (bold) b = b.setBold(true);
    return b.build();
  };
  const R = (col, w) => [sh.getRange(BOG_FIRST, col, n, w || 1)];
  const pass = `LEFT($Q${BOG_FIRST},1)="✓"`, fail = `LEFT($Q${BOG_FIRST},1)="✗"`;
  // цены, прибыль и маржа красятся по вердикту «Проходит?»
  const priceCells = [sh.getRange(BOG_FIRST, BOG.IN_ACTION, n, 2),        // цена в акции и макс. цена
                      sh.getRange(BOG_FIRST, BOG.PROFIT, n, 2)];          // прибыль и маржа в акции
  sh.setConditionalFormatRules([
    rule(`=AND(ISNUMBER(${tl}${BOG_FIRST}),ISNUMBER($I${BOG_FIRST}),${tl}${BOG_FIRST}<$I${BOG_FIRST})`, '#F8CBAD', '#9C0006', [grid]),
    rule(`=AND(ISNUMBER(${tl}${BOG_FIRST}),ISNUMBER($I${BOG_FIRST}),${tl}${BOG_FIRST}>=$I${BOG_FIRST})`, '#C6EFCE', '#006100', [grid]),
    rule(`=${pass}`, '#C6EFCE', '#006100', R(BOG.PASS), true),
    rule(`=${fail}`, '#F8CBAD', '#9C0006', R(BOG.PASS), true),
    // статус — по участию: участвует зелёным, «нет в акции» серым, кандидат без заливки
    rule(`=$P${BOG_FIRST}="Участвует"`, '#C6EFCE', '#006100', R(BOG.STATUS), true),
    rule(`=$P${BOG_FIRST}="Нет в акции"`, '#F1F3F4', '#9AA0A6', R(BOG.STATUS)),
    rule(`=${pass}`, '#E6F4EA', '#1E7B34', priceCells),
    rule(`=${fail}`, '#FCE8E6', '#B3261E', priceCells),
    rule(`=N($U${BOG_FIRST})>0`, '#F8CBAD', '#9C0006', R(21), true)
  ]);
}

/* ---------- Добавить все подходящие по марже ---------- */
function addEligible_(mode) {
  const dry = isDryRun_(), thr = thresholds_();
  const acts = listActions_();
  const targets = mode === 'selected'
    ? acts.filter(a => String(a.title).trim() === selectedTitle_())
    : acts.filter(a => !excludeRe_().test(a.title || ''));
  if (!targets.length) throw new Error('Акция не найдена — обновите «Бог акций»');

  const report = []; let total = 0;
  targets.forEach(a => {
    const inside = new Set(fetchActionProducts_(a.id).map(p => key_(p.id)));
    const items = [];
    fetchActionCandidates_(a.id).forEach(cnd => {
      const pid = key_(cnd.id), t = thr[pid], mx = Number(cnd.max_action_price) || 0;
      if (inside.has(pid) || !t || !(t.floor > 0) || !mx) return;
      if (mx < t.floor) return;
      const o = { product_id: Number(cnd.id), action_price: mx };
      if (/STOCK/i.test(a.action_type || '')) o.stock = Math.max(1, t.stock);
      items.push(o);
    });
    if (!items.length) return;
    if (dry) { report.push(`${a.title}: добавили бы ${items.length} — ${items.slice(0, 30).map(i => thr[key_(i.product_id)].art).join(', ')}`); total += items.length; return; }
    chunk_(items, 1000).forEach(part => {
      const r = ozon_('/v1/actions/products/activate', { action_id: Number(a.id), products: part });
      const ok = ((r.result && r.result.product_ids) || []).length, rej = (r.result && r.result.rejected) || [];
      total += ok;
      report.push(`${a.title}: +${ok}` + (rej.length ? `, отказ ${rej.length}: ${rej.slice(0, 10).map(x => (thr[key_(x.product_id)] || {}).art + ' — ' + x.reason).join('; ')}` : ''));
    });
  });
  log_('Акции: добавление', dry ? 'DRY' : 'INFO', report.join('\n') || 'подходящих кандидатов нет');
  if (!dry && total) refreshActions_();
  return `${dry ? '[ПРОВЕРКА] ' : ''}акций: ${targets.length}, товаров: ${total}. Подробно — лист «Лог»`;
}

/* ---------- Убрать из всех акций товары ниже порога ----------
 * Порог — «Порог акций» (PROMO_MARGIN); для ручных решений — «Мин. цена» (MIN_MARGIN).
 */
function removeIneligible_() {
  const dry = isDryRun_(), thr = thresholds_(), report = [];
  const manual = manualPromo_(), alive = new Set();
  let total = 0, noFloor = 0, keptManual = 0;
  listActions_().forEach(a => {
    const bad = [];
    fetchActionProducts_(a.id).forEach(p => {
      const t = thr[key_(p.id)], mk = manualKey_(a.id, p.id);
      if (manual.has(mk)) alive.add(mk);
      if (!t || !(t.floor > 0)) { noFloor++; return; }
      const floor = floorIn_(t, manual, a.id, p.id);
      if (Number(p.action_price) < floor) bad.push({ id: Number(p.id), art: t.art, price: p.action_price, floor, mk });
      else if (Number(p.action_price) < t.floor) keptManual++;
    });
    if (!bad.length) return;
    const list = bad.slice(0, 30).map(b => `${b.art} (${b.price} < ${b.floor})`).join(', ');
    if (dry) { report.push(`${a.title}: убрали бы ${bad.length} — ${list}`); total += bad.length; return; }
    bad.forEach(b => alive.delete(b.mk));
    chunk_(bad, 1000).forEach(part => {
      const r = ozon_('/v1/actions/products/deactivate', { action_id: Number(a.id), product_ids: part.map(b => b.id) });
      const ok = ((r.result && r.result.product_ids) || []).length, rej = (r.result && r.result.rejected) || [];
      total += ok;
      report.push(`${a.title}: −${ok} — ${list}` + (rej.length ? `; не удалось ${rej.length}: ${rej.slice(0, 10).map(x => x.reason).join('; ')}` : ''));
    });
  });
  if (keptManual) report.push(`Оставлены по ручному решению (маржа от ${pctCfg_('MIN_MARGIN', 0.10)}): ${keptManual}`);
  if (noFloor) report.push(`Товары в акциях без порога (нет закупа), не тронуты: ${noFloor}`);
  // из списка ручных решений убираем товары, которых в акции больше нет
  if (alive.size !== manual.size) saveManualPromo_(alive);
  log_('Акции: удаление', dry ? 'DRY' : 'INFO', report.join('\n') || 'всё в акциях проходит по марже');
  if (!dry && total) refreshActions_();
  return `${dry ? '[ПРОВЕРКА] ' : ''}ниже порога: ${total}. Подробно — лист «Лог»`;
}

/* ---------- Удалить из выбранной акции выделенные строки и строки с отметкой «Удалить» ---------- */
function removeMarked_() {
  const sh = sheet_(SHEETS.BOG), ss = SpreadsheetApp.getActive();
  const rows = [];
  if (ss.getActiveSheet().getName() === SHEETS.BOG) {
    const list = ss.getActiveRangeList();
    (list ? list.getRanges() : []).forEach(r => {
      for (let i = Math.max(r.getRow(), BOG_FIRST); i <= Math.min(r.getLastRow(), sh.getLastRow()); i++) rows.push(i);
    });
  }
  if (rows.length) {                                      // выделенные строки помечаем «Удалить»
    const h = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
    const c = (h.indexOf('Действие') >= 0 ? h.indexOf('Действие') : BOG.ACTION - 1) + 1;
    rows.forEach(i => { if (sh.getRange(i, 1).getValue() !== '') sh.getRange(i, c).setValue('Удалить'); });
  }
  return applySelectedActions_('del');
}

/* ---------- Ручные отметки в колонке «Действие» по выбранной акции ---------- */
/**
 * Обычный режим: добавляем, только если цена акции не ниже «Порога акций» (PROMO_MARGIN).
 * mode = 'min': порог — «Мин. цена» (MIN_MARGIN). Товары между «Мин. ценой» и «Порогом акций»
 * запоминаются как ручное решение, и ночная чистка их не убирает. Ниже «Мин. цены» — никогда.
 * mode = 'del': только отметки «Удалить», «Добавить» не трогаем; в боевом режиме — с подтверждением.
 */
function applySelectedActions_(mode) {
  const sh = sheet_(SHEETS.BOG), dry = isDryRun_(), toMin = mode === 'min';
  const minPct = pctCfg_('MIN_MARGIN', 0.10), promoPct = pctCfg_('PROMO_MARGIN', 0.12);
  const sel = listActions_().find(a => String(a.title).trim() === selectedTitle_());
  if (!sel) throw new Error('Акция в ячейке B1 не найдена — обновите «Бог акций»');
  const last = sh.getLastRow();
  if (last < BOG_FIRST) return 'нет данных';
  // колонки ищем по заголовкам: если лист сдвинулся, отметки всё равно найдутся
  const h = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  const at = (name, def) => { const i = h.indexOf(name); return i >= 0 ? i : def - 1; };
  const cPid = at('Product ID', BOG.PID), cAct = at('Действие', BOG.ACTION);
  const cMax = at('Макс. цена акции, ₽', BOG.MAX), cFloor = at('Порог акций, ₽', BOG.FLOOR);
  const cMin = at('Мин. цена, ₽', 8), cStock = at('Остаток', 5);
  const width = Math.max(sh.getLastColumn(), BOG.RESULT);
  const data = sh.getRange(BOG_FIRST, 1, last - BOG_FIRST + 1, width).getValues();

  const add = [], del = [], res = {};
  let marks = 0;
  data.forEach((row, i) => {
    const act = String(row[cAct] || '').trim(); if (!act) return;
    if (mode === 'del' && act !== 'Удалить') return;
    marks++;
    const rowNum = BOG_FIRST + i, pid = Number(row[cPid]);
    if (!pid) { res[rowNum] = '✗ нет Product ID в строке'; return; }
    if (act === 'Удалить') { del.push({ pid, row: rowNum }); return; }
    const price = Math.round(Number(row[cMax]) || 0);
    const floor = Number(row[cFloor]) || 0, minFloor = Number(row[cMin]) || 0;
    if (!price) { res[rowNum] = '✗ нет цены акции (товар не кандидат)'; return; }
    // без порога маржу не проверить — такой товар в акцию не пускаем
    if (!floor || (toMin && !minFloor)) { res[rowNum] = '✗ нет порога: не посчитан закуп или тарифы'; return; }
    if (toMin && price < minFloor) {
      res[rowNum] = `✗ ниже мин. цены ${minFloor}: маржа меньше ${minPct} — не добавляем`;
      return;
    }
    if (!toMin && price < floor) {
      res[rowNum] = `✗ ниже порога ${floor} (маржа меньше ${promoPct}). Если согласны на маржу от ${minPct} — «🏷 АКЦИИ → Применить отметки с маржой от ${minPct}»`;
      return;
    }
    add.push({ pid, row: rowNum, price, stock: Number(row[cStock]) || 1, below: price < floor });
  });
  if (!add.length && !del.length) {
    if (marks) {                                    // отметки есть, но ни одна не прошла проверку
      patchBogResult_(sh, res, last);
      return `отметок: ${marks}, ни одна не прошла: причины в колонке «Результат»`;
    }
    return `нет отметок в колонке «Действие» (ищу в колонке ${cAct + 1}: «${h[cAct] || '—'}»)`;
  }

  const belowNote = `маржа ниже ${promoPct}, но не ниже ${minPct} — ручное решение`;
  if (mode === 'del' && !dry) {
    const ui = SpreadsheetApp.getUi();
    const a = ui.alert('Удалить из акции', `Убрать из акции «${sel.title}» товаров: ${del.length}?`, ui.ButtonSet.YES_NO);
    if (a !== ui.Button.YES) return 'отменено';
  }
  if (dry) {
    const note = 'ПРОВЕРКА — режим проверки, в Ozon не ушло: ';
    add.forEach(x => res[x.row] = `${note}добавили бы по ${x.price}` + (x.below ? ` (${belowNote})` : ''));
    del.forEach(x => res[x.row] = `${note}убрали бы из акции`);
  } else {
    const manual = manualPromo_();
    chunk_(add, 1000).forEach(part => {
      const products = part.map(x => { const o = { product_id: x.pid, action_price: x.price };
        if (/STOCK/i.test(sel.action_type || '')) o.stock = Math.max(1, x.stock); return o; });
      const r = ozon_('/v1/actions/products/activate', { action_id: Number(sel.id), products });
      const rej = {}; ((r.result && r.result.rejected) || []).forEach(x => rej[key_(x.product_id)] = x.reason);
      part.forEach(x => {
        if (rej[key_(x.pid)]) { res[x.row] = '✗ ' + rej[key_(x.pid)]; return; }
        res[x.row] = `✓ в акции по ${x.price}` + (x.below ? ` (${belowNote})` : '');
        if (x.below) manual.add(manualKey_(sel.id, x.pid));
      });
    });
    chunk_(del, 1000).forEach(part => {
      const r = ozon_('/v1/actions/products/deactivate', { action_id: Number(sel.id), product_ids: part.map(x => x.pid) });
      const rej = {}; ((r.result && r.result.rejected) || []).forEach(x => rej[key_(x.product_id)] = x.reason);
      part.forEach(x => {
        if (rej[key_(x.pid)]) { res[x.row] = '✗ ' + rej[key_(x.pid)]; return; }
        res[x.row] = '✓ убран';
        manual.delete(manualKey_(sel.id, x.pid));
      });
    });
    saveManualPromo_(manual);
  }
  const cRes = at('Результат', BOG.RESULT);
  const colRes = sh.getRange(BOG_FIRST, cRes + 1, last - BOG_FIRST + 1, 1), vRes = colRes.getValues();
  const colAct = sh.getRange(BOG_FIRST, cAct + 1, last - BOG_FIRST + 1, 1), vAct = colAct.getValues();
  Object.keys(res).forEach(row => {
    const i = Number(row) - BOG_FIRST;
    vRes[i][0] = res[row];
    if (res[row].indexOf('✓') === 0) vAct[i][0] = '';
  });
  colRes.setValues(vRes); colAct.setValues(vAct);
  if (!dry) {
    // после действий лист перестраивается и колонка «Результат» очищается — возвращаем результаты по Product ID
    const byPid = {};
    Object.keys(res).forEach(row => { byPid[key_(data[Number(row) - BOG_FIRST][cPid])] = res[row]; });
    refreshActions_();
    writeBogResultsByPid_(byPid);
  }
  const nBelow = add.filter(x => x.below).length;
  return `${dry ? '[ПРОВЕРКА] ' : ''}${toMin ? `[МАРЖА ОТ ${minPct}] ` : ''}добавить: ${add.length}` +
    (nBelow ? ` (из них с маржой ниже ${promoPct}: ${nBelow})` : '') + `, убрать: ${del.length} (акция «${sel.title}»)`;
}

/* ---------- Распределение по самым выгодным акциям ----------
 * Для каждого товара смотрим все акции (кроме ACTIONS_EXCLUDE), где он может участвовать,
 * и выбираем ту, где максимальная цена участия выше всего — то есть скидка минимальная,
 * а маржа лучшая. Товар добавляется в неё и убирается из акций, которые перебили бы эту цену
 * (покупатель видит самую низкую цену из всех акций товара).
 * Если ни одна акция не проходит по «Порогу акций» — товар убирается из акций совсем.
 * Для ручных решений («маржа от 10%») порог в этой акции — «Мин. цена».
 */
function distributeBestActions_() {
  const dry = isDryRun_(), thr = thresholds_(), excl = excludeRe_(), manual = manualPromo_();
  const acts = listActions_().filter(a => !excl.test(a.title || ''));
  if (!acts.length) throw new Error('Нет акций для распределения (проверьте ACTIONS_EXCLUDE)');

  // текущее участие и максимальные цены по всем акциям
  const inAction = {}, maxPrice = {};
  acts.forEach(a => {
    inAction[a.id] = {}; maxPrice[a.id] = {};
    fetchActionProducts_(a.id).forEach(p => {
      inAction[a.id][key_(p.id)] = Number(p.action_price) || 0;
      if (p.max_action_price) maxPrice[a.id][key_(p.id)] = Number(p.max_action_price);
    });
    fetchActionCandidates_(a.id).forEach(p => {
      if (maxPrice[a.id][key_(p.id)] === undefined) maxPrice[a.id][key_(p.id)] = Number(p.max_action_price) || 0;
    });
  });

  const addBy = {}, delBy = {}, note = {}, report = [];
  let moved = 0, kept = 0, dropped = 0, skipped = 0;

  Object.keys(thr).forEach(pid => {
    const t = thr[pid];
    if (!(t.floor > 0)) { skipped++; return; }                       // нет закупа — не трогаем

    // лучшая акция: максимальная цена участия выше всего и не ниже порога
    let best = null;
    acts.forEach(a => {
      const mx = maxPrice[a.id][pid];
      if (!mx || mx < floorIn_(t, manual, a.id, pid)) return;
      if (!best || mx > best.price) best = { action: a, price: mx };
    });

    const participates = acts.filter(a => inAction[a.id][pid] !== undefined);

    if (!best) {                                                     // ничего не проходит по марже
      participates.forEach(a => { (delBy[a.id] = delBy[a.id] || []).push({ pid, art: t.art }); });
      if (participates.length) { dropped++; note[pid] = `✗ ни одна акция не проходит по порогу ${t.floor} — убираем из ${participates.length}`; }
      return;
    }

    const cur = inAction[best.action.id][pid];
    if (cur === undefined || Math.abs(cur - best.price) > 0.5) {
      const o = { product_id: Number(pid), action_price: best.price };
      if (/STOCK/i.test(best.action.action_type || '')) o.stock = Math.max(1, t.stock);
      (addBy[best.action.id] = addBy[best.action.id] || []).push({ o, art: t.art });
      moved++;
      note[pid] = `→ «${best.action.title}» по ${best.price}` + (cur !== undefined ? ` (было ${cur})` : '');
    } else { kept++; note[pid] = `✓ уже в «${best.action.title}» по ${cur}`; }

    // убираем из акций, где цена ниже выбранной — иначе она перебьёт выгодную
    participates.forEach(a => {
      if (a.id === best.action.id) return;
      if (inAction[a.id][pid] < best.price) {
        (delBy[a.id] = delBy[a.id] || []).push({ pid, art: t.art });
        note[pid] = (note[pid] || '') + `; убираем из «${a.title}» (${inAction[a.id][pid]})`;
      }
    });
  });

  // выполняем
  let added = 0, removed = 0;
  Object.keys(addBy).forEach(aid => {
    const title = acts.find(a => String(a.id) === String(aid)).title;
    if (dry) { added += addBy[aid].length; report.push(`${title}: добавили бы ${addBy[aid].length}`); return; }
    chunk_(addBy[aid], 1000).forEach(part => {
      const r = ozon_('/v1/actions/products/activate', { action_id: Number(aid), products: part.map(x => x.o) });
      const ok = ((r.result && r.result.product_ids) || []).length, rej = (r.result && r.result.rejected) || [];
      added += ok;
      report.push(`${title}: +${ok}` + (rej.length ? `, отказ ${rej.length}: ${rej.slice(0, 10).map(x => (thr[key_(x.product_id)] || {}).art + ' — ' + x.reason).join('; ')}` : ''));
      rej.forEach(x => note[key_(x.product_id)] = '✗ ' + x.reason);
    });
  });
  Object.keys(delBy).forEach(aid => {
    const title = acts.find(a => String(a.id) === String(aid)).title;
    if (dry) { removed += delBy[aid].length; report.push(`${title}: убрали бы ${delBy[aid].length}`); return; }
    chunk_(delBy[aid], 1000).forEach(part => {
      const r = ozon_('/v1/actions/products/deactivate', { action_id: Number(aid), product_ids: part.map(x => Number(x.pid)) });
      removed += ((r.result && r.result.product_ids) || []).length;
      report.push(`${title}: −${((r.result && r.result.product_ids) || []).length}`);
    });
  });

  // пишем решение по каждому товару в колонку «Результат»
  writeBogNotes_(note, dry);
  log_('Распределение по акциям', dry ? 'DRY' : 'INFO',
    `Лучшая акция найдена для ${moved + kept} товаров, уже стоят верно ${kept}, переставляем ${moved}, ` +
    `убираем совсем ${dropped}, без закупа пропущено ${skipped}\n` + report.join('\n'));
  if (!dry && (added || removed)) refreshActions_();
  return `${dry ? '[ПРОВЕРКА] ' : ''}переставить: ${moved}, уже верно: ${kept}, добавлено: ${added}, убрано: ${removed}. Подробно — «Лог» и колонка «Результат»`;
}

/** Пишет результаты в колонку «Результат» по Product ID (после перестройки листа) */
function writeBogResultsByPid_(byPid) {
  const sh = sheet_(SHEETS.BOG), last = sh.getLastRow();
  if (last < BOG_FIRST || !Object.keys(byPid).length) return;
  const pids = sh.getRange(BOG_FIRST, BOG.PID, last - BOG_FIRST + 1, 1).getValues();
  const rng = sh.getRange(BOG_FIRST, BOG.RESULT, pids.length, 1), v = rng.getValues();
  pids.forEach((r, i) => { const t = byPid[key_(r[0])]; if (t !== undefined) v[i][0] = t; });
  rng.setValues(v);
}

/** Записывает пояснение по каждому товару в колонку «Результат» листа «Бог акций» */
function writeBogNotes_(note, dry) {
  const sh = sheet_(SHEETS.BOG), last = sh.getLastRow();
  if (last < BOG_FIRST) return;
  const pids = sh.getRange(BOG_FIRST, BOG.PID, last - BOG_FIRST + 1, 1).getValues();
  const out = pids.map(r => {
    const t = note[key_(r[0])];
    return [t ? (dry ? 'ПРОВЕРКА: ' + t : t) : ''];
  });
  sh.getRange(BOG_FIRST, BOG.RESULT, out.length, 1).setValues(out);
}


/* ================== 04_prices_discounts.gs ================== */
/** =====================================================================
 *  ВЫГРУЗКА ЦЕН (с листа Ozon) и ЗАЯВКИ НА СКИДКУ
 * ===================================================================== */
function uploadPrices()       { run_('Отправка отмеченных цен', () => uploadPrices_('checked')); }
function uploadAllPrices()    { run_('Выгрузка всех цен', () => uploadPrices_('all')); }
function uploadManualPrices() { run_('Выгрузка ручных цен', () => uploadPrices_('manual')); }
function setPricesToPromoMargin()   { run_('Цены по целевой марже', () => setPricesToMargin_('PROMO_MARGIN', 'Порог акций, ₽')); }
function setPricesToMinMargin()     { run_('Цены по минимальной марже', () => setPricesToMargin_('MIN_MARGIN', 'Мин. цена, ₽')); }
function explainProduct()           { explainProduct_(); }
function pricePromoAndUpload()      { run_('Цены 12% + выгрузка', pricePromoAndUpload_); }
function loadDiscountTasks()        { run_('Заявки на скидку', loadDiscountTasks_); }
function autoDecideDiscounts()      { run_('Заявки: авто-решение', autoDecideDiscounts_); }
function processDiscountDecisions() { run_('Заявки: применить решения', processDiscountDecisions_); }

/** Выставить всем маржу PROMO_MARGIN, вывести из неподходящих акций и сразу отправить цены в Ozon */
function pricePromoAndUpload_() {
  const step1 = setPricesToMargin_('PROMO_MARGIN', 'Порог акций, ₽');
  SpreadsheetApp.flush();
  const step2 = uploadPrices_('all');
  return step1 + ' || ' + step2;
}

/** Разбор расчёта по одному товару: что во что складывается и откуда берётся маржа */
function explainProduct_() {
  const ui = SpreadsheetApp.getUi();
  const q = ui.prompt('Разбор товара', 'Введите артикул (как в колонке «Артикул»)', ui.ButtonSet.OK_CANCEL);
  if (q.getSelectedButton() !== ui.Button.OK) return;
  const art = String(q.getResponseText()).trim().toLowerCase();
  const m = readMain_();
  const r = m.rows.find(x => String(x['Артикул']).trim().toLowerCase() === art);
  if (!r) { ui.alert('Товар не найден: ' + art); return; }

  const acq = Number(cfg_('ACQUIRING_RATE', 0.01)), pack = Number(cfg_('PACKAGING_RUB', 20));
  const defBuyout = Number(cfg_('DEFAULT_BUYOUT', 0.92));
  const logMode = String(cfg_('LOGISTICS_MODE', 'MAX')).toUpperCase();
  const promo = Number(cfg_('PROMO_MARGIN', 0.12));
  const tariff = {}; readTable_(SHEETS.TARIFFS).rows.forEach(t => tariff[key_(t['product_id'])] = t);
  const t = tariff[key_(r['Product ID'])] || {};

  const cost = Number(r['Закуп, ₽']) || 0;
  const comm = (Number(t['Комиссия FBS, %']) || 0) / 100;
  const lmin = Number(t['Логистика FBS мин, ₽']) || 0, lmax = Number(t['Логистика FBS макс, ₽']) || 0;
  const logi = logMode === 'MIN' ? lmin : logMode === 'AVG' ? (lmin + lmax) / 2 : lmax;
  const proc = Number(t['Обработка FBS, ₽']) || 0, last = Number(t['Последняя миля FBS, ₽']) || 0;
  const ret = (Number(t['Обработка возврата FBS, ₽']) || 0) + (Number(t['Обратная логистика FBS, ₽']) || 0);
  const buy = Number(r['Выкуп, %']) || defBuyout;
  const delivery = (logi + proc + last) / buy + (1 / buy - 1) * ret;
  const fix = cost + delivery + pack;
  const target = priceForMargin_(r, t, promo, acq, pack, defBuyout, logMode);

  const sale = Number(r['Цена продажи, ₽']) || 0;
  const actionPrice = Number(r['Мин. цена в акциях, ₽']) || 0;
  const effective = (actionPrice > 0 && actionPrice < sale) ? actionPrice : sale;
  const marginAt = p => p > 0 ? (p - p * (comm + acq) - fix) / p : 0;
  const pc = x => Math.round(x * 1000) / 10 + '%';

  const txt = [
    `Товар: ${r['Артикул']} (Product ID ${r['Product ID']})`,
    `Закуп: ${Math.round(cost)} ₽`,
    `Логистика ${Math.round(logi)} + обработка ${proc} + последняя миля ${last}, возврат ${ret}, выкуп ${pc(buy)} → доставка на проданную штуку ${Math.round(delivery)} ₽`,
    `Упаковка: ${pack} ₽ → всего постоянных затрат ${Math.round(fix)} ₽`,
    `Комиссия ${pc(comm)}, эквайринг ${pc(acq)}`,
    `Цена под маржу ${pc(promo)}: ${Math.round(target)} ₽`,
    `Цена продажи сейчас: ${Math.round(sale)} ₽ → маржа ${pc(marginAt(sale))}`,
    actionPrice ? `Цена в акции: ${Math.round(actionPrice)} ₽ → маржа ${pc(marginAt(actionPrice))}` : 'В акциях не участвует',
    `Считается от цены ${Math.round(effective)} ₽ → итоговая маржа ${pc(marginAt(effective))}`,
    actionPrice && actionPrice < target
      ? `ПРИЧИНА НИЗКОЙ МАРЖИ: цена акции ${Math.round(actionPrice)} ₽ ниже нужной ${Math.round(target)} ₽. Цена продажи тут не поможет — товар надо убрать из акции.`
      : 'Акции маржу не занижают.'
  ].join('\n');
  log_('Разбор товара', 'INFO', txt);
  ui.alert('Разбор ' + r['Артикул'], txt, ui.ButtonSet.OK);
}

/**
 * Проставляет всем товарам цену продажи, дающую заданную маржу.
 * Если колонка с расчётом пустая или в ошибке, цена считается прямо в скрипте
 * по тем же данным: закуп, тарифы Ozon, выкуп, упаковка, эквайринг.
 * Строки с отметкой «Цена вручную» не трогаются.
 */
function setPricesToMargin_(cfgKey, sourceCol) {
  const m = readMain_();
  const margin = Number(cfg_(cfgKey, 0.12));
  const pct = Math.round(margin * 1000) / 10;
  const acq = Number(cfg_('ACQUIRING_RATE', 0.01));
  const pack = Number(cfg_('PACKAGING_RUB', 20));
  const defBuyout = Number(cfg_('DEFAULT_BUYOUT', 0.92));
  const logMode = String(cfg_('LOGISTICS_MODE', 'MAX')).toUpperCase();
  const hasManual = m.h.indexOf('Цена вручную') >= 0;

  // тарифы по product_id — на случай, если формулы не посчитаны
  const tariff = {};
  readTable_(SHEETS.TARIFFS).rows.forEach(t => tariff[key_(t['product_id'])] = t);

  const patch = {}, res = {};
  let manual = 0, skipped = 0;
  m.rows.forEach(r => {
    if (hasManual && r['Цена вручную'] === true) { manual++; return; }
    let target = Number(r[sourceCol]);
    if (!(target > 0)) target = priceForMargin_(r, tariff[key_(r['Product ID'])], margin, acq, pack, defBuyout, logMode);
    if (!(target > 0)) { skipped++; return; }
    patch[r._row] = Math.ceil(target);
    const act = Number(r['Мин. цена в акциях, ₽']) || 0;
    res[r._row] = act > 0 && act < target
      ? `цена по марже ${pct}% = ${Math.round(target)}; в акции было ${Math.round(act)} — выводим из акции`
      : `цена по марже ${pct}%`;
  });
  if (!Object.keys(patch).length) return `нечего менять: нет закупа или тарифов (пропущено ${skipped}, вручную ${manual})`;
  mainPatch_(m, 'Цена продажи, ₽', patch);
  mainPatch_(m, 'Результат', res);
  log_('Цены по марже', 'INFO', `Выставлено ${Object.keys(patch).length} цен по марже ${pct}%. Вручную пропущено: ${manual}, без расчёта: ${skipped}`);

  // товары, которые сидят в акциях ниже этой маржи, сразу выводим из акций
  SpreadsheetApp.flush();
  let cleanup = '';
  try { cleanup = ' | из акций: ' + removeIneligible_(); }
  catch (e) { cleanup = ' | чистку акций выполнить не удалось: ' + e.message; log_('Цены по марже', 'WARN', e.stack || e); }

  return `выставлено: ${Object.keys(patch).length}, вручную пропущено: ${manual}, без закупа: ${skipped}${cleanup}`;
}

/** Цена, дающая нужную маржу: считаем так же, как формулы листа */
function priceForMargin_(row, t, margin, acq, pack, defBuyout, logMode) {
  const cost = Number(row['Закуп, ₽']);
  if (!(cost > 0) || !t) return 0;
  const comm = (Number(t['Комиссия FBS, %']) || 0) / 100;
  const lmin = Number(t['Логистика FBS мин, ₽']) || 0, lmax = Number(t['Логистика FBS макс, ₽']) || 0;
  const logi = logMode === 'MIN' ? lmin : logMode === 'AVG' ? (lmin + lmax) / 2 : lmax;
  const proc = Number(t['Обработка FBS, ₽']) || 0, last = Number(t['Последняя миля FBS, ₽']) || 0;
  const ret = (Number(t['Обработка возврата FBS, ₽']) || 0) + (Number(t['Обратная логистика FBS, ₽']) || 0);
  const buy = Number(row['Выкуп, %']) || defBuyout;
  const delivery = (logi + proc + last) / buy + (1 / buy - 1) * ret;
  const fix = cost + delivery + pack;
  const den = 1 - margin - comm - acq;
  return den > 0 ? fix / den : 0;
}


/* ---------- Отправка цен в Ozon ----------
 * mode: 'checked' — только строки с галочкой «Отправить»
 *       'all'     — все товары, у которых заполнена цена продажи
 *       'manual'  — только строки с галочкой «Цена вручную»
 */
function uploadPrices_(mode) {
  const m = readMain_(), dry = isDryRun_();
  const maxChg = Number(cfg_('MAX_PRICE_CHANGE', 0.3));
  const res = {}, batch = [];
  const pick = r => mode === 'all' ? Number(r['Цена к отправке, ₽']) > 0
    : mode === 'manual' ? r['Цена вручную'] === true
    : r['Отправить'] === true;

  m.rows.filter(pick).forEach(r => {
    const p = {
      price: Math.round(Number(r['Цена к отправке, ₽'])),
      old: Math.round(Number(r['Зачёркнутая, ₽']) || 0),
      min: Math.round(Number(r['min_price, ₽']) || 0),
      cur: Number(r['Цена на Ozon, ₽']) || 0
    };
    const err = validatePrice_(p, maxChg);
    if (err) { res[r._row] = '✗ ' + err; return; }
    batch.push({ row: r._row, item: {
      offer_id: String(r['Артикул']), price: String(p.price), old_price: String(p.old),
      min_price: String(p.min), currency_code: 'RUB' } });
  });

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM HH:mm');
  if (dry) {
    batch.forEach(b => res[b.row] = `ПРОВЕРКА: уйдёт ${b.item.price} / зачёркн. ${b.item.old_price} / min ${b.item.min_price}`);
  } else {
    chunk_(batch, 1000).forEach(part => {
      const r = ozon_('/v1/product/import/prices', { prices: part.map(b => b.item) });
      const by = {}; (r.result || []).forEach(x => by[String(x.offer_id)] = x);
      part.forEach(b => {
        const x = by[b.item.offer_id];
        res[b.row] = !x ? '✗ нет ответа от Ozon'
          : (x.errors && x.errors.length) ? '✗ ' + x.errors.map(e => e.message || e.code).join('; ')
          : `✓ ${b.item.price} — ${stamp}`;
      });
    });
  }

  mainPatch_(m, 'Результат', res);
  const off = {};
  Object.keys(res).forEach(row => { if (res[row].charAt(0) === '✓') off[row] = false; });
  if (m.h.indexOf('Отправить') >= 0) mainPatch_(m, 'Отправить', off);
  if (!dry && Object.keys(off).length) { Utilities.sleep(2000); syncTariffs_(); }

  const errs = Object.keys(res).filter(k => res[k].charAt(0) === '✗').length;
  const what = mode === 'all' ? 'все цены' : mode === 'manual' ? 'ручные цены' : 'отмеченные цены';
  return `${dry ? '[ПРОВЕРКА, в Ozon ничего не ушло] ' : ''}${what}: взято ${Object.keys(res).length}, ` +
         `отправлено ${Object.keys(off).length}, отклонено ${errs}. Причины отказов — в колонке «Результат»`;
}

/** Проверки перед отправкой: возвращает текст ошибки или ''. Ниже мин. цены (MIN_MARGIN) — никогда. */
function validatePrice_(p, maxChg) {
  if (!(p.price > 0)) return 'нет цены продажи';
  if (!p.min) return 'нет мин. цены (нет закупа?)';
  if (p.price < p.min) return `ниже мин. цены ${p.min}`;
  if (p.cur && maxChg && Math.abs(p.price / p.cur - 1) > maxChg)
    return `изменение ${Math.round((p.price / p.cur - 1) * 100)}% больше лимита MAX_PRICE_CHANGE`;
  if (p.old) {   // правило Ozon для зачёркнутой цены
    const need = p.price < 400 ? 20 : p.price <= 10000 ? p.price * 0.05 : 500;
    if (p.old - p.price < need) return `зачёркнутая ${p.old} слишком близко к цене`;
  }
  return '';
}

/* ---------- Заявки на скидку (решения привязаны к id, не к строке) ---------- */
function loadDiscountTasks_() {
  const keep = {}; readTable_(SHEETS.DISCOUNTS).rows.forEach(r => keep[String(r['id'])] = r);
  const tasks = [];
  ['NEW', 'SEEN'].forEach(status => {
    for (let page = 1; page <= 200; page++) {
      const part = ozon_('/v1/actions/discounts-task/list', { status, page, limit: 50 }).result || [];
      part.forEach(x => tasks.push(x)); if (part.length < 50) break;
    }
  });
  writeTable_(SHEETS.DISCOUNTS, tasks.map(x => { const o = keep[String(x.id)] || {}; return {
    'id': String(x.id), 'Создана': date_(x.created_at), 'Статус': x.status, 'Артикул': x.offer_id, 'Покупатель': x.customer_name,
    'Запрошенная цена': x.requested_price, 'Исходная цена': x.original_price, 'Скидка, %': x.discount_percent,
    'Кол-во мин': x.requested_quantity_min, 'Кол-во макс': x.requested_quantity_max,
    'Решение': o['Решение'] || '', 'Одобренная цена': o['Одобренная цена'] || '', 'Результат': o['Результат'] || '' }; }));
  return `активных заявок: ${tasks.length}`;
}
/** Как старый DiscountsTrigger: загрузить → одобрить всё ≥ порога, отклонить остальное → отправить */
function autoDecideDiscounts_() {
  loadDiscountTasks_(); SpreadsheetApp.flush();
  const t = readTable_(SHEETS.DISCOUNTS), patch = {};
  t.rows.forEach(r => { if (r['Решение']) return; const rec = String(r['Рекомендация'] || '');
    if (rec.startsWith('✓')) patch[r._row] = 'Одобрить'; else if (rec.startsWith('✗')) patch[r._row] = 'Отклонить'; });
  patchColumn_(t.sh, t.h, 'Решение', patch);
  return processDiscountDecisions_();
}
function processDiscountDecisions_() {
  const t = readTable_(SHEETS.DISCOUNTS), dry = isDryRun_();
  const approve = [], decline = [], res = {}, rowById = {};
  t.rows.forEach(r => {
    const d = String(r['Решение'] || '').trim(); if (!d) return;
    const id = String(r['id']); rowById[id] = r._row;
    if (d === 'Отклонить') { decline.push({ id }); return; }
    if (d !== 'Одобрить') return;
    const price = Math.round(Number(r['Одобренная цена']) || Number(r['Запрошенная цена']));
    const floor = Number(r['Порог акций, ₽']) || 0;
    if (!floor) { res[r._row] = '✗ нет порога (нет закупа)'; return; }
    if (price < floor) { res[r._row] = `✗ ниже порога ${floor}`; return; }
    approve.push({ id, approved_price: price, approved_quantity_min: Number(r['Кол-во мин']) || 1, approved_quantity_max: Number(r['Кол-во макс']) || 1 });
  });
  const send = (path, list, txt) => chunk_(list, 50).forEach(part => {
    if (dry) { part.forEach(x => res[rowById[x.id]] = 'ПРОВЕРКА: ' + txt); return; }
    const r = ozon_(path, { tasks: part }), fails = {};
    ((r.result && r.result.fail_details) || []).forEach(f => fails[String(f.task_id)] = f.error_for_user);
    part.forEach(x => res[rowById[x.id]] = fails[x.id] ? '✗ ' + fails[x.id] : '✓ ' + txt);
  });
  send('/v1/actions/discounts-task/approve', approve, 'одобрено');
  send('/v1/actions/discounts-task/decline', decline, 'отклонено');
  patchColumn_(t.sh, t.h, 'Результат', res);
  const done = {}; Object.keys(res).forEach(row => { if (res[row].startsWith('✓')) done[row] = ''; });
  patchColumn_(t.sh, t.h, 'Решение', done);
  return `${dry ? '[ПРОВЕРКА] ' : ''}одобрить: ${approve.length}, отклонить: ${decline.length}`;
}


/* ================== 05_menu.gs ================== */
/** =====================================================================
 *  МЕНЮ, ЕЖЕДНЕВНЫЙ ЗАПУСК, ОФОРМЛЕНИЕ
 *  Четыре отдельных меню в строке меню таблицы: ЦЕНЫ, АКЦИИ, ОСТАТКИ, НАСТРОЙКИ.
 * ===================================================================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // проценты и интервал в подписях берём из «Настроек», чтобы меню не расходилось с ними
  const promo = pctCfg_('PROMO_MARGIN', 0.12), min = pctCfg_('MIN_MARGIN', 0.10);
  let stockMin = 5;
  try { stockMin = stockRefreshMin_(); } catch (e) {}

  ui.createMenu('💰 ЦЕНЫ')
    .addItem(`Выгрузить цены по марже ${promo}`, 'pricePromoAndUpload')
    .addItem(`Посчитать цены по марже ${promo} (без отправки)`, 'setPricesToPromoMargin')
    .addSubMenu(ui.createMenu('Отправить в Ozon')
      .addItem('Отмеченные галочкой', 'uploadPrices')
      .addItem('Заполненные вручную', 'uploadManualPrices'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Обновить данные')
      .addItem('Тарифы и цены Ozon', 'syncTariffs')
      .addItem('Закуп', 'importCosts')
      .addItem('Новые товары из Ozon', 'addMissingProducts')
      .addItem('Источники закупа из старой таблицы', 'sourcesFromOldTable'))
    .addSubMenu(ui.createMenu('Отчёты')
      .addItem('Разобрать расчёт по товару', 'explainProduct')
      .addItem('Аудит: сверить расчёт с Ozon', 'runAudit')
      .addItem('План-факт по финотчёту', 'syncFinance')
      .addItem('Выкуп и возвраты', 'syncBuyout'))
    .addToUi();

  ui.createMenu('🏷 АКЦИИ')
    .addItem('Обновить лист «Бог акций»', 'refreshActions')
    .addSeparator()
    .addItem(`Применить отметки «Действие» (маржа от ${promo})`, 'applySelectedActions')
    .addItem(`Применить отметки с маржой от ${min}`, 'applySelectedMin')
    .addItem('Удалить из акции выделенные и отмеченные', 'removeMarkedFromAction')
    .addSeparator()
    .addSubMenu(ui.createMenu('Массово по всем акциям')
      .addItem('Добавить все подходящие по марже', 'addEligibleToAll')
      .addItem('Убрать неподходящие по марже', 'removeIneligible')
      .addItem('Распределить по самым выгодным акциям', 'distributeBestActions'))
    .addSubMenu(ui.createMenu('Заявки на скидку')
      .addItem('Загрузить', 'loadDiscountTasks')
      .addItem('Решить автоматически', 'autoDecideDiscounts')
      .addItem('Применить мои решения', 'processDiscountDecisions'))
    .addToUi();

  ui.createMenu('📦 ОСТАТКИ')
    .addItem('Обновить остатки сейчас', 'syncStocks')
    .addSubMenu(ui.createMenu('Автообновление')
      .addItem(`Включить автообновление (каждые ${stockMin} мин)`, 'installStockTrigger')
      .addItem('Выключить автообновление', 'stopStockTrigger'))
    .addItem('Вернуть остатки, обнулённые аудитом', 'restoreStocks')
    .addToUi();

  ui.createMenu('⚙ НАСТРОЙКИ')
    .addItem('▶ Обновить все данные', 'syncAll')
    .addItem('Режим проверки (вкл/выкл)', 'toggleDryRun')
    .addItem('Скрыть / показать служебные колонки', 'toggleView')
    .addSeparator()
    .addSubMenu(ui.createMenu('Подключение и структура')
      .addItem('Ключи API', 'setupCredentials')
      .addItem('Проверить подключение', 'testConnection')
      .addItem('Обновить структуру таблицы', 'applySchema'))
    .addSubMenu(ui.createMenu('Автозапуски')
      .addItem('Ночное обновление акций (00:05 и 01:00 МСК)', 'installActionTriggers')
      .addItem('Ежедневное обновление (6:00)', 'installDailyTrigger')
      .addItem('Выключить все автозапуски', 'removeAllTriggers'))
    .addToUi();
}

/** Одна кнопка вместо двух: прячет служебные колонки, если они видны, и показывает, если спрятаны */
function toggleView() {
  const sh = sheet_(SHEETS.MAIN), h = headersAt_(sh, MAIN_HDR_ROW);
  const i = MAIN_HIDE.map(n => h.indexOf(n)).find(x => x >= 0);
  const hidden = i !== undefined && sh.isColumnHiddenByUser(i + 1);
  run_(hidden ? 'Показать все колонки' : 'Компактный вид', () => toggleColumns_(!hidden));
}

/** Режим проверки DRY_RUN: показывает, что включено сейчас, и переключает после подтверждения */
function toggleDryRun() {
  const ui = SpreadsheetApp.getUi(), dry = isDryRun_();
  const a = ui.alert(dry ? 'Режим проверки сейчас ВКЛЮЧЁН' : 'Режим проверки сейчас ВЫКЛЮЧЕН',
    dry ? 'В Ozon ничего не уходит. Выключить режим проверки? Цены, акции, остатки и заявки начнут по-настоящему уходить в Ozon.'
        : 'Изменения уходят в Ozon. Включить режим проверки? Тогда в Ozon ничего уходить не будет, только проверка.',
    ui.ButtonSet.YES_NO);
  if (a !== ui.Button.YES) return;
  setCfg_('DRY_RUN', !dry);
  const msg = dry ? 'выключен — изменения уходят в Ozon' : 'включён — в Ozon ничего не уходит';
  log_('Режим проверки', 'INFO', msg);
  toast_('Режим проверки ' + msg, 10);
}

/** Выключает все автозапуски разом */
function removeAllTriggers() {
  const n = ScriptApp.getProjectTriggers().length;
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  toast_(`Автозапуски выключены: ${n}`, 8);
}

function syncAll() {
  run_('Обновить всё', () => {
    const out = [];
    const step = (name, fn) => { try { out.push(`${name}: ${fn()}`); } catch (e) { out.push(`${name}: ОШИБКА ${e.message}`); log_(name, 'ERROR', e.stack || e); } };
    step('Курс', updateUsdRate_);
    step('Тарифы', syncTariffs_);
    step('Остатки', syncStocks_);
    step('Закуп', importCosts_);
    step('Заказы', syncOrders60_);
    step('Выкуп', syncBuyout_);
    step('План-факт', syncFinance_);
    SpreadsheetApp.flush();
    if (cfg_('AUTO_REMOVE_FROM_ACTIONS', false) === true) step('Чистка акций', removeIneligible_);
    step('Бог акций', refreshActions_);
    SpreadsheetApp.flush();
    step('Аудит', runAudit_);
    return out.join(' | ');
  });
}

function testConnection() {
  run_('Проверка подключения', () => `ок, товаров в кабинете: ${ozon_('/v3/product/list', { filter: { visibility: 'ALL' }, limit: 1 }).result.total}`);
}

/* ---------- Ночное обновление акций ---------- */
/** Смещение часового пояса скрипта в минутах (для пересчёта московского времени) */
function tzOffsetMin_(tz) {
  const z = Utilities.formatDate(new Date(), tz, 'Z');            // например +0700
  const sign = z.charAt(0) === '-' ? -1 : 1;
  return sign * (Number(z.substr(1, 2)) * 60 + Number(z.substr(3, 2)));
}
/** Московское время → местное время скрипта */
function mskToLocal_(h, m) {
  const off = tzOffsetMin_(Session.getScriptTimeZone()) - 180;    // Москва = UTC+3
  let t = (h * 60 + m + off + 1440) % 1440;
  return { h: Math.floor(t / 60), m: t % 60 };
}

/** Ночной прогон: новые акции + чистка тех, кто не проходит по марже */
function actionsNightly() {
  run_('Ночное обновление акций', () => {
    const out = [];
    try { out.push('акции: ' + refreshActions_()); }
    catch (e) { out.push('акции: ОШИБКА ' + e.message); log_('Ночное обновление акций', 'ERROR', e.stack || e); }
    if (cfg_('AUTO_REMOVE_FROM_ACTIONS', true) === true) {
      try { out.push('чистка: ' + removeIneligible_()); }
      catch (e) { out.push('чистка: ОШИБКА ' + e.message); log_('Ночное обновление акций', 'ERROR', e.stack || e); }
    }
    return out.join(' | ');
  });
}

function installActionTriggers() {
  removeActionTriggers();
  const times = [[0, 5], [1, 0]];                                  // по московскому времени
  const made = times.map(([h, m]) => {
    const loc = mskToLocal_(h, m);
    ScriptApp.newTrigger('actionsNightly').timeBased().everyDays(1).atHour(loc.h).nearMinute(loc.m).create();
    return `${h}:${m < 10 ? '0' + m : m} МСК (местное ${loc.h}:${loc.m < 10 ? '0' + loc.m : loc.m})`;
  });
  toast_('Ночное обновление акций включено: ' + made.join(', '), 10);
  log_('Триггеры', 'INFO', 'Ночное обновление акций: ' + made.join(', ') +
    '. Google запускает такие триггеры с точностью примерно ±15 минут.');
}
function removeActionTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'actionsNightly').forEach(t => ScriptApp.deleteTrigger(t));
}

/* ---------- Автообновление остатков ----------
 * Ozon ставит товар в резерв сразу при заказе, поэтому «доступно = present − reserved»
 * меняется в тот же момент. Таймер просто почаще перечитывает это значение.
 */
function stocksTick() {
  try { syncStocks_(); }
  catch (e) { log_('Остатки (авто)', 'ERROR', e.message); }
}

/** Интервал автообновления из «Настроек»; Google принимает только 1, 5, 10, 15 и 30 минут */
function stockRefreshMin_() {
  const mins = Number(cfg_('STOCKS_REFRESH_MIN', 5));
  return [1, 5, 10, 15, 30].indexOf(mins) >= 0 ? mins : 5;
}
function installStockTrigger() {
  removeStockTrigger();
  const mins = stockRefreshMin_();
  ScriptApp.newTrigger('stocksTick').timeBased().everyMinutes(mins).create();
  log_('Триггеры', 'INFO', `Остатки обновляются каждые ${mins} мин.`);
  toast_(`Остатки будут обновляться каждые ${mins} мин.`, 8);
}
function stopStockTrigger() {
  removeStockTrigger();
  log_('Триггеры', 'INFO', 'Автообновление остатков выключено');
  toast_('Автообновление остатков выключено', 8);
}
function removeStockTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'stocksTick').forEach(t => ScriptApp.deleteTrigger(t));
}

function installDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger('syncAll').timeBased().everyDays(1).atHour(6).create();
  toast_('Ежедневное обновление включено (~6:00)');
}
function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncAll').forEach(t => ScriptApp.deleteTrigger(t));
}


/* ================== 06_buyout_finance.gs ================== */
/** =====================================================================
 *  ВЫКУП, ВОЗВРАТЫ И ПЛАН-ФАКТ
 *  Выкуп → колонки «Выкуп, %» и «Выкуп: основа» листа Ozon → «Логистика с выкупом» → цена.
 *  План-факт → лист «План-факт» из финансового отчёта Ozon (реальные удержания).
 * ===================================================================== */
function syncBuyout()  { run_('Выкуп и возвраты', syncBuyout_); }
function syncFinance() { run_('План-факт', syncFinance_); }

/** Интервалы [from, to] не длиннее stepDays — у Ozon есть лимиты на период запроса */
function periods_(days, stepDays) {
  const out = [], end = new Date();
  let from = new Date(end.getTime() - days * 864e5);
  while (from < end) {
    const to = new Date(Math.min(from.getTime() + stepDays * 864e5 - 1000, end.getTime()));
    out.push([from, to]); from = new Date(to.getTime() + 1000);
  }
  return out;
}

/* ---------- Выкуп ----------
 * Выкуп = (доставлено − вернули после покупки) / (доставлено + не выкупили).
 * Доставлено — отправления FBS в статусе delivered; невыкуп/отмены после отгрузки и возвраты — /v1/returns/list.
 * Мало отгрузок у товара → выкуп его категории → DEFAULT_BUYOUT.
 */
function syncBuyout_() {
  const days = Number(cfg_('BUYOUT_DAYS', 180)), minShip = Number(cfg_('BUYOUT_MIN_SHIPMENTS', 5));
  const delivered = {}, notPicked = {}, returned = {};
  const add = (map, k, q) => { if (k) map[k] = (map[k] || 0) + (Number(q) || 1); };

  periods_(days, 30).forEach(([from, to]) => {
    for (let offset = 0; offset < 200000; offset += 1000) {
      const r = ozon_('/v3/posting/fbs/list', { dir: 'ASC', filter: { since: from.toISOString(), to: to.toISOString(), status: 'delivered' }, limit: 1000, offset });
      const part = (r.result && r.result.postings) || [];
      part.forEach(p => (p.products || []).forEach(pr => add(delivered, key_(pr.sku), pr.quantity)));
      if (!r.result || !r.result.has_next || !part.length) break;
    }
    let lastId = 0;
    for (let page = 0; page < 500; page++) {
      const r = ozon_('/v1/returns/list', { filter: { logistic_return_date: { time_from: from.toISOString(), time_to: to.toISOString() } }, limit: 500, last_id: lastId });
      const list = r.returns || [];
      list.forEach(x => {
        if (String(x.schema || '').toLowerCase() === 'fbo') return;        // работаем только по FBS
        const sku = key_(x.product && x.product.sku), q = x.product && x.product.quantity;
        if (String(x.type) === 'Cancellation') add(notPicked, sku, q);      // не забрали / отмена после отгрузки
        else add(returned, sku, q);                                         // вернули после покупки
      });
      if (!r.has_next || !list.length) break;
      lastId = list[list.length - 1].id;
    }
  });

  const m = readMain_();
  const calc = (d, n, rt) => (d + n) > 0 ? Math.max(0, Math.min(1, (d - rt) / (d + n))) : null;
  // выкуп по категориям
  const cat = {};
  m.rows.forEach(r => {
    const k = key_(r['SKU']), c = String(r['Категория'] || '').trim() || NEW_BLOCK;
    cat[c] = cat[c] || { d: 0, n: 0, rt: 0 };
    cat[c].d += delivered[k] || 0; cat[c].n += notPicked[k] || 0; cat[c].rt += returned[k] || 0;
  });
  const def = Number(cfg_('DEFAULT_BUYOUT', 0.92));
  const pct = {}, basis = {};
  let totD = 0, totN = 0, totR = 0;
  m.rows.forEach(r => {
    const k = key_(r['SKU']), d = delivered[k] || 0, n = notPicked[k] || 0, rt = returned[k] || 0;
    totD += d; totN += n; totR += rt;
    const c = cat[String(r['Категория'] || '').trim() || NEW_BLOCK];
    if (d + n >= minShip) { pct[r._row] = Math.round(calc(d, n, rt) * 1000) / 1000; basis[r._row] = `товар: ${d + n} отгр.`; }
    else if (c && c.d + c.n >= minShip * 3) { pct[r._row] = Math.round(calc(c.d, c.n, c.rt) * 1000) / 1000; basis[r._row] = 'категория'; }
    else { pct[r._row] = ''; basis[r._row] = 'по умолч.'; }
  });
  mainPatch_(m, 'Выкуп, %', pct); mainPatch_(m, 'Выкуп: основа', basis);
  const shop = calc(totD, totN, totR);
  const msg = `выкуп магазина за ${days} дн.: ${shop === null ? 'нет данных' : Math.round(shop * 1000) / 10 + '%'} (доставлено ${totD}, невыкуп ${totN}, возвраты ${totR}); по умолчанию сейчас ${Math.round(def * 100)}%`;
  log_('Выкуп', 'INFO', msg + '. Если сильно отличается — поправьте DEFAULT_BUYOUT в «Настройках».');
  return msg;
}

/* ---------- План-факт по финансовому отчёту ----------
 * Ozon отключил /v3/finance/transaction/list 06.07.2026.
 * Источник — /v1/finance/accrual/by-day: одна дата за запрос, пагинация по last_id.
 * В ответе три вида начислений:
 *   POSTING  — продажа: выручка (sale_amount), комиссия (sale_commission), логистика (delivery);
 *   ITEM     — сборы по товару (эквайринг, реклама и т.п.);
 *   NON_ITEM — списания без товара (реклама, хранение, штрафы) → строка «БЕЗ ТОВАРА».
 */

/** Справочник типов начислений: type_id → название */
function financeTypes_() {
  const map = {};
  try {
    const r = ozon_('/v1/finance/accrual/types', {});
    const list = r.accrual_types || r.types || r.items || (r.result && (r.result.accrual_types || r.result.types)) || [];
    list.forEach(t => { const k = t.id !== undefined ? t.id : (t.type_id || t.code); if (k !== undefined) map[String(k)] = t.description || t.name || String(k); });
  } catch (e) { log_('План-факт', 'WARN', 'Справочник типов недоступен: ' + e.message); }
  return map;
}

/** {amount:"-12.88",currency:"RUB"} → -12.88 */
function money_(v) { return v ? (Number(String(v.amount).replace(',', '.')) || 0) : 0; }
/** Группы статей по type_id из /v1/finance/accrual/types */
var FEE_GROUPS = {
  back: [2, 6, 9, 40, 45, 53, 59, 60, 65, 71, 78, 79, 102, 103, 113, 115],   // возвраты, невыкупы, обратная логистика
  ads:  [3, 4, 5, 19, 23, 33, 36, 41, 47, 49, 54, 55, 61, 70, 74, 75, 80, 87, 95, 96, 116, 118],  // реклама и продвижение
  fine: [14, 89, 90, 91, 92, 93, 94],                                        // штрафы и операционные ошибки
  log:  [12, 13, 16, 17, 21, 28, 29, 30, 32, 42, 43, 44, 56, 58, 73, 77, 82, 84, 85, 86, 97, 98, 101, 108, 109, 110, 111, 112, 114, 120, 121],
  com:  [69]
};
function feeGroup_(id) {
  const n = Number(id);
  for (const g in FEE_GROUPS) if (FEE_GROUPS[g].indexOf(n) >= 0) return g;
  return 'other';
}
function isReturnType_(types, id) { return feeGroup_(id) === 'back'; }
function aggFor_(agg, sku) {
  const k = key_(sku);
  return agg[k] = agg[k] || { sold: 0, ret: 0, rev: 0, com: 0, log: 0, back: 0, ads: 0, fine: 0, other: 0, pay: 0 };
}

/** Раскладывает одно начисление по товарам */
function applyAccrual_(a, types, agg, nonItem) {
  const total = money_(a.total_amount);
  const cat = String(a.accrued_category || '');

  if (cat === 'POSTING' && a.posting) {
    const prods = a.posting.products || [];
    const n = prods.length || 1;
    prods.forEach(p => {
      const g = aggFor_(agg, p.sku), q = Number(p.quantity) || 1;
      const c = p.commission || {};
      const rev = money_(c.sale_amount), com = money_(c.sale_commission);
      if (rev > 0) g.sold += q; else if (rev < 0) g.ret += q;
      g.rev += rev; g.com += com;
      const d = p.delivery || {};
      (d.services || []).forEach(sv => {
        const v = money_(sv.accrued), gr = feeGroup_(sv.type_id);
        if (gr === 'back') g.back += v; else if (gr === 'ads') g.ads += v; else if (gr === 'fine') g.fine += v; else g.log += v;
      });
      if (!(d.services || []).length) g.log += money_(d.total_accrued);
      g.pay += total / n;
    });
    return;
  }
  if (cat === 'ITEM' && a.item_fees) {
    (a.item_fees.fees || []).forEach(f => {
      const g = aggFor_(agg, f.sku);
      (f.fees || []).forEach(x => {
        const v = money_(x.accrued), gr = feeGroup_(x.type_id);
        if (gr === 'back') g.back += v; else if (gr === 'ads') g.ads += v;
        else if (gr === 'fine') g.fine += v; else if (gr === 'log') g.log += v; else g.other += v;
        g.pay += v;
      });
    });
    return;
  }
  // без привязки к товару: реклама, хранение, штрафы, тара
  const fee = a.non_item_fee || (a.container_fees && a.container_fees.fees && a.container_fees.fees[0]);
  const id = fee ? String(fee.type_id) : '';
  const name = types[id] || (id ? 'тип ' + id : 'прочее');
  nonItem[name] = (nonItem[name] || 0) + (fee ? money_(fee.accrued) : total);
  nonItemGroup_[name] = id ? feeGroup_(id) : 'other';
}

/** название статьи → группа (для раскраски по колонкам) */
var nonItemGroup_ = {};

function syncFinance_() {
  const days = Math.min(Number(cfg_('FACT_DAYS', 90)), 92);
  const types = financeTypes_();
  const tz = Session.getScriptTimeZone();
  const agg = {}, nonItem = {};
  const t0 = Date.now();
  let daysDone = 0, records = 0, stopped = false;

  for (let i = 1; i <= days; i++) {
    if (Date.now() - t0 > 4 * 60000) { stopped = true; break; }          // не упираемся в лимит времени Apps Script
    const day = Utilities.formatDate(new Date(Date.now() - i * 864e5), tz, 'yyyy-MM-dd');
    let lastId = '';
    for (let page = 0; page < 500; page++) {
      const body = { date: day };
      if (lastId) body.last_id = lastId;
      const r = ozon_('/v1/finance/accrual/by-day', body);
      const list = r.accruals || (r.result && r.result.accruals) || [];
      list.forEach(a => { applyAccrual_(a, types, agg, nonItem); records++; });
      lastId = r.last_id || (r.result && r.result.last_id) || '';
      if (!lastId || !list.length) break;
    }
    daysDone++;
  }

  const now = new Date(), R = v => Math.round(v);
  const rows = Object.keys(agg).map(sku => { const a = agg[sku]; return {
    'SKU': Number(sku) || sku, 'Продано, шт': a.sold, 'Возвращено, шт': a.ret, 'Выручка, ₽': R(a.rev),
    'Комиссия, ₽': R(a.com), 'Логистика и услуги, ₽': R(a.log), 'Возвраты, ₽': R(a.back),
    'Реклама, ₽': R(a.ads), 'Штрафы, ₽': R(a.fine), 'Прочее, ₽': R(a.other),
    'К выплате, ₽': R(a.pay), 'Обновлено': now }; })
    .sort((x, y) => y['Выручка, ₽'] - x['Выручка, ₽']);
  // списания без товара — отдельной строкой на каждую статью
  const nonItemSum = Object.keys(nonItem).reduce((s, k) => s + nonItem[k], 0);
  Object.keys(nonItem).sort((a, b) => nonItem[a] - nonItem[b]).forEach(name => {
    const v = R(nonItem[name]), gr = nonItemGroup_[name] || 'other';
    const row = { 'SKU': 'БЕЗ ТОВАРА', 'Название': name, 'К выплате, ₽': v, 'Обновлено': now };
    row[gr === 'ads' ? 'Реклама, ₽' : gr === 'fine' ? 'Штрафы, ₽' : gr === 'back' ? 'Возвраты, ₽' : 'Прочее, ₽'] = v;
    rows.push(row);
  });
  writeTable_('План-факт', rows);

  const tot = rows.reduce((s, r) => ({ rev: s.rev + (r['Выручка, ₽'] || 0), pay: s.pay + (r['К выплате, ₽'] || 0) }), { rev: 0, pay: 0 });
  log_('План-факт', 'INFO', `Дней ${daysDone}${stopped ? ' (прервано по времени, запустите ещё раз с меньшим FACT_DAYS)' : ''}, ` +
    `начислений ${records}, товаров ${rows.length - 1}, выручка ${R(tot.rev)}, к выплате ${R(tot.pay)}.\n` +
    'Без товара: ' + Object.keys(nonItem).map(k => `${k}: ${R(nonItem[k])}`).join('; '));
  return `дней: ${daysDone}, товаров: ${rows.length - 1}, выручка ${R(tot.rev).toLocaleString('ru-RU')} ₽, без товара ${R(nonItemSum).toLocaleString('ru-RU')} ₽`;
}


/* ================== 08_schema.gs ================== */
/** =====================================================================
 *  СТРУКТУРА ТАБЛИЦЫ: колонки, формулы, оформление.
 *  Меняется структура → замените этот файл и нажмите
 *  ⚙ НАСТРОЙКИ → «Обновить структуру таблицы».
 *  Данные не теряются: колонки переставляются вместе со значениями.
 * ===================================================================== */
function applySchema() { run_('Обновление структуры', applySchema_); }
function compactView() { run_('Компактный вид', () => toggleColumns_(true)); }
function fullView()    { run_('Показать все колонки', () => toggleColumns_(false)); }

/* ---------- Палитра: спокойные цвета, акцент только там, где нужно действие ---------- */
var OZ_UI = {
  band: '#ECEFF1', bandText: '#455A64',        // строка 1 — группы колонок
  header: '#455A64', headerText: '#FFFFFF',    // строка 2 — заголовки
  block: '#C7D7EA', blockText: '#0B2545',      // строки-заголовки блоков
  input: '#FFFDE7', inputText: '#1A237E',      // ячейки ручного ввода
  key: '#E8F0FE',                              // ключевые колонки: закуп, остаток, цена, маржа
  bad: '#FCE8E6', badText: '#B3261E',
  warn: '#FEF7E0', warnText: '#8A6100',
  good: '#E6F4EA', goodText: '#1E7B34',
  muted: '#9AA0A6', grid: '#DADCE0', font: 'Arial', size: 10
};

/* ---------- Списки колонок вспомогательных листов ---------- */
var TARIFF_COLS = ['product_id', 'Артикул', 'Цена', 'Цена до скидки', 'Мин. цена', 'Комиссия FBS, %',
  'Логистика FBS мин, ₽', 'Логистика FBS макс, ₽', 'Обработка FBS, ₽', 'Последняя миля FBS, ₽',
  'Мин. цена конкурента на Ozon', 'Мин. цена на других площадках', 'Индекс цены',
  'Обновлено', 'Обработка возврата FBS, ₽', 'Обратная логистика FBS, ₽'];

var PF_COLS = ['SKU', 'Артикул', 'Название', 'Категория', 'Продано, шт', 'Возвращено, шт', 'Выручка, ₽',
  'Комиссия, ₽', 'Логистика и услуги, ₽', 'Возвраты, ₽', 'Реклама, ₽', 'Штрафы, ₽', 'Прочее, ₽', 'К выплате, ₽', 'Удержания факт, %',
  'Удержания план, %', 'Разница, п.п.', 'Закуп проданного, ₽', 'Прибыль факт, ₽', 'Прибыль факт/шт, ₽',
  'Маржа факт, %', 'Прибыль план/шт, ₽', 'Факт − план/шт, ₽', 'Обновлено'];

var DISCOUNT_COLS = ['id', 'Создана', 'Статус', 'Артикул', 'Покупатель', 'Запрошенная цена', 'Исходная цена',
  'Скидка, %', 'Кол-во мин', 'Кол-во макс', 'Порог акций, ₽', 'Рекомендация', 'Решение', 'Одобренная цена', 'Результат'];

var SOURCE_COLS = ['Источник', 'ID таблицы', 'Лист', 'Колонка кода', 'Колонка закупа', 'Колонка РРЦ',
  'Валюта', 'Товаров', 'Комментарий'];

var MAIN_COLS = ['Артикул', 'SKU', 'Product ID', 'Код в прайсе', 'Название', 'Категория', 'Поставщик', 'Источник закупа',
  'Остаток FBS', 'Закуп, ₽', 'РРЦ, ₽', 'Комиссия, %', 'Логистика, ₽',
  'Обработка Ozon, ₽', 'Посл. миля, ₽', 'Выкуп, %', 'Выкуп: основа', 'Возврат, ₽', 'Логистика с выкупом, ₽',
  'Упаковка, ₽', 'Эквайринг, %', 'Затраты фикс., ₽', 'Мин. цена, ₽', 'Порог акций, ₽',
  'Цена продажи, ₽', 'Цена вручную', 'Цена на Ozon, ₽', 'Мин. цена в акциях, ₽', 'В акции', 'Цена факт., ₽', 'Прибыль, ₽', 'Маржа, %',
  'Статус', 'Прибыль факт/шт, ₽', 'Маржа факт, %', 'Конкурент Ozon, ₽', 'Другие площадки, ₽',
  'Разница с рынком', 'Позиция', 'Отправить', 'Цена к отправке, ₽', 'Зачёркнутая, ₽', 'min_price, ₽', 'Результат'];

var SETTINGS_DEFAULTS = [
  ['MIN_MARGIN', 0.10, 'Минимальная маржа: «Мин. цена», min_price в Ozon и стартовая цена новых товаров.'],
  ['PROMO_MARGIN', 0.12, 'Маржа для акций и заявок на скидку. Ниже — не добавляем и убираем из акций (кроме ручных решений от MIN_MARGIN).'],
  ['ACQUIRING_RATE', 0.01, 'Эквайринг, доля от цены.'],
  ['PACKAGING_RUB', 20, 'Своя упаковка и сборка, ₽ на единицу.'],
  ['LOGISTICS_MODE', 'MAX', 'Логистика FBS: MAX, MIN или AVG.'],
  ['OLD_PRICE_ROUND', 100, 'Зачёркнутая цена: округление цены до…'],
  ['OLD_PRICE_ADD', 1600, '…и надбавка, ₽.'],
  ['MAX_PRICE_CHANGE', 0.30, 'Не отправлять цену, если она меняется больше чем на эту долю.'],
  ['DRY_RUN', true, 'TRUE — режим проверки, в Ozon ничего не уходит.'],
  ['FBS_WAREHOUSE', 'ПОРТ', 'FBS-склад для колонки «Остаток FBS».'],
  ['USD_RATE', 84.0954, 'Курс USD, обновляется из меню.'],
  ['COMPETITOR_TOLERANCE', 0.03, 'Разница с конкурентом в пределах ± этого значения = «на уровне».'],
  ['AUTO_REMOVE_FROM_ACTIONS', true, 'Ежедневно убирать из акций товары ниже порога.'],
  ['ACTIONS_EXCLUDE', 'FBO|для складов', 'Акции, которые массовые кнопки пропускают (регулярное выражение).'],
  ['COST_1C_SHEET_ID', '', 'Таблица с себестоимостью из 1С.'],
  ['COST_1C_SHEET', 'Prices', 'Имя листа в таблице 1С.'],
  ['ORDERS_DAYS', 60, 'За сколько дней считать заказы для «Хватит, дней».'],
  ['DEFAULT_BUYOUT', 0.92, 'Выкуп по умолчанию, если по товару мало данных.'],
  ['BUYOUT_DAYS', 180, 'За сколько дней считать выкуп.'],
  ['BUYOUT_MIN_SHIPMENTS', 5, 'Меньше отгрузок — берём выкуп категории.'],
  ['FACT_DAYS', 90, 'Период план-факта (финотчёт Ozon).'],
  ['AUDIT_TOLERANCE', 0.02, 'Допустимое расхождение маржи и комиссии с Ozon (2 п.п.).'],
  ['AUDIT_REMOVE_FROM_ACTIONS', true, 'Убирать из акций товары, у которых маржа по живым данным Ozon ниже PROMO_MARGIN.'],
  ['AUDIT_ZERO_STOCK', false, 'ОПАСНО: обнулять остаток у товаров, которые по данным Ozon продаются в убыток.'],
  ['AUDIT_MAX_ZERO', 10, 'Максимум обнулений остатка за один запуск аудита.'],
  ['AUDIT_MAX_REMOVE', 50, 'Максимум снятий с акций за один запуск аудита.'],
  ['STOCKS_REFRESH_MIN', 5, 'Как часто автоматически обновлять остатки, минут. Допустимо: 1, 5, 10, 15, 30.']
];
// ключи, которые скрипт больше не читает: «Обновить структуру таблицы» убирает их из «Настроек»
var SETTINGS_OBSOLETE = ['ALLOW_BELOW_MIN', 'OLD_TABLE_ID'];

var MAIN_WIDTHS = { 'В акции': 60, 'Артикул': 150, 'Название': 300, 'Категория': 110, 'Поставщик': 100, 'Источник закупа': 120,
  'Статус': 105, 'Позиция': 100, 'Выкуп: основа': 110, 'Результат': 210 };
var MAIN_PCT = ['Комиссия, %', 'Выкуп, %', 'Эквайринг, %', 'Маржа, %', 'Маржа факт, %', 'Разница с рынком'];
var MAIN_MONEY = ['Закуп, ₽', 'РРЦ, ₽', 'Логистика, ₽', 'Обработка Ozon, ₽', 'Посл. миля, ₽', 'Возврат, ₽',
  'Логистика с выкупом, ₽', 'Упаковка, ₽', 'Затраты фикс., ₽', 'Мин. цена, ₽', 'Порог акций, ₽',
  'Цена продажи, ₽', 'Цена вручную', 'Цена на Ozon, ₽', 'Мин. цена в акциях, ₽', 'Цена факт., ₽', 'Прибыль, ₽',
  'Прибыль факт/шт, ₽', 'Конкурент Ozon, ₽', 'Другие площадки, ₽', 'Цена к отправке, ₽', 'Зачёркнутая, ₽', 'min_price, ₽'];
var MAIN_INPUTS = ['Код в прайсе', 'Категория', 'Поставщик', 'Источник закупа', 'Цена продажи, ₽', 'Цена вручную', 'Отправить'];
var MAIN_HIDE = ['SKU', 'Product ID', 'Код в прайсе', 'Поставщик', 'Источник закупа', 'РРЦ, ₽', 'Комиссия, %',
  'Логистика, ₽', 'Обработка Ozon, ₽', 'Посл. миля, ₽', 'Выкуп, %', 'Выкуп: основа', 'Возврат, ₽',
  'Логистика с выкупом, ₽', 'Упаковка, ₽', 'Эквайринг, %', 'Мин. цена в акциях, ₽',
  'Конкурент Ozon, ₽', 'Другие площадки, ₽', 'Цена к отправке, ₽', 'Зачёркнутая, ₽', 'min_price, ₽'];
// колонки, которые нужно удалить из листов при обновлении структуры
var DROP_COLS = { 'Ozon': ['Комиссия факт, %', 'Δ комиссия, п.п.', 'Ozon-карта, %'], 'План-факт': ['Комиссия факт, %'],
  'Тарифы': ['Скидка Ozon-карта, %'] };

var MAIN_BANDS = [['Товар', 'Артикул', 'Остаток FBS'], ['Закуп', 'Закуп, ₽', 'РРЦ, ₽'],
  ['Расходы Ozon и свои', 'Комиссия, %', 'Затраты фикс., ₽'], ['Цены', 'Мин. цена, ₽', 'Цена факт., ₽'],
  ['Итог: план и факт', 'Прибыль, ₽', 'Маржа факт, %'], ['Конкуренты', 'Конкурент Ozon, ₽', 'Позиция'],
  ['Выгрузка цен', 'Отправить', 'Результат']];

/* ---------- Хелперы ---------- */
function letter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
function cfgRef_(key) { return `VLOOKUP("${key}",Настройки!$A:$B,2,FALSE)`; }

/** Формулы главного листа в нотации A1 для строки r (h — фактические заголовки листа) */
function mainFormulas_(h, r) {
  const c = n => '$' + letter_(h.indexOf(n) + 1) + r;
  const pidRef = c('Product ID'), skuRef = c('SKU');
  const tc = n => '$' + letter_(TARIFF_COLS.indexOf(n) + 1);
  const t = n => `INDEX(Тарифы!${tc(n)}:${tc(n)},MATCH(${pidRef},Тарифы!$A:$A,0))`;
  const pc = n => '$' + letter_(PF_COLS.indexOf(n) + 1);
  const pf = n => `INDEX('План-факт'!${pc(n)}:${pc(n)},MATCH(${skuRef},'План-факт'!$A:$A,0))`;
  const actFrom = letter_(BOG_ACT_COL), actTo = letter_(BOG_ACT_COL + 60);
  const bog = `MIN(INDEX('Бог акций'!$${actFrom}:$${actTo},MATCH(${pidRef},'Бог акций'!$A:$A,0),0))`;
  const B = `IF(N(${c('Выкуп, %')})>0,${c('Выкуп, %')},${cfgRef_('DEFAULT_BUYOUT')})`;
  const base = `IF(${c('Цена продажи, ₽')}="",${c('Цена на Ozon, ₽')},${c('Цена продажи, ₽')})`;
  const priceAt = m => {
    const den = `(1-${m}-${c('Комиссия, %')}-${c('Эквайринг, %')})`;
    return `=IF(OR(${c('Затраты фикс., ₽')}="",${c('Комиссия, %')}=""),"",IF(${den}<=0,"нереально",ROUNDUP(${c('Затраты фикс., ₽')}/${den},0)))`;
  };
  const F = {};
  F['Комиссия, %'] = `=IFERROR(${t('Комиссия FBS, %')}/100,"")`;
  F['Логистика, ₽'] = `=IFERROR(IF(${cfgRef_('LOGISTICS_MODE')}="MIN",${t('Логистика FBS мин, ₽')},IF(${cfgRef_('LOGISTICS_MODE')}="AVG",(${t('Логистика FBS мин, ₽')}+${t('Логистика FBS макс, ₽')})/2,${t('Логистика FBS макс, ₽')})),"")`;
  F['Обработка Ozon, ₽'] = `=IFERROR(N(${t('Обработка FBS, ₽')}),0)`;
  F['Посл. миля, ₽'] = `=IFERROR(N(${t('Последняя миля FBS, ₽')}),0)`;
  F['Упаковка, ₽'] = `=${cfgRef_('PACKAGING_RUB')}`;
  F['Эквайринг, %'] = `=${cfgRef_('ACQUIRING_RATE')}`;
  F['Возврат, ₽'] = `=IFERROR(N(${t('Обработка возврата FBS, ₽')})+N(${t('Обратная логистика FBS, ₽')}),0)`;
  F['Логистика с выкупом, ₽'] = `=IF(${c('Логистика, ₽')}="","",ROUND((${c('Логистика, ₽')}+${c('Обработка Ozon, ₽')}+${c('Посл. миля, ₽')})/${B}+(1/${B}-1)*${c('Возврат, ₽')},0))`;
  F['Затраты фикс., ₽'] = `=IF(OR(NOT(ISNUMBER(${c('Закуп, ₽')})),${c('Логистика с выкупом, ₽')}=""),"",${c('Закуп, ₽')}+${c('Логистика с выкупом, ₽')}+${c('Упаковка, ₽')})`;
  F['Мин. цена, ₽'] = priceAt(cfgRef_('MIN_MARGIN'));
  F['Порог акций, ₽'] = priceAt(cfgRef_('PROMO_MARGIN'));
  F['Цена на Ozon, ₽'] = `=IFERROR(${t('Цена')},"")`;
  F['Мин. цена в акциях, ₽'] = `=IFERROR(IF(${bog}=0,"",${bog}),"")`;
  // как в старом «Командном пункте»: товар в акции продаётся по самой низкой цене участия,
  // поэтому прибыль и маржа считаются от неё; вне акций — от цены продажи
  const act = c('Мин. цена в акциях, ₽');
  F['Цена факт., ₽'] = `=IF(NOT(ISNUMBER(${base})),"",IF(AND(ISNUMBER(${act}),N(${act})>0,N(${act})<${base}),${act},${base}))`;
  F['В акции'] = `=IF(N(${c('Мин. цена в акциях, ₽')})>0,"🟩","")`;
  F['Прибыль, ₽'] = `=IF(OR(${c('Затраты фикс., ₽')}="",N(${c('Цена факт., ₽')})=0,${c('Комиссия, %')}=""),"",ROUND(${c('Цена факт., ₽')}-${c('Цена факт., ₽')}*(${c('Комиссия, %')}+${c('Эквайринг, %')})-${c('Затраты фикс., ₽')},0))`;
  F['Маржа, %'] = `=IF(${c('Прибыль, ₽')}="","",${c('Прибыль, ₽')}/${c('Цена факт., ₽')})`;
  // допуск 0,2 п.п.: цена округляется до рубля, и маржа может выйти 11,996% вместо 12%
  F['Статус'] = `=IF(NOT(ISNUMBER(${c('Закуп, ₽')})),"❔ нет закупа",IF(${c('Маржа, %')}="","❔ нет цены",IF(${c('Прибыль, ₽')}<0,"⛔ убыток",IF(${c('Маржа, %')}<${cfgRef_('MIN_MARGIN')}-2/1000,"⚠ ниже мин.","✓ норма"))))`;
  F['Прибыль факт/шт, ₽'] = `=IFERROR(IF(${pf('Прибыль факт/шт, ₽')}="","",${pf('Прибыль факт/шт, ₽')}),"")`;
  F['Маржа факт, %'] = `=IFERROR(IF(${pf('Маржа факт, %')}="","",${pf('Маржа факт, %')}),"")`;
  F['Конкурент Ozon, ₽'] = `=IFERROR(IF(N(${t('Мин. цена конкурента на Ozon')})=0,"",${t('Мин. цена конкурента на Ozon')}),"")`;
  F['Другие площадки, ₽'] = `=IFERROR(IF(N(${t('Мин. цена на других площадках')})=0,"",${t('Мин. цена на других площадках')}),"")`;
  F['Разница с рынком'] = `=IF(OR(N(${c('Цена на Ozon, ₽')})=0,MIN(${c('Конкурент Ozon, ₽')},${c('Другие площадки, ₽')})=0),"",${c('Цена на Ozon, ₽')}/MIN(${c('Конкурент Ozon, ₽')},${c('Другие площадки, ₽')})-1)`;
  F['Позиция'] = `=IF(${c('Разница с рынком')}="","— нет данных",IF(${c('Разница с рынком')}<-${cfgRef_('COMPETITOR_TOLERANCE')},"✓ дешевле",IF(${c('Разница с рынком')}<=${cfgRef_('COMPETITOR_TOLERANCE')},"≈ на уровне","▲ дороже")))`;
  F['Цена к отправке, ₽'] = `=IF(NOT(ISNUMBER(${c('Цена продажи, ₽')})),"",ROUND(${c('Цена продажи, ₽')},0))`;
  F['Зачёркнутая, ₽'] = `=IF(${c('Цена к отправке, ₽')}="","",ROUND(${c('Цена к отправке, ₽')}/${cfgRef_('OLD_PRICE_ROUND')},0)*${cfgRef_('OLD_PRICE_ROUND')}+${cfgRef_('OLD_PRICE_ADD')})`;
  F['min_price, ₽'] = `=IF(ISNUMBER(${c('Мин. цена, ₽')}),${c('Мин. цена, ₽')},"")`;
  return F;
}

/* ---------- Применение ---------- */
function applySchema_() {
  const out = [];
  out.push('листы: ' + ensureSheets_());
  out.push('Настройки: ' + ensureSettings_());
  out.push('Ozon: ' + applyMainSchema_());
  return out.join(' | ');
}

/** Удаляет колонки по названию заголовка */
function dropColumns_(sheetName, names, headerRow) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || !names || !names.length) return 0;
  let removed = 0;
  names.forEach(name => {
    for (let guard = 0; guard < 5; guard++) {
      const h = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
      const i = h.indexOf(name);
      if (i < 0) break;
      sh.deleteColumn(i + 1);
      removed++;
    }
  });
  return removed;
}

function applyMainSchema_() {
  const sh = sheet_(SHEETS.MAIN);
  dropColumns_(SHEETS.MAIN, DROP_COLS['Ozon'], MAIN_HDR_ROW);
  // 0. объединённые ячейки мешают двигать колонки — разъединяем всё заранее
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  if (sh.getFilter()) sh.getFilter().remove();
  // 1. порядок колонок: недостающие вставляем, существующие двигаем на место
  MAIN_COLS.forEach((name, i) => {
    const target = i + 1;
    const h = headersAt_(sh, MAIN_HDR_ROW);
    const cur = h.indexOf(name) + 1;
    if (!cur) {
      sh.insertColumnBefore(Math.min(target, sh.getMaxColumns()));
      sh.getRange(MAIN_HDR_ROW, target).setValue(name);
    } else if (cur !== target) {
      sh.moveColumns(sh.getRange(1, cur, sh.getMaxRows(), 1), cur < target ? target + 1 : target);
    }
  });
  const h = headersAt_(sh, MAIN_HDR_ROW), W = MAIN_COLS.length;
  styleMain_(sh, h, W);

  // 2. формулы: в первую товарную строку, дальше regroup_ разносит по всем
  const m = readMain_();
  if (!m.rows.length) return 'структура обновлена, товаров пока нет';
  const first = m.rows[0]._row;
  const F = mainFormulas_(h, first);
  Object.keys(F).forEach(n => sh.getRange(first, h.indexOf(n) + 1).setFormula(fx_(F[n])));
  PropertiesService.getDocumentProperties().deleteProperty('TPL_MAIN');
  SpreadsheetApp.flush();
  const res = regroup_([]);
  paintColumns_(sh, h);
  styleBlocks_(sh, W);
  return res;
}

/** Оформление листа Ozon */
function styleMain_(sh, h, W) {
  const rows = Math.max(sh.getMaxRows() - MAIN_FIRST + 1, 1);
  const colRange = n => sh.getRange(MAIN_FIRST, h.indexOf(n) + 1, rows, 1);

  sh.getRange(MAIN_FIRST, 1, rows, W).setBackground(null).setFontColor('#202124')
    .setFontFamily(OZ_UI.font).setFontSize(OZ_UI.size).setFontWeight('normal').setVerticalAlignment('middle')
    .setBorder(null, null, null, null, true, false, OZ_UI.grid, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(1, 1, 1, sh.getMaxColumns()).clearContent().setBackground(null);
  MAIN_BANDS.forEach(([title, from, to]) => {
    const a = h.indexOf(from) + 1, b = h.indexOf(to) + 1;
    if (a < 1 || b < 1) return;
    sh.getRange(1, a, 1, b - a + 1).merge().setValue(title).setBackground(OZ_UI.band)
      .setFontColor(OZ_UI.bandText).setFontWeight('bold').setFontSize(9)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.setRowHeight(1, 22);

  sh.getRange(MAIN_HDR_ROW, 1, 1, W).setBackground(OZ_UI.header).setFontColor(OZ_UI.headerText)
    .setFontFamily(OZ_UI.font).setFontSize(9).setFontWeight('bold').setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(MAIN_HDR_ROW, 40);
  sh.setFrozenRows(2);
  // граница закрепления должна совпадать с краем объединённой группы в строке 1,
  // иначе Google ругается «часть объединённых ячеек»
  const lastBand = MAIN_BANDS[0][2];
  try { sh.setFrozenColumns(h.indexOf(lastBand) + 1); } catch (e) { sh.setFrozenColumns(0); }

  Object.keys(MAIN_WIDTHS).forEach(n => { if (h.indexOf(n) >= 0) sh.setColumnWidth(h.indexOf(n) + 1, MAIN_WIDTHS[n]); });
  MAIN_PCT.forEach(n => colRange(n).setNumberFormat('0.0%'));
  MAIN_MONEY.forEach(n => colRange(n).setNumberFormat('#,##0'));
  colRange('Остаток FBS').setNumberFormat('#,##0');
  ['Название', 'Категория', 'Поставщик', 'Источник закупа', 'Артикул', 'Результат', 'Статус', 'Позиция']
    .forEach(n => colRange(n).setHorizontalAlignment('left'));
  paintColumns_(sh, h);
  colRange('Результат').setFontColor(OZ_UI.muted).setFontSize(9);
  toggleColumns_(true);

  // подсветка: фон только у «Статуса», дальше — цвет текста
  const stL = '$' + letter_(h.indexOf('Статус') + 1) + MAIN_FIRST;
  const posL = '$' + letter_(h.indexOf('Позиция') + 1) + MAIN_FIRST;
  const stkL = '$' + letter_(h.indexOf('Остаток FBS') + 1) + MAIN_FIRST;
  const rule = (formula, ranges, bg, fc, bold, italic) => {
    let b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges(ranges);
    if (bg) b = b.setBackground(bg);
    if (fc) b = b.setFontColor(fc);
    if (bold) b = b.setBold(true);
    if (italic) b = b.setItalic(true);
    return b.build();
  };
  const mgL = '$' + letter_(h.indexOf('Маржа, %') + 1) + MAIN_FIRST;
  const st = [colRange('Статус')], margin = [colRange('Маржа, %')];
  const wholeRow = [sh.getRange(MAIN_FIRST, 1, rows, W)];
  const notBlock = `AND($A${MAIN_FIRST}<>"",LEFT($A${MAIN_FIRST},1)<>"▌")`;
  sh.setConditionalFormatRules([
    // маржа: 10–20% зелёная, ниже 10% красная
    rule(fx_(`=AND(ISNUMBER(${mgL}),${mgL}>=1/10,${mgL}<=2/10)`), margin, OZ_UI.good, OZ_UI.goodText, true),
    rule(fx_(`=AND(ISNUMBER(${mgL}),${mgL}<98/1000)`), margin, OZ_UI.bad, OZ_UI.badText, true),
    rule(fx_(`=AND(ISNUMBER(${mgL}),${mgL}>2/10)`), margin, OZ_UI.key, '#0B2545', true),
    // статус
    rule(fx_(`=LEFT(${stL},1)="⛔"`), st, OZ_UI.bad, OZ_UI.badText, true),
    rule(fx_(`=LEFT(${stL},1)="⚠"`), st, OZ_UI.warn, OZ_UI.warnText, true),
    rule(fx_(`=LEFT(${stL},1)="✓"`), st, OZ_UI.good, OZ_UI.goodText),
    rule(fx_(`=LEFT(${stL},1)="❔"`), st, null, OZ_UI.muted),
    rule(fx_(`=LEFT(${posL},1)="▲"`), [colRange('Позиция')], null, OZ_UI.badText),
    rule(fx_(`=LEFT(${posL},1)="✓"`), [colRange('Позиция')], null, OZ_UI.goodText),
    // товары без остатка — вся строка серая
    rule(fx_(`=AND(${notBlock},N(${stkL})=0)`), wholeRow, '#DDE1E6', '#6B7075', false, true)
  ]);
}

/** Заливка колонок: ручной ввод и ключевые колонки — на все строки */
function paintColumns_(sh, h) {
  const rows = Math.max(sh.getMaxRows() - MAIN_FIRST + 1, 1);
  const col = n => { const i = h.indexOf(n); return i < 0 ? null : sh.getRange(MAIN_FIRST, i + 1, rows, 1); };
  MAIN_INPUTS.forEach(n => { const r = col(n); if (r) r.setBackground(OZ_UI.input).setFontColor(OZ_UI.inputText); });
  ['Закуп, ₽', 'Остаток FBS', 'Маржа, %'].forEach(n => { const r = col(n); if (r) r.setBackground(OZ_UI.key).setFontWeight('bold'); });
  const cp = col('Цена продажи, ₽'); if (cp) cp.setFontWeight('bold');
  const inAct = col('В акции'); if (inAct) inAct.setHorizontalAlignment('center').setFontSize(12);
  ['Закуп, ₽', 'Остаток FBS', 'Цена продажи, ₽', 'Маржа, %'].forEach(n => {
    const i = h.indexOf(n); if (i >= 0) sh.getRange(MAIN_HDR_ROW, i + 1).setBackground('#1A3B5C');
  });
}

/** Строки-заголовки блоков — спокойный серый */
function styleBlocks_(sh, W) {
  const last = sh.getLastRow();
  if (last < MAIN_FIRST) return;
  const col = sh.getRange(MAIN_FIRST, 1, last - MAIN_FIRST + 1, 1).getValues();
  const rows = [];
  col.forEach((r, i) => { if (isBlockRow_(r[0])) rows.push(MAIN_FIRST + i); });
  chunk_(rows, 150).forEach(part => {
    sh.getRangeList(part.map(r => sh.getRange(r, 1, 1, W).getA1Notation()))
      .setBackground(OZ_UI.block).setFontColor(OZ_UI.blockText).setFontWeight('bold').setFontSize(11)
      .setBorder(true, null, true, null, null, null, '#5B7DA6', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
}

/** Прячет или показывает служебные колонки */
function toggleColumns_(hide) {
  const sh = sheet_(SHEETS.MAIN), h = headersAt_(sh, MAIN_HDR_ROW);
  let n = 0;
  MAIN_HIDE.forEach(name => {
    const i = h.indexOf(name);
    if (i < 0) return;
    if (hide) sh.hideColumns(i + 1); else sh.showColumns(i + 1);
    n++;
  });
  const bog = SpreadsheetApp.getActive().getSheetByName(SHEETS.BOG);
  if (bog) for (let c = 22; c <= 24; c++) { if (hide) bog.hideColumns(c); else bog.showColumns(c); }
  return `${hide ? 'скрыто' : 'показано'} колонок: ${n}`;
}

/** Добавляет недостающие ключи настроек и убирает устаревшие, значения остальных не трогает */
function ensureSettings_() {
  const sh = sheet_(SHEETS.SETTINGS);
  const keys = () => sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  let have = keys();
  const dropped = [];
  for (let i = have.length - 1; i >= 0; i--) {             // снизу вверх, чтобы номера строк не съезжали
    if (SETTINGS_OBSOLETE.indexOf(have[i]) >= 0) { sh.deleteRow(i + 2); dropped.push(have[i]); }
  }
  if (dropped.length) have = keys();
  const add = SETTINGS_DEFAULTS.filter(d => have.indexOf(d[0]) < 0);
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  SETTINGS_CACHE_ = null;
  const out = [];
  if (add.length) out.push('добавлены ключи: ' + add.map(a => a[0]).join(', '));
  if (dropped.length) out.push('убраны устаревшие: ' + dropped.join(', '));
  return out.length ? out.join('; ') : 'всё на месте';
}

/** Прайсы, которые должны быть на листе «Источники закупа» (из формул старой таблицы) */
var SOURCES_DEFAULTS = [
  // Источник, ID таблицы, Лист, Код, Закуп, РРЦ, Валюта
  ['Прайс АТОЛ',      '1wOJa0O7Rdxgh_Kyy_Gcvsfq-1xq_X6BVg2ytC4pRx1M', 'АТОЛ', 'B', 'D', 'C', 'RUB'],
  ['Прайс MERTECH',   '1wZqNiJmKV17l9SLTl-5WS6Ppu1CHhzMtN9EF69X5wgw', 'Mertech Products', 'C', 'G', 'F', 'RUB'],
  ['Прайс POSCENTER', '1MhVhzJ3JqXm8acnn5APaz1a1Zfwd1TZ3JKQBb1HoAIo', 'Лист1', 'A', 'D', '', 'RUB'],
  ['Прайс CAS',       '1Z0Sfk_h2b4PTe0JOM2Zzfk2aigesl3l5xK5ZInSfizQ', 'Прайс', 'B', 'E', 'L', 'USD'],
  ['Прайс Cassida',   '1QJU_fjOJqJwc9y7NgTx43bZtPzlJA47e58gpnmi3iXI', 'Лист 1', 'B', 'D', 'F', 'RUB'],
  ['Прайс PayTor',    '14i7YHri2W9Q3s9qlskYYhp9anmCwc5xJCvqug0cSnTk', 'Прайс_full, Сенсорные терминалы, POS-периферия, Сканеры и принтеры шк', 'A', 'H', 'F', 'RUB'],
  ['Прайс Компас',    '12F715p61VJSIUK6A1rKyb0UYc43LQqTuxI-_MMfxfog', 'TDSheet', 'A', 'N', 'P', 'USD'],
  ['Прайс Dors',      '1tyLYsUYbdK7vGkJ0y95UnMtFIiwt_RDflJmCcogDk2c', 'Table 1', 'A', 'C', 'H', 'RUB'],   // смешанная валюта: $ определяется по ячейке
  ['Прайс Скансити',  '18J6EL2QqJaaATKa0-fsHxcA9XC2gyPEPPkS7H-7cgMk', 'Лист1', 'B', 'D', '', 'USD'],
  ['Прайс Эвотор',    '1dbmU69YY-5FFqoLFGioaCZ6p3kAbpOP1AHBoHYQIPd0', 'Лист1', 'C', 'E', '', 'RUB']
];

/**
 * Приводит лист «Источники закупа» в порядок: по ID таблицы находит строку прайса
 * и выставляет лист, колонки и валюту; недостающие прайсы добавляет.
 * Строки, которые вы добавили сами (с другим ID), не трогает.
 */
function ensureSources_() {
  const sh = sheet_(SHEETS.SOURCES);
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, SOURCE_COLS.length).getValues() : [];
  const byId = {};
  data.forEach((r, i) => { const id = String(r[1]).trim(); if (id) byId[id] = i; });
  let fixed = 0, added = 0;
  SOURCES_DEFAULTS.forEach(d => {
    const row = [d[0], d[1], d[2], d[3], d[4], d[5], d[6], '', ''];
    if (byId[d[1]] !== undefined) {
      const i = byId[d[1]];
      row[0] = String(data[i][0]).trim() || d[0];          // название оставляем ваше
      data[i] = row;
      fixed++;
    } else { data.push(row); added++; }
  });
  // остальные ваши строки: колонка «Валюта» могла съехать — чиним на RUB, если там не валюта
  data.forEach(r => { if (['RUB', 'USD'].indexOf(String(r[6]).trim().toUpperCase()) < 0) r[6] = 'RUB'; r[7] = ''; r[8] = r[8] || ''; });
  if (last > 1) sh.getRange(2, 1, last - 1, SOURCE_COLS.length).clearContent();
  if (data.length) sh.getRange(2, 1, data.length, SOURCE_COLS.length).setValues(data);
  sh.getRange(1, 1, 1, SOURCE_COLS.length).setValues([SOURCE_COLS]);
  return `прайсов: ${data.length} (обновлено ${fixed}, добавлено ${added})`;
}

/** Создаёт недостающие листы, ставит заголовки и строки-шаблоны формул */
function ensureSheets_() {
  const ss = SpreadsheetApp.getActive(), made = [];
  const ensure = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); made.push(name); }
    const cur = sh.getLastColumn() ? headersAt_(sh, 1) : [];
    if (cur.join('|') !== headers.join('|')) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground(OZ_UI.header)
        .setFontColor(OZ_UI.headerText).setFontSize(9).setWrap(true);
      sh.setFrozenRows(1);
    }
    return sh;
  };
  dropColumns_(SHEETS.TARIFFS, DROP_COLS['Тарифы'], 1);
  ensure(SHEETS.TARIFFS, TARIFF_COLS);
  ensure('План-факт', PF_COLS);
  ensure(SHEETS.DISCOUNTS, DISCOUNT_COLS);
  ensure(SHEETS.SOURCES, SOURCE_COLS);
  const srcNote = ensureSources_();
  ensure(SHEETS.LOG, ['Время', 'Операция', 'Уровень', 'Сообщение']);
  dropColumns_('План-факт', DROP_COLS['План-факт'], 1);
  ensure('План-факт', PF_COLS);
  ensureFormulaRow_('План-факт', PF_COLS, pfFormulas_(2));
  ensureFormulaRow_(SHEETS.DISCOUNTS, DISCOUNT_COLS, discountFormulas_(2));
  formatPlanFact_();
  return (made.length ? 'созданы ' + made.join(', ') : 'все на месте') + ' | источники: ' + srcNote;
}

/** Форматы колонок «План-факта»: деньги — числами, доли — процентами */
function formatPlanFact_() {
  const sh = sheet_('План-факт');
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  const col = n => { const i = PF_COLS.indexOf(n); return i < 0 ? null : sh.getRange(2, i + 1, rows, 1); };
  ['Выручка, ₽', 'Комиссия, ₽', 'Логистика и услуги, ₽', 'Возвраты, ₽', 'Реклама, ₽', 'Штрафы, ₽', 'Прочее, ₽',
   'К выплате, ₽', 'Закуп проданного, ₽', 'Прибыль факт, ₽', 'Прибыль факт/шт, ₽', 'Прибыль план/шт, ₽',
   'Факт − план/шт, ₽'].forEach(n => { const r = col(n); if (r) r.setNumberFormat('#,##0'); });
  ['Продано, шт', 'Возвращено, шт'].forEach(n => { const r = col(n); if (r) r.setNumberFormat('#,##0'); });
  ['Удержания факт, %', 'Удержания план, %', 'Разница, п.п.', 'Маржа факт, %']
    .forEach(n => { const r = col(n); if (r) r.setNumberFormat('0.0%'); });
  const upd = col('Обновлено'); if (upd) upd.setNumberFormat('dd.MM.yyyy HH:mm');
}

/** Строка 2 служит шаблоном формул для writeTable_ */
function ensureFormulaRow_(name, cols, F) {
  const sh = sheet_(name);
  // чистим строку целиком: после добавления колонок старые формулы съезжают и дают #ERROR!
  sh.getRange(2, 1, 1, Math.max(sh.getMaxColumns(), cols.length)).clearContent().clearFormat();
  sh.getRange(2, 1, 1, cols.length).setFontFamily(OZ_UI.font).setFontSize(OZ_UI.size);
  Object.keys(F).forEach(n => sh.getRange(2, cols.indexOf(n) + 1).setFormula(fx_(F[n])));
  PropertiesService.getDocumentProperties().deleteProperty('TPL_' + name);
  SpreadsheetApp.flush();
  templateRow_(sh, cols, 2, 'TPL_' + name);
}

function pfFormulas_(r) {
  const C = n => '$' + letter_(PF_COLS.indexOf(n) + 1) + r;
  const mc = n => '$' + letter_(MAIN_COLS.indexOf(n) + 1);
  const oz = n => `INDEX(Ozon!${mc(n)}:${mc(n)},MATCH($A${r},Ozon!${mc('SKU')}:${mc('SKU')},0))`;
  const net = `(${C('Продано, шт')}-${C('Возвращено, шт')})`;
  return {
    'Артикул': `=IFERROR(${oz('Артикул')}&"","")`,
    'Название': `=IFERROR(${oz('Название')}&"","")`,
    'Категория': `=IFERROR(${oz('Категория')}&"","")`,
    'Удержания факт, %': `=IF(N(${C('Выручка, ₽')})<=0,"",-(${C('Комиссия, ₽')}+${C('Логистика и услуги, ₽')}+${C('Возвраты, ₽')}+${C('Реклама, ₽')}+${C('Штрафы, ₽')}+${C('Прочее, ₽')})/${C('Выручка, ₽')})`,
    'Удержания план, %': `=IFERROR((${oz('Цена факт., ₽')}*(${oz('Комиссия, %')}+${oz('Эквайринг, %')})+${oz('Логистика с выкупом, ₽')})/${oz('Цена факт., ₽')},"")`,
    'Разница, п.п.': `=IF(OR(${C('Удержания факт, %')}="",${C('Удержания план, %')}=""),"",${C('Удержания факт, %')}-${C('Удержания план, %')})`,
    'Закуп проданного, ₽': `=IFERROR(${oz('Закуп, ₽')}*${net},"")`,
    'Прибыль факт, ₽': `=IF(OR(${C('К выплате, ₽')}="",${C('Закуп проданного, ₽')}=""),"",ROUND(${C('К выплате, ₽')}-${C('Закуп проданного, ₽')}-${cfgRef_('PACKAGING_RUB')}*${net},0))`,
    'Прибыль факт/шт, ₽': `=IF(OR(${C('Прибыль факт, ₽')}="",${net}<=0),"",ROUND(${C('Прибыль факт, ₽')}/${net},0))`,
    'Маржа факт, %': `=IF(OR(${C('Прибыль факт, ₽')}="",N(${C('Выручка, ₽')})<=0),"",${C('Прибыль факт, ₽')}/${C('Выручка, ₽')})`,
    'Прибыль план/шт, ₽': `=IFERROR(IF(${oz('Прибыль, ₽')}="","",${oz('Прибыль, ₽')}),"")`,
    'Факт − план/шт, ₽': `=IF(OR(${C('Прибыль факт/шт, ₽')}="",${C('Прибыль план/шт, ₽')}=""),"",${C('Прибыль факт/шт, ₽')}-${C('Прибыль план/шт, ₽')})`
  };
}

function discountFormulas_(r) {
  const C = n => '$' + letter_(DISCOUNT_COLS.indexOf(n) + 1) + r;
  const mc = n => '$' + letter_(MAIN_COLS.indexOf(n) + 1);
  const oz = n => `INDEX(Ozon!${mc(n)}:${mc(n)},MATCH(${C('Артикул')},Ozon!$A:$A,0))`;
  return {
    'Порог акций, ₽': `=IFERROR(${oz('Порог акций, ₽')},"")`,
    'Рекомендация': `=IF(AND(${C('Статус')}<>"NEW",${C('Статус')}<>"SEEN"),"— обработана",IF(NOT(ISNUMBER(${C('Порог акций, ₽')})),"❔ нет порога",IF(${C('Запрошенная цена')}>=${C('Порог акций, ₽')},"✓ можно одобрить","✗ ниже порога")))`
  };
}


/* ================== 09_audit.gs ================== */
/** =====================================================================
 *  АУДИТ: сверка расчёта таблицы с живыми данными Ozon.
 *  Лист «Аудит» — расширенный: все расходы из таблицы, те же данные с Ozon
 *  и расхождения между ними.
 *  Что делает автоматика (каждое действие отключается в «Настройках»):
 *    • маржа по живым данным Ozon ниже порога акций → товар убирается из акций;
 *    • маржа ниже нуля (продажа в убыток) → остаток ставится в 0, чтобы не купили;
 *    • всё записывается в лист и в «Лог», остаток до обнуления сохраняется
 *      для отката кнопкой «Вернуть остатки».
 * ===================================================================== */
function runAudit()       { run_('Аудит и защита', runAudit_); }
function restoreStocks()  { run_('Вернуть остатки', restoreStocks_); }

var AUDIT = 'Аудит';
var AUDIT_COLS = ['Артикул', 'SKU', 'Product ID', 'Название', 'Категория', 'Остаток FBS',
  // как считает таблица
  'Закуп, ₽', 'Комиссия табл., %', 'Логистика табл., ₽', 'Обработка табл., ₽', 'Посл. миля табл., ₽',
  'Возврат табл., ₽', 'Выкуп, %', 'Затраты табл., ₽', 'Цена продажи, ₽', 'Мин. цена, ₽', 'Порог акций, ₽',
  'Цена в акциях табл., ₽', 'Маржа табл., %',
  // что отдаёт Ozon прямо сейчас
  'Комиссия Ozon, %', 'Логистика Ozon, ₽', 'Обработка Ozon, ₽', 'Посл. миля Ozon, ₽', 'Возврат Ozon, ₽',
  'Цена Ozon, ₽', 'Мин. цена в акциях Ozon, ₽', 'Затраты Ozon, ₽', 'Прибыль по Ozon, ₽', 'Маржа по Ozon, %',
  // расхождения и вердикт
  'Δ комиссия, п.п.', 'Δ логистика, ₽', 'Δ цена, ₽', 'Δ маржа, п.п.',
  'Статус проверки', 'Что сделано', 'Остаток до обнуления', 'Проверено'];

var AUDIT_SETTINGS = [
  ['AUDIT_TOLERANCE', 0.02, 'Допустимое расхождение маржи и комиссии (2 п.п.). Больше — помечаем как ошибку.'],
  ['AUDIT_REMOVE_FROM_ACTIONS', true, 'Убирать из акций товары, у которых маржа по живым данным Ozon ниже PROMO_MARGIN.'],
  ['AUDIT_ZERO_STOCK', false, 'ОПАСНО: обнулять остаток у товаров, которые по данным Ozon продаются в убыток. Включайте осознанно.'],
  ['AUDIT_MAX_ZERO', 10, 'Максимум товаров, которым за один запуск можно обнулить остаток (защита от массового обнуления).'],
  ['AUDIT_MAX_REMOVE', 50, 'Максимум снятий с акций за один запуск аудита.']
];

/* ---------- Основной прогон ---------- */
function runAudit_() {
  ensureAuditSheet_();
  const dry = isDryRun_();
  const tol = Number(cfg_('AUDIT_TOLERANCE', 0.02));
  const promo = Number(cfg_('PROMO_MARGIN', 0.12)), minMargin = Number(cfg_('MIN_MARGIN', 0.10));
  const acq = Number(cfg_('ACQUIRING_RATE', 0.01));
  const pack = Number(cfg_('PACKAGING_RUB', 20));
  const defBuyout = Number(cfg_('DEFAULT_BUYOUT', 0.92));
  const logMode = String(cfg_('LOGISTICS_MODE', 'MAX')).toUpperCase();
  const manual = manualPromo_();

  // 1. живые данные Ozon: цены, комиссии, логистика
  const live = {};
  ozonAll_('/v5/product/info/prices', { filter: { visibility: 'ALL' }, limit: 1000 },
    r => ({ items: r.items, next: r.cursor }), 'cursor').forEach(o => {
      const p = o.price || {}, c = o.commissions || {};
      live[key_(o.product_id)] = {
        price: num_(p.price) || 0, comm: (num_(c.sales_percent_fbs) || 0) / 100,
        logMin: num_(c.fbs_direct_flow_trans_min_amount) || 0, logMax: num_(c.fbs_direct_flow_trans_max_amount) || 0,
        proc: num_(c.fbs_first_mile_max_amount) || 0, last: num_(c.fbs_deliv_to_customer_amount) || 0,
        ret: (num_(c.fbs_return_flow_amount) || 0) + (num_(c.fbs_return_flow_trans_max_amount) || 0)
      };
    });

  // 2. живые цены участия в акциях
  const actions = listActions_(), actMin = {}, actWhere = {};
  actions.forEach(a => fetchActionProducts_(a.id).forEach(p => {
    const k = key_(p.id), price = Number(p.action_price) || 0;
    if (!price) return;
    (actWhere[k] = actWhere[k] || []).push({ id: a.id, title: a.title, price });
    if (!actMin[k] || price < actMin[k]) actMin[k] = price;
  }));

  // 3. сверка по каждому товару
  const m = readMain_(), rows = [], toRemove = {}, toZero = [];
  const prevStock = auditPrevStock_();
  m.rows.forEach(r => {
    const pid = key_(r['Product ID']), L = live[pid];
    const cost = Number(r['Закуп, ₽']) || 0;
    const buy = Number(r['Выкуп, %']) || defBuyout;
    const tablMargin = Number(r['Маржа, %']);
    const row = { 'Артикул': r['Артикул'], 'SKU': r['SKU'], 'Product ID': r['Product ID'], 'Название': r['Название'],
      'Категория': r['Категория'], 'Остаток FBS': r['Остаток FBS'], 'Закуп, ₽': cost || '',
      'Комиссия табл., %': r['Комиссия, %'], 'Логистика табл., ₽': r['Логистика, ₽'], 'Обработка табл., ₽': r['Обработка Ozon, ₽'],
      'Посл. миля табл., ₽': r['Посл. миля, ₽'], 'Возврат табл., ₽': r['Возврат, ₽'], 'Выкуп, %': buy,
      'Затраты табл., ₽': r['Затраты фикс., ₽'], 'Цена продажи, ₽': r['Цена продажи, ₽'], 'Мин. цена, ₽': r['Мин. цена, ₽'],
      'Порог акций, ₽': r['Порог акций, ₽'], 'Цена в акциях табл., ₽': r['Мин. цена в акциях, ₽'],
      'Маржа табл., %': tablMargin, 'Проверено': new Date(),
      'Остаток до обнуления': prevStock[key_(r['Артикул'])] || '' };

    if (!L) { row['Статус проверки'] = '❔ нет в ответе Ozon'; rows.push(row); return; }
    if (!cost) { row['Статус проверки'] = '❔ нет закупа'; rows.push(row); return; }

    // расходы по живым данным Ozon
    const logi = logMode === 'MIN' ? L.logMin : logMode === 'AVG' ? (L.logMin + L.logMax) / 2 : L.logMax;
    const delivery = (logi + L.proc + L.last) / buy + (1 / buy - 1) * L.ret;
    const fix = cost + delivery + pack;
    const priceNow = Math.min(L.price || Infinity, actMin[pid] || Infinity);
    const revenue = priceNow === Infinity ? 0 : priceNow;
    const profit = revenue ? Math.round(revenue - priceNow * (L.comm + acq) - fix) : '';
    const margin = revenue ? profit / revenue : '';

    Object.assign(row, {
      'Комиссия Ozon, %': L.comm, 'Логистика Ozon, ₽': logi, 'Обработка Ozon, ₽': L.proc, 'Посл. миля Ozon, ₽': L.last,
      'Возврат Ozon, ₽': L.ret, 'Цена Ozon, ₽': L.price || '', 'Мин. цена в акциях Ozon, ₽': actMin[pid] || '',
      'Затраты Ozon, ₽': Math.round(fix), 'Прибыль по Ozon, ₽': profit, 'Маржа по Ozon, %': margin,
      'Δ комиссия, п.п.': isFinite(Number(r['Комиссия, %'])) ? L.comm - Number(r['Комиссия, %']) : '',
      'Δ логистика, ₽': isFinite(Number(r['Логистика, ₽'])) ? Math.round(logi - Number(r['Логистика, ₽'])) : '',
      'Δ цена, ₽': isFinite(Number(r['Цена на Ozon, ₽'])) ? Math.round((L.price || 0) - Number(r['Цена на Ozon, ₽'])) : '',
      'Δ маржа, п.п.': (margin !== '' && isFinite(tablMargin)) ? margin - tablMargin : ''
    });

    // вердикт
    const dComm = Math.abs(Number(row['Δ комиссия, п.п.']) || 0);
    const dMargin = Math.abs(Number(row['Δ маржа, п.п.']) || 0);
    const manualHere = (actWhere[pid] || []).some(a => manual.has(manualKey_(a.id, pid)));
    if (margin === '') row['Статус проверки'] = '❔ нет цены на Ozon';
    else if (margin < 0) row['Статус проверки'] = '⛔ убыток по данным Ozon';
    else if (margin < promo && actMin[pid]) row['Статус проверки'] = manualHere && margin >= minMargin
      ? '✓ в акции от мин. маржи (ручное решение)' : '⚠ в акции ниже порога';
    else if (dComm > tol || dMargin > tol) row['Статус проверки'] = '⚠ расхождение с таблицей';
    else row['Статус проверки'] = '✓ сходится';

    // что делаем
    const acts = [];
    // снимаем с акции, только если цена этой акции реально ниже порога из таблицы
    // и маржа по живым данным ниже нужной с запасом на допуск. Ручное решение «маржа от MIN_MARGIN»
    // сверяем с «Мин. ценой» и MIN_MARGIN, остальные — с «Порогом акций» и PROMO_MARGIN
    const floor = Number(r['Порог акций, ₽']) || 0, minFloor = Number(r['Мин. цена, ₽']) || 0;
    const badIn = (actWhere[pid] || []).filter(a => {
      const own = manual.has(manualKey_(a.id, pid)) && minFloor > 0;
      const fl = own ? minFloor : floor, mg = own ? minMargin : promo;
      return margin !== '' && fl > 0 && a.price < fl && margin < mg - tol;
    });
    if (cfg_('AUDIT_REMOVE_FROM_ACTIONS', true) === true && badIn.length) {
      badIn.forEach(a => { (toRemove[a.id] = toRemove[a.id] || []).push({ pid: Number(r['Product ID']), art: r['Артикул'], title: a.title, price: a.price }); });
      acts.push(`убираем из акций: ${badIn.map(a => a.title).join(', ')}`);
    }
    if (cfg_('AUDIT_ZERO_STOCK', false) === true && margin !== '' && margin < 0 && Number(r['Остаток FBS']) > 0) {
      toZero.push({ art: String(r['Артикул']), sku: Number(r['SKU']), stock: Number(r['Остаток FBS']), margin });
      acts.push('обнуляем остаток');
    }
    row['Что сделано'] = acts.length ? (dry ? 'ПРОВЕРКА: ' : '') + acts.join('; ') : '';
    rows.push(row);
  });

  // 4. снятие с акций (с ограничителем на один запуск)
  let removed = 0, removeSkipped = 0;
  const maxRemove = Number(cfg_('AUDIT_MAX_REMOVE', 50));
  let budget = maxRemove;
  Object.keys(toRemove).forEach(aid => {
    let list = toRemove[aid];
    if (list.length > budget) { removeSkipped += list.length - budget; list = list.slice(0, budget); }
    budget -= list.length;
    if (!list.length) return;
    if (dry) { removed += list.length; return; }
    chunk_(list, 1000).forEach(part => {
      const r = ozon_('/v1/actions/products/deactivate', { action_id: Number(aid), product_ids: part.map(x => x.pid) });
      removed += ((r.result && r.result.product_ids) || []).length;
    });
  });

  // 5. обнуление остатков — с ограничителем
  const maxZero = Number(cfg_('AUDIT_MAX_ZERO', 10));
  let zeroed = 0, zeroSkipped = 0;
  if (toZero.length) {
    toZero.sort((a, b) => a.margin - b.margin);                 // сначала самые убыточные
    if (toZero.length > maxZero) { zeroSkipped = toZero.length - maxZero; toZero.length = maxZero; }
    if (!dry) zeroed = setStocks_(toZero.map(x => ({ offer_id: x.art, stock: 0 })));
    else zeroed = toZero.length;
    saveAuditPrevStock_(toZero);
  }

  writeTable_(AUDIT, rows);
  const bad = rows.filter(r => String(r['Статус проверки']).charAt(0) === '⛔').length;
  const warn = rows.filter(r => String(r['Статус проверки']).charAt(0) === '⚠').length;
  log_('Аудит', dry ? 'DRY' : 'INFO',
    `Проверено ${rows.length}: убыточных по Ozon ${bad}, расхождений ${warn}. Снято с акций ${removed}` +
    (removeSkipped ? ` (ещё ${removeSkipped} не тронуты: лимит AUDIT_MAX_REMOVE)` : '') + `, обнулено остатков ${zeroed}` +
    (zeroSkipped ? `, пропущено из-за лимита AUDIT_MAX_ZERO ${zeroSkipped}` : '') +
    (toZero.length ? '\n' + toZero.map(x => `${x.art}: маржа ${Math.round(x.margin * 1000) / 10}%, было ${x.stock} шт`).join('\n') : ''));
  return `${dry ? '[ПРОВЕРКА] ' : ''}убыточных: ${bad}, расхождений: ${warn}, снято с акций: ${removed}, обнулено: ${zeroed}` +
    (zeroSkipped ? ` (ещё ${zeroSkipped} ждут: лимит за запуск)` : '');
}

/* ---------- Остатки FBS ---------- */
/** Склад FBS по названию из настроек */
function warehouseId_() {
  const name = String(cfg_('FBS_WAREHOUSE', '')).trim().toLowerCase();
  let list = [];
  try { const r = ozon_('/v2/warehouse/list', {}); list = r.result || r.warehouses || []; }
  catch (e) { const r = ozon_('/v1/warehouse/list', {}); list = r.result || []; }   // старый метод отключён 20.03.2026
  const w = list.find(x => String(x.name || '').trim().toLowerCase() === name) || list[0];
  if (!w) throw new Error('Не найден FBS-склад');
  return w.warehouse_id;
}
/** Ставит остатки: items = [{offer_id, stock}] */
function setStocks_(items) {
  const wid = warehouseId_();
  let ok = 0;
  chunk_(items, 100).forEach(part => {
    const r = ozon_('/v2/products/stocks', { stocks: part.map(x => ({ offer_id: String(x.offer_id), stock: Number(x.stock), warehouse_id: wid })) });
    ((r.result) || []).forEach(x => { if (x.updated) ok++; else log_('Остатки', 'WARN', `${x.offer_id}: ${JSON.stringify(x.errors || [])}`); });
  });
  return ok;
}
/** Возвращает остатки, обнулённые аудитом */
function restoreStocks_() {
  const t = readTable_(AUDIT);
  const back = t.rows.filter(r => Number(r['Остаток до обнуления']) > 0)
    .map(r => ({ offer_id: String(r['Артикул']), stock: Number(r['Остаток до обнуления']) }));
  if (!back.length) return 'нечего возвращать';
  if (isDryRun_()) return `[ПРОВЕРКА] вернули бы остатки: ${back.length}`;
  const ok = setStocks_(back);
  const patch = {}; t.rows.forEach(r => { if (Number(r['Остаток до обнуления']) > 0) patch[r._row] = ''; });
  patchColumn_(t.sh, t.h, 'Остаток до обнуления', patch);
  log_('Остатки', 'INFO', 'Возвращены остатки: ' + back.map(x => `${x.offer_id}=${x.stock}`).join(', '));
  return `возвращено: ${ok}`;
}
/** Запомненные остатки до обнуления (по артикулу) */
function auditPrevStock_() {
  const ss = SpreadsheetApp.getActive().getSheetByName(AUDIT);
  const out = {};
  if (!ss || ss.getLastRow() < 2) return out;
  const t = readTable_(AUDIT);
  t.rows.forEach(r => { if (Number(r['Остаток до обнуления']) > 0) out[key_(r['Артикул'])] = Number(r['Остаток до обнуления']); });
  return out;
}
function saveAuditPrevStock_(list) {
  const p = PropertiesService.getDocumentProperties();
  p.setProperty('AUDIT_ZEROED', JSON.stringify(list.map(x => ({ a: x.art, s: x.stock }))));
}

/* ---------- Лист «Аудит» ---------- */
function ensureAuditSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(AUDIT);
  if (!sh) sh = ss.insertSheet(AUDIT);
  const cur = sh.getLastColumn() ? headersAt_(sh, 1) : [];
  if (cur.join('|') !== AUDIT_COLS.join('|')) {
    sh.getRange(1, 1, 1, AUDIT_COLS.length).setValues([AUDIT_COLS])
      .setFontWeight('bold').setBackground(OZ_UI.header).setFontColor(OZ_UI.headerText).setFontSize(9).setWrap(true);
    sh.setFrozenRows(1); sh.setFrozenColumns(1);
    sh.setColumnWidth(1, 150); sh.setColumnWidth(4, 280);
    sh.setColumnWidth(AUDIT_COLS.indexOf('Статус проверки') + 1, 170);
    sh.setColumnWidth(AUDIT_COLS.indexOf('Что сделано') + 1, 260);
  }
  const rows = Math.max(sh.getMaxRows() - 1, 1);
  const col = n => sh.getRange(2, AUDIT_COLS.indexOf(n) + 1, rows, 1);
  ['Комиссия табл., %', 'Выкуп, %', 'Маржа табл., %', 'Комиссия Ozon, %', 'Маржа по Ozon, %',
   'Δ комиссия, п.п.', 'Δ маржа, п.п.'].forEach(n => col(n).setNumberFormat('0.0%'));
  const stL = '$' + letter_(AUDIT_COLS.indexOf('Статус проверки') + 1) + '2';
  const rule = (f, bg, fc) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f)
    .setBackground(bg).setFontColor(fc).setRanges([sh.getRange(2, 1, rows, AUDIT_COLS.length)]).build();
  sh.setConditionalFormatRules([
    rule(`=LEFT(${stL},1)="⛔"`, OZ_UI.bad, OZ_UI.badText),
    rule(`=LEFT(${stL},1)="⚠"`, OZ_UI.warn, OZ_UI.warnText)
  ]);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), AUDIT_COLS.length).createFilter();
  return sh;
}
