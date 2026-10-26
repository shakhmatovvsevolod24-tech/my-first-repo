/** =====================================================================
 *  OZON: КАРТОЧКИ И ПРОДВИЖЕНИЕ — отдельная таблица, весь код в одном файле.
 *  Где теряются покупатели и что поправить в карточке. Только чтение:
 *  в кабинете Ozon скрипт ничего не меняет. Нужна подписка Ozon Premium.
 *    «Товары»             — список товаров из кабинета: тип Ozon, остаток FBS, заказы;
 *    «Воронка»            — показы → посетители карточки → корзина → заказ за FUNNEL_DAYS дней,
 *                           сравнение с предыдущим таким же периодом и диагноз по каждому товару;
 *    «Поисковые запросы»  — по каким запросам находят товар и каких слов запроса нет в названии;
 *    «Качество карточек»  — контент-рейтинг Ozon и что именно его снижает.
 *  Ключи API хранятся в свойствах скрипта, в коде их нет.
 * ===================================================================== */

/* ================== 01_core ================== */
var OZON_HOST = 'https://api-seller.ozon.ru';
var SHEETS = { PRODUCTS: 'Товары', SETTINGS: 'Настройки', LOG: 'Лог' };

var OZ_UI = {
  header: '#455A64', headerText: '#FFFFFF',
  bad: '#FCE8E6', badText: '#B3261E',
  warn: '#FEF7E0', warnText: '#8A6100',
  good: '#E6F4EA', goodText: '#1E7B34'
};

var SETTINGS_DEFAULTS = [
  ['FUNNEL_DAYS', 28, 'Период «Воронки» и «Поисковых запросов», дней. Сравнивается с предыдущим таким же периодом.'],
  ['FUNNEL_MIN_VIEWS', 300, 'Меньше показов за период — диагноз «Не видят».'],
  ['FUNNEL_MIN_VISITS', 30, 'Меньше посетителей карточки — конверсии не оцениваем.'],
  ['FUNNEL_WEAK', 0.5, 'Конверсия ниже этой доли от медианы категории считается проблемой.'],
  ['QUERIES_PER_SKU', 15, 'Сколько самых частых запросов брать на товар (Ozon отдаёт не больше 15).'],
  ['QUERIES_TIP_MIN', 20, 'Подсказку «добавьте слова» даём, если запрос искали хотя бы столько покупателей.'],
  ['ORDERS_DAYS', 60, 'За сколько дней считать заказы на листе «Товары».']
];

/* ---------- Ключи ---------- */
function getCreds_() {
  const p = PropertiesService.getScriptProperties();
  const clientId = p.getProperty('OZON_CLIENT_ID'), apiKey = p.getProperty('OZON_API_KEY');
  if (!clientId || !apiKey) throw new Error('Не заданы ключи API: 🔎 КАРТОЧКИ → Подключение → Ключи API');
  return { clientId, apiKey };
}
function setupCredentials() {
  const ui = SpreadsheetApp.getUi();
  const c = ui.prompt('Ozon: Client-Id', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  const k = ui.prompt('Ozon: Api-Key', ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperties({
    OZON_CLIENT_ID: c.getResponseText().trim(), OZON_API_KEY: k.getResponseText().trim()
  });
  ui.alert('Ключи сохранены в свойствах скрипта.');
}

/* ---------- Настройки: лист создаётся сам, недостающие ключи дописываются ---------- */
let SETTINGS_CACHE_ = null;
function ensureSettings_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.SETTINGS);
    sh.getRange(1, 1, 1, 3).setValues([['Ключ', 'Значение', 'Описание']]).setFontWeight('bold')
      .setBackground(OZ_UI.header).setFontColor(OZ_UI.headerText);
    sh.setFrozenRows(1); sh.setColumnWidth(1, 170); sh.setColumnWidth(3, 520);
  }
  const have = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  const add = SETTINGS_DEFAULTS.filter(d => have.indexOf(d[0]) < 0);
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
  SETTINGS_CACHE_ = null;
  return sh;
}
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
    catch (e) {                                   // сбой сети — повторяем, как при 5xx
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

/* ---------- Мелочи ---------- */
function chunk_(arr, n) { const o = []; for (let i = 0; i < arr.length; i += n) o.push(arr.slice(i, i + n)); return o; }
function key_(v) { return String(v === null || v === undefined ? '' : v).trim().replace(/\.0+$/, '').toLowerCase(); }
function toast_(msg, sec) { try { SpreadsheetApp.getActive().toast(msg, 'Карточки Ozon', sec || 5); } catch (e) {} }

/* ---------- Лог и запуск ---------- */
function log_(op, level, msg) {
  console.log(`[${level}] ${op}: ${msg}`);
  try {
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(SHEETS.LOG);
    if (!sh) { sh = ss.insertSheet(SHEETS.LOG); sh.getRange(1, 1, 1, 4).setValues([['Время', 'Операция', 'Уровень', 'Сообщение']]).setFontWeight('bold'); }
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
    ensureSettings_();
    const msg = fn() || 'готово';
    log_(title, 'OK', msg);
    toast_(`${title}: ${msg} (${Math.round((Date.now() - t0) / 1000)} с)`, 10);
  } catch (e) {
    log_(title, 'ERROR', e.stack || e);
    toast_(`${title}: ОШИБКА — ${e.message}`, 15);
    throw e;
  } finally { lock.releaseLock(); }
}

/* ================== 02_menu ================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔎 КАРТОЧКИ')
    .addItem('▶ Обновить всё', 'cardsAll')
    .addSeparator()
    .addItem('Товары: список, остатки, заказы', 'refreshProducts')
    .addItem('Воронка: где теряются покупатели', 'funnelReport')
    .addItem('Поисковые запросы и слова для названий', 'queriesReport')
    .addItem('Качество карточек (контент-рейтинг)', 'contentReport')
    .addSubMenu(ui.createMenu('Подключение и автозапуск')
      .addItem('Ключи API', 'setupCredentials')
      .addItem('Проверить подключение', 'testConnection')
      .addItem('Обновлять всё каждый день (~7:00)', 'installDailyTrigger')
      .addItem('Выключить автообновление', 'removeDailyTrigger'))
    .addToUi();
}
function testConnection() {
  run_('Проверка подключения', () => `ок, товаров в кабинете: ${ozon_('/v3/product/list', { filter: { visibility: 'ALL' }, limit: 1 }).result.total}`);
}
function installDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger('cardsAll').timeBased().everyDays(1).atHour(7).create();
  toast_('Ежедневное обновление включено (~7:00)');
}
function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'cardsAll').forEach(t => ScriptApp.deleteTrigger(t));
}
function cardsAll() {
  run_('Обновить всё', () => {
    const out = [];
    const step = (name, fn) => { try { out.push(`${name}: ${fn()}`); } catch (e) { out.push(`${name}: ОШИБКА ${e.message}`); log_(name, 'ERROR', e.stack || e); } };
    step('Товары', refreshProducts_);
    step('Воронка', funnelReport_);
    step('Запросы', queriesReport_);
    step('Качество', contentReport_);
    return out.join(' | ');
  });
}

/* ================== 03_products ================== */
/** Лист «Товары»: всё, что нужно отчётам о товаре, прямо из кабинета */
function refreshProducts() { run_('Товары', refreshProducts_); }
var PRODUCT_COLS = ['Артикул', 'SKU', 'Product ID', 'Название', 'Категория', 'Остаток FBS', 'Заказы 60 дн', 'Обновлено'];

/** Названия типов Ozon: {type_id: «Весы торговые»} из дерева категорий */
function typeNames_() {
  const out = {};
  const walk = list => (list || []).forEach(n => {
    if (n.type_id) out[key_(n.type_id)] = n.type_name || '';
    walk(n.children);
  });
  try { walk(ozon_('/v1/description-category/tree', { language: 'DEFAULT' }).result); }
  catch (e) { log_('Товары', 'WARN', 'типы товаров не получены, категория останется пустой: ' + e.message); }
  return out;
}
/** Заказанные штуки по SKU за days дней — из аналитики Ozon (старый список отправлений Ozon отключил) */
function ordersBySku_(days) {
  const p = funnelPeriods_(days), out = {}, data = analyticsBySku_(p.cur[0], p.cur[1], ['ordered_units']);
  Object.keys(data).forEach(sku => out[sku] = data[sku].ordered_units);
  return out;
}

function refreshProducts_() {
  const list = ozonAll_('/v3/product/list', { filter: { visibility: 'ALL' }, limit: 1000 },
    r => ({ items: r.result.items, next: r.result.last_id }), 'last_id').filter(x => !x.archived);
  const info = {};
  chunk_(list.map(x => x.product_id), 1000).forEach(part =>
    (ozon_('/v3/product/info/list', { product_id: part }).items || []).forEach(it => info[key_(it.id)] = it));
  const stock = {};
  ozonAll_('/v4/product/info/stocks', { filter: { visibility: 'ALL' }, limit: 1000 }, r => ({ items: r.items, next: r.cursor }), 'cursor')
    .forEach(it => (it.stocks || []).filter(s => s.type === 'fbs')
      .forEach(s => stock[key_(it.product_id)] = (stock[key_(it.product_id)] || 0) + (Number(s.present) || 0) - (Number(s.reserved) || 0)));
  const types = typeNames_();
  // без заказов лист всё равно нужен остальным отчётам: сбой здесь не останавливает обновление
  const days = Number(cfg_('ORDERS_DAYS', 60));
  let orders = null;
  try { orders = ordersBySku_(days); }
  catch (e) { log_('Товары', 'WARN', 'заказы не получены, колонка «Заказы 60 дн» пустая: ' + e.message); }
  const now = new Date();
  const rows = list.map(x => {
    const it = info[key_(x.product_id)] || {};
    const sku = it.sku || (it.sources && it.sources[0] && it.sources[0].sku) || '';
    return [String(x.offer_id), sku, x.product_id, it.name || '', types[key_(it.type_id)] || '',
      stock[key_(x.product_id)] || 0, orders ? orders[key_(sku)] || 0 : '', now];
  }).sort((a, b) => (Number(b[6]) || 0) - (Number(a[6]) || 0) || b[5] - a[5]);
  writeReport_(SHEETS.PRODUCTS, PRODUCT_COLS, rows, sh => sh.getRange(2, 8, rows.length, 1).setNumberFormat('dd.MM.yyyy HH:mm'));
  return `товаров: ${rows.length}, с остатком: ${rows.filter(r => r[5] > 0).length}, с заказами за ${days} дн.: ${orders ? rows.filter(r => r[6] > 0).length : 'не получены, см. «Лог»'}`;
}

/** Товары с листа «Товары» как объекты по заголовкам */
function readProducts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  if (!sh || sh.getLastRow() < 2) throw new Error('Лист «Товары» пустой — сначала «🔎 КАРТОЧКИ → Товары»');
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  return sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getValues()
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => { const o = {}; h.forEach((k, j) => o[k] = r[j]); return o; });
}

/* ================== 04_reports ================== */
function funnelReport()   { run_('Воронка', funnelReport_); }
function queriesReport()  { run_('Поисковые запросы', queriesReport_); }
function contentReport()  { run_('Качество карточек', contentReport_); }

/* Сравнение идёт с медианой своей категории (тип товара Ozon),
 * если в категории мало товаров с данными — с медианой всего магазина. */
var FUNNEL_SHEET = 'Воронка', QUERIES_SHEET = 'Поисковые запросы', CONTENT_SHEET = 'Качество карточек';
var FUNNEL_COLS = ['Артикул', 'Название', 'Категория', 'Остаток FBS', 'Показы в поиске и каталоге', 'Посетители карточки',
  'Переход из показа, %', 'Добавили в корзину', 'Карточка → корзина, %', 'Заказано, шт', 'Посетитель → заказ, %',
  'Выручка, ₽', 'Позиция в выдаче', 'Показы к прошлому периоду', 'Заказы к прошлому периоду', 'Диагноз', 'Что проверить'];
var QUERIES_COLS = ['Артикул', 'Название', 'Запрос', 'Искали, польз.', 'Видели товар, польз.', 'Позиция', 'Конверсия, %',
  'Заказы', 'Выручка, ₽', 'Нет в названии', 'Подсказка'];
var CONTENT_COLS = ['Артикул', 'Название', 'Категория', 'Остаток FBS', 'Заказы 60 дн', 'Контент-рейтинг', 'По группам',
  'Что улучшить', 'Заполнить характеристики', 'Приоритет'];

/** Метрики аналитики Ozon в том порядке, в каком их вернёт /v1/analytics/data */
var FUNNEL_METRICS = ['hits_view_search', 'session_view_pdp', 'session_view', 'hits_tocart', 'hits_tocart_pdp',
  'ordered_units', 'revenue', 'position_category'];

/** Диагнозы: порядок в списке = порядок на листе (сначала то, что чинится в карточке) */
var DIAG = {
  CTR:   ['🖼 Видят, но не заходят', 'главное фото, цена в выдаче, рейтинг и число отзывов'],
  CART:  ['📄 Заходят, но не берут', 'фото, описание и характеристики, отзывы, вопросы без ответа, цена рядом с конкурентами'],
  ORDER: ['🛒 Кладут в корзину, но не заказывают', 'срок доставки, цена у конкурентов, наличие'],
  VIEWS: ['🔍 Не видят', 'мало показов: запросы в названии и характеристиках, верная категория, реклама'],
  STOCK: ['⏸ Нет остатка', 'без остатка Ozon товар не показывает'],
  FEW:   ['… мало данных', 'посетителей пока слишком мало для выводов'],
  OK:    ['✓ В норме', '']
};
var DIAG_ORDER = ['CTR', 'CART', 'ORDER', 'VIEWS', 'STOCK', 'FEW', 'OK'];

function ymd_(d) {
  try { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
  catch (e) { return d.toISOString().slice(0, 10); }
}
/** Текущий и предыдущий периоды по days дней, оба заканчиваются вчера / перед текущим */
function funnelPeriods_(days, now) {
  const day = 864e5, end = new Date((now || new Date()).getTime() - day);
  const from = new Date(end.getTime() - (days - 1) * day);
  const pEnd = new Date(from.getTime() - day), pFrom = new Date(pEnd.getTime() - (days - 1) * day);
  return { cur: [ymd_(from), ymd_(end)], prev: [ymd_(pFrom), ymd_(pEnd)] };
}

/** Аналитика Ozon по SKU за период: {sku: {метрика: число}}. Ozon ограничивает частоту — при отказе ждём минуту */
function analyticsBySku_(from, to, metrics) {
  metrics = metrics || FUNNEL_METRICS;
  const out = {};
  for (let offset = 0; offset < 100000; offset += 1000) {
    const body = { date_from: from, date_to: to, metrics, dimension: ['sku'], filters: [],
      sort: [{ key: metrics[0], order: 'DESC' }], limit: 1000, offset };
    let r;
    for (let attempt = 1; ; attempt++) {
      try { r = ozon_('/v1/analytics/data', body); break; }
      catch (e) {
        if (/HTTP 403|premium/i.test(e.message)) throw new Error('Ozon не отдаёт аналитику: проверьте подписку Premium и права ключа API. ' + e.message);
        if (attempt >= 3 || !/5 попыток|429/.test(e.message)) throw e;
        Utilities.sleep(65000);
      }
    }
    const data = (r.result && r.result.data) || [];
    data.forEach(d => {
      const sku = key_(d.dimensions && d.dimensions[0] && d.dimensions[0].id);
      if (!sku) return;
      const m = {}; metrics.forEach((k, i) => m[k] = Number((d.metrics || [])[i]) || 0);
      out[sku] = m;
    });
    if (data.length < 1000) break;
  }
  return out;
}

function median_(arr) {
  const a = arr.filter(x => isFinite(x)).sort((x, y) => x - y);
  if (!a.length) return NaN;
  const i = Math.floor(a.length / 2);
  return a.length % 2 ? a[i] : (a[i - 1] + a[i]) / 2;
}
var ratio_ = (a, b) => b > 0 ? a / b : NaN;
var pctTxt_ = x => isFinite(x) ? (Math.round(x * 1000) / 10).toString().replace('.', ',') + '%' : '—';

/** Конверсии товара из метрик */
function funnelRates_(m) {
  return { ctr: ratio_(m.session_view_pdp, m.hits_view_search), cart: ratio_(m.hits_tocart_pdp, m.session_view_pdp),
           order: ratio_(m.ordered_units, m.hits_tocart) };
}
/** Медианы конверсий по группе товаров (считаются только товары с достаточными данными) */
function funnelMedians_(list, t) {
  return {
    ctr: median_(list.filter(m => m.hits_view_search >= t.minViews).map(m => funnelRates_(m).ctr)),
    cart: median_(list.filter(m => m.session_view_pdp >= t.minVisits).map(m => funnelRates_(m).cart)),
    order: median_(list.filter(m => m.hits_tocart >= t.minCart).map(m => funnelRates_(m).order))
  };
}
/**
 * Диагноз товара: [ключ DIAG, пояснение с цифрами].
 * t — пороги {minViews, minVisits, minCart, weak}; med — медианы категории.
 */
function funnelDiagnosis_(m, stock, med, t) {
  const r = funnelRates_(m);
  const weak = (x, base) => isFinite(x) && isFinite(base) && base > 0 && x < base * t.weak;
  const vs = (x, base) => `${pctTxt_(x)} при ${pctTxt_(base)} в категории`;
  if (!(Number(stock) > 0) && m.hits_view_search < t.minViews) return ['STOCK', ''];
  if (m.hits_view_search < t.minViews) return ['VIEWS', `показов ${m.hits_view_search} за период`];
  if (weak(r.ctr, med.ctr)) return ['CTR', 'переход из показа ' + vs(r.ctr, med.ctr)];
  if (m.session_view_pdp < t.minVisits) return ['FEW', ''];
  if (weak(r.cart, med.cart)) return ['CART', 'в корзину ' + vs(r.cart, med.cart)];
  if (m.hits_tocart >= t.minCart && weak(r.order, med.order)) return ['ORDER', 'заказ из корзины ' + vs(r.order, med.order)];
  return ['OK', ''];
}
/** Изменение к прошлому периоду, если база не крошечная: доля или '' */
function trend_(cur, prev, minBase) { return prev >= minBase ? cur / prev - 1 : ''; }

function funnelReport_() {
  const days = Number(cfg_('FUNNEL_DAYS', 28));
  const t = { minViews: Number(cfg_('FUNNEL_MIN_VIEWS', 300)), minVisits: Number(cfg_('FUNNEL_MIN_VISITS', 30)),
              minCart: 10, weak: Number(cfg_('FUNNEL_WEAK', 0.5)) };
  const p = funnelPeriods_(days);
  const cur = analyticsBySku_(p.cur[0], p.cur[1]);
  const prev = analyticsBySku_(p.prev[0], p.prev[1]);
  if (!Object.keys(cur).length) throw new Error(`Ozon не вернул данных за ${p.cur[0]} — ${p.cur[1]}`);

  const zero = {}; FUNNEL_METRICS.forEach(k => zero[k] = 0);
  const items = readProducts_().filter(r => key_(r['SKU'])).map(r => ({ r, m: cur[key_(r['SKU'])] || zero, pm: prev[key_(r['SKU'])] || zero }));
  // медианы: по категории, если в ней хватает товаров с показами, иначе по магазину
  const all = funnelMedians_(items.map(x => x.m), t), byCat = {};
  const cats = {}; items.forEach(x => (cats[x.r['Категория'] || ''] = cats[x.r['Категория'] || ''] || []).push(x.m));
  Object.keys(cats).forEach(c => {
    const enough = cats[c].filter(m => m.hits_view_search >= t.minViews).length >= 5;
    byCat[c] = enough ? funnelMedians_(cats[c], t) : all;
  });

  const rows = items.map(({ r, m, pm }) => {
    const rt = funnelRates_(m), med = byCat[r['Категория'] || ''];
    const [d, why] = funnelDiagnosis_(m, r['Остаток FBS'], med, t);
    const num = x => isFinite(x) ? x : '';
    return { o: DIAG_ORDER.indexOf(d), views: m.hits_view_search, row: [r['Артикул'], r['Название'], r['Категория'], r['Остаток FBS'],
      m.hits_view_search, m.session_view_pdp, num(rt.ctr), m.hits_tocart, num(rt.cart), m.ordered_units,
      num(ratio_(m.ordered_units, m.session_view)), Math.round(m.revenue), m.position_category ? Math.round(m.position_category) : '',
      trend_(m.hits_view_search, pm.hits_view_search, t.minViews), trend_(m.ordered_units, pm.ordered_units, 5),
      DIAG[d][0], [why, DIAG[d][1]].filter(Boolean).join(' — ')] };
  }).sort((a, b) => a.o - b.o || b.views - a.views);

  writeReport_(FUNNEL_SHEET, FUNNEL_COLS, rows.map(x => x.row), sh => {
    const n = rows.length, c = name => sh.getRange(2, FUNNEL_COLS.indexOf(name) + 1, n, 1);
    ['Показы в поиске и каталоге', 'Посетители карточки', 'Добавили в корзину', 'Выручка, ₽'].forEach(k => c(k).setNumberFormat('#,##0'));
    ['Переход из показа, %', 'Карточка → корзина, %', 'Посетитель → заказ, %'].forEach(k => c(k).setNumberFormat('0.0%'));
    ['Показы к прошлому периоду', 'Заказы к прошлому периоду'].forEach(k => c(k).setNumberFormat('+0%;-0%;0%'));
    const diag = c('Диагноз'), trends = sh.getRange(2, FUNNEL_COLS.indexOf('Показы к прошлому периоду') + 1, n, 2);
    const starts = (text, bg, fc) => SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(text)
      .setBackground(bg).setFontColor(fc).setRanges([diag]).build();
    sh.setConditionalFormatRules([
      starts(DIAG.CTR[0], OZ_UI.bad, OZ_UI.badText), starts(DIAG.CART[0], OZ_UI.bad, OZ_UI.badText),
      starts(DIAG.ORDER[0], OZ_UI.warn, OZ_UI.warnText), starts(DIAG.VIEWS[0], OZ_UI.warn, OZ_UI.warnText),
      starts(DIAG.OK[0], OZ_UI.good, OZ_UI.goodText),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(-0.3).setFontColor(OZ_UI.badText).setRanges([trends]).build()
    ]);
    sh.setColumnWidth(FUNNEL_COLS.indexOf('Диагноз') + 1, 200);
    sh.setColumnWidth(FUNNEL_COLS.indexOf('Что проверить') + 1, 380);
  });
  const cnt = k => rows.filter(x => x.o === DIAG_ORDER.indexOf(k)).length;
  const fell = rows.filter(x => x.row[14] !== '' && x.row[14] <= -0.3).length;
  return `${p.cur[0]} — ${p.cur[1]}: не заходят ${cnt('CTR')}, не берут ${cnt('CART')}, не заказывают ${cnt('ORDER')}, ` +
    `не видят ${cnt('VIEWS')}, в норме ${cnt('OK')}; заказы упали на 30%+ у ${fell}. Лист «${FUNNEL_SHEET}»`;
}

/* ---------- Поисковые запросы ---------- */
var STOP_WORDS_ = ['для', 'и', 'в', 'на', 'с', 'со', 'по', 'от', 'до', 'без', 'под', 'из', 'к', 'а', 'или', 'шт'];
/** Слова для сравнения: нижний регистр, ё→е, длинные русские слова — по первым 5 буквам (грубо отрезаем окончания) */
function stems_(text) {
  return String(text || '').toLowerCase().replace(/ё/g, 'е').split(/[^a-zа-я0-9]+/i)
    .filter(w => w && STOP_WORDS_.indexOf(w) < 0)
    .map(w => /^[а-я]{6,}$/.test(w) ? w.slice(0, 5) : w);
}
/** Слова запроса, которых нет в названии (в исходном написании) */
function missingWords_(query, name) {
  const have = {}; stems_(name).forEach(s => have[s] = true);
  const words = String(query || '').toLowerCase().replace(/ё/g, 'е').split(/[^a-zа-я0-9]+/i).filter(w => w && STOP_WORDS_.indexOf(w) < 0);
  return words.filter(w => !have[stems_(w)[0]]);
}

/** Все страницы одного запроса /v1/analytics/product-queries/details */
function fetchQueries_(skus, from, to, perSku) {
  const out = [];
  let size = 1000;
  for (let page = 0; page < 500; page++) {
    let r;
    const body = { date_from: from, date_to: to, skus, limit_by_sku: perSku, page, page_size: size };
    try { r = ozon_('/v1/analytics/product-queries/details', body); }
    catch (e) {
      if (/HTTP 403|premium/i.test(e.message)) throw new Error('Ozon не отдаёт поисковые запросы: проверьте подписку Premium и права ключа API. ' + e.message);
      if (page === 0 && size > 100 && /HTTP 400/.test(e.message)) { size = 100; page--; continue; }   // большая страница не принята
      throw e;
    }
    const q = r.queries || (r.result && r.result.queries) || [];
    q.forEach(x => out.push(x));
    const pages = Number(r.page_count || (r.result && r.result.page_count)) || 0;
    if (!q.length || page + 1 >= pages) break;
  }
  return out;
}

function queriesReport_() {
  const days = Number(cfg_('FUNNEL_DAYS', 28)), perSku = Math.min(Number(cfg_('QUERIES_PER_SKU', 15)), 15);
  const p = funnelPeriods_(days);
  const from = p.cur[0] + 'T00:00:00Z', to = p.cur[1] + 'T23:59:59Z';
  // только товары, которые есть в наличии или продавались: по остальным запросов всё равно нет
  const bySku = {};
  readProducts_().forEach(r => {
    if (key_(r['SKU']) && (Number(r['Остаток FBS']) > 0 || Number(r['Заказы 60 дн']) > 0)) bySku[key_(r['SKU'])] = r;
  });
  const skus = Object.keys(bySku);
  if (!skus.length) throw new Error('Нет товаров с остатком или заказами — обновите лист «Товары»');
  const raw = [];
  chunk_(skus, 1000).forEach(part => fetchQueries_(part, from, to, perSku).forEach(x => raw.push(x)));
  if (!raw.length) return `Ozon не вернул запросов за ${p.cur[0]} — ${p.cur[1]}`;

  const n = v => Number(v) || 0;
  const rows = raw.map(q => {
    const r = bySku[key_(q.sku)] || {}, name = r['Название'] || q.name || '';
    const miss = missingWords_(q.query, name), searches = n(q.unique_search_users);
    const tip = miss.length && searches >= Number(cfg_('QUERIES_TIP_MIN', 20))
      ? `добавьте «${miss.join(' ')}» в название или описание, если подходит товару` : '';
    return [r['Артикул'] || q.offer_id || q.sku, name, q.query || '', searches, n(q.unique_view_users),
      q.position ? Math.round(n(q.position) * 10) / 10 : '', n(q.view_conversion) / (n(q.view_conversion) > 1 ? 100 : 1),
      n(q.order_count), Math.round(n(q.gmv)), miss.join(' '), tip];
  }).sort((a, b) => String(a[0]).localeCompare(String(b[0])) || b[3] - a[3]);

  writeReport_(QUERIES_SHEET, QUERIES_COLS, rows, sh => {
    const k = rows.length;
    sh.getRange(2, 4, k, 2).setNumberFormat('#,##0');
    sh.getRange(2, 7, k, 1).setNumberFormat('0.0%');
    sh.getRange(2, 9, k, 1).setNumberFormat('#,##0');
    sh.setColumnWidth(3, 240); sh.setColumnWidth(11, 360);
    const tip = sh.getRange(2, 11, k, 1);
    sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty()
      .setBackground(OZ_UI.warn).setFontColor(OZ_UI.warnText).setRanges([tip]).build()]);
  });
  return `${p.cur[0]} — ${p.cur[1]}: запросов ${rows.length} по ${skus.length} товарам, подсказок ${rows.filter(r => r[10]).length}. Лист «${QUERIES_SHEET}»`;
}

/* ---------- Качество карточек: контент-рейтинг Ozon ---------- */
/** Строка отчёта из ответа /v1/product/rating-by-sku для одного товара */
function contentRow_(p) {
  const groups = p.groups || [];
  const weakGroups = groups.filter(g => Number(g.rating) < 100)
    .map(g => `${g.name || g.key} ${Math.round(Number(g.rating) || 0)}`).join(', ');
  const todo = [];
  groups.forEach(g => (g.conditions || []).forEach(c => { if (!c.fulfilled && c.description) todo.push(c.description); }));
  const attrs = [];
  groups.forEach(g => {
    const names = (g.improve_attributes || []).map(a => a.name).filter(Boolean);
    if (!names.length) return;
    const atLeast = Number(g.improve_at_least) || 0;
    attrs.push((atLeast ? `минимум ${atLeast} из: ` : '') + names.join(', '));
  });
  return { rating: Number(p.rating) || 0, weakGroups, todo: todo.join('; '), attrs: attrs.join('; ') };
}

function contentReport_() {
  const main = readProducts_().filter(r => key_(r['SKU']));
  const bySku = {}; main.forEach(r => bySku[key_(r['SKU'])] = r);
  const got = {};
  chunk_(Object.keys(bySku).map(Number).filter(isFinite), 100).forEach(part => {
    const r = ozon_('/v1/product/rating-by-sku', { skus: part });
    (r.products || (r.result && r.result.products) || []).forEach(p => got[key_(p.sku)] = contentRow_(p));
  });
  if (!Object.keys(got).length) throw new Error('Ozon не вернул контент-рейтинг');
  const rows = main.filter(r => got[key_(r['SKU'])]).map(r => {
    const c = got[key_(r['SKU'])], orders = Number(r['Заказы 60 дн']) || 0, stock = Number(r['Остаток FBS']) || 0;
    // приоритет: сколько рейтинга недобрано × насколько товар важен (продажи, иначе хотя бы наличие)
    const prio = Math.round((100 - c.rating) * (1 + orders + (stock > 0 ? 1 : 0)));
    return [r['Артикул'], r['Название'], r['Категория'], r['Остаток FBS'], r['Заказы 60 дн'], c.rating, c.weakGroups, c.todo, c.attrs, prio];
  }).sort((a, b) => b[9] - a[9]);

  writeReport_(CONTENT_SHEET, CONTENT_COLS, rows, sh => {
    const k = rows.length, rt = sh.getRange(2, 6, k, 1);
    rt.setNumberFormat('0');
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(70).setBackground(OZ_UI.bad).setFontColor(OZ_UI.badText).setRanges([rt]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(90).setBackground(OZ_UI.warn).setFontColor(OZ_UI.warnText).setRanges([rt]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(90).setBackground(OZ_UI.good).setFontColor(OZ_UI.goodText).setRanges([rt]).build()
    ]);
    sh.setColumnWidth(7, 220); sh.setColumnWidth(8, 420); sh.setColumnWidth(9, 320);
  });
  const avg = rows.length ? Math.round(rows.reduce((s, r) => s + r[5], 0) / rows.length) : 0;
  return `товаров ${rows.length}, средний рейтинг ${avg}, ниже 70: ${rows.filter(r => r[5] < 70).length}. Лист «${CONTENT_SHEET}»`;
}

/** Перезаписывает лист-отчёт: заголовок, строки, фильтр; format(sh) — форматы колонок, если строки есть */
function writeReport_(name, cols, rows, format) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getFilter()) sh.getFilter().remove();             // clear() фильтр не снимает
  sh.clear();
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold')
    .setBackground(OZ_UI.header).setFontColor(OZ_UI.headerText).setFontSize(9).setWrap(true);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 280);
  if (!rows.length) return sh;
  sh.getRange(2, 1, rows.length, cols.length).setValues(rows);
  if (format) format(sh);
  sh.getRange(1, 1, rows.length + 1, cols.length).createFilter();
  return sh;
}
