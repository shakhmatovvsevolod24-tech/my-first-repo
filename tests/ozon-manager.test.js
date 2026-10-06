/**
 * Проверки скрипта apps-script/ozon-manager.gs без Google: сервисы Apps Script подменены заглушками.
 * Запуск: node --test tests/
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'ozon-manager.gs'), 'utf8');
/** Объекты из песочницы vm — с чужими прототипами; для сравнения приводим к обычным */
const plain = x => JSON.parse(JSON.stringify(x));

/* ---------- Заглушки Apps Script ---------- */
class Props {
  constructor() { this.m = {}; }
  getProperty(k) { return k in this.m ? this.m[k] : null; }
  setProperty(k, v) { this.m[k] = String(v); return this; }
  setProperties(o) { Object.assign(this.m, o); return this; }
  deleteProperty(k) { delete this.m[k]; return this; }
}

/** Лист как двумерный массив; строки и колонки с 1, как в Apps Script */
class Sheet {
  constructor(rows) { this.rows = rows.map(r => r.slice()); }
  getLastRow() { let n = 0; this.rows.forEach((r, i) => { if (r.some(v => v !== '' && v !== null && v !== undefined)) n = i + 1; }); return n; }
  getLastColumn() { return Math.max(0, ...this.rows.map(r => r.length)); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {                       // только вид «B1»
      const m = a.match(/^([A-Z]+)(\d+)$/);
      const col = m[1].split('').reduce((s, ch) => s * 26 + ch.charCodeAt(0) - 64, 0);
      return new Range(this, Number(m[2]), col, 1, 1);
    }
    return new Range(this, a, b, c || 1, d || 1);
  }
  getName() { return this.name || ''; }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  cell(r, c) { const row = this.rows[r - 1] || []; const v = row[c - 1]; return v === undefined ? '' : v; }
  put(r, c, v) { while (this.rows.length < r) this.rows.push([]); const row = this.rows[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
}
class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.cell(this.r + i, this.c + j)); out.push(row); } return out; }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.put(this.r + i, this.c + j, x))); return this; }
  getValue() { return this.getValues()[0][0]; }
  getFormulas() { return this.getValues().map(r => r.map(v => (typeof v === 'string' && v.startsWith('=') ? v : ''))); }
  getDisplayValues() { return this.getValues().map(r => r.map(v => String(v))); }
  setValue(x) { return this.setValues([[x]]); }
}

/** Меню: запоминаем, что построил onOpen */
function mockUi(menus) {
  const builder = name => {
    const m = { name, items: [] };
    const api = {
      addItem(label, fn) { m.items.push({ label, fn }); return api; },
      addSeparator() { m.items.push({ sep: true }); return api; },
      addSubMenu(sub) { m.items.push({ sub: sub._m }); return api; },
      addToUi() { menus.push(m); },
      _m: m
    };
    return api;
  };
  return { createMenu: builder, alert: () => 'YES', ButtonSet: { YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' }, Button: { YES: 'YES', OK: 'OK' } };
}

/** Загружает скрипт в отдельный контекст; overrides подменяют функции скрипта после загрузки */
function load({ settings = {}, fetch } = {}) {
  const logs = [], menus = [], docProps = new Props(), scriptProps = new Props();
  scriptProps.setProperties({ OZON_CLIENT_ID: '1', OZON_API_KEY: 'k' });
  const ctx = {
    console: { log: m => logs.push(m) },
    PropertiesService: { getDocumentProperties: () => docProps, getScriptProperties: () => scriptProps },
    SpreadsheetApp: {
      getActive: () => ({ getSheetByName: () => null, toast() {} }),
      getUi: () => mockUi(menus),
      flush() {}
    },
    UrlFetchApp: { fetch: fetch || (() => { throw new Error('сеть не нужна в этом тесте'); }) },
    Utilities: { sleep() {}, formatDate: d => new Date(d).toISOString() },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => 'Europe/Moscow' },
    ScriptApp: { getProjectTriggers: () => [] }
  };
  vm.createContext(ctx);
  vm.runInContext(SRC, ctx);
  const cfg = Object.assign({ DRY_RUN: true, MIN_MARGIN: 0.10, PROMO_MARGIN: 0.12 }, settings);
  ctx.cfg_ = (k, fb) => (k in cfg ? cfg[k] : fb);
  ctx.log_ = (op, level, msg) => logs.push(`[${level}] ${op}: ${msg}`);
  return { ctx, logs, menus, docProps, cfg };
}

/* ---------- Лист «Бог акций» ---------- */
const BOG_HDR = ['Product ID', 'Артикул', 'Название', 'Категория', 'Остаток', 'Заказы 60 дн', 'Цена продажи, ₽',
  'Мин. цена, ₽', 'Порог акций, ₽', 'Цена в акции сейчас, ₽', 'Макс. цена акции, ₽', 'Скидка от цены, %',
  'Прибыль в акции, ₽', 'Маржа в акции, %', 'Потеря прибыли, ₽', 'Статус', 'Проходит?', 'Действие', 'Результат'];
/** row: [pid, art, мин. цена, порог, макс. цена акции, действие] */
function bogSheet(title, rows) {
  const data = [['Акция ▶', title], BOG_HDR];
  rows.forEach(([pid, art, min, floor, max, act]) => {
    const r = new Array(BOG_HDR.length).fill('');
    r[0] = pid; r[1] = art; r[4] = 3; r[7] = min; r[8] = floor; r[10] = max; r[17] = act;
    data.push(r);
  });
  return new Sheet(data);
}
const resultOf = (sh, art) => { const r = sh.rows.find(x => x[1] === art); return r[18]; };
const actionOf = (sh, art) => { const r = sh.rows.find(x => x[1] === art); return r[17]; };

/* ---------- Меню ---------- */
test('onOpen: пять отдельных меню, у каждой кнопки есть функция', () => {
  const { ctx, menus } = load();
  ctx.onOpen();
  assert.deepEqual(menus.map(m => m.name), ['💰 ЦЕНЫ', '✍ РУЧНАЯ НАСТРОЙКА', '🏷 АКЦИИ', '📦 ОСТАТКИ', '⚙ НАСТРОЙКИ']);
  const flat = list => list.flatMap(i => i.sub ? flat(i.sub.items) : [i]);
  const items = menus.flatMap(m => flat(m.items).filter(i => i.fn));
  menus.forEach(m => assert.ok(m.items.filter(i => i.fn || i.sub).length <= 7, `в меню «${m.name}» больше 7 пунктов`));
  items.forEach(i => assert.equal(typeof ctx[i.fn], 'function', `нет функции ${i.fn} для «${i.label}»`));
  const fns = items.map(i => i.fn);
  ['applySelectedForce', 'removeAllFromActions', 'setupSheetUi', 'analyzeShippingSpeed', 'compareCosts']
    .forEach(fn => assert.ok(!fns.includes(fn), `кнопка ${fn} должна быть убрана`));
  const labels = items.map(i => i.label);
  assert.ok(labels.includes('Применить отметки с маржой от 10%'));
  assert.ok(labels.includes('Выгрузить цены по марже 12%'));
  assert.ok(labels.includes('Включить автообновление (каждые 5 мин)'));
  assert.ok(labels.includes('Обновить закуп вручную у выделенных'));
  const manual = menus.find(m => m.name === '✍ РУЧНАЯ НАСТРОЙКА');
  assert.equal(manual.items[0].fn, 'addMissingProducts', 'загрузка новых товаров — первой кнопкой в «Ручной настройке»');
  assert.equal(manual.items[0].label, 'Загрузить новые товары в таблицу');
  assert.equal(new Set(fns).size, fns.length, 'одна функция висит на двух кнопках');
});

test('onOpen: подписи берут проценты и интервал из «Настроек»', () => {
  const { ctx, menus } = load({ settings: { MIN_MARGIN: 0.11, PROMO_MARGIN: 0.15, STOCKS_REFRESH_MIN: 7 } });
  ctx.onOpen();
  const flat = list => list.flatMap(i => i.sub ? flat(i.sub.items) : [i]);
  const labels = menus.flatMap(m => flat(m.items).filter(i => i.fn).map(i => i.label));
  assert.ok(labels.includes('Применить отметки с маржой от 11%'));
  assert.ok(labels.includes('Применить отметки «Действие» (маржа от 15%)'));
  assert.ok(labels.includes('Включить автообновление (каждые 5 мин)'), '7 минут Google не принимает — берём 5');
});

/* ---------- Ручные отметки на «Бог акций» ---------- */
const ROWS = [
  [101, 'A-ok', 100, 110, 120, 'Добавить'],      // выше порога 12%
  [102, 'B-mid', 100, 110, 105, 'Добавить'],     // между 10% и 12%
  [103, 'C-low', 100, 110, 95, 'Добавить'],      // ниже 10%
  [104, 'D-nofloor', '', 'нереально', 90, 'Добавить'],  // порог не посчитан
  [105, 'E-del', 100, 110, 130, 'Удалить']
];
function setupBog(opts) {
  const env = load(opts);
  const sh = bogSheet('Акция X', ROWS);
  env.ctx.sheet_ = () => sh;
  env.ctx.listActions_ = () => [{ id: 7, title: 'Акция X', action_type: 'DISCOUNT' }];
  env.refreshed = 0;
  env.ctx.refreshActions_ = () => { env.refreshed++; return ''; };
  env.calls = [];
  env.ctx.ozon_ = (p, body) => { env.calls.push({ p, body }); return { result: { product_ids: [], rejected: [] } }; };
  return Object.assign(env, { sh });
}

test('обычные отметки: только от 12%, без порога — не добавляем', () => {
  const { ctx, sh, calls } = setupBog();
  const msg = ctx.applySelectedActions_();
  assert.match(msg, /^\[ПРОВЕРКА\] добавить: 1, убрать: 1/);
  assert.match(resultOf(sh, 'A-ok'), /^ПРОВЕРКА.*: добавили бы по 120$/);
  assert.match(resultOf(sh, 'B-mid'), /^✗ ниже порога 110 .*маржой от 10%/);
  assert.match(resultOf(sh, 'C-low'), /^✗ ниже порога 110/);
  assert.match(resultOf(sh, 'D-nofloor'), /^✗ нет порога/);
  assert.equal(calls.length, 0, 'в режиме проверки в Ozon ничего не уходит');
});

test('отметки с маржой от 10%: пропускает 10–12%, ниже 10% — никогда', () => {
  const { ctx, sh } = setupBog();
  const msg = ctx.applySelectedActions_('min');
  assert.match(msg, /\[МАРЖА ОТ 10%\] добавить: 2 \(из них с маржой ниже 12%: 1\), убрать: 1/);
  assert.match(resultOf(sh, 'A-ok'), /^ПРОВЕРКА.*: добавили бы по 120$/);
  assert.match(resultOf(sh, 'B-mid'), /^ПРОВЕРКА.*: добавили бы по 105 \(маржа ниже 12%, но не ниже 10%/);
  assert.match(resultOf(sh, 'C-low'), /^✗ ниже мин. цены 100/);
  assert.match(resultOf(sh, 'D-nofloor'), /^✗ нет порога/);
});

test('боевой режим: ручное решение запоминается, «Удалить» его снимает', () => {
  const env = setupBog({ settings: { DRY_RUN: false } });
  const { ctx, sh, calls, docProps } = env;
  docProps.setProperty('PROMO_MIN_OK', JSON.stringify(['7:105']));      // E-del раньше добавляли вручную
  ctx.applySelectedActions_('min');
  const act = calls.find(c => c.p === '/v1/actions/products/activate');
  assert.deepEqual(plain(act.body.products.map(p => [p.product_id, p.action_price])), [[101, 120], [102, 105]]);
  assert.deepEqual(plain(calls.find(c => c.p === '/v1/actions/products/deactivate').body.product_ids), [105]);
  assert.deepEqual(JSON.parse(docProps.getProperty('PROMO_MIN_OK')), ['7:102']);
  assert.match(resultOf(sh, 'B-mid'), /^✓ в акции по 105 \(маржа ниже 12%/);
  assert.equal(actionOf(sh, 'B-mid'), '', 'успешная отметка снимается');
  assert.equal(actionOf(sh, 'C-low'), 'Добавить', 'отказ оставляет отметку');
  assert.equal(env.refreshed, 1);
});

/* ---------- Чистка акций ---------- */
function setupClean(opts) {
  const env = load(opts);
  env.ctx.readMain_ = () => ({ rows: [
    { 'Product ID': 201, 'Артикул': 'P-manual', 'Порог акций, ₽': 110, 'Мин. цена, ₽': 100, 'Остаток FBS': 1 },
    { 'Product ID': 202, 'Артикул': 'P-low', 'Порог акций, ₽': 110, 'Мин. цена, ₽': 100, 'Остаток FBS': 1 },
    { 'Product ID': 203, 'Артикул': 'P-good', 'Порог акций, ₽': 110, 'Мин. цена, ₽': 100, 'Остаток FBS': 1 }
  ] });
  env.ctx.listActions_ = () => [{ id: 1, title: 'X' }, { id: 2, title: 'Y' }];
  const inside = {
    1: [{ id: 201, action_price: 105 }, { id: 202, action_price: 95 }, { id: 203, action_price: 115 }],
    2: [{ id: 201, action_price: 105 }]
  };
  env.ctx.fetchActionProducts_ = id => inside[id];
  env.refreshed = 0;
  env.ctx.refreshActions_ = () => { env.refreshed++; return ''; };
  env.calls = [];
  env.ctx.ozon_ = (p, body) => { env.calls.push({ p, body }); return { result: { product_ids: body.product_ids, rejected: [] } }; };
  // 201 вручную добавлен в X (не в Y), 202 — тоже, но цена ниже 10%; 999 — уже не в акции
  env.docProps.setProperty('PROMO_MIN_OK', JSON.stringify(['1:201', '1:202', '2:999']));
  return env;
}

test('ночная чистка: ручное решение живёт до 10%, остальное — до 12%', () => {
  const env = setupClean({ settings: { DRY_RUN: false } });
  const msg = env.ctx.removeIneligible_();
  const byAction = Object.fromEntries(env.calls.map(c => [c.body.action_id, c.body.product_ids]));
  assert.deepEqual(plain(byAction), { 1: [202], 2: [201] }, 'X: убираем только P-low; Y: P-manual без ручного решения');
  assert.deepEqual(JSON.parse(env.docProps.getProperty('PROMO_MIN_OK')), ['1:201'], 'снятые и устаревшие решения удалены');
  assert.match(msg, /ниже порога: 2/);
  assert.ok(env.logs.some(l => /Оставлены по ручному решению \(маржа от 10%\): 1/.test(l)));
});

test('ночная чистка в режиме проверки ничего не меняет', () => {
  const env = setupClean();
  const msg = env.ctx.removeIneligible_();
  assert.equal(env.calls.length, 0);
  assert.match(msg, /^\[ПРОВЕРКА\] ниже порога: 2/);
  assert.deepEqual(JSON.parse(env.docProps.getProperty('PROMO_MIN_OK')), ['1:201', '1:202'], 'убрали только решение по товару, которого нет в акции');
});

/* ---------- Распределение по акциям ---------- */
test('распределение не выкидывает ручное решение от 10%', () => {
  const env = setupClean({ settings: { DRY_RUN: true, ACTIONS_EXCLUDE: '' } });
  env.ctx.fetchActionCandidates_ = () => [];
  env.ctx.fetchActionProducts_ = id => id === 1 ? [{ id: 201, action_price: 105, max_action_price: 105 }] : [];
  env.ctx.writeBogNotes_ = () => {};
  env.ctx.distributeBestActions_();
  assert.ok(env.logs.some(l => /уже стоят верно 1/.test(l) && /убираем совсем 0/.test(l)), env.logs.join('\n'));
});

/* ---------- Цены ---------- */
test('цена ниже мин. цены не уходит, даже если в «Настройках» осталось ALLOW_BELOW_MIN = TRUE', () => {
  const { ctx } = load({ settings: { ALLOW_BELOW_MIN: true, MAX_PRICE_CHANGE: 0.3 } });
  const res = {};
  const item = (row, art, price) => ({ _row: row, 'Артикул': art, 'Product ID': row, 'Отправить': true, 'Цена к отправке, ₽': price,
    'Зачёркнутая, ₽': 0, 'min_price, ₽': 100, 'Цена на Ozon, ₽': 95 });
  ctx.readMain_ = () => ({ h: ['Артикул', 'Отправить', 'Результат'], rows: [item(3, 'low', 90), item(4, 'ok', 100)] });
  ctx.mainPatch_ = (m, name, patch) => { if (name === 'Результат') Object.assign(res, patch); };
  ctx.uploadPrices_('checked');
  assert.equal(res[3], '✗ ниже мин. цены 100');
  assert.match(res[4], /^ПРОВЕРКА: уйдёт 100 /);
});

/* ---------- Аудит ---------- */
test('аудит: снимает только акции ниже своего порога, ручное решение сверяет с 10%', () => {
  const env = load({ settings: { DRY_RUN: false, AUDIT_TOLERANCE: 0.02, ACQUIRING_RATE: 0.02, PACKAGING_RUB: 0,
    DEFAULT_BUYOUT: 1, AUDIT_REMOVE_FROM_ACTIONS: true, AUDIT_ZERO_STOCK: false } });
  const { ctx } = env;
  // закуп 68, комиссия 20% + эквайринг 2%, логистики нет → маржа = 0.78 − 68/цена; 10% = 100 ₽, 12% ≈ 104 ₽
  const row = (pid, art) => ({ 'Product ID': pid, 'Артикул': art, 'SKU': pid, 'Закуп, ₽': 68, 'Выкуп, %': 1,
    'Маржа, %': 0.2, 'Комиссия, %': 0.2, 'Логистика, ₽': 0, 'Цена на Ozon, ₽': 130,
    'Порог акций, ₽': 104, 'Мин. цена, ₽': 100, 'Остаток FBS': 5 });
  ctx.readMain_ = () => ({ rows: [row(301, 'manual-ok'), row(302, 'plain-low'), row(303, 'manual-too-low')] });
  ctx.ensureAuditSheet_ = () => {};
  ctx.auditPrevStock_ = () => ({});
  let written = [];
  ctx.writeTable_ = (name, rows) => { written = rows; };
  ctx.ozonAll_ = () => [301, 302, 303].map(id => ({ product_id: id, price: { price: 130 },
    commissions: { sales_percent_fbs: 20, fbs_direct_flow_trans_max_amount: 0, fbs_first_mile_max_amount: 0, fbs_deliv_to_customer_amount: 0 } }));
  ctx.listActions_ = () => [{ id: 1, title: 'X' }, { id: 2, title: 'Z' }, { id: 3, title: 'Y' }];
  const inside = {
    1: [{ id: 301, action_price: 101 }, { id: 303, action_price: 90 }],   // X: ручные решения
    2: [{ id: 301, action_price: 120 }, { id: 302, action_price: 120 }],  // Z: выше порога
    3: [{ id: 302, action_price: 95 }]                                    // Y: ниже порога
  };
  ctx.fetchActionProducts_ = id => inside[id];
  env.docProps.setProperty('PROMO_MIN_OK', JSON.stringify(['1:301', '1:303']));
  const calls = [];
  ctx.ozon_ = (p, body) => { calls.push(body); return { result: { product_ids: body.product_ids } }; };
  ctx.runAudit_();
  const removed = calls.map(b => [b.action_id, b.product_ids]).sort();
  assert.deepEqual(plain(removed), [[1, [303]], [3, [302]]], 'Z (120 ₽) не трогаем, ручное решение 301 на 101 ₽ оставляем');
  const status = Object.fromEntries(written.map(r => [r['Артикул'], r['Статус проверки']]));
  assert.equal(status['manual-ok'], '✓ в акции от мин. маржи (ручное решение)');
  assert.equal(status['plain-low'], '⚠ в акции ниже порога');
});

/* ---------- Сеть ---------- */
test('ozon_: повторяет запрос после сбоя сети', () => {
  let n = 0;
  const { ctx } = load({ fetch: () => {
    if (++n < 3) throw new Error('Адрес недоступен: https://api-seller.ozon.ru/...');
    return { getResponseCode: () => 200, getContentText: () => '{"ok":1}' };
  } });
  assert.deepEqual(plain(ctx.ozon_('/v1/test', {})), { ok: 1 });
  assert.equal(n, 3);
});

test('ozon_: после 5 сбоев подряд отдаёт ошибку', () => {
  let n = 0;
  const { ctx } = load({ fetch: () => { n++; throw new Error('Адрес недоступен'); } });
  assert.throws(() => ctx.ozon_('/v1/test', {}), /Адрес недоступен/);
  assert.equal(n, 5);
});

/* ---------- Закуп ---------- */
const SOURCES = [
  { 'Источник': 'Прайс POSCENTER', 'ID таблицы': 'ID_POS', 'Лист': 'Лист1', 'Колонка кода': 'A', 'Колонка закупа': 'B', 'Валюта': 'RUB' },
  { 'Источник': 'Прайс MERTECH', 'ID таблицы': 'ID_MER', 'Лист': 'Лист1', 'Колонка кода': 'A', 'Колонка закупа': 'B', 'Валюта': 'RUB' }
];
function priceBook(rows) { const sh = new Sheet(rows); sh.name = 'Лист1'; return { getSheets: () => [sh] }; }

test('закуп: код ищется только в своём прайсе, чужой прайс не подставляется', () => {
  const env = load({ settings: { COST_1C_SHEET_ID: '' } });
  const { ctx } = env;
  ctx.writeCostCheck_ = () => 0; ctx.appendCostHistory_ = () => {};
  ctx.SpreadsheetApp.openById = id => ({ ID_POS: priceBook([['4865', 21678], ['4863', 16300]]), ID_MER: priceBook([['9999', 500]]) })[id];
  ctx.readTable_ = () => ({ rows: SOURCES });
  ctx.readMain_ = () => ({ rows: [
    { _row: 3, 'Артикул': '4865- MERTECH 2310 P2D', 'Источник закупа': 'Прайс MERTECH', 'Код в прайсе': 4865, 'Product ID': 1 },
    { _row: 4, 'Артикул': '4863 - РИТЕЙЛ-02Ф', 'Источник закупа': 'Прайс POSCENTER', 'Код в прайсе': 4863, 'Product ID': 2 },
    { _row: 5, 'Артикул': 'только 1С', 'Источник закупа': '1С', 'Код в прайсе': 4863, 'Product ID': 3 }
  ] });
  const cost = {};
  ctx.mainPatch_ = (m, name, patch) => { if (name === 'Закуп, ₽') Object.assign(cost, patch); };
  ctx.importCosts_();
  assert.deepEqual(plain(cost), { 3: '', 4: 16300, 5: '' }, 'MERTECH-товар не берёт цену POSCENTER; «1С» не лезет в прайсы');
  assert.ok(env.logs.some(l => /Кода нет в своём прайсе.*4865- MERTECH 2310 P2D/.test(l)));
});

test('источники закупа из старой таблицы: по ID прайса в формуле', () => {
  const env = load({ settings: { COST_1C_SHEET_ID: 'ID_1C' } });
  const { ctx } = env;
  const old = new Sheet([
    ['555'],
    ['Артикул', 'SKU', 'Product ID', 'Прайс', 'Название', 'Категория', 'Поставщик', 'Остатки', 'OZON Карта', 'Закупочная цена'],
    ['A-pos', '', 11, 4538, '', '', '', '', '', '=IFERROR(VLOOKUP(D3,IMPORTRANGE("https://docs.google.com/spreadsheets/d/ID_POS/","Лист1!a:d"),4,0))'],
    ['A-1c', '', 12, '', '', '', '', '', '', '=VLOOKUP(C4,IMPORTRANGE("https://docs.google.com/spreadsheets/d/ID_1C/","Prices!A:ZZ"),5,0)'],
    ['A-mix', '', 13, 1, '', '', '', '', '', '=MAX(IMPORTRANGE("ID_POS","a"),IMPORTRANGE("ID_MER","a"))'],
    ['A-man', '', 14, 7, '', '', '', '', '', '=IMPORTRANGE("ID_MER","a")'],
    ['A-same', '', 15, 8, '', '', '', '', '', '=IMPORTRANGE("ID_MER","a")']
  ]);
  ctx.SpreadsheetApp.openById = () => ({ getSheetByName: () => old });
  ctx.readTable_ = () => ({ rows: SOURCES });
  ctx.readMain_ = () => ({ rows: [
    { _row: 3, 'Артикул': 'A-pos', 'Product ID': 11, 'Источник закупа': '1С', 'Код в прайсе': '' },
    { _row: 4, 'Артикул': 'A-1c', 'Product ID': 12, 'Источник закупа': 'Прайс MERTECH', 'Код в прайсе': 5 },
    { _row: 5, 'Артикул': 'A-mix', 'Product ID': 13, 'Источник закупа': '1С', 'Код в прайсе': 1 },
    { _row: 6, 'Артикул': 'A-man', 'Product ID': 14, 'Источник закупа': 'вручную', 'Код в прайсе': 7 },
    { _row: 7, 'Артикул': 'A-same', 'Product ID': 99, 'Источник закупа': 'Прайс MERTECH', 'Код в прайсе': 8 }
  ] });
  const patches = {};
  ctx.mainPatch_ = (m, name, patch) => { patches[name] = plain(patch); };
  const msg = ctx.sourcesFromOldTable_();
  assert.deepEqual(patches['Источник закупа'], { 3: 'Прайс POSCENTER', 4: '1С' });
  assert.deepEqual(patches['Код в прайсе'], { 3: 4538 });
  assert.match(msg, /источник поменян у 2, совпадал у 1, не тронуты .*: 1/);
});

/* ---------- Удаление из акции ---------- */
function withSelection(env, rows) {
  const ranges = rows.map(([a, b]) => ({ getRow: () => a, getLastRow: () => b }));
  env.ctx.SpreadsheetApp.getActive = () => ({
    getActiveSheet: () => ({ getName: () => 'Бог акций' }),
    getActiveRangeList: () => ({ getRanges: () => ranges }),
    getSheetByName: () => null, toast() {}
  });
}

test('удалить выделенные: только удаление, «Добавить» не трогаем', () => {
  const env = setupBog();
  withSelection(env, [[4, 5]]);                         // строки B-mid и C-low
  const msg = env.ctx.removeMarked_();
  assert.match(msg, /добавить: 0, убрать: 3/, 'две выделенные + E-del с отметкой');
  assert.match(resultOf(env.sh, 'B-mid'), /убрали бы из акции$/);
  assert.equal(resultOf(env.sh, 'A-ok'), '', 'отметка «Добавить» не обработана');
  assert.equal(actionOf(env.sh, 'A-ok'), 'Добавить');
});

test('боевой режим: результат виден после перестройки листа', () => {
  const env = setupBog({ settings: { DRY_RUN: false } });
  withSelection(env, []);
  env.ctx.refreshActions_ = () => { env.sh.rows.slice(2).forEach(r => { r[18] = ''; r[17] = ''; }); return ''; };
  env.ctx.ozon_ = (p, body) => ({ result: { product_ids: body.product_ids || [], rejected: [] } });
  env.ctx.removeMarked_();
  assert.equal(resultOf(env.sh, 'E-del'), '✓ убран');
});

/* ---------- Формулы листа Ozon ---------- */
test('цена факт. берёт цену в акции, если она ниже цены продажи (как в старой таблице)', () => {
  const { ctx } = load();
  const f = ctx.mainFormulas_(ctx.MAIN_COLS || vm.runInContext('MAIN_COLS', ctx), 3)['Цена факт., ₽'];
  const col = n => '$' + ctx.letter_(vm.runInContext('MAIN_COLS', ctx).indexOf(n) + 1) + '3';
  assert.ok(f.includes(`N(${col('Мин. цена в акциях, ₽')})<`), f);
  assert.ok(f.includes(col('Цена продажи, ₽')), f);
});

test('Ozon-карты нет ни в колонках, ни в формулах', () => {
  const { ctx } = load();
  const cols = vm.runInContext('MAIN_COLS.concat(TARIFF_COLS)', ctx);
  assert.ok(!cols.some(c => /карт/i.test(c)), cols.join(', '));
  const F = ctx.mainFormulas_(vm.runInContext('MAIN_COLS', ctx), 3);
  assert.ok(!Object.values(F).some(f => /карт/i.test(f)));
  assert.ok(!/карт/i.test(SRC.replace(/\/\/.*|\/\*[\s\S]*?\*\//g, '').replace(/'Ozon-карта'|'Ozon-карта, %'|'Скидка Ozon-карта, %'/g, '')));
});

test('закуп 1С: пустая колонка сегодняшней даты пропускается, берётся последняя заполненная', () => {
  const env = load({ settings: { COST_1C_SHEET_ID: 'ID_1C', COST_1C_SHEET: 'Prices' } });
  const { ctx } = env;
  ctx.writeCostCheck_ = () => 0; ctx.appendCostHistory_ = () => {};
  const today = '2026.10.01';
  ctx.Utilities.formatDate = () => today;
  const c1 = new Sheet([['product_id', '2026.09.30', today], [1, 500, ''], [2, 700, '']]);
  ctx.SpreadsheetApp.openById = id => (id === 'ID_1C' ? { getSheetByName: () => c1 } : null);
  ctx.readTable_ = () => ({ rows: [] });
  ctx.readMain_ = () => ({ rows: [{ _row: 3, 'Артикул': 'a', 'Источник закупа': '1С', 'Product ID': 1 },
                                  { _row: 4, 'Артикул': 'b', 'Источник закупа': '1С', 'Product ID': 2 }] });
  const cost = {};
  ctx.mainPatch_ = (m, name, patch) => { if (name === 'Закуп, ₽') Object.assign(cost, patch); };
  ctx.importCosts_();
  assert.deepEqual(plain(cost), { 3: 500, 4: 700 });
});

test('лист Ozon: цена безубыточности и ROI', () => {
  const { ctx } = load();
  const cols = vm.runInContext('MAIN_COLS', ctx);
  const F = ctx.mainFormulas_(cols, 3);
  const col = n => '$' + ctx.letter_(cols.indexOf(n) + 1) + '3';
  assert.match(F['Цена безубыточности, ₽'], /\(1-0-/);
  assert.equal(F['ROI, %'], `=IF(OR(${col('Прибыль, ₽')}="",N(${col('Закуп, ₽')})<=0),"",${col('Прибыль, ₽')}/${col('Закуп, ₽')})`);
});

test('ABC: группы по прибыли, убыточные отдельно', () => {
  const { ctx } = load();
  const g = ctx.abcGroups_([{ key: 'a', p: 70 }, { key: 'b', p: 20 }, { key: 'c', p: 7 }, { key: 'd', p: 3 }, { key: 'e', p: -5 }], 'p');
  assert.deepEqual(plain(['a', 'b', 'c', 'd', 'e'].map(k => g[k].g)), ['A', 'A', 'B', 'C', 'C ⛔']);
});

test('«Обновить структуру» не стирает первую строку данных «План-факта»', () => {
  const { ctx } = load();
  const sh = new Sheet([['SKU', 'Артикул', 'Выручка, ₽'], [1272218530, '=OLD()', 1069445], [2, '', 5]]);
  sh.getMaxColumns = () => 4;
  sh.setFontFamily = () => sh;
  const origRange = sh.getRange.bind(sh);
  sh.getRange = (...a) => { const r = origRange(...a); r.setFontFamily = () => r; r.setFontSize = () => r;
    r.clearContent = () => { r.setValues(r.getValues().map(x => x.map(() => ''))); return r; };
    r.setFormula = f => { r.setValue(f); return r; };
    r.getFormulasR1C1 = () => r.getFormulas(); return r; };
  ctx.sheet_ = () => sh;
  ctx.ensureFormulaRow_('План-факт', ['SKU', 'Артикул', 'Выручка, ₽'], { 'Артикул': '=NEW()' });
  assert.equal(sh.cell(2, 1), 1272218530);
  assert.equal(sh.cell(2, 3), 1069445);
  assert.equal(sh.cell(2, 2), '=NEW()');
});

test('демпинг: ниже закупа, ниже безубыточности, большой разрыв', () => {
  const { ctx } = load();
  const v = (...a) => { const r = ctx.dumpingVerdict_(...a); return r && plain(r); };
  assert.deepEqual(v(3269, 115, 829, 100, 0.2), [0, '⛔ ниже нашего закупа']);
  assert.deepEqual(v(3269, 115, 829, 502, 0.2), [1, '⚠ ниже нашей безубыточности']);
  assert.deepEqual(v(1000, 100, 300, 500, 0.2), [2, '▲ дешевле нас на 100%']);
  assert.equal(v(1000, 100, 300, 900, 0.2), null);
  assert.equal(v(3269, 115, 829, 0, 0.2, 234), null, 'другие площадки ниже безубыточности, но выше закупа — не демпинг');
  assert.deepEqual(v(3269, 115, 829, 0, 0.2, 100), [0, '⛔ ниже нашего закупа']);
});

test('AN/AO: разница только с Ozon и по цене для покупателя', () => {
  const { ctx } = load();
  const F = ctx.mainFormulas_(vm.runInContext('MAIN_COLS', ctx), 3);
  assert.ok(!F['Разница с Ozon'].includes('Другие'), F['Разница с Ozon']);
  const cols = vm.runInContext('MAIN_COLS', ctx);
  const col = n => '$' + ctx.letter_(cols.indexOf(n) + 1) + '3';
  assert.ok(F['Разница с Ozon'].includes(col('Мин. цена в акциях, ₽')));
  assert.ok(!F['Разница с Ozon'].includes(col('Другие площадки, ₽')));
  assert.equal(cols.indexOf('Разница с Ozon') + 1, 40, 'колонка AN');
  assert.equal(cols.indexOf('Позиция на Ozon') + 1, 41, 'колонка AO');
});

test('логистика: факт из «План-факта», иначе тариф', () => {
  const { ctx } = load();
  const cols = vm.runInContext('MAIN_COLS', ctx);
  const F = ctx.mainFormulas_(cols, 3);
  assert.match(F['Логистика с выкупом, ₽'], /План-факт/);
  assert.match(F['Логистика: основа'], /"факт"/);
  const pf = ctx.pfFormulas_(2)['Логистика факт/шт, ₽'];
  assert.match(pf, /LOGISTICS_FACT_MIN_SALES/);
  const t = { 'Комиссия FBS, %': 50, 'Логистика FBS мин, ₽': 100, 'Логистика FBS макс, ₽': 500, 'Обработка FBS, ₽': 0, 'Последняя миля FBS, ₽': 0 };
  const byFact = ctx.priceForMargin_({ 'Закуп, ₽': 100, 'Выкуп, %': 1, 'Логистика: основа': 'факт', 'Логистика с выкупом, ₽': 200 }, t, 0, 0, 0, 1, 'MAX');
  const byTariff = ctx.priceForMargin_({ 'Закуп, ₽': 100, 'Выкуп, %': 1 }, t, 0, 0, 0, 1, 'MAX');
  assert.equal(byFact, 600);   // (100 + 200) / 0.5
  assert.equal(byTariff, 1200); // (100 + 500) / 0.5
});

test('старая таблица без доступа: понятная ошибка со ссылкой', () => {
  const { ctx } = load();
  ctx.SpreadsheetApp.openById = () => ({ getSheetByName: () => { throw new Error('You do not have permission to access the requested document.'); } });
  ctx.Session.getActiveUser = () => ({ getEmail: () => 'user@example.com' });
  assert.throws(() => ctx.sourcesFromOldTable_(), /нет доступа к старой таблице у аккаунта user@example\.com.*1Rdh8/);
});

test('отчёты снимают старый фильтр перед новым', () => {
  ['abcAnalysis_', 'dumpingReport_'].forEach(fn => {
    const body = SRC.slice(SRC.indexOf(`function ${fn}()`)).split('\nfunction ')[0];
    assert.ok(body.indexOf('getFilter().remove()') >= 0 && body.indexOf('getFilter().remove()') < body.indexOf('createFilter'), fn);
  });
});

/* ---------- Закуп: правила, проверка, папка с прайсами ---------- */
const DAY = 864e5;
/** Прогон importCosts_: прайс POSCENTER (код → цена), 1С (product_id → цена), товары; возвращает закуп, проверку и историю */
function runCosts({ price = {}, c1 = {}, rows, sources = SOURCES, settings = {}, drive }) {
  const env = load({ settings: Object.assign({ COST_1C_SHEET_ID: 'ID_1C', COST_1C_SHEET: 'Prices' }, settings) });
  const { ctx } = env;
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '.');
  ctx.Utilities.formatDate = () => today;
  const c1Sheet = new Sheet([['product_id', today]].concat(Object.keys(c1).map(k => [Number(k), c1[k]])));
  ctx.SpreadsheetApp.openById = id => ({ ID_1C: { getSheetByName: () => c1Sheet },
    ID_POS: priceBook(Object.keys(price).map(k => [k, price[k]])), ID_MER: priceBook([]) })[id];
  if (drive) ctx.DriveApp = drive;
  ctx.readTable_ = () => ({ rows: sources });
  ctx.readMain_ = () => ({ rows });
  const cost = {}, out = {};
  ctx.mainPatch_ = (m, name, patch) => { if (name === 'Закуп, ₽') Object.assign(cost, patch); };
  ctx.writeCostCheck_ = (src, recs) => { out.src = src; out.recs = recs; return recs.filter(x => x.why.length).length; };
  ctx.appendCostHistory_ = h => { out.history = h; };
  out.msg = ctx.importCosts_();
  const why = art => (out.recs.find(x => x.r['Артикул'] === art) || { why: [] }).why.join('; ');
  return Object.assign(out, { cost: plain(cost), why, env });
}
const prod = (row, art, pid, extra) => Object.assign({ _row: row, 'Артикул': art, 'Источник закупа': 'Прайс POSCENTER', 'Код в прайсе': art, 'Product ID': pid }, extra);

test('правило закупа: у товара → у прайса → по умолчанию MAX', () => {
  const rows = [prod(3, 'a', 1), prod(4, 'b', 2, { 'Правило закупа': 'Прайс' }), prod(5, 'c', 3, { 'Правило закупа': '1С' }),
    prod(6, 'd', 4, { 'Правило закупа': 'прайс' }), prod(7, 'e', 5, { 'Правило закупа': 'Прайс' })];
  const price = { a: 800, b: 800, c: 800, d: 800 }, c1 = { 1: 1000, 2: 1000, 3: 1000, 4: 1000, 5: 1000 };
  const r = runCosts({ price, c1, rows });
  assert.deepEqual(r.cost, { 3: 1000, 4: 800, 5: 1000, 6: 800, 7: 1000 });
  assert.match(r.why('e'), /правило «Прайс», но в прайсе цены нет — взята 1С/);
  // правило прайса действует на все его товары без своего правила
  const src = SOURCES.map(x => Object.assign({}, x, x['Источник'] === 'Прайс POSCENTER' ? { 'Правило': 'Прайс' } : {}));
  assert.deepEqual(runCosts({ price, c1, rows: [prod(3, 'a', 1), prod(4, 'c', 3, { 'Правило закупа': '1С' })], sources: src }).cost, { 3: 800, 4: 1000 });
});

test('проверка: расхождение прайса и 1С, скачок закупа, история', () => {
  const rows = [prod(3, 'a', 1, { 'Закуп, ₽': 1000 }), prod(4, 'b', 2, { 'Закуп, ₽': 500 }), prod(5, 'c', 3, { 'Закуп, ₽': 1000 })];
  const r = runCosts({ price: { a: 1000, b: 800, c: 1050 }, c1: { 1: 700, 2: 790, 3: 1000 }, rows });
  assert.match(r.why('a'), /прайс и 1С расходятся на 43%/);
  assert.match(r.why('b'), /закуп изменился на 60%/);
  assert.equal(r.why('c'), '', 'мелкие расхождения не шумят');
  assert.deepEqual(plain(r.history.map(h => [h[1], h[2], h[3]])), [['b', 500, 800], ['c', 1000, 1050]]);
});

test('проверка: старый прайс — одна строка про источник, а не про каждый товар', () => {
  const drive = { getFileById: () => ({ getLastUpdated: () => new Date(Date.now() - 40 * DAY) }) };
  const r = runCosts({ price: { a: 100, b: 100 }, c1: {}, rows: [prod(3, 'a', 1), prod(4, 'b', 2)], drive });
  assert.equal(r.src.length, 1);
  assert.match(r.src[0], /«Прайс POSCENTER» \(товаров: 2\): не обновлялся 40 дн\..*срок — 30 дн\./);
  assert.equal(r.why('a'), '');
  // свой срок у прайса важнее общего
  const src = SOURCES.map(x => Object.assign({}, x, { 'Годен, дней': 60 }));
  assert.equal(runCosts({ price: { a: 100 }, rows: [prod(3, 'a', 1)], drive, sources: src }).src.length, 0);
});

test('проверка: нет доступа к прайсу — одна строка с источником, закуп из 1С', () => {
  const src = [{ 'Источник': 'Прайс X', 'ID таблицы': 'NOPE', 'Лист': 'Лист1', 'Колонка кода': 'A', 'Колонка закупа': 'B' }];
  const r = runCosts({ c1: { 1: 500 }, rows: [prod(3, 'a', 1, { 'Источник закупа': 'Прайс X' })], sources: src });
  assert.deepEqual(r.cost, { 3: 500 });
  assert.match(r.src[0], /«Прайс X» \(товаров: 1\): нет доступа/);
  assert.equal(r.why('a'), '');
});

test('папка с прайсами: берётся самый свежий файл поставщика, копии скрипта пропускаются', () => {
  const file = (id, name, days, mime) => ({ getId: () => id, getName: () => name, getMimeType: () => mime || 'application/vnd.google-apps.spreadsheet',
    getLastUpdated: () => new Date(Date.now() - days * DAY) });
  const files = [file('OLD', 'CAS прайс март', 20), file('NEW', 'CAS прайс апрель', 2), file('DORS', 'Dors', 0),
    file('CONV', '[копия для скрипта] CAS', 0), file('PDF', 'CAS.pdf', 0, 'application/pdf')];
  const drive = { getFolderById: () => ({ getName: () => 'Прайсы', getFiles: () => { let i = 0; return { hasNext: () => i < files.length, next: () => files[i++] }; } }) };
  const src = [{ 'Источник': 'Прайс CAS', 'Папка': 'https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz', 'Файл содержит': 'cas',
    'Лист': '', 'Колонка кода': 'A', 'Колонка закупа': 'B', 'Валюта': 'RUB' }];
  const env = load({ settings: { COST_1C_SHEET_ID: '' } });
  const { ctx } = env;
  const opened = [];
  ctx.DriveApp = drive;
  ctx.SpreadsheetApp.openById = id => { opened.push(id); return priceBook([['k1', id === 'NEW' ? 222 : 111]]); };
  ctx.readTable_ = () => ({ rows: src });
  ctx.readMain_ = () => ({ rows: [prod(3, 'k1', 1, { 'Источник закупа': 'Прайс CAS' })] });
  const cost = {};
  ctx.mainPatch_ = (m, name, patch) => { if (name === 'Закуп, ₽') Object.assign(cost, patch); };
  ctx.writeCostCheck_ = () => 0; ctx.appendCostHistory_ = () => {};
  ctx.importCosts_();
  assert.deepEqual(opened, ['NEW']);
  assert.deepEqual(plain(cost), { 3: 222 });
});

test('ссылка вместо ID и правило закупа разных написаний', () => {
  const { ctx } = load();
  assert.equal(ctx.idFrom_('https://docs.google.com/spreadsheets/d/1wOJa0O7Rdxgh_Kyy_Gcvsfq-1xq_X6BVg2ytC4pRx1M/edit#gid=0'), '1wOJa0O7Rdxgh_Kyy_Gcvsfq-1xq_X6BVg2ytC4pRx1M');
  assert.equal(ctx.idFrom_(' ID_POS '), 'ID_POS');
  assert.deepEqual(['max', 'Макс', 'Прайс', '1C', '1с', '', 'ерунда'].map(ctx.costRule_), ['MAX', 'MAX', 'Прайс', '1С', '1С', '', '']);
});

/** Лист с заглушками оформления: неизвестные методы диапазона и листа ничего не делают */
function fakeReportSheet(rows) {
  const sh = new Sheet(rows);
  const chain = target => new Proxy(target, { get: (t, p) => (p in t ? (typeof t[p] === 'function' ? t[p].bind(t) : t[p]) : () => chain(t)) });
  const getRange = sh.getRange.bind(sh);
  sh.getRange = (...a) => { const r = getRange(...a); r.clear = () => { for (let i = 0; i < r.nr; i++) for (let j = 0; j < r.nc; j++) sh.put(r.r + i, r.c + j, ''); return r; }; return chain(r); };
  return chain(Object.assign(sh, { getFilter: () => null }));
}

test('лист проверки: «Принято» прячет строку, пока данные не изменятся; скачок держится до галочки', () => {
  const { ctx } = load();
  const C = vm.runInContext('COST_CHECK_COLS', ctx);
  const row = o => C.map(c => (c in o ? o[c] : ''));
  const sh = fakeReportSheet([C,
    row({ 'Артикул': 'a', 'Закуп стал, ₽': 1000, 'Прайс, ₽': 1000, '1С, ₽': 700, 'Что проверить': 'прайс и 1С расходятся на 43%', 'Принято': true }),
    row({ 'Артикул': 'b', 'Закуп был, ₽': 500, 'Закуп стал, ₽': 800, 'Что проверить': 'закуп изменился на 60%', 'Принято': false })]);
  ctx.reportSheet_ = () => sh;
  ctx.readMain_ = () => ({ rows: [] });
  const rule = new Proxy({}, { get: (t, p) => () => (p === 'build' ? {} : rule) });
  ctx.SpreadsheetApp.newConditionalFormatRule = () => rule;
  const rec = (art, o) => Object.assign({ r: { 'Артикул': art }, rule: 'MAX', old: '', val: 0, p: 0, c: 0, why: [] }, o);
  const n = ctx.writeCostCheck_([], [
    rec('a', { old: 1000, val: 1000, p: 1000, c: 700, why: ['прайс и 1С расходятся на 43%'] }),   // принято, ничего не поменялось
    rec('b', { old: 800, val: 800, p: 800, c: 790, why: [] }),                                      // скачок с прошлого раза
    rec('c', { old: 1000, val: 1000, p: 1000, c: 600, why: ['прайс и 1С расходятся на 67%'] })]);
  assert.equal(n, 2);
  const out = sh.rows.slice(1).filter(r => r[0]).map(r => [r[0], r[C.indexOf('Закуп был, ₽')], r[C.indexOf('Что проверить')], r[C.indexOf('Принято')]]);
  assert.deepEqual(out, [['b', 500, 'закуп изменился на 60%', false], ['c', 1000, 'прайс и 1С расходятся на 67%', false],
    ['a', 1000, 'прайс и 1С расходятся на 43%', true]]);
});

/* ---------- Временно отключённые товары (строки без Product ID) ---------- */
test('отключённые: закуп их не трогает, в «Нет закупа» и на проверку не попадают', () => {
  const rows = [prod(3, 'a', 1), prod(4, 'off', ''), prod(5, 'off2', null, { 'Закуп, ₽': 500 })];
  const r = runCosts({ price: { a: 100 }, c1: { 1: 100 }, rows });
  assert.deepEqual(r.cost, { 3: 100 }, 'строки отключённых не переписываются');
  assert.deepEqual(plain(r.recs.map(x => x.r['Артикул'])), ['a']);
  assert.ok(!r.env.logs.some(l => /Нет закупа/.test(l)), r.env.logs.join('\n'));
  assert.match(r.msg, /отключённых пропущено: 2/);
});

test('отключённые: цена не уходит в Ozon, даже если она осталась в строке', () => {
  const { ctx } = load({ settings: { MAX_PRICE_CHANGE: 0.3 } });
  const res = {};
  const item = (row, art, pid) => ({ _row: row, 'Артикул': art, 'Product ID': pid, 'Цена к отправке, ₽': 100,
    'Зачёркнутая, ₽': 0, 'min_price, ₽': 90, 'Цена на Ozon, ₽': 95 });
  ctx.readMain_ = () => ({ h: ['Артикул', 'Результат'], rows: [item(3, 'on', 11), item(4, 'off', '')] });
  ctx.mainPatch_ = (m, name, patch) => { if (name === 'Результат') Object.assign(res, patch); };
  const msg = ctx.uploadPrices_('all');
  assert.match(res[3], /^ПРОВЕРКА: уйдёт 100 /);
  assert.match(res[4], /^⏸ товар отключён/);
  assert.match(msg, /взято 1, .*отключённых пропущено: 1/);
});

test('отключённые: аудит их не проверяет, статус на листе — «⏸ отключён»', () => {
  const env = load({ settings: { DRY_RUN: true, AUDIT_TOLERANCE: 0.02, DEFAULT_BUYOUT: 1 } });
  const { ctx } = env;
  ctx.readMain_ = () => ({ rows: [{ 'Product ID': '', 'Артикул': 'off', 'Закуп, ₽': 68 }] });
  ctx.ensureAuditSheet_ = () => {};
  ctx.auditPrevStock_ = () => ({});
  let written = null;
  ctx.writeTable_ = (name, rows) => { written = rows; };
  ctx.ozonAll_ = () => [];
  ctx.listActions_ = () => [];
  ctx.runAudit_();
  assert.deepEqual(plain(written || []), [], 'строки «❔ нет в ответе Ozon» для отключённых больше нет');
  const cols = vm.runInContext('MAIN_COLS', ctx);
  const F = ctx.mainFormulas_(cols, 3);
  const pid = '$' + ctx.letter_(cols.indexOf('Product ID') + 1) + '3';
  assert.ok(F['Статус'].startsWith(`=IF(${pid}="","⏸ отключён",`), F['Статус']);
});

/* ---------- Ручной закуп ---------- */
test('закуп вручную: автообновление не перезаписывает, но показывает, если прайс ушёл', () => {
  const rows = [prod(3, 'a', 1, { 'Закуп вручную': true, 'Закуп, ₽': 900 }), prod(4, 'b', 2, { 'Закуп вручную': true, 'Закуп, ₽': 900 }),
    prod(5, 'c', 3)];
  const r = runCosts({ price: { a: 950, b: 1200, c: 500 }, c1: { 1: 900, 2: 900, 3: 500 }, rows });
  assert.deepEqual(r.cost, { 5: 500 }, 'ручной закуп остался как есть');
  assert.equal(r.why('a'), '', 'расхождение 6% — не шумим');
  assert.match(r.why('b'), /закуп задан вручную; в прайсе сейчас 1200 ₽ \(\+33%\)/);
  assert.deepEqual(plain(r.history || []), [], 'ручные строки в историю автообновления не попадают');
});

/** Лист Ozon в песочнице: строка 2 — заголовки, товары с 3-й; выделение — строки sel */
function manualEnv(rows, sel) {
  const env = load({ settings: { DRY_RUN: true } });
  const { ctx } = env;
  const H = ['Артикул', 'Product ID', 'Закуп, ₽', 'Закуп вручную'];
  const sheet = new Sheet([['Товар'], H].concat(rows.map(r => H.map(k => (k in r ? r[k] : '')))));
  sheet.name = 'Ozon';
  ctx.SpreadsheetApp.getActive = () => ({
    getActiveSheet: () => sheet, getSheetByName: () => null, toast() {},
    getActiveRangeList: () => ({ getRanges: () => sel.map(([a, b]) => ({ getRow: () => a, getLastRow: () => b })) })
  });
  ctx.readMain_ = () => ({ sh: sheet, h: H, rows: rows.map((r, i) => Object.assign({ _row: i + 3 }, r)) });
  const patches = {}, history = [];
  ctx.mainPatch_ = (m, name, patch) => { patches[name] = Object.assign(patches[name] || {}, patch); };
  ctx.appendCostHistory_ = h => history.push(...h);
  return Object.assign(env, { sheet, patches, history, H });
}

test('закуп вручную: правка ячейки + кнопка → галочка, история «было → стало»', () => {
  const env = manualEnv([{ 'Артикул': 'a', 'Product ID': 1, 'Закуп, ₽': 900 }, { 'Артикул': 'b', 'Product ID': 2, 'Закуп, ₽': 500 }], [[3, 3]]);
  // пользователь поменял закуп у «a» с 1000 на 900: onEdit запомнил старое значение
  env.ctx.onEdit({ oldValue: '1000', range: { getNumRows: () => 1, getNumColumns: () => 1, getRow: () => 3, getColumn: () => 3, getSheet: () => env.sheet } });
  const msg = env.ctx.manualCostFromSelection_();
  assert.deepEqual(plain(env.patches['Закуп вручную']), { 3: true }, 'отмечен только выделенный товар');
  assert.deepEqual(plain(env.history.map(h => [h[1], h[2], h[3], h[5]])), [['a', 1000, 900, 'вручную']]);
  assert.match(msg, /закуп вручную: 1 — a: 1000 → 900 ₽/);
  assert.equal(env.docProps.getProperty('MANUAL_COST_OLD'), '{}', 'запомненное значение использовано и стёрто');
});

test('закуп вручную: пустая ячейка, отключённый товар и выделение не на листе Ozon', () => {
  const env = manualEnv([{ 'Артикул': 'a', 'Product ID': 1, 'Закуп, ₽': '' }, { 'Артикул': 'off', 'Product ID': '', 'Закуп, ₽': 700 }], [[3, 4]]);
  assert.match(env.ctx.manualCostFromSelection_(), /в «Закуп, ₽» не число: a/);
  assert.equal(env.patches['Закуп вручную'], undefined, 'отключённый товар не трогаем');
  env.sheet.name = 'Бог акций';
  assert.throws(() => env.ctx.manualCostFromSelection_(), /перейдите на лист «Ozon»/);
});

test('вернуть автоматический закуп: снимает галочку и сразу обновляет закуп', () => {
  const env = manualEnv([{ 'Артикул': 'a', 'Product ID': 1, 'Закуп, ₽': 900, 'Закуп вручную': true },
    { 'Артикул': 'b', 'Product ID': 2, 'Закуп, ₽': 500, 'Закуп вручную': false }], [[3, 4]]);
  env.ctx.importCosts_ = () => 'обновлено: 2';
  const msg = env.ctx.manualCostRevert_();
  assert.deepEqual(plain(env.patches['Закуп вручную']), { 3: false });
  assert.match(msg, /автоматический закуп возвращён: 1 \(a\) \| обновлено: 2/);
});

/* ---------- Новые товары из Ozon ---------- */
test('новые товары: отключённая строка снова в продаже — получает Product ID, архив не трогаем', () => {
  const { ctx } = load();
  const rows = [{ _row: 3, 'Артикул': 'X', 'Product ID': '' }, { _row: 4, 'Артикул': 'Y', 'Product ID': '' }, { _row: 5, 'Артикул': 'Z', 'Product ID': 5 }];
  ctx.readMain_ = () => ({ h: [], rows });
  ctx.ozonAll_ = () => [{ offer_id: 'X', product_id: 77, archived: false }, { offer_id: 'Y', product_id: 88, archived: true },
    { offer_id: 'Z', product_id: 5, archived: false }];
  ctx.ozon_ = (p, body) => ({ items: body.product_id.map(id => ({ id, sku: id * 100, name: 'n' + id })) });
  const patches = {}; let regrouped = false;
  ctx.mainPatch_ = (m, name, patch) => { patches[name] = Object.assign(patches[name] || {}, patch); };
  ctx.regroup_ = () => { regrouped = true; };
  ctx.syncTariffs_ = ctx.syncStocks_ = ctx.importCosts_ = () => '';
  const msg = ctx.addMissingProducts_();
  assert.deepEqual(plain(patches['Product ID']), { 3: 77 });
  assert.deepEqual(plain(patches['SKU']), { 3: 7700 });
  assert.equal(regrouped, false, 'новых строк нет — лист не перестраиваем');
  assert.match(msg, /снова в продаже \(вернули Product ID\): 1 — X/);
});

test('«Обновить всё» новые товары в таблицу не загружает — только кнопкой', () => {
  const { ctx } = load({ settings: { AUTO_REMOVE_FROM_ACTIONS: false } });
  const called = [];
  ['updateUsdRate_', 'syncTariffs_', 'syncStocks_', 'importCosts_', 'syncOrders60_', 'syncBuyout_', 'syncFinance_',
    'refreshActions_', 'runAudit_', 'dumpingReport_', 'addMissingProducts_'].forEach(fn => { ctx[fn] = () => { called.push(fn); return 'ok'; }; });
  ctx.syncAll();
  assert.ok(called.includes('importCosts_'), called.join(', '));
  assert.ok(!called.includes('addMissingProducts_'), 'новые товары заводятся только по кнопке');
});
