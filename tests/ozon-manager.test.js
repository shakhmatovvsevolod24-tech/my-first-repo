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
test('onOpen: четыре отдельных меню, у каждой кнопки есть функция', () => {
  const { ctx, menus } = load();
  ctx.onOpen();
  assert.deepEqual(menus.map(m => m.name), ['💰 ЦЕНЫ', '🏷 АКЦИИ', '📦 ОСТАТКИ', '⚙ НАСТРОЙКИ']);
  const items = menus.flatMap(m => m.items.filter(i => i.fn));
  items.forEach(i => assert.equal(typeof ctx[i.fn], 'function', `нет функции ${i.fn} для «${i.label}»`));
  const fns = items.map(i => i.fn);
  ['applySelectedForce', 'removeAllFromActions', 'setupSheetUi', 'analyzeShippingSpeed', 'compareCosts']
    .forEach(fn => assert.ok(!fns.includes(fn), `кнопка ${fn} должна быть убрана`));
  const labels = items.map(i => i.label);
  assert.ok(labels.includes('Применить отметки с маржой от 10%'));
  assert.ok(labels.includes('Выгрузить цены по марже 12%'));
  assert.ok(labels.includes('Включить автообновление (каждые 5 мин)'));
  assert.equal(new Set(fns).size, fns.length, 'одна функция висит на двух кнопках');
});

test('onOpen: подписи берут проценты и интервал из «Настроек»', () => {
  const { ctx, menus } = load({ settings: { MIN_MARGIN: 0.11, PROMO_MARGIN: 0.15, STOCKS_REFRESH_MIN: 7 } });
  ctx.onOpen();
  const labels = menus.flatMap(m => m.items.filter(i => i.fn).map(i => i.label));
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
  const item = (row, art, price) => ({ _row: row, 'Артикул': art, 'Отправить': true, 'Цена к отправке, ₽': price,
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
