/**
 * Проверки скрипта apps-script/ozon-cards.gs без Google: сервисы Apps Script подменены заглушками.
 * Запуск: node --test
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SRC = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'ozon-cards.gs'), 'utf8');
const plain = x => JSON.parse(JSON.stringify(x));

/** Лист: хранит значения, остальные методы (форматы, фильтры) — пустые и возвращают себя */
function makeSheet(name) {
  const sh = { name, rows: [] };
  const chain = target => new Proxy(target, { get: (t, k) => k in t ? t[k] : () => chain(t) });
  sh.getLastRow = () => sh.rows.length;
  sh.getLastColumn = () => Math.max(0, ...sh.rows.map(r => r.length));
  sh.getFilter = () => null;
  sh.clear = () => { sh.rows = []; return api; };
  sh.getRange = (r, c, nr, nc) => chain({
    setValues(v) { v.forEach((row, i) => row.forEach((x, j) => { while (sh.rows.length < r + i) sh.rows.push([]); sh.rows[r + i - 1][c + j - 1] = x; })); return chain(this); },
    getValues() { const o = []; for (let i = 0; i < (nr || 1); i++) { const row = []; for (let j = 0; j < (nc || 1); j++) row.push((sh.rows[r + i - 1] || [])[c + j - 1] ?? ''); o.push(row); } return o; }
  });
  const api = chain(sh);
  return api;
}
const rule = () => { const b = new Proxy({}, { get: () => () => b }); return b; };

function load({ fetch } = {}) {
  const sheets = {}, menus = [], props = {};
  const ctx = {
    console: { log() {} },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => ({ OZON_CLIENT_ID: '1', OZON_API_KEY: 'k' })[k] || props[k] || null, setProperties: o => Object.assign(props, o) }) },
    SpreadsheetApp: {
      getActive: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = makeSheet(n)), toast() {} }),
      getUi: () => ({ createMenu: name => { const m = { name, items: [] }; const b = {
        addItem(label, fn) { m.items.push({ label, fn }); return b; }, addSeparator() { return b; },
        addSubMenu(s) { m.items.push({ sub: s._m }); return b; }, addToUi() { menus.push(m); }, _m: m }; return b; } }),
      newConditionalFormatRule: rule, flush() {}
    },
    UrlFetchApp: { fetch: fetch || (() => { throw new Error('сеть не нужна в этом тесте'); }) },
    Utilities: { sleep() {}, formatDate: d => new Date(d).toISOString().slice(0, 10) },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => 'UTC' },
    ScriptApp: { getProjectTriggers: () => [] }
  };
  vm.createContext(ctx);
  vm.runInContext(SRC, ctx);
  ctx.log_ = () => {};
  return { ctx, sheets, menus };
}
/** Ответ Ozon по пути запроса */
const ozonFetch = routes => (url, opts) => {
  const p = url.replace('https://api-seller.ozon.ru', ''), body = opts.payload ? JSON.parse(opts.payload) : {};
  if (!routes[p]) throw new Error('нет заглушки для ' + p);
  return { getResponseCode: () => 200, getContentText: () => JSON.stringify(routes[p](body)) };
};

test('меню: у каждой кнопки есть функция', () => {
  const { ctx, menus } = load();
  ctx.onOpen();
  const flat = list => list.flatMap(i => i.sub ? flat(i.sub.items) : [i]);
  const items = flat(menus[0].items);
  assert.equal(menus[0].name, '🔎 КАРТОЧКИ');
  items.forEach(i => assert.equal(typeof ctx[i.fn], 'function', `нет функции ${i.fn}`));
});

const FT = { minViews: 300, minVisits: 30, minCart: 10, weak: 0.5 };
const fm = (views, visits, cart, cartPdp, orders) => ({ hits_view_search: views, session_view_pdp: visits, session_view: visits,
  hits_tocart: cart, hits_tocart_pdp: cartPdp, ordered_units: orders, revenue: 0, position_category: 0 });
const MED = { ctr: 0.05, cart: 0.10, order: 0.5 };

test('воронка: диагноз по шагу, где товар хуже категории', () => {
  const { ctx } = load();
  const d = (m, stock = 5) => ctx.funnelDiagnosis_(m, stock, MED, FT)[0];
  assert.equal(d(fm(100, 5, 0, 0, 0), 0), 'STOCK');
  assert.equal(d(fm(100, 5, 0, 0, 0)), 'VIEWS');
  assert.equal(d(fm(10000, 100, 20, 20, 10)), 'CTR');        // 1% при 5%
  assert.equal(d(fm(400, 20, 0, 0, 0)), 'FEW');              // переходы в норме, посетителей мало
  assert.equal(d(fm(1000, 100, 2, 2, 1)), 'CART');           // 2% при 10%
  assert.equal(d(fm(1000, 100, 20, 20, 2)), 'ORDER');        // 10% из корзины при 50%
  assert.equal(d(fm(1000, 100, 20, 20, 10)), 'OK');
  assert.match(ctx.funnelDiagnosis_(fm(10000, 100, 20, 20, 10), 5, MED, FT)[1], /1% при 5% в категории/);
});

test('воронка: периоды без пересечения, текущий кончается вчера', () => {
  const { ctx } = load();
  const p = plain(ctx.funnelPeriods_(28, new Date('2026-10-26T12:00:00Z')));
  assert.deepEqual(p, { cur: ['2026-09-28', '2026-10-25'], prev: ['2026-08-31', '2026-09-27'] });
});

test('воронка: метрики Ozon раскладываются по SKU в порядке запроса', () => {
  const calls = [];
  const { ctx } = load({ fetch: ozonFetch({ '/v1/analytics/data': b => { calls.push(b); return { result: { data: [
    { dimensions: [{ id: '111', name: 'x' }], metrics: [500, 40, 50, 8, 6, 3, 9000, 12.4] }] } }; } }) });
  const r = plain(ctx.analyticsBySku_('2026-09-28', '2026-10-25'));
  assert.equal(calls[0].dimension[0], 'sku');
  assert.deepEqual(r['111'], { hits_view_search: 500, session_view_pdp: 40, session_view: 50, hits_tocart: 8,
    hits_tocart_pdp: 6, ordered_units: 3, revenue: 9000, position_category: 12.4 });
});

test('воронка: без Premium — понятная ошибка', () => {
  const { ctx } = load({ fetch: () => ({ getResponseCode: () => 403, getContentText: () => '{"message":"premium required"}' }) });
  assert.throws(() => ctx.analyticsBySku_('2026-09-28', '2026-10-25'), /подписку Premium/);
});

test('запросы: слова запроса, которых нет в названии, с учётом окончаний', () => {
  const { ctx } = load();
  const miss = (q, n) => plain(ctx.missingWords_(q, n));
  assert.deepEqual(miss('весы торговые для магазина', 'MERTECH Торговые весы M-ER 326'), ['магазина']);
  assert.deepEqual(miss('сканер штрихкода беспроводной', 'Сканер штрих кода беспроводной MERTECH CL-2220'), []);
  assert.deepEqual(miss('сканер 2d qr', 'Сканер штрих кода беспроводной MERTECH CL-2220'), ['2d', 'qr']);
  assert.deepEqual(miss('ёмкость', 'Емкости'), []);
});

test('качество карточек: невыполненные условия и характеристики', () => {
  const { ctx } = load();
  const r = plain(ctx.contentRow_({ sku: 1, rating: 62.5, groups: [
    { key: 'media', name: 'Медиа', rating: 50, conditions: [{ description: 'Добавьте видео', fulfilled: false }, { description: 'Фото', fulfilled: true }] },
    { key: 'important_attributes', name: 'Важные атрибуты', rating: 70, improve_at_least: 2,
      improve_attributes: [{ id: 1, name: 'Гарантия' }, { id: 2, name: 'Интерфейс' }], conditions: [] },
    { key: 'text', name: 'Текст', rating: 100, conditions: [] }] }));
  assert.equal(r.weakGroups, 'Медиа 50, Важные атрибуты 70');
  assert.equal(r.todo, 'Добавьте видео');
  assert.equal(r.attrs, 'минимум 2 из: Гарантия, Интерфейс');
});

test('товары: список, тип Ozon, остаток FBS и заказы из аналитики', () => {
  const { ctx, sheets } = load({ fetch: ozonFetch({
    '/v3/product/list': b => ({ result: { items: b.last_id ? [] : [{ product_id: 1, offer_id: 'A-1' }, { product_id: 2, offer_id: 'B-2' }, { product_id: 3, offer_id: 'OLD', archived: true }], last_id: 'x' } }),
    '/v3/product/info/list': () => ({ items: [{ id: 1, name: 'Весы', sku: 111, type_id: 70 }, { id: 2, name: 'Сканер', sources: [{ sku: 222 }], type_id: 80 }] }),
    '/v4/product/info/stocks': b => ({ items: b.cursor ? [] : [{ product_id: 1, stocks: [{ type: 'fbs', present: 5, reserved: 1 }, { type: 'fbo', present: 9 }] }], cursor: 'c' }),
    '/v1/description-category/tree': () => ({ result: [{ category_name: 'Торговое оборудование', children: [{ type_id: 70, type_name: 'Весы торговые' }, { type_id: 80, type_name: 'Сканер штрихкода' }] }] }),
    '/v1/analytics/data': b => { assert.deepEqual(b.metrics, ['ordered_units']); return { result: { data: [{ dimensions: [{ id: '222' }], metrics: [2] }] } }; }
  }) });
  ctx.ensureSettings_();
  ctx.refreshProducts_();
  const rows = sheets['Товары'].rows;
  assert.deepEqual(rows[0].slice(0, 7), ['Артикул', 'SKU', 'Product ID', 'Название', 'Категория', 'Остаток FBS', 'Заказы 60 дн']);
  assert.deepEqual(rows.slice(1).map(r => r.slice(0, 7)), [
    ['B-2', 222, 2, 'Сканер', 'Сканер штрихкода', 0, 2],
    ['A-1', 111, 1, 'Весы', 'Весы торговые', 4, 0]]);
  const p = plain(ctx.readProducts_());
  assert.equal(p[1]['Остаток FBS'], 4);
});

test('товары: сбой заказов не мешает собрать лист', () => {
  const { ctx, sheets } = load({ fetch: ozonFetch({
    '/v3/product/list': b => ({ result: { items: b.last_id ? [] : [{ product_id: 1, offer_id: 'A-1' }], last_id: 'x' } }),
    '/v3/product/info/list': () => ({ items: [{ id: 1, name: 'Весы', sku: 111, type_id: 70 }] }),
    '/v4/product/info/stocks': () => ({ items: [] }),
    '/v1/description-category/tree': () => ({ result: [] })
  }) });
  ctx.ensureSettings_();
  assert.match(ctx.refreshProducts_(), /не получены/);
  assert.deepEqual(sheets['Товары'].rows[1].slice(0, 7), ['A-1', 111, 1, 'Весы', '', 0, '']);
});
