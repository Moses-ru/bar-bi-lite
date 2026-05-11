const tg = window.Telegram?.WebApp;
if (tg) {
  tg.ready();
  tg.expand();
  tg.setHeaderColor('#07111f');
  tg.setBackgroundColor('#07111f');
}

const CONFIG = { periodDays: 30, riskDays: 3, overstockDays: 30 };
const DEFAULT_SUPPLIER_SETTINGS = { deliveryDays: 2, orderDays: 7, reserveDays: 2 };
const state = {
  stock: [],
  suppliers: {},
  supplierProducts: {},
  manualSuppliers: {},
  supplierSettings: {},
  orderEdits: {},
  history: [],
  charts: {},
  lastFile: '',
  abcFilter: 'ALL'
};

const $ = (id) => document.getElementById(id);
const money = (v) => `${Math.round(Number(v) || 0).toLocaleString('ru-RU')} ₽`;
const num = (v, d = 3) => Number(v || 0).toLocaleString('ru-RU', { maximumFractionDigits: d });
const normalize = (s) => String(s ?? '').toLowerCase().replace(/ё/g, 'е').replace(/[.,;:()\[\]"'«»]/g, ' ').replace(/\s+/g, ' ').trim();
const productKey = (product) => normalize(product);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
function toNumber(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
function toast(text) { const t = $('toast'); t.textContent = text; t.classList.add('show'); setTimeout(() => t.classList.remove('show'), 2600); }
function today() { return new Date().toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric' }); }

function categoryOf(name) {
  const s = normalize(name);
  const tests = [
    ['Фрукты и овощи', ['апельсин', 'лимон', 'лайм', 'банан', 'гранат', 'мята', 'тархун', 'шпинат', 'огур', 'морковь', 'ягод', 'клубник', 'брусник', 'клюкв', 'зелень', 'базилик', 'имбирь', 'яблок', 'груша', 'малина', 'смородина', 'арбуз', 'дыня', 'перец', 'овощ', 'фрукт', 'лист']],
    ['Вино / игристое', ['вино', 'просекко', 'шампан', 'игрист', 'каберне', 'шардоне', 'киндзмараули', 'пино', 'рислинг', 'совиньон', 'мерло', 'кьянти']],
    ['Крепкий алкоголь', ['водка', 'виски', 'джин', 'ром', 'текила', 'коньяк', 'бренди', 'ликер', 'настойка', 'бурбон', 'абсент', 'самбука', 'егермейстер', 'мартини', 'апероль', 'кампари']],
    ['Пиво / сидр', ['пиво', 'сидр', 'эль', 'лагер', 'кроненбург', 'хофброй', 'бакалар']],
    ['Сиропы / пюре', ['сироп', 'пюре', 'pinch', 'drop', 'барбарис', 'маракуй', 'манго']],
    ['Соки / напитки', ['сок', 'кола', 'coca', 'pepsi', 'вода', 'боржоми', 'морс', 'тоник', 'лимонад', 'энергетик', 'rich', 'рич']],
    ['Молочка / сливки', ['молоко', 'сливки', 'сыр', 'мороженое']],
    ['Расходники', ['салфет', 'трубоч', 'стакан', 'крышк', 'перчат', 'мешок', 'чековая', 'уголь', 'фольга']]
  ];
  for (const [cat, words] of tests) if (words.some(w => s.includes(w))) return cat;
  return 'Другое';
}

function isValidSupplierName(value) {
  const raw = String(value ?? '').trim();
  const n = normalize(raw);
  if (!raw || raw.length < 3) return false;
  if (/^[\d\s.,\-/\\]+$/.test(raw)) return false;
  if (/^\d/.test(raw) && !/[a-zа-яё]{3,}/i.test(raw)) return false;
  const banned = ['артикул', 'наименование', 'тип', 'ед изм', 'единица', 'цена', 'товар', 'прайс', 'поставщик', 'товар поставщика', 'товар в системе', 'код', 'сумма', 'ндс'];
  if (banned.includes(n)) return false;
  if (n.includes('товар поставщика') || n.includes('товар в системе')) return false;
  return /[a-zа-яё]{3,}/i.test(raw);
}
function isProductHeader(value) { const n = normalize(value); return n === 'наименование' || n === 'название' || n === 'номенклатура'; }
function isSupplierHeader(value) { const n = normalize(value); return n === 'поставщик' || n === 'контрагент'; }
function isArticleHeader(value) { const n = normalize(value); return n === 'артикул' || n === 'код' || n === 'штрихкод'; }

function cleanupSuppliers() {
  for (const supplier of Object.keys(state.suppliers || {})) {
    if (supplier !== 'Прочее' && !isValidSupplierName(supplier)) {
      delete state.suppliers[supplier];
      delete state.supplierSettings[supplier];
    }
  }
  for (const [product, supplier] of Object.entries(state.supplierProducts || {})) if (!isValidSupplierName(supplier)) delete state.supplierProducts[product];
  for (const [product, supplier] of Object.entries(state.manualSuppliers || {})) if (supplier !== 'Прочее' && !isValidSupplierName(supplier)) delete state.manualSuppliers[product];
}
function ensureSupplierSettings(supplier) { if (!state.supplierSettings[supplier]) state.supplierSettings[supplier] = { ...DEFAULT_SUPPLIER_SETTINGS }; }
function bestSupplierMatch(product) {
  const key = normalize(product);
  if (!key) return '';
  if (state.supplierProducts[key]) return state.supplierProducts[key];
  let best = { score: 0, supplier: '' };
  for (const [p, sup] of Object.entries(state.supplierProducts)) {
    if (!p || p.length < 3) continue;
    let score = 0;
    if (key === p) score = 100;
    else if (key.includes(p) || p.includes(key)) score = Math.min(key.length, p.length) / Math.max(key.length, p.length) * 92;
    else {
      const a = new Set(key.split(' ').filter(w => w.length > 2));
      const b = new Set(p.split(' ').filter(w => w.length > 2));
      const inter = [...a].filter(x => b.has(x)).length;
      const union = new Set([...a, ...b]).size || 1;
      score = inter / union * 85;
    }
    if (score > best.score) best = { score, supplier: sup };
  }
  return best.score >= 54 ? best.supplier : '';
}
function supplierFor(product) { return state.manualSuppliers[productKey(product)] || bestSupplierMatch(product) || 'Прочее'; }
function getSupplierSettings(supplier) { return { ...DEFAULT_SUPPLIER_SETTINGS, ...(state.supplierSettings[supplier] || {}) }; }
function targetDaysForSupplier(supplier) { const s = getSupplierSettings(supplier); return Number(s.deliveryDays || 0) + Number(s.orderDays || 0) + Number(s.reserveDays || 0); }
function recalcOrders() {
  state.stock.forEach(r => {
    r.supplier = supplierFor(r.product);
    const targetDays = targetDaysForSupplier(r.supplier || 'Прочее');
    r.targetDays = targetDays;
    r.recommendedOrder = Math.max(0, r.dailySales * targetDays - r.stockEnd);
    r.deliveryDays = getSupplierSettings(r.supplier || 'Прочее').deliveryDays;
    const manualQty = state.orderEdits[productKey(r.product)];
    r.orderQty = manualQty !== undefined && manualQty !== null && manualQty !== '' ? Math.max(0, Number(manualQty) || 0) : r.recommendedOrder;
  });
}

async function readWorkbook(file) { const buf = await file.arrayBuffer(); return XLSX.read(buf, { type: 'array', cellDates: true }); }
function parseOSVWorkbook(wb) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  let start = 0;
  for (let i = 0; i < rows.length; i++) {
    const joined = rows[i].map(x => normalize(x)).join(' ');
    if (joined.includes('наименование') && joined.includes('ед изм')) { start = i + 1; break; }
  }
  const data = [];
  for (const r of rows.slice(start)) {
    const product = String(r[0] ?? '').trim();
    const unit = String(r[1] ?? '').trim();
    if (!product || normalize(product).includes('итого') || normalize(product).includes('товар')) continue;
    const stockEnd = toNumber(r[19]);
    const stockSum = toNumber(r[20]);
    const salesQty = Math.abs(toNumber(r[6]));
    const salesCost = Math.abs(toNumber(r[7]));
    const writeoffQty = Math.abs(toNumber(r[11]));
    const writeoffCost = Math.abs(toNumber(r[12]));
    const shortage = Math.abs(toNumber(r[17]));
    if (!unit && !stockEnd && !stockSum && !salesQty && !salesCost) continue;
    const dailySales = salesQty / CONFIG.periodDays;
    const daysLeft = dailySales > 0 ? stockEnd / dailySales : 9999;
    const category = categoryOf(product);
    const supplier = supplierFor(product);
    const overstockMoney = (daysLeft >= CONFIG.overstockDays && salesQty > 0) ? Math.max(0, stockSum) : 0;
    data.push({ product, unit, category, supplier, stockEnd, stockSum, salesQty, salesCost, dailySales, daysLeft, recommendedOrder: 0, orderQty: 0, writeoffQty, writeoffCost, shortage, overstockMoney });
  }
  return data;
}
function parsePriceWorkbook(wb, fileName) {
  let found = 0;
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
    let headerIdx = -1, supplierCol = -1, productCols = [], articleCols = [];
    for (let i = 0; i < Math.min(rows.length, 100); i++) {
      const row = rows[i] || [];
      const sCol = row.findIndex(cell => isSupplierHeader(cell));
      const pCols = [], aCols = [];
      row.forEach((cell, idx) => { if (isProductHeader(cell)) pCols.push(idx); if (isArticleHeader(cell)) aCols.push(idx); });
      if (sCol >= 0 && pCols.length) { headerIdx = i; supplierCol = sCol; productCols = pCols; articleCols = aCols; break; }
    }
    if (headerIdx < 0 || supplierCol < 0 || !productCols.length) { console.warn('Прайс пропущен: нет Поставщик/Наименование', fileName, sheetName); continue; }
    productCols = productCols.filter(c => c !== supplierCol && !articleCols.includes(c));
    for (const r of rows.slice(headerIdx + 1)) {
      const supplier = String(r[supplierCol] ?? '').trim();
      if (!isValidSupplierName(supplier)) continue;
      state.suppliers[supplier] = true;
      ensureSupplierSettings(supplier);
      for (const pcol of productCols) {
        const product = String(r[pcol] ?? '').trim();
        const n = normalize(product);
        if (!product || product.length < 3) continue;
        if (/^[\d\s.,\-/\\]+$/.test(product)) continue;
        if (n.includes('итого') || n.includes('наименование') || n.includes('артикул') || n.includes('товар поставщика') || n.includes('товар в системе')) continue;
        state.supplierProducts[n] = supplier;
        found++;
      }
    }
  }
  cleanupSuppliers();
  saveLocalSuppliers();
  return found;
}

function saveLocalSuppliers() {
  localStorage.setItem('barbi_suppliers', JSON.stringify(state.suppliers));
  localStorage.setItem('barbi_supplier_products', JSON.stringify(state.supplierProducts));
  localStorage.setItem('barbi_manual_suppliers', JSON.stringify(state.manualSuppliers));
  localStorage.setItem('barbi_supplier_settings', JSON.stringify(state.supplierSettings));
  localStorage.setItem('barbi_order_edits', JSON.stringify(state.orderEdits));
}
function loadLocalSuppliers() {
  try { state.suppliers = JSON.parse(localStorage.getItem('barbi_suppliers') || '{}'); } catch { state.suppliers = {}; }
  try { state.supplierProducts = JSON.parse(localStorage.getItem('barbi_supplier_products') || '{}'); } catch { state.supplierProducts = {}; }
  try { state.manualSuppliers = JSON.parse(localStorage.getItem('barbi_manual_suppliers') || '{}'); } catch { state.manualSuppliers = {}; }
  try { state.supplierSettings = JSON.parse(localStorage.getItem('barbi_supplier_settings') || '{}'); } catch { state.supplierSettings = {}; }
  try { state.orderEdits = JSON.parse(localStorage.getItem('barbi_order_edits') || '{}'); } catch { state.orderEdits = {}; }
  cleanupSuppliers();
  state.suppliers['Прочее'] = true;
  ensureSupplierSettings('Прочее');
  Object.keys(state.suppliers).forEach(ensureSupplierSettings);
  saveLocalSuppliers();
}
function saveLocalStock() {
  try { localStorage.setItem('barbi_last_stock', JSON.stringify({ lastFile: state.lastFile, stock: state.stock, savedAt: new Date().toISOString() })); }
  catch { toast('ОСВ слишком большая для localStorage. После обновления придется загрузить заново.'); }
}
function loadLocalStock() {
  try {
    const saved = JSON.parse(localStorage.getItem('barbi_last_stock') || '{}');
    if (Array.isArray(saved.stock) && saved.stock.length) { state.stock = saved.stock; state.lastFile = saved.lastFile || 'из localStorage'; return true; }
  } catch { }
  return false;
}
function loadHistory() { try { state.history = JSON.parse(localStorage.getItem('barbi_stock_history') || '[]'); } catch { state.history = []; } }
function saveHistorySnapshot() {
  if (!state.stock.length) return;
  loadHistory();
  const summary = calcSummary();
  const snapshot = { id: Date.now(), savedAt: new Date().toISOString(), file: state.lastFile, summary, stock: state.stock };
  state.history = [snapshot, ...state.history.filter(x => x.file !== state.lastFile)].slice(0, 10);
  try { localStorage.setItem('barbi_stock_history', JSON.stringify(state.history)); }
  catch { state.history = state.history.slice(0, 3); localStorage.setItem('barbi_stock_history', JSON.stringify(state.history)); }
}
function restoreHistory(id) {
  loadHistory();
  const item = state.history.find(x => String(x.id) === String(id));
  if (!item) return;
  state.stock = item.stock || [];
  state.lastFile = item.file || 'из истории';
  recalcOrders();
  saveLocalStock();
  renderAll();
  switchTab('dashboard');
  toast('ОСВ восстановлена из истории');
}

function calcSummary() {
  const s = state.stock;
  return {
    sku: s.length,
    stockSum: s.reduce((a, r) => a + r.stockSum, 0),
    stops: s.filter(r => r.stockEnd <= 0).length,
    risk: s.filter(r => r.stockEnd > 0 && r.daysLeft <= CONFIG.riskDays).length,
    over: s.filter(r => r.daysLeft >= CONFIG.overstockDays && r.salesQty > 0).length,
    orderQty: s.reduce((a, r) => a + (r.orderQty ?? r.recommendedOrder), 0),
    frozen: s.reduce((a, r) => a + r.overstockMoney, 0),
    salesCost: s.reduce((a, r) => a + r.salesCost, 0),
    writeoffCost: s.reduce((a, r) => a + (r.writeoffCost || 0), 0),
    noSupplier: s.filter(r => r.supplier === 'Прочее').length
  };
}
function byCategory() { const m = {}; state.stock.forEach(r => { if (!m[r.category]) m[r.category] = { category: r.category, stockSum: 0, salesQty: 0, sku: 0 }; m[r.category].stockSum += r.stockSum; m[r.category].salesQty += r.salesQty; m[r.category].sku++; }); return Object.values(m).sort((a, b) => b.stockSum - a.stockSum); }
function topBy(key, n = 10) { return [...state.stock].sort((a, b) => (b[key] || 0) - (a[key] || 0)).slice(0, n); }
function orderRows(supplier = 'Все поставщики') { let rows = state.stock.filter(r => (r.orderQty ?? r.recommendedOrder) > 0); if (supplier && supplier !== 'Все поставщики') rows = rows.filter(r => r.supplier === supplier); return rows.sort((a, b) => (b.orderQty ?? b.recommendedOrder) - (a.orderQty ?? a.recommendedOrder)); }
function abcRows() { const rows = [...state.stock].filter(r => r.salesCost > 0 || r.salesQty > 0).sort((a, b) => (b.salesCost || b.salesQty) - (a.salesCost || a.salesQty)); const total = rows.reduce((a, r) => a + (r.salesCost || r.salesQty), 0) || 1; let acc = 0; return rows.map(r => { const v = r.salesCost || r.salesQty; acc += v; const cum = acc / total; return { ...r, abc: cum <= .8 ? 'A' : cum <= .95 ? 'B' : 'C', share: v / total * 100, cumulative: cum * 100 }; }); }

function renderAll() {
  recalcOrders();
  const s = calcSummary();
  $('kpiSku').textContent = s.sku; $('kpiStock').textContent = money(s.stockSum); $('kpiStops').textContent = s.stops; $('kpiRisk').textContent = s.risk; $('kpiOver').textContent = s.over; $('kpiOrder').textContent = num(s.orderQty, 0);
  $('fileStatus').innerHTML = state.stock.length ? `<b>ОСВ: ${esc(state.lastFile)}</b><span>Поставщиков в базе: ${Object.keys(state.suppliers).length}. Последние данные, ручные привязки, редактор заявки и история хранятся в localStorage.</span>` : `<b>Данные не загружены</b><span>Сначала загрузи ОСВ. Прайсы нужны для заявок по поставщикам.</span>`;
  renderInsights(s); renderSupplierSelect(); renderOrders(); renderLists(); renderSupplierSettings(); renderAnalytics(s); renderHistory(); renderCharts();
}
function renderInsights(s) {
  const stopTop = state.stock.filter(r => r.stockEnd <= 0).sort((a, b) => b.salesQty - a.salesQty).map(r => r.product).slice(0, 4).join(', ');
  const html = [];
  if (s.stops) html.push(`<div class="insight critical"><b>Критично · Стопы</b><p>В стопе ${s.stops} позиций. Проверь: ${esc(stopTop || 'позиции с нулевым остатком')}.</p></div>`);
  if (s.risk) html.push(`<div class="insight warn"><b>Риск стопа</b><p>${s.risk} позиций закончатся примерно за ${CONFIG.riskDays} дня или быстрее.</p></div>`);
  if (s.frozen) html.push(`<div class="insight warn"><b>Оверсток</b><p>В лишнем запасе заморожено около ${money(s.frozen)}.</p></div>`);
  if (s.noSupplier) html.push(`<div class="insight"><b>Поставщики</b><p>${s.noSupplier} товаров не найдены в прайсах и попали в «Прочее». Их можно вручную привязать в заявке.</p></div>`);
  if (s.salesCost) html.push(`<div class="insight good"><b>Себестоимость продаж</b><p>По ОСВ общая себестоимость продаж: ${money(s.salesCost)}.</p></div>`);
  if (!html.length) html.push(`<div class="insight good"><b>Все спокойно</b><p>Критичных проблем по текущей ОСВ не найдено.</p></div>`);
  $('insights').innerHTML = html.join('');
}
function supplierOptions(selected) {
  cleanupSuppliers();
  const suppliers = ['Все поставщики', ...new Set(state.stock.map(r => r.supplier).filter(s => s === 'Прочее' || isValidSupplierName(s))), ...Object.keys(state.suppliers).filter(s => s === 'Прочее' || isValidSupplierName(s))].filter(Boolean);
  const unique = [...new Set(suppliers)].sort((a, b) => a.localeCompare(b, 'ru'));
  return unique.map(s => `<option value="${esc(s)}" ${s === selected ? 'selected' : ''}>${esc(s)}</option>`).join('');
}
function renderSupplierSelect() { const selected = $('supplierSelect').value; $('supplierSelect').innerHTML = supplierOptions(selected); }
function renderOrders() {
  const supplier = $('supplierSelect').value || 'Все поставщики';
  const rows = orderRows(supplier).slice(0, 160);
  const st = supplier !== 'Все поставщики' ? getSupplierSettings(supplier) : null;
  $('supplierHint').textContent = st ? `Доставка: ${st.deliveryDays} дн. · Закупаемся на: ${st.orderDays} дн. · Страховой запас: ${st.reserveDays} дн. · Итого запас: ${targetDaysForSupplier(supplier)} дн.` : 'Выбери поставщика, чтобы увидеть заявку и настройки запаса.';
  $('orderList').innerHTML = rows.length ? rows.map(r => {
    const key = productKey(r.product);
    const qty = r.orderQty ?? r.recommendedOrder;
    return `<div class="order-row editable" data-key="${esc(key)}"><b>${esc(r.product)}<small>${esc(r.category)}</small><select class="row-supplier">${supplierOptions(r.supplier)}</select></b><div class="num">${num(r.stockEnd)} ${esc(r.unit)}</div><div class="num"><input class="order-qty" inputmode="decimal" type="number" min="0" step="0.001" value="${Number(qty || 0).toFixed(3).replace(/\.?0+$/, '')}"><small>${esc(r.unit)}</small></div></div>`;
  }).join('') : '<div class="row"><b>Нет позиций к заказу</b><small>По выбранному поставщику заявка пустая.</small></div>';
  document.querySelectorAll('.row-supplier').forEach(sel => sel.addEventListener('change', e => {
    const key = e.target.closest('.order-row').dataset.key;
    state.manualSuppliers[key] = e.target.value;
    state.suppliers[e.target.value] = true;
    ensureSupplierSettings(e.target.value);
    saveLocalSuppliers(); recalcOrders(); saveLocalStock(); renderAll(); toast('Поставщик товара сохранен');
  }));
  document.querySelectorAll('.order-qty').forEach(inp => inp.addEventListener('change', e => {
    const key = e.target.closest('.order-row').dataset.key;
    state.orderEdits[key] = Number(e.target.value || 0);
    saveLocalSuppliers(); recalcOrders(); saveLocalStock(); renderAll(); toast('Количество в заявке сохранено');
  }));
}
function renderLists() {
  $('topStockList').innerHTML = topBy('stockSum', 10).map((r, i) => `<div class="row"><div><b>${i + 1}. ${esc(r.product)}</b><small>${esc(r.category)} · ${esc(r.supplier)}</small></div><div class="value">${money(r.stockSum)}</div></div>`).join('');
  $('topSalesList').innerHTML = topBy('salesQty', 10).map((r, i) => `<div class="row"><div><b>${i + 1}. ${esc(r.product)}</b><small>${esc(r.category)} · остаток: ${num(r.stockEnd)} ${esc(r.unit)}</small></div><div class="value">${num(r.salesQty)} ${esc(r.unit)}</div></div>`).join('');
  $('stopList').innerHTML = state.stock.filter(r => r.stockEnd <= 0).sort((a, b) => b.salesQty - a.salesQty).slice(0, 100).map(r => `<div class="row danger"><div><b>${esc(r.product)}</b><small>${esc(r.category)} · ${esc(r.supplier)}</small></div><div class="value">${num(r.stockEnd)} ${esc(r.unit)}</div></div>`).join('') || '<div class="row"><b>Стопов нет</b></div>';
  $('overList').innerHTML = state.stock.filter(r => r.daysLeft >= CONFIG.overstockDays && r.salesQty > 0).sort((a, b) => b.stockSum - a.stockSum).slice(0, 100).map(r => `<div class="row warn"><div><b>${esc(r.product)}</b><small>${esc(r.category)} · запас: ${num(r.daysLeft, 1)} дней</small></div><div class="value">${money(r.stockSum)}</div></div>`).join('') || '<div class="row"><b>Оверстоков нет</b></div>';
  const abc = abcRows().filter(r => state.abcFilter === 'ALL' || r.abc === state.abcFilter);
  $('abcList').innerHTML = abc.slice(0, 120).map(r => `<div class="row"><div><b>${esc(r.product)}</b><small>${esc(r.category)} · класс ${r.abc} · накопительно ${num(r.cumulative, 1)}%</small></div><div class="value">${num(r.share, 1)}%</div></div>`).join('') || '<div class="row"><b>Нет данных для ABC</b></div>';
}
function renderSupplierSettings() {
  cleanupSuppliers();
  const suppliers = [...new Set([...Object.keys(state.suppliers), ...state.stock.map(r => r.supplier)])].filter(s => s && (s === 'Прочее' || isValidSupplierName(s))).sort((a, b) => a.localeCompare(b, 'ru'));
  $('supplierSettingsList').innerHTML = suppliers.map(s => { const st = getSupplierSettings(s); return `<div class="supplier-setting" data-supplier="${esc(s)}"><b>${esc(s)}</b><div class="settings-fields"><label>Доставка<input type="number" min="0" value="${st.deliveryDays}" data-field="deliveryDays"></label><label>Закуп, дней<input type="number" min="1" value="${st.orderDays}" data-field="orderDays"></label><label>Запас, дней<input type="number" min="0" value="${st.reserveDays}" data-field="reserveDays"></label></div></div>`; }).join('') || '<div class="row"><b>Поставщиков пока нет</b></div>';
  document.querySelectorAll('.supplier-setting input').forEach(inp => inp.addEventListener('change', e => { const box = e.target.closest('.supplier-setting'); const supplier = box.dataset.supplier; ensureSupplierSettings(supplier); state.supplierSettings[supplier][e.target.dataset.field] = Number(e.target.value || 0); saveLocalSuppliers(); recalcOrders(); saveLocalStock(); renderAll(); toast('Настройки поставщика сохранены'); }));
}
function renderAnalytics(s) { $('metricSalesCost').textContent = money(s.salesCost); $('metricWriteoff').textContent = money(s.writeoffCost); $('metricFrozen').textContent = money(s.frozen); $('metricNoSupplier').textContent = s.noSupplier; }
function renderHistory() {
  loadHistory();
  const box = $('historyList');
  if (!box) return;
  box.innerHTML = state.history.length ? state.history.map(item => {
    const sm = item.summary || {}; const dt = new Date(item.savedAt).toLocaleString('ru-RU');
    return `<div class="row"><div><b>${esc(item.file || 'ОСВ')}</b><small>${dt} · SKU: ${sm.sku || 0} · Остатки: ${money(sm.stockSum || 0)} · Стопы: ${sm.stops || 0}</small></div><button class="small-action" data-history="${item.id}">Открыть</button></div>`;
  }).join('') : '<div class="row"><b>Истории пока нет</b><small>Загрузи ОСВ — последние 10 обработок сохранятся в браузере.</small></div>';
  document.querySelectorAll('[data-history]').forEach(btn => btn.addEventListener('click', () => restoreHistory(btn.dataset.history)));
}

const baseChart = { chart: { foreColor: '#dbe4f3', toolbar: { show: false }, background: 'transparent' }, theme: { mode: 'dark' }, grid: { borderColor: 'rgba(255,255,255,.08)' }, tooltip: { theme: 'dark' }, legend: { labels: { colors: '#dbe4f3' } } };
function mountChart(id, options) { if (state.charts[id]) state.charts[id].destroy(); state.charts[id] = new ApexCharts($(id), options); state.charts[id].render(); }
function renderCharts() {
  if (!state.stock.length) return;
  const cats = byCategory().slice(0, 8);
  mountChart('categoryChart', { ...baseChart, chart: { ...baseChart.chart, type: 'donut', height: 300 }, series: cats.map(x => Math.round(x.stockSum)), labels: cats.map(x => x.category), dataLabels: { style: { fontSize: '13px', fontWeight: 800 } }, plotOptions: { pie: { donut: { size: '58%', labels: { show: true, total: { show: true, label: 'Всего', formatter: () => money(cats.reduce((a, c) => a + c.stockSum, 0)) } } } } }, stroke: { width: 0 }, colors: ['#8b5cf6', '#38bdf8', '#34d399', '#f59e0b', '#fb7185', '#a78bfa', '#22c55e', '#eab308'] });
  const topSales = topBy('salesQty', 10).reverse();
  mountChart('salesChart', { ...baseChart, chart: { ...baseChart.chart, type: 'bar', height: 420 }, plotOptions: { bar: { horizontal: true, borderRadius: 6 } }, series: [{ name: 'Расход', data: topSales.map(r => Number(r.salesQty.toFixed(3))) }], xaxis: { categories: topSales.map(r => r.product), labels: { style: { fontSize: '12px' } } }, yaxis: { labels: { style: { fontSize: '13px' } } }, colors: ['#8b5cf6'] });
  const abc = abcRows(); const cnt = { A: 0, B: 0, C: 0 }; abc.forEach(r => cnt[r.abc]++);
  mountChart('abcChart', { ...baseChart, chart: { ...baseChart.chart, type: 'bar', height: 260 }, series: [{ name: 'Позиций', data: [cnt.A, cnt.B, cnt.C] }], xaxis: { categories: ['A — ключевые', 'B — средние', 'C — хвост'] }, plotOptions: { bar: { borderRadius: 8, columnWidth: '45%' } }, colors: ['#34d399'] });
}

function orderText() { const supplier = $('supplierSelect').value || 'Все поставщики'; const rows = orderRows(supplier); const lines = [`ЗАЯВКА · ${supplier}`, '', 'Наименование | Количество']; rows.forEach(r => lines.push(`${r.product} | ${num(r.orderQty ?? r.recommendedOrder)} ${r.unit}`)); return lines.join('\n'); }
function downloadCSV() { const supplier = $('supplierSelect').value || 'Все поставщики'; const rows = orderRows(supplier); const header = ['Наименование', 'Количество']; const csv = [header, ...rows.map(r => [r.product, `${num(r.orderQty ?? r.recommendedOrder)} ${r.unit}`])].map(row => row.map(x => `"${String(x).replace(/"/g, '""')}"`).join(';')).join('\n'); const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `zayavka_${supplier}.csv`; a.click(); URL.revokeObjectURL(url); }

async function handleOSV(file) { const wb = await readWorkbook(file); state.stock = parseOSVWorkbook(wb); state.lastFile = file.name; recalcOrders(); saveLocalStock(); saveHistorySnapshot(); Object.keys(state.charts).forEach(k => { try { state.charts[k].destroy(); } catch { } }); state.charts = {}; renderAll(); toast(`ОСВ загружена: ${state.stock.length} товаров`); }
async function handlePrices(files) { let total = 0; for (const file of files) { const wb = await readWorkbook(file); total += parsePriceWorkbook(wb, file.name); } if (state.stock.length) { state.stock.forEach(r => r.supplier = supplierFor(r.product)); recalcOrders(); saveLocalStock(); } renderAll(); toast(`Прайсы загружены: ${total} связей. Поставщики обновлены.`); }

function switchTab(tab) { document.querySelectorAll('[data-tab]').forEach(b => b.classList.toggle('active', b.dataset.tab === tab)); document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === tab)); closeDrawer(); setTimeout(() => { Object.values(state.charts).forEach(c => { try { c.windowResizeHandler(); } catch { } }); }, 50); }
function openDrawer() { $('drawer').classList.add('open'); $('drawerBackdrop').classList.add('open'); }
function closeDrawer() { $('drawer').classList.remove('open'); $('drawerBackdrop').classList.remove('open'); }
function resetSuppliersOnly() {
  if (confirm('Очистить только поставщиков, ручные привязки и редактор заявки? ОСВ останется.')) {
    ['barbi_suppliers', 'barbi_supplier_products', 'barbi_manual_suppliers', 'barbi_supplier_settings', 'barbi_order_edits'].forEach(k => localStorage.removeItem(k));
    state.suppliers = {}; state.supplierProducts = {}; state.manualSuppliers = {}; state.supplierSettings = {}; state.orderEdits = {};
    loadLocalSuppliers();
    if (state.stock.length) { state.stock.forEach(r => { r.supplier = 'Прочее'; }); recalcOrders(); saveLocalStock(); }
    renderAll(); toast('Поставщики очищены. Теперь загрузи прайс заново.');
  }
}
function resetAll() {
  if (confirm('Очистить поставщиков, настройки, историю и текущую ОСВ?')) {
    ['barbi_suppliers', 'barbi_supplier_products', 'barbi_manual_suppliers', 'barbi_supplier_settings', 'barbi_order_edits', 'barbi_last_stock', 'barbi_stock_history'].forEach(k => localStorage.removeItem(k));
    state.stock = []; state.suppliers = {}; state.supplierProducts = {}; state.manualSuppliers = {}; state.supplierSettings = {}; state.orderEdits = {}; state.history = []; state.lastFile = '';
    loadLocalSuppliers(); Object.keys(state.charts).forEach(k => { try { state.charts[k].destroy(); } catch { } }); state.charts = {}; renderAll(); toast('Данные очищены');
  }
}

document.addEventListener('DOMContentLoaded', () => {
  $('datePill').textContent = today();
  loadLocalSuppliers(); loadHistory();
  const restored = loadLocalStock();
  if (restored) { recalcOrders(); }
  renderAll();
  if (restored) toast('Последняя ОСВ восстановлена из localStorage');
  $('osvInput').addEventListener('change', e => e.target.files[0] && handleOSV(e.target.files[0]));
  $('priceInput').addEventListener('change', e => e.target.files.length && handlePrices(e.target.files));
  $('supplierSelect').addEventListener('change', renderOrders);
  $('copyOrderBtn').addEventListener('click', async () => { await navigator.clipboard.writeText(orderText()); toast('Заявка скопирована'); });
  $('downloadOrderBtn').addEventListener('click', downloadCSV);
  $('menuBtn').addEventListener('click', openDrawer); $('closeMenuBtn').addEventListener('click', closeDrawer); $('drawerBackdrop').addEventListener('click', closeDrawer); $('resetSuppliersBtn').addEventListener('click', resetSuppliersOnly); $('resetBtn').addEventListener('click', resetAll);
  document.querySelectorAll('[data-tab]').forEach(btn => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));
  document.querySelectorAll('#abcFilter button').forEach(btn => btn.addEventListener('click', () => { state.abcFilter = btn.dataset.abc; document.querySelectorAll('#abcFilter button').forEach(b => b.classList.toggle('active', b === btn)); renderLists(); }));
});
