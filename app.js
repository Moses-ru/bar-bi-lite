const tg = window.Telegram?.WebApp;
if (tg) { tg.ready(); tg.expand(); tg.setHeaderColor('#07111f'); tg.setBackgroundColor('#07111f'); }

const CONFIG = { periodDays: 30, targetDays: 10, riskDays: 3, overstockDays: 30 };
const state = { stock: [], suppliers: {}, supplierProducts: {}, charts: {}, lastFile: '' };

const $ = (id) => document.getElementById(id);
const money = (v) => `${Math.round(Number(v)||0).toLocaleString('ru-RU')} ₽`;
const num = (v, d=3) => Number(v || 0).toLocaleString('ru-RU', { maximumFractionDigits:d });
const normalize = (s) => String(s ?? '').toLowerCase().replace(/ё/g,'е').replace(/[.,;:()\[\]"'«»]/g,' ').replace(/\s+/g,' ').trim();
function toNumber(v){ if(v == null || v === '') return 0; if(typeof v === 'number') return isFinite(v) ? v : 0; const n = parseFloat(String(v).replace(/\s/g,'').replace(',', '.')); return isNaN(n) ? 0 : n; }
function toast(text){ const t=$('toast'); t.textContent=text; t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),2600); }

function categoryOf(name){
  const s = normalize(name);
  const tests = [
    ['Фрукты и овощи', ['апельсин','лимон','лайм','банан','гранат','мята','тархун','шпинат','огур','морковь','ягод','клубник','брусник','клюкв','зелень','базилик','имбирь','яблок','груша','малина','смородина','арбуз','дыня','перец']],
    ['Вино / игристое', ['вино','просекко','шампан','игрист','каберне','шардоне','киндзмараули','пино','рислинг','совиньон','мерло','кьянти']],
    ['Крепкий алкоголь', ['водка','виски','джин','ром','текила','коньяк','бренди','ликер','настойка','бурбон','абсент','самбука','егермейстер','мартини','апероль','кампари']],
    ['Пиво / сидр', ['пиво','сидр','эль','лагер','кроненбург','хофброй','бакалар']],
    ['Сиропы / пюре', ['сироп','пюре','pinch','drop','барбарис','маракуй','манго']],
    ['Соки / напитки', ['сок','кола','coca','pepsi','вода','боржоми','морс','тоник','лимонад','энергетик','rich','рич']],
    ['Молочка / сливки', ['молоко','сливки','сыр','мороженое']],
    ['Расходники', ['салфет','трубоч','стакан','крышк','перчат','мешок','чековая','уголь','фольга']]
  ];
  for (const [cat, words] of tests) if (words.some(w => s.includes(w))) return cat;
  return 'Другое';
}
function isVegFruit(row){ return row.category === 'Фрукты и овощи'; }

function supplierFor(product, category){
  const key = normalize(product);
  for (const [p, sup] of Object.entries(state.supplierProducts)) {
    if (key.includes(p) || p.includes(key)) return sup;
  }
  if (category === 'Фрукты и овощи') return 'Овощник';
  return 'Не указан';
}

async function readWorkbook(file){
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: 'array', cellDates: true });
}

function parseOSVWorkbook(wb){
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header:1, raw:true, defval:'' });
  let start = 0;
  for (let i=0; i<rows.length; i++) {
    const joined = rows[i].map(x => normalize(x)).join(' ');
    if (joined.includes('наименование') && joined.includes('ед изм')) { start = i+1; break; }
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
    const writeoff = Math.abs(toNumber(r[11]));
    const shortage = Math.abs(toNumber(r[17]));
    if (!unit && !stockEnd && !stockSum && !salesQty) continue;
    const dailySales = salesQty / CONFIG.periodDays;
    const daysLeft = dailySales > 0 ? stockEnd / dailySales : 9999;
    const category = categoryOf(product);
    const recommendedOrder = Math.max(0, dailySales * CONFIG.targetDays - stockEnd);
    const overstockMoney = (daysLeft >= CONFIG.overstockDays && salesQty > 0) ? Math.max(0, stockSum) : 0;
    data.push({ product, unit, category, supplier: supplierFor(product, category), stockEnd, stockSum, salesQty, salesCost, dailySales, daysLeft, recommendedOrder, writeoff, shortage, overstockMoney });
  }
  return data;
}

function parsePriceWorkbook(wb, fileName){
  let found = 0;
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header:1, raw:true, defval:'' });
    let headerIdx = -1, productCol = -1, supplierCol = -1;
    for (let i=0; i<Math.min(rows.length, 30); i++) {
      rows[i].forEach((cell, idx) => {
        const c = normalize(cell);
        if (['наименование','товар','номенклатура','продукт'].some(k => c.includes(k))) { headerIdx=i; productCol=idx; }
        if (c.includes('поставщик')) { supplierCol=idx; }
      });
      if (productCol >= 0) break;
    }
    const fallbackSupplier = sheetName.length > 2 && !normalize(sheetName).includes('лист') ? sheetName.trim() : fileName.replace(/\.(xlsx|xls)$/i,'').replace(/прайс|листы|поставщиков/gi,'').trim() || 'Поставщик';
    const start = headerIdx >= 0 ? headerIdx + 1 : 0;
    const pcol = productCol >= 0 ? productCol : 0;
    for (const r of rows.slice(start)) {
      const product = String(r[pcol] ?? '').trim();
      if (!product || product.length < 3 || normalize(product).includes('итого')) continue;
      const supplier = supplierCol >= 0 && r[supplierCol] ? String(r[supplierCol]).trim() : fallbackSupplier;
      if (!supplier) continue;
      state.suppliers[supplier] = true;
      state.supplierProducts[normalize(product)] = supplier;
      found++;
    }
  }
  localStorage.setItem('barbi_suppliers', JSON.stringify(state.suppliers));
  localStorage.setItem('barbi_supplier_products', JSON.stringify(state.supplierProducts));
  return found;
}

function loadLocalSuppliers(){
  try { state.suppliers = JSON.parse(localStorage.getItem('barbi_suppliers') || '{}'); } catch { state.suppliers = {}; }
  try { state.supplierProducts = JSON.parse(localStorage.getItem('barbi_supplier_products') || '{}'); } catch { state.supplierProducts = {}; }
  state.suppliers['Овощник'] = true;
}

function calcSummary(){
  const s = state.stock;
  return {
    sku: s.length,
    stockSum: s.reduce((a,r)=>a+r.stockSum,0),
    stops: s.filter(r=>r.stockEnd<=0).length,
    risk: s.filter(r=>r.stockEnd>0 && r.daysLeft<=CONFIG.riskDays).length,
    over: s.filter(r=>r.daysLeft>=CONFIG.overstockDays && r.salesQty>0).length,
    orderQty: s.reduce((a,r)=>a+r.recommendedOrder,0),
    frozen: s.reduce((a,r)=>a+r.overstockMoney,0)
  };
}
function byCategory(){
  const m = {};
  state.stock.forEach(r => { if(!m[r.category]) m[r.category]={category:r.category, stockSum:0, salesQty:0, sku:0}; m[r.category].stockSum += r.stockSum; m[r.category].salesQty += r.salesQty; m[r.category].sku++; });
  return Object.values(m).sort((a,b)=>b.stockSum-a.stockSum);
}
function topBy(key, n=10){ return [...state.stock].sort((a,b)=>(b[key]||0)-(a[key]||0)).slice(0,n); }
function orderRows(supplier='Все поставщики'){
  let rows = state.stock.filter(r=>r.recommendedOrder>0);
  if (supplier && supplier !== 'Все поставщики') rows = rows.filter(r => r.supplier === supplier);
  return rows.sort((a,b)=>b.recommendedOrder-a.recommendedOrder);
}
function abcRows(){
  const rows = [...state.stock].filter(r=>r.salesCost>0 || r.salesQty>0).sort((a,b)=>b.salesCost-a.salesCost);
  const total = rows.reduce((a,r)=>a+r.salesCost,0) || rows.reduce((a,r)=>a+r.salesQty,0) || 1;
  let acc = 0;
  return rows.map(r => { const v = r.salesCost || r.salesQty; acc += v; const share = acc/total; return {...r, abc: share <= .8 ? 'A' : share <= .95 ? 'B' : 'C', share: v/total*100}; });
}

function renderAll(){
  const s = calcSummary();
  $('kpiSku').textContent = s.sku;
  $('kpiStock').textContent = money(s.stockSum);
  $('kpiStops').textContent = s.stops;
  $('kpiRisk').textContent = s.risk;
  $('kpiOver').textContent = s.over;
  $('kpiOrder').textContent = num(s.orderQty, 0);
  $('fileStatus').innerHTML = state.stock.length ? `<b>ОСВ загружена: ${state.lastFile}</b><span>Поставщиков в базе: ${Object.keys(state.suppliers).length}. Овощник назначается автоматически для фруктов и овощей.</span>` : `<b>Данные не загружены</b><span>Сначала загрузи ОСВ. Прайсы нужны для заявок по поставщикам.</span>`;
  renderInsights(s); renderSupplierSelect(); renderOrders(); renderLists(); renderCharts();
}

function renderInsights(s){
  const stopTop = topBy('salesQty', 6).filter(r=>r.stockEnd<=0).map(r=>r.product).slice(0,4).join(', ');
  const noSup = state.stock.filter(r=>r.supplier === 'Не указан').length;
  const html = [];
  if (s.stops) html.push(`<div class="insight critical"><b>Критично · Стопы</b><p>В стопе ${s.stops} позиций. Проверь в первую очередь: ${stopTop || 'позиции с нулевым остатком'}.</p></div>`);
  if (s.risk) html.push(`<div class="insight warn"><b>Риск стопа</b><p>${s.risk} позиций закончатся примерно за ${CONFIG.riskDays} дня или быстрее.</p></div>`);
  if (s.frozen) html.push(`<div class="insight warn"><b>Оверсток</b><p>В лишнем запасе заморожено около ${money(s.frozen)}. Не заказывай эти позиции без необходимости.</p></div>`);
  if (noSup) html.push(`<div class="insight"><b>Поставщики</b><p>У ${noSup} товаров не найден поставщик. Загрузи прайс-листы, чтобы заявка стала точнее.</p></div>`);
  if (!html.length) html.push(`<div class="insight good"><b>Все спокойно</b><p>Критичных проблем по текущей ОСВ не найдено.</p></div>`);
  $('insights').innerHTML = html.join('');
}

function renderSupplierSelect(){
  const suppliers = ['Все поставщики', ...new Set(state.stock.map(r=>r.supplier).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'ru'));
  $('supplierSelect').innerHTML = suppliers.map(s=>`<option>${s}</option>`).join('');
}
function renderOrders(){
  const supplier = $('supplierSelect').value || 'Все поставщики';
  const rows = orderRows(supplier).slice(0,80);
  $('orderList').innerHTML = rows.length ? rows.map(r=>`<div class="row"><div><b>${r.product}</b><small>${r.category} · ${r.supplier} · остаток: ${num(r.stockEnd)} ${r.unit}</small></div><div class="value">${num(r.recommendedOrder)} ${r.unit}</div></div>`).join('') : '<div class="row"><b>Нет позиций к заказу</b><small>По выбранному поставщику заявка пустая.</small></div>';
}
function renderLists(){
  $('topStockList').innerHTML = topBy('stockSum', 10).map((r,i)=>`<div class="row"><div><b>${i+1}. ${r.product}</b><small>${r.category} · ${r.supplier}</small></div><div class="value">${money(r.stockSum)}</div></div>`).join('');
  $('topSalesList').innerHTML = topBy('salesQty', 10).map((r,i)=>`<div class="row"><div><b>${i+1}. ${r.product}</b><small>${r.category} · остаток: ${num(r.stockEnd)} ${r.unit}</small></div><div class="value">${num(r.salesQty)} ${r.unit}</div></div>`).join('');
  $('stopList').innerHTML = state.stock.filter(r=>r.stockEnd<=0).sort((a,b)=>b.salesQty-a.salesQty).slice(0,80).map(r=>`<div class="row danger"><div><b>${r.product}</b><small>${r.category} · ${r.supplier}</small></div><div class="value">${num(r.stockEnd)} ${r.unit}</div></div>`).join('') || '<div class="row"><b>Стопов нет</b></div>';
  $('overList').innerHTML = state.stock.filter(r=>r.daysLeft>=CONFIG.overstockDays && r.salesQty>0).sort((a,b)=>b.stockSum-a.stockSum).slice(0,80).map(r=>`<div class="row warn"><div><b>${r.product}</b><small>${r.category} · запас: ${num(r.daysLeft,1)} дней</small></div><div class="value">${money(r.stockSum)}</div></div>`).join('') || '<div class="row"><b>Оверстоков нет</b></div>';
  const abc = abcRows();
  $('abcList').innerHTML = abc.slice(0,30).map(r=>`<div class="row"><div><b>${r.product}</b><small>${r.category} · класс ${r.abc}</small></div><div class="value">${num(r.share,1)}%</div></div>`).join('') || '<div class="row"><b>Нет данных для ABC</b></div>';
}

const baseChart = { chart:{ foreColor:'#dbe4f3', toolbar:{show:false}, background:'transparent' }, theme:{mode:'dark'}, grid:{borderColor:'rgba(255,255,255,.08)'}, tooltip:{theme:'dark'}, legend:{labels:{colors:'#dbe4f3'}} };
function mountChart(id, options){ if (state.charts[id]) state.charts[id].destroy(); state.charts[id] = new ApexCharts($(id), options); state.charts[id].render(); }
function renderCharts(){
  if (!state.stock.length) return;
  const cats = byCategory().slice(0,8);
  mountChart('categoryChart', { ...baseChart, chart:{...baseChart.chart,type:'donut',height:300}, series:cats.map(x=>Math.round(x.stockSum)), labels:cats.map(x=>x.category), dataLabels:{style:{fontSize:'13px',fontWeight:800}}, plotOptions:{pie:{donut:{size:'58%',labels:{show:true,total:{show:true,label:'Всего',formatter:()=>money(cats.reduce((a,c)=>a+c.stockSum,0))}}}}}, stroke:{width:0}, colors:['#8b5cf6','#38bdf8','#34d399','#f59e0b','#fb7185','#a78bfa','#22c55e','#eab308'] });
  const topSales = topBy('salesQty', 10).reverse();
  mountChart('salesChart', { ...baseChart, chart:{...baseChart.chart,type:'bar',height:420}, plotOptions:{bar:{horizontal:true,borderRadius:6}}, series:[{name:'Расход',data:topSales.map(r=>Number(r.salesQty.toFixed(3)))}], xaxis:{categories:topSales.map(r=>r.product),labels:{style:{fontSize:'12px'}}}, yaxis:{labels:{style:{fontSize:'13px'}}}, colors:['#8b5cf6'] });
  const abc = abcRows(); const cnt = {A:0,B:0,C:0}; abc.forEach(r=>cnt[r.abc]++);
  mountChart('abcChart', { ...baseChart, chart:{...baseChart.chart,type:'bar',height:280}, series:[{name:'Позиций',data:[cnt.A,cnt.B,cnt.C]}], xaxis:{categories:['A — ключевые','B — средние','C — хвост']}, plotOptions:{bar:{borderRadius:8,columnWidth:'45%'}}, colors:['#34d399'] });
}

function orderText(){
  const supplier = $('supplierSelect').value || 'Все поставщики';
  const rows = orderRows(supplier);
  const lines = [`ЗАЯВКА · ${supplier}`, ''];
  rows.forEach(r => lines.push(`• ${r.product} — ${num(r.recommendedOrder)} ${r.unit}`));
  return lines.join('\n');
}
function downloadCSV(){
  const supplier = $('supplierSelect').value || 'Все поставщики';
  const rows = orderRows(supplier);
  const header = ['Поставщик','Товар','Категория','Ед. изм.','Остаток','Расход','К заказу'];
  const csv = [header, ...rows.map(r=>[r.supplier,r.product,r.category,r.unit,r.stockEnd,r.salesQty,r.recommendedOrder])].map(row=>row.map(x=>`"${String(x).replace(/"/g,'""')}"`).join(';')).join('\n');
  const blob = new Blob(['\ufeff'+csv], {type:'text/csv;charset=utf-8'});
  const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download=`zayavka_${supplier}.csv`; a.click(); URL.revokeObjectURL(url);
}

async function handleOSV(file){
  const wb = await readWorkbook(file);
  state.stock = parseOSVWorkbook(wb);
  state.lastFile = file.name;
  Object.keys(state.charts).forEach(k=>{ try{state.charts[k].destroy()}catch{} }); state.charts = {};
  renderAll(); toast(`ОСВ загружена: ${state.stock.length} товаров`);
}
async function handlePrices(files){
  let total = 0;
  for (const file of files) { const wb = await readWorkbook(file); total += parsePriceWorkbook(wb, file.name); }
  if (state.stock.length) state.stock.forEach(r => r.supplier = supplierFor(r.product, r.category));
  renderAll(); toast(`Прайсы загружены: ${total} позиций`);
}

function switchTab(tab){
  document.querySelectorAll('[data-tab]').forEach(b=>b.classList.toggle('active', b.dataset.tab===tab));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.toggle('active', p.id===tab));
  setTimeout(()=>{Object.values(state.charts).forEach(c=>{try{c.windowResizeHandler()}catch{}})},50);
}

document.addEventListener('DOMContentLoaded', () => {
  loadLocalSuppliers(); renderAll();
  $('osvInput').addEventListener('change', e => e.target.files[0] && handleOSV(e.target.files[0]));
  $('priceInput').addEventListener('change', e => e.target.files.length && handlePrices(e.target.files));
  $('supplierSelect').addEventListener('change', renderOrders);
  $('copyOrderBtn').addEventListener('click', async()=>{ await navigator.clipboard.writeText(orderText()); toast('Заявка скопирована'); });
  $('downloadOrderBtn').addEventListener('click', downloadCSV);
  $('resetBtn').addEventListener('click', ()=>{ if(confirm('Очистить поставщиков и текущую ОСВ?')){ localStorage.removeItem('barbi_suppliers'); localStorage.removeItem('barbi_supplier_products'); state.stock=[]; loadLocalSuppliers(); renderAll(); toast('Данные очищены'); }});
  document.querySelectorAll('[data-tab]').forEach(b=>b.addEventListener('click',()=>switchTab(b.dataset.tab)));
});
