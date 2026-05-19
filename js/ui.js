// ═══════════════════════════════════════════════════
// 智词引擎 · UI 渲染与交互
// ═══════════════════════════════════════════════════

var KC = window.KC || {};

// ── UI 常量与状态 ──
KC.UI = {};
KC.UI.L1_META = {
  '国内酒店': { cls: 'hotel', color: '#1a73e8', pct: 0 },
  '国内机票': { cls: 'flight', color: '#10b981', pct: 0 },
  '国际酒店': { cls: 'hotel-intl', color: '#8b5cf6', pct: 0 },
  '国际机票': { cls: 'flight-intl', color: '#f59e0b', pct: 0 },
  '景区': { cls: 'scenic', color: '#ef4444', pct: 0 },
  '火车票': { cls: 'train', color: '#ec4899', pct: 0 },
  '公共': { cls: 'public', color: '#06b6d4', pct: 0 },
  '待确认': { cls: 'unclassified', color: '#9ca3af', pct: 0 },
};
KC.UI.L1_ORDER = ['国内酒店','国内机票','国际酒店','国际机票','景区','火车票','公共','待确认'];
KC.UI.L2_COLORS = ['#1a73e8','#10b981','#8b5cf6','#f59e0b','#ef4444','#ec4899','#0d9488','#f97316','#6366f1','#84cc16','#06b6d4','#e11d48'];

// 局部别名（函数内裸引用 → 命名空间属性的快捷方式）
const L1_META = KC.UI.L1_META;
const L1_ORDER = KC.UI.L1_ORDER;
const L2_COLORS = KC.UI.L2_COLORS;
const PAGE_SIZE = 50;
KC.UI.PAGE_SIZE = PAGE_SIZE;

let allResults = [];
let filteredResults = [];
let currentPage = 1;
let charts = {};

KC.State = {
  allResults: [],
  filteredResults: [],
  currentPage: 1,
  charts: {},
  analysisCharts: {},
  crossCharts: {},
  geoCharts: {},
  geoSelectedL1: new Set(),
};

function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.getElementById('tab-paste').classList.toggle('hidden', tab !== 'paste');
  document.getElementById('tab-upload').classList.toggle('hidden', tab !== 'upload');
}

function loadSample() {
  const samples = [
    '上海酒店推荐','北京民宿性价比高','杭州旅游攻略','苏州住哪个地铁站比较方便',
    '广州长隆迎宾路城际酒店','三亚亚特兰蒂斯酒店','成都宽窄巷子附近酒店',
    '重庆洪崖洞住宿','武汉黄鹤楼附近酒店','南京总统府门票',
    '北京故宫门票多少钱','上海迪士尼乐园门票','张家界天门山门票',
    '五一机票特价','国庆去哪里玩便宜机票','北京飞上海机票',
    '从成都到重庆机票','杭州到厦门航班','深圳到海口特价机票',
    '国航值机选座','机票退改签手续费','儿童机票怎么买',
    '随心飞第三次卡','机票盲盒真的假的','高铁票候补成功率高吗',
    '火车票抢票软件哪个快','长沙南站怎么去市区','五一车票候补',
    '高铁贵宾厅有什么服务','济州岛酒店推荐','日本东京酒店性价比',
    '曼谷民宿近地铁口','普吉岛亲子酒店推荐','巴厘岛悬崖酒店',
    '首尔仁川机场附近酒店','巴黎酒店安全区域','五一日本往返机票',
    '五一去巴厘岛费用','济州岛购物免税店攻略','马尔代夫选岛攻略',
    '携程酒店预订','去哪儿机票比价','飞猪民宿','美团酒店',
    '同程旅行订房优惠','如家酒店汉庭全季','希尔顿万豪洲际',
    '五一天安门升旗仪式预约','国庆长城门票','春节三亚机票贵',
    '暑假带娃去哪里玩好','清明祭祖去哪里','端午粽子推荐',
    '元旦跨年烟花秀','情人节适合去哪里','圣诞节约会餐厅',
  ];
  document.getElementById('kw-text').value = samples.join('\n');
}

async function runClassify() {
  const t0 = performance.now();
  showLoading(true);

  // 让浏览器先渲染 loading 动画，避免同步代码阻塞 UI
  await new Promise(r => setTimeout(r, 0));

  let raw = [];
  const tab = document.querySelector('.tab-btn.active').dataset.tab;

  try {
    if (tab === 'paste') {
      const text = document.getElementById('kw-text').value.trim();
      if (!text) { showLoading(false); alert('请输入关键词'); return; }
      raw = text.split('\n').map(s => s.trim()).filter(Boolean);
    } else {
      if (typeof XLSX === 'undefined') { showLoading(false); alert('Excel 解析库未加载，请刷新页面重试或使用粘贴模式'); return; }
      const file = document.getElementById('file-input').files[0];
      if (!file) { showLoading(false); alert('请先上传文件'); return; }
      const ab = await file.arrayBuffer();
      const data = new Uint8Array(ab);
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const col = parseInt(document.getElementById('col-kw').value) || 1;
      const colCost = parseInt(document.getElementById('col-cost').value) || 0;
      const colImp = parseInt(document.getElementById('col-impression').value) || 0;
      const colClick = parseInt(document.getElementById('col-click').value) || 0;
      const rowStart = parseInt(document.getElementById('row-start').value) || 1;
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let r = range.s.r + rowStart - 1; r <= range.e.r; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c: col - 1 })];
        const val = cell ? (typeof cell.v === 'number' ? String(cell.v) : String(cell.v || '')) : '';
        if (val.trim()) {
          const item = { kw: val.trim(), cost: 0, impression: 0, click: 0 };
          if (colCost > 0) {
            const costCell = ws[XLSX.utils.encode_cell({ r, c: colCost - 1 })];
            item.cost = costCell ? (parseFloat(costCell.v) || 0) : 0;
          }
          if (colImp > 0) {
            const impCell = ws[XLSX.utils.encode_cell({ r, c: colImp - 1 })];
            item.impression = impCell ? (parseInt(impCell.v) || 0) : 0;
          }
          if (colClick > 0) {
            const clickCell = ws[XLSX.utils.encode_cell({ r, c: colClick - 1 })];
            item.click = clickCell ? (parseInt(clickCell.v) || 0) : 0;
          }
          raw.push(item);
        }
      }
    }
  } catch(e) {
    showLoading(false);
    alert('读取失败: ' + e.message);
    return;
  }

  if (raw.length === 0) { showLoading(false); alert('未找到关键词'); return; }

  const tClassify = performance.now();
  allResults = raw.map(item => {
    const result = KC.Classifier.classify(item.kw || item);
    if (typeof item === 'object') {
      result.cost = item.cost || 0;
      result.impression = item.impression || 0;
      result.click = item.click || 0;
    } else {
      result.cost = 0;
      result.impression = 0;
      result.click = 0;
    }
    return result;
  });
  const msClassify = Math.round(performance.now() - tClassify);

  const ms = Math.round(performance.now() - t0);
  console.log(`分类 ${raw.length} 条关键词耗时: ${msClassify}ms（总耗时: ${ms}ms，含文件读取 ${ms - msClassify}ms）`);
  showLoading(false);
  renderAll(allResults, ms);
}

function initFilters() {
  // 一级分类筛选器
  const l1Tags = document.getElementById('filter-l1-tags');
  const l1s = [...new Set(allResults.map(r => r.l1).filter(Boolean))].sort();
  l1Tags.innerHTML = '<span class="filter-tag active" data-value="" onclick="selectFilter(this, \'l1\')">全部</span>' +
    l1s.map(l1 => `<span class="filter-tag" data-value="${l1}" onclick="selectFilter(this, 'l1')">${l1}</span>`).join('');

  // 二级分类筛选器 - 初始显示全部
  updateL2FilterTags('');

  // 国家下拉
  const countrySelect = document.getElementById('filter-country');
  const countries = [...new Set(allResults.map(r => r.country).filter(Boolean))].sort();
  countrySelect.innerHTML = '<option value="">全部国家</option>' +
    countries.map(c => `<option value="${c}">${c}</option>`).join('');

  // 城市下拉（级联：受分区/等级控制）
  updateCityDropdown();

  // 分区下拉
  const zoneSelect = document.getElementById('filter-zone');
  const zones = [...new Set(allResults.map(r => r.zone).filter(Boolean))].sort();
  zoneSelect.innerHTML = '<option value="">全部分区</option>' +
    zones.map(z => `<option value="${z}">${z}</option>`).join('');

  // 等级下拉
  const tierSelect = document.getElementById('filter-tier');
  const tiers = [...new Set(allResults.map(r => r.tier).filter(Boolean))].sort();
  tierSelect.innerHTML = '<option value="">全部等级</option>' +
    tiers.map(t => `<option value="${t}">${t}</option>`).join('');
}

// 根据选中的一级分类更新二级分类标签
function updateL2FilterTags(selectedL1) {
  const l2Tags = document.getElementById('filter-l2-tags');
  let l2s;
  
  if (selectedL1 === '') {
    // 全部：显示所有二级分类
    l2s = [...new Set(allResults.map(r => r.l2).filter(Boolean))].sort();
  } else {
    // 只显示该一级分类下的二级分类
    l2s = [...new Set(allResults.filter(r => r.l1 === selectedL1).map(r => r.l2).filter(Boolean))].sort();
  }
  
  l2Tags.innerHTML = '<span class="filter-tag active" data-value="" onclick="selectFilter(this, \'l2\')">全部</span>' +
    l2s.map(l2 => `<span class="filter-tag" data-value="${l2}" onclick="selectFilter(this, 'l2')">${l2}</span>`).join('');
}

// 根据分区/等级筛选更新城市下拉
function updateCityDropdown() {
  const filterZone = document.getElementById('filter-zone')?.value || '';
  const filterTier = document.getElementById('filter-tier')?.value || '';
  let base = allResults;
  if (filterZone) base = base.filter(r => r.zone === filterZone);
  if (filterTier) base = base.filter(r => r.tier === filterTier);
  const cities = [...new Set(base.map(r => r.city).filter(Boolean))].sort();
  const citySelect = document.getElementById('filter-city');
  const prevCity = citySelect?.value || '';
  citySelect.innerHTML = '<option value="">全部城市</option>' +
    cities.map(c => `<option value="${c}">${c}</option>`).join('');
  // 恢复之前选中（如果还在列表中）
  if (prevCity && cities.includes(prevCity)) citySelect.value = prevCity;
}

// 分区/等级变化时的级联处理
function onZoneTierChange() {
  updateCityDropdown();
  renderTable(1);
}

function selectFilter(el, type) {
  const container = type === 'l1' ? document.getElementById('filter-l1-tags') : document.getElementById('filter-l2-tags');
  container.querySelectorAll('.filter-tag').forEach(tag => tag.classList.remove('active'));
  el.classList.add('active');
  
  // 如果点击的是一级分类，更新二级分类标签
  if (type === 'l1') {
    const selectedL1 = el.dataset.value;
    updateL2FilterTags(selectedL1);
  }
  
  renderTable(1);
}

function clearFilters() {
  document.getElementById('search-box').value = '';
  document.querySelectorAll('.filter-tag').forEach(tag => tag.classList.remove('active'));
  document.querySelector('#filter-l1-tags .filter-tag[data-value=""]').classList.add('active');
  // 重置二级分类标签为全部
  updateL2FilterTags('');
  document.getElementById('filter-country').value = '';
  document.getElementById('filter-zone').value = '';
  document.getElementById('filter-tier').value = '';
  document.getElementById('filter-city').value = '';
  // 重建城市下拉（无分区/等级限制）
  updateCityDropdown();
  renderTable(1);
}

function renderAll(results, ms) {
  ensureDataLabelsPlugin();
  document.getElementById('result-area').classList.remove('hidden');
  initFilters();
  currentPage = 1;
  document.getElementById('search-box').value = '';

  // 统计
  document.getElementById('stat-total').textContent = results.length.toLocaleString();
  const l1Set = new Set(results.map(r => r.l1));
  const l2Set = new Set(results.map(r => r.l2));
  document.getElementById('stat-l1').textContent = l1Set.size;
  document.getElementById('stat-l2').textContent = l2Set.size;
  document.getElementById('stat-time').textContent = ms + 'ms';

  // L1 计数
  const l1Count = {};
  const l2Count = {};
  results.forEach(r => {
    l1Count[r.l1] = (l1Count[r.l1] || 0) + 1;
    l2Count[r.l2] = (l2Count[r.l2] || 0) + 1;
  });

  // L1 百分比
  L1_ORDER.forEach(l1 => {
    if (l1Count[l1] !== undefined)
      L1_META[l1].pct = l1Count[l1] / results.length;
  });

  // 渲染表格（优先渲染，确保核心数据始终可见）
  filteredResults = [...results];
  renderTable(1);

  // 渲染图表（CDN 可能未加载，加容错）
  try { renderL1Chart(l1Count, results.length); } catch(e) { console.warn('L1图表渲染跳过:', e.message); }
  try { renderL2Chart(l2Count, results.length); } catch(e) { console.warn('L2图表渲染跳过:', e.message); }

  // 渲染数据分析（如果有花费数据）
  const hasCostData = results.some(r => r.cost > 0);
  if (hasCostData) {
    document.getElementById('analysis-section').classList.remove('hidden');
    document.getElementById('cross-analysis-section').classList.remove('hidden');
    try { renderCostAnalysis(); } catch(e) { console.warn('花费分析渲染跳过:', e.message); }
    initCrossFilters();
    try { renderCrossAnalysis(); } catch(e) { console.warn('交叉分析渲染跳过:', e.message); }
  } else {
    document.getElementById('analysis-section').classList.add('hidden');
    document.getElementById('cross-analysis-section').classList.add('hidden');
  }

  // 渲染关键词词云
  try { initWcFilters(); } catch(e) { console.warn('词云筛选器初始化跳过:', e.message); }
  try { renderWordCloud(); } catch(e) { console.warn('词云渲染跳过:', e.message); }

  // 渲染地理维度分析（有明确标签数据时显示）
  var hasLabeledData = results.some(function(r) { return r.l1 && r.l1 !== '待确认'; });
  if (hasLabeledData) {
    document.getElementById('geo-analysis-section').classList.remove('hidden');
    try { initGeoFilters(); renderGeoAnalysis(); } catch(e) { console.warn('地理分析渲染跳过:', e.message); }
  } else {
    document.getElementById('geo-analysis-section').classList.add('hidden');
  }

  // 重置词包生成器
  pkgResults = [];
  document.getElementById('pkg-stats').classList.add('hidden');
  document.getElementById('pkg-table-wrap').classList.add('hidden');
  document.getElementById('pkg-hint').textContent = `分类完成，可生成词包组合`;
  document.getElementById('pkg-search').value = '';
  initPkgFilters();
}

// Chart.js 数据标签插件
const dataLabelsPlugin = {
  id: 'dataLabels',
  afterDatasetsDraw(chart, args, options) {
    const { ctx } = chart;
    chart.data.datasets.forEach((dataset, i) => {
      const meta = chart.getDatasetMeta(i);
      meta.data.forEach((element, index) => {
        const value = dataset.data[index];
        if (value === 0 || value === undefined || value === null) return;
        
        ctx.save();
        ctx.font = 'bold 11px Arial';
        ctx.fillStyle = '#fff';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        
        const pos = element.tooltipPosition();
        const label = value.toLocaleString();
        
        // 根据图表类型调整位置
        if (chart.config.type === 'doughnut' || chart.config.type === 'pie') {
          // 环形图：显示百分比和数量
          const total = dataset.data.reduce((sum, v) => sum + (v || 0), 0);
          const pct = total > 0 ? (value / total * 100).toFixed(1) : 0;
          const labelText = pct + '%';
          const countText = '(' + value.toLocaleString() + ')';
          
          const angle = (element.startAngle + element.endAngle) / 2;
          const radius = (element.innerRadius + element.outerRadius) / 2;
          const x = element.x + Math.cos(angle) * radius;
          const y = element.y + Math.sin(angle) * radius;
          
          // 如果扇形太小就不显示
          const arcSize = element.endAngle - element.startAngle;
          if (arcSize > 0.15) {
            // 绘制百分比
            ctx.font = 'bold 12px Arial';
            ctx.fillText(labelText, x, y - 6);
            // 绘制数量
            ctx.font = '10px Arial';
            ctx.fillText(countText, x, y + 8);
          }
        } else if (chart.config.type === 'bar') {
          // 柱状图：显示在柱子顶部
          ctx.fillStyle = '#1f2937';
          ctx.fillText(label, pos.x, pos.y - 5);
        }
        
        ctx.restore();
      });
    });
  }
};

// 注册插件（延迟注册：Chart 由 defer CDN 加载，内联脚本执行时 Chart 可能尚未就绪）
let _dlpRegistered = false;
function ensureDataLabelsPlugin() {
  if (!_dlpRegistered && typeof Chart !== 'undefined') {
    Chart.register(dataLabelsPlugin);
    _dlpRegistered = true;
  }
}

function renderL1Chart(counts, total) {
  if (charts.l1) charts.l1.destroy();
  const labels = L1_ORDER.filter(l => counts[l]);
  const data = labels.map(l => counts[l]);
  const colors = labels.map(l => L1_META[l]?.color || '#999');

  charts.l1 = new Chart(document.getElementById('chart-l1'), {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colors,
        borderWidth: 2,
        borderColor: '#fff',
        hoverOffset: 6,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: { font: { size: 12 }, padding: 14, usePointStyle: true, pointStyleWidth: 10 }
        },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ${ctx.parsed.toLocaleString()} (${(ctx.parsed/total*100).toFixed(1)}%)`
          }
        }
      }
    }
  });
}

function renderL2Chart(counts, total) {
  if (charts.l2) charts.l2.destroy();
  // 合并 l1+l2 作为标签
  const l2Map = {};
  allResults.forEach(r => {
    const key = r.l1 + ' > ' + r.l2;
    l2Map[key] = (l2Map[key] || 0) + 1;
  });
  const sorted = Object.entries(l2Map).sort((a, b) => b[1] - a[1]);
  const labels = sorted.slice(0, 15).map(s => s[0]);
  const data = sorted.slice(0, 15).map(s => s[1]);

  charts.l2 = new Chart(document.getElementById('chart-l2'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '词数',
        data,
        backgroundColor: labels.map((_, i) => L2_COLORS[i % L2_COLORS.length]),
        borderRadius: 5,
        barThickness: 18,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.parsed.x.toLocaleString()} (${(ctx.parsed.x/total*100).toFixed(1)}%)`
          }
        }
      },
      scales: {
        x: { grid: { color: '#f3f4f6' }, ticks: { font: { size: 11 } } },
        y: { grid: { display: false }, ticks: { font: { size: 11 } } }
      }
    }
  });
}

function renderTable(page) {
  const search = document.getElementById('search-box').value.trim().toLowerCase();
  const filterL1 = document.querySelector('#filter-l1-tags .active')?.dataset.value || '';
  const filterL2 = document.querySelector('#filter-l2-tags .active')?.dataset.value || '';
  const filterCountry = document.getElementById('filter-country').value;
  const filterCity = document.getElementById('filter-city').value;
  const filterZone = document.getElementById('filter-zone').value;
  const filterTier = document.getElementById('filter-tier').value;

  filteredResults = allResults.filter(r => {
    if (search && !r.kw.toLowerCase().includes(search)) return false;
    if (filterL1 && r.l1 !== filterL1) return false;
    if (filterL2 && r.l2 !== filterL2) return false;
    if (filterCountry && r.country !== filterCountry) return false;
    if (filterCity && r.city !== filterCity) return false;
    if (filterZone && r.zone !== filterZone) return false;
    if (filterTier && r.tier !== filterTier) return false;
    return true;
  });
  currentPage = page;
  const start = (page - 1) * PAGE_SIZE;
  const end = start + PAGE_SIZE;
  const pageItems = filteredResults.slice(start, end);
  const totalPages = Math.max(1, Math.ceil(filteredResults.length / PAGE_SIZE));

  const tbody = document.getElementById('tbl-body');
  tbody.innerHTML = pageItems.map((r, i) => {
    const l1Meta = L1_META[r.l1] || { cls: '', color: '#999' };
    const ctr = r.impression > 0 ? (r.click / r.impression * 100).toFixed(2) + '%' : '-';
    const cpc = r.click > 0 ? '¥' + (r.cost / r.click).toFixed(2) : '-';
    return `<tr>
      <td class="num" style="color:#9ca3af">${start + i + 1}</td>
      <td style="max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.kw)}">${escHtml(r.kw)}</td>
      <td><span class="l1-badge ${l1Meta.cls}">${r.l1}</span></td>
      <td><span class="l2-tag">${r.l2}</span></td>
      <td style="color:#6366f1;font-weight:500">${r.country||'-'}</td>
      <td style="color:#059669;font-weight:500">${r.city||'-'}</td>
      <td style="color:#6366f1;font-weight:500">${r.zone||'-'}</td>
      <td style="color:#f59e0b;font-weight:500">${r.tier||'-'}</td>
      <td class="num" style="color:#ef4444">${r.cost ? '¥' + r.cost.toLocaleString() : '-'}</td>
      <td class="num" style="color:#3b82f6">${r.impression ? r.impression.toLocaleString() : '-'}</td>
      <td class="num" style="color:#10b981">${r.click ? r.click.toLocaleString() : '-'}</td>
      <td class="num">${ctr}</td>
      <td class="num">${cpc}</td>
    </tr>`;
  }).join('');

  // 分页
  const pagination = document.getElementById('pagination');
  if (filteredResults.length <= PAGE_SIZE) {
    pagination.innerHTML = '';
    return;
  }
  let html = `<span class="page-info">${filteredResults.length} 条结果，第 ${page}/${totalPages} 页</span>`;
  html += `<button class="page-btn" onclick="renderTable(1)">«</button>`;
  html += `<button class="page-btn" onclick="renderTable(${Math.max(1, page-1)})">‹</button>`;
  const pages = [];
  for (let p = Math.max(1, page-2); p <= Math.min(totalPages, page+2); p++) pages.push(p);
  pages.forEach(p => {
    html += `<button class="page-btn${p===page?' active':''}" onclick="renderTable(${p})">${p}</button>`;
  });
  html += `<button class="page-btn" onclick="renderTable(${Math.min(totalPages, page+1)})">›</button>`;
  html += `<button class="page-btn" onclick="renderTable(${totalPages})">»</button>`;
  pagination.innerHTML = html;
}

function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function exportCSV() {
  const rows = [['关键词','一级分类','二级分类','国家','城市','分区','等级','花费','曝光量','点击量','CTR','CPC']];
  filteredResults.forEach(r => {
    const ctr = r.impression > 0 ? (r.click / r.impression * 100).toFixed(2) + '%' : '-';
    const cpc = r.click > 0 ? (r.cost / r.click).toFixed(2) : '-';
    const kw = r.kw.replace(/[\r\n,]/g, ' ');
    rows.push([kw, r.l1, r.l2, r.country, r.city, r.zone, r.tier, r.cost, r.impression, r.click, ctr, cpc]);
  });
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  downloadFile('\ufeff' + csv, '关键词分类结果.csv', 'text/csv;charset=utf-8');
}

function downloadFile(content, filename, type) {
  const blob = new Blob([content], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function showLoading(show) {
  document.getElementById('loading').classList.toggle('show', show);
}

// ═══════════════════════════════════════════════════
// 数据分析模块
// ═══════════════════════════════════════════════════

let analysisCharts = {};

function switchAnalysisTab(tab) {
  document.querySelectorAll('[data-analysis]').forEach(btn => btn.classList.remove('active'));
  document.querySelector(`[data-analysis="${tab}"]`).classList.add('active');
  document.querySelectorAll('.analysis-content').forEach(el => el.classList.add('hidden'));
  document.getElementById(`analysis-${tab}`).classList.remove('hidden');
  
  // 渲染对应图表
  if (tab === 'cost') renderCostAnalysis();
  else if (tab === 'impression') renderImpressionAnalysis();
  else if (tab === 'click') renderClickAnalysis();
  else if (tab === 'efficiency') renderEfficiencyAnalysis();
}

function calculateMetrics() {
  const metrics = {};
  
  // 按一级分类汇总
  allResults.forEach(r => {
    const l1 = r.l1 || '未分类';
    if (!metrics[l1]) {
      metrics[l1] = {
        count: 0,
        cost: 0,
        impression: 0,
        click: 0
      };
    }
    metrics[l1].count++;
    metrics[l1].cost += r.cost || 0;
    metrics[l1].impression += r.impression || 0;
    metrics[l1].click += r.click || 0;
  });
  
  // 计算衍生指标
  Object.keys(metrics).forEach(l1 => {
    const m = metrics[l1];
    m.avgCost = m.count > 0 ? m.cost / m.count : 0;
    m.avgImpression = m.count > 0 ? m.impression / m.count : 0;
    m.avgClick = m.count > 0 ? m.click / m.count : 0;
    m.ctr = m.impression > 0 ? (m.click / m.impression * 100) : 0;
    m.cpc = m.click > 0 ? (m.cost / m.click) : 0;
  });
  
  return metrics;
}

function renderCostAnalysis() {
  const metrics = calculateMetrics();
  const l1s = Object.keys(metrics).sort((a, b) => metrics[b].cost - metrics[a].cost);
  const totalCost = Object.values(metrics).reduce((sum, m) => sum + m.cost, 0);
  
  // 渲染图表
  if (analysisCharts.cost) analysisCharts.cost.destroy();
  analysisCharts.cost = new Chart(document.getElementById('chart-cost'), {
    type: 'bar',
    data: {
      labels: l1s,
      datasets: [{
        label: '花费',
        data: l1s.map(l1 => metrics[l1].cost),
        backgroundColor: l1s.map(l1 => L1_META[l1]?.color || '#999'),
        borderRadius: 5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` 花费: ¥${ctx.parsed.y.toLocaleString()}`
          }
        }
      },
      scales: {
        y: { beginAtZero: true, ticks: { callback: v => '¥' + v.toLocaleString() } }
      }
    }
  });
  
  // 渲染表格
  const tbody = document.querySelector('#table-cost tbody');
  tbody.innerHTML = l1s.map(l1 => {
    const m = metrics[l1];
    const pct = totalCost > 0 ? (m.cost / totalCost * 100).toFixed(1) : 0;
    return `<tr>
      <td><span class="l1-badge ${L1_META[l1]?.cls || ''}">${l1}</span></td>
      <td class="num">${m.count.toLocaleString()}</td>
      <td class="num" style="color:#ef4444;font-weight:600">¥${m.cost.toLocaleString()}</td>
      <td class="num">${pct}%</td>
      <td class="num">¥${m.avgCost.toFixed(2)}</td>
      <td class="num">${m.impression.toLocaleString()}</td>
      <td class="num">${m.click.toLocaleString()}</td>
      <td class="num">¥${m.cpc.toFixed(2)}</td>
    </tr>`;
  }).join('');
}

function renderImpressionAnalysis() {
  const metrics = calculateMetrics();
  const l1s = Object.keys(metrics).sort((a, b) => metrics[b].impression - metrics[a].impression);
  const totalImp = Object.values(metrics).reduce((sum, m) => sum + m.impression, 0);
  
  // 渲染图表
  if (analysisCharts.impression) analysisCharts.impression.destroy();
  analysisCharts.impression = new Chart(document.getElementById('chart-impression'), {
    type: 'bar',
    data: {
      labels: l1s,
      datasets: [{
        label: '曝光量',
        data: l1s.map(l1 => metrics[l1].impression),
        backgroundColor: l1s.map(l1 => L1_META[l1]?.color || '#999'),
        borderRadius: 5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, ticks: { callback: v => v.toLocaleString() } }
      }
    }
  });
  
  // 渲染表格
  const tbody = document.querySelector('#table-impression tbody');
  tbody.innerHTML = l1s.map(l1 => {
    const m = metrics[l1];
    const pct = totalImp > 0 ? (m.impression / totalImp * 100).toFixed(1) : 0;
    return `<tr>
      <td><span class="l1-badge ${L1_META[l1]?.cls || ''}">${l1}</span></td>
      <td class="num">${m.count.toLocaleString()}</td>
      <td class="num" style="color:#3b82f6;font-weight:600">${m.impression.toLocaleString()}</td>
      <td class="num">${pct}%</td>
      <td class="num">${m.avgImpression.toFixed(0)}</td>
      <td class="num">${m.click.toLocaleString()}</td>
      <td class="num">${m.ctr.toFixed(2)}%</td>
    </tr>`;
  }).join('');
}

function renderClickAnalysis() {
  const metrics = calculateMetrics();
  const l1s = Object.keys(metrics).sort((a, b) => metrics[b].click - metrics[a].click);
  const totalClick = Object.values(metrics).reduce((sum, m) => sum + m.click, 0);
  
  // 渲染图表
  if (analysisCharts.click) analysisCharts.click.destroy();
  analysisCharts.click = new Chart(document.getElementById('chart-click'), {
    type: 'bar',
    data: {
      labels: l1s,
      datasets: [{
        label: '点击量',
        data: l1s.map(l1 => metrics[l1].click),
        backgroundColor: l1s.map(l1 => L1_META[l1]?.color || '#999'),
        borderRadius: 5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { beginAtZero: true, ticks: { callback: v => v.toLocaleString() } }
      }
    }
  });
  
  // 渲染表格
  const tbody = document.querySelector('#table-click tbody');
  tbody.innerHTML = l1s.map(l1 => {
    const m = metrics[l1];
    const pct = totalClick > 0 ? (m.click / totalClick * 100).toFixed(1) : 0;
    return `<tr>
      <td><span class="l1-badge ${L1_META[l1]?.cls || ''}">${l1}</span></td>
      <td class="num">${m.count.toLocaleString()}</td>
      <td class="num" style="color:#10b981;font-weight:600">${m.click.toLocaleString()}</td>
      <td class="num">${pct}%</td>
      <td class="num">${m.avgClick.toFixed(0)}</td>
      <td class="num">¥${m.cost.toLocaleString()}</td>
      <td class="num">¥${m.cpc.toFixed(2)}</td>
    </tr>`;
  }).join('');
}

function renderEfficiencyAnalysis() {
  const metrics = calculateMetrics();
  
  // 计算效率评分（综合CTR和CPC）
  Object.keys(metrics).forEach(l1 => {
    const m = metrics[l1];
    // 效率评分 = CTR / CPC * 100（CTR越高、CPC越低，效率越高）
    m.efficiency = m.cpc > 0 ? (m.ctr / m.cpc * 100) : 0;
  });
  
  const l1s = Object.keys(metrics).sort((a, b) => metrics[b].efficiency - metrics[a].efficiency);
  const totalCost = Object.values(metrics).reduce((sum, m) => sum + m.cost, 0);
  const totalImp = Object.values(metrics).reduce((sum, m) => sum + m.impression, 0);
  const totalClick = Object.values(metrics).reduce((sum, m) => sum + m.click, 0);
  
  // 渲染图表
  if (analysisCharts.efficiency) analysisCharts.efficiency.destroy();
  analysisCharts.efficiency = new Chart(document.getElementById('chart-efficiency'), {
    type: 'bar',
    data: {
      labels: l1s,
      datasets: [{
        label: '效率评分',
        data: l1s.map(l1 => metrics[l1].efficiency),
        backgroundColor: l1s.map(l1 => L1_META[l1]?.color || '#999'),
        borderRadius: 5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });
  
  // 渲染表格
  const tbody = document.querySelector('#table-efficiency tbody');
  tbody.innerHTML = l1s.map(l1 => {
    const m = metrics[l1];
    const costPct = totalCost > 0 ? (m.cost / totalCost * 100).toFixed(1) : 0;
    const impPct = totalImp > 0 ? (m.impression / totalImp * 100).toFixed(1) : 0;
    const clickPct = totalClick > 0 ? (m.click / totalClick * 100).toFixed(1) : 0;
    return `<tr>
      <td><span class="l1-badge ${L1_META[l1]?.cls || ''}">${l1}</span></td>
      <td class="num">${m.count.toLocaleString()}</td>
      <td class="num">${m.ctr.toFixed(2)}%</td>
      <td class="num">¥${m.cpc.toFixed(2)}</td>
      <td class="num">${costPct}%</td>
      <td class="num">${impPct}%</td>
      <td class="num">${clickPct}%</td>
      <td class="num" style="color:#8b5cf6;font-weight:600">${m.efficiency.toFixed(1)}</td>
    </tr>`;
  }).join('');
}

// ═══════════════════════════════════════════════════
// 地理维度分析模块
// ═══════════════════════════════════════════════════

var geoSelectedL1 = new Set();
var geoCharts = {};

// 区域固定顺序（去除"未识别"）
var GEO_REGION_ORDER = ['华东','华南','华北','华中','西南','西北','东北','海外'];
// 城市等级固定顺序（去除"未识别"）
var GEO_TIER_ORDER = ['一线城市','新一线城市','二线城市','三线城市','四线城市','五线城市','海外城市'];

function initGeoFilters() {
  // 获取所有非空、非"待确认"的业务线
  var l1Set = new Set();
  allResults.forEach(function(r) {
    if (r.l1 && r.l1 !== '待确认') {
      l1Set.add(r.l1);
    }
  });

  // 默认全选
  geoSelectedL1 = new Set(l1Set);

  var container = document.getElementById('geo-filter-l1');
  container.innerHTML = [...l1Set].sort().map(function(l1) {
    return '<span class="filter-tag active" data-l1="' + l1 + '" onclick="toggleGeoL1(this)">' + l1 + '</span>';
  }).join('');
}

function toggleGeoL1(el) {
  var l1 = el.dataset.l1;
  if (geoSelectedL1.has(l1)) {
    geoSelectedL1.delete(l1);
    el.classList.remove('active');
  } else {
    geoSelectedL1.add(l1);
    el.classList.add('active');
  }
  renderGeoAnalysis();
}

function clearGeoFilters() {
  var l1Set = new Set();
  allResults.forEach(function(r) {
    if (r.l1 && r.l1 !== '待确认') {
      l1Set.add(r.l1);
    }
  });
  geoSelectedL1 = new Set(l1Set);
  document.querySelectorAll('#geo-filter-l1 .filter-tag').forEach(function(el) {
    el.classList.add('active');
  });
  renderGeoAnalysis();
}

function calculateGeoMetrics() {
  var metric = document.getElementById('geo-metric-select').value;
  var data = {
    country: {},   // 国家分布
    region: {},    // 区域分布
    tier: {},      // 城市等级分布
    city: {}       // 城市TOP20
  };

  var CITY_TIER_MAP = KC.DictEntity.CITY_TIER_MAP;
  var CITY_REGION_MAP = KC.DictEntity.CITY_REGION_MAP;

  // 只分析有明确标签的数据
  allResults.forEach(function(r) {
    if (!r.l1 || r.l1 === '待确认') return;
    if (!geoSelectedL1.has(r.l1)) return;

    var value = r[metric] || 0;

    // 1. 国家分布（跳过无国家信息的数据）
    var country = r.country;
    if (!country || country === '') return; // 跳过未识别
    if (!data.country[country]) data.country[country] = 0;
    data.country[country] += value;

    // 2. 区域分布（跳过未识别的数据）
    var region = null;
    if (r.country !== '中国' && r.city) {
      region = '海外';
    } else if (r.city && CITY_REGION_MAP[r.city]) {
      region = CITY_REGION_MAP[r.city];
    }
    // 无法识别区域的数据跳过
    if (region !== null) {
      if (!data.region[region]) data.region[region] = 0;
      data.region[region] += value;
    }

    // 3. 城市等级分布（跳过未识别的数据）
    var tier = null;
    if (r.country !== '中国' && r.city) {
      tier = '海外城市';
    } else if (r.city && CITY_TIER_MAP[r.city]) {
      tier = CITY_TIER_MAP[r.city];
    }
    // 无法识别等级的数据跳过
    if (tier !== null) {
      if (!data.tier[tier]) data.tier[tier] = 0;
      data.tier[tier] += value;
    }

    // 4. 城市TOP20（有city且非空字符串）
    if (r.city && r.city !== '') {
      if (!data.city[r.city]) data.city[r.city] = 0;
      data.city[r.city] += value;
    }
  });

  return data;
}

function renderGeoAnalysis() {
  var data = calculateGeoMetrics();
  renderGeoCountryChart(data.country);
  renderGeoRegionChart(data.region);
  renderGeoTierChart(data.tier);
  renderGeoCityChart(data.city);
}

function renderGeoCountryChart(countryData) {
  if (geoCharts.country) geoCharts.country.destroy();
  var sorted = Object.entries(countryData).sort(function(a, b) { return b[1] - a[1]; });
  var labels = sorted.map(function(item) { return item[0]; });
  var values = sorted.map(function(item) { return item[1]; });

  geoCharts.country = new Chart(document.getElementById('chart-geo-country'), {
    type: 'pie',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: [
          '#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6',
          '#ec4899','#06b6d4','#f97316','#84cc16','#6366f1',
          '#14b8a6','#e11d48','#a855f7','#0ea5e9','#d946ef'
        ]
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12, padding: 8 } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              var total = ctx.dataset.data.reduce(function(s, v) { return s + v; }, 0);
              var pct = total > 0 ? (ctx.parsed / total * 100).toFixed(1) : 0;
              return ' ' + ctx.label + ': ' + pct + '%';
            }
          }
        }
      }
    }
  });
}

function renderGeoRegionChart(regionData) {
  if (geoCharts.region) geoCharts.region.destroy();
  // 按固定顺序展示
  var labels = [];
  var values = [];
  GEO_REGION_ORDER.forEach(function(region) {
    if (regionData[region] !== undefined) {
      labels.push(region);
      values.push(regionData[region]);
    }
  });

  geoCharts.region = new Chart(document.getElementById('chart-geo-region'), {
    type: 'pie',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: [
          '#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6',
          '#ec4899','#06b6d4','#f97316'
        ]
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12, padding: 8 } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              var total = ctx.dataset.data.reduce(function(s, v) { return s + v; }, 0);
              var pct = total > 0 ? (ctx.parsed / total * 100).toFixed(1) : 0;
              return ' ' + ctx.label + ': ' + pct + '%';
            }
          }
        }
      }
    }
  });
}

function renderGeoTierChart(tierData) {
  if (geoCharts.tier) geoCharts.tier.destroy();
  // 按固定顺序展示
  var labels = [];
  var values = [];
  GEO_TIER_ORDER.forEach(function(tier) {
    if (tierData[tier] !== undefined) {
      labels.push(tier);
      values.push(tierData[tier]);
    }
  });

  geoCharts.tier = new Chart(document.getElementById('chart-geo-tier'), {
    type: 'pie',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: [
          '#3b82f6','#ef4444','#10b981','#f59e0b','#8b5cf6',
          '#ec4899','#06b6d4'
        ]
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, boxWidth: 12, padding: 8 } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              var total = ctx.dataset.data.reduce(function(s, v) { return s + v; }, 0);
              var pct = total > 0 ? (ctx.parsed / total * 100).toFixed(1) : 0;
              return ' ' + ctx.label + ': ' + pct + '%';
            }
          }
        }
      }
    }
  });
}

function renderGeoCityChart(cityData) {
  if (geoCharts.city) geoCharts.city.destroy();
  // 取Top50按数值降序
  var sorted = Object.entries(cityData).sort(function(a, b) { return b[1] - a[1]; }).slice(0, 50);
  var total = sorted.reduce(function(sum, item) { return sum + item[1]; }, 0);
  var labels = sorted.map(function(item) { return item[0]; });
  var values = sorted.map(function(item) { return total > 0 ? (item[1] / total * 100) : 0; });

  geoCharts.city = new Chart(document.getElementById('chart-geo-city'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: '占比',
        data: values,
        backgroundColor: '#3b82f6',
        borderRadius: 4,
        barThickness: 10
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) { return ' ' + ctx.parsed.y.toFixed(1) + '%'; }
          }
        }
      },
      scales: {
        y: { beginAtZero: true, ticks: { callback: function(v) { return v.toFixed(1) + '%'; } } },
        x: { ticks: { font: { size: 9 }, maxRotation: 60, minRotation: 45 } }
      }
    }
  });
}

// ── 文件拖拽 ──
const dropzone = document.getElementById('dropzone');
dropzone.addEventListener('dragover', e => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', e => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  const files = e.dataTransfer.files;
  if (files.length > 0) {
    document.getElementById('file-input').files = files;
    // 手动触发 change 显示反馈
    document.getElementById('file-input').dispatchEvent(new Event('change'));
  }
});
document.getElementById('file-input').addEventListener('change', function() {
  const f = this.files[0];
  if (f) {
    const sizeStr = f.size > 1024*1024 ? (f.size/1024/1024).toFixed(1) + ' MB' : (f.size/1024).toFixed(0) + ' KB';
    // dropzone 内部显示文件名
    const dropzone = document.getElementById('dropzone');
    dropzone.classList.add('has-file');
    document.getElementById('filename-display').textContent = '✅ ' + f.name + '（' + sizeStr + '）';
    // 操作栏显示醒目标记
    document.getElementById('file-info').innerHTML = '<span class="file-info-badge">📁 ' + f.name + '（' + sizeStr + '）</span>';
  }
});

// ═══════════════════════════════════════════════════
// 交叉分析模块：业务线 × 词包（可筛选版）
// ═══════════════════════════════════════════════════

let crossCharts = {};
let crossSelectedL1 = new Set();
let crossSelectedL2 = new Set();

function initCrossFilters() {
  // 获取所有唯一的l1和l2
  const l1Set = new Set(allResults.map(r => r.l1).filter(Boolean));
  const l2Set = new Set(allResults.map(r => r.l2).filter(Boolean));
  
  // 默认全选
  crossSelectedL1 = new Set(l1Set);
  crossSelectedL2 = new Set(l2Set);
  
  // 渲染业务线筛选器
  const l1Container = document.getElementById('cross-filter-l1');
  l1Container.innerHTML = [...l1Set].sort().map(l1 => 
    `<span class="filter-tag active" data-l1="${l1}" onclick="toggleCrossL1(this)">${l1}</span>`
  ).join('');
  
  // 渲染词包筛选器
  const l2Container = document.getElementById('cross-filter-l2');
  l2Container.innerHTML = [...l2Set].sort().map(l2 => 
    `<span class="filter-tag active" data-l2="${l2}" onclick="toggleCrossL2(this)">${l2}</span>`
  ).join('');
}

function toggleCrossL1(el) {
  const l1 = el.dataset.l1;
  if (crossSelectedL1.has(l1)) {
    crossSelectedL1.delete(l1);
    el.classList.remove('active');
  } else {
    crossSelectedL1.add(l1);
    el.classList.add('active');
  }
  // 业务线变化后，联动更新词包标签（只显示选中业务线下存在的词包）
  updateCrossL2Tags();
  renderCrossAnalysis();
}

// 根据当前选中的业务线，动态更新词包筛选标签
function updateCrossL2Tags() {
  // 计算选中业务线下实际存在的词包集合
  let availableL2s;
  if (crossSelectedL1.size === 0) {
    availableL2s = new Set();
  } else {
    availableL2s = new Set(
      allResults
        .filter(r => r.l1 && crossSelectedL1.has(r.l1) && r.l2)
        .map(r => r.l2)
    );
  }

  const l2Container = document.getElementById('cross-filter-l2');

  // 如果选中了全部业务线（等于初始全量），直接显示全部词包
  const allL1Set = new Set(allResults.map(r => r.l1).filter(Boolean));
  const isAllSelected = [...allL1Set].every(l1 => crossSelectedL1.has(l1));

  if (isAllSelected) {
    availableL2s = new Set(allResults.map(r => r.l2).filter(Boolean));
  }

  // 重新渲染词包标签，默认全选可用词包
  crossSelectedL2 = new Set(availableL2s);
  l2Container.innerHTML = [...availableL2s].sort().map(l2 =>
    `<span class="filter-tag active" data-l2="${l2}" onclick="toggleCrossL2(this)">${l2}</span>`
  ).join('') || '<span style="color:#9ca3af;font-size:12px;">（当前业务线下无词包）</span>';
}

function toggleCrossL2(el) {
  const l2 = el.dataset.l2;
  if (crossSelectedL2.has(l2)) {
    crossSelectedL2.delete(l2);
    el.classList.remove('active');
  } else {
    crossSelectedL2.add(l2);
    el.classList.add('active');
  }
  renderCrossAnalysis();
}

function clearCrossFilters() {
  crossSelectedL1 = new Set(allResults.map(r => r.l1).filter(Boolean));
  crossSelectedL2 = new Set(allResults.map(r => r.l2).filter(Boolean));
  
  document.querySelectorAll('#cross-filter-l1 .filter-tag').forEach(el => {
    el.classList.add('active');
  });

  // 清除时恢复全量词包标签
  const l2Container = document.getElementById('cross-filter-l2');
  l2Container.innerHTML = [...crossSelectedL2].sort().map(l2 =>
    `<span class="filter-tag active" data-l2="${l2}" onclick="toggleCrossL2(this)">${l2}</span>`
  ).join('');
  
  renderCrossAnalysis();
}

function calculateCrossMetrics() {
  const metrics = {};
  
  // 只统计选中的业务线和词包
  allResults.forEach(r => {
    const l1 = r.l1 || '未分类';
    const l2 = r.l2 || '通用词';
    
    if (!crossSelectedL1.has(l1) || !crossSelectedL2.has(l2)) return;
    
    const key = l1 + '|||' + l2;
    if (!metrics[key]) {
      metrics[key] = { l1, l2, count: 0, cost: 0, impression: 0, click: 0 };
    }
    metrics[key].count++;
    metrics[key].cost += r.cost || 0;
    metrics[key].impression += r.impression || 0;
    metrics[key].click += r.click || 0;
  });
  
  // 计算衍生指标
  Object.values(metrics).forEach(m => {
    m.ctr = m.impression > 0 ? (m.click / m.impression * 100) : 0;
    m.cpc = m.click > 0 ? (m.cost / m.click) : 0;
    m.efficiency = m.cpc > 0 ? (m.ctr / m.cpc * 100) : 0;
  });
  
  return metrics;
}

function renderCrossAnalysis() {
  const metrics = calculateCrossMetrics();
  const metricField = document.getElementById('cross-metric-select').value;
  
  // 获取所有唯一的l1和l2（只包含选中的）
  const l1s = [...crossSelectedL1].sort();
  const l2s = [...crossSelectedL2].sort();
  
  if (l1s.length === 0 || l2s.length === 0) {
    document.getElementById('cross-summary-cards').innerHTML = '<p style="color:#9ca3af;text-align:center;">请选择至少一个业务线和词包</p>';
    return;
  }
  
  // 构建矩阵
  const matrix = {};
  l1s.forEach(l1 => {
    matrix[l1] = {};
    l2s.forEach(l2 => {
      const key = l1 + '|||' + l2;
      matrix[l1][l2] = metrics[key] || { l1, l2, count: 0, cost: 0, impression: 0, click: 0, ctr: 0, cpc: 0, efficiency: 0 };
    });
  });
  
  // 渲染概览卡片
  renderCrossSummaryCards(metrics);
  
  // 渲染图表
  renderCrossCharts(l1s, l2s, matrix, metricField);
  
  // 渲染热力图
  renderCrossHeatmap(l1s, l2s, matrix, metricField);
  
  // 渲染详细表格
  renderCrossDetailTable(metrics);
}

function renderCrossSummaryCards(metrics) {
  const values = Object.values(metrics);
  const totalCount = values.reduce((sum, m) => sum + m.count, 0);
  const totalCost = values.reduce((sum, m) => sum + m.cost, 0);
  const totalImp = values.reduce((sum, m) => sum + m.impression, 0);
  const totalClick = values.reduce((sum, m) => sum + m.click, 0);
  const avgCtr = totalImp > 0 ? (totalClick / totalImp * 100) : 0;
  const avgCpc = totalClick > 0 ? (totalCost / totalClick) : 0;
  
  const cards = [
    { label: '关键词数', value: totalCount.toLocaleString(), color: '#1a73e8' },
    { label: '总花费', value: '¥' + totalCost.toLocaleString(), color: '#ef4444' },
    { label: '总曝光', value: totalImp.toLocaleString(), color: '#3b82f6' },
    { label: '总点击', value: totalClick.toLocaleString(), color: '#10b981' },
    { label: '平均CTR', value: avgCtr.toFixed(2) + '%', color: '#f59e0b' },
    { label: '平均CPC', value: '¥' + avgCpc.toFixed(2), color: '#8b5cf6' },
  ];
  
  document.getElementById('cross-summary-cards').innerHTML = cards.map(c => `
    <div style="background:linear-gradient(135deg,${c.color}15,${c.color}08);border-radius:8px;padding:12px;text-align:center;border:1px solid ${c.color}30;transition:all .2s;"
      onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 4px 12px ${c.color}20'"
      onmouseout="this.style.transform='';this.style.boxShadow=''"
    >
      <div style="font-size:20px;font-weight:700;color:${c.color};margin-bottom:4px;">${c.value}</div>
      <div style="font-size:11px;color:#6b7280;">${c.label}</div>
    </div>
  `).join('');
}

function renderCrossCharts(l1s, l2s, matrix, metricField) {
  // 计算合计并排序
  const l1Totals = {};
  l1s.forEach(l1 => {
    let sum = 0;
    l2s.forEach(l2 => sum += matrix[l1][l2][metricField] || 0);
    l1Totals[l1] = sum;
  });
  
  const l2Totals = {};
  l2s.forEach(l2 => {
    let sum = 0;
    l1s.forEach(l1 => sum += matrix[l1][l2][metricField] || 0);
    l2Totals[l2] = sum;
  });
  
  // 按合计值降序排序
  const sortedL1s = [...l1s].sort((a, b) => l1Totals[b] - l1Totals[a]);
  const sortedL2s = [...l2s].sort((a, b) => l2Totals[b] - l2Totals[a]);
  
  // 业务线维度图表（已排序）
  const l1Data = sortedL1s.map(l1 => l1Totals[l1]);
  const l1Colors = sortedL1s.map(l1 => L1_META[l1]?.color || '#999');
  
  if (crossCharts.l1) crossCharts.l1.destroy();
  crossCharts.l1 = new Chart(document.getElementById('chart-cross-l1'), {
    type: 'bar',
    data: {
      labels: sortedL1s,
      datasets: [{
        label: getMetricLabel(metricField),
        data: l1Data,
        backgroundColor: l1Colors,
        borderRadius: 5
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { 
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = l1Data.reduce((a, b) => a + b, 0);
              const pct = total > 0 ? (ctx.parsed.y / total * 100).toFixed(1) : 0;
              return ` ${ctx.parsed.y.toLocaleString()} (${pct}%)`;
            }
          }
        }
      },
      scales: { y: { beginAtZero: true, ticks: { callback: v => formatMetricValue(v, metricField) } } }
    }
  });
  
  // 词包维度图表（已排序）
  const l2Data = sortedL2s.map(l2 => l2Totals[l2]);
  const totalL2 = l2Data.reduce((a, b) => a + b, 0);
  
  if (crossCharts.l2) crossCharts.l2.destroy();
  crossCharts.l2 = new Chart(document.getElementById('chart-cross-l2'), {
    type: 'doughnut',
    data: {
      labels: sortedL2s,
      datasets: [{
        data: l2Data,
        backgroundColor: sortedL2s.map((_, i) => L2_COLORS[i % L2_COLORS.length]),
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, padding: 10, usePointStyle: true } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const pct = totalL2 > 0 ? (ctx.parsed / totalL2 * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${ctx.parsed.toLocaleString()} (${pct}%)`;
            }
          }
        }
      }
    }
  });
}

function renderCrossHeatmap(l1s, l2s, matrix, metricField) {
  // 计算每个业务线的合计值，用于排序
  const l1Totals = {};
  l1s.forEach(l1 => {
    let sum = 0;
    l2s.forEach(l2 => {
      sum += matrix[l1][l2][metricField] || 0;
    });
    l1Totals[l1] = sum;
  });
  
  // 计算每个词包的合计值，用于排序
  const l2Totals = {};
  l2s.forEach(l2 => {
    let sum = 0;
    l1s.forEach(l1 => {
      sum += matrix[l1][l2][metricField] || 0;
    });
    l2Totals[l2] = sum;
  });
  
  // 按合计值降序排序
  const sortedL1s = [...l1s].sort((a, b) => l2Totals[b] - l2Totals[a]);
  const sortedL2s = [...l2s].sort((a, b) => l2Totals[b] - l2Totals[a]);
  
  // 计算所有数据的最大值用于颜色映射
  let maxValue = 0;
  sortedL1s.forEach(l1 => {
    sortedL2s.forEach(l2 => {
      maxValue = Math.max(maxValue, matrix[l1][l2][metricField] || 0);
    });
  });
  
  // 计算总计（用于百分比计算）
  let grandTotal = 0;
  sortedL1s.forEach(l1 => {
    sortedL2s.forEach(l2 => {
      grandTotal += matrix[l1][l2][metricField] || 0;
    });
  });
  
  // 生成热力图HTML
  let html = '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
  
  // 表头
  html += '<thead><tr><th style="background:#f9fafb;padding:8px;border:1px solid #e5e7eb;">业务线 \\ 词包</th>';
  sortedL2s.forEach(l2 => {
    const colTotal = l2Totals[l2] || 0;
    const colPct = grandTotal > 0 ? (colTotal / grandTotal * 100).toFixed(1) : 0;
    html += `<th style="background:#f9fafb;padding:8px;border:1px solid #e5e7eb;text-align:center;font-size:11px;">
      <div>${l2}</div>
      <div style="font-size:10px;color:#6b7280;font-weight:normal;">${formatMetricValue(colTotal, metricField)}</div>
      <div style="font-size:10px;color:#1a73e8;font-weight:normal;">${colPct}%</div>
    </th>`;
  });
  html += `<th style="background:#f9fafb;padding:8px;border:1px solid #e5e7eb;text-align:center;font-size:11px;">
    <div>合计</div>
    <div style="font-size:10px;color:#6b7280;font-weight:normal;">${formatMetricValue(grandTotal, metricField)}</div>
  </th></tr></thead>`;
  
  // 表体
  html += '<tbody>';
  sortedL1s.forEach(l1 => {
    const rowTotal = l1Totals[l1] || 0;
    const rowPct = grandTotal > 0 ? (rowTotal / grandTotal * 100).toFixed(1) : 0;
    
    html += `<tr><td style="padding:8px;border:1px solid #e5e7eb;font-weight:600;background:#f9fafb;">
      <div>${l1}</div>
      <div style="font-size:10px;color:#6b7280;font-weight:normal;">${formatMetricValue(rowTotal, metricField)}</div>
      <div style="font-size:10px;color:#1a73e8;font-weight:normal;">${rowPct}%</div>
    </td>`;
    
    sortedL2s.forEach(l2 => {
      const value = matrix[l1][l2][metricField] || 0;
      
      // 计算颜色强度
      const intensity = maxValue > 0 ? (value / maxValue) : 0;
      const color = getHeatmapColor(intensity, metricField);
      
      // 计算百分比（占总计）
      const pct = grandTotal > 0 ? (value / grandTotal * 100).toFixed(1) : 0;
      const rowPct = rowTotal > 0 ? (value / rowTotal * 100).toFixed(1) : 0;
      
      html += `<td style="padding:8px;border:1px solid #e5e7eb;text-align:center;background:${color};cursor:pointer;transition:all .15s;"
        onmouseover="this.style.transform='scale(1.05)';this.style.zIndex='10';this.style.boxShadow='0 2px 8px rgba(0,0,0,0.15)';this.style.position='relative'"
        onmouseout="this.style.transform='';this.style.zIndex='';this.style.boxShadow=''"
        title="${l1} × ${l2}: ${formatMetricValue(value, metricField)} (${pct}% of total)"
      >
        <div style="font-weight:600;font-size:13px;">${formatMetricValue(value, metricField)}</div>
        <div style="font-size:10px;color:#374151;margin-top:2px;">${pct}%</div>
        <div style="font-size:9px;color:#6b7280;">行${rowPct}%</div>
      </td>`;
    });
    
    html += `<td style="padding:8px;border:1px solid #e5e7eb;text-align:center;font-weight:600;background:#f3f4f6;">
      <div>${formatMetricValue(rowTotal, metricField)}</div>
      <div style="font-size:10px;color:#1a73e8;">${rowPct}%</div>
    </td></tr>`;
  });
  
  html += '</tbody></table>';
  
  document.getElementById('cross-heatmap').innerHTML = html;
}

let crossSortField = 'cost';
let crossSortDesc = true;

function sortCrossTable(field) {
  if (crossSortField === field) {
    crossSortDesc = !crossSortDesc;
  } else {
    crossSortField = field;
    crossSortDesc = true;
  }
  renderCrossAnalysis();
}

function renderCrossDetailTable(metrics) {
  const values = Object.values(metrics);
  
  // 排序
  values.sort((a, b) => {
    const aVal = a[crossSortField] || 0;
    const bVal = b[crossSortField] || 0;
    return crossSortDesc ? bVal - aVal : aVal - bVal;
  });
  
  // 计算总计用于百分比
  const totalCost = values.reduce((sum, m) => sum + m.cost, 0);
  const totalImp = values.reduce((sum, m) => sum + m.impression, 0);
  const totalClick = values.reduce((sum, m) => sum + m.click, 0);
  
  // 渲染动态表头（带排序指示器）
  const headers = [
    { field: null, label: '业务线' },
    { field: null, label: '词包' },
    { field: 'count', label: '关键词数' },
    { field: 'cost', label: '花费' },
    { field: 'impression', label: '曝光' },
    { field: 'click', label: '点击' },
    { field: 'ctr', label: 'CTR' },
    { field: 'cpc', label: 'CPC' },
    { field: 'efficiency', label: '效率评分' }
  ];
  
  const theadHtml = headers.map(h => {
    if (!h.field) return `<th>${h.label}</th>`;
    const arrow = crossSortField === h.field ? (crossSortDesc ? ' ▼' : ' ▲') : '';
    return `<th style="cursor:pointer;" onclick="sortCrossTable('${h.field}')" onmouseover="this.style.background='#e5e7eb'" onmouseout="this.style.background=''">${h.label}${arrow}</th>`;
  }).join('');
  document.getElementById('cross-detail-header').innerHTML = theadHtml;
  
  const tbody = document.querySelector('#table-cross-detail tbody');
  tbody.innerHTML = values.map(m => {
    const costPct = totalCost > 0 ? (m.cost / totalCost * 100).toFixed(1) : 0;
    const impPct = totalImp > 0 ? (m.impression / totalImp * 100).toFixed(1) : 0;
    const clickPct = totalClick > 0 ? (m.click / totalClick * 100).toFixed(1) : 0;
    
    return `
    <tr>
      <td><span class="l1-badge ${L1_META[m.l1]?.cls || ''}">${m.l1}</span></td>
      <td><span class="l2-tag">${m.l2}</span></td>
      <td class="num">${m.count.toLocaleString()}</td>
      <td class="num" style="color:#ef4444">
        ¥${m.cost.toLocaleString()}
        <div style="font-size:10px;color:#6b7280;">${costPct}%</div>
      </td>
      <td class="num">
        ${m.impression.toLocaleString()}
        <div style="font-size:10px;color:#6b7280;">${impPct}%</div>
      </td>
      <td class="num">
        ${m.click.toLocaleString()}
        <div style="font-size:10px;color:#6b7280;">${clickPct}%</div>
      </td>
      <td class="num">${m.ctr.toFixed(2)}%</td>
      <td class="num">¥${m.cpc.toFixed(2)}</td>
      <td class="num" style="color:${m.efficiency >= 10 ? '#10b981' : m.efficiency >= 5 ? '#f59e0b' : '#ef4444'};font-weight:600">
        ${m.efficiency.toFixed(1)}
      </td>
    </tr>`;
  }).join('');
  
  // 添加合计行
  const totalRow = `
    <tr style="background:#f3f4f6;font-weight:600;">
      <td colspan="2">合计</td>
      <td class="num">${values.reduce((sum, m) => sum + m.count, 0).toLocaleString()}</td>
      <td class="num">¥${totalCost.toLocaleString()}</td>
      <td class="num">${totalImp.toLocaleString()}</td>
      <td class="num">${totalClick.toLocaleString()}</td>
      <td class="num">${totalImp > 0 ? (totalClick/totalImp*100).toFixed(2) : 0}%</td>
      <td class="num">¥${totalClick > 0 ? (totalCost/totalClick).toFixed(2) : 0}</td>
      <td></td>
    </tr>
  `;
  tbody.innerHTML += totalRow;
}

function getHeatmapColor(intensity, metricField) {
  // 根据指标类型选择颜色主题
  if (metricField === 'efficiency') {
    // 效率：绿->黄->红
    if (intensity >= 0.7) return `rgba(16, 185, 129, ${0.2 + intensity * 0.5})`;
    if (intensity >= 0.3) return `rgba(245, 158, 11, ${0.2 + intensity * 0.5})`;
    return `rgba(239, 68, 68, ${0.2 + intensity * 0.5})`;
  } else if (metricField === 'ctr') {
    // CTR：蓝->绿
    return `rgba(59, 130, 246, ${0.15 + intensity * 0.55})`;
  } else if (metricField === 'cpc') {
    // CPC：红（高CPC=红色）
    return `rgba(239, 68, 68, ${0.15 + intensity * 0.55})`;
  } else {
    // 默认：蓝色渐变
    return `rgba(26, 115, 232, ${0.15 + intensity * 0.55})`;
  }
}

function getMetricLabel(metricField) {
  const labels = {
    count: '关键词数量',
    cost: '花费',
    impression: '曝光量',
    click: '点击量',
    ctr: 'CTR',
    cpc: 'CPC',
    efficiency: '效率评分'
  };
  return labels[metricField] || metricField;
}

function formatMetricValue(value, metricField) {
  if (value === 0 || value === undefined || value === null) return '-';
  
  switch (metricField) {
    case 'cost':
    case 'cpc':
      return '¥' + value.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    case 'ctr':
      return value.toFixed(2) + '%';
    case 'efficiency':
      return value.toFixed(1);
    default:
      return value.toLocaleString();
  }
}

;

// ═══════════════════════════════════════════════════
// 关键词词云模块
// ═══════════════════════════════════════════════════

let wcSelectedL1 = new Set();
let wcSelectedL2 = new Set();
let wcWordcloudInstance = null;

function initWcFilters() {
  const l1Set = new Set(allResults.map(r => r.l1).filter(Boolean));
  const l2Set = new Set(allResults.map(r => r.l2).filter(Boolean));
  wcSelectedL1 = new Set(l1Set);
  wcSelectedL2 = new Set(l2Set);

  const l1Container = document.getElementById('wc-filter-l1');
  l1Container.innerHTML = [...l1Set].sort().map(l1 =>
    `<span class="filter-tag active" data-l1="${l1}" onclick="toggleWcL1(this)">${l1}</span>`
  ).join('');

  updateWcL2Tags();
}

function toggleWcL1(el) {
  const l1 = el.dataset.l1;
  if (wcSelectedL1.has(l1)) {
    wcSelectedL1.delete(l1);
    el.classList.remove('active');
  } else {
    wcSelectedL1.add(l1);
    el.classList.add('active');
  }
  updateWcL2Tags();
  renderWordCloud();
}

function updateWcL2Tags() {
  let availableL2s;
  if (wcSelectedL1.size === 0) {
    availableL2s = new Set();
  } else {
    availableL2s = new Set(
      allResults
        .filter(r => r.l1 && wcSelectedL1.has(r.l1) && r.l2)
        .map(r => r.l2)
    );
  }
  const allL1Set = new Set(allResults.map(r => r.l1).filter(Boolean));
  const isAllSelected = [...allL1Set].every(l1 => wcSelectedL1.has(l1));
  if (isAllSelected) {
    availableL2s = new Set(allResults.map(r => r.l2).filter(Boolean));
  }
  wcSelectedL2 = new Set(availableL2s);
  const l2Container = document.getElementById('wc-filter-l2');
  l2Container.innerHTML = [...availableL2s].sort().map(l2 =>
    `<span class="filter-tag active" data-l2="${l2}" onclick="toggleWcL2(this)">${l2}</span>`
  ).join('') || '<span style="color:#9ca3af;font-size:12px;">（当前业务线下无词包）</span>';
}

function toggleWcL2(el) {
  const l2 = el.dataset.l2;
  if (wcSelectedL2.has(l2)) {
    wcSelectedL2.delete(l2);
    el.classList.remove('active');
  } else {
    wcSelectedL2.add(l2);
    el.classList.add('active');
  }
  renderWordCloud();
}

function clearWcFilters() {
  wcSelectedL1 = new Set(allResults.map(r => r.l1).filter(Boolean));
  wcSelectedL2 = new Set(allResults.map(r => r.l2).filter(Boolean));
  document.querySelectorAll('#wc-filter-l1 .filter-tag').forEach(el => el.classList.add('active'));
  updateWcL2Tags();
  renderWordCloud();
}

function renderWordCloud() {
  const canvas = document.getElementById('wordcloud-canvas');
  const wrap = document.getElementById('wordcloud-wrap');
  const tooltip = document.getElementById('wordcloud-tooltip');
  const metric = document.getElementById('wc-metric-select').value;
  const metricLabel = { cost: '消耗', impression: '展示', click: '点击' }[metric];

  // 根据筛选条件过滤数据
  const filtered = allResults.filter(r =>
    wcSelectedL1.has(r.l1) && wcSelectedL2.has(r.l2)
  );

  // 高清 canvas（处理 Retina 屏幕）
  const dpr = window.devicePixelRatio || 1;
  const width = wrap.offsetWidth;
  const height = wrap.offsetHeight;
  canvas.width = width * dpr;
  canvas.height = height * dpr;
  canvas.style.width = width + 'px';
  canvas.style.height = height + 'px';

  if (filtered.length === 0) {
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = '#9ca3af';
    ctx.font = '14px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('无匹配关键词', width / 2, height / 2);
    return;
  }

  // 判断是否有对应指标数据
  const hasMetricData = filtered.some(r => r[metric] > 0);

  // 构建词列表：[text, weight, l1, l2, kw]
  const wordMap = {};
  filtered.forEach(r => {
    const kw = r.kw.trim();
    if (!kw) return;
    if (!wordMap[kw]) {
      wordMap[kw] = { text: kw, weight: 0, l1: r.l1, l2: r.l2 };
    }
    wordMap[kw].weight += hasMetricData ? (r[metric] || 1) : 1;
  });

  const wordList = Object.values(wordMap);
  const maxWeight = Math.max(...wordList.map(w => w.weight));
  const minWeight = Math.min(...wordList.map(w => w.weight));
  const weightRange = maxWeight - minWeight || 1;

  // 构建 wordcloud2 所需格式 [[word, weight], ...]
  const wcData = wordList.map(w => {
    const normalized = (w.weight - minWeight) / weightRange;
    const fontSize = Math.round(14 + normalized * 50);
    const color = L1_META[w.l1] ? L1_META[w.l1].color : '#6b7280';
    return [w.text, fontSize, color, w.l1, w.l2, w.weight];
  });

  // 隐藏 tooltip
  tooltip.style.display = 'none';

  // 渲染词云
  try {
    WordCloud(canvas, {
      list: wcData.map(item => [item[0], item[1]]),
      gridSize: Math.round(10 * dpr * width / 600),
      weightFactor: function(size) { return size * dpr; },
      fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
      fontWeight: 600,
      color: function(word, weight) {
        const item = wcData.find(d => d[0] === word);
        return item ? item[2] : '#6b7280';
      },
      backgroundColor: 'transparent',
      rotateRatio: 0.3,
      rotationSteps: 2,
      minRotation: -Math.PI / 2,
      maxRotation: Math.PI / 2,
      drawOutOfBound: false,
      shrinkToFit: true,
      shuffle: true,
      hover: function(item, dimension, event) {
        if (item) {
          const wcItem = wcData.find(d => d[0] === item[0]);
          if (wcItem) {
            tooltip.innerHTML = `<b>${wcItem[0]}</b><br><span style="opacity:.7">${wcItem[3]} · ${wcItem[4]}</span><br>${metricLabel}: ${Number(wcItem[5]).toLocaleString()}`;
            tooltip.style.display = 'block';
            tooltip.style.left = (event.offsetX + 12) + 'px';
            tooltip.style.top = (event.offsetY - 50) + 'px';
          }
        } else {
          tooltip.style.display = 'none';
        }
      },
      click: function(item) {
        if (item) {
          const wcItem = wcData.find(d => d[0] === item[0]);
          if (wcItem) {
            document.getElementById('search-box').value = wcItem[0];
            renderTable(1);
            document.querySelector('.card.mb-4:last-child').scrollIntoView({ behavior: 'smooth' });
          }
        }
      }
    });
  } catch (e) {
    // wordcloud2 容错
    console.warn('词云渲染异常:', e);
  }
}

// ═══════════════════════════════════════════════════
// 词包生成器（v2：筛选 + 维度选择 + Jenks 最优分档）
// ═══════════════════════════════════════════════════

let pkgResults = [];
let pkgPage = 1;
const PKG_PAGE_SIZE = 30;

// 初始化词包筛选器（分类完成后调用）
function initPkgFilters() {
  // 业务线筛选
  const l1Set = [...new Set(allResults.map(r => r.l1).filter(Boolean))].sort();
  const l1Container = document.getElementById('pkg-filter-l1');
  l1Container.innerHTML = l1Set.map(l1 => {
    const meta = L1_META[l1] || { cls: '', color: '#999' };
    return `<span class="filter-tag active" data-l1="${l1}" data-color="${meta.color}" onclick="togglePkgFilter(this,'l1')">${l1}</span>`;
  }).join('') + `<span class="filter-tag" onclick="togglePkgFilterAll('l1')" style="font-size:11px;color:#6b7280;">全选/反选</span>`;

  // 意图词包筛选 — 初始化后由 updatePkgL2Filter 联动
  window._pkgSelectedL2 = new Set();
  updatePkgL2Filter();

  // 城市筛选 — 统一用可搜索下拉多选（由 updatePkgCityFilter 联动）
  window._pkgSelectedCities = new Set();
  updatePkgCityFilter();

  // 点击外部关闭下拉
  document.addEventListener('click', e => {
    document.querySelectorAll('.city-select-wrap').forEach(wrap => {
      if (!wrap.contains(e.target)) {
        wrap.querySelector('.city-dropdown')?.classList.remove('open');
      }
    });
  });

  // radio 样式联动
  document.querySelectorAll('input[name=pkg-dim]').forEach(radio => {
    radio.addEventListener('change', () => {
      document.querySelectorAll('input[name=pkg-dim]').forEach(r => {
        r.closest('label').style.borderColor = r.checked ? '#fbbf24' : '#d1d5db';
      });
    });
  });
}

// ── L2 意图词包联动 ──
function getAvailableL2s() {
  const selectedL1 = getPkgSelectedL1s();
  return [...new Set(
    allResults.filter(r => selectedL1.has(r.l1)).map(r => r.l2).filter(Boolean)
  )].sort();
}

function updatePkgL2Filter() {
  const l2Set = getAvailableL2s();
  const container = document.getElementById('pkg-filter-l2');
  // 全选新出现的 L2
  window._pkgSelectedL2 = new Set(l2Set);
  if (l2Set.length === 0) {
    container.innerHTML = '<span style="color:#9ca3af;font-size:12px;">无匹配意图词包</span>';
  } else {
    container.innerHTML = `
      <div class="city-select-wrap">
        <div class="city-select-btn" onclick="toggleL2Dropdown()">
          📋 选择意图词包 <span class="count" id="l2-count-badge">${l2Set.length}</span> <span style="font-size:10px;color:#9ca3af;">▼</span>
        </div>
        <div class="city-dropdown" id="l2-dropdown">
          <input type="text" class="city-search" id="l2-search-input" placeholder="搜索意图词包…" oninput="filterL2List()">
          <div class="city-actions">
            <a onclick="l2SelectAll()">全选</a>
            <a onclick="l2SelectNone()">全不选</a>
            <a onclick="l2SelectInvert()">反选</a>
          </div>
          <div class="city-list" id="l2-list"></div>
        </div>
      </div>`;
    renderL2List();
  }
}

function toggleL2Dropdown() {
  const dd = document.getElementById('l2-dropdown');
  dd.classList.toggle('open');
  if (dd.classList.contains('open')) {
    const input = document.getElementById('l2-search-input');
    input.value = '';
    renderL2List();
    setTimeout(() => input.focus(), 50);
  }
}

function renderL2List(filter) {
  const l2Set = getAvailableL2s();
  const selected = window._pkgSelectedL2 || new Set(l2Set);
  const keyword = (filter || '').toLowerCase();
  const list = document.getElementById('l2-list');
  const filtered = keyword ? l2Set.filter(c => c.toLowerCase().includes(keyword)) : l2Set;
  list.innerHTML = filtered.map(c =>
    `<label class="city-item"><input type="checkbox" ${selected.has(c) ? 'checked' : ''} onchange="toggleL2Item(this,'${c.replace(/'/g, "\\'")}')"> ${c}</label>`
  ).join('');
  if (filtered.length === 0) list.innerHTML = '<div style="padding:8px 12px;color:#9ca3af;font-size:12px;">无匹配意图词包</div>';
}

function filterL2List() {
  const keyword = document.getElementById('l2-search-input').value;
  renderL2List(keyword);
}

function toggleL2Item(cb, l2) {
  const selected = window._pkgSelectedL2;
  if (cb.checked) selected.add(l2); else selected.delete(l2);
  updateL2Badge();
}

function l2SelectAll() {
  getAvailableL2s().forEach(c => window._pkgSelectedL2.add(c));
  renderL2List(document.getElementById('l2-search-input')?.value);
  updateL2Badge();
}

function l2SelectNone() {
  window._pkgSelectedL2.clear();
  renderL2List(document.getElementById('l2-search-input')?.value);
  updateL2Badge();
}

function l2SelectInvert() {
  getAvailableL2s().forEach(c => {
    if (window._pkgSelectedL2.has(c)) window._pkgSelectedL2.delete(c);
    else window._pkgSelectedL2.add(c);
  });
  renderL2List(document.getElementById('l2-search-input')?.value);
  updateL2Badge();
}

function updateL2Badge() {
  const badge = document.getElementById('l2-count-badge');
  if (badge) badge.textContent = window._pkgSelectedL2.size;
}

// ── 城市联动 ──
function getAvailableCities() {
  const selectedL1 = getPkgSelectedL1s();
  return [...new Set(
    allResults.filter(r => selectedL1.has(r.l1)).map(r => r.city).filter(Boolean)
  )].sort();
}

function updatePkgCityFilter() {
  const citySet = getAvailableCities();
  const container = document.getElementById('pkg-filter-city');
  // 全选新出现的城市
  window._pkgSelectedCities = new Set(citySet);
  if (citySet.length === 0) {
    container.innerHTML = '<span style="color:#9ca3af;font-size:12px;">无匹配城市</span>';
  } else {
    container.innerHTML = `
      <div class="city-select-wrap">
        <div class="city-select-btn" onclick="toggleCityDropdown()">
          🏙 选择城市 <span class="count" id="city-count-badge">${citySet.length}</span> <span style="font-size:10px;color:#9ca3af;">▼</span>
        </div>
        <div class="city-dropdown" id="city-dropdown">
          <input type="text" class="city-search" id="city-search-input" placeholder="搜索城市…" oninput="filterCityList()">
          <div class="city-actions">
            <a onclick="citySelectAll()">全选</a>
            <a onclick="citySelectNone()">全不选</a>
            <a onclick="citySelectInvert()">反选</a>
          </div>
          <div class="city-list" id="city-list"></div>
        </div>
      </div>`;
    renderCityList();
  }
}

function toggleCityDropdown() {
  const dd = document.getElementById('city-dropdown');
  dd.classList.toggle('open');
  if (dd.classList.contains('open')) {
    const input = document.getElementById('city-search-input');
    input.value = '';
    renderCityList();
    setTimeout(() => input.focus(), 50);
  }
}

function renderCityList(filter) {
  const citySet = getAvailableCities();
  const selected = window._pkgSelectedCities || new Set(citySet);
  const keyword = (filter || '').toLowerCase();
  const list = document.getElementById('city-list');
  const filtered = keyword ? citySet.filter(c => c.toLowerCase().includes(keyword)) : citySet;
  list.innerHTML = filtered.map(c =>
    `<label class="city-item"><input type="checkbox" ${selected.has(c) ? 'checked' : ''} onchange="toggleCityItem(this,'${c.replace(/'/g, "\\'")}')"> ${c}</label>`
  ).join('');
  if (filtered.length === 0) list.innerHTML = '<div style="padding:8px 12px;color:#9ca3af;font-size:12px;">无匹配城市</div>';
}

function filterCityList() {
  const keyword = document.getElementById('city-search-input').value;
  renderCityList(keyword);
}

function toggleCityItem(cb, city) {
  const selected = window._pkgSelectedCities;
  if (cb.checked) selected.add(city); else selected.delete(city);
  updateCityBadge();
}

function citySelectAll() {
  getAvailableCities().forEach(c => window._pkgSelectedCities.add(c));
  renderCityList(document.getElementById('city-search-input')?.value);
  updateCityBadge();
}

function citySelectNone() {
  window._pkgSelectedCities.clear();
  renderCityList(document.getElementById('city-search-input')?.value);
  updateCityBadge();
}

function citySelectInvert() {
  getAvailableCities().forEach(c => {
    if (window._pkgSelectedCities.has(c)) window._pkgSelectedCities.delete(c);
    else window._pkgSelectedCities.add(c);
  });
  renderCityList(document.getElementById('city-search-input')?.value);
  updateCityBadge();
}

function updateCityBadge() {
  const badge = document.getElementById('city-count-badge');
  if (badge) badge.textContent = window._pkgSelectedCities.size;
}

function togglePkgFilter(el, type) {
  const wasActive = el.classList.contains('active');
  el.classList.toggle('active');
  // 业务线标签：active 时清除内联 color/border-color 让 CSS !important 生效；
  // 非 active 时恢复业务线自身颜色
  if (type === 'l1') {
    const meta = L1_META[el.dataset.l1] || { color: '#999' };
    if (wasActive) {
      el.style.borderColor = meta.color;
      el.style.color = meta.color;
    } else {
      el.style.borderColor = '';
      el.style.color = '';
    }
    // L1 变化时联动 L2 和城市
    updatePkgL2Filter();
    updatePkgCityFilter();
  }
}

function togglePkgFilterAll(type) {
  const container = document.getElementById('pkg-filter-l1');
  const tags = container.querySelectorAll('.filter-tag[data-l1]');
  const allActive = [...tags].every(t => t.classList.contains('active'));
  tags.forEach(t => {
    if (allActive) {
      t.classList.remove('active');
      const meta = L1_META[t.dataset.l1] || { color: '#999' };
      t.style.borderColor = meta.color;
      t.style.color = meta.color;
    } else {
      t.classList.add('active');
      t.style.borderColor = '';
      t.style.color = '';
    }
  });
  // L1 变化时联动 L2 和城市
  updatePkgL2Filter();
  updatePkgCityFilter();
}

function getPkgSelectedCities() {
  return window._pkgSelectedCities || new Set();
}

function getPkgSelectedL1s() {
  const tags = document.querySelectorAll('#pkg-filter-l1 .filter-tag[data-l1].active');
  return new Set([...tags].map(t => t.dataset.l1));
}

function getPkgSelectedL2s() {
  return window._pkgSelectedL2 || new Set();
}

// ── Jenks 自然断点法 ──
// 在一组排序数值中寻找 k 个断点，使组内方差最小、组间差异最大
function jenksBreaks(values, k) {
  // values: 已排序的数值数组（升序）
  const n = values.length;
  if (n <= k) {
    // 数据量不够分 k 档，直接每条一档
    return values.slice(0, n - 1);
  }

  // 动态规划求最优断点
  // mat1[i][j] = 前i个数据分成j组的组内方差之和最小值
  const mat1 = Array.from({ length: n + 1 }, () => Array(k + 1).fill(Infinity));
  const mat2 = Array.from({ length: n + 1 }, () => Array(k + 1).fill(0));

  mat1[0][0] = 0;

  // 前缀和，用于快速计算区间方差
  const prefixSum = [0];
  const prefixSq = [0];
  for (let i = 0; i < n; i++) {
    prefixSum.push(prefixSum[i] + values[i]);
    prefixSq.push(prefixSq[i] + values[i] * values[i]);
  }

  function variance(lo, hi) {
    // 区间 [lo, hi) 的方差 * n
    const count = hi - lo;
    if (count <= 1) return 0;
    const sum = prefixSum[hi] - prefixSum[lo];
    const sq = prefixSq[hi] - prefixSq[lo];
    return sq - sum * sum / count;
  }

  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= Math.min(i, k); j++) {
      for (let x = j - 1; x < i; x++) {
        const v = mat1[x][j - 1] + variance(x, i);
        if (v < mat1[i][j]) {
          mat1[i][j] = v;
          mat2[i][j] = x;
        }
      }
    }
  }

  // 回溯断点位置
  const breaks = [];
  let pos = n;
  for (let j = k; j >= 2; j--) {
    pos = mat2[pos][j];
    breaks.push(values[pos]);
  }
  return breaks.reverse();
}

// ── 主生成函数 ──
function generatePackages() {
  if (allResults.length === 0) { alert('请先进行关键词分类'); return; }

  const selectedL1 = getPkgSelectedL1s();
  const selectedL2 = getPkgSelectedL2s();
  const selectedCities = getPkgSelectedCities();
  const dim = document.querySelector('input[name=pkg-dim]:checked').value;
  const maxKw = parseInt(document.getElementById('pkg-max').value) || 1000;
  const minSize = parseInt(document.getElementById('pkg-min-size').value) || 3;
  const maxTiers = parseInt(document.getElementById('pkg-max-tiers').value) || 3;

  // 维度映射
  const dimLabel = { cost: '消耗', impression: '展示', click: '点击' }[dim];
  const dimKey = dim; // r.cost, r.impression, r.click

  // ① 按业务线 + 意图词包 + 城市筛选
  const hasCityFilter = selectedCities.size > 0 && selectedCities.size < new Set(allResults.map(r => r.city).filter(Boolean)).size;
  let filtered = allResults.filter(r => {
    if (!selectedL1.has(r.l1)) return false;
    if (r.l2 && !selectedL2.has(r.l2)) return false;
    // 城市筛选：如果用户做了城市选择（非全选），则只保留已选城市的关键词
    if (hasCityFilter) {
      if (!r.city || !selectedCities.has(r.city)) return false;
    }
    return true;
  });

  if (filtered.length === 0) { alert('筛选后无关键词，请调整筛选条件'); return; }

  // ①.5 关键词数量上限：超过时按维度值从高到低截取
  if (filtered.length > maxKw) {
    filtered.sort((a, b) => (b[dimKey] || 0) - (a[dimKey] || 0));
    filtered = filtered.slice(0, maxKw);
  }

  // ② 按 业务线 > 意图词包 > 城市 分组
  const groups = {};
  filtered.forEach(r => {
    const cityLabel = r.city || '未指定城市';
    const key = (r.l1 || '未分类') + '>' + (r.l2 || '通用词') + '>' + cityLabel;
    if (!groups[key]) groups[key] = { l1: r.l1, l2: r.l2, city: cityLabel, items: [] };
    groups[key].items.push(r);
  });

  // ③ 在每个分组内用 Jenks 自然断点法按维度分档
  const tierNames = ['头部','中上','中等','中下','尾部'];
  const tierClasses = ['high','vhigh','mid','low','vlow'];
  let packages = [];
  let totalVariance = 0;
  let totalItems = 0;

  Object.values(groups).forEach(group => {
    const items = group.items;
    const metricValues = items.map(r => r[dimKey] || 0).sort((a, b) => a - b);

    // 决定分档数：关键词少于 minSize*2 不分档，少于 minSize*3 分2档，依此类推
    let tiers = 1;
    for (let t = maxTiers; t >= 2; t--) {
      if (items.length >= minSize * t) { tiers = t; break; }
    }

    if (tiers <= 1) {
      // 不分档，整包输出
      const pkg = buildPackage(group.l1 + ' > ' + group.l2 + ' > ' + group.city, items, '不分档', dimLabel);
      packages.push(pkg);
      totalVariance += pkg.variance;
      totalItems += items.length;
    } else {
      // 用 Jenks 找断点
      const breaks = jenksBreaks(metricValues, tiers);

      // 将关键词分配到各档
      // 断点 breakPoints: [v1, v2, ...] 表示第一档 ≤ v1, 第二档 ≤ v2, ...
      const sortedItems = [...items].sort((a, b) => (a[dimKey] || 0) - (b[dimKey] || 0));
      const tierItems = Array.from({ length: tiers }, () => []);
      let breakIdx = 0;

      sortedItems.forEach(item => {
        const val = item[dimKey] || 0;
        // 找到所属档位
        while (breakIdx < breaks.length && val > breaks[breakIdx]) breakIdx++;
        const tier = Math.min(breakIdx, tiers - 1);
        tierItems[tier].push(item);
      });

      tierItems.forEach((tierGroup, t) => {
        if (tierGroup.length === 0) return;
        const label = tierNames[Math.round(t * (tierNames.length - 1) / Math.max(1, tiers - 1))];
        const name = group.l1 + ' > ' + group.l2 + ' > ' + group.city + ' > ' + label + dimLabel;
        const pkg = buildPackage(name, tierGroup, label, dimLabel);
        pkg.tierClass = tierClasses[Math.round(t * (tierClasses.length - 1) / Math.max(1, tiers - 1))];
        packages.push(pkg);
        totalVariance += pkg.variance;
        totalItems += tierGroup.length;
      });
    }
  });

  // ④ 按关键词数降序
  packages.sort((a, b) => b.items.length - a.items.length);

  pkgResults = packages;
  pkgPage = 1;

  // ⑤ 统计
  document.getElementById('pkg-stats').classList.remove('hidden');
  const totalKw = pkgResults.reduce((s, p) => s + p.items.length, 0);
  const avgVariance = totalItems > 0 ? totalVariance / totalItems : 0;
  // 均匀度：1 - 归一化方差（越小越均匀，转为百分比）
  const maxPossibleVar = calcMaxVariance(filtered, dimKey);
  const uniformity = maxPossibleVar > 0 ? Math.max(0, (1 - avgVariance / maxPossibleVar) * 100).toFixed(1) : '100.0';

  document.getElementById('pkg-stat-count').textContent = pkgResults.length;
  document.getElementById('pkg-stat-kw').textContent = totalKw.toLocaleString();
  document.getElementById('pkg-stat-avg').textContent = pkgResults.length > 0 ? Math.round(totalKw / pkgResults.length) : 0;
  document.getElementById('pkg-stat-cost').textContent = '¥' + pkgResults.reduce((s, p) => s + p.cost, 0).toLocaleString();
  document.getElementById('pkg-stat-optimality').textContent = uniformity + '%';

  document.getElementById('pkg-table-wrap').classList.remove('hidden');
  const filteredNote = filtered.length < allResults.length ? `（筛选后 ${filtered.length} 条）` : '';
  document.getElementById('pkg-hint').textContent = `已生成 ${pkgResults.length} 个词包（${dimLabel}维度），共 ${totalKw} 条关键词${filteredNote}`;
  document.getElementById('pkg-search').value = '';
  renderPkgTable(1);
}

function buildPackage(name, items, tierLabel, dimLabel) {
  const pkg = {
    name,
    tierLabel,
    dimLabel,
    tierClass: 'mid',
    items,
    cost: items.reduce((s, r) => s + (r.cost || 0), 0),
    impression: items.reduce((s, r) => s + (r.impression || 0), 0),
    click: items.reduce((s, r) => s + (r.click || 0), 0),
    variance: 0
  };
  // 计算档内离散度（CV = 标准差/均值）
  const dimKey = document.querySelector('input[name=pkg-dim]:checked').value;
  pkg.variance = calcVariance(items, dimKey);
  return pkg;
}

function calcVariance(items, dimKey) {
  if (items.length <= 1) return 0;
  const vals = items.map(r => r[dimKey] || 0);
  const mean = vals.reduce((s, v) => s + v, 0) / vals.length;
  if (mean === 0) return 0;
  const variance = vals.reduce((s, v) => s + (v - mean) ** 2, 0) / vals.length;
  // 返回变异系数 CV（归一化）
  return Math.sqrt(variance) / mean;
}

function calcMaxVariance(items, dimKey) {
  const vals = items.map(r => r[dimKey] || 0).filter(v => v > 0);
  if (vals.length <= 1) return 1;
  const mean = vals.reduce((s, v) => s + v, 0) / vals.length;
  if (mean === 0) return 1;
  const variance = vals.reduce((s, v) => s + (v - mean) ** 2, 0) / vals.length;
  return Math.sqrt(variance) / mean;
}

function renderPkgTable(page) {
  const search = (document.getElementById('pkg-search').value || '').trim().toLowerCase();
  const sortVal = document.getElementById('pkg-sort').value;

  let items = [...pkgResults];
  if (search) items = items.filter(p => p.name.toLowerCase().includes(search));

  const [field, dir] = sortVal.split('-');
  items.sort((a, b) => {
    let va, vb;
    if (field === 'kw') { va = a.items.length; vb = b.items.length; }
    else if (field === 'cost') { va = a.cost; vb = b.cost; }
    else if (field === 'impression') { va = a.impression; vb = b.impression; }
    else if (field === 'click') { va = a.click; vb = b.click; }
    else if (field === 'variance') { va = a.variance; vb = b.variance; }
    return dir === 'desc' ? vb - va : va - vb;
  });

  const totalPages = Math.max(1, Math.ceil(items.length / PKG_PAGE_SIZE));
  page = Math.max(1, Math.min(page, totalPages));
  const start = (page - 1) * PKG_PAGE_SIZE;
  const pageItems = items.slice(start, start + PKG_PAGE_SIZE);

  const dimKey = document.querySelector('input[name=pkg-dim]:checked').value;

  const tbody = document.getElementById('pkg-tbody');
  tbody.innerHTML = pageItems.map((pkg, i) => {
    const avgCpc = pkg.click > 0 ? '¥' + (pkg.cost / pkg.click).toFixed(2) : '-';
    const avgCtr = pkg.impression > 0 ? (pkg.click / pkg.impression * 100).toFixed(2) + '%' : '-';
    const sample = pkg.items.slice(0, 3).map(r => r.kw).join('、');
    const more = pkg.items.length > 3 ? `等${pkg.items.length}个` : '';
    const nameHtml = escHtml(pkg.name).replace(/^([^>]+)( > )?/, (m, l1, sep) => {
      const meta = L1_META[l1];
      return meta ? `<span class="l1-badge ${meta.cls}">${l1}</span>${sep || ''}` : m;
    });
    const tierHtml = pkg.tierLabel === '不分档' || pkg.tierLabel === '合并'
      ? `<span class="l2-tag">${pkg.tierLabel}</span>`
      : `<span class="pkg-tier ${pkg.tierClass}">${pkg.tierLabel}${pkg.dimLabel}</span>`;
    const cvPct = (pkg.variance * 100).toFixed(1);
    const cvColor = pkg.variance < 0.3 ? '#10b981' : pkg.variance < 0.7 ? '#f59e0b' : '#ef4444';
    return `<tr>
      <td class="num" style="color:#9ca3af">${start + i + 1}</td>
      <td style="font-weight:500;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(pkg.name)}">${nameHtml}</td>
      <td>${tierHtml}</td>
      <td class="num" style="font-weight:600;color:#1a73e8">${pkg.items.length}</td>
      <td class="num" style="color:#ef4444">${pkg.cost ? '¥' + pkg.cost.toLocaleString() : '-'}</td>
      <td class="num" style="color:#3b82f6">${pkg.impression ? pkg.impression.toLocaleString() : '-'}</td>
      <td class="num" style="color:#10b981">${pkg.click ? pkg.click.toLocaleString() : '-'}</td>
      <td class="num">${avgCpc}</td>
      <td class="num">${avgCtr}</td>
      <td class="num" style="color:${cvColor};font-weight:500">${cvPct}%</td>
      <td class="pkg-kw-sample" title="${escHtml(sample + more)}">${escHtml(sample)}${more ? '<span style="color:#1a73e8">' + more + '</span>' : ''}</td>
    </tr>`;
  }).join('');

  const pgEl = document.getElementById('pkg-pagination');
  if (items.length <= PKG_PAGE_SIZE) { pgEl.innerHTML = ''; return; }
  let html = `<span class="page-info">${items.length} 个词包，第 ${page}/${totalPages} 页</span>`;
  html += `<button class="page-btn" onclick="renderPkgTable(1)">«</button>`;
  html += `<button class="page-btn" onclick="renderPkgTable(${Math.max(1, page-1)})">‹</button>`;
  for (let p = Math.max(1, page-2); p <= Math.min(totalPages, page+2); p++) {
    html += `<button class="page-btn${p===page?' active':''}" onclick="renderPkgTable(${p})">${p}</button>`;
  }
  html += `<button class="page-btn" onclick="renderPkgTable(${Math.min(totalPages, page+1)})">›</button>`;
  html += `<button class="page-btn" onclick="renderPkgTable(${totalPages})">»</button>`;
  pgEl.innerHTML = html;
}

function exportPackages() {
  if (pkgResults.length === 0) { alert('请先生成词包'); return; }

  // 每个关键词一行，包含词包信息
  const rows = [['关键词','词包名称','档位','业务线','意图词包','城市','花费','曝光量','点击量','CTR','CPC']];
  pkgResults.forEach(pkg => {
    pkg.items.forEach(r => {
      const ctr = r.impression > 0 ? (r.click / r.impression * 100).toFixed(2) + '%' : '-';
      const cpc = r.click > 0 ? (r.cost / r.click).toFixed(2) : '-';
      const kw = String(r.kw).replace(/[\r\n,]/g, ' ');
      rows.push([kw, pkg.name, pkg.tierLabel + pkg.dimLabel, r.l1, r.l2, r.city, r.cost, r.impression, r.click, ctr, cpc]);
    });
  });

  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  downloadFile('\ufeff' + csv, '词包组合结果.csv', 'text/csv;charset=utf-8');
}