// Power BI Multi-Page Dashboard
// ================================

const FILE_NAME = "Solution List.xlsx";
const AUTO_REFRESH_MS = 60000;

// Supabase Configuration
const SUPABASE_URL = 'https://sbqmjzqwtzvrgujgswhs.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNicW1qenF3dHp2cmd1amdzd2hzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njk0MzAyMzEsImV4cCI6MjA4NTAwNjIzMX0.N1biiymjOz3T_mpGUiogQpsl0ts6PRBO_GwjVgg8cdY';
const SUPABASE_BUCKET = 'files';

// Column Keys
const KEYS = {
  division: "Division",
  name: "Solution Name",
  focus: "Focus Area",
  stage: "Stage",
  smv: "SMV Unlock",
  oh: "OH Reduction",
  other: "Other Savings"
};

// Colors
const COLORS = {
  blue: '#118DFF',
  green: '#12B76A',
  orange: '#E66C37',
  purple: '#8B5CF6',
  gold: '#F59E0B',
  teal: '#0891B2'
};

const CHART_COLORS = ['#118DFF', '#12B76A', '#E66C37', '#8B5CF6', '#F59E0B', '#0891B2', '#EF4444', '#EC4899'];

// State
let allRows = [];
let filteredRows = [];
let drillRows = [];
let charts = {};
let currentPage = 'home';
let lastRefreshAt = Date.now();
let previousPage = 'home';

// DOM Elements
const $ = id => document.getElementById(id);
const $$ = sel => document.querySelectorAll(sel);

// Utility Functions
function normalize(value) {
  if (value === undefined || value === null) return "Unspecified";
  return String(value).trim() || "Unspecified";
}

function toNumber(value) {
  if (value === undefined || value === null) return 0;
  const num = parseFloat(String(value).replace(/[^0-9.-]/g, ""));
  return isNaN(num) ? 0 : num;
}

function formatNumber(value, decimals = 2) {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  }).format(value);
}

function getUniqueValues(rows, key) {
  return [...new Set(rows.map(r => normalize(r[key])))].sort();
}

function groupBy(rows, key) {
  const groups = {};
  rows.forEach(row => {
    const label = normalize(row[key]);
    if (!groups[label]) groups[label] = { count: 0, smv: 0, oh: 0, other: 0, total: 0 };
    groups[label].count++;
    groups[label].smv += toNumber(row[KEYS.smv]);
    groups[label].oh += toNumber(row[KEYS.oh]);
    groups[label].other += toNumber(row[KEYS.other]);
    groups[label].total = groups[label].smv + groups[label].oh + groups[label].other;
  });
  return groups;
}

// Data Loading - fetches from Supabase Storage
async function loadData() {
  try {
    // Try to fetch from Supabase Storage first
    const supabaseFileUrl = `${SUPABASE_URL}/storage/v1/object/public/${SUPABASE_BUCKET}/${FILE_NAME}`;
    let response = await fetch(supabaseFileUrl, { cache: "no-store" });

    // Fallback to local file if Supabase file doesn't exist
    if (!response.ok) {
      response = await fetch(FILE_NAME, { cache: "no-store" });
    }

    if (!response.ok) throw new Error("File not found");

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    allRows = rawRows.map(row => {
      const normalized = {};
      Object.keys(row).forEach(key => normalized[key.trim()] = row[key]);
      return normalized;
    });

    filteredRows = [...allRows];
    lastRefreshAt = Date.now();

    populateFilters();
    updateAllPages();
    updateConnectionStatus(true);
    updateLastRefresh();

  } catch (error) {
    console.error("Load error:", error);
    updateConnectionStatus(false);
  }
}

function updateConnectionStatus(connected) {
  const statusEl = $('connectionStatus');
  if (statusEl) {
    statusEl.innerHTML = connected
      ? `<span class="status-indicator"></span><span class="status-text">Connected</span>`
      : `<span class="status-indicator" style="background:#EF4444"></span><span class="status-text" style="color:#EF4444">Error</span>`;
  }
}

function updateLastRefresh() {
  const el = $('lastRefresh');
  if (el) el.textContent = new Date().toLocaleTimeString();
}

// Filters
function populateFilters() {
  const divisions = getUniqueValues(allRows, KEYS.division);
  const stages = getUniqueValues(allRows, KEYS.stage);
  const focuses = getUniqueValues(allRows, KEYS.focus);

  populateSelect('divisionFilter', divisions);
  populateSelect('stageFilter', stages);
  populateSelect('focusFilter', focuses);
}

function populateSelect(id, values) {
  const select = $(id);
  if (!select) return;
  const current = select.value;
  select.innerHTML = '<option value="">All</option>';
  values.forEach(v => select.innerHTML += `<option value="${v}">${v}</option>`);
  if (current && values.includes(current)) select.value = current;
}

function applyFilters() {
  const division = $('divisionFilter')?.value || '';
  const stage = $('stageFilter')?.value || '';
  const focus = $('focusFilter')?.value || '';
  const search = $('searchFilter')?.value?.toLowerCase().trim() || '';

  filteredRows = allRows.filter(row => {
    if (division && normalize(row[KEYS.division]) !== division) return false;
    if (stage && normalize(row[KEYS.stage]) !== stage) return false;
    if (focus && normalize(row[KEYS.focus]) !== focus) return false;
    if (search && !normalize(row[KEYS.name]).toLowerCase().includes(search)) return false;
    return true;
  });

  updateFilterInfo();
  updateAllPages();
}

function updateFilterInfo() {
  const el = $('filterCount');
  if (el) {
    el.textContent = filteredRows.length === allRows.length
      ? 'Showing all solutions'
      : `Showing ${filteredRows.length} of ${allRows.length}`;
  }
}

function clearFilters() {
  ['divisionFilter', 'stageFilter', 'focusFilter', 'searchFilter'].forEach(id => {
    const el = $(id);
    if (el) el.value = '';
  });
  applyFilters();
}

// Page Navigation
function switchPage(page) {
  previousPage = currentPage;
  currentPage = page;

  $$('.page').forEach(p => p.classList.remove('active'));
  $$('.tab-btn').forEach(t => t.classList.remove('active'));

  $(`page-${page}`)?.classList.add('active');
  document.querySelector(`[data-page="${page}"]`)?.classList.add('active');

  // Rebuild charts when switching pages
  setTimeout(() => updatePageCharts(page), 100);
}

// Update All Pages
function updateAllPages() {
  updateKPIs();
  updateHomeCharts();
  updateMatrix();
  updateRankings();
  updateInsights();
  updateDrillPage();
}

function updatePageCharts(page) {
  switch(page) {
    case 'home': updateHomeCharts(); break;
    case 'matrix': updateMatrixCharts(); break;
    case 'summary': updateSummaryCharts(); break;
    case 'drillthrough': updateDrillCharts(); break;
  }
}

// KPIs
function updateKPIs() {
  const total = filteredRows.length;
  const smv = filteredRows.reduce((s, r) => s + toNumber(r[KEYS.smv]), 0);
  const oh = filteredRows.reduce((s, r) => s + toNumber(r[KEYS.oh]), 0);
  const other = filteredRows.reduce((s, r) => s + toNumber(r[KEYS.other]), 0);

  animateValue('kpiTotal', total, 0);
  animateValue('kpiSmv', smv, 3);
  animateValue('kpiOh', oh, 1);
  animateValue('kpiOther', other, 1);
  animateValue('kpiTotalSavings', smv + oh + other, 2);
}

function animateValue(id, target, decimals) {
  const el = $(id);
  if (!el) return;
  const start = parseFloat(el.textContent.replace(/,/g, '')) || 0;
  const duration = 400;
  const startTime = performance.now();

  function update(now) {
    const progress = Math.min((now - startTime) / duration, 1);
    const current = start + (target - start) * (1 - Math.pow(1 - progress, 3));
    el.textContent = decimals === 0 ? Math.round(current) : formatNumber(current, decimals);
    if (progress < 1) requestAnimationFrame(update);
  }
  requestAnimationFrame(update);
}

// Home Charts
function updateHomeCharts() {
  createDivisionChart();
  createStageChart();
  createFocusChart();
}

function createDivisionChart() {
  const ctx = $('divisionChart');
  if (!ctx) return;

  const groups = groupBy(filteredRows, KEYS.division);
  const labels = Object.keys(groups).sort();

  const config = {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'SMV Unlock', data: labels.map(l => groups[l].smv), backgroundColor: COLORS.blue, borderRadius: 4 },
        { label: 'OH Reduction', data: labels.map(l => groups[l].oh), backgroundColor: COLORS.green, borderRadius: 4 },
        { label: 'Other Savings', data: labels.map(l => groups[l].other), backgroundColor: COLORS.orange, borderRadius: 4 }
      ]
    },
    options: getBarOptions()
  };

  updateOrCreateChart('divisionChart', 'division', config);
}

function createStageChart() {
  const ctx = $('stageChart');
  if (!ctx) return;

  const groups = groupBy(filteredRows, KEYS.stage);
  const labels = Object.keys(groups).sort();

  const config = {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: labels.map(l => groups[l].count),
        backgroundColor: CHART_COLORS.slice(0, labels.length),
        borderWidth: 0
      }]
    },
    options: getDoughnutOptions()
  };

  updateOrCreateChart('stageChart', 'stage', config);
}

function createFocusChart() {
  const ctx = $('focusChart');
  if (!ctx) return;

  const groups = groupBy(filteredRows, KEYS.focus);
  const labels = Object.keys(groups).sort();

  const config = {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'SMV', data: labels.map(l => groups[l].smv), backgroundColor: COLORS.blue },
        { label: 'OH', data: labels.map(l => groups[l].oh), backgroundColor: COLORS.green },
        { label: 'Other', data: labels.map(l => groups[l].other), backgroundColor: COLORS.orange }
      ]
    },
    options: { ...getBarOptions(), indexAxis: 'y' }
  };

  updateOrCreateChart('focusChart', 'focus', config);
}

// Matrix
function updateMatrix() {
  const table = $('matrixTable');
  if (!table) return;

  const stages = getUniqueValues(filteredRows, KEYS.stage);
  const focuses = getUniqueValues(filteredRows, KEYS.focus);

  // Build data
  const data = {};
  const stageTotals = {};
  const focusTotals = {};
  let grand = { count: 0, smv: 0, oh: 0, other: 0 };

  stages.forEach(s => {
    data[s] = {};
    stageTotals[s] = { count: 0, smv: 0, oh: 0, other: 0 };
    focuses.forEach(f => data[s][f] = { count: 0, smv: 0, oh: 0, other: 0 });
  });
  focuses.forEach(f => focusTotals[f] = { count: 0, smv: 0, oh: 0, other: 0 });

  filteredRows.forEach(row => {
    const s = normalize(row[KEYS.stage]);
    const f = normalize(row[KEYS.focus]);
    const smv = toNumber(row[KEYS.smv]);
    const oh = toNumber(row[KEYS.oh]);
    const other = toNumber(row[KEYS.other]);

    if (data[s]?.[f]) {
      data[s][f].count++; data[s][f].smv += smv; data[s][f].oh += oh; data[s][f].other += other;
    }
    if (stageTotals[s]) {
      stageTotals[s].count++; stageTotals[s].smv += smv; stageTotals[s].oh += oh; stageTotals[s].other += other;
    }
    if (focusTotals[f]) {
      focusTotals[f].count++; focusTotals[f].smv += smv; focusTotals[f].oh += oh; focusTotals[f].other += other;
    }
    grand.count++; grand.smv += smv; grand.oh += oh; grand.other += other;
  });

  // Build HTML
  let html = `<thead><tr><th>Stage \\ Focus Area</th>`;
  focuses.forEach(f => html += `<th>${f}</th>`);
  html += `<th>Total</th></tr></thead><tbody>`;

  stages.forEach(s => {
    html += `<tr><th>${s}</th>`;
    focuses.forEach(f => html += createMatrixCell(data[s][f], s, f));
    html += createMatrixCell(stageTotals[s], s, null, true);
    html += `</tr>`;
  });
  html += `</tbody><tfoot><tr><th>Total</th>`;
  focuses.forEach(f => html += createMatrixCell(focusTotals[f], null, f, true));
  html += createMatrixCell(grand, null, null, false, true);
  html += `</tr></tfoot>`;

  table.innerHTML = html;

  // Add click events
  table.querySelectorAll('.matrix-cell').forEach(cell => {
    cell.addEventListener('click', () => {
      const stage = cell.dataset.stage || null;
      const focus = cell.dataset.focus || null;
      drillThrough(stage, focus);
    });
  });
}

function createMatrixCell(d, stage, focus, isTotal = false, isGrand = false) {
  const cls = isGrand ? 'matrix-cell grand-total' : isTotal ? 'matrix-cell total-row' : 'matrix-cell';
  const stageAttr = stage ? `data-stage="${stage}"` : '';
  const focusAttr = focus ? `data-focus="${focus}"` : '';

  return `<td><div class="${cls}" ${stageAttr} ${focusAttr}>
    <span class="cell-count">${d.count}</span>
    <div class="cell-metrics">
      <span class="cell-metric smv">${formatNumber(d.smv, 2)}</span>
      <span class="cell-metric oh">${formatNumber(d.oh, 1)}</span>
      <span class="cell-metric other">${formatNumber(d.other, 1)}</span>
    </div>
  </div></td>`;
}

function updateMatrixCharts() {
  // Stage chart
  const stageGroups = groupBy(filteredRows, KEYS.stage);
  const stageLabels = Object.keys(stageGroups).sort();

  updateOrCreateChart('matrixStageChart', 'matrixStage', {
    type: 'bar',
    data: {
      labels: stageLabels,
      datasets: [{ data: stageLabels.map(l => stageGroups[l].total), backgroundColor: CHART_COLORS }]
    },
    options: { ...getBarOptions(), plugins: { legend: { display: false } } }
  });

  // Focus chart
  const focusGroups = groupBy(filteredRows, KEYS.focus);
  const focusLabels = Object.keys(focusGroups).sort();

  updateOrCreateChart('matrixFocusChart', 'matrixFocus', {
    type: 'bar',
    data: {
      labels: focusLabels,
      datasets: [{ data: focusLabels.map(l => focusGroups[l].total), backgroundColor: CHART_COLORS }]
    },
    options: { ...getBarOptions(), indexAxis: 'y', plugins: { legend: { display: false } } }
  });
}

// Rankings
function updateRankings() {
  const tbody = document.querySelector('#rankingsTable tbody');
  if (!tbody) return;

  const groups = groupBy(filteredRows, KEYS.division);
  const sorted = Object.entries(groups).sort((a, b) => b[1].total - a[1].total);

  tbody.innerHTML = sorted.map(([div, d], i) => {
    const rank = i + 1;
    const badge = rank <= 3 ? `<span class="rank-badge rank-${rank}">${rank}</span>` : rank;
    return `<tr data-division="${div}">
      <td>${badge}</td>
      <td><strong>${div}</strong></td>
      <td>${d.count}</td>
      <td>${formatNumber(d.smv, 3)}</td>
      <td>${formatNumber(d.oh, 1)}</td>
      <td>${formatNumber(d.other, 1)}</td>
      <td><strong>${formatNumber(d.total, 2)}</strong></td>
    </tr>`;
  }).join('');

  // Click to drill through
  tbody.querySelectorAll('tr').forEach(row => {
    row.addEventListener('click', () => {
      const division = row.dataset.division;
      drillThroughDivision(division);
    });
  });
}

// Insights
function updateInsights() {
  const groups = groupBy(filteredRows, KEYS.division);
  const entries = Object.entries(groups);

  if (entries.length === 0) return;

  const topDiv = entries.sort((a, b) => b[1].total - a[1].total)[0];
  const bestSmv = entries.sort((a, b) => b[1].smv - a[1].smv)[0];
  const bestOh = entries.sort((a, b) => b[1].oh - a[1].oh)[0];
  const mostSolutions = entries.sort((a, b) => b[1].count - a[1].count)[0];

  const commercialized = filteredRows.filter(r => normalize(r[KEYS.stage]).toLowerCase() === 'commercialized').length;
  const rnd = filteredRows.filter(r => normalize(r[KEYS.stage]).toLowerCase() === 'r&d').length;

  setText('insightTopDiv', topDiv[0]);
  setText('insightBestSmv', bestSmv[0]);
  setText('insightBestOh', bestOh[0]);
  setText('insightMostSolutions', `${mostSolutions[0]} (${mostSolutions[1].count})`);
  setText('insightCommercialized', `${commercialized} solutions`);
  setText('insightRnd', `${rnd} solutions`);
}

function setText(id, text) {
  const el = $(id);
  if (el) el.textContent = text;
}

// Summary Charts
function updateSummaryCharts() {
  const groups = groupBy(filteredRows, KEYS.division);
  const sorted = Object.entries(groups).sort((a, b) => b[1].total - a[1].total);
  const labels = sorted.map(s => s[0]);

  updateOrCreateChart('summaryBarChart', 'summaryBar', {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'SMV', data: sorted.map(s => s[1].smv), backgroundColor: COLORS.blue },
        { label: 'OH', data: sorted.map(s => s[1].oh), backgroundColor: COLORS.green },
        { label: 'Other', data: sorted.map(s => s[1].other), backgroundColor: COLORS.orange }
      ]
    },
    options: getBarOptions()
  });

  updateOrCreateChart('summaryPieChart', 'summaryPie', {
    type: 'doughnut',
    data: {
      labels: ['SMV Unlock', 'OH Reduction', 'Other Savings'],
      datasets: [{
        data: [
          filteredRows.reduce((s, r) => s + toNumber(r[KEYS.smv]), 0),
          filteredRows.reduce((s, r) => s + toNumber(r[KEYS.oh]), 0),
          filteredRows.reduce((s, r) => s + toNumber(r[KEYS.other]), 0)
        ],
        backgroundColor: [COLORS.blue, COLORS.green, COLORS.orange],
        borderWidth: 0
      }]
    },
    options: getDoughnutOptions()
  });
}

// Drill Through
function drillThrough(stage, focus) {
  drillRows = filteredRows.filter(row => {
    if (stage && normalize(row[KEYS.stage]) !== stage) return false;
    if (focus && normalize(row[KEYS.focus]) !== focus) return false;
    return true;
  });

  let title = 'All Solutions';
  let subtitle = '';

  if (stage && focus) {
    title = `${stage} - ${focus}`;
    subtitle = `Solutions in ${stage} stage with ${focus} focus area`;
  } else if (stage) {
    title = `${stage} Stage`;
    subtitle = `All solutions in ${stage} stage`;
  } else if (focus) {
    title = focus;
    subtitle = `All solutions with ${focus} focus`;
  }

  setText('drillTitle', title);
  setText('drillSubtitle', subtitle);

  switchPage('drillthrough');
  updateDrillPage();
}

function drillThroughDivision(division) {
  drillRows = filteredRows.filter(row => normalize(row[KEYS.division]) === division);

  setText('drillTitle', division);
  setText('drillSubtitle', `All solutions in ${division} division`);

  switchPage('drillthrough');
  updateDrillPage();
}

function updateDrillPage() {
  if (drillRows.length === 0) drillRows = filteredRows;

  const smv = drillRows.reduce((s, r) => s + toNumber(r[KEYS.smv]), 0);
  const oh = drillRows.reduce((s, r) => s + toNumber(r[KEYS.oh]), 0);
  const other = drillRows.reduce((s, r) => s + toNumber(r[KEYS.other]), 0);

  setText('drillCount', drillRows.length);
  setText('drillSmv', formatNumber(smv, 3));
  setText('drillOh', formatNumber(oh, 1));
  setText('drillOther', formatNumber(other, 1));
  setText('drillTotal', formatNumber(smv + oh + other, 2));
  setText('drillRecordCount', `${drillRows.length} records`);

  // Table
  const tbody = document.querySelector('#drillTable tbody');
  if (tbody) {
    tbody.innerHTML = drillRows.map(row => {
      const smv = toNumber(row[KEYS.smv]);
      const oh = toNumber(row[KEYS.oh]);
      const other = toNumber(row[KEYS.other]);
      return `<tr>
        <td><strong>${normalize(row[KEYS.name])}</strong></td>
        <td>${normalize(row[KEYS.division])}</td>
        <td>${normalize(row[KEYS.focus])}</td>
        <td>${normalize(row[KEYS.stage])}</td>
        <td>${formatNumber(smv, 3)}</td>
        <td>${formatNumber(oh, 1)}</td>
        <td>${formatNumber(other, 1)}</td>
        <td><strong>${formatNumber(smv + oh + other, 2)}</strong></td>
      </tr>`;
    }).join('');
  }

  updateDrillCharts();
}

function updateDrillCharts() {
  const divGroups = groupBy(drillRows, KEYS.division);
  const divLabels = Object.keys(divGroups).sort();

  updateOrCreateChart('drillDivisionChart', 'drillDivision', {
    type: 'bar',
    data: {
      labels: divLabels,
      datasets: [{ data: divLabels.map(l => divGroups[l].total), backgroundColor: CHART_COLORS }]
    },
    options: { ...getBarOptions(), plugins: { legend: { display: false } } }
  });

  const stageGroups = groupBy(drillRows, KEYS.stage);
  const stageLabels = Object.keys(stageGroups).sort();

  updateOrCreateChart('drillStageChart', 'drillStage', {
    type: 'doughnut',
    data: {
      labels: stageLabels,
      datasets: [{ data: stageLabels.map(l => stageGroups[l].count), backgroundColor: CHART_COLORS }]
    },
    options: getDoughnutOptions()
  });
}

// Chart Helpers
function updateOrCreateChart(canvasId, chartKey, config) {
  const ctx = $(canvasId);
  if (!ctx) return;

  if (charts[chartKey]) {
    charts[chartKey].destroy();
  }

  charts[chartKey] = new Chart(ctx, config);
}

function getBarOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'bottom', labels: { usePointStyle: true, padding: 15, font: { size: 12 }, color: '#ffffff' } },
      tooltip: { backgroundColor: '#252423', padding: 12, cornerRadius: 4 }
    },
    scales: {
      x: { grid: { display: false }, ticks: { font: { size: 12 }, color: '#ffffff' } },
      y: { grid: { color: 'rgba(255, 255, 255, 0.15)' }, ticks: { font: { size: 12 }, color: '#ffffff' } }
    }
  };
}

function getDoughnutOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    cutout: '55%',
    plugins: {
      legend: { position: 'bottom', labels: { usePointStyle: true, padding: 12, font: { size: 12 }, color: '#ffffff' } },
      tooltip: {
        backgroundColor: '#252423',
        callbacks: {
          label: ctx => {
            const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
            const pct = ((ctx.parsed / total) * 100).toFixed(1);
            return ` ${ctx.label}: ${ctx.parsed} (${pct}%)`;
          }
        }
      }
    }
  };
}

function getPieOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'bottom', labels: { usePointStyle: true, padding: 12, font: { size: 11 }, color: '#ffffff' } },
      tooltip: {
        backgroundColor: '#252423',
        callbacks: {
          label: ctx => {
            const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
            const pct = ((ctx.parsed / total) * 100).toFixed(1);
            return ` ${ctx.label}: ${formatNumber(ctx.parsed)} (${pct}%)`;
          }
        }
      }
    }
  };
}

// Export
function exportCSV() {
  const rows = [['Solution Name', 'Division', 'Focus Area', 'Stage', 'SMV Unlock', 'OH Reduction', 'Other Savings', 'Total']];

  (currentPage === 'drillthrough' ? drillRows : filteredRows).forEach(row => {
    const smv = toNumber(row[KEYS.smv]);
    const oh = toNumber(row[KEYS.oh]);
    const other = toNumber(row[KEYS.other]);
    rows.push([
      normalize(row[KEYS.name]),
      normalize(row[KEYS.division]),
      normalize(row[KEYS.focus]),
      normalize(row[KEYS.stage]),
      smv, oh, other, smv + oh + other
    ]);
  });

  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'solution_savings_export.csv';
  link.click();
}

// Countdown
function updateCountdown() {
  const elapsed = Date.now() - lastRefreshAt;
  const remaining = Math.max(0, AUTO_REFRESH_MS - elapsed);
  setText('countdown', Math.ceil(remaining / 1000));
}

// Download Excel from Supabase Storage
function downloadExcel() {
  const supabaseFileUrl = `${SUPABASE_URL}/storage/v1/object/public/${SUPABASE_BUCKET}/${FILE_NAME}`;
  window.location.href = supabaseFileUrl;
}

// Upload Excel to Supabase Storage
async function uploadExcel(file) {
  if (!file) return;

  showLoading(true);

  try {
    // Upload file to Supabase Storage
    const uploadUrl = `${SUPABASE_URL}/storage/v1/object/${SUPABASE_BUCKET}/${FILE_NAME}`;

    const uploadResponse = await fetch(uploadUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'x-upsert': 'true' // Overwrite if exists
      },
      body: file
    });

    if (!uploadResponse.ok) {
      const errorData = await uploadResponse.json();
      throw new Error(errorData.message || 'Failed to save file to cloud storage');
    }

    // Process the file locally to update the UI immediately
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    allRows = rawRows.map(row => {
      const normalized = {};
      Object.keys(row).forEach(key => normalized[key.trim()] = row[key]);
      return normalized;
    });

    filteredRows = [...allRows];
    lastRefreshAt = Date.now();

    populateFilters();
    updateAllPages();
    updateConnectionStatus(true);
    updateLastRefresh();

    showLoading(false);
    showNotification('File uploaded and saved to cloud!', false);
  } catch (error) {
    showLoading(false);
    showNotification('Failed to upload file: ' + error.message, true);
  }
}

// Logout
function logout() {
  localStorage.removeItem('dashboard_auth');
  localStorage.removeItem('dashboard_user');
  window.location.href = '/login.html';
}

// Show loading overlay
function showLoading(show) {
  let overlay = document.querySelector('.loading-overlay');

  if (show && !overlay) {
    overlay = document.createElement('div');
    overlay.className = 'loading-overlay';
    overlay.innerHTML = '<div class="spinner"></div>';
    document.body.appendChild(overlay);
  } else if (!show && overlay) {
    overlay.remove();
  }
}

// Show notification
function showNotification(message, isError = false) {
  // Remove existing notification
  const existing = document.querySelector('.upload-notification');
  if (existing) existing.remove();

  const notification = document.createElement('div');
  notification.className = `upload-notification ${isError ? 'error' : ''}`;
  notification.innerHTML = `
    <div class="notification-icon">
      <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
        ${isError
          ? '<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>'
          : '<path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>'}
      </svg>
    </div>
    <span class="notification-text">${message}</span>
  `;

  document.body.appendChild(notification);

  setTimeout(() => notification.remove(), 4000);
}

// Get username from localStorage
function checkAuth() {
  const username = localStorage.getItem('dashboard_user');
  if (username) {
    const userNameEl = $('userName');
    if (userNameEl) userNameEl.textContent = username;
  }
}

// Event Listeners
function initEventListeners() {
  // Tab navigation
  $$('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchPage(btn.dataset.page));
  });

  // Filters
  ['divisionFilter', 'stageFilter', 'focusFilter', 'searchFilter'].forEach(id => {
    $(id)?.addEventListener('change', applyFilters);
    $(id)?.addEventListener('input', applyFilters);
  });

  $('clearFilters')?.addEventListener('click', clearFilters);
  $('refreshBtn')?.addEventListener('click', loadData);
  $('exportBtn')?.addEventListener('click', exportCSV);
  $('backBtn')?.addEventListener('click', () => switchPage(previousPage));

  // Download Excel
  $('downloadExcel')?.addEventListener('click', downloadExcel);

  // Upload Excel
  $('fileInput')?.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) uploadExcel(file);
    e.target.value = ''; // Reset to allow re-uploading same file
  });

  // Logout
  $('logoutBtn')?.addEventListener('click', logout);
}

// Initialize
function init() {
  initEventListeners();
  checkAuth();
  loadData();
  setInterval(loadData, AUTO_REFRESH_MS);
  setInterval(updateCountdown, 1000);
}

document.addEventListener('DOMContentLoaded', init);
