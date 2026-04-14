/* ============================================
   ABC-Analysis Application Logic
   ============================================ */

// ── State ──────────────────────────────────
let rawData = [];             // parsed rows from Excel
let globalABC = [];           // ABC result: global
let warehouseABC = [];        // ABC result: per warehouse
let warehouses = [];          // unique warehouse names
let activeTab = 'global';     // current tab
let paretoChart = null;       // Chart.js instance
let currentSort = { field: 'qty', dir: 'desc' };

// ── DOM refs ───────────────────────────────
const $ = (id) => document.getElementById(id);
const uploadSection  = $('uploadSection');
const uploadArea     = $('uploadArea');
const fileInput      = $('fileInput');
const uploadBtn      = $('uploadBtn');
const loadingOverlay = $('loadingOverlay');
const resultsSection = $('resultsSection');
const headerStats    = $('headerStats');

const tabGlobal    = $('tabGlobal');
const tabWarehouse = $('tabWarehouse');
const tabPareto    = $('tabPareto');

const panelGlobal    = $('panelGlobal');
const panelWarehouse = $('panelWarehouse');
const panelPareto    = $('panelPareto');

const filterWarehouse     = $('filterWarehouse');
const filterWarehouseWrap = $('filterWarehouseWrap');
const filterCategory      = $('filterCategory');
const filterSearch        = $('filterSearch');
const exportBtn           = $('exportBtn');
const paretoTopN          = $('paretoTopN');

// ── File Upload ────────────────────────────
uploadArea.addEventListener('click', () => fileInput.click());
uploadBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

// Drag & Drop
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});
uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});
uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

function handleFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) {
        alert('Пожалуйста, загрузите файл формата .xlsx или .xls');
        return;
    }

    uploadSection.style.display = 'none';
    loadingOverlay.style.display = 'flex';

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            parseExcel(e.target.result);
            performABC();
            showResults();
        } catch (err) {
            console.error(err);
            alert('Ошибка при обработке файла: ' + err.message);
            uploadSection.style.display = '';
            loadingOverlay.style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

// ── Parse Excel ────────────────────────────
function parseExcel(buffer) {
    const wb = XLSX.read(buffer, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

    rawData = [];

    // Find header row — look for row containing 'Товар' or 'Кол-во'
    let headerIdx = 0;
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
        const row = rows[i];
        if (row && row.some(c => typeof c === 'string' && (c.includes('Товар') || c.includes('Кол-во')))) {
            headerIdx = i;
            break;
        }
    }

    // Parse data rows (skip header)
    for (let i = headerIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        const product   = row[2]; // Column C (0-indexed: 2)
        const warehouse = row[3]; // Column D
        const unit      = row[4]; // Column E
        const qty       = row[5]; // Column F

        // Validation: must have product and qty
        if (!product || product === null) continue;
        if (qty === null || qty === undefined || qty === '' || typeof qty !== 'number') continue;

        // Only write-offs: F < 0
        if (qty >= 0) continue;

        rawData.push({
            product: String(product).trim(),
            warehouse: warehouse ? String(warehouse).trim() : '—',
            unit: unit ? String(unit).trim() : '',
            qty: Math.abs(qty)
        });
    }

    if (rawData.length === 0) {
        throw new Error('Не найдено записей списания (F < 0). Проверьте структуру файла.');
    }
}

// ── ABC Calculation ────────────────────────
function performABC() {
    // --- Global ---
    const globalMap = new Map();
    for (const r of rawData) {
        globalMap.set(r.product, (globalMap.get(r.product) || 0) + r.qty);
    }

    globalABC = Array.from(globalMap.entries())
        .map(([product, sumQty]) => ({ product, sumQty }))
        .sort((a, b) => b.sumQty - a.sumQty);

    const totalQtyGlobal = globalABC.reduce((s, r) => s + r.sumQty, 0);
    let cumGlobal = 0;
    for (const r of globalABC) {
        r.share = r.sumQty / totalQtyGlobal;
        cumGlobal += r.share;
        r.cumulative = cumGlobal;
        r.category = cumGlobal <= 0.8 ? 'A' : cumGlobal <= 0.95 ? 'B' : 'C';
    }

    // --- Per Warehouse ---
    const whMap = new Map();
    for (const r of rawData) {
        const key = `${r.product}|||${r.warehouse}`;
        whMap.set(key, (whMap.get(key) || 0) + r.qty);
    }

    warehouseABC = Array.from(whMap.entries())
        .map(([key, sumQty]) => {
            const [product, warehouse] = key.split('|||');
            return { product, warehouse, sumQty };
        })
        .sort((a, b) => b.sumQty - a.sumQty);

    const totalQtyWH = warehouseABC.reduce((s, r) => s + r.sumQty, 0);
    let cumWH = 0;
    for (const r of warehouseABC) {
        r.share = r.sumQty / totalQtyWH;
        cumWH += r.share;
        r.cumulative = cumWH;
        r.category = cumWH <= 0.8 ? 'A' : cumWH <= 0.95 ? 'B' : 'C';
    }

    // --- Unique warehouses ---
    warehouses = [...new Set(rawData.map(r => r.warehouse))].sort();
}

// ── Show Results ───────────────────────────
function showResults() {
    loadingOverlay.style.display = 'none';
    resultsSection.style.display = '';
    headerStats.style.display = 'flex';

    // Stats
    $('statRows').textContent = rawData.length.toLocaleString('ru-RU');
    $('statProducts').textContent = globalABC.length.toLocaleString('ru-RU');
    $('statWarehouses').textContent = warehouses.length.toLocaleString('ru-RU');

    // Summary cards
    updateSummaryCards(globalABC);

    // Populate warehouse filter
    filterWarehouse.innerHTML = '<option value="">Все склады</option>';
    for (const w of warehouses) {
        const opt = document.createElement('option');
        opt.value = w;
        opt.textContent = w;
        filterWarehouse.appendChild(opt);
    }

    // Render tables
    renderGlobalTable();
    renderWarehouseTable();

    // Set active tab
    switchTab('global');
}

// ── Summary Cards ──────────────────────────
function updateSummaryCards(data) {
    const countA = data.filter(r => r.category === 'A').length;
    const countB = data.filter(r => r.category === 'B').length;
    const countC = data.filter(r => r.category === 'C').length;
    const total = data.length || 1;

    $('countA').textContent = countA;
    $('countB').textContent = countB;
    $('countC').textContent = countC;

    // Animate bars
    setTimeout(() => {
        $('barA').style.width = ((countA / total) * 100) + '%';
        $('barB').style.width = ((countB / total) * 100) + '%';
        $('barC').style.width = ((countC / total) * 100) + '%';
    }, 100);
}

// ── Render Global Table ────────────────────
function renderGlobalTable() {
    const filtered = applyFilters(globalABC, false);
    const tbody = $('tbodyGlobal');
    tbody.innerHTML = '';

    $('globalRowCount').textContent = `${filtered.length} записей`;

    for (let i = 0; i < filtered.length; i++) {
        const r = filtered[i];
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="td-num">${i + 1}</td>
            <td class="td-product" title="${escapeHtml(r.product)}">${escapeHtml(r.product)}</td>
            <td class="td-qty">${formatNumber(r.sumQty)}</td>
            <td class="td-share">${(r.share * 100).toFixed(2)}%</td>
            <td class="td-cumulative">${(r.cumulative * 100).toFixed(2)}%</td>
            <td class="td-category"><span class="category-badge badge-${r.category.toLowerCase()}">${r.category}</span></td>
        `;
        tbody.appendChild(tr);
    }
}

// ── Render Warehouse Table ─────────────────
function renderWarehouseTable() {
    const filtered = applyFilters(warehouseABC, true);
    const tbody = $('tbodyWarehouse');
    tbody.innerHTML = '';

    $('warehouseRowCount').textContent = `${filtered.length} записей`;

    for (let i = 0; i < filtered.length; i++) {
        const r = filtered[i];
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="td-num">${i + 1}</td>
            <td class="td-product" title="${escapeHtml(r.product)}">${escapeHtml(r.product)}</td>
            <td class="td-warehouse" title="${escapeHtml(r.warehouse)}">${escapeHtml(r.warehouse)}</td>
            <td class="td-qty">${formatNumber(r.sumQty)}</td>
            <td class="td-share">${(r.share * 100).toFixed(2)}%</td>
            <td class="td-cumulative">${(r.cumulative * 100).toFixed(2)}%</td>
            <td class="td-category"><span class="category-badge badge-${r.category.toLowerCase()}">${r.category}</span></td>
        `;
        tbody.appendChild(tr);
    }
}

// ── Filters ────────────────────────────────
function applyFilters(data, hasWarehouse) {
    let result = [...data];
    const whFilter  = filterWarehouse.value;
    const catFilter = filterCategory.value;
    const search    = filterSearch.value.trim().toLowerCase();

    if (whFilter && hasWarehouse) {
        result = result.filter(r => r.warehouse === whFilter);
    }

    if (catFilter) {
        result = result.filter(r => r.category === catFilter);
    }

    if (search) {
        result = result.filter(r => r.product.toLowerCase().includes(search));
    }

    return result;
}

// Debounced filter handler
let filterTimeout;
function onFilterChange() {
    clearTimeout(filterTimeout);
    filterTimeout = setTimeout(() => {
        renderGlobalTable();
        renderWarehouseTable();
        if (activeTab === 'pareto') renderParetoChart();
    }, 200);
}

filterWarehouse.addEventListener('change', onFilterChange);
filterCategory.addEventListener('change', onFilterChange);
filterSearch.addEventListener('input', onFilterChange);

// ── Tab Switching ──────────────────────────
function switchTab(tab) {
    activeTab = tab;

    [tabGlobal, tabWarehouse, tabPareto].forEach(b => b.classList.remove('active'));
    [panelGlobal, panelWarehouse, panelPareto].forEach(p => p.style.display = 'none');

    if (tab === 'global') {
        tabGlobal.classList.add('active');
        panelGlobal.style.display = '';
        filterWarehouseWrap.style.display = 'none';
        updateSummaryCards(globalABC);
    } else if (tab === 'warehouse') {
        tabWarehouse.classList.add('active');
        panelWarehouse.style.display = '';
        filterWarehouseWrap.style.display = '';
        updateSummaryCards(warehouseABC);
    } else if (tab === 'pareto') {
        tabPareto.classList.add('active');
        panelPareto.style.display = '';
        filterWarehouseWrap.style.display = 'none';
        renderParetoChart();
    }
}

tabGlobal.addEventListener('click', () => switchTab('global'));
tabWarehouse.addEventListener('click', () => switchTab('warehouse'));
tabPareto.addEventListener('click', () => switchTab('pareto'));

// ── Pareto Chart ───────────────────────────
function renderParetoChart() {
    const topN = parseInt(paretoTopN.value) || 20;
    const allFilteredData = applyFilters(globalABC, false);
    
    let chartData = allFilteredData.slice(0, topN);
    const othersData = allFilteredData.slice(topN);

    // Group "Others"
    if (othersData.length > 0) {
        const othersSum = othersData.reduce((s, r) => s + r.sumQty, 0);
        chartData.push({
            product: `Прочие (${othersData.length} наимен.)`,
            sumQty: othersSum,
            category: 'C',
            share: othersSum / allFilteredData.reduce((s, r) => s + r.sumQty, 0),
            isOthers: true
        });
    }

    if (paretoChart) paretoChart.destroy();

    const canvas = $('paretoChart');
    const ctx = canvas.getContext('2d');

    // Dynamic width for scrollability — significantly increased for long labels
    const minBarWidth = 120; 
    const requiredWidth = Math.max(canvas.parentElement.clientWidth, chartData.length * minBarWidth);
    canvas.style.width = requiredWidth + 'px';

    const labels = chartData.map(r => truncate(r.product, 25));
    const quantities = chartData.map(r => r.sumQty);

    // Calculate cumulative for the chart (based on total data, not just TOP-N)
    const totalQty = allFilteredData.reduce((s, r) => s + r.sumQty, 0);
    let cum = 0;
    const cumulatives = chartData.map(r => {
        cum += (r.sumQty / totalQty) * 100;
        return cum;
    });

    // Category colors
    const barColors = chartData.map(r => {
        if (r.isOthers) return 'rgba(139, 146, 165, 0.6)';
        if (r.category === 'A') return 'rgba(34, 197, 94, 0.7)';
        if (r.category === 'B') return 'rgba(245, 158, 11, 0.7)';
        return 'rgba(239, 68, 68, 0.7)';
    });

    const barBorders = chartData.map(r => {
        if (r.isOthers) return 'rgba(139, 146, 165, 1)';
        if (r.category === 'A') return 'rgba(34, 197, 94, 1)';
        if (r.category === 'B') return 'rgba(245, 158, 11, 1)';
        return 'rgba(239, 68, 68, 1)';
    });

    paretoChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: 'Кол-во',
                    data: quantities,
                    backgroundColor: barColors,
                    borderColor: barBorders,
                    borderWidth: 1,
                    borderRadius: 4,
                    yAxisID: 'y',
                    order: 2,
                },
                {
                    label: 'Кумулятивная доля, %',
                    data: cumulatives,
                    type: 'line',
                    borderColor: '#A5B4FC',
                    backgroundColor: 'rgba(165, 180, 252, 0.1)',
                    pointBackgroundColor: '#A5B4FC',
                    pointRadius: 4,
                    pointHoverRadius: 6,
                    borderWidth: 2.5,
                    fill: true,
                    tension: 0.3,
                    yAxisID: 'y1',
                    order: 1,
                }
            ]
        },
        options: {
            responsive: false, // Set false to respect manual width
            maintainAspectRatio: false,
            interaction: {
                intersect: false,
                mode: 'index'
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        color: '#8B92A5',
                        font: { family: 'Inter', size: 12 },
                        padding: 20,
                        usePointStyle: true
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(22, 27, 39, 0.98)',
                    borderColor: 'rgba(99, 102, 241, 0.4)',
                    borderWidth: 1,
                    titleColor: '#E8EAED',
                    bodyColor: '#E8EAED',
                    padding: 14,
                    cornerRadius: 8,
                    titleFont: { family: 'Inter', weight: '700', size: 14 },
                    bodyFont: { family: 'Inter', size: 13 },
                    callbacks: {
                        title: (items) => chartData[items[0].dataIndex].product,
                        label: function(ctx) {
                            const item = chartData[ctx.dataIndex];
                            if (ctx.dataset.label === 'Кол-во') {
                                const lines = [
                                    ` Кол-во: ${formatNumber(item.sumQty)}`,
                                    ` Доля: ${(item.share * 100).toFixed(2)}%`
                                ];
                                if (!item.isOthers) {
                                    lines.push(` Категория: ${item.category}`);
                                }
                                return lines;
                            }
                            return ` Кумулятивно: ${ctx.parsed.y.toFixed(1)}%`;
                        }
                    }
                }
            },
            layout: {
                padding: {
                    bottom: 120, // Significantly increased space for long rotated labels
                    left: 10,
                    right: 20
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: '#8B92A5',
                        font: { family: 'Inter', size: 11 },
                        maxRotation: 45,
                        minRotation: 45,
                        autoSkip: false,
                        padding: 8,
                        align: 'right',
                    },
                    grid: { display: false }
                },
                y: {
                    position: 'left',
                    beginAtZero: true,
                    ticks: {
                        color: '#5C6478',
                        font: { family: 'Inter', size: 11 },
                        callback: (v) => formatNumber(v)
                    },
                    grid: { color: 'rgba(255, 255, 255, 0.04)' },
                    title: {
                        display: true,
                        text: 'Количество списаний',
                        color: '#8B92A5',
                        font: { family: 'Inter', size: 12 }
                    }
                },
                y1: {
                    position: 'right',
                    beginAtZero: true,
                    max: 105,
                    ticks: {
                        color: '#5C6478',
                        font: { family: 'Inter', size: 11 },
                        callback: (v) => v + '%'
                    },
                    grid: { drawOnChartArea: false },
                    title: {
                        display: true,
                        text: 'Кумулятивная доля',
                        color: '#8B92A5',
                        font: { family: 'Inter', size: 12 }
                    }
                }
            }
        },
        plugins: [{
            id: 'paretoLine',
            afterDraw(chart) {
                const y1 = chart.scales.y1;
                const yPixel = y1.getPixelForValue(80);
                const ctx = chart.ctx;
                ctx.save();
                ctx.beginPath();
                ctx.setLineDash([8, 4]);
                ctx.strokeStyle = 'rgba(239, 68, 68, 0.6)';
                ctx.lineWidth = 1.5;
                ctx.moveTo(chart.chartArea.left, yPixel);
                ctx.lineTo(chart.chartArea.right, yPixel);
                ctx.stroke();

                ctx.fillStyle = 'rgba(239, 68, 68, 0.9)';
                ctx.font = 'bold 11px Inter';
                ctx.textAlign = 'right';
                ctx.fillText('Порог 80%', chart.chartArea.right - 8, yPixel - 8);
                ctx.restore();
            }
        }]
    });
}

paretoTopN.addEventListener('change', () => {
    if (activeTab === 'pareto') renderParetoChart();
});

// ── Sorting ────────────────────────────────
document.querySelectorAll('.data-table th.sortable').forEach(th => {
    th.addEventListener('click', () => {
        const table = th.closest('.data-table');
        const field = th.dataset.sort;
        const isGlobal = table.id === 'tableGlobal';
        const data = isGlobal ? globalABC : warehouseABC;

        // Toggle direction
        if (currentSort.field === field) {
            currentSort.dir = currentSort.dir === 'desc' ? 'asc' : 'desc';
        } else {
            currentSort.field = field;
            currentSort.dir = field === 'product' || field === 'warehouse' ? 'asc' : 'desc';
        }

        // Update header styles
        table.querySelectorAll('th').forEach(h => h.classList.remove('sorted-asc', 'sorted-desc'));
        th.classList.add(currentSort.dir === 'asc' ? 'sorted-asc' : 'sorted-desc');

        // Sort
        const sortFn = getSortFunction(currentSort.field, currentSort.dir);
        data.sort(sortFn);

        // Re-calculate cumulative shares after re-sorting
        const total = data.reduce((s, r) => s + r.sumQty, 0);
        let cum = 0;
        for (const r of data) {
            r.share = r.sumQty / total;
            cum += r.share;
            r.cumulative = cum;
            r.category = cum <= 0.8 ? 'A' : cum <= 0.95 ? 'B' : 'C';
        }

        if (isGlobal) renderGlobalTable();
        else renderWarehouseTable();
    });
});

function getSortFunction(field, dir) {
    return (a, b) => {
        let va, vb;
        switch (field) {
            case 'product':    va = a.product; vb = b.product; break;
            case 'warehouse':  va = a.warehouse || ''; vb = b.warehouse || ''; break;
            case 'qty':        va = a.sumQty; vb = b.sumQty; break;
            case 'share':      va = a.share; vb = b.share; break;
            case 'cumulative': va = a.cumulative; vb = b.cumulative; break;
            default:           va = a.sumQty; vb = b.sumQty;
        }
        if (typeof va === 'string') {
            return dir === 'asc' ? va.localeCompare(vb, 'ru') : vb.localeCompare(va, 'ru');
        }
        return dir === 'asc' ? va - vb : vb - va;
    };
}

// ── Export to Excel ────────────────────────
exportBtn.addEventListener('click', () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Global
    const globalData = applyFilters(globalABC, false).map((r, i) => ({
        '№': i + 1,
        'Товар': r.product,
        'Кол-во (SUM)': r.sumQty,
        'Доля, %': +(r.share * 100).toFixed(2),
        'Кумулятивная доля, %': +(r.cumulative * 100).toFixed(2),
        'Категория': r.category
    }));
    const ws1 = XLSX.utils.json_to_sheet(globalData);
    ws1['!cols'] = [
        { wch: 6 }, { wch: 50 }, { wch: 14 },
        { wch: 10 }, { wch: 20 }, { wch: 12 }
    ];
    XLSX.utils.book_append_sheet(wb, ws1, 'ABC_Общий');

    // Sheet 2: Per Warehouse
    const whData = applyFilters(warehouseABC, true).map((r, i) => ({
        '№': i + 1,
        'Товар': r.product,
        'Склад': r.warehouse,
        'Кол-во (SUM)': r.sumQty,
        'Доля, %': +(r.share * 100).toFixed(2),
        'Кумулятивная доля, %': +(r.cumulative * 100).toFixed(2),
        'Категория': r.category
    }));
    const ws2 = XLSX.utils.json_to_sheet(whData);
    ws2['!cols'] = [
        { wch: 6 }, { wch: 50 }, { wch: 35 },
        { wch: 14 }, { wch: 10 }, { wch: 20 }, { wch: 12 }
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'ABC_По_Складам');

    XLSX.writeFile(wb, 'ABC_Анализ_Списаний.xlsx');
});

// ── Helpers ────────────────────────────────
function escapeHtml(str) {
    const d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
}

function formatNumber(n) {
    if (n === undefined || n === null) return '0';
    return Number(n).toLocaleString('ru-RU', { maximumFractionDigits: 2 });
}

function truncate(str, maxLen) {
    if (!str) return '';
    return str.length > maxLen ? str.substring(0, maxLen) + '…' : str;
}
