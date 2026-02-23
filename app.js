// ============================================
// Constants
// ============================================
const CORS_PROXY = 'https://corsproxy.io/?';
const YAHOO_API_BASE = 'https://query1.finance.yahoo.com/v8/finance/chart/';
const BATCH_SIZE = 5;
const BATCH_DELAY_MS = 1500;

// XLS の列インデックス (data_j.xls 形式)
const COL = {
    DATE: 0,       // 日付
    CODE: 1,       // コード
    NAME: 2,       // 銘柄名
    MARKET: 3,     // 市場・商品区分
    SEC33_CODE: 4, // 33業種コード
    SEC33_NAME: 5, // 33業種区分
    SEC17_CODE: 6, // 17業種コード
    SEC17_NAME: 7, // 17業種区分
    SCALE_CODE: 8, // 規模コード
    SCALE_NAME: 9, // 規模区分
};

// ============================================
// State
// ============================================
let xlsData = null;        // { header: string[], rows: any[][] }
let closingPrices = {};    // { stockCode: { price, change1d, change7d, change14d, change30d, change180d, error } }
let errorMessages = [];
let isFetching = false;
let sortColIdx = -1;
let sortAsc = true;
let priceTargetDate = null; // 終値取得に使用した日付 (YYYYMMDD 文字列)

// ============================================
// DOM References
// ============================================
const $ = (id) => document.getElementById(id);

const dom = {
    dropZone: $('dropZone'),
    fileInput: $('fileInput'),
    fileInfo: $('fileInfo'),
    fileName: $('fileName'),
    fileSize: $('fileSize'),
    rowCount: $('rowCount'),
    fileRemove: $('fileRemove'),
    actionSection: $('actionSection'),
    fetchBtn: $('fetchBtn'),
    downloadBtn: $('downloadBtn'),
    progressSection: $('progressSection'),
    progressFill: $('progressFill'),
    progressLabel: $('progressLabel'),
    progressValue: $('progressValue'),
    progressDetail: $('progressDetail'),
    resultsSection: $('resultsSection'),
    successCount: $('successCount'),
    naCount: $('naCount'),
    errorCount: $('errorCount'),
    errorSection: $('errorSection'),
    errorLog: $('errorLog'),
    tableSection: $('tableSection'),
    tableRowCount: $('tableRowCount'),
    tableHead: $('tableHead'),
    tableBody: $('tableBody'),
};

// ============================================
// XLS Parsing (SheetJS)
// ============================================

/**
 * XLS/XLSX ファイルを解析して { header, rows } を返す
 */
function parseXLS(arrayBuffer) {
    const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // JSON 形式で取得（ヘッダー付き）
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (jsonData.length < 2) {
        throw new Error('データが不足しています（ヘッダー＋最低1行のデータが必要です）。');
    }

    const header = jsonData[0].map(h => String(h).trim());
    const rows = jsonData.slice(1).filter(row => row.some(cell => cell !== ''));

    return { header, rows };
}

// ============================================
// Stock Code Utilities
// ============================================

/**
 * コード列から Yahoo Finance 用ティッカーを生成
 * 例: 1301 → "1301.T"
 */
function toTicker(rawCode) {
    const code = String(rawCode).trim();
    if (code.length < 4) return null;

    // 先頭4文字を取得
    const ticker = code.substring(0, 4);

    // 英数字のみか確認
    if (!/^[A-Za-z0-9]+$/.test(ticker)) return null;

    return ticker + '.T';
}

/**
 * データからユニークな銘柄コード一覧を取得
 */
function getUniqueStocks(rows) {
    const seen = new Map();

    for (const row of rows) {
        if (row.length <= COL.CODE) continue;
        const rawCode = String(row[COL.CODE]).trim();
        if (!rawCode || seen.has(rawCode)) continue;

        const ticker = toTicker(rawCode);
        if (!ticker) continue;

        const name = String(row[COL.NAME] || '').trim();

        seen.set(rawCode, { ticker, rawCode, name });
    }

    return Array.from(seen.values());
}

/**
 * 終値取得に使う日付を決定する
 * - 東京市場の取引時間（9:00〜15:00 JST）中は前日の日付
 * - それ以外は当日の日付
 * 戻り値: YYYYMMDD 形式の文字列
 */
function getTargetDate() {
    const now = new Date();
    // JST は UTC+9
    const jstOffset = 9 * 60;
    const utcMs = now.getTime() + now.getTimezoneOffset() * 60 * 1000;
    const jstDate = new Date(utcMs + jstOffset * 60 * 1000);

    const hour = jstDate.getHours();
    const minute = jstDate.getMinutes();
    const currentMinutes = hour * 60 + minute;

    // 9:00 (540分) 〜 15:00 (900分) の間は前日
    let targetDate = jstDate;
    if (currentMinutes >= 540 && currentMinutes < 900) {
        targetDate = new Date(jstDate.getTime() - 24 * 60 * 60 * 1000);
    }

    const y = targetDate.getFullYear();
    const m = String(targetDate.getMonth() + 1).padStart(2, '0');
    const d = String(targetDate.getDate()).padStart(2, '0');
    return `${y}${m}${d}`;
}

/**
 * YYYYMMDD を YYYY/MM/DD に変換
 */
function formatDateStr(yyyymmdd) {
    const s = String(yyyymmdd);
    return `${s.substring(0, 4)}/${s.substring(4, 6)}/${s.substring(6, 8)}`;
}

// ============================================
// Yahoo Finance API
// ============================================

/**
 * タイムスタンプ配列から、指定タイムスタンプに最も近い（以前の）終値を探す
 */
function findClosestPrice(timestamps, closes, targetTs) {
    let bestIdx = -1;
    let bestDiff = Infinity;

    // 対象日以前で最も近いデータを優先
    for (let i = 0; i < timestamps.length; i++) {
        if (closes[i] === null || closes[i] === undefined) continue;
        const diff = targetTs - timestamps[i];
        if (diff >= 0 && diff < bestDiff) {
            bestDiff = diff;
            bestIdx = i;
        }
    }

    // 対象日以前にない場合、以降の最も近いデータ
    if (bestIdx === -1) {
        for (let i = 0; i < timestamps.length; i++) {
            if (closes[i] === null || closes[i] === undefined) continue;
            const diff = Math.abs(targetTs - timestamps[i]);
            if (diff < bestDiff) {
                bestDiff = diff;
                bestIdx = i;
            }
        }
    }

    return bestIdx >= 0 ? closes[bestIdx] : null;
}

/**
 * 変動率を計算 (%)：(現在値 - 過去値) / 過去値 * 100
 */
function calcChangeRate(currentPrice, pastPrice) {
    if (currentPrice === null || pastPrice === null || pastPrice === 0) return null;
    return Math.round((currentPrice - pastPrice) / pastPrice * 10000) / 100;
}

/**
 * 日付文字列（YYYYMMDD 数値）を Date に変換
 */
function parseDateValue(dateVal) {
    const s = String(dateVal).trim();
    // YYYYMMDD 形式
    if (/^\d{8}$/.test(s)) {
        const y = parseInt(s.substring(0, 4));
        const m = parseInt(s.substring(4, 6)) - 1;
        const d = parseInt(s.substring(6, 8));
        return new Date(y, m, d);
    }
    // 数値（Excel シリアル日付ではなくそのまま YYYYMMDD として格納されることを想定）
    const n = parseInt(s);
    if (n > 19000000 && n < 21000000) {
        const y = Math.floor(n / 10000);
        const m = Math.floor((n % 10000) / 100) - 1;
        const d = n % 100;
        return new Date(y, m, d);
    }
    // フォールバック
    return new Date(s);
}

/**
 * 指定ティッカーの終値・株価変動率を取得
 */
async function fetchClosingPrice(ticker, targetDateStr) {
    const nullResult = {
        price: null,
        actualDate: null,
        change1d: null,
        change7d: null,
        change14d: null,
        change30d: null,
        change180d: null,
        error: null
    };

    if (!ticker) return { ...nullResult, error: '無効なティッカー' };

    try {
        const targetDate = parseDateValue(targetDateStr);
        const targetTs = Math.floor(targetDate.getTime() / 1000);

        // 6ヶ月比のため200日前まで取得
        const startTs = targetTs - 200 * 86400;
        const endTs = targetTs + 14 * 86400;

        const apiUrl = `${YAHOO_API_BASE}${encodeURIComponent(ticker)}?period1=${startTs}&period2=${endTs}&interval=1d`;
        const proxyUrl = `${CORS_PROXY}${encodeURIComponent(apiUrl)}`;

        const response = await fetch(proxyUrl, {
            headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) {
            return { ...nullResult, error: `HTTP ${response.status}` };
        }

        const data = await response.json();

        if (!data.chart || !data.chart.result || data.chart.result.length === 0) {
            return { ...nullResult, error: 'データなし' };
        }

        const result = data.chart.result[0];
        const timestamps = result.timestamp || [];
        const closes = result.indicators?.quote?.[0]?.close || [];

        if (timestamps.length === 0 || closes.length === 0) {
            return { ...nullResult, error: 'チャートデータなし' };
        }

        // 有効な取引日データのみを抽出（タイムスタンプ昇順）
        const tradingDays = [];
        for (let i = 0; i < timestamps.length; i++) {
            if (closes[i] !== null && closes[i] !== undefined) {
                tradingDays.push({ ts: timestamps[i], close: closes[i] });
            }
        }
        tradingDays.sort((a, b) => a.ts - b.ts);

        if (tradingDays.length === 0) {
            return { ...nullResult, error: '有効な終値なし' };
        }

        // 当日（targetTs）以前で最も近い取引日を探す
        let currentIdx = -1;
        for (let i = tradingDays.length - 1; i >= 0; i--) {
            if (tradingDays[i].ts <= targetTs) {
                currentIdx = i;
                break;
            }
        }
        // 対象日以前になければ最も近い取引日
        if (currentIdx === -1) currentIdx = 0;

        const currentPrice = tradingDays[currentIdx].close;
        const actualTs = tradingDays[currentIdx].ts;

        const actualDate = new Date(actualTs * 1000);
        const formattedDate = `${actualDate.getFullYear()}/${String(actualDate.getMonth() + 1).padStart(2, '0')}/${String(actualDate.getDate()).padStart(2, '0')}`;

        // 前日比: 実際の1つ前の取引日と比較（カレンダー日付ではなく取引日ベース）
        const price1d = currentIdx >= 1 ? tradingDays[currentIdx - 1].close : null;

        // 1週間〜6ヶ月比: 実際の取引日のタイムスタンプから期間を引いて最も近い取引日と比較
        const price7d = findClosestPrice(timestamps, closes, actualTs - 7 * 86400);
        const price14d = findClosestPrice(timestamps, closes, actualTs - 14 * 86400);
        const price30d = findClosestPrice(timestamps, closes, actualTs - 30 * 86400);
        const price180d = findClosestPrice(timestamps, closes, actualTs - 180 * 86400);

        // 変動率を計算
        const change1d = calcChangeRate(currentPrice, price1d);
        const change7d = calcChangeRate(currentPrice, price7d);
        const change14d = calcChangeRate(currentPrice, price14d);
        const change30d = calcChangeRate(currentPrice, price30d);
        const change180d = calcChangeRate(currentPrice, price180d);

        return {
            price: Math.round(currentPrice * 10) / 10,
            actualDate: formattedDate,
            change1d,
            change7d,
            change14d,
            change30d,
            change180d,
            error: null
        };
    } catch (err) {
        return { ...nullResult, error: err.message };
    }
}

// ============================================
// Batch Processing
// ============================================

async function fetchAllPrices(stocks) {
    isFetching = true;
    closingPrices = {};
    errorMessages = [];

    // 最新の取引日の終値を取得（市場開場中は前日）
    const targetDateStr = getTargetDate();
    priceTargetDate = targetDateStr;

    const total = stocks.length;
    let completed = 0;

    showSection(dom.progressSection, true);
    updateProgress(0, total, `基準日: ${formatDateStr(targetDateStr)} — 準備中...`);

    dom.fetchBtn.disabled = true;
    dom.fetchBtn.classList.add('loading');

    for (let i = 0; i < total; i += BATCH_SIZE) {
        const batch = stocks.slice(i, i + BATCH_SIZE);

        const promises = batch.map(async (stock) => {
            const result = await fetchClosingPrice(stock.ticker, targetDateStr);
            closingPrices[stock.rawCode] = result;

            if (result.error) {
                errorMessages.push({
                    code: stock.rawCode,
                    name: stock.name,
                    ticker: stock.ticker || 'N/A',
                    error: result.error
                });
            }

            completed++;
            updateProgress(completed, total, `${stock.name} (${stock.ticker}) を取得中...`);
        });

        await Promise.all(promises);

        // レート制限対策
        if (i + BATCH_SIZE < total) {
            await sleep(BATCH_DELAY_MS);
        }
    }

    isFetching = false;
    dom.fetchBtn.disabled = false;
    dom.fetchBtn.classList.remove('loading');

    updateProgress(total, total, '完了！');
    showResults();
    renderTable();
    dom.downloadBtn.disabled = false;
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ============================================
// UI Helpers
// ============================================

function showSection(el, show = true) {
    el.style.display = show ? '' : 'none';
    if (show) {
        el.style.animation = 'none';
        el.offsetHeight; // reflow
        el.style.animation = '';
    }
}

function updateProgress(current, total, detail) {
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;
    dom.progressFill.style.width = pct + '%';
    dom.progressValue.textContent = pct + '%';
    dom.progressLabel.textContent = `${current} / ${total} 銘柄`;
    dom.progressDetail.textContent = detail;
}

function showResults() {
    const successN = Object.values(closingPrices).filter(v => v.price !== null).length;
    const errorN = errorMessages.length;
    const naN = Object.values(closingPrices).filter(v => v.price === null).length;

    dom.successCount.textContent = successN;
    dom.naCount.textContent = naN;
    dom.errorCount.textContent = errorN;

    showSection(dom.resultsSection);

    if (errorMessages.length > 0) {
        dom.errorLog.innerHTML = errorMessages.map(e =>
            `<div class="error-entry"><span class="error-code">${escapeHTML(e.code)}</span> ${escapeHTML(e.name)} (${escapeHTML(e.ticker)}) — ${escapeHTML(e.error)}</div>`
        ).join('');
        showSection(dom.errorSection);
    } else {
        showSection(dom.errorSection, false);
    }
}

function escapeHTML(str) {
    const div = document.createElement('div');
    div.textContent = String(str);
    return div.innerHTML;
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
}

// ============================================
// Table Rendering
// ============================================

// 表に表示するXLS列
const DISPLAY_COLS = [
    { key: COL.DATE, label: '日付' },
    { key: COL.CODE, label: 'コード' },
    { key: COL.NAME, label: '銘柄名' },
    { key: COL.MARKET, label: '市場・商品区分' },
    { key: COL.SEC33_NAME, label: '33業種区分' },
    { key: COL.SEC17_NAME, label: '17業種区分' },
    { key: COL.SCALE_NAME, label: '規模区分' },
];

// 追加列（価格取得後に表示）
const EXTRA_COLS = [
    { key: 'price', label: '終値' },
    { key: 'change1d', label: '前日比(%)' },
    { key: 'change7d', label: '1週間比(%)' },
    { key: 'change14d', label: '2週間比(%)' },
    { key: 'change30d', label: '1ヶ月比(%)' },
    { key: 'change180d', label: '6ヶ月比(%)' },
];

function renderTable() {
    if (!xlsData) return;

    const { rows } = xlsData;
    const hasPrices = Object.keys(closingPrices).length > 0;

    // ヘッダー構築
    const allCols = [...DISPLAY_COLS];
    if (hasPrices) {
        allCols.push(...EXTRA_COLS.map((c, i) => ({
            ...c,
            colIdx: DISPLAY_COLS.length + i,
            isExtra: true
        })));
    }

    dom.tableHead.innerHTML = '<tr>' + allCols.map((col, idx) => {
        let indicator = '';
        if (idx === sortColIdx) {
            indicator = sortAsc ? ' ▲' : ' ▼';
        }
        return `<th data-col="${idx}">${escapeHTML(col.label)}<span class="sort-indicator">${indicator}</span></th>`;
    }).join('') + '</tr>';

    // ヘッダークリックイベント
    dom.tableHead.querySelectorAll('th').forEach(th => {
        th.addEventListener('click', () => {
            const colIdx = parseInt(th.dataset.col);
            if (sortColIdx === colIdx) {
                sortAsc = !sortAsc;
            } else {
                sortColIdx = colIdx;
                sortAsc = true;
            }
            renderTable();
        });
    });

    // ソート
    const sortedRows = getSortedRows(rows, hasPrices);

    // 最大200行表示
    const displayRows = sortedRows.slice(0, 200);
    dom.tableBody.innerHTML = displayRows.map(row => {
        const cells = DISPLAY_COLS.map(col => {
            let val = row[col.key];
            if (col.key === COL.DATE) {
                // YYYYMMDD を YYYY/MM/DD に変換
                const s = String(val).trim();
                if (/^\d{8}$/.test(s)) {
                    val = `${s.substring(0, 4)}/${s.substring(4, 6)}/${s.substring(6, 8)}`;
                }
            }
            return `<td>${escapeHTML(String(val ?? ''))}</td>`;
        });

        if (hasPrices) {
            const code = String(row[COL.CODE] || '').trim();
            const pd = closingPrices[code];

            // 終値
            if (pd && pd.price !== null) {
                cells.push(`<td class="price-cell has-price">${pd.price.toLocaleString()}</td>`);
            } else {
                cells.push(`<td class="price-cell no-price">N/A</td>`);
            }

            // 変動率列
            for (const key of ['change1d', 'change7d', 'change14d', 'change30d', 'change180d']) {
                const val = pd?.[key];
                if (val !== null && val !== undefined) {
                    const sign = val > 0 ? '+' : '';
                    const colorClass = val > 0 ? 'change-up' : val < 0 ? 'change-down' : '';
                    cells.push(`<td class="price-cell has-price ${colorClass}">${sign}${val.toFixed(2)}%</td>`);
                } else {
                    cells.push(`<td class="price-cell no-price">N/A</td>`);
                }
            }
        }

        return '<tr>' + cells.join('') + '</tr>';
    }).join('');

    dom.tableRowCount.textContent = `${rows.length} 行${rows.length > 200 ? '（200行まで表示）' : ''}`;
    showSection(dom.tableSection);
}

function getSortedRows(rows, hasPrices) {
    if (sortColIdx < 0) return rows;

    const sorted = [...rows];
    const baseCount = DISPLAY_COLS.length;

    sorted.sort((a, b) => {
        let valA, valB;

        if (sortColIdx < baseCount) {
            // XLS 列
            const colKey = DISPLAY_COLS[sortColIdx].key;
            valA = String(a[colKey] ?? '').trim();
            valB = String(b[colKey] ?? '').trim();
        } else if (hasPrices) {
            // 追加列
            const codeA = String(a[COL.CODE] || '').trim();
            const codeB = String(b[COL.CODE] || '').trim();
            const pdA = closingPrices[codeA];
            const pdB = closingPrices[codeB];

            const extraIdx = sortColIdx - baseCount;
            const extraKey = EXTRA_COLS[extraIdx]?.key;

            if (extraKey) {
                valA = pdA?.[extraKey] ?? null;
                valB = pdB?.[extraKey] ?? null;
            }

            // null は常に末尾
            if (valA === null && valB === null) return 0;
            if (valA === null) return 1;
            if (valB === null) return -1;
            return sortAsc ? valA - valB : valB - valA;
        }

        // 文字列比較（数値として解釈可能なら数値比較）
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
            return sortAsc ? numA - numB : numB - numA;
        }

        const cmp = String(valA).localeCompare(String(valB), 'ja');
        return sortAsc ? cmp : -cmp;
    });

    return sorted;
}

// ============================================
// CSV Export
// ============================================

function generateOutputCSV() {
    if (!xlsData) return '';

    const { header, rows } = xlsData;
    const lines = [];

    // ヘッダー行
    const csvHeader = [
        ...DISPLAY_COLS.map(c => c.label),
        ...EXTRA_COLS.map(c => c.label)
    ];
    lines.push(csvHeader.join(','));

    // データ行
    for (const row of rows) {
        const cells = DISPLAY_COLS.map(col => {
            let val = String(row[col.key] ?? '').trim();
            if (col.key === COL.DATE) {
                const s = val;
                if (/^\d{8}$/.test(s)) {
                    val = `${s.substring(0, 4)}/${s.substring(4, 6)}/${s.substring(6, 8)}`;
                }
            }
            // カンマや引用符を含む場合は引用符で囲む
            if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                return '"' + val.replace(/"/g, '""') + '"';
            }
            return val;
        });

        const code = String(row[COL.CODE] || '').trim();
        const pd = closingPrices[code];

        // 終値
        cells.push(pd && pd.price !== null ? String(pd.price) : 'N/A');

        // 変動率
        for (const key of ['change1d', 'change7d', 'change14d', 'change30d', 'change180d']) {
            const val = pd?.[key];
            cells.push(val !== null && val !== undefined ? val.toFixed(2) : 'N/A');
        }

        lines.push(cells.join(','));
    }

    return lines.join('\r\n');
}

function downloadCSV() {
    const csvContent = generateOutputCSV();
    if (!csvContent) return;

    // BOM付きUTF-8
    const bom = '\uFEFF';
    const blob = new Blob([bom + csvContent], { type: 'text/csv;charset=utf-8;' });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;

    const now = new Date();
    const ts = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    a.download = `kabuka_data_${ts}.csv`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ============================================
// File Handling
// ============================================

function handleFile(file) {
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xls', 'xlsx'].includes(ext)) {
        alert('XLS または XLSX ファイルを選択してください。');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            xlsData = parseXLS(e.target.result);
            closingPrices = {};
            errorMessages = [];
            sortColIdx = -1;
            sortAsc = true;

            // ファイル情報を表示
            dom.fileName.textContent = file.name;
            dom.fileSize.textContent = formatFileSize(file.size);
            dom.rowCount.textContent = xlsData.rows.length;

            showSection(dom.fileInfo, true);
            dom.dropZone.style.display = 'none';

            showSection(dom.actionSection, true);
            dom.downloadBtn.disabled = true;

            // テーブルをすぐに表示
            renderTable();

            // 結果・エラー・進捗を非表示
            showSection(dom.resultsSection, false);
            showSection(dom.errorSection, false);
            showSection(dom.progressSection, false);
        } catch (err) {
            alert('ファイルの読み込みに失敗しました: ' + err.message);
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

function resetFile() {
    xlsData = null;
    closingPrices = {};
    errorMessages = [];
    sortColIdx = -1;
    sortAsc = true;

    dom.dropZone.style.display = '';
    showSection(dom.fileInfo, false);
    showSection(dom.actionSection, false);
    showSection(dom.progressSection, false);
    showSection(dom.resultsSection, false);
    showSection(dom.errorSection, false);
    showSection(dom.tableSection, false);

    dom.fileInput.value = '';
    dom.downloadBtn.disabled = true;
}

// ============================================
// Event Listeners
// ============================================

// ドラッグ＆ドロップ
dom.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dom.dropZone.classList.add('dragover');
});

dom.dropZone.addEventListener('dragleave', () => {
    dom.dropZone.classList.remove('dragover');
});

dom.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dom.dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    handleFile(file);
});

dom.dropZone.addEventListener('click', () => {
    dom.fileInput.click();
});

// ファイル選択
dom.fileInput.addEventListener('change', (e) => {
    handleFile(e.target.files[0]);
});

// ファイル削除
dom.fileRemove.addEventListener('click', resetFile);

// 終値取得ボタン
dom.fetchBtn.addEventListener('click', async () => {
    if (!xlsData || isFetching) return;

    const stocks = getUniqueStocks(xlsData.rows);
    if (stocks.length === 0) {
        alert('有効な銘柄コードが見つかりません。');
        return;
    }

    await fetchAllPrices(stocks);
});

// CSVダウンロードボタン
dom.downloadBtn.addEventListener('click', downloadCSV);
