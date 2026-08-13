// forecast.js
// 讀取 localStorage key "itemMonthlyUsage"（跟主頁「✅ 記錄這個月」寫入的是同一份資料），
// 結構：{ "2026-08": { "2026-08-00": { "FSG-3-01": {qty, meters, name}, ... } }, ... }
// （"2026-08-00" 代表整月合計的固定key，一個月一筆，不用逐日累積）
// 計算：建議月叫貨量 = 近3個月平均用量 × 70% ＋ 全部歷史月平均用量 × 30%
// 需要累積至少 1 個月的記錄資料才會顯示估算結果。

const ITEM_STORAGE_KEY = "itemMonthlyUsage";
const RECENT_MONTHS_COUNT = 3;
const RECENT_WEIGHT = 0.7;
const HISTORY_WEIGHT = 0.3;
const MIN_MONTHS_REQUIRED = 1;

function escapeHtml(text) {
  if (text == null) return "";
  return text
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function formatQtyInt(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return Math.round(num).toLocaleString("zh-TW");
}

function formatMeters(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return (Math.round(num * 1000) / 1000).toLocaleString("zh-TW");
}

function loadItemUsage() {
  try {
    const raw = localStorage.getItem(ITEM_STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.error("讀取品項用量歷史失敗：", e);
    return {};
  }
}

// ============================================================
// 以下的系列判斷邏輯跟 product.js 保持一致（同一套規則、同一份例外對照表），
// 目的是把同系列不同顏色/長度變體的物品編號合併成一個系列再加總用量，
// 這樣「FSG-3-035 這個系列總共要叫多少」才看得出正確答案，
// 不會被拆成 FSG-3-035、FSG-3-035B、FSG-3-035R...一堆分開的小數字。
// product.js 那邊如果有更新這套規則，這裡也要同步更新。
// ============================================================
function extractMmInfo(name) {
  if (!name) return { specMm: null, cutMm: null };
  let specMm = null;
  let cutMm = null;
  let m = /mm\s*[x*]\s*(\d+(?:\.\d+)?)/i.exec(name);
  if (m) {
    cutMm = parseFloat(m[1]);
  } else {
    m = /[x*]\s*(\d+(?:\.\d+)?)(?=\s*mm\b)/i.exec(name);
    if (m) cutMm = parseFloat(m[1]);
  }
  return { specMm, cutMm };
}

const SERIES_CODE_ALIAS_MAP = {
  "FSG-3-042-028M": "FSG-3-04",
  "H12": "H120",
  "H15CB": "H150CB",
  "HST-046": "HST-045",
  "ATM-012BK": "ATM-012",
  "FSG-2-017": "FSG-2-015",
  "H140CB": "H14CB",
  "H14": "H140",
  "FSG-3-035": "FSG-3-037",
  "HST-043": "HST-045",
  "HST-023": "HST-025",
};
const SERIES_NAME_ALIAS_MAP = {
  "矽套管 3.7mm* 38mm": "FSG-3-035",
};

function applyKnownSeriesAlias(row) {
  const code = (row.itemCode || "").toString().trim();
  if (code && SERIES_CODE_ALIAS_MAP[code]) return SERIES_CODE_ALIAS_MAP[code];
  const name = (row.name || "").toString().trim();
  if (name && SERIES_NAME_ALIAS_MAP[name]) return SERIES_NAME_ALIAS_MAP[name];
  return null;
}

function extractProductSeries(row) {
  const alias = applyKnownSeriesAlias(row);
  if (alias) return alias;

  let code = String(row.itemCode || "").trim();
  if (!code) return "未填寫物品編號";
  const name = (row.name || "").trim();

  if (/[0-9]+[Mm]$/.test(code)) {
    return code.replace(/[-_\s]*\d+(?:\.\d+)?(?:[Mm])?$/, "");
  }
  if (/[A-Za-z]M$/.test(code)) {
    code = code.replace(/M$/, "");
  }

  const segments = code.split(/[-_]/);
  if (segments.length > 1) {
    const last = segments[segments.length - 1];
    const lastNumMatch = last.match(/^(\d+(?:\.\d+)?)([A-Za-z]*)$/);

    if (lastNumMatch) {
      const numStr = lastNumMatch[1];
      const letters = lastNumMatch[2];
      let shouldCut = false;

      const info = extractMmInfo(name);
      if (info.cutMm != null && parseFloat(numStr) === info.cutMm) {
        shouldCut = true;
      }
      const explicitLengths = ['115', '053', '200', '530', '1000', '1650'];
      if (explicitLengths.includes(numStr)) {
        shouldCut = true;
      }
      if (letters === "" && numStr.length >= 3 && !numStr.startsWith("0")) {
        shouldCut = true;
      }

      if (shouldCut) {
        segments.pop();
        if (letters) {
          segments[segments.length - 1] += letters;
        }
        return segments.join("-");
      }
    }
  }
  return code;
}


// 把 { month: { 記錄key: { itemCode: {qty,meters,name} } } }
// 依「產品系列」（不是原始物品編號）彙總成 { month: { series: {qty,meters,name} } }。
// 記錄key可能是整月合計(YYYY-MM-00)，也可能是舊版逐日記錄留下的日期，
// 這裡不管是哪種都直接加總，兩種資料格式都能正確彙總。
function aggregateByMonth(all) {
  const monthly = {};

  Object.keys(all).forEach((month) => {
    const dates = all[month] || {};
    const dateKeys = Object.keys(dates);

    if (!monthly[month]) monthly[month] = {};

    dateKeys.forEach((date) => {
      const items = dates[date] || {};
      Object.keys(items).forEach((code) => {
        const rec = items[code];
        const series = extractProductSeries({ itemCode: code, name: rec.name || "" });
        if (!monthly[month][series]) {
          monthly[month][series] = { qty: 0, meters: 0, name: rec.name || "" };
        }
        monthly[month][series].qty += rec.qty || 0;
        monthly[month][series].meters += rec.meters || 0;
        if (rec.name) monthly[month][series].name = rec.name;
      });
    });
  });

  return { monthly };
}

function computeForecast(monthly) {
  const sortedMonths = Object.keys(monthly).sort(); // 升冪，最後面是最新月份
  const recentMonths = sortedMonths.slice(-RECENT_MONTHS_COUNT);
  const allMonths = sortedMonths;

  // 收集所有出現過的物品編號
  const allCodes = new Set();
  allMonths.forEach((m) => {
    Object.keys(monthly[m]).forEach((code) => allCodes.add(code));
  });

  const results = [];
  allCodes.forEach((code) => {
    let name = "";
    let recentQtySum = 0;
    let recentMetersSum = 0;
    recentMonths.forEach((m) => {
      const rec = monthly[m][code];
      if (rec) {
        recentQtySum += rec.qty;
        recentMetersSum += rec.meters;
        if (rec.name) name = rec.name;
      }
    });
    const recentQtyAvg = recentMonths.length ? recentQtySum / recentMonths.length : 0;
    const recentMetersAvg = recentMonths.length ? recentMetersSum / recentMonths.length : 0;

    let allQtySum = 0;
    let allMetersSum = 0;
    allMonths.forEach((m) => {
      const rec = monthly[m][code];
      if (rec) {
        allQtySum += rec.qty;
        allMetersSum += rec.meters;
        if (rec.name) name = rec.name;
      }
    });
    const allQtyAvg = allMonths.length ? allQtySum / allMonths.length : 0;
    const allMetersAvg = allMonths.length ? allMetersSum / allMonths.length : 0;

    const suggestedQty = recentQtyAvg * RECENT_WEIGHT + allQtyAvg * HISTORY_WEIGHT;
    const suggestedMeters = recentMetersAvg * RECENT_WEIGHT + allMetersAvg * HISTORY_WEIGHT;

    results.push({
      code,
      name,
      recentQtyAvg,
      recentMetersAvg,
      allQtyAvg,
      allMetersAvg,
      suggestedQty,
      suggestedMeters,
      monthsWithData: allMonths.filter((m) => monthly[m][code]).length,
    });
  });

  // 依建議叫貨量由高到低排序，量大的品項最需要優先關注
  results.sort((a, b) => b.suggestedQty - a.suggestedQty);
  return results;
}

let forecastResults = [];

function renderTable() {
  const tbody = document.getElementById("forecastTbody");
  tbody.innerHTML = "";
  forecastResults.forEach((r) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="border px-2 py-1">${escapeHtml(r.code)}</td>
      <td class="border px-2 py-1 text-slate-500 text-[10px]">${escapeHtml(r.name)}</td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(r.recentQtyAvg)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(r.recentMetersAvg)}</td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(r.allQtyAvg)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(r.allMetersAvg)}</td>
      <td class="border px-2 py-1 text-right font-semibold">${formatQtyInt(r.suggestedQty)}</td>
      <td class="border px-2 py-1 text-right font-semibold">${formatMeters(r.suggestedMeters)}</td>
      <td class="border px-2 py-1 text-right">${r.monthsWithData}</td>
    `;
    tbody.appendChild(tr);
  });

  const downloadBtn = document.getElementById("downloadForecastExcel");
  if (forecastResults.length) downloadBtn.classList.remove("hidden");
}

function downloadExcel() {
  if (!forecastResults.length) return;
  const roundM = (v) => Number((Math.round(v * 1000) / 1000).toFixed(3));
  const aoa = [
    ["產品系列", "品名", "近3月平均(數量)", "近3月平均(米數)", "全歷史平均(數量)", "全歷史平均(米數)", "建議月叫貨量(數量)", "建議月叫貨量(米數)", "已有月份數"],
    ...forecastResults.map((r) => [
      r.code,
      r.name,
      Math.round(r.recentQtyAvg),
      roundM(r.recentMetersAvg),
      Math.round(r.allQtyAvg),
      roundM(r.allMetersAvg),
      Math.round(r.suggestedQty),
      roundM(r.suggestedMeters),
      r.monthsWithData,
    ]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "叫貨量預估");
  XLSX.writeFile(wb, "叫貨量預估.xlsx");
}

document.addEventListener("DOMContentLoaded", () => {
  const statusEl = document.getElementById("status");
  const all = loadItemUsage();
  const { monthly } = aggregateByMonth(all);
  const monthCount = Object.keys(monthly).length;

  if (monthCount < MIN_MONTHS_REQUIRED) {
    statusEl.textContent =
      `目前還沒有已記錄的整月資料，請先到主頁上傳一整個月的銷貨明細並按「✅ 記錄這個月」，` +
      `累積至少 ${MIN_MONTHS_REQUIRED} 個月才會顯示估算結果（月數越多，近3月/全歷史的比較越準）。`;
    return;
  }

  forecastResults = computeForecast(monthly);

  if (!forecastResults.length) {
    statusEl.textContent = "已有記錄的月份，但目前沒有任何品項用量資料。";
    return;
  }

  statusEl.textContent =
    `已累積 ${monthCount} 個月的記錄資料，共 ${forecastResults.length} 個品項。` +
    `建議月叫貨量 = 近${RECENT_MONTHS_COUNT}個月平均 × ${RECENT_WEIGHT * 100}% ＋ 全歷史平均 × ${HISTORY_WEIGHT * 100}%` +
    (monthCount < RECENT_MONTHS_COUNT ? `（目前月數還不到${RECENT_MONTHS_COUNT}個月，近期平均會用現有的全部月份計算）。` : "。");

  renderTable();

  const downloadBtn = document.getElementById("downloadForecastExcel");
  downloadBtn.addEventListener("click", downloadExcel);
});
