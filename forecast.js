// forecast.js
// 讀取 localStorage key "itemMonthlyUsage"（跟主頁「✅ 記錄這批明細」寫入的是同一份資料），
// 結構：{ "2026-08": { "2026-08-10": { "FSG-3-01": {qty, meters, name}, ... }, ... }, ... }
// 計算：建議月叫貨量 = 近3個月平均用量 × 70% ＋ 全部歷史月平均用量 × 30%
// 需要累積至少 30 天的記錄資料才會顯示估算結果。

const ITEM_STORAGE_KEY = "itemMonthlyUsage";
const RECENT_MONTHS_COUNT = 3;
const RECENT_WEIGHT = 0.7;
const HISTORY_WEIGHT = 0.3;
const MIN_DAYS_REQUIRED = 30;

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

// 把 { month: { date: { itemCode: {qty,meters,name} } } }
// 彙總成 { month: { itemCode: {qty,meters,name} } }，並統計總天數。
function aggregateByMonth(all) {
  const monthly = {};
  let totalDays = 0;

  Object.keys(all).forEach((month) => {
    const dates = all[month] || {};
    const dateKeys = Object.keys(dates);
    totalDays += dateKeys.length;

    if (!monthly[month]) monthly[month] = {};

    dateKeys.forEach((date) => {
      const items = dates[date] || {};
      Object.keys(items).forEach((code) => {
        const rec = items[code];
        if (!monthly[month][code]) {
          monthly[month][code] = { qty: 0, meters: 0, name: rec.name || "" };
        }
        monthly[month][code].qty += rec.qty || 0;
        monthly[month][code].meters += rec.meters || 0;
        if (rec.name) monthly[month][code].name = rec.name;
      });
    });
  });

  return { monthly, totalDays };
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
      allQtyAvg,
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
      <td class="border px-2 py-1 text-right">${formatQtyInt(r.allQtyAvg)}</td>
      <td class="border px-2 py-1 text-right font-semibold">${formatQtyInt(r.suggestedQty)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(r.suggestedMeters)}</td>
      <td class="border px-2 py-1 text-right">${r.monthsWithData}</td>
    `;
    tbody.appendChild(tr);
  });

  const downloadBtn = document.getElementById("downloadForecastExcel");
  if (forecastResults.length) downloadBtn.classList.remove("hidden");
}

function downloadExcel() {
  if (!forecastResults.length) return;
  const aoa = [
    ["物品編號", "品名", "近3月平均(數量)", "全歷史平均(數量)", "建議月叫貨量(數量)", "建議月叫貨量(米數)", "已有月份數"],
    ...forecastResults.map((r) => [
      r.code,
      r.name,
      Math.round(r.recentQtyAvg),
      Math.round(r.allQtyAvg),
      Math.round(r.suggestedQty),
      Number((Math.round(r.suggestedMeters * 1000) / 1000).toFixed(3)),
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
  const { monthly, totalDays } = aggregateByMonth(all);

  if (totalDays < MIN_DAYS_REQUIRED) {
    statusEl.textContent =
      `目前累積了 ${totalDays} 天的記錄資料，還需要至少 ${MIN_DAYS_REQUIRED} 天才會顯示估算結果` +
      `（請持續在主頁上傳明細後按「✅ 記錄這批明細」，累積越多天資料越準）。`;
    return;
  }

  forecastResults = computeForecast(monthly);

  if (!forecastResults.length) {
    statusEl.textContent = "已有足夠天數的記錄，但目前沒有任何品項用量資料。";
    return;
  }

  const monthCount = Object.keys(monthly).length;
  statusEl.textContent =
    `已累積 ${totalDays} 天、涵蓋 ${monthCount} 個月的記錄資料，共 ${forecastResults.length} 個品項。` +
    `建議月叫貨量 = 近${RECENT_MONTHS_COUNT}個月平均 × ${RECENT_WEIGHT * 100}% ＋ 全歷史平均 × ${HISTORY_WEIGHT * 100}%。`;

  renderTable();

  const downloadBtn = document.getElementById("downloadForecastExcel");
  downloadBtn.addEventListener("click", downloadExcel);
});
