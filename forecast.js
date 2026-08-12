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

// 把 { month: { 記錄key: { itemCode: {qty,meters,name} } } }
// 彙總成 { month: { itemCode: {qty,meters,name} } }。
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
        if (!monthly[month][code]) {
          monthly[month][code] = { qty: 0, meters: 0, name: rec.name || "" };
        }
        monthly[month][code].qty += rec.qty || 0;
        monthly[month][code].meters += rec.meters || 0;
        if (rec.name) monthly[month][code].name = rec.name;
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
