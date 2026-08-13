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
let currentSortKey = "suggestedQty";
let currentSortDir = "desc"; // "asc" | "desc"

function sortForecastResults() {
  const key = currentSortKey;
  const dir = currentSortDir === "asc" ? 1 : -1;
  const th = document.querySelector(`th[data-sort-key="${key}"]`);
  const type = th ? th.getAttribute("data-sort-type") : "num";

  forecastResults.sort((a, b) => {
    if (type === "text") {
      const av = (a[key] || "").toString();
      const bv = (b[key] || "").toString();
      return av.localeCompare(bv, "zh-Hant") * dir;
    }
    return ((a[key] || 0) - (b[key] || 0)) * dir;
  });
}

function updateSortArrows() {
  document.querySelectorAll("th[data-sort-key]").forEach((th) => {
    const arrow = th.querySelector(".sortArrow");
    if (!arrow) return;
    if (th.getAttribute("data-sort-key") === currentSortKey) {
      arrow.textContent = currentSortDir === "asc" ? "▲" : "▼";
    } else {
      arrow.textContent = "";
    }
  });
}

function setupSortableHeaders() {
  document.querySelectorAll("th[data-sort-key]").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.getAttribute("data-sort-key");
      if (currentSortKey === key) {
        currentSortDir = currentSortDir === "asc" ? "desc" : "asc";
      } else {
        currentSortKey = key;
        // 數量類欄位預設由大到小看比較直覺，品名/系列預設由小到大(字母排序)
        currentSortDir = th.getAttribute("data-sort-type") === "text" ? "asc" : "desc";
      }
      sortForecastResults();
      updateSortArrows();
      renderTable(true);
    });
  });
}

function renderTable(skipSort) {
  if (!skipSort) {
    sortForecastResults();
  }
  updateSortArrows();
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

function refreshDisplay() {
  const statusEl = document.getElementById("status");
  const all = loadItemUsage();
  const { monthly } = aggregateByMonth(all);
  const monthCount = Object.keys(monthly).length;

  if (monthCount < MIN_MONTHS_REQUIRED) {
    statusEl.textContent =
      `目前還沒有已記錄的整月資料，請先到主頁上傳一整個月的銷貨明細並按「✅ 記錄這個月」，` +
      `累積至少 ${MIN_MONTHS_REQUIRED} 個月才會顯示估算結果（月數越多，近3月/全歷史的比較越準）。`;
    forecastResults = [];
    const tbody = document.getElementById("forecastTbody");
    if (tbody) tbody.innerHTML = "";
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
}

// ============================================================
// 📥 匯入整年歷史資料：解析Excel，把年度總量平均拆成12個月，
// 灌進 itemMonthlyUsage。已經有真實記錄的月份不覆蓋，直接跳過。
//
// 支援兩種常見格式：
//   ① 「產品別統計」彙總格式：標題就在第一列，欄位是
//      產品系列/物品編號、品名、總計數量、總計米數。
//   ② 「客戶銷售明細表」原始明細格式：標題前面通常還有公司名稱、
//      報表標題、日期範圍等好幾列，欄位是物品編號、品名、銷貨量
//      （通常沒有米數欄位），資料按客戶分組，中間穿插「客戶名稱:」
//      跟「小計」這種不是真正品項的列，且同一個物品編號會在不同
//      客戶底下重複出現，需要把「銷貨量」全部加總起來才是年度總量。
// ============================================================
function findHeaderRow(rows) {
  const scanLimit = Math.min(rows.length, 30);
  for (let i = 0; i < scanLimit; i++) {
    // 先把每個儲存格轉成字串，並把所有空白（含全形空白）都拿掉再比對，
    // 因為有些報表的欄位標題中間會有排版用的空格，例如「品　名」，
    // 直接比對「品名」會抓不到，要先去空白才比對得到。
    const row = Array.from(rows[i] || []).map((c) =>
      c == null ? "" : c.toString().replace(/[\s\u3000]/g, "")
    );
    const codeIdx = row.findIndex((h) => h.includes("編號") || h.includes("系列"));
    const qtyIdx = row.findIndex((h) => (h.includes("數量") || h.includes("銷貨量")) && !h.includes("米"));
    if (codeIdx !== -1 && qtyIdx !== -1) {
      const nameIdx = row.findIndex((h) => h.includes("品名"));
      const meterIdx = row.findIndex((h) => h.includes("米數"));
      return { headerRowIndex: i, codeIdx, nameIdx, qtyIdx, meterIdx };
    }
  }
  return null;
}

function parseImportFile(file, callback) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (!rows.length) { callback(null, "檔案是空的。"); return; }

      const found = findHeaderRow(rows);
      if (!found) {
        callback(null, "找不到必要欄位（物品編號/產品系列、總計數量/銷貨量），請確認檔案格式。");
        return;
      }
      const { headerRowIndex, codeIdx, nameIdx, qtyIdx, meterIdx } = found;

      // 同一個物品編號可能在明細表裡分散出現很多次（不同客戶各買一些），
      // 這裡用累加的方式把同一個編號的銷貨量全部加總，才是年度總量。
      const items = {};
      for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row[codeIdx] == null || row[codeIdx] === "") continue;

        const codeRaw = String(row[codeIdx]).trim();
        // 跳過「客戶名稱:(CHxxx)xxx」這種分組標籤列（不是真正的品項資料）
        if (codeRaw.includes("客戶名稱")) continue;

        const name = nameIdx !== -1 ? String(row[nameIdx] || "").trim() : "";
        // 跳過「小計」列：物品編號欄位剛好落在小計文字，或品名欄位是「小計」
        if (codeRaw === "小計" || name === "小計") continue;
        // 跳過「合計」列：整份報表最後面的總加總列，同樣不是真正的品項，
        // 不擋掉的話會被當成一筆巨大的假品項，混進所有平均值計算裡。
        if (codeRaw === "合計" || name === "合計") continue;

        const qty = parseFloat(row[qtyIdx]) || 0;
        // 這種明細表通常沒有米數欄位，沒有的話用數量頂替（跟系統其他地方一致的作法）
        const meters = meterIdx !== -1 ? parseFloat(row[meterIdx]) || 0 : qty;

        if (!items[codeRaw]) {
          items[codeRaw] = { qty: 0, meters: 0, name };
        }
        items[codeRaw].qty += qty;
        items[codeRaw].meters += meters;
        if (name) items[codeRaw].name = name;
      }

      if (!Object.keys(items).length) {
        callback(null, "沒有解析到任何品項資料，請確認檔案內容。");
        return;
      }
      callback(items, null);
    } catch (err) {
      console.error(err);
      callback(null, "讀取檔案時發生錯誤，請確認是正確的 Excel 檔案。");
    }
  };
  reader.readAsArrayBuffer(file);
}

document.addEventListener("DOMContentLoaded", () => {
  setupSortableHeaders();
  refreshDisplay();

  const downloadBtn = document.getElementById("downloadForecastExcel");
  if (downloadBtn) downloadBtn.addEventListener("click", downloadExcel);

  const importBtn = document.getElementById("importBtn");
  if (importBtn) {
    importBtn.addEventListener("click", () => {
      const fileInput = document.getElementById("importFile");
      const yearInput = document.getElementById("importYear");
      const importStatus = document.getElementById("importStatus");

      const file = fileInput.files && fileInput.files[0];
      const year = parseInt(yearInput.value, 10);

      if (!file) {
        importStatus.textContent = "請先選擇檔案。";
        return;
      }
      if (!Number.isFinite(year) || year < 2000 || year > 2100) {
        importStatus.textContent = "請輸入有效的年份，例如 2025。";
        return;
      }

      importStatus.textContent = "匯入中…";

      parseImportFile(file, (items, err) => {
        if (err) {
          importStatus.textContent = err;
          return;
        }

        const all = loadItemUsage();
        let importedMonths = 0;
        let skippedMonths = 0;

        for (let m = 1; m <= 12; m++) {
          const month = `${year}-${String(m).padStart(2, "0")}`;
          if (all[month] && Object.keys(all[month]).length) {
            skippedMonths++;
            continue; // 已有真實記錄的月份，不覆蓋
          }
          const key = `${month}-00`;
          const monthItems = {};
          Object.keys(items).forEach((code) => {
            monthItems[code] = {
              qty: items[code].qty / 12,
              meters: items[code].meters / 12,
              name: items[code].name,
            };
          });
          all[month] = { [key]: monthItems };
          importedMonths++;
        }

        try {
          localStorage.setItem(ITEM_STORAGE_KEY, JSON.stringify(all));
        } catch (e) {
          importStatus.textContent = "儲存失敗，可能是瀏覽器儲存空間不足。";
          return;
        }

        importStatus.textContent =
          `匯入完成：${year}年共 ${Object.keys(items).length} 個品項，已灌入 ${importedMonths} 個月` +
          (skippedMonths ? `（${skippedMonths} 個月已有真實記錄資料，已跳過保留）。` : "。");

        refreshDisplay();
      });
    });
  }
});
