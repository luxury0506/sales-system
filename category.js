function escapeHtml(text) {
  if (text == null) return "";
  return text
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function formatMoney(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return Math.round(num).toLocaleString("zh-TW");
}

function formatQtyInt(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return Math.round(num);
}

function formatMeters(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return (Math.round(num * 1000) / 1000).toLocaleString("zh-TW");
}

function formatPercent(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return (num * 100).toFixed(1) + "%";
}

const statusEl = document.getElementById("status");
const categoryTbody = document.getElementById("categoryTbody");
const downloadBtn = document.getElementById("downloadCategoryExcel");
let categorySummary = [];

function loadFromLocalStorage() {
  const raw = localStorage.getItem("salesAnalysisData");
  if (!raw) {
    statusEl.textContent = "找不到分析資料，請先回主頁上傳銷售檔並完成計算。";
    return null;
  }

  try {
    const payload = JSON.parse(raw);
    if (!payload.processedRows || !Array.isArray(payload.processedRows)) {
      statusEl.textContent = "資料格式不正確，請在主頁重新執行一次計算。";
      return null;
    }
    statusEl.textContent =
      "已載入主頁分析結果。" +
      (payload.savedAt ? `（最後更新：${new Date(payload.savedAt).toLocaleString()}）` : "");
    return payload.processedRows;
  } catch (e) {
    console.error(e);
    statusEl.textContent = "讀取分析資料時發生錯誤。";
    return null;
  }
}

// 從物品編號抓出「大分類」，例如：
//   FSG-3-03B   -> FSG-3   (三段以上：取前兩段「系列-家族碼」)
//   CFT-6-3/8A  -> CFT-6
//   SRG-015     -> SRG     (兩段：取第一段，因為第二段是規格號不是家族碼)
//   HST-01      -> HST
//   H025C       -> H       (沒有連字號：只取開頭的英文字母)
//   PDK0001323  -> PDK
//   200011000282-> 未分類(其他)
function extractCategory(row) {
  const code = String(row.itemCode || "").trim();
  if (!code) return "未填寫物品編號";

  const segments = code.split(/[-_]/);

  // 第二段若剛好是「單一數字」(3、2、6...)，代表這是系列的家族碼，
  // 例如 FSG-3-03B / FSG-2-01BL / CFT-3-5A / CFT-6-3/8A，
  // 分類要保留到「FSG-3」「CFT-6」這一層。
  if (segments.length >= 3 && /^\d$/.test(segments[1])) {
    return `${segments[0]}-${segments[1]}`;
  }
  // 其餘情況（含只有兩段的，如 SRG-01、HST-01、SR-01；
  // 或第二段是完整規格號如 HST-046-050-T01、SR-025-030C-T08），
  // 第二段其實是規格/長度號，不是家族碼，只取第一段當分類。
  if (segments.length >= 2) return segments[0];
  // 完全沒有連字號，例如 H03、H025C、PDK0001323
  const m = code.match(/^[A-Za-z]+/);
  if (m) return m[0];
  return "未分類(其他)";
}

function buildCategorySummary(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const key = extractCategory(row);

    if (!map.has(key)) {
      map.set(key, {
        category: key,
        codes: new Set(), // 收集這個分類底下涵蓋哪些物品編號/產品系列
        totalQty: 0,
        totalMeters: 0,
        totalAmount: 0,
        totalCost: 0,
        totalProfit: 0,
      });
    }

    const agg = map.get(key);
    if (row.itemCode) agg.codes.add(String(row.itemCode).trim());
    agg.totalQty += Number(row.qty) || 0;
    agg.totalMeters += Number(row.meters) || 0;
    agg.totalAmount += Number(row.amount) || 0;
    agg.totalCost += Number(row.cost) || 0;
    agg.totalProfit += Number(row.profit) || 0;
  });

  const list = Array.from(map.values()).map((x) => {
    const marginRate = x.totalAmount > 0 ? x.totalProfit / x.totalAmount : 0;
    const codesArray = Array.from(x.codes).sort();
    const isTruncated = codesArray.length > 6;
    const shortCodes = isTruncated ? codesArray.slice(0, 6).join("、 ") : codesArray.join("、 ");
    const allCodesText = codesArray.join("、 ");
    return {
      ...x,
      marginRate,
      codeCount: codesArray.length,
      displayCodes: allCodesText, // 匯出 Excel 用，維持完整清單
      shortCodesText: shortCodes,
      allCodesText,
      isTruncated,
    };
  });

  // 依銷售額由高到低排序
  list.sort((a, b) => b.totalAmount - a.totalAmount);
  return list;
}

// 渲染「涵蓋物品編號參考」這一格：預設精簡顯示，超過6個編號時
// 附上「顯示全部」按鈕，點擊可展開完整清單／再點一次收合。
function renderCodesCell(cell, c, expanded) {
  if (!cell) return;
  if (!c.isTruncated) {
    cell.innerHTML = `<span>${escapeHtml(c.allCodesText)}</span>`;
    return;
  }
  if (expanded) {
    cell.innerHTML =
      `<span>${escapeHtml(c.allCodesText)}</span> ` +
      `<button type="button" class="toggleCodesBtn text-blue-600 hover:underline whitespace-nowrap">收合 ▲</button>`;
  } else {
    cell.innerHTML =
      `<span>${escapeHtml(c.shortCodesText)}...</span> ` +
      `<button type="button" class="toggleCodesBtn text-blue-600 hover:underline whitespace-nowrap">顯示全部 ${c.codeCount} 個 ▼</button>`;
  }
  const btn = cell.querySelector(".toggleCodesBtn");
  btn.addEventListener("click", () => renderCodesCell(cell, c, !expanded));
}

function renderTable() {
  categoryTbody.innerHTML = "";
  categorySummary.forEach((c, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="border px-2 py-1 font-medium">${escapeHtml(c.category)}</td>
      <td class="border px-2 py-1 text-right">${c.codeCount}</td>
      <td class="border px-2 py-1 text-slate-500 text-[10px] break-words max-w-[220px]" data-idx="${idx}"></td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(c.totalQty)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(c.totalMeters)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalAmount)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalCost)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalProfit)}</td>
      <td class="border px-2 py-1 text-right">${formatPercent(c.marginRate)}</td>
    `;
    categoryTbody.appendChild(tr);
    renderCodesCell(tr.querySelector(`td[data-idx="${idx}"]`), c, false);
  });

  // 合計列
  if (categorySummary.length) {
    const total = categorySummary.reduce(
      (acc, c) => {
        acc.totalQty += c.totalQty;
        acc.totalMeters += c.totalMeters;
        acc.totalAmount += c.totalAmount;
        acc.totalCost += c.totalCost;
        acc.totalProfit += c.totalProfit;
        return acc;
      },
      { totalQty: 0, totalMeters: 0, totalAmount: 0, totalCost: 0, totalProfit: 0 }
    );
    const totalMargin = total.totalAmount > 0 ? total.totalProfit / total.totalAmount : 0;
    const tr = document.createElement("tr");
    tr.className = "bg-slate-50 font-semibold";
    tr.innerHTML = `
      <td class="border px-2 py-1">合計</td>
      <td class="border px-2 py-1 text-right">-</td>
      <td class="border px-2 py-1">-</td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(total.totalQty)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(total.totalMeters)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(total.totalAmount)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(total.totalCost)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(total.totalProfit)}</td>
      <td class="border px-2 py-1 text-right">${formatPercent(totalMargin)}</td>
    `;
    categoryTbody.appendChild(tr);

    downloadBtn.classList.remove("hidden");
  }
}

let categoryChartInstance = null;

function renderChart(limitMode = "10") {
  const canvas = document.getElementById("categoryAmountChart");
  if (!canvas || !categorySummary.length) return;

  let list = [];
  if (limitMode === "all") {
    list = categorySummary;
  } else {
    const n = parseInt(limitMode, 10);
    list = categorySummary.slice(0, n);
  }

  const labels = list.map((c) => c.category);
  const data = list.map((c) => Math.round(c.totalAmount));

  if (categoryChartInstance) {
    categoryChartInstance.destroy();
  }

  categoryChartInstance = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "銷售總額",
          data,
          backgroundColor: "#F59E0B", // 橘色，跟產品頁的藍色區分
        },
      ],
    },
    options: {
      responsive: true,
      plugins: { legend: { display: true } },
      scales: {
        x: { ticks: { autoSkip: false, maxRotation: 60, minRotation: 20 } },
        y: { beginAtZero: true },
      },
      animation: { duration: 500 },
    },
  });
}

function downloadExcel() {
  if (!categorySummary.length) return;
  const aoa = [
    ["分類", "涵蓋編號數", "涵蓋物品編號參考", "總計數量", "總計米數", "銷售總額", "成本總額", "總毛利", "毛利率"],
    ...categorySummary.map((c) => [
      c.category,
      c.codeCount,
      c.displayCodes,
      Math.round(c.totalQty),
      Number((Math.round(c.totalMeters * 1000) / 1000).toFixed(3)),
      Math.round(c.totalAmount),
      Math.round(c.totalCost),
      Math.round(c.totalProfit),
      (c.marginRate * 100).toFixed(1) + "%",
    ]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "分類統計");
  XLSX.writeFile(wb, "產品分類加總統計.xlsx");
}

document.addEventListener("DOMContentLoaded", () => {
  const rows = loadFromLocalStorage();
  if (!rows) return;

  categorySummary = buildCategorySummary(rows);
  renderTable();
  renderChart("10");

  downloadBtn.addEventListener("click", downloadExcel);

  const chartSelect = document.getElementById("chartLimitSelect");
  if (chartSelect) {
    chartSelect.addEventListener("change", () => {
      renderChart(chartSelect.value);
    });
  }
});
