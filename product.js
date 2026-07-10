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
const productTbody = document.getElementById("productTbody");
const downloadBtn = document.getElementById("downloadProductExcel");
let productSummary = [];

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

function extractProductSeries(itemCode) {
  if (!itemCode) return "未填寫物品編號";
  
  // 將結尾的裁切長度或米數後綴移除，例如：
  // "CFT-3-0A-200" -> "CFT-3-0A"
  // "CFT-3-0A-530M" -> "CFT-3-0A"
  // "CFT-3-0A-200m" -> "CFT-3-0A"
  return itemCode.trim().replace(/-\d+(?:\.\d+)?(?:[Mm])?$/, "");
}

function buildProductSummary(rows) {
  const map = new Map();

  rows.forEach((row) => {
    // 使用正規化的物品編號來當作分類 Key
    const key = extractProductSeries(row.itemCode);
    
    if (!map.has(key)) {
      map.set(key, {
        product: key,
        totalQty: 0,
        totalMeters: 0,
        totalAmount: 0,
        totalCost: 0,
        totalProfit: 0,
      });
    }
    
    const agg = map.get(key);
    agg.totalQty += Number(row.qty) || 0;
    agg.totalMeters += Number(row.meters) || 0;
    agg.totalAmount += Number(row.amount) || 0;
    agg.totalCost += Number(row.cost) || 0;
    agg.totalProfit += Number(row.profit) || 0;
  });

  const list = Array.from(map.values()).map((x) => {
    const marginRate = x.totalAmount > 0 ? x.totalProfit / x.totalAmount : 0;
    return { ...x, marginRate };
  });

  // 依總米數由高到低排序，因為通常看產品會比較關注銷貨的米數
  list.sort((a, b) => b.totalMeters - a.totalMeters);
  return list;
}

function renderTable() {
  productTbody.innerHTML = "";
  productSummary.forEach((c) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="border px-2 py-1">${escapeHtml(c.product)}</td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(c.totalQty)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(c.totalMeters)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalAmount)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalCost)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalProfit)}</td>
      <td class="border px-2 py-1 text-right">${formatPercent(c.marginRate)}</td>
    `;
    productTbody.appendChild(tr);
  });

  if (productSummary.length) {
    downloadBtn.classList.remove("hidden");
  }
}

let metersChartInstance = null;

function renderChart(limitMode = "10") {
  const canvas = document.getElementById("metersChart");
  if (!canvas || !productSummary.length) return;

  // 依下拉選單決定顯示筆數
  let list = [];
  if (limitMode === "all") {
    list = productSummary;
  } else {
    const n = parseInt(limitMode, 10);
    list = productSummary.slice(0, n);
  }

  const labels = list.map((c) => c.product);
  const data = list.map((c) => Math.round(c.totalMeters));

  if (metersChartInstance) {
    metersChartInstance.destroy();
  }

  metersChartInstance = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "總計米數 (M)",
          data,
          backgroundColor: "#3B82F6", // 使用藍色
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
  if (!productSummary.length) return;
  const aoa = [
    ["產品系列", "總計數量", "總計米數", "銷售總額", "成本總額", "總毛利", "毛利率"],
    ...productSummary.map((c) => [
      c.product,
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
  XLSX.utils.book_append_sheet(wb, ws, "產品統計");
  XLSX.writeFile(wb, "產品別米數統計.xlsx");
}

document.addEventListener("DOMContentLoaded", () => {
  const rows = loadFromLocalStorage();
  if (!rows) return;

  productSummary = buildProductSummary(rows);
  renderTable();
  renderChart("10"); // 預設 Top10

  downloadBtn.addEventListener("click", downloadExcel);

  const chartSelect = document.getElementById("chartLimitSelect");
  if (chartSelect) {
    chartSelect.addEventListener("change", () => {
      renderChart(chartSelect.value);
    });
  }
});
