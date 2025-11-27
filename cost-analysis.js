let itemProfitChartInstance = null;

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
const itemTbody = document.getElementById("itemTbody");
const downloadBtn = document.getElementById("downloadItemExcel");
let itemSummary = [];

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
      (payload.savedAt ? `（最後更新：${payload.savedAt}）` : "");
    return payload.processedRows;
  } catch (e) {
    console.error(e);
    statusEl.textContent = "讀取分析資料時發生錯誤。";
    return null;
  }
}

function buildItemSummary(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const key = (row.itemCode || "") + "||" + (row.name || "");
    if (!map.has(key)) {
      map.set(key, {
        itemCode: row.itemCode || "",
        name: row.name || "",
        totalMeters: 0,
        totalAmount: 0,
        totalCost: 0,
        totalProfit: 0,
      });
    }
    const agg = map.get(key);
    agg.totalMeters += Number(row.meters) || 0;
    agg.totalAmount += Number(row.amount) || 0;
    agg.totalCost += Number(row.cost) || 0;
    agg.totalProfit += Number(row.profit) || 0;
  });

  const list = Array.from(map.values()).map((x) => {
    const avgSell =
      x.totalMeters > 0 ? x.totalAmount / x.totalMeters : 0;
    const avgCost =
      x.totalMeters > 0 ? x.totalCost / x.totalMeters : 0;
    const marginRate =
      x.totalAmount > 0 ? x.totalProfit / x.totalAmount : 0;
    return { ...x, avgSell, avgCost, marginRate };
  });

  // 依毛利由高到低排序
  list.sort((a, b) => b.totalProfit - a.totalProfit);
  return list;
}

function renderTable() {
  itemTbody.innerHTML = "";

  itemSummary.forEach((p) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="border px-2 py-1">${escapeHtml(p.itemCode)}</td>
      <td class="border px-2 py-1">${escapeHtml(p.name)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(p.totalMeters)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(p.totalAmount)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(p.totalCost)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(p.totalProfit)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(p.avgSell)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(p.avgCost)}</td>
      <td class="border px-2 py-1 text-right">${formatPercent(p.marginRate)}</td>
    `;
    itemTbody.appendChild(tr);
  });

  if (itemSummary.length) {
    downloadBtn.classList.remove("hidden");
  }
}

function renderChart(limitMode = "10") {
  const canvas = document.getElementById("itemProfitChart");
  if (!canvas || !itemSummary.length) return;

  let list = [];
  if (limitMode === "all") {
    list = itemSummary;
  } else {
    const n = parseInt(limitMode);
    list = itemSummary.slice(0, n);
  }

  const labels = list.map((p) =>
    p.itemCode.length > 12 ? p.itemCode.slice(0, 12) + "…" : p.itemCode
  );

  const data = list.map((p) => Math.round(p.totalProfit));

  const colors = data.map((v) =>
    v >= 0 ? "#4CAF50" : "#E53935"
  );

  if (itemProfitChartInstance) {
    itemProfitChartInstance.destroy();
  }

  itemProfitChartInstance = new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "毛利（元）",
          data,
          backgroundColor: colors,
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
    },
  });
}


function downloadExcel() {
  if (!itemSummary.length) return;

  const aoa = [
    [
      "物品編號",
      "品名",
      "總米數",
      "銷售額",
      "銷貨成本",
      "毛利",
      "平均售價/米",
      "平均成本/米",
      "毛利率",
    ],
    ...itemSummary.map((p) => [
      p.itemCode,
      p.name,
      Number((Math.round(p.totalMeters * 1000) / 1000).toFixed(3)),
      Math.round(p.totalAmount),
      Math.round(p.totalCost),
      Math.round(p.totalProfit),
      Math.round(p.avgSell),
      Math.round(p.avgCost),
      (p.marginRate * 100).toFixed(1) + "%",
    ]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "品項成本比較");
  XLSX.writeFile(wb, "品項成本比較.xlsx");
}

document.addEventListener("DOMContentLoaded", () => {
  const rows = loadFromLocalStorage();
  if (!rows) return;

  itemSummary = buildItemSummary(rows);
  renderTable();
  renderChart("10");

  downloadBtn.addEventListener("click", downloadExcel);

  const chartSelect = document.getElementById("itemChartLimitSelect");
  if (chartSelect) {
    chartSelect.addEventListener("change", () => {
      renderChart(chartSelect.value);
    });
  }
});

