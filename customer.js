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
const customerTbody = document.getElementById("customerTbody");
const downloadBtn = document.getElementById("downloadCustomerExcel");
let customerSummary = [];

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

function buildCustomerSummary(rows) {
  const map = new Map();

  rows.forEach((row) => {
    const key = row.customer || "未填寫客戶";
    if (!map.has(key)) {
      map.set(key, {
        customer: key,
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
    const marginRate =
      x.totalAmount > 0 ? x.totalProfit / x.totalAmount : 0;
    return { ...x, marginRate };
  });

  // 依毛利由高到低排序
  list.sort((a, b) => b.totalProfit - a.totalProfit);
  return list;
}

function renderTable() {
  customerTbody.innerHTML = "";
  customerSummary.forEach((c) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="border px-2 py-1">${escapeHtml(c.customer)}</td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(c.totalQty)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(c.totalMeters)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalAmount)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalCost)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(c.totalProfit)}</td>
      <td class="border px-2 py-1 text-right">${formatPercent(c.marginRate)}</td>
    `;
    customerTbody.appendChild(tr);
  });

  if (customerSummary.length) {
    downloadBtn.classList.remove("hidden");
  }
}

function renderChart() {
  const canvas = document.getElementById("profitChart");
  if (!canvas || !customerSummary.length) return;

  const top10 = customerSummary.slice(0, 10);
  const labels = top10.map((c) => c.customer);
  const data = top10.map((c) => Math.round(c.totalProfit));

  new Chart(canvas.getContext("2d"), {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "毛利（元）",
          data,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: true },
      },
      scales: {
        x: { ticks: { autoSkip: false, maxRotation: 60, minRotation: 30 } },
        y: { beginAtZero: true },
      },
    },
  });
}

function downloadExcel() {
  if (!customerSummary.length) return;
  const aoa = [
    ["客戶", "總銷貨量", "總米數", "銷售額", "銷貨成本", "銷貨毛利", "毛利率"],
    ...customerSummary.map((c) => [
      c.customer,
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
  XLSX.utils.book_append_sheet(wb, ws, "客戶統計");
  XLSX.writeFile(wb, "客戶毛利統計.xlsx");
}

document.addEventListener("DOMContentLoaded", () => {
  const rows = loadFromLocalStorage();
  if (!rows) return;

  customerSummary = buildCustomerSummary(rows);
  renderTable();
  renderChart();

  downloadBtn.addEventListener("click", downloadExcel);
});
