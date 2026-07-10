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

function extractProductSeries(row) {
  let code = (row.itemCode || "").trim();
  if (!code) return "未填寫物品編號";
  const name = (row.name || "").trim();
  
  // 1. 若明顯包含 M 長度單位結尾 (如 200M, 530M)，直接切除
  if (/[0-9]+[Mm]$/.test(code)) {
    return code.replace(/[-_\s]*\d+(?:\.\d+)?(?:[Mm])?$/, "");
  }

  const segments = code.split(/[-_]/);
  if (segments.length > 1) {
    const last = segments[segments.length - 1];
    const lastNumMatch = last.match(/^(\d+(?:\.\d+)?)([A-Za-z]*)$/);
    
    if (lastNumMatch) {
      const numStr = lastNumMatch[1];
      const letters = lastNumMatch[2];
      
      let shouldCut = false;
      
      // 2. 判斷是否有裁切長度 (從品名取得)
      const info = extractMmInfo(name);
      if (info.cutMm != null && parseFloat(numStr) === info.cutMm) {
        shouldCut = true;
      }
      
      // 3. 備用：您剛剛提供的明確長度清單
      const explicitLengths = ['115', '053', '200', '530', '1000', '1650'];
      if (explicitLengths.includes(numStr)) {
        shouldCut = true;
      }
      
      // 4. 若純數字且大於三位數且沒有前導零 (避免切到 065)，也當作長度
      if (letters === "" && numStr.length >= 3 && !numStr.startsWith("0")) {
        shouldCut = true;
      }
      
      if (shouldCut) {
        segments.pop();
        // 如果有字母(例如 b, c)，把它加回倒數第二段，保留顏色！
        if (letters) {
          segments[segments.length - 1] += letters;
        }
        return segments.join("-");
      }
    }
  }
  return code;
}

function buildProductSummary(rows) {
  const map = new Map();

  rows.forEach((row) => {
    // 傳入完整的 row，讓 extractProductSeries 可以利用 row.name 進行裁切分析
    const key = extractProductSeries(row);
    
    if (!map.has(key)) {
      map.set(key, {
        product: key,
        names: new Set(), // 收集同屬性下的所有品名
        totalQty: 0,
        totalMeters: 0,
        totalAmount: 0,
        totalCost: 0,
        totalProfit: 0,
      });
    }
    
    const agg = map.get(key);
    if (row.name) {
      // 避免品名太長或包含長度綴飾，先簡單收納
      agg.names.add(row.name.trim());
    }
    agg.totalQty += Number(row.qty) || 0;
    agg.totalMeters += Number(row.meters) || 0;
    agg.totalAmount += Number(row.amount) || 0;
    agg.totalCost += Number(row.cost) || 0;
    agg.totalProfit += Number(row.profit) || 0;
  });

  const list = Array.from(map.values()).map((x) => {
    const marginRate = x.totalAmount > 0 ? x.totalProfit / x.totalAmount : 0;
    // 將收集到的 Set 轉回逗號分隔字串
    const namesArray = Array.from(x.names);
    // 如果名字太多限制一下長度，或者全顯示
    const joinedNames = namesArray.length > 5 ? namesArray.slice(0, 5).join("、 ") + "..." : namesArray.join("、 ");
    
    return { ...x, marginRate, displayNames: joinedNames };
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
      <td class="border px-2 py-1 text-slate-500 text-[10px] break-words max-w-[200px]" title="${escapeHtml(c.displayNames)}">${escapeHtml(c.displayNames)}</td>
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
    ["產品系列", "涵蓋之品名參考", "總計數量", "總計米數", "銷售總額", "成本總額", "總毛利", "毛利率"],
    ...productSummary.map((c) => [
      c.product,
      c.displayNames,
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
