/***********************
 * 共用小工具
 ************************/
function safeCell(v) {
  if (v == null) return "";
  return v.toString().trim();
}

function escapeHtml(text) {
  if (text == null) return "";
  return text
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function parseNumber(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string") {
    const cleaned = v.replace(/,/g, "");
    const num = parseFloat(cleaned);
    return isNaN(num) ? NaN : num;
  }
  return NaN;
}

function formatMeters(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return (Math.round(num * 1000) / 1000).toLocaleString("zh-TW");
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

/***********************
 * 從品名抓 mm 規格與裁切長度
 *
 * 規則：
 *  - 先抓第一個出現的「xx mm」當規格 specMm
 *  - 只有在有「*」或「x」時，才抓後面的數字當裁切長度 cutMm
 *  - 若找不到裁切長度 → 視為「銷貨量單位就是米」
 *
 * 範例：
 *  - "熱收縮套管 3.5mm * 85 (白)"
 *      → specMm = 3.5, cutMm = 85
 *  - "熱收縮套管 2.0mm * 180mm"
 *      → specMm = 2.0, cutMm = 180
 *  - "玻璃纖維矽套管 10.0mm 1.5KV"
 *      → specMm = 10.0, cutMm = null（⇒ 米數 = 銷貨量）
 ************************/
function extractMmInfo(name) {
  if (!name) return { specMm: null, cutMm: null };

  // 1. 抓「有 mm 的數字」當規格
  const mmRegex = /([\d\.]+)\s*mm/gi;
  let mmMatch = mmRegex.exec(name);
  let specMm = null;

  if (mmMatch) {
    const v = parseFloat(mmMatch[1]);
    if (!isNaN(v)) {
      specMm = v;
    }
  }

  // 2. 尋找裁切長度：只接受有 * 或 x 的情況
  let cutMm = null;

  // 2-1 形式："...mm * 85" 或 "...mm x 85"
  let m = /mm\s*[x*]\s*(\d+(?:\.\d+)?)/i.exec(name);
  if (m) {
    const v = parseFloat(m[1]);
    if (!isNaN(v)) {
      cutMm = v;
    }
  } else {
    // 2-2 形式："* 180mm" 或 "x180mm"
    m = /[x*]\s*(\d+(?:\.\d+)?)(?=\s*mm\b)/i.exec(name);
    if (m) {
      const v = parseFloat(m[1]);
      if (!isNaN(v)) {
        cutMm = v;
      }
    }
  }

  return { specMm, cutMm };
}

/***********************
 * 判斷供應商（順博 / 瑞普，需匯率）
 ************************/
function getSupplierFromItemCode(itemCode) {
  if (!itemCode) return null;
  const code = itemCode.toUpperCase();
  if (code.includes("FSG-3") || code.includes("HST")) return "shunbo";
  if (code.includes("FSG-2") || code.includes("SRG")) return "ruipu";
  return null;
}

/***********************
 * 雲林電子 G5 熱縮價格（不需匯率）
 *
 * 規則：
 *   - H + 數字 開頭 → 雲林熱縮（H01, H015, H02, H035-085W...）
 *   - 結尾 CB       → 超薄 thin
 *   - 結尾 C        → 透明 transparent
 *   - 結尾 R/BL/G/Y/W → 彩色 color（含白色 W）
 *   - 其他          → 黑色 black
 ************************/
function getYunlinUnitPrice(itemCode, specMm) {
  if (!itemCode || specMm == null) return null;
  if (typeof YUNLIN_G5 === "undefined") return null;

  const code = itemCode.toUpperCase();

  // H + 數字 開頭 → 雲林熱縮
  if (!/^H\d+/.test(code)) return null;

  let colorType = "black";

  if (code.endsWith("CB")) {
    colorType = "thin";           // 超薄
  } else if (code.endsWith("C")) {
    colorType = "transparent";    // 透明
  } else if (/(R|BL|G|Y|W)$/.test(code)) {
    colorType = "color";          // 彩色（含白色 W）
  }

  const mmKey = String(specMm);
  const row = YUNLIN_G5[mmKey];
  if (!row) return null;

  const price = row[colorType];
  if (typeof price === "number" && price > 0) {
    return price;                 // 台幣 / 米
  }

  return null;
}

/***********************
 * DOM 元素
 ************************/
const salesFileInput = document.getElementById("salesFile");
const exchangeRateInput = document.getElementById("exchangeRate");
const applyRateBtn = document.getElementById("applyRateBtn");
const statusEl = document.getElementById("status");
const tableContainer = document.getElementById("tableContainer");
const resultTbody = document.getElementById("resultTbody");
const downloadBtn = document.getElementById("downloadBtn");

/***********************
 * 全域資料
 ************************/
let baseRows = [];      // 原始每筆銷貨（含客戶、品名、米數…）
let processedRows = []; // 加上成本、毛利後的資料

// 順博 / 瑞普 成本表，從 cost-data.js 來
const costMap = window.COST_MAP || new Map();

/***********************
 * 將分析完的資料存到 localStorage
 * 給 customer.html / cost.html 使用
 ************************/
function saveToLocalStorage() {
  try {
    const rateVal = parseFloat(exchangeRateInput.value);
    const payload = {
      exchangeRate: Number.isFinite(rateVal) ? rateVal : null,
      baseRows,
      processedRows,
      savedAt: new Date().toISOString(),
    };
    localStorage.setItem("salesAnalysisData", JSON.stringify(payload));
  } catch (e) {
    console.error("儲存到 localStorage 失敗：", e);
  }
}

/***********************
 * 綁定事件
 ************************/
if (salesFileInput) {
  salesFileInput.addEventListener("change", handleSalesFile);
}
if (applyRateBtn) {
  applyRateBtn.addEventListener("click", () => recalcAndRender());
}
if (exchangeRateInput) {
  exchangeRateInput.addEventListener("change", () => recalcAndRender());
}
if (downloadBtn) {
  downloadBtn.addEventListener("click", downloadExcel);
}

/***********************
 * 讀取銷售明細 Excel
 ************************/
function handleSalesFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  statusEl.textContent = "讀取銷售明細中...";
  tableContainer.classList.add("hidden");
  downloadBtn.classList.add("hidden");
  resultTbody.innerHTML = "";
  baseRows = [];
  processedRows = [];

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = evt.target.result;

      if (typeof XLSX === "undefined") {
        statusEl.textContent = "找不到 XLSX 函式庫，請確認 index.html 有載入 SheetJS。";
        return;
      }

      const workbook = XLSX.read(data, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true });

      if (!rows || !rows.length) {
        statusEl.textContent = "銷售檔案內容為空。";
        return;
      }

      // 找標題列
      let headerRowIndex = -1;
      let itemCodeColIndex = -1;
      let nameColIndex = -1;
      let qtyColIndex = -1;
      let amountColIndex = -1;

      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;

        const qIndex = row.findIndex(
          (cell) => typeof cell === "string" && cell.toString().includes("銷貨量")
        );
        if (qIndex !== -1) {
          headerRowIndex = r;
          qtyColIndex = qIndex;

          itemCodeColIndex = row.findIndex((cell) => {
            if (typeof cell !== "string") return false;
            const text = cell.toString().replace(/\s/g, "");
            return text.includes("物品編號");
          });

          nameColIndex = row.findIndex((cell) => {
            if (typeof cell !== "string") return false;
            const text = cell.toString().replace(/\s/g, "");
            return text.includes("品名");
          });

          amountColIndex = row.findIndex((cell) => {
            if (typeof cell !== "string") return false;
            const text = cell.toString().replace(/\s/g, "");
            return text.includes("銷貨金額");
          });

          break;
        }
      }

      if (
        headerRowIndex === -1 ||
        nameColIndex === -1 ||
        qtyColIndex === -1
      ) {
        statusEl.textContent =
          "找不到標題列（需要至少有「品名」、「銷貨量」欄位），請確認銷售報表格式。";
        return;
      }

      const results = [];
      let currentCustomer = "";

      for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;

        // 客戶名稱:(CH049)世僖
        const firstCell = safeCell(row[0]);
        const customerMatch = firstCell.match(/^客戶名稱[:：]\((.+?)\)/);
        if (customerMatch) {
          // 原字串是 "(CH049)世僖"
          const raw = customerMatch[1].trim(); // CH049)世僖
          // 這裡先不拆 code / name，直接整串存起來即可
          currentCustomer = raw;
          continue;
        }

        const itemCode =
          itemCodeColIndex !== -1 ? safeCell(row[itemCodeColIndex]) : "";
        const name = safeCell(row[nameColIndex]);
        if (!name) continue;

        const qty = parseNumber(row[qtyColIndex]);
        if (!Number.isFinite(qty)) continue;

        const amount =
          amountColIndex !== -1 ? parseNumber(row[amountColIndex]) : 0;

        const { specMm, cutMm } = extractMmInfo(name);
        const meters = cutMm != null ? qty * (cutMm / 1000) : qty;

        results.push({
          customer: currentCustomer, // 例如 "CH049)世僖"（之後統計頁再細拆也可以）
          itemCode,
          name,
          qty,
          meters,
          amount,
          specMm,
        });
      }

      if (!results.length) {
        statusEl.textContent = "沒有找到有效的銷售品項資料。";
        return;
      }

      baseRows = results;
      statusEl.textContent = `銷售明細讀取完成，共 ${results.length} 筆品項。`;
      recalcAndRender();
    } catch (err) {
      console.error(err);
      statusEl.textContent = "讀取銷售檔案時發生錯誤，請確認檔案格式。";
    }
  };

  reader.readAsArrayBuffer(file);
}

/***********************
 * 主計算：雲林(不套匯率) + 順博/瑞普(要匯率)
 ************************/
function recalcAndRender() {
  if (!baseRows.length) {
    tableContainer.classList.add("hidden");
    downloadBtn.classList.add("hidden");
    return;
  }

  const rateVal = parseFloat(exchangeRateInput.value);
  const hasRate = Number.isFinite(rateVal) && rateVal > 0;

  if (!hasRate) {
    statusEl.textContent =
      "提醒：尚未輸入有效匯率，順博 / 瑞普的銷貨成本與毛利會顯示為 0（雲林熱縮不受影響）。";
  }

  processedRows = baseRows.map((row) => {
    let unitPrice = 0;  // 最後都視為「台幣 / 米」
    let cost = 0;

    // 1️⃣ 雲林電子熱縮（Hxx 開頭，不需匯率）
    const yunlinUnit = getYunlinUnitPrice(row.itemCode, row.specMm);
    if (yunlinUnit != null) {
      unitPrice = yunlinUnit;
      cost = unitPrice * row.meters;
    } else {
      // 2️⃣ 順博 / 瑞普（FSG-2, FSG-3, HST, SRG，需要匯率）
      const supplierFromCode = getSupplierFromItemCode(row.itemCode);
      let supplier = supplierFromCode;
      const mmKey = row.specMm != null ? String(row.specMm) : null;

      if (mmKey && costMap.size) {
        // 2-1 依物品編號指定的供應商尋找
        if (supplier) {
          const key = `${supplier}|${mmKey}`;
          const basePrice = costMap.get(key);
          if (Number.isFinite(basePrice) && hasRate) {
            unitPrice = basePrice * rateVal; // 台幣 / 米
          }
        }

        // 2-2 找不到就 fallback：先順博再瑞普
        if (!unitPrice && hasRate) {
          const keyShunbo = `shunbo|${mmKey}`;
          const keyRuipu = `ruipu|${mmKey}`;
          if (costMap.has(keyShunbo)) {
            unitPrice = costMap.get(keyShunbo) * rateVal;
          } else if (costMap.has(keyRuipu)) {
            unitPrice = costMap.get(keyRuipu) * rateVal;
          }
        }

        if (unitPrice && hasRate) {
          cost = unitPrice * row.meters;
        }
      }
    }

    const profit = row.amount - cost;

    return {
      ...row,
      unitPrice,
      cost,
      profit,
    };
  });

  // ❌ 排除不需要列出的品項 Z043 / Z044
  processedRows = processedRows.filter((row) => {
    const code = (row.itemCode || "").toUpperCase();
    return !code.startsWith("Z043") && !code.startsWith("Z044");
  });

  // ✅ 儲存到 localStorage，給其他頁面用
  saveToLocalStorage();

  renderTable();
}

/***********************
 * 畫表格（含物品編號＋總計）
 ************************/
function renderTable() {
  resultTbody.innerHTML = "";

  let totalQty = 0;
  let totalMeters = 0;
  let totalAmount = 0;
  let totalCost = 0;
  let totalProfit = 0;

  processedRows.forEach((row) => {
    totalQty += row.qty;
    totalMeters += row.meters;
    totalAmount += row.amount;
    totalCost += row.cost;
    totalProfit += row.profit;

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="border px-2 py-1">${escapeHtml(row.itemCode)}</td>
      <td class="border px-2 py-1">${escapeHtml(row.name)}</td>
      <td class="border px-2 py-1 text-right">${formatQtyInt(row.qty)}</td>
      <td class="border px-2 py-1 text-right">${formatMeters(row.meters)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(row.amount)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(row.cost)}</td>
      <td class="border px-2 py-1 text-right">${formatMoney(row.profit)}</td>
    `;
    resultTbody.appendChild(tr);
  });

  // 總計列
  const totalTr = document.createElement("tr");
  totalTr.classList.add("bg-yellow-100", "font-semibold");
  totalTr.innerHTML = `
    <td class="border px-2 py-1 text-right">總計：</td>
    <td class="border px-2 py-1"></td>
    <td class="border px-2 py-1 text-right">${formatQtyInt(totalQty)}</td>
    <td class="border px-2 py-1 text-right">${formatMeters(totalMeters)}</td>
    <td class="border px-2 py-1 text-right">${formatMoney(totalAmount)}</td>
    <td class="border px-2 py-1 text-right">${formatMoney(totalCost)}</td>
    <td class="border px-2 py-1 text-right">${formatMoney(totalProfit)}</td>
  `;
  resultTbody.appendChild(totalTr);

  tableContainer.classList.remove("hidden");
  downloadBtn.classList.remove("hidden");
}

/***********************
 * 匯出 Excel（含物品編號＋總計）
 ************************/
function downloadExcel() {
  if (!processedRows.length) return;

  let totalQty = 0;
  let totalMeters = 0;
  let totalAmount = 0;
  let totalCost = 0;
  let totalProfit = 0;

  const bodyRows = processedRows.map((r) => {
    totalQty += r.qty;
    totalMeters += r.meters;
    totalAmount += r.amount;
    totalCost += r.cost;
    totalProfit += r.profit;

    return [
      r.itemCode,
      r.name,
      formatQtyInt(r.qty),
      Number((Math.round(r.meters * 1000) / 1000).toFixed(3)),
      Math.round(r.amount),
      Math.round(r.cost),
      Math.round(r.profit),
    ];
  });

  const aoa = [
    ["物品編號", "品名", "銷貨量", "換算米數(米)", "銷貨金額", "銷貨成本", "銷貨毛利"],
    ...bodyRows,
    [
      "總計",
      "",
      formatQtyInt(totalQty),
      Number((Math.round(totalMeters * 1000) / 1000).toFixed(3)),
      Math.round(totalAmount),
      Math.round(totalCost),
      Math.round(totalProfit),
    ],
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "銷貨成本毛利試算");
  XLSX.writeFile(wb, "銷貨成本毛利試算.xlsx");
}
