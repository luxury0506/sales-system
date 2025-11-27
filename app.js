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
 *
 * 優先順序：
 *  1) 先看物品編號包含 FSG-2 / FSG-3 / HST / SRG
 *  2) 如果物品編號看不出來，再用「品名關鍵字」判斷
 *
 * 規則（你提供的對應）：
 *  - 外玻內矽套管 / 外玻內矽絕緣套管 → HST → 順博 shunbo
 *  - 玻璃纖維矽套管 / 矽套管          → FSG 系列（預設當 FSG-3 → 順博）
 *  - 外矽內玻套管 / 外矽內玻           → SRG → 瑞普 ruipu
 *
 *  ⚠ PVC高壓套管 / PVC套管（CFT-3 / CFT-6）一律不計成本 → 這邊直接回 null
 ************************/
function getSupplierFromRow(row) {
  const codeU = (row.itemCode || "").toUpperCase();
  const name = row.name || "";

  // PVC 套管系列（CFT-3 / CFT-6），一律不計成本
  if (
    codeU.startsWith("CFT-3") ||
    codeU.startsWith("CFT-6") ||
    name.includes("PVC高壓套管") ||
    name.includes("PVC套管")
  ) {
    return null;
  }

  // 1️⃣ 先用物品編號判斷
  if (codeU.includes("FSG-3") || codeU.includes("HST")) {
    // FSG-3 & HST → 順博
    return "shunbo";
  }
  if (codeU.includes("FSG-2") || codeU.includes("SRG")) {
    // FSG-2 & SRG → 瑞普
    return "ruipu";
  }

  // 2️⃣ 再用品名關鍵字判斷

  // 外玻內矽 → HST → 順博
  if (name.includes("外玻內矽套管") || name.includes("外玻內矽絕緣套管")) {
    return "shunbo";
  }

  // 外矽內玻 → SRG → 瑞普
  if (name.includes("外矽內玻套管") || name.includes("外矽內玻")) {
    return "ruipu";
  }

  // 玻璃纖維矽套管 / 矽套管 → FSG 系列
  // （這裡假設預設用 FSG-3 → 順博；如果你之後有確定哪幾個是 FSG-2，再個別加特例）
  if (name.includes("玻璃纖維矽套管") || name.includes("矽套管")) {
    return "shunbo"; // 預設 FSG-3，用順博價格
  }

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
function getYunlinUnitPrice(itemCode, specMm, name) {
  if (!itemCode || specMm == null) return null;
  if (typeof YUNLIN_G5 === "undefined") return null;

  const code = itemCode.toUpperCase();
  const text = (name || "").toString();

  // H + 數字 開頭 → 雲林熱縮
  if (!/^H\d+/.test(code)) return null;

  let colorType = "black";

  // 1️⃣ 物品編號末尾的類型判斷（優先）
  if (code.endsWith("CB")) {
    colorType = "thin";          // 超薄
  } else if (code.endsWith("C")) {
    colorType = "transparent";   // 透明
  } else if (/(R|BL|G|Y|W)$/.test(code)) {
    colorType = "color";         // 彩色（含白色 W）
  } else {
    // 2️⃣ 再看品名文字裡有沒有顏色關鍵字
    if (
      text.includes("（黑") ||
      text.includes("黑色") ||
      text.includes("紅色") ||
      text.includes("（紅") ||
      text.includes("藍色") ||
      text.includes("（藍") ||
      text.includes("綠色") ||
      text.includes("（綠") ||
      text.includes("黃色") ||
      text.includes("（黃") ||
      /[Ww]/.test(text) // 品名裡出現 W/w 也當彩色（例如標註 White）
    ) {
      colorType = "color";
    }
  }

  const mmKey = String(specMm);
  const row = YUNLIN_G5[mmKey];
  if (!row) return null;

  const price = row[colorType];
  if (typeof price === "number" && price > 0) {
    return price; // 台幣 / 米
  }

  return null;
}

/***********************
 * 從 COST_MAP 取順博 / 瑞普的「每米人民幣單價」
 * 支援兩種格式：
 *  - Map：      new Map([[ "shunbo|3.0", 0.12 ], ...])
 *  - 一般物件： { "shunbo|3.0": 0.12, "3.0": 0.12, ... }
 ************************/
function getBasePriceFromCostTable(mmKey, supplier) {
  const raw = window.COST_MAP;
  if (!raw) return null;

  const mmStr = String(mmKey);
  const mmFloat = parseFloat(mmKey);
  const candidates = [];

  // 有供應商資訊就先試 supplier|mm
  if (supplier) {
    candidates.push(`${supplier}|${mmStr}`);
    if (!Number.isNaN(mmFloat)) {
      candidates.push(`${supplier}|${mmFloat}`);
    }
  }
  // 再退而求其次只用 mm 當 key
  candidates.push(mmStr);
  if (!Number.isNaN(mmFloat)) {
    candidates.push(String(mmFloat));
  }

  // 1) COST_MAP 是 Map
  if (raw instanceof Map) {
    for (const k of candidates) {
      if (raw.has(k)) {
        const val = raw.get(k);
        if (Number.isFinite(val)) return val;
      }
    }
  }
  // 2) COST_MAP 是一般物件
  else if (typeof raw === "object") {
    for (const k of candidates) {
      if (Object.prototype.hasOwnProperty.call(raw, k)) {
        const val = raw[k];
        if (Number.isFinite(val)) return val;
      }
    }
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
const clearDataBtn = document.getElementById("clearDataBtn");

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
if (clearDataBtn) {
  clearDataBtn.addEventListener("click", clearAnalysisData);
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

        // 解析客戶名稱列，例如：
        // 1) "客戶名稱:(CH049)世僖"
        // 2) "客戶名稱: (TC107)台中某公司"
        // 3) "客戶名稱:TC105"
        // 4) "客戶名稱: TC105 台中某公司"
        const firstCell = safeCell(row[0]);
        const customerLineMatch = firstCell.match(/^客戶名稱[:：]\s*(.+)$/);
        if (customerLineMatch) {
          const body = customerLineMatch[1].trim(); // 拿掉「客戶名稱:」

          let code = "";
          let name = "";

          // 情況一："(CH049)世僖"
          let mParen = body.match(/^\(([^)]+)\)\s*(.*)$/);
          if (mParen) {
            code = mParen[1].trim();        // CH049
            name = mParen[2].trim();        // 世僖
          } else {
            // 情況二："TC105 台中某公司" 或 "TC105"
            let mCodeName = body.match(/^([A-Za-z0-9]+)\s*(.*)$/);
            if (mCodeName) {
              code = mCodeName[1].trim();   // TC105
              name = mCodeName[2].trim();   // 台中某公司 (可能為空字串)
            } else {
              // 其它奇怪格式，就整串當成名稱
              code = body;
              name = "";
            }
          }

          const full = name ? `${code} ${name}` : code; // 有名字就 "代碼 名稱"，沒有就單純代碼
          currentCustomer = full;
          continue; // 這一列只用來設定客戶，不是銷貨資料
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
          customer: currentCustomer,
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
 * 主計算：雲林(不套匯率) + 順博/瑞普(要匯率) + CFT-3/6 不計成本
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
    let unitPrice = 0; // 台幣 / 米
    let cost = 0;

    const codeUpper = (row.itemCode || "").toUpperCase();

    // ⭐ CFT-3 / CFT-6 一律不計成本
    if (codeUpper.startsWith("CFT-3") || codeUpper.startsWith("CFT-6")) {
      const profit = row.amount; // 成本=0 → 毛利=銷售額
      return {
        ...row,
        unitPrice: 0,
        cost: 0,
        profit,
      };
    }

    // 1️⃣ 雲林電子熱縮（Hxx 開頭，不需匯率）
    const yunlinUnit = getYunlinUnitPrice(row.itemCode, row.specMm, row.name);
    if (yunlinUnit != null) {
      unitPrice = yunlinUnit;
      cost = unitPrice * row.meters;
} else {
  // 2️⃣ 順博 / 瑞普（含顏色加價：FSG-3 +5%，HST +8%，白色不加價）
  const supplier = getSupplierFromRow(row);
  const mmKey = row.specMm != null ? String(row.specMm) : null;

  if (supplier && mmKey && hasRate) {
    let basePrice = getBasePriceFromCostTable(mmKey, supplier);

    if (Number.isFinite(basePrice)) {
      const codeUpper = (row.itemCode || "").toUpperCase();
      const nameText = (row.name || "").toString();

      // ✅ 白色（不加價）
      const isWhite =
        nameText.includes("白") ||
        /W$/.test(codeUpper);

      // ✅ 彩色（不含白色、透明不列入）
      const isColor =
        !isWhite &&
        /(黑|紅|藍|綠|黃)/.test(nameText) ||
        /(R|BL|G|Y)$/.test(codeUpper);

      // ✅ FSG-3 彩色 +5%
      if (supplier === "shunbo" && codeUpper.includes("FSG-3") && isColor) {
        basePrice *= 1.05;
      }

      // ✅ HST 彩色 +8%
      if (supplier === "shunbo" && codeUpper.includes("HST") && isColor) {
        basePrice *= 1.08;
      }

      unitPrice = basePrice * rateVal; // 台幣 / 米
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

  // ❌ 排除不需要列出的品項 Z043 / Z044 / A1
  processedRows = processedRows.filter((row) => {
    const code = (row.itemCode || "").toUpperCase();
    return !code.startsWith("Z043") &&
       !code.startsWith("Z044") &&
       code !== "A1";
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
 * 手動清除分析資料
 ************************/
function clearAnalysisData() {
  try {
    localStorage.removeItem("salesAnalysisData");
  } catch (e) {
    console.error("清除 localStorage 失敗：", e);
  }
  baseRows = [];
  processedRows = [];
  resultTbody.innerHTML = "";
  tableContainer.classList.add("hidden");
  downloadBtn.classList.add("hidden");
  statusEl.textContent = "已清除分析資料，請重新上傳銷售檔。";
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