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
 ************************/
function extractMmInfo(name) {
  if (!name) return { specMm: null, cutMm: null };

  const mmRegex = /([\d\.]+)\s*mm/gi;
  let mmMatch = mmRegex.exec(name);
  let specMm = null;

  if (mmMatch) {
    const v = parseFloat(mmMatch[1]);
    if (!isNaN(v)) {
      specMm = v;
    }
  }

  // 裁切長度
  let cutMm = null;

  let m = /mm\s*[x*]\s*(\d+(?:\.\d+)?)/i.exec(name);
  if (m) {
    const v = parseFloat(m[1]);
    if (!isNaN(v)) {
      cutMm = v;
    }
  } else {
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
function getSupplierFromRow(row) {
  const codeU = (row.itemCode || "").toUpperCase();
  const name = row.name || "";

  // PVC 套管系列 -> 在 PVC 成本處理，不在這裡判斷供應商
  if (
    codeU.startsWith("CFT-3") ||
    codeU.startsWith("CFT-6") ||
    name.includes("PVC高壓套管") ||
    name.includes("PVC套管")
  ) {
    return null;
  }

  // 物品編號
  if (codeU.includes("FSG-3") || codeU.includes("HST")) {
    return "shunbo";
  }
  if (codeU.includes("FSG-2") || codeU.includes("SRG")) {
    return "ruipu";
  }

  // 品名關鍵字
  if (name.includes("外玻內矽套管") || name.includes("外玻內矽絕緣套管")) {
    return "shunbo";
  }

  if (name.includes("外矽內玻套管") || name.includes("外矽內玻")) {
    return "ruipu";
  }

  if (name.includes("玻璃纖維矽套管") || name.includes("矽套管")) {
    return "shunbo"; // 預設 FSG-3
  }

  return null;
}

/***********************
 * 雲林電子 G5 熱縮價格（不需匯率）
 ************************/
function getYunlinUnitPrice(itemCode, specMm, name) {
  if (!itemCode || specMm == null) return null;
  if (typeof YUNLIN_G5 === "undefined") return null;

  const code = itemCode.toUpperCase();
  const text = (name || "").toString();

  if (!/^H\d+/.test(code)) return null;

  let colorType = "black";

  if (code.endsWith("CB")) {
    colorType = "thin";
  } else if (code.endsWith("C")) {
    colorType = "transparent";
  } else if (/(R|BL|G|Y|W)$/.test(code)) {
    colorType = "color";
  } else {
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
      /[Ww]/.test(text)
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
 ************************/
function getBasePriceFromCostTable(mmKey, supplier) {
  const raw = window.COST_MAP;
  if (!raw) return null;

  const mmStr = String(mmKey);
  const mmFloat = parseFloat(mmKey);
  const candidates = [];

  if (supplier) {
    candidates.push(`${supplier}|${mmStr}`);
    if (!Number.isNaN(mmFloat)) {
      candidates.push(`${supplier}|${mmFloat}`);
    }
  }
  candidates.push(mmStr);
  if (!Number.isNaN(mmFloat)) {
    candidates.push(String(mmFloat));
  }

  if (raw instanceof Map) {
    for (const k of candidates) {
      if (raw.has(k)) {
        const val = raw.get(k);
        if (Number.isFinite(val)) return val;
      }
    }
  } else if (typeof raw === "object") {
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
 * PVC 成本（不吃匯率），由 pvc.html 設定
 * localStorage 結構：
 * {
 *   cft3: { "8.380": 單價元/M, ... },
 *   cft6: { "10.000": 單價元/M, ... },
 *   updatedAt: ISOString
 * }
 ************************/
const PVC_STORAGE_KEY = "PVC_COST_TABLE";
let PVC_COST_CFT3 = {};
let PVC_COST_CFT6 = {};

function loadPvcCostFromLocalStorage() {
  try {
    const raw = localStorage.getItem(PVC_STORAGE_KEY);
    if (!raw) return;
    const obj = JSON.parse(raw);
    PVC_COST_CFT3 = obj.cft3 || {};
    PVC_COST_CFT6 = obj.cft6 || {};
  } catch (e) {
    console.error("讀取 PVC 成本失敗：", e);
  }
}

/**
 * 根據物品編號 + 規格(mm) 取得 PVC 每米成本（元/M，不吃匯率）
 */
function getPvcUnitPrice(itemCode, specMm, name) {
  const codeUpper = (itemCode || "").toUpperCase();
  const text = (name || "").toString();
  const d = specMm != null ? parseFloat(specMm) : NaN;
  if (!Number.isFinite(d)) return null;

  let table = null;

  if (codeUpper.startsWith("CFT-3")) {
    table = PVC_COST_CFT3;
  } else if (codeUpper.startsWith("CFT-6")) {
    table = PVC_COST_CFT6;
  } else if (text.includes("PVC高壓套管") || text.includes("PVC套管")) {
    // 若之後有其它 PVC 物品編號格式，再補判斷；目前先限制 CFT-3/6
    return null;
  } else {
    return null;
  }

  const keys = Object.keys(table);
  if (!keys.length) return null;

  let bestKey = null;
  let bestDiff = Infinity;
  for (const k of keys) {
    const v = parseFloat(k);
    if (!Number.isFinite(v)) continue;
    const diff = Math.abs(v - d);
    if (diff < bestDiff) {
      bestDiff = diff;
      bestKey = k;
    }
  }

  if (!bestKey) return null;
  const unit = table[bestKey];
  return Number.isFinite(unit) ? unit : null;
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
let baseRows = [];
let processedRows = [];

// 順博 / 瑞普 成本表
const costMap = window.COST_MAP || new Map();

// 一載入就讀 PVC 成本
loadPvcCostFromLocalStorage();

/***********************
 * 將分析完的資料存到 localStorage
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

        const firstCell = safeCell(row[0]);
        const customerLineMatch = firstCell.match(/^客戶名稱[:：]\s*(.+)$/);
        if (customerLineMatch) {
          const body = customerLineMatch[1].trim();

          let code = "";
          let name = "";

          let mParen = body.match(/^\(([^)]+)\)\s*(.*)$/);
          if (mParen) {
            code = mParen[1].trim();
            name = mParen[2].trim();
          } else {
            let mCodeName = body.match(/^([A-Za-z0-9]+)\s*(.*)$/);
            if (mCodeName) {
              code = mCodeName[1].trim();
              name = mCodeName[2].trim();
            } else {
              code = body;
              name = "";
            }
          }

          const full = name ? `${code} ${name}` : code;
          currentCustomer = full;
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

// ========= PVC 成本利潤計算（在 index 上用） =========

const pvcWeightPerMInput = document.getElementById("pvcWeightPerM");
const pvcWastePercentInput = document.getElementById("pvcWastePercent");
const pvcPelletPriceInput = document.getElementById("pvcPelletPrice");
const pvcProfitPercentInput = document.getElementById("pvcProfitPercent");
const pvcSeriesSelect = document.getElementById("pvcSeries");
const pvcSpecMmInput = document.getElementById("pvcSpecMm");
const pvcCalcSaveBtn = document.getElementById("pvcCalcSaveBtn");
const pvcCalcStatus = document.getElementById("pvcCalcStatus");

// 這裡沿用你原本讀進來的 PVC_COST_CFT3 / PVC_COST_CFT6
// 如果一開始還沒呼叫 loadPvcCostFromLocalStorage，就先呼叫一次
if (typeof loadPvcCostFromLocalStorage === "function") {
  loadPvcCostFromLocalStorage();
}

// 把目前 PVC 成本寫回 localStorage，維持原本結構
function savePvcCostToLocalStorage() {
  try {
    const payload = {
      cft3: PVC_COST_CFT3 || {},
      cft6: PVC_COST_CFT6 || {},
      updatedAt: new Date().toISOString(),
    };
    localStorage.setItem(PVC_STORAGE_KEY, JSON.stringify(payload));
  } catch (e) {
    console.error("儲存 PVC 成本失敗：", e);
  }
}

if (pvcCalcSaveBtn) {
  pvcCalcSaveBtn.addEventListener("click", () => {
    const weightPerM = parseFloat(pvcWeightPerMInput.value); // g/M
    const wastePercent = parseFloat(pvcWastePercentInput.value) || 0;
    const pelletPrice = parseFloat(pvcPelletPriceInput.value);
    const profitPercent = parseFloat(pvcProfitPercentInput.value) || 0;
    const series = pvcSeriesSelect.value; // CFT-3 / CFT-6
    const specMm = parseFloat(pvcSpecMmInput.value);

    if (!Number.isFinite(weightPerM) || weightPerM <= 0) {
      pvcCalcStatus.textContent = "請輸入有效的『每米重 (g/M)』。";
      return;
    }
    if (!Number.isFinite(pelletPrice) || pelletPrice <= 0) {
      pvcCalcStatus.textContent = "請輸入有效的『PVC 粒 (元/Kg)』。";
      return;
    }
    if (!Number.isFinite(specMm) || specMm <= 0) {
      pvcCalcStatus.textContent = "請輸入有效的『規格內徑 (mm)』。";
      return;
    }

    const wasteRate = wastePercent / 100;
    const profitRate = profitPercent / 100;

    // 每米重量 Kg
    const weightKgPerM = weightPerM / 1000;

    // 成本 = 重量 × PVC粒 × (1 + 廢料%)
    const materialPricePerKg = pelletPrice * (1 + wasteRate);
    const costPerM = weightKgPerM * materialPricePerKg;

    // 售價（只是顯示給你看，真正存起來的是成本）
    const pricePerM = costPerM * (1 + profitRate);

    const key = specMm.toFixed(3); // 內徑 mm 當 key

    if (series === "CFT-3") {
      PVC_COST_CFT3[key] = costPerM;
    } else {
      PVC_COST_CFT6[key] = costPerM;
    }

    savePvcCostToLocalStorage();

    pvcCalcStatus.textContent =
      `已儲存 ${series} 規格 ${key} mm；成本 ${costPerM.toFixed(4)} 元/M，售價約 ${pricePerM.toFixed(4)} 元/M。` +
      " 下次上傳銷貨檔時會自動套用這個成本。";
  });
}

/***********************
 * 主計算：
 *  0) PVC 成本（CFT-3 / CFT-6，不吃匯率，來自 PVC_STORAGE）
 *  1) 雲林熱縮（不吃匯率）
 *  2) 順博 / 瑞普（吃匯率，含顏色加價）
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
      "提醒：尚未輸入有效匯率，順博 / 瑞普的銷貨成本與毛利會顯示為 0（雲林熱縮與 PVC 不受影響）。";
  }

  processedRows = baseRows.map((row) => {
    let unitPrice = 0;
    let cost = 0;

    const codeUpper = (row.itemCode || "").toUpperCase();
    const nameText = (row.name || "").toString();

    // 0️⃣ 先試 PVC 成本（CFT-3 / CFT-6）
    const pvcUnit = getPvcUnitPrice(row.itemCode, row.specMm, row.name);
    if (pvcUnit != null) {
      unitPrice = pvcUnit;           // 元 / 米（不吃匯率）
      cost = unitPrice * row.meters;
    } else {
      // 1️⃣ 雲林熱縮（Hxx，不吃匯率）
      const yunlinUnit = getYunlinUnitPrice(
        row.itemCode,
        row.specMm,
        row.name
      );
      if (yunlinUnit != null) {
        unitPrice = yunlinUnit;
        cost = unitPrice * row.meters;
      } else {
        // 2️⃣ 順博 / 瑞普（吃匯率）
        const supplier = getSupplierFromRow(row);
        const mmKey = row.specMm != null ? String(row.specMm) : null;

        if (supplier && mmKey && hasRate) {
          let basePrice = getBasePriceFromCostTable(mmKey, supplier);

          if (Number.isFinite(basePrice)) {
            const isWhite =
              nameText.includes("白") ||
              /W$/.test(codeUpper);

            const isColor =
              !isWhite &&
              (/(黑|紅|藍|綠|黃)/.test(nameText) ||
                /(R|BL|G|Y)$/.test(codeUpper));

            // FSG-3 彩色 +5%
            if (supplier === "shunbo" && codeUpper.includes("FSG-3") && isColor) {
              basePrice *= 1.05;
            }

            // HST 彩色 +8%
            if (supplier === "shunbo" && codeUpper.includes("HST") && isColor) {
              basePrice *= 1.08;
            }

            unitPrice = basePrice * rateVal; // 台幣 / 米
            cost = unitPrice * row.meters;
          }
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

  // 排除 Z043 / Z044 / A1
  processedRows = processedRows.filter((row) => {
    const code = (row.itemCode || "").toUpperCase();
    return (
      !code.startsWith("Z043") &&
      !code.startsWith("Z044") &&
      code !== "A1"
    );
  });

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
