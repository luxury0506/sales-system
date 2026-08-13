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


/**
 * 計算單筆資料的換算米數
 * @param {string} itemName - 品名或物品編號 (例如: "CFT-3-0A-530M")
 * @param {number} salesQty - 銷貨量
 * @} returns {number} 換算後的總米數
 */
function calculateMeters(itemName, salesQty) {
  // 確保 itemName 是字串且去除前後空白
  const name = String(itemName || '').trim();
  
  // 1. 優先規則：利用正規表示法檢查是否以數字加 M 或 m 結尾
  // 匹配範例：CFT-3-0A-530M -> 抓出 530
  const matchM = name.match(/-(\d+(?:\.\d+)?)[Mm]$/); 
  
  if (matchM) {
    // 抓到品名尾碼的米數
    const unitMeter = parseFloat(matchM[1]); 
    
    // 總米數 = 品名單件米數 * 銷貨量
    // (如果明細中的 530M 本身就是總量、銷貨量只是1，乘上去也沒影響；如果出貨2捲就是 530 * 2)
    return unitMeter * salesQty; 
  }

  // 2. 次要規則：如果是 PVC 系列的特殊計算 (從主頁面 inputs 抓取或帶入公式)
  // if (isPvcSeries) { ... }

  // 3. 基礎規則：查閱 cost-data.js 的預設對照表
  // const defaultRate = costData[name]?.conversionRate || 1;
  // return salesQty * defaultRate;
  
  return salesQty; // 預設防呆
}

/***********************
 * CFT-3 / CFT-6 規格資料（取自 2604銷售分析.xlsx「PVC原料成本」工作表）
 * 用來：
 *   ① 從物品編號解析 AWG／規格碼 → 內徑(mm)、壁厚(mm)
 *      （銷售明細的品名常常只有 AWG 代碼，沒有直接寫內徑 mm，
 *        例如「PVC高壓套管 0A* 100mm」「PVC高壓套管 8AWG(透明)」）
 *   ② 一鍵計算並存入 PVC_COST_CFT3 / PVC_COST_CFT6
 ************************/
const CFT_SPEC_TABLE = {
  "CFT-3": [
    ["0A", 8.38, 0.6], ["1A", 7.47, 0.6], ["2A", 6.68, 0.6], ["3A", 5.94, 0.6],
    ["4A", 5.28, 0.55], ["5A", 4.72, 0.55], ["6A", 4.22, 0.55], ["7A", 3.76, 0.55],
    ["8A", 3.38, 0.55], ["9A", 3, 0.55], ["10A", 2.69, 0.55], ["11A", 2.41, 0.45],
    ["12A", 2.16, 0.45], ["13A", 1.93, 0.45], ["14A", 1.68, 0.45], ["15A", 1.5, 0.45],
    ["16A", 1.35, 0.45], ["17A", 1.19, 0.45], ["18A", 1.07, 0.45], ["19A", 0.97, 0.45],
    ["20A", 0.86, 0.45], ["5/16A", 7.94, 0.6], ["3/8A", 9.53, 0.6],
  ],
  "CFT-6": [
    ["0A", 8.38, 0.7], ["1A", 7.47, 0.7], ["2A", 6.68, 0.7], ["3A", 5.94, 0.7],
    ["4A", 5.28, 0.7], ["5A", 4.72, 0.7], ["6A", 4.22, 0.7], ["7A", 3.76, 0.7],
    ["8A", 3.38, 0.7], ["9A", 3, 0.7], ["10A", 2.69, 0.7], ["11A", 2.41, 0.7],
    ["12A", 2.16, 0.7], ["13A", 1.93, 0.7], ["14A", 1.68, 0.7], ["15A", 1.5, 0.7],
    ["16A", 1.35, 0.7], ["17A", 1.19, 0.7], ["18A", 1.07, 0.7], ["19A", 0.97, 0.7],
    ["20A", 0.86, 0.7], ["5/16A", 7.94, 0.7], ["3/8A", 9.53, 0.75], ["7/16A", 11.1, 0.75],
    ["1/2A", 12.7, 0.75], ["9/16A", 14, 0.85], ["5/8A", 15.9, 0.85],
  ],
};

// AWG／規格碼 -> 內徑(mm)。兩個系列的內徑相同，只有壁厚不同，合併成一張查表方便從物品編號反查。
const AWG_TO_INNER_MM = {};
Object.values(CFT_SPEC_TABLE).forEach((rows) => {
  rows.forEach(([label, innerMm]) => {
    AWG_TO_INNER_MM[label.toUpperCase()] = innerMm;
  });
});

/**
 * 從物品編號解析出 CFT 系列與 AWG／規格碼，回傳內徑(mm)。
 * 例："CFT-3-0A-100" -> "0A" -> 8.38
 *     "CFT-3-2AC"     -> 去掉尾端顏色字母 "C" -> "2A" -> 6.68
 *     "CFT-6-7/16A-550" -> "7/16A" -> 11.1
 *     "CFT-3-09-300"  -> 純數字規格 "09" -> 9（沒有 AWG 字尾的內徑代號）
 * 找不到就回傳 null，呼叫端會退回原本用品名文字抓 mm 的方式。
 */
function getCftInnerMmFromCode(itemCode) {
  const code = (itemCode || "").toString().trim().toUpperCase();
  if (!code.startsWith("CFT-3") && !code.startsWith("CFT-6")) return null;

  const parts = code.split("-");
  // parts: ["CFT","3","0A","100"] 或 ["CFT","6","7/16A","550"] 等
  if (parts.length < 3) return null;
  let specToken = parts[2];
  if (!specToken) return null;

  // 直接命中（"0A"、"7/16A"、"3/8A" 這種）
  if (AWG_TO_INNER_MM[specToken] != null) return AWG_TO_INNER_MM[specToken];

  // 去掉尾端顏色／材質字母後再試一次（"2AC" -> "2A"，"8AC" -> "8A"）
  if (specToken.length > 1 && /A[A-Z]$/.test(specToken)) {
    const withoutColorSuffix = specToken.slice(0, -1);
    if (AWG_TO_INNER_MM[withoutColorSuffix] != null) {
      return AWG_TO_INNER_MM[withoutColorSuffix];
    }
  }

  // 純數字規格碼（沒有 AWG 字尾），例如 "09" -> 9、"9" -> 9
  if (/^\d+(\.\d+)?$/.test(specToken)) {
    const v = parseFloat(specToken);
    if (Number.isFinite(v) && v > 0) return v;
  }

  // 純數字規格碼 + 尾端顏色字母，例如 "12C" -> "12" -> 12
  if (/^\d+(\.\d+)?[A-Z]$/.test(specToken)) {
    const v = parseFloat(specToken.slice(0, -1));
    if (Number.isFinite(v) && v > 0) return v;
  }

  return null;
}


function extractMmInfo(name) {
  if (!name) return { specMm: null, cutMm: null };

  let specMm = null;
  let cutMm = null;

  // 先找「x」或「*」分隔符，把品名拆成「直徑部分」跟「長度部分」，
  // 例如 "3.0mm* 2300mm" -> 前段"3.0mm"、後段"2300mm"。
  // 這樣即使前段漏打"mm"（例："3.0* 2300mm"），也能正確抓到直徑是3.0，
  // 不會被後面的長度數字"2300"誤判成直徑。
  const sepIdx = name.search(/[x*]/i);
  if (sepIdx !== -1) {
    const beforePart = name.slice(0, sepIdx);
    const afterPart = name.slice(sepIdx + 1);

    // 前段：抓最後一個數字，"mm"可有可無
    const diaMatch = /([\d.]+)\s*(?:mm)?\s*$/i.exec(beforePart.trim());
    if (diaMatch) {
      const v = parseFloat(diaMatch[1]);
      if (!isNaN(v)) specMm = v;
    }

    // 後段：抓第一個「數字+mm」當裁切長度
    const cutMatch = /([\d.]+)\s*mm/i.exec(afterPart);
    if (cutMatch) {
      const v = parseFloat(cutMatch[1]);
      if (!isNaN(v)) cutMm = v;
    }
  }

  // 沒有分隔符時，才退回「抓第一個 XXmm」的舊邏輯；
  // 有分隔符但前段抓不到數字（例如AWG式代號），交給下面的AWG換算處理，
  // 不要在整個品名裡亂抓，否則可能又抓到後段的裁切長度。
  if (specMm == null && sepIdx === -1) {
    const mmMatch = /([\d.]+)\s*mm/i.exec(name);
    if (mmMatch) {
      const v = parseFloat(mmMatch[1]);
      if (!isNaN(v)) specMm = v;
    }
  }

  // ② 如果沒有寫 mm，但有 "3/8AWG"、"7/16AWG" 這類字樣，就把英吋換算成 mm
  if (specMm == null) {
    const awgMatch = /(\d+)\s*\/\s*(\d+)\s*AWG/i.exec(name);
    if (awgMatch) {
      const num = parseFloat(awgMatch[1]);
      const den = parseFloat(awgMatch[2]);
      if (Number.isFinite(num) && Number.isFinite(den) && den > 0) {
        // 1 inch = 25.4 mm
        specMm = (25.4 * num) / den;
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
 * PET／AIS／AISC／YG（雲林電子，不吃匯率）
 * 編碼規則（由實際物品編號範例確認）：
 *   PET-032 = 3.2mm、AIS-254R = 25.4mm（R=顏色字尾，比對前先去掉）
 *   規則：去掉可能的顏色字尾字母，剩下的數字 ÷10 = mm
 *   例：PET-032 -> 032 -> 3.2mm；AIS-127 -> 127 -> 12.7mm
 * YG 的實際編碼規則尚未經過真實範例確認，先比照同一家供應商(雲林電子)
 * 的 PET/AIS 編碼慣例套用，如果之後發現對不上，要再調整。
 ************************/
function parseTenthsMmCode(specToken) {
  if (!specToken) return null;
  // 去掉尾端顏色／材質字母（例："254R" -> "254"）
  const stripped = specToken.replace(/[A-Za-z]+$/, "");
  if (!/^\d+$/.test(stripped)) return null;
  // 標準格式是3碼（mm×10），但實務上偶爾會把尾端的0省略掉，
  // 例如 PET-16 其實代表 PET-160（=16.0mm），不是 1.6mm。
  // 所以不足3碼時，先在後面補0再除以10。
  const padded = stripped.length < 3 ? stripped.padEnd(3, "0") : stripped;
  const v = parseInt(padded, 10) / 10;
  return Number.isFinite(v) && v > 0 ? v : null;
}

function getMmTableUnitPrice(itemCode, prefix, table) {
  if (!itemCode || typeof table === "undefined") return null;
  const code = itemCode.toString().trim().toUpperCase();
  const parts = code.split("-");
  if (parts.length < 2 || parts[0] !== prefix) return null;

  const mm = parseTenthsMmCode(parts[1]);
  if (mm == null) return null;

  const mmKey = String(mm);
  const row = table[mmKey];
  if (row == null) return null;

  // YG_TABLE / PET_TABLE / AIS_TABLE / AISC_TABLE 都是「mm -> 價格數字」，
  // 直接回傳；如果是 {cost:...} 這種物件形式（像 HTK_TABLE）在這裡不適用。
  const price = typeof row === "number" ? row : null;
  return price != null && price > 0 ? price : null;
}

function getPetUnitPrice(itemCode) {
  if (typeof PET_TABLE === "undefined") return null;
  return getMmTableUnitPrice(itemCode, "PET", PET_TABLE);
}

function getAisUnitPrice(itemCode) {
  if (typeof AIS_TABLE === "undefined") return null;
  return getMmTableUnitPrice(itemCode, "AIS", AIS_TABLE);
}

function getAiscUnitPrice(itemCode) {
  if (typeof AISC_TABLE === "undefined") return null;
  return getMmTableUnitPrice(itemCode, "AISC", AISC_TABLE);
}

function getYgUnitPrice(itemCode) {
  if (typeof YG_TABLE === "undefined") return null;
  return getMmTableUnitPrice(itemCode, "YG", YG_TABLE);
}

/***********************
 * HTK／ATM（雲林電子，不吃匯率）
 * 實際銷貨編號跟 PET/AIS 同一套「PREFIX-XXX」(mm×10) 規則，
 * 例：HTK-032 = 3.2mm、ATM-080 = 8.0mm。
 ************************/
function getHtkUnitPrice(itemCode) {
  if (typeof HTK_TABLE === "undefined") return null;
  return getMmTableUnitPrice(itemCode, "HTK", HTK_TABLE);
}

function getAtmUnitPrice(itemCode) {
  // ATM 系列跟 PET/AIS/HTK 不一樣：代號是「直接整數mm」，不是mm×10。
  // 由實際品名對照確認：ATM-012BK = 12.0mm（不是1.2mm）、ATM-022BK = 22.0mm。
  if (typeof ATM_TABLE === "undefined" || !itemCode) return null;
  const code = itemCode.toString().trim().toUpperCase();
  const parts = code.split("-");
  if (parts.length < 2 || parts[0] !== "ATM") return null;

  const stripped = parts[1].replace(/[A-Za-z]+$/, "");
  if (!/^\d+$/.test(stripped)) return null;
  const mm = parseInt(stripped, 10);
  if (!Number.isFinite(mm) || mm <= 0) return null;

  const row = ATM_TABLE[String(mm)];
  return typeof row === "number" && row > 0 ? row : null;
}

/***********************
 * 統一入口：依序試 PET / AIS / AISC / YG / HTK / ATM
 ************************/
function getOtherHeatShrinkUnitPrice(itemCode) {
  return (
    getPetUnitPrice(itemCode) ??
    getAisUnitPrice(itemCode) ??
    getAiscUnitPrice(itemCode) ??
    getYgUnitPrice(itemCode) ??
    getHtkUnitPrice(itemCode) ??
    getAtmUnitPrice(itemCode)
  );
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
 * localStorage 結構（依顏色分開儲存）：
 * {
 *   cft3: {
 *     black:       { "8.380": 單價元/M, ... },
 *     transparent: { "8.380": 單價元/M, ... },
 *     color:       { "8.380": 單價元/M, ... }
 *   },
 *   cft6: { black: {...}, transparent: {...}, color: {...} },
 *   updatedAt: ISOString
 * }
 ************************/
const PVC_STORAGE_KEY = "PVC_COST_TABLE";
let PVC_COST_CFT3 = { black: {}, transparent: {}, color: {} };
let PVC_COST_CFT6 = { black: {}, transparent: {}, color: {} };

function normalizePvcSeriesTable(raw) {
  // 兼容舊格式（沒分顏色、mm直接對成本的扁平物件）：當作黑色資料處理。
  if (!raw) return { black: {}, transparent: {}, color: {} };
  const looksNested =
    raw.black || raw.transparent || raw.color;
  if (looksNested) {
    return {
      black: raw.black || {},
      transparent: raw.transparent || {},
      color: raw.color || {},
    };
  }
  return { black: raw, transparent: {}, color: {} };
}

function loadPvcCostFromLocalStorage() {
  try {
    const raw = localStorage.getItem(PVC_STORAGE_KEY);
    if (!raw) return;
    const obj = JSON.parse(raw);
    PVC_COST_CFT3 = normalizePvcSeriesTable(obj.cft3);
    PVC_COST_CFT6 = normalizePvcSeriesTable(obj.cft6);
  } catch (e) {
    console.error("讀取 PVC 成本失敗：", e);
  }
}

/**
 * 從物品編號／品名判斷 PVC 顏色分類：black／transparent／color。
 * 規則跟雲林熱縮的顏色判斷比照：字尾 C=透明，R/BL/G/Y=彩色，其餘=黑色。
 */
function getPvcColorType(itemCode, name) {
  const code = (itemCode || "").toString().toUpperCase();
  const text = (name || "").toString();

  if (/C$/.test(code) || text.includes("透") || text.includes("透明")) {
    return "transparent";
  }
  if (
    /(R|BL|G|Y)$/.test(code) ||
    /(紅|藍|綠|黃)/.test(text)
  ) {
    return "color";
  }
  return "black";
}

/**
 * 根據物品編號 + 規格(mm) 取得 PVC 每米成本（元/M，不吃匯率），
 * 會依顏色（黑/透明/彩色）分開查對應的成本表。
 */
function getPvcUnitPrice(itemCode, specMm, name) {
  const codeUpper = (itemCode || "").toUpperCase();
  const text = (name || "").toString();
  const d = specMm != null ? parseFloat(specMm) : NaN;
  if (!Number.isFinite(d)) return null;

  let seriesTable = null;

  if (codeUpper.startsWith("CFT-3")) {
    seriesTable = PVC_COST_CFT3;
  } else if (codeUpper.startsWith("CFT-6")) {
    seriesTable = PVC_COST_CFT6;
  } else if (text.includes("PVC高壓套管") || text.includes("PVC套管")) {
    // 若之後有其它 PVC 物品編號格式，再補判斷；目前先限制 CFT-3/6
    return null;
  } else {
    return null;
  }

  const colorType = getPvcColorType(itemCode, name);
  let table = seriesTable[colorType];

  // 該顏色沒有資料就退回黑色（通常黑色最完整，總比查不到好）
  if (!table || !Object.keys(table).length) {
    table = seriesTable.black;
  }
  if (!table) return null;

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

        // 跳過「小計」「合計」「總計」這種每個客戶群組結尾的加總列，
        // 這種列本身就是上面幾筆交易的加總，不是真正的商品交易；
        // 如果不跳過，會被當成一筆新的交易再加一次，導致總量被重複計算。
        if (name === "小計" || name === "合計" || name === "總計") continue;

        const qty = parseNumber(row[qtyColIndex]);
        if (!Number.isFinite(qty)) continue;

        const amount =
          amountColIndex !== -1 ? parseNumber(row[amountColIndex]) : 0;

        const { specMm: textSpecMm, cutMm } = extractMmInfo(name);

        // CFT-3 / CFT-6 品項：優先用物品編號解析出來的內徑(mm)，
        // 因為品名常常只寫 AWG 代碼（例如「0A* 100mm」「8AWG(透明)」），
        // 用文字規則抓到的第一個 mm 數字常常是「裁切長度」而不是「內徑」。
        const codeSpecMm = getCftInnerMmFromCode(itemCode);
        const specMm = codeSpecMm != null ? codeSpecMm : textSpecMm;

        let meters = qty;
        // 新增規則：如果物品編號或品名結尾有數字+M (排除 mm)，代表銷貨數量就是總米數，不須做裁切(÷1000)換算
        const isMeterUnit = 
          (/[0-9][Mm]$/.test(itemCode) && !/mm$/i.test(itemCode)) || 
          (/[0-9][Mm]$/.test(name)     && !/mm$/i.test(name));

        if (isMeterUnit) {
          meters = qty;
        } else if (cutMm != null) {
          meters = qty * (cutMm / 1000);
        }

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

// 成本表(順博/瑞普)裡沒有登記、但確認要比照鄰近規格計價的直徑對應。
// 例：FSG-2-017(1.7mm) 比照 FSG-2-015(1.5mm) 計價。
// 有新的例外請直接加在這裡。
const MM_SPEC_COST_ALIAS = {
  "1.7": "1.5",  // FSG-2-017 -> FSG-2-015
  "3.2": "3",    // FSG-3-032 -> FSG-3-03
  "3.7": "3.5",  // FSG-3-035-038M(品名3.7mm) -> FSG-3-035
  "4.3": "4.5",  // HST-043 -> HST-045
  "4.6": "4.5",  // HST-046系列 -> HST-045
};

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
      "提醒：尚未輸入有效匯率，順博 / 瑞普的銷貨成本與毛利會顯示為 0（雲林熱縮、PET/AIS/AISC/YG/HTK/ATM 與 PVC 不受影響）。";
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
        // 1.5️⃣ 其他熱縮系列：PET / AIS / AISC / YG / HTK / ATM（不吃匯率）
        const otherUnit = getOtherHeatShrinkUnitPrice(row.itemCode);
        if (otherUnit != null) {
          unitPrice = otherUnit;
          cost = unitPrice * row.meters;
        } else {
        // 2️⃣ 順博 / 瑞普（吃匯率）
        const supplier = getSupplierFromRow(row);
        let mmKey = row.specMm != null ? String(row.specMm) : null;
        if (mmKey != null && MM_SPEC_COST_ALIAS[mmKey] != null) {
          mmKey = MM_SPEC_COST_ALIAS[mmKey];
        }

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

  // 讓頁面上其他獨立功能（例如「本月損益總覽」）可以讀到這批明細的合計，
  // 不用重新解析一次表格。
  window.__currentFileTotals = {
    totalQty,
    totalMeters,
    totalAmount,
    totalCost,
    totalProfit,
    rowCount: processedRows.length,
    updatedAt: new Date().toISOString(),
  };

  // 依物品編號彙總這批明細的用量，供「叫貨量預估」功能記錄歷史用。
  const itemBreakdown = {};
  processedRows.forEach((row) => {
    const code = row.itemCode || "未填寫物品編號";
    if (!itemBreakdown[code]) {
      itemBreakdown[code] = { qty: 0, meters: 0, name: row.name || "" };
    }
    itemBreakdown[code].qty += row.qty || 0;
    itemBreakdown[code].meters += row.meters || 0;
  });
  window.__currentFileItemBreakdown = itemBreakdown;

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
  window.__currentFileTotals = null;
  window.__currentFileItemBreakdown = null;
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
