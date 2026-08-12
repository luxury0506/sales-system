// cost-data.js - 從《大陸成富報價單(2022表格).xlsx》與《雲林電子2021報價單..xlsx》整理
// 單位說明：
//   順博 / 瑞普：每米 單價（人民幣）→ 需要乘匯率（台幣／人民幣）
//   雲林 G5：每米 單價（新台幣）→ 不需要匯率

// =====================================================
// 一、順博 / 瑞普：FSG-2、FSG-3、HST、SRG 用的成本
// 結構：COST_TABLE[廠商][mm規格] = 每米單價（人民幣）
// =====================================================
const COST_TABLE = {
  shunbo: {
    "0.8": 0.1536,
    "1": 0.1455,
    "1.2": 0.1546,
    "1.5": 0.1636,
    "2": 0.1818,
    "2.5": 0.2182,
    "3": 0.2364,
    "3.5": 0.2818,
    "4": 0.3455,
    "4.5": 0.4091,
    "5": 0.4636,
    "5.5": 0.5472,
    "6": 0.5727,
    "6.5": 0.5864,
    "7": 0.6,
    "8": 0.7,
    "9": 0.8,
    "10": 0.9,
    "11": 1.0,
    "12": 1.1,
    "13": 1.2,
    "14": 1.3,
    "15": 1.4,
    "16": 1.5,
    "18": 1.7,
    "20": 1.9,
    "22": 2.1,
    "25": 2.4,
    "28": 2.7,
    "30": 2.9,
    "32": 3.2,
    "35": 3.5,
    "38": 3.8,
    "40": 4.0,
    "45": 4.5
  },
  ruipu: {
    "1": 0.15,
    "1.5": 0.16,
    "2": 0.17,
    "2.5": 0.2,
    "3": 0.22,
    "3.5": 0.26,
    "4": 0.32,
    "4.5": 0.41,
    "5": 0.45,
    "6": 0.57,
    "7": 0.58,
    "8": 0.7,
    "10": 0.95,
    "12": 1.2,
    "14": 1.5,
    "16": 2.0,
    "18": 2.37,
    "20": 2.8,
    "22": 3.4,
    "25": 4.6,
    "30": 7.0,
    "35": 8.0
  }
};

// 轉成 Map 方便查詢： key = "shunbo|2.0" → 0.1818
window.COST_MAP = new Map();
for (const [supplier, mmTable] of Object.entries(COST_TABLE)) {
  for (const [mm, price] of Object.entries(mmTable)) {
    window.COST_MAP.set(`${supplier}|${mm}`, price);
  }
}

// =====================================================
// 二、雲林電子 G5 熱收縮套管單價（每米，新台幣）
// 結構：YUNLIN_G5[內徑mm] = { black, color, transparent, thin, no_print }
//   black       : 黑色
//   color       : 彩色 (R / BL / G / Y)
//   transparent : 透明 (料號結尾 C)
//   thin        : 超薄 (料號結尾 CB)
//   no_print    : 不印字（若有用到再擴充）
// =====================================================

const YUNLIN_G5 = {
  "0.8":  { black: 1.11,  color: 1.35, transparent: null,  thin: null,  no_print: null },
  "1":    { black: 0.72,  color: 0.87, transparent: 0.86, thin: 0.85,  no_print: 1.01 },
  "1.5":  { black: 0.81,  color: 0.95, transparent: 0.93, thin: 1.01,  no_print: 1.09 },
  "2":    { black: 0.87,  color: 1.02, transparent: 0.86, thin: 0.87,  no_print: 1.34 },
  "2.5":  { black: 0.95,  color: 1.17, transparent: 1.01, thin: 0.95,  no_print: null },
  "3":    { black: 1.03,  color: 1.33, transparent: 1.09, thin: 1.03,  no_print: null },
  "3.5":  { black: 1.19,  color: 1.6,  transparent: 1.24, thin: 1.19,  no_print: null },
  "4":    { black: 1.35,  color: 1.71, transparent: 1.4,  thin: 1.35,  no_print: null },
  "4.5":  { black: 1.43,  color: 1.86, transparent: 1.48, thin: 1.43,  no_print: null },
  "5":    { black: 1.75,  color: 2.23, transparent: 1.79, thin: 1.75,  no_print: null },
  "6":    { black: 1.99,  color: 2.73, transparent: 2.1,  thin: 1.99,  no_print: null },
  "7":    { black: 2.31,  color: 3.21, transparent: 2.41, thin: 2.31,  no_print: null },
  "8":    { black: 2.62,  color: 3.48, transparent: 2.72, thin: 2.62,  no_print: null },
  "9":    { black: 2.78,  color: 3.71, transparent: 2.88, thin: 2.78,  no_print: null },
  "10":   { black: 3.02,  color: 4.09, transparent: 3.11, thin: 3.02,  no_print: null },
  "11":   { black: 3.38,  color: 4.58, transparent: 3.47, thin: 3.34,  no_print: null },
  "12":   { black: 3.74,  color: 5.06, transparent: 3.83, thin: 3.58,  no_print: null },
  "13":   { black: 4.08,  color: 5.51, transparent: 4.16, thin: 4.13,  no_print: null },
  "14":   { black: 4.43,  color: 5.99, transparent: 4.51, thin: 4.45,  no_print: null },
  "15":   { black: 4.77,  color: 6.47, transparent: 4.86, thin: 4.85,  no_print: null },
  "16":   { black: 5.11,  color: 6.94, transparent: 5.19, thin: 5.49,  no_print: null },
  "18":   { black: 5.79,  color: 7.86, transparent: 5.88, thin: null,  no_print: null },
  "20":   { black: 6.47,  color: 8.79, transparent: 6.56, thin: null,  no_print: null },
  "22":   { black: 7.15,  color: 9.71, transparent: 7.24, thin: null,  no_print: null },
  "25":   { black: 8.28,  color: 11.35,transparent: 8.37, thin: null,  no_print: null },
  "28":   { black: 9.52,  color: 12.94,transparent: 9.61, thin: null,  no_print: null },
  "30":   { black: 10.16, color: 13.74,transparent: 10.25,thin: null,  no_print: null },
  "32":   { black: 10.8,  color: 14.52,transparent: 10.9, thin: null,  no_print: null },
  "35":   { black: 11.91, color: 16.01,transparent: 12.01,thin: null,  no_print: null },
  "38":   { black: 13.01, color: 17.49,transparent: 13.12,thin: null,  no_print: null },
  "40":   { black: 13.95, color: 18.74,transparent: 14.06,thin: null,  no_print: null },
  "45":   { black: 15.78, color: 21.03,transparent: 15.9, thin: null,  no_print: null },
  "50":   { black: 17.6,  color: 23.44,transparent: 17.72,thin: null,  no_print: null },
  "60":   { black: 21.62, color: 28.87,transparent: 21.76,thin: null,  no_print: null },
  "70":   { black: 25.1,  color: 33.51,transparent: 25.24,thin: null,  no_print: null },
  "80":   { black: 28.29, color: 37.81,transparent: 28.44,thin: null,  no_print: null },
  "90":   { black: 32.2,  color: 43.04,transparent: 32.36,thin: null,  no_print: null },
  "100":  { black: 35.96, color: 48.01,transparent: 36.13,thin: null,  no_print: null },
  "120":  { black: 44.55, color: 59.47,transparent: 44.75,thin: null,  no_print: null },
  "150":  { black: 58.26, color: 77.64,transparent: 58.49,thin: null,  no_print: null },
  "180":  { black: 72.3,  color: 96.64,transparent: 72.56,thin: null,  no_print: null }
};

// =====================================================
// 三、其他熱縮套管系列進價（每米，新台幣，不吃匯率）
// 資料來源：《熱縮管折數換算表o_2021.xlsx》，雲林電子 2021_11。
// 結構統一為：規格(mm，取到小數第1位當key) → 進價/M（元）
// 這些表目前只是「資料」，還沒有串進 app.js 的銷貨成本計算流程，
// 因為還不確定這幾個系列在實際銷貨明細裡的物品編號規則，
// 等確認後再由 app.js 呼叫對應的查價函式。
// =====================================================

// PET／擴充編織網管（雲林電子，規格為 mm，已從英吋規格換算）
const PET_TABLE = {
  "3.2": 3.36,
  "6.4": 3.9,
  "9.5": 4.5,
  "12.7": 5.7,
  "16": 8.41, // 原始表格這列缺規格標籤、mm寫成15，但實際物品編號PET-16確認是16.0mm(5/8")
  "19.1": 9.7,
  "32": 20.8,
  "38.1": 28.8,
  "44.5": 31.7,
  "50": 32.3,
};

// AIS-3X 含膠熱縮套管（雲林電子，規格為英吋標示的內徑，已換算成mm）
const AIS_TABLE = {
  "3.2": 8.2,     // 1/8"
  "4.8": 10.2,    // 3/16"
  "6.4": 13.3,    // 1/4"
  "9.5": 17.8,    // 3/8"
  "12.7": 34.7,   // 1/2"
  "15": 50.2,     // 5/8"
  "19.1": 52.4,   // 3/4"
  "25.4": 82.7,   // 1"
  "38.1": 110,    // 1-1/2"
};

// AIS-C 含膠熱縮套管（雲林電子，規格為英吋標示的內徑，已換算成mm）
const AISC_TABLE = {
  "3.2": 4.56,    // 1/8"
  "4.8": 6.6,     // 3/16"
  "6.4": 7.92,    // 1/4"
  "9.5": 13.2,    // 3/8"
  "12.7": 20.4,   // 1/2"
  "15": 26.4,     // 5/8"
  "19.1": 31.7,   // 3/4"
  "25.4": 43.4,   // 1"
  "38.1": 101.9,  // 1-1/2"
};

// YG 黃綠熱縮套管（雲林電子，規格mm已直接標示於原始表格）
const YG_TABLE = {
  "1.5": 2.4,
  "2.5": 2.3,
  "3": 2.49,
  "4.5": 3.23,
  "6": 3.96,
  "8": 5.25,
  "9": 5.62,
  "12": 7.92,
  "16": 11.98,
  "18": 14.01,
  "25": 21.84,
  "30": 29.12,
  "40": 40.92,
  "50": 56.03,
};

// HTK PVDF熱收縮套管（雲林電子）
// 實際銷貨編號格式跟 PET/AIS 一樣是「HTK-XXX」(mm×10)，例：HTK-032 = 3.2mm。
// 原始供應商型號（如 HTK-150-0032）僅留在註解對照，不用來查價。
const HTK_TABLE = {
  "1.2": 12.99,   // HTK-150-0012
  "1.6": 14.84,   // HTK-150-0016
  "2.4": 18.81,   // HTK-150-0024
  "3.2": 23.05,   // HTK-150-0032
  "4.8": 31.54,   // HTK-150-0048
  "6.4": 37.63,   // HTK-150-0064
  "9.5": 55.39,   // HTK-150-0095
  "12.7": 75.26,  // HTK-150-0127
  "19.1": 130.11, // HTK-150-0191
  "25.4": 191.2,  // HTK-150-0254
};

// ATM 通用型中壁含膠熱收縮套管（雲林電子）
// 實際銷貨編號格式是「ATM-XXX」(mm×10)，例：ATM-080 = 8.0mm。
// 原始供應商型號（如 ATM008BK）僅留在註解對照，不用來查價。
const ATM_TABLE = {
  "8": 34,     // ATM008BK
  "12": 48,    // ATM012BK
  "16": 59,    // ATM016BK
  "19": 64,    // ATM019BK
  "22": 69,    // ATM022BK
  "28": 71,    // ATM028BK
  "33": 77,    // ATM033BK
  "40": 89,    // ATM040BK
  "44": 99,    // ATM044BK
  "55": 115,   // ATM055BK
  "65": 146,   // ATM065BK
  "75": 165,   // ATM075BK
  "85": 216,   // ATM085BK
  "95": 228,   // ATM095BK
  "115": 283,  // ATM115BK
  "140": 374,  // ATM140BK
  "160": 508,  // ATM160BK
};

// 勳旺電子 G5黑（跟雲林電子YUNLIN_G5.black是同類產品的「另一個供應商」報價，
// 兩邊價格不同，暫不自動套用，先保留原始資料，等確認實際進貨都跟哪家供應商採購再決定怎麼用）
const XUNWANG_G5_BLACK_TABLE = {
  "1": 0.74,
  "1.5": 0.89,
  "2": 0.98,
  "2.5": 1.28,
  "3": 1.43,
  "3.5": 1.52,
  "4": 1.72,
  "4.5": 1.97,
  "5": 2.16,
  "5.5": 2.31,
  "6": 2.51,
};
