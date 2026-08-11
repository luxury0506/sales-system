// pvc.js
// PVC 管重量 + 成本利潤計算，並儲存每米成本到 localStorage
//
// 本頁是 PVC 成本的唯一編輯入口：主頁 index.html 只會「讀」這裡存的
// PVC_COST_TABLE（透過 app.js 的 loadPvcCostFromLocalStorage /
// getPvcUnitPrice），計算銷貨成本時使用；成本設定本身統一在這裡改。

const PVC_STORAGE_KEY = "PVC_COST_TABLE";

let pvcCost = {
  cft3: {},      // 規格(mm) → 每米成本（台幣）
  cft6: {},      // 規格(mm) → 每米成本（台幣）
  updatedAt: null,
};

// CFT-3／CFT-6 全部 AWG／內徑規格表（比照《2604銷售分析.xlsx》PVC原料成本工作表）
// label 為 "-" 代表沒有對應的 AWG 代號，只是額外的內徑規格，一樣可以存成本、
// 之後主頁用「最接近的內徑」去查表時仍然找得到。
const CFT_SPEC_TABLE = {
  "CFT-3": [
    ["0A", 8.38, 0.6], ["1A", 7.47, 0.6], ["2A", 6.68, 0.6], ["3A", 5.94, 0.6],
    ["4A", 5.28, 0.55], ["5A", 4.72, 0.55], ["6A", 4.22, 0.55], ["7A", 3.76, 0.55],
    ["8A", 3.38, 0.55], ["9A", 3, 0.55], ["10A", 2.69, 0.55], ["11A", 2.41, 0.45],
    ["12A", 2.16, 0.45], ["13A", 1.93, 0.45], ["14A", 1.68, 0.45], ["15A", 1.5, 0.45],
    ["16A", 1.35, 0.45], ["17A", 1.19, 0.45], ["18A", 1.07, 0.45], ["19A", 0.97, 0.45],
    ["20A", 0.86, 0.45], ["5/16A", 7.94, 0.6], ["3/8A", 9.53, 0.6],
    ["-", 2, 0.45], ["-", 2.6, 0.55], ["-", 2.8, 0.55], ["-", 5.5, 0.55], ["-", 6.7, 0.6],
    ["-", 7.8, 0.6], ["-", 8, 0.6], ["-", 9, 0.6], ["-", 7.3, 0.6],
  ],
  "CFT-6": [
    ["0A", 8.38, 0.7], ["1A", 7.47, 0.7], ["2A", 6.68, 0.7], ["3A", 5.94, 0.7],
    ["4A", 5.28, 0.7], ["5A", 4.72, 0.7], ["6A", 4.22, 0.7], ["7A", 3.76, 0.7],
    ["8A", 3.38, 0.7], ["9A", 3, 0.7], ["10A", 2.69, 0.7], ["11A", 2.41, 0.7],
    ["12A", 2.16, 0.7], ["13A", 1.93, 0.7], ["14A", 1.68, 0.7], ["15A", 1.5, 0.7],
    ["16A", 1.35, 0.7], ["17A", 1.19, 0.7], ["18A", 1.07, 0.7], ["19A", 0.97, 0.7],
    ["20A", 0.86, 0.7], ["5/16A", 7.94, 0.7], ["3/8A", 9.53, 0.75], ["7/16A", 11.1, 0.75],
    ["1/2A", 12.7, 0.75], ["9/16A", 14, 0.85], ["5/8A", 15.9, 0.85],
    ["3/8A-T112", 9.53, 1.2], ["9A-T08", 3, 0.8], ["083-T10", 8.3, 1],
    ["105-T10", 10.5, 1], ["15A-T11", 1.5, 1.2], ["09-T01", 9, 1], ["10-T10", 10, 1],
    ["-", 10, 0.75], ["-", 12, 0.75], ["-", 15, 0.85], ["-", 7.8, 1], ["-", 9, 0.8], ["-", 16, 0.85],
  ],
};

function loadPvcFromStorage() {
  const raw = localStorage.getItem(PVC_STORAGE_KEY);
  if (!raw) return;
  try {
    const obj = JSON.parse(raw);
    pvcCost = {
      cft3: obj.cft3 || {},
      cft6: obj.cft6 || {},
      updatedAt: obj.updatedAt || null,
    };
  } catch (e) {
    console.error("PVC_STORAGE 解析失敗：", e);
  }
}

function savePvcToStorage() {
  pvcCost.updatedAt = new Date().toISOString();
  localStorage.setItem(PVC_STORAGE_KEY, JSON.stringify(pvcCost));
}

function renderPvcTables() {
  const infoEl = document.getElementById("savedInfo");
  const body3 = document.getElementById("tableCft3");
  const body6 = document.getElementById("tableCft6");

  body3.innerHTML = "";
  body6.innerHTML = "";

  const keys3 = Object.keys(pvcCost.cft3);
  const keys6 = Object.keys(pvcCost.cft6);

  if (!keys3.length && !keys6.length) {
    infoEl.textContent = "目前尚未儲存任何 PVC 成本資料。";
    return;
  }

  infoEl.textContent =
    `CFT-3 規格 ${keys3.length} 筆，CFT-6 規格 ${keys6.length} 筆` +
    (pvcCost.updatedAt
      ? `，最後更新：${new Date(pvcCost.updatedAt).toLocaleString()}`
      : "");

  keys3
    .sort((a, b) => parseFloat(a) - parseFloat(b))
    .slice(0, 50)
    .forEach((k) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${k}</td><td>${pvcCost.cft3[k].toFixed(4)}</td>`;
      body3.appendChild(tr);
    });

  keys6
    .sort((a, b) => parseFloat(a) - parseFloat(b))
    .slice(0, 50)
    .forEach((k) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${k}</td><td>${pvcCost.cft6[k].toFixed(4)}</td>`;
      body6.appendChild(tr);
    });
}

document.addEventListener("DOMContentLoaded", () => {
  loadPvcFromStorage();
  renderPvcTables();

  const innerDiameterInput = document.getElementById("innerDiameter");
  const innerRadiusInput = document.getElementById("innerRadius");
  const innerAreaInput = document.getElementById("innerArea");
  const thicknessInput = document.getElementById("thickness");
  const outerDiameterInput = document.getElementById("outerDiameter");
  const outerAreaInput = document.getElementById("outerArea");
  const densityInput = document.getElementById("density");
  const weightPerMInput = document.getElementById("weightPerM");
  const weightStatus = document.getElementById("weightStatus");

  const weight305Input = document.getElementById("weight305");
  const weight100Input = document.getElementById("weight100");
  const wastePercentInput = document.getElementById("wastePercent");
  const pelletPriceInput = document.getElementById("pelletPrice");
  const costPerMInput = document.getElementById("costPerM");
  const profitPercentInput = document.getElementById("profitPercent");
  const pricePerMInput = document.getElementById("pricePerM");
  const costStatus = document.getElementById("costStatus");

  const seriesSelect = document.getElementById("seriesSelect");
  const specForSaveInput = document.getElementById("specForSave");
  const saveStatus = document.getElementById("saveStatus");

  let latestCostPerM = null; // 目前計算出的成本（元/M）

  // 計算管重量
  document.getElementById("calcWeightBtn").addEventListener("click", () => {
    const inner = parseFloat(innerDiameterInput.value);
    const t = parseFloat(thicknessInput.value);
    const density = parseFloat(densityInput.value);

    if (!Number.isFinite(inner) || inner <= 0) {
      weightStatus.innerHTML = `<span class="error">請輸入有效的內徑。</span>`;
      return;
    }
    if (!Number.isFinite(t) || t <= 0) {
      weightStatus.innerHTML = `<span class="error">請輸入有效的厚度。</span>`;
      return;
    }
    if (!Number.isFinite(density) || density <= 0) {
      weightStatus.innerHTML = `<span class="error">請輸入有效的比重。</span>`;
      return;
    }

    const innerRadius = inner / 2;
    const outer = inner + 2 * t;
    const outerRadius = outer / 2;

    const innerArea = Math.PI * innerRadius * innerRadius; // mm²
    const outerArea = Math.PI * outerRadius * outerRadius; // mm²

    // ✅ 每米重(g) = (外圓面積 - 內圓面積) × 比重
    const areaDiff = outerArea - innerArea; // mm²
    const weightPerM = areaDiff * density;  // g/M （因為 mm²→cm² 與長度 1M 抵消）

    innerRadiusInput.value = innerRadius.toFixed(3);
    innerAreaInput.value = innerArea.toFixed(3);
    outerDiameterInput.value = outer.toFixed(3);
    outerAreaInput.value = outerArea.toFixed(3);
    weightPerMInput.value = weightPerM.toFixed(3);

    // 順便算出 305M / 100M 重量
    const weight305 = (weightPerM * 305) / 1000; // Kg
    const weight100 = (weightPerM * 100) / 1000; // Kg
    weight305Input.value = weight305.toFixed(3);
    weight100Input.value = weight100.toFixed(3);

    weightStatus.innerHTML = `<span class="ok">已完成重量計算。</span>`;

    // 若 specForSave 還沒填，預設帶入內徑
    if (!specForSaveInput.value) {
      specForSaveInput.value = inner.toFixed(3);
    }
  });

  // 計算成本與售價
  document.getElementById("calcCostBtn").addEventListener("click", () => {
    const weightPerM = parseFloat(weightPerMInput.value); // g/M
    if (!Number.isFinite(weightPerM) || weightPerM <= 0) {
      costStatus.innerHTML = `<span class="error">請先計算每米重量。</span>`;
      return;
    }

    const wastePercent = parseFloat(wastePercentInput.value) || 0;
    const pelletPrice = parseFloat(pelletPriceInput.value);
    const profitPercent = parseFloat(profitPercentInput.value) || 0;

    if (!Number.isFinite(pelletPrice) || pelletPrice <= 0) {
      costStatus.innerHTML = `<span class="error">請輸入有效的 PVC 粒單價。</span>`;
      return;
    }

    const wasteRate = wastePercent / 100;
    const profitRate = profitPercent / 100;

    // 每米重量 (Kg)
    const weightKgPerM = weightPerM / 1000;

    // ✅ 材料成本(元/M) = 重量(Kg/M) × PVC粒(元/Kg) × (1 + 廢料%)
    const materialPricePerKg = pelletPrice * (1 + wasteRate);
    const costPerM = weightKgPerM * materialPricePerKg;

    // ✅ 售價(元/M) = 成本 × (1 + 利潤%)
    const pricePerM = costPerM * (1 + profitRate);

    latestCostPerM = costPerM;

    costPerMInput.value = costPerM.toFixed(4);
    pricePerMInput.value = pricePerM.toFixed(4);

    costStatus.innerHTML = `<span class="ok">已完成成本與售價計算。</span>`;
  });

  // 儲存為 PVC 成本（給主系統用）
  document.getElementById("saveCurrentPvcBtn").addEventListener("click", () => {
    if (!Number.isFinite(latestCostPerM) || latestCostPerM <= 0) {
      saveStatus.innerHTML = `<span class="error">請先完成成本計算（成本 元/M）。</span>`;
      return;
    }

    const series = seriesSelect.value; // "CFT-3" or "CFT-6"
    let specVal = parseFloat(specForSaveInput.value);
    if (!Number.isFinite(specVal) || specVal <= 0) {
      // 若沒輸入，就用內徑
      specVal = parseFloat(innerDiameterInput.value);
    }

    if (!Number.isFinite(specVal) || specVal <= 0) {
      saveStatus.innerHTML = `<span class="error">請先輸入有效的規格（內徑 mm）。</span>`;
      return;
    }

    const key = specVal.toFixed(3); // key 用 mm，保留三位小數

    if (series === "CFT-3") {
      pvcCost.cft3[key] = latestCostPerM;
    } else {
      pvcCost.cft6[key] = latestCostPerM;
    }

    savePvcToStorage();
    renderPvcTables();

    saveStatus.innerHTML =
      `<span class="ok">已儲存 ${series} 規格 ${key} mm，每米成本 ${latestCostPerM.toFixed(4)} 元。</span>`;
  });

  // ---- 🚀 一鍵存入 CFT-3／CFT-6 全部規格成本 ----
  const bulkFillBtn = document.getElementById("bulkFillBtn");
  const bulkFillStatus = document.getElementById("bulkFillStatus");

  if (bulkFillBtn) {
    bulkFillBtn.addEventListener("click", () => {
      const pelletPrice = parseFloat(pelletPriceInput.value);
      if (!Number.isFinite(pelletPrice) || pelletPrice <= 0) {
        bulkFillStatus.innerHTML = `<span class="error">請先在上面輸入有效的『PVC 粒 (元/Kg)』，再按一鍵存入。</span>`;
        return;
      }

      let count3 = 0;
      let count6 = 0;

      Object.entries(CFT_SPEC_TABLE).forEach(([series, rows]) => {
        rows.forEach(([, innerMm, wallMm]) => {
          const outerMm = innerMm + wallMm * 2;
          const weightPerM = (outerMm - wallMm) * wallMm * 3.14 * 1.3; // g/M，1.3已含耗損係數
          const costPerM = (weightPerM / 1000) * pelletPrice; // 元/M

          const key = innerMm.toFixed(3);
          if (series === "CFT-3") {
            pvcCost.cft3[key] = costPerM;
            count3++;
          } else {
            pvcCost.cft6[key] = costPerM;
            count6++;
          }
        });
      });

      savePvcToStorage();
      renderPvcTables();

      bulkFillStatus.innerHTML =
        `<span class="ok">一鍵存入完成：CFT-3 共 ${count3} 個規格、CFT-6 共 ${count6} 個規格（單價 ${pelletPrice} 元/Kg）。</span>`;
    });
  }

  // ---- 📋 PVC 每米成本速查表 ----
  let specTableActiveSeries = "CFT-3";
  const specTableBody = document.getElementById("specTableBody");
  const tabCFT3 = document.getElementById("tabCFT3");
  const tabCFT6 = document.getElementById("tabCFT6");

  function specCalcRow(innerD, wall, pelletPrice) {
    const outerD = innerD + wall * 2;
    const weightPerM = (outerD - wall) * wall * 3.14 * 1.3;
    const costPerM = (weightPerM / 1000) * pelletPrice;
    const cost305 = costPerM * 305;
    return { outerD, weightPerM, costPerM, cost305 };
  }

  function renderSpecTable() {
    if (!specTableBody) return;
    const pelletPrice = parseFloat(pelletPriceInput.value);
    const validPrice = Number.isFinite(pelletPrice) && pelletPrice > 0 ? pelletPrice : 0;
    const rows = CFT_SPEC_TABLE[specTableActiveSeries] || [];

    specTableBody.innerHTML = "";
    rows.forEach(([label, innerD, wall]) => {
      const { outerD, weightPerM, costPerM, cost305 } = specCalcRow(innerD, wall, validPrice);
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${label}</td>
        <td>${innerD}</td>
        <td>${wall}</td>
        <td>${outerD.toFixed(3)}</td>
        <td>${weightPerM.toFixed(3)}</td>
        <td>${costPerM.toFixed(3)}</td>
        <td>${cost305.toFixed(1)}</td>
      `;
      specTableBody.appendChild(tr);
    });
  }

  if (tabCFT3) tabCFT3.addEventListener("click", () => { specTableActiveSeries = "CFT-3"; renderSpecTable(); });
  if (tabCFT6) tabCFT6.addEventListener("click", () => { specTableActiveSeries = "CFT-6"; renderSpecTable(); });
  if (pelletPriceInput) {
    pelletPriceInput.addEventListener("input", renderSpecTable);
    pelletPriceInput.addEventListener("change", renderSpecTable);
  }
  renderSpecTable();
});
