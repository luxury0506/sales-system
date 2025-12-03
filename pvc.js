// pvc.js
// PVC 管重量 + 成本利潤計算，並儲存每米成本到 localStorage

const PVC_STORAGE_KEY = "PVC_COST_TABLE";

let pvcCost = {
  cft3: {},      // 規格(mm) → 每米成本（台幣）
  cft6: {},      // 規格(mm) → 每米成本（台幣）
  updatedAt: null,
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
});
