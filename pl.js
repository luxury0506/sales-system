// pl.js
// 本月損益總覽（獨立頁面）。
// 讀寫 localStorage key "monthlyPL"，跟主頁 index.html 記錄「這批明細」用的是同一把 key，
// 資料結構：{ "2026-08": { "2026-08-10": {amount,cost,profit,rowCount,recordedAt}, ... }, ... }

const STORAGE_KEY = "monthlyPL";

function todayStr() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

function currentMonthStr() {
  return todayStr().slice(0, 7);
}

function loadAll() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.error("讀取本月損益資料失敗：", e);
    return {};
  }
}

function saveAll(data) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (e) {
    console.error("儲存本月損益資料失敗：", e);
  }
}

function formatMoney(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return Math.round(num).toLocaleString("zh-TW");
}

document.addEventListener("DOMContentLoaded", () => {
  const monthSelect = document.getElementById("plMonthSelect");
  const banner = document.getElementById("plMonthBanner");
  const tbody = document.getElementById("plMonthTbody");

  if (monthSelect) monthSelect.value = currentMonthStr();

  function renderMonth() {
    if (!monthSelect || !tbody || !banner) return;
    const month = monthSelect.value || currentMonthStr();
    const all = loadAll();
    const monthData = all[month] || {};
    const dates = Object.keys(monthData).sort();

    tbody.innerHTML = "";
    let totalAmount = 0;
    let totalCost = 0;
    let totalProfit = 0;

    dates.forEach((date) => {
      const d = monthData[date];
      totalAmount += d.amount || 0;
      totalCost += d.cost || 0;
      totalProfit += d.profit || 0;

      const marginRate = d.amount > 0 ? (d.profit / d.amount) * 100 : 0;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="border px-2 py-1">${date}</td>
        <td class="border px-2 py-1 text-right">${formatMoney(d.amount)}</td>
        <td class="border px-2 py-1 text-right">${formatMoney(d.cost)}</td>
        <td class="border px-2 py-1 text-right ${d.profit < 0 ? "text-red-600" : ""}">${formatMoney(d.profit)}</td>
        <td class="border px-2 py-1 text-right">${marginRate.toFixed(1)}%</td>
        <td class="border px-2 py-1 text-center">
          <button type="button" data-date="${date}" class="plDeleteBtn text-red-500 hover:text-red-700 text-xs">刪除</button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    if (!dates.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="border px-2 py-3 text-center text-slate-400">這個月還沒有記錄任何一天，請先到主頁上傳明細算完後按「✅ 記錄這批明細」。</td></tr>`;
    }

    tbody.querySelectorAll(".plDeleteBtn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const date = btn.getAttribute("data-date");
        const all2 = loadAll();
        if (all2[month]) {
          delete all2[month][date];
          if (!Object.keys(all2[month]).length) delete all2[month];
          saveAll(all2);
        }
        renderMonth();
      });
    });

    const marginRate = totalAmount > 0 ? (totalProfit / totalAmount) * 100 : 0;
    const isProfit = totalProfit >= 0;
    banner.className =
      "rounded-lg p-4 text-center font-semibold text-lg " +
      (dates.length === 0
        ? "bg-slate-50 text-slate-400"
        : isProfit
        ? "bg-emerald-50 text-emerald-700"
        : "bg-red-50 text-red-700");

    if (!dates.length) {
      banner.textContent = `${month} 目前還沒有任何記錄`;
    } else {
      banner.textContent =
        `${month} 本月至今：` +
        (isProfit ? "🟢 賺 " : "🔴 賠 ") +
        formatMoney(Math.abs(totalProfit)) +
        ` 元　｜　銷貨金額 ${formatMoney(totalAmount)}　銷貨成本 ${formatMoney(totalCost)}　毛利率 ${marginRate.toFixed(1)}%　（已記錄 ${dates.length} 天）`;
    }
  }

  if (monthSelect) monthSelect.addEventListener("change", renderMonth);

  renderMonth();
});
