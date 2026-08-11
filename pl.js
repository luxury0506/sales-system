// pl.js
// 損益總覽（不分月份，全部日期攤平在同一張表）。
// 讀寫 localStorage key "monthlyPL"，跟主頁 index.html 記錄「這批明細」用的是同一把 key，
// 資料結構：{ "2026-08": { "2026-08-10": {amount,cost,profit,rowCount,recordedAt}, ... }, ... }
// 這裡把所有月份底下的日期攤平成一個陣列一起顯示。

const STORAGE_KEY = "monthlyPL";

function loadAll() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) {
    console.error("讀取損益資料失敗：", e);
    return {};
  }
}

function saveAll(data) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (e) {
    console.error("儲存損益資料失敗：", e);
  }
}

function formatMoney(v) {
  const num = Number(v);
  if (!Number.isFinite(num)) return "";
  return Math.round(num).toLocaleString("zh-TW");
}

// 把 { "2026-08": { "2026-08-10": {...} }, "2026-07": {...} } 攤平成
// [ { month: "2026-08", date: "2026-08-10", ...record }, ... ]，依日期由舊到新排序。
function flattenAll(all) {
  const rows = [];
  Object.keys(all).forEach((month) => {
    const monthData = all[month] || {};
    Object.keys(monthData).forEach((date) => {
      rows.push({ month, date, ...monthData[date] });
    });
  });
  rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  return rows;
}

document.addEventListener("DOMContentLoaded", () => {
  const banner = document.getElementById("plBanner");
  const tbody = document.getElementById("plTbody");

  function render() {
    if (!tbody || !banner) return;
    const all = loadAll();
    const rows = flattenAll(all);

    tbody.innerHTML = "";
    let totalAmount = 0;
    let totalCost = 0;
    let totalProfit = 0;

    rows.forEach((r) => {
      totalAmount += r.amount || 0;
      totalCost += r.cost || 0;
      totalProfit += r.profit || 0;

      const marginRate = r.amount > 0 ? (r.profit / r.amount) * 100 : 0;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="border px-2 py-1">${r.date}</td>
        <td class="border px-2 py-1 text-right">${formatMoney(r.amount)}</td>
        <td class="border px-2 py-1 text-right">${formatMoney(r.cost)}</td>
        <td class="border px-2 py-1 text-right ${r.profit < 0 ? "text-red-600" : ""}">${formatMoney(r.profit)}</td>
        <td class="border px-2 py-1 text-right">${marginRate.toFixed(1)}%</td>
        <td class="border px-2 py-1 text-center">
          <button type="button" data-month="${r.month}" data-date="${r.date}" class="plDeleteBtn text-red-500 hover:text-red-700 text-xs">刪除</button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    if (!rows.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="border px-2 py-3 text-center text-slate-400">還沒有任何記錄，請先到主頁上傳明細算完後按「✅ 記錄這批明細」。</td></tr>`;
    }

    tbody.querySelectorAll(".plDeleteBtn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const month = btn.getAttribute("data-month");
        const date = btn.getAttribute("data-date");
        const all2 = loadAll();
        if (all2[month]) {
          delete all2[month][date];
          if (!Object.keys(all2[month]).length) delete all2[month];
          saveAll(all2);
        }
        render();
      });
    });

    const marginRate = totalAmount > 0 ? (totalProfit / totalAmount) * 100 : 0;
    const isProfit = totalProfit >= 0;
    banner.className =
      "rounded-lg p-4 text-center font-semibold text-lg " +
      (rows.length === 0
        ? "bg-slate-50 text-slate-400"
        : isProfit
        ? "bg-emerald-50 text-emerald-700"
        : "bg-red-50 text-red-700");

    if (!rows.length) {
      banner.textContent = "目前還沒有任何記錄";
    } else {
      banner.textContent =
        `累計至今：` +
        (isProfit ? "🟢 賺 " : "🔴 賠 ") +
        formatMoney(Math.abs(totalProfit)) +
        ` 元　｜　銷貨金額 ${formatMoney(totalAmount)}　銷貨成本 ${formatMoney(totalCost)}　毛利率 ${marginRate.toFixed(1)}%　（已記錄 ${rows.length} 天，${r0(rows)}）`;
    }
  }

  function r0(rows) {
    if (!rows.length) return "";
    return `${rows[0].date} ~ ${rows[rows.length - 1].date}`;
  }

  render();
});
