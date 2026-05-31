const data = window.dashboardData;
const red = "#d4002a";
const darkRed = "#8a0016";
const gray = "#6b6b6b";
const palette = ["#d4002a", "#8a0016", "#444", "#777", "#b4001e", "#cfcfcf"];
const charts = [];

const fmt = new Intl.NumberFormat("ar-SA");
const $ = (selector) => document.querySelector(selector);

function total(rows) {
  return rows.reduce((sum, row) => sum + row[1], 0);
}

function percent(value, base) {
  return base ? `${((value / base) * 100).toFixed(1)}%` : "0%";
}

function setText(selector, value) {
  const element = $(selector);
  if (element) element.textContent = value;
}

function initHeader() {
  const org = data.organization;
  document.title = `${org.name} - ${org.dashboardTitle}`;
  setText("#brand-mark", org.shortName);
  setText("#hero-mark", org.shortName);
  setText("#nav-title", org.dashboardTitle);
  setText("#nav-subtitle", org.subtitle);
  setText("#date-pill", `🗓️ ${org.month}`);
  setText("#hero-org", org.name);
  setText("#hero-badge", `📊 ${org.reportBadge}`);
  $("#hero-title").innerHTML = org.heroTitle;

  const registrations = data.kpis[0].value;
  const members = data.kpis[2].value;
  setText("#hero-summary", `${fmt.format(registrations)} ${org.recordLabel || "تسجيل"} · ${fmt.format(members)} ${org.memberLabel || "عضو ميداني"} · ${fmt.format(data.forms.length)} نماذج · ${org.dateRange}`);
  setText("#footer", `${org.name} · ${org.dashboardTitle} · ${org.month}`);
  setText("#daily-title", org.dailyTitle || "التوزيع اليومي للتسجيلات");
  setText("#cases-title", org.casesTitle || "أنواع الحالات الطبية");
  setText("#gender-title", org.genderTitle || "توزيع الجنس");
  setText("#source-title", org.sourceTitle || "مصدر التسجيل");
  setText("#roles-title", org.rolesTitle || "صفة المصاب");
  setText("#roles-full-title", org.rolesTitle || "صفة المصاب");
  setText("#allergy-title", org.allergyTitle || "حساسية الأدوية");
  setText("#medicine-title", org.medicineTitle || "استخدام الأدوية");
  setText("#members-section-title", org.membersSectionTitle || "الأعضاء الميدانيون");

  $("#tabs").innerHTML = data.tabs
    .map((tab, index) => `<button class="tab ${index === 0 ? "active" : ""}" data-tab="${tab.id}">${tab.icon} ${tab.label}</button>`)
    .join("");

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab === button));
      document.querySelectorAll(".page").forEach((page) => page.classList.toggle("active", page.dataset.page === button.dataset.tab));
      charts.forEach((chart) => chart.resize());
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  });
}

function renderKpis() {
  $("#kpi-strip").innerHTML = data.kpis
    .map((kpi) => `
      <article class="kpi">
        <div class="kpi-value">${fmt.format(kpi.value)}</div>
        <div class="kpi-label">${kpi.label}</div>
        ${kpi.badge ? `<span class="kpi-badge ${kpi.good ? "good" : ""}">${kpi.badge}</span>` : ""}
      </article>
    `)
    .join("");
}

function renderBars(selector, rows) {
  const max = Math.max(...rows.map((row) => row[1]), 1);
  const base = total(rows);
  $(selector).innerHTML = rows
    .map(([name, value]) => `
      <div class="bar-item">
        <div class="bar-meta">
          <span class="bar-name">${name}</span>
          <span class="bar-side">
            <span class="bar-percent">${percent(value, base)}</span>
            <span class="bar-num">${fmt.format(value)}</span>
          </span>
        </div>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.max((value / max) * 100, 3)}%"></div></div>
      </div>
    `)
    .join("");
}

function renderMiniStats(selector, rows) {
  $(selector).innerHTML = rows
    .map(([label, value]) => `<div><strong>${fmt.format(value)}</strong><span>${label}</span></div>`)
    .join("");
}

function renderMembers() {
  setText("#members-count", `${fmt.format(data.kpis[2].value)} عضو`);
  $("#podium").innerHTML = data.members
    .slice(0, 3)
    .map(([name, value], index) => `
      <article class="podium-card">
        <div class="podium-rank">${index + 1}</div>
        <div class="podium-name">${name}</div>
        <div class="podium-value">${fmt.format(value)} تسجيل</div>
      </article>
    `)
    .join("");

  $("#member-list").innerHTML = data.members
    .map(([name, value], index) => `
      <div class="member-row">
        <span class="member-name">${index + 1}. ${name}</span>
        <span class="member-score">${fmt.format(value)}</span>
      </div>
    `)
    .join("");
}

function renderForms() {
  setText("#forms-count", `${fmt.format(data.forms.length)} نماذج`);
  $("#form-cards").innerHTML = data.forms
    .map((form) => `
      <article class="form-card">
        <div class="form-icon">${form.icon}</div>
        <h3>${form.name}</h3>
        <p>${form.description}</p>
        <div class="row"><span>إجمالي التسجيلات</span><strong>${fmt.format(form.total)}</strong></div>
        <div class="row"><span>مكتمل</span><strong>${fmt.format(form.completed)}</strong></div>
      </article>
    `)
    .join("");
}

function commonOptions({ indexAxis = "x", legend = false } = {}) {
  return {
    indexAxis,
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: legend, position: "bottom", labels: { font: { family: "Cairo" }, usePointStyle: true } },
      tooltip: {
        rtl: true,
        titleFont: { family: "Cairo" },
        bodyFont: { family: "Cairo" },
        callbacks: { label: (ctx) => `${ctx.label || ctx.dataset.label}: ${fmt.format(ctx.raw)}` }
      }
    },
    scales: indexAxis === "y" ? {
      x: { grid: { color: "#eee" }, ticks: { font: { family: "Cairo" } } },
      y: { grid: { display: false }, ticks: { font: { family: "Cairo" } } }
    } : {
      x: { grid: { display: false }, ticks: { font: { family: "Cairo" } } },
      y: { beginAtZero: true, grid: { color: "#eee" }, ticks: { font: { family: "Cairo" } } }
    }
  };
}

function chart(id, config) {
  const ctx = document.getElementById(id);
  if (!ctx || !window.Chart) return;
  charts.push(new Chart(ctx, config));
}

function barChart(id, rows, options = {}) {
  chart(id, {
    type: "bar",
    data: {
      labels: rows.map((row) => row[0]),
      datasets: [{
        label: "عدد التسجيلات",
        data: rows.map((row) => row[1]),
        borderRadius: 8,
        backgroundColor: options.color || red,
        hoverBackgroundColor: darkRed
      }]
    },
    options: commonOptions(options)
  });
}

function doughnut(id, rows) {
  chart(id, {
    type: "doughnut",
    data: {
      labels: rows.map((row) => row[0]),
      datasets: [{ data: rows.map((row) => row[1]), backgroundColor: palette, borderColor: "#fff", borderWidth: 3 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "58%",
      plugins: {
        legend: { display: true, position: "bottom", labels: { font: { family: "Cairo" }, usePointStyle: true } },
        tooltip: { rtl: true, titleFont: { family: "Cairo" }, bodyFont: { family: "Cairo" } }
      }
    }
  });
}

function initCharts() {
  barChart("dailyChart", data.daily);
  doughnut("genderChart", data.gender);
  doughnut("sourceChart", data.sources);
  doughnut("rolesChart", data.roles);
  doughnut("allergyChart", data.allergy);
  doughnut("medChart", data.medicineUse);
  barChart("casesFullChart", data.cases, { indexAxis: "y" });
  doughnut("rolesFullChart", data.roles);
  barChart("locChart", data.locations, { indexAxis: "y", color: gray });
  barChart("membersChart", data.members, { indexAxis: "y" });
  barChart("formsChart", data.forms.map((form) => [form.name, form.total]));
}

function init() {
  initHeader();
  renderKpis();
  setText("#daily-count", data.organization.dailyCount || `${fmt.format(data.daily.length)} يوم`);
  setText("#cases-count", `${fmt.format(total(data.cases))} حالة`);
  renderBars("#case-bars", data.cases.slice(0, 6));
  renderBars("#action-bars", data.actions);
  renderMiniStats("#allergy-stat", data.allergy);
  renderMiniStats("#medicine-stat", data.medicineUse);
  renderBars("#cases-full-bars", data.cases);
  renderBars("#location-bars", data.locations);
  renderMembers();
  renderForms();
  initCharts();
}

init();
