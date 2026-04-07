(async function () {
  const res = await fetch('/api/me');
  if (!res.ok) {
    window.location.href = '/login.html';
  }
})();

const $ = (id) => document.getElementById(id);

// Forms
const projectForm = $("projectForm");
const transactionForm = $("transactionForm");

// Project form
const editId = $("editId");
const keepLogoPath = $("keepLogoPath");
const projectNo = $("projectNo");
const projectCode = $("projectCode");
const projectName = $("projectName");
const category = $("category");
const owner = $("owner");
const startDate = $("startDate");
const endDate = $("endDate");
const remark = $("remark");
const companyLogo = $("companyLogo");
const logoPreview = $("logoPreview");

const contractCurrency = $("contractCurrency");
const totalPrice = $("totalPrice");
const totalPriceRaw = $("totalPriceRaw");
const vatPercent = $("vatPercent");
const totalPriceWithVatDisplay = $("totalPriceWithVatDisplay");

// Transaction form
const selectedProjectId = $("selectedProjectId");
const selectedProjectText = $("selectedProjectText");

const txType = $("txType");
const txCategory = $("txCategory");
const txDescription = $("txDescription");
const txCurrency = $("txCurrency");
const txAmount = $("txAmount");
const txAmountRaw = $("txAmountRaw");
const txDate = $("txDate");
const billFile = $("billFile");

// Buttons
const saveBtn = $("saveBtn");
const resetBtn = $("resetBtn");
const addTxBtn = $("addTxBtn");
const clearTxBtn = $("clearTxBtn");
const refreshBtn = $("refreshBtn");
const downloadExcelBtn = $("downloadExcelBtn");
const logoutBtn = $("logoutBtn");

// List / filters
const projectList = $("projectList");
const searchInput = $("searchInput");
const filterCategory = $("filterCategory");
const filterOwner = $("filterOwner");
const sortBy = $("sortBy");

// KPI
const kpiProjects = $("kpiProjects");
const kpiIncome = $("kpiIncome");
const kpiInvestment = $("kpiInvestment");
const kpiExpense = $("kpiExpense");

// Datalists
const projectCategoryList = $("projectCategoryList");
const txCategoryList = $("txCategoryList");
const filterCategoryList = $("filterCategoryList");
const filterOwnerList = $("filterOwnerList");

// Project modal
const projectModal = $("projectModal");
const projectModalBackdrop = $("projectModalBackdrop");
const closeProjectModalBtn = $("closeProjectModal");
const editProjectBtn = $("editProjectBtn");
const openAddTxFromDetailBtn = $("openAddTxFromDetailBtn");
const deleteProjectBtn = $("deleteProjectBtn");
const projectModalHistory = $("projectModalHistory");
const detailHistoryCount = $("detailHistoryCount");

const detailProjectCode = $("detailProjectCode");
const detailProjectCategory = $("detailProjectCategory");
const detailProjectOwner = $("detailProjectOwner");
const detailProjectName = $("detailProjectName");
const detailProjectRemark = $("detailProjectRemark");
const detailStartDate = $("detailStartDate");
const detailEndDate = $("detailEndDate");
const detailCurrency = $("detailCurrency");
const detailTransactionCount = $("detailTransactionCount");
const projectDetailLogo = $("projectDetailLogo");

const detailTotalPrice = $("detailTotalPrice");
const detailVatPercent = $("detailVatPercent");
const detailTotalWithVat = $("detailTotalWithVat");
const detailActualCost = $("detailActualCost");
const detailEstimatedProfit = $("detailEstimatedProfit");
const detailBalance = $("detailBalance");
const detailIncome = $("detailIncome");
const detailInvestment = $("detailInvestment");
const detailExpense = $("detailExpense");

// Tx modal
const txModal = $("txModal");
const txModalBackdrop = $("txModalBackdrop");
const closeTxModalBtn = $("closeTxModal");
const txModalHistory = $("txModalHistory");
const txHistoryCount = $("txHistoryCount");

let allProjects = [];
let currentProjectId = "";

const DEFAULT_PROJECT_CATEGORIES = ["Administrative expenses"];
const DEFAULT_TX_CATEGORIES = ["Administrative expenses"];

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * Strict number parser.
 * Keeps Excel data untouched.
 * Frontend only converts display/calculation values safely.
 */
function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const raw = String(value).trim();
  if (!raw) return 0;

  // remove commas first
  const cleaned = raw.replace(/,/g, "");

  // direct numeric string
  if (/^-?\d+(\.\d+)?$/.test(cleaned)) {
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }

  // fallback: pick first numeric fragment only
  // example bad value: "0 2026-03-17T09:48:36.429Z" -> 0
  const match = cleaned.match(/-?\d+(\.\d+)?/);
  if (!match) return 0;

  const n = Number(match[0]);
  return Number.isFinite(n) ? n : 0;
}

function getCurrencySymbol(currency) {
  const map = {
    LAK: "₭",
    THB: "฿",
    USD: "$",
    CNY: "¥",
  };
  return map[String(currency || "").toUpperCase()] || "";
}

function formatDisplayNumber(value) {
  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: 2,
  }).format(toNumber(value));
}

function formatMoney(value, currency = "") {
  const code = String(currency || "").toUpperCase();
  const symbol = getCurrencySymbol(code);
  const text = formatDisplayNumber(value);
  if (symbol) return `${symbol}${text}`;
  return `${text}${code ? ` ${code}` : ""}`;
}

function formatDateRange(start, end) {
  const startText = start || "-";
  const endText = end || "-";
  return `${escapeHtml(startText)} - ${escapeHtml(endText)}`;
}

function setButtonLoading(button, isLoading, text, loadingText) {
  if (!button) return;
  button.disabled = isLoading;
  button.textContent = isLoading ? loadingText : text;
}

async function fetchJSON(url, options = {}) {
  const response = await fetch(url, options);
  let data = null;

  try {
    data = await response.json();
  } catch {
    data = null;
  }

  if (!response.ok) {
    throw new Error(data?.error || data?.message || "Request failed");
  }

  return data;
}

function buildDatalist(listEl, values, pinnedValues = []) {
  const uniqueValues = [...new Set(values.filter(Boolean).map((v) => String(v).trim()).filter(Boolean))];
  const pinned = pinnedValues
    .filter(Boolean)
    .map((v) => String(v).trim())
    .filter((v) => uniqueValues.includes(v));

  const normal = uniqueValues
    .filter((value) => !pinned.includes(value))
    .sort((a, b) => a.localeCompare(b));

  const finalValues = [...pinned, ...normal];

  listEl.innerHTML = finalValues
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
}

function refreshDatalists(projects) {
  const projectCategories = [
    ...DEFAULT_PROJECT_CATEGORIES,
    ...projects.map((item) => item.category).filter(Boolean),
  ];

  const owners = [...new Set(projects.map((item) => item.owner).filter(Boolean))];

  const txCategories = [
    ...DEFAULT_TX_CATEGORIES,
    ...projects
      .flatMap((item) => (item.transactions || []).map((tx) => tx.category))
      .filter(Boolean),
  ];

  buildDatalist(projectCategoryList, projectCategories, DEFAULT_PROJECT_CATEGORIES);
  buildDatalist(filterCategoryList, projectCategories, DEFAULT_PROJECT_CATEGORIES);
  buildDatalist(filterOwnerList, owners);
  buildDatalist(txCategoryList, txCategories, DEFAULT_TX_CATEGORIES);
}

function calcTotalWithVat(total, vat) {
  const t = toNumber(total);
  const v = toNumber(vat);
  return t + (t * v) / 100;
}

/**
 * Recalculate all project summary values safely on frontend.
 * We do NOT trust broken precomputed fields if backend returns mixed text.
 */
function buildProjectSummary(project) {
  const currency = project?.contractCurrency || "LAK";
  const safeVatPercent = toNumber(project?.vatPercent);
  const safeTotalPrice = toNumber(project?.totalPrice);

  const transactions = Array.isArray(project?.transactions) ? project.transactions : [];

  let income = 0;
  let investment = 0;
  let expense = 0;

  for (const tx of transactions) {
    const amount = toNumber(tx?.amount);
    if (tx?.type === "income") income += amount;
    else if (tx?.type === "investment") investment += amount;
    else if (tx?.type === "expense") expense += amount;
  }

  const actualCost = investment + expense;
  const totalPriceWithVat = calcTotalWithVat(safeTotalPrice, safeVatPercent);
  const estimatedProfit = income - totalPriceWithVat;
  const balance = income - actualCost;

  return {
    currency,
    totalPrice: safeTotalPrice,
    vatPercent: safeVatPercent,
    totalPriceWithVat,
    actualCost,
    estimatedProfit,
    balance,
    totals: {
      income,
      investment,
      expense,
    },
    transactionCount: transactions.length,
  };
}

function normalizeTransaction(tx) {
  return {
    ...tx,
    amount: toNumber(tx?.amount),
  };
}

function normalizeProject(project) {
  const normalizedTransactions = Array.isArray(project?.transactions)
    ? project.transactions.map(normalizeTransaction)
    : [];

  const summary = buildProjectSummary({
    ...project,
    transactions: normalizedTransactions,
  });

  return {
    ...project,
    contractCurrency: project?.contractCurrency || "LAK",
    totalPrice: summary.totalPrice,
    vatPercent: summary.vatPercent,
    totalPriceWithVat: summary.totalPriceWithVat,
    actualCost: summary.actualCost,
    estimatedProfit: summary.estimatedProfit,
    balance: summary.balance,
    transactionCount: summary.transactionCount,
    totals: summary.totals,
    transactions: normalizedTransactions,
  };
}

function updateKPIs(projects) {
  const total = projects.reduce(
    (sum, project) => {
      sum.income += toNumber(project.totals?.income);
      sum.investment += toNumber(project.totals?.investment);
      sum.expense += toNumber(project.totals?.expense);
      return sum;
    },
    { income: 0, investment: 0, expense: 0 }
  );

  kpiProjects.textContent = String(projects.length);
  kpiIncome.textContent = formatDisplayNumber(total.income);
  kpiInvestment.textContent = formatDisplayNumber(total.investment);
  kpiExpense.textContent = formatDisplayNumber(total.expense);
}

async function loadNextProjectCode() {
  const data = await fetchJSON("/api/next-project-code");
  projectNo.value = data.no || "";
  projectCode.value = data.projectCode || "";
}

function clearLogoPreview() {
  logoPreview.className = "logo-preview empty";
  logoPreview.innerHTML = "No logo selected";
}

function showLogoPreview(src) {
  if (!src) {
    clearLogoPreview();
    return;
  }

  logoPreview.className = "logo-preview";
  logoPreview.innerHTML = `<img src="${src}" alt="Logo Preview">`;
}

function showDetailLogo(src) {
  if (!projectDetailLogo) return;

  if (!src) {
    projectDetailLogo.className = "logo-preview empty small";
    projectDetailLogo.innerHTML = "No logo";
    return;
  }

  projectDetailLogo.className = "logo-preview small";
  projectDetailLogo.innerHTML = `<img src="${src}" alt="Project Logo">`;
}

function parseInputNumber(value) {
  return String(value || "")
    .replace(/,/g, "")
    .replace(/[^\d.]/g, "");
}

function formatInputNumber(value) {
  if (!value) return "";
  const num = Number(value);
  if (Number.isNaN(num)) return "";
  return num.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

function updateProjectPricePreview() {
  const raw = parseInputNumber(totalPrice.value);
  totalPriceRaw.value = raw;

  if (document.activeElement === totalPrice) {
    totalPrice.value = raw ? formatInputNumber(raw) : "";
  }

  const totalWithVat = calcTotalWithVat(raw, vatPercent.value);
  const currency = contractCurrency?.value || "";
  totalPriceWithVatDisplay.value = totalWithVat
    ? formatMoney(totalWithVat, currency)
    : "";
}

function resetProjectForm() {
  editId.value = "";
  keepLogoPath.value = "";
  projectName.value = "";
  category.value = "";
  owner.value = "";
  startDate.value = "";
  endDate.value = "";
  remark.value = "";
  companyLogo.value = "";
  contractCurrency.value = "LAK";
  totalPrice.value = "";
  totalPriceRaw.value = "";
  vatPercent.value = "10";
  totalPriceWithVatDisplay.value = "";
  saveBtn.textContent = "Save Project";
  clearLogoPreview();
  loadNextProjectCode().catch(console.error);
}

function getDefaultTxCategoryByType(type) {
  if (type === "expense") return "Administrative expenses";
  return "";
}

function resetTransactionForm(keepProject = true) {
  txType.value = "income";
  txCategory.value = getDefaultTxCategoryByType(txType.value);
  txDescription.value = "";
  txCurrency.value = "LAK";
  txAmount.value = "";
  txAmountRaw.value = "";
  txDate.value = todayISO();
  billFile.value = "";

  if (!keepProject) {
    selectedProjectId.value = "";
    selectedProjectText.value = "";
    renderTxModalHistory(null);
  }
}

function populateProjectForm(project) {
  editId.value = project.id || "";
  keepLogoPath.value = project.logoPath || "";
  projectNo.value = project.no || "";
  projectCode.value = project.projectCode || "";
  projectName.value = project.projectName || "";
  category.value = project.category || "";
  owner.value = project.owner || "";
  startDate.value = project.startDate || "";
  endDate.value = project.endDate || "";
  remark.value = project.remark || "";
  contractCurrency.value = project.contractCurrency || "LAK";
  totalPriceRaw.value = String(toNumber(project.totalPrice || 0) || "");
  totalPrice.value = totalPriceRaw.value ? formatInputNumber(totalPriceRaw.value) : "";
  vatPercent.value = String(toNumber(project.vatPercent || 0));
  saveBtn.textContent = "Update Project";
  showLogoPreview(project.logoPath || "");
  updateProjectPricePreview();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function getFilteredProjects() {
  const keyword = searchInput.value.trim().toLowerCase();
  const categoryValue = filterCategory.value.trim().toLowerCase();
  const ownerValue = filterOwner.value.trim().toLowerCase();
  const sortValue = sortBy.value;

  let items = allProjects.filter((item) => {
    const haystack = [
      item.projectCode,
      item.projectName,
      item.category,
      item.owner,
      item.contractCurrency
    ]
      .join(" ")
      .toLowerCase();

    const matchesKeyword = !keyword || haystack.includes(keyword);
    const matchesCategory =
      !categoryValue || String(item.category || "").toLowerCase() === categoryValue;
    const matchesOwner =
      !ownerValue || String(item.owner || "").toLowerCase() === ownerValue;

    return matchesKeyword && matchesCategory && matchesOwner;
  });

  items = [...items];

  if (sortValue === "oldest") {
    items.sort((a, b) => toNumber(a.no) - toNumber(b.no));
  } else if (sortValue === "name") {
    items.sort((a, b) =>
      String(a.projectName || "").localeCompare(String(b.projectName || ""))
    );
  } else if (sortValue === "code") {
    items.sort((a, b) =>
      String(a.projectCode || "").localeCompare(String(b.projectCode || ""))
    );
  } else if (sortValue === "profitHigh") {
    items.sort((a, b) => toNumber(b.estimatedProfit) - toNumber(a.estimatedProfit));
  } else if (sortValue === "profitLow") {
    items.sort((a, b) => toNumber(a.estimatedProfit) - toNumber(b.estimatedProfit));
  } else {
    items.sort((a, b) => toNumber(b.no) - toNumber(a.no));
  }

  return items;
}

function getLogoCardHtml(item) {
  if (item.logoPath) {
    return `<img class="project-logo" src="${item.logoPath}" alt="${escapeHtml(
      item.projectName || "Project Logo"
    )}">`;
  }

  return `<div class="project-logo" style="display:grid;place-items:center;font-size:11px;font-weight:800;color:#94a3b8;">NO LOGO</div>`;
}

function getProjectRemarkText(item) {
  const text = String(item.remark || "").trim();
  return text || "No remark";
}

function renderProjectCard(item) {
  const isActive = item.id === currentProjectId;
  const currency = item.contractCurrency || "LAK";
  const profitClass = toNumber(item.estimatedProfit) >= 0 ? "profit-positive" : "profit-negative";

  return `
    <article class="project-card ${isActive ? "active" : ""}" data-project-id="${item.id}">
      <div class="project-head">
        ${getLogoCardHtml(item)}
        <div class="project-title-wrap">
          <div class="project-code">${escapeHtml(item.projectCode || "")}</div>
          <h3 class="project-name">${escapeHtml(item.projectName || "")}</h3>
        </div>
      </div>

      <div class="project-meta">
        <div class="meta-box">
          <div class="meta-label">Category</div>
          <div class="meta-value">${escapeHtml(item.category || "-")}</div>
        </div>
        <div class="meta-box">
          <div class="meta-label">Owner</div>
          <div class="meta-value">${escapeHtml(item.owner || "-")}</div>
        </div>
      </div>

      <div class="amount-row">
        <div class="amount-box">
          <div class="label">Total Price</div>
          <div class="amt">${formatMoney(item.totalPrice, currency)}</div>
        </div>

        <div class="amount-box">
          <div class="label">Total + VAT</div>
          <div class="amt">${formatMoney(item.totalPriceWithVat, currency)}</div>
        </div>

        <div class="amount-box ${profitClass}">
          <div class="label">Profit</div>
          <div class="amt">${formatMoney(item.estimatedProfit, currency)}</div>
        </div>
      </div>

      <div class="amount-row">
        <div class="amount-box income">
          <div class="label">Income</div>
          <div class="amt">${formatDisplayNumber(item.totals?.income)}</div>
        </div>

        <div class="amount-box investment">
          <div class="label">Investment</div>
          <div class="amt">${formatDisplayNumber(item.totals?.investment)}</div>
        </div>

        <div class="amount-box expense">
          <div class="label">Expense</div>
          <div class="amt">${formatDisplayNumber(item.totals?.expense)}</div>
        </div>
      </div>

      <div class="project-foot">
        <div class="project-stat">Transactions <strong>${toNumber(item.transactionCount)}</strong></div>
        <div class="project-period">${formatDateRange(item.startDate, item.endDate)}</div>
      </div>

      <div class="project-remark">
        <div class="project-remark-label">Remark</div>
        <p>${escapeHtml(getProjectRemarkText(item))}</p>
      </div>

      <div class="project-actions">
        <button class="btn btn-light btn-small" data-action="open" data-id="${item.id}" type="button">Open</button>
        <button class="btn btn-light btn-small" data-action="add-tx" data-id="${item.id}" type="button">Add Transaction</button>
        <button class="btn btn-light btn-small" data-action="edit" data-id="${item.id}" type="button">Edit</button>
        <button class="btn btn-danger btn-small" data-action="delete" data-id="${item.id}" type="button">Delete</button>
      </div>
    </article>
  `;
}

function renderProjectList() {
  const items = getFilteredProjects();

  if (!items.length) {
    projectList.innerHTML = `
      <div class="empty-state" style="grid-column:1 / -1;">
        <h3>No projects found</h3>
        <p>Try another search or create a new project.</p>
      </div>
    `;
    return;
  }

  projectList.innerHTML = items.map(renderProjectCard).join("");
}

function getTypeBadgeClass(type) {
  if (type === "income") return "badge badge-income";
  if (type === "investment") return "badge badge-investment";
  if (type === "expense") return "badge badge-expense";
  return "badge";
}

function renderHistoryList(container, countEl, project = null) {
  if (!container || !countEl) return;

  const current = project?.id
    ? (allProjects.find((item) => item.id === project.id) || project)
    : (allProjects.find((item) => item.id === currentProjectId) || null);

  if (!current) {
    countEl.textContent = "0 item(s)";
    container.innerHTML = `
      <div class="empty-state small">
        <h3>No transactions yet</h3>
        <p>Add the first transaction for this project.</p>
      </div>
    `;
    return;
  }

  const transactions = Array.isArray(current.transactions) ? current.transactions : [];
  countEl.textContent = `${transactions.length} item(s)`;

  if (!transactions.length) {
    container.innerHTML = `
      <div class="empty-state small">
        <h3>No transactions yet</h3>
        <p>Add the first transaction for this project.</p>
      </div>
    `;
    return;
  }

  container.innerHTML = transactions.map((tx) => {
    const fileLink = tx.billPath
      ? `<a href="${tx.billPath}" target="_blank" rel="noopener">Open Bill</a>`
      : `<span style="color:#94a3b8;">No file</span>`;

    return `
      <div class="history-item">
        <div class="history-top">
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
            <span class="${getTypeBadgeClass(tx.type)}">${escapeHtml((tx.type || "").toUpperCase())}</span>
            <strong style="font-size:16px;">${formatMoney(tx.amount, tx.currency)}</strong>
          </div>
          <button class="btn btn-danger btn-small" data-action="delete-tx" data-id="${tx.id}" type="button">Delete</button>
        </div>

        <div class="history-meta">
          <div><strong>No:</strong> ${escapeHtml(tx.no || "-")}</div>
          <div><strong>Category:</strong> ${escapeHtml(tx.category || "-")}</div>
          <div><strong>Date:</strong> ${escapeHtml(tx.date || "-")}</div>
        </div>

        <div class="project-remark" style="margin-bottom:10px;">
          <div class="project-remark-label">Description</div>
          <p>${escapeHtml(tx.description || "No description")}</p>
        </div>

        <div class="history-files">${fileLink}</div>
      </div>
    `;
  }).join("");
}

function renderTxModalHistory(project = null) {
  renderHistoryList(txModalHistory, txHistoryCount, project);
}

function renderProjectModalHistory(project = null) {
  renderHistoryList(projectModalHistory, detailHistoryCount, project);
}

function fillProjectModal(project = null) {
  const currentProject = project || allProjects.find((item) => item.id === currentProjectId);

  if (!currentProject) return;

  const currency = currentProject.contractCurrency || "LAK";
  const profit = toNumber(currentProject.estimatedProfit);

  showDetailLogo(currentProject.logoPath || "");

  detailProjectCode.textContent = currentProject.projectCode || "-";
  detailProjectCategory.textContent = currentProject.category || "-";
  detailProjectOwner.textContent = currentProject.owner || "-";
  detailProjectName.textContent = currentProject.projectName || "-";
  detailProjectRemark.textContent = getProjectRemarkText(currentProject);
  detailStartDate.textContent = currentProject.startDate || "-";
  detailEndDate.textContent = currentProject.endDate || "-";
  detailCurrency.textContent = currency;
  detailTransactionCount.textContent = String(toNumber(currentProject.transactionCount));

  detailTotalPrice.textContent = formatMoney(currentProject.totalPrice, currency);
  detailVatPercent.textContent = `${formatDisplayNumber(currentProject.vatPercent)}%`;
  detailTotalWithVat.textContent = formatMoney(currentProject.totalPriceWithVat, currency);
  detailActualCost.textContent = formatMoney(currentProject.actualCost, currency);
  detailEstimatedProfit.textContent = formatMoney(profit, currency);
  detailBalance.textContent = formatMoney(currentProject.balance, currency);

  detailIncome.textContent = formatDisplayNumber(currentProject.totals?.income);
  detailInvestment.textContent = formatDisplayNumber(currentProject.totals?.investment);
  detailExpense.textContent = formatDisplayNumber(currentProject.totals?.expense);

  if (detailEstimatedProfit) {
    detailEstimatedProfit.classList.toggle("text-success", profit >= 0);
    detailEstimatedProfit.classList.toggle("text-danger", profit < 0);
  }

  renderProjectModalHistory(currentProject);
}

function openProjectModal(project) {
  if (!project || !projectModal) return;

  currentProjectId = project.id || "";
  fillProjectModal(project);
  renderProjectList();

  projectModal.classList.remove("hidden");
  projectModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeProjectModal() {
  if (!projectModal) return;
  projectModal.classList.add("hidden");
  projectModal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
}

function openTxModal(project) {
  if (!project || !txModal) return;

  currentProjectId = project.id || "";
  const fullProject = allProjects.find((item) => item.id === currentProjectId) || project;

  selectedProjectId.value = fullProject.id || "";
  selectedProjectText.value = `${fullProject.projectCode || "-"} - ${fullProject.projectName || "-"}`;
  txDate.value = todayISO();

  renderProjectList();
  renderTxModalHistory(fullProject);

  txModal.classList.remove("hidden");
  txModal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeTxModal() {
  if (!txModal) return;
  txModal.classList.add("hidden");
  txModal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
}

async function loadProjects(keepSelection = true) {
  const projects = await fetchJSON("/api/projects");
  allProjects = (Array.isArray(projects) ? projects : []).map(normalizeProject);

  refreshDatalists(allProjects);
  updateKPIs(allProjects);

  if (
    keepSelection &&
    currentProjectId &&
    !allProjects.find((item) => item.id === currentProjectId)
  ) {
    currentProjectId = "";
  }

  renderProjectList();

  if (currentProjectId) {
    const selected = allProjects.find((item) => item.id === currentProjectId);
    if (selected) {
      selectedProjectId.value = selected.id || "";
      selectedProjectText.value = `${selected.projectCode || "-"} - ${selected.projectName || "-"}`;
      fillProjectModal(selected);
      renderTxModalHistory(selected);
      return;
    }
  }

  currentProjectId = "";
  selectedProjectId.value = "";
  selectedProjectText.value = "";
  renderProjectModalHistory(null);
  renderTxModalHistory(null);
}

async function submitProjectForm(event) {
  event.preventDefault();

  const formData = new FormData();
  formData.append("projectName", projectName.value.trim());
  formData.append("category", category.value.trim());
  formData.append("owner", owner.value.trim());
  formData.append("startDate", startDate.value);
  formData.append("endDate", endDate.value);
  formData.append("remark", remark.value.trim());
  formData.append("keepLogoPath", keepLogoPath.value.trim());
  formData.append("contractCurrency", contractCurrency.value);
  formData.append("totalPrice", totalPriceRaw.value || "0");
  formData.append("vatPercent", vatPercent.value || "0");

  if (companyLogo.files?.[0]) {
    formData.append("companyLogo", companyLogo.files[0]);
  }

  const isEdit = Boolean(editId.value.trim());
  const url = isEdit ? `/api/projects/${editId.value.trim()}` : "/api/projects";
  const method = isEdit ? "PUT" : "POST";

  try {
    setButtonLoading(
      saveBtn,
      true,
      isEdit ? "Update Project" : "Save Project",
      isEdit ? "Updating..." : "Saving..."
    );

    const result = await fetchJSON(url, { method, body: formData });
    await loadProjects(false);
    resetProjectForm();

    if (result?.project) {
      const safeProject = normalizeProject(result.project);
      currentProjectId = safeProject.id || "";
      openProjectModal(safeProject);
    }

    alert(isEdit ? "Project updated successfully" : "Project saved successfully");
  } catch (error) {
    alert(error.message);
  } finally {
    const nowEdit = Boolean(editId.value.trim());
    setButtonLoading(
      saveBtn,
      false,
      nowEdit ? "Update Project" : "Save Project",
      ""
    );
  }
}

async function submitTransactionForm(event) {
  event.preventDefault();

  const projectId = selectedProjectId.value.trim();
  if (!projectId) {
    alert("Please select a project first");
    return;
  }

  const formData = new FormData();
  formData.append("type", txType.value);
  formData.append("category", txCategory.value.trim());
  formData.append("description", txDescription.value.trim());
  formData.append("currency", txCurrency.value);
  formData.append("amount", txAmountRaw.value || txAmount.value || "0");
  formData.append("date", txDate.value);

  if (billFile.files?.[0]) {
    formData.append("billFile", billFile.files[0]);
  }

  try {
    setButtonLoading(addTxBtn, true, "Add Transaction", "Saving...");

    await fetchJSON(`/api/projects/${projectId}/transactions`, {
      method: "POST",
      body: formData,
    });

    await loadProjects(false);

    const selected = allProjects.find((item) => item.id === projectId);
    if (selected) {
      currentProjectId = selected.id || "";
      fillProjectModal(selected);
      renderTxModalHistory(selected);
    }

    resetTransactionForm(true);
    closeTxModal();
    alert("Transaction saved successfully");
  } catch (error) {
    alert(error.message);
  } finally {
    setButtonLoading(addTxBtn, false, "Add Transaction", "");
  }
}

async function deleteProjectById(id) {
  const project = allProjects.find((item) => item.id === id);
  if (!project) return;

  if (!confirm(`Delete project "${project.projectName}" and all related transactions?`)) {
    return;
  }

  try {
    await fetchJSON(`/api/projects/${id}`, { method: "DELETE" });

    if (currentProjectId === id) {
      currentProjectId = "";
      selectedProjectId.value = "";
      selectedProjectText.value = "";
      closeProjectModal();
      closeTxModal();
    }

    await loadProjects(true);
    resetProjectForm();
    alert("Project deleted");
  } catch (error) {
    alert(error.message);
  }
}

async function handleProjectListClick(event) {
  const actionButton = event.target.closest("button[data-action]");
  if (actionButton) {
    const action = actionButton.dataset.action;
    const id = actionButton.dataset.id;
    const project = allProjects.find((item) => item.id === id);

    if (!project) return;

    if (action === "open") {
      openProjectModal(project);
      return;
    }

    if (action === "add-tx") {
      currentProjectId = project.id || "";
      openTxModal(project);
      return;
    }

    if (action === "edit") {
      currentProjectId = project.id || "";
      populateProjectForm(project);
      closeProjectModal();
      return;
    }

    if (action === "delete") {
      await deleteProjectById(id);
      return;
    }

    return;
  }

  const card = event.target.closest(".project-card");
  if (!card) return;

  const id = card.dataset.projectId;
  const project = allProjects.find((item) => item.id === id);
  if (!project) return;

  openProjectModal(project);
}

async function handleHistoryClick(event) {
  const button = event.target.closest('button[data-action="delete-tx"]');
  if (!button) return;

  const txId = button.dataset.id;
  if (!confirm("Delete this transaction?")) return;

  try {
    await fetchJSON(`/api/transactions/${txId}`, { method: "DELETE" });
    await loadProjects(false);

    const selected = allProjects.find((item) => item.id === currentProjectId);
    if (selected) {
      fillProjectModal(selected);
      renderTxModalHistory(selected);
    } else {
      renderProjectModalHistory(null);
      renderTxModalHistory(null);
      closeProjectModal();
    }

    alert("Transaction deleted");
  } catch (error) {
    alert(error.message);
  }
}

function handleLogoInputChange() {
  const file = companyLogo.files?.[0];

  if (!file) {
    if (keepLogoPath.value) showLogoPreview(keepLogoPath.value);
    else clearLogoPreview();
    return;
  }

  const reader = new FileReader();
  reader.onload = () => showLogoPreview(reader.result);
  reader.readAsDataURL(file);
}

function handleAmountInput(event) {
  let raw = parseInputNumber(event.target.value);

  const parts = raw.split(".");
  if (parts.length > 2) {
    raw = `${parts[0]}.${parts.slice(1).join("")}`;
  }

  txAmountRaw.value = raw;
  txAmount.value = raw ? formatInputNumber(raw) : "";
}

function handleProjectAmountInput(event) {
  let raw = parseInputNumber(event.target.value);

  const parts = raw.split(".");
  if (parts.length > 2) {
    raw = `${parts[0]}.${parts.slice(1).join("")}`;
  }

  totalPriceRaw.value = raw;
  totalPrice.value = raw ? formatInputNumber(raw) : "";
  updateProjectPricePreview();
}

function handleTxTypeChange() {
  const defaultCategory = getDefaultTxCategoryByType(txType.value);
  if (defaultCategory && !txCategory.value.trim()) {
    txCategory.value = defaultCategory;
  }
}

async function handleLogout() {
  try {
    await fetchJSON("/api/logout", { method: "POST" });
    window.location.href = "/login.html";
  } catch (error) {
    alert(error.message);
  }
}

function attachEvents() {
  projectForm.addEventListener("submit", submitProjectForm);
  transactionForm.addEventListener("submit", submitTransactionForm);

  resetBtn.addEventListener("click", resetProjectForm);
  clearTxBtn.addEventListener("click", () => resetTransactionForm(true));

  refreshBtn.addEventListener("click", () => {
    loadProjects(true).catch((error) => alert(error.message));
  });

  downloadExcelBtn.addEventListener("click", () => {
    window.location.href = "/api/download-excel";
  });

  logoutBtn?.addEventListener("click", handleLogout);

  companyLogo.addEventListener("change", handleLogoInputChange);
  txAmount.addEventListener("input", handleAmountInput);
  txType.addEventListener("change", handleTxTypeChange);

  totalPrice?.addEventListener("input", handleProjectAmountInput);
  vatPercent?.addEventListener("input", updateProjectPricePreview);
  contractCurrency?.addEventListener("change", updateProjectPricePreview);

  projectList.addEventListener("click", handleProjectListClick);
  projectModalHistory?.addEventListener("click", handleHistoryClick);
  txModalHistory?.addEventListener("click", handleHistoryClick);

  searchInput.addEventListener("input", renderProjectList);
  filterCategory.addEventListener("input", renderProjectList);
  filterOwner.addEventListener("input", renderProjectList);
  sortBy.addEventListener("change", renderProjectList);

  closeTxModalBtn?.addEventListener("click", closeTxModal);
  txModalBackdrop?.addEventListener("click", closeTxModal);

  closeProjectModalBtn?.addEventListener("click", closeProjectModal);
  projectModalBackdrop?.addEventListener("click", closeProjectModal);

  editProjectBtn?.addEventListener("click", () => {
    const project = allProjects.find((item) => item.id === currentProjectId);
    if (!project) return;
    populateProjectForm(project);
    closeProjectModal();
  });

  openAddTxFromDetailBtn?.addEventListener("click", () => {
    const project = allProjects.find((item) => item.id === currentProjectId);
    if (!project) return;
    closeProjectModal();
    openTxModal(project);
  });

  deleteProjectBtn?.addEventListener("click", async () => {
    if (!currentProjectId) return;
    await deleteProjectById(currentProjectId);
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;

    if (txModal && !txModal.classList.contains("hidden")) {
      closeTxModal();
      return;
    }

    if (projectModal && !projectModal.classList.contains("hidden")) {
      closeProjectModal();
    }
  });
}

async function init() {
  txDate.value = todayISO();
  attachEvents();
  resetProjectForm();
  resetTransactionForm(false);
  closeTxModal();
  closeProjectModal();
  updateProjectPricePreview();
  await loadProjects(true);
}

init().catch((error) => {
  console.error(error);
  alert(error.message || "Failed to load app");
});