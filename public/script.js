(async function () {
  const res = await fetch('/api/me');
  if (!res.ok) {
    window.location.href = '/login.html';
  }
})();

const $ = (id) => document.getElementById(id);

const projectForm = $("projectForm");
const transactionForm = $("transactionForm");

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

const saveBtn = $("saveBtn");
const resetBtn = $("resetBtn");
const addTxBtn = $("addTxBtn");
const clearTxBtn = $("clearTxBtn");
const refreshBtn = $("refreshBtn");
const downloadExcelBtn = $("downloadExcelBtn");

const projectList = $("projectList");
const projectDetail = $("projectDetail");
const detailCard = $("detailCard");

const searchInput = $("searchInput");
const filterCategory = $("filterCategory");
const filterOwner = $("filterOwner");
const sortBy = $("sortBy");

const kpiProjects = $("kpiProjects");
const kpiIncome = $("kpiIncome");
const kpiInvestment = $("kpiInvestment");
const kpiExpense = $("kpiExpense");

const projectCategoryList = $("projectCategoryList");
const txCategoryList = $("txCategoryList");
const filterCategoryList = $("filterCategoryList");
const filterOwnerList = $("filterOwnerList");

const txModal = $("txModal");
const txModalBackdrop = $("txModalBackdrop");
const closeTxModalBtn = $("closeTxModal");

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

function toNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function formatDisplayNumber(value) {
  return new Intl.NumberFormat("en-US", {
    maximumFractionDigits: 2,
  }).format(toNumber(value));
}

function formatMoney(value, currency = "") {
  return `${formatDisplayNumber(value)}${currency ? ` ${currency}` : ""}`;
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
    throw new Error(data?.error || "Request failed");
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
  saveBtn.textContent = "Update Project";
  showLogoPreview(project.logoPath || "");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function getFilteredProjects() {
  const keyword = searchInput.value.trim().toLowerCase();
  const categoryValue = filterCategory.value.trim().toLowerCase();
  const ownerValue = filterOwner.value.trim().toLowerCase();
  const sortValue = sortBy.value;

  let items = allProjects.filter((item) => {
    const haystack = [item.projectCode, item.projectName, item.category, item.owner]
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
        <div class="amount-box income">
          <div class="label">Income</div>
          <div class="amt">${formatDisplayNumber(item.totals?.income || 0)}</div>
        </div>

        <div class="amount-box investment">
          <div class="label">Investment</div>
          <div class="amt">${formatDisplayNumber(item.totals?.investment || 0)}</div>
        </div>

        <div class="amount-box expense">
          <div class="label">Expense</div>
          <div class="amt">${formatDisplayNumber(item.totals?.expense || 0)}</div>
        </div>

        <div class="amount-box balance">
          <div class="label">Balance</div>
          <div class="amt">${formatDisplayNumber(item.balance || 0)}</div>
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

function renderProjectDetail(project = null) {
  const currentProject = project || allProjects.find((item) => item.id === currentProjectId);

  if (!currentProject) {
    detailCard.classList.add("hidden");
    projectDetail.innerHTML = "";
    return;
  }

  detailCard.classList.remove("hidden");

  const transactions = currentProject.transactions || [];

  const historyHtml = transactions.length
    ? transactions
        .map((tx) => {
          const fileLink = tx.billPath
            ? `<a href="${tx.billPath}" target="_blank" rel="noopener">Open Bill</a>`
            : `<span style="color:#94a3b8;">No file</span>`;

          return `
            <div class="history-item">
              <div class="history-top">
                <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                  <span class="${getTypeBadgeClass(tx.type)}">${escapeHtml(
                    (tx.type || "").toUpperCase()
                  )}</span>
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
        })
        .join("")
    : `
      <div class="empty-state small">
        <h3>No transactions yet</h3>
        <p>Add the first transaction for this project.</p>
      </div>
    `;

  const logoHtml = currentProject.logoPath
    ? `<img src="${currentProject.logoPath}" alt="${escapeHtml(
        currentProject.projectName
      )}" style="width:100px;height:100px;object-fit:contain;border-radius:18px;border:1px solid #dbe4f0;background:#f8fbff;padding:8px;">`
    : `<div style="width:100px;height:100px;display:grid;place-items:center;border:1px solid #dbe5f0;border-radius:18px;background:#f7fbff;font-size:12px;font-weight:800;color:#94a3b8;">NO LOGO</div>`;

  projectDetail.innerHTML = `
    <div class="project-head" style="margin-bottom:18px;">
      ${logoHtml}
      <div class="project-title-wrap">
        <div class="project-code">${escapeHtml(currentProject.projectCode || "")}</div>
        <h3 class="project-name">${escapeHtml(currentProject.projectName || "")}</h3>
      </div>
    </div>

    <div class="project-meta" style="margin-bottom:14px;">
      <div class="meta-box">
        <div class="meta-label">Category</div>
        <div class="meta-value">${escapeHtml(currentProject.category || "-")}</div>
      </div>
      <div class="meta-box">
        <div class="meta-label">Owner</div>
        <div class="meta-value">${escapeHtml(currentProject.owner || "-")}</div>
      </div>
      <div class="meta-box">
        <div class="meta-label">Start Date</div>
        <div class="meta-value">${escapeHtml(currentProject.startDate || "-")}</div>
      </div>
      <div class="meta-box">
        <div class="meta-label">End Date</div>
        <div class="meta-value">${escapeHtml(currentProject.endDate || "-")}</div>
      </div>
    </div>

    <div class="amount-row" style="margin-bottom:14px;">
      <div class="amount-box income">
        <div class="label">Income</div>
        <div class="amt">${formatDisplayNumber(currentProject.totals?.income || 0)}</div>
      </div>

      <div class="amount-box investment">
        <div class="label">Investment</div>
        <div class="amt">${formatDisplayNumber(currentProject.totals?.investment || 0)}</div>
      </div>

      <div class="amount-box expense">
        <div class="label">Expense</div>
        <div class="amt">${formatDisplayNumber(currentProject.totals?.expense || 0)}</div>
      </div>

      <div class="amount-box balance">
        <div class="label">Balance</div>
        <div class="amt">${formatDisplayNumber(currentProject.balance || 0)}</div>
      </div>
    </div>

    <div class="project-foot" style="margin-bottom:14px;">
      <div class="project-stat">Transactions <strong>${toNumber(
        currentProject.transactionCount
      )}</strong></div>
      <div class="project-period">${formatDateRange(
        currentProject.startDate,
        currentProject.endDate
      )}</div>
    </div>

    <div class="project-remark" style="margin-bottom:18px;">
      <div class="project-remark-label">Remark</div>
      <p>${escapeHtml(getProjectRemarkText(currentProject))}</p>
    </div>

    <div class="section-head" style="margin-bottom:14px;">
      <h3>Transaction History</h3>
      <p>${toNumber(currentProject.transactionCount)} item(s)</p>
    </div>

    <div class="history-list">${historyHtml}</div>
  `;
}

function openProject(project, scrollToDetail = true) {
  if (!project) return;

  currentProjectId = project.id || "";
  selectedProjectId.value = currentProjectId;
  selectedProjectText.value = `${project.projectCode || "-"} - ${project.projectName || "-"}`;

  renderProjectList();
  renderProjectDetail(project);

  if (scrollToDetail) {
    detailCard.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function openTxModal(project) {
  if (!project || !txModal) return;

  currentProjectId = project.id || "";
  selectedProjectId.value = currentProjectId;
  selectedProjectText.value = `${project.projectCode || "-"} - ${project.projectName || "-"}`;
  txDate.value = todayISO();

  renderProjectList();
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
  allProjects = Array.isArray(projects) ? projects : [];

  refreshDatalists(allProjects);
  updateKPIs(allProjects);

  if (
    keepSelection &&
    currentProjectId &&
    !allProjects.find((item) => item.id === currentProjectId)
  ) {
    currentProjectId = "";
  }

  if (currentProjectId) {
    const selected = allProjects.find((item) => item.id === currentProjectId);
    if (selected) {
      selectedProjectId.value = selected.id || "";
      selectedProjectText.value = `${selected.projectCode || "-"} - ${selected.projectName || "-"}`;
      renderProjectList();
      renderProjectDetail(selected);
      return;
    }
  }

  currentProjectId = "";
  selectedProjectId.value = "";
  selectedProjectText.value = "";
  renderProjectList();
  renderProjectDetail(null);
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
    openProject(result.project, false);
    resetProjectForm();
    alert(isEdit ? "Project updated successfully" : "Project saved successfully");
  } catch (error) {
    alert(error.message);
  } finally {
    setButtonLoading(
      saveBtn,
      false,
      editId.value ? "Update Project" : "Save Project",
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
    if (selected) openProject(selected, false);

    resetTransactionForm(true);
    closeTxModal();
    alert("Transaction saved successfully");
  } catch (error) {
    alert(error.message);
  } finally {
    setButtonLoading(addTxBtn, false, "Add Transaction", "");
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
      openProject(project);
      return;
    }

    if (action === "add-tx") {
      openProject(project, false);
      openTxModal(project);
      return;
    }

    if (action === "edit") {
      openProject(project, false);
      populateProjectForm(project);
      return;
    }

    if (action === "delete") {
      if (!confirm(`Delete project "${project.projectName}" and all related transactions?`)) {
        return;
      }

      try {
        await fetchJSON(`/api/projects/${id}`, { method: "DELETE" });

        if (currentProjectId === id) {
          currentProjectId = "";
          selectedProjectId.value = "";
          selectedProjectText.value = "";
          renderProjectDetail(null);
          closeTxModal();
        }

        await loadProjects(true);
        resetProjectForm();
        alert("Project deleted");
      } catch (error) {
        alert(error.message);
      }
    }

    return;
  }

  const card = event.target.closest(".project-card");
  if (!card) return;

  const id = card.dataset.projectId;
  const project = allProjects.find((item) => item.id === id);
  if (!project) return;

  openProject(project, false);
  openTxModal(project);
}

async function handleDetailClick(event) {
  const button = event.target.closest('button[data-action="delete-tx"]');
  if (!button) return;

  const txId = button.dataset.id;
  if (!confirm("Delete this transaction?")) return;

  try {
    await fetchJSON(`/api/transactions/${txId}`, { method: "DELETE" });
    await loadProjects(false);

    const selected = allProjects.find((item) => item.id === currentProjectId);
    if (selected) openProject(selected, false);
    else renderProjectDetail(null);

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

function handleAmountInput(event) {
  let raw = parseInputNumber(event.target.value);

  const parts = raw.split(".");
  if (parts.length > 2) {
    raw = `${parts[0]}.${parts.slice(1).join("")}`;
  }

  txAmountRaw.value = raw;
  txAmount.value = raw ? formatInputNumber(raw) : "";
}

function handleTxTypeChange() {
  const defaultCategory = getDefaultTxCategoryByType(txType.value);

  if (defaultCategory && !txCategory.value.trim()) {
    txCategory.value = defaultCategory;
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

  companyLogo.addEventListener("change", handleLogoInputChange);
  txAmount.addEventListener("input", handleAmountInput);
  txType.addEventListener("change", handleTxTypeChange);

  projectList.addEventListener("click", handleProjectListClick);
  projectDetail.addEventListener("click", handleDetailClick);

  searchInput.addEventListener("input", renderProjectList);
  filterCategory.addEventListener("input", renderProjectList);
  filterOwner.addEventListener("input", renderProjectList);
  sortBy.addEventListener("change", renderProjectList);

  closeTxModalBtn?.addEventListener("click", closeTxModal);
  txModalBackdrop?.addEventListener("click", closeTxModal);

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && txModal && !txModal.classList.contains("hidden")) {
      closeTxModal();
    }
  });
}

async function init() {
  txDate.value = todayISO();
  attachEvents();
  resetProjectForm();
  resetTransactionForm(false);
  closeTxModal();
  await loadProjects(true);
}

init().catch((error) => {
  console.error(error);
  alert(error.message || "Failed to load app");
});