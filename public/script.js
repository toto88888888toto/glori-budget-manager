const projectForm = document.getElementById("projectForm");
const editId = document.getElementById("editId");
const projectCode = document.getElementById("projectCode");
const projectName = document.getElementById("projectName");
const department = document.getElementById("department");
const budgetCategory = document.getElementById("budgetCategory");
const owner = document.getElementById("owner");
const budgetAmount = document.getElementById("budgetAmount");
const spentAmount = document.getElementById("spentAmount");
const status = document.getElementById("status");
const priority = document.getElementById("priority");
const startDate = document.getElementById("startDate");
const endDate = document.getElementById("endDate");
const remark = document.getElementById("remark");

const saveBtn = document.getElementById("saveBtn");
const resetBtn = document.getElementById("resetBtn");
const cancelEditBtn = document.getElementById("cancelEditBtn");
const downloadExcelBtn = document.getElementById("downloadExcelBtn");
const formTitle = document.getElementById("formTitle");

const searchInput = document.getElementById("searchInput");
const filterStatus = document.getElementById("filterStatus");
const filterDepartment = document.getElementById("filterDepartment");
const sortBy = document.getElementById("sortBy");
const resultCount = document.getElementById("resultCount");
const projectList = document.getElementById("projectList");
const projectCardTemplate = document.getElementById("projectCardTemplate");

const departmentSuggestions = document.getElementById("departmentSuggestions");
const categorySuggestions = document.getElementById("categorySuggestions");
const ownerSuggestions = document.getElementById("ownerSuggestions");

const kpiProjects = document.getElementById("kpiProjects");
const kpiBudget = document.getElementById("kpiBudget");
const kpiSpent = document.getElementById("kpiSpent");
const kpiBalance = document.getElementById("kpiBalance");

let allProjects = [];

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function parseNumber(value) {
  return Number(String(value || "").replace(/,/g, "")) || 0;
}

function formatNumberInput(input) {
  const numeric = parseNumber(input.value);
  input.value = numeric ? numeric.toLocaleString("en-US") : "";
}

function formatMoney(value) {
  const number = Number(value || 0);
  return number.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function textOrDash(value) {
  return value ? value : "-";
}

function formatDateLabel(value, prefix) {
  return `${prefix}: ${value || "-"}`;
}

function setFormMode(isEdit) {
  formTitle.textContent = isEdit ? "Edit Budget Project" : "Add Budget Project";
  saveBtn.textContent = isEdit ? "Update Project" : "Save Project";
  cancelEditBtn.classList.toggle("hidden", !isEdit);
}

async function fetchNextProjectCode() {
  if (editId.value) return;

  try {
    const res = await fetch("/api/next-project-code");
    const data = await res.json();
    projectCode.value = data.projectCode || "GB-0001";
  } catch (error) {
    console.error(error);
    projectCode.value = "GB-0001";
  }
}

function resetForm() {
  editId.value = "";
  projectForm.reset();
  status.value = "Planned";
  priority.value = "Medium";
  setFormMode(false);
  fetchNextProjectCode();
}

function buildSuggestionList(values, target) {
  target.innerHTML = values.map(value => `<option value="${escapeHtml(value)}"></option>`).join("");
}

function populateFiltersAndSuggestions() {
  const departments = [...new Set(allProjects.map(item => item.department).filter(Boolean))].sort();
  const categories = [...new Set(allProjects.map(item => item.budgetCategory).filter(Boolean))].sort();
  const owners = [...new Set(allProjects.map(item => item.owner).filter(Boolean))].sort();

  buildSuggestionList(departments, departmentSuggestions);
  buildSuggestionList(categories, categorySuggestions);
  buildSuggestionList(owners, ownerSuggestions);

  const currentDepartment = filterDepartment.value;
  filterDepartment.innerHTML = '<option value="">All Departments</option>' +
    departments.map(item => `<option value="${escapeHtml(item)}">${escapeHtml(item)}</option>`).join("");
  filterDepartment.value = currentDepartment;
}

function getFilteredProjects() {
  const keyword = searchInput.value.trim().toLowerCase();
  const statusValue = filterStatus.value;
  const departmentValue = filterDepartment.value;
  const sortValue = sortBy.value;

  let rows = allProjects.filter(project => {
    const haystack = [
      project.projectCode,
      project.projectName,
      project.department,
      project.budgetCategory,
      project.owner,
      project.status,
      project.priority,
      project.remark
    ].join(" ").toLowerCase();

    const matchKeyword = !keyword || haystack.includes(keyword);
    const matchStatus = !statusValue || project.status === statusValue;
    const matchDepartment = !departmentValue || project.department === departmentValue;

    return matchKeyword && matchStatus && matchDepartment;
  });

  rows.sort((a, b) => {
    if (sortValue === "budget-desc") return Number(b.budgetAmount || 0) - Number(a.budgetAmount || 0);
    if (sortValue === "budget-asc") return Number(a.budgetAmount || 0) - Number(b.budgetAmount || 0);
    if (sortValue === "balance-desc") return Number(b.balanceAmount || 0) - Number(a.balanceAmount || 0);
    if (sortValue === "name-asc") return String(a.projectName || "").localeCompare(String(b.projectName || ""));
    return new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime();
  });

  return rows;
}

function renderSummary() {
  const totalBudget = allProjects.reduce((sum, item) => sum + Number(item.budgetAmount || 0), 0);
  const totalSpent = allProjects.reduce((sum, item) => sum + Number(item.spentAmount || 0), 0);
  const totalBalance = totalBudget - totalSpent;

  kpiProjects.textContent = allProjects.length.toLocaleString("en-US");
  kpiBudget.textContent = formatMoney(totalBudget);
  kpiSpent.textContent = formatMoney(totalSpent);
  kpiBalance.textContent = formatMoney(totalBalance);
}

function renderProjects() {
  const projects = getFilteredProjects();
  resultCount.textContent = `${projects.length} project${projects.length === 1 ? "" : "s"}`;
  projectList.innerHTML = "";

  if (!projects.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.innerHTML = "<h3>No projects found</h3><p>Try changing your search or add a new project.</p>";
    projectList.appendChild(empty);
    return;
  }

  projects.forEach(project => {
    const node = projectCardTemplate.content.firstElementChild.cloneNode(true);
    const balance = Number(project.balanceAmount ?? (Number(project.budgetAmount || 0) - Number(project.spentAmount || 0)));

    node.querySelector(".project-code").textContent = project.projectCode || "-";
    node.querySelector(".project-title").textContent = project.projectName || "-";
    node.querySelector(".status-pill").textContent = project.status || "-";
    node.querySelector(".project-department").textContent = textOrDash(project.department);
    node.querySelector(".project-category").textContent = textOrDash(project.budgetCategory);
    node.querySelector(".project-owner").textContent = textOrDash(project.owner);
    node.querySelector(".project-priority").textContent = textOrDash(project.priority);
    node.querySelector(".project-budget").textContent = formatMoney(project.budgetAmount);
    node.querySelector(".project-spent").textContent = formatMoney(project.spentAmount);
    node.querySelector(".project-balance").textContent = formatMoney(balance);
    node.querySelector(".project-start").textContent = formatDateLabel(project.startDate, "Start");
    node.querySelector(".project-end").textContent = formatDateLabel(project.endDate, "End");
    node.querySelector(".project-remark").textContent = project.remark || "No remark";

    const statusPill = node.querySelector(".status-pill");
    if (project.status === "Completed") {
      statusPill.style.background = "#effcf5";
      statusPill.style.color = "#166534";
    } else if (project.status === "On Hold") {
      statusPill.style.background = "#fff8eb";
      statusPill.style.color = "#b45309";
    } else if (project.status === "Cancelled") {
      statusPill.style.background = "#fff1f2";
      statusPill.style.color = "#b91c1c";
    }

    node.querySelector(".btn-edit").addEventListener("click", () => startEdit(project));
    node.querySelector(".btn-delete").addEventListener("click", () => deleteProject(project));

    projectList.appendChild(node);
  });
}

async function loadProjects() {
  const res = await fetch("/api/projects");
  allProjects = await res.json();
  populateFiltersAndSuggestions();
  renderSummary();
  renderProjects();
  if (!editId.value) fetchNextProjectCode();
}

function startEdit(project) {
  editId.value = project.id || "";
  projectCode.value = project.projectCode || "";
  projectName.value = project.projectName || "";
  department.value = project.department || "";
  budgetCategory.value = project.budgetCategory || "";
  owner.value = project.owner || "";
  budgetAmount.value = project.budgetAmount ? Number(project.budgetAmount).toLocaleString("en-US") : "";
  spentAmount.value = project.spentAmount ? Number(project.spentAmount).toLocaleString("en-US") : "";
  status.value = project.status || "Planned";
  priority.value = project.priority || "Medium";
  startDate.value = project.startDate || "";
  endDate.value = project.endDate || "";
  remark.value = project.remark || "";
  setFormMode(true);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function deleteProject(project) {
  const ok = window.confirm(`Delete ${project.projectName || project.projectCode}?`);
  if (!ok) return;

  const res = await fetch(`/api/projects/${encodeURIComponent(project.id)}`, { method: "DELETE" });
  const data = await res.json();

  if (!res.ok) {
    alert(data.error || "Delete failed");
    return;
  }

  if (editId.value === project.id) resetForm();
  await loadProjects();
}

async function submitForm(event) {
  event.preventDefault();

  const payload = {
    editId: editId.value.trim(),
    projectCode: projectCode.value.trim(),
    projectName: projectName.value.trim(),
    department: department.value.trim(),
    budgetCategory: budgetCategory.value.trim(),
    owner: owner.value.trim(),
    budgetAmount: parseNumber(budgetAmount.value),
    spentAmount: parseNumber(spentAmount.value),
    status: status.value,
    priority: priority.value,
    startDate: startDate.value,
    endDate: endDate.value,
    remark: remark.value.trim()
  };

  saveBtn.disabled = true;
  saveBtn.textContent = editId.value ? "Updating..." : "Saving...";

  try {
    const res = await fetch("/api/projects", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Save failed");

    resetForm();
    await loadProjects();
  } catch (error) {
    console.error(error);
    alert(error.message || "Save failed");
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = editId.value ? "Update Project" : "Save Project";
  }
}

function attachEvents() {
  projectForm.addEventListener("submit", submitForm);
  resetBtn.addEventListener("click", resetForm);
  cancelEditBtn.addEventListener("click", resetForm);
  downloadExcelBtn.addEventListener("click", () => {
    window.location.href = "/api/download-excel";
  });

  [searchInput, filterStatus, filterDepartment, sortBy].forEach(element => {
    element.addEventListener("input", renderProjects);
    element.addEventListener("change", renderProjects);
  });

  [budgetAmount, spentAmount].forEach(input => {
    input.addEventListener("blur", () => formatNumberInput(input));
  });
}

attachEvents();
resetForm();
loadProjects().catch(error => {
  console.error(error);
  alert("Cannot load projects");
});
