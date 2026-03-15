const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid");

const app = express();
const PORT = process.env.PORT || 3000;

const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const DATA_DIR = path.join(ROOT_DIR, "data");
const EXCEL_FILE = path.join(DATA_DIR, "budget.xlsx");
const SHEET_NAME = "Projects";

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(PUBLIC_DIR));

function headers() {
  return [
    "id",
    "projectCode",
    "projectName",
    "department",
    "budgetCategory",
    "owner",
    "budgetAmount",
    "spentAmount",
    "balanceAmount",
    "status",
    "priority",
    "startDate",
    "endDate",
    "remark",
    "createdAt",
    "updatedAt"
  ];
}

async function ensureWorkbook() {
  const workbook = new ExcelJS.Workbook();

  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }

  let sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) sheet = workbook.addWorksheet(SHEET_NAME);

  const wantedHeaders = headers();
  const hasRows = sheet.rowCount > 0 && sheet.getRow(1).cellCount > 0;

  if (!hasRows) {
    sheet.addRow(wantedHeaders);
    styleWorksheet(sheet);
  } else {
    const currentHeaders = sheet.getRow(1).values.slice(1).map(v => String(v || "").trim());
    const same = JSON.stringify(currentHeaders) === JSON.stringify(wantedHeaders);

    if (!same) {
      sheet.spliceRows(1, sheet.rowCount);
      sheet.addRow(wantedHeaders);
      styleWorksheet(sheet);
    }
  }

  await workbook.xlsx.writeFile(EXCEL_FILE);
  return { workbook, sheet };
}

function styleWorksheet(sheet) {
  sheet.columns = [
    { width: 16 },
    { width: 14 },
    { width: 28 },
    { width: 18 },
    { width: 18 },
    { width: 20 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 12 },
    { width: 14 },
    { width: 14 },
    { width: 34 },
    { width: 24 },
    { width: 24 }
  ];

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF2563EB" }
  };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };
  headerRow.height = 22;

  sheet.views = [{ state: "frozen", ySplit: 1 }];
}

function rowToProject(row) {
  const values = row.values;
  const budgetAmount = Number(values[7] || 0);
  const spentAmount = Number(values[8] || 0);
  return {
    id: values[1] || "",
    projectCode: values[2] || "",
    projectName: values[3] || "",
    department: values[4] || "",
    budgetCategory: values[5] || "",
    owner: values[6] || "",
    budgetAmount,
    spentAmount,
    balanceAmount: Number(values[9] || budgetAmount - spentAmount || 0),
    status: values[10] || "Planned",
    priority: values[11] || "Medium",
    startDate: values[12] || "",
    endDate: values[13] || "",
    remark: values[14] || "",
    createdAt: values[15] || "",
    updatedAt: values[16] || ""
  };
}

function projectToRow(project) {
  const budgetAmount = Number(project.budgetAmount || 0);
  const spentAmount = Number(project.spentAmount || 0);
  const balanceAmount = budgetAmount - spentAmount;

  return [
    project.id,
    project.projectCode,
    project.projectName,
    project.department,
    project.budgetCategory,
    project.owner,
    budgetAmount,
    spentAmount,
    balanceAmount,
    project.status,
    project.priority,
    project.startDate,
    project.endDate,
    project.remark,
    project.createdAt,
    project.updatedAt
  ];
}

async function readAllProjects() {
  const { sheet } = await ensureWorkbook();
  const rows = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (!row.getCell(1).value) return;
    rows.push({ ...rowToProject(row), _rowNumber: rowNumber });
  });

  return rows.sort((a, b) => {
    const aDate = new Date(a.createdAt || 0).getTime();
    const bDate = new Date(b.createdAt || 0).getTime();
    return bDate - aDate;
  });
}

function nextProjectCode(projects) {
  const max = projects.reduce((acc, project) => {
    const num = Number(String(project.projectCode || "").replace(/\D/g, "")) || 0;
    return Math.max(acc, num);
  }, 0);
  return `GB-${String(max + 1).padStart(4, "0")}`;
}

function normalizeProjectInput(body, projects, existing) {
  const budgetAmount = Number(body.budgetAmount || 0);
  const spentAmount = Number(body.spentAmount || 0);

  return {
    id: existing?.id || uuidv4(),
    projectCode: String(body.projectCode || existing?.projectCode || "").trim() || nextProjectCode(projects),
    projectName: String(body.projectName || "").trim(),
    department: String(body.department || "").trim(),
    budgetCategory: String(body.budgetCategory || "").trim(),
    owner: String(body.owner || "").trim(),
    budgetAmount,
    spentAmount,
    status: String(body.status || existing?.status || "Planned").trim() || "Planned",
    priority: String(body.priority || existing?.priority || "Medium").trim() || "Medium",
    startDate: String(body.startDate || "").trim(),
    endDate: String(body.endDate || "").trim(),
    remark: String(body.remark || "").trim(),
    createdAt: existing?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function validateProject(project) {
  if (!project.projectName) return "Project Name is required";
  if (!project.department) return "Department is required";
  if (!project.budgetCategory) return "Budget Category is required";
  if (Number.isNaN(project.budgetAmount) || project.budgetAmount < 0) return "Budget Amount is invalid";
  if (Number.isNaN(project.spentAmount) || project.spentAmount < 0) return "Spent Amount is invalid";
  return "";
}

function buildSummary(projects) {
  const totalBudget = projects.reduce((sum, p) => sum + Number(p.budgetAmount || 0), 0);
  const totalSpent = projects.reduce((sum, p) => sum + Number(p.spentAmount || 0), 0);
  const totalBalance = totalBudget - totalSpent;
  const projectCount = projects.length;
  const activeCount = projects.filter(p => ["Planned", "In Progress", "On Hold"].includes(p.status)).length;
  const completedCount = projects.filter(p => p.status === "Completed").length;

  return {
    projectCount,
    activeCount,
    completedCount,
    totalBudget,
    totalSpent,
    totalBalance
  };
}

app.get("/api/projects", async (req, res) => {
  try {
    const projects = await readAllProjects();
    res.json(projects.map(({ _rowNumber, ...rest }) => rest));
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Cannot read projects" });
  }
});

app.get("/api/summary", async (req, res) => {
  try {
    const projects = await readAllProjects();
    res.json(buildSummary(projects));
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Cannot build summary" });
  }
});

app.get("/api/next-project-code", async (req, res) => {
  try {
    const projects = await readAllProjects();
    res.json({ projectCode: nextProjectCode(projects) });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Cannot generate project code" });
  }
});

app.post("/api/projects", async (req, res) => {
  try {
    const { workbook, sheet } = await ensureWorkbook();
    const projects = await readAllProjects();
    const editId = String(req.body.editId || "").trim();

    if (editId) {
      const existing = projects.find(project => project.id === editId);
      if (!existing) return res.status(404).json({ error: "Project not found" });

      const updated = normalizeProjectInput(req.body, projects, existing);
      const errorMessage = validateProject(updated);
      if (errorMessage) return res.status(400).json({ error: errorMessage });

      const row = sheet.getRow(existing._rowNumber);
      row.values = [null, ...projectToRow(updated)];
      row.commit();
      await workbook.xlsx.writeFile(EXCEL_FILE);
      return res.json({ ok: true, project: updated });
    }

    const created = normalizeProjectInput(req.body, projects);
    const errorMessage = validateProject(created);
    if (errorMessage) return res.status(400).json({ error: errorMessage });

    sheet.addRow(projectToRow(created));
    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ ok: true, project: created });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Cannot save project" });
  }
});

app.delete("/api/projects/:id", async (req, res) => {
  try {
    const { workbook, sheet } = await ensureWorkbook();
    const projects = await readAllProjects();
    const target = projects.find(project => project.id === req.params.id);

    if (!target) return res.status(404).json({ error: "Project not found" });

    sheet.spliceRows(target._rowNumber, 1);
    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ ok: true });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Cannot delete project" });
  }
});

app.get("/api/download-excel", async (req, res) => {
  try {
    await ensureWorkbook();
    return res.download(EXCEL_FILE, "glori-budget.xlsx");
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Cannot download Excel" });
  }
});

app.get("*", (req, res) => {
  res.sendFile(path.join(PUBLIC_DIR, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Glori Budget Manager running on http://localhost:${PORT}`);
});
