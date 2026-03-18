const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const session = require('express-session');
const bcrypt = require('bcrypt');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 3000;
const IS_PROD = process.env.NODE_ENV === 'production';

const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, 'public');
const DATA_DIR = path.join(ROOT_DIR, 'data');
const UPLOAD_DIR = path.join(ROOT_DIR, 'uploads');
const EXCEL_FILE = path.join(DATA_DIR, 'budget.xlsx');

const LOGIN_HTML = path.join(PUBLIC_DIR, 'login.html');
const INDEX_HTML = path.join(PUBLIC_DIR, 'index.html');

const PROJECT_SHEET = 'Projects';
const TRANSACTION_SHEET = 'Transactions';

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

app.set('trust proxy', 1);

app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true }));
app.use('/uploads', express.static(UPLOAD_DIR));
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "login.html"));
});

app.use(
  session({
    name: 'glori.sid',
    secret: process.env.SESSION_SECRET || 'glori_secret_2026_change_me',
    resave: false,
    saveUninitialized: false,
    rolling: true,
    proxy: true,
    cookie: {
      httpOnly: true,
      secure: IS_PROD,
      sameSite: 'lax',
      maxAge: 1000 * 60 * 60 * 8
    }
  })
);

app.use((req, res, next) => {
  console.log(
    `[${new Date().toISOString()}] ${req.method} ${req.url} | secure=${req.secure} | xfwd=${
      req.headers['x-forwarded-proto'] || '-'
    }`
  );
  next();
});

const USERS = [
  {
    username: 'bin',
    passwordHash:
      '$2b$10$6nm6.uHG3zmS7/u4nliIMukMkuZYJsTpnOE2ugpwvXKRifDKbrhCS'
  }
];

function isPageRequest(req) {
  const accept = String(req.headers.accept || '');
  return accept.includes('text/html');
}

function requireAuth(req, res, next) {
  if (req.session?.user) return next();

  if (isPageRequest(req)) {
    return res.redirect('/login.html');
  }

  return res.status(401).json({ ok: false, message: 'Unauthorized' });
}

function sendError(res, message, status = 500) {
  return res.status(status).json({ ok: false, error: message });
}

function toNumber(value) {
  return Number(String(value ?? '').replace(/,/g, '').trim()) || 0;
}

function normalizeDate(value) {
  if (!value) return '';
  if (typeof value === 'string') return value.slice(0, 10);
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  return String(value).slice(0, 10);
}

function toText(value) {
  return String(value ?? '').trim();
}

function publicPathFromFile(file) {
  return file ? `/uploads/${file.filename}` : '';
}

function removePublicFile(filePath) {
  try {
    if (!filePath || typeof filePath !== 'string' || !filePath.startsWith('/uploads/')) return;
    const fullPath = path.join(UPLOAD_DIR, filePath.replace('/uploads/', ''));
    if (fs.existsSync(fullPath)) fs.unlinkSync(fullPath);
  } catch (error) {
    console.error('Failed deleting file', filePath, error.message);
  }
}

function styleSheet(sheet, count) {
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.columns = Array.from({ length: count }, () => ({ width: 22 }));

  const row = sheet.getRow(1);
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.alignment = { vertical: 'middle', horizontal: 'center' };
  row.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF2563EB' }
  };
  row.height = 22;
}

function ensureHeaders(sheet, headers) {
  if (sheet.rowCount === 0) {
    sheet.addRow(headers);
  } else {
    const current = sheet
      .getRow(1)
      .values.slice(1)
      .map((v) => String(v || '').trim());

    const matches = JSON.stringify(current) === JSON.stringify(headers);

    if (!matches) {
      if (sheet.rowCount >= 1) sheet.spliceRows(1, 1);
      sheet.insertRow(1, headers);
    }
  }

  styleSheet(sheet, headers.length);
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname || '');
    const base = path
      .basename(file.originalname || 'file', ext)
      .replace(/[^a-zA-Z0-9_-]/g, '_')
      .slice(0, 80);

    cb(null, `${Date.now()}-${base}${ext}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 20 * 1024 * 1024, files: 10 }
});

const projectUpload = upload.fields([{ name: 'companyLogo', maxCount: 1 }]);
const transactionUpload = upload.fields([{ name: 'billFile', maxCount: 1 }]);

const PROJECT_HEADERS = [
  'id',
  'no',
  'projectCode',
  'projectName',
  'category',
  'owner',
  'startDate',
  'endDate',
  'remark',
  'logoPath',
  'createdAt',
  'updatedAt'
];

const TRANSACTION_HEADERS = [
  'id',
  'no',
  'projectId',
  'type',
  'category',
  'description',
  'currency',
  'amount',
  'date',
  'billPath',
  'createdAt',
  'updatedAt'
];

let workbookCache = null;
let projectSheetCache = null;
let transactionSheetCache = null;

function invalidateWorkbookCache() {
  workbookCache = null;
  projectSheetCache = null;
  transactionSheetCache = null;
}

async function openWorkbook() {
  if (workbookCache && projectSheetCache && transactionSheetCache) {
    return {
      workbook: workbookCache,
      projectSheet: projectSheetCache,
      transactionSheet: transactionSheetCache
    };
  }

  const workbook = new ExcelJS.Workbook();

  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }

  let projectSheet = workbook.getWorksheet(PROJECT_SHEET);
  let transactionSheet = workbook.getWorksheet(TRANSACTION_SHEET);

  if (!projectSheet) projectSheet = workbook.addWorksheet(PROJECT_SHEET);
  if (!transactionSheet) transactionSheet = workbook.addWorksheet(TRANSACTION_SHEET);

  ensureHeaders(projectSheet, PROJECT_HEADERS);
  ensureHeaders(transactionSheet, TRANSACTION_HEADERS);

  workbookCache = workbook;
  projectSheetCache = projectSheet;
  transactionSheetCache = transactionSheet;

  return { workbook, projectSheet, transactionSheet };
}

let writeQueue = Promise.resolve();

function queueWrite(task) {
  const run = writeQueue.then(task, task);
  writeQueue = run.catch(() => {});
  return run;
}

async function saveWorkbook(workbook) {
  await workbook.xlsx.writeFile(EXCEL_FILE);
  invalidateWorkbookCache();
}

function rowToProject(row) {
  const values = row.values;
  return {
    id: toText(values[1]),
    no: toNumber(values[2]),
    projectCode: toText(values[3]),
    projectName: toText(values[4]),
    category: toText(values[5]),
    owner: toText(values[6]),
    startDate: normalizeDate(values[7]),
    endDate: normalizeDate(values[8]),
    remark: toText(values[9]),
    logoPath: toText(values[10]),
    createdAt: toText(values[11]),
    updatedAt: toText(values[12]),
    _rowNumber: row.number
  };
}

function rowToTransaction(row) {
  const values = row.values;
  return {
    id: toText(values[1]),
    no: toNumber(values[2]),
    projectId: toText(values[3]),
    type: toText(values[4]).toLowerCase(),
    category: toText(values[5]),
    description: toText(values[6]),
    currency: toText(values[7]) || 'LAK',
    amount: toNumber(values[8]),
    date: normalizeDate(values[9]),
    billPath: toText(values[10]),
    createdAt: toText(values[11]),
    updatedAt: toText(values[12]),
    _rowNumber: row.number
  };
}

function projectToRow(project) {
  return [
    project.id,
    toNumber(project.no),
    project.projectCode,
    project.projectName,
    project.category,
    project.owner,
    project.startDate,
    project.endDate,
    project.remark,
    project.logoPath,
    project.createdAt,
    project.updatedAt
  ];
}

function transactionToRow(tx) {
  return [
    tx.id,
    toNumber(tx.no),
    tx.projectId,
    tx.type,
    tx.category,
    tx.description,
    tx.currency,
    toNumber(tx.amount),
    tx.date,
    tx.billPath,
    tx.createdAt,
    tx.updatedAt
  ];
}

async function getAllData() {
  const { projectSheet, transactionSheet } = await openWorkbook();
  const projects = [];
  const transactions = [];

  projectSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (!row.getCell(1).value && !row.getCell(4).value) return;
    projects.push(rowToProject(row));
  });

  transactionSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (!row.getCell(1).value && !row.getCell(3).value) return;
    transactions.push(rowToTransaction(row));
  });

  return { projects, transactions };
}

function nextProjectNo(projects) {
  return projects.reduce((max, item) => Math.max(max, toNumber(item.no)), 0) + 1;
}

function nextProjectCode(projects) {
  const max = projects.reduce((current, item) => {
    const match = String(item.projectCode || '').match(/GB-(\d+)/i);
    const num = match ? Number(match[1]) : 0;
    return Math.max(current, num);
  }, 0);

  return `GB-${String(max + 1).padStart(5, '0')}`;
}

function nextTransactionNo(transactions, projectId) {
  return (
    transactions
      .filter((tx) => tx.projectId === projectId)
      .reduce((max, item) => Math.max(max, toNumber(item.no)), 0) + 1
  );
}

function validateProject(project) {
  if (!project.projectName) return 'Project name is required';
  if (!project.category) return 'Category is required';
  return '';
}

function validateTransaction(tx) {
  if (!tx.projectId) return 'Project is required';
  if (!['income', 'investment', 'expense'].includes(tx.type)) return 'Type is invalid';
  if (!tx.category) return 'Category is required';
  if (toNumber(tx.amount) <= 0) return 'Amount must be greater than 0';
  if (!tx.date) return 'Date is required';
  return '';
}

function projectPayload(body, existing, req, projects) {
  const uploadedLogo = publicPathFromFile(req.files?.companyLogo?.[0]);

  return {
    id: existing?.id || uuidv4(),
    no: existing?.no || nextProjectNo(projects),
    projectCode: existing?.projectCode || nextProjectCode(projects),
    projectName: toText(body.projectName || existing?.projectName),
    category: toText(body.category || existing?.category),
    owner: toText(body.owner || existing?.owner),
    startDate: normalizeDate(body.startDate || existing?.startDate),
    endDate: normalizeDate(body.endDate || existing?.endDate),
    remark: toText(body.remark || existing?.remark),
    logoPath: uploadedLogo || toText(body.keepLogoPath || existing?.logoPath),
    createdAt: existing?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function transactionPayload(projectId, body, req, transactions) {
  return {
    id: uuidv4(),
    no: nextTransactionNo(transactions, projectId),
    projectId,
    type: toText(body.type).toLowerCase(),
    category: toText(body.category),
    description: toText(body.description),
    currency: toText(body.currency || 'LAK').toUpperCase(),
    amount: toNumber(body.amount),
    date: normalizeDate(body.date),
    billPath: publicPathFromFile(req.files?.billFile?.[0]),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function buildProjectSummary(project, transactions) {
  const related = transactions
    .filter((tx) => tx.projectId === project.id)
    .sort((a, b) => {
      const aTime = new Date(a.date || a.createdAt || 0).getTime();
      const bTime = new Date(b.date || b.createdAt || 0).getTime();
      return bTime - aTime;
    });

  const totals = related.reduce(
    (sum, tx) => {
      const amount = toNumber(tx.amount);
      if (tx.type === 'income') sum.income += amount;
      if (tx.type === 'investment') sum.investment += amount;
      if (tx.type === 'expense') sum.expense += amount;
      return sum;
    },
    { income: 0, investment: 0, expense: 0 }
  );

  return {
    ...project,
    transactions: related.map(({ _rowNumber, ...tx }) => tx),
    totals,
    balance: totals.income - totals.investment - totals.expense,
    transactionCount: related.length
  };
}

app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    app: 'Glori Budget Manager',
    uptime: process.uptime(),
    timestamp: new Date().toISOString(),
    env: process.env.NODE_ENV || 'development'
  });
});

app.post('/api/login', async (req, res) => {
  try {
    const username = String(req.body.username || '').trim();
    const password = String(req.body.password || '');

    const user = USERS.find((u) => u.username === username);
    if (!user) {
      return res.status(401).json({ ok: false, message: 'Invalid username or password' });
    }

    const matched = await bcrypt.compare(password, user.passwordHash);
    if (!matched) {
      return res.status(401).json({ ok: false, message: 'Invalid username or password' });
    }

    req.session.user = { username: user.username };

    req.session.save((err) => {
      if (err) {
        console.error('Session save error:', err);
        return res.status(500).json({ ok: false, message: 'Login failed' });
      }
      return res.json({ ok: true, user: req.session.user });
    });
  } catch (error) {
    console.error('Login error:', error);
    return res.status(500).json({ ok: false, message: 'Login failed' });
  }
});

app.post('/api/logout', requireAuth, (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      console.error('Logout error:', err);
      return res.status(500).json({ ok: false, message: 'Logout failed' });
    }

    res.clearCookie('glori.sid', {
      httpOnly: true,
      secure: IS_PROD,
      sameSite: 'lax'
    });

    return res.json({ ok: true });
  });
});

app.get('/api/me', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ ok: false });
  }
  return res.json({ ok: true, user: req.session.user });
});

app.get('/', (req, res) => {
  if (req.session?.user) {
    return res.redirect('/index.html');
  }
  return res.sendFile(LOGIN_HTML);
});

app.get('/login.html', (req, res) => {
  if (req.session?.user) {
    return res.redirect('/index.html');
  }
  return res.sendFile(LOGIN_HTML);
});

app.get('/index.html', requireAuth, (req, res) => {
  return res.sendFile(INDEX_HTML);
});

app.get('/api/projects', requireAuth, async (req, res) => {
  const started = Date.now();
  try {
    const { projects, transactions } = await getAllData();
    const items = projects
      .map((project) => buildProjectSummary(project, transactions))
      .sort((a, b) => toNumber(b.no) - toNumber(a.no));

    console.log('/api/projects took', Date.now() - started, 'ms');
    return res.json(items);
  } catch (error) {
    console.error('Cannot read projects:', error);
    return sendError(res, 'Cannot read projects');
  }
});

app.get('/api/projects/:id', requireAuth, async (req, res) => {
  const started = Date.now();
  try {
    const { projects, transactions } = await getAllData();
    const project = projects.find((item) => item.id === req.params.id);
    if (!project) return sendError(res, 'Project not found', 404);

    console.log('/api/projects/:id took', Date.now() - started, 'ms');
    return res.json(buildProjectSummary(project, transactions));
  } catch (error) {
    console.error('Cannot read project:', error);
    return sendError(res, 'Cannot read project');
  }
});

app.get('/api/next-project-code', requireAuth, async (req, res) => {
  try {
    const { projects } = await getAllData();
    return res.json({
      no: nextProjectNo(projects),
      projectCode: nextProjectCode(projects)
    });
  } catch (error) {
    console.error('Cannot generate project code:', error);
    return sendError(res, 'Cannot generate project code');
  }
});

app.post('/api/projects', requireAuth, projectUpload, async (req, res) => {
  try {
    const result = await queueWrite(async () => {
      const { workbook, projectSheet } = await openWorkbook();
      const { projects, transactions } = await getAllData();

      const project = projectPayload(req.body, null, req, projects);
      const validation = validateProject(project);
      if (validation) {
        return { status: 400, body: { ok: false, error: validation } };
      }

      projectSheet.addRow(projectToRow(project));
      await saveWorkbook(workbook);

      return {
        status: 200,
        body: { ok: true, project: buildProjectSummary(project, transactions) }
      };
    });

    return res.status(result.status).json(result.body);
  } catch (error) {
    console.error('Cannot save project:', error);
    return sendError(res, 'Cannot save project');
  }
});

app.put('/api/projects/:id', requireAuth, projectUpload, async (req, res) => {
  try {
    const result = await queueWrite(async () => {
      const { workbook, projectSheet } = await openWorkbook();
      const { projects, transactions } = await getAllData();

      const existing = projects.find((item) => item.id === req.params.id);
      if (!existing) {
        return { status: 404, body: { ok: false, error: 'Project not found' } };
      }

      const updated = projectPayload(req.body, existing, req, projects);
      const validation = validateProject(updated);
      if (validation) {
        return { status: 400, body: { ok: false, error: validation } };
      }

      if (req.files?.companyLogo?.[0] && existing.logoPath && existing.logoPath !== updated.logoPath) {
        removePublicFile(existing.logoPath);
      }

      const row = projectSheet.getRow(existing._rowNumber);
      row.values = [null, ...projectToRow(updated)];
      row.commit();

      await saveWorkbook(workbook);

      return {
        status: 200,
        body: { ok: true, project: buildProjectSummary(updated, transactions) }
      };
    });

    return res.status(result.status).json(result.body);
  } catch (error) {
    console.error('Cannot update project:', error);
    return sendError(res, 'Cannot update project');
  }
});

app.delete('/api/projects/:id', requireAuth, async (req, res) => {
  try {
    const result = await queueWrite(async () => {
      const { workbook, projectSheet, transactionSheet } = await openWorkbook();
      const { projects, transactions } = await getAllData();

      const existing = projects.find((item) => item.id === req.params.id);
      if (!existing) {
        return { status: 404, body: { ok: false, error: 'Project not found' } };
      }

      if (existing.logoPath) removePublicFile(existing.logoPath);

      const related = transactions.filter((tx) => tx.projectId === existing.id);
      related.forEach((tx) => {
        if (tx.billPath) removePublicFile(tx.billPath);
      });

      related
        .map((tx) => tx._rowNumber)
        .sort((a, b) => b - a)
        .forEach((rowNumber) => transactionSheet.spliceRows(rowNumber, 1));

      projectSheet.spliceRows(existing._rowNumber, 1);
      await saveWorkbook(workbook);

      return { status: 200, body: { ok: true } };
    });

    return res.status(result.status).json(result.body);
  } catch (error) {
    console.error('Cannot delete project:', error);
    return sendError(res, 'Cannot delete project');
  }
});

app.post('/api/projects/:id/transactions', requireAuth, transactionUpload, async (req, res) => {
  try {
    const result = await queueWrite(async () => {
      const { workbook, transactionSheet } = await openWorkbook();
      const { projects, transactions } = await getAllData();

      const project = projects.find((item) => item.id === req.params.id);
      if (!project) {
        return { status: 404, body: { ok: false, error: 'Project not found' } };
      }

      const tx = transactionPayload(project.id, req.body, req, transactions);
      const validation = validateTransaction(tx);
      if (validation) {
        return { status: 400, body: { ok: false, error: validation } };
      }

      transactionSheet.addRow(transactionToRow(tx));
      await saveWorkbook(workbook);

      return {
        status: 200,
        body: {
          ok: true,
          transaction: tx,
          project: buildProjectSummary(project, [...transactions, tx])
        }
      };
    });

    return res.status(result.status).json(result.body);
  } catch (error) {
    console.error('Cannot save transaction:', error);
    return sendError(res, 'Cannot save transaction');
  }
});

app.delete('/api/transactions/:id', requireAuth, async (req, res) => {
  try {
    const result = await queueWrite(async () => {
      const { workbook, transactionSheet } = await openWorkbook();
      const { transactions } = await getAllData();

      const tx = transactions.find((item) => item.id === req.params.id);
      if (!tx) {
        return { status: 404, body: { ok: false, error: 'Transaction not found' } };
      }

      if (tx.billPath) removePublicFile(tx.billPath);
      transactionSheet.spliceRows(tx._rowNumber, 1);
      await saveWorkbook(workbook);

      return { status: 200, body: { ok: true } };
    });

    return res.status(result.status).json(result.body);
  } catch (error) {
    console.error('Cannot delete transaction:', error);
    return sendError(res, 'Cannot delete transaction');
  }
});

app.get('/api/download-excel', requireAuth, async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      await queueWrite(async () => {
        const { workbook } = await openWorkbook();
        await saveWorkbook(workbook);
      });
    }

    return res.download(EXCEL_FILE, 'glori-budget.xlsx');
  } catch (error) {
    console.error('Cannot download Excel:', error);
    return sendError(res, 'Cannot download Excel');
  }
});

app.use('/api', (req, res) => {
  return res.status(404).json({ ok: false, error: 'API route not found' });
});

app.use((req, res) => {
  if (isPageRequest(req)) {
    return res.redirect('/');
  }
  return res.status(404).send('Page not found');
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Glori Budget Manager running on port ${PORT}`);
  console.log(`NODE_ENV=${process.env.NODE_ENV || 'development'}`);
  console.log(`PUBLIC_DIR=${PUBLIC_DIR}`);
  console.log(`DATA_DIR=${DATA_DIR}`);
  console.log(`UPLOAD_DIR=${UPLOAD_DIR}`);
  console.log(`EXCEL_FILE=${EXCEL_FILE}`);
});