require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');
const http = require('http');
const crypto = require('crypto');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const multer = require('multer');
const XLSX = require('xlsx');
const initSqlJs = require('sql.js');
let nodemailer = null;
try { nodemailer = require('nodemailer'); } catch (e) {}

const app = express();
const PORT = process.env.PORT || 4000;
const JWT_SECRET = process.env.JWT_SECRET || 'nova_super_secret_jwt_key_2026_production';
const MASTER_ADMIN_SECRET = process.env.MASTER_ADMIN_SECRET || 'NOVA-MASTER-KEY-9988';
const DB_FILE = path.join(__dirname, 'novacrm_production.sqlite');

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
const upload = multer({ dest: uploadDir, limits: { fileSize: 50 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use('/uploads', express.static(uploadDir));
app.use(express.static(__dirname));

const uid = (prefix = 'id') => `${prefix}_${crypto.randomBytes(6).toString('hex')}`;
const addDaysISO = (days, baseDate = new Date()) => {
  const d = new Date(baseDate);
  d.setDate(d.getDate() + Number(days));
  return d.toISOString();
};

// Detect Local Network IPv4 Address so other computers/phones on Wi-Fi/LAN can open the CRM
function getLocalNetworkUrls() {
  const nets = os.networkInterfaces();
  const urls = [];
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === 'IPv4' && !net.internal) {
        urls.push(`http://${net.address}:${PORT}`);
      }
    }
  }
  return urls;
}

// Binary/Text Downloader with Redirect Support for Google Sheets (.xlsx & .csv)
function fetchUrlBuffer(targetUrl, maxRedirects = 8) {
  return new Promise((resolve, reject) => {
    const client = targetUrl.startsWith('https') ? https : http;
    const req = client.get(targetUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36',
        'Accept': '*/*'
      }
    }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location) {
        if (maxRedirects <= 0) return reject(new Error('Too many redirects'));
        const nextUrl = new URL(res.headers.location, targetUrl).toString();
        return resolve(fetchUrlBuffer(nextUrl, maxRedirects - 1));
      }
      if (res.statusCode < 200 || res.statusCode >= 300) {
        return reject(new Error(`HTTP Status ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve({
        buffer: Buffer.concat(chunks),
        contentType: res.headers['content-type'] || ''
      }));
    });
    req.on('error', err => reject(err));
    req.setTimeout(20000, () => { req.destroy(); reject(new Error('Connection timed out')); });
  });
}

let sqlDb = null;

function persistDb() {
  if (!sqlDb) return;
  try {
    const data = sqlDb.export();
    fs.writeFileSync(DB_FILE, Buffer.from(data));
  } catch (e) {
    console.error('DB Persist Error:', e.message);
  }
}

function queryAll(sql, params = []) {
  try {
    const stmt = sqlDb.prepare(sql);
    const cleanParams = params.map(p => (p === undefined ? null : p));
    stmt.bind(cleanParams);
    const rows = [];
    while (stmt.step()) rows.push(stmt.getAsObject());
    stmt.free();
    return rows;
  } catch (e) {
    console.error('SQL Query Error:', e.message, sql);
    return [];
  }
}

function queryOne(sql, params = []) {
  const rows = queryAll(sql, params);
  return rows[0] || null;
}

function runSql(sql, params = [], skipSave = false) {
  const cleanParams = params.map(p => (p === undefined ? null : p));
  sqlDb.run(sql, cleanParams);
  if (!skipSave) persistDb();
}

const SCHEMA_SQL = `
CREATE TABLE IF NOT EXISTS tenants (
  id TEXT PRIMARY KEY,
  company_name TEXT NOT NULL,
  contact_name TEXT NOT NULL,
  email TEXT UNIQUE NOT NULL,
  phone TEXT,
  plan_type TEXT NOT NULL DEFAULT 'trial_7d',
  status TEXT NOT NULL DEFAULT 'active',
  access_start TEXT NOT NULL,
  access_end TEXT NOT NULL,
  public_share_url TEXT DEFAULT '',
  smtp_email TEXT DEFAULT '',
  smtp_app_password TEXT DEFAULT '',
  max_users INTEGER DEFAULT 25,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS users (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  username TEXT UNIQUE NOT NULL,
  email TEXT UNIQUE NOT NULL,
  password_hash TEXT,
  plain_pass_hint TEXT DEFAULT '',
  google_id TEXT,
  full_name TEXT NOT NULL,
  role TEXT NOT NULL DEFAULT 'admin',
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS leads (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  name TEXT NOT NULL,
  company TEXT,
  email TEXT,
  phone TEXT,
  address TEXT DEFAULT '',
  links TEXT DEFAULT '',
  stage TEXT NOT NULL DEFAULT 'New',
  interest_level TEXT DEFAULT 'Mid / Follow-up',
  talk_summary TEXT DEFAULT '',
  followup_required TEXT DEFAULT 'Required',
  followup_date TEXT DEFAULT '',
  notes_history_json TEXT DEFAULT '[]',
  extra_data_json TEXT DEFAULT '{}',
  qualified_by TEXT DEFAULT '',
  deal_value REAL DEFAULT 0,
  source TEXT DEFAULT 'Manual',
  priority TEXT DEFAULT 'medium',
  owner_name TEXT DEFAULT 'Unassigned',
  notes TEXT DEFAULT '',
  sort_order INTEGER DEFAULT 0,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_leads_tenant_stage ON leads(tenant_id, stage);
CREATE TABLE IF NOT EXISTS calendar_events (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  title TEXT NOT NULL,
  event_type TEXT DEFAULT 'followup',
  start_at TEXT NOT NULL,
  end_at TEXT,
  lead_id TEXT DEFAULT '',
  notes TEXT DEFAULT '',
  status TEXT DEFAULT 'Scheduled',
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS employees (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  emp_code TEXT NOT NULL,
  full_name TEXT NOT NULL,
  designation TEXT NOT NULL,
  department TEXT NOT NULL,
  role TEXT DEFAULT 'agent',
  login_username TEXT DEFAULT '',
  login_password_hint TEXT DEFAULT '',
  email TEXT NOT NULL,
  phone TEXT NOT NULL,
  joining_date TEXT,
  monthly_salary REAL DEFAULT 0,
  location TEXT,
  education TEXT,
  emergency_contact TEXT,
  address TEXT,
  bank_details TEXT,
  background_summary TEXT,
  previous_companies_json TEXT DEFAULT '[]',
  documents_json TEXT DEFAULT '[]',
  status TEXT DEFAULT 'Active',
  created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS messages (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  lead_name TEXT DEFAULT '',
  recipient TEXT NOT NULL,
  channel TEXT NOT NULL,
  subject TEXT DEFAULT '',
  body TEXT NOT NULL,
  sender_name TEXT DEFAULT '',
  status TEXT DEFAULT 'Sent',
  sent_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS subscription_payments (
  id TEXT PRIMARY KEY,
  tenant_id TEXT NOT NULL,
  amount REAL NOT NULL,
  currency TEXT DEFAULT 'INR',
  plan_granted TEXT NOT NULL,
  payment_gateway TEXT NOT NULL,
  transaction_ref TEXT,
  valid_until TEXT NOT NULL,
  paid_at TEXT DEFAULT CURRENT_TIMESTAMP
);
`;

function ensureAllColumnsAndTables() {
  const ensureCol = (table, col, def) => {
    const existing = queryAll(`PRAGMA table_info(${table})`).map(c => c.name);
    if (!existing.includes(col)) {
      try { sqlDb.run(`ALTER TABLE ${table} ADD COLUMN ${col} ${def}`); } catch (e) {}
    }
  };

  ensureCol('tenants', 'public_share_url', "TEXT DEFAULT ''");
  ensureCol('tenants', 'smtp_email', "TEXT DEFAULT ''");
  ensureCol('tenants', 'smtp_app_password', "TEXT DEFAULT ''");
  ensureCol('users', 'plain_pass_hint', "TEXT DEFAULT ''");
  ensureCol('leads', 'address', "TEXT DEFAULT ''");
  ensureCol('leads', 'links', "TEXT DEFAULT ''");
  ensureCol('leads', 'interest_level', "TEXT DEFAULT 'Mid / Follow-up'");
  ensureCol('leads', 'talk_summary', "TEXT DEFAULT ''");
  ensureCol('leads', 'followup_required', "TEXT DEFAULT 'Required'");
  ensureCol('leads', 'followup_date', "TEXT DEFAULT ''");
  ensureCol('leads', 'notes_history_json', "TEXT DEFAULT '[]'");
  ensureCol('leads', 'extra_data_json', "TEXT DEFAULT '{}'");
  ensureCol('leads', 'qualified_by', "TEXT DEFAULT ''");
  ensureCol('leads', 'owner_name', "TEXT DEFAULT 'Admin'");
  ensureCol('leads', 'sort_order', "INTEGER DEFAULT 0");

  ensureCol('employees', 'role', "TEXT DEFAULT 'agent'");
  ensureCol('employees', 'login_username', "TEXT DEFAULT ''");
  ensureCol('employees', 'login_password_hint', "TEXT DEFAULT ''");
  ensureCol('employees', 'previous_companies_json', "TEXT DEFAULT '[]'");
  ensureCol('employees', 'documents_json', "TEXT DEFAULT '[]'");

  ensureCol('calendar_events', 'lead_id', "TEXT DEFAULT ''");
  ensureCol('calendar_events', 'notes', "TEXT DEFAULT ''");
  ensureCol('calendar_events', 'status', "TEXT DEFAULT 'Scheduled'");

  ensureCol('messages', 'lead_name', "TEXT DEFAULT ''");
  ensureCol('messages', 'sender_name', "TEXT DEFAULT ''");
  persistDb();
}

function authMiddleware(req, res, next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (!token) return res.status(401).json({ error: 'Authentication required' });

  try {
    const decoded = jwt.verify(token, JWT_SECRET);
    const tenant = queryOne('SELECT * FROM tenants WHERE id = ?', [decoded.tenantId]);
    if (!tenant) return res.status(401).json({ error: 'Workspace not found' });

    const msLeft = new Date(tenant.access_end).getTime() - Date.now();
    const daysLeft = Math.ceil(msLeft / 86400000);

    req.user = decoded;
    req.tenant = { ...tenant, daysLeft };
    next();
  } catch (e) {
    return res.status(401).json({ error: 'Invalid or expired session token' });
  }
}

function masterAdminMiddleware(req, res, next) {
  const key = req.headers['x-master-secret'];
  if (!key || key !== MASTER_ADMIN_SECRET) {
    return res.status(403).json({ error: 'Unauthorized Master Admin secret key' });
  }
  next();
}

/* ================= MASTER ADMIN ROUTES ================= */
app.get('/api/master/tenants', masterAdminMiddleware, (req, res) => {
  const rows = queryAll(`
    SELECT t.*,
      (SELECT COUNT(*) FROM leads l WHERE l.tenant_id = t.id) AS lead_count,
      (SELECT COUNT(*) FROM employees e WHERE e.tenant_id = t.id) AS emp_count,
      (SELECT username FROM users u WHERE u.tenant_id = t.id LIMIT 1) AS login_id
    FROM tenants t
    ORDER BY t.created_at DESC
  `);
  res.json({ tenants: rows });
});

app.post('/api/master/provision-client', masterAdminMiddleware, (req, res) => {
  const { companyName, contactName, email, phone, username, password, accessDays, planType, maxUsers } = req.body;
  if (!companyName || !email || !username || !password || !accessDays) {
    return res.status(400).json({ error: 'Missing required fields' });
  }
  const tenantId = uid('tnt');
  const userId = uid('usr');
  const now = new Date().toISOString();
  const accessEnd = addDaysISO(Number(accessDays));
  const hash = bcrypt.hashSync(password, 10);

  try {
    runSql(`
      INSERT INTO tenants (id, company_name, contact_name, email, phone, plan_type, status, access_start, access_end, max_users, created_at)
      VALUES (?, ?, ?, ?, ?, ?, 'active', ?, ?, ?, ?)
    `, [tenantId, companyName, contactName, email.toLowerCase(), phone || '', planType || 'custom_access', now, accessEnd, maxUsers || 25, now], true);

    runSql(`
      INSERT INTO users (id, tenant_id, username, email, password_hash, plain_pass_hint, full_name, role, created_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, 'admin', ?)
    `, [userId, tenantId, username.toLowerCase(), email.toLowerCase(), hash, password, contactName, now], true);

    runSql(`
      INSERT INTO employees (id, tenant_id, emp_code, full_name, designation, department, role, login_username, login_password_hint, email, phone, joining_date, status, created_at)
      VALUES (?, ?, 'EMP-001', ?, 'Workspace Admin', 'Management & Sales', 'admin', ?, ?, ?, ?, ?, 'Active', ?)
    `, [uid('emp'), tenantId, contactName, username.toLowerCase(), password, email.toLowerCase(), phone || '', now.slice(0, 10), now], false);

    res.json({ success: true, tenantId, username, accessEnd });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.post('/api/master/update-access', masterAdminMiddleware, (req, res) => {
  const { tenantId, addDays, planType, status } = req.body;
  const tenant = queryOne('SELECT * FROM tenants WHERE id = ?', [tenantId]);
  if (!tenant) return res.status(404).json({ error: 'Client not found' });

  const base = new Date(tenant.access_end) > new Date() ? new Date(tenant.access_end) : new Date();
  const newEnd = addDays ? addDaysISO(Number(addDays), base) : tenant.access_end;

  runSql(`UPDATE tenants SET access_end = ?, plan_type = ?, status = ? WHERE id = ?`,
    [newEnd, planType || tenant.plan_type, status || tenant.status, tenantId]);

  res.json({ success: true, accessEnd: newEnd });
});

/* ================= CLIENT & TEAM MEMBER AUTH ================= */
app.post('/api/auth/login', (req, res) => {
  const { loginId, password } = req.body;
  if (!loginId || !password) return res.status(400).json({ error: 'Enter User ID/Email and Password' });

  const clean = loginId.toLowerCase().trim();
  const user = queryOne('SELECT * FROM users WHERE username = ? OR email = ?', [clean, clean]);
  if (!user || !user.password_hash || !bcrypt.compareSync(password, user.password_hash)) {
    return res.status(401).json({ error: 'Invalid User ID or Password' });
  }

  const tenant = queryOne('SELECT * FROM tenants WHERE id = ?', [user.tenant_id]);
  const token = jwt.sign({ userId: user.id, tenantId: tenant.id, role: user.role, name: user.full_name }, JWT_SECRET, { expiresIn: '7d' });

  res.json({ token, user: { id: user.id, name: user.full_name, email: user.email, role: user.role }, tenant });
});

app.post('/api/auth/google-or-signup', (req, res) => {
  const { companyName, fullName, email, phone, password, googleId, planChoice } = req.body;
  if (!email || !fullName) return res.status(400).json({ error: 'Name and email are required' });

  const cleanEmail = email.toLowerCase().trim();
  let user = queryOne('SELECT * FROM users WHERE email = ?', [cleanEmail]);

  if (!user) {
    const tenantId = uid('tnt');
    const userId = uid('usr');
    const now = new Date().toISOString();
    const accessEnd = addDaysISO(7);
    const hash = password ? bcrypt.hashSync(password, 10) : null;

    runSql(`
      INSERT INTO tenants (id, company_name, contact_name, email, phone, plan_type, status, access_start, access_end, max_users, created_at)
      VALUES (?, ?, ?, ?, ?, 'trial_7d', 'active', ?, ?, 25, ?)
    `, [tenantId, companyName || `${fullName}'s Workspace`, fullName, cleanEmail, phone || '', now, accessEnd, now], true);

    runSql(`
      INSERT INTO users (id, tenant_id, username, email, password_hash, plain_pass_hint, google_id, full_name, role, created_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'admin', ?)
    `, [userId, tenantId, cleanEmail, cleanEmail, hash, password || '', googleId || null, fullName, now], true);

    runSql(`
      INSERT INTO employees (id, tenant_id, emp_code, full_name, designation, department, role, login_username, email, phone, joining_date, status, created_at)
      VALUES (?, ?, 'EMP-001', ?, 'Founder / Admin', 'Management & Sales', 'admin', ?, ?, ?, ?, 'Active', ?)
    `, [uid('emp'), tenantId, fullName, cleanEmail, cleanEmail, phone || '', now.slice(0, 10), now], false);

    user = queryOne('SELECT * FROM users WHERE id = ?', [userId]);
  }

  const tenant = queryOne('SELECT * FROM tenants WHERE id = ?', [user.tenant_id]);
  const token = jwt.sign({ userId: user.id, tenantId: tenant.id, role: user.role, name: user.full_name }, JWT_SECRET, { expiresIn: '7d' });

  res.json({ token, user: { id: user.id, name: user.full_name, email: user.email, role: user.role }, tenant, requirePayment: planChoice === 'annual_1yr' });
});

/* ================= WORKSPACE STATUS & NETWORK SHARE URLS ================= */
app.get('/api/workspace/status', authMiddleware, (req, res) => {
  const existingEmp = queryOne('SELECT COUNT(*) as c FROM employees WHERE tenant_id = ?', [req.tenant.id]);
  if (!existingEmp || existingEmp.c === 0) {
    runSql(`
      INSERT INTO employees (id, tenant_id, emp_code, full_name, designation, department, role, login_username, email, phone, joining_date, status, created_at)
      VALUES (?, ?, 'EMP-001', ?, 'Workspace Admin', 'Sales & Management', 'admin', ?, ?, ?, ?, 'Active', ?)
    `, [uid('emp'), req.tenant.id, req.user.name, req.tenant.email || 'admin', req.tenant.email || '', req.tenant.phone || '', new Date().toISOString().slice(0, 10), new Date().toISOString()]);
  }

  const leadCount = (queryOne('SELECT COUNT(*) as c FROM leads WHERE tenant_id = ?', [req.tenant.id]) || { c: 0 }).c;
  const totalValue = (queryOne('SELECT SUM(deal_value) as s FROM leads WHERE tenant_id = ?', [req.tenant.id]) || { s: 0 }).s || 0;
  const wonValue = (queryOne("SELECT SUM(deal_value) as s FROM leads WHERE tenant_id = ? AND stage = 'Won'", [req.tenant.id]) || { s: 0 }).s || 0;
  const wonCount = (queryOne("SELECT COUNT(*) as c FROM leads WHERE tenant_id = ? AND stage = 'Won'", [req.tenant.id]) || { c: 0 }).c || 0;
  const qualifiedCount = (queryOne("SELECT COUNT(*) as c FROM leads WHERE tenant_id = ? AND stage IN ('Qualified','Proposal Sent','Negotiation','Won')", [req.tenant.id]) || { c: 0 }).c || 0;
  const empCount = (queryOne('SELECT COUNT(*) as c FROM employees WHERE tenant_id = ?', [req.tenant.id]) || { c: 0 }).c;

  const followups = queryAll(`
    SELECT id, company, name, phone, stage, interest_level, talk_summary, followup_date, owner_name
    FROM leads
    WHERE tenant_id = ? AND followup_required = 'Required' AND followup_date != '' AND followup_date IS NOT NULL
    ORDER BY followup_date ASC
  `, [req.tenant.id]);

  const stageBreakdown = queryAll('SELECT stage, COUNT(*) as count, SUM(deal_value) as total FROM leads WHERE tenant_id = ? GROUP BY stage', [req.tenant.id]);
  const sourceBreakdown = queryAll('SELECT source, COUNT(*) as count FROM leads WHERE tenant_id = ? GROUP BY source', [req.tenant.id]);
  // IMPORTANT: Order by ROWID ASC so the FIRST row of the uploaded sheet appears FIRST!
  const recentLeads = queryAll('SELECT * FROM leads WHERE tenant_id = ? ORDER BY ROWID ASC LIMIT 25', [req.tenant.id]);
  const teamList = queryAll('SELECT id, full_name, designation, role, email, phone, login_username FROM employees WHERE tenant_id = ? ORDER BY full_name ASC', [req.tenant.id]);

  res.json({
    tenant: req.tenant,
    user: req.user,
    networkUrls: getLocalNetworkUrls(),
    teamList,
    followups,
    stats: {
      leadCount, totalValue, wonValue, wonCount, qualifiedCount, empCount,
      eventCount: followups.length, stageBreakdown, sourceBreakdown, recentLeads
    }
  });
});

// Save Custom Share URL or SMTP Email Configuration
app.post('/api/workspace/share-settings', authMiddleware, (req, res) => {
  const { public_share_url, smtp_email, smtp_app_password } = req.body;
  runSql(`
    UPDATE tenants SET
      public_share_url = COALESCE(?, public_share_url),
      smtp_email = COALESCE(?, smtp_email),
      smtp_app_password = COALESCE(?, smtp_app_password)
    WHERE id = ?
  `, [
    public_share_url !== undefined ? public_share_url.trim() : null,
    smtp_email !== undefined ? smtp_email.trim() : null,
    smtp_app_password !== undefined ? smtp_app_password.trim() : null,
    req.tenant.id
  ]);
  res.json({ success: true });
});

/* ================= LEADS CRUD (ORDERED BY EXACT SHEET ORDER: ROWID ASC) ================= */
app.get('/api/leads', authMiddleware, (req, res) => {
  const page = Math.max(1, parseInt(req.query.page || '1', 10));
  const limit = Math.min(500, Math.max(10, parseInt(req.query.limit || '50', 10)));
  const offset = (page - 1) * limit;
  const search = (req.query.q || '').trim();
  const stage = (req.query.stage || '').trim();
  const owner = (req.query.owner || '').trim();

  let where = 'WHERE tenant_id = ?';
  const params = [req.tenant.id];

  if (stage) { where += ' AND stage = ?'; params.push(stage); }
  if (owner) { where += ' AND owner_name = ?'; params.push(owner); }
  if (search) {
    where += ' AND (company LIKE ? OR name LIKE ? OR phone LIKE ? OR email LIKE ? OR address LIKE ? OR talk_summary LIKE ? OR owner_name LIKE ?)';
    const q = `%${search}%`;
    params.push(q, q, q, q, q, q, q);
  }

  const total = (queryOne(`SELECT COUNT(*) as c FROM leads ${where}`, params) || { c: 0 }).c;
  // Ordered by ROWID ASC so Row 1 of Excel/Google Sheet is always shown at #1!
  const rows = queryAll(`SELECT * FROM leads ${where} ORDER BY ROWID ASC LIMIT ? OFFSET ?`, [...params, limit, offset]);

  res.json({ leads: rows, total, page, totalPages: Math.ceil(total / limit) || 1 });
});

app.get('/api/leads/:id', authMiddleware, (req, res) => {
  const lead = queryOne('SELECT * FROM leads WHERE id = ? AND tenant_id = ?', [req.params.id, req.tenant.id]);
  if (!lead) return res.status(404).json({ error: 'Lead not found' });
  res.json({ lead });
});

app.post('/api/leads', authMiddleware, (req, res) => {
  const {
    id, name, company, email, phone, address, links, stage, interest_level,
    talk_summary, followup_required, followup_date, deal_value, source, priority, owner_name, notes
  } = req.body;

  if (!company && !name) return res.status(400).json({ error: 'Company Name is required' });
  const now = new Date().toISOString();
  const assignedOwner = owner_name || req.user.name;
  const newId = id || uid('ld');
  const activeStage = stage || 'New';
  const activeInterest = interest_level || 'Mid / Follow-up';
  const activeTalk = talk_summary || notes || '';
  const fReq = followup_required || (followup_date ? 'Required' : 'Not Required');
  const fDate = followup_date || '';

  const initialHistory = activeTalk ? [{
    id: uid('nt'),
    text: activeTalk,
    stage: activeStage,
    interest: activeInterest,
    author: assignedOwner,
    at: now
  }] : [];

  runSql(`
    INSERT INTO leads (
      id, tenant_id, name, company, email, phone, address, links, stage,
      interest_level, talk_summary, followup_required, followup_date, notes_history_json,
      extra_data_json, qualified_by, deal_value, source, priority, owner_name, notes, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '{}', ?, ?, ?, ?, ?, ?, ?)
  `, [
    newId, req.tenant.id, name || company, company || name, email || '', phone || '',
    address || '', links || '', activeStage, activeInterest,
    activeTalk, fReq, fDate, JSON.stringify(initialHistory),
    ['Qualified', 'Proposal Sent', 'Negotiation', 'Won'].includes(activeStage) ? assignedOwner : '',
    Number(deal_value || 0), source || 'Manual',
    priority || 'medium', assignedOwner, activeTalk, now
  ]);

  if (fReq === 'Required' && fDate) {
    runSql(`
      INSERT INTO calendar_events (id, tenant_id, title, event_type, start_at, end_at, lead_id, notes, created_at)
      VALUES (?, ?, ?, 'followup', ?, ?, ?, ?, ?)
    `, [uid('evt'), req.tenant.id, `Follow-up: ${company || name}`, fDate, fDate, newId, activeTalk, now]);
  }

  return res.json({ success: true, id: newId });
});

app.post('/api/leads/:id/quick-update', authMiddleware, (req, res) => {
  const lead = queryOne('SELECT * FROM leads WHERE id = ? AND tenant_id = ?', [req.params.id, req.tenant.id]);
  if (!lead) return res.status(404).json({ error: 'Lead not found' });

  const body = req.body || {};
  const nextCompany = body.company !== undefined ? String(body.company) : String(lead.company || '');
  const nextName = body.name !== undefined ? String(body.name) : String(lead.name || '');
  const nextPhone = body.phone !== undefined ? String(body.phone) : String(lead.phone || '');
  const nextEmail = body.email !== undefined ? String(body.email) : String(lead.email || '');
  const nextStage = body.stage !== undefined ? String(body.stage) : String(lead.stage || 'New');
  const nextInterest = body.interest_level !== undefined ? String(body.interest_level) : String(lead.interest_level || 'Mid / Follow-up');
  const nextOwner = body.owner_name !== undefined ? String(body.owner_name) : String(lead.owner_name || req.user.name);
  const nextAddress = body.address !== undefined ? String(body.address) : String(lead.address || '');
  const nextLinks = body.links !== undefined ? String(body.links) : String(lead.links || '');
  const nextDealVal = body.deal_value !== undefined ? Number(body.deal_value || 0) : Number(lead.deal_value || 0);

  let nextFollowDate = body.followup_date !== undefined ? String(body.followup_date) : String(lead.followup_date || '');
  let nextFollowReq = body.followup_required !== undefined ? String(body.followup_required) : String(lead.followup_required || 'Required');
  if (body.followup_date && body.followup_required === undefined) {
    nextFollowReq = 'Required';
  }
  if (nextFollowReq === 'Not Required' && body.followup_date === undefined) {
    nextFollowDate = '';
  }

  let history = [];
  try { history = JSON.parse(lead.notes_history_json || '[]'); } catch (e) { history = []; }
  const now = new Date().toISOString();

  if (body.new_note && String(body.new_note).trim() !== '') {
    history.unshift({
      id: uid('nt'),
      text: String(body.new_note).trim(),
      stage: nextStage,
      interest: nextInterest,
      author: req.user.name,
      at: now
    });
  } else if (body.stage && body.stage !== lead.stage) {
    history.unshift({
      id: uid('nt'),
      text: `Stage changed from "${lead.stage}" to "${nextStage}"`,
      stage: nextStage,
      interest: nextInterest,
      author: req.user.name,
      at: now
    });
  }

  const nextTalk = body.new_note !== undefined
    ? String(body.new_note).trim()
    : (body.talk_summary !== undefined ? String(body.talk_summary) : String(lead.talk_summary || ''));

  let qualBy = String(lead.qualified_by || '');
  if (['Qualified', 'Proposal Sent', 'Negotiation', 'Won'].includes(nextStage)) {
    qualBy = nextOwner || req.user.name;
  }

  runSql(`
    UPDATE leads SET
      company = ?, name = ?, phone = ?, email = ?,
      stage = ?, interest_level = ?, talk_summary = ?,
      followup_required = ?, followup_date = ?,
      notes_history_json = ?, qualified_by = ?,
      address = ?, links = ?, owner_name = ?, deal_value = ?
    WHERE id = ? AND tenant_id = ?
  `, [
    nextCompany, nextName, nextPhone, nextEmail,
    nextStage, nextInterest, nextTalk,
    nextFollowReq, nextFollowDate,
    JSON.stringify(history), qualBy,
    nextAddress, nextLinks, nextOwner, nextDealVal,
    lead.id, req.tenant.id
  ]);

  runSql('DELETE FROM calendar_events WHERE lead_id = ? AND tenant_id = ?', [lead.id, req.tenant.id], true);
  if (nextFollowReq === 'Required' && nextFollowDate) {
    runSql(`
      INSERT INTO calendar_events (id, tenant_id, title, event_type, start_at, end_at, lead_id, notes, status, created_at)
      VALUES (?, ?, ?, 'followup', ?, ?, ?, ?, 'Scheduled', ?)
    `, [
      uid('evt'), req.tenant.id, `📞 Follow-up: ${nextCompany || nextName}`,
      nextFollowDate, nextFollowDate, lead.id, nextTalk || `Assigned: ${nextOwner}`, now
    ], false);
  } else {
    persistDb();
  }

  const updated = queryOne('SELECT * FROM leads WHERE id = ? AND tenant_id = ?', [lead.id, req.tenant.id]);
  res.json({ success: true, lead: updated });
});

app.post('/api/leads/bulk-assign', authMiddleware, (req, res) => {
  const { ids, owner_name } = req.body;
  if (!Array.isArray(ids) || !ids.length || !owner_name) {
    return res.status(400).json({ error: 'Select leads and a team member' });
  }
  sqlDb.run('BEGIN TRANSACTION');
  for (const id of ids) {
    sqlDb.run('UPDATE leads SET owner_name = ? WHERE id = ? AND tenant_id = ?', [owner_name, id, req.tenant.id]);
  }
  sqlDb.run('COMMIT');
  persistDb();
  res.json({ success: true, count: ids.length });
});

/* ================= MULTI-SHEET GOOGLE SHEET FETCHER (RETURNS ALL SHEETS/TABS!) ================= */
app.post('/api/leads/fetch-google-sheet', authMiddleware, async (req, res) => {
  const { url } = req.body;
  if (!url) return res.status(400).json({ error: 'Please paste a Google Sheet link' });

  try {
    const cleanUrl = url.trim();
    const idMatch = cleanUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const gidMatch = cleanUrl.match(/[#&?]gid=([0-9]+)/);
    const gid = gidMatch ? gidMatch[1] : '0';

    if (!idMatch || !idMatch[1]) {
      return res.status(400).json({ error: 'Invalid Google Sheet URL.' });
    }

    const sheetId = idMatch[1];
    // Try full .xlsx export first so ALL sheets/tabs in the Google Sheet are detected!
    const endpoints = [
      { type: 'xlsx', url: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx` },
      { type: 'csv', url: `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&gid=${gid}` },
      { type: 'csv', url: `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}` }
    ];

    for (const ep of endpoints) {
      try {
        const { buffer, contentType } = await fetchUrlBuffer(ep.url);
        const sampleStr = buffer.slice(0, 500).toString('utf8');
        if (contentType.includes('text/html') || sampleStr.includes('accounts.google.com') || sampleStr.trim().startsWith('<!DOCTYPE html>')) {
          continue;
        }
        if (ep.type === 'xlsx') {
          return res.json({ success: true, base64Xlsx: buffer.toString('base64') });
        } else {
          const csvText = buffer.toString('utf8');
          if (csvText && csvText.trim().length > 0) {
            return res.json({ success: true, csvText });
          }
        }
      } catch (innerErr) {}
    }

    throw new Error('This Google Sheet is Restricted! In Google Sheets, click "Share" (top right) -> change General Access to "Anyone with the link" -> then click Analyze.');
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// Preserves Exact Sheet Order (Row 1 first -> Row N last)
app.post('/api/leads/import-mapped', authMiddleware, (req, res) => {
  const { leads, source, assignTo } = req.body;
  if (!Array.isArray(leads) || leads.length === 0) return res.status(400).json({ error: 'No leads to import' });

  const baseTime = Date.now();
  const defaultOwner = assignTo || req.user.name;
  let importedCount = 0;

  sqlDb.run('BEGIN TRANSACTION');
  for (let i = 0; i < leads.length; i++) {
    const item = leads[i];
    const company = String(item.company || item.name || '').trim();
    const name = String(item.name || item.company || '').trim();
    const phone = String(item.phone || '').trim();
    const email = String(item.email || '').trim();
    const address = String(item.address || '').trim();
    const links = String(item.links || '').trim();
    const stage = String(item.stage || 'New').trim();
    const interest = String(item.interest_level || 'Mid / Follow-up').trim();
    const talk = String(item.talk_summary || item.notes || '').trim();
    const fReq = String(item.followup_required || (item.followup_date ? 'Required' : 'Not Required')).trim();
    const fDate = String(item.followup_date || '').trim();
    const val = Number(String(item.deal_value || '0').replace(/[^0-9.]/g, '')) || 0;
    const owner = String(item.owner_name || defaultOwner).trim();
    const extraJson = JSON.stringify(item.extra_data || {});

    if (!company && !name && !phone && !email) continue;

    const rowTime = new Date(baseTime + i).toISOString();
    const history = talk ? [{ id: uid('nt'), text: talk, stage, interest, author: owner, at: rowTime }] : [];

    sqlDb.run(`
      INSERT INTO leads (
        id, tenant_id, name, company, email, phone, address, links, stage,
        interest_level, talk_summary, followup_required, followup_date, notes_history_json,
        extra_data_json, qualified_by, deal_value, source, owner_name, notes, sort_order, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?, ?, ?)
    `, [
      uid('ld'), req.tenant.id, name || company, company || name, email, phone,
      address, links, stage, interest, talk, fReq, fDate, JSON.stringify(history),
      extraJson, val, source || 'Import', owner, talk, i + 1, rowTime
    ]);
    importedCount++;
  }
  sqlDb.run('COMMIT');
  persistDb();

  res.json({ success: true, importedCount });
});

app.delete('/api/leads/:id', authMiddleware, (req, res) => {
  runSql('DELETE FROM calendar_events WHERE lead_id = ? AND tenant_id = ?', [req.params.id, req.tenant.id], true);
  runSql('DELETE FROM leads WHERE id = ? AND tenant_id = ?', [req.params.id, req.tenant.id], false);
  res.json({ success: true });
});

app.post('/api/leads/bulk-delete', authMiddleware, (req, res) => {
  const { ids, deleteAll } = req.body;
  if (deleteAll === true) {
    runSql('DELETE FROM calendar_events WHERE tenant_id = ? AND lead_id != ""', [req.tenant.id], true);
    runSql('DELETE FROM leads WHERE tenant_id = ?', [req.tenant.id], false);
    return res.json({ success: true });
  }
  if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: 'No leads selected' });
  sqlDb.run('BEGIN TRANSACTION');
  for (const id of ids) {
    sqlDb.run('DELETE FROM calendar_events WHERE lead_id = ? AND tenant_id = ?', [id, req.tenant.id]);
    sqlDb.run('DELETE FROM leads WHERE id = ? AND tenant_id = ?', [id, req.tenant.id]);
  }
  sqlDb.run('COMMIT');
  persistDb();
  res.json({ success: true, deletedCount: ids.length });
});

/* ================= CALENDAR & FOLLOW-UPS API ================= */
app.get('/api/calendar', authMiddleware, (req, res) => {
  const calEvents = queryAll('SELECT * FROM calendar_events WHERE tenant_id = ? ORDER BY start_at ASC', [req.tenant.id]);
  const leadFollowups = queryAll(`
    SELECT id, company, name, phone, stage, interest_level, talk_summary, followup_date, owner_name
    FROM leads
    WHERE tenant_id = ? AND followup_required = 'Required' AND followup_date != '' AND followup_date IS NOT NULL
    ORDER BY followup_date ASC
  `, [req.tenant.id]);

  res.json({ events: calEvents, leadFollowups });
});

app.post('/api/calendar', authMiddleware, (req, res) => {
  const { title, event_type, start_at, end_at, notes } = req.body;
  if (!title || !start_at) return res.status(400).json({ error: 'Title and Date required' });
  const id = uid('evt');
  runSql(`
    INSERT INTO calendar_events (id, tenant_id, title, event_type, start_at, end_at, notes, status, created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, 'Scheduled', ?)
  `, [id, req.tenant.id, title, event_type || 'followup', start_at, end_at || start_at, notes || '', new Date().toISOString()]);
  res.json({ success: true, id });
});

app.delete('/api/calendar/:id', authMiddleware, (req, res) => {
  runSql('DELETE FROM calendar_events WHERE id = ? AND tenant_id = ?', [req.params.id, req.tenant.id]);
  res.json({ success: true });
});

/* ================= MESSAGES & DIRECT EMAIL INVITE SENDER ================= */
app.get('/api/messages', authMiddleware, (req, res) => {
  const messages = queryAll('SELECT * FROM messages WHERE tenant_id = ? ORDER BY sent_at DESC LIMIT 100', [req.tenant.id]);
  res.json({ messages });
});

app.post('/api/messages', authMiddleware, async (req, res) => {
  const { lead_name, recipient, channel, subject, body } = req.body;
  if (!recipient || !body) return res.status(400).json({ error: 'Recipient and message body required' });
  const id = uid('msg');

  let emailSentViaSmtp = false;
  if (channel === 'Email' && nodemailer && (req.tenant.smtp_email || process.env.SMTP_USER)) {
    try {
      const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
          user: req.tenant.smtp_email || process.env.SMTP_USER,
          pass: req.tenant.smtp_app_password || process.env.SMTP_PASS
        }
      });
      await transporter.sendMail({
        from: `"${req.tenant.company_name}" <${req.tenant.smtp_email || process.env.SMTP_USER}>`,
        to: recipient,
        subject: subject || `Invitation / Update from ${req.tenant.company_name}`,
        text: body
      });
      emailSentViaSmtp = true;
    } catch (e) {
      console.error('SMTP Send Error:', e.message);
    }
  }

  runSql(`
    INSERT INTO messages (id, tenant_id, lead_name, recipient, channel, subject, body, sender_name, status, sent_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `, [id, req.tenant.id, lead_name || '', recipient, channel || 'WhatsApp', subject || '', body, req.user.name, emailSentViaSmtp ? 'Sent via SMTP' : 'Logged', new Date().toISOString()]);

  res.json({ success: true, id, emailSentViaSmtp });
});

/* ================= TEAM + INVITE LOGIN ACCESS + HRMS PROFILE ================= */
app.get('/api/team', authMiddleware, (req, res) => {
  const rows = queryAll('SELECT * FROM employees WHERE tenant_id = ? ORDER BY created_at DESC', [req.tenant.id]);
  const allLeads = queryAll('SELECT id, company, name, phone, stage, interest_level, talk_summary, deal_value, followup_date, owner_name, qualified_by FROM leads WHERE tenant_id = ? ORDER BY ROWID ASC', [req.tenant.id]);

  const employees = rows.map(r => {
    let prevCos = [];
    let docs = [];
    try { prevCos = JSON.parse(r.previous_companies_json || '[]'); } catch (e) {}
    try { docs = JSON.parse(r.documents_json || '[]'); } catch (e) {}

    const memberNameLower = String(r.full_name || '').toLowerCase().trim();
    const myLeads = allLeads.filter(l => String(l.owner_name || '').toLowerCase().trim() === memberNameLower);
    const workingLeads = myLeads.filter(l => ['New', 'Contacted', 'Proposal Sent', 'Negotiation'].includes(l.stage));
    const qualifiedLeads = allLeads.filter(l =>
      String(l.qualified_by || '').toLowerCase().trim() === memberNameLower ||
      (String(l.owner_name || '').toLowerCase().trim() === memberNameLower && ['Qualified', 'Proposal Sent', 'Negotiation', 'Won'].includes(l.stage))
    );
    const completedLeads = myLeads.filter(l => l.stage === 'Won');
    const revenueWon = completedLeads.reduce((s, l) => s + Number(l.deal_value || 0), 0);

    return {
      ...r,
      previous_companies: prevCos,
      documents: docs,
      metrics: {
        totalLeads: myLeads.length,
        workingCount: workingLeads.length,
        qualifiedCount: qualifiedLeads.length,
        completedCount: completedLeads.length,
        revenueWon,
        assignedLeads: myLeads
      }
    };
  });

  res.json({ employees });
});

app.post('/api/team', authMiddleware, (req, res) => {
  const {
    id, emp_code, full_name, designation, department, role,
    login_username, login_password, email, phone, joining_date,
    monthly_salary, location, education, emergency_contact, address, bank_details,
    background_summary, previous_companies
  } = req.body;

  if (!full_name || !designation) return res.status(400).json({ error: 'Full Name and Designation are required' });
  const prevJson = JSON.stringify(Array.isArray(previous_companies) ? previous_companies : []);
  const cleanEmail = String(email || `${full_name.toLowerCase().replace(/[^a-z0-9]/g, '')}@team.local`).toLowerCase().trim();
  const cleanUser = String(login_username || cleanEmail).toLowerCase().trim();
  const now = new Date().toISOString();

  if (id) {
    const existing = queryOne('SELECT * FROM employees WHERE id = ? AND tenant_id = ?', [id, req.tenant.id]);
    if (!existing) return res.status(404).json({ error: 'Team member not found' });

    const passHint = login_password ? login_password : (existing.login_password_hint || 'team123');
    runSql(`
      UPDATE employees SET emp_code=?, full_name=?, designation=?, department=?, role=?,
      login_username=?, login_password_hint=?, email=?, phone=?, joining_date=?,
      monthly_salary=?, location=?, education=?, emergency_contact=?, address=?,
      bank_details=?, background_summary=?, previous_companies_json=?
      WHERE id=? AND tenant_id=?
    `, [
      emp_code || existing.emp_code, full_name, designation, department || 'Sales & Revenue', role || 'agent',
      cleanUser, passHint, cleanEmail, phone || '', joining_date || existing.joining_date,
      Number(monthly_salary || 0), location || '', education || '', emergency_contact || '',
      address || '', bank_details || '', background_summary || '', prevJson, id, req.tenant.id
    ]);

    if (login_password) {
      const hash = bcrypt.hashSync(login_password, 10);
      const uRow = queryOne('SELECT * FROM users WHERE username = ? OR email = ?', [cleanUser, cleanEmail]);
      if (uRow) {
        runSql('UPDATE users SET username=?, email=?, password_hash=?, plain_pass_hint=?, full_name=?, role=? WHERE id=?',
          [cleanUser, cleanEmail, hash, login_password, full_name, role || 'agent', uRow.id]);
      } else {
        runSql(`
          INSERT INTO users (id, tenant_id, username, email, password_hash, plain_pass_hint, full_name, role, created_at)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        `, [uid('usr'), req.tenant.id, cleanUser, cleanEmail, hash, login_password, full_name, role || 'agent', now]);
      }
    }

    return res.json({ success: true, id, login_username: cleanUser, login_password: passHint });
  } else {
    const newId = uid('emp');
    const passToUse = login_password || 'team123';
    const hash = bcrypt.hashSync(passToUse, 10);

    runSql(`
      INSERT INTO employees (
        id, tenant_id, emp_code, full_name, designation, department, role,
        login_username, login_password_hint, email, phone,
        joining_date, monthly_salary, location, education, emergency_contact, address,
        bank_details, background_summary, previous_companies_json, documents_json, status, created_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '[]', 'Active', ?)
    `, [
      newId, req.tenant.id, emp_code || `EMP-${Date.now().toString().slice(-4)}`,
      full_name, designation, department || 'Sales & Revenue', role || 'agent',
      cleanUser, passToUse, cleanEmail, phone || '',
      joining_date || now.slice(0, 10), Number(monthly_salary || 0),
      location || '', education || '', emergency_contact || '', address || '',
      bank_details || '', background_summary || '', prevJson, now
    ], true);

    const existingU = queryOne('SELECT * FROM users WHERE username = ? OR email = ?', [cleanUser, cleanEmail]);
    if (!existingU) {
      runSql(`
        INSERT INTO users (id, tenant_id, username, email, password_hash, plain_pass_hint, full_name, role, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `, [uid('usr'), req.tenant.id, cleanUser, cleanEmail, hash, passToUse, full_name, role || 'agent', now], false);
    } else {
      runSql('UPDATE users SET password_hash = ?, plain_pass_hint = ? WHERE id = ?', [hash, passToUse, existingU.id], false);
    }

    return res.json({ success: true, id: newId, login_username: cleanUser, login_password: passToUse });
  }
});

app.post('/api/team/:id/documents', authMiddleware, upload.single('doc'), (req, res) => {
  const emp = queryOne('SELECT * FROM employees WHERE id = ? AND tenant_id = ?', [req.params.id, req.tenant.id]);
  if (!emp || !req.file) return res.status(404).json({ error: 'Team member or file not found' });

  let docs = [];
  try { docs = JSON.parse(emp.documents_json || '[]'); } catch (e) {}
  docs.push({
    id: uid('doc'),
    title: req.file.originalname,
    category: req.body.category || 'HRMS Document',
    size: Math.round(req.file.size / 1024) + ' KB',
    uploadedAt: new Date().toISOString().slice(0, 10),
    fileUrl: `/uploads/${req.file.filename}`
  });

  runSql('UPDATE employees SET documents_json = ? WHERE id = ?', [JSON.stringify(docs), emp.id]);
  res.json({ success: true, documents: docs });
});

app.delete('/api/team/:id', authMiddleware, (req, res) => {
  runSql('DELETE FROM employees WHERE id = ? AND tenant_id = ?', [req.params.id, req.tenant.id]);
  res.json({ success: true });
});

app.post('/api/subscriptions/upgrade-1yr', authMiddleware, (req, res) => {
  const { paymentGateway, transactionRef } = req.body;
  const currentEnd = new Date(req.tenant.access_end) > new Date() ? new Date(req.tenant.access_end) : new Date();
  const validUntil = addDaysISO(365, currentEnd);

  runSql(`UPDATE tenants SET plan_type = 'annual_1yr', status = 'active', access_end = ? WHERE id = ?`, [validUntil, req.tenant.id], true);
  runSql(`
    INSERT INTO subscription_payments (id, tenant_id, amount, plan_granted, payment_gateway, transaction_ref, valid_until, paid_at)
    VALUES (?, ?, 24999, 'annual_1yr', ?, ?, ?, ?)
  `, [uid('pay'), req.tenant.id, paymentGateway || 'Razorpay', transactionRef || uid('txn'), validUntil, new Date().toISOString()], false);

  res.json({ success: true, validUntil });
});

initSqlJs().then((SQL) => {
  if (fs.existsSync(DB_FILE)) {
    const fileBuffer = fs.readFileSync(DB_FILE);
    sqlDb = new SQL.Database(fileBuffer);
  } else {
    sqlDb = new SQL.Database();
  }
  sqlDb.run(SCHEMA_SQL);
  ensureAllColumnsAndTables();

  const netUrls = getLocalNetworkUrls();
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`✅ NovaCRM + HRMS Server v7 running!`);
    console.log(`💻 On this PC: http://localhost:${PORT}`);
    if (netUrls.length) console.log(`🌐 Share on Wi-Fi / Office LAN: ${netUrls[0]}`);
    console.log(`🔒 Hidden Master Admin Portal: http://localhost:${PORT}/master-admin.html`);
  });
});
