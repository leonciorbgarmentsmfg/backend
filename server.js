import { createHash, randomBytes, scryptSync } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Readable } from 'node:stream';
import express from 'express';
import multer from 'multer';
import ExcelJS from 'exceljs';
import Database from 'better-sqlite3';
import { v4 as uuidv4 } from 'uuid';
import db, { dbPath, verifyAdminCredentials } from './db/init.js';

const app = express();
const PORT = process.env.PORT || 5000;
const HOST = process.env.HOST || '0.0.0.0';
if (process.env.TRUST_PROXY) {
  app.set('trust proxy', process.env.TRUST_PROXY === 'true' ? 1 : process.env.TRUST_PROXY);
}
const SESSION_COOKIE = 'bpsms_session';
const SESSION_DURATION_MS = 8 * 60 * 60 * 1000;
const PAID_LAYOUT_FEE_THRESHOLD = 1;
const MINIMUM_LAYOUT_FEE = 700;
const orderPriorities = ['Urgent', 'High', 'Normal'];
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const uploadDir = process.env.BPSMS_UPLOAD_DIR || path.join(__dirname, 'uploads');
const frontendDist = path.resolve(__dirname, '../frontend/dist');
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 },
});
const databaseUpload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 },
});
const roleStations = {
  print_operator: 'printing',
  pressing_operator: 'pressing',
  cutting_operator: 'cutting',
  sewing_operator: 'sewing',
  packing_operator: 'packing',
};
const consumableCatalog = {
  subli_ink_cyan_atexco: { business_unit: 'sublimation', label: 'Sublimation ink · Atexco · Cyan', unit: 'L' },
  subli_ink_magenta_atexco: { business_unit: 'sublimation', label: 'Sublimation ink · Atexco · Magenta', unit: 'L' },
  subli_ink_yellow_atexco: { business_unit: 'sublimation', label: 'Sublimation ink · Atexco · Yellow', unit: 'L' },
  subli_ink_black_atexco: { business_unit: 'sublimation', label: 'Sublimation ink · Atexco · Black', unit: 'L' },
  subli_paper_atexco: { business_unit: 'sublimation', label: 'Sublimation paper · Atexco', unit: 'm' },
  subli_solution_atexco: { business_unit: 'sublimation', label: 'Sublimation solution · Atexco', unit: 'L' },
  subli_ink_cyan_creons: { business_unit: 'sublimation', label: 'Sublimation ink · Creons · Cyan', unit: 'L' },
  subli_ink_magenta_creons: { business_unit: 'sublimation', label: 'Sublimation ink · Creons · Magenta', unit: 'L' },
  subli_ink_yellow_creons: { business_unit: 'sublimation', label: 'Sublimation ink · Creons · Yellow', unit: 'L' },
  subli_ink_black_creons: { business_unit: 'sublimation', label: 'Sublimation ink · Creons · Black', unit: 'L' },
  subli_paper_creons: { business_unit: 'sublimation', label: 'Sublimation paper · Creons', unit: 'm' },
  subli_solution_creons: { business_unit: 'sublimation', label: 'Sublimation solution · Creons', unit: 'L' },
  dtf_ink_cyan: { business_unit: 'dtf', label: 'DTF ink · Cyan', unit: 'L' },
  dtf_ink_magenta: { business_unit: 'dtf', label: 'DTF ink · Magenta', unit: 'L' },
  dtf_ink_yellow: { business_unit: 'dtf', label: 'DTF ink · Yellow', unit: 'L' },
  dtf_ink_black: { business_unit: 'dtf', label: 'DTF ink · Black', unit: 'L' },
  dtf_ink_white: { business_unit: 'dtf', label: 'DTF ink · White', unit: 'L' },
  dtf_film: { business_unit: 'dtf', label: 'DTF film', unit: 'm' },
  dtf_solution: { business_unit: 'dtf', label: 'DTF solution', unit: 'L' },
  subli_ink_legacy_atexco: { business_unit: 'sublimation', label: 'Sublimation ink · Atexco · legacy (color not recorded)', unit: 'L', legacy: true },
  subli_ink_legacy_creons: { business_unit: 'sublimation', label: 'Sublimation ink · Creons · legacy (color not recorded)', unit: 'L', legacy: true },
};
const machineOrderFields = {
  printing: 'printer_machine',
  pressing: 'pressing_machine',
  cutting: 'cutting_machine',
};
const stages = ['printing', 'pressing', 'cutting', 'sewing', 'packing'];
const stageNext = {
  sizing: 'printing',
  printing: 'pressing',
  pressing: 'cutting',
  cutting: 'sewing',
  sewing: 'packing',
  packing: 'complete',
};
const styleOptionFields = {
  product_category: 'category',
  product_style: 'style',
  product_cut: 'cut',
  product_neck: 'neck',
  product_sleeve: 'sleeve_type',
  product_length: 'sleeve_length',
  product_others: 'other_details',
};

fs.mkdirSync(uploadDir, { recursive: true });
app.use(express.json({ limit: '1mb' }));

function hashToken(token) {
  return createHash('sha256').update(token).digest('hex');
}

function getSessionToken(req) {
  const cookies = req.headers.cookie?.split(';') ?? [];
  const sessionCookie = cookies
    .map((cookie) => cookie.trim())
    .find((cookie) => cookie.startsWith(`${SESSION_COOKIE}=`));
  const token = sessionCookie?.slice(SESSION_COOKIE.length + 1);
  return token && /^[a-f0-9]{64}$/.test(token) ? token : null;
}

function requireAuth(req, res, next) {
  const token = getSessionToken(req);
  if (!token) {
    return res.status(401).json({ error: 'Authentication required' });
  }

  const session = db
    .prepare(
      `SELECT users.id, users.name, users.username, users.role,
              users.assigned_station, users.business_unit,
              users.must_change_password, sessions.expires_at
       FROM sessions JOIN users ON users.id = sessions.user_id
       WHERE sessions.token_hash = ?`,
    )
    .get(hashToken(token));

  if (!session || session.expires_at <= Date.now()) {
    if (session) {
      db.prepare('DELETE FROM sessions WHERE token_hash = ?').run(
        hashToken(token),
      );
    }
    res.clearCookie(SESSION_COOKIE, { httpOnly: true, sameSite: 'lax' });
    return res.status(401).json({ error: 'Session expired; please log in again' });
  }

  if (
    session.must_change_password === 1 &&
    !['/api/auth/me', '/api/auth/change-password', '/api/auth/logout'].includes(req.path)
  ) {
    return res.status(403).json({ error: 'Change your temporary password before continuing' });
  }

  req.user = session;
  next();
}

function hasAdminAccess(user) {
  return user.role === 'admin' || user.role === 'supervisor';
}

function allowRoles(...roles) {
  return (req, res, next) => {
    if (hasAdminAccess(req.user) || roles.includes(req.user.role)) {
      return next();
    }
    return res.status(403).json({ error: 'You do not have permission for this action' });
  };
}

function allowOrderCrud(req, res, next) {
  if (req.user.role === 'supervisor' && req.user.business_unit === 'sublimation') {
    return res.status(403).json({ error: 'Sublimation supervisors have read-only access to orders' });
  }
  next();
}

function hashPassword(password) {
  const salt = randomBytes(16).toString('hex');
  return `${salt}:${scryptSync(password, salt, 64).toString('hex')}`;
}

function createId(prefix, table, column) {
  const year = new Date().getFullYear();
  const idPrefix = `${prefix}-${year}-`;
  const row = db
    .prepare(
      `SELECT COALESCE(MAX(CAST(substr(${column}, ?) AS INTEGER)), 0) + 1 AS next
       FROM ${table} WHERE ${column} LIKE ?`,
    )
    .get(idPrefix.length + 1, `${idPrefix}%`);
  return `${idPrefix}${String(row.next).padStart(4, '0')}`;
}

function resolveProductStyle(selectionIds, currentValues, existingOrder = null) {
  if (!selectionIds || typeof selectionIds !== 'object' || Array.isArray(selectionIds)) {
    return { error: 'Select all seven product style-number options' };
  }
  const values = {};
  const codes = [];
  let canGenerateNumber = true;
  for (const [category, field] of Object.entries(styleOptionFields)) {
    const rawId = selectionIds[category];
    if (rawId === '' || rawId === null || rawId === undefined) {
      if (
        existingOrder &&
        String(currentValues[field] ?? '') === String(existingOrder[field] ?? '')
      ) {
        values[field] = existingOrder[field];
        canGenerateNumber = false;
        continue;
      }
      return { error: `Select a ${category.replace('product_', '').replace('_', ' ')} option` };
    }
    const optionId = Number(rawId);
    if (!Number.isInteger(optionId) || optionId < 1) {
      return { error: 'One or more product style-number selections are invalid' };
    }
    const option = db
      .prepare('SELECT id, category, name, code FROM style_number_options WHERE id = ?')
      .get(optionId);
    if (!option || option.category !== category) {
      return { error: 'One or more product style-number selections are invalid' };
    }
    values[field] = option.name;
    codes.push(option.code);
  }
  return {
    values,
    style_number: canGenerateNumber
      ? `${codes.slice(0, 3).join('')}-${codes[3]}-${codes.slice(4).join('')}`
      : existingOrder.style_number,
  };
}

function formatItemPart(value) {
  return String(value ?? '')
    .toLocaleLowerCase()
    .replace(/(^|[\s-])([a-z])/g, (_, separator, letter) => `${separator}${letter.toLocaleUpperCase()}`);
}

function deriveItem(category, neck) {
  const formattedNeck = String(neck ?? '').toLocaleUpperCase() === 'ROUND'
    ? 'Round Neck'
    : formatItemPart(neck);
  return category && neck
    ? `${formatItemPart(category)} o ${formattedNeck}`
    : '';
}

function fetchOrder(orderId) {
  return db.prepare('SELECT * FROM orders WHERE order_id = ?').get(orderId);
}

function canAccessOrder(user, order, options = {}) {
  if (!order) return false;
  if (user.role === 'supervisor') {
    return order.business_unit === user.business_unit;
  }
  if (hasAdminAccess(user)) return true;
  if (user.role === 'sales') {
    return order.sales_rep_id === user.id;
  }
  if (user.role === 'artist') {
    return order.assigned_artist_id === user.id && Number(order.layout_fee) >= PAID_LAYOUT_FEE_THRESHOLD;
  }

  const station = roleStations[user.role];
  if (!station) return false;
  if (options.allowReprint && station === 'printing') {
    return true;
  }
  return order.order_status === station;
}

function publicOrder(order, user) {
  const client = db.prepare(
    `SELECT company_name, contact_person, contact_number, shipping_address
     FROM clients WHERE client_id = ?`,
  ).get(order.client_id);
  const safeOrder = {
    ...order,
    layout_fee_paid: Number(order.layout_fee) >= PAID_LAYOUT_FEE_THRESHOLD ? 1 : 0,
    sales_rep_name: db
      .prepare('SELECT name FROM users WHERE id = ?')
      .get(order.sales_rep_id)?.name ?? '',
    client_company_name: client?.company_name ?? '',
    client_contact_person: client?.contact_person ?? '',
    client_contact_number: client?.contact_number ?? '',
    client_shipping_address: client?.shipping_address ?? '',
  };
  if (!hasAdminAccess(user) && user.role !== 'sales') {
    delete safeOrder.layout_fee;
    delete safeOrder.layout_fee_paid;
    delete safeOrder.down_payment;
    delete safeOrder.per_piece_price;
    delete safeOrder.sales_rep_id;
  }
  return safeOrder;
}

function getOrderForUser(req, res, options) {
  const order = fetchOrder(req.params.orderId);
  if (!canAccessOrder(req.user, order, options)) {
    res.status(order ? 403 : 404).json({
      error: order ? 'You do not have access to this order' : 'Order not found',
    });
    return null;
  }
  return order;
}

function recordOrderEvent(orderId, actorId, eventType, fromStage, toStage, details = '') {
  db.prepare(
    `INSERT INTO order_events
      (order_id, actor_id, event_type, from_stage, to_stage, details)
     VALUES (?, ?, ?, ?, ?, ?)`,
  ).run(orderId, actorId, eventType, fromStage, toStage, details);
}

function handoffOrderToPressing(orderId, actorId, eventType, details) {
  const order = fetchOrder(orderId);
  if (!order) throw new Error(`Cannot hand off missing order ${orderId}`);
  db.prepare(
    `UPDATE orders SET order_status = 'pressing', updated_at = CURRENT_TIMESTAMP
     WHERE order_id = ?`,
  ).run(orderId);
  recordOrderEvent(orderId, actorId, eventType, order.order_status, 'pressing', details);
}

function canManageLineup(user, order) {
  if (hasAdminAccess(user)) return true;
  if (user.role === 'sales') return order.sales_rep_id === user.id;
  return user.role === 'artist' &&
    order.assigned_artist_id === user.id &&
    order.order_status === 'sizing';
}

async function parseLineup(buffer, ext, options) {
  const workbook = new ExcelJS.Workbook();
  if (ext === '.csv') {
    await workbook.csv.read(Readable.from([buffer]));
  } else {
    await workbook.xlsx.load(buffer);
  }
  const sheet = workbook.worksheets[0];
  if (!sheet || sheet.rowCount < 2) {
    throw new Error('The spreadsheet must contain a header and at least one lineup row');
  }

  const cellText = (value) => {
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    if (typeof value === 'object' && value !== null) {
      if ('result' in value) return String(value.result ?? '');
      if ('text' in value) return String(value.text ?? '');
      if ('richText' in value) return value.richText.map((part) => part.text).join('');
    }
    return String(value ?? '');
  };
  const headers = sheet.getRow(1).values.slice(1).map((value) => cellText(value).trim());
  const rows = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    rows.push(
      headers.map((_, index) => cellText(row.getCell(index + 1).value).trim()),
    );
  }
  return validateLineupTable(headers, rows, options);
}

function calculateLineupQuantity(headers, rows) {
  const sizeHeader = /^(xxs|xs|s|m|l|xl|xxl|[2-9]xl)$/i;
  const sizeColumns = headers.flatMap((header, index) => sizeHeader.test(header) ? [index] : []);
  const quantityColumn = headers.findIndex((header) => /^(qty|quantity|total quantity)$/i.test(header));
  const numericValue = (value) => {
    const quantity = Number(value);
    return Number.isFinite(quantity) && quantity > 0 ? quantity : 0;
  };
  const total = rows.reduce((sum, row) => {
    if (sizeColumns.length) {
      return sum + sizeColumns.reduce((rowTotal, column) => rowTotal + numericValue(row[column]), 0);
    }
    if (quantityColumn >= 0) return sum + numericValue(row[quantityColumn]);
    return sum + (row.some((value) => String(value ?? '').trim()) ? 1 : 0);
  }, 0);
  return total;
}

function validateFabricInventoryEntry(body) {
  const fabric = typeof body?.fabric === 'string' ? body.fabric.trim() : '';
  const fabricWidth = typeof body?.fabric_width === 'string' ? body.fabric_width.trim() : '';
  const gsm = Number(body?.gsm);
  const rollQuantity = Number(body?.roll_quantity);
  const measurementType = body?.measurement_type;
  const entryDate = body?.entry_date;
  if (!fabric || fabric.length > 100) {
    return { error: 'Fabric name is required and must be 100 characters or fewer' };
  }
  if (!fabricWidth || fabricWidth.length > 50) {
    return { error: 'Fabric width is required and must be 50 characters or fewer' };
  }
  if (!Number.isSafeInteger(gsm) || gsm <= 0) {
    return { error: 'GSM must be a positive whole number' };
  }
  if (!Number.isSafeInteger(rollQuantity) || rollQuantity <= 0) {
    return { error: 'Roll quantity must be a positive whole number' };
  }
  if (!['meter', 'kilo'].includes(measurementType)) {
    return { error: 'Select meter or kilo measurement' };
  }
  if (
    typeof entryDate !== 'string' ||
    !/^\d{4}-\d{2}-\d{2}$/.test(entryDate) ||
    Number.isNaN(new Date(`${entryDate}T00:00:00Z`).getTime()) ||
    new Date(`${entryDate}T00:00:00Z`).toISOString().slice(0, 10) !== entryDate
  ) {
    return { error: 'Enter a valid fabric entry date' };
  }

  if (measurementType === 'meter') {
    const metersPerRoll = Number(body?.meters_per_roll);
    if (!Number.isFinite(metersPerRoll) || metersPerRoll <= 0) {
      return { error: 'Meters per roll must be greater than zero' };
    }
    const totalMeter = metersPerRoll * rollQuantity;
    if (!Number.isFinite(totalMeter) || totalMeter <= 0) {
      return { error: 'Total meters must be a finite positive amount' };
    }
    return {
      value: {
        fabric,
        fabric_width: fabricWidth,
        gsm,
        roll_quantity: rollQuantity,
        measurement_type: measurementType,
        meters_per_roll: metersPerRoll,
        kilo_per_meter: null,
        kilo_per_roll: null,
        total_meter: totalMeter,
        entry_date: entryDate,
      },
    };
  }

  const metersPerKilo = Number(body?.kilo_per_meter);
  const kiloPerRoll = Number(body?.kilo_per_roll);
  if (!Number.isFinite(metersPerKilo) || metersPerKilo <= 0) {
    return { error: 'Meters per kilo must be greater than zero' };
  }
  if (!Number.isFinite(kiloPerRoll) || kiloPerRoll <= 0) {
    return { error: 'Kilo per roll must be greater than zero' };
  }
  const totalMeter = kiloPerRoll * rollQuantity * metersPerKilo;
  if (!Number.isFinite(totalMeter) || totalMeter <= 0) {
    return { error: 'Total meters must be a finite positive amount' };
  }
  return {
    value: {
      fabric,
      fabric_width: fabricWidth,
      gsm,
      roll_quantity: rollQuantity,
      measurement_type: measurementType,
      meters_per_roll: null,
      kilo_per_meter: metersPerKilo,
      kilo_per_roll: kiloPerRoll,
      total_meter: totalMeter,
      entry_date: entryDate,
    },
  };
}

function validateLineupTable(headers, rows, { requirePositiveQuantity = true } = {}) {
  if (
    !Array.isArray(headers) ||
    headers.length < 1 ||
    headers.length > 100 ||
    !headers.every((header) => typeof header === 'string' && header.trim())
  ) {
    throw new Error('Lineup must have between 1 and 100 non-empty column names');
  }
  if (!Array.isArray(rows) || rows.length < 1 || rows.length > 5000) {
    throw new Error('Lineup must contain between 1 and 5,000 rows');
  }
  if (!rows.every((row) => Array.isArray(row) && row.length === headers.length &&
    row.every((cell) => typeof cell === 'string'))) {
    throw new Error('Lineup rows must match the column count and contain text values');
  }
  const cleanHeaders = headers.map((header) => header.trim());
  const cleanRows = rows.map((row) => row.map((cell) => cell.trim()));
  const totalQuantity = calculateLineupQuantity(cleanHeaders, cleanRows);
  if (requirePositiveQuantity && totalQuantity < 1) {
    throw new Error('Could not calculate a positive total quantity');
  }
  return { headers: cleanHeaders, rows: cleanRows, totalQuantity };
}

async function encodeLineupTable(headers, rows, ext) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Lineup');
  sheet.addRow(headers);
  rows.forEach((row) => sheet.addRow(row));
  const contents = ext === '.csv'
    ? await workbook.csv.writeBuffer()
    : await workbook.xlsx.writeBuffer();
  return Buffer.from(contents);
}

function fileExtension(file) {
  const ext = path.extname(file.originalname).toLowerCase();
  if (!['.csv', '.xlsx', '.xls', '.png', '.jpg', '.jpeg', '.pdf'].includes(ext)) {
    throw new Error('Unsupported file type');
  }
  return ext;
}

function saveUpload(file) {
  const filename = `${uuidv4()}${fileExtension(file)}`;
  fs.writeFileSync(path.join(uploadDir, filename), file.buffer, { flag: 'wx' });
  return filename;
}

app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body ?? {};
  if (typeof username !== 'string' || typeof password !== 'string') {
    return res.status(400).json({ error: 'Username and password are required' });
  }

  const user = verifyAdminCredentials(username.trim(), password);
  if (!user) {
    return res.status(401).json({ error: 'Invalid username or password' });
  }

  const token = randomBytes(32).toString('hex');
  const expiresAt = Date.now() + SESSION_DURATION_MS;
  db.prepare(
    'INSERT INTO sessions (token_hash, user_id, expires_at) VALUES (?, ?, ?)',
  ).run(hashToken(token), user.id, expiresAt);

  res.cookie(SESSION_COOKIE, token, {
    httpOnly: true,
    sameSite: 'lax',
    secure: process.env.NODE_ENV === 'production',
    maxAge: SESSION_DURATION_MS,
    path: '/',
  });
  res.json({ user });
});

app.post('/api/auth/change-password', requireAuth, (req, res) => {
  const { password } = req.body ?? {};
  if (req.user.must_change_password !== 1) {
    return res.status(409).json({ error: 'A password change is not required' });
  }
  if (typeof password !== 'string' || password.length < 8) {
    return res.status(400).json({ error: 'New password must be at least 8 characters' });
  }
  db.prepare(
    'UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?',
  ).run(hashPassword(password), req.user.id);
  res.json({ message: 'Password updated. Please continue.' });
});

app.get('/api/auth/me', requireAuth, (req, res) => {
  const { expires_at: _expiresAt, ...user } = req.user;
  res.json({ user });
});

app.post('/api/auth/logout', (req, res) => {
  const token = getSessionToken(req);
  if (token) {
    db.prepare('DELETE FROM sessions WHERE token_hash = ?').run(hashToken(token));
  }
  res.clearCookie(SESSION_COOKIE, { httpOnly: true, sameSite: 'lax', path: '/' });
  res.status(204).end();
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'Backend is running' });
});

app.get('/api/staff', requireAuth, allowRoles('admin'), (req, res) => {
  const staff = db
    .prepare(
      `SELECT id, name, username, role, assigned_station, business_unit, is_protected
       FROM users ORDER BY name`,
    )
    .all();
  res.json({ staff });
});

app.get('/api/artists', requireAuth, allowRoles('sales'), (req, res) => {
  const artists = db
    .prepare("SELECT id, name FROM users WHERE role = 'artist' ORDER BY name")
    .all();
  res.json({ artists });
});

app.get(
  '/api/settings/machines',
  requireAuth,
  allowRoles('sales', 'supervisor', 'print_operator', 'pressing_operator', 'cutting_operator'),
  (req, res) => {
    const station = roleStations[req.user.role];
    const machines = station
      ? db.prepare('SELECT id, category, name, business_unit FROM machines WHERE category = ? ORDER BY name').all(station)
      : db.prepare('SELECT id, category, name, business_unit FROM machines ORDER BY category, business_unit, name').all();
  res.json({ machines });
  },
);

function getDatabaseFilename(value) {
  const requestedName = String(value ?? '').trim();
  const baseName = requestedName.toLocaleLowerCase().endsWith('.db')
    ? requestedName.slice(0, -3)
    : requestedName;
  if (
    !baseName ||
    baseName.length > 80 ||
    !/^[A-Za-z0-9][A-Za-z0-9 _-]*$/.test(baseName) ||
    path.basename(requestedName) !== requestedName
  ) {
    return null;
  }
  return `${baseName}.db`;
}

app.get('/api/settings/database/export', requireAuth, allowRoles('admin'), async (req, res) => {
  const filename = getDatabaseFilename(req.query.file_name);
  if (!filename) return res.status(400).json({ error: 'Enter a valid database filename' });
  const exportPath = path.join(path.dirname(dbPath), `${uuidv4()}-${filename}`);
  try {
    await db.backup(exportPath);
  } catch (error) {
    fs.rmSync(exportPath, { force: true });
    console.error('Unable to export the database:', error);
    return res.status(500).json({ error: 'Unable to export the database' });
  }
  res.download(exportPath, filename, (error) => {
    fs.rmSync(exportPath, { force: true });
    if (error && !res.headersSent) {
      res.status(500).json({ error: 'Unable to deliver the database export' });
    }
  });
});

app.post(
  '/api/settings/database/import',
  requireAuth,
  allowRoles('admin'),
  databaseUpload.single('database'),
  (req, res) => {
    const filename = getDatabaseFilename(req.body?.file_name);
    if (!filename) return res.status(400).json({ error: 'Enter a valid database filename' });
    if (!req.file || path.extname(req.file.originalname).toLocaleLowerCase() !== '.db') {
      return res.status(400).json({ error: 'Select a SQLite .db database file' });
    }
    const mode = req.body?.mode;
    if (!['replace', 'merge'].includes(mode)) {
      return res.status(400).json({ error: 'Select Replace or Merge import mode' });
    }

    const temporaryPath = path.join(path.dirname(dbPath), `${uuidv4()}-import.db`);
    let importedDb;
    let attached = false;
    try {
      fs.writeFileSync(temporaryPath, req.file.buffer, { flag: 'wx' });
      importedDb = new Database(temporaryPath, { readonly: true, fileMustExist: true });
      const integrity = importedDb.pragma('integrity_check');
      if (integrity.length !== 1 || integrity[0].integrity_check !== 'ok') {
        return res.status(400).json({ error: 'The selected file is not a valid SQLite database' });
      }
      const importedTables = importedDb
        .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
        .all()
        .map((row) => row.name);
      const currentTables = db
        .prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%' ORDER BY name")
        .all()
        .map((row) => row.name);
      if (importedTables.length !== currentTables.length ||
        importedTables.some((table, index) => table !== currentTables[index])) {
        return res.status(400).json({
          error: 'The database schema does not match this BPSMS version',
        });
      }
      for (const table of currentTables) {
        const quote = (identifier) => `"${identifier.replaceAll('"', '""')}"`;
        const currentColumns = db.prepare(`PRAGMA table_info(${quote(table)})`).all().map((column) => column.name);
        const incomingColumns = importedDb.prepare(`PRAGMA table_info(${quote(table)})`).all().map((column) => column.name);
        if (currentColumns.length !== incomingColumns.length ||
          currentColumns.some((column, index) => column !== incomingColumns[index])) {
          return res.status(400).json({ error: `The ${table} table is incompatible with this BPSMS version` });
        }
      }
      if (!importedDb.prepare("SELECT 1 FROM users WHERE is_protected = 1 AND role = 'admin' LIMIT 1").get()) {
        return res.status(400).json({ error: 'The database does not contain a protected administrator account' });
      }

      const protectedUserTriggers = db.prepare(
        `SELECT sql FROM sqlite_master
         WHERE type = 'trigger'
           AND name IN ('prevent_protected_user_update', 'prevent_protected_user_delete')
         ORDER BY name`,
      ).all().map((trigger) => trigger.sql).filter(Boolean);
      db.prepare('ATTACH DATABASE ? AS bpsms_import').run(temporaryPath);
      attached = true;
      const quote = (identifier) => `"${identifier.replaceAll('"', '""')}"`;
      db.pragma('foreign_keys = OFF');
      try {
        const importRows = db.transaction(() => {
          const dataTables = currentTables.filter((table) => table !== 'sessions');
          db.exec(
            'DROP TRIGGER IF EXISTS prevent_protected_user_update; DROP TRIGGER IF EXISTS prevent_protected_user_delete;',
          );
          if (mode === 'replace') {
            db.prepare('DELETE FROM sessions').run();
            for (const table of [...dataTables].reverse()) {
              db.exec(`DELETE FROM main.${quote(table)}`);
            }
          }
          for (const table of dataTables) {
            db.exec(
              `INSERT OR REPLACE INTO main.${quote(table)}
               SELECT * FROM bpsms_import.${quote(table)}`,
            );
          }
          const violations = db.prepare('PRAGMA main.foreign_key_check').all();
          if (violations.length) {
            throw new Error('The imported data would violate database relationships');
          }
          if (!db.prepare("SELECT 1 FROM main.users WHERE is_protected = 1 AND role = 'admin' LIMIT 1").get()) {
            throw new Error('The import must preserve a protected administrator account');
          }
          for (const trigger of protectedUserTriggers) db.exec(trigger);
        });
        importRows();
      } finally {
        db.pragma('foreign_keys = ON');
      }
      db.prepare('DETACH DATABASE bpsms_import').run();
      attached = false;
      res.json({
        message: mode === 'replace'
          ? `Database ${filename} restored. Sign in again to continue.`
          : `Database ${filename} merged successfully.`,
        mode,
      });
    } catch (error) {
      console.error('Unable to import the database:', error);
      if (error.message === 'The imported data would violate database relationships') {
        return res.status(409).json({ error: error.message });
      }
      return res.status(400).json({ error: 'Unable to import this database file' });
    } finally {
      if (attached) {
        try {
          db.prepare('DETACH DATABASE bpsms_import').run();
        } catch (error) {
          console.error('Unable to detach the imported database:', error);
        }
      }
      if (importedDb?.open) importedDb.close();
      fs.rmSync(temporaryPath, { force: true });
    }
  },
);

app.get('/api/settings/style-number-options', requireAuth, allowRoles('sales'), (req, res) => {
  const options = db
    .prepare('SELECT id, category, name, code FROM style_number_options ORDER BY category, id')
    .all();
  res.json({ options });
});

app.post('/api/settings/style-number-options', requireAuth, allowRoles('admin'), (req, res) => {
  const { category, name, code } = req.body ?? {};
  if (
    !Object.hasOwn(styleOptionFields, category) ||
    typeof name !== 'string' ||
    !name.trim() ||
    typeof code !== 'string' ||
    !/^[a-zA-Z0-9]{1,6}$/.test(code.trim())
  ) {
    return res.status(400).json({ error: 'Select a product category and provide a name and a 1–6 character alphanumeric code' });
  }
  try {
    const result = db
      .prepare('INSERT INTO style_number_options (category, name, code) VALUES (?, ?, ?)')
      .run(category, name.trim(), code.trim().toUpperCase());
    const option = db
      .prepare('SELECT id, category, name, code FROM style_number_options WHERE id = ?')
      .get(result.lastInsertRowid);
    res.status(201).json({ option });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'That name or code already exists in this product category' });
    }
    throw error;
  }
});

app.delete('/api/settings/style-number-options/:id', requireAuth, allowRoles('admin'), (req, res) => {
  const optionId = Number(req.params.id);
  if (!Number.isInteger(optionId)) {
    return res.status(400).json({ error: 'Invalid product style option ID' });
  }
  const option = db.prepare('SELECT * FROM style_number_options WHERE id = ?').get(optionId);
  if (!option) return res.status(404).json({ error: 'Product style option not found' });
  const orderField = styleOptionFields[option.category];
  if (db.prepare(`SELECT 1 FROM orders WHERE ${orderField} = ? LIMIT 1`).get(option.name)) {
    return res.status(409).json({ error: 'This option is used by an order and cannot be deleted' });
  }
  db.prepare('DELETE FROM style_number_options WHERE id = ?').run(optionId);
  res.status(204).end();
});

app.post('/api/settings/machines', requireAuth, allowRoles('admin'), (req, res) => {
  const { category, name } = req.body ?? {};
  const businessUnit = category === 'printing' ? req.body?.business_unit ?? null : null;
  if (
    !Object.hasOwn(machineOrderFields, category) ||
    typeof name !== 'string' ||
    !name.trim() ||
    (category === 'printing' && businessUnit !== null && !['dtf', 'sublimation'].includes(businessUnit))
  ) {
    return res.status(400).json({ error: 'Select a valid machine category and provide a name and printer section' });
  }
  try {
    const result = db
      .prepare('INSERT INTO machines (category, name, business_unit) VALUES (?, ?, ?)')
      .run(category, name.trim(), businessUnit);
    const machine = db.prepare('SELECT id, category, name, business_unit FROM machines WHERE id = ?').get(result.lastInsertRowid);
    res.status(201).json({ machine });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'That machine already exists in this category' });
    }
    throw error;
  }
});

app.put('/api/settings/machines/:id', requireAuth, allowRoles('admin'), (req, res) => {
  const machineId = Number(req.params.id);
  const { name } = req.body ?? {};
  const machine = Number.isInteger(machineId)
    ? db.prepare('SELECT * FROM machines WHERE id = ?').get(machineId)
    : null;
  const businessUnit = machine?.category === 'printing'
    ? req.body?.business_unit ?? machine.business_unit
    : null;
  if (!Number.isInteger(machineId) || typeof name !== 'string' || !name.trim() ||
    (machine?.category === 'printing' && businessUnit !== null && !['dtf', 'sublimation'].includes(businessUnit))) {
    return res.status(400).json({ error: 'Provide a valid machine ID, name, and printer section' });
  }
  if (!machine) return res.status(404).json({ error: 'Machine not found' });
  const field = machineOrderFields[machine.category];
  if (db.prepare(`SELECT 1 FROM orders WHERE ${field} = ? LIMIT 1`).get(machine.name)) {
    return res.status(409).json({ error: 'This machine is assigned to an order and cannot be renamed' });
  }
  try {
    db.prepare('UPDATE machines SET name = ?, business_unit = ? WHERE id = ?')
      .run(name.trim(), businessUnit, machineId);
    res.json({
      machine: db.prepare('SELECT id, category, name, business_unit FROM machines WHERE id = ?').get(machineId),
    });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'That machine already exists in this category' });
    }
    throw error;
  }
});

app.delete('/api/settings/machines/:id', requireAuth, allowRoles('admin'), (req, res) => {
  const machineId = Number(req.params.id);
  if (!Number.isInteger(machineId)) {
    return res.status(400).json({ error: 'Invalid machine ID' });
  }
  const machine = db.prepare('SELECT * FROM machines WHERE id = ?').get(machineId);
  if (!machine) return res.status(404).json({ error: 'Machine not found' });
  const field = machineOrderFields[machine.category];
  if (db.prepare(`SELECT 1 FROM orders WHERE ${field} = ? LIMIT 1`).get(machine.name)) {
    return res.status(409).json({ error: 'This machine is assigned to an order and cannot be deleted' });
  }
  db.prepare('DELETE FROM machines WHERE id = ?').run(machineId);
  res.status(204).end();
});

app.post('/api/staff', requireAuth, allowRoles('admin'), (req, res) => {
  const { name, username, password, role, assigned_station, business_unit } = req.body ?? {};
  const validRoles = [
    'admin',
    'supervisor',
    'sales',
    'artist',
    'print_operator',
    'pressing_operator',
    'cutting_operator',
    'sewing_operator',
    'packing_operator',
  ];

  if (
    typeof name !== 'string' ||
    typeof username !== 'string' ||
    typeof password !== 'string' ||
    !validRoles.includes(role) ||
    name.trim().length === 0 ||
    username.trim().length === 0 ||
    password.length < 8
  ) {
    return res.status(400).json({
      error: 'Provide a name, username, password of at least 8 characters, and a valid role',
    });
  }

  if (role === 'supervisor' && !['dtf', 'sublimation'].includes(business_unit)) {
    return res.status(400).json({ error: 'Select whether this Supervisor manages DTF or Sublimation reports' });
  }
  const station = roleStations[role] ?? null;
  if (station && assigned_station && assigned_station !== station) {
    return res.status(400).json({ error: `This role is assigned to ${station}` });
  }

  try {
    const result = db
      .prepare(
        `INSERT INTO users
          (name, username, password_hash, role, assigned_station, business_unit, must_change_password)
         VALUES (?, ?, ?, ?, ?, ?, 1)`,
      )
      .run(
        name.trim(),
        username.trim(),
        hashPassword(password),
        role,
        station,
        role === 'supervisor' ? business_unit : null,
      );
    const staff = db
      .prepare(
        `SELECT id, name, username, role, assigned_station, business_unit, must_change_password
         FROM users WHERE id = ?`,
      )
      .get(result.lastInsertRowid);
    res.status(201).json({ staff });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'Username is already in use' });
    }
    throw error;
  }
});

app.put('/api/staff/:id', requireAuth, allowRoles('admin'), (req, res) => {
  const staffId = Number(req.params.id);
  const { name, username, password, role, business_unit } = req.body ?? {};
  const validRoles = [
    'admin',
    'supervisor',
    'sales',
    'artist',
    'print_operator',
    'pressing_operator',
    'cutting_operator',
    'sewing_operator',
    'packing_operator',
  ];

  if (
    !Number.isInteger(staffId) ||
    typeof name !== 'string' ||
    typeof username !== 'string' ||
    !validRoles.includes(role) ||
    name.trim().length === 0 ||
    username.trim().length === 0 ||
    (password !== undefined && password !== '' &&
      (typeof password !== 'string' || password.length < 8))
  ) {
    return res.status(400).json({
      error: 'Provide a name, username, valid role, and an optional password of at least 8 characters',
    });
  }
  if (role === 'supervisor' && !['dtf', 'sublimation'].includes(business_unit)) {
    return res.status(400).json({ error: 'Select whether this Supervisor manages DTF or Sublimation reports' });
  }

  const existing = db
    .prepare('SELECT id, is_protected FROM users WHERE id = ?')
    .get(staffId);
  if (!existing) return res.status(404).json({ error: 'Staff account not found' });
  if (staffId === req.user.id) {
    return res.status(409).json({ error: 'You cannot edit the account you are currently using' });
  }
  if (existing.is_protected) {
    return res.status(409).json({ error: 'The protected bootstrap administrator cannot be edited' });
  }

  const station = roleStations[role] ?? null;
  const updateStaff = db.transaction(() => {
    if (typeof password === 'string' && password.length > 0) {
      db.prepare(
        `UPDATE users
         SET name = ?, username = ?, role = ?, assigned_station = ?, business_unit = ?,
             password_hash = ?, must_change_password = 1
         WHERE id = ?`,
      ).run(
        name.trim(),
        username.trim(),
        role,
        station,
        role === 'supervisor' ? business_unit : null,
        hashPassword(password),
        staffId,
      );
    } else {
      db.prepare(
        `UPDATE users
         SET name = ?, username = ?, role = ?, assigned_station = ?, business_unit = ?
         WHERE id = ?`,
      ).run(name.trim(), username.trim(), role, station, role === 'supervisor' ? business_unit : null, staffId);
    }

    if (typeof password === 'string' && password.length > 0) {
      const currentToken = getSessionToken(req);
      db.prepare('DELETE FROM sessions WHERE user_id = ? AND token_hash != ?').run(
        staffId,
        currentToken ? hashToken(currentToken) : '',
      );
    }
  });

  try {
    updateStaff();
    const staff = db
      .prepare(
        `SELECT id, name, username, role, assigned_station, business_unit, must_change_password,
                is_protected
         FROM users WHERE id = ?`,
      )
      .get(staffId);
    res.json({ staff });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'Username is already in use' });
    }
    throw error;
  }
});

app.delete('/api/staff/:id', requireAuth, allowRoles('admin'), (req, res) => {
  const staffId = Number(req.params.id);
  if (!Number.isInteger(staffId)) {
    return res.status(400).json({ error: 'Invalid staff account ID' });
  }
  if (staffId === req.user.id) {
    return res.status(409).json({ error: 'You cannot delete the account you are currently using' });
  }

  const staff = db
    .prepare('SELECT id, name, is_protected FROM users WHERE id = ?')
    .get(staffId);
  if (!staff) return res.status(404).json({ error: 'Staff account not found' });
  if (staff.is_protected) {
    return res.status(409).json({ error: 'The protected bootstrap administrator cannot be deleted' });
  }

  try {
    db.prepare('DELETE FROM users WHERE id = ?').run(staffId);
    res.status(204).end();
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_FOREIGNKEY') {
      return res.status(409).json({
        error: 'This account is linked to clients, orders, production logs, or consumable transactions and cannot be deleted without removing that history',
      });
    }
    throw error;
  }
});

app.get('/api/clients', requireAuth, allowRoles('sales', 'supervisor'), (req, res) => {
  const query =
    hasAdminAccess(req.user)
      ? `SELECT clients.*, users.name AS sales_rep_name
         FROM clients JOIN users ON users.id = clients.sales_rep_id
         ORDER BY clients.company_name`
      : `SELECT clients.*, users.name AS sales_rep_name
         FROM clients JOIN users ON users.id = clients.sales_rep_id
         WHERE clients.sales_rep_id = ? ORDER BY clients.company_name`;
  const clients =
    hasAdminAccess(req.user)
      ? db.prepare(query).all()
      : db.prepare(query).all(req.user.id);
  res.json({ clients });
});

app.post('/api/clients', requireAuth, allowRoles('sales'), (req, res) => {
  const { company_name, contact_person, contact_number, shipping_address } =
    req.body ?? {};
  if (
    ![company_name, contact_person, contact_number, shipping_address].every(
      (value) => typeof value === 'string' && value.trim(),
    )
  ) {
    return res.status(400).json({ error: 'All client profile fields are required' });
  }

  const salesRepId = hasAdminAccess(req.user)
    ? Number(req.body?.sales_rep_id)
    : req.user.id;
  const salesRep = db
    .prepare("SELECT id FROM users WHERE id = ? AND role = 'sales'")
    .get(salesRepId);
  if (!salesRep) {
    return res.status(400).json({ error: 'Select a valid Sales representative' });
  }

  const clientId = createId('CLI', 'clients', 'client_id');
  try {
    db.prepare(
      `INSERT INTO clients
        (client_id, company_name, contact_person, contact_number, shipping_address, sales_rep_id)
       VALUES (?, ?, ?, ?, ?, ?)`,
    ).run(
      clientId,
      company_name.trim(),
      contact_person.trim(),
      contact_number.trim(),
      shipping_address.trim(),
      salesRepId,
    );
    const client = db.prepare('SELECT * FROM clients WHERE client_id = ?').get(clientId);
    res.status(201).json({ client });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'A client with that company name already exists' });
    }
    throw error;
  }
});

app.put('/api/clients/:clientId', requireAuth, allowRoles('sales'), (req, res) => {
  const { company_name, contact_person, contact_number, shipping_address } =
    req.body ?? {};
  if (
    ![company_name, contact_person, contact_number, shipping_address].every(
      (value) => typeof value === 'string' && value.trim(),
    )
  ) {
    return res.status(400).json({ error: 'All client profile fields are required' });
  }

  const client = db
    .prepare('SELECT * FROM clients WHERE client_id = ?')
    .get(req.params.clientId);
  if (!client || (req.user.role === 'sales' && client.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Client not found' });
  }

  const salesRepId = hasAdminAccess(req.user)
    ? Number(req.body?.sales_rep_id)
    : client.sales_rep_id;
  const salesRep = db
    .prepare("SELECT id FROM users WHERE id = ? AND role = 'sales'")
    .get(salesRepId);
  if (!salesRep) {
    return res.status(400).json({ error: 'Select a valid Sales representative' });
  }

  try {
    db.prepare(
      `UPDATE clients
       SET company_name = ?, contact_person = ?, contact_number = ?, shipping_address = ?, sales_rep_id = ?
       WHERE client_id = ?`,
    ).run(
      company_name.trim(),
      contact_person.trim(),
      contact_number.trim(),
      shipping_address.trim(),
      salesRepId,
      client.client_id,
    );
    res.json({
      client: db.prepare('SELECT * FROM clients WHERE client_id = ?').get(client.client_id),
    });
  } catch (error) {
    if (error.code === 'SQLITE_CONSTRAINT_UNIQUE') {
      return res.status(409).json({ error: 'A client with that company name already exists' });
    }
    throw error;
  }
});

app.delete('/api/clients/:clientId', requireAuth, allowRoles('sales'), (req, res) => {
  const client = db
    .prepare('SELECT * FROM clients WHERE client_id = ?')
    .get(req.params.clientId);
  if (!client || (req.user.role === 'sales' && client.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Client not found' });
  }
  const hasOrders = db
    .prepare('SELECT 1 FROM orders WHERE client_id = ? LIMIT 1')
    .get(client.client_id);
  if (hasOrders) {
    return res.status(409).json({
      error: 'This client has production orders and cannot be deleted',
    });
  }

  db.prepare('DELETE FROM clients WHERE client_id = ?').run(client.client_id);
  res.status(204).end();
});

function canManageFabricInventory(user) {
  return hasAdminAccess(user) || user.role === 'pressing_operator';
}

function isOrderFabricAvailable(fabric, fabricWidth, gsm, existingOrder = null) {
  if (
    existingOrder &&
    existingOrder.fabric === fabric &&
    (existingOrder.fabric_width || '') === fabricWidth &&
    Number(existingOrder.gsm) === Number(gsm)
  ) {
    return true;
  }
  return Boolean(db.prepare(
    `SELECT 1 FROM fabric_inventory
     WHERE fabric = ? COLLATE NOCASE
       AND fabric_width = ? COLLATE NOCASE
       AND gsm = ?
     LIMIT 1`,
  ).get(fabric, fabricWidth, Number(gsm)));
}

app.get('/api/fabric-inventory', requireAuth, allowRoles('pressing_operator', 'supervisor'), (req, res) => {
  const filters = [];
  const params = [];
  if (typeof req.query.date === 'string' && req.query.date) {
    filters.push('fabric_inventory.entry_date = ?');
    params.push(req.query.date);
  }
  if (typeof req.query.fabric === 'string' && req.query.fabric.trim()) {
    filters.push('fabric_inventory.fabric LIKE ? COLLATE NOCASE');
    params.push(`%${req.query.fabric.trim()}%`);
  }
  const where = filters.length ? `WHERE ${filters.join(' AND ')}` : '';
  const entries = db.prepare(
    `SELECT fabric_inventory.*, created_user.name AS created_by_name,
            updated_user.name AS updated_by_name
     FROM fabric_inventory
     JOIN users AS created_user ON created_user.id = fabric_inventory.created_by
     JOIN users AS updated_user ON updated_user.id = fabric_inventory.updated_by
     ${where}
     ORDER BY fabric_inventory.entry_date DESC, fabric_inventory.id DESC`,
  ).all(...params);
  res.json({ entries });
});

app.get('/api/fabric-inventory/options', requireAuth, (req, res) => {
  const options = db.prepare(
    `SELECT DISTINCT fabric, fabric_width, gsm
     FROM fabric_inventory
     ORDER BY fabric COLLATE NOCASE, fabric_width COLLATE NOCASE, gsm`,
  ).all();
  res.json({ options });
});

app.post('/api/fabric-inventory', requireAuth, (req, res, next) => {
  if (!canManageFabricInventory(req.user)) {
    return res.status(403).json({ error: 'You do not have permission to manage fabric inventory' });
  }
  next();
}, (req, res) => {
  const result = validateFabricInventoryEntry(req.body);
  if (result.error) return res.status(400).json({ error: result.error });
  const entry = result.value;
  const insert = db.prepare(
    `INSERT INTO fabric_inventory (
       fabric, fabric_width, gsm, roll_quantity, measurement_type, meters_per_roll,
       kilo_per_meter, kilo_per_roll, total_meter, entry_date, created_by, updated_by
     ) VALUES (
       @fabric, @fabric_width, @gsm, @roll_quantity, @measurement_type, @meters_per_roll,
       @kilo_per_meter, @kilo_per_roll, @total_meter, @entry_date, @user_id, @user_id
     )`,
  ).run({ ...entry, user_id: req.user.id });
  const saved = db.prepare('SELECT * FROM fabric_inventory WHERE id = ?').get(insert.lastInsertRowid);
  res.status(201).json({ entry: saved });
});

app.put('/api/fabric-inventory/:id', requireAuth, (req, res, next) => {
  if (!canManageFabricInventory(req.user)) {
    return res.status(403).json({ error: 'You do not have permission to manage fabric inventory' });
  }
  next();
}, (req, res) => {
  const entryId = Number(req.params.id);
  if (!Number.isInteger(entryId) || entryId <= 0) {
    return res.status(400).json({ error: 'Fabric entry ID is invalid' });
  }
  const existing = db.prepare('SELECT id FROM fabric_inventory WHERE id = ?').get(entryId);
  if (!existing) return res.status(404).json({ error: 'Fabric entry not found' });
  const result = validateFabricInventoryEntry(req.body);
  if (result.error) return res.status(400).json({ error: result.error });
  db.prepare(
    `UPDATE fabric_inventory SET
       fabric = @fabric, fabric_width = @fabric_width, gsm = @gsm, roll_quantity = @roll_quantity,
       measurement_type = @measurement_type, meters_per_roll = @meters_per_roll,
       kilo_per_meter = @kilo_per_meter, kilo_per_roll = @kilo_per_roll,
       total_meter = @total_meter, entry_date = @entry_date,
       updated_by = @user_id, updated_at = CURRENT_TIMESTAMP
     WHERE id = @id`,
  ).run({ ...result.value, user_id: req.user.id, id: entryId });
  const saved = db.prepare('SELECT * FROM fabric_inventory WHERE id = ?').get(entryId);
  res.json({ entry: saved });
});

app.delete('/api/fabric-inventory/:id', requireAuth, (req, res, next) => {
  if (!canManageFabricInventory(req.user)) {
    return res.status(403).json({ error: 'You do not have permission to manage fabric inventory' });
  }
  next();
}, (req, res) => {
  const entryId = Number(req.params.id);
  if (!Number.isInteger(entryId) || entryId <= 0) {
    return res.status(400).json({ error: 'Fabric entry ID is invalid' });
  }
  const result = db.prepare('DELETE FROM fabric_inventory WHERE id = ?').run(entryId);
  if (result.changes === 0) return res.status(404).json({ error: 'Fabric entry not found' });
  res.status(204).end();
});

app.get('/api/consumables', requireAuth, allowRoles('supervisor', 'print_operator'), (req, res) => {
  const requestedUnit = req.query.business_unit;
  const businessUnits = hasAdminAccess(req.user)
    ? (['dtf', 'sublimation'].includes(requestedUnit) ? [requestedUnit] : ['dtf', 'sublimation'])
    : req.user.role === 'print_operator'
      ? ['sublimation']
      : [];
  if (businessUnits.some((unit) => !['dtf', 'sublimation'].includes(unit))) {
    return res.status(403).json({ error: 'Supervisor business unit is not configured' });
  }

  const placeholders = businessUnits.map(() => '?').join(', ');
  const balances = db.prepare(
    `SELECT consumable_code,
            COALESCE(SUM(CASE
              WHEN transaction_type = 'stock_in' THEN quantity
              WHEN transaction_type = 'usage' THEN -quantity
              ELSE quantity
            END), 0) AS on_hand,
            COALESCE(SUM(CASE WHEN transaction_type = 'stock_in' THEN quantity ELSE 0 END), 0) AS stock_received,
            COALESCE(SUM(CASE WHEN transaction_type = 'usage' THEN quantity ELSE 0 END), 0) AS used
     FROM consumable_transactions
     WHERE business_unit IN (${placeholders})
     GROUP BY consumable_code`,
  ).all(...businessUnits);
  const balanceByCode = new Map(balances.map((entry) => [entry.consumable_code, entry]));
  const inventory = Object.entries(consumableCatalog)
    .filter(([, product]) => businessUnits.includes(product.business_unit))
    .map(([code, product]) => ({
      code,
      ...product,
      ...(balanceByCode.get(code) || { on_hand: 0, stock_received: 0, used: 0 }),
    }));
  const transactions = db.prepare(
    `SELECT consumable_transactions.*, users.name AS recorded_by,
            orders.team_project_name
     FROM consumable_transactions
     JOIN users ON users.id = consumable_transactions.created_by
     LEFT JOIN orders ON orders.order_id = consumable_transactions.order_id
     WHERE consumable_transactions.business_unit IN (${placeholders})
     ORDER BY consumable_transactions.entry_date DESC, consumable_transactions.id DESC
     LIMIT 500`,
  ).all(...businessUnits).map((entry) => ({
    ...entry,
    consumable_label: consumableCatalog[entry.consumable_code]?.label || entry.consumable_code,
    unit: consumableCatalog[entry.consumable_code]?.unit || '',
  }));
  res.json({ inventory, transactions, business_units: businessUnits });
});

app.post('/api/consumables/transactions', requireAuth, allowRoles('supervisor'), (req, res) => {
  const { business_unit: businessUnit, consumable_code: consumableCode, transaction_type: type, entry_date: entryDate } = req.body ?? {};
  const product = consumableCatalog[consumableCode];
  const quantity = Number(req.body?.quantity);
  const note = typeof req.body?.note === 'string' ? req.body.note.trim() : '';
  if (!product || product.legacy || product.business_unit !== businessUnit) {
    return res.status(400).json({ error: 'Select a valid consumable for the business unit' });
  }
  if (
    !['stock_in', 'adjustment'].includes(type) ||
    !Number.isFinite(quantity) ||
    (type === 'stock_in' && quantity <= 0) ||
    (type === 'adjustment' && quantity === 0)
  ) {
    return res.status(400).json({ error: 'Enter a valid stock receipt or non-zero stock adjustment quantity' });
  }
  if (
    typeof entryDate !== 'string' ||
    !/^\d{4}-\d{2}-\d{2}$/.test(entryDate) ||
    Number.isNaN(new Date(`${entryDate}T00:00:00Z`).getTime()) ||
    new Date(`${entryDate}T00:00:00Z`).toISOString().slice(0, 10) !== entryDate
  ) {
    return res.status(400).json({ error: 'Enter a valid consumable transaction date' });
  }
  if (type === 'adjustment' && !note) {
    return res.status(400).json({ error: 'Enter a reason for the stock adjustment' });
  }

  const insert = db.transaction(() => {
    const currentBalance = db.prepare(
      `SELECT COALESCE(SUM(CASE
         WHEN transaction_type = 'stock_in' THEN quantity
         WHEN transaction_type = 'usage' THEN -quantity
         ELSE quantity
       END), 0) AS on_hand
       FROM consumable_transactions
       WHERE business_unit = ? AND consumable_code = ?`,
    ).get(businessUnit, consumableCode).on_hand;
    if (currentBalance + quantity < 0) {
      return { error: `Adjustment would make ${product.label} stock negative` };
    }
    const result = db.prepare(
      `INSERT INTO consumable_transactions
         (business_unit, consumable_code, transaction_type, quantity, entry_date, note, created_by)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
    ).run(businessUnit, consumableCode, type, quantity, entryDate, note, req.user.id);
    return { id: result.lastInsertRowid };
  });
  const result = insert();
  if (result.error) return res.status(400).json({ error: result.error });
  res.status(201).json({ id: result.id });
});

function isValidInventoryDate(entryDate) {
  return typeof entryDate === 'string' &&
    /^\d{4}-\d{2}-\d{2}$/.test(entryDate) &&
    !Number.isNaN(new Date(`${entryDate}T00:00:00Z`).getTime()) &&
    new Date(`${entryDate}T00:00:00Z`).toISOString().slice(0, 10) === entryDate;
}

function getConsumableTransactionImpact(type, quantity) {
  return type === 'usage' ? -quantity : quantity;
}

function getConsumableBalance(businessUnit, consumableCode, excludedId = null) {
  return db.prepare(
    `SELECT COALESCE(SUM(CASE
       WHEN transaction_type = 'stock_in' THEN quantity
       WHEN transaction_type = 'usage' THEN -quantity
       ELSE quantity
     END), 0) AS on_hand
     FROM consumable_transactions
     WHERE business_unit = ? AND consumable_code = ?
       AND (? IS NULL OR id != ?)`,
  ).get(businessUnit, consumableCode, excludedId, excludedId).on_hand;
}

function syncPaperUsageForStationLog(transaction) {
  if (
    transaction.transaction_type !== 'usage' ||
    !transaction.station_log_id ||
    !transaction.consumable_code.startsWith('subli_paper_')
  ) return;
  const paperMeters = db.prepare(
    `SELECT COALESCE(SUM(quantity), 0) AS paper_meters
     FROM consumable_transactions
     WHERE station_log_id = ? AND transaction_type = 'usage'
       AND consumable_code LIKE 'subli_paper_%'`,
  ).get(transaction.station_log_id).paper_meters;
  db.prepare('UPDATE station_logs SET paper_meters = ? WHERE id = ?')
    .run(paperMeters, transaction.station_log_id);
}

app.put('/api/consumables/transactions/:id', requireAuth, allowRoles('supervisor'), (req, res) => {
  const transactionId = Number(req.params.id);
  if (!Number.isInteger(transactionId) || transactionId <= 0) {
    return res.status(400).json({ error: 'Consumable transaction ID is invalid' });
  }
  const existing = db.prepare('SELECT * FROM consumable_transactions WHERE id = ?').get(transactionId);
  if (!existing) return res.status(404).json({ error: 'Consumable transaction not found' });
  const quantity = Number(req.body?.quantity);
  const entryDate = req.body?.entry_date;
  const note = typeof req.body?.note === 'string' ? req.body.note.trim() : '';
  if (
    !Number.isFinite(quantity) ||
    (existing.transaction_type !== 'adjustment' && quantity <= 0) ||
    (existing.transaction_type === 'adjustment' && quantity === 0)
  ) {
    return res.status(400).json({ error: 'Enter a valid non-zero transaction quantity' });
  }
  if (!isValidInventoryDate(entryDate)) {
    return res.status(400).json({ error: 'Enter a valid consumable transaction date' });
  }
  if (existing.transaction_type === 'adjustment' && !note) {
    return res.status(400).json({ error: 'Enter a reason for the stock adjustment' });
  }

  const update = db.transaction(() => {
    const balanceWithoutEntry = getConsumableBalance(
      existing.business_unit,
      existing.consumable_code,
      transactionId,
    );
    const resultingBalance = balanceWithoutEntry +
      getConsumableTransactionImpact(existing.transaction_type, quantity);
    if (resultingBalance < -1e-9) {
      return { error: `This change would make ${consumableCatalog[existing.consumable_code]?.label || existing.consumable_code} stock negative` };
    }
    db.prepare(
      `UPDATE consumable_transactions
       SET quantity = ?, entry_date = ?, note = ?
       WHERE id = ?`,
    ).run(quantity, entryDate, note, transactionId);
    syncPaperUsageForStationLog(existing);
    return {};
  });
  const result = update();
  if (result.error) return res.status(400).json({ error: result.error });
  res.json({ id: transactionId });
});

app.delete('/api/consumables/transactions/:id', requireAuth, allowRoles('supervisor'), (req, res) => {
  const transactionId = Number(req.params.id);
  if (!Number.isInteger(transactionId) || transactionId <= 0) {
    return res.status(400).json({ error: 'Consumable transaction ID is invalid' });
  }
  const existing = db.prepare('SELECT * FROM consumable_transactions WHERE id = ?').get(transactionId);
  if (!existing) return res.status(404).json({ error: 'Consumable transaction not found' });
  const remove = db.transaction(() => {
    const resultingBalance = getConsumableBalance(existing.business_unit, existing.consumable_code, transactionId);
    if (resultingBalance < -1e-9) {
      return { error: `This transaction cannot be deleted because it would make ${consumableCatalog[existing.consumable_code]?.label || existing.consumable_code} stock negative` };
    }
    db.prepare('DELETE FROM consumable_transactions WHERE id = ?').run(transactionId);
    syncPaperUsageForStationLog(existing);
    return {};
  });
  const result = remove();
  if (result.error) return res.status(409).json({ error: result.error });
  res.status(204).end();
});

app.get('/api/orders', requireAuth, (req, res) => {
  let orders;
  if (hasAdminAccess(req.user)) {
    orders = db.prepare('SELECT * FROM orders ORDER BY created_at DESC').all();
  } else if (req.user.role === 'sales') {
    orders = db
      .prepare('SELECT * FROM orders WHERE sales_rep_id = ? ORDER BY created_at DESC')
      .all(req.user.id);
  } else if (req.user.role === 'artist') {
    orders = db
      .prepare(
        `SELECT * FROM orders
         WHERE assigned_artist_id = ? AND layout_fee >= ?
         ORDER BY created_at DESC`,
      )
      .all(req.user.id, PAID_LAYOUT_FEE_THRESHOLD);
  } else {
    const station = roleStations[req.user.role];
    if (!station) return res.status(403).json({ error: 'You do not have order access' });

    orders =
      station === 'printing'
        ? db
            .prepare(
              `SELECT DISTINCT orders.* FROM orders
               LEFT JOIN reprint_jobs ON reprint_jobs.order_id = orders.order_id
               WHERE orders.order_status = 'printing'
                  OR reprint_jobs.status != 'complete'
                  OR EXISTS (
                    SELECT 1 FROM order_events
                    WHERE order_events.order_id = orders.order_id
                      AND order_events.to_stage IN ('printing', 'pressing', 'cutting', 'sewing', 'packing', 'complete')
                  )
               ORDER BY orders.created_at DESC`,
            )
            .all()
        : db
            .prepare('SELECT * FROM orders WHERE order_status = ? ORDER BY created_at DESC')
            .all(station);
  }
  const stationLogs = orders.length
    ? db
        .prepare(
          `SELECT station_logs.id, station_logs.order_id, station_logs.station,
                  station_logs.fabric_used, station_logs.printer_brand, station_logs.paper_meters,
                  station_logs.defect_description, station_logs.reprint_job_id,
                  station_logs.created_at, users.name AS operator
           FROM station_logs
           JOIN users ON users.id = station_logs.operator_id
           WHERE station_logs.order_id IN (${orders.map(() => '?').join(', ')})
           ORDER BY station_logs.id DESC`,
        )
        .all(...orders.map((order) => order.order_id))
    : [];
  const logsByOrder = new Map();
  for (const log of stationLogs) {
    const logs = logsByOrder.get(log.order_id) || [];
    logs.push(log);
    logsByOrder.set(log.order_id, logs);
  }
  const productionItems = orders.length
    ? db
        .prepare(
          `SELECT production_log_items.*, station_logs.order_id
           FROM production_log_items
           JOIN station_logs ON station_logs.id = production_log_items.station_log_id
           WHERE station_logs.order_id IN (${orders.map(() => '?').join(', ')})
           ORDER BY production_log_items.id`,
        )
        .all(...orders.map((order) => order.order_id))
    : [];
  const itemsByLog = new Map();
  for (const item of productionItems) {
    const items = itemsByLog.get(item.station_log_id) || [];
    items.push(item);
    itemsByLog.set(item.station_log_id, items);
  }
  for (const logs of logsByOrder.values()) {
    for (const log of logs) {
      log.production_items = itemsByLog.get(log.id) || [];
    }
  }
  const pressingNotes = orders.length
    ? db
        .prepare(
          `SELECT order_events.order_id, order_events.event_type, order_events.details,
                  order_events.created_at, users.name AS actor_name
           FROM order_events
           JOIN users ON users.id = order_events.actor_id
           WHERE order_events.order_id IN (${orders.map(() => '?').join(', ')})
             AND order_events.to_stage = 'pressing'
             AND order_events.event_type IN (
               'stage_completed', 'reprint_completed', 'test_print_completed'
             )
             AND TRIM(order_events.details) != ''
           ORDER BY order_events.id DESC`,
        )
        .all(...orders.map((order) => order.order_id))
    : [];
  const pressingNotesByOrder = new Map();
  for (const note of pressingNotes) {
    const notes = pressingNotesByOrder.get(note.order_id) || [];
    notes.push(note);
    pressingNotesByOrder.set(note.order_id, notes);
  }
  res.json({
    orders: orders.map((order) => ({
      ...publicOrder(order, req.user),
      station_logs: logsByOrder.get(order.order_id) || [],
      pressing_notes: pressingNotesByOrder.get(order.order_id) || [],
    })),
  });
});

app.get('/api/order-priorities', requireAuth, (req, res) => {
  const activeOrders = "order_status NOT IN ('complete', 'pickup_delivered')";
  let where;
  let params = [];
  if (hasAdminAccess(req.user)) {
    where = activeOrders;
  } else if (req.user.role === 'sales') {
    where = `${activeOrders} AND sales_rep_id = ?`;
    params = [req.user.id];
  } else if (req.user.role === 'artist') {
    where = `${activeOrders} AND assigned_artist_id = ? AND layout_fee >= ?`;
    params = [req.user.id, PAID_LAYOUT_FEE_THRESHOLD];
  } else if (roleStations[req.user.role]) {
    where = `${activeOrders} AND order_status = ?`;
    params = [roleStations[req.user.role]];
  } else {
    return res.json({ orders: [] });
  }
  const orders = db.prepare(
    `SELECT order_id, team_project_name, order_status, due_date, order_priority, print_mode
     FROM orders
     WHERE ${where}
     ORDER BY CASE order_priority
       WHEN 'Urgent' THEN 0
       WHEN 'High' THEN 1
       ELSE 2
     END, due_date, order_id`,
  ).all(...params);
  res.json({ orders });
});

app.put('/api/orders/:orderId/priority', requireAuth, allowRoles('admin', 'sales', 'supervisor'), (req, res) => {
  const order = fetchOrder(req.params.orderId);
  if (
    !order ||
    (req.user.role === 'sales' && order.sales_rep_id !== req.user.id)
  ) {
    return res.status(404).json({ error: 'Order not found' });
  }
  if (['complete', 'pickup_delivered'].includes(order.order_status)) {
    return res.status(409).json({ error: 'Completed orders cannot have their priority changed' });
  }
  const priority = req.body?.order_priority;
  if (!orderPriorities.includes(priority)) {
    return res.status(400).json({ error: `Select one of: ${orderPriorities.join(', ')}` });
  }
  if (order.order_priority !== priority) {
    db.prepare(
      'UPDATE orders SET order_priority = ?, updated_at = CURRENT_TIMESTAMP WHERE order_id = ?',
    ).run(priority, order.order_id);
    recordOrderEvent(
      order.order_id,
      req.user.id,
      'priority_changed',
      order.order_status,
      order.order_status,
      `Priority changed from ${order.order_priority} to ${priority}`,
    );
  }
  res.json({ order: fetchOrder(order.order_id) });
});

app.post('/api/orders', requireAuth, allowRoles('sales', 'supervisor', 'admin'), allowOrderCrud, (req, res) => {
  const isDirectPrintingOrder = hasAdminAccess(req.user) && req.body?.direct_print === true;
  if (isDirectPrintingOrder) {
    const businessUnit = req.body?.business_unit;
    if (!['dtf', 'sublimation'].includes(businessUnit)) {
      return res.status(400).json({ error: 'Select DTF or Sublimation for this print order' });
    }
    const {
      client_id: clientId,
      team_project_name: projectName,
      due_date: dueDate,
      print_mode: printMode,
      printer_machine: printerMachine,
      pressing_machine: pressingMachine,
    } = req.body ?? {};
    if (![clientId, projectName, dueDate, printMode, printerMachine].every(
      (value) => typeof value === 'string' && value.trim(),
    )) {
      return res.status(400).json({
        error: 'Client, name, date, print mode, and printer brand are required',
      });
    }
    if (!['print_only', 'print_press'].includes(printMode)) {
      return res.status(400).json({ error: 'Select Print only or Print and press' });
    }
    if (printMode === 'print_press' &&
      (typeof pressingMachine !== 'string' || !pressingMachine.trim())) {
      return res.status(400).json({
        error: 'Select a configured pressing machine for Print and press orders',
      });
    }
    const parsedDueDate = /^\d{4}-\d{2}-\d{2}$/.test(dueDate)
      ? new Date(`${dueDate}T00:00:00Z`)
      : null;
    if (
      !parsedDueDate ||
      Number.isNaN(parsedDueDate.getTime()) ||
      parsedDueDate.toISOString().slice(0, 10) !== dueDate
    ) {
      return res.status(400).json({ error: 'Due date must be a valid YYYY-MM-DD date' });
    }
    const quantity = Number(req.body?.total_quantity);
    if (!Number.isInteger(quantity) || quantity <= 0) {
      return res.status(400).json({ error: 'Order quantity must be a positive whole number' });
    }
    const client = db.prepare('SELECT * FROM clients WHERE client_id = ?').get(clientId);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    if (!db.prepare(
      "SELECT 1 FROM machines WHERE category = 'printing' AND name = ? AND (business_unit IS NULL OR business_unit = ?)",
    ).get(printerMachine, businessUnit)) {
      return res.status(400).json({ error: 'Select a configured printer brand for this order type' });
    }
    if (
      printMode === 'print_press' &&
      !db.prepare("SELECT 1 FROM machines WHERE category = 'pressing' AND name = ?").get(pressingMachine)
    ) {
      return res.status(400).json({ error: 'Select a configured pressing machine' });
    }
    const orderId = createId('ORD', 'orders', 'order_id');
    db.prepare(
      `INSERT INTO orders (
        order_id, client_id, sales_rep_id, team_project_name, due_date,
        style_number, category, style, cut, neck, sleeve_type, sleeve_length,
        fabric, gsm, item, total_quantity, business_unit, print_mode,
        printer_machine, pressing_machine, order_status
      ) VALUES (?, ?, ?, ?, ?, ?, ?, 'PRINT', 'N/A', 'N/A',
        'N/A', 'N/A', 'N/A', 1, ?, ?, ?, ?, ?, ?, 'printing')`,
    ).run(
      orderId,
      clientId,
      client.sales_rep_id,
      projectName.trim(),
      dueDate,
      businessUnit === 'dtf' ? 'DTF-ORDER' : 'SUBLI-ORDER',
      businessUnit === 'dtf' ? 'DTF' : 'Sublimation',
      `${businessUnit === 'dtf' ? 'DTF' : 'Sublimation'} ${printMode === 'print_only' ? 'print only' : 'print and press'}`,
      quantity,
      businessUnit,
      printMode,
      printerMachine,
      printMode === 'print_press' ? pressingMachine : null,
    );
    recordOrderEvent(
      orderId,
      req.user.id,
      'created',
      null,
      'printing',
      `${req.user.role === 'supervisor' ? 'Supervisor' : 'Admin'} order · ${businessUnit} · ${printMode}`,
    );
    return res.status(201).json({ order: fetchOrder(orderId) });
  }
  const productStyle = req.body?.style_option_ids
    ? resolveProductStyle(req.body.style_option_ids, req.body)
    : null;
  if (productStyle?.error) return res.status(400).json({ error: productStyle.error });
  const fields = [
    'client_id',
    'team_project_name',
    'due_date',
    'fabric',
    'fabric_width',
    'gsm',
  ];
  if (!productStyle) {
    fields.push('style_number', 'category', 'style', 'cut', 'neck', 'sleeve_type', 'sleeve_length');
  }
  for (const field of fields) {
    if (req.body?.[field] === undefined || String(req.body[field]).trim() === '') {
      return res.status(400).json({ error: `${field} is required` });
    }
  }

  const { client_id, assigned_artist_id, printer_machine, pressing_machine, cutting_machine } =
    req.body;
  const parsedDueDate =
    typeof req.body.due_date === 'string' &&
    /^\d{4}-\d{2}-\d{2}$/.test(req.body.due_date)
      ? new Date(`${req.body.due_date}T00:00:00Z`)
      : null;
  if (
    !parsedDueDate ||
    Number.isNaN(parsedDueDate.getTime()) ||
    parsedDueDate.toISOString().slice(0, 10) !== req.body.due_date
  ) {
    return res.status(400).json({ error: 'Due date must be a valid YYYY-MM-DD date' });
  }
  const layoutFee = Number(req.body.layout_fee ?? 0);
  const downPayment = Number(req.body.down_payment ?? 0);
  const layoutFeePaid = layoutFee >= PAID_LAYOUT_FEE_THRESHOLD ? 1 : 0;
  if (
    !Number.isFinite(layoutFee) ||
    layoutFee < 0 ||
    !Number.isFinite(downPayment) ||
    downPayment < 0
  ) {
    return res.status(400).json({ error: 'Layout fee and down payment must be non-negative amounts' });
  }
  if (layoutFee > 0 && layoutFee < MINIMUM_LAYOUT_FEE) {
    return res.status(400).json({ error: `Layout fee must be 0 or at least ${MINIMUM_LAYOUT_FEE}` });
  }
  const perPiecePrice = req.body.per_piece_price == null || req.body.per_piece_price === ''
    ? null
    : Number(req.body.per_piece_price);
  if (
    (layoutFee > 0 && (perPiecePrice == null || !Number.isFinite(perPiecePrice) || perPiecePrice < 0)) ||
    (perPiecePrice != null && (!Number.isFinite(perPiecePrice) || perPiecePrice < 0))
  ) {
    return res.status(400).json({ error: 'Enter a valid non-negative per-piece price when the layout fee is paid' });
  }
  const client = db.prepare('SELECT * FROM clients WHERE client_id = ?').get(client_id);
  if (!client || (req.user.role === 'sales' && client.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Client not found' });
  }
  if (!productStyle && (
    !['Sports', 'Corporate', 'Casual'].includes(req.body.category) ||
    !['Mens', 'Womens', 'Kids', 'Unisex'].includes(req.body.cut) ||
    !['V-Neck', 'Round Neck', 'Polo Collar'].includes(req.body.neck) ||
    !['Raglan', 'Set-in'].includes(req.body.sleeve_type) ||
    !['Sleeveless', 'Short Sleeve', 'Long Sleeve'].includes(req.body.sleeve_length)
  )) {
    return res.status(400).json({ error: 'One or more order selections are invalid' });
  }
  if (!Number.isInteger(Number(req.body.gsm)) || Number(req.body.gsm) <= 0) {
    return res.status(400).json({ error: 'GSM must be a positive whole number' });
  }
  if (
    String(req.body.fabric_width).trim().length > 50 ||
    !isOrderFabricAvailable(
      String(req.body.fabric).trim(),
      String(req.body.fabric_width).trim(),
      req.body.gsm,
    )
  ) {
    return res.status(400).json({ error: 'Select a fabric, width, and GSM combination from Fabric inventory' });
  }

  const artistId = layoutFeePaid && assigned_artist_id ? Number(assigned_artist_id) : null;
  if (layoutFeePaid) {
    if (!artistId) {
      return res.status(400).json({ error: 'Assign an Artist when the layout fee is at least 1' });
    }
    if (!printer_machine || !pressing_machine) {
      return res.status(400).json({
        error: 'Select both a printer machine and a pressing machine when the layout fee is at least 1',
      });
    }
    const artist = db
      .prepare("SELECT id FROM users WHERE id = ? AND role = 'artist'")
      .get(artistId);
    if (!artist) return res.status(400).json({ error: 'Assigned artist is invalid' });
  }
  for (const [category, name] of Object.entries({
    printing: layoutFeePaid ? printer_machine : null,
    pressing: layoutFeePaid ? pressing_machine : null,
    cutting: layoutFeePaid ? cutting_machine : null,
  })) {
    if (
      name &&
      !db.prepare('SELECT 1 FROM machines WHERE category = ? AND name = ?').get(category, name)
    ) {
      return res.status(400).json({ error: `Select a configured ${category} machine` });
    }
  }

  const salesRepId =
    req.user.role === 'sales'
      ? req.user.id
      : Number(req.body.sales_rep_id || (hasAdminAccess(req.user) ? client.sales_rep_id : req.user.id));
  if (salesRepId !== req.user.id) {
    const salesRep = db
      .prepare("SELECT id FROM users WHERE id = ? AND role = 'sales'")
      .get(salesRepId);
    if (!salesRep) {
      return res.status(400).json({ error: 'Assigned sales representative is invalid' });
    }
  }

  const orderId = createId('ORD', 'orders', 'order_id');
  db.prepare(
    `INSERT INTO orders (
      order_id, client_id, sales_rep_id, team_project_name, due_date, style_number,
      category, style, cut, neck, sleeve_type, sleeve_length, fabric, fabric_width, gsm, item,
      other_details, layout_fee, layout_fee_paid, per_piece_price, down_payment, assigned_artist_id, printer_machine,
      pressing_machine, cutting_machine, business_unit, print_mode
    ) VALUES (
      @order_id, @client_id, @sales_rep_id, @team_project_name, @due_date, @style_number,
      @category, @style, @cut, @neck, @sleeve_type, @sleeve_length, @fabric, @fabric_width, @gsm, @item,
      @other_details, @layout_fee, @layout_fee_paid, @per_piece_price, @down_payment, @assigned_artist_id,
      @printer_machine, @pressing_machine, @cutting_machine, 'sublimation', 'print_press'
    )`,
  ).run({
    order_id: orderId,
    client_id,
    sales_rep_id: salesRepId,
    team_project_name: String(req.body.team_project_name).trim(),
    due_date: req.body.due_date,
    style_number: productStyle?.style_number ?? String(req.body.style_number).trim(),
    category: productStyle?.values.category ?? req.body.category,
    style: productStyle?.values.style ?? String(req.body.style).trim(),
    cut: productStyle?.values.cut ?? req.body.cut,
    neck: productStyle?.values.neck ?? req.body.neck,
    sleeve_type: productStyle?.values.sleeve_type ?? req.body.sleeve_type,
    sleeve_length: productStyle?.values.sleeve_length ?? req.body.sleeve_length,
    fabric: String(req.body.fabric).trim(),
    fabric_width: String(req.body.fabric_width).trim(),
    gsm: Number(req.body.gsm),
    item: deriveItem(
      productStyle?.values.category ?? req.body.category,
      productStyle?.values.neck ?? req.body.neck,
    ),
    other_details: productStyle?.values.other_details ?? (req.body.other_details || null),
    layout_fee: layoutFee,
    layout_fee_paid: layoutFeePaid,
    per_piece_price: perPiecePrice,
    down_payment: layoutFeePaid ? downPayment : 0,
    assigned_artist_id: layoutFeePaid ? artistId : null,
    printer_machine: layoutFeePaid ? printer_machine || null : null,
    pressing_machine: layoutFeePaid ? pressing_machine || null : null,
    cutting_machine: layoutFeePaid ? cutting_machine || null : null,
  });
  recordOrderEvent(orderId, req.user.id, 'created', null, 'mockup');
  res.status(201).json({ order: fetchOrder(orderId) });
});

app.put('/api/orders/:orderId', requireAuth, allowRoles('sales'), allowOrderCrud, (req, res) => {
  const order = fetchOrder(req.params.orderId);
  if (!order || (req.user.role === 'sales' && order.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Order not found' });
  }

  const productStyle = req.body?.style_option_ids
    ? resolveProductStyle(req.body.style_option_ids, req.body, order)
    : null;
  if (productStyle?.error) return res.status(400).json({ error: productStyle.error });
  const fields = [
    'client_id',
    'team_project_name',
    'due_date',
    'fabric',
    'fabric_width',
    'gsm',
  ];
  if (!productStyle) {
    fields.push('style_number', 'category', 'style', 'cut', 'neck', 'sleeve_type', 'sleeve_length');
  }
  for (const field of fields) {
    if (req.body?.[field] === undefined || String(req.body[field]).trim() === '') {
      return res.status(400).json({ error: `${field} is required` });
    }
  }

  const salesRepId = hasAdminAccess(req.user)
    ? Number(req.body.sales_rep_id || order.sales_rep_id)
    : order.sales_rep_id;
  const salesRep = db
    .prepare("SELECT id FROM users WHERE id = ? AND role = 'sales'")
    .get(salesRepId);
  if (!salesRep) {
    return res.status(400).json({ error: 'Select a valid Sales representative' });
  }

  const client = db.prepare('SELECT * FROM clients WHERE client_id = ?').get(req.body.client_id);
  if (
    !client ||
    (req.user.role === 'sales' && client.sales_rep_id !== req.user.id)
  ) {
    return res.status(400).json({ error: 'Select a valid client for this order' });
  }

  const parsedDueDate =
    typeof req.body.due_date === 'string' &&
    /^\d{4}-\d{2}-\d{2}$/.test(req.body.due_date)
      ? new Date(`${req.body.due_date}T00:00:00Z`)
      : null;
  if (
    !parsedDueDate ||
    Number.isNaN(parsedDueDate.getTime()) ||
    parsedDueDate.toISOString().slice(0, 10) !== req.body.due_date
  ) {
    return res.status(400).json({ error: 'Due date must be a valid YYYY-MM-DD date' });
  }
  if (!productStyle && (
    !['Sports', 'Corporate', 'Casual'].includes(req.body.category) ||
    !['Mens', 'Womens', 'Kids', 'Unisex'].includes(req.body.cut) ||
    !['V-Neck', 'Round Neck', 'Polo Collar'].includes(req.body.neck) ||
    !['Raglan', 'Set-in'].includes(req.body.sleeve_type) ||
    !['Sleeveless', 'Short Sleeve', 'Long Sleeve'].includes(req.body.sleeve_length)
  )) {
    return res.status(400).json({ error: 'One or more order selections are invalid' });
  }
  const gsm = Number(req.body.gsm);
  const layoutFee = Number(req.body.layout_fee ?? 0);
  const downPayment = Number(req.body.down_payment ?? 0);
  const layoutFeePaid = layoutFee >= PAID_LAYOUT_FEE_THRESHOLD ? 1 : 0;
  if (!Number.isInteger(gsm) || gsm <= 0) {
    return res.status(400).json({ error: 'GSM must be a positive whole number' });
  }
  if (
    String(req.body.fabric_width).trim().length > 50 ||
    !isOrderFabricAvailable(
      String(req.body.fabric).trim(),
      String(req.body.fabric_width).trim(),
      gsm,
      order,
    )
  ) {
    return res.status(400).json({ error: 'Select a fabric, width, and GSM combination from Fabric inventory' });
  }
  if (
    !Number.isFinite(layoutFee) ||
    layoutFee < 0 ||
    !Number.isFinite(downPayment) ||
    downPayment < 0
  ) {
    return res.status(400).json({ error: 'Layout fee and down payment must be non-negative amounts' });
  }
  if (layoutFee > 0 && layoutFee < MINIMUM_LAYOUT_FEE) {
    return res.status(400).json({ error: `Layout fee must be 0 or at least ${MINIMUM_LAYOUT_FEE}` });
  }
  const perPiecePrice = req.body.per_piece_price == null || req.body.per_piece_price === ''
    ? null
    : Number(req.body.per_piece_price);
  if (
    (layoutFee > 0 && (perPiecePrice == null || !Number.isFinite(perPiecePrice) || perPiecePrice < 0)) ||
    (perPiecePrice != null && (!Number.isFinite(perPiecePrice) || perPiecePrice < 0))
  ) {
    return res.status(400).json({ error: 'Enter a valid non-negative per-piece price when the layout fee is paid' });
  }

  const assignedArtistId = layoutFeePaid
    ? Number(req.body.assigned_artist_id)
    : null;
  const downPaymentToSave = layoutFeePaid ? downPayment : 0;
  if (layoutFeePaid) {
    if (!assignedArtistId) {
      return res.status(400).json({ error: 'Assign an Artist when the layout fee is at least 1' });
    }
    if (!req.body.printer_machine || !req.body.pressing_machine) {
      return res.status(400).json({
        error: 'Select both a printer machine and a pressing machine when the layout fee is at least 1',
      });
    }
    const artist = db
      .prepare("SELECT id FROM users WHERE id = ? AND role = 'artist'")
      .get(assignedArtistId);
    if (!artist) return res.status(400).json({ error: 'Assigned artist is invalid' });
  }
  const machineValues = layoutFeePaid
    ? {
        printing: req.body.printer_machine,
        pressing: req.body.pressing_machine,
        cutting: req.body.cutting_machine,
      }
    : {
        printing: null,
        pressing: null,
        cutting: null,
      };
  for (const [category, name] of Object.entries(machineValues)) {
    if (
      name &&
      !db.prepare('SELECT 1 FROM machines WHERE category = ? AND name = ?').get(category, name)
    ) {
      return res.status(400).json({ error: `Select a configured ${category} machine` });
    }
  }

  db.prepare(
    `UPDATE orders SET
      client_id = ?, sales_rep_id = ?, team_project_name = ?, due_date = ?, style_number = ?,
      category = ?, style = ?, cut = ?, neck = ?, sleeve_type = ?, sleeve_length = ?,
      fabric = ?, fabric_width = ?, gsm = ?, item = ?, other_details = ?, layout_fee = ?, layout_fee_paid = ?, per_piece_price = ?, down_payment = ?,
      assigned_artist_id = ?, printer_machine = ?, pressing_machine = ?, cutting_machine = ?,
      updated_at = CURRENT_TIMESTAMP
     WHERE order_id = ?`,
  ).run(
    client.client_id,
    salesRepId,
    String(req.body.team_project_name).trim(),
    req.body.due_date,
    productStyle?.style_number ?? String(req.body.style_number).trim(),
    productStyle?.values.category ?? req.body.category,
    productStyle?.values.style ?? String(req.body.style).trim(),
    productStyle?.values.cut ?? req.body.cut,
    productStyle?.values.neck ?? req.body.neck,
    productStyle?.values.sleeve_type ?? req.body.sleeve_type,
    productStyle?.values.sleeve_length ?? req.body.sleeve_length,
    String(req.body.fabric).trim(),
    String(req.body.fabric_width).trim(),
    gsm,
    deriveItem(
      productStyle?.values.category ?? req.body.category,
      productStyle?.values.neck ?? req.body.neck,
    ),
    productStyle?.values.other_details ?? (typeof req.body.other_details === 'string' ? req.body.other_details.trim() || null : null),
    layoutFee,
    layoutFeePaid,
    perPiecePrice,
    downPaymentToSave,
    assignedArtistId,
    machineValues.printing || null,
    machineValues.pressing || null,
    machineValues.cutting || null,
    order.order_id,
  );
  recordOrderEvent(
    order.order_id,
    req.user.id,
    'details_updated',
    order.order_status,
    order.order_status,
  );
  res.json({ order: fetchOrder(order.order_id) });
});

app.put('/api/orders/:orderId/payments', requireAuth, allowRoles('sales'), allowOrderCrud, (req, res) => {
  const order = fetchOrder(req.params.orderId);
  if (!order || (req.user.role === 'sales' && order.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Order not found' });
  }
  if (order.order_status !== 'sales_approval') {
    return res.status(409).json({ error: 'Payments can only be updated while Sales approval is in progress' });
  }
  const layoutFee = Number(req.body?.layout_fee);
  const downPayment = Number(req.body?.down_payment);
  if (
    !Number.isFinite(layoutFee) ||
    layoutFee < MINIMUM_LAYOUT_FEE ||
    !Number.isFinite(downPayment) ||
    downPayment < 0
  ) {
    return res.status(400).json({
      error: `Layout fee must remain at least ${MINIMUM_LAYOUT_FEE}, and down payment must be a non-negative amount`,
    });
  }
  db.prepare(
    `UPDATE orders SET layout_fee = ?, layout_fee_paid = 1, down_payment = ?,
       updated_at = CURRENT_TIMESTAMP
     WHERE order_id = ?`,
  ).run(layoutFee, downPayment, order.order_id);
  res.json({ order: fetchOrder(order.order_id) });
});

app.delete('/api/orders/:orderId', requireAuth, allowRoles('sales'), allowOrderCrud, (req, res) => {
  const order = fetchOrder(req.params.orderId);
  if (!order || (req.user.role === 'sales' && order.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Order not found' });
  }
  const reprintImages = db
    .prepare('SELECT image_path FROM reprint_jobs WHERE order_id = ? AND image_path IS NOT NULL')
    .all(order.order_id)
    .map((job) => job.image_path);
  const testPrintMockups = db
    .prepare('SELECT mockup_path FROM test_print_requests WHERE order_id = ?')
    .all(order.order_id)
    .map((job) => job.mockup_path);
  db.transaction(() => {
    db.prepare('DELETE FROM orders WHERE order_id = ?').run(order.order_id);
  })();
  for (const filename of [
    order.spreadsheet_path,
    order.mockup_path,
    order.swatch_path,
    ...reprintImages,
    ...testPrintMockups,
  ]) {
    if (filename && path.basename(filename) === filename) {
      fs.rmSync(path.join(uploadDir, filename), { force: true });
    }
  }
  res.status(204).end();
});

app.post('/api/lineup/preview', requireAuth, allowRoles('sales'), upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Select a spreadsheet file' });
  const ext = path.extname(req.file.originalname).toLowerCase();
  if (!['.csv', '.xlsx'].includes(ext)) {
    return res.status(400).json({ error: 'Lineup must be a CSV or XLSX spreadsheet' });
  }
  try {
    const table = await parseLineup(req.file.buffer, ext, { requirePositiveQuantity: false });
    res.json({ ...table, filename: req.file.originalname, extension: ext });
  } catch (error) {
    res.status(400).json({ error: `Invalid lineup spreadsheet: ${error.message}` });
  }
});

app.get('/api/orders/:orderId/lineup', requireAuth, allowRoles('sales', 'artist', 'print_operator'), async (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  const assignedArtistCanView = req.user.role === 'artist' &&
    order.assigned_artist_id === req.user.id;
  if (
    req.user.role !== 'print_operator' &&
    !canManageLineup(req.user, order) &&
    !assignedArtistCanView
  ) {
    return res.status(403).json({ error: 'You cannot access this order lineup' });
  }
  if (!order.spreadsheet_path || path.basename(order.spreadsheet_path) !== order.spreadsheet_path) {
    return res.status(404).json({ error: 'No lineup spreadsheet is attached to this order' });
  }
  const filepath = path.join(uploadDir, order.spreadsheet_path);
  if (!fs.existsSync(filepath)) {
    return res.status(404).json({ error: 'The attached lineup spreadsheet could not be found' });
  }
  const ext = path.extname(filepath).toLowerCase();
  if (!['.csv', '.xlsx'].includes(ext)) {
    return res.status(400).json({ error: 'The attached lineup is not a supported spreadsheet' });
  }
  try {
    const table = await parseLineup(fs.readFileSync(filepath), ext);
    res.json({ ...table, filename: path.basename(filepath), extension: ext });
  } catch (error) {
    res.status(400).json({ error: `Invalid lineup spreadsheet: ${error.message}` });
  }
});

app.post('/api/orders/:orderId/lineup', requireAuth, allowRoles('sales', 'artist'), upload.single('file'), async (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  if (!['mockup', 'sales_approval', 'sizing'].includes(order.order_status)) {
    return res.status(409).json({ error: 'The lineup can only be updated before printing begins' });
  }
  if (Number(order.down_payment) < PAID_LAYOUT_FEE_THRESHOLD) {
    return res.status(409).json({
      error: 'The down payment must be at least 1 before uploading the lineup',
    });
  }
  if (!canManageLineup(req.user, order)) {
    return res.status(403).json({ error: 'You cannot update this order lineup' });
  }
  if (!req.file) return res.status(400).json({ error: 'Select a spreadsheet file' });

  const ext = path.extname(req.file.originalname).toLowerCase();
  if (!['.csv', '.xlsx'].includes(ext)) {
    return res.status(400).json({ error: 'Lineup must be a CSV or XLSX spreadsheet' });
  }
  let lineup;
  try {
    lineup = await parseLineup(req.file.buffer, ext);
  } catch (error) {
    return res.status(400).json({ error: `Invalid lineup spreadsheet: ${error.message}` });
  }
  const filename = saveUpload(req.file);
  db.prepare(
    'UPDATE orders SET spreadsheet_path = ?, total_quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE order_id = ?',
  ).run(filename, lineup.totalQuantity, order.order_id);
  if (order.spreadsheet_path) {
    fs.rmSync(path.join(uploadDir, order.spreadsheet_path), { force: true });
  }
  res.json({ total_quantity: lineup.totalQuantity, order: fetchOrder(order.order_id) });
});

app.put('/api/orders/:orderId/lineup', requireAuth, allowRoles('sales', 'artist'), async (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  if (!['mockup', 'sales_approval', 'sizing'].includes(order.order_status)) {
    return res.status(409).json({ error: 'The lineup can only be edited before printing begins' });
  }
  if (Number(order.down_payment) < PAID_LAYOUT_FEE_THRESHOLD) {
    return res.status(409).json({
      error: 'The down payment must be at least 1 before editing the lineup',
    });
  }
  if (!canManageLineup(req.user, order)) {
    return res.status(403).json({ error: 'You cannot edit this order lineup' });
  }
  const existingPath = order.spreadsheet_path;
  if (existingPath && path.basename(existingPath) !== existingPath) {
    return res.status(404).json({ error: 'No valid lineup spreadsheet is attached to this order' });
  }
  const ext = existingPath
    ? path.extname(existingPath).toLowerCase()
    : String(req.body?.extension || '.xlsx').toLowerCase();
  if (!['.csv', '.xlsx'].includes(ext)) {
    return res.status(400).json({ error: 'The attached lineup is not a supported spreadsheet' });
  }
  let table;
  try {
    table = validateLineupTable(req.body?.headers, req.body?.rows, {
      requirePositiveQuantity: false,
    });
  } catch (error) {
    return res.status(400).json({ error: error.message });
  }
  const filepath = existingPath ? path.join(uploadDir, existingPath) : null;
  const filename = `${uuidv4()}${ext}`;
  const nextPath = path.join(uploadDir, filename);
  let contents;
  try {
    contents = await encodeLineupTable(table.headers, table.rows, ext);
    fs.writeFileSync(nextPath, contents, { flag: 'wx' });
    db.prepare(
      'UPDATE orders SET spreadsheet_path = ?, total_quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE order_id = ?',
    ).run(filename, table.totalQuantity, order.order_id);
  } catch (error) {
    fs.rmSync(nextPath, { force: true });
    console.error('Unable to save edited lineup:', error);
    return res.status(500).json({ error: 'Unable to save the edited lineup spreadsheet' });
  }
  if (filepath) fs.rmSync(filepath, { force: true });
  res.json({ ...table, filename, extension: ext, order: fetchOrder(order.order_id) });
});

app.post(
  '/api/orders/:orderId/design',
  requireAuth,
  allowRoles('artist'),
  upload.single('mockup'),
  (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  if (order.order_status !== 'mockup') {
    return res.status(409).json({ error: 'This order is not awaiting a mock-up' });
  }
  const mockup = req.file;
  if (!mockup) {
    return res.status(400).json({ error: 'Select a concept mock-up file' });
  }

  const mockupExt = path.extname(mockup.originalname).toLowerCase();
  const designTypes = ['.png', '.jpg', '.jpeg', '.pdf'];
  if (!designTypes.includes(mockupExt)) {
    return res.status(400).json({ error: 'Mock-up file must be PNG, JPG, JPEG, or PDF' });
  }
  const mockupFilename = saveUpload(mockup);
  try {
    db.transaction(() => {
      db.prepare(
        `UPDATE orders
         SET mockup_path = ?, order_status = 'sales_approval',
             updated_at = CURRENT_TIMESTAMP
         WHERE order_id = ?`,
      ).run(mockupFilename, order.order_id);
      db.prepare(
        'INSERT INTO test_print_requests (order_id, artist_id, mockup_path) VALUES (?, ?, ?)',
      ).run(order.order_id, req.user.id, mockupFilename);
    })();
  } catch (error) {
    fs.rmSync(path.join(uploadDir, mockupFilename), { force: true });
    throw error;
  }
  const oldMockupIsNeeded = order.mockup_path && db.prepare(
    `SELECT 1 FROM test_print_requests
     WHERE mockup_path = ? AND status != 'complete' LIMIT 1`,
  ).get(order.mockup_path);
  if (order.mockup_path && !oldMockupIsNeeded && path.basename(order.mockup_path) === order.mockup_path) {
    fs.rmSync(path.join(uploadDir, order.mockup_path), { force: true });
  }
  recordOrderEvent(order.order_id, req.user.id, 'mockup_uploaded', 'mockup', 'sales_approval');
  res.json({ order: fetchOrder(order.order_id) });
  },
);

app.post('/api/orders/:orderId/swatch', requireAuth, allowRoles('sales'), upload.single('file'), (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  if (order.order_status !== 'sales_approval') {
    return res.status(409).json({ error: 'A swatch can only be uploaded while the order awaits Sales approval' });
  }
  if (!req.file) return res.status(400).json({ error: 'Select a swatch file' });
  const ext = path.extname(req.file.originalname).toLowerCase();
  if (!['.png', '.jpg', '.jpeg', '.pdf'].includes(ext)) {
    return res.status(400).json({ error: 'Swatch file must be PNG, JPG, JPEG, or PDF' });
  }
  const filename = saveUpload(req.file);
  db.prepare(
    'UPDATE orders SET swatch_path = ?, updated_at = CURRENT_TIMESTAMP WHERE order_id = ?',
  ).run(filename, order.order_id);
  if (order.swatch_path && path.basename(order.swatch_path) === order.swatch_path) {
    fs.rmSync(path.join(uploadDir, order.swatch_path), { force: true });
  }
  res.json({ order: fetchOrder(order.order_id) });
});

app.get('/api/orders/:orderId/files/:kind', requireAuth, (req, res) => {
  const order = getOrderForUser(req, res, { allowReprint: true });
  if (!order) return;
  const file = req.params.kind === 'mockup'
    ? order.mockup_path
    : req.params.kind === 'swatch'
      ? order.swatch_path
    : req.params.kind === 'lineup'
      ? order.spreadsheet_path
    : req.params.kind === 'production-spreadsheet'
      ? order.production_spreadsheet_path
      : null;
  if (!file || path.basename(file) !== file) {
    return res.status(404).json({ error: 'File not found' });
  }
  const filePath = path.join(uploadDir, file);
  if (req.params.kind === 'production-spreadsheet') {
    return res.download(filePath, `production-${order.order_id}.xlsx`, (error) => {
      if (error && !res.headersSent) res.status(error.statusCode || 404).end();
    });
  }
  res.sendFile(filePath, (error) => {
    if (error && !res.headersSent) res.status(error.statusCode || 404).end();
  });
});

app.get('/api/orders/:orderId/print-data', requireAuth, async (req, res) => {
  const order = getOrderForUser(req, res, { allowReprint: true });
  if (!order) return;

  const client = db
    .prepare(
      `SELECT company_name, contact_person, contact_number, shipping_address
       FROM clients WHERE client_id = ?`,
    )
    .get(order.client_id);
  const artist = order.assigned_artist_id
    ? db.prepare('SELECT name FROM users WHERE id = ?').get(order.assigned_artist_id)
    : null;
  let lineup = { headers: [], rows: [] };
  if (order.spreadsheet_path) {
    if (path.basename(order.spreadsheet_path) !== order.spreadsheet_path) {
      return res.status(400).json({ error: 'The attached lineup path is invalid' });
    }
    const filepath = path.join(uploadDir, order.spreadsheet_path);
    if (!fs.existsSync(filepath)) {
      return res.status(404).json({ error: 'The attached lineup spreadsheet could not be found' });
    }
    try {
      const extension = path.extname(filepath).toLowerCase();
      if (!['.csv', '.xlsx'].includes(extension)) {
        return res.status(400).json({ error: 'The attached lineup is not a supported spreadsheet' });
      }
      lineup = await parseLineup(fs.readFileSync(filepath), extension);
    } catch (error) {
      return res.status(400).json({ error: `Invalid lineup spreadsheet: ${error.message}` });
    }
  }
  res.json({
    order: publicOrder(order, req.user),
    client,
    artist_name: artist?.name || '',
    lineup,
  });
});

app.post('/api/orders/:orderId/approve', requireAuth, allowRoles('sales'), (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  if (order.order_status !== 'sales_approval') {
    return res.status(409).json({ error: 'This order is not awaiting approval' });
  }
  const approve = req.body?.approved === true;
  if (approve && !order.swatch_path) {
    return res.status(409).json({ error: 'Upload a swatch before approving the mock-up' });
  }
  if (approve && (!order.spreadsheet_path || order.total_quantity < 1)) {
    return res.status(409).json({ error: 'Upload a valid lineup before approving the mock-up' });
  }
  const nextStage = approve ? 'sizing' : 'mockup';
  db.prepare(
    'UPDATE orders SET order_status = ?, updated_at = CURRENT_TIMESTAMP WHERE order_id = ?',
  ).run(nextStage, order.order_id);
  recordOrderEvent(
    order.order_id,
    req.user.id,
    approve ? 'mockup_approved' : 'mockup_rejected',
    'sales_approval',
    nextStage,
    String(req.body?.notes || ''),
  );
  res.json({ order: fetchOrder(order.order_id) });
});

app.post('/api/orders/:orderId/complete-sizing', requireAuth, allowRoles('artist'), (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  if (order.order_status !== 'sizing') {
    return res.status(409).json({ error: 'This order is not awaiting sizing' });
  }
  const sizingNotes = String(req.body?.sizing_notes || '').trim();
  if (!sizingNotes) {
    return res.status(400).json({ error: 'Enter the completed layout and sizing notes' });
  }
  db.prepare(
    `UPDATE orders SET order_status = 'printing', sizing_notes = ?,
       updated_at = CURRENT_TIMESTAMP
     WHERE order_id = ?`,
  ).run(sizingNotes, order.order_id);
  recordOrderEvent(
    order.order_id,
    req.user.id,
    'sizing_completed',
    'sizing',
    'printing',
    sizingNotes,
  );
  res.json({ order: fetchOrder(order.order_id) });
});

function verifyStationAssignment(user, station) {
  return hasAdminAccess(user) || roleStations[user.role] === station;
}

app.post('/api/orders/:orderId/station-log', requireAuth, (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  const station = order.order_status;
  if (!stages.includes(station) || !verifyStationAssignment(req.user, station)) {
    return res.status(403).json({ error: 'Only the assigned station can log production data' });
  }

  const legacyPaperMeters = Number(req.body?.paper_meters || 0);
  const fabricUsed = order.business_unit === 'dtf' || order.print_mode === 'print_only'
    ? ''
    : String(req.body?.fabric_used || '').trim();
  const printerBrand = station === 'printing'
    ? String(req.body?.printer_brand || order.printer_machine || '').trim()
    : '';
  if (
    station === 'printing' &&
    !db.prepare(
      "SELECT 1 FROM machines WHERE category = 'printing' AND name = ? AND (business_unit IS NULL OR business_unit = ?)",
    ).get(printerBrand, order.business_unit || 'sublimation')
  ) {
    return res.status(400).json({ error: 'Select a configured printer brand' });
  }
  const defect = String(req.body?.defect_description || '').trim();
  if (!Number.isFinite(legacyPaperMeters) || legacyPaperMeters < 0) {
    return res.status(400).json({ error: 'Paper usage must be a non-negative number of meters' });
  }
  const rawConsumables = req.body?.consumables ?? [];
  if (!Array.isArray(rawConsumables) || rawConsumables.length > 7) {
    return res.status(400).json({ error: 'Production log contains an invalid consumables list' });
  }
  const consumables = [];
  const seenConsumables = new Set();
  for (const entry of rawConsumables) {
    const product = consumableCatalog[entry?.code];
    const quantity = Number(entry?.quantity);
    if (
      !product ||
      product.legacy ||
      seenConsumables.has(entry.code) ||
      !Number.isFinite(quantity) ||
      quantity <= 0
    ) {
      return res.status(400).json({ error: 'Enter valid positive quantities for the selected Printing consumables' });
    }
    seenConsumables.add(entry.code);
    consumables.push({ code: entry.code, quantity, product });
  }
  const consumablePaperMeters = consumables
    .filter(({ code }) => code.startsWith('subli_paper_'))
    .reduce((total, entry) => total + entry.quantity, 0);
  const paperMeters = consumables.length ? consumablePaperMeters : legacyPaperMeters;
  if (!fabricUsed && paperMeters === 0 && !defect && consumables.length === 0) {
    return res.status(400).json({ error: 'Enter fabric usage, paper usage, or describe a defect' });
  }

  const log = db.transaction(() => {
    for (const entry of consumables) {
      const onHand = db.prepare(
        `SELECT COALESCE(SUM(CASE
           WHEN transaction_type = 'stock_in' THEN quantity
           WHEN transaction_type = 'usage' THEN -quantity
           ELSE quantity
         END), 0) AS on_hand
         FROM consumable_transactions
         WHERE business_unit = ? AND consumable_code = ?`,
      ).get(entry.product.business_unit, entry.code).on_hand;
      if (onHand < entry.quantity) {
        return { error: `Not enough ${entry.product.label} in stock. Available: ${onHand} ${entry.product.unit}` };
      }
    }
    const result = db
      .prepare(
        `INSERT INTO station_logs
          (order_id, station, operator_id, fabric_used, printer_brand, paper_meters, defect_description)
         VALUES (?, ?, ?, ?, ?, ?, ?)`,
      )
      .run(order.order_id, station, req.user.id, fabricUsed, printerBrand, paperMeters, defect || null);
    const stationLogId = result.lastInsertRowid;
    for (const entry of consumables) {
      db.prepare(
        `INSERT INTO consumable_transactions
           (business_unit, consumable_code, transaction_type, quantity, order_id,
            station_log_id, entry_date, created_by)
         VALUES (?, ?, 'usage', ?, ?, ?, date('now', 'localtime'), ?)`,
      ).run(
        entry.product.business_unit,
        entry.code,
        entry.quantity,
        order.order_id,
        stationLogId,
        req.user.id,
      );
      db.prepare(
        `INSERT INTO production_log_items
           (station_log_id, consumable_code, label, quantity, unit)
         VALUES (?, ?, ?, ?, ?)`,
      ).run(
        stationLogId,
        entry.code,
        entry.product.label,
        entry.code.includes('_ink_') ? entry.quantity * 1000 : entry.quantity,
        entry.code.includes('_ink_') ? 'ml' : entry.product.unit,
      );
    }
    let reprint = null;
    if (defect) {
      const reprintResult = db
        .prepare(
          'INSERT INTO reprint_jobs (order_id, station_log_id, created_by) VALUES (?, ?, ?)',
        )
        .run(order.order_id, result.lastInsertRowid, req.user.id);
      reprint = db
        .prepare('SELECT * FROM reprint_jobs WHERE id = ?')
        .get(reprintResult.lastInsertRowid);
    }
    recordOrderEvent(
      order.order_id,
      req.user.id,
      defect ? 'defect_logged' : 'consumable_logged',
      station,
      station,
      defect,
    );
    return {
      log: db.prepare('SELECT * FROM station_logs WHERE id = ?').get(stationLogId),
      reprint,
    };
  });
  const result = log();
  if (result.error) return res.status(400).json({ error: result.error });
  res.status(201).json(result);
});

app.post('/api/orders/:orderId/reprint-log', requireAuth, upload.single('image'), (req, res) => {
  const order = getOrderForUser(req, res, { allowReprint: true });
  if (!order) return;
  const stageCanRequest =
    ['pressing', 'cutting', 'sewing', 'packing'].includes(order.order_status) &&
    (hasAdminAccess(req.user) || roleStations[req.user.role] === order.order_status);
  const lateRequestCanRequest =
    (hasAdminAccess(req.user) || req.user.role === 'print_operator') &&
    stageSequence.indexOf(order.order_status) >= stageSequence.indexOf('printing');
  const completedOrderCanRequest =
    order.order_status === 'complete' &&
    (hasAdminAccess(req.user) || ['sales', 'artist'].includes(req.user.role));
  if (!stageCanRequest && !lateRequestCanRequest && !completedOrderCanRequest) {
    return res.status(403).json({ error: 'Only the assigned production station or Printing can request a reprint' });
  }
  const hasReachedPrinting = db
    .prepare(
      `SELECT 1 FROM order_events
       WHERE order_id = ?
         AND to_stage IN ('printing', 'pressing', 'cutting', 'sewing', 'packing', 'complete')
       LIMIT 1`,
    )
    .get(order.order_id);
  if (!hasReachedPrinting) {
    return res.status(409).json({ error: 'The order must reach Printing before a reprint can be logged' });
  }

  const fabricUsed = order.business_unit === 'dtf' || order.print_mode === 'print_only'
    ? ''
    : String(req.body?.fabric_used || '').trim();
  const defect = String(req.body?.defect_description || '').trim();
  if ((!fabricUsed && order.business_unit !== 'dtf' && order.print_mode !== 'print_only') || !defect) {
    return res.status(400).json({ error: 'Enter the fabric when applicable and the reason for the reprint' });
  }
  const imageTypes = new Map([
    ['.png', 'image/png'],
    ['.jpg', 'image/jpeg'],
    ['.jpeg', 'image/jpeg'],
  ]);
  const imageExtension = req.file ? path.extname(req.file.originalname).toLowerCase() : '';
  if (req.file && imageTypes.get(imageExtension) !== req.file.mimetype) {
    return res.status(400).json({ error: 'Reprint image must be a PNG or JPEG file' });
  }
  const imagePath = req.file ? saveUpload(req.file) : null;

  try {
    const reprint = db.transaction(() => {
      const logResult = db
        .prepare(
          `INSERT INTO station_logs
            (order_id, station, operator_id, fabric_used, paper_meters, defect_description)
           VALUES (?, ?, ?, ?, ?, ?)`,
        )
        .run(order.order_id, order.order_status, req.user.id, fabricUsed, 0, defect);
      const jobResult = db
        .prepare(
          `INSERT INTO reprint_jobs (order_id, station_log_id, created_by, image_path)
           VALUES (?, ?, ?, ?)`,
        )
        .run(order.order_id, logResult.lastInsertRowid, req.user.id, imagePath);
      recordOrderEvent(
        order.order_id,
        req.user.id,
        'reprint_requested',
        order.order_status,
        order.order_status,
        defect,
      );
      return {
        log: db.prepare('SELECT * FROM station_logs WHERE id = ?').get(logResult.lastInsertRowid),
        reprint: db.prepare('SELECT * FROM reprint_jobs WHERE id = ?').get(jobResult.lastInsertRowid),
      };
    })();
    res.status(201).json(reprint);
  } catch (error) {
    if (imagePath) fs.rmSync(path.join(uploadDir, imagePath), { force: true });
    throw error;
  }
});

app.post('/api/orders/:orderId/complete-stage', requireAuth, async (req, res) => {
  const order = getOrderForUser(req, res);
  if (!order) return;
  const station = order.order_status;
  const completedPrintHandoff = station === 'pressing'
    ? db
        .prepare(
          `SELECT event_type FROM order_events
           WHERE order_id = ? AND to_stage = 'pressing'
             AND event_type IN ('stage_completed', 'reprint_completed', 'test_print_completed')
           ORDER BY id DESC LIMIT 1`,
        )
        .get(order.order_id)
    : null;
  const nextStage = station === 'pressing' &&
    completedPrintHandoff?.event_type === 'test_print_completed'
    ? 'sales_approval'
    : station === 'printing' && order.print_mode === 'print_only'
      ? 'complete'
      : stageNext[station];
  if (!nextStage || !verifyStationAssignment(req.user, station)) {
    return res.status(403).json({ error: 'Only the assigned station can complete this stage' });
  }
  const pressingNote = station === 'printing' &&
    order.business_unit !== 'dtf' &&
    order.print_mode !== 'print_only'
    ? String(req.body?.pressing_note || '').trim()
    : '';
  if (pressingNote.length > 1000) {
    return res.status(400).json({ error: 'The note for Pressing must be 1,000 characters or fewer' });
  }
  if (!['pressing', 'cutting', 'sewing', 'packing'].includes(station)) {
    const stationLog = db
      .prepare('SELECT id FROM station_logs WHERE order_id = ? AND station = ? LIMIT 1')
      .get(order.order_id, station);
    if (!stationLog) {
      return res.status(409).json({ error: 'Record station production data before completing this stage' });
    }
  }

  let productionSpreadsheetPath = null;
  const previousSpreadsheetPath = order.production_spreadsheet_path;
  try {
    if (station === 'printing') {
      const logs = db.prepare(
        `SELECT station_logs.id, station_logs.station, station_logs.fabric_used,
                station_logs.printer_brand,
                station_logs.paper_meters, station_logs.defect_description,
                station_logs.created_at, users.name AS operator
         FROM station_logs
         JOIN users ON users.id = station_logs.operator_id
         WHERE station_logs.order_id = ? AND station_logs.station = 'printing'
         ORDER BY station_logs.id`,
      ).all(order.order_id);
      const itemRows = db.prepare(
        `SELECT production_log_items.station_log_id, production_log_items.label,
                production_log_items.quantity, production_log_items.unit
         FROM production_log_items
         JOIN station_logs ON station_logs.id = production_log_items.station_log_id
         WHERE station_logs.order_id = ? AND station_logs.station = 'printing'
         ORDER BY production_log_items.id`,
      ).all(order.order_id);
      const itemsByLog = new Map();
      for (const item of itemRows) {
        const items = itemsByLog.get(item.station_log_id) || [];
        items.push(item);
        itemsByLog.set(item.station_log_id, items);
      }
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'BPSMS';
      workbook.subject = `Printing production log for ${order.order_id}`;
      const worksheet = workbook.addWorksheet('Printing Production');
      worksheet.addRow([
        'Order ID', 'Client', 'Project', 'Business unit', 'Print mode',
        'Recorded at', 'Operator', 'Printer brand', 'Input', 'Quantity', 'Unit',
        'Fabric used', 'Paper (m)', 'Defect / issue',
      ]);
      const clientName = db.prepare('SELECT company_name FROM clients WHERE client_id = ?')
        .get(order.client_id)?.company_name || '';
      for (const log of logs) {
        const logItems = itemsByLog.get(log.id) || [];
        const rows = logItems.length ? logItems : [{ label: '', quantity: '', unit: '' }];
        for (const item of rows) {
          worksheet.addRow([
            order.order_id,
            clientName,
            order.team_project_name,
            order.business_unit,
            order.print_mode,
            log.created_at,
            log.operator,
            log.printer_brand,
            item.label,
            item.quantity,
            item.unit,
            log.fabric_used,
            log.paper_meters,
            log.defect_description || '',
          ]);
        }
      }
      worksheet.getRow(1).font = { bold: true };
      worksheet.views = [{ state: 'frozen', ySplit: 1 }];
      worksheet.columns.forEach((column) => {
        column.width = Math.min(
          Math.max(...column.values.map((value) => String(value ?? '').length), 12) + 2,
          40,
        );
      });
      productionSpreadsheetPath = `${uuidv4()}.xlsx`;
      fs.writeFileSync(
        path.join(uploadDir, productionSpreadsheetPath),
        await workbook.xlsx.writeBuffer(),
        { flag: 'wx' },
      );
    }

    db.transaction(() => {
      db.prepare(
        `UPDATE orders
         SET order_status = ?,
             production_spreadsheet_path = COALESCE(?, production_spreadsheet_path),
             updated_at = CURRENT_TIMESTAMP
         WHERE order_id = ?`,
      ).run(nextStage, productionSpreadsheetPath, order.order_id);
      recordOrderEvent(
        order.order_id,
        req.user.id,
        'stage_completed',
        station,
        nextStage,
        pressingNote,
      );
    })();
  } catch (error) {
    if (productionSpreadsheetPath) {
      fs.rmSync(path.join(uploadDir, productionSpreadsheetPath), { force: true });
    }
    console.error('Unable to complete production stage:', error);
    return res.status(500).json({ error: 'Unable to save the production completion data' });
  }
  if (
    productionSpreadsheetPath &&
    previousSpreadsheetPath &&
    path.basename(previousSpreadsheetPath) === previousSpreadsheetPath
  ) {
    fs.rmSync(path.join(uploadDir, previousSpreadsheetPath), { force: true });
  }
  res.json({ order: fetchOrder(order.order_id) });
});

app.put('/api/orders/:orderId/packed-quantity', requireAuth, allowRoles('sales'), (req, res) => {
  const order = fetchOrder(req.params.orderId);
  if (!order || (req.user.role === 'sales' && order.sales_rep_id !== req.user.id)) {
    return res.status(404).json({ error: 'Order not found' });
  }
  if (order.order_status !== 'complete') {
    return res.status(409).json({ error: 'Packed quantity can only be updated for completed orders' });
  }
  const packedQuantity = Number(req.body?.packed_quantity);
  if (
    !Number.isInteger(packedQuantity) ||
    packedQuantity !== Number(order.total_quantity)
  ) {
    return res.status(400).json({
      error: `Packed pieces must equal the designated order quantity of ${order.total_quantity}`,
    });
  }
  const perPiecePrice = Number(req.body?.per_piece_price);
  if (!Number.isFinite(perPiecePrice) || perPiecePrice < 0) {
    return res.status(400).json({ error: 'Per-piece price must be a non-negative amount' });
  }
  db.transaction(() => {
    db.prepare(
      `UPDATE orders
       SET packed_quantity = ?, per_piece_price = ?,
           order_status = 'pickup_delivered', updated_at = CURRENT_TIMESTAMP
       WHERE order_id = ?`,
    ).run(packedQuantity, perPiecePrice, order.order_id);
    recordOrderEvent(
      order.order_id,
      req.user.id,
      'packed_quantity_updated',
      'complete',
      'pickup_delivered',
      `Packed pieces: ${packedQuantity}; per-piece price: ${perPiecePrice.toFixed(2)}; total amount: ${(perPiecePrice * packedQuantity).toFixed(2)}; balance after layout fee and down payment: ${(perPiecePrice * packedQuantity - Number(order.layout_fee || 0) - Number(order.down_payment || 0)).toFixed(2)}`,
    );
  })();
  res.json({ order: fetchOrder(order.order_id) });
});

app.get('/api/test-prints', requireAuth, (req, res) => {
  if (!verifyStationAssignment(req.user, 'printing')) {
    return res.status(403).json({ error: 'Only Printing can access test print requests' });
  }
  const jobs = db.prepare(
    `SELECT test_print_requests.id, test_print_requests.order_id, test_print_requests.artist_id,
            test_print_requests.status, test_print_requests.created_at,
          orders.team_project_name, orders.due_date, orders.printer_machine,
            test_print_requests.mockup_path,
            users.name AS artist_name
     FROM test_print_requests
     JOIN orders ON orders.order_id = test_print_requests.order_id
     JOIN users ON users.id = test_print_requests.artist_id
     WHERE test_print_requests.status != 'complete'
     ORDER BY test_print_requests.created_at`,
  ).all();
  res.json({ jobs });
});

app.get('/api/test-prints/:id/mockup', requireAuth, (req, res) => {
  if (!verifyStationAssignment(req.user, 'printing')) {
    return res.status(403).json({ error: 'Only Printing can access test print files' });
  }
  const request = db.prepare(
    'SELECT mockup_path FROM test_print_requests WHERE id = ?',
  ).get(req.params.id);
  if (!request?.mockup_path || path.basename(request.mockup_path) !== request.mockup_path) {
    return res.status(404).json({ error: 'Test print mock-up not found' });
  }
  const mockupPath = path.join(uploadDir, request.mockup_path);
  if (!fs.existsSync(mockupPath)) {
    return res.status(404).json({ error: 'Test print mock-up not found' });
  }
  res.sendFile(mockupPath, (error) => {
    if (error && !res.headersSent) res.status(error.statusCode || 404).end();
  });
});

app.post('/api/test-prints/:id/:action', requireAuth, (req, res) => {
  if (!verifyStationAssignment(req.user, 'printing')) {
    return res.status(403).json({ error: 'Only Printing can process test print requests' });
  }
  const request = db.prepare(
    'SELECT * FROM test_print_requests WHERE id = ?',
  ).get(req.params.id);
  if (!request) return res.status(404).json({ error: 'Test print request not found' });
  const action = req.params.action;
  const pressingNote = String(req.body?.pressing_note || '').trim();
  if (action === 'complete' && pressingNote.length > 1000) {
    return res.status(400).json({ error: 'The note for Pressing must be 1,000 characters or fewer' });
  }
  if (action === 'start' && request.status === 'open') {
    db.prepare("UPDATE test_print_requests SET status = 'in_progress' WHERE id = ?").run(request.id);
  } else if (action === 'complete' && ['open', 'in_progress'].includes(request.status)) {
    db.transaction(() => {
      db.prepare(
        `UPDATE test_print_requests SET status = 'complete', completed_at = CURRENT_TIMESTAMP
         WHERE id = ?`,
      ).run(request.id);
      handoffOrderToPressing(
        request.order_id,
        req.user.id,
        'test_print_completed',
        pressingNote,
      );
    })();
  } else {
    return res.status(409).json({ error: 'This test print action is not valid for its current status' });
  }
  if (action === 'complete') {
    const order = fetchOrder(request.order_id);
    if (order?.mockup_path !== request.mockup_path) {
      const activeRequest = db.prepare(
        `SELECT 1 FROM test_print_requests
         WHERE mockup_path = ? AND status != 'complete' LIMIT 1`,
      ).get(request.mockup_path);
      if (!activeRequest && path.basename(request.mockup_path) === request.mockup_path) {
        fs.rmSync(path.join(uploadDir, request.mockup_path), { force: true });
      }
    }
  }
  res.json({
    job: db.prepare('SELECT * FROM test_print_requests WHERE id = ?').get(request.id),
  });
});

app.get('/api/reprints', requireAuth, (req, res) => {
  if (!hasAdminAccess(req.user) && req.user.role !== 'print_operator') {
    return res.status(403).json({ error: 'Only Printing can access reprint jobs' });
  }
  const reprints = db
    .prepare(
      `SELECT reprint_jobs.*, station_logs.station, station_logs.fabric_used,
              station_logs.defect_description, station_logs.paper_meters,
             orders.team_project_name, orders.order_status, orders.printer_machine,
             orders.business_unit, orders.print_mode,
             orders.business_unit, orders.print_mode
       FROM reprint_jobs
       JOIN station_logs ON station_logs.id = reprint_jobs.station_log_id
       JOIN orders ON orders.order_id = reprint_jobs.order_id
       WHERE reprint_jobs.status != 'complete'
       ORDER BY reprint_jobs.created_at`,
    )
    .all();
  const logsByJob = new Map();
  if (reprints.length) {
    const logs = db
      .prepare(
        `SELECT station_logs.id, station_logs.reprint_job_id, station_logs.fabric_used,
                station_logs.paper_meters, station_logs.defect_description,
                station_logs.created_at, users.name AS operator
         FROM station_logs
         JOIN users ON users.id = station_logs.operator_id
         WHERE station_logs.reprint_job_id IN (${reprints.map(() => '?').join(', ')})
         ORDER BY station_logs.id DESC`,
      )
      .all(...reprints.map((job) => job.id));
    for (const log of logs) {
      const jobLogs = logsByJob.get(log.reprint_job_id) || [];
      jobLogs.push(log);
      logsByJob.set(log.reprint_job_id, jobLogs);
    }
  }
  res.json({
    reprints: reprints.map((job) => ({
      ...job,
      station_logs: logsByJob.get(job.id) || [],
    })),
  });
});

app.get('/api/reprints/:id/image', requireAuth, (req, res) => {
  if (!verifyStationAssignment(req.user, 'printing')) {
    return res.status(403).json({ error: 'Only Printing can view reprint images' });
  }
  const reprint = db.prepare('SELECT image_path FROM reprint_jobs WHERE id = ?').get(req.params.id);
  if (!reprint) return res.status(404).json({ error: 'Reprint job not found' });
  if (!reprint.image_path || path.basename(reprint.image_path) !== reprint.image_path) {
    return res.status(404).json({ error: 'Reprint image not found' });
  }
  const imagePath = path.join(uploadDir, reprint.image_path);
  if (!fs.existsSync(imagePath)) {
    return res.status(404).json({ error: 'Reprint image not found' });
  }
  res.type(path.extname(reprint.image_path)).sendFile(imagePath);
});

app.post('/api/reprints/:id/station-log', requireAuth, (req, res) => {
  if (!verifyStationAssignment(req.user, 'printing')) {
    return res.status(403).json({ error: 'Only Printing can log reprint production data' });
  }
  const reprint = db.prepare('SELECT * FROM reprint_jobs WHERE id = ?').get(req.params.id);
  if (!reprint) return res.status(404).json({ error: 'Reprint job not found' });
  if (reprint.status === 'complete') {
    return res.status(409).json({ error: 'Completed reprint jobs cannot receive production logs' });
  }
  const order = fetchOrder(reprint.order_id);
  const fabricUsed = order?.business_unit === 'dtf' || order?.print_mode === 'print_only'
    ? ''
    : String(req.body?.fabric_used || '').trim();
  const printerBrand = String(req.body?.printer_brand || order?.printer_machine || '').trim();
  const legacyPaperMeters = Number(req.body?.paper_meters || 0);
  const defect = String(req.body?.defect_description || '').trim();
  if (!Number.isFinite(legacyPaperMeters) || legacyPaperMeters < 0) {
    return res.status(400).json({ error: 'Paper usage must be a non-negative number of meters' });
  }
  if (!order || !db.prepare(
    "SELECT 1 FROM machines WHERE category = 'printing' AND name = ? AND (business_unit IS NULL OR business_unit = ?)",
  ).get(printerBrand, order.business_unit || 'sublimation')) {
    return res.status(400).json({ error: 'Select a configured printer brand' });
  }
  const rawConsumables = req.body?.consumables ?? [];
  if (!Array.isArray(rawConsumables) || rawConsumables.length > 7) {
    return res.status(400).json({ error: 'Production log contains an invalid consumables list' });
  }
  const consumables = [];
  const seenConsumables = new Set();
  for (const entry of rawConsumables) {
    const product = consumableCatalog[entry?.code];
    const quantity = Number(entry?.quantity);
    if (
      !product ||
      product.legacy ||
      seenConsumables.has(entry.code) ||
      !Number.isFinite(quantity) ||
      quantity <= 0
    ) {
      return res.status(400).json({ error: 'Enter valid positive quantities for the selected Printing consumables' });
    }
    seenConsumables.add(entry.code);
    consumables.push({ code: entry.code, quantity, product });
  }
  const paperMeters = consumables.length
    ? consumables
      .filter(({ code }) => code.startsWith('subli_paper_'))
      .reduce((total, entry) => total + entry.quantity, 0)
    : legacyPaperMeters;
  if (!fabricUsed && paperMeters === 0 && !defect && consumables.length === 0) {
    return res.status(400).json({ error: 'Enter fabric usage, paper usage, or describe a defect' });
  }

  const saveLog = db.transaction(() => {
    for (const entry of consumables) {
      const onHand = db.prepare(
        `SELECT COALESCE(SUM(CASE
           WHEN transaction_type = 'stock_in' THEN quantity
           WHEN transaction_type = 'usage' THEN -quantity
           ELSE quantity
         END), 0) AS on_hand
         FROM consumable_transactions
         WHERE business_unit = ? AND consumable_code = ?`,
      ).get(entry.product.business_unit, entry.code).on_hand;
      if (onHand < entry.quantity) {
        return { error: `Not enough ${entry.product.label} in stock. Available: ${onHand} ${entry.product.unit}` };
      }
    }
    const result = db
      .prepare(
        `INSERT INTO station_logs
          (order_id, station, operator_id, fabric_used, printer_brand, paper_meters, defect_description, reprint_job_id)
         VALUES (?, 'printing', ?, ?, ?, ?, ?, ?)`,
      )
      .run(reprint.order_id, req.user.id, fabricUsed, printerBrand, paperMeters, defect || null, reprint.id);
    for (const entry of consumables) {
      db.prepare(
        `INSERT INTO consumable_transactions
           (business_unit, consumable_code, transaction_type, quantity, order_id,
            station_log_id, entry_date, created_by)
         VALUES (?, ?, 'usage', ?, ?, ?, date('now', 'localtime'), ?)`,
      ).run(
        entry.product.business_unit,
        entry.code,
        entry.quantity,
        reprint.order_id,
        result.lastInsertRowid,
        req.user.id,
      );
      db.prepare(
        `INSERT INTO production_log_items
           (station_log_id, consumable_code, label, quantity, unit)
         VALUES (?, ?, ?, ?, ?)`,
      ).run(
        result.lastInsertRowid,
        entry.code,
        entry.product.label,
        entry.code.includes('_ink_') ? entry.quantity * 1000 : entry.quantity,
        entry.code.includes('_ink_') ? 'ml' : entry.product.unit,
      );
    }
    return { logId: result.lastInsertRowid };
  });
  const result = saveLog();
  if (result.error) return res.status(400).json({ error: result.error });
  recordOrderEvent(
    reprint.order_id,
    req.user.id,
    defect ? 'reprint_defect_logged' : 'reprint_production_logged',
    'printing',
    'printing',
    defect,
  );
  res.status(201).json({
    log: db
      .prepare(
        `SELECT station_logs.*, users.name AS operator
         FROM station_logs JOIN users ON users.id = station_logs.operator_id
         WHERE station_logs.id = ?`,
      )
      .get(result.logId),
  });
});

app.post('/api/reprints/:id/:action', requireAuth, (req, res) => {
  if (!verifyStationAssignment(req.user, 'printing')) {
    return res.status(403).json({ error: 'Only Printing can process reprints' });
  }
  const reprint = db.prepare('SELECT * FROM reprint_jobs WHERE id = ?').get(req.params.id);
  if (!reprint) return res.status(404).json({ error: 'Reprint job not found' });
  const order = fetchOrder(reprint.order_id);
  const action = req.params.action;
  const canHandoffToPressing = order?.print_mode !== 'print_only';
  const pressingNote = canHandoffToPressing && order?.business_unit !== 'dtf'
    ? String(req.body?.pressing_note || '').trim()
    : '';
  if (action === 'complete' && pressingNote.length > 1000) {
    return res.status(400).json({ error: 'The note for Pressing must be 1,000 characters or fewer' });
  }
  if (action === 'start' && reprint.status === 'open') {
    db.prepare("UPDATE reprint_jobs SET status = 'in_progress' WHERE id = ?").run(reprint.id);
  } else if (
    action === 'complete' &&
    ['open', 'in_progress'].includes(reprint.status)
  ) {
    db.transaction(() => {
      db.prepare(
        `UPDATE reprint_jobs SET status = 'complete', completed_at = CURRENT_TIMESTAMP
         WHERE id = ?`,
      ).run(reprint.id);
      if (canHandoffToPressing) {
        handoffOrderToPressing(
          reprint.order_id,
          req.user.id,
          'reprint_completed',
          pressingNote,
        );
      } else {
        recordOrderEvent(
          reprint.order_id,
          req.user.id,
          'reprint_completed',
          order.order_status,
          order.order_status,
        );
      }
    })();
  } else {
    return res.status(409).json({ error: 'This reprint action is not valid for its current status' });
  }
  res.json({ reprint: db.prepare('SELECT * FROM reprint_jobs WHERE id = ?').get(reprint.id) });
});

app.get('/api/reports', requireAuth, allowRoles('sales'), (req, res) => {
  const salesScope = req.user.role === 'sales'
    ? 'WHERE sales_rep_id = ?'
    : '';
  const joinedOrderScope = req.user.role === 'sales'
    ? 'WHERE orders.sales_rep_id = ?'
    : '';
  const params = req.user.role === 'sales' ? [req.user.id] : [];
  const totals = db
    .prepare(
      `SELECT COUNT(*) AS order_count,
              SUM(CASE WHEN order_status NOT IN ('complete', 'pickup_delivered') THEN 1 ELSE 0 END) AS active_orders,
              SUM(CASE WHEN order_status IN ('complete', 'pickup_delivered') THEN 1 ELSE 0 END) AS completed_orders,
              SUM(total_quantity) AS total_quantity,
              SUM(layout_fee) AS layout_fees,
              SUM(down_payment) AS down_payments
       FROM orders ${salesScope}`,
    )
    .get(...params);
  const byStage = db
    .prepare(
      `SELECT order_status, COUNT(*) AS count,
              COALESCE(SUM(total_quantity), 0) AS total_quantity
       FROM orders ${salesScope}
       GROUP BY order_status`,
    )
    .all(...params);
  const floorMetrics = db
    .prepare(
      `SELECT COALESCE(SUM(station_logs.paper_meters), 0) AS paper_meters,
              SUM(CASE WHEN NULLIF(TRIM(station_logs.defect_description), '') IS NOT NULL THEN 1 ELSE 0 END) AS defects
       FROM station_logs
       JOIN orders ON orders.order_id = station_logs.order_id
       ${joinedOrderScope}`,
    )
    .get(...params);
  const consumptionByStage = db
    .prepare(
    `SELECT station, COUNT(*) AS entries,
            COALESCE(SUM(station_logs.paper_meters), 0) AS paper_meters,
            SUM(CASE WHEN NULLIF(TRIM(station_logs.defect_description), '') IS NOT NULL THEN 1 ELSE 0 END) AS defects,
            GROUP_CONCAT(DISTINCT NULLIF(TRIM(station_logs.fabric_used), '')) AS fabric_used
     FROM station_logs
     JOIN orders ON orders.order_id = station_logs.order_id
     ${joinedOrderScope}
     GROUP BY station
     ORDER BY station`,
    )
    .all(...params);
  res.json({
    report: {
      ...totals,
      total_layout_fees: totals.layout_fees || 0,
      total_down_payments: totals.down_payments || 0,
      paper_meters: floorMetrics.paper_meters,
      defects: floorMetrics.defects || 0,
      by_stage: byStage,
      consumption_by_stage: consumptionByStage,
    },
  });
});

app.get('/api/orders/:orderId/activity', requireAuth, (req, res) => {
  const order = getOrderForUser(req, res, { allowReprint: true });
  if (!order) return;
  const events = db
    .prepare(
      `SELECT order_events.id, order_events.event_type, order_events.from_stage,
              order_events.to_stage, order_events.details, order_events.created_at,
              users.name AS actor_name
       FROM order_events JOIN users ON users.id = order_events.actor_id
       WHERE order_events.order_id = ? ORDER BY order_events.id DESC`,
    )
    .all(order.order_id);
  const stationLogs = db
    .prepare(
      `SELECT station_logs.id, station_logs.station, station_logs.fabric_used,
              station_logs.printer_brand,
              station_logs.paper_meters, station_logs.defect_description,
              station_logs.created_at, users.name AS operator
       FROM station_logs JOIN users ON users.id = station_logs.operator_id
       WHERE station_logs.order_id = ? ORDER BY station_logs.id DESC`,
    )
    .all(order.order_id);
  const productionItems = db.prepare(
    `SELECT production_log_items.*
    FROM production_log_items
    JOIN station_logs ON station_logs.id = production_log_items.station_log_id
    WHERE station_logs.order_id = ?
    ORDER BY production_log_items.id`,
  ).all(order.order_id);
  const itemsByLog = new Map();
  for (const item of productionItems) {
    const items = itemsByLog.get(item.station_log_id) || [];
    items.push(item);
    itemsByLog.set(item.station_log_id, items);
  }
  for (const log of stationLogs) {
    log.production_items = itemsByLog.get(log.id) || [];
  }
  res.json({ events, station_logs: stationLogs });
});

app.use((error, req, res, next) => {
  if (res.headersSent) return next(error);
  if (error instanceof multer.MulterError) {
    return res.status(400).json({ error: error.message });
  }
  console.error(error);
  return res.status(500).json({ error: 'An unexpected server error occurred' });
});

if (fs.existsSync(frontendDist)) {
  app.use(express.static(frontendDist));
  app.get('*', (req, res, next) => {
    if (req.path === '/api' || req.path.startsWith('/api/')) return next();
    res.sendFile(path.join(frontendDist, 'index.html'), (error) => {
      if (error) next(error);
    });
  });
}

app.listen(PORT, HOST, () => {
  console.log(`Backend listening on http://${HOST}:${PORT}`);
});
