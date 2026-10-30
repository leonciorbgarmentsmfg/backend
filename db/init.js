import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { scryptSync, timingSafeEqual } from 'node:crypto';
import Database from 'better-sqlite3';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const dataDir = process.env.BPSMS_DATA_DIR || path.join(__dirname, '..', 'data');
const dbPath = path.join(dataDir, 'app.db');
const ADMIN_USERNAME = 'ADMINISTRATOR';
const ADMIN_PASSWORD_SALT = 'a1f744672f84784ae3c17225bc6acac0';
const ADMIN_PASSWORD_HASH =
  '0351abd22b285c2e1b0ab26f2fbe413bc19dd84b3234278dd9d6c703c4b7dd835039e0bd275399da9613fc8b2fa78911e1ac1543c3c6f204a1b9f097ee136e3a';
const styleNumberSeedData = {
  product_category: [
    ['T-SHIRT', 'TS'], ['POLO SHIRT', 'PS'], ['SINGLET', 'SL'],
    ['JERSEY SET', 'SET'], ['OUTERWEAR', 'OW'], ['DRESS', 'DR'],
    ['SHORTS', 'SH'], ['PANTS', 'PA'], ['BOMBER', 'BB'], ['FINISHER SHIRT', 'FS'],
  ],
  product_style: [
    ['REGULAR', '1'], ['BASKETBALL', '2'], ['VOLLEYBALL', '3'],
    ['BASEBALL', '4'], ['SWIM', '5'], ['FOOTBALL', '6'], ['PE', '7'], ['BIKE', '8'],
  ],
  product_cut: [
    ['REGULAR', '1'], ['LADIES', '2'], ['OVERSIZED', '3'], ['CROPPED', '4'],
    ['NBA', '5'], ['RACERBACK', '6'], ['KIDS', '7'], ['SNUG FIT', '8'], ['JOGGER', '9'],
  ],
  product_neck: [
    ['ROUND', '01'], ['REGULAR COLLAR WITH BUTTON', '02'],
    ['CHINESE COLLAR WITH ZIPPER', '03'], ['V-NECK', '04'], ['V-NECK FLAT', '05'],
    ['V-NECK WITH COLLAR', '06'], ['REGULAR COLLAR WITH ZIPPER', '07'],
    ['CHINESE COLLAR WITH BUTTONS', '08'], ['WISHBONE', '09'],
    ['V-NECK WITH INSERT', '10'], ['CREW (ROUND)', '11'], ['ASYMMETRICAL', '12'],
    ['WITH HOODIE', '13'], ['KNITTED-PLAIN', '14'], ['KNITTED-WITH DESIGN', '15'],
    ['STAND COLLAR', '16'],
  ],
  product_sleeve: [
    ['SLEEVELESS', '0'], ['SET-IN (REGULAR)', '1'], ['RAGLAN', '2'],
    ['SLEEVELESS W/O BIAS', '3'], ['SLEEVELESS WITH BIAS', '4'],
  ],
  product_length: [
    ['SLEEVELESS', '0'], ['CAP', '1'], ['SHORT', '2'], ['ELBOW (ABOVE)', '3'],
    ['MID (BELOW ELBOW)', '4'], ['THREE QUARTER', '5'], ['LONG', '6'],
    ['SHORT WITH CUFF', '7'], ['LONG WITH CUFF', '8'],
  ],
  product_others: [
    ['N/A', '0'], ['SHORT WITH SLIT', '1'], ['STRAIGHT PANTS', '2'],
    ['JOGGER', '3'], ['MEN CYCLING', '4'], ['REGULAR', '5'], ['NBA', '6'],
    ['LADIES CYCLING', '7'], ['REVERSIBLE', '8'], ['LEG CUFF RIBBING', '9'],
    ['KNITTED WAISTBAND', '10'], ['ZIP FRONT', '11'],
  ],
};

fs.mkdirSync(dataDir, { recursive: true });

const db = new Database(dbPath);
db.pragma('foreign_keys = ON');
const existingColumns = db
  .prepare("PRAGMA table_info('users')")
  .all()
  .map((column) => column.name);
const hasProtectedAdminSchema =
  existingColumns.includes('password_hash') &&
  existingColumns.includes('is_protected') &&
  existingColumns.includes('username');

const initializeDatabase = db.transaction(() => {
  if (!hasProtectedAdminSchema && existingColumns.length > 0) {
    db.exec('DROP TABLE users');
  }

  if (
    hasProtectedAdminSchema &&
    !existingColumns.includes('assigned_station')
  ) {
    db.exec('ALTER TABLE users ADD COLUMN assigned_station TEXT');
  }
  if (
    hasProtectedAdminSchema &&
    !existingColumns.includes('business_unit')
  ) {
    db.exec("ALTER TABLE users ADD COLUMN business_unit TEXT CHECK (business_unit IS NULL OR business_unit IN ('dtf', 'sublimation'))");
  }
  if (
    hasProtectedAdminSchema &&
    !existingColumns.includes('must_change_password')
  ) {
    db.exec(
      'ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 0 CHECK (must_change_password IN (0, 1))',
    );
  }

  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      username TEXT NOT NULL UNIQUE,
      password_hash TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'sales',
      assigned_station TEXT,
      business_unit TEXT CHECK (business_unit IS NULL OR business_unit IN ('dtf', 'sublimation')),
      must_change_password INTEGER NOT NULL DEFAULT 0 CHECK (must_change_password IN (0, 1)),
      is_protected INTEGER NOT NULL DEFAULT 0 CHECK (is_protected IN (0, 1))
    );

    CREATE TABLE IF NOT EXISTS sessions (
      token_hash TEXT PRIMARY KEY,
      user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      expires_at INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS clients (
      client_id TEXT PRIMARY KEY,
      company_name TEXT NOT NULL UNIQUE,
      contact_person TEXT NOT NULL,
      contact_number TEXT NOT NULL,
      shipping_address TEXT NOT NULL,
      sales_rep_id INTEGER NOT NULL REFERENCES users(id),
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS orders (
      order_id TEXT PRIMARY KEY,
      client_id TEXT NOT NULL REFERENCES clients(client_id),
      sales_rep_id INTEGER NOT NULL REFERENCES users(id),
      team_project_name TEXT NOT NULL,
      order_date TEXT NOT NULL DEFAULT (date('now')),
      due_date TEXT NOT NULL,
      style_number TEXT NOT NULL,
      category TEXT NOT NULL,
      style TEXT NOT NULL,
      cut TEXT NOT NULL,
      neck TEXT NOT NULL,
      sleeve_type TEXT NOT NULL,
      sleeve_length TEXT NOT NULL,
      fabric TEXT NOT NULL,
      fabric_width TEXT NOT NULL DEFAULT '',
      gsm INTEGER NOT NULL,
      order_priority TEXT NOT NULL DEFAULT 'Normal'
        CHECK (order_priority IN ('Urgent', 'High', 'Normal')),
      item TEXT NOT NULL,
      total_quantity INTEGER NOT NULL DEFAULT 0,
      spreadsheet_path TEXT,
      production_spreadsheet_path TEXT,
      business_unit TEXT NOT NULL DEFAULT 'sublimation'
        CHECK (business_unit IN ('dtf', 'sublimation')),
      print_mode TEXT NOT NULL DEFAULT 'print_press'
        CHECK (print_mode IN ('print_only', 'print_press')),
      mockup_path TEXT,
      swatch_path TEXT,
      sizing_notes TEXT NOT NULL DEFAULT '',
      other_details TEXT,
      layout_fee REAL NOT NULL DEFAULT 0,
      layout_fee_paid INTEGER NOT NULL DEFAULT 0 CHECK (layout_fee_paid IN (0, 1)),
      down_payment REAL NOT NULL DEFAULT 0,
      assigned_artist_id INTEGER REFERENCES users(id),
      printer_machine TEXT,
      pressing_machine TEXT,
      cutting_machine TEXT,
      packed_quantity INTEGER CHECK (packed_quantity IS NULL OR packed_quantity >= 0),
      per_piece_price REAL CHECK (per_piece_price IS NULL OR per_piece_price >= 0),
      order_status TEXT NOT NULL DEFAULT 'mockup'
        CHECK (order_status IN ('mockup', 'sales_approval', 'sizing', 'printing', 'pressing', 'cutting', 'sewing', 'packing', 'complete', 'pickup_delivered')),
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS order_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      order_id TEXT NOT NULL REFERENCES orders(order_id) ON DELETE CASCADE,
      actor_id INTEGER NOT NULL REFERENCES users(id),
      event_type TEXT NOT NULL,
      from_stage TEXT,
      to_stage TEXT,
      details TEXT,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS station_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      order_id TEXT NOT NULL REFERENCES orders(order_id) ON DELETE CASCADE,
      station TEXT NOT NULL,
      operator_id INTEGER NOT NULL REFERENCES users(id),
      fabric_used TEXT NOT NULL DEFAULT '',
      printer_brand TEXT NOT NULL DEFAULT '',
      paper_meters REAL NOT NULL DEFAULT 0 CHECK (paper_meters >= 0),
      defect_description TEXT,
      reprint_job_id INTEGER REFERENCES reprint_jobs(id) ON DELETE SET NULL,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS production_log_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      station_log_id INTEGER NOT NULL REFERENCES station_logs(id) ON DELETE CASCADE,
      consumable_code TEXT NOT NULL,
      label TEXT NOT NULL,
      quantity REAL NOT NULL CHECK (quantity > 0),
      unit TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS fabric_inventory (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      fabric TEXT NOT NULL,
      fabric_width TEXT NOT NULL DEFAULT '',
      gsm INTEGER NOT NULL CHECK (gsm > 0),
      roll_quantity INTEGER NOT NULL CHECK (roll_quantity > 0),
      measurement_type TEXT NOT NULL CHECK (measurement_type IN ('meter', 'kilo')),
      meters_per_roll REAL CHECK (meters_per_roll IS NULL OR meters_per_roll > 0),
      kilo_per_meter REAL CHECK (kilo_per_meter IS NULL OR kilo_per_meter > 0),
      kilo_per_roll REAL CHECK (kilo_per_roll IS NULL OR kilo_per_roll > 0),
      total_meter REAL NOT NULL CHECK (total_meter > 0),
      entry_date TEXT NOT NULL,
      created_by INTEGER NOT NULL REFERENCES users(id),
      updated_by INTEGER NOT NULL REFERENCES users(id),
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    UPDATE fabric_inventory
    SET total_meter = kilo_per_roll * roll_quantity * kilo_per_meter
    WHERE measurement_type = 'kilo'
      AND kilo_per_meter IS NOT NULL
      AND kilo_per_roll IS NOT NULL;
    CREATE INDEX IF NOT EXISTS idx_fabric_inventory_date
      ON fabric_inventory(entry_date DESC, id DESC);

    CREATE TABLE IF NOT EXISTS consumable_transactions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      business_unit TEXT NOT NULL CHECK (business_unit IN ('dtf', 'sublimation')),
      consumable_code TEXT NOT NULL,
      transaction_type TEXT NOT NULL CHECK (transaction_type IN ('stock_in', 'adjustment', 'usage')),
      quantity REAL NOT NULL CHECK (quantity != 0),
      order_id TEXT REFERENCES orders(order_id) ON DELETE CASCADE,
      station_log_id INTEGER REFERENCES station_logs(id) ON DELETE SET NULL,
      entry_date TEXT NOT NULL,
      note TEXT NOT NULL DEFAULT '',
      created_by INTEGER NOT NULL REFERENCES users(id),
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
    CREATE INDEX IF NOT EXISTS idx_consumable_transactions_product
      ON consumable_transactions(business_unit, consumable_code, entry_date DESC, id DESC);
    UPDATE consumable_transactions
    SET consumable_code = 'subli_ink_legacy_atexco',
        quantity = quantity / 1000
    WHERE consumable_code = 'subli_ink_atexco';
    UPDATE consumable_transactions
    SET consumable_code = 'subli_ink_legacy_creons',
        quantity = quantity / 1000
    WHERE consumable_code = 'subli_ink_creons';

    CREATE TABLE IF NOT EXISTS reprint_jobs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      order_id TEXT NOT NULL REFERENCES orders(order_id) ON DELETE CASCADE,
      station_log_id INTEGER NOT NULL REFERENCES station_logs(id) ON DELETE CASCADE,
      created_by INTEGER NOT NULL REFERENCES users(id),
      image_path TEXT,
      status TEXT NOT NULL DEFAULT 'open' CHECK (status IN ('open', 'in_progress', 'complete')),
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      completed_at TEXT
    );

    CREATE TABLE IF NOT EXISTS test_print_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      order_id TEXT NOT NULL REFERENCES orders(order_id) ON DELETE CASCADE,
      artist_id INTEGER NOT NULL REFERENCES users(id),
      mockup_path TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'open' CHECK (status IN ('open', 'in_progress', 'complete')),
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      completed_at TEXT
    );

    CREATE TABLE IF NOT EXISTS machines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      category TEXT NOT NULL CHECK (category IN ('printing', 'pressing', 'cutting')),
      name TEXT NOT NULL,
      business_unit TEXT CHECK (business_unit IN ('dtf', 'sublimation')),
      UNIQUE (category, name, business_unit)
    );

    CREATE TABLE IF NOT EXISTS style_number_options (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      category TEXT NOT NULL CHECK (category IN (
        'product_category', 'product_style', 'product_cut', 'product_neck',
        'product_sleeve', 'product_length', 'product_others'
      )),
      name TEXT NOT NULL,
      code TEXT NOT NULL,
      UNIQUE (category, name),
      UNIQUE (category, code)
    );

    DROP TRIGGER IF EXISTS prevent_protected_user_update;
    CREATE TRIGGER prevent_protected_user_update
    BEFORE UPDATE ON users
    WHEN OLD.is_protected = 1 AND (
      NEW.name IS NOT OLD.name OR
      NEW.username IS NOT OLD.username OR
      NEW.role IS NOT OLD.role OR
      NEW.assigned_station IS NOT OLD.assigned_station OR
      NEW.business_unit IS NOT OLD.business_unit OR
      NEW.is_protected IS NOT OLD.is_protected
    )
    BEGIN
      SELECT RAISE(ABORT, 'Protected administrator cannot be edited');
    END;

    CREATE TRIGGER IF NOT EXISTS prevent_protected_user_delete
    BEFORE DELETE ON users
    WHEN OLD.is_protected = 1 AND (
      NEW.name IS NOT OLD.name OR
      NEW.username IS NOT OLD.username OR
      NEW.role IS NOT OLD.role OR
      NEW.assigned_station IS NOT OLD.assigned_station OR
      NEW.business_unit IS NOT OLD.business_unit OR
      NEW.is_protected IS NOT OLD.is_protected
    )
    BEGIN
      SELECT RAISE(ABORT, 'Protected administrator cannot be deleted');
    END;
  `);

  const clientColumns = db
    .prepare("PRAGMA table_info('clients')")
    .all()
    .map((column) => column.name);
  if (!clientColumns.includes('sales_rep_id')) {
    db.exec('ALTER TABLE clients ADD COLUMN sales_rep_id INTEGER REFERENCES users(id)');
  }

  const orderColumns = db
    .prepare("PRAGMA table_info('orders')")
    .all()
    .map((column) => column.name);
  if (orderColumns.length > 0 && !orderColumns.includes('swatch_path')) {
    db.exec('ALTER TABLE orders ADD COLUMN swatch_path TEXT');
  }
  if (orderColumns.length > 0 && !orderColumns.includes('fabric_width')) {
    db.exec("ALTER TABLE orders ADD COLUMN fabric_width TEXT NOT NULL DEFAULT ''");
  }
  if (orderColumns.length > 0 && !orderColumns.includes('order_priority')) {
    db.exec(`ALTER TABLE orders ADD COLUMN order_priority TEXT NOT NULL DEFAULT 'Normal'
      CHECK (order_priority IN ('Urgent', 'High', 'Normal'))`);
  }
  if (orderColumns.length > 0 && !orderColumns.includes('sizing_notes')) {
    db.exec("ALTER TABLE orders ADD COLUMN sizing_notes TEXT NOT NULL DEFAULT ''");
  }
  if (orderColumns.length > 0 && !orderColumns.includes('cutting_machine')) {
    db.exec('ALTER TABLE orders ADD COLUMN cutting_machine TEXT');
  }
  if (orderColumns.length > 0 && !orderColumns.includes('packed_quantity')) {
    db.exec('ALTER TABLE orders ADD COLUMN packed_quantity INTEGER CHECK (packed_quantity IS NULL OR packed_quantity >= 0)');
  }
  if (orderColumns.length > 0 && !orderColumns.includes('per_piece_price')) {
    db.exec('ALTER TABLE orders ADD COLUMN per_piece_price REAL CHECK (per_piece_price IS NULL OR per_piece_price >= 0)');
  }
  const stationLogColumns = db
    .prepare("PRAGMA table_info('station_logs')")
    .all()
    .map((column) => column.name);
  if (!stationLogColumns.includes('fabric_used')) {
    db.exec("ALTER TABLE station_logs ADD COLUMN fabric_used TEXT NOT NULL DEFAULT ''");
  }
  if (!stationLogColumns.includes('printer_brand')) {
    db.exec("ALTER TABLE station_logs ADD COLUMN printer_brand TEXT NOT NULL DEFAULT ''");
  }
  if (!stationLogColumns.includes('reprint_job_id')) {
    db.exec('ALTER TABLE station_logs ADD COLUMN reprint_job_id INTEGER REFERENCES reprint_jobs(id) ON DELETE SET NULL');
  }
  const fabricInventoryColumns = db
    .prepare("PRAGMA table_info('fabric_inventory')")
    .all()
    .map((column) => column.name);
  if (!fabricInventoryColumns.includes('fabric_width')) {
    db.exec("ALTER TABLE fabric_inventory ADD COLUMN fabric_width TEXT NOT NULL DEFAULT ''");
  }
  db.exec(`CREATE INDEX IF NOT EXISTS idx_fabric_inventory_catalog
    ON fabric_inventory(fabric, fabric_width, gsm)`);

  db.prepare("INSERT OR IGNORE INTO machines (category, name) VALUES ('pressing', 'ES')").run();
  for (const order of db.prepare('SELECT DISTINCT cutting_machine FROM orders WHERE cutting_machine IS NOT NULL AND cutting_machine != \'\'').all()) {
    db.prepare("INSERT OR IGNORE INTO machines (category, name) VALUES ('cutting', ?)").run(order.cutting_machine);
  }

  const protectedAdmin = db
    .prepare('SELECT id, username FROM users WHERE is_protected = 1 LIMIT 1')
    .get();
  if (protectedAdmin) {
    if (protectedAdmin.username !== ADMIN_USERNAME) {
      const usernameInUse = db
        .prepare('SELECT id FROM users WHERE username = ? AND id != ?')
        .get(ADMIN_USERNAME, protectedAdmin.id);
      if (usernameInUse) {
        throw new Error('Cannot rename the protected administrator: username "Administrator" is already in use');
      }
      db.exec('DROP TRIGGER IF EXISTS prevent_protected_user_update');
      db.prepare('UPDATE users SET username = ? WHERE id = ?').run(
        ADMIN_USERNAME,
        protectedAdmin.id,
      );
      db.exec(`
        CREATE TRIGGER prevent_protected_user_update
        BEFORE UPDATE ON users
        WHEN OLD.is_protected = 1 AND (
          NEW.name IS NOT OLD.name OR
          NEW.username IS NOT OLD.username OR
          NEW.role IS NOT OLD.role OR
          NEW.assigned_station IS NOT OLD.assigned_station OR
          NEW.is_protected IS NOT OLD.is_protected
        )
        BEGIN
          SELECT RAISE(ABORT, 'Protected administrator cannot be edited');
        END;
      `);
    }
  } else {
    const usernameInUse = db
      .prepare('SELECT id FROM users WHERE username = ?')
      .get(ADMIN_USERNAME);
    if (usernameInUse) {
      throw new Error('Cannot create the protected administrator: username "Administrator" is already in use');
    }
    db.prepare(`
      INSERT INTO users (name, username, password_hash, role, is_protected)
      VALUES (?, ?, ?, 'admin', 1)
    `).run(
      'Administrator',
      ADMIN_USERNAME,
      `${ADMIN_PASSWORD_SALT}:${ADMIN_PASSWORD_HASH}`,
    );
  }
});

initializeDatabase.immediate();

const machineColumns = db
  .prepare("PRAGMA table_info('machines')")
  .all()
  .map((column) => column.name);
if (machineColumns.length > 0 && !machineColumns.includes('business_unit')) {
  db.exec("ALTER TABLE machines ADD COLUMN business_unit TEXT CHECK (business_unit IN ('dtf', 'sublimation'))");
}
const machinesSchema = db
  .prepare("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'machines'")
  .get()?.sql || '';
if (!/UNIQUE\s*\(\s*category\s*,\s*name\s*,\s*business_unit\s*\)/i.test(machinesSchema)) {
  db.transaction(() => {
    db.exec(`
      CREATE TABLE machines_replacement (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT NOT NULL CHECK (category IN ('printing', 'pressing', 'cutting')),
        name TEXT NOT NULL,
        business_unit TEXT CHECK (business_unit IN ('dtf', 'sublimation')),
        UNIQUE (category, name, business_unit)
      );
      INSERT INTO machines_replacement (id, category, name, business_unit)
        SELECT id, category, name, business_unit FROM machines;
      DROP TABLE machines;
      ALTER TABLE machines_replacement RENAME TO machines;
    `);
  })();
}
db.exec(`
  CREATE UNIQUE INDEX IF NOT EXISTS idx_machines_category_name_unit
  ON machines(category, name, ifnull(business_unit, ''));
`);
db.transaction(() => {
  db.prepare(
    `UPDATE machines SET business_unit = 'sublimation'
     WHERE category = 'printing' AND business_unit IS NULL
       AND upper(name) IN ('ATEXCO', 'CREONS')
       AND NOT EXISTS (
         SELECT 1 FROM machines assigned
         WHERE assigned.category = machines.category
           AND assigned.name = machines.name
           AND assigned.business_unit = 'sublimation'
       )`,
  ).run();
  db.prepare(
    "DELETE FROM machines WHERE category = 'printing' AND business_unit IS NULL AND upper(name) IN ('ATEXCO', 'CREONS')",
  ).run();
  db.prepare(
    "INSERT OR IGNORE INTO machines (category, name, business_unit) VALUES ('printing', 'CREONS', 'sublimation')",
  ).run();
})();
db.prepare(
  "INSERT OR IGNORE INTO machines (category, name, business_unit) VALUES ('printing', ?, 'dtf')",
).run('Creons');
db.prepare(
  "INSERT OR IGNORE INTO machines (category, name, business_unit) VALUES ('printing', ?, 'dtf')",
).run('Grando');
db.transaction(() => {
  for (const brand of ['CREONS', 'GRANDO']) {
    const matches = db.prepare(
      "SELECT id, name FROM machines WHERE category = 'printing' AND business_unit = 'dtf' AND upper(name) = ? ORDER BY (name = ?) DESC, id",
    ).all(brand, brand);
    if (matches.length > 0) {
      const canonical = matches[0];
      db.prepare(
        "UPDATE orders SET printer_machine = ? WHERE business_unit = 'dtf' AND upper(printer_machine) = ?",
      ).run(brand, brand);
      if (canonical.name !== brand) {
        db.prepare('UPDATE machines SET name = ? WHERE id = ?').run(brand, canonical.id);
      }
      for (const duplicate of matches.slice(1)) {
        db.prepare('DELETE FROM machines WHERE id = ?').run(duplicate.id);
      }
    } else {
      db.prepare(
        "INSERT INTO machines (category, name, business_unit) VALUES ('printing', ?, 'dtf')",
      ).run(brand);
    }
  }

  db.prepare(
    "DELETE FROM machines WHERE category = 'printing' AND business_unit = 'dtf' AND name NOT IN ('CREONS', 'GRANDO')",
  ).run();
  db.prepare("DELETE FROM machines WHERE lower(name) = 'other'").run();
})()

const initialOrderColumns = db
  .prepare("PRAGMA table_info('orders')")
  .all()
  .map((column) => column.name);
if (initialOrderColumns.length > 0 && !initialOrderColumns.includes('production_spreadsheet_path')) {
  db.exec('ALTER TABLE orders ADD COLUMN production_spreadsheet_path TEXT');
}
if (initialOrderColumns.length > 0 && !initialOrderColumns.includes('business_unit')) {
  db.exec("ALTER TABLE orders ADD COLUMN business_unit TEXT NOT NULL DEFAULT 'sublimation' CHECK (business_unit IN ('dtf', 'sublimation'))");
}
if (initialOrderColumns.length > 0 && !initialOrderColumns.includes('print_mode')) {
  db.exec("ALTER TABLE orders ADD COLUMN print_mode TEXT NOT NULL DEFAULT 'print_press' CHECK (print_mode IN ('print_only', 'print_press'))");
}

const reprintJobColumns = db
  .prepare("PRAGMA table_info('reprint_jobs')")
  .all()
  .map((column) => column.name);
if (reprintJobColumns.length > 0 && !reprintJobColumns.includes('image_path')) {
  db.exec('ALTER TABLE reprint_jobs ADD COLUMN image_path TEXT');
}

const ordersSchema = db
  .prepare("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'orders'")
  .get()?.sql || '';
if (
  !ordersSchema.includes("'pressing'") ||
  !ordersSchema.includes("'pickup_delivered'") ||
  /CHECK\s*\(\s*category\s+IN/i.test(ordersSchema) ||
  /CHECK\s*\(\s*cut\s+IN/i.test(ordersSchema)
) {
  db.pragma('foreign_keys = OFF');
  try {
    const migrateOrders = db.transaction(() => {
      db.exec(`
        CREATE TABLE orders_stage_migration (
          order_id TEXT PRIMARY KEY,
          client_id TEXT NOT NULL REFERENCES clients(client_id),
          sales_rep_id INTEGER NOT NULL REFERENCES users(id),
          team_project_name TEXT NOT NULL,
          order_date TEXT NOT NULL DEFAULT (date('now')),
          due_date TEXT NOT NULL,
          style_number TEXT NOT NULL,
          category TEXT NOT NULL,
          style TEXT NOT NULL,
          cut TEXT NOT NULL,
          neck TEXT NOT NULL,
          sleeve_type TEXT NOT NULL,
          sleeve_length TEXT NOT NULL,
          fabric TEXT NOT NULL,
          fabric_width TEXT NOT NULL DEFAULT '',
          gsm INTEGER NOT NULL,
          order_priority TEXT NOT NULL DEFAULT 'Normal'
            CHECK (order_priority IN ('Urgent', 'High', 'Normal')),
          item TEXT NOT NULL,
          total_quantity INTEGER NOT NULL DEFAULT 0,
          spreadsheet_path TEXT,
          production_spreadsheet_path TEXT,
          business_unit TEXT NOT NULL DEFAULT 'sublimation'
            CHECK (business_unit IN ('dtf', 'sublimation')),
          print_mode TEXT NOT NULL DEFAULT 'print_press'
            CHECK (print_mode IN ('print_only', 'print_press')),
          mockup_path TEXT,
          swatch_path TEXT,
          sizing_notes TEXT NOT NULL DEFAULT '',
          other_details TEXT,
          layout_fee REAL NOT NULL DEFAULT 0,
          layout_fee_paid INTEGER NOT NULL DEFAULT 1 CHECK (layout_fee_paid IN (0, 1)),
          down_payment REAL NOT NULL DEFAULT 0,
          assigned_artist_id INTEGER REFERENCES users(id),
          printer_machine TEXT,
          pressing_machine TEXT,
          cutting_machine TEXT,
          packed_quantity INTEGER CHECK (packed_quantity IS NULL OR packed_quantity >= 0),
          per_piece_price REAL CHECK (per_piece_price IS NULL OR per_piece_price >= 0),
          order_status TEXT NOT NULL DEFAULT 'mockup'
            CHECK (order_status IN ('mockup', 'sales_approval', 'sizing', 'printing', 'pressing', 'cutting', 'sewing', 'packing', 'complete', 'pickup_delivered')),
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        INSERT INTO orders_stage_migration (
          order_id, client_id, sales_rep_id, team_project_name, order_date, due_date,
          style_number, category, style, cut, neck, sleeve_type, sleeve_length,
          fabric, fabric_width, gsm, order_priority, item, total_quantity, spreadsheet_path,
          production_spreadsheet_path, business_unit, print_mode, mockup_path, swatch_path,
          sizing_notes, other_details, layout_fee, layout_fee_paid, down_payment, assigned_artist_id,
          printer_machine, pressing_machine, cutting_machine, packed_quantity, per_piece_price,
          order_status, created_at, updated_at
        )
        SELECT
          order_id, client_id, sales_rep_id, team_project_name, order_date, due_date,
          style_number, category, style, cut, neck, sleeve_type, sleeve_length,
          fabric, fabric_width, gsm, order_priority, item, total_quantity, spreadsheet_path,
          production_spreadsheet_path, business_unit, print_mode, mockup_path, swatch_path,
          sizing_notes, other_details, layout_fee, 1, down_payment, assigned_artist_id,
          printer_machine, pressing_machine, cutting_machine, packed_quantity, per_piece_price,
          order_status, created_at, updated_at
        FROM orders;

        DROP TABLE orders;
        ALTER TABLE orders_stage_migration RENAME TO orders;
      `);
    });
    migrateOrders.immediate();
  } finally {
    db.pragma('foreign_keys = ON');
  }
}

const migratedOrderColumns = db
  .prepare("PRAGMA table_info('orders')")
  .all()
  .map((column) => column.name);
if (migratedOrderColumns.length && !migratedOrderColumns.includes('layout_fee_paid')) {
  db.exec(
    'ALTER TABLE orders ADD COLUMN layout_fee_paid INTEGER NOT NULL DEFAULT 1 CHECK (layout_fee_paid IN (0, 1))',
  );
}

const insertStyleNumberOption = db.prepare(
  'INSERT OR IGNORE INTO style_number_options (category, name, code) VALUES (?, ?, ?)',
);
for (const [category, options] of Object.entries(styleNumberSeedData)) {
  for (const [name, code] of options) {
    insertStyleNumberOption.run(category, name, code);
  }
}

const styleOptionCodes = new Map(
  db
    .prepare('SELECT category, name, code FROM style_number_options')
    .all()
    .map((option) => [`${option.category}:${option.name}`, option.code]),
);
const styleNumberOrderFields = [
  ['product_category', 'category'],
  ['product_style', 'style'],
  ['product_cut', 'cut'],
  ['product_neck', 'neck'],
  ['product_sleeve', 'sleeve_type'],
  ['product_length', 'sleeve_length'],
  ['product_others', 'other_details'],
];
const updateLegacyStyleNumber = db.prepare(
  'UPDATE orders SET style_number = ? WHERE order_id = ?',
);
for (const order of db.prepare('SELECT * FROM orders').all()) {
  const codes = styleNumberOrderFields.map(([category, field]) =>
    styleOptionCodes.get(`${category}:${order[field]}`),
  );
  if (codes.every(Boolean) && order.style_number === codes.join('-')) {
    updateLegacyStyleNumber.run(
      `${codes.slice(0, 3).join('')}-${codes[3]}-${codes.slice(4).join('')}`,
      order.order_id,
    );
  }
}

const userColumns = db
  .prepare("PRAGMA table_info('users')")
  .all()
  .map((column) => column.name);
if (userColumns.includes('email')) {
  db.pragma('foreign_keys = OFF');
  try {
    const removeUserEmail = db.transaction(() => {
      db.exec(`
        CREATE TABLE users_without_email (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          username TEXT NOT NULL UNIQUE,
          password_hash TEXT NOT NULL,
          role TEXT NOT NULL DEFAULT 'sales',
          assigned_station TEXT,
          must_change_password INTEGER NOT NULL DEFAULT 0 CHECK (must_change_password IN (0, 1)),
          is_protected INTEGER NOT NULL DEFAULT 0 CHECK (is_protected IN (0, 1))
        );

        INSERT INTO users_without_email
          (id, name, username, password_hash, role, assigned_station, must_change_password, is_protected)
        SELECT id, name, username, password_hash, role, assigned_station, must_change_password, is_protected
        FROM users;

        DROP TABLE users;
        ALTER TABLE users_without_email RENAME TO users;

        CREATE TRIGGER prevent_protected_user_update
        BEFORE UPDATE ON users
        WHEN OLD.is_protected = 1
        BEGIN
          SELECT RAISE(ABORT, 'Protected administrator cannot be edited');
        END;

        CREATE TRIGGER prevent_protected_user_delete
        BEFORE DELETE ON users
        WHEN OLD.is_protected = 1
        BEGIN
          SELECT RAISE(ABORT, 'Protected administrator cannot be deleted');
        END;
      `);
    });
    removeUserEmail.immediate();
  } finally {
    db.pragma('foreign_keys = ON');
  }
}

db.exec(`
  DROP TRIGGER IF EXISTS prevent_protected_user_update;
  CREATE TRIGGER prevent_protected_user_update
  BEFORE UPDATE ON users
  WHEN OLD.is_protected = 1 AND (
    NEW.name IS NOT OLD.name OR
    NEW.username IS NOT OLD.username OR
    NEW.role IS NOT OLD.role OR
    NEW.assigned_station IS NOT OLD.assigned_station OR
    NEW.is_protected IS NOT OLD.is_protected
  )
  BEGIN
    SELECT RAISE(ABORT, 'Protected administrator profile cannot be edited');
  END;
`);

const foreignKeyViolations = db.pragma('foreign_key_check');
if (foreignKeyViolations.length > 0) {
  throw new Error('Database migration failed: user records have broken foreign key references');
}

export function verifyAdminCredentials(username, password) {
  const user = db
    .prepare(
      `SELECT id, name, username, role, assigned_station, must_change_password,
              is_protected, password_hash
       FROM users
       WHERE username = ? COLLATE NOCASE`,
    )
    .get(username);

  if (!user || typeof password !== 'string') {
    return null;
  }

  const [salt, storedHash] = user.password_hash.split(':');
  const candidateHash = scryptSync(password, salt, 64);
  const expectedHash = Buffer.from(storedHash, 'hex');

  if (
    candidateHash.length !== expectedHash.length ||
    !timingSafeEqual(candidateHash, expectedHash)
  ) {
    return null;
  }

  const { password_hash: _passwordHash, is_protected: _isProtected, ...safeUser } = user;
  return safeUser;
}

export default db;
export { dbPath };
