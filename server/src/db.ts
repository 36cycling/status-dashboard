import initSqlJs, { Database } from 'sql.js';
import fs from 'fs';
import path from 'path';

let db: Database;

const dbPath = process.env.DATABASE_PATH || path.join(__dirname, '..', 'data', 'dashboard.db');

export async function initDb(): Promise<Database> {
  if (db) return db;

  const SQL = await initSqlJs();

  // Ensure data directory exists
  const dir = path.dirname(dbPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  // Load existing database or create new
  if (fs.existsSync(dbPath)) {
    const buffer = fs.readFileSync(dbPath);
    db = new SQL.Database(buffer);
  } else {
    db = new SQL.Database();
  }

  db.run('PRAGMA foreign_keys = ON');

  db.run(`
    CREATE TABLE IF NOT EXISTS customers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL DEFAULT '',
      company TEXT NOT NULL DEFAULT '',
      email TEXT NOT NULL UNIQUE,
      archived INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS timeline_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      customer_id INTEGER NOT NULL,
      type TEXT NOT NULL CHECK(type IN ('email_in', 'email_out', 'tl_contact', 'tl_deal')),
      subject TEXT NOT NULL DEFAULT '',
      summary TEXT NOT NULL DEFAULT '',
      date TEXT NOT NULL,
      is_replied INTEGER NOT NULL DEFAULT 0,
      outlook_message_id TEXT,
      metadata TEXT NOT NULL DEFAULT '{}',
      FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS auth_tokens (
      service TEXT PRIMARY KEY,
      access_token TEXT NOT NULL,
      refresh_token TEXT,
      expires_at TEXT
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS sync_state (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    )
  `);

  // Create indexes (ignore if exist)
  try { db.run('CREATE INDEX idx_events_customer ON timeline_events(customer_id)'); } catch {}
  try { db.run('CREATE INDEX idx_events_message_id ON timeline_events(outlook_message_id)'); } catch {}
  try { db.run('CREATE INDEX idx_customers_email ON customers(email)'); } catch {}

  saveDb();
  return db;
}

export function getDb(): Database {
  if (!db) throw new Error('Database not initialized. Call initDb() first.');
  return db;
}

export function saveDb() {
  if (!db) return;
  const data = db.export();
  const buffer = Buffer.from(data);
  fs.writeFileSync(dbPath, buffer);
}

// Helper functions to make queries easier
export function runQuery(sql: string, params: any[] = []) {
  const d = getDb();
  d.run(sql, params);
  saveDb();
}

export function getOne(sql: string, params: any[] = []): any | undefined {
  const d = getDb();
  const stmt = d.prepare(sql);
  stmt.bind(params);
  if (stmt.step()) {
    const row = stmt.getAsObject();
    stmt.free();
    return row;
  }
  stmt.free();
  return undefined;
}

export function getAll(sql: string, params: any[] = []): any[] {
  const d = getDb();
  const results: any[] = [];
  const stmt = d.prepare(sql);
  stmt.bind(params);
  while (stmt.step()) {
    results.push(stmt.getAsObject());
  }
  stmt.free();
  return results;
}
