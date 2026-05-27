CREATE TABLE IF NOT EXISTS mobile_capture (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  mobile_id TEXT NOT NULL UNIQUE,
  source_device TEXT NOT NULL DEFAULT '',
  user_name TEXT NOT NULL DEFAULT '',
  equipment_code TEXT NOT NULL DEFAULT '',
  component_name TEXT NOT NULL DEFAULT '',
  work_date TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  received_at TEXT NOT NULL DEFAULT '',
  desktop_imported_at TEXT,
  payload_json TEXT NOT NULL DEFAULT '{}'
);

CREATE INDEX IF NOT EXISTS idx_mobile_capture_imported ON mobile_capture (desktop_imported_at, id);
CREATE INDEX IF NOT EXISTS idx_mobile_capture_work_date ON mobile_capture (work_date);

CREATE TABLE IF NOT EXISTS mobile_photo (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  capture_id INTEGER NOT NULL,
  mobile_id TEXT NOT NULL DEFAULT '',
  file_name TEXT NOT NULL DEFAULT '',
  mime_type TEXT NOT NULL DEFAULT 'image/jpeg',
  captured_at TEXT NOT NULL DEFAULT '',
  r2_key TEXT NOT NULL DEFAULT '',
  data_url TEXT NOT NULL DEFAULT '',
  created_at TEXT NOT NULL DEFAULT '',
  FOREIGN KEY (capture_id) REFERENCES mobile_capture(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_mobile_photo_capture ON mobile_photo (capture_id);

CREATE TABLE IF NOT EXISTS catalog_snapshot (
  name TEXT PRIMARY KEY,
  payload_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS filter_inventory_snapshot (
  name TEXT PRIMARY KEY,
  payload_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS epp_snapshot (
  name TEXT PRIMARY KEY,
  payload_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS portal_snapshot (
  name TEXT PRIMARY KEY,
  payload_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS hose_change (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  external_id TEXT UNIQUE,
  change_date TEXT NOT NULL DEFAULT '',
  equipment TEXT NOT NULL DEFAULT '',
  system TEXT NOT NULL DEFAULT '',
  part_type TEXT NOT NULL DEFAULT 'MANGUERA',
  diameter TEXT NOT NULL DEFAULT '',
  length_m REAL NOT NULL DEFAULT 0,
  quantity REAL NOT NULL DEFAULT 1,
  unit_cost REAL NOT NULL DEFAULT 0,
  estimated_life_days REAL NOT NULL DEFAULT 30,
  estimated_weekly_qty REAL NOT NULL DEFAULT 0,
  failure_reason TEXT NOT NULL DEFAULT '',
  technician TEXT NOT NULL DEFAULT '',
  notes TEXT NOT NULL DEFAULT '',
  source TEXT NOT NULL DEFAULT 'web',
  deleted INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_hose_change_period ON hose_change (deleted, change_date, equipment);
CREATE INDEX IF NOT EXISTS idx_hose_change_external ON hose_change (external_id);
