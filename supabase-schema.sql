-- ================================================================
-- Waffarha Guard — Supabase Schema
-- Run this in: Supabase Dashboard → SQL Editor → New Query
-- ================================================================

-- ── 1. CUSTOMERS ────────────────────────────────────────────────
-- Normalized entity keyed by email.
-- One row per unique email across all platforms and all uploads.
-- Avoids re-storing domain/disposable metadata on every alert.
CREATE TABLE IF NOT EXISTS customers (
  id             UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  email          TEXT        UNIQUE NOT NULL,
  is_disposable  BOOLEAN     NOT NULL DEFAULT false,
  first_seen_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  last_seen_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── 2. UPLOAD_SESSIONS ──────────────────────────────────────────
-- One row per file import event.
-- All fraud alerts from that import reference this row.
CREATE TABLE IF NOT EXISTS upload_sessions (
  id               UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  platform         TEXT        NOT NULL,   -- 'paytabs' | 'noon' | 'paymob_wallets' | 'paymob_bnpl' | 'fawry' | 'admin'
  uploaded_by      TEXT        NOT NULL,   -- username from session
  filename         TEXT,
  record_count     INTEGER     NOT NULL DEFAULT 0,
  high_count       INTEGER     NOT NULL DEFAULT 0,
  mid_count        INTEGER     NOT NULL DEFAULT 0,
  high_amt_count   INTEGER     NOT NULL DEFAULT 0,
  fake_dom_count   INTEGER     NOT NULL DEFAULT 0,
  other_count      INTEGER     NOT NULL DEFAULT 0,  -- wallet abusers, recharge abusers, etc.
  created_at       TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── 3. FRAUD_ALERTS ─────────────────────────────────────────────
-- One row per flagged entity per upload session.
-- entity_email is NULL for wallet-abuser and recharge-abuser alerts
-- (those are keyed on a wallet number or recharge number instead).
-- The `detail` JSONB column stores the full result object so the
-- UI can reconstruct cards without joins.
CREATE TABLE IF NOT EXISTS fraud_alerts (
  id                UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  upload_session_id UUID        NOT NULL REFERENCES upload_sessions(id) ON DELETE CASCADE,
  platform          TEXT        NOT NULL,
  alert_type        TEXT        NOT NULL,
  -- alert_type values:
  --   'multi_cc'           – used >3 different CC/wallets/methods
  --   'high_amount'        – single transaction ≥ threshold
  --   'fake_domain'        – non-whitelisted email domain
  --   'wallet_abuser'      – same wallet used by multiple emails
  --   'bnpl_fraud'         – multiple failed BNPL attempts
  --   'pay_method_abuse'   – Admin: >3 payment methods
  --   'suspected_trials'   – Admin: ≥5 failed/trial attempts
  --   'recharge_abuser'    – Admin: same recharge # by multiple emails
  --   'fawry_suspected'    – Fawry: ≥3 transactions by same email
  risk_level          TEXT        NOT NULL,  -- 'High' | 'Mid' | 'HighSuspicious' | 'HighAmount' | 'FakeDomain'
  entity_email        TEXT,                  -- NULL for wallet/recharge abusers
  entity_identifier   TEXT,                  -- wallet #, recharge #, user ID, or same as email
  customer_names      TEXT[]      NOT NULL DEFAULT '{}',
  payment_methods     TEXT[]      NOT NULL DEFAULT '{}',  -- CC cards / wallets / methods flagged
  total_amount        NUMERIC(14,2) NOT NULL DEFAULT 0,
  transaction_count   INTEGER     NOT NULL DEFAULT 0,
  reasons             TEXT[]      NOT NULL DEFAULT '{}',
  detail              JSONB,                 -- full result object for card rendering
  created_at          TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── 4. AUDIT_LOGS ───────────────────────────────────────────────
-- Append-only log of every user action in the app.
CREATE TABLE IF NOT EXISTS audit_logs (
  id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  username     TEXT        NOT NULL,
  action       TEXT        NOT NULL,   -- 'Import' | 'Export' | 'Login' | 'UserAdded' | etc.
  platform     TEXT,
  record_count INTEGER,
  details      TEXT,
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── 5. APP_USERS ────────────────────────────────────────────────
-- Replaces the sessionStorage-based user list in the component.
-- Passwords stored as bcrypt hashes (never plaintext).
-- Seeded with initial users via INSERT below.
CREATE TABLE IF NOT EXISTS app_users (
  id              UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  username        TEXT        UNIQUE NOT NULL,
  password_hash   TEXT        NOT NULL,
  role            TEXT        NOT NULL DEFAULT 'user' CHECK (role IN ('user','superadmin')),
  security_answer TEXT,
  is_active       BOOLEAN     NOT NULL DEFAULT true,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ── INDEXES ─────────────────────────────────────────────────────
-- Optimise the most common query patterns:
-- • look up all alerts for a session            (cascade deletes + listing)
-- • look up all alerts for a specific email     (customer history)
-- • filter alerts by platform or risk level     (dashboard filters)
-- • recent audit log entries                    (audit tab)

CREATE INDEX IF NOT EXISTS idx_fraud_alerts_session
  ON fraud_alerts(upload_session_id);

CREATE INDEX IF NOT EXISTS idx_fraud_alerts_email
  ON fraud_alerts(entity_email)
  WHERE entity_email IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_fraud_alerts_platform
  ON fraud_alerts(platform);

CREATE INDEX IF NOT EXISTS idx_fraud_alerts_risk
  ON fraud_alerts(risk_level);

CREATE INDEX IF NOT EXISTS idx_fraud_alerts_type
  ON fraud_alerts(alert_type);

CREATE INDEX IF NOT EXISTS idx_upload_sessions_platform
  ON upload_sessions(platform);

CREATE INDEX IF NOT EXISTS idx_upload_sessions_created
  ON upload_sessions(created_at DESC);

CREATE INDEX IF NOT EXISTS idx_audit_logs_username
  ON audit_logs(username);

CREATE INDEX IF NOT EXISTS idx_audit_logs_created
  ON audit_logs(created_at DESC);

CREATE INDEX IF NOT EXISTS idx_customers_email
  ON customers(email);

-- ── 6. BLOCKED_ENTITIES ─────────────────────────────────────────
-- Globally blocked emails. Filtered out at render time across all platforms.
CREATE TABLE IF NOT EXISTS blocked_entities (
  id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  entity_value TEXT        UNIQUE NOT NULL,   -- lowercased email
  entity_type  TEXT        NOT NULL DEFAULT 'email',
  blocked_by   TEXT        NOT NULL,          -- username who blocked
  platform     TEXT,                          -- platform where it was flagged
  note         TEXT,                          -- optional reason
  created_at   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_blocked_entities_value
  ON blocked_entities(entity_value);

-- ── SEED: initial app users ─────────────────────────────────────
-- Passwords below are PLAINTEXT placeholders.
-- Replace password_hash values with bcrypt hashes before going live,
-- or keep as-is for initial testing (the app will validate raw strings
-- until bcrypt is wired in).
INSERT INTO app_users (username, password_hash, role) VALUES
  ('Yaheia.adel',   'asd123456**',  'user'),
  ('Traek.nabil',   'asd789**',     'user'),
  ('Fatma.saad',    'asd101112**',  'user'),
  ('Mostafa.ezzat', 'asd131415**',  'user')
ON CONFLICT (username) DO NOTHING;
