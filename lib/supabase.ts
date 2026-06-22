// @ts-nocheck
import { createClient } from "@supabase/supabase-js";

const supabaseUrl = process.env.NEXT_PUBLIC_SUPABASE_URL!;
const supabaseAnonKey = process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!;

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

// ── Types ──────────────────────────────────────────────────────────────────────

export type Platform =
  | "paytabs"
  | "noon"
  | "paymob_wallets"
  | "paymob_bnpl"
  | "fawry"
  | "admin"
  | "promo";

export type AlertType =
  | "multi_cc"
  | "high_amount"
  | "fake_domain"
  | "wallet_abuser"
  | "bnpl_fraud"
  | "pay_method_abuse"
  | "suspected_trials"
  | "recharge_abuser"
  | "fawry_suspected"
  | "promo_high_discount"
  | "promo_same_card"
  | "promo_same_wallet"
  | "promo_fake_domain";

export interface UploadSessionPayload {
  platform: Platform;
  uploaded_by: string;
  filename?: string;
  record_count: number;
  high_count: number;
  mid_count: number;
  high_amt_count: number;
  fake_dom_count: number;
  other_count?: number;
}

export interface FraudAlertPayload {
  platform: Platform;
  alert_type: AlertType;
  risk_level: string;
  entity_email?: string;
  entity_identifier?: string;
  customer_names?: string[];
  payment_methods?: string[];
  total_amount?: number;
  transaction_count?: number;
  reasons?: string[];
  detail?: object;
}

// ── upsertCustomer ─────────────────────────────────────────────────────────────
// Ensures a customer row exists for an email; updates last_seen_at on conflict.
async function upsertCustomer(email: string, isDisposable: boolean) {
  if (!email) return;
  await supabase.from("customers").upsert(
    { email, is_disposable: isDisposable, last_seen_at: new Date().toISOString() },
    { onConflict: "email" }
  );
}

// ── logUploadSession ───────────────────────────────────────────────────────────
// Creates an upload_session row and returns its id.
export async function logUploadSession(
  payload: UploadSessionPayload
): Promise<string | null> {
  const { data, error } = await supabase
    .from("upload_sessions")
    .insert({
      platform: payload.platform,
      uploaded_by: payload.uploaded_by,
      filename: payload.filename ?? null,
      record_count: payload.record_count,
      high_count: payload.high_count,
      mid_count: payload.mid_count,
      high_amt_count: payload.high_amt_count,
      fake_dom_count: payload.fake_dom_count,
      other_count: payload.other_count ?? 0,
    })
    .select("id")
    .single();

  if (error) {
    console.error("[Supabase] logUploadSession error:", error.message);
    return null;
  }
  return data?.id ?? null;
}

// ── logFraudAlerts ─────────────────────────────────────────────────────────────
// Inserts all fraud alert rows for a session in one batch call.
export async function logFraudAlerts(
  sessionId: string,
  alerts: FraudAlertPayload[]
) {
  if (!sessionId || alerts.length === 0) return;

  // Upsert customers in parallel (fire-and-forget, best-effort)
  const emailSet = new Set(
    alerts.map((a) => a.entity_email).filter(Boolean) as string[]
  );
  await Promise.allSettled(
    [...emailSet].map((email) =>
      upsertCustomer(
        email,
        alerts.find((a) => a.entity_email === email)?.alert_type === "fake_domain"
      )
    )
  );

  const rows = alerts.map((a) => ({
    upload_session_id: sessionId,
    platform: a.platform,
    alert_type: a.alert_type,
    risk_level: a.risk_level,
    entity_email: a.entity_email ?? null,
    entity_identifier: a.entity_identifier ?? a.entity_email ?? null,
    customer_names: a.customer_names ?? [],
    payment_methods: a.payment_methods ?? [],
    total_amount: a.total_amount ?? 0,
    transaction_count: a.transaction_count ?? 0,
    reasons: a.reasons ?? [],
    detail: a.detail ?? null,
  }));

  const { error } = await supabase.from("fraud_alerts").insert(rows);
  if (error) {
    console.error("[Supabase] logFraudAlerts error:", error.message);
  }
}

// ── logAuditEntry ──────────────────────────────────────────────────────────────
export async function logAuditEntry(entry: {
  username: string;
  action: string;
  platform?: string;
  record_count?: number;
  details?: string;
}) {
  const { error } = await supabase.from("audit_logs").insert({
    username: entry.username,
    action: entry.action,
    platform: entry.platform ?? null,
    record_count: entry.record_count ?? null,
    details: entry.details ?? null,
  });
  if (error) {
    console.error("[Supabase] logAuditEntry error:", error.message);
  }
}

// ── buildAlerts helpers ────────────────────────────────────────────────────────
// Converts the app's detection result arrays into FraudAlertPayload rows.

// ── Blocked entities ───────────────────────────────────────────────────────────

export async function loadBlockedEmails(): Promise<
  { entity_value: string; blocked_by: string; platform: string | null; note: string | null; created_at: string }[]
> {
  const { data, error } = await supabase
    .from("blocked_entities")
    .select("entity_value, blocked_by, platform, note, created_at")
    .eq("entity_type", "email")
    .order("created_at", { ascending: false });
  if (error) {
    console.error("[Supabase] loadBlockedEmails error:", error.message);
    return [];
  }
  return data ?? [];
}

export async function blockEmail(
  email: string,
  blocked_by: string,
  platform?: string,
  note?: string
): Promise<boolean> {
  const { error } = await supabase.from("blocked_entities").upsert(
    {
      entity_value: email.toLowerCase().trim(),
      entity_type: "email",
      blocked_by,
      platform: platform ?? null,
      note: note ?? null,
    },
    { onConflict: "entity_value" }
  );
  if (error) {
    console.error("[Supabase] blockEmail error:", error.message);
    return false;
  }
  return true;
}

export async function unblockEmail(email: string): Promise<boolean> {
  const { error } = await supabase
    .from("blocked_entities")
    .delete()
    .eq("entity_value", email.toLowerCase().trim());
  if (error) {
    console.error("[Supabase] unblockEmail error:", error.message);
    return false;
  }
  return true;
}

export function buildCCAlerts(
  fraud: any[],
  platform: Platform
): FraudAlertPayload[] {
  return fraud.map((r) => ({
    platform,
    alert_type: "multi_cc",
    risk_level: r.risk,
    entity_email: r.email,
    entity_identifier: r.email,
    customer_names: r.custNames ?? [],
    payment_methods: r.uniqueCCs ?? [],
    total_amount: r.totalAmt ?? 0,
    transaction_count: r.txCount ?? 0,
    reasons: r.reasons ?? [],
    detail: r,
  }));
}

export function buildHighAmtAlerts(
  rows: any[],
  platform: Platform
): FraudAlertPayload[] {
  return rows.map((r) => ({
    platform,
    alert_type: "high_amount",
    risk_level: "HighAmount",
    entity_email: r.email,
    entity_identifier: r.email,
    customer_names: r.custNames ?? [],
    payment_methods: r.uniqueCCs ?? [],
    total_amount: r.totalAmt ?? 0,
    transaction_count: r.txCount ?? 0,
    reasons: r.reasons ?? [],
    detail: r,
  }));
}

export function buildFakeDomAlerts(
  rows: any[],
  platform: Platform
): FraudAlertPayload[] {
  return rows.map((r) => ({
    platform,
    alert_type: "fake_domain",
    risk_level: "FakeDomain",
    entity_email: r.email,
    entity_identifier: r.email,
    customer_names: r.custNames ?? [],
    payment_methods: r.uniqueCCs ?? [],
    total_amount: r.totalAmt ?? 0,
    transaction_count: r.txCount ?? 0,
    reasons: r.reasons ?? [],
    detail: r,
  }));
}

export function buildWalletAbuserAlerts(
  rows: any[],
  platform: Platform
): FraudAlertPayload[] {
  return rows.map((r) => ({
    platform,
    alert_type: "wallet_abuser",
    risk_level: r.risk,
    entity_email: null,
    entity_identifier: r.wallet,
    customer_names: r.emails ?? [],
    payment_methods: [r.wallet],
    total_amount: r.totalAmt ?? 0,
    transaction_count: r.txCount ?? 0,
    reasons: r.reasons ?? [],
    detail: r,
  }));
}

export function buildBNPLAlerts(
  rows: any[],
  platform: Platform
): FraudAlertPayload[] {
  return rows.map((r) => ({
    platform,
    alert_type: "bnpl_fraud",
    risk_level: r.risk ?? "HighSuspicious",
    entity_email: r.email,
    entity_identifier: r.email,
    customer_names: r.custNames ?? [],
    payment_methods: r.uniqueCCs ?? [],
    total_amount: r.totalAmt ?? 0,
    transaction_count: r.txCount ?? 0,
    reasons: r.reasons ?? [],
    detail: r,
  }));
}

export function buildAdminAlerts(
  payMethods: any[],
  suspected: any[],
  highAmt: any[],
  fakeDom: any[],
  recharge: any[]
): FraudAlertPayload[] {
  return [
    ...payMethods.map((r) => ({
      platform: "admin" as Platform,
      alert_type: "pay_method_abuse" as AlertType,
      risk_level: r.risk,
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: r.uniqueMethods ?? [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...suspected.map((r) => ({
      platform: "admin" as Platform,
      alert_type: "suspected_trials" as AlertType,
      risk_level: "HighSuspicious",
      entity_email: r.email,
      entity_identifier: r.userId,
      customer_names: r.custNames ?? [],
      payment_methods: [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...highAmt.map((r) => ({
      platform: "admin" as Platform,
      alert_type: "high_amount" as AlertType,
      risk_level: "HighAmount",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: r.uniqueMethods ?? [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...fakeDom.map((r) => ({
      platform: "admin" as Platform,
      alert_type: "fake_domain" as AlertType,
      risk_level: "FakeDomain",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...recharge.map((r) => ({
      platform: "admin" as Platform,
      alert_type: "recharge_abuser" as AlertType,
      risk_level: r.risk,
      entity_email: null,
      entity_identifier: r.recharge,
      customer_names: r.emails ?? [],
      payment_methods: [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
  ];
}

export function buildPromoAlerts(
  highDiscount: any[],
  sameCard: any[],
  sameWallet: any[],
  fakeDom: any[]
): FraudAlertPayload[] {
  return [
    ...highDiscount.map((r) => ({
      platform: "promo" as Platform,
      alert_type: "promo_high_discount" as AlertType,
      risk_level: "HighDiscount",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: r.promoCodes ?? [],
      total_amount: r.totalDiscount ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...sameCard.map((r) => ({
      platform: "promo" as Platform,
      alert_type: "promo_same_card" as AlertType,
      risk_level: r.risk,
      entity_email: null,
      entity_identifier: r.card,
      customer_names: r.emails ?? [],
      payment_methods: [r.card],
      total_amount: r.totalDiscount ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...sameWallet.map((r) => ({
      platform: "promo" as Platform,
      alert_type: "promo_same_wallet" as AlertType,
      risk_level: r.risk,
      entity_email: null,
      entity_identifier: r.wallet,
      customer_names: r.emails ?? [],
      payment_methods: [r.wallet],
      total_amount: r.totalDiscount ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...fakeDom.map((r) => ({
      platform: "promo" as Platform,
      alert_type: "promo_fake_domain" as AlertType,
      risk_level: "FakeDomain",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: r.promoCodes ?? [],
      total_amount: r.totalDiscount ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
  ];
}

export function buildFawryAlerts(
  highAmt: any[],
  suspected: any[],
  fakeDom: any[]
): FraudAlertPayload[] {
  return [
    ...highAmt.map((r) => ({
      platform: "fawry" as Platform,
      alert_type: "high_amount" as AlertType,
      risk_level: "HighAmount",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...suspected.map((r) => ({
      platform: "fawry" as Platform,
      alert_type: "fawry_suspected" as AlertType,
      risk_level: "HighSuspicious",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
    ...fakeDom.map((r) => ({
      platform: "fawry" as Platform,
      alert_type: "fake_domain" as AlertType,
      risk_level: "FakeDomain",
      entity_email: r.email,
      entity_identifier: r.email,
      customer_names: r.custNames ?? [],
      payment_methods: [],
      total_amount: r.totalAmt ?? 0,
      transaction_count: r.txCount ?? 0,
      reasons: r.reasons ?? [],
      detail: r,
    })),
  ];
}
