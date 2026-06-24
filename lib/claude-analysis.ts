import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.SUPABASE_SERVICE_ROLE_KEY!
);

export interface HistoricalAlert {
  entity_email: string | null;
  platform: string;
  alert_type: string;
  risk_level: string;
  total_amount: number;
  transaction_count: number;
  created_at: string;
  reasons: string[];
}

export interface CleanHistory {
  email: string;
  totalAmount: number;
  totalTxCount: number;
  platforms: string[];
  uploadCount: number;
}

export interface HistoricalContext {
  repeatOffenders: Record<string, HistoricalAlert[]>;
  totalHistoricalHits: number;
  crossPlatformEmails: string[];
  cleanHistory: Record<string, CleanHistory>;
}

export async function getHistoricalContext(
  emails: string[]
): Promise<HistoricalContext> {
  if (!emails.length) {
    return { repeatOffenders: {}, totalHistoricalHits: 0, crossPlatformEmails: [], cleanHistory: {} };
  }

  const emailSlice = emails.slice(0, 50);

  const [fraudResult, cleanResult] = await Promise.all([
    supabase
      .from("fraud_alerts")
      .select("entity_email, platform, alert_type, risk_level, total_amount, transaction_count, created_at, reasons")
      .in("entity_email", emailSlice)
      .order("created_at", { ascending: false })
      .limit(300),
    supabase
      .from("all_transactions")
      .select("entity_email, platform, amount, transaction_count, created_at")
      .in("entity_email", emailSlice)
      .order("created_at", { ascending: false })
      .limit(500),
  ]);

  if (fraudResult.error) {
    console.error("[claude-analysis] fraud_alerts query error:", fraudResult.error.message);
  }

  const data = fraudResult.data ?? [];
  const cleanData = cleanResult.data ?? [];

  const repeatOffenders: Record<string, HistoricalAlert[]> = {};
  data.forEach((alert) => {
    if (!alert.entity_email) return;
    if (!repeatOffenders[alert.entity_email]) {
      repeatOffenders[alert.entity_email] = [];
    }
    repeatOffenders[alert.entity_email].push(alert as HistoricalAlert);
  });

  const crossPlatformEmails = Object.entries(repeatOffenders)
    .filter(([, alerts]) => {
      const platforms = new Set(alerts.map((a) => a.platform));
      return platforms.size > 1;
    })
    .map(([email]) => email);

  const cleanHistory: Record<string, CleanHistory> = {};
  cleanData.forEach((row) => {
    if (!row.entity_email) return;
    if (!cleanHistory[row.entity_email]) {
      cleanHistory[row.entity_email] = { email: row.entity_email, totalAmount: 0, totalTxCount: 0, platforms: [], uploadCount: 0 };
    }
    const ch = cleanHistory[row.entity_email];
    ch.totalAmount += row.amount || 0;
    ch.totalTxCount += row.transaction_count || 0;
    ch.uploadCount += 1;
    if (!ch.platforms.includes(row.platform)) ch.platforms.push(row.platform);
  });

  return {
    repeatOffenders,
    totalHistoricalHits: data.length,
    crossPlatformEmails,
    cleanHistory,
  };
}

export function buildHistoricalContextString(
  context: HistoricalContext,
  currentEmails: string[]
): string {
  const { repeatOffenders, crossPlatformEmails, cleanHistory } = context;
  const lines: string[] = [];

  const repeats = currentEmails.filter((e) => repeatOffenders[e]?.length > 0).slice(0, 20);

  if (repeats.length > 0) {
    lines.push("REPEAT OFFENDERS (previously flagged in system):");
    repeats.forEach((email) => {
      const history = repeatOffenders[email];
      const platforms = [...new Set(history.map((h) => h.platform))].join(", ");
      const totalPrevAmt = history.reduce((s, h) => s + (h.total_amount || 0), 0);
      const alertTypes = [...new Set(history.map((h) => h.alert_type))].join(", ");
      lines.push(
        `  • ${email}: ${history.length} prior flags · platforms=[${platforms}] · types=[${alertTypes}] · historical total=EGP ${totalPrevAmt.toLocaleString()}`
      );
    });
  }

  if (crossPlatformEmails.length > 0) {
    lines.push("\nCROSS-PLATFORM OFFENDERS (same email flagged on multiple platforms):");
    crossPlatformEmails.slice(0, 10).forEach((email) => {
      const history = repeatOffenders[email];
      const byPlatform: Record<string, number> = {};
      history.forEach((h) => {
        byPlatform[h.platform] = (byPlatform[h.platform] || 0) + 1;
      });
      const detail = Object.entries(byPlatform)
        .map(([p, n]) => `${p}×${n}`)
        .join(", ");
      const totalAmt = history.reduce((s, h) => s + (h.total_amount || 0), 0);
      lines.push(`  • ${email}: [${detail}] — total EGP ${totalAmt.toLocaleString()}`);
    });
  }

  // Clean history: flagged emails that also have normal transaction history
  const cleanEntries = currentEmails
    .filter((e) => cleanHistory[e] && cleanHistory[e].totalTxCount > 0)
    .slice(0, 15);

  if (cleanEntries.length > 0) {
    lines.push("\nCLEAN TRANSACTION HISTORY (all-time activity for currently flagged emails):");
    cleanEntries.forEach((email) => {
      const ch = cleanHistory[email];
      lines.push(
        `  • ${email}: ${ch.totalTxCount} total transactions · EGP ${Math.round(ch.totalAmount).toLocaleString()} lifetime · seen across [${ch.platforms.join(", ")}] · ${ch.uploadCount} upload(s)`
      );
    });
  }

  if (lines.length === 0) {
    return "No historical records found for the currently flagged emails — these appear to be first-time detections.";
  }

  return lines.join("\n");
}
