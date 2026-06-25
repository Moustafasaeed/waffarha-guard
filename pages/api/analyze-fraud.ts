// @ts-nocheck
import type { NextApiRequest, NextApiResponse } from "next";

export const config = {
  api: { bodyParser: { sizeLimit: "1mb" } },
};

import {
  getSessionData,
  getHistoricalContext,
  buildHistoricalContextString,
} from "../../lib/claude-analysis";

const FRAUD_TYPES = new Set(["multi_cc","wallet_abuser","bnpl_fraud","pay_method_abuse","suspected_trials","recharge_abuser","fawry_suspected","promo_same_card","promo_same_wallet"]);
const HIGH_AMT_TYPES = new Set(["high_amount","promo_high_discount"]);
const FAKE_DOM_TYPES = new Set(["fake_domain","promo_fake_domain"]);

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { platform, sessionId, recordCount } = req.body;

    if (!platform || !sessionId) {
      return res.status(400).json({ error: "platform and sessionId are required" });
    }

    // Pull current session data from Supabase
    const { session, alerts, allStats } = await getSessionData(sessionId);

    // Group alerts by type
    const fraudAlerts = alerts.filter((a: any) => FRAUD_TYPES.has(a.alert_type));
    const highAmtAlerts = alerts.filter((a: any) => HIGH_AMT_TYPES.has(a.alert_type));
    const fakeDomAlerts = alerts.filter((a: any) => FAKE_DOM_TYPES.has(a.alert_type));
    const otherAlerts = alerts.filter((a: any) => !FRAUD_TYPES.has(a.alert_type) && !HIGH_AMT_TYPES.has(a.alert_type) && !FAKE_DOM_TYPES.has(a.alert_type));

    // Extract unique emails for RAG lookup
    const emails = [...new Set(alerts.map((a: any) => a.entity_email).filter(Boolean))] as string[];

    // Pull historical context from Supabase (RAG)
    const historicalContext = await getHistoricalContext(emails);
    const historicalText = buildHistoricalContextString(historicalContext, emails);

    // Build top-cases summaries (cap size for prompt efficiency)
    const topFraud = fraudAlerts.slice(0, 10).map((a: any) => ({
      email: a.entity_email ?? a.entity_identifier,
      type: a.alert_type,
      risk: a.risk_level,
      methods: a.payment_methods?.length ?? 0,
      totalAmt: a.total_amount ?? 0,
      txCount: a.transaction_count ?? 0,
      reasons: a.reasons ?? [],
    }));

    const topHighAmt = highAmtAlerts.slice(0, 5).map((a: any) => ({
      email: a.entity_email,
      totalAmt: a.total_amount ?? 0,
      txCount: a.transaction_count ?? 0,
      reasons: a.reasons ?? [],
    }));

    const topFakeDom = fakeDomAlerts.slice(0, 5).map((a: any) => ({
      email: a.entity_email,
      domain: a.entity_email?.split("@")[1] ?? "unknown",
      totalAmt: a.total_amount ?? 0,
    }));

    const topOther = otherAlerts.slice(0, 5).map((a: any) => ({
      identifier: a.entity_email ?? a.entity_identifier,
      type: a.alert_type,
      risk: a.risk_level,
      txCount: a.transaction_count ?? 0,
      reasons: a.reasons ?? [],
    }));

    const highRiskCount = fraudAlerts.filter((a: any) => a.risk_level === "High").length;
    const midRiskCount = fraudAlerts.filter((a: any) => a.risk_level === "Mid").length;

    const systemPrompt = `You are a senior fraud analyst for Waffarha, an Egyptian e-commerce coupon and savings platform. Waffarha partners with PayTabs, Noon, PayMob, and Fawry. Transactions are in EGP.

Your role:
1. Analyze detected fraud patterns across the current batch
2. Correlate with historical data to identify repeat and cross-platform offenders
3. Predict likely next fraud moves based on observed patterns
4. Provide specific, actionable recommendations

Always cite evidence. Be direct. No filler. Respond ONLY with raw JSON — no markdown, no code fences.`;

    const allStatsBlock = `
ALL RECORDS STATISTICS (entire uploaded file — not just flagged):
- Total records: ${allStats.totalRecords}
- Unique customers: ${allStats.uniqueEmails}
- Total transaction volume: EGP ${allStats.totalAmount.toLocaleString()}
- Amount distribution: p50=EGP ${allStats.amountP50}, p90=EGP ${allStats.amountP90}, p99=EGP ${allStats.amountP99}, max=EGP ${allStats.maxAmount.toLocaleString()}
- Zero-amount records: ${allStats.zeroAmountCount}
- Flagged rate: ${alerts.length} flagged entities out of ${allStats.uniqueEmails} unique customers
- Top email domains: ${allStats.topDomains.join(", ")}`;

    const userMessage = `Analyze this fraud detection batch from ${platform}:

BATCH STATISTICS:
- Total records: ${allStats.totalRecords}
- High Risk (CC/wallet/method abuse): ${highRiskCount}
- Mid Risk: ${midRiskCount}
- High Amounts: ${highAmtAlerts.length}
- Fake Domains: ${fakeDomAlerts.length}
- Other flags: ${otherAlerts.length}${allStatsBlock}

TOP CC/WALLET/METHOD FRAUD CASES:
${JSON.stringify(topFraud, null, 2)}

TOP HIGH AMOUNT CASES:
${JSON.stringify(topHighAmt, null, 2)}

FAKE DOMAIN CASES:
${JSON.stringify(topFakeDom, null, 2)}

OTHER FLAGS (wallet abusers / BNPL / suspected trials / recharge):
${JSON.stringify(topOther, null, 2)}

HISTORICAL CONTEXT FROM DATABASE:
${historicalText}

Respond with this exact JSON:
{
  "executive_summary": "2-3 sentences summarizing this batch's fraud picture",
  "risk_level": "low|medium|high|critical",
  "key_patterns": ["specific pattern observed 1", "pattern 2", "..."],
  "repeat_offenders": ["email@...: N prior flags on [platforms], total EGP X"],
  "cross_platform_activity": ["email@...: flagged on Platform A and B with combined EGP X"],
  "predicted_next_moves": ["specific likely next fraud action 1", "..."],
  "priority_actions": [
    { "target": "email or identifier", "action": "block|escalate|monitor|report_to_processor", "reason": "specific reason" }
  ],
  "recommendations": ["actionable system-level recommendation 1", "..."]
}`;

    const claudeRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": process.env.ANTHROPIC_API_KEY || "",
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
      },
      body: JSON.stringify({
        model: "claude-opus-4-8",
        max_tokens: 2048,
        thinking: { type: "adaptive" },
        system: systemPrompt,
        messages: [{ role: "user", content: userMessage }],
      }),
    });

    if (!claudeRes.ok) {
      const errBody = await claudeRes.text();
      return res.status(500).json({ error: `Anthropic API error ${claudeRes.status}: ${errBody}` });
    }

    const claudeData = await claudeRes.json();
    const textBlock = (claudeData.content || []).find((b: any) => b.type === "text");
    if (!textBlock) {
      return res.status(500).json({ error: "No text in Claude response" });
    }

    let analysis: any;
    try {
      const raw = textBlock.text
        .replace(/^```json\s*/i, "")
        .replace(/^```\s*/i, "")
        .replace(/\s*```$/i, "")
        .trim();
      analysis = JSON.parse(raw);
    } catch {
      return res.status(500).json({ error: "Claude returned non-JSON response", raw: textBlock.text });
    }

    return res.status(200).json({
      analysis,
      meta: {
        emailsChecked: emails.length,
        totalAlerts: alerts.length,
        historicalHits: historicalContext.totalHistoricalHits,
        crossPlatformCount: historicalContext.crossPlatformEmails.length,
        repeatOffenderCount: Object.keys(historicalContext.repeatOffenders).length,
      },
    });

  } catch (err: any) {
    console.error("[analyze-fraud API]", err);
    return res.status(500).json({ error: err.message ?? "Unexpected server error" });
  }
}
