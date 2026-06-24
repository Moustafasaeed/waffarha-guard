// @ts-nocheck
import type { NextApiRequest, NextApiResponse } from "next";

export const config = {
  api: { bodyParser: { sizeLimit: "10mb" } },
};
import {
  getHistoricalContext,
  buildHistoricalContextString,
} from "../../lib/claude-analysis";

export default async function handler(
  req: NextApiRequest,
  res: NextApiResponse
) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
  const {
    platform,
    recordCount,
    fraud = [],
    highAmt = [],
    fakeDom = [],
    otherFlags = [],
    allStats = null,
  } = req.body;

  if (!platform) {
    return res.status(400).json({ error: "platform is required" });
  }

  // Collect all flagged emails for RAG lookup
  const allFlagged = [...fraud, ...highAmt, ...fakeDom, ...otherFlags];
  const emails = [
    ...new Set(
      allFlagged
        .map((r: any) => r.email || r.entity_email)
        .filter(Boolean) as string[]
    ),
  ];

  // Pull historical fraud data from Supabase (RAG context)
  const historicalContext = await getHistoricalContext(emails);
  const historicalText = buildHistoricalContextString(historicalContext, emails);

  // Summarize current batch (top cases only, to keep prompt size manageable)
  const highRisk = fraud.filter((r: any) => r.risk === "High");
  const midRisk = fraud.filter((r: any) => r.risk === "Mid");

  const topFraud = [...highRisk.slice(0, 6), ...midRisk.slice(0, 4)].map(
    (r: any) => ({
      email: r.email,
      risk: r.risk,
      uniqueMethods: r.uniqueCCs?.length ?? 0,
      totalAmt: r.totalAmt ?? 0,
      txCount: r.txCount ?? 0,
      reasons: r.reasons ?? [],
    })
  );

  const topHighAmt = highAmt.slice(0, 5).map((r: any) => ({
    email: r.email,
    maxAmt: r.maxAmt ?? r.totalAmt ?? 0,
    txCount: r.txCount ?? 0,
    reasons: r.reasons ?? [],
  }));

  const topFakeDom = fakeDom.slice(0, 5).map((r: any) => ({
    email: r.email,
    domain: r.domain ?? (r.email?.split("@")[1] || "unknown"),
    totalAmt: r.totalAmt ?? 0,
  }));

  const topOther = otherFlags.slice(0, 5).map((r: any) => ({
    identifier: r.email ?? r.wallet ?? r.recharge ?? r.userId,
    risk: r.risk,
    txCount: r.txCount ?? 0,
    reasons: r.reasons ?? [],
  }));

  const systemPrompt = `You are a senior fraud analyst for Waffarha, an Egyptian e-commerce coupon and savings platform. Waffarha partners with payment processors like PayTabs, Noon, PayMob, and Fawry. Transactions are in EGP.

Your role:
1. Analyze detected fraud patterns across the current batch
2. Correlate with historical data to identify repeat and cross-platform offenders
3. Predict likely next fraud moves based on observed patterns
4. Provide specific, actionable recommendations

Always cite evidence. Be direct. No filler. Respond ONLY with raw JSON — no markdown, no code fences.`;

  const allStatsBlock = allStats ? `
ALL RECORDS STATISTICS (entire uploaded file — not just flagged):
- Total records: ${allStats.totalRecords}
- Unique customers: ${allStats.uniqueEmails}
- Total transaction volume: EGP ${allStats.totalAmount.toLocaleString()}
- Amount distribution: p50=EGP ${allStats.amountP50}, p90=EGP ${allStats.amountP90}, p99=EGP ${allStats.amountP99}, max=EGP ${allStats.maxAmount.toLocaleString()}
- Zero-amount records: ${allStats.zeroAmountCount}
- Flagged rate: ${[...fraud, ...highAmt, ...fakeDom, ...otherFlags].length} flagged entities out of ${allStats.uniqueEmails} unique customers
- Top email domains: ${allStats.topDomains.join(", ")}` : "";

  const userMessage = `Analyze this fraud detection batch from ${platform}:

BATCH STATISTICS:
- Total records: ${recordCount}
- High Risk (CC/wallet abuse): ${highRisk.length}
- Mid Risk: ${midRisk.length}
- High Amounts: ${highAmt.length}
- Fake Domains: ${fakeDom.length}
- Other flags: ${otherFlags.length}${allStatsBlock}

TOP CC/WALLET FRAUD CASES:
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

  try {
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
      return res.status(500).json({
        error: "Claude returned non-JSON response",
        raw: textBlock.text,
      });
    }

    return res.status(200).json({
      analysis,
      meta: {
        emailsChecked: emails.length,
        historicalHits: historicalContext.totalHistoricalHits,
        crossPlatformCount: historicalContext.crossPlatformEmails.length,
        repeatOffenderCount: Object.keys(historicalContext.repeatOffenders).length,
      },
    });
  } catch (err: any) {
    console.error("[analyze-fraud API] inner:", err);
    return res.status(500).json({ error: err.message ?? "Claude API error" });
  }

  } catch (err: any) {
    console.error("[analyze-fraud API] outer:", err);
    return res.status(500).json({ error: err.message ?? "Unexpected server error" });
  }
}
