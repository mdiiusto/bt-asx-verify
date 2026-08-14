// Blackwood & Terrace — ASX Verification Server (Railway version)
// -----------------------------------------------------------------
// Same job as the Cloudflare Worker version, built as a small Node/Express
// app for Railway deployment instead.

const express = require("express");
const app = express();
app.use(express.json());

const MODEL = "claude-sonnet-5";
const BATCH_SIZE = 15;
const PORT = process.env.PORT || 3000;
const API_KEY = process.env.ANTHROPIC_API_KEY;

// CORS — allow calls from your hosted tool. Once deployed, you can tighten
// this to your exact domain (e.g. "https://blackwoodterrace.com.au")
// instead of "*" for better security.
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(200);
  next();
});

app.get("/", (req, res) => {
  res.send("BT ASX Verification server is running. POST to /verify with { names: [...] }.");
});

app.post("/verify", async (req, res) => {
  if (!API_KEY) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY environment variable not set on this server." });
  }

  const names = Array.isArray(req.body.names) ? req.body.names.filter(Boolean) : [];
  if (!names.length) {
    return res.status(400).json({ error: "No names provided" });
  }

  try {
    let results = [];
    for (let i = 0; i < names.length; i += BATCH_SIZE) {
      const batch = names.slice(i, i + BATCH_SIZE);
      const batchResults = await verifyBatch(batch);
      results = results.concat(batchResults);
    }
    res.json({ results });
  } catch (err) {
    res.status(500).json({ error: err.message || "Verification failed" });
  }
});

async function verifyBatch(names) {
  const systemPrompt = `You are checking Western Australian mining company/holder names against public records to determine if each is ASX-listed, or a subsidiary of an ASX-listed or other publicly-listed company, or foreign state-owned.

For each name given, use web search to determine:
- is_listed: true if ASX-listed, a subsidiary of an ASX-listed company, listed on another exchange, or foreign state-owned. false if it appears to be a genuinely private company with no listed or state-owned parent found.
- parent: the ultimate listed/state-owned parent company name if is_listed is true, else null.
- confidence: "high" if you found clear, specific evidence; "low" if uncertain or nothing specific was found.
- note: one short sentence explaining the finding.

Respond with ONLY a JSON array, no other text, no markdown code fences. Format:
[{"name":"...","is_listed":true,"parent":"...","confidence":"high","note":"..."}, ...]

Every name in the input must appear exactly once in the output, using the exact same name string given.`;

  const userMessage = "Check these names:\n" + names.map(n => "- " + n).join("\n");

  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 4000,
      system: systemPrompt,
      messages: [{ role: "user", content: userMessage }],
      tools: [{ type: "web_search_20250305", name: "web_search" }],
    }),
  });

  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`Anthropic API error (${resp.status}): ${errText.slice(0, 300)}`);
  }

  const data = await resp.json();
  const textBlocks = (data.content || []).filter(b => b.type === "text").map(b => b.text);
  const fullText = textBlocks.join("\n").trim();

  const jsonMatch = fullText.match(/\[[\s\S]*\]/);
  if (!jsonMatch) {
    return names.map(n => ({ name: n, is_listed: null, parent: null, confidence: "low", note: "Could not parse a result — try again or check manually." }));
  }

  try {
    const parsed = JSON.parse(jsonMatch[0]);
    const byName = {};
    parsed.forEach(r => { byName[r.name] = r; });
    return names.map(n => byName[n] || { name: n, is_listed: null, parent: null, confidence: "low", note: "No result returned for this name — check manually." });
  } catch (e) {
    return names.map(n => ({ name: n, is_listed: null, parent: null, confidence: "low", note: "Could not parse a result — try again or check manually." }));
  }
}

app.listen(PORT, () => console.log(`BT ASX verification server listening on port ${PORT}`));
