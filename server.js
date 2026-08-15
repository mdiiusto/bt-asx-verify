// Blackwood & Terrace — Deal Sourcing Server
// -----------------------------------------------------------------
// (Node/Express; run with `npm start`.)
// Serves the workbench frontend and runs three AI-backed jobs:
//   POST /verify          — bulk ASX/listed-status check on a list of names
//   POST /research        — deep research on a single opportunity
//   POST /generate-brief  — compiles selected opportunities into a Word doc

const express = require("express");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle } = require("docx");
const app = express();
app.use(express.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "public")));

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

app.post("/research", async (req, res) => {
  if (!API_KEY) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY environment variable not set on this server." });
  }
  const opp = req.body.opportunity;
  if (!opp || (!opp.site_name && !opp.project_name)) {
    return res.status(400).json({ error: "No opportunity data provided" });
  }

  try {
    const text = await researchOpportunity(opp);
    res.json({ research: text });
  } catch (err) {
    res.status(500).json({ error: err.message || "Research failed" });
  }
});

app.post("/generate-brief", async (req, res) => {
  const { clientName, opportunities } = req.body;
  if (!Array.isArray(opportunities) || !opportunities.length) {
    return res.status(400).json({ error: "No opportunities provided" });
  }
  try {
    const buffer = await buildBriefDocx(clientName || "Client", opportunities);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="BT_Brief_${(clientName||'Client').replace(/[^a-z0-9]/gi,'_')}.docx"`);
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message || "Brief generation failed" });
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

async function researchOpportunity(opp) {
  const systemPrompt = `You are a research analyst for Blackwood & Terrace, an Australian deal origination firm connecting offshore private capital with off-market Australian mining opportunities. You research WA mining projects and tenement holders for a private client brief.

For the project/tenement given, use web search to find and report on:
1. Ownership — who actually owns/controls it (private, ASX-listed, foreign parent, state-owned). Be specific and cite what you find.
2. Scale — resource estimate, mine life, production figures, DFS/PFS economics (NPV, IRR) if publicly available.
3. Recent developments — funding rounds, JV agreements, corporate activity, news in the last 12-24 months.
4. Contact — any named individuals, investor relations contacts, or company addresses/emails found in public records.
5. Assessment — one paragraph: is this genuinely off-market and approachable, or already spoken for (listed, JV'd, state-owned)? State this plainly.

Write in plain prose, organized under short headers, 300-500 words total. This is for a private research file, not a public document — be direct and factual, cite where information came from (company name, article, or "MINEDEX record").`;

  const userMessage = `Research this WA mining opportunity:
Site/Project: ${opp.site_name || opp.project_name || ''}
Project name: ${opp.project_name || ''}
Commodity: ${opp.commodity_group || ''} (target: ${opp.target_commodity || ''})
Stage: ${opp.site_stage || ''}
Tenement: ${opp.tenement_id || ''}
Registered holder: ${opp.holder1 || ''}
${opp.web_link ? 'MINEDEX record: ' + opp.web_link : ''}`;

  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 2000,
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
  return textBlocks.join("\n\n").trim() || "No research output returned — try again.";
}

const CHARCOAL = "14110D";
const BRASS = "A98D4B";
const WARMGREY = "7A7264";

async function buildBriefDocx(clientName, opportunities) {
  const children = [];

  children.push(
    new Paragraph({
      spacing: { after: 60 },
      children: [new TextRun({ text: "Blackwood & Terrace", bold: true, size: 44, color: CHARCOAL, font: "Georgia" })],
    }),
    new Paragraph({
      spacing: { after: 300 },
      children: [new TextRun({ text: "Off-Market Deal Origination — Australia", size: 20, color: BRASS, font: "Georgia" })],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [new TextRun({ text: `Prepared for: ${clientName}`, bold: true, size: 24 })],
    }),
    new Paragraph({
      spacing: { after: 300 },
      children: [new TextRun({ text: `Date: ${new Date().toLocaleDateString('en-AU', { day: 'numeric', month: 'long', year: 'numeric' })}`, size: 20, color: WARMGREY })],
    }),
  );

  opportunities.forEach((opp, idx) => {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 120 },
        children: [new TextRun({ text: `${idx + 1}. ${opp.site_name || opp.project_name || 'Untitled opportunity'}`, color: CHARCOAL, font: "Georgia" })],
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({ text: `${opp.commodity_group || ''}${opp.target_commodity ? ' — ' + opp.target_commodity : ''}${opp.site_stage ? ' — ' + opp.site_stage : ''}`, italics: true, color: WARMGREY, size: 20 })],
      }),
    );

    const factRows = [
      ["Project", opp.project_name || ''],
      ["Tenement", opp.tenement_id || ''],
      ["Type / Status", `${opp.tenement_type || ''} / ${opp.tenement_status || ''}`],
      ["Holder", opp.holder1 || ''],
      ["Area", opp.area ? `${opp.area} ${opp.area_unit || ''}` : ''],
      ["Expires", opp.end_date || ''],
    ].filter(r => r[1] && r[1].trim() !== ' /');

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 1, color: "E5E0D5" },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: "E5E0D5" },
        left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "E5E0D5" },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: factRows.map(([label, value]) => new TableRow({
        children: [
          new TableCell({ width: { size: 25, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, color: WARMGREY })] })] }),
          new TableCell({ width: { size: 75, type: WidthType.PERCENTAGE }, children: [new Paragraph({ children: [new TextRun({ text: value, size: 18 })] })] }),
        ],
      })),
    }));

    if (opp.research) {
      children.push(
        new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: "Research", bold: true, size: 20, color: CHARCOAL })] }),
        ...opp.research.split(/\n+/).filter(Boolean).map(line => new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: line, size: 19 })] })),
      );
    }
    if (opp.sales_pitch) {
      children.push(
        new Paragraph({ spacing: { before: 160, after: 80 }, children: [new TextRun({ text: "Why this fits", bold: true, size: 20, color: CHARCOAL })] }),
        new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text: opp.sales_pitch, size: 19 })] }),
      );
    }
    if (opp.web_link) {
      children.push(new Paragraph({ spacing: { before: 120 }, children: [new TextRun({ text: "Source: " + opp.web_link, size: 16, color: WARMGREY })] }));
    }
  });

  const doc = new Document({
    sections: [{ properties: {}, children }],
  });

  return await Packer.toBuffer(doc);
}

app.listen(PORT, () => console.log(`BT ASX verification server listening on port ${PORT}`));
