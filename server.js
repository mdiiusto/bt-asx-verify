// Blackwood & Terrace — Deal Sourcing Server
// -----------------------------------------------------------------
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

// DMIRS SLIP public endpoints — layer 0 is Minedex sites, layer 3 is tenements.
const MINEDEX_URL = "https://services.slip.wa.gov.au/public/rest/services/SLIP_Public_Services/Industry_and_Mining/MapServer/0/query";
const TENEMENT_URL = "https://services.slip.wa.gov.au/public/rest/services/SLIP_Public_Services/Industry_and_Mining/MapServer/3/query";

// Kept in sync with the same list in public/index.html.
const MAJOR_HOLDERS = [
  'BHP', 'RIO TINTO', 'FORTESCUE', 'FMG', 'NORTHERN STAR', 'NEWCREST',
  'SOUTH32', 'IGO LIMITED', 'IGO LTD', 'MINERAL RESOURCES', 'EVOLUTION MINING',
  'REGIS RESOURCES', 'GOLD FIELDS', 'NEWMONT', 'ILUKA', 'SANDFIRE',
  'OZ MINERALS', '29METALS', '29 METALS', 'GOLDEN GROVE OPERATIONS',
  'PANORAMIC', 'CHALICE MINING', 'RAMELIUS', 'ST BARBARA', 'SILVER LAKE',
  'GENESIS MINERALS', 'CAPRICORN METALS', 'WEST AFRICAN RESOURCES',
  'PERSEUS MINING', 'ALCOA', 'ALUMINA LIMITED', 'TIANQI', 'ALBEMARLE',
  'PILBARA MINERALS', 'LIONTOWN', 'MINRES', 'WESFARMERS', 'CITIC',
  'HANCOCK', 'ROY HILL', 'ATLAS IRON', 'MOUNT GIBSON', 'CHAMPION IRON',
  'GINA RINEHART', 'TALISON', 'GALAXY RESOURCES', 'ALLKEM', 'ARCADIUM',
  'IMDEX', 'ERAMET', 'GLENCORE', 'ANGLO AMERICAN', 'VALE',
  'AERIS RESOURCES', 'ROUND OAK', 'LITHIUM AUSTRALIA', 'LITHOPHILE',
  'VENTUREX', 'DEVELOP GLOBAL', 'ANAX METALS', 'WHIM CREEK METALS',
  'HORSESHOE METALS', 'MURCHISON COPPER MINES',
];

function isMajorHolder(name) {
  if (!name) return false;
  const upper = name.toUpperCase();
  return MAJOR_HOLDERS.some(m => upper.includes(m));
}

function isMajorProject(row) {
  const haystack = [row.project_name, row.site_name, row.short_name].filter(Boolean).join(' ');
  return isMajorHolder(haystack);
}

function esriMsToDate(ms) {
  if (!ms) return '';
  try { return new Date(ms).toISOString().slice(0, 10); } catch (e) { return ''; }
}

// Pulls the full Minedex site list from DMIRS, paging until a short page comes
// back. Mirrors fetchLiveMinedex() in public/index.html.
async function fetchLiveMinedexServer() {
  const outFields = 'oid,gid,site_code,site_title,short_name,site_commo,site_type_,site_sub_t,site_stage,target_com,commodity,proj_code,proj_title,confidenti,point_conf,latitude,longitude,web_link,extract_da';
  let offset = 0;
  const pageSize = 2000;
  const all = [];
  while (true) {
    const params = new URLSearchParams({
      where: '1=1',
      outFields,
      returnGeometry: 'false',
      f: 'json',
      resultOffset: offset,
      resultRecordCount: pageSize,
      orderByFields: 'oid',
    });
    const resp = await fetch(MINEDEX_URL + '?' + params.toString());
    const data = await resp.json();
    if (data.error) throw new Error(data.error.message || 'Server error');
    const feats = data.features || [];
    feats.forEach(f => all.push(f.attributes));
    if (feats.length < pageSize) break;
    offset += feats.length;
  }
  return all.map(a => ({
    site_code: a.site_code, site_name: a.site_title, short_name: a.short_name,
    site_commodity_summary: a.site_commo, site_type: a.site_type_, site_sub_type: a.site_sub_t,
    site_stage: a.site_stage, target_commodity: a.target_com, commodity_group: a.commodity,
    project_code: a.proj_code, project_name: a.proj_title, confidentiality: a.confidenti,
    point_confidence: a.point_conf, latitude: a.latitude, longitude: a.longitude,
    web_link: a.web_link, extract_date: esriMsToDate(a.extract_da),
  }));
}

// Mirrors fetchTenementsAtPoint() in public/index.html.
async function fetchTenementsAtPointServer(lat, lon) {
  if (isNaN(lat) || isNaN(lon)) return [];
  const url = `${TENEMENT_URL}?geometry=${lon}%2C${lat}&geometryType=esriGeometryPoint&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=tenid,fmt_tenid,type,tenstatus,holder1,legal_area,unit_of_me,enddate&returnGeometry=false&f=json`;
  try {
    const resp = await fetch(url);
    const data = await resp.json();
    return ((data && data.features) || []).map(f => f.attributes);
  } catch (e) {
    return [];
  }
}

// Mirrors collectClientCandidates() in public/index.html. The client-side
// exclusions list is a browser-local preference, so it has no equivalent here.
function collectCandidatesServer(rows, minerals, stages) {
  const tokens = minerals.map(m => m.trim().toUpperCase()).filter(Boolean);
  const stageSet = new Set(stages);
  const stageOrder = ['Operating', 'Under Development', 'Care and Maintenance', 'Proposed', 'Undeveloped'];
  let candidates = rows.filter(r => {
    if (!stageSet.has(r.site_stage)) return false;
    const hay = ((r.target_commodity || '') + ' ' + (r.commodity_group || '')).toUpperCase();
    if (tokens.length && !tokens.some(t => hay.includes(t))) return false;
    if (isMajorProject(r)) return false;
    const t = (r.site_type || '').toUpperCase();
    if (t && t !== 'MINE' && t !== 'DEPOSIT') return false;
    return true;
  });
  const seen = new Set();
  candidates = candidates.filter(r => {
    const key = (r.project_name || r.site_name || '').toUpperCase();
    if (!key) return true;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  candidates.sort((a, b) => stageOrder.indexOf(a.site_stage) - stageOrder.indexOf(b.site_stage));
  return candidates;
}

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
  const clientCriteria = req.body.clientCriteria || null;
  if (!opp || (!opp.site_name && !opp.project_name)) {
    return res.status(400).json({ error: "No opportunity data provided" });
  }

  try {
    const text = await researchOpportunity(opp, clientCriteria);
    res.json({ research: text });
  } catch (err) {
    res.status(500).json({ error: err.message || "Research failed" });
  }
});

app.post("/find-candidates", async (req, res) => {
  const minerals = Array.isArray(req.body.minerals) ? req.body.minerals : [];
  const stages = Array.isArray(req.body.stages) ? req.body.stages : [];
  if (!minerals.length || !stages.length) {
    return res.status(400).json({ error: "minerals and stages are required" });
  }
  try {
    const rows = await fetchLiveMinedexServer();
    console.log(`[find-candidates] Pulled ${rows.length} live Minedex records from DMIRS`);
    const candidates = collectCandidatesServer(rows, minerals, stages);
    console.log(`[find-candidates] Filtered to ${candidates.length} candidates matching minerals=[${minerals.join(',')}] stages=[${stages.join(',')}]`);
    const capped = candidates.slice(0, 30);
    const enriched = [];
    for (const cand of capped) {
      const lat = parseFloat(cand.latitude), lon = parseFloat(cand.longitude);
      const tenements = isNaN(lat) || isNaN(lon) ? [] : await fetchTenementsAtPointServer(lat, lon);
      const primaryTen = tenements[0] || null;
      console.log(`[find-candidates] ${cand.site_name || cand.project_name}: tenement=${primaryTen ? (primaryTen.fmt_tenid||primaryTen.tenid) : 'none found'}, holder=${primaryTen ? primaryTen.holder1 : 'n/a'}`);
      const holderIsMajor = primaryTen ? isMajorHolder(primaryTen.holder1) : false;
      enriched.push({
        candidate: cand,
        tenement: primaryTen,
        holderIsMajor,
      });
    }
    res.json({
      totalMatched: candidates.length,
      researched: enriched.length,
      candidates: enriched,
    });
  } catch (err) {
    res.status(500).json({ error: err.message || "Candidate search failed" });
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

async function researchOpportunity(opp, clientCriteria) {
  const systemPrompt = `You are a research analyst for Blackwood & Terrace, an Australian deal origination firm connecting offshore private capital with off-market Australian mining opportunities. You research WA mining projects and tenement holders for a private client brief.

For the project/tenement given, use web search to find and report on:
${clientCriteria ? `0. FIT VERDICT — the very first line of your response, exactly in this format: "FIT: Yes" or "FIT: Maybe" or "FIT: No" followed by a colon and one sentence why, judged against the client criteria given below. Be honest and specific — if it's ASX-listed, JV'd to a major, foreign state-owned, or clearly outside the value range, that's a No.\n` : ''}1. OWNERSHIP VERIFICATION — this is the most important section and requires genuine multi-step investigation, not a single search. Do ALL of the following before concluding:
   a. Search the holder company name directly, plus the words "ASX" and "subsidiary".
   b. Search the holder company name plus "owned by" and separately plus "parent company".
   c. If the holder name doesn't obviously match a public company, search the associated project/site name plus "ASX announcement" — many ASX-listed companies hold WA tenements through obscure shell or joint-venture subsidiary names that don't resemble the parent (e.g. a company called "Round Oak" turning out to be a subsidiary of ASX-listed "Aeris Resources", or "Venturex" turning out to be "Develop Global"). A generic-sounding Pty Ltd name is NOT sufficient evidence of being private — you must actively rule out a listed parent, not just fail to find one on the first search.
   d. If initial searches find nothing, try the ASIC company name search pattern ("[name] ASIC") to check registered company type and any linked entities.
   e. Check for recent ASX announcements, investor relations pages, or annual reports mentioning this specific project or tenement number, since a listed company will usually have investor-facing material about assets it holds even if the operating subsidiary name looks generic.

Only after these steps, report OFF-MARKET STATUS as one of: "Confirmed private" (active search for a listed/state parent found none), "Confirmed listed/controlled" (found a specific listed or state parent, name it), or "Unable to confirm" (searches were inconclusive — this must be flagged for manual follow-up, never defaulted to "appears private"). Do not conclude "private" merely from absence of a quick match — only from having actively searched for and failed to find a parent across the steps above.
2. Scale — resource estimate, mine life, production figures, DFS/PFS economics (NPV, IRR) if publicly available.
3. Recent developments — funding rounds, JV agreements, corporate activity, news in the last 12-24 months.
4. Contact — any named individuals, investor relations contacts, or company addresses/emails found in public records.
SEARCHES PERFORMED: List the specific search queries you actually ran, so the depth of the check is auditable.
5. Assessment — one paragraph: is this genuinely off-market and approachable, or already spoken for (listed, JV'd, state-owned)? State this plainly.

Write in plain prose, organized under short headers, 300-500 words total. This is for a private research file, not a public document — be direct and factual, cite where information came from (company name, article, or "MINEDEX record").`;

  const userMessage = `Research this WA mining opportunity:
Site/Project: ${opp.site_name || opp.project_name || ''}
Project name: ${opp.project_name || ''}
Commodity: ${opp.commodity_group || ''} (target: ${opp.target_commodity || ''})
Stage: ${opp.site_stage || ''}
Tenement: ${opp.tenement_id || ''}
Registered holder: ${opp.holder1 || ''}
${opp.web_link ? 'MINEDEX record: ' + opp.web_link : ''}
${clientCriteria ? '\nCLIENT CRITERIA to judge fit against:\n' + clientCriteria : ''}`;

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
