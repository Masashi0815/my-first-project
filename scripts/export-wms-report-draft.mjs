/**
 * WMS 3週目.xlsx をマスタに、レポート HTML 向け JSON を標準出力する。
 * 累計差分は WMS.xlsx（基準）との行番号対応で算出。
 */
import fs from "fs";
import path from "path";
import { execFileSync } from "child_process";

function extractXlsx(absXlsx, destDir) {
  fs.rmSync(destDir, { recursive: true, force: true });
  fs.mkdirSync(destDir, { recursive: true });
  const zipPath = path.join(destDir, "_pack.zip");
  fs.copyFileSync(absXlsx, zipPath);
  const ps = [
    "-NoProfile",
    "-Command",
    `Expand-Archive -LiteralPath '${zipPath.replace(/'/g, "''")}' -DestinationPath '${destDir.replace(/'/g, "''")}' -Force`,
  ];
  execFileSync("powershell", ps, { stdio: "ignore" });
  fs.unlinkSync(zipPath);
}

function parseSST(xml) {
  const out = [];
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml))) {
    const firstT = /<t[^>]*>([^<]*)<\/t>/.exec(m[1]);
    out.push(firstT ? firstT[1] : "");
  }
  return out;
}

function parseRowCells(inner) {
  const c = {};
  const re =
    /<c r="([A-Z]+)(\d+)"([^>]*)>\s*<v>([^<]*)<\/v>/g;
  let m;
  while ((m = re.exec(inner))) {
    const col = m[1];
    const attrs = m[3];
    const vv = m[4];
    const isStr = /t="s"/.test(attrs);
    c[col] = { raw: vv, isStr };
  }
  return c;
}

function resolveCell(cell, sst) {
  if (!cell) return undefined;
  if (cell.isStr) return sst[Number(cell.raw)];
  return Number(cell.raw);
}

function readSheetRows(xlsxAbs, minRow = 3) {
  const dest = path.join(
    process.env.TEMP || "/tmp",
    "wms_export_" + Math.random().toString(36).slice(2),
  );
  extractXlsx(xlsxAbs, dest);
  const base = path.join(dest, "xl");
  const sst = parseSST(
    fs.readFileSync(path.join(base, "sharedStrings.xml"), "utf8"),
  );
  const sheet = fs.readFileSync(
    path.join(base, "worksheets", "sheet1.xml"),
    "utf8",
  );
  const rowRe = /<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  const rows = [];
  let rm;
  while ((rm = rowRe.exec(sheet))) {
    const ri = Number(rm[1]);
    if (ri < minRow) continue;
    const cells = parseRowCells(rm[2]);
    rows.push({
      row: ri,
      rep: resolveCell(cells.B, sst),
      doctor: resolveCell(cells.F, sst),
      facility: resolveCell(cells.E, sst),
      cases: resolveCell(cells.I, sst),
      callGuide: resolveCell(cells.M, sst),
      plan: resolveCell(cells.N, sst),
      promo: resolveCell(cells.O, sst),
    });
  }
  fs.rmSync(dest, { recursive: true, force: true });
  rows.sort((a, b) => a.row - b.row);
  return rows;
}

function promoStatus(plan, promo) {
  const p = Number(plan);
  const a = Number(promo);
  if (!Number.isFinite(p) || !Number.isFinite(a))
    return { kind: "unknown", gap: null };
  const gap = p - a;
  if (gap <= 0) return { kind: "done", gap };
  if (gap === 1) return { kind: "warn", gap };
  return { kind: "bad", gap };
}

function findAiDir() {
  const candidates = [
    path.join(process.env.USERPROFILE || "", "OneDrive", "Desktop", "AI"),
    path.join(process.env.USERPROFILE || "", "Desktop", "AI"),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

const ai = findAiDir();
if (!ai) {
  console.error(JSON.stringify({ error: "AI folder not found" }));
  process.exit(1);
}

const xlsxFiles = fs
  .readdirSync(ai)
  .filter((f) => f.endsWith(".xlsx") && /^WMS/i.test(f))
  .map((f) => ({ f, full: path.join(ai, f) }));

const baseFile = xlsxFiles.find((x) => /^WMS\.xlsx$/i.test(x.f));
const week3File = xlsxFiles.find(
  (x) => /3|三|Ⅲ/i.test(x.f) && !/^WMS\.xlsx$/i.test(x.f),
);

if (!baseFile || !week3File) {
  console.error(
    JSON.stringify({
      error: "need WMS.xlsx and WMS week3",
      found: xlsxFiles.map((x) => x.f),
    }),
  );
  process.exit(1);
}

const newer = readSheetRows(week3File.full);
const olderRows = readSheetRows(baseFile.full);
const prevByRow = new Map(olderRows.map((r) => [r.row, r]));

const byRep = new Map();
for (const r of newer) {
  if (!r.rep) continue;
  if (!byRep.has(r.rep)) byRep.set(r.rep, []);
  byRep.get(r.rep).push(r);
}

const guideBreakdown = {};
for (const [rep, list] of byRep) {
  const counts = { 4: 0, 2: 0, 1: 0 };
  for (const r of list) {
    const g = Number(r.callGuide);
    if (g === 4) counts[4]++;
    else if (g === 2) counts[2]++;
    else if (g === 1) counts[1]++;
  }
  guideBreakdown[rep] = counts;
}

let achieved = 0;
const nList = newer.length;
const doctorRows = newer.map((n) => {
  const o = prevByRow.get(n.row);
  const prevCases =
    o && typeof o.cases === "number" && Number.isFinite(o.cases)
      ? o.cases
      : null;
  const currCases =
    typeof n.cases === "number" && Number.isFinite(n.cases) ? n.cases : null;
  const delta =
    prevCases !== null && currCases !== null ? currCases - prevCases : null;
  const st = promoStatus(n.plan, n.promo);
  if (st.kind === "done") achieved++;
  return {
    ...n,
    casesPrev: prevCases,
    delta,
    status: st,
  };
});

const rate = nList ? Math.round((achieved / nList) * 1000) / 10 : 0;

const withIncrease = doctorRows.filter((d) => (d.delta ?? 0) !== 0);
const totalDelta = withIncrease.reduce((s, d) => s + (d.delta ?? 0), 0);

const deltaByRep = new Map();
for (const d of withIncrease) {
  if (!d.rep) continue;
  deltaByRep.set(d.rep, (deltaByRep.get(d.rep) || 0) + (d.delta ?? 0));
}

const out = {
  files: { base: baseFile.f, week3: week3File.f },
  rowCount: nList,
  achievedCount: achieved,
  achievementRatePct: rate,
  guideBreakdown,
  doctorRows,
  cumulativeSummary: {
    doctorsWithIncrease: withIncrease.length,
    totalDelta,
    deltaByRep: Object.fromEntries(deltaByRep),
  },
};

console.log(JSON.stringify(out, null, 2));
