/**
 * 2つの WMS の Sheet1（同じ行構成）で列「累計症例数」（列 I）の差分を出す。
 * 比較：前回 = WMS.xlsx、今回 = 「WMS 3週目」相当のブック（名前に 3 が含まれる .xlsx）。
 */
import fs from "fs";
import path from "path";
import { execFileSync } from "child_process";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

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
  execFileSync("powershell", ps, { stdio: "inherit" });
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

function readDataRows(xlsxAbs) {
  const dest = path.join(
    process.env.TEMP || "/tmp",
    "wms_compare_" + Math.random().toString(36).slice(2),
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
    if (ri < 3) continue;
    const cells = parseRowCells(rm[2]);
    rows.push({
      row: ri,
      rep: resolveCell(cells.B, sst),
      doctor: resolveCell(cells.F, sst),
      facility: resolveCell(cells.E, sst),
      cases: resolveCell(cells.I, sst),
    });
  }
  fs.rmSync(dest, { recursive: true, force: true });
  return rows;
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
  console.error("AI フォルダが見つかりません。");
  process.exit(1);
}

const xlsxFiles = fs
  .readdirSync(ai)
  .filter((f) => f.endsWith(".xlsx") && /^WMS/i.test(f))
  .map((f) => ({ f, full: path.join(ai, f) }));

const baseFile = xlsxFiles.find((x) => /^WMS\.xlsx$/i.test(x.f));
const week3File = xlsxFiles.find((x) => /3|三|Ⅲ/i.test(x.f) && !/^WMS\.xlsx$/i.test(x.f));

if (!baseFile || !week3File) {
  console.error("WMS.xlsx と WMS 第3週相当の xlsx の両方が必要です。", xlsxFiles.map((x) => x.f));
  process.exit(1);
}

const older = readDataRows(baseFile.full);
const newer = readDataRows(week3File.full);

const byRow = new Map(older.map((r) => [r.row, r]));
const out = [];
for (const n of newer) {
  const o = byRow.get(n.row);
  const prev = o && typeof o.cases === "number" ? o.cases : null;
  const curr = typeof n.cases === "number" ? n.cases : null;
  const delta =
    prev !== null && curr !== null ? curr - prev : null;
  out.push({
    row: n.row,
    rep: n.rep,
    doctor: n.doctor,
    facility: n.facility,
    casesPrev: prev,
    casesNow: curr,
    delta,
  });
}

console.log(
  JSON.stringify(
    {
      files: {
        previous: baseFile.f,
        current: week3File.f,
      },
      rows: out,
    },
    null,
    2,
  ),
);
