/** 開発用：%TEMP%/wms_xlsx_extract に WMS.xlsx を ZIP 展開したあと実行。JSON で集計を確認できます。 */
import fs from "fs";
import path from "path";

const xlsxDir = path.join(
  process.env.TEMP || "/tmp",
  "wms_xlsx_extract",
  "xl",
);

function parseSST(xml) {
  const out = [];
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml))) {
    const inner = m[1];
    const firstT = /<t[^>]*>([^<]*)<\/t>/.exec(inner);
    out.push(firstT ? firstT[1] : "");
  }
  return out;
}

function parseRowCells(inner) {
  const c = {};
  const re =
    /<c r="([A-Z]+)(\d+)"([^>]*)>\s*<v>([^<]*)<\/v>|<c r="([A-Z]+)(\d+)"([^>]*)\/>/g;
  let m;
  while ((m = re.exec(inner))) {
    if (m[1]) {
      const col = m[1];
      const attrs = m[3];
      const vv = m[4];
      const isStr = /t="s"/.test(attrs);
      const num = Number(vv);
      c[col] = { raw: vv, isStr, val: isStr ? null : num };
    }
  }
  return c;
}

function resolveCell(cell, sst) {
  if (!cell) return undefined;
  if (cell.isStr) return sst[Number(cell.raw)];
  return cell.val;
}

const sstXml = fs.readFileSync(path.join(xlsxDir, "sharedStrings.xml"), "utf8");
const sheetXml = fs.readFileSync(
  path.join(xlsxDir, "worksheets", "sheet1.xml"),
  "utf8",
);
const sst = parseSST(sstXml);

const rowRe = /<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
const rows = [];
let rm;
while ((rm = rowRe.exec(sheetXml))) {
  const ri = Number(rm[1]);
  if (ri < 3 || ri > 22) continue;
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

/** 達成判定: 宣伝実績 >= 計画（計画は面談計画列） */
function promoStatus(plan, promo) {
  const p = Number(plan);
  const a = Number(promo);
  if (!Number.isFinite(p) || !Number.isFinite(a)) return { kind: "unknown", gap: null };
  const gap = p - a;
  if (gap <= 0) return { kind: "done", gap };
  if (gap === 1) return { kind: "warn", gap };
  return { kind: "bad", gap };
}

/** コールガイダンス度数別・担当者ごとの医師数 */
const byRep = new Map();
for (const r of rows) {
  if (!r.rep) continue;
  if (!byRep.has(r.rep)) byRep.set(r.rep, []);
  byRep.get(r.rep).push(r);
}

const guideBuckets = [4, 2, 1];
const guideBreakdown = {};
for (const [rep, list] of byRep) {
  const counts = { 4: 0, 2: 0, 1: 0, other: 0 };
  for (const r of list) {
    const g = Number(r.callGuide);
    if (g === 4) counts[4]++;
    else if (g === 2) counts[2]++;
    else if (g === 1) counts[1]++;
    else counts.other++;
  }
  guideBreakdown[rep] = counts;
}

let achieved = 0;
const doctorRows = rows.map((r) => {
  const st = promoStatus(r.plan, r.promo);
  if (st.kind === "done") achieved++;
  return { ...r, status: st };
});
const rate = rows.length ? Math.round((achieved / rows.length) * 1000) / 10 : 0;

const out = {
  rowCount: rows.length,
  achievedCount: achieved,
  achievementRatePct: rate,
  guideBreakdown,
  doctorRows,
};

console.log(JSON.stringify(out, null, 2));
