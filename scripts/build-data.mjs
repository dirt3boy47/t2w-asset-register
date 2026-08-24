#!/usr/bin/env node
/* ==========================================================================
   Excel -> data.json

   Reads the master workbook and writes the dataset the register page fetches.
   Run by Cloudflare Pages at deploy time, so data.json is never committed and
   can never drift from the workbook it came from.

   This is not "save as JSON". It reproduces the same derivations the register
   depends on:
     - Record Key and every populated column, empty columns dropped
     - the alignment sheet for each asset, derived from chainage where the
       source register did not state one (700 m per sheet)
     - the Data Quality exceptions, linked to the specific records they affect
     - the alignment sheet index, including sheets that hold no records

   Usage: node scripts/build-data.mjs [path/to/workbook.xlsx] [path/to/out.json]
   ========================================================================== */

import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import * as XLSX from 'xlsx';

const INTERVAL = 700;                       // metres of chainage per alignment sheet
const DATA_SHEET = 'Master Register';
const DQ_SHEET   = 'Data Quality';

const args = process.argv.slice(2);
const XL  = args[0] || findWorkbook();
const OUT = args[1] || 'public/data.json';

function findWorkbook(){
  const candidates = fs.readdirSync('.').filter(f => /\.xlsx$/i.test(f) && !/^~\$/.test(f));
  const preferred = candidates.find(f => /Master Asset Register/i.test(f));
  const pick = preferred || candidates[0];
  if(!pick) fail('No .xlsx file found in the repository root. Commit the workbook, or pass its path as an argument.');
  return pick;
}
function fail(msg){
  console.error('\n  BUILD FAILED: ' + msg + '\n');
  process.exit(1);
}
function snake(h){
  if(h === 'ID') return 'utility_id';
  return String(h).toLowerCase()
    .replace(/&/g,'and').replace(/\(m\)/g,'m').replace(/\(mm\)/g,'mm')
    .replace(/\(surveyed\)/g,'surveyed').replace(/\(start\)/g,'start').replace(/\(end\)/g,'end')
    .replace(/\(source\)/g,'source').replace(/\(superseded\)/g,'superseded')
    .replace(/\?/g,'').replace(/[^a-z0-9]+/g,'_').replace(/^_|_$/g,'');
}
const DWG_RE = /^D-DWG-(PHW|PWG)-(\d+)(\.(\d+)|_C)?$/;
function parseDwg(v){
  if(v == null) return null;
  const m = DWG_RE.exec(String(v).trim().toUpperCase());
  if(!m) return null;
  return { sec:m[1], no:parseInt(m[2],10), rev:m[4] ? m[4] : (m[3] === '_C' ? 'C' : '1') };
}
const pad3 = n => String(n).padStart(3,'0');

/* ---------- read ---------- */
if(!fs.existsSync(XL)) fail(`Workbook not found: ${XL}`);
console.log(`  reading  ${XL}`);
const wb = XLSX.read(fs.readFileSync(XL), { type:'buffer' });
if(!wb.Sheets[DATA_SHEET])
  fail(`The workbook has no "${DATA_SHEET}" sheet. Found: ${wb.SheetNames.join(', ')}`);

const aoa = XLSX.utils.sheet_to_json(wb.Sheets[DATA_SHEET], { header:1, raw:true, defval:null });
if(aoa.length < 2) fail(`"${DATA_SHEET}" has no data rows.`);

const head = aoa[0].map(h => h == null ? '' : String(h));
const cols = head.map(snake);
const Ci = {};
cols.forEach((c,i) => { Ci[c] = i; });

for(const required of ['record_key','chainage_start_m','register','pipeline_section']){
  if(Ci[required] == null)
    fail(`"${DATA_SHEET}" is missing a required column (${required}). Has the header row changed?`);
}

const body = aoa.slice(1).filter(r => r && r.some(v => v != null && v !== ''));
console.log(`  records  ${body.length}`);

/* ---------- derive ---------- */
const EXTRA = ['drawing_base','drawing_end_base','drawing_section','drawing_no','drawing_rev',
               'drawing_is_derived','drawing_rule_offset','is_ranged'];
const allCols = cols.concat(EXTRA);

let derivedCount = 0, noSheetCount = 0;
const rows = body.map(r => {
  const row = cols.map((c,i) => {
    const v = r[i];
    return (v === '' || v === undefined) ? null : v;
  });
  const sec = row[Ci.pipeline_section];
  const ch  = row[Ci.chainage_start_m];
  const p   = parseDwg(row[Ci.drawing_start]);

  let base=null, dsec=null, dno=null, drev=null, derived=null, offset=null;
  if(p){
    dsec = p.sec; dno = p.no; drev = p.rev;
    base = `D-DWG-${p.sec}-${pad3(p.no)}`;
    derived = 0;
    if(typeof ch === 'number') offset = p.no - (Math.floor(ch / INTERVAL) + 1);
  } else if((sec === 'PHW' || sec === 'PWG') && typeof ch === 'number'){
    dno = Math.floor(ch / INTERVAL) + 1;
    dsec = sec;
    base = `D-DWG-${sec}-${pad3(dno)}`;
    derived = 1;
    derivedCount++;
  } else {
    noSheetCount++;
  }
  const pe = parseDwg(row[Ci.drawing_end]);
  row.push(base,
           pe ? `D-DWG-${pe.sec}-${pad3(pe.no)}` : null,
           dsec, dno, drev, derived, offset,
           row[Ci.chainage_end_m] != null ? 1 : 0);
  return row;
});
console.log(`  sheets   ${derivedCount} derived from chainage, ${noSheetCount} cannot be placed`);

/* ---------- data quality, and which records each exception touches ---------- */
const dq = [], flags = {};
const ds = wb.Sheets[DQ_SHEET];
if(ds){
  const d2 = XLSX.utils.sheet_to_json(ds, { header:1, raw:true, defval:null });
  /* Locate the header rather than assuming a position. sheet_to_json trims leading
     blank rows and columns, so this table can start at index 0 even though it sits
     in column B of row 5. Assuming an offset shifts every field by one. */
  let hRow = -1, hCol = 0;
  for(let i = 0; i < d2.length && hRow < 0; i++){
    const rw = d2[i] || [];
    for(let c = 0; c < rw.length; c++){
      if(String(rw[c] || '').trim() === 'Register' &&
         String(rw[c+1] || '').startsWith('Source row')){ hRow = i; hCol = c; break; }
    }
  }
  if(hRow < 0){
    console.log(`  note     "${DQ_SHEET}" found but no header row recognised; exceptions skipped`);
  } else {
    for(let i = hRow + 1; i < d2.length; i++){
      const q = d2[i]; if(!q) continue;
      const v = [q[hCol], q[hCol+1], q[hCol+2], q[hCol+3], q[hCol+4]];
      if(!v.some(x => x != null && x !== '')) continue;
      dq.push(v.map(x => x == null ? null : String(x)));
    }
    const iAid = allCols.indexOf('asset_id');
    const iKey = allCols.indexOf('record_key');
    const iReg = allCols.indexOf('register');
    const iSr  = allCols.indexOf('source_row');
    const byId = {}, byReg = {};
    rows.forEach((r,k) => {
      [iAid, iKey].forEach(ix => {
        if(ix >= 0 && r[ix]) byId[String(r[ix]).trim()] = k + 1;
      });
      const g = r[iReg];
      if(!byReg[g]) byReg[g] = {};
      if(iSr >= 0) byReg[g][r[iSr]] = k + 1;
    });
    dq.forEach((q, qi) => {
      const hay = `${q[1] || ''} ${q[4] || ''}`;
      for(const id of Object.keys(byId)){
        if(id && hay.includes(id)) (flags[byId[id]] = flags[byId[id]] || []).push(qi + 1);
      }
      const base = String(q[0] || '').split(' / ')[0];
      if(byReg[base] && q[1]){
        (String(q[1]).match(/\d+/g) || []).forEach(t => {
          const k = byReg[base][+t];
          if(k) (flags[k] = flags[k] || []).push(qi + 1);
        });
      }
    });
    Object.keys(flags).forEach(k => {
      flags[k] = flags[k].filter((v,i,a) => a.indexOf(v) === i);
    });
    console.log(`  quality  ${dq.length} exceptions, ${Object.keys(flags).length} records flagged`);
  }
}

/* ---------- alignment sheet index, including sheets holding no records ---------- */
const iBase = allCols.indexOf('drawing_base');
const iCh   = allCols.indexOf('chainage_start_m');
const iRev  = allCols.indexOf('drawing_rev');
const iDer  = allCols.indexOf('drawing_is_derived');
const iSec  = allCols.indexOf('pipeline_section');

const agg = {};
rows.forEach(r => {
  const b = r[iBase]; if(!b) return;
  const e = agg[b] || (agg[b] = { n:0, min:null, max:null, revs:{}, stated:0, derived:0 });
  e.n++;
  const c = r[iCh];
  if(typeof c === 'number'){
    e.min = e.min == null ? c : Math.min(e.min, c);
    e.max = e.max == null ? c : Math.max(e.max, c);
  }
  if(r[iRev]) e.revs[r[iRev]] = 1;
  if(r[iDer] === 1) e.derived++; else e.stated++;
});
const top = {};
Object.keys(agg).forEach(b => {
  const m = /^D-DWG-(PHW|PWG)-(\d+)$/.exec(b);
  if(m) top[m[1]] = Math.max(top[m[1]] || 0, +m[2]);
});
rows.forEach(r => {
  const s = r[iSec], c = r[iCh];
  if((s === 'PHW' || s === 'PWG') && typeof c === 'number')
    top[s] = Math.max(top[s] || 0, Math.floor(c / INTERVAL) + 1);
});
Object.keys(top).forEach(s => {
  for(let n = 1; n <= top[s]; n++){
    const b = `D-DWG-${s}-${pad3(n)}`;
    if(!agg[b]) agg[b] = { n:0, min:null, max:null, revs:{}, stated:0, derived:0 };
  }
});
const drawings = Object.keys(agg).sort().map(b => {
  const m = /^D-DWG-(PHW|PWG)-(\d+)$/.exec(b);
  const no = m ? parseInt(m[2],10) : 0;
  const e = agg[b];
  return [b, m ? m[1] : '', no, (no-1)*INTERVAL, no*INTERVAL, e.min, e.max, e.n,
          Object.keys(e.revs).sort().join(',') || null, e.stated, e.derived];
});
const emptySheets = drawings.filter(d => d[7] === 0).map(d => d[0]);
console.log(`  drawings ${drawings.length} sheets` +
            (emptySheets.length ? `, ${emptySheets.length} holding no records (${emptySheets.join(', ')})` : ''));

/* ---------- drop columns that are empty everywhere, to keep the payload small ----------
   Progress columns are exempt. They start empty by design, and you have to be able to
   filter on "not installed" - which is impossible if the column has been dropped. */
const ALWAYS_KEEP = new Set(["installed", "install_date", "installed_by", "test_document_no", "test_status", "test_date", "complete", "completion_date", "progress_notes"]);
const keep = allCols.filter((c,i) =>
  ALWAYS_KEEP.has(c) || rows.some(r => r[i] != null && r[i] !== ''));
const keepIx = keep.map(c => allCols.indexOf(c));
const trimmed = rows.map(r => keepIx.map(i => r[i]));
console.log(`  columns  ${keep.length} kept of ${allCols.length} (${allCols.length - keep.length} empty in every record)`);

/* ---------- write ---------- */
const out = {
  meta: {
    records: trimmed.length,
    interval: INTERVAL,
    workbook: path.basename(XL),
    generated: new Date().toLocaleDateString('en-AU', { day:'2-digit', month:'short', year:'numeric' })
  },
  cols: keep,
  rows: trimmed,
  flags,
  dq,
  drawings
};
fs.mkdirSync(path.dirname(OUT), { recursive:true });
fs.writeFileSync(OUT, JSON.stringify(out));
const kb = (fs.statSync(OUT).size / 1024).toFixed(0);
console.log(`\n  wrote    ${OUT}  (${kb} KB, ${trimmed.length} records)\n`);
