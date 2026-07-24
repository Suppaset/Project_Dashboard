import React from "react";
import { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, PieChart, Pie, Cell, RadarChart, Radar, PolarGrid,
  PolarAngleAxis, PolarRadiusAxis, ScatterChart, Scatter, ComposedChart, ReferenceLine
} from "recharts";

/* ═══════════════════════════════════════════
   CONSTANTS & UTILS
   ═══════════════════════════════════════════ */
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const PAL = ["#2563eb","#0ea5e9","#8b5cf6","#f59e0b","#10b981","#ec4899","#6366f1","#14b8a6","#f97316","#a855f7","#ef4444","#84cc16"];
const C = { pri:"#0f172a", sec:"#2563eb", acc:"#0ea5e9", pos:"#10b981", neg:"#ef4444", warn:"#f59e0b", mut:"#94a3b8", bg:"#f1f5f9", bor:"#e2e8f0" };
const sum = a => a.filter(v => v != null && typeof v === "number").reduce((s, v) => s + v, 0);
const avg = a => { const f = a.filter(v => v != null && typeof v === "number"); return f.length ? sum(f) / f.length : 0; };
const fmt = n => n == null ? "—" : typeof n === "number" ? Math.round(n).toLocaleString() : String(n);
const fmtDec = (n, d=1) => n == null ? "—" : typeof n === "number" ? n.toFixed(d) : String(n);
const pct = n => n == null ? "—" : (n >= 0 ? "+" : "") + n.toFixed(1) + "%";
const readVal = v => (v != null && typeof v === "number") ? v : null;
const readAny = v => (v != null && v !== "" && v !== "-") ? v : null;

/* ═══════════════════════════════════════════
   EXCEL PARSER — strict per-sheet extraction
   ═══════════════════════════════════════════ */
function parseExcel(ab) {
  const wb = XLSX.read(ab, { type: "array", cellDates: true });
  const missing = ["Overview","DC","Compare"].filter(s => !wb.SheetNames.includes(s));
  if (missing.length) return { error: "Missing sheet(s): " + missing.join(", ") };

  /* ── OVERVIEW SHEET ── */
  const ov = XLSX.utils.sheet_to_json(wb.Sheets["Overview"], { header:1, defval:null });
  let ovHdr = -1;
  for (let i = 0; i < ov.length; i++) { if (ov[i] && String(ov[i][0]||"").includes("Yearly")) { ovHdr = i; break; } }
  if (ovHdr < 0) return { error: "Overview: cannot find 'Yearly / Month' header" };

  const overview = { yearly:{}, forecast:null, diff:null, pctAct:null, pctFore:null, tripsGrow:null, growPct:null };
  const ovYears = [];
  for (let i = ovHdr + 1; i < ov.length; i++) {
    const label = String(ov[i]?.[0]||"").trim();
    if (!label) continue;
    const vals = MONTHS.map((_, mi) => readVal(ov[i]?.[mi+1]));
    const total = readVal(ov[i]?.[13]);
    if (/^Y\d{4}$/i.test(label)) {
      const key = label.toUpperCase();
      overview.yearly[key] = { months: vals, total };
      ovYears.push(key);
    }
    else if (/^forecast$/i.test(label)) overview.forecast = { months: vals, total };
    else if (/^diff$/i.test(label)) overview.diff = { months: vals, total };
    else if (/^%\s*\(act\)/i.test(label)) overview.pctAct = { months: vals, total };
    else if (/^%\s*\(fore\)/i.test(label)) overview.pctFore = { months: vals, total };
    else if (/^trips\s*grows?\s*up/i.test(label)) overview.tripsGrow = { months: vals, total };
    else if (label === "%") overview.growPct = { months: vals, total };
  }
  overview.years = ovYears;
  const ovLatest = ovYears[ovYears.length - 1] || null;
  overview.latestYear = ovLatest;
  overview.actualMonths = ovLatest ? overview.yearly[ovLatest].months.filter(v => v != null).length : 0;

  /* ── DC SHEET ── */
  const dc = XLSX.utils.sheet_to_json(wb.Sheets["DC"], { header:1, defval:null });
  let dcHdr = -1;
  for (let i = 0; i < dc.length; i++) {
    const row = (dc[i]||[]).map(v => String(v||"").trim());
    if (row[0] === "Month") { dcHdr = i; break; }
  }
  if (dcHdr < 0) return { error: "DC: cannot find 'Month' header" };

  const dcHeader = dc[dcHdr];
  // Find DC names from "BY DC" section (columns 1..n until "Total")
  const dcSheet = { byDC: {}, byYear: {}, avgPerMonth: {}, growthPct: {}, pctByDC: {}, dcList: [], years: [] };
  // Parse DC column names
  const dcNames = [];
  for (let ci = 1; ci < dcHeader.length; ci++) {
    const h = String(dcHeader[ci]||"").trim();
    if (!h || h === "Total" || h === "None") break;
    dcNames.push({ name: h, col: ci });
  }
  dcSheet.dcList = dcNames.map(d => d.name);
  // Also find Total column for BY DC
  let dcTotalCol = -1;
  for (let ci = 1; ci < dcHeader.length; ci++) {
    if (String(dcHeader[ci]||"").trim() === "Total") { dcTotalCol = ci; break; }
  }

  // Parse BY DC monthly data
  const dcMonthlyRows = [];
  for (let i = dcHdr + 1; i < dc.length; i++) {
    const row = dc[i]; if (!row) continue;
    const dateVal = row[0];
    if (!(dateVal instanceof Date)) continue;
    const yr = dateVal.getUTCFullYear();
    const mo = dateVal.getUTCMonth();
    const entry = { year: yr, monthIdx: mo, month: MONTHS[mo] };
    dcNames.forEach(d => { entry[d.name] = readAny(row[d.col]); });
    if (dcTotalCol >= 0) entry.Total = readVal(row[dcTotalCol]);
    dcMonthlyRows.push(entry);
  }
  dcSheet.byDC = dcMonthlyRows;

  // Find "BY Year" sub-table (col index 9 = "Year")
  let byYearCol = -1;
  for (let ci = 0; ci < dcHeader.length; ci++) {
    if (String(dcHeader[ci]||"").trim() === "Year") { byYearCol = ci; break; }
  }
  if (byYearCol >= 0) {
    // DC names for this sub-table are at cols byYearCol+1, +2, ...
    const byYearDCs = [];
    for (let ci = byYearCol + 1; ci < dcHeader.length; ci++) {
      const h = String(dcHeader[ci]||"").trim();
      if (!h || h === "Total" || h === "None") break;
      byYearDCs.push({ name: h, col: ci });
    }
    let byYearTotalCol = -1;
    for (let ci = byYearCol + 1; ci < dcHeader.length; ci++) {
      if (String(dcHeader[ci]||"").trim() === "Total") { byYearTotalCol = ci; break; }
    }
    for (let i = dcHdr + 1; i < dc.length; i++) {
      const row = dc[i]; if (!row) continue;
      const yrLabel = String(row[byYearCol]||"").trim();
      if (!/^Y\d{4}$/i.test(yrLabel)) continue;
      const key = yrLabel.toUpperCase();
      if (!dcSheet.years.includes(key)) dcSheet.years.push(key);
      const entry = { year: key };
      byYearDCs.forEach(d => { entry[d.name] = readVal(row[d.col]); });
      if (byYearTotalCol >= 0) entry.Total = readVal(row[byYearTotalCol]);
      dcSheet.byYear[key] = entry;
    }
  }

  // Find "Average per Month" sub-table
  let avgCol = -1;
  for (let ci = 0; ci < dcHeader.length; ci++) {
    if (String(dcHeader[ci]||"").trim() === "Average" && ci > 10) { avgCol = ci; break; }
  }
  if (avgCol >= 0) {
    const avgDCs = [];
    for (let ci = avgCol + 1; ci < dcHeader.length; ci++) {
      const h = String(dcHeader[ci]||"").trim();
      if (!h || h === "Total" || h === "None") break;
      avgDCs.push({ name: h, col: ci });
    }
    let avgTotalCol = -1;
    for (let ci = avgCol + 1; ci < dcHeader.length; ci++) {
      if (String(dcHeader[ci]||"").trim() === "Total") { avgTotalCol = ci; break; }
    }
    for (let i = dcHdr + 1; i < dc.length; i++) {
      const row = dc[i]; if (!row) continue;
      const yrLabel = String(row[avgCol]||"").trim();
      if (!/^Y\d{4}$/i.test(yrLabel)) continue;
      const entry = { year: yrLabel.toUpperCase() };
      avgDCs.forEach(d => { entry[d.name] = readVal(row[d.col]); });
      if (avgTotalCol >= 0) entry.Total = readVal(row[avgTotalCol]);
      dcSheet.avgPerMonth[yrLabel.toUpperCase()] = entry;
    }
  }

  // Find "Growth percentage" sub-table
  let growCol = -1;
  for (let ci = 0; ci < dcHeader.length; ci++) {
    if (String(dcHeader[ci]||"").trim() === "Growth") { growCol = ci; break; }
  }
  if (growCol >= 0) {
    const growDCs = [];
    for (let ci = growCol + 1; ci < dcHeader.length; ci++) {
      const h = String(dcHeader[ci]||"").trim();
      if (!h || h === "Total" || h === "None") break;
      growDCs.push({ name: h, col: ci });
    }
    for (let i = dcHdr + 1; i < dc.length; i++) {
      const row = dc[i]; if (!row) continue;
      const yrLabel = String(row[growCol]||"").trim();
      if (!/^Y\d{4}$/i.test(yrLabel)) continue;
      const entry = { year: yrLabel.toUpperCase() };
      growDCs.forEach(d => { const v = row[d.col]; entry[d.name] = (v != null && typeof v === "number") ? v : (typeof v === "string" ? v : null); });
      dcSheet.growthPct[yrLabel.toUpperCase()] = entry;
    }
  }

  // Find "Percentage by DC" sub-table (last "Average" column)
  let pctCol = -1;
  for (let ci = dcHeader.length - 1; ci >= 0; ci--) {
    if (String(dcHeader[ci]||"").trim() === "Average") { pctCol = ci; break; }
  }
  if (pctCol >= 0 && pctCol !== avgCol) {
    const pctDCs = [];
    for (let ci = pctCol + 1; ci < dcHeader.length; ci++) {
      const h = String(dcHeader[ci]||"").trim();
      if (!h || h === "Total" || h === "None") break;
      pctDCs.push({ name: h, col: ci });
    }
    for (let i = dcHdr + 1; i < dc.length; i++) {
      const row = dc[i]; if (!row) continue;
      const yrLabel = String(row[pctCol]||"").trim();
      if (!/^Y\d{4}$/i.test(yrLabel)) continue;
      const entry = { year: yrLabel.toUpperCase() };
      pctDCs.forEach(d => { entry[d.name] = readVal(row[d.col]); });
      dcSheet.pctByDC[yrLabel.toUpperCase()] = entry;
    }
  }

  /* ── COMPARE SHEET ── */
  const cp = XLSX.utils.sheet_to_json(wb.Sheets["Compare"], { header:1, defval:null });
  let cpHdr = -1;
  for (let i = 0; i < cp.length; i++) {
    const row = (cp[i]||[]).map(v => String(v||"").trim());
    if (row.includes("DC") && row.includes("Year")) { cpHdr = i; break; }
  }
  if (cpHdr < 0) return { error: "Compare: cannot find 'DC'/'Year' header" };

  const cpHeader = cp[cpHdr];
  const cpDCCol = cpHeader.findIndex(v => String(v||"").trim() === "DC");
  const cpYrCol = cpHeader.findIndex(v => String(v||"").trim() === "Year");
  const cpJanCol = cpHeader.findIndex(v => String(v||"").trim() === "Jan");
  const cpTotalCol = cpHeader.findIndex(v => String(v||"").trim() === "Total");

  const compare = { dcList: [], data: {} };
  let curDC = null;
  for (let i = cpHdr + 1; i < cp.length; i++) {
    const row = cp[i]; if (!row) continue;
    const dcVal = row[cpDCCol] ? String(row[cpDCCol]).trim() : null;
    const yrVal = row[cpYrCol] ? String(row[cpYrCol]).trim() : null;
    if (dcVal && dcVal !== "") {
      if (dcVal === "Total") { curDC = "_Total"; } else { curDC = dcVal; }
    }
    if (!curDC || curDC === "_Total" || !yrVal) continue;
    if (!compare.dcList.includes(curDC)) { compare.dcList.push(curDC); compare.data[curDC] = {}; }
    const months = MONTHS.map((_, mi) => readVal(row[cpJanCol + mi]));
    const total = cpTotalCol >= 0 ? readVal(row[cpTotalCol]) : null;
    if (/^Y\d{4}$/i.test(yrVal)) {
      compare.data[curDC][yrVal.toUpperCase()] = { months, total };
    } else if (/^diff/i.test(yrVal)) {
      compare.data[curDC].diff = { months, total };
    }
  }
  // Also read Total rows (not as DC but for reference)
  let totalSection = false;
  for (let i = cpHdr + 1; i < cp.length; i++) {
    const row = cp[i]; if (!row) continue;
    const dcVal = row[cpDCCol] ? String(row[cpDCCol]).trim() : null;
    if (dcVal === "Total") totalSection = true;
    else if (dcVal && dcVal !== "") totalSection = false;
    if (!totalSection) continue;
    const yrVal = row[cpYrCol] ? String(row[cpYrCol]).trim() : null;
    if (!yrVal) continue;
    if (!compare.data._Total) compare.data._Total = {};
    if (/^Y\d{4}$/i.test(yrVal)) {
      compare.data._Total[yrVal.toUpperCase()] = { months: MONTHS.map((_, mi) => readVal(row[cpJanCol + mi])), total: cpTotalCol >= 0 ? readVal(row[cpTotalCol]) : null };
    } else if (/^diff/i.test(yrVal)) {
      compare.data._Total.diff = { months: MONTHS.map((_, mi) => readVal(row[cpJanCol + mi])), total: cpTotalCol >= 0 ? readVal(row[cpTotalCol]) : null };
    }
  }
  // Determine compare years
  const cpYears = [];
  if (compare.dcList.length) {
    Object.keys(compare.data[compare.dcList[0]]).forEach(k => { if (/^Y\d{4}$/.test(k) && !cpYears.includes(k)) cpYears.push(k); });
  }
  cpYears.sort();
  compare.years = cpYears;

  return { overview, dcSheet, compare, error: null };
}

/* ═══════════════════════════════════════════
   UI COMPONENTS
   ═══════════════════════════════════════════ */
const KPI = ({ label, value, sub, color }) => (
  <div style={{ background:"#fff", borderRadius:12, padding:"15px 17px", boxShadow:"0 1px 4px rgba(0,0,0,.05)", borderLeft:`4px solid ${color||C.sec}`, flex:"1 1 0", minWidth:130 }}>
    <div style={{ fontSize:10, color:C.mut, fontWeight:700, letterSpacing:.6, textTransform:"uppercase", marginBottom:3 }}>{label}</div>
    <div style={{ fontSize:22, fontWeight:800, color:C.pri, lineHeight:1.15 }}>{value}</div>
    {sub && <div style={{ fontSize:11, color: typeof sub==="string"&&sub.startsWith("-")?C.neg:typeof sub==="string"&&sub.startsWith("+")?C.pos:C.mut, fontWeight:600, marginTop:3 }}>{sub}</div>}
  </div>
);
const Card = ({ title, children, span, minH, style: sx }) => (
  <div style={{ background:"#fff", borderRadius:12, padding:"14px 16px", boxShadow:"0 1px 4px rgba(0,0,0,.05)", gridColumn:span?`span ${span}`:undefined, minHeight:minH||240, display:"flex", flexDirection:"column", ...sx }}>
    {title && <div style={{ fontSize:12, fontWeight:700, color:C.pri, marginBottom:10 }}>{title}</div>}
    <div style={{ flex:1, minHeight:0 }}>{children}</div>
  </div>
);
const Insight = ({ items }) => (
  <div style={{ background:`linear-gradient(135deg,${C.pri},${C.sec})`, borderRadius:12, padding:"15px 17px", color:"#fff" }}>
    <div style={{ fontSize:12, fontWeight:700, marginBottom:7 }}>💡 Insights</div>
    {items.map((t,i)=><div key={i} style={{ fontSize:11, lineHeight:1.65, opacity:.93, marginBottom:2 }}>• {t}</div>)}
  </div>
);
const TT = ({ active, payload, label }) => {
  if (!active||!payload?.length) return null;
  return <div style={{ background:"#fff", border:`1px solid ${C.bor}`, borderRadius:8, padding:"7px 11px", fontSize:11, boxShadow:"0 4px 12px rgba(0,0,0,.1)" }}>
    <div style={{ fontWeight:700, marginBottom:3 }}>{label}</div>
    {payload.map((p,i)=><div key={i} style={{ color:p.color||C.pri, marginBottom:1 }}>{p.name}: {typeof p.value==="number"?fmt(p.value):p.value}</div>)}
  </div>;
};
const Btn = ({ active, onClick, children }) => (
  <button onClick={onClick} style={{ padding:"5px 15px", borderRadius:7, border:active?`2px solid ${C.sec}`:`1px solid ${C.bor}`, background:active?C.sec:"#fff", color:active?"#fff":C.pri, fontSize:11.5, fontWeight:600, cursor:"pointer" }}>{children}</button>
);

/* ═══════════════════════════════════════════
   PAGE 1: OVERVIEW (Sheet "Overview" only)
   ═══════════════════════════════════════════ */
function Overview({ data }) {
  const ov = data.overview;
  const { years, yearly, forecast, diff, pctAct, pctFore, latestYear: ly, actualMonths: actLen } = ov;
  const latest = yearly[ly];
  const totalActual = latest?.total;
  const totalForecast = forecast?.total;
  const diffTotal = diff?.total;
  const prevYear = years.length >= 2 ? years[years.length - 2] : null;
  const prevTotal = prevYear ? yearly[prevYear]?.total : null;
  const growthVsLY = prevTotal && totalActual != null ? ((totalActual - prevTotal) / prevTotal * 100) : null;

  const monthlyData = MONTHS.map((m, i) => ({
    month: m,
    actual: latest?.months[i],
    forecast: forecast?.months[i] != null ? Math.round(forecast.months[i]) : null,
    diff: diff?.months[i],
    pctAct: pctAct?.months[i],
    pctFore: pctFore?.months[i],
    ...Object.fromEntries(years.map(y => [y, yearly[y]?.months[i]]))
  }));
  const yearlyTotals = years.map(y => ({ year: y.slice(1), trips: yearly[y]?.total }));

  const achievement = totalForecast ? (totalActual / totalForecast * 100) : null;
  const bestMonth = monthlyData.filter(d => d.actual != null).sort((a, b) => b.actual - a.actual)[0];
  const worstDiff = monthlyData.filter(d => d.diff != null).sort((a, b) => a.diff - b.diff)[0];
  const bestDiff = monthlyData.filter(d => d.diff != null).sort((a, b) => b.diff - a.diff)[0];

  return <div>
    <div style={{ display:"flex", gap:11, marginBottom:15, flexWrap:"wrap" }}>
      <KPI label={`Actual Trips (${ly?.slice(1)||""})`} value={fmt(totalActual)} sub={`${actLen} เดือน`} color={C.sec} />
      {forecast && <KPI label="Forecast" value={fmt(totalForecast)} color={C.acc} />}
      {diff && <KPI label="Diff (Act−Fore)" value={fmt(diffTotal)} sub={diffTotal!=null?pct(diffTotal/(totalForecast||1)*100):undefined} color={(diffTotal||0)>=0?C.pos:C.neg} />}
      {achievement!=null && <KPI label="Achievement" value={fmtDec(achievement)+"%"} color={achievement>=100?C.pos:C.neg} />}
      {growthVsLY!=null && <KPI label={`vs ${prevYear?.slice(1)||"LY"}`} value={pct(growthVsLY)} color={growthVsLY>=0?C.pos:C.neg} />}
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:11, marginBottom:11 }}>
      <Card title={`📊 Actual vs Forecast — Monthly ${ly?.slice(1)||""}`}>
        <ResponsiveContainer width="100%" height={215}><ComposedChart data={monthlyData} margin={{ left:-10, right:8 }}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" /><XAxis dataKey="month" tick={{ fontSize:11 }} /><YAxis tick={{ fontSize:11 }} /><Tooltip content={<TT />} /><Legend wrapperStyle={{ fontSize:10 }} /><Bar dataKey="actual" name="Actual" fill={C.sec} radius={[4,4,0,0]} barSize={25} />{forecast && <Line dataKey="forecast" name="Forecast" stroke={C.neg} strokeWidth={2} strokeDasharray="6 3" dot={{ r:3 }} />}</ComposedChart></ResponsiveContainer>
      </Card>
      <Card title="📈 Yearly Total Trips">
        <ResponsiveContainer width="100%" height={215}><BarChart data={yearlyTotals} margin={{ left:-10 }}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" /><XAxis dataKey="year" tick={{ fontSize:11 }} /><YAxis tick={{ fontSize:11 }} /><Tooltip content={<TT />} /><Bar dataKey="trips" name="Trips" fill={C.acc} radius={[4,4,0,0]} barSize={28}>{yearlyTotals.map((d,i)=><Cell key={i} fill={i===yearlyTotals.length-1?C.warn:C.acc} />)}</Bar></BarChart></ResponsiveContainer>
      </Card>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1.5fr", gap:11, marginBottom:11 }}>
      <Card title="📉 Monthly Diff (Actual − Forecast)">
        <ResponsiveContainer width="100%" height={195}><BarChart data={monthlyData.filter(d=>d.diff!=null)} margin={{ left:-10 }}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" /><XAxis dataKey="month" tick={{ fontSize:11 }} /><YAxis tick={{ fontSize:11 }} /><Tooltip content={<TT />} /><ReferenceLine y={0} stroke={C.mut} /><Bar dataKey="diff" name="Diff" radius={[4,4,0,0]} barSize={25}>{monthlyData.filter(d=>d.diff!=null).map((d,i)=><Cell key={i} fill={(d.diff||0)>=0?C.pos:C.neg} />)}</Bar></BarChart></ResponsiveContainer>
      </Card>
      <Card title="📈 Overall Trend (All Years)">
        <ResponsiveContainer width="100%" height={195}><LineChart data={monthlyData} margin={{ left:-10, right:8 }}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" /><XAxis dataKey="month" tick={{ fontSize:11 }} /><YAxis tick={{ fontSize:11 }} /><Tooltip content={<TT />} /><Legend wrapperStyle={{ fontSize:9.5 }} />{years.map((y,i)=><Line key={y} dataKey={y} stroke={i===years.length-1?C.warn:PAL[i%PAL.length]} strokeWidth={i===years.length-1?2.5:1.5} dot={i===years.length-1?{r:3}:false} connectNulls={false} />)}</LineChart></ResponsiveContainer>
      </Card>
    </div>
    <Insight items={[
      bestMonth && `เดือนที่มี Trips สูงสุด: ${bestMonth.month} (${fmt(bestMonth.actual)})`,
      worstDiff && `เดือนที่ Diff ต่ำสุด: ${worstDiff.month} (${fmt(worstDiff.diff)})`,
      bestDiff && (bestDiff.diff>0 ? `เดือนที่ทำได้เกิน Forecast: ${bestDiff.month} (+${fmt(bestDiff.diff)})` : "ยังไม่มีเดือนที่ทำได้เกิน Forecast"),
      achievement!=null && `Achievement: ${fmtDec(achievement)}%`,
      growthVsLY!=null && `เทียบกับ ${prevYear?.slice(1)||"ปีก่อน"}: ${growthVsLY<0?"ลดลง":"เพิ่มขึ้น"} ${Math.abs(growthVsLY).toFixed(1)}%`,
    ].filter(Boolean)} />
  </div>;
}

/* ═══════════════════════════════════════════
   PAGE 2: DC ANALYSIS (Sheet "DC" only)
   ═══════════════════════════════════════════ */
function DCAnalysis({ data }) {
  const ds = data.dcSheet;
  const { dcList, years, byDC, byYear, avgPerMonth, growthPct, pctByDC } = ds;
  const [selDC, setSelDC] = useState(dcList[0]||"");
  const [trendMode, setTrendMode] = useState("month");
  const latestYear = years[years.length-1]||"";

  // Monthly data for selected DC from byDC rows
  const dcYearsAvail = [...new Set(byDC.map(r => "Y"+r.year))].filter(y => years.includes(y)).sort();
  const trendMonth = MONTHS.map((m, mi) => {
    const row = { month: m };
    dcYearsAvail.forEach(y => {
      const found = byDC.find(r => "Y"+r.year === y && r.monthIdx === mi);
      row[y] = found ? readAny(found[selDC]) : null;
      if (row[y] != null && typeof row[y] === "string") row[y] = null; // '-' etc
    });
    return row;
  });

  // BY Year bar (from byYear sub-table)
  const trendYear = years.map(y => ({ year: y.slice(1), trips: byYear[y]?.[selDC] })).filter(d => d.trips != null);

  // Avg per month (from avgPerMonth sub-table)
  const avgBar = latestYear && avgPerMonth[latestYear] ? dcList.filter(d => avgPerMonth[latestYear][d] != null).map(d => ({ name: d, avg: avgPerMonth[latestYear][d], fill: PAL[dcList.indexOf(d)%PAL.length] })).sort((a,b) => b.avg - a.avg) : [];

  // Growth % (from growthPct sub-table)
  const growthBar = latestYear && growthPct[latestYear] ? dcList.filter(d => growthPct[latestYear][d] != null && typeof growthPct[latestYear][d] === "number").map(d => ({ name: d, growth: parseFloat((growthPct[latestYear][d]*100).toFixed(1)), fill: growthPct[latestYear][d] >= 0 ? C.pos : C.neg })).sort((a,b) => a.growth - b.growth) : [];

  // Contribution donut (from pctByDC sub-table)
  const donutData = latestYear && pctByDC[latestYear] ? dcList.filter(d => pctByDC[latestYear][d] != null && pctByDC[latestYear][d] > 0).map(d => ({ name: d, value: parseFloat((pctByDC[latestYear][d]*100).toFixed(1)), fill: PAL[dcList.indexOf(d)%PAL.length] })) : [];

  // KPIs from sub-tables
  const selTotal = byYear[latestYear]?.[selDC];
  const selAvg = avgPerMonth[latestYear]?.[selDC];
  const selGrowth = growthPct[latestYear]?.[selDC];
  const selContrib = pctByDC[latestYear]?.[selDC];

  const topG = growthBar.length ? [...growthBar].sort((a,b)=>b.growth-a.growth)[0] : null;
  const botG = growthBar.length ? [...growthBar].sort((a,b)=>a.growth-b.growth)[0] : null;

  return <div>
    <div style={{ display:"flex", gap:6, marginBottom:13, flexWrap:"wrap" }}>
      {dcList.map(d => <Btn key={d} active={selDC===d} onClick={()=>setSelDC(d)}>{d}</Btn>)}
    </div>
    <div style={{ display:"flex", gap:11, marginBottom:14, flexWrap:"wrap" }}>
      <KPI label={`${selDC} Total (${latestYear.slice(1)})`} value={fmt(selTotal)} color={C.sec} />
      <KPI label="Avg / Month" value={selAvg!=null?fmtDec(selAvg,1):"—"} color={C.acc} />
      <KPI label="Growth %" value={selGrowth!=null&&typeof selGrowth==="number"?pct(selGrowth*100):(typeof selGrowth==="string"?selGrowth:"—")} color={selGrowth!=null&&typeof selGrowth==="number"?(selGrowth>=0?C.pos:C.neg):C.mut} />
      <KPI label="Contribution %" value={selContrib!=null?fmtDec(selContrib*100,1)+"%":"—"} color={C.warn} />
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"1fr 2fr", gap:11, marginBottom:11 }}>
      <Card title={`🍩 DC Contribution (${latestYear.slice(1)})`}>
        {donutData.length ? <ResponsiveContainer width="100%" height={205}><PieChart><Pie data={donutData} cx="50%" cy="50%" innerRadius={44} outerRadius={80} paddingAngle={2} dataKey="value" label={({name,value})=>`${name} ${value}%`} labelLine={false} style={{fontSize:9}}>{donutData.map((d,i)=><Cell key={i} fill={d.fill}/>)}</Pie><Tooltip formatter={v=>v+"%"}/></PieChart></ResponsiveContainer> : <div style={{color:C.mut,fontSize:12,padding:20}}>No data</div>}
      </Card>
      <Card title={`📈 ${selDC} Trend`}>
        <div style={{ display:"flex", gap:5, marginBottom:6 }}>
          <Btn active={trendMode==="month"} onClick={()=>setTrendMode("month")}>Monthly</Btn>
          <Btn active={trendMode==="year"} onClick={()=>setTrendMode("year")}>Yearly</Btn>
        </div>
        {trendMode==="month" ? (
          <ResponsiveContainer width="100%" height={175}><LineChart data={trendMonth} margin={{left:-10,right:8}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="month" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9.5}}/>{dcYearsAvail.map((y,i)=><Line key={y} dataKey={y} stroke={i===dcYearsAvail.length-1?C.warn:PAL[i%PAL.length]} strokeWidth={i===dcYearsAvail.length-1?2.5:1.5} dot={i===dcYearsAvail.length-1?{r:3}:false} connectNulls={false}/>)}</LineChart></ResponsiveContainer>
        ) : (
          <ResponsiveContainer width="100%" height={175}><BarChart data={trendYear} margin={{left:-10}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="year" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip content={<TT/>}/><Bar dataKey="trips" name="Trips" fill={C.sec} radius={[4,4,0,0]} barSize={32}/></BarChart></ResponsiveContainer>
        )}
      </Card>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:11, marginBottom:11 }}>
      <Card title={`📊 Avg Trips / Month (${latestYear.slice(1)})`}>
        {avgBar.length ? <ResponsiveContainer width="100%" height={190}><BarChart data={avgBar} margin={{left:-10}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="name" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip content={<TT/>}/><Bar dataKey="avg" name="Avg/Month" radius={[4,4,0,0]} barSize={22}>{avgBar.map((d,i)=><Cell key={i} fill={d.fill}/>)}</Bar></BarChart></ResponsiveContainer> : <div style={{color:C.mut,fontSize:12}}>No data</div>}
      </Card>
      <Card title={`📉 Growth % by DC (${latestYear.slice(1)})`}>
        {growthBar.length ? <ResponsiveContainer width="100%" height={190}><BarChart data={growthBar} layout="vertical" margin={{left:8,right:14}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis type="number" tick={{fontSize:11}} tickFormatter={v=>v+"%"}/><YAxis dataKey="name" type="category" tick={{fontSize:11}} width={46}/><Tooltip formatter={v=>v+"%"}/><ReferenceLine x={0} stroke={C.mut}/><Bar dataKey="growth" name="Growth %" radius={[0,4,4,0]} barSize={15}>{growthBar.map((d,i)=><Cell key={i} fill={d.fill}/>)}</Bar></BarChart></ResponsiveContainer> : <div style={{color:C.mut,fontSize:12}}>No data</div>}
      </Card>
    </div>
    <Insight items={[
      topG && `DC เติบโตมากที่สุด: ${topG.name} (${pct(topG.growth)})`,
      botG && `DC ลดลงมากที่สุด: ${botG.name} (${pct(botG.growth)})`,
      selTotal!=null && `${selDC} Total ${latestYear.slice(1)}: ${fmt(selTotal)} trips`,
    ].filter(Boolean)} />
  </div>;
}

/* ═══════════════════════════════════════════
   PAGE 3: COMPARE (Sheet "Compare" only)
   ═══════════════════════════════════════════ */
function ComparePage({ data }) {
  const cp = data.compare;
  const { dcList, years: cpYears } = cp;
  const [selected, setSelected] = useState(()=>dcList.slice(0,Math.min(3,dcList.length)));
  const toggle = d => setSelected(s=>s.includes(d)?s.filter(x=>x!==d):[...s,d]);
  const active = selected.length ? selected : dcList;
  const ly = cpYears[cpYears.length-1]||"";
  const prevYr = cpYears.length>=2 ? cpYears[cpYears.length-2] : null;

  const hasDataMonths = MONTHS.map((_,mi) => active.some(d => cp.data[d]?.[ly]?.months[mi] != null)).filter(Boolean).length;
  const dataMonths = MONTHS.slice(0, hasDataMonths);

  const trendData = MONTHS.map((m,mi)=>{ const r={month:m}; active.forEach(d=>{r[d]=cp.data[d]?.[ly]?.months[mi]??null;}); return r; });
  const groupedData = dataMonths.map((m,mi)=>{ const r={month:m}; active.forEach(d=>{r[d]=cp.data[d]?.[ly]?.months[mi]??null;}); return r; });

  // Ranking by total from Compare sheet
  const ranking = dcList.map(d=>({dc:d, total:cp.data[d]?.[ly]?.total})).filter(d=>d.total!=null).sort((a,b)=>b.total-a.total);

  // Diff data from Compare sheet
  const diffData = active.map(d => ({ name:d, diff:cp.data[d]?.diff?.total, fill:(cp.data[d]?.diff?.total||0)>=0?C.pos:C.neg })).filter(d=>d.diff!=null);

  // Growth calc from Compare totals (current vs prev year)
  const growthComp = active.filter(d => cp.data[d]?.[ly]?.total!=null && prevYr && cp.data[d]?.[prevYr]?.total).map(d => {
    const cur = cp.data[d][ly].total;
    const prev = cp.data[d][prevYr].total;
    const g = ((cur-prev)/prev*100);
    return { name:d, growth:parseFloat(g.toFixed(1)), fill:g>=0?C.pos:C.neg };
  });

  // Contribution from Compare totals
  const totalAll = sum(dcList.map(d=>cp.data[d]?.[ly]?.total||0));
  const contribData = active.map(d=>({ name:d, contribution:totalAll?(parseFloat(((cp.data[d]?.[ly]?.total||0)/totalAll*100).toFixed(1))):0, fill:PAL[dcList.indexOf(d)%PAL.length] }));

  // Scatter
  const scatterData = active.filter(d=>cp.data[d]?.[ly]?.total!=null).map(d=>{
    const m=cp.data[d][ly].months.filter(v=>v!=null);
    const a=m.length?sum(m)/m.length:0;
    const prev=prevYr&&cp.data[d]?.[prevYr]?.total;
    const g=prev?((cp.data[d][ly].total-prev)/prev*100):0;
    return { name:d, avgTrips:Math.round(a), growth:parseFloat(g.toFixed(1)), fill:PAL[dcList.indexOf(d)%PAL.length] };
  });

  // Heatmap
  const heatVals=[]; active.forEach(d=>dataMonths.forEach((_,mi)=>{const v=cp.data[d]?.[ly]?.months[mi]; if(v!=null) heatVals.push(v);}));
  const maxHeat=Math.max(...heatVals,1);

  // Radar
  const radarData = useMemo(()=>{
    const mets=["Trips","Growth","Contribution","Consistency"];
    const vals={};
    active.forEach(d=>{
      const t=cp.data[d]?.[ly]?.total||0;
      const prev=prevYr?cp.data[d]?.[prevYr]?.total:null;
      const g=prev?((t-prev)/prev*100+100):100;
      const con=totalAll?(t/totalAll*100):0;
      const m=cp.data[d]?.[ly]?.months.filter(v=>v!=null)||[];
      const sd=m.length>1?Math.sqrt(m.reduce((s,v)=>s+Math.pow(v-avg(m),2),0)/m.length):0;
      const cv=avg(m)>0?(1-sd/avg(m))*100:0;
      vals[d]={Trips:t,Growth:g,Contribution:con,Consistency:Math.max(0,cv)};
    });
    const mx={}; mets.forEach(mt=>{mx[mt]=Math.max(...active.map(d=>vals[d]?.[mt]||0),1);});
    return mets.map(mt=>{const r={metric:mt}; active.forEach(d=>{r[d]=parseFloat(((vals[d]?.[mt]||0)/mx[mt]*100).toFixed(1));}); return r;});
  },[selected,ly,prevYr,totalAll]);

  return <div>
    <div style={{ display:"flex", gap:6, marginBottom:13, flexWrap:"wrap", alignItems:"center" }}>
      <span style={{ fontSize:11, color:C.mut, fontWeight:600, marginRight:2 }}>Select DCs:</span>
      {dcList.map((d,i)=>{const col=PAL[i%PAL.length];const on=selected.includes(d); return <button key={d} onClick={()=>toggle(d)} style={{padding:"5px 13px",borderRadius:7,border:on?`2px solid ${col}`:`1px solid ${C.bor}`,background:on?col+"18":"#fff",color:on?col:C.mut,fontSize:11,fontWeight:600,cursor:"pointer"}}>{d}</button>;})}
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"3fr 2fr", gap:11, marginBottom:11 }}>
      <Card title={`📈 Trend Comparison (${ly.slice(1)})`}>
        <ResponsiveContainer width="100%" height={205}><LineChart data={trendData} margin={{left:-10,right:8}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="month" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9.5}}/>{active.map(d=><Line key={d} dataKey={d} stroke={PAL[dcList.indexOf(d)%PAL.length]} strokeWidth={2} dot={{r:2}} connectNulls={false}/>)}</LineChart></ResponsiveContainer>
      </Card>
      <Card title={`🏆 Ranking (${ly.slice(1)})`}>
        <div style={{display:"flex",flexDirection:"column",gap:6,paddingTop:2}}>
          {ranking.map((r,i)=>(
            <div key={r.dc} style={{display:"flex",alignItems:"center",gap:7}}>
              <span style={{fontSize:12,fontWeight:700,color:i===0?C.warn:C.mut,width:20}}>#{i+1}</span>
              <span style={{fontSize:11,fontWeight:600,width:48,color:C.pri}}>{r.dc}</span>
              <div style={{flex:1,background:"#f1f5f9",borderRadius:4,height:15,overflow:"hidden"}}><div style={{width:`${ranking[0].total?(r.total/ranking[0].total)*100:0}%`,height:"100%",background:PAL[dcList.indexOf(r.dc)%PAL.length],borderRadius:4}}/></div>
              <span style={{fontSize:10.5,fontWeight:600,color:C.pri,width:48,textAlign:"right"}}>{fmt(r.total)}</span>
            </div>
          ))}
        </div>
      </Card>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:11, marginBottom:11 }}>
      <Card title="📊 Grouped Comparison">
        <ResponsiveContainer width="100%" height={190}><BarChart data={groupedData} margin={{left:-10}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="month" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/><Tooltip content={<TT/>}/><Legend wrapperStyle={{fontSize:9}}/>{active.map(d=><Bar key={d} dataKey={d} fill={PAL[dcList.indexOf(d)%PAL.length]} radius={[2,2,0,0]} barSize={Math.max(6,Math.floor(36/active.length))}/>)}</BarChart></ResponsiveContainer>
      </Card>
      <Card title="📉 Diff (Latest vs Previous)">
        {diffData.length ? <ResponsiveContainer width="100%" height={190}><BarChart data={diffData} margin={{left:-10}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="name" tick={{fontSize:11}}/><YAxis tick={{fontSize:11}}/><Tooltip content={<TT/>}/><ReferenceLine y={0} stroke={C.mut}/><Bar dataKey="diff" name="Diff" radius={[4,4,0,0]} barSize={22}>{diffData.map((d,i)=><Cell key={i} fill={d.fill}/>)}</Bar></BarChart></ResponsiveContainer> : <div style={{color:C.mut}}>No diff data</div>}
      </Card>
      <Card title="🍩 Contribution %">
        <ResponsiveContainer width="100%" height={190}><PieChart><Pie data={contribData} cx="50%" cy="50%" innerRadius={36} outerRadius={70} paddingAngle={2} dataKey="contribution" label={({name,value})=>`${name} ${value}%`} labelLine={false} style={{fontSize:9}}>{contribData.map((d,i)=><Cell key={i} fill={d.fill}/>)}</Pie><Tooltip formatter={v=>v+"%"}/></PieChart></ResponsiveContainer>
      </Card>
    </div>
    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:11, marginBottom:11 }}>
      <Card title="🔵 Scatter: Avg Trips vs Growth %">
        <ResponsiveContainer width="100%" height={205}><ScatterChart margin={{left:-5,right:8,bottom:5}}><CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/><XAxis dataKey="avgTrips" type="number" name="Avg Trips" tick={{fontSize:11}}/><YAxis dataKey="growth" type="number" name="Growth %" tick={{fontSize:11}} tickFormatter={v=>v+"%"}/><Tooltip cursor={{strokeDasharray:"3 3"}} formatter={(v,n)=>n==="growth"?v+"%":v} labelFormatter={()=>""}/><Scatter data={scatterData} shape="circle">{scatterData.map((d,i)=><Cell key={i} fill={d.fill} r={8}/>)}</Scatter></ScatterChart></ResponsiveContainer>
        <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>{scatterData.map(d=><span key={d.name} style={{fontSize:9.5,color:d.fill,fontWeight:600}}>● {d.name}</span>)}</div>
      </Card>
      <Card title="🕸️ Radar: Multi-Dimension">
        <ResponsiveContainer width="100%" height={220}><RadarChart data={radarData}><PolarGrid stroke={C.bor}/><PolarAngleAxis dataKey="metric" tick={{fontSize:10}}/><PolarRadiusAxis tick={{fontSize:8}} domain={[0,100]}/>{active.map(d=><Radar key={d} name={d} dataKey={d} stroke={PAL[dcList.indexOf(d)%PAL.length]} fill={PAL[dcList.indexOf(d)%PAL.length]} fillOpacity={.12} strokeWidth={2}/>)}<Legend wrapperStyle={{fontSize:9.5}}/><Tooltip/></RadarChart></ResponsiveContainer>
      </Card>
    </div>
    <Card title={`🗺️ Heatmap (${ly.slice(1)})`} style={{marginBottom:11}}>
      <div style={{overflowX:"auto"}}><div style={{display:"grid",gridTemplateColumns:`68px repeat(${dataMonths.length},1fr)`,gap:2,fontSize:11}}>
        <div/>
        {dataMonths.map(m=><div key={m} style={{fontWeight:700,padding:4,textAlign:"center",color:C.pri}}>{m}</div>)}
        {active.map(d=><React.Fragment key={d}>
          <div style={{fontWeight:600,padding:4,color:C.pri,display:"flex",alignItems:"center"}}>{d}</div>
          {dataMonths.map((m,mi)=>{const v=cp.data[d]?.[ly]?.months[mi];const int=v!=null?Math.max(.08,v/maxHeat):0;return <div key={d+m} style={{padding:4,textAlign:"center",background:v!=null?`rgba(37,99,235,${int*.85})`:"#f8fafc",color:int>.45?"#fff":C.pri,borderRadius:4,fontWeight:600,fontSize:11}}>{v!=null?v:"—"}</div>;})}
        </React.Fragment>)}
      </div></div>
    </Card>
    <Insight items={[
      ranking[0] && `DC อันดับ 1: ${ranking[0].dc} (${fmt(ranking[0].total)} trips)`,
      growthComp.length && `เติบโตมากที่สุด: ${[...growthComp].sort((a,b)=>b.growth-a.growth)[0]?.name} (${pct([...growthComp].sort((a,b)=>b.growth-a.growth)[0]?.growth)})`,
      growthComp.length && `ลดลงมากที่สุด: ${[...growthComp].sort((a,b)=>a.growth-b.growth)[0]?.name} (${pct([...growthComp].sort((a,b)=>a.growth-b.growth)[0]?.growth)})`,
    ].filter(Boolean)} />
  </div>;
}

/* ═══════════════════════════════════════════
   MAIN APP
   ═══════════════════════════════════════════ */
export default function App() {
  const [data, setData] = useState(null);
  const [page, setPage] = useState(0);
  const [uploadMsg, setUploadMsg] = useState(null);
  const fileRef = useRef();

  const handleFile = useCallback(async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    setUploadMsg({ type:"loading", text:"⏳ Processing..." });
    try {
      const ab = await file.arrayBuffer();
      const result = parseExcel(ab);
      if (result.error) { setUploadMsg({ type:"error", text: result.error }); }
      else { setData(result); setUploadMsg({ type:"success", text:`✓ "${file.name}" loaded — Overview: ${result.overview.years.length} years, DC: ${result.dcSheet.dcList.length} DCs, Compare: ${result.compare.dcList.length} DCs` }); setPage(0); }
    } catch (err) { setUploadMsg({ type:"error", text:"Parse failed: "+err.message }); }
    if (fileRef.current) fileRef.current.value = "";
  }, []);

  const pages = ["Overview","DC Analysis","Compare"];

  // Landing screen
  if (!data) return (
    <div style={{ fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif", background:C.bg, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ textAlign:"center", maxWidth:460, padding:40 }}>
        <div style={{ fontSize:56, marginBottom:16 }}>🚛</div>
        <div style={{ fontSize:22, fontWeight:800, color:C.pri, marginBottom:8 }}>BH Truck Volume Dashboard</div>
        <div style={{ fontSize:14, color:C.mut, marginBottom:28, lineHeight:1.6 }}>อัปโหลดไฟล์ Excel (.xlsx) เพื่อเริ่มวิเคราะห์ข้อมูล<br/>ระบบรองรับ Template: Overview, DC, Compare</div>
        {uploadMsg && <div style={{ padding:"10px 16px", borderRadius:8, marginBottom:16, background:uploadMsg.type==="error"?"#fef2f2":"#fffbeb", color:uploadMsg.type==="error"?C.neg:C.warn, fontSize:13, fontWeight:600 }}>{uploadMsg.text}</div>}
        <label style={{ display:"inline-flex", alignItems:"center", gap:8, padding:"14px 32px", borderRadius:10, background:`linear-gradient(135deg,${C.pri},${C.sec})`, color:"#fff", fontSize:15, fontWeight:700, cursor:"pointer", boxShadow:"0 4px 14px rgba(37,99,235,.3)" }}>
          📁 Upload Excel File
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />
        </label>
        <div style={{ fontSize:11, color:C.mut, marginTop:16 }}>รองรับไฟล์ .xlsx และ .xls</div>
      </div>
    </div>
  );

  return (
    <div style={{ fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif", background:C.bg, minHeight:"100vh", color:C.pri }}>
      <div style={{ background:`linear-gradient(135deg,#0f172a,#1e3a5f)`, padding:"13px 20px", display:"flex", alignItems:"center", justifyContent:"space-between", flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ fontSize:18, fontWeight:800, color:"#fff" }}>🚛 BH Truck Volume Dashboard</div>
          <div style={{ fontSize:10, color:"rgba(255,255,255,.5)", marginTop:1 }}>Data source: Excel file • Strict sheet-to-page mapping</div>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:6 }}>
          <div style={{ display:"flex", gap:2 }}>
            {pages.map((p,i)=><button key={p} onClick={()=>setPage(i)} style={{padding:"6px 16px",borderRadius:7,border:"none",background:page===i?"rgba(255,255,255,.18)":"transparent",color:page===i?"#fff":"rgba(255,255,255,.4)",fontSize:12,fontWeight:600,cursor:"pointer"}}>{p}</button>)}
          </div>
          <label style={{ padding:"6px 14px", borderRadius:7, background:"rgba(255,255,255,.12)", color:"#fff", fontSize:11.5, fontWeight:600, cursor:"pointer", border:"1px solid rgba(255,255,255,.2)", display:"flex", alignItems:"center", gap:4 }}>
            📁 Upload Excel
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display:"none" }} />
          </label>
        </div>
      </div>
      {uploadMsg && uploadMsg.type!=="loading" && (
        <div style={{ padding:"8px 20px", background:uploadMsg.type==="error"?"#fef2f2":"#f0fdf4", color:uploadMsg.type==="error"?C.neg:C.pos, fontSize:11.5, fontWeight:600, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <span>{uploadMsg.text}</span>
          <button onClick={()=>setUploadMsg(null)} style={{background:"none",border:"none",cursor:"pointer",fontSize:15,color:"inherit"}}>×</button>
        </div>
      )}
      <div style={{ padding:"14px 18px", maxWidth:1280, margin:"0 auto" }}>
        {page === 0 && <Overview data={data} />}
        {page === 1 && <DCAnalysis data={data} />}
        {page === 2 && <ComparePage data={data} />}
      </div>
      <div style={{ textAlign:"center", padding:"8px", fontSize:9, color:C.mut }}>
        Overview → Sheet "Overview" | DC Analysis → Sheet "DC" | Compare → Sheet "Compare" | No cross-sheet data
      </div>
    </div>
  );
}
