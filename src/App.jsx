import { useState, useEffect, useMemo } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

// ── API ───────────────────────────────────────────────────────────────────────
const API = "https://giiava-warehouse-api.azurewebsites.net/api/api";
async function apiGet(action, params = {}) {
  const qs = new URLSearchParams({ action, ...params }).toString();
  const res = await fetch(`${API}?${qs}`);
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

// ── Styles ───────────────────────────────────────────────────────────────────
const S = {
  bg:"#060f1a", surface:"#0c1a2e", border:"#1a3a5c",
  text:"#e8f0fe", muted:"#5a7a9a",
  orange:"#f97316", green:"#22c55e", blue:"#38bdf8",
  red:"#ef4444", purple:"#a78bfa", yellow:"#fbbf24",
};
const mono = "'IBM Plex Mono','Courier New',monospace";
const sans = "'Barlow','IBM Plex Sans',system-ui,sans-serif";

// ── Helpers ───────────────────────────────────────────────────────────────────
function fmtDate(d) {
  if (!d) return "—";
  const dt = new Date(d);
  return dt.toLocaleDateString("en-GB", { day:"2-digit", month:"short" });
}
function isoDay(d) { return d instanceof Date ? d.toISOString().split("T")[0] : (d||"").split("T")[0]; }
function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate()+n); return r; }
function isSunday(d) { return new Date(d).getDay() === 0; }
function lastWorkDay(days) {
  const today = isoDay(new Date());
  const sorted = [...days].sort();
  for (let i = sorted.length-1; i >= 0; i--) {
    if (sorted[i] <= today && !isSunday(sorted[i])) return sorted[i];
  }
  return null;
}
function weekKey(dateStr) {
  const d = new Date(dateStr);
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  const mon = new Date(d); mon.setDate(d.getDate() + diff);
  return isoDay(mon);
}
function avg(arr) { return arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0; }
function kgToBags(kg) { return Math.round(kg / 25); }

// ── Card wrapper ──────────────────────────────────────────────────────────────
function Card({ title, subtitle, accent=S.blue, children }) {
  return (
    <div style={{ background:S.surface, border:`1px solid ${S.border}`, borderRadius:12,
      marginBottom:16, overflow:"hidden" }}>
      <div style={{ padding:"12px 16px", borderBottom:`1px solid ${S.border}`,
        display:"flex", justifyContent:"space-between", alignItems:"baseline" }}>
        <div style={{ fontFamily:mono, fontSize:12, fontWeight:700, color:accent,
          letterSpacing:"0.08em", textTransform:"uppercase" }}>{title}</div>
        {subtitle && <div style={{ fontFamily:mono, fontSize:10, color:S.muted }}>{subtitle}</div>}
      </div>
      <div style={{ padding:"14px 16px" }}>{children}</div>
    </div>
  );
}

// ── Stat pill ─────────────────────────────────────────────────────────────────
function Stat({ label, value, color=S.text, dim=false }) {
  return (
    <div style={{ background:"#04080f", borderRadius:8, padding:"8px 12px",
      textAlign:"center", opacity:dim?0.6:1 }}>
      <div style={{ fontFamily:mono, fontSize:10, color:S.muted, marginBottom:2 }}>{label}</div>
      <div style={{ fontFamily:mono, fontSize:17, fontWeight:700, color }}>{value}</div>
    </div>
  );
}

// ── SVG Bar Chart ─────────────────────────────────────────────────────────────
function BarChart({ data, highlight, avgVal, color=S.blue, height=80, unit="" }) {
  if (!data.length) return <div style={{fontFamily:mono,fontSize:12,color:S.muted}}>No data</div>;
  const max = Math.max(...data.map(d=>d.value), 1);
  const w = 100 / data.length;
  return (
    <svg viewBox={`0 0 100 ${height+12}`} style={{ width:"100%", display:"block" }}
      preserveAspectRatio="none">
      {avgVal != null && (
        <line x1="0" y1={height - (avgVal/max)*height} x2="100"
          y2={height - (avgVal/max)*height}
          stroke={S.yellow} strokeWidth="0.4" strokeDasharray="2,1" />
      )}
      {data.map((d, i) => {
        const bh = Math.max((d.value/max)*height, 0.5);
        const isHL = d.day === highlight;
        return (
          <g key={d.day}>
            <rect x={i*w+0.5} y={height-bh} width={w-1} height={bh}
              fill={isHL ? S.orange : color} rx="0.5" opacity={isHL?1:0.75} />
            {isHL && (
              <text x={i*w+w/2} y={height-bh-2} textAnchor="middle"
                fontSize="3.5" fill={S.orange} fontFamily={mono}>
                {d.value}{unit}
              </text>
            )}
          </g>
        );
      })}
      {data.filter((_,i) => i===0 || i===data.length-1 || data[i].day===highlight).map((d,_i,arr) => {
        const i = data.indexOf(d);
        return (
          <text key={"l"+d.day} x={i*w+w/2} y={height+10}
            textAnchor="middle" fontSize="3" fill={S.muted} fontFamily={mono}>
            {fmtDate(d.day)}
          </text>
        );
      })}
    </svg>
  );
}

// ── Sparkline ─────────────────────────────────────────────────────────────────
function Sparkline({ data, color=S.blue, height=40 }) {
  if (data.length < 2) return null;
  const max = Math.max(...data.map(d=>d.value), 1);
  const min = Math.min(...data.map(d=>d.value), 0);
  const range = max - min || 1;
  const pts = data.map((d,i) => {
    const x = (i/(data.length-1))*100;
    const y = height - ((d.value - min)/range)*height;
    return `${x},${y}`;
  }).join(" ");
  return (
    <svg viewBox={`0 0 100 ${height}`} style={{ width:"100%", display:"block" }}
      preserveAspectRatio="none">
      <polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" />
    </svg>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 1: Daily Production Output
// ─────────────────────────────────────────────────────────────────────────────
function ProductionSection({ production }) {
  const last30 = production.slice(-30);
  const hl = lastWorkDay(production.map(d=>d.day));
  const hlRow = production.find(d=>d.day===hl);
  const avgBags = Math.round(avg(last30.map(d=>d.bags||0)));
  const best = last30.reduce((a,b)=>(b.bags||0)>(a.bags||0)?b:a, last30[0]||{});
  return (
    <Card title="Daily Production Output" subtitle="last 30 days" accent={S.blue}>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8, marginBottom:12 }}>
        <Stat label="Yesterday" value={`${hlRow?.bags??'—'} bags`} color={S.orange} />
        <Stat label="30d avg" value={`${avgBags} bags`} color={S.blue} />
        <Stat label="Best day" value={`${best?.bags??'—'} bags`} color={S.green} />
      </div>
      <BarChart
        data={last30.map(d=>({ day:d.day, value:d.bags||0 }))}
        highlight={hl} avgVal={avgBags} color={S.blue} height={80} />
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 2: Utilities per MT
// ─────────────────────────────────────────────────────────────────────────────
function UtilitiesSection({ utilities, production, acetoneTanks }) {
  const prodMap = useMemo(() => {
    const m = {};
    production.forEach(d => { m[d.day] = (d.csl_kg||0)/1000; }); // MT crude processed
    return m;
  }, [production]);

  // Build an acetone consumed map keyed by day
  const acetoneConsumedMap = useMemo(() => {
    const m = {};
    acetoneTanks.forEach(d => { if (d.acetone_consumed_kg != null) m[d.day] = d.acetone_consumed_kg; });
    return m;
  }, [acetoneTanks]);

  const joined = useMemo(() => {
    // Union of utility days and acetone days
    const daySet = new Set([
      ...utilities.map(u => u.day),
      ...Object.keys(acetoneConsumedMap),
    ]);
    return Array.from(daySet).sort().map(day => {
      const u  = utilities.find(x => x.day === day) || {};
      const mt = prodMap[day] || 0;
      const ac = acetoneConsumedMap[day] ?? null;
      return {
        day,
        diesel:      mt > 0 && u.diesel_L       != null ? +((u.diesel_L      )/mt).toFixed(1) : null,
        electricity: mt > 0 && u.electricity_kwh != null ? +((u.electricity_kwh)/mt).toFixed(1) : null,
        water:       mt > 0 && u.water_m3        != null ? +((u.water_m3      )/mt).toFixed(1) : null,
        acetone:     mt > 0 && ac               != null ? +(ac/mt).toFixed(1) : null,
      };
    }).filter(d => d.diesel!==null || d.electricity!==null || d.acetone!==null);
  }, [utilities, prodMap, acetoneConsumedMap]);

  const hl = lastWorkDay(joined.map(d=>d.day));
  const hlRow = joined.find(d=>d.day===hl);
  const hlAcetone = acetoneTanks.find(d=>d.day===hl);
  const avg30 = (key) => +avg(joined.slice(-30).map(d=>d[key]||0)).toFixed(1);

  function SubChart({ label, dataKey, unit, color }) {
    const vals = joined.slice(-30).map(d=>({ day:d.day, value:d[dataKey]||0 }));
    const hlVal = hlRow?.[dataKey];
    const a = avg30(dataKey);
    return (
      <div style={{ background:"#04080f", borderRadius:8, padding:"10px 12px" }}>
        <div style={{ fontFamily:mono, fontSize:9, color:S.muted, marginBottom:4,
          letterSpacing:"0.1em", textTransform:"uppercase" }}>{label}</div>
        <div style={{ fontFamily:mono, fontSize:20, fontWeight:700, color }}>
          {hlVal != null ? `${hlVal}` : "—"}
          <span style={{ fontSize:11, color:S.muted, marginLeft:4 }}>{unit}</span>
        </div>
        <div style={{ fontFamily:mono, fontSize:9, color:S.muted, marginBottom:6 }}>
          30d avg: {a} {unit}
        </div>
        <Sparkline data={vals} color={color} height={32} />
      </div>
    );
  }

  const acetoneLast = acetoneTanks.slice(-1)[0];
  const acetoneStockAvg = +avg(acetoneTanks.slice(-30).map(d=>d.total_acetone_kg||0)).toFixed(0);
  return (
    <Card title="Utilities per MT Produced" subtitle="last working day + 30d avg" accent={S.purple}>
      {/* 4 per-MT metrics in 2×2 grid */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:10 }}>
        <SubChart label="Diesel"       dataKey="diesel"      unit="L/MT"   color={S.orange} />
        <SubChart label="Electricity"  dataKey="electricity" unit="kWh/MT" color={S.yellow} />
        <SubChart label="Water"        dataKey="water"       unit="m³/MT"  color={S.blue}   />
        <SubChart label="Acetone"      dataKey="acetone"     unit="kg/MT"  color={S.green}  />
      </div>
      {/* Acetone stock level — full-width below */}
      <div style={{ background:"#04080f", borderRadius:8, padding:"10px 12px" }}>
        <div style={{ fontFamily:mono, fontSize:9, color:S.muted, marginBottom:4,
          letterSpacing:"0.1em", textTransform:"uppercase" }}>Total Acetone Stock</div>
        <div style={{ fontFamily:mono, fontSize:20, fontWeight:700, color:S.green }}>
          {hlAcetone?.total_acetone_kg != null
            ? `${Math.round(hlAcetone.total_acetone_kg).toLocaleString()}`
            : acetoneLast?.total_acetone_kg != null
            ? `${Math.round(acetoneLast.total_acetone_kg).toLocaleString()}`
            : "—"}
          <span style={{ fontSize:11, color:S.muted, marginLeft:4 }}>kg available</span>
        </div>
        <div style={{ fontFamily:mono, fontSize:9, color:S.muted, marginBottom:6 }}>
          30d avg: {acetoneStockAvg.toLocaleString()} kg · RAT+RMT6+0.88×RMT7+0.95×IMT7
        </div>
        <Sparkline
          data={acetoneTanks.slice(-30).map(d=>({ day:d.day, value:d.total_acetone_kg||0 }))}
          color={S.green} height={28} />
      </div>
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 3: Decanter 7 Uptime Calendar
// ─────────────────────────────────────────────────────────────────────────────
function DecanterSection({ decanter }) {
  const decMap = useMemo(() => {
    const m = {};
    decanter.forEach(d => { m[isoDay(d.day)] = d.dec7_ran; });
    return m;
  }, [decanter]);

  const today = new Date(); today.setHours(0,0,0,0);
  const start = addDays(today, -89);
  // Align start to Monday
  const dow = start.getDay(); // 0=Sun
  const alignedStart = addDays(start, dow === 0 ? -6 : 1-dow);

  const days = [];
  for (let d = new Date(alignedStart); d <= today; d = addDays(d, 1)) {
    days.push(isoDay(d));
  }

  const todayStr = isoDay(today);
  const last30Days = days.filter(d => d >= isoDay(addDays(today,-30)) && d <= todayStr && !isSunday(d));
  const gaps = last30Days.filter(d => decMap[d] === 0 || (decMap[d] === undefined && d < todayStr)).length;

  function cellColor(day) {
    if (isSunday(day)) return "#111";
    if (day > todayStr) return S.surface;
    const ran = decMap[day];
    if (ran === 1) return "#0a2218";
    if (ran === 0) return "#2d0000";
    return "#1a1a1a"; // no data logged yet
  }
  function textColor(day) {
    if (isSunday(day)) return "#333";
    if (day > todayStr) return S.muted;
    const ran = decMap[day];
    if (ran === 1) return S.green;
    if (ran === 0) return S.red;
    return "#333";
  }

  const weeks = [];
  for (let i = 0; i < days.length; i += 7) {
    weeks.push(days.slice(i, i+7));
  }

  return (
    <Card title="Decanter 7 — Uptime Calendar" subtitle="90 days" accent={S.red}>
      <div style={{ display:"flex", gap:12, marginBottom:12, flexWrap:"wrap" }}>
        <div style={{ fontFamily:mono, fontSize:22, fontWeight:700, color:gaps>5?S.red:gaps>2?S.yellow:S.green }}>
          {gaps} gaps
        </div>
        <div style={{ fontFamily:sans, fontSize:13, color:S.muted, alignSelf:"center" }}>
          in last 30 working days
        </div>
      </div>
      <div style={{ display:"flex", gap:4, marginBottom:8 }}>
        {["Mon","Tue","Wed","Thu","Fri","Sat","Sun"].map(d => (
          <div key={d} style={{ flex:1, fontFamily:mono, fontSize:9, color:S.muted, textAlign:"center" }}>{d}</div>
        ))}
      </div>
      {weeks.map((week, wi) => (
        <div key={wi} style={{ display:"flex", gap:4, marginBottom:4 }}>
          {week.map(day => (
            <div key={day} style={{ flex:1, aspectRatio:"1", background:cellColor(day),
              borderRadius:4, display:"flex", alignItems:"center", justifyContent:"center",
              border: day===todayStr ? `1px solid ${S.orange}` : "1px solid transparent" }}>
              <span style={{ fontFamily:mono, fontSize:9, color:textColor(day) }}>
                {new Date(day).getDate()}
              </span>
            </div>
          ))}
          {week.length < 7 && Array.from({length:7-week.length}).map((_,i) => (
            <div key={"pad"+i} style={{ flex:1 }} />
          ))}
        </div>
      ))}
      <div style={{ display:"flex", gap:16, marginTop:10, flexWrap:"wrap" }}>
        {[["#0a2218",S.green,"Ran"],["#2d0000",S.red,"Gap"],["#1a1a1a","#333","No log"],["#111","#333","Sunday"]]
          .map(([bg,c,l]) => (
            <div key={l} style={{ display:"flex", alignItems:"center", gap:5 }}>
              <div style={{ width:10, height:10, background:bg, borderRadius:2 }} />
              <span style={{ fontFamily:mono, fontSize:9, color:c }}>{l}</span>
            </div>
          ))}
      </div>
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 4: Packaging Efficiency
// ─────────────────────────────────────────────────────────────────────────────
function PackagingEfficiencySection({ packagingOut, attendance }) {
  const attMap = useMemo(() => {
    const m = {};
    attendance.forEach(d => { m[isoDay(d.day)] = d.workers_present||0; });
    return m;
  }, [attendance]);

  const data = useMemo(() => packagingOut
    .filter(d => attMap[isoDay(d.day)] > 0)
    .map(d => {
      const workers = attMap[isoDay(d.day)] || 0;
      const benchmark = workers * 40;
      const pct = benchmark > 0 ? Math.round((d.bags_packed/benchmark)*100) : null;
      return { day: isoDay(d.day), pct, bags: d.bags_packed, workers };
    })
    .filter(d => d.pct !== null)
    .slice(-30), [packagingOut, attMap]);

  const hl = lastWorkDay(data.map(d=>d.day));
  const hlRow = data.find(d=>d.day===hl);
  const avg30 = Math.round(avg(data.map(d=>d.pct||0)));

  function barColor(pct) {
    if (pct >= 100) return S.green;
    if (pct >= 75) return S.yellow;
    return S.red;
  }

  return (
    <Card title="Packaging Efficiency" subtitle="bags packed ÷ (workers × 40)" accent={S.green}>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:12 }}>
        <Stat label="Last working day" value={hlRow ? `${hlRow.pct}%` : "—"}
          color={hlRow ? barColor(hlRow.pct) : S.muted} />
        <Stat label="30d avg" value={`${avg30}%`} color={barColor(avg30)} />
      </div>
      {data.length === 0
        ? <div style={{ fontFamily:mono, fontSize:12, color:S.muted }}>No data yet</div>
        : <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
            {data.slice(-14).map(d => (
              <div key={d.day} style={{ display:"flex", alignItems:"center", gap:8 }}>
                <div style={{ fontFamily:mono, fontSize:10, color:S.muted, width:52, flexShrink:0 }}>
                  {fmtDate(d.day)}
                </div>
                <div style={{ flex:1, background:"#04080f", borderRadius:4, height:16, position:"relative" }}>
                  <div style={{ width:`${Math.min(d.pct,150)}%`, maxWidth:"100%", height:"100%",
                    background: d.day===hl ? S.orange : barColor(d.pct),
                    borderRadius:4, opacity: d.day===hl?1:0.8 }} />
                </div>
                <div style={{ fontFamily:mono, fontSize:10,
                  color: d.day===hl ? S.orange : barColor(d.pct), width:36, textAlign:"right" }}>
                  {d.pct}%
                </div>
              </div>
            ))}
          </div>
      }
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 5: Packaging Surplus / Shortfall
// ─────────────────────────────────────────────────────────────────────────────
function SurplusSection({ production, packagingOut }) {
  const data = useMemo(() => {
    const prodMap = {};
    production.forEach(d => { prodMap[isoDay(d.day)] = d.bags||0; });
    const packMap = {};
    packagingOut.forEach(d => { packMap[isoDay(d.day)] = d.bags_packed||0; });

    const allDays = [...new Set([
      ...production.map(d=>isoDay(d.day)),
      ...packagingOut.map(d=>isoDay(d.day)),
    ])].sort();

    let cumulative = 0;
    return allDays.map(day => {
      cumulative += (packMap[day]||0) - (prodMap[day]||0);
      return { day, cumulative };
    });
  }, [production, packagingOut]);

  const current = data[data.length-1]?.cumulative ?? 0;
  const max = Math.max(...data.map(d=>Math.abs(d.cumulative)), 1);

  return (
    <Card title="Packaging Surplus / Shortfall" subtitle="cumulative bags packed − produced"
      accent={current >= 0 ? S.green : S.red}>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:12 }}>
        <Stat label="Current balance"
          value={`${current >= 0 ? "+" : ""}${current} bags`}
          color={current >= 0 ? S.green : S.red} />
        <Stat label="Trend"
          value={current >= 0 ? "Surplus" : "Shortfall"}
          color={current >= 0 ? S.green : S.red} />
      </div>
      <svg viewBox={`0 0 100 60`} style={{ width:"100%", display:"block" }} preserveAspectRatio="none">
        {/* Zero line */}
        <line x1="0" y1="30" x2="100" y2="30" stroke={S.border} strokeWidth="0.5" />
        {data.map((d, i) => {
          const x = (i/(data.length-1||1))*100;
          const y = 30 - (d.cumulative/max)*28;
          const x2 = i < data.length-1 ? ((i+1)/(data.length-1||1))*100 : x;
          const y2 = i < data.length-1 ? 30 - (data[i+1].cumulative/max)*28 : y;
          const col = d.cumulative >= 0 ? S.green : S.red;
          return <line key={d.day} x1={x} y1={y} x2={x2} y2={y2}
            stroke={col} strokeWidth="1.2" />;
        })}
        {data.length > 0 && (() => {
          const last = data[data.length-1];
          const lx = 100;
          const ly = 30 - (last.cumulative/max)*28;
          return <circle cx={lx} cy={ly} r="1.5"
            fill={last.cumulative>=0?S.green:S.red} />;
        })()}
      </svg>
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 6: Crude Availability
// ─────────────────────────────────────────────────────────────────────────────
const CRUDE_LOW_KG      = 50000;
const CRUDE_CRITICAL_KG = 10000;

function CrudeSection({ crude }) {
  const projected = useMemo(() => {
    if (!crude || crude.latestTankKg == null) return [];
    let balance = crude.latestTankKg;
    const planMap = {};
    (crude.plan||[]).forEach(p => {
      const d = (p.PlanDate||"").split("T")[0];
      planMap[d] = (p.ConsumptionMT||0)*1000;
    });
    const purchaseMap = {};
    (crude.purchases||[]).forEach(p => {
      const d = (p.ExpectedArrival||"").split("T")[0];
      purchaseMap[d] = (purchaseMap[d]||0) + (p.QuantityMT||0)*1000;
    });
    const today = new Date(); today.setHours(0,0,0,0);
    const weeks = [];
    for (let w = 0; w < 6; w++) {
      for (let d = 0; d < 7; d++) {
        const day = isoDay(addDays(today, w*7+d));
        balance += (purchaseMap[day]||0);
        balance -= (planMap[day]||0);
      }
      weeks.push({ week: w+1, balance: Math.round(balance) });
    }
    return weeks;
  }, [crude]);

  if (!crude || crude.latestTankKg == null) {
    return (
      <Card title="Crude Oil Availability" subtitle="6-week projection" accent={S.yellow}>
        <div style={{ fontFamily:sans, color:S.yellow, fontSize:13 }}>
          ⚠ No kg readings logged yet. Have the Utility operator log RMT6 and RMT7 in the Tank Levels panel.
        </div>
      </Card>
    );
  }

  const shortage = projected.find(w => w.balance <= 0);
  const maxKg = Math.max(...projected.map(w=>w.balance), crude.latestTankKg, 1);

  function wkColor(kg) {
    if (kg <= 0) return S.red;
    if (kg < CRUDE_CRITICAL_KG) return "#fb923c";
    if (kg < CRUDE_LOW_KG) return S.yellow;
    return S.green;
  }

  return (
    <Card title="Crude Oil Availability" subtitle="6-week projection" accent={S.purple}>
      {shortage && (
        <div style={{ background:"#2d0000", border:`1px solid ${S.red}`, borderRadius:8,
          padding:"10px 14px", marginBottom:12, fontFamily:sans, fontSize:13, color:S.red }}>
          ⚠ Projected shortfall at Week {shortage.week} ({(shortage.balance/1000).toFixed(1)} MT)
        </div>
      )}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:12 }}>
        <Stat label="Current stock"
          value={`${(crude.latestTankKg/1000).toFixed(1)} MT`}
          color={crude.latestTankKg < CRUDE_CRITICAL_KG ? S.red :
                 crude.latestTankKg < CRUDE_LOW_KG ? S.yellow : S.green} />
        <Stat label="As of" value={fmtDate(crude.tankReadingAt)} color={S.muted} />
      </div>
      <div style={{ display:"flex", gap:6, alignItems:"flex-end", height:80 }}>
        {projected.map(w => {
          const h = Math.max((w.balance/maxKg)*72, 2);
          const col = wkColor(w.balance);
          return (
            <div key={w.week} style={{ flex:1, display:"flex", flexDirection:"column",
              alignItems:"center", gap:4 }}>
              <div style={{ fontFamily:mono, fontSize:9, color:col }}>
                {w.balance >= 1000 ? `${(w.balance/1000).toFixed(0)}t` : `${w.balance}kg`}
              </div>
              <div style={{ width:"100%", height:h, background:col,
                borderRadius:"3px 3px 0 0", opacity:0.85 }} />
              <div style={{ fontFamily:mono, fontSize:9, color:S.muted }}>W{w.week}</div>
            </div>
          );
        })}
      </div>
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7: Sales Pipeline
// ─────────────────────────────────────────────────────────────────────────────
function SalesSection({ sales }) {
  const today = isoDay(new Date());

  const weeks = useMemo(() => {
    const grouped = {};
    sales.forEach(o => {
      const wk = weekKey(isoDay(o.requested_crd));
      if (!grouped[wk]) grouped[wk] = { orders:[], totalBags:0 };
      grouped[wk].orders.push(o);
      grouped[wk].totalBags += kgToBags(o.quantity_kg||0);
    });
    // Build 10 weeks starting from this week's Monday
    const todayDate = new Date();
    const dow = todayDate.getDay();
    const thisMon = addDays(todayDate, dow===0?-6:1-dow);
    const result = [];
    for (let w=0; w<10; w++) {
      const mon = isoDay(addDays(thisMon, w*7));
      const fri = isoDay(addDays(thisMon, w*7+4));
      result.push({ week: mon, label:`${fmtDate(mon)}–${fmtDate(fri)}`,
        ...(grouped[mon]||{orders:[],totalBags:0}) });
    }
    return result;
  }, [sales, today]);

  const maxBags = Math.max(...weeks.map(w=>w.totalBags), 1);

  return (
    <Card title="Sales Pipeline" subtitle="next 10 weeks" accent={S.orange}>
      {weeks.map(w => {
        const hasBags = w.totalBags > 0;
        const pct = (w.totalBags/maxBags)*100;
        const customers = [...new Set(w.orders.map(o=>o.company_name||o.customer_id))];
        return (
          <div key={w.week} style={{ marginBottom:10 }}>
            <div style={{ display:"flex", justifyContent:"space-between",
              alignItems:"baseline", marginBottom:4 }}>
              <div style={{ fontFamily:mono, fontSize:11, color:hasBags?S.text:S.muted }}>
                {w.label}
              </div>
              <div style={{ fontFamily:mono, fontSize:12, fontWeight:700,
                color:hasBags?S.orange:S.muted }}>
                {hasBags ? `${w.totalBags} bags` : "—"}
              </div>
            </div>
            <div style={{ background:"#04080f", borderRadius:4, height:8, marginBottom:hasBags?4:0 }}>
              <div style={{ width:`${pct}%`, height:"100%",
                background:hasBags?S.orange:"#1a1a1a", borderRadius:4 }} />
            </div>
            {hasBags && (
              <div style={{ fontFamily:sans, fontSize:11, color:S.muted }}>
                {customers.join(" · ")} ({w.orders.length} order{w.orders.length!==1?"s":""})
              </div>
            )}
          </div>
        );
      })}
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 8: Batch Viewer
// ─────────────────────────────────────────────────────────────────────────────
function BatchViewerSection() {
  const [batchList, setBatchList] = useState([]);
  const [selected, setSelected]   = useState("");
  const [details, setDetails]     = useState(null);
  const [loading, setLoading]     = useState(false);
  const [listError, setListError] = useState(null);

  useEffect(() => {
    apiGet("getBatchList")
      .then(setBatchList)
      .catch(e => setListError(e.message));
  }, []);

  useEffect(() => {
    if (!selected) { setDetails(null); return; }
    setLoading(true);
    apiGet("getBatchDetails", { code: selected })
      .then(setDetails)
      .catch(() => setDetails(null))
      .finally(() => setLoading(false));
  }, [selected]);

  function fmtTs(ts) {
    if (!ts) return "—";
    return new Date(ts).toLocaleString("en-GB", { day:"2-digit", month:"short", hour:"2-digit", minute:"2-digit" });
  }

  const root = details?.root;
  const entries = details?.entries || [];
  const totalBags = entries.reduce((s, e) => s + (e.NumberOfBags||0), 0);
  const totalCsl  = entries.reduce((s, e) => s + (e.csl_kg||0), 0) / 1000;

  return (
    <Card title="Batch Viewer" subtitle="select a root batch to inspect" accent={S.blue}>
      {listError && (
        <div style={{ fontFamily:mono, fontSize:12, color:S.red, marginBottom:10 }}>
          ⚠ Could not load batch list: {listError}
        </div>
      )}

      <select
        value={selected}
        onChange={e => setSelected(e.target.value)}
        style={{ width:"100%", background:"#04080f", border:`1px solid ${S.border}`,
          color: selected ? S.text : S.muted, borderRadius:8, padding:"10px 12px",
          fontFamily:mono, fontSize:13, marginBottom:14, cursor:"pointer" }}>
        <option value="">— Select batch —</option>
        {batchList.map(b => (
          <option key={b.RootBatchCode} value={b.RootBatchCode}>
            {b.RootBatchCode}
            {b.first_date ? `  ·  ${fmtDate(b.first_date)}` : "  ·  no production"}
            {b.total_bags ? `  ·  ${b.total_bags} bags` : ""}
          </option>
        ))}
      </select>

      {loading && (
        <div style={{ fontFamily:mono, fontSize:12, color:S.muted, textAlign:"center", padding:20 }}>
          Loading…
        </div>
      )}

      {!loading && root && (
        <>
          {/* Header summary */}
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:14 }}>
            <Stat label="Allocated" value={`${root.AllocatedQty_MT ?? "—"} MT`} color={S.blue} />
            <Stat label="Status" value={root.IsActive ? "Active" : "Closed"}
              color={root.IsActive ? S.green : S.muted} />
            <Stat label="Total Bags" value={totalBags || "—"} color={S.orange} />
            <Stat label="CSL Processed" value={totalCsl > 0 ? `${totalCsl.toFixed(2)} MT` : "—"} color={S.purple} />
          </div>
          <div style={{ fontFamily:mono, fontSize:10, color:S.muted, marginBottom:14 }}>
            Allocated by {root.CreatedByName || "—"} · Created {fmtDate(root.CreatedAt)}
            {root.Year ? `  ·  Year ${root.Year}  ·  Block ${root.Block}  ·  Receipt ${root.ReceiptNo}  ·  Material ${root.MaterialType}` : ""}
          </div>

          {/* Production entries table */}
          {entries.length === 0 ? (
            <div style={{ fontFamily:mono, fontSize:12, color:S.muted }}>No submitted production entries.</div>
          ) : (
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontFamily:mono, fontSize:11 }}>
                <thead>
                  <tr style={{ borderBottom:`1px solid ${S.border}` }}>
                    {["Date","Batch No.","Decanter","Operator","Start","Stop","Bags","CSL (kg)"].map(h => (
                      <th key={h} style={{ padding:"6px 8px", color:S.muted, fontWeight:600,
                        textAlign:"left", whiteSpace:"nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {entries.map((e, i) => (
                    <tr key={e.EntryID ?? i}
                      style={{ borderBottom:`1px solid ${S.border}22`,
                        background: i % 2 === 0 ? "transparent" : "#04080f" }}>
                      <td style={{ padding:"6px 8px", color:S.text, whiteSpace:"nowrap" }}>
                        {e.production_date ? fmtDate(e.production_date) : "—"}
                      </td>
                      <td style={{ padding:"6px 8px", color:S.orange, whiteSpace:"nowrap" }}>
                        {e.FullBatchNumber || e.LotNumber || "—"}
                      </td>
                      <td style={{ padding:"6px 8px", color:S.blue }}>{e.Decanter || "—"}</td>
                      <td style={{ padding:"6px 8px", color:S.muted }}>{e.OperatorName || "—"}</td>
                      <td style={{ padding:"6px 8px", color:S.muted, whiteSpace:"nowrap" }}>
                        {fmtTs(e.DecanterStartTS)}
                      </td>
                      <td style={{ padding:"6px 8px", color:S.muted, whiteSpace:"nowrap" }}>
                        {fmtTs(e.DecanterStopTS)}
                      </td>
                      <td style={{ padding:"6px 8px", color:S.text, textAlign:"right" }}>
                        {e.NumberOfBags ?? "—"}
                      </td>
                      <td style={{ padding:"6px 8px", color:S.green, textAlign:"right" }}>
                        {e.csl_kg != null ? Math.round(e.csl_kg).toLocaleString() : "—"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      )}

      {!loading && selected && !root && (
        <div style={{ fontFamily:mono, fontSize:12, color:S.muted }}>No data found for this batch.</div>
      )}
    </Card>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const account = accounts[0];
  const roles = account?.idTokenClaims?.roles ?? [];
  const allowed = roles.includes("GS.Management") || roles.includes("GS.PlantManager");

  const [data, setData]       = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState(null);
  const [lastRefresh, setLastRefresh] = useState(null);

  async function load() {
    setError(null);
    try {
      const d = await apiGet("getManagementDashboard");
      setData(d);
      setLastRefresh(new Date());
    } catch(e) { setError(e.message); }
    finally { setLoading(false); }
  }

  useEffect(() => {
    if (isAuthenticated && allowed) {
      load();
      const interval = setInterval(load, 5*60*1000);
      return () => clearInterval(interval);
    }
    if (isAuthenticated) setLoading(false);
  }, [isAuthenticated, allowed]);

  function handleLogin()  { instance.loginPopup(loginRequest).catch(e=>console.error(e)); }
  function handleLogout() { instance.logoutPopup().catch(e=>console.error(e)); }

  if (!isAuthenticated) {
    return (
      <div style={{ minHeight:"100vh", background:S.bg, display:"flex", flexDirection:"column",
        alignItems:"center", justifyContent:"center", padding:24,
        fontFamily:sans, color:S.text }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Barlow:wght@400;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');`}</style>
        <div style={{ color:S.orange, fontSize:11, letterSpacing:4, fontWeight:700, marginBottom:8,
          fontFamily:mono }}>GIIAVA</div>
        <div style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:30, fontWeight:800,
          color:S.text, textAlign:"center", lineHeight:1.1, marginBottom:6 }}>
          Management<br/>Dashboard
        </div>
        <div style={{ color:S.muted, marginBottom:36, fontSize:14 }}>
          Sign in with your Microsoft account
        </div>
        <button onClick={handleLogin} style={{ background:"#2563eb", color:"#fff", border:"none",
          borderRadius:12, padding:"16px 40px", fontSize:16, fontWeight:700, cursor:"pointer",
          fontFamily:sans }}>
          Sign in with Microsoft
        </button>
      </div>
    );
  }

  if (!allowed) {
    return (
      <div style={{ minHeight:"100vh", background:S.bg, display:"flex", flexDirection:"column",
        alignItems:"center", justifyContent:"center", padding:24, fontFamily:sans }}>
        <div style={{ color:S.red, fontSize:18, fontWeight:700 }}>⚠ Access Denied</div>
        <div style={{ color:S.muted, fontSize:14, marginTop:8, textAlign:"center" }}>
          This dashboard requires the GS.Management or GS.PlantManager role.
        </div>
        <button onClick={handleLogout} style={{ marginTop:24, background:"#1f2937", color:S.muted,
          border:"none", borderRadius:8, padding:"10px 24px", cursor:"pointer", fontSize:14 }}>
          Sign Out
        </button>
      </div>
    );
  }

  return (
    <div style={{ maxWidth:700, margin:"0 auto", minHeight:"100vh", background:S.bg,
      color:S.text, fontFamily:sans, padding:"0 0 40px" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Barlow:wght@400;600;700;800&family=Barlow+Condensed:wght@600;700;800&display=swap');*{box-sizing:border-box}input,select{outline:none}`}</style>

      {/* Header */}
      <div style={{ background:"#0d0d0d", borderBottom:`1px solid #1f2937`,
        padding:"12px 16px", position:"sticky", top:0, zIndex:10 }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
          <div>
            <div style={{ fontFamily:mono, fontSize:9, color:S.orange, letterSpacing:4,
              fontWeight:700 }}>GIIAVA</div>
            <div style={{ fontFamily:"'Barlow Condensed',sans-serif", fontSize:16, fontWeight:800,
              color:S.text, lineHeight:1 }}>Management Dashboard</div>
          </div>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            {lastRefresh && (
              <div style={{ fontFamily:mono, fontSize:9, color:S.muted }}>
                {lastRefresh.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}
              </div>
            )}
            <button onClick={load} style={{ background:"#1f2937", border:"none", color:S.muted,
              padding:"5px 10px", borderRadius:6, cursor:"pointer", fontSize:11 }}>↻</button>
            <button onClick={handleLogout} style={{ background:"#1f2937", border:"none",
              color:S.muted, padding:"5px 10px", borderRadius:6, cursor:"pointer", fontSize:11 }}>
              Sign Out
            </button>
          </div>
        </div>
      </div>

      <div style={{ padding:"16px 12px 0" }}>
        {loading && (
          <div style={{ textAlign:"center", padding:60, fontFamily:mono,
            fontSize:13, color:S.muted }}>Loading dashboard…</div>
        )}
        {error && (
          <div style={{ background:"#1f0000", border:`1px solid ${S.red}`, borderRadius:8,
            padding:"12px 16px", fontFamily:mono, fontSize:13, color:S.red, marginBottom:16 }}>
            ⚠ {error}
            <button onClick={load} style={{ marginLeft:12, background:"none", border:"none",
              color:S.blue, cursor:"pointer", fontSize:12 }}>Retry</button>
          </div>
        )}
        {data && !loading && (
          <>
            <ProductionSection production={data.production} />
            <UtilitiesSection utilities={data.utilities} production={data.production}
              acetoneTanks={data.acetoneTanks} />
            <DecanterSection decanter={data.decanter} />
            <PackagingEfficiencySection packagingOut={data.packagingOut}
              attendance={data.attendance} />
            <SurplusSection production={data.production} packagingOut={data.packagingOut} />
            <CrudeSection crude={data.crude} />
            <SalesSection sales={data.sales} />
            <BatchViewerSection />
          </>
        )}
      </div>
    </div>
  );
}
