// Tender Cost — ชิ้นส่วนหน้าจอที่ใช้ร่วมกัน (แถบหัว, ปุ่ม, การ์ด, กล่องข้อความ, เครื่องคิดเลข, ช่องกรอกเงิน/วันที่)
import { Component, useEffect, useLayoutEffect, useRef, useState } from "react";
import { ROLE_LABELS } from "./auth.js";
import { BNAV_H, BOTTOM, DialogStore, FAB_GAP, FAB_SIZE, GROUPS, PO_STATUS, STATUS_BG, STATUS_CLR, T, _LANG, accountOf, bnavH, fmt, fmt0, poStatusLabel, t, toggleLang, uiLocale, useLang } from "./core.jsx";

// กันจอขาว: ถ้าหน้าจอส่วนใดโยน error ตอน render จะโชว์กล่องแจ้ง + ปุ่มลองใหม่
// แทนที่จะพังทั้งแอพ
class ErrorBoundary extends Component {
  constructor(props){ super(props); this.state = { err:null }; }
  static getDerivedStateFromError(err){ return { err }; }
  componentDidCatch(err, info){ console.error("UI error:", err, info); }
  render(){
    if (this.state.err) {
      return (
        <div style={{minHeight:"60vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14,padding:24,textAlign:"center"}}>
          <div style={{color:T.amber}}><Ico name="alert" size={40} sw={1.6} /></div>
          <div style={{fontSize:16,fontWeight:650,color:"#0f172a"}}>{t("เกิดข้อผิดพลาดในการแสดงผลหน้านี้","Something went wrong displaying this page")}</div>
          <div style={{fontSize:13,color:"#64748b",maxWidth:460}}>{t("ข้อมูลของคุณยังปลอดภัย ลองกดปุ่มด้านล่างเพื่อโหลดใหม่ ถ้ายังเป็นอยู่ให้แจ้งผู้ดูแลระบบ","Your data is safe. Press the button below to reload; if it keeps happening, contact your admin")}</div>
          <button onClick={()=>{ this.setState({err:null}); if(typeof window!=="undefined") window.location.reload(); }}
            style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:10,padding:"9px 20px",fontSize:14,fontWeight:600,cursor:"pointer"}}>{t("โหลดหน้าใหม่","Reload page")}</button>
        </div>
      );
    }
    return this.props.children;
  }
}

// ─── Loader ───────────────────────────────────────────────────────────────────
function Loader() {
  return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:T.bg}}>
      <div style={{textAlign:"center"}}>
        <div style={{width:40,height:40,border:`3px solid ${T.blueMid}`,borderTopColor:T.blue,borderRadius:"50%",animation:"spin 0.7s linear infinite",margin:"0 auto 14px"}}/>
        <div style={{fontSize:13,color:T.textSecondary}}>{t("กำลังโหลด...","Loading...")}</div>
      </div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

// ─── SyncBadge ────────────────────────────────────────────────────────────────
function SyncBadge({ syncing, syncedAt, light = false }) {
  return (
    <div className="sync-badge" title={syncedAt ? `${t("บันทึก/ซิงก์ล่าสุด","Last saved / synced")} ${syncedAt.toLocaleTimeString(uiLocale())}` : undefined}
      style={light ? {display:"flex",alignItems:"center",gap:6,fontSize:12,color:T.textMuted,whiteSpace:"nowrap"}
                   : {display:"flex",alignItems:"center",gap:6,background:"rgba(255,255,255,0.15)",backdropFilter:"blur(8px)",borderRadius:8,padding:"5px 12px",fontSize:11,color:"rgba(255,255,255,0.85)"}}>
      <span style={{width:6,height:6,borderRadius:"50%",background:syncing?"#fbbf24":"#34d399",display:"inline-block",boxShadow:syncing?"0 0 6px #fbbf24":"0 0 6px #34d399",animation:syncing?"pulse 0.8s ease-in-out infinite":"none"}}/>
      {syncing ? t("กำลัง sync...","Syncing...") : syncedAt ? `sync ${syncedAt.toLocaleTimeString(uiLocale(),{hour:"2-digit",minute:"2-digit",second:"2-digit"})}` : ""}
    </div>
  );
}

// ─── Search input with a clear (×) button ──────────────────────────────────
// Small wrapper around the standard .input-base search box used across QS,
// Procurement, and Accounting toolbars — shows an × to instantly clear the
// text once something has been typed, instead of having to select-and-delete.
function SearchInput({ value, onChange, placeholder, width = 240, big = false }) {
  // big = เด่นขึ้น (กรอบชัด + เงา) แต่ "ขนาดเท่าเดิม"
  const bigStyle = big ? {
    border:`2px solid ${value?T.blue:"#94a3b8"}`, borderRadius:10, background:"#fff",
    boxShadow:"0 1px 4px rgba(15,23,42,0.07)",
  } : {};
  return (
    <div style={{position:"relative",width}}>
      <Ico name="search" size={16} color={T.textMuted} style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",pointerEvents:"none"}} />
      <input value={value} onChange={e=>onChange(e.target.value)} placeholder={String(placeholder||"").replace(/^\u{1F50D}\s*/u,"")}
        className="input-base" style={{width:"100%",paddingLeft:34,paddingRight:value?30:13,...bigStyle}}
        onFocus={big?(e=>{e.currentTarget.style.borderColor=T.blue;e.currentTarget.style.boxShadow="0 0 0 3px rgba(37,99,235,0.15)";}):undefined}
        onBlur={big?(e=>{e.currentTarget.style.borderColor=value?T.blue:"#94a3b8";e.currentTarget.style.boxShadow="0 1px 4px rgba(15,23,42,0.07)";}):undefined}/>
      {value && (
        <button type="button" onClick={()=>onChange("")} title={t("ล้างคำค้นหา","Clear search")}
          style={{position:"absolute",right:6,top:"50%",transform:"translateY(-50%)",width:20,height:20,border:"none",
            borderRadius:"50%",background:"transparent",color:T.textMuted,fontSize:15,lineHeight:1,cursor:"pointer",
            display:"flex",alignItems:"center",justifyContent:"center",padding:0}}
          onMouseEnter={e=>{e.currentTarget.style.background="#e2e8f0";e.currentTarget.style.color=T.textPrimary;}}
          onMouseLeave={e=>{e.currentTarget.style.background="transparent";e.currentTarget.style.color=T.textMuted;}}>
          ×
        </button>
      )}
    </div>
  );
}

// ─── Group filter (multi-select dropdown) ─────────────────────────────────────
// แทนแถวชิปหมวดยาว ๆ ที่รก — เป็นปุ่มเดียวเปิด dropdown ติ๊กเลือกได้หลายหมวด
// selected = อาเรย์ของหมวดที่เลือก (ว่าง = ทุกหมวด)
function GroupFilter({ selected, onChange, options = GROUPS, color = T.blue }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [open]);
  const toggle = (g) => onChange(selected.includes(g) ? selected.filter(x => x !== g) : [...selected, g]);
  const has = selected.length > 0;
  const label = !has ? t("ทุกหมวด","All groups") : selected.length === 1 ? selected[0] : `${selected.length} ${t("หมวด","groups")}`;
  const rowStyle = (on) => ({ display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",border:"none",
    background: on ? T.blueLight : "transparent", color: on ? color : T.textSecondary, cursor:"pointer",
    padding:"7px 10px", borderRadius:8, fontSize:12, fontWeight: on ? 650 : 500 });
  return (
    <div ref={ref} style={{position:"relative",flexShrink:0}}>
      <button onClick={()=>setOpen(o=>!o)} title={t("กรองตามหมวด (เลือกได้หลายหมวด)","Filter by group (multi-select)")}
        style={{display:"flex",alignItems:"center",gap:7,padding:"6px 12px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap",
          border:`1.5px solid ${has?color:T.cardBorder}`, background: has?color:T.card, color: has?"#fff":T.textSecondary}}>
        <Ico name="tag" size={14} /> {label}
        {has && <span style={{fontSize:11,opacity:0.85}}>({selected.length})</span>}
        <span style={{fontSize:9,opacity:0.8}}>▼</span>
      </button>
      {open && (
        <div style={{position:"absolute",top:"calc(100% + 6px)",left:0,zIndex:60,background:T.card,border:`1px solid ${T.cardBorder}`,
          borderRadius:12,boxShadow:"0 10px 32px rgba(15,23,42,0.18)",padding:6,minWidth:210,maxHeight:340,overflowY:"auto"}}>
          <button onClick={()=>{ onChange([]); }} style={rowStyle(!has)}>
            <span style={{fontSize:13}}>{!has?"◉":"◯"}</span> {t("ทุกหมวด","All groups")}
          </button>
          <div style={{height:1,background:T.cardBorder,margin:"4px 2px"}}/>
          {options.map(g => {
            const on = selected.includes(g);
            return (
              <button key={g} onClick={()=>toggle(g)} style={rowStyle(on)}>
                <span style={{fontSize:13}}>{on?"☑":"☐"}</span> {g}
              </button>
            );
          })}
          {has && (
            <>
              <div style={{height:1,background:T.cardBorder,margin:"4px 2px"}}/>
              <button onClick={()=>onChange([])} style={{...rowStyle(false),color:T.red,justifyContent:"center",fontWeight:600}}>
                ✕ {t("ล้างตัวเลือก","Clear selection")}
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
}

// ─── StatCard ─────────────────────────────────────────────────────────────────
// ชุดไอคอนเส้น (แทนอีโมจิ ซึ่งหน้าตาต่างกันแต่ละเครื่อง/ระบบ)
const ICON_PATHS = {
  download: "M12 4v11M7 10l5 5 5-5M5 20h14",
  edit: "M4 20h4L19 9l-4-4L4 16v4zM13.5 6.5l4 4",
  home: "M4 11l8-7 8 7v9h-5v-6H9v6H4z",
  grid: "M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z",
  eyeoff: "M3 3l18 18M10.6 6.1A9.8 9.8 0 0 1 12 6c5 0 9 6 9 6a15 15 0 0 1-3.2 3.7M6.6 6.6C4.3 8.1 3 12 3 12s4 6 9 6c1.6 0 3-.4 4.3-1",
  ruler: "M4 17L17 4l3 3L7 20zM8 13l2 2M11 10l2 2M14 7l2 2",
  gear: "M12 9a3 3 0 1 0 0 6 3 3 0 0 0 0-6zM12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1",
  plus: "M12 5v14M5 12h14",
  check: "M20 12a8 8 0 1 1-16 0 8 8 0 0 1 16 0zM8.5 12.5l2.5 2.5 5-5",
  clipboard: "M9 3h6v3H9zM7 4.5H5V21h14V4.5h-2M8 11h8M8 15h5",
  box: "M3 7l9-4 9 4v10l-9 4-9-4zM3 7l9 4 9-4M12 11v10",
  calc: "M6 3h12v18H6zM9 7h6M9 12h.01M12 12h.01M15 12h.01M9 16h.01M12 16h.01M15 16h.01",
  receipt: "M6 3h12v18l-3-2-3 2-3-2-3 2zM9 8h6M9 12h6",
  wallet: "M3 7h15a3 3 0 0 1 3 3v7a3 3 0 0 1-3 3H3zM3 7l12-3v3M16 14h2",
  chart: "M4 20V11M10 20V4M16 20v-8M3 20h18",
  chevrons: "M7 10l5 5 5-5",
  alert: "M12 4l9 16H3zM12 10v4M12 17h.01",
  calendar: "M4 6h16v14H4zM4 10h16M8 3v4M16 3v4",
  trash: "M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13M10 11v6M14 11v6",
  lock: "M6 11h12v9H6zM8 11V8a4 4 0 0 1 8 0v3",
  search: "M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM20 20l-4.2-4.2",
  copy: "M8 8h11v12H8zM5 16V4h11",
  clock: "M12 3a9 9 0 1 0 0 18 9 9 0 0 0 0-18zM12 7v5l3 2",
  truck: "M3 6h11v10H3zM14 10h4l3 3v3h-7M7 19a2 2 0 1 0 0-4 2 2 0 0 0 0 4zM17 19a2 2 0 1 0 0-4 2 2 0 0 0 0 4z",
  tag: "M3 12V4h8l10 10-8 8zM7.5 7.5h.01",
  filter: "M4 5h16l-6 8v6l-4-2v-4z",
  users: "M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8zM2 21v-1a6 6 0 0 1 12 0v1M16 3.5a4 4 0 0 1 0 7.5M22 21v-1a6 6 0 0 0-4-5.6",
  trend: "M3 17l6-6 4 4 8-8M15 7h6v6",
};
function Ico({ name, size = 16, color = "currentColor", sw = 1.9, style }) {
  const d = ICON_PATHS[name]; if (!d) return null;
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden="true" style={{ flexShrink: 0, verticalAlign: "-3px", ...(style || {}) }}>
    <path d={d} stroke={color} strokeWidth={sw} strokeLinecap="round" strokeLinejoin="round" />
  </svg>;
}
const EMOJI_ICON = { "📐":"ruler", "⚙️":"gear", "➕":"plus", "✅":"check", "📋":"clipboard", "📦":"box", "🧮":"calc", "🧾":"receipt", "💰":"wallet", "📊":"chart", "⚠️":"alert" };
function StatCard({ label, value, sub, color, icon, accent, thb, rate, progress = null, lead = false }) {
  // ถ้าใส่ยอดบาท (thb) + อัตราแลกเปลี่ยน (rate = บาท/USD) จะโชว์ ≈ $ ควบคู่ให้
  const usd = (rate && rate > 0 && typeof thb === "number") ? thb / rate : null;
  // เว้นช่องเล็ก ๆ ระหว่าง ฿ กับตัวเลข กันสัญลักษณ์ ฿ ทับหลักแรก (ฟอนต์ mono บางตัว ฿ ยื่น)
  const shownValue = (typeof value === "string" && value.startsWith("฿"))
    ? <><span style={{marginRight:3}}>฿</span>{value.slice(1)}</> : value;
  return (
    <div className={`stat-card${lead ? " stat-lead" : ""}`} style={{background:T.card,borderRadius:14,padding:"20px 22px",border:`1px solid ${T.cardBorder}`,position:"relative",overflow:"hidden"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
        <div style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{label}</div>
        {icon && <div className="stat-icon" style={{width:34,height:34,borderRadius:10,background:accent||T.blueLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:color||T.blue}}>{EMOJI_ICON[icon] ? <Ico name={EMOJI_ICON[icon]} size={18} /> : icon}</div>}
      </div>
      <div className="stat-val" style={{fontSize:lead?28:22,fontWeight:lead?700:650,color:T.textPrimary,letterSpacing:"-0.5px",fontVariantNumeric:"tabular-nums"}}>{shownValue}</div>
      {progress != null && (
        <div role="progressbar" aria-valuemin={0} aria-valuemax={100} aria-valuenow={Math.round(Math.min(100, Math.max(0, progress)))}
          style={{height:6,borderRadius:3,background:"#e9edf3",marginTop:10,overflow:"hidden"}}>
          <div style={{height:"100%",width:`${Math.min(100, Math.max(0, progress))}%`,background:progress>100?T.red:progress>=80?T.amber:(color||T.blue),borderRadius:3}}/>
        </div>
      )}
      {usd != null && <div style={{fontSize:14,color:T.green,fontWeight:650,fontVariantNumeric:"tabular-nums",marginTop:3}}>≈ ${fmt0(usd)}</div>}
      {sub && <div style={{fontSize:11,color:T.textMuted,marginTop:5}}>{sub}</div>}
    </div>
  );
}

// แสดงบรรทัดเป็นดอลลาร์ ($) ใต้ยอดบาทในตาราง — ขนาดราวครึ่งหนึ่งของบาท, ทศนิยม 2 ตำแหน่ง
// คืน null ถ้าไม่ได้เปิดใช้อัตราแลกเปลี่ยน
function usdLine(thb, rate) {
  if (!rate || rate <= 0 || typeof thb !== "number" || !isFinite(thb)) return null;
  return <div className="usd-sub" style={{fontSize:11,color:T.green,fontWeight:600,fontVariantNumeric:"tabular-nums",lineHeight:1.2,marginTop:2}}>≈ ${fmt(thb/rate)}</div>;
}

// อัตราแลกเปลี่ยนของโปรเจกต์ที่ควรใช้แสดงผล (0 = ปิด/ไม่แสดง $)
function effRate(project) {
  return (project?.showUsd !== false) ? (parseFloat(project?.usdRate) || 0) : 0;
}

// ตัวควบคุมค่าเงิน: สลับเปิด-ปิดการแสดง $ + แก้ไขอัตราแลกเปลี่ยนได้ (วางข้างปุ่ม Export)
function CurrencyControl({ project, updateProject }) {
  const on = project?.showUsd !== false;   // ค่าเริ่มต้น: เปิด
  const [txt, setTxt] = useState(project?.usdRate ?? "");
  useEffect(() => { setTxt(project?.usdRate ?? ""); }, [project?.usdRate]);
  const commitRate = () => {
    const v = String(txt).trim();
    if (v !== String(project?.usdRate ?? "")) updateProject({ usdRate: v });
  };
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,padding:"2px 10px",border:`1px solid ${T.cardBorder}`,borderRadius:10,background:T.card}}>
      <button onClick={()=>updateProject({ showUsd: !on })} title={t("เปิด/ปิดการแสดงเป็นดอลลาร์ ($)","Show/hide US dollar ($)")}
        style={{display:"flex",alignItems:"center",gap:6,border:"none",background:"transparent",cursor:"pointer",padding:"6px 2px",minHeight:32}}>
        <span style={{width:34,height:18,borderRadius:99,background:on?T.green:"#cbd5e1",position:"relative",transition:"all .15s",display:"inline-block",flexShrink:0}}>
          <span style={{position:"absolute",top:2,left:on?18:2,width:14,height:14,borderRadius:99,background:"#fff",transition:"all .15s"}}/>
        </span>
        <span style={{fontSize:12,fontWeight:650,color:on?T.green:T.textMuted}}>USD</span>
      </button>
      <span style={{fontSize:11,color:T.textMuted,whiteSpace:"nowrap"}}>฿/$</span>
      <input type="number" step="any" min="0" value={txt} placeholder={t("อัตรา","Rate")}
        onChange={e=>setTxt(e.target.value)} onBlur={commitRate}
        onKeyDown={e=>{ if(e.key==="Enter") e.currentTarget.blur(); }}
        style={{width:70,fontSize:13,padding:"6px 8px",minHeight:32,border:`1px solid ${T.cardBorder}`,borderRadius:7,fontVariantNumeric:"tabular-nums",textAlign:"right"}}/>
    </div>
  );
}

// ─── ปุ่มสลับภาษา ไทย / EN ────────────────────────────────────────────────────
// ─── เครื่องคิดเลข (ป๊อปอัพลอย ลากย้ายได้ ไม่บังการทำงาน) ─────────────────────────
// คำนวณแบบปลอดภัย (ไม่ใช้ eval): + − × ÷ วงเล็บ ทศนิยม และ % แบบเครื่องคิดเลข
//   100+7% = 107 · 100−10% = 90 · 200×15% = 30 · 50% = 0.5
const calcEval = (src) => {
  const s = String(src ?? "").replace(/,/g, "").replace(/×/g, "*").replace(/÷/g, "/").replace(/[−–]/g, "-").replace(/\s+/g, "");
  if (!s) return { ok: false, empty: true };
  let i = 0;
  const peek = () => s[i];
  const num = () => {
    const m = /^(\d+\.?\d*|\.\d+)/.exec(s.slice(i));
    if (!m) throw new Error("num");
    i += m[0].length; return parseFloat(m[0]);
  };
  // factor → { v, pct }  (pct = ค่านี้เป็น "เปอร์เซ็นต์" ล้วน ๆ ไว้คิดแบบ a ± b%)
  const factor = () => {
    const c = peek();
    if (c === "-") { i++; const f = factor(); return { v: -f.v, pct: f.pct }; }
    if (c === "+") { i++; return factor(); }
    let v;
    if (c === "(") { i++; v = expr().v; if (peek() !== ")") throw new Error("paren"); i++; }
    else v = num();
    if (peek() === "%") { i++; return { v: v / 100, pct: true }; }
    return { v, pct: false };
  };
  const term = () => {
    let f = factor(), v = f.v, pct = f.pct;
    while (peek() === "*" || peek() === "/") {
      const op = s[i++]; const r = factor(); pct = false;
      if (op === "*") v *= r.v; else { if (r.v === 0) throw new Error("div0"); v /= r.v; }
    }
    return { v, pct };
  };
  const expr = () => {
    let { v } = term();
    while (peek() === "+" || peek() === "-") {
      const op = s[i++]; const r = term();
      const d = r.pct ? v * r.v : r.v;               // 100+7% → 100 + 100×7%
      v = op === "+" ? v + d : v - d;
    }
    return { v };
  };
  try {
    const { v } = expr();
    if (i !== s.length || !isFinite(v)) return { ok: false };
    return { ok: true, v: Math.round(v * 1e8) / 1e8 };
  } catch (e) { return { ok: false, div0: e.message === "div0" }; }
};
const fmtCalc = (v) => Number(v).toLocaleString("en-US", { maximumFractionDigits: 4 });
// ตัวเลขเป็นข้อความธรรมดา (ไม่มี , และไม่เป็น 1e-8 / 1e+21) — ใช้ใส่กลับเข้าสูตร/คัดลอก
const plainNum = (v, dp = 8) => { const r = Math.round(v * 10 ** dp) / 10 ** dp; return Number.isInteger(r) ? String(r) : r.toFixed(dp).replace(/\.?0+$/, ""); };
const CALC_MAX = 1e15;   // เกินนี้ความละเอียดของตัวเลขไม่พอ
// แสดงสูตรให้อ่านง่าย: ใส่ , หลักพัน · × ÷ − แทน * / -
const prettyExpr = (s) => String(s).replace(/\d+(\.\d*)?/g, (m) => { const [a, b] = m.split("."); return Number(a).toLocaleString("en-US") + (b !== undefined ? "." + b : ""); })
  .replace(/\*/g, " × ").replace(/\//g, " ÷ ").replace(/([\d%)])-/g, "$1 − ").replace(/\+/g, " + ");

const CalcStore = { open: false, subs: new Set(), set(v) { this.open = v; this.subs.forEach(f => f(v)); } };
const useCalcOpen = () => {
  const [o, setO] = useState(CalcStore.open);
  useEffect(() => { CalcStore.subs.add(setO); return () => { CalcStore.subs.delete(setO); }; }, []);
  return o;
};
// ปุ่มลอยมุมขวาล่าง: เครื่องคิดเลข (ล่างสุด) + กลับไปด้านบน (อยู่เหนือ แสดงเมื่อเลื่อนลงมาแล้ว)
// z-index ต่ำกว่าหน้าต่างป๊อปอัพ (100+) จึงไม่บังปุ่มบันทึกของฟอร์ม PO
const fabStyle = (active) => ({ position:"fixed", right:FAB_GAP, width:FAB_SIZE, height:FAB_SIZE, borderRadius:16, border:"none", cursor:"pointer",
  background: active ? "#2563eb" : "#1e3a8a", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", zIndex:95,
  boxShadow:"0 10px 24px rgba(30,58,138,0.35), 0 2px 6px rgba(15,23,42,0.18)", transition:"background .15s, transform .12s, opacity .2s" });
function CalcFab() {
  useLang();
  const open = useCalcOpen();
  return (
    <button onClick={() => CalcStore.set(!open)} title={t("เครื่องคิดเลข","Calculator")} aria-label={t("เครื่องคิดเลข","Calculator")} aria-pressed={open} data-calc-toggle
      className="fab-btn" style={{ ...fabStyle(open), bottom:BOTTOM(FAB_GAP) }}>
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <rect x="4.5" y="2.5" width="15" height="19" rx="2.5" stroke="#fff" strokeWidth="1.8"/>
        <rect x="7.5" y="5.5" width="9" height="3.5" rx="0.8" fill="#fff"/>
        {[0,1,2].flatMap(r => [0,1,2].map(c => <circle key={`${r}${c}`} cx={8.5 + c*3.5} cy={12.5 + r*3} r="1.05" fill="#fff"/>))}
      </svg>
    </button>
  );
}
// ปุ่ม ↑ มุมขวาล่าง = กลับบนสุดของ "หน้าเพจ" เท่านั้น (ตารางมีปุ่มของตัวเอง — TableTopButton)
function ScrollTopFab() {
  useLang();
  const [show, setShow] = useState(false);
  useEffect(() => {
    const onScroll = () => setShow(window.scrollY > 300);
    onScroll();
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);
  const toTop = () => { try { window.scrollTo({ top: 0, behavior: "smooth" }); } catch { window.scrollTo(0, 0); } };
  return (
    <button onClick={toTop} title={t("กลับบนสุดของหน้า","Back to top of page")} aria-label={t("กลับบนสุดของหน้า","Back to top of page")} data-scroll-top tabIndex={show ? 0 : -1}
      className="fab-btn" style={{ ...fabStyle(false), bottom:BOTTOM(FAB_GAP + FAB_SIZE + 12), opacity: show ? 1 : 0, pointerEvents: show ? "auto" : "none", transform: show ? "none" : "translateY(8px)" }}>
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M12 19V5M5.5 11.5 12 5l6.5 6.5" stroke="#fff" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"/>
      </svg>
    </button>
  );
}

// ปุ่ม "บนสุดของตาราง" — ลอยอยู่มุมขวาล่างของ "ตัวตาราง" ที่เลื่อนอยู่ในกล่องของตัวเอง (เช่นตารางรายเดือน)
// โผล่เมื่อเลื่อนตารางลงไปแล้ว กดแล้วตารางกลับแถวแรก (หน้าเพจไม่ขยับ)
function TableTopButton() {
  useLang();
  const [box, setBox] = useState(null);   // { left, top } ตำแหน่งปุ่ม หรือ null = ซ่อน
  const elRef = useRef(null);
  useEffect(() => {
    let raf = 0;
    const W = 44, H = 44, M = 14, GAP = 16;
    const place = () => {
      raf = 0;
      const el = elRef.current;
      if (!el || !el.isConnected || el.scrollTop < 200) { setBox(null); return; }
      const r = el.getBoundingClientRect();
      let visBottom = Math.min(r.top + el.clientTop + el.clientHeight, window.innerHeight - bnavH());
      const foot = el.querySelector("tfoot");                          // แถวรวมที่ตรึงไว้ด้านล่าง → วางปุ่มเหนือแถวรวม
      const fcell = foot && foot.querySelector("td,th");                 // ใช้ตำแหน่ง "ช่อง" (ตรึงแบบ sticky) ไม่ใช่กล่อง tfoot
      if (fcell) { const fr = fcell.getBoundingClientRect(); if (fr.height && fr.top < visBottom) visBottom = fr.top; }
      const head = el.querySelector("thead");
      const visTop = Math.max(r.top, head ? head.getBoundingClientRect().bottom : r.top, 0);
      if (visBottom - visTop < H + 60) { setBox(null); return; }   // ตารางแทบไม่อยู่ในจอ
      const vw = document.documentElement.clientWidth || window.innerWidth;   // ไม่นับแถบเลื่อนของหน้า (ปุ่มมุมจอวัดจากตรงนี้)
      let left = r.left + el.clientLeft + el.clientWidth - W - M;
      left = Math.min(left, vw - FAB_GAP - FAB_SIZE - GAP - W);          // เว้นระยะจากปุ่มมุมจอ ไม่ทับกัน
      setBox({ left: Math.max(4, left), top: visBottom - H - M });
    };
    const req = () => { if (!raf) raf = requestAnimationFrame(place); };
    const onAnyScroll = (e) => {
      const el = e.target;
      if (el && el.nodeType === 1 && !/^(INPUT|TEXTAREA|SELECT)$/.test(el.tagName) && !el.closest("[data-calc],[role=listbox],[data-table-top]")
          && el.scrollHeight > el.clientHeight + 4) elRef.current = el;
      req();
    };
    document.addEventListener("scroll", onAnyScroll, { capture: true, passive: true });
    window.addEventListener("resize", req);
    return () => { document.removeEventListener("scroll", onAnyScroll, { capture: true }); window.removeEventListener("resize", req); if (raf) cancelAnimationFrame(raf); };
  }, []);
  if (!box) return null;
  const toTop = () => { const el = elRef.current; if (!el) return; try { el.scrollTo({ top: 0, behavior: "smooth" }); } catch { el.scrollTop = 0; } };
  return (
    <button onClick={toTop} data-table-top title={t("กลับแถวแรกของตาราง","Back to the first row of the table")} aria-label={t("กลับแถวแรกของตาราง","Back to the first row of the table")}
      style={{ position:"fixed", left:box.left, top:box.top, zIndex:94, width:44, height:44, padding:0, borderRadius:14, border:"1.5px solid #1e3a8a",
        background:"rgba(255,255,255,0.97)", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center",
        boxShadow:"0 6px 18px rgba(30,58,138,0.22)" }}>
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M5 4h14M12 20V9M6.5 14.5 12 9l5.5 5.5" stroke="#1e3a8a" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round"/>
      </svg>
    </button>
  );
}

function CalculatorPopup({ selSum = null }) {
  useLang();
  const open = useCalcOpen();
  const [expr, setExpr] = useState("");
  const [done, setDone] = useState(null);            // { expr, v } หลังกด =
  const [err, setErr] = useState("");
  const [hist, setHist] = useState([]);              // ประวัติ 10 รายการล่าสุด (เฉพาะรอบนี้)
  const [showHist, setShowHist] = useState(false);
  const [copied, setCopied] = useState(false);
  const [pos, setPos] = useState(null);              // null = มุมขวาล่าง / {x,y} หลังลากย้าย
  const boxRef = useRef(null);
  const drag = useRef(null);

  const exprRef = useRef(null);
  useEffect(() => { if (open) setTimeout(() => boxRef.current?.focus(), 0); }, [open]);
  // วางตำแหน่งด้วย top/left เสมอ (เปิดครั้งแรกที่มุมขวาล่าง) — เวลาความสูงเปลี่ยน ปุ่มจะไม่เลื่อนหนีนิ้ว/เมาส์
  const [, setVp] = useState(0);
  useEffect(() => {   // หมุนจอ/ย่อหน้าต่าง → วาดใหม่แล้วดึงกลับเข้าจอ
    if (!open) return;
    const f = () => setVp(v => v + 1);
    window.addEventListener("resize", f); window.addEventListener("orientationchange", f);
    return () => { window.removeEventListener("resize", f); window.removeEventListener("orientationchange", f); };
  }, [open]);
  useLayoutEffect(() => {   // ทุกครั้งที่วาด: ถ้าหลุดขอบ (ความสูงเปลี่ยน / จอเล็กลง / เปิดใหม่) ดึงกลับเข้าจอ
    if (!open || !pos || !boxRef.current) return;
    const r = boxRef.current.getBoundingClientRect();
    const x = Math.min(Math.max(4, pos.x), Math.max(4, window.innerWidth - r.width - 4));
    const y = Math.min(Math.max(4, pos.y), Math.max(4, window.innerHeight - r.height - 4));
    if (x !== pos.x || y !== pos.y) setPos({ x, y });
  });
  useLayoutEffect(() => {
    if (!open || pos || !boxRef.current) return;
    const r = boxRef.current.getBoundingClientRect();
    const side = window.innerWidth >= r.width + FAB_SIZE + FAB_GAP * 2 + 8;   // มีที่ข้างปุ่มลอยไหม
    setPos(side
      ? { x: window.innerWidth - r.width - FAB_SIZE - FAB_GAP - 12, y: Math.max(4, window.innerHeight - bnavH() - r.height - FAB_GAP) }
      : { x: Math.max(4, window.innerWidth - r.width - 12), y: Math.max(4, window.innerHeight - bnavH() - r.height - FAB_SIZE - FAB_GAP - 12) });
  }, [open, pos]);
  useEffect(() => { const el = exprRef.current; if (el) el.scrollLeft = el.scrollWidth; });   // สูตรยาว → เลื่อนให้เห็นท้ายสุด
  if (!open) return null;

  const preview = calcEval(expr);
  const OPS = "+-*/";
  const press = (k) => {
    setErr(""); setCopied(false);
    if (k === "AC") { setExpr(""); setDone(null); return; }
    if (k === "⌫") { if (done) { setDone(null); return; } setExpr(e => e.slice(0, -1)); return; }
    if (k === "=") {
      const r = calcEval(expr);
      if (!r.ok) { if (!r.empty) setErr(r.div0 ? t("หารด้วย 0 ไม่ได้","Can't divide by 0") : t("สูตรไม่ครบ","Incomplete formula")); return; }
      if (Math.abs(r.v) >= CALC_MAX) { setErr(t("ตัวเลขใหญ่เกินไป","Number too large")); return; }
      if (done && expr === plainNum(done.v)) return;                       // กด = ซ้ำ → ไม่ต้องทำอะไร
      setDone({ expr, v: r.v });
      if (/[+\-*/%]/.test(expr.replace(/^-/, ""))) setHist(h => [{ expr, v: r.v }, ...h].slice(0, 10));   // เก็บเฉพาะที่เป็นสูตรจริง
      setExpr(plainNum(r.v)); return;
    }
    // หลังกด = : พิมพ์ตัวเลขต่อ = เริ่มใหม่ · พิมพ์เครื่องหมายต่อ = ใช้ผลลัพธ์เดิมต่อ
    let base = expr;
    if (done) { setDone(null); if (/[\d.(]/.test(k)) base = ""; }
    if (OPS.includes(k)) {
      if (!base && k !== "-") return;
      if (OPS.includes(base.slice(-1))) base = base.slice(0, -1);   // กดเครื่องหมายซ้ำ → แทนตัวเดิม
      setExpr(base + k); return;
    }
    if (k === ".") { const seg = base.split(/[+\-*/()%]/).pop(); if (seg.includes(".")) return; setExpr(base + (seg === "" ? "0." : ".")); return; }
    if (k === "%") { if (!/[\d)]$/.test(base)) return; setExpr(base + "%"); return; }
    if (/\d/.test(k) && (base.split(/[+\-*/()%]/).pop() || "").replace(".", "").length >= 15) return;   // ยาวเกินความละเอียดตัวเลข
    setExpr(base + k);
  };
  const insertNumber = (v) => {
    const n = plainNum(Math.round(v * 100) / 100, 2);
    setDone(null); setErr("");
    setExpr(e => (!e || /[+\-*/(]$/.test(e)) ? e + (n.startsWith("-") && e ? `(${n})` : n) : e + "+" + n);
  };
  const result = done ? done.v : preview.ok ? preview.v : null;
  const copy = async () => {
    if (result == null) return;
    const txt = plainNum(result, 4);   // ความละเอียดเท่าที่แสดงบนจอ
    try { await navigator.clipboard.writeText(txt); } catch { /* บางเบราว์เซอร์ไม่อนุญาต */ }
    setCopied(true); setTimeout(() => setCopied(false), 1500);
  };
  const onKey = (e) => {
    e.stopPropagation();                                 // ไม่ให้ไปโดนคีย์ลัดของหน้า (Undo, Esc ปิดฟอร์ม)
    const k = e.key;
    if ((e.ctrlKey || e.metaKey) && k.toLowerCase() === "c" && !String(window.getSelection?.() || "")) { e.preventDefault(); copy(); return; }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (k === "Escape") { e.preventDefault(); CalcStore.set(false); return; }
    if (k === " " || k === "Spacebar") { e.preventDefault(); return; }   // กันเผลอกดปุ่มที่โฟกัสอยู่ (เช่นปิด)
    // แป้นพิมพ์ภาษาไทย: แถวตัวเลขให้ ๅ / - ภ ถ … และ x c . / - = ให้ ป แ ใ ฝ ข ช → ดูจากตำแหน่งปุ่ม (e.code) ประกอบ
    const code = e.code || "";
    const map = { "Enter": "=", "Backspace": "⌫", "Delete": "AC", "x": "*", "X": "*", "c": "AC", "C": "AC", ",": "" };
    const byCode = { NumpadAdd:"+", NumpadSubtract:"-", NumpadMultiply:"*", NumpadDivide:"/", NumpadDecimal:".", NumpadEnter:"=", NumpadEqual:"=",
      KeyX:"*", KeyC:"AC", Period:".", Slash:"/", Minus:"-", Equal:"=" };
    const dm = /^Digit(\d)$/.exec(code);
    let m = null;
    if (/^[0-9]$/.test(k)) m = k;
    else if (dm && !"%*()".includes(k)) m = dm[1];                 // แถวตัวเลข (ยกเว้น Shift+5/8/9/0 บนแป้นอังกฤษ = % * ( ))
    else if (/^Numpad\d$/.test(code)) m = code.slice(6);
    else if (k in map) m = map[k];
    else if (/^[.+\-*/%()=]$/.test(k)) m = k;
    else if (byCode[code]) m = byCode[code];
    if (m === "") { e.preventDefault(); return; }
    if (/^[0-9.+\-*/%()=⌫]$/.test(m) || m === "AC" || m === "⌫") { e.preventDefault(); press(m); }
  };
  const onPaste = (e) => {
    const raw = e.clipboardData?.getData("text") || "";
    if (!raw.trim()) return;
    e.preventDefault(); setCopied(false);
    // หลายช่องจาก Excel (แท็บ/ขึ้นบรรทัด) → บวกกัน · "(1,234.50)" แบบบัญชี = ติดลบ · × ÷ − x → เครื่องหมายคิดเลข
    const cells = raw.split(/[\t\r\n]+/).map(c => c.trim()).filter(Boolean).map(c => {
      const neg = /^\(\s*[\d,.]+\s*\)$/.test(c);
      let v = c.replace(/[฿$\s]/g, "").replace(/,/g, "").replace(/×|[xX]/g, "*").replace(/÷/g, "/").replace(/[−–]/g, "-");
      if (neg) v = "-" + v.replace(/[()]/g, "");
      return v;
    });
    const txt = cells.length > 1 ? cells.map(c => (/^-/.test(c) ? `(${c})` : c)).join("+") : (cells[0] || "");
    if (!txt || /[^\d.+\-*/%()]/.test(txt)) { setErr(t("วางได้เฉพาะตัวเลข/สูตร","Only numbers or formulas can be pasted")); return; }
    setErr("");
    if (done) { setDone(null); setExpr(/^[\d.(]/.test(txt) ? txt : expr + txt); return; }   // หลังกด = : ตัวเลข → เริ่มใหม่
    setExpr(x => x + txt);
  };
  // ลากย้ายจากแถบหัว
  const onDragStart = (e) => {
    if (e.button !== 0 || e.target.closest("button")) return;
    const r = boxRef.current.getBoundingClientRect();
    drag.current = { dx: e.clientX - r.left, dy: e.clientY - r.top, w: r.width, h: r.height };
    e.currentTarget.setPointerCapture?.(e.pointerId);
  };
  const onDragMove = (e) => {
    const d = drag.current; if (!d) return;
    const x = Math.min(Math.max(4, e.clientX - d.dx), window.innerWidth - d.w - 4);
    const y = Math.min(Math.max(4, e.clientY - d.dy), window.innerHeight - d.h - 4);
    setPos({ x, y });
  };
  const onDragEnd = () => { drag.current = null; };

  const KEYS = [
    ["AC","⌫","%","/"],
    ["7","8","9","*"],
    ["4","5","6","-"],
    ["1","2","3","="],
    ["0",".","+"],
  ];
  const label = { "/": "÷", "*": "×", "-": "−" };
  const keyStyle = (k) => {
    const op = OPS.includes(k) || k === "%";
    const base = { height:46, borderRadius:14, border:"1px solid rgba(255,255,255,0.9)", background:"rgba(255,255,255,0.9)",
      boxShadow:"0 1px 2px rgba(15,23,42,0.06), inset 0 1px 0 rgba(255,255,255,0.8)", fontSize:18, fontWeight:600, color:T.textPrimary,
      cursor:"pointer", fontFamily:"inherit", display:"flex", alignItems:"center", justifyContent:"center", userSelect:"none", transition:"transform .06s, background .12s" };
    if (k === "=") return { ...base, gridRow:"span 2", height:"auto", fontSize:24, color:"#fff", border:"none", background:"linear-gradient(180deg,#3b82f6 0%,#1d4ed8 100%)", boxShadow:"0 6px 16px rgba(37,99,235,0.35)" };
    if (k === "AC") return { ...base, color:"#dc2626", fontSize:15, fontWeight:700 };
    if (k === "⌫") return { ...base, color:T.textSecondary, fontSize:16 };
    if (op) return { ...base, color:T.blue, background:"rgba(219,234,254,0.95)", fontSize:20 };
    return base;
  };
  const place = pos ? { left: pos.x, top: pos.y } : { right: 20, bottom: 80, visibility: "hidden" };

  return (
    <div ref={boxRef} data-calc tabIndex={-1} onKeyDown={onKey} onPaste={onPaste} role="dialog" aria-label={t("เครื่องคิดเลข","Calculator")}
      onPointerDownCapture={() => { const b = boxRef.current; if (b && !b.contains(document.activeElement)) b.focus({ preventScroll: true }); }}
      className="calc-pop"
      style={{position:"fixed", ...place, zIndex:250, width:292, maxWidth:"calc(100vw - 24px)", maxHeight:"calc(100vh - 8px)", overflowY:"auto", outline:"none",
        background:"rgba(246,249,253,0.93)", backdropFilter:"blur(18px) saturate(160%)", WebkitBackdropFilter:"blur(18px) saturate(160%)",
        border:"1px solid rgba(255,255,255,0.7)", borderRadius:24, boxShadow:"0 24px 60px rgba(15,23,42,0.28), inset 0 1px 0 rgba(255,255,255,0.6)",
        padding:14, animation:"fadeIn 0.15s ease"}}>
      {/* แถบหัว: ลากย้ายได้ */}
      <div onPointerDown={onDragStart} onPointerMove={onDragMove} onPointerUp={onDragEnd} onPointerCancel={onDragEnd}
        style={{display:"flex",alignItems:"center",gap:6,marginBottom:6,cursor:"grab",touchAction:"none"}}>
        <span style={{fontSize:12,fontWeight:700,color:T.textMuted,letterSpacing:0.5,flex:1}}><Ico name="calc" size={14} /> {t("เครื่องคิดเลข","Calculator")}</span>
        <button onClick={()=>setShowHist(v=>!v)} title={t("ประวัติการคำนวณ","History")} aria-label={t("ประวัติการคำนวณ","History")} tabIndex={-1}
          style={{border:"none",background:showHist?"rgba(37,99,235,0.12)":"transparent",color:showHist?T.blue:T.textMuted,borderRadius:8,minWidth:32,height:30,cursor:"pointer",fontSize:14,display:"inline-grid",placeItems:"center"}}><Ico name="clock" size={16} /></button>
        <button onClick={copy} title={t("คัดลอกผลลัพธ์","Copy result")} aria-label={t("คัดลอกผลลัพธ์","Copy result")} tabIndex={-1} disabled={result==null}
          style={{border:"none",background:"transparent",color:copied?T.green:T.textMuted,borderRadius:8,minWidth:32,height:30,cursor:result==null?"default":"pointer",fontSize:copied?12:14,fontWeight:700}}>{copied ? t("คัดลอกแล้ว","Copied") : <Ico name="copy" size={16} />}</button>
        <button onClick={()=>CalcStore.set(false)} title={t("ปิด (Esc)","Close (Esc)")} aria-label={t("ปิด","Close")} tabIndex={-1}
          style={{border:"none",background:"transparent",color:T.textMuted,borderRadius:8,minWidth:32,height:30,cursor:"pointer",fontSize:15,lineHeight:1}}>✕</button>
      </div>

      {/* จอแสดงผล */}
      <div style={{textAlign:"right",padding:"4px 6px 10px"}}>
        <div data-calc-expr ref={exprRef} style={{fontSize:14,color:T.textSecondary,fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap",overflowX:"auto",overflowY:"hidden",height:22,lineHeight:"20px",scrollbarWidth:"none"}}>
          {done ? prettyExpr(done.expr) : (expr ? prettyExpr(expr) : " ")}
        </div>
        <div data-calc-result style={{fontSize: err ? 20 : result != null && fmtCalc(result).length > 16 ? 20 : result != null && fmtCalc(result).length > 11 ? 25 : 32, fontWeight:750, color: err ? T.red : done ? T.textPrimary : result != null ? T.textSecondary : T.textMuted,
          fontVariantNumeric:"tabular-nums",letterSpacing:-0.5,height:42,lineHeight:"42px",marginTop:2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
          {err ? err : result != null ? `=${fmtCalc(result)}` : "0"}
        </div>
      </div>

      {selSum != null && (
        <button onClick={()=>insertNumber(selSum)} tabIndex={-1} title={t("ใส่ผลรวมของช่องที่ลากเลือกในตาราง","Insert the sum of the cells selected in the table")}
          style={{width:"100%",marginBottom:8,border:"1px dashed rgba(37,99,235,0.45)",background:"rgba(219,234,254,0.55)",color:T.blue,borderRadius:12,padding:"7px 10px",fontSize:12.5,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
          Σ {t("ใส่ผลรวมที่เลือก","Insert selected sum")} · {fmtCalc(Math.round(selSum*100)/100)}
        </button>
      )}

      {showHist ? (
        <div style={{maxHeight:258,overflowY:"auto",display:"flex",flexDirection:"column",gap:6}}>
          {hist.length === 0 && <div style={{fontSize:12.5,color:T.textMuted,textAlign:"center",padding:"30px 0"}}>{t("ยังไม่มีประวัติ","No history yet")}</div>}
          {hist.map((h, i) => (
            <button key={i} tabIndex={-1} onClick={()=>{ setExpr(String(h.v)); setDone(null); setShowHist(false); }}
              style={{textAlign:"right",border:"1px solid rgba(255,255,255,0.75)",background:"rgba(255,255,255,0.6)",borderRadius:12,padding:"8px 10px",cursor:"pointer",fontVariantNumeric:"tabular-nums"}}>
              <div style={{fontSize:12,color:T.textMuted}}>{prettyExpr(h.expr)}</div>
              <div style={{fontSize:15,fontWeight:700,color:T.textPrimary}}>={fmtCalc(h.v)}</div>
            </button>
          ))}
        </div>
      ) : (
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}}>
          {KEYS.flat().map(k => (
            <button key={k} tabIndex={-1} onMouseDown={e=>e.preventDefault()} onClick={()=>press(k)} style={keyStyle(k)}
              aria-label={k === "⌫" ? t("ลบทีละตัว","Backspace") : k === "AC" ? t("ล้าง","Clear") : undefined}>
              {label[k] || k}
            </button>
          ))}
        </div>
      )}
      <div style={{fontSize:11,color:T.textMuted,textAlign:"center",marginTop:8}}>{t("พิมพ์จากคีย์บอร์ดได้ · 100+7% = 107 · Esc ปิด","Keyboard works · 100+7% = 107 · Esc to close")}</div>
    </div>
  );
}

function LangToggle({ dark = true }) {
  useLang();
  const on  = dark ? "#fff" : T.textPrimary;
  const bg  = dark ? "rgba(255,255,255,0.15)" : "#fff";
  const bd  = dark ? "rgba(255,255,255,0.3)"  : T.cardBorder;
  // ปุ่มสลับภาษาแบบช่องเดียว — โชว์ภาษาที่ใช้อยู่ตอนนี้ (TH หรือ EN) กดแล้วสลับ ตัวหนังสือก็สลับตาม
  return (
    <button onClick={toggleLang} title={_LANG==="th"?"เปลี่ยนเป็น English":"Switch to ไทย"}
      style={{background:bg,border:`1px solid ${bd}`,borderRadius:8,padding:"5px 13px",minHeight:32,fontSize:13,fontWeight:800,color:on,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6,whiteSpace:"nowrap",minWidth:66,justifyContent:"center"}}>
      🌐 {_LANG==="en" ? "EN" : "TH"}
    </button>
  );
}

// จอมือถือ (≤640px) — ใช้สลับตาราง ↔ การ์ด และแสดงแถบแท็บด้านล่าง (ต้องเปลี่ยน "โครงหน้า" ไม่ใช่แค่ CSS)
const PHONE_MQ = "(max-width: 640px)";
function useIsPhone() {
  const get = () => (typeof window !== "undefined" && window.matchMedia) ? window.matchMedia(PHONE_MQ).matches : false;
  const [v, setV] = useState(get);
  useEffect(() => {
    if (typeof window === "undefined" || !window.matchMedia) return;
    const m = window.matchMedia(PHONE_MQ); const f = () => setV(m.matches); f();
    if (m.addEventListener) m.addEventListener("change", f); else if (m.addListener) m.addListener(f);
    return () => { if (m.removeEventListener) m.removeEventListener("change", f); else if (m.removeListener) m.removeListener(f); };
  }, []);
  return v;
}

// แถบแท็บด้านล่าง (มือถือเท่านั้น) — เมนูหลักของแผนกอยู่ใต้นิ้วโป้ง แทนแท็บด้านบน
function BottomNav({ items }) {
  useLang();
  useEffect(() => { document.body.classList.add("has-bnav"); return () => document.body.classList.remove("has-bnav"); }, []);
  return (
    <nav data-bottom-nav aria-label={t("เมนูหลัก","Main menu")}
      style={{position:"fixed",left:0,right:0,bottom:0,zIndex:93,background:"#fff",borderTop:`1px solid ${T.cardBorder}`,display:"flex",
        paddingBottom:"env(safe-area-inset-bottom, 0px)",boxShadow:"0 -4px 16px rgba(15,23,42,0.06)"}}>
      {items.map(it => (
        <button key={it.key} onClick={it.onClick} aria-current={it.on ? "page" : undefined} data-bnav={it.key}
          style={{flex:1,minWidth:0,height:BNAV_H,border:"none",background:"none",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:3,
            cursor:"pointer",color:it.on?T.blue:T.textSecondary,fontSize:11,fontWeight:it.on?700:500,position:"relative"}}>
          {it.on && <span aria-hidden="true" style={{position:"absolute",top:0,left:"25%",right:"25%",height:3,borderRadius:"0 0 3px 3px",background:T.blue}}/>}
          <Ico name={it.icon} size={20} />
          <span style={{maxWidth:"100%",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",padding:"0 4px"}}>{it.label}</span>
        </button>
      ))}
    </nav>
  );
}

// ที่แสดงกล่องข้อความในแอป (uiAlert / uiConfirm / uiPrompt) — ทีละกล่อง ตามลำดับ
function DialogHost() {
  useLang();
  const [q, setQ] = useState(DialogStore.q);
  const [text, setText] = useState("");
  const okRef = useRef(null), inRef = useRef(null), cancelRef = useRef(null);
  useEffect(() => { DialogStore.subs.add(setQ); return () => DialogStore.subs.delete(setQ); }, []);
  const d = q[0];
  const matchOk = !d || d.kind !== "prompt" || !d.match || text.trim() === String(d.match).trim();
  const close = (val) => {
    if (!d) return;
    const r = d.resolve; setText(""); DialogStore.shift();
    r(d.kind === "alert" ? undefined : d.kind === "confirm" ? !!val : (val === null ? null : val));
  };
  // โฟกัสปุ่ม/ช่องในกล่องตอนเปิด แล้วคืนโฟกัสเดิมตอนปิด
  useEffect(() => {
    if (!d) return;
    const prev = document.activeElement;
    const f = setTimeout(() => { (d.kind === "prompt" ? inRef.current : d.danger && cancelRef.current ? cancelRef.current : okRef.current)?.focus(); }, 0);
    return () => { clearTimeout(f); if (prev && prev.focus && prev.isConnected) try { prev.focus(); } catch {} };
  }, [d]);
  // จับปุ่มก่อนส่วนอื่นของแอป (Esc ในตาราง / เครื่องคิดเลข จะได้ไม่ทำงานซ้อน) + วนโฟกัสอยู่ในกล่อง
  useEffect(() => {
    if (!d) return;
    const onKey = (e) => {
      if (e.key === "Escape") { e.preventDefault(); e.stopImmediatePropagation(); close(d.kind === "prompt" ? null : false); }
      else if (e.key === "Enter" && !(e.target && e.target.tagName === "BUTTON")) { e.preventDefault(); e.stopImmediatePropagation(); if (matchOk) close(d.kind === "prompt" ? text : true); }
      else if (e.key === "Tab") {
        const els = [inRef.current, cancelRef.current, okRef.current].filter(x => x && !x.disabled);
        if (!els.length) return;
        const i = els.indexOf(document.activeElement);
        e.preventDefault(); els[(i + (e.shiftKey ? -1 : 1) + els.length) % els.length].focus();
      }
    };
    window.addEventListener("keydown", onKey, true);
    return () => window.removeEventListener("keydown", onKey, true);
  }, [d, text, matchOk]);
  if (!d) return null;
  const title = d.title || (d.kind === "alert" ? t("แจ้งเตือน","Notice") : d.kind === "prompt" ? t("ยืนยัน","Confirm") : t("ยืนยัน","Confirm"));
  const okLabel = d.okLabel || t("ตกลง","OK");
  return (
    <div data-ui-dialog={d.kind} onMouseDown={e => { if (e.target === e.currentTarget && d.kind === "alert") close(); }}
      style={{position:"fixed",inset:0,zIndex:400,background:"rgba(15,23,42,0.45)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div role={d.kind === "alert" ? "alertdialog" : "dialog"} aria-modal="true" aria-labelledby="ui-dlg-t" aria-describedby="ui-dlg-m"
        style={{background:"#fff",borderRadius:16,width:"min(440px, 100%)",maxHeight:"calc(100vh - 32px)",overflowY:"auto",boxShadow:"0 24px 64px rgba(15,23,42,0.28)",padding:"20px 22px 18px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          {(d.danger || d.kind === "alert") && <span style={{width:32,height:32,borderRadius:10,display:"grid",placeItems:"center",background:d.danger?T.redBg:T.amberBg,color:d.danger?T.red:T.amber,flexShrink:0}}><Ico name="alert" size={18} /></span>}
          <h2 id="ui-dlg-t" style={{margin:0,fontSize:16,fontWeight:650,color:T.textPrimary}}>{title}</h2>
        </div>
        <div id="ui-dlg-m" data-ui-message style={{fontSize:14,lineHeight:1.6,color:T.textSecondary,whiteSpace:"pre-line",overflowWrap:"anywhere"}}>{String(d.message ?? "").replace(/^\u26A0\uFE0F?\s*/, "")}</div>
        {d.kind === "prompt" && (
          <input ref={inRef} data-ui-input value={text} onChange={e => setText(e.target.value)} placeholder={d.placeholder || ""} className="input-base"
            aria-label={title} style={{marginTop:12}} />
        )}
        <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:18,flexWrap:"wrap"}}>
          {d.kind !== "alert" && <button ref={cancelRef} data-ui-cancel className="btn-ghost" onClick={() => close(d.kind === "prompt" ? null : false)}>{d.cancelLabel || t("ยกเลิก","Cancel")}</button>}
          <button ref={okRef} data-ui-ok className="btn-primary" disabled={!matchOk} onClick={() => close(d.kind === "prompt" ? text : true)}
            style={{...(d.danger ? {background:T.red} : {}), ...(matchOk ? {} : {opacity:0.5,cursor:"not-allowed"})}}>{okLabel}</button>
        </div>
      </div>
    </div>
  );
}

// หัวข้อส่วนของฟอร์ม (เลขวงกลม + ชื่อ + คำอธิบายสั้น) — ใช้ในฟอร์ม PO
function FormStep({ n, title, hint, first = false }) {
  return (
    <div data-form-step={n} style={{gridColumn:"1/-1",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginTop:first?0:8,paddingTop:first?0:16,borderTop:first?"none":`1px solid ${T.cardBorder}`}}>
      <span aria-hidden="true" style={{width:24,height:24,borderRadius:"50%",background:T.blueLight,color:T.blue,fontSize:12,fontWeight:700,display:"inline-grid",placeItems:"center",flexShrink:0}}>{n}</span>
      <h3 style={{margin:0,fontSize:14,fontWeight:650,color:T.textPrimary}}>{title}</h3>
      {hint && <span style={{fontSize:12,color:T.textMuted}}>{hint}</span>}
    </div>
  );
}

// ช่องวันที่ที่แสดงเป็นวันที่ไทย (26 ก.ย. 2569) — ปฏิทินของเบราว์เซอร์แสดงตามภาษาเครื่อง (มักเป็น mm/dd/yyyy)
// ยังใช้ <input type="date"> ตัวเดิม (ค่าเป็น yyyy-mm-dd เหมือนเดิม) แค่ซ้อนข้อความวันที่ไทยไว้ตอนที่ไม่ได้พิมพ์อยู่
const fmtDateInput = (iso) => {
  if (!iso) return "";
  const [y, m, d] = String(iso).slice(0, 10).split("-").map(Number);
  if (!y || !m || !d) return String(iso);
  return new Date(y, m - 1, d).toLocaleDateString(_LANG === "en" ? "en-GB" : "th-TH", { day: "numeric", month: "short", year: "numeric" });
};
function DateInput({ value, onChange, disabled, style, ...rest }) {
  useLang();
  const [focus, setFocus] = useState(false);
  const cover = !focus;
  return (
    <span style={{ position: "relative", display: "block", minWidth: 0 }}>
      <input type="date" value={value || ""} onChange={onChange} disabled={disabled} className="input-base date-th"
        onFocus={() => setFocus(true)} onBlur={() => setFocus(false)} data-cover={cover ? "1" : undefined}
        style={{ width: "100%", ...(style || {}), ...(cover ? { color: "transparent" } : {}) }} {...rest} />
      {cover && (
        <span aria-hidden="true" data-date-text className="date-th-text" style={{ position: "absolute", left: 14, right: 38, top: "50%", transform: "translateY(-50%)", pointerEvents: "none",
          color: value ? (disabled ? T.textMuted : T.textPrimary) : T.textMuted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", fontVariantNumeric: "tabular-nums" }}>
          {value ? fmtDateInput(value) : t("เลือกวันที่", "Pick a date")}
        </span>
      )}
    </span>
  );
}

// แถบหัวสีขาวของหน้าที่ไม่ใช่หน้าแผนก (หน้ารายชื่อโครงการ / Admin / เลือกแผนก) — หน้าตาเดียวกับ Shell
function TopBar({ onBack, title, sub, brand = false, children }) {
  return (
    <div className="app-header shell-bar" style={{background:"#fff",borderTop:`3px solid ${T.blue}`,borderBottom:`1px solid ${T.cardBorder}`,padding:"9px 28px",display:"flex",alignItems:"center",gap:12,minHeight:58}}>
      {onBack && (
        <button onClick={onBack} title={t("กลับหน้าก่อนหน้า","Go back")} aria-label={t("กลับ","Back")}
          style={{border:`1px solid ${T.cardBorder}`,background:"#fff",color:T.textSecondary,cursor:"pointer",borderRadius:8,width:36,height:36,display:"grid",placeItems:"center",flexShrink:0}}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M15 5l-7 7 7 7" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"/></svg>
        </button>
      )}
      {brand && <span aria-hidden="true" style={{width:34,height:34,borderRadius:9,background:T.blue,color:"#fff",display:"grid",placeItems:"center",fontSize:13,fontWeight:800,letterSpacing:0.5,flexShrink:0}}>TC</span>}
      <div style={{minWidth:0,flex:1}}>
        <h1 style={{margin:0,fontSize:15,fontWeight:650,color:T.textPrimary,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{title}</h1>
        {sub && <div className="crumb-txt" style={{fontSize:12,color:T.textMuted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{sub}</div>}
      </div>
      {children}
    </div>
  );
}

// ─── Shell ────────────────────────────────────────────────────────────────────
// เมนูผู้ใช้ (มุมขวาของแถบหัว): ชื่อ/บทบาท · ภาษา · ออกจากระบบ — รวมไว้ที่เดียว แถบหัวจะได้โล่ง
function UserMenu({ session, onLogout }) {
  useLang();
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const off = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    const esc = (e) => { if (e.key === "Escape") { e.stopPropagation(); setOpen(false); } };
    document.addEventListener("mousedown", off); document.addEventListener("keydown", esc, true);
    return () => { document.removeEventListener("mousedown", off); document.removeEventListener("keydown", esc, true); };
  }, [open]);
  if (!session) return null;
  const initial = (session.name || "?").trim().charAt(0).toUpperCase();
  return (
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={() => setOpen(v => !v)} aria-haspopup="menu" aria-expanded={open} data-user-menu title={session.name}
        style={{display:"flex",alignItems:"center",gap:8,border:`1px solid ${T.cardBorder}`,background:"#fff",borderRadius:999,padding:"3px 10px 3px 3px",cursor:"pointer",minHeight:36}}>
        <span style={{width:28,height:28,borderRadius:"50%",background:T.blueLight,color:T.blue,display:"grid",placeItems:"center",fontWeight:700,fontSize:12}}>{initial}</span>
        <span className="um-name" style={{fontSize:13,fontWeight:600,color:T.textPrimary,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{session.name}</span>
        <Ico name="chevrons" size={14} color={T.textMuted} />
      </button>
      {open && (
        <div role="menu" style={{position:"absolute",right:0,top:"calc(100% + 6px)",zIndex:60,minWidth:220,background:"#fff",border:`1px solid ${T.cardBorder}`,borderRadius:12,boxShadow:"0 12px 32px rgba(15,23,42,0.16)",padding:8}}>
          <div style={{padding:"6px 8px 10px",borderBottom:`1px solid ${T.cardBorder}`,marginBottom:6}}>
            <div style={{fontSize:13,fontWeight:650,color:T.textPrimary}}>{session.name}</div>
            <div style={{fontSize:12,color:T.textMuted}}>{ROLE_LABELS[session.role] || session.role}</div>
          </div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,padding:"4px 8px"}}>
            <span style={{fontSize:13,color:T.textSecondary}}>{t("ภาษา","Language")}</span>
            <LangToggle dark={false} />
          </div>
          <button role="menuitem" onClick={() => { setOpen(false); onLogout && onLogout(); }}
            style={{width:"100%",textAlign:"left",marginTop:6,border:"none",background:"none",padding:"9px 8px",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600,color:T.red}}>
            {t("ออกจากระบบ","Logout")}
          </button>
        </div>
      )}
    </div>
  );
}

// ─── Shell ────────────────────────────────────────────────────────────────────
// แถบหัวแบบเรียบ: ← กลับ · โครงการ / ชื่อโครงการ · ป้ายแผนก — สีแผนกเหลือแค่ป้าย + เส้นบนบาง ๆ
// (เดิมพื้นไล่สีเต็มแถบ + ปุ่มหลายปุ่มน้ำหนักเท่ากัน) · ภาษา/ออกจากระบบ อยู่ในเมนูผู้ใช้
const DEPT_STYLE = {
  qs:          { color:"#1d4ed8", bg:"#e8efff", label:() => "QS" },
  procurement: { color:"#b45309", bg:"#fdf1e3", label:() => t("จัดซื้อ","Procurement") },
  accounting:  { color:"#047857", bg:"#e3f5ec", label:() => t("บัญชี","Accounting") },
};
function Shell({ role, color, project, onBack, onHome, onDept, children, syncedAt, syncing, session, onLogout }) {
  const dept = DEPT_STYLE[role] || DEPT_STYLE.qs;
  const iconBtn = { border:`1px solid ${T.cardBorder}`, background:"#fff", color:T.textSecondary, cursor:"pointer", borderRadius:8, width:36, height:36, display:"grid", placeItems:"center", flexShrink:0 };
  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",flexDirection:"column"}}>
      <div className="app-header shell-bar" style={{background:"#fff",borderTop:`3px solid ${dept.color}`,borderBottom:`1px solid ${T.cardBorder}`,padding:"9px 28px",display:"flex",alignItems:"center",gap:12,minHeight:58}}>
        <button onClick={onBack} title={t("กลับหน้าก่อนหน้า","Go back")} aria-label={t("กลับ","Back")} style={iconBtn}>
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M15 5l-7 7 7 7" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"/></svg>
        </button>
        <nav aria-label={t("ตำแหน่งหน้า","Breadcrumb")} style={{display:"flex",alignItems:"center",gap:8,minWidth:0,flex:1}}>
          {onHome && (
            <button className="crumb-root" onClick={onHome} title={t("ไปหน้าเลือกโครงการ","Go to projects")}
              style={{border:"none",background:"none",padding:"4px 2px",cursor:"pointer",fontSize:14,color:T.textSecondary,display:"flex",alignItems:"center",gap:6,whiteSpace:"nowrap"}}>
              <Ico name="home" size={16} /><span className="crumb-txt">{t("หน้าโครงการ","Projects")}</span>
            </button>
          )}
          {onHome && <span className="crumb-sep" style={{color:T.textMuted}}>/</span>}
          <span title={project.name} style={{fontSize:15,fontWeight:650,color:T.textPrimary,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{project.name}</span>
          {onDept ? (
            <button onClick={onDept} title={t("เปลี่ยนแผนก","Switch department")}
              style={{border:"none",background:dept.bg,color:dept.color,borderRadius:6,padding:"3px 8px",fontSize:12,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:4,whiteSpace:"nowrap",flexShrink:0}}>
              {dept.label()} <Ico name="chevrons" size={12} />
            </button>
          ) : (
            <span style={{background:dept.bg,color:dept.color,borderRadius:6,padding:"3px 8px",fontSize:12,fontWeight:700,whiteSpace:"nowrap",flexShrink:0}}>{dept.label()}</span>
          )}
        </nav>
        {project.area && <span className="hdr-meta" style={{fontSize:12,color:T.textMuted,whiteSpace:"nowrap"}}>{project.area} ft² · {project.panels} Panels</span>}
        <SyncBadge light syncing={syncing} syncedAt={syncedAt}/>
        <UserMenu session={session} onLogout={onLogout} />
      </div>
      <div style={{flex:1,overflow:"auto"}}>{children}</div>
    </div>
  );
}

// ─── Money input ────────────────────────────────────────────────────────────
// ช่องกรอกยอดเงินที่ (1) โชว์ , คั่นหลักพันให้อ่านง่าย และ (2) พิมพ์บวก/ลบได้
// เช่น "20000+10000" แล้วกด Enter → รวมเป็น 30,000 ให้อัตโนมัติ
// เก็บค่าเป็นตัวเลขล้วน (string ไม่มี ,) ไว้เบื้องหลัง โค้ดส่วนอื่นใช้ parseFloat ได้ตามเดิม
// allowNegative: ช่องที่ติดลบได้ (งานลด/ตัดงานในยอดเพิ่มรายเดือน) — ช่องอื่น (งบ/ยอด PO/ยอดรับ) ติดลบไม่ได้
const evalMoney = (expr, allowNegative = false) => {
  let cleaned = String(expr ?? "").replace(/[,\s]/g, "").replace(/−/g, "-");   // − (ลบแบบ Unicode จาก Excel) → -
  if (!cleaned) return "";
  // รูปแบบบัญชี "(12,000)" = -12,000 (เดิมอ่านเป็น +12,000)
  let sign = 1;
  const paren = cleaned.match(/^\((.*)\)$/);
  if (paren) { sign = -1; cleaned = paren[1]; }
  const terms = cleaned.match(/[+-]?\d*\.?\d+/g);
  if (!terms) return "";
  const sum = sign * terms.reduce((s, t) => s + (parseFloat(t) || 0), 0);
  if (isNaN(sum)) return "";
  // ช่องที่ติดลบไม่ได้: ยังพิมพ์สูตรลบได้ (เช่น 100-20=80) แต่ถ้าผลรวมติดลบ ให้เป็น 0 กันข้อมูลเสียหาย
  // ปัดเป็นสตางค์ (2 ตำแหน่ง) เสมอ — เดิมช่องที่ติดลบไม่ได้เก็บทศนิยมเกิน (เช่น 1,643,254.568) ทำให้ยอดรวมมีเศษ
  const r2 = Math.round(sum * 100) / 100;
  return String(allowNegative ? r2 : Math.max(0, r2));
};
const fmtMoneyInput = (v) => {
  if (v === "" || v == null || isNaN(Number(v))) return "";
  return Number(v).toLocaleString("en-US", { maximumFractionDigits: 2 });
};
function MoneyInput({ value, onChange, placeholder = "0", disabled, className = "input-base", style, onPaste, allowNegative = false }) {
  const [focused, setFocused] = useState(false);
  const [text, setText] = useState("");
  const focusedRef = useRef(false);
  // บันทึกเฉพาะเมื่อค่าเปลี่ยนจริง (แค่คลิก/tab ผ่านช่องไม่ต้องเขียนเซิร์ฟเวอร์) และกัน Enter+blur ยิงซ้ำ
  const commit = () => {
    if (!focusedRef.current) return;
    focusedRef.current = false;
    setFocused(false);
    const v = evalMoney(text, allowNegative);
    const cur = (value === "" || value == null) ? "" : String(parseFloat(value));
    if (v !== cur) onChange(v);
  };
  return (
    <input
      type="text"
      inputMode="decimal"
      className={className}
      disabled={disabled}
      placeholder={placeholder}
      style={{ textAlign: "right", fontVariantNumeric: "tabular-nums", ...(!focused && parseFloat(value) < 0 ? { color: T.red } : {}), ...(style || {}) }}
      value={focused ? text : fmtMoneyInput(value)}
      onFocus={() => { focusedRef.current = true; setFocused(true); setText(value != null && value !== "" ? String(value) : ""); }}
      onClick={(e) => e.stopPropagation()}
      onChange={(e) => setText(e.target.value.replace(/[^0-9.+\-,\s()\u2212]/g, ""))}
      onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); commit(); e.currentTarget.blur(); } }}
      onBlur={commit}
      onPaste={(e) => {
        // Excel-style block paste: if the clipboard holds a grid (tabs / newlines),
        // let the parent distribute it across many cells instead of pasting into one.
        const raw = e.clipboardData?.getData("text") ?? "";
        const isGrid = /[\t\n\r]/.test(raw.replace(/\s+$/, ""));
        if (!onPaste) {
          // ช่องที่วางทั้งบล็อกไม่ได้: ใช้แค่ค่าแรก (กันตัวเลขหลายช่องต่อกันเป็นเลขเดียว เช่น 1,000⏎2,000 → 10,002,000)
          if (isGrid) { e.preventDefault(); setText(raw.split(/[\t\n\r]/)[0].replace(/[^0-9.+\-,\s()\u2212]/g, "")); }
          return;
        }
        if (isGrid) {
          e.preventDefault();
          e.currentTarget.blur();
          onPaste(raw);
        }
      }}
      title={t("พิมพ์บวก/ลบได้ เช่น 20000+10000 แล้วกด Enter เพื่อรวมยอด · วางจาก Excel ได้ทั้งบล็อก","Type +/- e.g. 20000+10000 then Enter to sum · paste a whole block from Excel")}
    />
  );
}

// ─── Procurement: PO Detail Modal ──────────────────────────────────────────────
// Read-only detail view opened by clicking any PO row. Lets the user confirm
// exactly what was entered without hunting through a wide table, and offers
// Edit / Delete from the same place.
// ── ช่องเลือก Account Code แบบพิมพ์ค้นหาได้ (แทน <select> เดิม) ──────────────────
function AccountPicker({ value, onChange, options }) {
  useLang();
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");
  const [hi, setHi] = useState(0);            // แถวที่ไฮไลต์ (เลือกด้วยคีย์บอร์ด)
  const ref = useRef(null);
  const listRef = useRef(null);
  const btnRef = useRef(null);
  // รหัสที่เลือกไว้แต่ไม่อยู่ในรายการ (เช่นถูกซ่อน) ยังแสดงชื่อได้
  const sel = options.find(a => a.code === value) || (value ? { code: value, name: accountOf(value)?.name || "" } : null);
  const ql = q.trim().toLowerCase();
  const list = ql ? options.filter(a => (`${a.code} ${a.name}`).toLowerCase().includes(ql)) : options;
  useEffect(() => { setHi(0); }, [q, open]);
  useEffect(() => {
    if (!open) return;
    const h = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [open]);
  useEffect(() => { listRef.current?.querySelector(`[data-i="${hi}"]`)?.scrollIntoView({ block: "nearest" }); }, [hi]);
  const pick = (code) => { onChange(code); setOpen(false); btnRef.current?.focus(); };
  // ↑↓ เลื่อน · Enter เลือก · Esc ปิดเฉพาะรายการ (ไม่ให้ไปปิดฟอร์ม PO ทั้งฟอร์ม) · Tab ปิด
  const onSearchKey = (e) => {
    if (e.key === "ArrowDown") { e.preventDefault(); setHi(h => Math.min(h + 1, Math.max(list.length - 1, 0))); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setHi(h => Math.max(h - 1, 0)); }
    else if (e.key === "Enter") { e.preventDefault(); if (list[hi]) pick(list[hi].code); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); setOpen(false); btnRef.current?.focus(); }
    else if (e.key === "Tab") setOpen(false);
  };
  return (
    <div ref={ref} style={{position:"relative"}}>
      <button type="button" ref={btnRef} onClick={()=>{ setOpen(o=>!o); setQ(""); }}
        onKeyDown={e=>{ if (e.key==="ArrowDown" && !open) { e.preventDefault(); setOpen(true); setQ(""); } }}
        aria-haspopup="listbox" aria-expanded={open}
        className="input-base" style={{width:"100%",textAlign:"left",cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"space-between",gap:6}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color: sel?T.textPrimary:T.textMuted}}>
          {sel ? `${sel.code} · ${sel.name}` : t("— เลือก Account Code —","— Select Account Code —")}
        </span>
        <span style={{color:T.textMuted,fontSize:11}}>▾</span>
      </button>
      {open && (
        <div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,zIndex:80,background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:10,boxShadow:"0 14px 34px rgba(15,23,42,0.2)",overflow:"hidden"}}>
          <div style={{padding:8,borderBottom:`1px solid ${T.cardBorder}`}}>
            <input autoFocus value={q} onChange={e=>setQ(e.target.value)} onKeyDown={onSearchKey} placeholder={t("พิมพ์ค้นหา รหัส / ชื่อบัญชี · ↑↓ Enter","Type to search code / name · ↑↓ Enter")}
              className="input-base" style={{width:"100%",fontSize:13}} />
          </div>
          <div ref={listRef} role="listbox" className="mscroll" style={{maxHeight:260,overflowY:"auto"}}>
            {value && (
              <div onClick={()=>{ onChange(""); setOpen(false); }} style={{padding:"8px 12px",cursor:"pointer",fontSize:12,color:T.textMuted,borderBottom:`1px solid ${T.cardBorder}`}}>
                {t("— ล้างการเลือก —","— Clear selection —")}
              </div>
            )}
            {list.length===0 && <div style={{padding:12,fontSize:12,color:T.textMuted}}>{t("ไม่พบรหัสที่ค้นหา","No matching code")}</div>}
            {list.map((a, i) => (
              <div key={a.code} data-i={i} role="option" aria-selected={a.code===value} onClick={()=>pick(a.code)}
                onMouseEnter={()=>setHi(i)}
                style={{padding:"8px 12px",cursor:"pointer",fontSize:13,display:"flex",gap:8,alignItems:"baseline",
                  background: i===hi ? T.bg : (a.code===value ? T.blueLight : "transparent"),
                  boxShadow: i===hi ? `inset 3px 0 0 ${T.blue}` : "none"}}>
                <span style={{fontVariantNumeric:"tabular-nums",color:T.blue,fontWeight:600,flexShrink:0}}>{a.code}</span>
                <span style={{color:T.textSecondary,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.name}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// A status badge that's also a dropdown — lets anyone change a PO's status
// in one click from wherever it's shown, instead of opening the full edit form.
function StatusPicker({ status, onChange, compact, disabled }) {
  return (
    <select value={status} disabled={disabled} onClick={e=>e.stopPropagation()} onChange={e=>{ e.stopPropagation(); onChange(e.target.value); }}
      style={{background:STATUS_BG[status],color:STATUS_CLR[status],fontSize:compact?11:12,padding:compact?"5px 8px":"6px 10px",minHeight:compact?30:34,
        borderRadius:20,fontWeight:600,border:`1px solid ${STATUS_CLR[status]}40`,cursor:disabled?"not-allowed":"pointer",outline:"none",
        opacity:disabled?0.65:1}} title={disabled?t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only"):undefined}>
      {PO_STATUS.map(s=><option key={s} value={s}>{poStatusLabel(s)}</option>)}
    </select>
  );
}

export { ErrorBoundary, Loader, SyncBadge, SearchInput, GroupFilter, ICON_PATHS, Ico, EMOJI_ICON, StatCard, usdLine, effRate, CurrencyControl, calcEval, fmtCalc, plainNum, CALC_MAX, prettyExpr, CalcStore, useCalcOpen, fabStyle, CalcFab, ScrollTopFab, TableTopButton, CalculatorPopup, LangToggle, PHONE_MQ, useIsPhone, BottomNav, DialogHost, FormStep, fmtDateInput, DateInput, TopBar, UserMenu, DEPT_STYLE, Shell, evalMoney, fmtMoneyInput, MoneyInput, AccountPicker, StatusPicker };
