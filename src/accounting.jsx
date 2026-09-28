// Tender Cost — แผนกบัญชี (ต้องจ่าย, ภาพรวมงบ, ตารางรวมเดือน)
import { Fragment, useEffect, useState } from "react";
import { Bar, BarChart, Cell, Legend, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { ACCOUNTS, GROUPS, GRP_COLORS, ORPHAN_NAME, PO_STATUS, STATUS_CLR, T, WASTE_LBL, buildCombinedBudget, exportAccountList, fmt, fmt0, fmtDate, fmtK, hiddenSafeForPO, monthShortLabel, paymentStatus, poBilledAmount, poItems, poPaidAmount, poPayLines, poStatusLabel, poTotal, roundReceived, t, todayStr } from "./core.jsx";
import { BottomNav, CurrencyControl, Ico, SearchInput, Shell, StatCard, effRate, usdLine, useIsPhone } from "./ui.jsx";

// ─── Accounting: ตารางรวมรายเดือน (ต้นทุน + Incoming/Received + Payment + PO) ────
//  ต่อ Acc. Code: Tender Cost (งบ), Balance Pending PO (งบ − PO), Stock (มีใน
//  store), Balance Cost (Balance Pending PO − Stock). ตามด้วย 2 กลุ่มเดือน —
//  Incoming/Received (รับจริง=ดำ, ยังเป็นแผน=แดง) และ Payment Plan — แล้วปิดท้าย
//  ด้วยสรุป PO: Total PO (ยอดผูกพัน) และ PO Balance (Total PO − รับจริง).
//  โชว์เฉพาะเดือนที่มีข้อมูล + TOTAL แต่ละกลุ่ม.
function AccountingMatrixTab({ tenderCosts, additions, poEntries, extraItems, hiddenAccounts, incomingPlan = [], usdRate = 0 }) {
  const accounts = exportAccountList(extraItems, hiddenSafeForPO(hiddenAccounts, poEntries), poEntries, incomingPlan)
    .map(a => a.orphan ? { ...a, name: t(ORPHAN_NAME, "(code no longer in the list)") } : a);
  const combined = buildCombinedBudget(tenderCosts, additions);
  const [aSearch, setASearch] = useState("");                       // ค้นหา Acc. Code/ชื่อบัญชี
  const [aSort, setASort] = useState({ key: "code", dir: "asc" });  // เรียงตามหัวคอลัมน์
  const committedByCode = {}, stockByCode = {}, plannedByCode = {};
  const payplan = {};
  const bump = (obj, code, mk, amt) => { if (!mk || !amt) return; (obj[code] = obj[code] || {}); obj[code][mk] = (obj[code][mk] || 0) + amt; };
  const today = todayStr();
  const lateOf = (r) => !!(r.planDate && r.planDate < today && !r.actualDate);
  // แต่ละช่องเดือน (ต่อ code) เก็บแยก rec(รับแล้ว)/po(PO รอเข้า)/plan(แผน) + ธง late — สีเหมือนตารางจัดซื้อ
  const mCell = {};
  const mBucket = (code, mk) => { const c = (mCell[code] = mCell[code] || {}); return (c[mk] = c[mk] || { rec:0, po:0, poLate:false, plan:0, planLate:false }); };
  // แผนของเข้า = อ็อบเจ็กต์รูปเดียวกับ PO → บัคเก็ตตามวันแผนรับของแต่ละงวด
  // (ไม่มีวันแผนก็ใช้วันในฟอร์ม) รวมยอดที่วางแผนไว้ต่อ Acc code ต่อเดือน
  // และเก็บยอดแผนรวมต่อ Acc code (plannedByCode) ใช้คำนวณ Balance Cost ให้ตรงกับ
  // หน้า "แผนของเข้า" ของจัดซื้อ (งบ − Stock − PO − แผน)
  (Array.isArray(incomingPlan) ? incomingPlan : []).forEach(pl => {
    poItems(pl).forEach(it => {
      plannedByCode[it.code] = (plannedByCode[it.code] || 0) + (parseFloat(it.amount) || 0);
      (it.rounds || []).forEach(r => {
        const amt = parseFloat(r.planAmount) || 0;
        if (amt > 0) { const cc = mBucket(it.code, (r.planDate || pl.date || "").slice(0, 7)); cc.plan += amt; if (lateOf(r)) cc.planLate = true; }
      });
    });
  });
  poEntries.forEach(p => {
    poItems(p).forEach(it => {
      const code = it.code;
      committedByCode[code] = (committedByCode[code] || 0) + (parseFloat(it.amount) || 0);
      stockByCode[code] = (stockByCode[code] || 0) + (parseFloat(it.store) || 0);
      (it.rounds || []).forEach(r => {
        if (roundReceived(r)) {
          const cc = mBucket(code, r.actualDate.slice(0, 7)); cc.rec += parseFloat(r.actualAmount) || 0; // รับแล้ว (เขียว)
        } else {
          // PO ที่สั่งแล้วแต่ยังไม่รับ = "PO รอเข้า" (ดำ) · ใช้ยอดจริงที่กรอกไว้ก่อน ไม่มีค่อยใช้แผน
          const amt = parseFloat(r.actualAmount) || parseFloat(r.planAmount) || 0;
          if (amt > 0) { const cc = mBucket(code, (r.actualDate || r.planDate || p.date || "").slice(0, 7)); cc.po += amt; if (lateOf(r)) cc.poLate = true; }
        }
      });
    });
    poPayLines(p).forEach(l => bump(payplan, l.code, l.month, l.amount || 0));
  });
  const monthsOf = (obj) => [...new Set(Object.values(obj).flatMap(m => Object.keys(m)))].sort();
  const mgM = monthsOf(mCell);            // เดือนที่มีของเข้า (รับ/PO/แผน)
  const payM = monthsOf(payplan);
  const cellTot = (c) => c ? (c.rec + c.po + c.plan) : 0;
  const lbl = (mk) => monthShortLabel(mk); // เดือนไทย + ปี พ.ศ. (เช่น "ส.ค. 69") ให้ตรงกับการ์ดแผน/Excel
  const money = (n) => !n ? "-" : (n < 0 ? `(${fmt(Math.abs(n))})` : fmt(n));

  const rows = accounts.map(a => {
    const budget = parseFloat(combined[a.code]) || 0;
    const committed = committedByCode[a.code] || 0;                     // Total PO (สั่งแล้ว)
    const stock = stockByCode[a.code] || 0;
    const planned = plannedByCode[a.code] || 0;                         // ยอดที่วางแผนจะเข้า (ยังไม่เป็น PO)
    const balPO = budget - committed;                                   // Balance Pending PO (งบ − PO)
    const balCost = budget - stock - committed - planned;               // Balance Cost = เหลือต้องสั่งจริง (งบ − Stock − PO − แผน) ตรงกับหน้าจัดซื้อ
    const balPOout = budget - stock - committed;                        // PO Balance = งบ − Stock − PO ที่สั่งแล้ว (ยังไม่คิดแผน)
    const mgRow = mgM.map(mk => mCell[a.code]?.[mk] || { rec:0, po:0, poLate:false, plan:0, planLate:false });
    const pyRow = payM.map(mk => payplan[a.code]?.[mk] || 0);
    const mgTot = mgRow.reduce((s, c) => s + cellTot(c), 0);
    const pyTot = pyRow.reduce((s, x) => s + x, 0);
    return { a, budget, committed, balPO, balPOout, stock, balCost, mgRow, pyRow, mgTot, pyTot };
  }).filter(r => r.budget || r.committed || r.stock || r.mgTot || r.pyTot);

  // ── ค้นหา + เรียงลำดับตามหัวคอลัมน์ ──────────────────────────────────────────
  const aQ = aSearch.trim().toLowerCase();
  const sortVal = (r, key) => {
    switch (key) {
      case "code":      return r.a.code;
      case "name":      return r.a.name || "";
      case "budget":    return r.budget;
      case "balPO":     return r.balPO;
      case "stock":     return r.stock;
      case "balCost":   return r.balCost;
      case "mgTot":     return r.mgTot;
      case "committed": return r.committed;
      case "balPOout":  return r.balPOout;
      case "pyTot":     return r.pyTot;
      default:
        if (key.startsWith("im:")) return cellTot(r.mgRow[+key.slice(3)]);
        if (key.startsWith("pm:")) return r.pyRow[+key.slice(3)] || 0;
        return r.a.code;
    }
  };
  const shownRows = rows
    .filter(r => !aQ || r.a.code.toLowerCase().includes(aQ) || (r.a.name || "").toLowerCase().includes(aQ))
    .sort((x, y) => {
      const vx = sortVal(x, aSort.key), vy = sortVal(y, aSort.key);
      const d = (typeof vx === "string" || typeof vy === "string") ? String(vx).localeCompare(String(vy), "th") : (vx - vy);
      return aSort.dir === "asc" ? d : -d;
    });
  const toggleSort = (key) => setASort(s => s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: (key === "code" || key === "name") ? "asc" : "desc" });
  const arrow = (key) => aSort.key === key ? (aSort.dir === "asc" ? " ▲" : " ▼") : "";
  const mgColSum = (i) => shownRows.reduce((s, r) => s + cellTot(r.mgRow[i]), 0);
  const pyColSum = (i) => shownRows.reduce((s, r) => s + (r.pyRow[i] || 0), 0);
  const totOf = (pick) => shownRows.reduce((s, r) => s + pick(r), 0);

  const bCost = "#f4e9ef", bMg = "#eef3ee", bPy = "#fdf1e2", bPO = "#eaeef5";
  const cell = { border: "1px solid #d9e0ea", padding: "8px 13px", fontSize:13, whiteSpace: "nowrap" };
  const num  = { ...cell, textAlign: "right", fontVariantNumeric: "tabular-nums" };
  const hCell = (bg) => ({ ...cell, background: bg, fontWeight: 650, color: T.textSecondary, textAlign: "center", position: "sticky", top: 0 });
  // ── ตรึงคอลัมน์แรก 2 ช่อง (รหัส/ชื่อบัญชี) ให้ไม่เลื่อนหายตอนดูเดือนไกล ๆ ──────
  const COL1_W = 86;
  const stickyBody0 = { position: "sticky", left: 0, background: "#fff", zIndex: 1 };
  const stickyBody1 = { position: "sticky", left: COL1_W, background: "#fff", zIndex: 1 };
  const stickyHead0 = { left: 0, zIndex: 3 };
  const stickyHead1 = { left: COL1_W, zIndex: 3 };
  const numCell = (v, bg) => (
    <td style={{ ...num, background: bg, color: v < 0 ? T.red : (v ? T.textPrimary : T.textMuted), fontWeight: v ? 500 : 450 }}>{money(v)}{v ? usdLine(Math.abs(v), usdRate) : null}</td>
  );
  // ช่องเดือน: รับแล้ว=เขียว · PO รอเข้า=ดำ(⚠=ล่าช้า) · แผน=แดง มี * — เหมือนตารางจัดซื้อ
  const mgCell = (c, bg) => {
    const tot = cellTot(c);
    return (
      <td style={{ ...num, background: bg, color: tot ? T.textPrimary : T.textMuted, fontWeight: 600 }}>
        {!tot ? "-" : (<>
          {c.rec > 0 && <div style={{ color: T.green }}>{fmt(c.rec)}</div>}
          {c.po > 0 && <div style={{ color: c.poLate ? T.amber : T.textPrimary }}>{fmt(c.po)}{c.poLate ? " ⚠" : ""}</div>}
          {c.plan > 0 && <div style={{ color: c.planLate ? T.amber : T.red }}>{fmt(c.plan)} *{c.planLate ? "⚠" : ""}</div>}
          {usdLine(tot, usdRate)}
        </>)}
      </td>
    );
  };

  return (
    <div>
      <div style={{ fontSize:12, color: T.textMuted, marginBottom: 8 }}>
        {t("โชว์เฉพาะเดือนที่มีข้อมูล · Pending PO = งบ − Stock − PO − แผน (เหลือต้องสั่งจริง) · Balance Cost = งบ − Stock − PO (ตรงกับหน้าจัดซื้อ) · Incoming = รับแล้ว(เขียว) + PO รอเข้า(ดำ) + แผน(แดง) · Total PO = ยอดที่สั่งแล้ว","Only months with data · Pending PO = Budget − Stock − PO − Plan (real remaining to order) · Balance Cost = Budget − Stock − PO (matches Procurement) · Incoming = Received(green) + PO awaiting(black) + Plan(red) · Total PO = ordered")}
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
        {[[t("รับแล้ว","Received"), T.green, "#eafaf1"], [t("ล่าช้า ⚠","Late ⚠"), T.amber, "#fff6e6"], [t("PO รอเข้า","PO awaiting"), T.textPrimary, "#eef2f7"], [t("แผน (มี * ต่อท้าย)","Plan (with *)"), T.red, "#fdecec"]].map(([label, clr, bg]) => (
          <span key={label} style={{ display: "inline-flex", alignItems: "center", gap: 7, background: bg, border: `1.5px solid ${clr}`, borderRadius: 20, padding: "5px 12px", fontSize: 13, fontWeight: 700, color: clr }}>
            <span style={{ width: 14, height: 14, borderRadius: 4, background: clr, display: "inline-block" }}/>{label}
          </span>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
        <SearchInput value={aSearch} onChange={setASearch} placeholder={t("ค้นหา Acc. Code / ชื่อบัญชี","Search Acc. Code / account name")} width={260} big/>
        <span style={{ fontSize: 11, color: T.textMuted }}>{t("คลิกหัวคอลัมน์เพื่อเรียงลำดับ · แสดง","Click a header to sort · showing")} {shownRows.length}/{rows.length} {t("รายการ","items")}</span>
      </div>
      <div className="fatscroll" style={{ border: `1px solid ${T.cardBorder}`, borderRadius: 12 }}>
        <table style={{ borderCollapse: "collapse", width: "max-content", minWidth: "100%" }}>
          <thead>
            <tr>
              <th colSpan={6} style={{ ...hCell("#eef2f7"), textAlign: "left" }}>{t("ต้นทุน / งบประมาณ","Cost / Budget")}</th>
              <th colSpan={mgM.length + 1} style={hCell(bMg)}>{t("ของเข้า (รับ/PO/แผน)","Incoming (Recv/PO/Plan)")}</th>
              <th colSpan={payM.length + 1} style={hCell(bPy)}>{t("แผนจ่ายเงิน","Payment plan")}</th>
              <th colSpan={2} style={hCell(bPO)}>{t("สรุป PO","PO summary")}</th>
            </tr>
            <tr>
              <th onClick={()=>toggleSort("code")}      style={{ ...hCell("#f1f5f9"), ...stickyHead0, textAlign: "left", minWidth: COL1_W, cursor:"pointer", userSelect:"none" }}>Acc. Code{arrow("code")}</th>
              <th onClick={()=>toggleSort("name")}      style={{ ...hCell("#f1f5f9"), ...stickyHead1, textAlign: "left", minWidth: 190, cursor:"pointer", userSelect:"none" }}>Acc. Name{arrow("name")}</th>
              <th onClick={()=>toggleSort("budget")}    style={{ ...hCell(bCost), minWidth: 120, cursor:"pointer", userSelect:"none" }} title={t(`งบ QS = ราคาเดิม + เผื่อเศษ ${WASTE_LBL} (ของราคาเดิม) + งานเพิ่ม`,`QS budget = baseline + ${WASTE_LBL} wastage (on baseline) + additions`)}>Tender Cost<span style={{fontSize:11,fontWeight:600,opacity:0.8,marginLeft:4}}>{t(`รวมเผื่อ ${WASTE_LBL}`,`incl. ${WASTE_LBL}`)}</span>{arrow("budget")}</th>
              <th onClick={()=>toggleSort("balPO")}     style={{ ...hCell(bCost), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Balance Pending PO{arrow("balPO")}</th>
              <th onClick={()=>toggleSort("stock")}     style={{ ...hCell(bCost), minWidth: 90, cursor:"pointer", userSelect:"none" }}>Stock{arrow("stock")}</th>
              <th onClick={()=>toggleSort("balCost")}   style={{ ...hCell(bCost), minWidth: 100, cursor:"pointer", userSelect:"none" }}>Pending PO{arrow("balCost")}</th>
              {mgM.map((mk,i) => <th key={"m" + mk} onClick={()=>toggleSort("im:"+i)} style={{ ...hCell(bMg), cursor:"pointer", userSelect:"none" }}>{lbl(mk)}{arrow("im:"+i)}</th>)}
              <th onClick={()=>toggleSort("mgTot")}     style={{ ...hCell(bMg), fontWeight: 700, cursor:"pointer", userSelect:"none" }}>TOTAL{arrow("mgTot")}</th>
              {payM.map((mk,i) => <th key={"p" + mk} onClick={()=>toggleSort("pm:"+i)} style={{ ...hCell(bPy), cursor:"pointer", userSelect:"none" }}>{lbl(mk)}{arrow("pm:"+i)}</th>)}
              <th onClick={()=>toggleSort("pyTot")}     style={{ ...hCell(bPy), fontWeight: 700, cursor:"pointer", userSelect:"none" }}>TOTAL{arrow("pyTot")}</th>
              <th onClick={()=>toggleSort("committed")} style={{ ...hCell(bPO), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Total PO{arrow("committed")}</th>
              <th onClick={()=>toggleSort("balPOout")}  style={{ ...hCell(bPO), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Balance Cost{arrow("balPOout")}</th>
            </tr>
          </thead>
          <tbody>
            {shownRows.map(r => (
              <tr key={r.a.code}>
                <td style={{ ...cell, ...stickyBody0, fontVariantNumeric: "tabular-nums", fontWeight: 600 }}>{r.a.code}</td>
                <td style={{ ...cell, ...stickyBody1 }}>{r.a.name}</td>
                {numCell(r.budget, bCost)}
                {numCell(r.balPO, bCost)}
                {numCell(r.stock, bCost)}
                {numCell(r.balCost, bCost)}
                {r.mgRow.map((c, i) => <Fragment key={"m" + i}>{mgCell(c, bMg)}</Fragment>)}
                <td style={{ ...num, background: bMg, fontWeight: 650, color: T.textPrimary }}>{money(r.mgTot)}{r.mgTot ? usdLine(r.mgTot, usdRate) : null}</td>
                {r.pyRow.map((v, i) => <Fragment key={"p" + i}>{numCell(v, bPy)}</Fragment>)}
                <td style={{ ...num, background: bPy, fontWeight: 650, color: r.pyTot < 0 ? T.red : T.textPrimary }}>{money(r.pyTot)}{r.pyTot ? usdLine(r.pyTot, usdRate) : null}</td>
                {numCell(r.committed, bPO)}
                {numCell(r.balPOout, bPO)}
              </tr>
            ))}
            {shownRows.length === 0 && (
              <tr><td style={{ ...cell, textAlign: "center", color: T.textMuted }} colSpan={mgM.length + payM.length + 10}>{rows.length === 0 ? t("— ยังไม่มีข้อมูล —","— No data —") : t("— ไม่พบรายการที่ตรงกับการค้นหา —","— No matches —")}</td></tr>
            )}
          </tbody>
          {shownRows.length > 0 && (
            <tfoot>
              <tr>
                <td style={{ ...cell, ...stickyBody0, fontWeight: 700, background: "#f1f5f9" }} colSpan={2}>TOTAL</td>
                {(() => { const v = totOf(r => r.budget); return <td style={{ ...num, fontWeight: 700, background: "#eef2f7" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })()}
                {(() => { const v = totOf(r => r.balPO); return <td style={{ ...num, fontWeight: 700, background: "#eef2f7", color: v < 0 ? T.red : T.textPrimary }}>{money(v)}{v ? usdLine(Math.abs(v), usdRate) : null}</td>; })()}
                {(() => { const v = totOf(r => r.stock); return <td style={{ ...num, fontWeight: 700, background: "#eef2f7" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })()}
                {(() => { const v = totOf(r => r.balCost); return <td style={{ ...num, fontWeight: 700, background: "#eef2f7", color: v < 0 ? T.red : T.textPrimary }}>{money(v)}{v ? usdLine(Math.abs(v), usdRate) : null}</td>; })()}
                {mgM.map((mk, i) => { const v = mgColSum(i); return <td key={"tm" + mk} style={{ ...num, fontWeight: 650, background: "#e6ede6" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })}
                {(() => { const v = totOf(r => r.mgTot); return <td style={{ ...num, fontWeight: 700, background: "#e6ede6" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })()}
                {payM.map((mk, i) => { const v = pyColSum(i); return <td key={"tp" + mk} style={{ ...num, fontWeight: 650, background: "#fbe9d4" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })}
                {(() => { const v = totOf(r => r.pyTot); return <td style={{ ...num, fontWeight: 700, background: "#fbe9d4" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })()}
                {(() => { const v = totOf(r => r.committed); return <td style={{ ...num, fontWeight: 700, background: "#e2e8f2" }}>{money(v)}{v ? usdLine(v, usdRate) : null}</td>; })()}
                {(() => { const v = totOf(r => r.balPOout); return <td style={{ ...num, fontWeight: 700, background: "#e2e8f2", color: v < 0 ? T.red : T.textPrimary }}>{money(v)}{v ? usdLine(Math.abs(v), usdRate) : null}</td>; })()}
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}

// ─── Accounting View ──────────────────────────────────────────────────────────
function AccountingView({ project, updateProject, tenderCosts, additions, poEntries, onBack, onHome, onDept, onExport, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[], incomingPlan=[] }) {
  // บัญชี = อ่านอย่างเดียว (RLS ไม่ให้เขียน tcs-projects) → ปุ่มสกุลเงินจึงเป็นค่า
  // "ดูเฉพาะเครื่องนี้" ไม่บันทึกกลับไปที่โครงการร่วม กันไม่ให้บัญชีแก้ข้อมูลโครงการ
  const [curOverride, setCurOverride] = useState({});
  const curProject   = { ...project, ...curOverride };
  const setCurrency  = (fields) => setCurOverride(o => ({ ...o, ...fields }));
  const usdRate = effRate(curProject);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const [view, setView] = useState("dashboard");
  const isPhone = useIsPhone();
  const [viewHist, setViewHist] = useState([]);   // ประวัติแท็บที่ดูมาก่อน — ปุ่มกลับจะย้อนทีละหน้า
  const goView = (v) => { if (v !== view) { setViewHist(h => [...h, view]); setView(v); } };
  const backView = () => { if (viewHist.length) { const h = [...viewHist]; const prev = h.pop(); setViewHist(h); setView(prev); } else onBack(); };
  const [sortKey, setSortKey] = useState(null);  // "code" | "name" | "group" | "budget" | "committed" | "pct" | null
  const [sortDir, setSortDir] = useState(1);
  const [accShow, setAccShow] = useState(null);   // ตารางงบ: "focus" = เฉพาะหมวดที่ต้องดู · "all" · null = อัตโนมัติ (มีหมวดต้องดู → focus)
  const [payListOpen, setPayListOpen] = useState(false);   // เปิดรายการ "ต้องจ่ายใคร" (แยกตาม Supplier) ใต้แผง "ต้องจ่าย"
  useEffect(() => { setPayListOpen(false); }, [view]);   // เปลี่ยนแท็บ → พับรายการ ไม่ให้ดันตารางของแท็บใหม่ลงล่าง

  // Budget = baseline Tender Cost + every monthly addition (ค่าธรรมดา + คอลัมน์
  // ย่อย) combined per Acc. Code — ใช้ตัวช่วยกลางเดียวกับ Export ให้ตัวเลขตรงกัน
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);

  // Sum only top-level codes — see note in ProcurementView. Object.values()
  // over the whole combinedBudget double-counts sub-items (EX-xxxx rows),
  // since their value is already folded into their parent account's total.
  // ไม่ซ่อน Acc.Code ที่ยังมี PO อยู่ (กันยอด committed/งบหายจากยอดรวมหน้าบัญชี)
  const effHidden = hiddenSafeForPO(hiddenAccounts, poEntries);
  const topLevelCodes = [
    ...ACCOUNTS.filter(a => !effHidden.includes(a.code)).map(a => a.code),
    ...extraItems.filter(e => !e.parentCode).map(e => e.code),
  ];
  const tenderTotal   = topLevelCodes.reduce((s,c) => s + (parseFloat(combinedBudget[c]) || 0), 0);
  const totalComm     = poEntries.reduce((s,p)=>s+poTotal(p),0);
  // จ่ายแล้ว = ทุกงวดที่ถือว่าจ่ายแล้ว (สถานะ Paid = จ่ายทันที, หรือถึงกำหนดจ่าย)
  // ใช้เกณฑ์ roundPaid ตัวเดียวให้ตรงกับหน้าจัดซื้อและไฟล์ Excel ทุกไฟล์
  const totalPaid     = poEntries.reduce((s,p)=> s + poPaidAmount(p), 0);
  const paidPOCount   = poEntries.filter(p=>paymentStatus(p)==="paid").length;
  const totalInvoiced = poEntries.reduce((s,p)=>s+poBilledAmount(p),0);   // ของเข้าแล้ว/วางบิล (ดู poBilledAmount)
  const pct           = tenderTotal>0?(totalComm/tenderTotal*100):0;

  // รวม "งานเพิ่ม" (standalone extra) เข้าไปในกราฟตามกลุ่มด้วย ไม่งั้นยอดในกราฟ
  // จะไม่ตรงกับการ์ดสรุป (ที่นับ topLevelCodes รวม extra) — และเคารพบัญชีที่ซ่อนไว้
  const chartGroups = [...new Set([...GROUPS, ...extraItems.filter(e=>!e.parentCode).map(e=>e.group||"Other")])];
  const groupData = chartGroups.map((g,i)=>{
    const codes=[
      ...ACCOUNTS.filter(a=>a.group===g && !effHidden.includes(a.code)).map(a=>a.code),
      ...extraItems.filter(e=>!e.parentCode && (e.group||"Other")===g).map(e=>e.code),
    ];
    const committed = poEntries.reduce((s,p)=>s+poItems(p).filter(it=>codes.includes(it.code)).reduce((s2,it)=>s2+(parseFloat(it.amount)||0),0),0);
    return {group:g,budget:codes.reduce((s,c)=>s+(parseFloat(combinedBudget[c])||0),0),committed,color:GRP_COLORS[i%GRP_COLORS.length]};
  }).filter(g=>g.budget>0||g.committed>0);

  // รวมบัญชีมาตรฐาน + "งานเพิ่ม" (standalone extra ที่ไม่ใช่รายการย่อย) ให้ยอดรวม
  // หน้าบัญชีตรงกับหน้า QS/ภาพรวม ที่นับ topLevelCodes เหมือนกัน
  const acctRows = exportAccountList(extraItems, effHidden, poEntries).map(a => a.orphan ? { ...a, name: t(ORPHAN_NAME, "(code no longer in the list)") } : a);
  const accountData = acctRows.map(a=>{
    const budget=parseFloat(combinedBudget[a.code])||0;
    // Every PO line item booked to this Account Code, whether the PO is
    // single-code or split across several — pos.length still counts POs (a
    // PO with two lines on the same code only counts once).
    const items = poEntries.flatMap(p=>poItems(p).filter(it=>it.code===a.code));
    const poCount = new Set(poEntries.filter(p=>poItems(p).some(it=>it.code===a.code)).map(p=>p.id)).size;
    const committed = items.reduce((s,it)=>s+(parseFloat(it.amount)||0),0);
    return {...a,budget,committed,pos:{length:poCount},over:committed>budget&&budget>0};
  }).filter(a=>a.budget>0||a.pos.length>0);
  const pctUsedOf = (a) => a.budget>0 ? (a.committed/a.budget*100) : (a.committed>0 ? 999 : 0);

  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => -d);
    else { setSortKey(key); setSortDir(1); }
  };
  // หมวดที่ต้องดู = เกินงบ · ใช้ ≥80% · มี PO แต่ไม่มีงบ — แสดงก่อน (ตารางเต็มมีหลายสิบแถว ส่วนใหญ่ยังไม่มี PO)
  const needsLook = (a) => a.over || (a.budget>0 && a.committed/a.budget>=0.8) || (a.budget<=0 && a.committed>0);
  const lookCount = accountData.filter(needsLook).length;
  const accMode = accShow || (lookCount>0 ? "focus" : "all");
  const displayAccountData = (() => {
    if (!sortKey) return accMode==="focus" ? accountData.filter(needsLook).sort((a,b)=>pctUsedOf(b)-pctUsedOf(a)) : accountData;
    const arr = [...accountData];
    arr.sort((a, b) => {
      let av, bv;
      if (sortKey === "code")           { av = a.code; bv = b.code; }
      else if (sortKey === "group")     { av = GROUPS.indexOf(a.group); bv = GROUPS.indexOf(b.group); }
      else if (sortKey === "name")      { av = a.name; bv = b.name; }
      else if (sortKey === "budget")    { av = a.budget; bv = b.budget; }
      else if (sortKey === "committed") { av = a.committed; bv = b.committed; }
      else if (sortKey === "variance")  { av = a.budget-a.committed; bv = b.budget-b.committed; }
      else                              { av = pctUsedOf(a); bv = pctUsedOf(b); }
      if (typeof av === "string") return av.localeCompare(bv) * sortDir;
      return (av - bv) * sortDir;
    });
    return accMode==="focus" ? arr.filter(needsLook) : arr;
  })();

  // ─── แผนจ่ายเงินรายเดือน (Payment forecast) ──────────────────────────────
  // ใช้ตัวช่วย poPayLines() ตัวเดียวกับ Export เพื่อให้ตัวเลขตรงกัน และกันการนับ
  // ซ้ำเมื่อ PO มีงวดส่งของซ้ำ (ยึดยอด item.amount เป็นหลัก).
  const payToday = todayStr();
  const payLines = poEntries.flatMap(poPayLines);
  const payMonthKeys = [...new Set(payLines.map(l=>l.month||"9999-99"))].sort();
  const payByMonth = payMonthKeys.map(mk => {
    const lines  = payLines.filter(l=>(l.month||"9999-99")===mk).sort((a,b)=>(a.payDate||"9999").localeCompare(b.payDate||"9999"));
    const cash   = lines.filter(l=>l.isCash).reduce((s,l)=>s+l.amount,0);
    const credit = lines.filter(l=>!l.isCash).reduce((s,l)=>s+l.amount,0);
    const sum    = cash+credit;
    const paidA  = lines.reduce((s,l)=>s+(l.paidAmount||0),0); // รวมยอดจ่ายจริง (รองรับจ่ายบางส่วน)
    return { mk, label: mk==="9999-99"?t("ยังไม่ระบุวันจ่าย","No pay date"):monthShortLabel(mk), lines, cash, credit, sum, paid:paidA, remain:Math.max(0,sum-paidA) };
  });
  const thisMonthKey = payToday.slice(0,7);
  // "ครบกำหนดเดือนนี้" = คงเหลือของเดือนนี้ + ยอดที่เลยกำหนดจากเดือนก่อน ๆ ที่ยังไม่จ่าย
  const dueThisMonth = payByMonth.filter(m=>m.mk!=="9999-99" && m.mk<=thisMonthKey).reduce((s,m)=>s+m.remain,0);
  // เดือนถัดไป — สำหรับแจ้งเตือนให้บัญชีเตรียมเงินล่วงหน้า
  const nextMonthKey = (() => { const [y,m]=thisMonthKey.split("-").map(Number); const ny=m===12?y+1:y, nm=m===12?1:m+1; return `${ny}-${String(nm).padStart(2,"0")}`; })();
  const nextBucket   = payByMonth.find(m=>m.mk===nextMonthKey);
  const dueNextMonth = nextBucket?.remain || 0;

  // ใช้ทั้งในแผง "ต้องจ่าย" และรายการแยกตาม Supplier — สถานะมีคำกำกับเสมอ (ไม่ต้องพึ่งคำอธิบายสี)
  const payRemainOf = (l) => Math.max(0, (l.amount||0) - (l.paidAmount||0));
  const payMethodOf = (l) => l.isCash ? t("เงินสด","Cash") : t(`เครดิต ${l.term} วัน`, `Credit ${l.term}d`);
  const payStatusOf = (l) => !l.received
    ? (l.short ? [t("ยอดรับยังไม่ครบ","Received less than ordered"), T.amber, T.amberBg] : [t("ของยังไม่เข้า (ตามแผน)","Goods not in yet (plan)"), T.textSecondary, "#f1f5f9"])
    : (l.payDate && l.payDate < payToday ? [t("เลยกำหนด","Overdue"), T.red, T.redBg] : [t("รอจ่าย","To pay"), T.amber, T.amberBg]);

  const pieData = PO_STATUS.map(s=>({name:poStatusLabel(s),value:poEntries.filter(p=>p.status===s).reduce((sum,p)=>sum+poTotal(p),0),color:STATUS_CLR[s]})).filter(d=>d.value>0);

  const CT = ({active,payload}) => {
    if (!active||!payload?.length) return null;
    return (
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:10,padding:"10px 14px",fontSize:12,boxShadow:"0 4px 16px rgba(0,0,0,0.1)"}}>
        <div style={{color:T.textMuted,marginBottom:4,fontWeight:600}}>{payload[0]?.payload?.group}</div>
        {payload.map(p=><div key={p.name} style={{color:p.fill||p.color,fontVariantNumeric:"tabular-nums"}}>{p.name}: {fmt(p.value)}</div>)}
      </div>
    );
  };


  return (
    <Shell role="accounting" color={T.green} project={project} onBack={backView} onHome={onHome} onDept={onDept} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"24px 28px"}}>
        {/* Tabs + Export */}
        <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
          {!isPhone && <div className="seg-tabs" role="tablist">
          {[["dashboard",t("ภาพรวมงบ","Budget overview"),t("ภาพรวม: งบประมาณ vs ที่ผูกพันแล้ว (PO) ทั้งโครงการ","Overview: budget vs committed (PO) for the whole project")],["matrix",t("ตารางรวมเดือน","Monthly matrix"),t("ตารางรวม: ต้นทุน + Incoming Plan / Actual Received / Payment Plan รายเดือน (เฉพาะเดือนที่มีข้อมูล)","Matrix: cost + Incoming/Received/Payment per month (only months with data)")]].map(([v,l,tip])=>(
            <button key={v} role="tab" aria-selected={view===v} onClick={()=>goView(v)} title={tip} className={`seg-tab${view===v?" on":""}`}>{l}</button>
          ))}
          </div>}
          <div style={{marginLeft:"auto"}}><CurrencyControl project={curProject} updateProject={setCurrency}/></div>
          {!isPhone && (
          <button onClick={onExport} className="btn-ghost" style={{display:"flex",alignItems:"center",gap:6}}>
            <Ico name="download" /> Export Excel
          </button>
          )}
        </div>
        {isPhone && <BottomNav items={[
          { key:"dashboard", icon:"chart", label:t("ภาพรวมงบ","Budget overview"), on:view==="dashboard", onClick:()=>goView("dashboard") },
          { key:"matrix",    icon:"grid",  label:t("ตารางรวมเดือน","Monthly matrix"), on:view==="matrix", onClick:()=>goView("matrix") },
          { key:"export",    icon:"download", label:"Export", onClick:onExport },
        ]} />}

        {/* ต้องจ่าย — งานหลักของบัญชีจึงอยู่บนสุด: ยอดเดือนนี้/เดือนหน้า + รายการที่ต้องจ่ายเรียงตามวันครบกำหนด
            (เดิมเป็นแถบเหลืองไล่สี กดแล้วค่อยเห็นรายการ) · "ดูทั้งหมดแยกตาม Supplier" เปิดรายการเต็มด้านล่าง */}
        {(dueThisMonth>0 || dueNextMonth>0) && (() => {
          const soon = payLines.filter(l => l.month && l.month <= nextMonthKey && payRemainOf(l) > 0.005)
            .sort((a,b) => (a.payDate||"9999").localeCompare(b.payDate||"9999"));
          const SHOW = 5;
          const th = { padding:"8px 12px", fontSize:12, color:T.textMuted, fontWeight:600, textAlign:"left", borderBottom:`1px solid ${T.cardBorder}`, whiteSpace:"nowrap" };
          const td = { padding:"8px 12px", fontSize:13, borderBottom:"1px solid #f1f5f9", whiteSpace:"nowrap" };
          return (
          <section data-pay-panel aria-label={t("ต้องจ่าย","To pay")}
            style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:payListOpen?"14px 14px 0 0":14,padding:"16px 18px",marginBottom:payListOpen?0:20}}>
            <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
              <h2 className="pay-h" style={{margin:0,fontSize:15,fontWeight:650,color:T.textPrimary,display:"flex",alignItems:"center",gap:8}}><Ico name="wallet" size={18} color={T.textSecondary} />{t("ต้องจ่าย","To pay")}</h2>
              <div data-due="this" className="due-box" style={{background:dueThisMonth>0?T.redBg:"#f8fafc",borderRadius:10,padding:"6px 12px",minWidth:150}}>
                <div style={{fontSize:12,color:T.textSecondary}}>{t("ครบกำหนดเดือนนี้","Due this month")} · {monthShortLabel(thisMonthKey)} <span style={{color:T.textMuted}}>({t("รวมค้างจ่าย","incl. overdue")})</span></div>
                <div style={{fontSize:20,fontWeight:700,color:dueThisMonth>0?T.red:T.textPrimary,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(dueThisMonth)}</div>
                {usdLine(dueThisMonth, usdRate)}
              </div>
              <div data-due="next" className="due-box" style={{background:"#f8fafc",borderRadius:10,padding:"6px 12px",minWidth:150}}>
                <div style={{fontSize:12,color:T.textSecondary}}>{t("เตรียมเดือนหน้า","Next month")} · {monthShortLabel(nextMonthKey)}</div>
                <div style={{fontSize:20,fontWeight:700,color:T.textPrimary,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(dueNextMonth)}</div>
                {usdLine(dueNextMonth, usdRate)}
              </div>
              <div style={{flex:1}}/>
              <button onClick={()=>goView("matrix")} className="btn-ghost" style={{fontSize:12,padding:"6px 12px",whiteSpace:"nowrap"}}>{t("ตารางรวมเดือน","Monthly matrix")} →</button>
            </div>
            {!payListOpen && soon.length>0 && (
              <div style={{marginTop:14}}>
                <div style={{fontSize:13,fontWeight:650,color:T.textPrimary,marginBottom:6}}>{t("ต้องจ่ายเร็ว ๆ นี้","Coming up")} <span style={{fontWeight:400,color:T.textMuted,fontSize:12}}>· {t("เรียงตามวันครบกำหนด","by due date")}</span></div>
                <div className="hscroll"><table data-pay-soon style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr>
                    {[["ครบกำหนด","Due"],["Supplier","Supplier"],["PO No.","PO No."],["วิธีจ่าย","Method"],["สถานะ","Status"],["ยอดต้องจ่าย (THB)","To pay (THB)"]].map(([h,he],k)=>(
                      <th key={h} style={{...th, textAlign: k===5?"right":"left"}}>{t(h,he)}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {soon.slice(0,SHOW).map((l,k) => { const [st, sc, sb] = payStatusOf(l); return (
                      <tr key={k}>
                        <td style={{...td,fontVariantNumeric:"tabular-nums",color:l.payDate&&l.payDate<payToday?T.red:T.textPrimary,fontWeight:l.payDate&&l.payDate<payToday?650:400}}>{l.payDate ? fmtDate(l.payDate) : "—"}</td>
                        <td style={{...td,color:T.textPrimary}}>{l.supplier || "—"}</td>
                        <td style={{...td,color:T.textSecondary}}>{l.poNo || "—"}</td>
                        <td style={td}>{payMethodOf(l)}</td>
                        <td style={td}><span style={{background:sb,color:sc,fontSize:12,padding:"2px 8px",borderRadius:20,fontWeight:600}}>{st}</span></td>
                        <td style={{...td,textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:600}}>{fmt(payRemainOf(l))}{usdLine(payRemainOf(l), usdRate)}</td>
                      </tr>
                    ); })}
                  </tbody>
                </table></div>
              </div>
            )}
            <button onClick={()=>setPayListOpen(o=>!o)} aria-expanded={payListOpen} className="btn-ghost"
              style={{marginTop:12,fontSize:12,padding:"6px 12px",display:"inline-flex",alignItems:"center",gap:6}}>
              {payListOpen ? t("ซ่อนรายการ","Hide list") : <>{t("ดูว่าต้องจ่ายใคร","Who to pay")} · {t(`ทั้งหมด ${soon.length} งวด แยกตาม Supplier`,`all ${soon.length} rounds by supplier`)}</>}
              <Ico name="chevrons" size={14} style={{transform:payListOpen?"rotate(180deg)":"none"}} />
            </button>
          </section>
          );
        })()}
        {/* 📋 ต้องจ่ายใคร — แตกยอดในแถบ 🔔 เป็นรายงวดต่อ Supplier (ใช้ payLines ชุดเดียวกัน ยอดรวมจึงตรงกับแถบ) */}
        {payListOpen && (dueThisMonth>0 || dueNextMonth>0) && (() => {
          const remainOf = payRemainOf;
          const open = payLines.filter(l => l.month && remainOf(l) > 0.005);
          const sections = [
            { key:"now",  title: t(`ครบกำหนดเดือนนี้ + ค้างจ่าย (ถึง ${monthShortLabel(thisMonthKey)})`, `Due this month + overdue (to ${monthShortLabel(thisMonthKey)})`),
              lines: open.filter(l => l.month <= thisMonthKey), clr: T.red, bg: T.redBg },
            { key:"next", title: t(`เดือนหน้า (${monthShortLabel(nextMonthKey)})`, `Next month (${monthShortLabel(nextMonthKey)})`),
              lines: open.filter(l => l.month === nextMonthKey), clr: T.amber, bg: T.amberBg },
          ];
          const methodOf = payMethodOf, statusOf = payStatusOf;
          const th = { padding:"8px 12px", fontSize:11, color:T.textMuted, fontWeight:600, textAlign:"left", borderBottom:`1px solid ${T.cardBorder}`, whiteSpace:"nowrap" };
          const td = { padding:"7px 12px", fontSize:13, borderBottom:"1px solid #f1f5f9", whiteSpace:"nowrap" };
          return (
            <div onClick={e=>e.stopPropagation()} style={{border:`1px solid ${T.cardBorder}`,borderTop:"none",borderRadius:"0 0 14px 14px",background:T.card,padding:"8px 18px 16px",marginBottom:20}}>
              {sections.map(sec => {
                const bySup = {};
                sec.lines.forEach(l => { const k = l.supplier || "—"; (bySup[k] = bySup[k] || []).push(l); });
                const groups = Object.entries(bySup)
                  .map(([sup, ls]) => ({ sup, ls: ls.sort((a,b)=>(a.payDate||"").localeCompare(b.payDate||"")), total: ls.reduce((s,l)=>s+remainOf(l),0) }))
                  .sort((a,b) => (a.ls[0].payDate||"").localeCompare(b.ls[0].payDate||""));
                const secTotal = groups.reduce((s,g)=>s+g.total,0);
                return (
                  <div key={sec.key} style={{marginTop:12}}>
                    <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:6}}>
                      <span style={{fontSize:13,fontWeight:700,color:sec.clr}}>{sec.title}</span>
                      <span style={{fontSize:12,color:T.textMuted}}>{groups.length} {t("ราย","suppliers")} · {sec.lines.length} {t("งวด","rounds")}</span>
                      <span style={{marginLeft:"auto",fontVariantNumeric:"tabular-nums",fontWeight:700,color:sec.clr}}>฿{fmt(secTotal)}</span>
                    </div>
                    {groups.length === 0 ? (
                      <div style={{fontSize:12,color:T.textMuted,padding:"6px 0"}}>{t("ไม่มียอดที่ต้องจ่าย","Nothing to pay")}</div>
                    ) : (
                      <div className="hscroll"><table style={{width:"100%",borderCollapse:"collapse"}}>
                        <thead><tr>
                          {[["ครบกำหนด","Due"],["PO No.","PO No."],["Acc. Code","Acc. Code"],["วิธีจ่าย","Method"],["สถานะ","Status"],["ยอดต้องจ่าย (THB)","To pay (THB)"]].map(([h,he],i)=>(
                            <th key={h} style={{...th, textAlign: i===5?"right":"left"}}>{t(h,he)}</th>
                          ))}
                        </tr></thead>
                        <tbody>
                          {groups.map(g => (
                            <Fragment key={g.sup}>
                              <tr style={{background:sec.bg}}>
                                <td colSpan={5} style={{...td,fontWeight:700,color:T.textPrimary}}>{g.sup} <span style={{fontWeight:500,color:T.textMuted,fontSize:12}}>· {g.ls.length} {t("งวด","rounds")}</span></td>
                                <td style={{...td,textAlign:"right",fontWeight:700,fontVariantNumeric:"tabular-nums",color:sec.clr}}>{fmt(g.total)}</td>
                              </tr>
                              {g.ls.map((l,i) => { const [st, sc, sb] = statusOf(l); return (
                                <tr key={i}>
                                  <td style={{...td,fontVariantNumeric:"tabular-nums",color:l.payDate<payToday?T.red:T.textPrimary}}>{l.payDate ? fmtDate(l.payDate) : "—"}</td>
                                  <td style={td}>{l.poNo}</td>
                                  <td style={td}><span style={{fontVariantNumeric:"tabular-nums",color:T.blue,fontWeight:600}}>{l.code}</span> <span style={{color:T.textMuted,fontSize:12}}>{l.accName}</span></td>
                                  <td style={td}>{methodOf(l)}</td>
                                  <td style={td}><span style={{background:sb,color:sc,fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:600}}>{st}</span></td>
                                  <td style={{...td,textAlign:"right",fontVariantNumeric:"tabular-nums"}}>{fmt(remainOf(l))}{usdLine(remainOf(l), usdRate)}</td>
                                </tr>
                              ); })}
                            </Fragment>
                          ))}
                        </tbody>
                      </table></div>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })()}

        {view==="dashboard" ? (
          <>
            <div className="stat-grid" style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:24}}>
              <StatCard label={t("งบประมาณ (QS)","Budget (QS)")} value={"฿"+fmt0(tenderTotal)} thb={tenderTotal} rate={usdRate} sub={t(`เดิม + เผื่อเศษ ${WASTE_LBL} + เพิ่มรายเดือนทุกเดือน`,`Baseline + ${WASTE_LBL} wastage + all monthly additions`)} color={T.blue} icon="📋" accent={T.blueLight}/>
              <StatCard label={t("ผูกพันแล้ว (PO)","Committed (PO)")} value={"฿"+fmt0(totalComm)} thb={totalComm} rate={usdRate} progress={tenderTotal>0?pct:null}
                sub={tenderTotal>0 ? `${pct.toFixed(1)}% ${t("ของงบ","of budget")} · ${tenderTotal-totalComm<0 ? t("เกินงบ","over by") : t("คงเหลือ","left")} ฿${fmt0(Math.abs(tenderTotal-totalComm))}` : `${pct.toFixed(1)}% ${t("ของงบ","of budget")}`} color={T.amber} icon="📦" accent={T.amberBg}/>
              <StatCard label={t("วางบิลแล้ว","Invoiced")} value={"฿"+fmt0(totalInvoiced)} thb={totalInvoiced} rate={usdRate}
                sub={`${t("นับจากของที่รับแล้ว","from goods received")} · ${t("ค้างจ่าย","unpaid")} ฿${fmt0(Math.max(totalInvoiced-totalPaid,0))}`} color={T.purple} icon="🧾" accent={T.purpleBg}/>
              <StatCard label={t("ชำระแล้ว","Paid")} value={"฿"+fmt0(totalPaid)} thb={totalPaid} rate={usdRate} sub={`${paidPOCount} ${t("รายการ","items")}`} color={T.green} icon="✅" accent={T.greenBg}/>
            </div>

            {/* ต้องรีบดู: หมวดเกินงบ (ชิปนับจำนวน แบบเดียวกับหน้าจัดซื้อ) — กดแล้วเรียงตารางด้านล่างตามส่วนต่าง */}
            {(() => {
              const overCount = accountData.filter(a=>a.over).length;
              if (!overCount) return null;
              return (
                <div role="group" aria-label={t("เรื่องที่ต้องรีบดู","Needs attention")} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:20}}>
                  <span style={{fontSize:13,color:T.textSecondary,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6}}><Ico name="alert" size={16} color={T.red} />{t("ต้องรีบดู","Needs attention")}</span>
                  <button className="att-chip" data-attention="over-budget" title={t("เรียงตารางตามส่วนต่าง","Sort by variance")}
                    onClick={()=>{ setAccShow("focus"); setSortKey("variance"); setSortDir(1); document.querySelector("[data-acc-table]")?.scrollIntoView({behavior:"smooth",block:"start"}); }}>
                    <b>{overCount}</b> {t("หมวดเกินงบ","categories over budget")}
                  </button>
                </div>
              );
            })()}


            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
              <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:22}}>
                <p style={{margin:"0 0 16px",fontSize:13,color:T.textPrimary,fontWeight:600}}>{t("Budget vs Committed ตาม Group","Budget vs Committed by Group")}</p>
                {groupData.length===0
                  ? <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>{t("QS ยังไม่ได้ลง Tender Cost","QS has not entered Tender Cost")}</div>
                  : <ResponsiveContainer width="100%" height={260}>
                      <BarChart data={groupData} margin={{left:0,right:0,top:4,bottom:44}}>
                        <XAxis dataKey="group" tick={{fill:T.textMuted,fontSize:11}} angle={-30} textAnchor="end" interval={0}/>
                        <YAxis tick={{fill:T.textMuted,fontSize:11}} tickFormatter={fmtK} width={60}/>
                        <Tooltip content={<CT/>}/>
                        <Bar dataKey="budget" name="Budget" fill={T.blue} radius={[5,5,0,0]}/>
                        <Bar dataKey="committed" name="Committed" fill={T.amber} radius={[5,5,0,0]}/>
                      </BarChart>
                    </ResponsiveContainer>
                }
              </div>
              <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:22}}>
                <p style={{margin:"0 0 16px",fontSize:13,color:T.textPrimary,fontWeight:600}}>{t("สถานะ PO","PO status")}</p>
                {pieData.length===0
                  ? <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>{t("ยังไม่มี PO","No PO yet")}</div>
                  : <ResponsiveContainer width="100%" height={220}>
                      <PieChart>
                        <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={78} innerRadius={36}>
                          {pieData.map((d,i)=><Cell key={i} fill={d.color}/>)}
                        </Pie>
                        <Tooltip formatter={v=>fmt(v)} contentStyle={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:10,fontSize:11,boxShadow:"0 4px 16px rgba(0,0,0,0.08)"}}/>
                        <Legend iconType="circle" wrapperStyle={{fontSize:11,color:T.textSecondary}}/>
                      </PieChart>
                    </ResponsiveContainer>
                }
              </div>
            </div>
            <div data-acc-table style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden",marginTop:20,scrollMarginTop:16}}>
            <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",padding:"14px 16px",borderBottom:`1px solid ${T.cardBorder}`}}>
              <h2 style={{margin:0,fontSize:15,fontWeight:650,color:T.textPrimary}}>{t("งบรายหมวด","Budget by code")}</h2>
              <div className="seg-tabs" role="tablist" aria-label={t("แสดงหมวด","Show codes")}>
                <button role="tab" data-acc-show="focus" aria-selected={accMode==="focus"} className={`seg-tab${accMode==="focus"?" on":""}`} onClick={()=>setAccShow("focus")}>{t("ต้องดู","Needs a look")} <span style={{fontWeight:500,color:T.textMuted}}>{lookCount}</span></button>
                <button role="tab" data-acc-show="all" aria-selected={accMode==="all"} className={`seg-tab${accMode==="all"?" on":""}`} onClick={()=>setAccShow("all")}>{t("ทั้งหมด","All")} <span style={{fontWeight:500,color:T.textMuted}}>{accountData.length}</span></button>
              </div>
              <span style={{fontSize:12,color:T.textMuted}}>{accMode==="focus" ? t("เกินงบ · ใช้ไป 80% ขึ้นไป · มี PO แต่ไม่มีงบ — เรียงจากใช้งบมากสุด","Over · 80%+ used · PO without budget — most used first") : t("ทุกหมวดที่มีงบหรือมี PO","Every code with budget or PO")}</span>
            </div>
            <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {[
                    {label:"Acc. Code", key:"code"},
                    {label:"Account Name", key:"name"},
                    {label:"Group", key:"group"},
                    {label:t("งบประมาณ (QS)","Budget (QS)"), key:"budget"},
                    {label:t("ผูกพันแล้ว (PO)","Committed (PO)"), key:"committed"},
                    {label:t("ส่วนต่าง","Variance"), key:"variance"},
                  ].map(({label,key})=>(
                    <th key={label||"__actions"}
                      style={{padding:"11px 16px",textAlign:["budget","committed","variance"].includes(key)?"right":"left",color:sortKey===key?T.green:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                      <span onClick={()=>key&&handleSort(key)} style={{cursor:key?"pointer":"default",userSelect:"none"}}>{label}{key && sortKey===key ? (sortDir===1?" ▲":" ▼") : ""}</span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {displayAccountData.length===0 && (
                  <tr><td colSpan={6} style={{padding:"28px 16px",textAlign:"center",color:T.textMuted,fontSize:13}}>
                    {t("ไม่มีหมวดที่เกินงบหรือใกล้เต็ม","No codes over or near budget")} · <button className="btn-ghost" style={{fontSize:12,padding:"4px 10px"}} onClick={()=>setAccShow("all")}>{t("ดูทั้งหมด","Show all")}</button>
                  </td></tr>
                )}
                {displayAccountData.map((a,i)=>{
                  const variance = a.budget - a.committed;
                  // มี PO แต่ไม่มีงบ → ส่วนต่างแสดงเป็นลบเต็มจำนวน (เดิมขึ้นแค่คำว่า "ไม่มีงบ" ทำให้บวกคอลัมน์แล้วไม่เท่าแถวรวม)
                  const noBudget = a.budget<=0 && a.committed>0;
                  return (
                    <tr key={a.code} data-acc-row={a.code} style={{background:(a.over||noBudget)?"#fff5f5":i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`}}>
                      <td style={{padding:"10px 16px",color:T.blue,fontVariantNumeric:"tabular-nums",fontSize:13,fontWeight:500}}>{a.code}</td>
                      <td style={{padding:"10px 16px",color:T.textPrimary}}>{a.name}</td>
                      <td style={{padding:"10px 16px"}}>
                        {(i===0 || displayAccountData[i-1]?.group!==a.group) && <span style={{background:T.blueLight,color:T.blue,fontSize:12,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{a.group}</span>}
                      </td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:T.blue,fontWeight:500}}>{a.budget>0?fmt(a.budget):"—"}{a.budget>0&&usdLine(a.budget, usdRate)}</td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:(a.over||noBudget)?T.red:T.amber,fontWeight:(a.over||noBudget)?650:500}}>{a.committed>0?fmt(a.committed):"—"}{a.committed>0&&usdLine(a.committed, usdRate)}</td>
                      <td data-variance title={noBudget ? t("ไม่มีงบ — มี PO แต่รหัสนี้ไม่มีงบ ทั้งยอดนับเป็นส่วนเกิน","No budget — POs on a code with no budget; the whole amount counts as over") : a.over ? t("เกินงบ","Over budget") : (a.budget>0 && a.committed/a.budget>=0.8) ? t(`ใช้ไป ${Math.round(a.committed/a.budget*100)}% ของงบ`,`${Math.round(a.committed/a.budget*100)}% of budget used`) : undefined}
                        style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:variance<0?T.red:T.textSecondary,fontWeight:variance<0?650:500}}>
                        {(a.budget>0||a.committed>0)?`${variance<0?"-":""}${fmt(Math.abs(variance))}`:"—"}
                        {(a.budget>0||a.committed>0)&&usdLine(Math.abs(variance), usdRate)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
                  <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:13}}>
                    {accMode==="focus"
                      ? <>{t(`แสดง ${displayAccountData.length} จาก ${accountData.length} หมวด`,`Showing ${displayAccountData.length} of ${accountData.length}`)} · {t("ยอดรวมทุกหมวด","totals for all codes")} <button className="btn-ghost" data-acc-showall style={{fontSize:12,padding:"3px 10px",marginLeft:6}} onClick={()=>setAccShow("all")}>{t("ดูทั้งหมด","Show all")}</button></>
                      : <>{accountData.length} {t("รายการ","items")}</>}
                  </td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:T.blue,fontWeight:650,fontSize:14}}>{fmt(accountData.reduce((s,a)=>s+a.budget,0))}{usdLine(accountData.reduce((s,a)=>s+a.budget,0), usdRate)}</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:T.amber,fontWeight:650,fontSize:14}}>{fmt(accountData.reduce((s,a)=>s+a.committed,0))}{usdLine(accountData.reduce((s,a)=>s+a.committed,0), usdRate)}</td>
                  {(() => {
                    const totalVariance = accountData.reduce((s,a)=>s+(a.budget-a.committed),0);
                    return (
                      <td style={{padding:"12px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:totalVariance<0?T.red:T.textSecondary,fontWeight:650,fontSize:14}}>
                        {totalVariance<0?"-":""}{fmt(Math.abs(totalVariance))}
                        {usdLine(Math.abs(totalVariance), usdRate)}
                      </td>
                    );
                  })()}
                </tr>
              </tfoot>
            </table></div>
            </div>
          </>
        ) : view==="matrix" ? (
          <AccountingMatrixTab tenderCosts={tenderCosts} additions={additions} poEntries={poEntries} extraItems={extraItems} hiddenAccounts={hiddenAccounts} incomingPlan={incomingPlan} usdRate={usdRate} />
        ) : null}
      </div>
    </Shell>
  );
}

export { AccountingMatrixTab, AccountingView };
