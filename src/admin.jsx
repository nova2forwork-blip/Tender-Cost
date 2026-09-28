// Tender Cost — หน้า Admin (ผู้ใช้, รหัสบัญชี, Log, กู้คืนข้อมูล)
import { useCallback, useEffect, useState } from "react";
import { loadKvSnapshots, loadSeen, restoreKvSnapshot, sg, ssOrThrow } from "./supabase.js";
import { ROLE_LABELS, createUser, deleteUser, loadLogs, loadUsers, resetPassword, toggleActive } from "./auth.js";
import { ACCOUNTS, GROUPS, T, appBuild, applyAccountList, migrateAccountCodes, t, uiAlert, uiConfirm, uiLocale, uid } from "./core.jsx";
import { Ico, TopBar, UserMenu } from "./ui.jsx";

// ─── User row (admin panel) ────────────────────────────────────────────────
function UserRow({ u, onReset, onToggle, onDelete, isSelf }) {
  const [resetting, setResetting] = useState(false);
  const [pw, setPw] = useState("");
  return (
    <tr style={{borderBottom:`1px solid #f1f5f9`}}>
      <td style={{padding:"10px 14px",color:T.textPrimary,fontWeight:600}}>{u.username}{isSelf && <span style={{marginLeft:6,fontSize:11,color:T.textMuted}}>({t("คุณ","you")})</span>}</td>
      <td style={{padding:"10px 14px",color:T.textSecondary}}>{u.name}</td>
      <td style={{padding:"10px 14px"}}>
        <span style={{background:T.blueLight,color:T.blue,fontSize:11,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{ROLE_LABELS[u.role]}</span>
      </td>
      <td style={{padding:"10px 14px"}}>
        <span style={{background:u.active?T.greenBg:T.redBg,color:u.active?T.green:T.red,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>
          {u.active ? t("ใช้งานได้","Active") : t("ระงับแล้ว","Suspended")}
        </span>
      </td>
      <td style={{padding:"10px 14px"}}>
        {resetting ? (
          <div style={{display:"flex",gap:6,alignItems:"center"}}>
            <input className="input-base" type="password" autoComplete="new-password" placeholder={t("รหัสผ่านใหม่ (≥ 8 ตัว)","New password (≥ 8 chars)")} value={pw} onChange={e=>setPw(e.target.value)} style={{width:150,padding:"6px 10px"}} />
            <button className="btn-primary" style={{padding:"6px 12px"}} onClick={async ()=>{ if(pw.trim().length<8){ uiAlert(t("รหัสผ่านต้องยาวอย่างน้อย 8 ตัวอักษร","Password must be at least 8 characters")); return; } if (await onReset(u.id,pw)) { setPw(""); setResetting(false); } }}>{t("บันทึก","Save")}</button>
            <button className="btn-ghost" style={{padding:"6px 10px"}} onClick={()=>{setResetting(false);setPw("");}}>{t("ยกเลิก","Cancel")}</button>
          </div>
        ) : (
          <div style={{display:"flex",gap:8}}>
            <button className="btn-ghost" style={{padding:"6px 12px",fontSize:12}} onClick={()=>setResetting(true)}>{t("รีเซ็ตรหัส","Reset password")}</button>
            {/* ห้ามระงับบัญชีตัวเอง (แอดมินคนสุดท้ายจะล็อกทุกคนออก) · ระงับคนอื่นถามยืนยันก่อน */}
            {!(isSelf && u.active) && <button className="btn-ghost" style={{padding:"6px 12px",fontSize:12}} onClick={async ()=>{ if (u.active && !(await uiConfirm(t(`ระงับผู้ใช้ "${u.username}"? ผู้ใช้นี้จะล็อกอินไม่ได้จนกว่าจะเปิดใช้อีกครั้ง`,`Suspend "${u.username}"? They won't be able to sign in until re-enabled`), { danger: true, okLabel: t("ระงับ","Suspend") }))) return; onToggle(u.id); }}>{u.active?t("ระงับ","Suspend"):t("เปิดใช้","Enable")}</button>}
            {!isSelf && <button className="btn-ghost" style={{padding:"6px 12px",fontSize:12,color:T.red,borderColor:T.red}} onClick={async ()=>{if((await uiConfirm(t(`ลบผู้ใช้ "${u.username}" ถาวร?\n\nย้อนกลับไม่ได้ — ผู้ใช้นี้จะเข้าระบบไม่ได้อีก`,`Delete user "${u.username}" permanently?\n\nCannot be undone — this user can no longer sign in`), { danger: true, okLabel: t("ลบผู้ใช้","Delete user") }))) onDelete(u.id);}}>{t("ลบ","Delete")}</button>}
          </div>
        )}
      </td>
    </tr>
  );
}

// ─── Admin: กู้คืนข้อมูล ─────────────────────────────────────────────────────
// แสดงประวัติทุกการแก้/ลบจากตาราง kv_history (ต้องรัน kv-history.sql ก่อน)
// ให้ admin เลือกคีย์ → เลือกเวอร์ชันก่อนหน้า → กดกู้คืนกลับเข้า kv_store
// การอ่านประวัติและการเขียนคืนถูกจำกัดเฉพาะ admin ด้วย RLS ฝั่ง DB อยู่แล้ว
function AdminRestoreTab() {
  const [snaps, setSnaps]     = useState([]);
  const [projMap, setProjMap] = useState({});
  const [dept, setDept]       = useState("all"); // กรองตามแผนก
  const [loading, setLoading] = useState(true);
  const [busy, setBusy]       = useState(false);
  const [msg, setMsg]         = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const [s, projs] = await Promise.all([loadKvSnapshots(), sg("tcs-projects")]);
      const map = {}; (projs || []).forEach(p => { map[p.id] = p.name; });
      setProjMap(map); setSnaps(s || []);
    } catch (e) {
      // เดิมไม่มี catch → ค้างที่ "กำลังโหลด" ตลอด และหลังกู้คืนปุ่มค้างสถานะ busy
      console.warn("โหลดสแนปช็อตไม่สำเร็จ:", e);
      setMsg(t("⚠ โหลดรายการสแนปช็อตไม่สำเร็จ — ตรวจเน็ต/สิทธิ์ แล้วลองเปิดหน้านี้ใหม่","⚠ Couldn't load snapshots — check connection/permissions and reopen this tab"));
    } finally {
      setLoading(false);
    }
  }, []);
  useEffect(() => { load(); }, [load]);

  const keyLabel = (key) => {
    if (key === "tcs-projects") return t("📁 รายชื่อโครงการ","📁 Project list");
    const m = key.match(/^tcs-(tenders|po|additions|extra|hidden|inplan)-(.+)$/);
    if (m) {
      const kt = { tenders:"Tender Cost", po:t("PO / จัดซื้อ","PO / Procurement"), additions:t("ยอดเพิ่มรายเดือน","Monthly additions"), extra:t("รายการเพิ่ม","Extra items"), hidden:t("หมวดที่ซ่อน","Hidden categories"), inplan:t("แผนของเข้า","Incoming plan") }[m[1]] || m[1];
      return `${kt} — ${projMap[m[2]] || m[2]}`;
    }
    if (key === "tcs-users") return t("ผู้ใช้ (คีย์เก่า)","Users (legacy key)");
    if (key === "tcs-logs")  return t("Log (คีย์เก่า)","Log (legacy key)");
    return key;
  };

  // จับคู่คีย์ข้อมูล → แผนกเจ้าของ
  const deptOf = (key) => {
    if (/^tcs-(tenders|additions|extra|hidden|columns)-/.test(key)) return "qs";
    if (/^tcs-(po|inplan)-/.test(key)) return "procurement";
    return "central"; // tcs-projects, tcs-users, tcs-logs, อื่น ๆ
  };
  const DEPTS = [["all",t("ทั้งหมด","All")],["qs","QS"],["procurement",t("จัดซื้อ","Procurement")],["central",t("ส่วนกลาง","Central")]];
  const deptTag = { qs:{label:"QS",color:T.blue,bg:T.blueLight}, procurement:{label:t("จัดซื้อ","Procurement"),color:T.amber,bg:T.amberBg}, central:{label:t("ส่วนกลาง","Central"),color:T.purple,bg:T.purpleBg} };

  // นับจำนวนไฟล์ข้อมูล (คีย์) ต่อแผนก — ใช้โชว์บนแท็บ
  const allKeys = [...new Set(snaps.map(s => s.key))];
  const deptCount = { all: allKeys.length, qs:0, procurement:0, central:0 };
  allKeys.forEach(k => { deptCount[deptOf(k)] = (deptCount[deptOf(k)]||0) + 1; });

  // สแนปช็อตเฉพาะแผนกที่เลือก แล้วรวมเป็น "รอบ" (วันเดียวกัน + รอบเวลาเดียวกัน = 1 รอบ)
  // แต่ละรอบเก็บเวอร์ชันล่าสุดของแต่ละคีย์ในรอบนั้น เพื่อกู้คืนทั้งชุดในคลิกเดียว
  const deptSnaps = snaps.filter(s => dept === "all" || deptOf(s.key) === dept);
  const roundMap = {};
  deptSnaps.forEach(r => {
    const rk = `${new Date(r.taken_at).toDateString()}|${r.slot}`;
    const g = roundMap[rk] || (roundMap[rk] = { rk, slot:r.slot, taken_at:r.taken_at, byKey:{} });
    const ex = g.byKey[r.key];
    if (!ex || r.taken_at > ex.taken_at) g.byKey[r.key] = r;
    if (r.taken_at > g.taken_at) g.taken_at = r.taken_at;
  });
  const rounds = Object.values(roundMap).sort((a,b) => b.taken_at.localeCompare(a.taken_at));
  const roundDateLabel = (r) => new Date(r.taken_at).toLocaleDateString(uiLocale(),{weekday:"short",day:"numeric",month:"short",year:"numeric"});
  const deptLabelOf = (id) => (DEPTS.find(([d])=>d===id)||[])[1] || t("ข้อมูล","Data");

  // กู้คืนทั้งชุดของแผนกที่เลือก กลับไปยังรอบเวลาที่กด — ย้อนทุกไฟล์พร้อมกัน
  const doRestoreRound = async (round) => {
    const rows = Object.values(round.byKey);
    const dl = deptLabelOf(dept);
    if (!(await uiConfirm(t(`กู้คืน "${dl}" ทั้งชุด (${rows.length} รายการ)\nกลับเป็นสแนปช็อต ${roundDateLabel(round)} · ${round.slot}?\n\nข้อมูลปัจจุบันของทุกไฟล์ในชุดนี้จะถูกแทนที่ด้วยข้อมูลจากรอบที่เลือก`,`Restore the whole "${dl}" set (${rows.length} items)\nback to snapshot ${roundDateLabel(round)} · ${round.slot}?\n\nCurrent data for every file in this set will be replaced with the selected round`), { okLabel: t("กู้คืน","Restore") }))) return;
    setBusy(true); setMsg("");
    let ok = 0, fail = 0;
    for (const row of rows) {
      try { await restoreKvSnapshot(row); ok++; }
      catch (e) { fail++; }
    }
    setMsg(fail === 0
      ? t(`✅ กู้คืน ${dl} สำเร็จ ${ok} รายการ — กลับไปหน้าหลักเพื่อดูข้อมูลที่กู้คืน`,`✅ Restored ${dl}: ${ok} items — go back to the main page to see restored data`)
      : t(`⚠️ กู้คืนสำเร็จ ${ok} รายการ · ไม่สำเร็จ ${fail} รายการ`,`⚠️ Restored ${ok} items · failed ${fail} items`));
    await load();
    setBusy(false);
  };

  if (loading) return <div style={{color:T.textMuted,fontSize:13}}>{t("กำลังโหลดสแนปช็อต...","Loading snapshots...")}</div>;

  if (!snaps.length) return (
    <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:24,fontSize:13,color:T.textSecondary,lineHeight:1.7}}>
      {t("ยังไม่มีสแนปช็อต","No snapshots yet")}<br/>
      <span style={{color:T.textMuted}}>{t("ระบบจะถ่ายสแนปช็อตอัตโนมัติวันละ 2 รอบ (12:00 และ 18:00) หลังจากรันไฟล์","Auto snapshots run twice a day (12:00 and 18:00) after running")} <b>kv-snapshots.sql</b> {t("ใน Supabase","in Supabase")}</span>
    </div>
  );

  return (
    <div>
      {msg && (
        <div style={{marginBottom:14,padding:"10px 14px",borderRadius:10,fontSize:13,fontWeight:600,
          background:msg.startsWith("✅")?T.greenBg:msg.startsWith("⚠️")?T.amberBg:T.redBg,
          color:msg.startsWith("✅")?T.green:msg.startsWith("⚠️")?T.amber:T.red}}>{msg}</div>
      )}
      <div style={{fontSize:12,color:T.textMuted,marginBottom:12}}>
        {t('สแนปช็อตอัตโนมัติวันละ 2 รอบ — 12:00 และ 18:00 · เลือกแผนก แล้วกด "กู้คืนทั้งชุด" กลับไปยังรอบเวลาที่ต้องการ — ทุกไฟล์ของแผนกนั้นจะย้อนกลับพร้อมกัน','Auto snapshots twice a day — 12:00 and 18:00 · pick a department and press "Restore set" to roll back to a chosen round — all its files roll back together')}
      </div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
        {DEPTS.map(([id,label])=>{
          const active = dept===id;
          return (
            <button key={id} onClick={()=>setDept(id)}
              style={{padding:"7px 16px",borderRadius:999,border:`1.5px solid ${active?T.blue:T.cardBorder}`,cursor:"pointer",fontSize:12,fontWeight:600,
                background:active?T.blue:T.card,color:active?"#fff":T.textSecondary}}>
              {label} <span style={{opacity:0.7,fontWeight:500}}>({deptCount[id]||0})</span>
            </button>
          );
        })}
      </div>

      {!rounds.length ? (
        <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:24,fontSize:13,color:T.textMuted}}>
          {t("แผนก","Department")} "{deptLabelOf(dept)}" {t("ยังไม่มีสแนปช็อตให้กู้คืน","has no snapshots to restore")}
        </div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div style={{fontSize:12,fontWeight:650,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>
            {t("รอบสแนปช็อตของ","Snapshot rounds of")} {deptLabelOf(dept)} ({rounds.length} {t("รอบ","rounds")})
          </div>
          {rounds.map(round => {
            const rows = Object.values(round.byKey);
            return (
              <div key={round.rk} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px"}}>
                <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                  <span style={{flexShrink:0,fontSize:11,fontWeight:650,padding:"3px 11px",borderRadius:8,background:T.blueLight,color:T.blue}}>{round.slot}</span>
                  <div style={{fontSize:14,fontWeight:650,color:T.textPrimary}}>{roundDateLabel(round)}</div>
                  <span style={{fontSize:12,color:T.textMuted}}>· {rows.length} {t("ไฟล์ในชุดนี้","files in this set")}</span>
                  <div style={{flex:1,minWidth:12}}/>
                  <button onClick={()=>doRestoreRound(round)} disabled={busy}
                    className="btn-primary" style={{flexShrink:0,padding:"8px 18px",fontSize:13,opacity:busy?0.5:1,cursor:busy?"default":"pointer"}}>
                    ↩︎ {t("กู้คืนทั้งชุด","Restore set")} ({rows.length})
                  </button>
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:12}}>
                  {rows.map(r => {
                    const dt = deptTag[deptOf(r.key)];
                    return (
                      <span key={r.id} style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:11,padding:"3px 9px",borderRadius:7,background:T.bg,border:`1px solid ${T.cardBorder}`,color:T.textSecondary,maxWidth:260}}>
                        <span style={{flexShrink:0,fontSize:11,fontWeight:650,padding:"0 5px",borderRadius:4,background:dt.bg,color:dt.color}}>{dt.label}</span>
                        <span style={{whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{keyLabel(r.key)}</span>
                      </span>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── Admin: จัดการรหัสบัญชี (แก้รหัส+ชื่อ, ห้ามซ้ำ, ย้ายข้อมูลให้) ─────────────
function AdminAccountsTab() {
  const [rows, setRows] = useState(() => ACCOUNTS.map(a => ({ rid: uid(), code: a.code, name: a.name, group: a.group, orig: a.code })));
  const [busy, setBusy] = useState(false);
  const [msg, setMsg]   = useState("");
  const groupOptions = [...new Set([...GROUPS, ...rows.map(r => r.group).filter(Boolean), "Other"])];
  const setCell = (rid, k, v) => setRows(rs => rs.map(r => r.rid === rid ? { ...r, [k]: v } : r));
  const addRow  = () => setRows(rs => [...rs, { rid: uid(), code: "", name: "", group: "Other", orig: "" }]);
  const delRow  = (rid) => setRows(rs => rs.filter(r => r.rid !== rid));
  const dupCodes = (() => { const seen = {}, dup = new Set(); rows.forEach(r => { const c = (r.code || "").trim(); if (!c) return; if (seen[c]) dup.add(c); seen[c] = 1; }); return dup; })();

  const save = async () => {
    const clean = rows.map(r => ({ ...r, code: (r.code || "").trim(), name: (r.name || "").trim(), group: (r.group || "Other").trim() }));
    if (clean.some(r => !r.code)) { setMsg(t("⚠ มีรหัสว่าง — กรอกรหัสให้ครบ","⚠ Some codes are empty — fill them in")); return; }
    if (clean.some(r => !r.name)) { setMsg(t("⚠ มีชื่อว่าง — กรอกชื่อให้ครบ","⚠ Some names are empty — fill them in")); return; }
    const codes = clean.map(r => r.code);
    if (new Set(codes).size !== codes.length) { setMsg(t("⚠ มีรหัสซ้ำกัน — Acc code ห้ามซ้ำ","⚠ Duplicate codes — each Acc. Code must be unique")); return; }
    const renameMap = {}; clean.forEach(r => { if (r.orig && r.orig !== r.code) renameMap[r.orig] = r.code; });
    const nRen = Object.keys(renameMap).length;
    if (!(await uiConfirm((t(`บันทึกรายการบัญชี ${clean.length} รายการ?`,`Save ${clean.length} account codes?`)) + (nRen ? t(`\n\nเปลี่ยนรหัส ${nRen} รายการ — ระบบจะย้ายข้อมูลเดิม (Tender Cost / PO / รายเดือน / แผน) ของทุกโครงการให้อัตโนมัติ`,`\n\n${nRen} codes changed — existing data (Tender Cost / PO / monthly / plan) for all projects will be migrated automatically`) : "") + t(`\n\nเสร็จแล้วหน้าจะรีเฟรชใหม่`,`\n\nThe page will refresh when done`), { okLabel: t("บันทึก","Save") }))) return;
    setBusy(true); setMsg(t("กำลังบันทึก…","Saving…"));
    try {
      if (nRen) { setMsg(t("กำลังย้ายข้อมูลข้ามทุกโครงการ…","Migrating data across all projects…")); await migrateAccountCodes(renameMap); }
      const list = clean.map(r => ({ code: r.code, name: r.name, group: r.group }));
      await ssOrThrow("tcs-accounts", list);   // OrThrow: บันทึกพลาดต้องแจ้ง (เดิมเงียบแล้วรีเฟรชหน้าเหมือนสำเร็จ)
      applyAccountList(list);
      setMsg(t("✓ บันทึกเรียบร้อย กำลังรีเฟรช…","✓ Saved, refreshing…"));
      setTimeout(() => { if (typeof window !== "undefined") window.location.reload(); }, 700);
    } catch (e) { console.warn("save accounts failed", e); setBusy(false); setMsg(t("บันทึกไม่สำเร็จ: ","Save failed: ") + (e.message || t("ลองใหม่อีกครั้ง","try again"))); }
  };

  const inp = { padding: "6px 8px", fontSize:12, border: `1px solid ${T.cardBorder}`, borderRadius: 8, width: "100%", fontFamily: "inherit" };
  return (
    <div style={{ background: T.card, border: `1px solid ${T.cardBorder}`, borderRadius: 14, overflow: "hidden" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 18px", borderBottom: `1px solid ${T.cardBorder}`, flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 600, color: T.textPrimary }}>{t("รหัสบัญชีทั้งหมด","All account codes")} ({rows.length})</div>
          <div style={{ fontSize:11, color: T.textMuted, marginTop: 2 }}>{t("แก้รหัส/ชื่อได้ · รหัสห้ามซ้ำ · เปลี่ยนรหัสแล้วระบบย้ายข้อมูลเดิมให้ทุกโครงการ · ใช้ร่วมกันทุกโครงการ","Edit code/name · codes must be unique · changing a code migrates existing data across projects · shared by all projects")}</div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={addRow} className="btn-ghost">+ {t("เพิ่มรหัส","Add code")}</button>
          <button onClick={save} disabled={busy || dupCodes.size > 0} className="btn-primary" style={{ background: dupCodes.size ? T.textMuted : T.blue }}>{busy ? t("กำลังบันทึก…","Saving…") : t("บันทึก","Save")}</button>
        </div>
      </div>
      {msg && <div style={{ padding: "10px 18px", fontSize:12, fontWeight: 600, color: msg.startsWith("✓") ? T.green : (msg.startsWith("⚠") || msg.startsWith("บันทึกไม่") || msg.startsWith("Save failed")) ? T.red : T.textSecondary, background: "#f8fafc" }}>{msg}</div>}
      <div style={{ maxHeight: "58vh", overflow: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize:12 }}>
          <thead>
            <tr>
              {[["Acc. Code","Acc. Code"], ["ชื่อ / คำอธิบาย","Name / description"], ["กลุ่ม","Group"], ["",""]].map(([h,he], i) => (
                <th key={h + i} style={{ position: "sticky", top: 0, background: "#f1f5f9", textAlign: "left", padding: "9px 12px", fontSize: 11, color: T.textMuted, fontWeight: 650, borderBottom: `1px solid ${T.cardBorder}`, width: h === "" ? 40 : (h === "Acc. Code" ? 130 : (h === "กลุ่ม" ? 160 : "auto")) }}>{t(h,he)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(r => {
              const isDup = dupCodes.has((r.code || "").trim());
              return (
                <tr key={r.rid} style={{ borderBottom: `1px solid ${T.cardBorder}` }}>
                  <td style={{ padding: "5px 10px" }}>
                    <input value={r.code} onChange={e => setCell(r.rid, "code", e.target.value)} style={{ ...inp, fontVariantNumeric: "tabular-nums", fontWeight: 600, borderColor: isDup ? T.red : T.cardBorder, background: isDup ? T.redBg : "#fff" }} />
                    {r.orig && r.orig !== (r.code || "").trim() && <div style={{ fontSize:11, color: T.amber, marginTop: 2 }}>{t("เดิม","was")} {r.orig}</div>}
                  </td>
                  <td style={{ padding: "5px 10px" }}><input value={r.name} onChange={e => setCell(r.rid, "name", e.target.value)} style={inp} /></td>
                  <td style={{ padding: "5px 10px" }}>
                    <input list="acc-groups" value={r.group} onChange={e => setCell(r.rid, "group", e.target.value)} style={inp} />
                  </td>
                  <td style={{ padding: "5px 10px", textAlign: "center" }}>
                    <button onClick={() => delRow(r.rid)} title={t("ลบรหัสนี้","Delete this code")} aria-label={t("ลบรหัสนี้","Delete this code")} className="icon-danger" style={{ background: "none", border: "none", color: T.textMuted, cursor: "pointer", padding: 6, borderRadius: 8, display: "inline-grid", placeItems: "center" }}><Ico name="trash" size={16} /></button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        <datalist id="acc-groups">{groupOptions.map(g => <option key={g} value={g} />)}</datalist>
      </div>
      {dupCodes.size > 0 && <div style={{ padding: "10px 18px", fontSize: 12, color: T.red, fontWeight: 600 }}>⚠ {t("มีรหัสซ้ำ","Duplicate codes")}: {[...dupCodes].join(", ")} — {t("แก้ให้ไม่ซ้ำก่อนบันทึก","make them unique before saving")}</div>}
    </div>
  );
}

// ─── Admin Panel ────────────────────────────────────────────────────────────
// ─── ใครเปิดแอปเวอร์ชันไหน (คีย์ tcs-seen-<user id> ที่แอปบันทึกตอนเข้าใช้) ─────────────
function AdminVersionsTab() {
  const [rows, setRows] = useState(null);
  const load = useCallback(async () => { setRows(null); setRows(await loadSeen()); }, []);
  useEffect(() => { load(); }, [load]);
  const me = appBuild();
  const list = (rows || []).slice().sort((a, b) => String(b.at || b.updated_at || "").localeCompare(String(a.at || a.updated_at || "")));
  const same = (r) => r.build === me.build && r.version === me.version;
  const nOld = list.filter(r => !same(r)).length;
  const th = { padding:"10px 14px", textAlign:"left", color:T.textMuted, fontWeight:600, fontSize:11, letterSpacing:0.6, textTransform:"uppercase", borderBottom:`1px solid ${T.cardBorder}` };
  const td = { padding:"9px 14px", fontSize:13, borderBottom:"1px solid #f1f5f9", fontVariantNumeric:"tabular-nums" };
  return (
    <div data-versions-tab style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
      <div style={{padding:"14px 18px",borderBottom:`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
        <div style={{fontSize:13,fontWeight:650,color:T.textPrimary}}>{t("เวอร์ชันที่แต่ละคนเปิดอยู่","Version each person is using")}</div>
        <div style={{fontSize:12,color:T.textMuted}}>{t("ของคุณ","Yours")}: <b style={{color:T.textPrimary}}>v{me.version} · build {me.build}</b>{me.commit ? ` · ${me.commit}` : ""}</div>
        {rows && list.length > 0 && (nOld === 0
          ? <span data-versions-ok style={{background:T.greenBg,color:T.green,fontSize:12,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{t("ทุกคนใช้เวอร์ชันเดียวกัน","Everyone is on the same version")}</span>
          : <span data-versions-old style={{background:T.amberBg,color:"#92400e",fontSize:12,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{t(`${nOld} คนยังใช้เวอร์ชันอื่น — ให้รีเฟรช`,`${nOld} on another version — ask them to refresh`)}</span>)}
        <button className="btn-ghost" onClick={load} style={{marginLeft:"auto",fontSize:12,padding:"5px 12px"}}>{t("โหลดใหม่","Reload")}</button>
      </div>
      {rows === null ? (
        <div style={{padding:24,color:T.textMuted,fontSize:13}}>{t("กำลังโหลด...","Loading...")}</div>
      ) : list.length === 0 ? (
        <div style={{padding:24,color:T.textMuted,fontSize:13,lineHeight:1.6}}>
          {t("ยังไม่มีข้อมูล — ต้องรันไฟล์ tender-cost-version.sql ใน Supabase ครั้งเดียว แล้วให้แต่ละคนเปิดแอปใหม่","No data yet — run tender-cost-version.sql in Supabase once, then have each person reopen the app")}
        </div>
      ) : (
        <div className="hscroll"><table style={{width:"100%",minWidth:640,borderCollapse:"collapse"}}>
          <thead><tr style={{background:"#f8fafc"}}>
            {[["ผู้ใช้","User"],["แผนก","Department"],["เวอร์ชัน","Version"],["Build","Build"],["เปิดล่าสุด","Last opened"]].map(([h,he]) => <th key={h} style={th}>{t(h,he)}</th>)}
          </tr></thead>
          <tbody>
            {list.map(r => (
              <tr key={r.key} data-seen-row data-same={same(r) ? "1" : "0"} style={{background:same(r) ? "transparent" : "#fffbeb"}}>
                <td style={{...td,fontWeight:600,color:T.textPrimary}}>{r.name || "—"}</td>
                <td style={{...td,color:T.textSecondary}}>{ROLE_LABELS[r.role] || r.role || "—"}</td>
                <td style={td}>{r.version ? `v${r.version}` : "—"}</td>
                <td style={{...td,color:same(r) ? T.green : "#92400e",fontWeight:650}}>{r.build || "—"}{r.commit ? <span style={{color:T.textMuted,fontWeight:400}}> · {r.commit}</span> : null}{!same(r) && <span style={{marginLeft:8,fontSize:11}}>{t("ไม่ตรง","differs")}</span>}</td>
                <td style={{...td,color:T.textSecondary,fontSize:12}}>{(r.at || r.updated_at) ? new Date(r.at || r.updated_at).toLocaleString(uiLocale()) : "—"}</td>
              </tr>
            ))}
          </tbody>
        </table></div>
      )}
    </div>
  );
}

function AdminPanel({ onBack, onLogout, session }) {
  const [tab,   setTab]   = useState("users");
  const [users, setUsers] = useState([]);
  const [logs,  setLogs]  = useState([]);
  const [loaded, setLoadedU] = useState(false);
  const [addOpen, setAddOpen] = useState(false);
  const [draft, setDraft] = useState({ username:"", name:"", role:"qs", password:"" });
  const [err, setErr] = useState("");

  const refresh = useCallback(async () => {
    try {
      const [u, l] = await Promise.all([loadUsers(), loadLogs()]);
      setUsers(u); setLogs(l);
    } catch (e) {
      // ยังไม่ได้ deploy Edge Function admin-users → แสดงหน้าเปล่าแทนจอขาว
      console.warn("โหลดผู้ใช้/log ไม่สำเร็จ (ยังไม่ได้ deploy edge function admin-users?)", e);
    } finally {
      setLoadedU(true);
    }
  }, []);
  useEffect(() => { refresh(); }, [refresh]);

  // การกระทำกับผู้ใช้ — ถ้าล้มเหลว (เน็ต/สิทธิ์/Edge Function) แจ้งให้รู้ (เดิมเงียบ ดูเหมือนสำเร็จ)
  const adminAct = async (fn, failMsg) => {
    try { setUsers(await fn()); return true; }
    catch (e) { console.warn(failMsg, e); uiAlert(`${failMsg}\n\n${e?.message || ""}`.trim()); return false; }
  };
  const handleReset   = (id, pw) => adminAct(() => resetPassword(users, id, pw), t("⚠ รีเซ็ตรหัสผ่านไม่สำเร็จ","⚠ Password reset failed"));
  const handleToggle  = (id)     => adminAct(() => toggleActive(users, id),      t("⚠ เปลี่ยนสถานะผู้ใช้ไม่สำเร็จ","⚠ Couldn't change user status"));
  const handleDelete  = (id)     => adminAct(() => deleteUser(users, id),        t("⚠ ลบผู้ใช้ไม่สำเร็จ","⚠ Couldn't delete user"));
  const handleCreate  = async () => {
    setErr("");
    if (!draft.username.trim() || !draft.password) { setErr(t("กรอก Username และ Password","Enter Username and Password")); return; }
    if (draft.password.trim().length < 8) { setErr(t("รหัสผ่านต้องยาวอย่างน้อย 8 ตัวอักษร","Password must be at least 8 characters")); return; }
    try {
      const next = await createUser(draft);
      setUsers(next); setAddOpen(false); setDraft({ username:"", name:"", role:"qs", password:"" });
    } catch (e) { setErr(e.message || t("สร้างผู้ใช้ไม่สำเร็จ","Failed to create user")); }
  };

  return (
    <div style={{minHeight:"100vh",background:T.bg}}>
      <TopBar onBack={onBack} title="Admin" sub={t("ผู้ใช้ · รหัสบัญชี · Log · กู้คืนข้อมูล","Users · account codes · log · restore")}>
        <UserMenu session={session} onLogout={onLogout} />
      </TopBar>

      <div style={{padding:"28px 32px"}}>
        <div style={{display:"flex",gap:8,marginBottom:22}}>
          {[["users",t("จัดการผู้ใช้","Manage users")],["accounts",t("รหัสบัญชี","Account codes")],["logs",t("Log การเข้าใช้งาน","Access log")],["versions",t("เวอร์ชันที่ใช้","Versions in use")],["restore",t("กู้คืนข้อมูล","Restore data")]].map(([id,label])=>(
            <button key={id} onClick={()=>setTab(id)}
              style={{background:tab===id?T.blue:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.blue:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {label}
            </button>
          ))}
        </div>

        {!loaded ? (
          <div style={{color:T.textMuted,fontSize:13}}>{t("กำลังโหลด...","Loading...")}</div>
        ) : tab === "restore" ? (
          <AdminRestoreTab />
        ) : tab === "versions" ? (
          <AdminVersionsTab />
        ) : tab === "accounts" ? (
          <AdminAccountsTab />
        ) : tab === "users" ? (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"16px 18px",borderBottom:`1px solid ${T.cardBorder}`}}>
              <div style={{fontSize:13,fontWeight:600,color:T.textPrimary}}>{t("ผู้ใช้ทั้งหมด","All users")} ({users.length})</div>
              <button className="btn-primary" onClick={()=>setAddOpen(v=>!v)}>+ {t("เพิ่มผู้ใช้","Add user")}</button>
            </div>
            {addOpen && (
              <div style={{padding:18,borderBottom:`1px solid ${T.cardBorder}`,background:"#fafbfd"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr auto",gap:10,alignItems:"end"}}>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>Username</span>
                    <input className="input-base" value={draft.username} onChange={e=>setDraft(d=>({...d,username:e.target.value}))} />
                  </label>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>{t("ชื่อที่แสดง","Display name")}</span>
                    <input className="input-base" value={draft.name} onChange={e=>setDraft(d=>({...d,name:e.target.value}))} />
                  </label>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>{t("แผนก","Department")}</span>
                    <select className="input-base" value={draft.role} onChange={e=>setDraft(d=>({...d,role:e.target.value}))}>
                      <option value="qs">QS</option>
                      <option value="procurement">{t("จัดซื้อ","Procurement")}</option>
                      <option value="accounting">{t("บัญชี","Accounting")}</option>
                      <option value="admin">Admin</option>
                    </select>
                  </label>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>Password (≥ 8 {t("ตัว","chars")})</span>
                    <input className="input-base" type="password" autoComplete="new-password" value={draft.password} onChange={e=>setDraft(d=>({...d,password:e.target.value}))} />
                  </label>
                  <button className="btn-primary" onClick={handleCreate}>{t("สร้าง","Create")}</button>
                </div>
                {err && <div style={{color:T.red,fontSize:12,marginTop:8,fontWeight:500}}>{err}</div>}
              </div>
            )}
            <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {[["Username","Username"],["ชื่อที่แสดง","Display name"],["แผนก","Department"],["สถานะ","Status"],["",""]].map(([h,he])=>(
                    <th key={h} style={{padding:"10px 14px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{t(h,he)}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {users.map(u => (
                  <UserRow key={u.id} u={u} onReset={handleReset} onToggle={handleToggle} onDelete={handleDelete} isSelf={u.id===session.id} />
                ))}
              </tbody>
            </table></div>
          </div>
        ) : (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
            <div style={{padding:"16px 18px",borderBottom:`1px solid ${T.cardBorder}`,fontSize:13,fontWeight:600,color:T.textPrimary}}>
              {t("ประวัติการเข้าใช้งานล่าสุด","Recent access history")} ({logs.length})
            </div>
            <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {[["เวลา","Time"],["Username","Username"],["แผนก","Department"],["ผลลัพธ์","Result"]].map(([h,he])=>(
                    <th key={h} style={{padding:"10px 14px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{t(h,he)}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {logs.length===0 ? (
                  <tr><td colSpan={4} style={{padding:"30px",textAlign:"center",color:T.textMuted}}>{t("ยังไม่มีข้อมูล","No data")}</td></tr>
                ) : logs.map(l => (
                  <tr key={l.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                    <td style={{padding:"9px 14px",fontVariantNumeric:"tabular-nums",fontSize:12,color:T.textSecondary}}>{new Date(l.time).toLocaleString(uiLocale())}</td>
                    <td style={{padding:"9px 14px",color:T.textPrimary,fontWeight:500}}>{l.username}</td>
                    <td style={{padding:"9px 14px",color:T.textSecondary}}>{ROLE_LABELS[l.role]||l.role}</td>
                    <td style={{padding:"9px 14px"}}>
                      <span style={{background:l.result==="success"?T.greenBg:T.redBg,color:l.result==="success"?T.green:T.red,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>
                        {l.result==="success"?t("สำเร็จ","Success"):l.result==="inactive"?t("บัญชีถูกระงับ","Suspended"):t("ล้มเหลว","Failed")}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table></div>
          </div>
        )}
      </div>
    </div>
  );
}

export { UserRow, AdminRestoreTab, AdminAccountsTab, AdminVersionsTab, AdminPanel };
