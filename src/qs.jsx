// Tender Cost — แผนก QS (ราคาเดิม + รายการเพิ่มรายเดือน)
import { Fragment, useCallback, useEffect, useLayoutEffect, useRef, useState } from "react";
import { Bar, BarChart, CartesianGrid, Cell, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { ACCOUNTS, GROUPS, T, UnsavedGuard, WASTE_LBL, _LANG, buildCombinedBudget, codeText, extraCodeError, historyEntry, renameProjectCode, fmt, fmt0, fmtK, fmtZ, hiddenSafeForPO, inThai, leaveIfDirty, monthAddValue, monthShortLabel, t, todayStr, uiAlert, uiConfirm, uid, wasteOf, withWaste } from "./core.jsx";
import { exportQSMonthExcel } from "./excel.js";
import { BottomNav, CurrencyControl, GroupFilter, Ico, MoneyInput, SearchInput, Shell, StatCard, effRate, evalMoney, usdLine, useIsPhone } from "./ui.jsx";

// ─── QS View ─────────────────────────────────────────────────────────────────
function QSView({ project, updateProject, tenderCosts, saveTenders, poEntries = [], savePO, incomingPlan = [], saveIncomingPlan, additions, saveAdditions, extraItems, saveExtraItems, hiddenAccounts, saveHiddenAccounts, onBack, onHome, onDept, syncedAt, syncing, session, onLogout, onExport, runExportFn, setEditMode }) {
  const [tab, setTab] = useState("baseline"); // "baseline" | "monthly"
  const isPhone = useIsPhone();
  const [tabHist, setTabHist] = useState([]);  // ประวัติแท็บ — ปุ่มกลับย้อนทีละหน้า
  const goTab   = (id) => { if (id !== tab) leaveIfDirty(() => { setTabHist(h => [...h, tab]); setTab(id); }); };
  const backTab = () => { if (tabHist.length) leaveIfDirty(() => { const h = [...tabHist]; const p = h.pop(); setTabHist(h); setTab(p); }); else onBack(); };
  // ปุ่ม "Export เดือนนี้" ของแท็บรายเดือน ถูกยกขึ้นมาไว้ข้างปุ่ม Export หลักด้านบน
  const monthlyExportRef = useRef(null);
  const registerMonthExport = useCallback(fn => { monthlyExportRef.current = fn; }, []);

  // Shared "add / remove line item" logic — used by both Baseline and Monthly tabs,
  // and kept in sync with tenderCosts + every month's additions on delete.
  // Two kinds of extra item:
  //  - standalone (has `group`): a brand-new scope item with its own Acc-like code
  //  - sub-item   (has `parentCode`): a breakdown line that rolls up INTO an existing Acc. Code
  // Sub-items are shared across the Baseline and Monthly tabs — one added in
  // either place shows up in both, and in every month going forward.
  // `addedInMonth` (set only when created from the Monthly tab) records which
  // month it first appeared in, so the UI can show "เพิ่มเมื่อ ..." vs.
  // "ตั้งแต่เริ่มต้น" for ones that were already in the Baseline.
  const handleAddExtraItem = ({ name, group, parentCode, code, addedInMonth, seedName }) => {
    if (!name.trim()) return;
    const item = parentCode
      ? { code:`EX-${uid()}`, name:name.trim(), parentCode, ...(addedInMonth ? { addedInMonth } : {}) }
      : { code: code || `EX-${uid()}`, name:name.trim(), group };
    // รายการย่อย "แรก" ของแถวที่มียอดอยู่แล้ว: สร้างรายการย่อย "รายการเดิม" มารับยอดเดิมไว้ด้วย
    // (บันทึกพร้อมกันในครั้งเดียว) — ไม่งั้นแถวแม่จะกลายเป็นผลรวมรายการย่อย = 0 แล้วยอดเดิมหาย
    if (parentCode && seedName) {
      const seed = { code:`EX-${uid()}`, name: seedName, parentCode, ...(addedInMonth ? { addedInMonth } : {}) };
      saveExtraItems([...extraItems, seed, item]);
      return { code: item.code, seedCode: seed.code };
    }
    saveExtraItems([...extraItems, item]);
    return item.code;
  };

  const handleDeleteExtraItem = async (code) => {
    if (!(await uiConfirm(t("ลบรายการนี้? ยอดเงินทุกส่วนของรายการนี้ (ราคาเดิม + รายเดือนทุกเดือน) จะถูกลบด้วย","Delete this item? All its amounts (baseline + every month) will be deleted too"), { danger: true, okLabel: t("ลบ","Delete") }))) return false;
    const item = extraItems.find(e => e.code === code);
    const parent = item?.parentCode;
    const remaining = extraItems.filter(e => e.code !== code);
    saveExtraItems(remaining);
    const kidsOf = (p) => remaining.filter(e => e.parentCode === p);
    const isKeyOf = (k) => k === code || k.startsWith(code + ":");   // ค่าหลัก + ค่าคอลัมน์ย่อยของรายการนี้
    // ราคาเดิม: ลบค่าของรายการนี้ + คำนวณยอดแถวแม่ใหม่ (เดิมยอดแม่ที่บันทึกไว้ยังรวมรายการที่ลบไป
    // ฝ่ายจัดซื้อ/บัญชีจึงเห็นยอดเก่า จนกว่า QS จะกดบันทึกอีกรอบ)
    const nextTenders = { ...tenderCosts }; delete nextTenders[code];
    if (parent && !item.addedInMonth) {
      const v = kidsOf(parent).filter(k => !k.addedInMonth).reduce((s, k) => s + (parseFloat(nextTenders[k.code]) || 0), 0);
      if (v > 0) nextTenders[parent] = v; else delete nextTenders[parent];
    }
    saveTenders(nextTenders);
    // รายเดือน: ลบค่า/คอลัมน์ย่อย/สีไฮไลต์ของรายการนี้ + คำนวณยอดแม่ของเดือนที่รายการนี้มีผลใหม่
    // (คีย์ระดับโปรเจกต์ที่ขึ้นต้นด้วย $ เช่น $columns เดิมถูกแปลงจาก array เป็น object — คงไว้ตามเดิม)
    const nextAdd = {};
    Object.entries(additions).forEach(([m, obj]) => {
      if (m.startsWith("$") || !obj || typeof obj !== "object" || Array.isArray(obj)) { nextAdd[m] = obj; return; }
      const o = {};
      Object.entries(obj).forEach(([k, v]) => { if (!isKeyOf(k)) o[k] = v; });
      if (obj.$fmt) o.$fmt = Object.fromEntries(Object.entries(obj.$fmt).filter(([k]) => !isKeyOf(k)));
      if (parent && (!item.addedInMonth || item.addedInMonth <= m)) {
        const kids = kidsOf(parent).filter(k => !k.addedInMonth || k.addedInMonth <= m);
        const cols = obj.$columns ?? additions.$columns ?? [];
        const v = kids.length
          ? kids.reduce((s, k) => s + (parseFloat(o[k.code]) || 0), 0)
          : cols.reduce((s, c) => s + (parseFloat(o[`${parent}:${c.id}`]) || 0), 0);
        if (v !== 0) o[parent] = v; else delete o[parent];
      }
      nextAdd[m] = o;
    });
    saveAdditions(nextAdd);
  };

  // เปลี่ยน Acc. Code ของ "รายการใหม่" (เดิมไม่ได้ใส่รหัส → ระบบให้รหัสภายใน EX-xxxx และแสดง "—")
  // ย้ายข้อมูลทุกส่วนของโครงการนี้ตามรหัสใหม่: ราคาเดิม · รายเดือน (รวมคอลัมน์ย่อย/สีไฮไลต์) · PO · แผนของเข้า
  // คืนข้อความผิดพลาด (แสดงใต้ช่อง) หรือ null เมื่อสำเร็จ / "cancel" เมื่อผู้ใช้ไม่ยืนยัน
  const plansArr = Array.isArray(incomingPlan) ? incomingPlan : [];
  const handleRenameExtraCode = async (from, rawTo) => {
    const to = String(rawTo || "").trim();
    const err = extraCodeError(to, from, { extra: extraItems, po: poEntries, plans: plansArr });
    if (err) return err;
    if (to === from) return null;
    const item = extraItems.find(e => e.code === from);
    const hist = historyEntry(session, "edited", t(`QS เปลี่ยน Acc. Code ของ "${item?.name || ""}" เป็น ${to}`, `QS changed the Acc. Code of "${item?.name || ""}" to ${to}`));
    const next = renameProjectCode({ tenders: tenderCosts, additions, extra: extraItems, po: poEntries, plans: plansArr, hidden: hiddenAccounts }, from, to, hist);
    if ((next.poCount || next.planCount) && !(await uiConfirm(
      t(`เปลี่ยน Acc. Code ของ "${item?.name || ""}" เป็น ${to}?\n\nPO ${next.poCount} ใบ${next.planCount ? ` และแผนของเข้า ${next.planCount} รายการ` : ""} ที่ใช้รายการนี้จะเปลี่ยนเป็นรหัสใหม่ด้วย`,
        `Change the Acc. Code of "${item?.name || ""}" to ${to}?\n\n${next.poCount} PO(s)${next.planCount ? ` and ${next.planCount} plan(s)` : ""} using this item will switch to the new code too`),
      { okLabel: t("เปลี่ยนรหัส","Change code") }))) return "cancel";
    saveTenders(next.tenders);
    saveAdditions(next.additions);
    if (next.poCount) savePO?.(next.po);
    if (next.planCount) saveIncomingPlan?.(next.plans);
    if ((hiddenAccounts || []).includes(from)) saveHiddenAccounts(next.hidden);
    saveExtraItems(next.extra);
    return null;
  };

  // Hide / restore a fixed Acc. Code (511010 ... etc). Hiding doesn't erase its stored
  // numbers — it's reversible — it just removes it from the QS entry list and from
  // downstream totals, in case a project doesn't use that code at all.
  const handleHideAccount = async (code) => {
    if (!(await uiConfirm(t("นำ Acc. Code นี้ออกจากรายการหลัก? (กู้คืนได้ภายหลัง ตัวเลขที่เคยกรอกไว้จะยังไม่หาย)","Remove this Acc. Code from the main list? (restorable later; entered numbers are kept)"), { okLabel: t("นำออก","Remove") }))) return;
    saveHiddenAccounts([...hiddenAccounts, code]);
  };
  const handleRestoreAccount = (code) => saveHiddenAccounts(hiddenAccounts.filter(c => c !== code));

  return (
    <Shell role="qs" color={T.blue} project={project} onBack={backTab} onHome={onHome} onDept={onDept} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"20px 28px 0"}}>
        <div style={{display:"flex",gap:8,marginBottom:20,alignItems:"center",flexWrap:"wrap"}}>
          {!isPhone && (
          <div className="seg-tabs" role="tablist">
          {[["baseline",t("ราคาเดิม (Baseline)","Baseline")],["monthly",t("รายการเพิ่มรายเดือน","Monthly additions")]].map(([id,label])=>(
            <button key={id} role="tab" aria-selected={tab===id} onClick={()=>goTab(id)} className={`seg-tab${tab===id?" on":""}`}>
              {label}
            </button>
          ))}
          </div>
          )}
          <div style={{marginLeft:"auto"}}><CurrencyControl project={project} updateProject={updateProject}/></div>
          {tab==="monthly" && (
            <button onClick={()=>{ const fn = monthlyExportRef.current; if (!fn) return; if (runExportFn) runExportFn(fn); else inThai(fn); }} className="btn-ghost"
              style={{display:"flex",alignItems:"center",gap:6}}>
              <Ico name="download" /> {t("Export เดือนนี้","Export this month")}
            </button>
          )}
          {!isPhone && (
          <button onClick={onExport} className="btn-ghost" style={{display:"flex",alignItems:"center",gap:6}}>
            <Ico name="download" /> Export Excel
          </button>
          )}
        </div>
      </div>
      {isPhone && <BottomNav items={[
        { key:"baseline", icon:"ruler", label:t("ราคาเดิม (Baseline)","Baseline"), on:tab==="baseline", onClick:()=>goTab("baseline") },
        { key:"monthly",  icon:"calendar", label:t("รายการเพิ่มรายเดือน","Monthly additions"), on:tab==="monthly", onClick:()=>goTab("monthly") },
        { key:"export",   icon:"download", label:"Export", onClick:onExport },
      ]} />}
      {/* รหัสที่ QS ซ่อนไว้ แต่มี PO อยู่ → หน้าจัดซื้อ/บัญชียังนับงบของรหัสนี้ (กันยอด PO หาย) ยอดรวมหน้านี้จึงน้อยกว่า — บอกให้รู้ */}
      {(() => {
        const safe = hiddenSafeForPO(hiddenAccounts, poEntries);
        const codes = (hiddenAccounts || []).filter(c => !safe.includes(c));
        if (!codes.length) return null;
        const comb = buildCombinedBudget(tenderCosts, additions);
        const amt = codes.reduce((sum, c) => sum + (parseFloat(comb[c]) || 0), 0);
        return (
          <div style={{padding:"0 28px"}}>
            <div data-qs-hidden-note role="note" style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",background:"#f8fafc",border:`1px solid ${T.cardBorder}`,borderRadius:10,padding:"8px 12px",margin:"4px 0 12px",fontSize:13,color:T.textSecondary}}>
              <Ico name="eyeoff" size={16} color={T.textMuted} />
              <span>
                {t(`ยอดรวมหน้านี้ไม่รวม ${codes.length} รหัสที่ซ่อนไว้แต่มี PO`,`Totals here exclude ${codes.length} hidden code(s) that have POs`)} (<b style={{color:T.textPrimary,fontWeight:600}}>{codes.join(", ")}</b> · {t("งบ","budget")} <b style={{color:T.textPrimary,fontWeight:600,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(amt)}</b>)
                {" "}— {t("หน้าจัดซื้อ/บัญชียังนับรหัสนี้ในงบ","Procurement/Accounting still count it in the budget")}
              </span>
              <button className="btn-ghost" style={{fontSize:12,padding:"4px 10px",marginLeft:"auto"}} onClick={()=>saveHiddenAccounts(hiddenAccounts.filter(c=>!codes.includes(c)))}>{t("นำกลับมาแสดง","Show it again")}</button>
            </div>
          </div>
        );
      })()}
      {tab === "baseline"
        ? <QSBaselineTab project={project} tenderCosts={tenderCosts} saveTenders={saveTenders} extraItems={extraItems} additions={additions}
                         onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem} onRenameExtra={handleRenameExtraCode}
                         hiddenAccounts={hiddenAccounts} onHideAccount={handleHideAccount} onRestoreAccount={handleRestoreAccount} setEditMode={setEditMode} />
        : <QSMonthlyTab tenderCosts={tenderCosts} additions={additions} saveAdditions={saveAdditions}
                         extraItems={extraItems} onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} setEditMode={setEditMode} project={project} registerMonthExport={registerMonthExport} />}
    </Shell>
  );
}

// ─── QS Tab 1: Baseline (original tender cost) ────────────────────────────────
function QSBaselineTab({ project, tenderCosts, saveTenders, extraItems, additions = {}, onAddExtra, onDeleteExtra, onRenameExtra, hiddenAccounts, onHideAccount, onRestoreAccount, setEditMode }) {
  const usdRate = effRate(project);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const [draft,  setDraft]  = useState({...tenderCosts});
  const [filter, setFilter] = useState([]);   // อาเรย์หมวดที่เลือก (ว่าง = ทุกหมวด) — เลือกได้หลายหมวด
  const [hideEmpty, setHideEmpty] = useState(false);   // ซ่อนแถวที่ไม่มีค่า (ราคาเดิม = 0)
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState(null);   // "code" | "group" | "name" | "value" | null
  const [sortDir, setSortDir] = useState(1);       // 1 = asc, -1 = desc
  const [saved,  setSaved]  = useState(false);
  const [addOpen, setAddOpen] = useState(false);
  const [addDraft, setAddDraft] = useState({ code:"", name:"", group:GROUPS[0] });
  const [subFor, setSubFor] = useState(null);       // code of account currently adding a sub-item
  const [subName, setSubName] = useState("");
  const [collapsed, setCollapsed] = useState({});   // code -> true means sub-items hidden
  const [showHidden, setShowHidden] = useState(false);
  const [forceEdit, setForceEdit] = useState(false); // user explicitly clicked "แก้ไข" to unlock an already-saved baseline
  const [codeEdit, setCodeEdit] = useState(null);   // { code, value, err } — แก้ Acc. Code ของรายการใหม่ในตาราง
  const commitCode = async () => {
    if (!codeEdit || !onRenameExtra) return;
    const { code: from, value } = codeEdit;
    const err = await onRenameExtra(from, value);
    if (err === "cancel") return;
    if (err) { setCodeEdit(c => c && ({ ...c, err })); return; }
    const to = String(value).trim();
    // ย้ายค่าที่กำลังพิมพ์ค้าง (draft) ไปใช้รหัสใหม่ด้วย ไม่งั้นกดบันทึกแล้วจะเขียนรหัสเดิมกลับ
    if (to !== from) setDraft(d => { const n = { ...d }; if (from in n) { n[to] = n[from]; delete n[from]; } return n; });
    setCodeEdit(null);
  };

  // The baseline counts as "saved" (and therefore locked, requiring "แก้ไข"
  // to unlock) once it carries the explicit $saved flag, or — for baselines
  // saved before this flag existed — once it already has any real value.
  const hasData = Object.entries(tenderCosts||{}).some(([k,v]) => !k.startsWith("$") && parseFloat(v));
  const baselineSaved = tenderCosts.$saved === true || hasData;
  const editingUnlocked = !baselineSaved || forceEdit;
  useEffect(() => { setEditMode?.(editingUnlocked); return () => setEditMode?.(false); }, [editingUnlocked, setEditMode]);

  // ซิงก์ draft จากค่าที่บันทึกไว้ — แต่ "ห้าม" เขียนทับสิ่งที่กำลังพิมพ์ค้างระหว่างแก้ไข
  // (เช่น realtime refetch ตอนคนอื่นเซฟ PO ในโครงการเดียวกัน) ยกเว้นตอนสลับโครงการ รีเซ็ตเสมอ
  const prevProjRef = useRef(project?.id);
  useEffect(() => {
    const switched = prevProjRef.current !== project?.id;
    prevProjRef.current = project?.id;
    if (switched || !editingUnlocked) setDraft({ ...tenderCosts });
  }, [tenderCosts, project?.id, editingUnlocked]);

  // ตั้งธง "แก้ค้างยังไม่บันทึก" เมื่อ draft ต่างจากค่าที่บันทึกไว้ (ระหว่างโหมดแก้ไข)
  const isDirty = editingUnlocked && JSON.stringify(draft) !== JSON.stringify(tenderCosts);
  useEffect(() => { UnsavedGuard.dirty = isDirty; return () => { UnsavedGuard.dirty = false; }; }, [isDirty]);

  // ยกเลิกการแก้ไข: ทิ้งค่าที่พิมพ์ค้าง คืนกลับเป็นค่าที่บันทึกไว้ล่าสุด แล้วล็อกกลับ
  const canCancel = baselineSaved; // มีค่าที่บันทึกไว้ให้ย้อนกลับได้
  const handleCancel = () => { setDraft({ ...tenderCosts }); setForceEdit(false); setAddOpen(false); setSubFor(null); };
  useEffect(() => {
    if (!editingUnlocked) return;
    const onEsc = (e) => {
      if (e.key !== "Escape" || e.defaultPrevented || !canCancel) return;
      const ae = document.activeElement;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT" || ae.isContentEditable)) return; // กำลังพิมพ์ — ไม่ทิ้งข้อมูล
      e.preventDefault();
      if (UnsavedGuard.dirty) { uiConfirm(t("ยกเลิกการแก้ไข? ค่าที่พิมพ์ไว้แต่ยังไม่บันทึกจะหายไป","Cancel editing? Unsaved values will be lost"), { danger: true, okLabel: t("ทิ้งการแก้ไข","Discard changes"), cancelLabel: t("แก้ต่อ","Keep editing") }).then(ok => { if (ok) handleCancel(); }); return; }
      handleCancel();
    };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [editingUnlocked, canCancel, tenderCosts]);

  // Sub-items (e.g. "Silicone Structure") roll up into an existing Acc. Code (e.g. 511025).
  // Standalone extras (no parentCode) are brand-new items with their own group, shown as their own row.
  // Shared with the Monthly tab — a sub-item added on either tab shows up on both.
  const subItemsByParent = {};
  extraItems.forEach(e => {
    if (e.parentCode) (subItemsByParent[e.parentCode] = subItemsByParent[e.parentCode] || []).push(e);
  });
  const standaloneExtras = extraItems.filter(e => !e.parentCode);
  // Baseline only ever shows sub-items that were created as part of the baseline
  // itself (no addedInMonth). Ones added later from the Monthly tab live only
  // there, starting from the month they were added — they don't belong to
  // "ราคาเดิม (Baseline)" and would be confusing to show here with a 0.00 baseline value.
  const childrenOf = (code) => (subItemsByParent[code] || []).filter(k => !k.addedInMonth);

  // Effective value of a row: sum of its sub-items if it has any, else its own draft value.
  const effectiveValue = (row) => {
    const kids = !row.isExtra ? childrenOf(row.code) : [];
    if (kids.length) return kids.reduce((s,k)=>s+(parseFloat(draft[k.code])||0),0);
    return parseFloat(draft[row.code]) || 0;
  };

  const visibleAccounts = ACCOUNTS.filter(a => !hiddenAccounts.includes(a.code));
  const hiddenList = ACCOUNTS.filter(a => hiddenAccounts.includes(a.code));
  const allRows = [...visibleAccounts, ...standaloneExtras.map(e => ({ code:e.code, name:e.name, group:e.group, isExtra:true }))];

  const base  = allRows.reduce((s,r)=>s+effectiveValue(r),0);
  const adj3  = allRows.reduce((s,r)=>s+wasteOf(effectiveValue(r)),0);   // เผื่อเศษ/สูญเสีย — รวมในงบ (ทุน) แล้ว · ผลรวมรายรหัส (ปัดสตางค์) ให้ตรงกับตาราง

  const q = search.toLowerCase();
  const filtered = allRows.filter(a => {
    if (filter.length && !filter.includes(a.group)) return false;
    const selfMatch = a.name.toLowerCase().includes(q) || a.code.includes(search);
    const childMatch = !a.isExtra && childrenOf(a.code).some(k=>k.name.toLowerCase().includes(q));
    return selfMatch || childMatch;
  });

  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => -d);
    else { setSortKey(key); setSortDir(1); }
  };
  const hiddenEmptyCount = hideEmpty ? filtered.length - filtered.filter(a => effectiveValue(a) !== 0).length : 0;

  // ── Monthly additions rolled up per Acc. Code (read-only on the Baseline tab) ──
  // Additions are entered on the "Monthly additions" tab; the monthly Save rolls a
  // code's sub-items and per-item columns into its plain `code` key each month, so
  // the all-time total for a code is just the sum of that key across every month
  // (skip the "$…" meta keys). Same rule buildCombinedBudget() uses, so
  // Baseline + Additions here matches the grand total shown elsewhere.
  const rowAddTotal = (row) => Object.keys(additions).reduce((s,m)=> m.startsWith("$") ? s : s + monthAddValue(additions, m, row.code), 0);
  const rowGrand    = (row) => withWaste(effectiveValue(row)) + rowAddTotal(row);   // ราคาเดิม + เผื่อเศษ + งานเพิ่ม = งบจริง
  const addAll      = allRows.reduce((s,r)=> s + rowAddTotal(r), 0);   // งานเพิ่มรวมทุก Code ทุกเดือน

  const displayRows = (() => {
    // ซ่อนแถวที่ไม่มีค่า = ราคาเดิม (รวมรายการย่อย) เป็น 0
    const baseRows = hideEmpty ? filtered.filter(a => effectiveValue(a) !== 0) : filtered;
    if (!sortKey) return baseRows;
    const arr = [...baseRows];
    arr.sort((a, b) => {
      let av, bv;
      if (sortKey === "code")       { av = a.code; bv = b.code; }
      else if (sortKey === "group") { av = GROUPS.indexOf(a.group); bv = GROUPS.indexOf(b.group); }
      else if (sortKey === "name")  { av = a.name; bv = b.name; }
      else if (sortKey === "add")   { av = rowAddTotal(a); bv = rowAddTotal(b); }
      else if (sortKey === "grand") { av = rowGrand(a); bv = rowGrand(b); }
      else                          { av = effectiveValue(a); bv = effectiveValue(b); }
      if (typeof av === "string") return av.localeCompare(bv) * sortDir;
      return (av - bv) * sortDir;
    });
    return arr;
  })();

  const handleSave = () => {
    const merged = {...draft};
    ACCOUNTS.forEach(a => {
      const kids = childrenOf(a.code);
      if (kids.length) merged[a.code] = kids.reduce((s,k)=>s+(parseFloat(merged[k.code])||0),0);
    });
    const clean = {};
    Object.entries(merged).forEach(([k,v]) => {
      if (k.startsWith("$")) return; // meta keys ($saved) are re-added explicitly below
      if(v!==""&&!isNaN(v)&&parseFloat(v)>0) clean[k]=parseFloat(v);
    });
    clean.$saved = true;
    saveTenders(clean);
    setForceEdit(false); setAddOpen(false); setSubFor(null);
    setSaved(true); setTimeout(()=>setSaved(false),2000);
  };

  const handleAddRow = () => {
    if (!addDraft.name.trim()) return;
    const code = addDraft.code.trim();
    if (code) {
      const taken = ACCOUNTS.some(a=>a.code===code) || extraItems.some(e=>e.code===code);
      if (taken) { uiAlert(t(`Acc. Code "${code}" มีอยู่แล้ว กรุณาใช้รหัสอื่น`, `Acc. Code "${code}" already exists, please use another`)); return; }
    }
    onAddExtra({ name:addDraft.name, group:addDraft.group, code: code || undefined });
    setAddDraft({ code:"", name:"", group:GROUPS[0] }); setAddOpen(false);
  };

  const handleAddSub = (parentCode) => {
    if (!subName.trim()) return;
    const parentVal = parseFloat(draft[parentCode]) || 0;
    if (childrenOf(parentCode).length === 0 && parentVal !== 0) {
      // รายการย่อยแรก: ย้ายราคาเดิมของแถวแม่ไปไว้ในรายการย่อย "รายการเดิม" (ยอดรวมเท่าเดิม)
      const res = onAddExtra({ name:subName, parentCode, seedName: t("รายการเดิม","Original item") });
      if (res?.seedCode) {
        saveTenders({ ...tenderCosts, [res.seedCode]: parentVal }); // บันทึกทันที กันยกเลิกแล้วยอดแม่เป็น 0
        setDraft(d => ({ ...d, [res.seedCode]: parentVal }));
      }
    } else {
      onAddExtra({ name:subName, parentCode });
    }
    setCollapsed(c => ({...c, [parentCode]: false})); // reveal the newly-added sub-item
    setSubName(""); setSubFor(null);
  };

  const handleDeleteRow = (code) => {
    Promise.resolve(onDeleteExtra(code)).then(ok => { if (ok === false) return; setDraft(d => { const n = {...d}; delete n[code]; return n; }); });
  };

  return (
    <div style={{padding:"4px 28px 24px"}}>
      {/* Stats */}
      <div className="stat-grid" style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:24}}>
        <StatCard label={t("ราคาเดิมรวม (Tender Cost)","Total Tender Cost")} value={"฿"+fmt0(base)} thb={base} rate={usdRate} sub={t(`ราคาเดิมทั้งหมด (ก่อนบวกเผื่อเศษ ${WASTE_LBL})`,`All baseline prices (before the ${WASTE_LBL} wastage)`)} color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label={t(`เผื่อเศษ/สูญเสีย ${WASTE_LBL}`,`Wastage allowance ${WASTE_LBL}`)} value={"฿"+fmt0(adj3)} thb={adj3} rate={usdRate} sub={t("รวมในงบแล้ว — บวกเข้าแต่ละ Acc. Code","Included in the budget — added to each Acc. Code")} color={T.amber} icon="⚙️" accent={T.amberBg}/>
        <StatCard label={t("งานเพิ่ม (รวมทุกเดือน)","Additions (all months)")} value={"฿"+fmt0(addAll)} thb={addAll} rate={usdRate} sub={t("รวมยอดที่เพิ่มจากแท็บรายเดือน","Total added from the Monthly tab")} color={T.purple} icon="➕" accent={T.purpleBg}/>
        <StatCard label={t("รวมทั้งหมด","Grand total")} value={"฿"+fmt0(base+adj3+addAll)} thb={base+adj3+addAll} rate={usdRate} sub={t(`ราคาเดิม + เผื่อเศษ ${WASTE_LBL} + งานเพิ่ม (งบจริง)`,`Baseline + ${WASTE_LBL} wastage + additions (actual budget)`)} color={T.green} icon="✅" accent={T.greenBg}/>
      </div>

      {/* Filters + Add row + Save */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <SearchInput value={search} onChange={setSearch} placeholder={t("ค้นหา Account Code / ชื่อ...","Search Account Code / name...")} width={240}/>
        <button onClick={()=>setHideEmpty(v=>!v)}
          title={t("ซ่อน/แสดงแถวที่ไม่มีค่า (ราคาเดิม = 0)","Hide/show empty rows (baseline = 0)")}
          style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:600,cursor:"pointer",
            border:`1.5px solid ${hideEmpty?T.blue:T.cardBorder}`,background:hideEmpty?T.blue:T.card,color:hideEmpty?"#fff":T.textSecondary,whiteSpace:"nowrap"}}>
          {hideEmpty ? `✓ ${t("เฉพาะที่มีค่า","With value only")}${hiddenEmptyCount?` (${t("ซ่อน","hidden")} ${hiddenEmptyCount})`:""}` : t("เฉพาะที่มีค่า","With value only")}
        </button>
        <GroupFilter selected={filter} onChange={setFilter}/>
        <div style={{flex:1}}/>
        {hiddenList.length > 0 && (
          <button className="btn-ghost" onClick={()=>setShowHidden(v=>!v)} style={{color:T.textMuted}}>
            <Ico name="eyeoff" /> {t("ที่ซ่อนไว้","Hidden")} ({hiddenList.length})
          </button>
        )}
        <button className="btn-ghost" onClick={()=>setAddOpen(v=>!v)} disabled={!editingUnlocked}
          style={!editingUnlocked?{opacity:0.4,cursor:"not-allowed"}:undefined}>+ {t("เพิ่มรายการหลักใหม่","Add new main item")}</button>
        {!editingUnlocked && (
          <span style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:T.textMuted,background:"#f1f5f9",padding:"6px 12px",borderRadius:8,fontWeight:600}}>
            <Ico name="lock" size={14} /> {t("บันทึกแล้ว","Saved")}
          </span>
        )}
        {editingUnlocked ? (
          <>
            <button onClick={handleSave} className="btn-primary"
              style={{background:saved?T.green:T.blue,minWidth:140}}>
              {saved?t("✓ บันทึกแล้ว","✓ Saved"):t("บันทึก Tender Cost","Save Tender Cost")}
            </button>
            {canCancel && (
              <button onClick={handleCancel} className="btn-ghost" title={t("ยกเลิกการแก้ไข (Esc)","Cancel editing (Esc)")}
                style={{color:T.red,borderColor:T.red}}>✕ {t("ยกเลิก","Cancel")}</button>
            )}
          </>
        ) : (
          <button onClick={()=>setForceEdit(true)} className="btn-primary" style={{minWidth:140,display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6}}>
            <Ico name="edit" /> {t("แก้ไข Tender Cost","Edit Tender Cost")}
          </button>
        )}
      </div>

      {/* Hidden accounts panel */}
      {showHidden && hiddenList.length > 0 && (
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:14,marginBottom:16}}>
          <div style={{fontSize:11,color:T.textSecondary,marginBottom:8}}>{t('Acc. Code ที่ซ่อนไว้ — ตัวเลขที่เคยกรอกยังอยู่ กด "กู้คืน" เพื่อนำกลับมาแสดง','Hidden Acc. Codes — entered values are kept; click "Restore" to bring them back')}</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
            {hiddenList.map(a=>(
              <div key={a.code} style={{display:"flex",alignItems:"center",gap:8,background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:8,padding:"6px 10px"}}>
                <span style={{fontVariantNumeric:"tabular-nums",fontSize:11,color:T.textMuted}}>{a.code}</span>
                <span style={{fontSize:12,color:T.textPrimary}}>{a.name}</span>
                <button onClick={()=>onRestoreAccount(a.code)} className="btn-ghost" style={{padding:"3px 9px",fontSize:11}}>↺ {t("กู้คืน","Restore")}</button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Inline "add standalone row" form */}
      {addOpen && (
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"1fr 2fr 1fr auto",gap:10,alignItems:"end"}}>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>{t("Acc. Code (ถ้ามี)","Acc. Code (if any)")}</span>
            <input className="input-base" value={addDraft.code} onChange={e=>setAddDraft(d=>({...d,code:e.target.value}))}
              placeholder={t("เช่น 511099","e.g. 511099")} style={{fontVariantNumeric:"tabular-nums"}}
              onKeyDown={e=>e.key==="Enter"&&handleAddRow()} />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>{t("ชื่อรายการใหม่ (งานที่ไม่มี Acc. Code เดิมรองรับ)","New item name (work without an existing Acc. Code)")}</span>
            <input className="input-base" value={addDraft.name} onChange={e=>setAddDraft(d=>({...d,name:e.target.value}))}
              placeholder={t("พิมพ์ชื่อรายการที่ต้องการเพิ่ม","Type the item name to add")} onKeyDown={e=>e.key==="Enter"&&handleAddRow()} autoFocus />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>Group</span>
            <select className="input-base" value={addDraft.group} onChange={e=>setAddDraft(d=>({...d,group:e.target.value}))}>
              {GROUPS.map(g => <option key={g} value={g}>{g}</option>)}
            </select>
          </label>
          <button className="btn-primary" onClick={handleAddRow}>+ {t("เพิ่ม","Add")}</button>
        </div>
      )}

      {/* Table */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
        <div className="hscroll"><table style={{width:"100%",minWidth:1100,borderCollapse:"collapse",fontSize:13}}>
          <thead>
            <tr style={{background:"#f8fafc"}}>
              {[
                {label:"Acc. Code", key:"code", align:"left"},
                {label:"Group", key:"group", align:"left"},
                {label:"Account Name", key:"name", align:"left"},
                {label:t("ราคาเดิม (THB)","Tender Cost (THB)"), key:"value", align:"right"},
                {label:t(`เผื่อเศษ ${WASTE_LBL}`,`Wastage ${WASTE_LBL}`), key:"waste", align:"right"},
                {label:t("รวมงานเพิ่ม","Total Additions"), key:"add", align:"right"},
                {label:t("รวมทั้งหมด","Grand total"), key:"grand", align:"right"},
                {label:"", key:null, align:"center"},
              ].map(({label,key,align})=>(
                <th key={label||"__actions"}
                  style={{padding:"11px 16px",textAlign:align,color:sortKey===key?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                  <span onClick={()=>key&&handleSort(key)} style={{cursor:key?"pointer":"default",userSelect:"none"}}>{label}{key && sortKey===key ? (sortDir===1?" ▲":" ▼") : ""}</span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayRows.map((a,i)=>{
              const kids = !a.isExtra ? childrenOf(a.code) : [];
              const hasKids = kids.length > 0;
              const isCollapsed = hasKids && collapsed[a.code];
              const rowVal = effectiveValue(a);
              return (
                <Fragment key={a.code}>
                  <tr onClick={()=>hasKids && setCollapsed(c=>({...c,[a.code]:!c[a.code]}))}
                      style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(hasKids&&!isCollapsed)||subFor===a.code?"none":"1px solid #f1f5f9",cursor:hasKids?"pointer":"default"}}>
                    <td style={{padding:"10px 16px",color:a.isExtra?T.amber:T.blue,fontVariantNumeric:"tabular-nums",fontSize:13,fontWeight:500,whiteSpace:"nowrap"}}>
                      {hasKids && (
                        <span title={isCollapsed?t("ขยายรายการย่อย","Expand sub-items"):t("ย่อรายการย่อย","Collapse sub-items")}
                          style={{color:T.textMuted,fontSize:12,marginRight:6,verticalAlign:"middle",display:"inline-block"}}>
                          {isCollapsed?"▸":"▾"}
                        </span>
                      )}
                      {a.isExtra && codeEdit?.code === a.code ? (
                        <span onClick={e=>e.stopPropagation()} style={{display:"inline-flex",flexDirection:"column",gap:3}}>
                          <span style={{display:"inline-flex",alignItems:"center",gap:4}}>
                            <input data-code-input className="input-base" autoFocus value={codeEdit.value} placeholder={t("เช่น 511099","e.g. 511099")}
                              aria-label={t(`Acc. Code ของ ${a.name}`,`Acc. Code for ${a.name}`)} aria-invalid={codeEdit.err ? true : undefined}
                              onChange={e=>setCodeEdit(c=>({ ...c, value:e.target.value, err:null }))}
                              onKeyDown={e=>{ if (e.key==="Enter") { e.preventDefault(); commitCode(); } else if (e.key==="Escape") { e.preventDefault(); e.stopPropagation(); setCodeEdit(null); } }}
                              style={{width:112,padding:"5px 8px",fontVariantNumeric:"tabular-nums",...(codeEdit.err?{borderColor:T.red}:{})}}/>
                            <button type="button" data-code-save onClick={commitCode} aria-label={t("บันทึกรหัส","Save code")} title={t("บันทึกรหัส","Save code")}
                              style={{background:T.blue,border:"none",color:"#fff",borderRadius:7,width:28,height:28,cursor:"pointer",display:"inline-grid",placeItems:"center"}}><Ico name="check" size={14} /></button>
                            <button type="button" onClick={()=>setCodeEdit(null)} aria-label={t("ยกเลิก","Cancel")} title={t("ยกเลิก","Cancel")}
                              style={{background:"none",border:`1px solid ${T.cardBorder}`,color:T.textMuted,borderRadius:7,width:28,height:28,cursor:"pointer"}}>✕</button>
                          </span>
                          {codeEdit.err && <span role="alert" data-code-error style={{fontSize:12,color:T.red,maxWidth:260,whiteSpace:"normal",lineHeight:1.35,fontWeight:400}}>{codeEdit.err}</span>}
                        </span>
                      ) : (
                        <>
                          {a.isExtra ? codeText(a.code) : a.code}
                          {a.isExtra && editingUnlocked && onRenameExtra && (
                            <button type="button" data-code-edit onClick={e=>{ e.stopPropagation(); setCodeEdit({ code:a.code, value:/^EX-/.test(a.code) ? "" : a.code, err:null }); }}
                              aria-label={t(`แก้ Acc. Code ของ ${a.name}`,`Edit Acc. Code for ${a.name}`)} title={/^EX-/.test(a.code) ? t("ใส่ Acc. Code","Set Acc. Code") : t("แก้ Acc. Code","Edit Acc. Code")}
                              style={{marginLeft:6,background:"none",border:`1px dashed ${T.cardBorder}`,borderRadius:6,color:T.textMuted,cursor:"pointer",padding:"2px 6px",verticalAlign:"middle",display:"inline-flex",alignItems:"center",gap:3,fontSize:11}}>
                              <Ico name="edit" size={12} />{/^EX-/.test(a.code) ? t("ใส่รหัส","Set code") : ""}
                            </button>
                          )}
                        </>
                      )}
                    </td>
                    <td style={{padding:"10px 16px"}}>
                      {(i===0 || displayRows[i-1]?.group!==a.group) && <span style={{background:T.blueLight,color:T.blue,fontSize:12,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{a.group}</span>}
                    </td>
                    <td style={{padding:"10px 16px",color:T.textPrimary}}>
                      {a.name}
                      {a.isExtra && <span style={{marginLeft:7,fontSize:12,background:T.amberBg,color:T.amber,padding:"1px 8px",borderRadius:6,fontWeight:600}}>{t("รายการใหม่","New item")}</span>}
                      {hasKids && <span style={{marginLeft:7,fontSize:12,background:T.greenBg,color:T.green,padding:"1px 8px",borderRadius:6,fontWeight:600}}>{kids.length} {t("รายการย่อย","sub-items")}</span>}
                      {!a.isExtra && editingUnlocked && (
                        <button onClick={(e)=>{e.stopPropagation(); setSubFor(subFor===a.code?null:a.code); setSubName(""); setCollapsed(c=>({...c,[a.code]:false}));}} title={t("เพิ่มรายการย่อยใต้ Acc. Code นี้","Add a sub-item under this Acc. Code")}
                          style={{marginLeft:9,background:"none",border:`1px dashed ${T.cardBorder}`,borderRadius:6,color:T.textMuted,cursor:"pointer",fontSize:12,padding:"1px 7px"}}>
                          + {t("รายการย่อย","Sub-item")}
                        </button>
                      )}
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right"}}>
                      {hasKids ? (
                        <div style={{width:160,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",background:T.blueLight,borderRadius:8,color:T.blue,fontWeight:650,fontSize:13}}>
                          {fmtZ(rowVal)}
                          {usdLine(rowVal, usdRate)}
                        </div>
                      ) : editingUnlocked ? (
                        <MoneyInput value={draft[a.code]??""} onChange={v=>setDraft(d=>({...d,[a.code]:v}))}
                          style={{width:160,background:(parseFloat(draft[a.code])||0)>0?T.blueLight:T.bg}}/>
                      ) : (
                        <div style={{width:160,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13,color:draft[a.code]>0?T.textPrimary:T.textMuted}}>{fmtZ(rowVal)}{usdLine(rowVal, usdRate)}</div>
                      )}
                    </td>
                    {(() => { const addV = rowAddTotal(a); const wV = wasteOf(rowVal); const grandV = rowVal + wV + addV; return (<>
                    <td style={{padding:"8px 16px",textAlign:"right"}}>
                      <div style={{width:120,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13,color:wV?T.amber:T.textMuted}}>{fmtZ(wV)}{usdLine(wV, usdRate)}</div>
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right"}}>
                      <div style={{width:150,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13, ...(addV>0?{background:T.amberBg,color:T.amber,fontWeight:650,borderRadius:8}:{color:T.textMuted})}}>{fmtZ(addV)}{usdLine(addV, usdRate)}</div>
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right"}}>
                      <div style={{width:160,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13,fontWeight:700, ...(grandV>0?{background:T.greenBg,color:T.green,borderRadius:8}:{color:T.textMuted})}}>{fmtZ(grandV)}{usdLine(grandV, usdRate)}</div>
                    </td>
                    </>); })()}
                    <td style={{padding:"8px 16px",textAlign:"center"}}>
                      {editingUnlocked && (a.isExtra
                        ? <button onClick={(e)=>{e.stopPropagation(); handleDeleteRow(a.code);}} title={t("ลบรายการนี้","Delete this item")}
                            style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:14}}>✕</button>
                        : <button onClick={(e)=>{e.stopPropagation(); onHideAccount(a.code);}} title={t("นำ Acc. Code นี้ออกจากรายการหลัก (กู้คืนได้)","Remove this Acc. Code from the main list (restorable)")}
                            style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",fontSize:14}}>✕</button>)}
                    </td>
                  </tr>

                  {/* Sub-items — roll up into the parent Acc. Code's total above */}
                  {!isCollapsed && kids.map((k,ki)=>(
                    <tr key={k.code} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(ki===kids.length-1 && subFor!==a.code)?"1px solid #f1f5f9":"none"}}>
                      <td style={{padding:"6px 16px 6px 30px",color:T.green,fontSize:13}}>↳</td>
                      <td/>
                      <td style={{padding:"6px 16px",color:T.green,fontSize:13,fontStyle:"italic"}}>
                        {k.name}
                        {k.addedInMonth && (
                          <span title={t("เพิ่มเข้ามาระหว่างทาง ไม่ได้มีมาตั้งแต่ต้น","Added later, not from the start")} style={{marginLeft:7,fontSize:12,background:T.amberBg,color:T.amber,padding:"1px 7px",borderRadius:6,fontWeight:600,fontStyle:"normal"}}>
                            {t("เพิ่มเมื่อ","Added")} {monthShortLabel(k.addedInMonth)}
                          </span>
                        )}
                      </td>
                      <td style={{padding:"6px 16px",textAlign:"right"}}>
                        {editingUnlocked ? (
                          <MoneyInput value={draft[k.code]??""} onChange={v=>setDraft(d=>({...d,[k.code]:v}))}
                            style={{width:160,fontSize:13,background:(parseFloat(draft[k.code])||0)>0?T.greenBg:T.bg}}/>
                        ) : (
                          <div style={{width:160,marginLeft:"auto",padding:"6px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13,color:draft[k.code]>0?T.textPrimary:T.textMuted}}>{fmtZ(parseFloat(draft[k.code])||0)}{usdLine(parseFloat(draft[k.code])||0, usdRate)}</div>
                        )}
                      </td>
                      {(() => { const kBase = parseFloat(draft[k.code])||0; const kAdd = rowAddTotal(k); const kW = wasteOf(kBase); const kGrand = kBase + kW + kAdd; return (<>
                      <td style={{padding:"6px 16px",textAlign:"right"}}>
                        <div style={{width:120,marginLeft:"auto",padding:"6px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13,color:kW?T.amber:T.textMuted}}>{fmtZ(kW)}{usdLine(kW, usdRate)}</div>
                      </td>
                      <td style={{padding:"6px 16px",textAlign:"right"}}>
                        <div style={{width:150,marginLeft:"auto",padding:"6px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13, ...(kAdd>0?{color:T.amber,fontWeight:600}:{color:T.textMuted})}}>{fmtZ(kAdd)}{usdLine(kAdd, usdRate)}</div>
                      </td>
                      <td style={{padding:"6px 16px",textAlign:"right"}}>
                        <div style={{width:160,marginLeft:"auto",padding:"6px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13,fontWeight:650, ...(kGrand>0?{color:T.green}:{color:T.textMuted})}}>{fmtZ(kGrand)}{usdLine(kGrand, usdRate)}</div>
                      </td>
                      </>); })()}
                      <td style={{padding:"6px 16px",textAlign:"center"}}>
                        {editingUnlocked && (
                          <button onClick={()=>handleDeleteRow(k.code)} title={t("ลบรายการย่อยนี้","Delete this sub-item")}
                            style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
                        )}
                      </td>
                    </tr>
                  ))}

                  {/* Inline "add sub-item" form for this account */}
                  {subFor===a.code && (
                    <tr style={{background:T.greenBg,borderBottom:"1px solid #f1f5f9"}}>
                      <td/><td/>
                      <td style={{padding:"7px 16px"}} colSpan={1}>
                        <input className="input-base" value={subName} onChange={e=>setSubName(e.target.value)}
                          placeholder={t("ชื่อรายการย่อย เช่น Silicone Structure","Sub-item name e.g. Silicone Structure")} style={{width:"100%",fontSize:13}}
                          onKeyDown={e=>e.key==="Enter"&&handleAddSub(a.code)} autoFocus />
                      </td>
                      <td colSpan={5} style={{padding:"7px 16px",display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button className="btn-primary" style={{padding:"5px 12px",fontSize:13}} onClick={()=>handleAddSub(a.code)}>+ {t("เพิ่ม","Add")}</button>
                        <button className="btn-ghost" style={{padding:"5px 12px",fontSize:13}} onClick={()=>setSubFor(null)}>{t("ยกเลิก","Cancel")}</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
              <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:13}}>{filtered.length} {t("รายการ","items")}</td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.blue,fontVariantNumeric:"tabular-nums",fontWeight:650,fontSize:14}}>
                {fmtZ(filtered.reduce((s,a)=>s+effectiveValue(a),0))}
                {usdLine(filtered.reduce((s,a)=>s+effectiveValue(a),0), usdRate)}
              </td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.amber,fontVariantNumeric:"tabular-nums",fontWeight:650,fontSize:14}}>
                {fmtZ(filtered.reduce((s,a)=>s+wasteOf(effectiveValue(a)),0))}
                {usdLine(filtered.reduce((s,a)=>s+wasteOf(effectiveValue(a)),0), usdRate)}
              </td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.amber,fontVariantNumeric:"tabular-nums",fontWeight:650,fontSize:14}}>
                {fmtZ(filtered.reduce((s,a)=>s+rowAddTotal(a),0))}
                {usdLine(filtered.reduce((s,a)=>s+rowAddTotal(a),0), usdRate)}
              </td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.green,fontVariantNumeric:"tabular-nums",fontWeight:700,fontSize:14}}>
                {fmtZ(filtered.reduce((s,a)=>s+rowGrand(a),0))}
                {usdLine(filtered.reduce((s,a)=>s+rowGrand(a),0), usdRate)}
              </td>
              <td/>
            </tr>
          </tfoot>
        </table></div>
      </div>
    </div>
  );
}

// ป้ายแกน X ของกราฟแนวโน้ม — เดือนที่เลือกอยู่จะเป็นชิปสีน้ำเงินเด่นชัด
function MonthAxisTick({ x, y, payload, selectedLabel }) {
  const sel = payload && payload.value === selectedLabel;
  if (sel) {
    const w = Math.max(52, String(payload.value).length * 8 + 20);
    return (
      <g transform={`translate(${x},${y})`}>
        <rect x={-w/2} y={5} width={w} height={22} rx={11} fill={T.blue}/>
        <text x={0} y={20} textAnchor="middle" fontSize={11} fontWeight={700} fill="#fff">{payload.value}</text>
      </g>
    );
  }
  return <text x={x} y={y} dy={17} textAnchor="middle" fontSize={11} fill={T.textMuted}>{payload && payload.value}</text>;
}

// ─── QS Tab 2: Monthly additions (เดิม / เพิ่มเดือนนี้ / รวมสะสม) ─────────────
// ── ตรึงคอลัมน์ซ้ายของตารางงบ QS (Acc.Code / Group / Account Name / ยอดก่อนหน้า) ──
//   4 คอลัมน์แรกไม่เลื่อนซ้าย-ขวา · คอลัมน์รายการที่เหลือเลื่อนได้
const QSF_W = [120, 118, 320, 150];                 // ความกว้างคงที่ 4 คอลัมน์ที่ตรึง
const QSF_L = [0, 120, 238, 558];                   // left สะสม (0, 120, 120+118, 238+320)
const QSF_SPAN3 = QSF_W[0] + QSF_W[1] + QSF_W[2];   // = 558 (สำหรับ footer colSpan=3)
const qsFrz = (i, bg, z = 3) => ({
  position: "sticky", left: QSF_L[i], width: QSF_W[i], minWidth: QSF_W[i], maxWidth: QSF_W[i],
  background: bg, zIndex: z, ...(i === 3 ? { boxShadow: "3px 0 5px -2px rgba(15,23,42,0.13)" } : {}),
});
const qsFrzSpan3 = (bg, z = 3) => ({
  position: "sticky", left: 0, width: QSF_SPAN3, minWidth: QSF_SPAN3, maxWidth: QSF_SPAN3, background: bg, zIndex: z,
  boxShadow: "3px 0 5px -2px rgba(15,23,42,0.13)",
});
// แถวรวมด้านล่าง: ตรึงไว้ (sticky bottom) ไม่ต้องเลื่อนลงไปดู
const QSF_FOOT = { position: "sticky", bottom: 0, zIndex: 5, background: "#eef2f7" };
function QSMonthlyTab({ tenderCosts, additions, saveAdditions, extraItems, onAddExtra, onDeleteExtra, hiddenAccounts, setEditMode, project, registerMonthExport }) {
  const usdRate = effRate(project);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const thisMonth = todayStr().slice(0,7);
  const months = Object.keys(additions).filter(k=>!k.startsWith("$")).sort();
  const [month, setMonth] = useState(months.length ? months[months.length-1] : thisMonth);
  // หัวตาราง 2 ชั้นตรึงตอนเลื่อน: ชั้นที่ 2 ต้องติดอยู่ใต้ชั้นแรกพอดี — วัดความสูงจริงของชั้นแรก
  // (เดิมตั้งตายตัว 33px แต่ชั้นแรกสูงกว่านั้น ชั้นที่ 2 จึงเลื่อนไปทับ/เหลื่อมกัน)
  const mTableRef = useRef(null), mHead1Ref = useRef(null);
  useLayoutEffect(() => {
    const tb = mTableRef.current, tr = mHead1Ref.current;
    if (!tb) return;
    if (!tr) { tb.style.removeProperty("--mh1"); return; }
    const set = () => tb.style.setProperty("--mh1", `${tr.getBoundingClientRect().height}px`);
    set();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(set); ro.observe(tr);
    return () => ro.disconnect();
  });
  const [newMonth, setNewMonth] = useState("");
  const [filter, setFilter] = useState([]);   // อาเรย์หมวดที่เลือก (ว่าง = ทุกหมวด) — เลือกได้หลายหมวด
  const [hideEmpty, setHideEmpty] = useState(false);   // ซ่อนแถวที่ไม่มีค่า (รวมสะสม = 0)
  const [search, setSearch] = useState("");
  const [sortKey, setSortKey] = useState(null);   // "code" | "group" | "name" | "before" | "add" | "cum" | null
  const [sortDir, setSortDir] = useState(1);
  const [draftAdd, setDraftAdd] = useState({...(additions[month]||{})});
  const [saved, setSaved] = useState(false);
  const [addExtraOpen, setAddExtraOpen] = useState(false);
  const [extraDraft, setExtraDraft] = useState({ code:"", name:"", group:GROUPS[0] });
  const [subFor, setSubFor] = useState(null);   // code of row currently adding a sub-item
  const [subName, setSubName] = useState("");
  const [rowCollapsed, setRowCollapsed] = useState({}); // code -> true means sub-items hidden
  const [addColOpen, setAddColOpen] = useState(false);  // "+ เพิ่มรายการ" inline form open?
  const [newColName, setNewColName] = useState("");
  const [forceEdit, setForceEdit] = useState(false);    // user explicitly clicked "แก้ไข" to unlock an already-saved month

  useEffect(() => { if (!months.includes(month) && months.length) setMonth(months[months.length-1]); }, [months]); // eslint-disable-line
  useEffect(() => { setForceEdit(false); setAddColOpen(false); }, [month]); // switching months always re-locks until "แก้ไข" is clicked again
  // ผูกปุ่ม "Export เดือนนี้" ที่ยกไปไว้บนหัว (QSView) ให้ยิง export ของเดือนที่เลือกอยู่
  useEffect(() => {
    if (!registerMonthExport) return;
    registerMonthExport(() => exportQSMonthExcel(project, tenderCosts, additions, month, extraItems, hiddenAccounts));
    return () => registerMonthExport(null);
  }, [registerMonthExport, project, tenderCosts, additions, month, extraItems, hiddenAccounts]);

  // "เพิ่มรายการ" — named sub-columns (e.g. CC#16, CC#17), each holding its own
  // set of per-Account-Code entries that add up to a row's monthly total.
  // The column set itself is stored once at the project level ($columns on
  // the additions object, a sibling of the month keys) so a column created
  // in any month automatically carries forward into every other month too —
  // it isn't something you have to re-create month by month. Projects
  // created before this feature has no $columns and behave exactly as
  // before: one plain entry field per row.
  // คอลัมน์ (รายการย่อย) เก็บ "ต่อเดือน" แล้ว → additions[month].$columns
  // ของเก่าเคยเก็บระดับโปรเจกต์ (additions.$columns) ยัง fallback ให้เดือนที่ยัง
  // ไม่มีของตัวเอง เพื่อไม่ให้ข้อมูลเดิมหาย พอเดือนไหนถูกบันทึกก็จะได้ชุดคอลัมน์
  // เป็นของตัวเอง (self-contained)
  const columnsOf = (m) => additions[m]?.$columns ?? additions.$columns ?? [];
  const columns = draftAdd.$columns ?? columnsOf(month);
  const isMultiCol = columns.length > 0;

  // A month counts as "saved" (and therefore locked, requiring "แก้ไข" to
  // unlock) once it carries the explicit $saved flag, or — for months saved
  // before this flag existed — once it already has any real entered value.
  const monthHasData = Object.entries(additions[month]||{}).some(([k,v]) => !k.startsWith("$") && parseFloat(v));
  const monthSaved = additions[month]?.$saved === true || monthHasData;
  const editingUnlocked = !monthSaved || forceEdit;

  // ซิงก์ draftAdd จากค่าที่บันทึกไว้ — รีเซ็ตเสมอเมื่อ "สลับเดือน" แต่ "ห้าม" เขียนทับ
  // สิ่งที่กำลังพิมพ์ค้างระหว่างแก้ไข (เช่น realtime refetch ตอนคนอื่นเซฟในโครงการเดียวกัน)
  const prevMonthRef = useRef(month);
  useEffect(() => {
    const switched = prevMonthRef.current !== month;
    prevMonthRef.current = month;
    if (switched || !editingUnlocked) setDraftAdd({ ...(additions[month] || {}) });
  }, [month, additions, editingUnlocked]);

  // ตั้งธง "แก้ค้างยังไม่บันทึก" เมื่อ draftAdd ต่างจากค่าที่บันทึกไว้ของเดือนนี้
  const isDirty = editingUnlocked && JSON.stringify(draftAdd) !== JSON.stringify(additions[month] || {});
  useEffect(() => { UnsavedGuard.dirty = isDirty; return () => { UnsavedGuard.dirty = false; }; }, [isDirty]);

  const [monthEditMode, setMonthEditMode] = useState(false); // โหมดจัดการเดือน (เพิ่ม/ลบเดือน) แยกจากการแก้ค่าในตาราง
  // เลื่อนแถวชิปเดือนให้เดือนที่เลือกอยู่ในสายตาเสมอ (เช่นตอนคลิกแท่งกราฟ)
  const activeChipRef = useRef(null);
  useEffect(() => { activeChipRef.current?.scrollIntoView({ behavior:"smooth", inline:"center", block:"nearest" }); }, [month]);
  useEffect(() => { setMonthEditMode(false); }, [month]);     // สลับเดือนแล้วปิดโหมดจัดการเดือน
  useEffect(() => { setEditMode?.(editingUnlocked || monthEditMode); return () => setEditMode?.(false); }, [editingUnlocked, monthEditMode, setEditMode]);

  // ยกเลิกการแก้ไข: ทิ้งค่าที่พิมพ์ค้างของเดือนนี้ คืนเป็นค่าที่บันทึกไว้ แล้วล็อก/ออกจากโหมดจัดการเดือน
  const canCancel = monthSaved;
  const handleCancel = () => { setDraftAdd({ ...(additions[month] || {}) }); setForceEdit(false); setMonthEditMode(false); setAddExtraOpen(false); setSubFor(null); setAddColOpen(false); };
  useEffect(() => {
    if (!editingUnlocked && !monthEditMode) return;
    // Esc = ยกเลิกการแก้ไข — แต่ไม่ทิ้งข้อมูลโดยไม่ตั้งใจ:
    //  • ถ้า Esc ถูกใช้ไปแล้ว (ยกเลิกการเลือกช่อง) หรือกำลังพิมพ์ในช่องกรอก → ไม่ทำอะไร
    //  • เดือนที่ยังไม่เคยบันทึก (ไม่มีปุ่มยกเลิก) → ไม่ทำอะไร
    //  • มีค่าที่พิมพ์ค้าง → ถามยืนยันก่อนทิ้ง
    const onEsc = (e) => {
      if (e.key !== "Escape" || e.defaultPrevented) return;
      const ae = document.activeElement;
      if (ae && (ae.tagName === "INPUT" || ae.tagName === "TEXTAREA" || ae.tagName === "SELECT" || ae.isContentEditable)) return;
      if (!canCancel) return;
      e.preventDefault();
      if (UnsavedGuard.dirty) { uiConfirm(t("ยกเลิกการแก้ไขเดือนนี้? ค่าที่พิมพ์ไว้แต่ยังไม่บันทึกจะหายไป","Cancel editing this month? Unsaved values will be lost"), { danger: true, okLabel: t("ทิ้งการแก้ไข","Discard changes"), cancelLabel: t("แก้ต่อ","Keep editing") }).then(ok => { if (ok) handleCancel(); }); return; }
      handleCancel();
    };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [editingUnlocked, monthEditMode, month, additions, canCancel]);

  // All rows = original 70 account codes + standalone extra items.
  // Sub-items (parentCode set) can be added right here for a monthly
  // breakdown, or on the Baseline tab for a baseline breakdown — either way
  // they roll up into their parent row's figures and aren't listed on their own.
  const allRows = [...ACCOUNTS.filter(a=>!hiddenAccounts.includes(a.code)), ...extraItems.filter(e=>!e.parentCode).map(e => ({ code:e.code, name:e.name, group:e.group, isExtra:true }))];

  const subItemsByParent = {};
  extraItems.forEach(e => { if (e.parentCode) (subItemsByParent[e.parentCode] = subItemsByParent[e.parentCode] || []).push(e); });
  const childrenOf = (code) => subItemsByParent[code] || [];

  // A row's monthly figure is: the sum of its sub-items (if it has any) —
  // else the sum across that month's named columns (if the month uses them) —
  // else its own plain entered value. Mirrors the Baseline tab's rollup logic.
  const kidsAsOf = (code, m) => childrenOf(code).filter(k => !k.addedInMonth || k.addedInMonth <= m);
  const rowMonthValue = (code, m, draft) => {
    // ── ระหว่างแก้ไข (มี draft): คิดสด ๆ จากค่าที่พิมพ์ = ผลรวมรายการย่อย/คอลัมน์
    //    (คอลัมน์ที่ถูกลบออกจากร่างจะไม่ถูกนับ เพราะไม่อยู่ใน columns) ──
    if (draft) {
      const kids = kidsAsOf(code, m);
      if (kids.length) return kids.reduce((s,k)=>s+(parseFloat(draft[k.code])||0),0);
      if (columns.length) return columns.reduce((s,c)=>s+(parseFloat(draft[`${code}:${c.id}`])||0),0);
      return parseFloat(draft[code])||0;
    }
    // ── ข้อมูลที่บันทึกแล้ว: ค่าธรรมดา (code) คือ "ยอดรวมที่ roll-up ไว้แล้ว"
    //    (handleSave ตั้งค่านี้ = ผลรวมคอลัมน์/รายการย่อยเสมอ) จึงอ่านตัวเดียวพอ
    //    — ไม่บวกคอลัมน์ซ้ำ (กันนับซ้ำ) และคอลัมน์ที่ลบไปแล้วก็ถูก roll-up ใหม่ไม่รวมมัน ──
    return parseFloat(additions[m]?.[code]) || 0;
  };

  const monthTotal = (m) => allRows.reduce((s,r) => s + rowMonthValue(r.code, m), 0);

  // Sum only top-level rows (accounts + standalone extras). Do NOT sum
  // Object.values(tenderCosts) directly — sub-item codes (EX-xxxx with a
  // parentCode) also have their own entries in tenderCosts, and their total
  // is already rolled up into their parent's value, so a wholesale sum
  // double-counts every account that has sub-items.
  const baseTotal = allRows.reduce((s,r) => s + withWaste(tenderCosts[r.code]), 0);   // ราคาเดิม + เผื่อเศษ (รวมในงบ)
  const thisMonthAdd = allRows.reduce((s,r) => s + rowMonthValue(r.code, month, draftAdd), 0);
  const cumulativeSoFar = months.filter(m=>m<month).reduce((s,m)=>s+monthTotal(m),0) + thisMonthAdd + baseTotal;

  // "Live" versions that use the currently-edited draft for the selected month
  // (instead of the last-saved value) so the top summary updates as you type.
  const monthTotalLive = (m) => m===month ? thisMonthAdd : monthTotal(m);
  const sortedMonths = months.length ? months : [thisMonth];
  const cumulativeLive = (uptoMonth) => baseTotal + sortedMonths.filter(m=>m<=uptoMonth).reduce((s,m)=>s+monthTotalLive(m),0);
  const grandTotal = cumulativeLive(sortedMonths[sortedMonths.length-1]);
  // Each bar = one stacked column: "previous" (running total up to the
  // month before) + "added" (that month's increment) in a different color,
  // so growth is visible within a single bar instead of a smooth area line.
  const chartData = [
    { label:t("เริ่มต้น","Start"), cumulative: baseTotal, previous: baseTotal, added: 0 },
    ...sortedMonths.map(m => {
      const added = monthTotalLive(m);
      const cumulative = cumulativeLive(m);
      return { label: monthShortLabel(m), monthKey: m, cumulative, previous: cumulative - added, added };
    }),
  ];
  const selectedLabel = (chartData.find(e => e.monthKey === month) || {}).label;   // ป้ายเดือนที่กำลังเลือกอยู่บนกราฟ

  // "ราคาเดิม (Baseline)" should reflect the running total as of the month
  // BEFORE the one currently selected — not the fixed original baseline —
  // so it moves forward as prior months get their additions saved.
  const priorMonths      = sortedMonths.filter(m => m < month);
  const prevMonthLabel   = priorMonths.length ? monthShortLabel(priorMonths[priorMonths.length-1]) : t("เริ่มต้น","Start");
  const baselineForMonth = baseTotal + priorMonths.reduce((s,m)=>s+monthTotalLive(m),0);

  const filtered = allRows.filter(r => {
    if (filter.length && !filter.includes(r.group)) return false;
    const q = search.toLowerCase();
    const selfMatch = r.name.toLowerCase().includes(q) || r.code.includes(search);
    const childMatch = kidsAsOf(r.code, month).some(k=>k.name.toLowerCase().includes(q));
    return selfMatch || childMatch;
  });

  // Mirrors the per-row figures computed inline in the table body, so header
  // sorting can order rows by the same "ยอดก่อนหน้า / เพิ่มเดือนนี้ / รวมสะสม" values shown.
  const cumBeforeOf = (r) => months.filter(m=>m<month).reduce((s,m)=>s+rowMonthValue(r.code, m),0) + withWaste(tenderCosts[r.code]);
  const cumOf = (r) => cumBeforeOf(r) + rowMonthValue(r.code, month, draftAdd);

  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => -d);
    else { setSortKey(key); setSortDir(1); }
  };
  const displayRows = (() => {
    // ซ่อนแถวที่ไม่มีค่า = รวมสะสมของเดือนนี้เป็น 0 (ทั้งยอดยกมาและเพิ่มเดือนนี้ว่าง)
    const base = hideEmpty ? filtered.filter(r => cumOf(r) !== 0) : filtered;
    if (!sortKey) return base;
    const arr = [...base];
    arr.sort((a, b) => {
      let av, bv;
      if (sortKey === "code")        { av = a.code; bv = b.code; }
      else if (sortKey === "group")  { av = GROUPS.indexOf(a.group); bv = GROUPS.indexOf(b.group); }
      else if (sortKey === "name")   { av = a.name; bv = b.name; }
      else if (sortKey === "before") { av = cumBeforeOf(a); bv = cumBeforeOf(b); }
      else if (sortKey === "add")    { av = rowMonthValue(a.code, month, draftAdd); bv = rowMonthValue(b.code, month, draftAdd); }
      else                           { av = cumOf(a); bv = cumOf(b); }
      if (typeof av === "string") return av.localeCompare(bv) * sortDir;
      return (av - bv) * sortDir;
    });
    return arr;
  })();
  const hiddenEmptyCount = hideEmpty ? filtered.length - filtered.filter(r => cumOf(r) !== 0).length : 0;

  // สลับเดือน — ถ้ามีค่าที่พิมพ์ค้างยังไม่บันทึก ถามก่อน (เดิมสลับแล้วค่าหายเงียบ ๆ)
  const goMonth = (m) => {
    if (!m || m === month) return;
    leaveIfDirty(() => setMonth(m));
  };
  const handleAddMonth = () => {
    if (!newMonth) return;
    if (newMonth !== month) leaveIfDirty(addMonthNow); else { UnsavedGuard.dirty = false; addMonthNow(); }
  };
  const addMonthNow = async () => {
    if (months.includes(newMonth)) {
      // ห้ามซ้ำ — ถ้ามีเดือนนี้อยู่แล้ว แค่กระโดดไปที่เดือนนั้นแทนการสร้างซ้ำ
      uiAlert(t(`มีเดือน ${monthShortLabel(newMonth)} อยู่แล้ว`, `${monthShortLabel(newMonth)} already exists`));
      setMonth(newMonth); setNewMonth("");
      return;
    }
    // ถ้าเดือนล่าสุดมีคอลัมน์อยู่ ให้ถามก่อนว่าจะคัดลอกมาที่เดือนใหม่ไหม
    const prevMonth = months.length ? months[months.length - 1] : null;
    const prevCols = prevMonth ? columnsOf(prevMonth) : [];
    const monthObj = {};
    if (prevCols.length) {
      if (await uiConfirm(t(`คัดลอกคอลัมน์จากเดือน ${monthShortLabel(prevMonth)} มาที่เดือนใหม่ไหม?\n(${prevCols.map(c=>c.name).join(", ")})\n\nOK = คัดลอกคอลัมน์ (ยอดเริ่มที่ว่าง) · Cancel = เริ่มเดือนใหม่แบบไม่มีคอลัมน์`, `Copy columns from ${monthShortLabel(prevMonth)} into the new month?\n(${prevCols.map(c=>c.name).join(", ")})\n\nOK = copy columns (values start empty) · Cancel = start the new month with no columns`), { okLabel: t("คัดลอกคอลัมน์","Copy columns"), cancelLabel: t("เริ่มแบบว่าง","Start blank") })) {
        monthObj.$columns = prevCols.map(c => ({ ...c }));
      } else {
        monthObj.$columns = []; // เริ่มใหม่แบบไม่มีคอลัมน์ (กัน fallback ไป global เดิม)
      }
    }
    saveAdditions({ ...additions, [newMonth]: monthObj });
    setMonth(newMonth); setNewMonth("");
  };

  // ลบเดือน — เอาข้อมูลที่เพิ่มในเดือนนั้นออกทั้งหมด (คีย์ meta อย่าง $columns
  // ที่เป็นระดับโปรเจกต์ไม่ถูกแตะ) แล้วถ้าลบเดือนที่กำลังดูอยู่ก็ย้ายไปเดือนอื่น
  const handleDeleteMonth = async (m) => {
    if (!(await uiConfirm(t(`ลบเดือน ${monthShortLabel(m)} และข้อมูลที่เพิ่มในเดือนนี้ทั้งหมด?\n(ราคาเดิม/Baseline ไม่ได้รับผลกระทบ)`, `Delete ${monthShortLabel(m)} and all additions entered in this month?\n(Baseline is not affected)`), { danger: true, okLabel: t("ลบเดือน","Delete month") }))) return;
    const next = { ...additions };
    delete next[m];
    saveAdditions(next);
    if (month === m) {
      const remaining = Object.keys(next).filter(k=>!k.startsWith("$")).sort();
      setMonth(remaining.length ? remaining[remaining.length-1] : thisMonth);
    }
  };

  const handleSave = () => {
    const merged = {...draftAdd};
    allRows.forEach(r => {
      const kids = kidsAsOf(r.code, month);
      if (kids.length) merged[r.code] = kids.reduce((s,k)=>s+(parseFloat(merged[k.code])||0),0);
      else if (columns.length) merged[r.code] = columns.reduce((s,c)=>s+(parseFloat(merged[`${r.code}:${c.id}`])||0),0);
    });
    const clean = {};
    Object.entries(merged).forEach(([k,v]) => {
      if (k.startsWith("$")) return; // meta keys ($saved) are re-added explicitly below
      if(v!==""&&!isNaN(v)&&parseFloat(v)!==0) clean[k]=parseFloat(v);
    });
    clean.$saved = true;
    const ownCols = draftAdd.$columns;
    if (columns.length) clean.$columns = columns;
    else if (Array.isArray(ownCols)) clean.$columns = []; // เดือนนี้ตั้งใจไม่มีคอลัมน์ (กัน fallback ไป global เดิม)
    if (draftAdd.$fmt && Object.keys(draftAdd.$fmt).length) clean.$fmt = draftAdd.$fmt; // เก็บสีไฮไลต์ที่ผู้ใช้ทำไว้
    saveAdditions({ ...additions, [month]: clean });
    setForceEdit(false); setAddExtraOpen(false); setSubFor(null); setAddColOpen(false);
    setSaved(true); setTimeout(()=>setSaved(false),2000);
  };

  const handleCreateExtra = () => {
    if (!extraDraft.name.trim()) return;
    const code = extraDraft.code.trim();
    if (code) {
      const taken = ACCOUNTS.some(a=>a.code===code) || extraItems.some(e=>e.code===code);
      if (taken) { uiAlert(t(`Acc. Code "${code}" มีอยู่แล้ว กรุณาใช้รหัสอื่น`, `Acc. Code "${code}" already exists, please use another`)); return; }
    }
    onAddExtra({ name:extraDraft.name, group:extraDraft.group, code: code || undefined });
    setExtraDraft({ code:"", name:"", group:GROUPS[0] }); setAddExtraOpen(false);
  };

  const handleAddSub = (parentCode) => {
    if (!subName.trim()) return;
    // รายการย่อยแรกของแถวที่มียอดอยู่แล้ว: ย้ายยอดเดิม (เดือนนี้ + เดือนหลังจากนี้ที่ยังไม่มีรายการย่อย)
    // ไปไว้ในรายการย่อย "รายการเดิม" — ยอดรวมทุกเดือนเท่าเดิม ไม่หายตอนบันทึก
    const firstKid = kidsAsOf(parentCode, month).length === 0;
    const draftVal = columns.length
      ? columns.reduce((s,c)=>s+(parseFloat(draftAdd[`${parentCode}:${c.id}`])||0),0)
      : (parseFloat(draftAdd[parentCode])||0);
    const laterSaved = months.filter(m => m >= month && kidsAsOf(parentCode, m).length === 0 && monthAddValue(additions, m, parentCode) !== 0);
    if (firstKid && (draftVal !== 0 || laterSaved.length)) {
      const res = onAddExtra({ name:subName, parentCode, addedInMonth: month, seedName: t("รายการเดิม","Original item") });
      if (res?.seedCode) {
        if (laterSaved.length) {
          const nextAdd = { ...additions };
          laterSaved.forEach(m => { nextAdd[m] = { ...nextAdd[m], [res.seedCode]: monthAddValue(additions, m, parentCode) }; });
          saveAdditions(nextAdd);
        }
        setDraftAdd(d => { const n = { ...d, [res.seedCode]: draftVal }; columns.forEach(c => { delete n[`${parentCode}:${c.id}`]; }); return n; });
      }
    } else {
      onAddExtra({ name:subName, parentCode, addedInMonth: month });
    }
    setRowCollapsed(c => ({...c, [parentCode]: false})); // reveal the newly-added sub-item
    setSubName(""); setSubFor(null);
  };

  const handleDeleteExtra = (code) => {
    Promise.resolve(onDeleteExtra(code)).then(ok => { if (ok === false) return; setDraftAdd(d => { const n = {...d}; delete n[code]; return n; }); });
  };

  // เพิ่ม "รายการ" (คอลัมน์ย่อย) เฉพาะเดือนที่กำลังดูอยู่ (ต่อเดือน ไม่ลามไปเดือนอื่น)
  // ครั้งแรกที่สร้างคอลัมน์ในเดือนนี้ จะพับค่าที่กรอกแบบช่องเดียวเดิมของเดือนนี้เข้า
  // เป็นคอลัมน์ "รายการหลัก" ก่อน เพื่อไม่ให้ค่าที่กรอกไว้หาย
  const handleAddColumn = () => {
    const name = newColName.trim();
    if (!name) return;
    const newCol = { id: uid(), name };
    const nextDraft = { ...draftAdd };
    if (columns.length === 0) {
      const seedCol = { id: "legacy", name: t("รายการหลัก","Main item") };
      allRows.forEach(r => {
        const v = nextDraft[r.code];
        if (v !== undefined && v !== "" && parseFloat(v)) nextDraft[`${r.code}:legacy`] = v;
      });
      nextDraft.$columns = [seedCol, newCol];
    } else {
      nextDraft.$columns = [...columns, newCol];
    }
    // เป็นแค่ "ร่าง" เหมือนการลบคอลัมน์ — มีผลจริงเมื่อกด "บันทึก" (ซึ่งจะรวมยอดคอลัมน์ลงคีย์หลัก
    // ให้ฝ่ายอื่นเห็นยอดถูก) · เดิมบันทึกทันทีแบบไม่รวมยอด ทำให้ยอดหายจากงบของฝ่ายอื่น/ยกเลิกไม่ได้
    setDraftAdd(nextDraft);
    setNewColName(""); setAddColOpen(false);
  };

  // ลบคอลัมน์ — เฉพาะเดือนนี้ และเป็นแค่ "ร่าง" เท่านั้น จะมีผลจริงเมื่อกด "บันทึก"
  // ถ้ากด "ยกเลิก" คอลัมน์และค่าที่กรอกไว้จะกลับคืนมา (ไม่โดนลบ) และคอลัมน์ที่ลบ
  // ไปแล้วจะไม่ถูกนำไปคิดยอด (เพราะยอด roll-up ตอนบันทึกจะไม่รวมคอลัมน์นั้น)
  const handleRemoveColumn = async (colId) => {
    if (!(await uiConfirm(t("ลบคอลัมน์นี้เฉพาะเดือนนี้?\n\n• จะมีผลจริงเมื่อกด \"บันทึก\"\n• กด \"ยกเลิก\" เพื่อคืนคอลัมน์และค่าที่กรอกไว้","Delete this column for this month only?\n\n• Takes effect when you press \"Save\"\n• Press \"Cancel\" to restore the column and entered values"), { danger: true, okLabel: t("ลบคอลัมน์","Delete column") }))) return;
    const nextCols = columns.filter(c => c.id !== colId);
    const nextDraft = { ...draftAdd };
    Object.keys(nextDraft).forEach(k => { if (k.endsWith(`:${colId}`)) delete nextDraft[k]; });
    nextDraft.$columns = nextCols;
    setDraftAdd(nextDraft);
    // ไม่ saveAdditions ที่นี่ — รอกด "บันทึก" (handleSave) เท่านั้น เพื่อให้ยกเลิกได้
  };

  // Excel-style block paste: paste a copied range from Excel/Sheets straight into
  // the grid. Starting from the focused cell, values flow down (rows) and right
  // (columns), matching whatever the user copied. Only leaf rows take a value —
  // parent rows (with sub-items) show a roll-up total and are skipped. Blank
  // cells in the pasted block are left untouched so pasting one column can't wipe
  // the others. Values land in the draft; the user still presses "Save this month".
  const handleGridPaste = async (startRowIdx, startColIdx, raw) => {
    if (!editingUnlocked) return;
    const grid = String(raw ?? "")
      .replace(/\r\n?/g, "\n")
      .replace(/\n+$/, "")
      .split("\n")
      .map(line => line.split("\t"));
    if (!grid.length) return;
    // วางหลายแถว = ลงตาม "ลำดับแถวที่เห็นบนจอ" — ถ้ากำลังกรอง/ค้นหา/เรียง/ซ่อนแถวว่าง ลำดับจะไม่ตรงกับ
    // ชีต Excel ต้นทาง ค่าอาจลงผิด Acc. Code จึงถามก่อน และบอกว่าแถวแรก/สุดท้ายจะลงที่รหัสไหน
    const viewChanged = filter.length > 0 || !!search.trim() || hideEmpty || !!sortKey;
    if (grid.length > 1 && viewChanged) {
      const first = displayRows[startRowIdx], last = displayRows[Math.min(startRowIdx + grid.length - 1, displayRows.length - 1)];
      if (!(await uiConfirm(t(
        `ตอนนี้ตารางถูกกรอง/ค้นหา/เรียง/ซ่อนแถวว่างอยู่ — ค่า ${grid.length} แถวจะลงตามลำดับที่เห็นบนจอ\nแถวแรก → ${first?.code || "-"} · แถวสุดท้าย → ${last?.code || "-"}\n\nถ้าคัดลอกมาจากชีตที่เรียงตาม Acc. Code ให้ยกเลิก แล้วล้างตัวกรอง/การเรียงก่อนวาง\nวางต่อไหม?`,
        `The table is filtered/searched/sorted or hiding empty rows — ${grid.length} rows will be pasted in on-screen order\nFirst row → ${first?.code || "-"} · last row → ${last?.code || "-"}\n\nIf you copied from a sheet in Acc. Code order, cancel and clear the filter/sort first.\nPaste anyway?`), { okLabel: t("วางเลย","Paste anyway") }))) return;
    }
    setDraftAdd(d => {
      const next = { ...d };
      grid.forEach((cells, ri) => {
        const row = displayRows[startRowIdx + ri];
        if (!row) return;
        if (kidsAsOf(row.code, month).length > 0) return; // parent roll-up row — no direct input
        cells.forEach((cellRaw, ci) => {
          if (String(cellRaw).trim() === "") return; // don't overwrite with blanks
          const val = evalMoney(cellRaw, true);   // ยอดเพิ่มรายเดือนติดลบได้ (งานลด) · "(12,000)" = -12,000
          if (val === "") return;
          if (isMultiCol) {
            const col = columns[startColIdx + ci];
            if (!col) return; // ignore columns beyond the current ones
            next[`${row.code}:${col.id}`] = val;
          } else {
            if (ci > 0) return; // single-column grid — only first pasted column applies
            next[row.code] = val;
          }
        });
      });
      return next;
    });
  };

  // ── Excel-style cell selection ───────────────────────────────────────────────
  // Selection is a Set of "ri:ci" keys (row-index in displayRows × column-index),
  // so it supports non-rectangular multi-selection like Excel:
  //   • plain drag / click  → replace with a rectangle
  //   • Shift + click        → extend the rectangle from the anchor
  //   • Ctrl/Cmd + click     → toggle a single cell (add, or de-select it)
  //   • Ctrl/Cmd + drag      → add a rectangle to what's already selected
  // Then Ctrl/Cmd+C copies it out as TSV (pastes cleanly into Excel/Sheets) and
  // Delete/Backspace clears it. Parent roll-up rows carry no direct value.
  const colCount = isMultiCol ? columns.length : 1;
  const [selSet, setSelSet] = useState(() => new Set());
  const selDragRef = useRef(false);
  const anchorRef = useRef(null);   // {ri,ci} for shift-extend / drag origin
  const dragModeRef = useRef(null); // {mode:"replace"|"add", base:Set}
  const inSel = (ri, ci) => selSet.has(ri + ":" + ci);
  const selCount = selSet.size;
  // คืน null ถ้าคอลัมน์ไม่มีอยู่แล้ว (เช่นเลือกช่องค้างไว้แล้วลบคอลัมน์/สลับเดือน) — กันหน้าพัง
  const cellKeyOf = (row, ci) => {
    if (!row || ci < 0) return null;
    if (!isMultiCol) return ci === 0 ? row.code : null;
    const col = columns[ci];
    return col ? `${row.code}:${col.id}` : null;
  };
  const cellValStr = (row, ci) => {
    if (kidsAsOf(row.code, month).length > 0) return ""; // parent roll-up — no direct value
    const raw = draftAdd[cellKeyOf(row, ci)];
    const n = parseFloat(raw);
    return (raw == null || raw === "" || isNaN(n)) ? "" : String(n);
  };
  const rectKeys = (a, b) => {
    const keys = [];
    const r0 = Math.min(a.ri, b.ri), r1 = Math.max(a.ri, b.ri);
    const c0 = Math.min(a.ci, b.ci), c1 = Math.max(a.ci, b.ci);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) keys.push(r + ":" + c);
    return keys;
  };
  const onCellDown = (ri, ci, e) => {
    // เลือกเซลล์ได้ทั้งโหมดดูและแก้ไข (โหมดดูใช้สำหรับไฮไลต์/เปลี่ยนสีอย่างเดียว)
    const additive = e.ctrlKey || e.metaKey;
    const ranged   = e.shiftKey && anchorRef.current;
    // pure selection op — don't focus/edit the input or start a text selection.
    // In view mode there's no input to focus, so always prevent (avoids selecting the numbers while dragging).
    if (ranged || additive || !editingUnlocked) e.preventDefault();
    if (ranged) {
      setSelSet(new Set(rectKeys(anchorRef.current, { ri, ci })));
      dragModeRef.current = { mode: "replace" };
    } else if (additive) {
      const base = new Set(selSet);
      const k = ri + ":" + ci;
      if (base.has(k)) base.delete(k); else base.add(k); // toggle → lets you de-select
      setSelSet(base);
      anchorRef.current = { ri, ci };
      dragModeRef.current = { mode: "add", base: new Set(base) };
    } else {
      setSelSet(new Set([ri + ":" + ci]));
      anchorRef.current = { ri, ci };
      dragModeRef.current = { mode: "replace" };
    }
    selDragRef.current = true;
  };
  const onCellEnter = (ri, ci) => {
    if (!selDragRef.current || !dragModeRef.current || !anchorRef.current) return;
    const rk = rectKeys(anchorRef.current, { ri, ci });
    if (dragModeRef.current.mode === "add") {
      const s = new Set(dragModeRef.current.base);
      rk.forEach(k => s.add(k));
      setSelSet(s);
    } else {
      setSelSet(new Set(rk));
    }
    const ae = typeof document !== "undefined" ? document.activeElement : null;
    if (ae && ae.tagName === "INPUT" && (ri !== anchorRef.current.ri || ci !== anchorRef.current.ci)) ae.blur();
  };
  const selectAll = () => {
    const s = new Set();
    for (let r = 0; r < displayRows.length; r++) for (let c = 0; c < colCount; c++) s.add(r + ":" + c);
    setSelSet(s); anchorRef.current = { ri: 0, ci: 0 };
  };
  const deselectAll = () => { setSelSet(new Set()); anchorRef.current = null; };
  // การเลือกช่องอ้างอิง "ตำแหน่งแถว/คอลัมน์บนจอ" — ถ้าเดือน/คอลัมน์/การเรียง/ตัวกรอง/ค้นหา/ซ่อนแถวว่างเปลี่ยน
  // ตำแหน่งจะชี้ไปคนละ Acc. Code แล้ว จึงล้างการเลือกทิ้ง (กันกด Delete แล้วลบผิดแถว)
  useEffect(() => { deselectAll(); }, [month, columns.length, sortKey, sortDir, filter, search, hideEmpty, displayRows.length]); // eslint-disable-line
  const buildSelTSV = () => {
    if (!selSet.size) return "";
    let r0 = Infinity, r1 = -Infinity, c0 = Infinity, c1 = -Infinity;
    selSet.forEach(k => { const [r, c] = k.split(":").map(Number); r0 = Math.min(r0, r); r1 = Math.max(r1, r); c0 = Math.min(c0, c); c1 = Math.max(c1, c); });
    const out = [];
    for (let r = r0; r <= r1; r++) {
      const row = displayRows[r];
      const cells = [];
      for (let c = c0; c <= c1; c++) cells.push(row && selSet.has(r + ":" + c) ? cellValStr(row, c) : "");
      out.push(cells.join("\t"));
    }
    return out.join("\n");
  };
  const clearSelection = () => {
    if (!selSet.size || !editingUnlocked) return;
    setDraftAdd(d => {
      const next = { ...d };
      selSet.forEach(k => {
        const [r, c] = k.split(":").map(Number);
        const row = displayRows[r]; if (!row) return;
        if (kidsAsOf(row.code, month).length > 0) return;
        const key = cellKeyOf(row, c); if (!key) return;
        next[key] = "";
      });
      return next;
    });
  };
  const copySelection = async () => {
    const tsv = buildSelTSV();
    if (!tsv) return;
    try { await navigator.clipboard.writeText(tsv); }
    catch { /* clipboard API blocked — user can still use Ctrl/Cmd+C */ }
  };
  // ── ไฮไลต์เอง: ลากเลือกเซลล์แล้วใส่สีพื้น (bg) หรือสีตัวอักษร (fg) เก็บไว้ใน
  // additions[month].$fmt (คีย์ $… ถูกข้ามจากการรวมยอดอยู่แล้ว) → บันทึกติดไปกับเดือน
  const cellFmt = draftAdd.$fmt || {};
  // แปลง fmt ของเซลล์ → CSS (พื้น/สีตัวอักษร/หนา/เอียง/ขีดเส้นใต้/ขนาด)
  const cellFmtStyle = (key) => {
    const f = cellFmt[key]; if (!f) return {};
    return {
      ...(f.bg?{background:f.bg}:{}), ...(f.fg?{color:f.fg}:{}),
      ...(f.b?{fontWeight:800}:{}), ...(f.i?{fontStyle:"italic"}:{}),
      ...(f.u?{textDecoration:"underline"}:{}), ...(f.sz?{fontSize:f.sz}:{}),
    };
  };
  // รายคีย์เซลล์ (leaf) ที่กำลังเลือกอยู่ — ใช้ร่วมกันทุกเครื่องมือจัดรูปแบบ
  const selectedCellKeys = () => {
    const keys = [];
    selSet.forEach(k => { const [r,c]=k.split(":").map(Number); const row=displayRows[r]; if(!row) return; if(kidsAsOf(row.code,month).length>0) return; const key=cellKeyOf(row,c); if(key) keys.push(key); });
    return keys;
  };
  const mutateFmt = (fn) => { // fn(cur) → คืน object ใหม่ (หรือ null เพื่อลบ) · ใช้ได้ทั้งโหมดดู/แก้ไข
    if (!selSet.size) return;
    const fmt = { ...(draftAdd.$fmt || {}) };
    selectedCellKeys().forEach(key => {
      let cur = fn({ ...(fmt[key] || {}) }) || {};
      Object.keys(cur).forEach(p => { if (cur[p] == null || cur[p] === false) delete cur[p]; });
      if (Object.keys(cur).length) fmt[key] = cur; else delete fmt[key];
    });
    setDraftAdd(d => ({ ...d, $fmt: fmt }));
    // โหมดดู (ไม่ได้แก้ไข): ค่าตัวเลขล็อกอยู่ จึงเซฟ "เฉพาะสี" ขึ้น backend ทันที ไม่แตะค่าที่บันทึกไว้
    if (!editingUnlocked) {
      const base = additions[month] || {};
      saveAdditions({ ...additions, [month]: { ...base, $fmt: fmt } });
    }
  };
  const applyCellFmt = (patch) => mutateFmt(cur => ({ ...cur, ...patch })); // {bg}/{fg}/{bg:null}…
  const toggleCellFmt = (prop) => { // สลับ หนา/เอียง/ขีดเส้นใต้
    const keys = selectedCellKeys();
    const allOn = keys.length>0 && keys.every(key => (cellFmt[key]||{})[prop]);
    mutateFmt(cur => ({ ...cur, [prop]: allOn ? null : true }));
  };
  const bumpFontSize = (delta) => mutateFmt(cur => { const base = cur.sz || 13; return { ...cur, sz: Math.max(9, Math.min(22, base + delta)) }; });
  useEffect(() => {
    const up = () => { selDragRef.current = false; dragModeRef.current = null; };
    const onCopy = (e) => {
      if (!selSet.size) return;
      const ae = document.activeElement;
      if (ae && ae.tagName === "INPUT" && selSet.size === 1) return; // single active cell → let the input copy normally
      const tsv = buildSelTSV();
      if (!tsv) return;
      e.clipboardData.setData("text/plain", tsv);
      e.preventDefault();
    };
    const onKey = (e) => {
      if (!selSet.size) return;
      // Esc ตอนมีช่องที่เลือก = แค่ยกเลิกการเลือก (preventDefault บอกตัวจัดการ Esc ของหน้าว่า "ใช้ไปแล้ว" ไม่ต้องยกเลิกการแก้ไข)
      if (e.key === "Escape") { e.preventDefault(); deselectAll(); return; }
      const ae = document.activeElement;
      const editing = ae && ae.tagName === "INPUT";
      if ((e.key === "Delete" || e.key === "Backspace") && !editing && editingUnlocked) {
        e.preventDefault();
        clearSelection();
      }
    };
    document.addEventListener("mouseup", up);
    document.addEventListener("copy", onCopy);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mouseup", up);
      document.removeEventListener("copy", onCopy);
      document.removeEventListener("keydown", onKey);
    };
  });

  return (
    <div style={{padding:"4px 28px 24px"}}>
      {/* Trend chart — the whole project's cost growth over time, at a glance */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"18px 20px 8px",marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap",gap:8}}>
          <span style={{fontSize:13,fontWeight:650,color:T.textPrimary}}><Ico name="trend" size={16} color={T.textSecondary} /> {t("แนวโน้มต้นทุนสะสม","Cumulative cost trend")}</span>
          <span style={{fontSize:12,color:T.textMuted}}>{t("รวมล่าสุดทั้งโปรเจกต์","Project latest total")}: <b style={{color:T.green,fontVariantNumeric:"tabular-nums",fontSize:15}}>฿{fmt0(grandTotal)}</b>{usdRate>0 && <b className="usd-sub" style={{color:T.green,fontVariantNumeric:"tabular-nums",fontSize:12,marginLeft:6}}>≈ ${fmt(grandTotal/usdRate)}</b>}</span>
        </div>
        <div style={{display:"flex",gap:16,marginBottom:6,fontSize:11,color:T.textMuted,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{display:"inline-flex",alignItems:"center",gap:5}}><span style={{width:10,height:10,borderRadius:2,background:T.blue,display:"inline-block"}}/>{t("ยอดก่อนหน้า (สะสม)","Previous (cumulative)")}</span>
          <span style={{display:"inline-flex",alignItems:"center",gap:5}}><span style={{width:10,height:10,borderRadius:2,background:T.amber,display:"inline-block"}}/>{t("เพิ่มเดือนนี้","Added this month")}</span>
          <span style={{color:T.textMuted,fontSize:11}}>· {t("คลิกที่แท่งเพื่อเลือกเดือน (เดือนที่เลือกจะมีกรอบ)","Click a bar to select a month (selected has an outline)")}</span>
        </div>
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={chartData} margin={{top:14,right:8,left:-14,bottom:6}} barCategoryGap="22%"
            onClick={(st)=>{ const mk = st && st.activePayload && st.activePayload[0] && st.activePayload[0].payload && st.activePayload[0].payload.monthKey; if (mk) goMonth(mk); }}
            style={{cursor:"pointer"}}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f7"/>
            <XAxis dataKey="label" tick={<MonthAxisTick selectedLabel={selectedLabel}/>} height={34} axisLine={false} tickLine={false} interval={0}/>
            <YAxis tick={{fontSize:11,fill:T.textMuted}} axisLine={false} tickLine={false} tickFormatter={fmtK}/>
            <Tooltip cursor={{fill:"rgba(37,99,235,0.06)"}} formatter={(v,name)=>[`${fmt(v)} THB`,name]} labelStyle={{color:T.textPrimary,fontWeight:600,marginBottom:2}}
              contentStyle={{borderRadius:10,border:`1px solid ${T.cardBorder}`,fontSize:12,boxShadow:"0 4px 14px rgba(0,0,0,0.08)"}}/>
            <Bar dataKey="previous" stackId="cum" name={t("ยอดก่อนหน้า","Previous")} radius={[0,0,0,0]}>
              {chartData.map((e,i)=>{ const sel = e.monthKey===month; return <Cell key={i} fill={T.blue} stroke={sel?"#0f172a":"none"} strokeWidth={sel?2.5:0} cursor="pointer"/>; })}
            </Bar>
            <Bar dataKey="added" stackId="cum" name={t("เพิ่มเดือนนี้","Added this month")} radius={[5,5,0,0]}>
              {chartData.map((e,i)=>{ const sel = e.monthKey===month; return <Cell key={i} fill={T.amber} stroke={sel?"#0f172a":"none"} strokeWidth={sel?2.5:0} cursor="pointer"/>; })}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Month picker — คลิกสลับเดือน · เพิ่ม/ลบเดือนได้ทันที (ลบมีเตือนก่อน) */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16,flexWrap:"wrap"}}>
        <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:6,flex:1,minWidth:0}}>
          {/* Start (baseline) — read-only reference: what date the project began */}
          <div style={{flexShrink:0,textAlign:"left",padding:"10px 16px",borderRadius:12,border:`1.5px solid ${T.cardBorder}`,background:"#f8fafc",minWidth:140}}>
            <div style={{fontSize:15,fontWeight:750,color:T.textSecondary,marginBottom:3,letterSpacing:0.2}}>{t("เริ่มต้น","Start")}</div>
            <div style={{fontSize:15,fontWeight:650,color:T.textSecondary,fontVariantNumeric:"tabular-nums"}}>{fmtK(baseTotal)}</div>
            <div style={{fontSize:11,color:T.textMuted,marginTop:2}}>{project?.createdAt ? new Date(project.createdAt).toLocaleDateString(_LANG==="en"?"en-US":"th-TH",{day:"numeric",month:"short",year:"2-digit"}) : t("ราคาเดิม","baseline")}</div>
          </div>
          {sortedMonths.map(m=>{
            const active = m===month;
            const add = monthTotalLive(m);
            const exists = months.includes(m); // เดือนที่มีจริง (ไม่ใช่ default เปล่า) ถึงลบได้
            return (
              <div key={m} onClick={()=>goMonth(m)} ref={active?activeChipRef:null}
                style={{position:"relative",flexShrink:0,textAlign:"left",padding:"10px 36px 10px 16px",borderRadius:12,border:`1.5px solid ${active?T.blue:T.cardBorder}`,
                  background:active?T.blue:T.card,cursor:"pointer",minWidth:140,transition:"all 0.15s"}}>
                <div style={{fontSize:15,fontWeight:750,color:active?"#fff":T.textPrimary,marginBottom:3,letterSpacing:0.2}}>{monthShortLabel(m)}</div>
                <div style={{fontSize:15,fontWeight:650,color:active?"#dbeafe":T.textSecondary,fontVariantNumeric:"tabular-nums"}}>{fmtK(cumulativeLive(m))}</div>
                <div style={{fontSize:13,fontWeight:700,color:active?(add<0?"#fecaca":"#fff"):(add>0?T.amber:add<0?T.red:T.textMuted),marginTop:3}}>{add>0?"+":""}{fmtK(add)} {t("เดือนนี้","this mo.")}</div>
                {exists && (
                  <button onClick={(e)=>{e.stopPropagation(); handleDeleteMonth(m);}} title={t("ลบเดือนนี้ (มีเตือนก่อนลบ)","Delete this month (asks first)")}
                    style={{position:"absolute",top:4,right:4,width:28,height:28,borderRadius:8,border:"none",lineHeight:1,
                      background:active?"rgba(255,255,255,0.2)":T.redBg,color:active?"#fff":T.red,cursor:"pointer",fontSize:13,padding:0,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
                )}
              </div>
            );
          })}
          {/* เพิ่มเดือน — มีป้ายบอก + แสดงเดือนที่เลือกเป็นคำอ่าน (เดิมเห็นแค่ช่องว่าง "-------") */}
          <div style={{flexShrink:0,display:"flex",flexDirection:"column",justifyContent:"center",gap:4,padding:"8px 12px",borderRadius:12,border:`1.5px dashed ${T.blue}`,background:T.blueLight}}>
            <span style={{fontSize:12,color:T.blue,fontWeight:650,whiteSpace:"nowrap"}}><Ico name="calendar" size={14} /> {t("เพิ่มเดือนใหม่","Add a month")}{newMonth ? ` · ${monthShortLabel(newMonth)}` : ""}</span>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <input type="month" value={newMonth} onChange={e=>setNewMonth(e.target.value)} className="input-base" title={t("เลือกเดือนที่จะเพิ่ม","Pick the month to add")}
                style={{border:`1px solid ${T.cardBorder}`,background:"#fff",padding:"6px 6px",width:136,fontSize:13,minHeight:32}}/>
              <button className="btn-primary" disabled={!newMonth} style={{padding:"6px 12px",fontSize:12,whiteSpace:"nowrap",background:newMonth?T.blue:"#94a3b8",cursor:newMonth?"pointer":"not-allowed",minHeight:32}} onClick={handleAddMonth}>+ {t("เพิ่ม","Add")}</button>
            </div>
          </div>
        </div>
      </div>

      {/* Stats for selected month */}
      <div className="stat-grid" style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
        <StatCard label={t("ยอดยกมา (ก่อนเดือนนี้)","Brought forward (before this month)")} value={"฿"+fmt0(baselineForMonth)} thb={baselineForMonth} rate={usdRate} sub={`${t("สะสมถึง","up to")} ${prevMonthLabel}`} color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label={t("เพิ่มเดือนนี้","Added this month")} value={"฿"+fmt0(thisMonthAdd)} thb={thisMonthAdd} rate={usdRate} sub={new Date(month+"-01").toLocaleDateString(_LANG==="en"?"en-US":"th-TH",{year:"numeric",month:"long"})} color={T.amber} icon="➕" accent={T.amberBg}/>
        <StatCard label={t("รวมสะสมถึงเดือนนี้","Cumulative to this month")} value={"฿"+fmt0(cumulativeSoFar)} thb={cumulativeSoFar} rate={usdRate} sub={t(`เดิม (+${WASTE_LBL}) + เพิ่มสะสมถึงเดือนที่เลือก`,`Baseline (+${WASTE_LBL}) + additions up to selected month`)} color={T.green} icon="✅" accent={T.greenBg}/>
        <StatCard label={t("รวมทั้งหมด","Grand total")} value={"฿"+fmt0(grandTotal)} thb={grandTotal} rate={usdRate} sub={t(`เดิม (+${WASTE_LBL}) + ทุกเดือนที่มีข้อมูล (ล่าสุด)`,`Baseline (+${WASTE_LBL}) + all months (latest)`)} color={T.purple} icon="🧮" accent={T.purpleBg}/>
      </div>

      {/* Toolbar: search + group filter + actions */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <SearchInput value={search} onChange={setSearch} placeholder={t("ค้นหา Account Code / ชื่อ...","Search Account Code / name...")} width={220}/>
        <button onClick={()=>setHideEmpty(v=>!v)}
          title={t("ซ่อน/แสดงแถวที่ไม่มีค่า (รวมสะสม = 0)","Hide/show empty rows (total = 0)")}
          style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:8,fontSize:11,fontWeight:600,cursor:"pointer",
            border:`1.5px solid ${hideEmpty?T.blue:T.cardBorder}`,background:hideEmpty?T.blue:T.card,color:hideEmpty?"#fff":T.textSecondary,whiteSpace:"nowrap"}}>
          {hideEmpty ? `✓ ${t("เฉพาะที่มีค่า","With value only")}${hiddenEmptyCount?` (${t("ซ่อน","hidden")} ${hiddenEmptyCount})`:""}` : t("เฉพาะที่มีค่า","With value only")}
        </button>
        <GroupFilter selected={filter} onChange={setFilter}/>
        <div style={{flex:1}}/>
        <button className="btn-ghost" onClick={()=>setAddExtraOpen(v=>!v)} disabled={!editingUnlocked}
          style={!editingUnlocked?{opacity:0.4,cursor:"not-allowed"}:undefined}>+ {t("งานพิเศษ","Extra item")}</button>
        {!editingUnlocked && (
          <span style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:T.textMuted,background:"#f1f5f9",padding:"6px 12px",borderRadius:8,fontWeight:600}}>
            <Ico name="lock" size={14} /> {t("บันทึกแล้ว","Saved")}
          </span>
        )}
        {editingUnlocked ? (
          <>
            <button onClick={handleSave} className="btn-primary" style={{background:saved?T.green:T.blue,minWidth:170}}>
              {saved?t("✓ บันทึกแล้ว","✓ Saved"):t("บันทึกรายการเดือนนี้","Save this month")}
            </button>
            {canCancel && (
              <button onClick={handleCancel} className="btn-ghost" title={t("ยกเลิกการแก้ไข (Esc)","Cancel editing (Esc)")}
                style={{color:T.red,borderColor:T.red}}>✕ {t("ยกเลิก","Cancel")}</button>
            )}
          </>
        ) : (
          <button onClick={()=>setForceEdit(true)} className="btn-primary" style={{minWidth:170,display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6}}>
            <Ico name="edit" /> {t("แก้ไขเดือนนี้","Edit this month")}
          </button>
        )}
      </div>

      {editingUnlocked && (
        <div style={{display:"flex",alignItems:"center",gap:8,margin:"-6px 2px 14px",fontSize:12,color:T.textMuted,flexWrap:"wrap"}}>
          <span style={{background:T.greenBg,color:T.green,fontWeight:700,fontSize:11,padding:"2px 8px",borderRadius:6,whiteSpace:"nowrap"}}>Excel</span>
          <span>{t("ลากคลุมเลือก · Shift+คลิก ขยายช่วง · Ctrl/Cmd+คลิก เลือก/ยกเลิกทีละช่อง · Ctrl/Cmd+C คัดลอก · Delete ล้าง · วางจาก Excel เติมทั้งบล็อก","Drag to select · Shift+click to extend · Ctrl/Cmd+click to toggle a cell · Ctrl/Cmd+C to copy · Delete to clear · paste from Excel to fill a block")}</span>
          <div style={{flex:1,minWidth:8}}/>
          <button onClick={selectAll} className="btn-ghost" style={{padding:"3px 10px",fontSize:11,whiteSpace:"nowrap"}}>{t("เลือกทั้งหมด","Select all")}</button>
          {selCount>0 && (
            <>
              <span style={{background:T.blueLight,color:T.blue,fontWeight:700,fontSize:11,padding:"3px 9px",borderRadius:6,whiteSpace:"nowrap"}}>{t("เลือก","Selected")} {selCount}</span>
              <button onClick={copySelection} className="btn-ghost" style={{padding:"3px 10px",fontSize:11,whiteSpace:"nowrap"}}><Ico name="copy" size={13} /> {t("คัดลอก","Copy")}</button>
              <button onClick={clearSelection} className="btn-ghost" style={{padding:"3px 10px",fontSize:11,whiteSpace:"nowrap",color:T.red,borderColor:T.red}}><Ico name="trash" size={13} /> {t("ล้างที่เลือก","Clear")}</button>
              <button onClick={deselectAll} className="btn-ghost" style={{padding:"3px 8px",fontSize:11,whiteSpace:"nowrap"}}>✕</button>
            </>
          )}
        </div>
      )}

      {/* ไฮไลต์เอง: ลากเลือกเซลล์แล้วจัดรูปแบบ (หนา/เอียง/ขีดเส้นใต้/ขนาด/สีพื้น/สีตัวอักษร) — ใช้ได้แม้ไม่ได้อยู่โหมดแก้ไข */}
      {selCount>0 && (
        <div style={{display:"flex",alignItems:"center",gap:6,margin:"-8px 2px 14px",fontSize:11,color:T.textMuted,flexWrap:"wrap"}}>
          {!editingUnlocked && (
            <span style={{background:T.blueLight,color:T.blue,fontWeight:700,fontSize:11,padding:"3px 9px",borderRadius:6,whiteSpace:"nowrap"}}>{t("เลือก","Selected")} {selCount}</span>
          )}
          {(() => {
            const on = { padding:0,width:26,height:24,borderRadius:6,border:`1px solid ${T.cardBorder}`,background:"#fff",cursor:"pointer",fontSize:13,lineHeight:1,color:T.textPrimary };
            return (<>
              <button onClick={()=>toggleCellFmt("b")} title={t("ตัวหนา","Bold")} style={{...on,fontWeight:900}}>B</button>
              <button onClick={()=>toggleCellFmt("i")} title={t("ตัวเอียง","Italic")} style={{...on,fontStyle:"italic",fontFamily:"Georgia,serif"}}>I</button>
              <button onClick={()=>toggleCellFmt("u")} title={t("ขีดเส้นใต้","Underline")} style={{...on,textDecoration:"underline"}}>U</button>
              <button onClick={()=>bumpFontSize(-1)} title={t("ลดขนาดตัวอักษร","Smaller")} style={{...on,fontSize:11}}>A−</button>
              <button onClick={()=>bumpFontSize(1)} title={t("เพิ่มขนาดตัวอักษร","Larger")} style={{...on,fontSize:15,fontWeight:700}}>A+</button>
            </>);
          })()}
          <span style={{width:1,height:18,background:T.cardBorder,margin:"0 2px"}}/>
          <span style={{fontWeight:700,color:T.textSecondary,whiteSpace:"nowrap"}}>{t("ไฮไลต์พื้น","Fill")}:</span>
          {["#FEF3C7","#D1FAE5","#FEE2E2","#DBEAFE","#E5E7EB"].map(bg=>(
            <button key={bg} onClick={()=>applyCellFmt({bg})} title={t("ใส่สีพื้นให้ช่องที่เลือก","Fill selected cells")}
              style={{width:22,height:22,borderRadius:6,border:`1px solid ${T.cardBorder}`,background:bg,cursor:"pointer",padding:0}}/>
          ))}
          <button onClick={()=>applyCellFmt({bg:null})} className="btn-ghost" style={{padding:"3px 8px",fontSize:11,whiteSpace:"nowrap"}} title={t("ล้างสีพื้น","Remove fill")}>{t("ล้างพื้น","No fill")}</button>
          <span style={{width:1,height:18,background:T.cardBorder,margin:"0 2px"}}/>
          <span style={{fontWeight:700,color:T.textSecondary,whiteSpace:"nowrap"}}>{t("สีตัวอักษร","Text")}:</span>
          {[["#DC2626","แดง","red"],["#059669","เขียว","green"],["#2563EB","น้ำเงิน","blue"],["#0F172A","ดำ","black"]].map(([fg,nm,en])=>(
            <button key={fg} onClick={()=>applyCellFmt({fg})} title={t(`สีตัวอักษร ${nm}`,`Text color ${en}`)}
              style={{width:22,height:22,borderRadius:6,border:`1px solid ${T.cardBorder}`,background:"#fff",color:fg,cursor:"pointer",padding:0,fontWeight:800,fontSize:13,lineHeight:1}}>A</button>
          ))}
          <button onClick={()=>applyCellFmt({fg:null})} className="btn-ghost" style={{padding:"3px 8px",fontSize:11,whiteSpace:"nowrap"}} title={t("ล้างสีตัวอักษร","Reset text color")}>↺ {t("สีปกติ","Default")}</button>
          {!editingUnlocked && (
            <>
              <span style={{width:1,height:18,background:T.cardBorder,margin:"0 2px"}}/>
              <button onClick={selectAll} className="btn-ghost" style={{padding:"3px 10px",fontSize:11,whiteSpace:"nowrap"}}>{t("เลือกทั้งหมด","Select all")}</button>
              <button onClick={copySelection} className="btn-ghost" style={{padding:"3px 10px",fontSize:11,whiteSpace:"nowrap"}}><Ico name="copy" size={13} /> {t("คัดลอก","Copy")}</button>
              <button onClick={deselectAll} className="btn-ghost" style={{padding:"3px 8px",fontSize:11,whiteSpace:"nowrap"}}>✕</button>
            </>
          )}
        </div>
      )}

      {addExtraOpen && (
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"1fr 2fr 1fr auto",gap:10,alignItems:"end"}}>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>{t("Acc. Code (เว้นว่างให้สร้างอัตโนมัติ)","Acc. Code (leave blank = auto)")}</span>
            <input className="input-base" value={extraDraft.code} onChange={e=>setExtraDraft(d=>({...d,code:e.target.value}))} placeholder={t("เช่น 511099","e.g. 511099")} />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>{t("ชื่อรายการงานเพิ่ม","Extra item name")}</span>
            <input className="input-base" value={extraDraft.name} onChange={e=>setExtraDraft(d=>({...d,name:e.target.value}))} placeholder={t("เช่น งานเพิ่มกระจกโค้งพิเศษ","e.g. Extra curved glass work")} onKeyDown={e=>e.key==="Enter"&&handleCreateExtra()} />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>Group</span>
            <select className="input-base" value={extraDraft.group} onChange={e=>setExtraDraft(d=>({...d,group:e.target.value}))}>
              {GROUPS.map(g => <option key={g} value={g}>{g}</option>)}
            </select>
          </label>
          <button className="btn-primary" onClick={handleCreateExtra}>+ {t("สร้างรายการ","Create item")}</button>
        </div>
      )}

      {/* Main table: เดิม + เพิ่มเดือนนี้ = รวมสะสม */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
        <div className="mscroll">
        {/* กว้างเต็มกล่องเสมอ (เดิม max-content → คอลัมน์น้อยแล้วเหลือที่ว่างขาวด้านข้าง แถวสีสลับจบไม่เท่ากัน) */}
        <table ref={mTableRef} style={{minWidth:"100%", width: isMultiCol ? "max-content" : "100%", borderCollapse:"collapse", fontSize:13}}>
          <thead>
            {isMultiCol ? (
              <>
                <tr ref={mHead1Ref} style={{background:"#f8fafc"}}>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"left",color:sortKey==="code"?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap", ...qsFrz(0,"#f8fafc",7)}}>
                    <span onClick={()=>handleSort("code")} style={{cursor:"pointer",userSelect:"none"}}>Acc. Code{sortKey==="code"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"left",color:sortKey==="group"?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap", ...qsFrz(1,"#f8fafc",7)}}>
                    <span onClick={()=>handleSort("group")} style={{cursor:"pointer",userSelect:"none"}}>Group{sortKey==="group"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"left",color:sortKey==="name"?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap", ...qsFrz(2,"#f8fafc",7)}}>
                    <span onClick={()=>handleSort("name")} style={{cursor:"pointer",userSelect:"none"}}>Account Name{sortKey==="name"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"right",color:sortKey==="before"?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap", ...qsFrz(3,"#f8fafc",7)}}>
                    <span onClick={()=>handleSort("before")} style={{cursor:"pointer",userSelect:"none"}}>{t("ยอดก่อนหน้า","Previous")}{sortKey==="before"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"center",width:20,color:T.textMuted,borderBottom:`1px solid ${T.cardBorder}`}}>+</th>
                  <th colSpan={columns.length+1} style={{padding:"9px 16px",textAlign:"center",color:T.textMuted,fontWeight:650,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>
                    + {t("เพิ่มเดือนนี้","Add this month")} · {monthShortLabel(month)}
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"center",width:20,color:T.textMuted,borderBottom:`1px solid ${T.cardBorder}`}}>=</th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"right",color:sortKey==="cum"?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                    <span onClick={()=>handleSort("cum")} style={{cursor:"pointer",userSelect:"none"}}>{t("รวมสะสม","Total")}{sortKey==="cum"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{width:20,borderBottom:`1px solid ${T.cardBorder}`}}></th>
                </tr>
                <tr style={{background:"#f8fafc"}}>
                  {columns.map(c=>(
                    <th key={c.id} style={{padding:"6px 18px",textAlign:"right",color:T.textMuted,fontWeight:600,fontSize:12,borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:5}}>
                        <span>{c.name}</span>
                        {editingUnlocked && <button onClick={()=>handleRemoveColumn(c.id)} title={t("ลบรายการนี้ (เฉพาะเดือนนี้)","Delete this item (this month only)")} style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:12,padding:0}}>✕</button>}
                      </div>
                    </th>
                  ))}
                  <th style={{padding:"6px 10px",textAlign:"right",borderBottom:`1px solid ${T.cardBorder}`}}>
                    {editingUnlocked && (addColOpen ? (
                      <div style={{display:"flex",gap:4,alignItems:"center",justifyContent:"flex-end"}}>
                        <input autoFocus value={newColName} onChange={e=>setNewColName(e.target.value)} placeholder={t("ชื่อ เช่น CC#17","Name e.g. CC#17")}
                          className="input-base" style={{width:88,fontSize:12,padding:"4px 6px"}}
                          onKeyDown={e=>e.key==="Enter"&&handleAddColumn()} />
                        <button onClick={handleAddColumn} className="btn-primary" style={{padding:"4px 9px",fontSize:12}}>+</button>
                        <button onClick={()=>setAddColOpen(false)} className="btn-ghost" style={{padding:"4px 7px",fontSize:12}}>×</button>
                      </div>
                    ) : (
                      <button onClick={()=>setAddColOpen(true)} className="btn-ghost" style={{padding:"4px 10px",fontSize:12,whiteSpace:"nowrap"}}>+ {t("เพิ่มรายการ","Add item")}</button>
                    ))}
                  </th>
                </tr>
              </>
            ) : (
              <tr style={{background:"#f8fafc"}}>
                {[
                  {label:"Acc. Code", key:"code", align:"left"},
                  {label:"Group", key:"group", align:"left"},
                  {label:"Account Name", key:"name", align:"left"},
                  {label:t("ยอดก่อนหน้า","Previous"), key:"before", align:"right"},
                  {label:"+", key:null, align:"center", width:20},
                  {label:t("+ เพิ่มเดือนนี้","+ Add this month"), key:"add", align:"right"},
                  {label:"=", key:null, align:"center", width:20},
                  {label:t("รวมสะสม","Total"), key:"cum", align:"right"},
                  {label:"", key:null, width:20},
                ].map(({label,key,align,width},idx)=>(
                  <th key={idx}
                    style={{padding:"11px 16px",textAlign:align||"left",color:key&&sortKey===key?T.blue:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:label.length>2?0.8:0,textTransform:label.length>2?"uppercase":"none",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap",...(width?{width}:{}),...(idx<4?qsFrz(idx,"#f8fafc",7):{})}}>
                    {key==="add" ? (
                      <div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:8}}>
                        <span onClick={()=>handleSort("add")} style={{cursor:"pointer",userSelect:"none"}}>{label}{sortKey==="add"?(sortDir===1?" ▲":" ▼"):""}</span>
                        {editingUnlocked && (addColOpen ? (
                          <div style={{display:"flex",gap:4,alignItems:"center"}} onClick={e=>e.stopPropagation()}>
                            <input autoFocus value={newColName} onChange={e=>setNewColName(e.target.value)} placeholder={t("ชื่อ เช่น CC#17","Name e.g. CC#17")}
                              className="input-base" style={{width:88,fontSize:12,padding:"4px 6px",textTransform:"none"}}
                              onKeyDown={e=>e.key==="Enter"&&handleAddColumn()} />
                            <button onClick={handleAddColumn} className="btn-primary" style={{padding:"4px 9px",fontSize:12}}>+</button>
                            <button onClick={()=>setAddColOpen(false)} className="btn-ghost" style={{padding:"4px 7px",fontSize:12}}>×</button>
                          </div>
                        ) : (
                          <button onClick={(e)=>{e.stopPropagation();setAddColOpen(true);}} className="btn-ghost" style={{padding:"3px 8px",fontSize:12,whiteSpace:"nowrap",textTransform:"none"}}>+ {t("เพิ่มรายการ","Add item")}</button>
                        ))}
                      </div>
                    ) : (
                      <span onClick={()=>key&&handleSort(key)} style={{cursor:key?"pointer":"default",userSelect:"none"}}>{label}{key && sortKey===key ? (sortDir===1?" ▲":" ▼") : ""}</span>
                    )}
                  </th>
                ))}
              </tr>
            )}
          </thead>
          <tbody>
            {displayRows.map((r,i) => {
              const kids = kidsAsOf(r.code, month);
              const hasKids = kids.length > 0;
              const isCollapsed = hasKids && rowCollapsed[r.code];
              const cumBefore = cumBeforeOf(r);
              const thisVal = rowMonthValue(r.code, month, draftAdd);
              const cum = cumBefore + thisVal;
              const rowBg = i % 2 === 0 ? T.card : "#fafbfd";
              return (
                <Fragment key={r.code}>
                  <tr onClick={()=>hasKids && setRowCollapsed(c=>({...c,[r.code]:!c[r.code]}))}
                      style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(hasKids&&!isCollapsed)||subFor===r.code?"none":"1px solid #f1f5f9",cursor:hasKids?"pointer":"default"}}>
                    <td style={{padding:"10px 16px",color:T.blue,fontVariantNumeric:"tabular-nums",fontSize:13,fontWeight:500, ...qsFrz(0,rowBg)}}>
                      {hasKids && (
                        <span title={isCollapsed?t("ขยายรายการย่อย","Expand sub-items"):t("ย่อรายการย่อย","Collapse sub-items")}
                          style={{color:T.textMuted,fontSize:12,marginRight:6,verticalAlign:"middle",display:"inline-block"}}>
                          {isCollapsed?"▸":"▾"}
                        </span>
                      )}
                      {r.isExtra ? codeText(r.code) : r.code}
                    </td>
                    <td style={{padding:"10px 16px", ...qsFrz(1,rowBg)}}>
                      {(i===0 || displayRows[i-1]?.group!==r.group) && <span style={{background:T.blueLight,color:T.blue,fontSize:12,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{r.group}</span>}
                    </td>
                    <td style={{padding:"10px 16px",color:T.textPrimary, ...qsFrz(2,rowBg)}}>
                      {r.name}
                      {r.isExtra && <span style={{marginLeft:7,fontSize:12,background:T.amberBg,color:T.amber,padding:"1px 8px",borderRadius:6,fontWeight:600}}>{t("งานเพิ่ม","Extra")}</span>}
                      {hasKids && <span style={{marginLeft:7,fontSize:12,background:T.greenBg,color:T.green,padding:"1px 8px",borderRadius:6,fontWeight:600}}>{kids.length} {t("รายการย่อย","sub-items")}</span>}
                      {editingUnlocked && (
                        <button onClick={(e)=>{e.stopPropagation(); setSubFor(subFor===r.code?null:r.code); setSubName(""); setRowCollapsed(c=>({...c,[r.code]:false}));}} title={t("เพิ่มรายการย่อยใต้ Acc. Code นี้","Add a sub-item under this Acc. Code")}
                          style={{marginLeft:9,background:"none",border:`1px dashed ${T.cardBorder}`,borderRadius:6,color:T.textMuted,cursor:"pointer",fontSize:12,padding:"1px 7px"}}>
                          + {t("รายการย่อย","Sub-item")}
                        </button>
                      )}
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right",color:cumBefore!==0?T.textPrimary:T.textMuted,fontVariantNumeric:"tabular-nums", ...qsFrz(3,rowBg)}} title={t(`ราคาเดิม + เผื่อเศษ ${WASTE_LBL} + ยอดเพิ่มของทุกเดือนก่อนหน้ารวมกัน`,`Baseline + ${WASTE_LBL} wastage + additions of all previous months`)}>{fmtZ(cumBefore)}{usdLine(cumBefore, usdRate)}</td>
                    <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                    {isMultiCol ? (
                      hasKids ? (
                        <td colSpan={columns.length+1} style={{padding:"8px 16px",textAlign:"right"}}>
                          <div style={{width:"100%",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",background:T.amberBg,borderRadius:8,color:T.amber,fontWeight:650,fontSize:13}}>
                            {fmtZ(thisVal)}
                            {usdLine(thisVal, usdRate)}
                          </div>
                        </td>
                      ) : columns.map((c,ci)=>{
                        const ck = `${r.code}:${c.id}`;
                        const cv = parseFloat(draftAdd[ck])||0;
                        const on = inSel(i,ci);
                        const hl = cellFmtStyle(ck);
                        return (
                          <td key={c.id}
                            onMouseDown={e=>onCellDown(i,ci,e)}
                            onMouseEnter={()=>onCellEnter(i,ci)}
                            style={{padding:"8px 10px",textAlign:"right",...(on?{background:"#dbeafe",boxShadow:`inset 0 0 0 1.5px ${T.blue}`}:{})}}>
                            {editingUnlocked ? (
                              <MoneyInput allowNegative value={draftAdd[ck]??""} onChange={v=>setDraftAdd(d=>({...d,[ck]:v}))}
                                onPaste={raw=>handleGridPaste(i,ci,raw)}
                                style={{width:104,fontSize:13,background:cv!==0?T.amberBg:(on?"transparent":T.bg),...hl}}/>
                            ) : (
                              <div style={{width:104,marginLeft:"auto",padding:"7px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13, ...(cv!==0?{background:T.amberBg,color:T.amber,fontWeight:700,borderRadius:8}:{color:T.textMuted}),...hl}}>{fmtZ(cv)}{usdLine(cv, usdRate)}</div>
                            )}
                          </td>
                        );
                      }).concat(<td key="__addcol" aria-hidden="true" />)   /* ช่องใต้ปุ่ม "+ เพิ่มรายการ" ในหัวตาราง — ให้จำนวนช่องเท่ากับหัว (เดิมขาด 1 ช่อง คอลัมน์หลังจากนี้เลยเหลื่อม) */
                    ) : (
                      <td
                        onMouseDown={!hasKids?e=>onCellDown(i,0,e):undefined}
                        onMouseEnter={!hasKids?()=>onCellEnter(i,0):undefined}
                        style={{padding:"8px 16px",textAlign:"right",...(inSel(i,0)&&!hasKids?{background:"#dbeafe",boxShadow:`inset 0 0 0 1.5px ${T.blue}`}:{})}}>
                        {hasKids ? (
                          <div style={{width:130,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",background:T.amberBg,borderRadius:8,color:T.amber,fontWeight:650,fontSize:13}}>
                            {fmtZ(thisVal)}
                            {usdLine(thisVal, usdRate)}
                          </div>
                        ) : editingUnlocked ? (
                          <MoneyInput allowNegative value={draftAdd[r.code]??""} onChange={v=>setDraftAdd(d=>({...d,[r.code]:v}))}
                            onPaste={raw=>handleGridPaste(i,0,raw)}
                            style={{width:130,background:thisVal!==0?T.amberBg:(inSel(i,0)?"transparent":T.bg),...cellFmtStyle(r.code)}}/>
                        ) : (
                          <div style={{width:130,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13, ...(thisVal!==0?{background:T.amberBg,color:T.amber,fontWeight:700,borderRadius:8}:{color:T.textMuted}),...cellFmtStyle(r.code)}}>{fmtZ(thisVal)}{usdLine(thisVal, usdRate)}</div>
                        )}
                      </td>
                    )}
                    <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                    <td style={{padding:"8px 16px",textAlign:"right",color:cum!==0?T.textPrimary:T.textMuted,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{fmtZ(cum)}{usdLine(cum, usdRate)}</td>
                    <td style={{padding:"8px 16px",textAlign:"center"}}>
                      {r.isExtra && editingUnlocked && (
                        <button onClick={(e)=>{e.stopPropagation(); handleDeleteExtra(r.code);}} title={t("ลบรายการงานเพิ่ม","Delete extra item")}
                          style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
                      )}
                    </td>
                  </tr>

                  {/* Sub-items — this month's value rolls up into the parent row above.
                      Only sub-items added on/before the currently-viewed month appear here
                      (kidsAsOf already filtered them), so a sub-item created in ก.ย. simply
                      doesn't exist in ส.ค. or earlier — no ghost "0.00" row. */}
                  {!isCollapsed && kids.map((k,ki) => {
                    const kBaseVal = withWaste(tenderCosts[k.code]);   // ราคาเดิม + เผื่อเศษ
                    const kCumBefore = months.filter(m=>m<month).reduce((s,m)=>s+(parseFloat(additions[m]?.[k.code])||0),0) + kBaseVal;
                    const kThisVal = parseFloat(draftAdd[k.code]) || 0;
                    const kCum = kCumBefore + kThisVal;
                    const isNewThisMonth = k.addedInMonth === month;
                    const subBg = isNewThisMonth ? T.greenBg : (i%2===0?T.card:"#fafbfd");
                    return (
                      <tr key={k.code} style={{background:subBg,borderLeft:`3px solid ${isNewThisMonth?T.green:"#e2e8f0"}`,borderBottom:(ki===kids.length-1 && subFor!==r.code)?"1px solid #f1f5f9":"none",transition:"background 0.2s"}}>
                        <td style={{padding:"7px 16px 7px 27px",color:T.green,fontSize:13, ...qsFrz(0,subBg)}}>↳</td>
                        <td style={qsFrz(1,subBg)}/>
                        <td style={{padding:"7px 16px",color:T.green,fontSize:13,fontStyle:"italic", ...qsFrz(2,subBg)}}>
                          {k.name}
                          {k.addedInMonth && (
                            isNewThisMonth ? (
                              <span title={t("รายการนี้เพิ่งเพิ่มเข้ามาในเดือนนี้","Added this month")} style={{marginLeft:8,fontSize:12,background:T.green,color:"#fff",padding:"2px 8px",borderRadius:6,fontWeight:650,fontStyle:"normal",letterSpacing:0.2}}>
                                {t("ใหม่เดือนนี้","New this month")}
                              </span>
                            ) : (
                              <span title={t("เพิ่มเข้ามาระหว่างทาง ไม่ได้มีมาตั้งแต่ต้น — เดือนก่อนหน้านั้นจะไม่แสดงรายการนี้","Added later, not from the start — earlier months don't show it")} style={{marginLeft:8,fontSize:12,background:T.amberBg,color:T.amber,padding:"2px 8px",borderRadius:6,fontWeight:600,fontStyle:"normal"}}>
                                {t("เพิ่มเมื่อ","Added")} {monthShortLabel(k.addedInMonth)}
                              </span>
                            )
                          )}
                        </td>
                        <td style={{padding:"7px 16px",textAlign:"right",color:kCumBefore!==0?T.textPrimary:T.textMuted,fontVariantNumeric:"tabular-nums",fontSize:13, ...qsFrz(3,subBg)}}>{fmtZ(kCumBefore)}{usdLine(kCumBefore, usdRate)}</td>
                        <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                        <td colSpan={isMultiCol ? columns.length+1 : 1} style={{padding:"7px 16px",textAlign:"right"}}>
                          {editingUnlocked ? (
                            <MoneyInput allowNegative value={draftAdd[k.code]??""} onChange={v=>setDraftAdd(d=>({...d,[k.code]:v}))}
                              style={{width:130,fontSize:13,background:kThisVal!==0?T.greenBg:T.bg}}/>
                          ) : (
                            <div style={{width:130,marginLeft:"auto",padding:"7px 8px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontSize:13, ...(kThisVal!==0?{background:T.greenBg,color:T.green,fontWeight:700,borderRadius:8}:{color:T.textMuted})}}>{fmtZ(kThisVal)}{usdLine(kThisVal, usdRate)}</div>
                          )}
                        </td>
                        <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                        <td style={{padding:"7px 16px",textAlign:"right",color:kCum!==0?T.textPrimary:T.textMuted,fontVariantNumeric:"tabular-nums",fontWeight:650,fontSize:13}}>{fmtZ(kCum)}{usdLine(kCum, usdRate)}</td>
                        <td style={{padding:"7px 16px",textAlign:"center"}}>
                          {editingUnlocked && (
                            <button onClick={()=>handleDeleteExtra(k.code)} title={t("ลบรายการย่อยนี้","Delete this sub-item")}
                              style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:14,opacity:0.7}}
                              onMouseEnter={e=>e.currentTarget.style.opacity=1} onMouseLeave={e=>e.currentTarget.style.opacity=0.7}>✕</button>
                          )}
                        </td>
                      </tr>
                    );
                  })}

                  {/* Inline "add sub-item" form for this row */}
                  {subFor===r.code && (
                    <tr style={{background:T.greenBg,borderBottom:"1px solid #f1f5f9"}}>
                      <td style={qsFrz(0,T.greenBg)}/><td style={qsFrz(1,T.greenBg)}/>
                      <td style={{padding:"7px 16px", ...qsFrz(2,T.greenBg)}}>
                        <input className="input-base" value={subName} onChange={e=>setSubName(e.target.value)}
                          placeholder={t("ชื่อรายการย่อย เช่น Silicone Structure","Sub-item name e.g. Silicone Structure")} style={{width:"100%",fontSize:13}}
                          onKeyDown={e=>e.key==="Enter"&&handleAddSub(r.code)} autoFocus />
                      </td>
                      <td colSpan={(isMultiCol ? 9+columns.length : 9)-3} style={{padding:"7px 16px",display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button className="btn-primary" style={{padding:"5px 12px",fontSize:13}} onClick={()=>handleAddSub(r.code)}>+ {t("เพิ่ม","Add")}</button>
                        <button className="btn-ghost" style={{padding:"5px 12px",fontSize:13}} onClick={()=>setSubFor(null)}>{t("ยกเลิก","Cancel")}</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={isMultiCol ? 9+columns.length : 9} style={{padding:"28px 16px",textAlign:"center",color:T.textMuted,fontSize:13}}>{t("ไม่พบรายการที่ตรงกับการค้นหา","No items match your search")}</td></tr>
            )}
          </tbody>
          <tfoot>
            <tr style={{background:"#eef2f7",borderTop:`2px solid ${T.textMuted}`}}>
              <td colSpan={3} style={{padding:"14px 16px",color:T.textSecondary,fontSize:13,fontWeight:700, ...qsFrzSpan3("#eef2f7"), bottom:0, zIndex:6}}>{t("รวม","Total")} {filtered.length} {t("รายการ","items")}</td>
              <td style={{padding:"14px 16px",textAlign:"right",color:T.textPrimary,fontVariantNumeric:"tabular-nums",fontWeight:700,fontSize:14, ...qsFrz(3,"#eef2f7"), bottom:0, zIndex:6}}>
                {fmtZ(filtered.reduce((s,r)=>s+cumBeforeOf(r),0))}
                {usdLine(filtered.reduce((s,r)=>s+cumBeforeOf(r),0), usdRate)}
              </td>
              <td style={QSF_FOOT}/>
              {isMultiCol
                ? columns.map(c => { const ct = filtered.reduce((s,r)=> s + (parseFloat(draftAdd[`${r.code}:${c.id}`])||0), 0); return (
                    <td key={c.id} style={{padding:"10px 14px",textAlign:"right",whiteSpace:"nowrap", ...QSF_FOOT}}>
                      <span style={{display:"inline-block",fontVariantNumeric:"tabular-nums",fontSize:14, ...(ct!==0?{background:T.amber,color:"#fff",fontWeight:800,padding:"5px 10px",borderRadius:8}:{color:T.textMuted,fontWeight:600})}}>{fmt(ct)}</span>
                      {usdLine(ct, usdRate)}
                    </td>
                  ); }).concat(<td key="__addcol" style={QSF_FOOT} />)
                : (() => { const ct = filtered.reduce((s,r)=>s+rowMonthValue(r.code, month, draftAdd),0); return (
                    <td style={{padding:"10px 16px",textAlign:"right",whiteSpace:"nowrap", ...QSF_FOOT}}>
                      <span style={{display:"inline-block",fontVariantNumeric:"tabular-nums",fontSize:14, ...(ct!==0?{background:T.amber,color:"#fff",fontWeight:800,padding:"5px 10px",borderRadius:8}:{color:T.textMuted,fontWeight:600})}}>{fmt(ct)}</span>
                      {usdLine(ct, usdRate)}
                    </td>
                  ); })()
              }
              <td style={QSF_FOOT}/>
              <td style={{padding:"10px 16px",textAlign:"right",whiteSpace:"nowrap", ...QSF_FOOT}}>
                {(() => { const g = filtered.reduce((s,r)=>s+cumBeforeOf(r)+rowMonthValue(r.code, month, draftAdd),0); return (<>
                  <span style={{display:"inline-block",background:T.green,color:"#fff",fontWeight:800,padding:"5px 11px",borderRadius:8,fontVariantNumeric:"tabular-nums",fontSize:15}}>{fmt(g)}</span>
                  {usdLine(g, usdRate)}
                </>); })()}
              </td>
              <td style={QSF_FOOT}/>
            </tr>
          </tfoot>
        </table>
        </div>
      </div>
    </div>
  );
}

export { QSView, QSBaselineTab, MonthAxisTick, QSF_W, QSF_L, QSF_SPAN3, qsFrz, qsFrzSpan3, QSF_FOOT, QSMonthlyTab };
