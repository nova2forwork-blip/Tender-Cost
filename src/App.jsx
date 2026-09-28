// Tender Cost — ตัวหลักของแอป: โหลด/บันทึกข้อมูล, สลับหน้า, สิทธิ์ตามแผนก
// ไฟล์ย่อยทั้งหมดต้องอยู่โฟลเดอร์เดียวกับไฟล์นี้: core.jsx excel.js ui.jsx home.jsx admin.jsx qs.jsx procurement.jsx accounting.jsx
import { useCallback, useEffect, useRef, useState } from "react";
import { recordSeen, sdOrThrow, sgOrThrow, ssMerge, ssOrThrow, supabase } from "./supabase.js";
import { clearSession, getSession, setSession } from "./auth.js";
import { BOTTOM, FAB_GAP, FAB_SIZE, GLOBAL_CSS, T, appBuild, buildFromHtml, UnsavedGuard, applyAccountList, fmt, inThai, leaveIfDirty, rollupAdditions, rollupTenders, setExtraRegistry, t, uiAlert, uiPrompt, useLang } from "./core.jsx";
import { buildTSV, exportAccountingExcel, exportProcurementExcel, exportQSExcel } from "./excel.js";
import { CalcFab, CalcStore, CalculatorPopup, DialogHost, ErrorBoundary, Ico, Loader, ScrollTopFab, TableTopButton, effRate } from "./ui.jsx";
import { HomeScreen, LoginScreen, RoleSelect } from "./home.jsx";
import { AdminPanel } from "./admin.jsx";
import { QSView } from "./qs.jsx";
import { ProcurementView } from "./procurement.jsx";
import { AccountingView } from "./accounting.jsx";

export default function App() {
  useLang();                                            // re-render ทั้งแอปเมื่อสลับภาษา
  const [session,  setSessionState] = useState(null);   // โหลดแบบ async ด้านล่าง
  const [authReady, setAuthReady]   = useState(false);  // true เมื่อเช็ค session เสร็จ
  const [screen,   setScreen]   = useState("home");
  // ── เวอร์ชันใหม่บนเซิร์ฟเวอร์: เช็กตอนเปิดแอป / ทุก 5 นาที / ตอนกลับมาที่แท็บ ─────────
  // ถ้าไฟล์ JS บนเซิร์ฟเวอร์ (index-XXXX.js) ไม่ใช่ตัวที่เปิดอยู่ = มีการ deploy ใหม่ → แสดงแถบให้รีเฟรช
  // ทุกคนจะได้ใช้เวอร์ชันเดียวกัน (เดิมแท็บที่เปิดค้างไว้ใช้โค้ดเก่าไปเรื่อย ๆ จนกว่าจะรีเฟรชเอง)
  const [newBuild, setNewBuild] = useState("");
  const [updateLater, setUpdateLater] = useState(false);
  useEffect(() => {
    const cur = appBuild().build;
    if (!cur || cur === "dev") return;
    let stop = false;
    const fetchIndex = () => window.__fetchIndex ? window.__fetchIndex() : fetch(`/?_v=${Date.now()}`, { cache: "no-store" }).then(r => r.ok ? r.text() : "");   // __fetchIndex = ชุดทดสอบ
    const check = async () => { try { const b = buildFromHtml(await fetchIndex()); if (!stop && b && b !== cur) { setNewBuild(b); setUpdateLater(false); } } catch { /* ออฟไลน์ — ลองใหม่รอบหน้า */ } };
    const onVis = () => { if (document.visibilityState === "visible") check(); };
    check();
    const iv = setInterval(check, 5 * 60 * 1000);
    document.addEventListener("visibilitychange", onVis);
    return () => { stop = true; clearInterval(iv); document.removeEventListener("visibilitychange", onVis); };
  }, []);
  const [projects, setProjects] = useState([]);
  const [accountsRev, setAccountsRev] = useState(0); // bump เมื่อรายการบัญชี (ACCOUNTS) ถูกแก้ → re-render ทั้งแอป
  const [activeId, setActiveId] = useState(null);
  const activeIdRef = useRef(null); activeIdRef.current = activeId;   // ใช้กันผลโหลดของโครงการเก่ามาทับ
  const [projReadyId, setProjReadyId] = useState(null);  // ข้อมูลใน state ตอนนี้เป็นของโครงการไหน (null = ยังไม่มี)
  const [projLoadErr, setProjLoadErr] = useState("");    // โหลดข้อมูลโครงการไม่สำเร็จ
  const [listLoadErr, setListLoadErr] = useState(false);  // โหลดรายการโครงการไม่สำเร็จ (อย่าโชว์ "ยังไม่มีโครงการ" หลอก ๆ)
  const [role,     setRole]     = useState(null);
  const [tenderCosts, setTCosts]= useState({});
  const [additions,   setAdditions]  = useState({});
  const [extraItems,  setExtraItems] = useState([]);
  const [hiddenAccounts, setHiddenAccounts] = useState([]); // codes of fixed Acc. Codes QS has removed for this project
  const [incomingPlan, setIncomingPlan] = useState([]); // แผนของเข้าทั้งโปรเจค (จัดซื้อวางแผนก่อนออก PO): [{ id, date, items:[{id,code,amount}] }]
  const [poEntries,   setPO]    = useState([]);
  setExtraRegistry(extraItems);   // ให้ตัวช่วยระดับโมดูลหา "ชื่อ" ของรหัสงานเพิ่มของโครงการนี้ได้
  const [loaded,   setLoaded]   = useState(false);
  const [newProjModal, setNewProjModal] = useState(false);
  const [syncedAt,    setSyncedAt]    = useState(null);
  const [syncing,     setSyncing]     = useState(false);
  const [syncError,   setSyncError]   = useState("");   // ข้อความเตือนเมื่อบันทึก/โหลดพลาด
  const [exportMsg,   setExportMsg]   = useState("");   // สถานะตอนกด Export (กำลังสร้าง/เสร็จ/พลาด)
  const runExport = async (fn) => {
    setExportMsg(t("⏳ กำลังสร้างไฟล์ Excel…","⏳ Building Excel file…"));
    try { await new Promise(r => setTimeout(r, 40)); await inThai(fn); setExportMsg(t("✓ สร้างไฟล์เรียบร้อย — ดูที่โฟลเดอร์ดาวน์โหลด","✓ File created — check your Downloads folder")); setTimeout(()=>setExportMsg(""), 3500); }
    catch (e) { console.warn("export failed:", e); setExportMsg(t("⚠ สร้างไฟล์ไม่สำเร็จ ลองใหม่อีกครั้ง","⚠ Couldn't create the file — please try again")); setTimeout(()=>setExportMsg(""), 4500); }
  };

  const handleLogin = (user) => { setSession(user); setSessionState(user); };
  const handleLogout = () => leaveIfDirty(() => {
    clearSession(); setSessionState(null);
    setScreen("home"); setRole(null); setActiveId(null);
  });

  // มีปุ่มลอยมุมขวาล่าง (หลังล็อกอิน) → ให้ CSS เว้นพื้นที่ไว้ ไม่ให้ปุ่มทับตาราง
  useEffect(() => { document.body.classList.toggle("has-fab", !!session); return () => document.body.classList.remove("has-fab"); }, [session]);
  // เตือนก่อนปิด/รีเฟรช/กดปุ่มย้อนของเบราว์เซอร์ ขณะยังมีการแก้ไขที่ไม่ได้บันทึก
  useEffect(() => {
    const onBeforeUnload = (e) => { if (UnsavedGuard.dirty) { e.preventDefault(); e.returnValue = ""; } };
    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, []);

  // โหลด session ตอนเปิดแอป — รองรับ auth.js ได้ทั้งสองแบบ:
  //  • ตัวเดิม: getSession() เป็น synchronous (คืน object/null จาก localStorage)
  //  • ตัวใหม่: getSession() เป็น async (คืน Promise จาก Supabase Auth)
  // Promise.resolve() ครอบให้ทำงานได้ทั้งคู่ ส่วน listener จะ logout เฉพาะตอน
  // เกิดเหตุการณ์ SIGNED_OUT จริง ๆ เท่านั้น (ไม่เผลอล้าง session บน stack เดิม)
  useEffect(() => {
    let mounted = true;
    Promise.resolve(getSession()).then((u) => { if (mounted) { setSessionState(u); setAuthReady(true); } });
    let subscription;
    try {
      const res = supabase.auth?.onAuthStateChange?.((evt, s) => {
        if (evt === "SIGNED_OUT" && mounted) setSessionState(null);
      });
      subscription = res?.data?.subscription;
    } catch { /* auth.js เดิมไม่ได้ใช้ Supabase Auth — ข้ามได้ */ }
    return () => { mounted = false; subscription?.unsubscribe?.(); };
  }, []);

  const fetchProjectData = useCallback(async (id) => {
    // ดึง 6 ส่วนพร้อมกัน (เดิมทีละส่วน ช้ากว่า 6 เท่า)
    const [tc, po, ad, ex, hid, inp] = await Promise.all(
      ["tenders","po","additions","extra","hidden","inplan"].map(k => sgOrThrow(`tcs-${k}-${id}`)));   // OrThrow: โหลดพลาดต้องขึ้นปุ่ม "ลองใหม่" ไม่ใช่โชว์โครงการว่าง
    // ผู้ใช้สลับไปโครงการอื่นระหว่างรอ → ทิ้งผลลัพธ์ของโครงการเก่า (กันข้อมูลโครงการ A ทับโครงการ B)
    if (activeIdRef.current !== id) return;
    setTCosts(tc || {});
    setPO(po || []);
    setAdditions(ad || {});
    setExtraItems(ex || []);
    setHiddenAccounts(hid || []);
    setIncomingPlan(Array.isArray(inp) ? inp : []);
    setProjReadyId(id);
    setProjLoadErr("");
  }, []);

  const fetchProjects = useCallback(async () => {
    const list = await sgOrThrow("tcs-projects");
    if (list) setProjects(list);
  }, []);
  // โหลดรายการบัญชีที่แอดมินแก้ (ใช้ร่วมทุกโครงการ) แล้วทับ ACCOUNTS ในที่
  const fetchAccounts = useCallback(async () => {
    const list = await sgOrThrow("tcs-accounts");
    if (Array.isArray(list) && list.length) { applyAccountList(list); setAccountsRev(v => v + 1); }
  }, []);

  const loadList = async () => {
    setListLoadErr(false);
    try { await fetchAccounts(); await fetchProjects(); setSyncedAt(new Date()); setSyncError(""); }
    catch (e) { console.warn("โหลดรายการโครงการไม่สำเร็จ:", e); setListLoadErr(true); }
    finally { setLoaded(true); }   // กันจอโหลดค้างเสมอ แม้ดึงข้อมูลพลาด
  };
  // โหลดรายการโครงการ "หลังล็อกอินเสร็จ" — สำคัญมากตอนใช้ RLS: ถ้าอ่านก่อน
  // Supabase แนบ token จะโดน DB ปฏิเสธแล้วขึ้น 0 โครงการ ทั้งที่มีสิทธิ์อ่าน
  // ผูกกับ session ไว้ พอล็อกอินเสร็จ (session มีค่า) จะดึงข้อมูลใหม่อัตโนมัติ
  useEffect(() => {
    if (!session) {
      // ออกจากระบบ (กดเอง หรือ token หมดอายุ) → ล้างหน้าจอ/ข้อมูลทั้งหมด ให้คนถัดไปที่ล็อกอินเริ่มที่หน้า
      // รายการโครงการเสมอ (เดิมค้างหน้าเก่า เช่นหน้า Admin → จอว่าง, หรือเห็นรายการโครงการของคนก่อน)
      UnsavedGuard.dirty = false; CalcStore.set(false);
      setScreen("home"); setRole(null); setActiveId(null); setProjReadyId(null); setProjLoadErr(""); setListLoadErr(false);
      setProjects([]); setTCosts({}); setPO([]); setAdditions({}); setExtraItems([]); setHiddenAccounts([]); setIncomingPlan([]);
      undoRef.current = []; redoRef.current = [];
      setLoaded(true); return;
    }
    setLoaded(false);   // ล็อกอินใหม่ → แสดง "กำลังโหลด" จนได้รายการโครงการ (ไม่โชว์ "ยังไม่มีโครงการ" หลอก ๆ)
    loadList();
  }, [fetchProjects, fetchAccounts, session]); // eslint-disable-line

  // เปิดโครงการ: ล้างข้อมูลของโครงการก่อนหน้าทิ้งก่อน แล้วค่อยโหลด — ระหว่างโหลดหน้าแผนกแสดง "กำลังโหลด"
  // (เดิมยังโชว์ตัวเลขโครงการเก่าใต้ชื่อโครงการใหม่ และถ้ากด Export ตอนนั้นจะได้ข้อมูลผิดโครงการ)
  const loadProject = useCallback((id) => {
    setProjLoadErr("");
    return fetchProjectData(id).catch(e => {
      console.warn("โหลดข้อมูลโครงการไม่สำเร็จ:", e);
      if (activeIdRef.current === id) setProjLoadErr(t("โหลดข้อมูลโครงการไม่สำเร็จ — ตรวจเน็ตแล้วกดลองใหม่","Couldn't load project data — check your connection and retry"));
    });
  }, [fetchProjectData]);
  useEffect(() => {
    if (!activeId || !session) return;
    if (projReadyId !== activeId) {
      setProjReadyId(null);
      setTCosts({}); setPO([]); setAdditions({}); setExtraItems([]); setHiddenAccounts([]); setIncomingPlan([]);
    }
    loadProject(activeId);
  }, [activeId, loadProject, session]); // eslint-disable-line

  // ── กัน "กรอกแล้วหาย/เด้งกลับ" ตอนเปิดหลายเครื่อง (ไอดีเดียว โครงการเดียวกัน) ──
  // realtime ของ Supabase ส่ง event การเขียน "กลับมาหาเครื่องที่เขียนเองด้วย" ถ้าเครื่อง
  // นั้นรีบดึงข้อมูลกลับมาทับทันที บางจังหวะจะอ่านได้ค่าเก่า (ก่อน merge ลงเสร็จ) แล้ว
  // setState ทับของที่เพิ่งกรอก → แถวที่เพิ่งบันทึก "เด้งหาย". เปิด 2 เครื่องยิ่งหนักเพราะ
  // event วิ่งชนกันถี่ขึ้น. แก้ 2 ชั้น: (1) ข้าม echo ของการเขียนจากเครื่องตัวเองภายในไม่กี่
  // วินาที (2) ถ้ากำลังกรอก/แก้ค้างอยู่ ให้เลื่อน refetch ไปทำหลังบันทึก ไม่ทับกลางคัน
  const lastWriteRef   = useRef({});    // key -> เวลาเขียนล่าสุดจากเครื่องนี้
  const pendingSyncRef = useRef(false); // มี refetch ค้างรอตอนเลิกแก้ไข
  const SELF_ECHO_MS   = 6000;
  const runProjectSync = useCallback(async (key) => {
    if (Date.now() - (lastWriteRef.current[key] || 0) < SELF_ECHO_MS) return; // echo ของเราเอง — ข้าม
    const isProjKey = activeId && (key === `tcs-tenders-${activeId}` || key === `tcs-po-${activeId}` || key === `tcs-additions-${activeId}` || key === `tcs-extra-${activeId}` || key === `tcs-hidden-${activeId}` || key === `tcs-inplan-${activeId}`);
    if (!(key === "tcs-projects" || key === "tcs-accounts" || isProjKey)) return;
    if (UnsavedGuard.dirty || editModeRef.current) { pendingSyncRef.current = true; return; } // กำลังกรอก — อย่าทับ
    setSyncing(true);
    try {
      if (key === "tcs-projects") await fetchProjects();
      else if (key === "tcs-accounts") { await fetchAccounts(); if (activeId) { await fetchProjectData(activeId); dropUndoFor(null); } }
      else if (isProjKey) { await fetchProjectData(activeId); dropUndoFor(key); }
      setSyncedAt(new Date());
    } catch (e) {
      console.warn("sync realtime ล้มเหลว:", e);
    } finally {
      setSyncing(false); // กันสปินเนอร์ค้างเมื่อ fetch ล้มเหลว
    }
  }, [activeId, fetchProjects, fetchAccounts, fetchProjectData]);

  useEffect(() => {
    if (!session) return; // subscribe realtime หลังล็อกอิน เพื่อให้ RLS ยอมส่ง event
    const channel = supabase.channel("kv_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "kv_store" }, (payload) => {
        runProjectSync(payload.new?.key || payload.old?.key || "");
      }).subscribe();
    return () => supabase.removeChannel(channel);
  }, [session, runProjectSync]);

  // ─── Undo / Redo ───────────────────────────────────────────────────────────
  // ทุกการบันทึกวิ่งผ่าน commit() ซึ่งจดค่าเดิมไว้ก่อนเขียนทับ → กด Ctrl+Z หรือ
  // ปุ่มย้อนกลับ เพื่อคืนค่าเดิมได้ทุกอย่าง (ลบข้อมูล/ลบคอลัมน์/ลบแถว/แก้ตัวเลข/
  // เพิ่มรายการ ฯลฯ) เก็บได้หลายขั้น (สูงสุด 60) และทำซ้ำ (redo) ได้
  const undoRef = useRef([]);
  const redoRef = useRef([]);
  const currentRef = useRef({});
  const [undoInfo, setUndoInfo] = useState({ u: 0, r: 0, label: "" });
  const [editMode, setEditMode] = useState(false); // true เมื่ออยู่ในโหมดแก้ไข — undo/Ctrl+Z ใช้ได้เฉพาะตอนนี้
  const editModeRef = useRef(false); editModeRef.current = editMode;
  // เลิกโหมดแก้ไขแล้ว → ถ้ามี refetch ค้างไว้ระหว่างที่กำลังกรอก ค่อยดึงข้อมูลล่าสุดมาแสดง
  useEffect(() => {
    if (editMode || !pendingSyncRef.current) return;
    pendingSyncRef.current = false;
    if (activeId) fetchProjectData(activeId).then(() => { dropUndoFor(null); setSyncedAt(new Date()); }).catch(() => {});
  }, [editMode, activeId, fetchProjectData]);
  const [selStats, setSelStats] = useState(null); // สรุปตัวเลขที่ลากเลือก (แบบ Excel)
  const [marquee, setMarquee]   = useState(null); // กรอบสี่เหลี่ยมขณะลากเลือก
  const [copied, setCopied]     = useState(false); // สถานะ "คัดลอกแล้ว"
  const dragRef = useRef({ pending:false, active:false, ax:0, ay:0, lastX:0, lastY:0, raf:0, scrollRAF:0, scrollEl:null, suppressClick:false });
  const hiliteRef = useRef([]); // ช่องที่กำลังไฮไลต์ (ไว้คืนค่าเดิมตอนล้าง)
  const selCellsRef = useRef([]); // เซลล์ที่เลือก {top,left,text} ไว้คัดลอก
  currentRef.current = {
    "tcs-projects": projects,
    [`tcs-tenders-${activeId}`]: tenderCosts,
    [`tcs-additions-${activeId}`]: additions,
    [`tcs-po-${activeId}`]: poEntries,
    [`tcs-extra-${activeId}`]: extraItems,
    [`tcs-hidden-${activeId}`]: hiddenAccounts,
    [`tcs-inplan-${activeId}`]: incomingPlan,
  };
  // มีข้อมูลของคนอื่นเข้ามาแทนที่ state ของ key นี้ (ดึงจากเซิร์ฟเวอร์) → ทิ้งขั้น undo/redo ของ key นั้น
  // เพราะ undo = "ย้อนจากค่าปัจจุบันกลับเป็นค่าเก่า" ถ้าค่าปัจจุบันมีงานของคนอื่นปนอยู่ จะย้อนงานเขาทิ้งไปด้วย
  // (key = null → ทิ้งทั้งหมด)
  const dropUndoFor = (key) => {
    const keep = (e) => key !== null && e.key !== key;
    const u = undoRef.current.filter(keep), r = redoRef.current.filter(keep);
    if (u.length === undoRef.current.length && r.length === redoRef.current.length) return;
    undoRef.current = u; redoRef.current = r; syncUndo();
  };
  const syncUndo = () => setUndoInfo({
    u: undoRef.current.length, r: redoRef.current.length,
    label: undoRef.current.length ? undoRef.current[undoRef.current.length - 1].label : "",
  });
  // เขียนลงเซิร์ฟเวอร์ พร้อมลองใหม่อัตโนมัติ 1 ครั้งเมื่อเน็ตสะดุดชั่วคราว ก่อนค่อย
  // แจ้งเตือน (กันเซฟหลุดเพราะ blip เล็ก ๆ). ถ้าส่ง prev มาด้วย จะใช้ ssMerge เพื่อ
  // "รวม" การแก้ของเราลงบนของล่าสุดบนเซิร์ฟเวอร์ (กันทับงานคนอื่นที่แก้พร้อมกัน).
  const persist = (key, value, prev, setState) => {
    // จำเวลาที่เครื่องนี้เขียน key นี้ ทั้งก่อนยิงและหลังสำเร็จ — เพื่อให้ echo ของ realtime
    // ที่วิ่งกลับมา (ซึ่งมาหลังเขียนเสร็จ) ยังอยู่ในกรอบเวลา แล้วถูกข้าม ไม่ดึงมาทับตัวเอง
    const mark = () => { lastWriteRef.current[key] = Date.now(); };
    mark();
    // หลังรวมงานของเครื่องอื่นเข้ามา: คำนวณยอดแม่ใหม่ (Tender / รายเดือน) ให้ตรงกับข้อมูลย่อย
    const pid = (key.match(/^tcs-(?:tenders|additions)-(.+)$/) || [])[1];
    const normalize = pid ? (v) => {
      const extra = currentRef.current[`tcs-extra-${pid}`] || [];
      return key.startsWith("tcs-tenders-") ? rollupTenders(v, extra) : rollupAdditions(v, extra);
    } : undefined;
    const attempt = () => { mark(); return (prev !== undefined ? ssMerge(key, prev, value, { normalize }) : ssOrThrow(key, value)); };
    const done = (written) => {
      mark(); setSyncedAt(new Date()); setSyncError("");
      // ค่าที่เขียนจริงมีงานของเครื่องอื่นรวมอยู่ → แสดงบนจอด้วย (เฉพาะถ้าเรายังไม่ได้แก้ต่อ กันทับงานที่เพิ่งพิมพ์)
      if (setState && written !== undefined && currentRef.current[key] === value && JSON.stringify(written) !== JSON.stringify(value)) {
        setState(written); dropUndoFor(key);
      }
    };
    const noRetry = (e) => e?.code === "42501" || e?.code === "GONE";   // ไม่มีสิทธิ์ / ถูกลบไปแล้ว → ลองซ้ำก็ไม่ผ่าน
    return attempt()
      .then(done)
      .catch(e0 => (noRetry(e0) ? Promise.reject(e0) : new Promise(res => setTimeout(res, 900)).then(attempt))
        .then(done)
        .catch(e => {
          console.warn("บันทึกไม่สำเร็จ:", key, e);
          setSyncError(e?.code === "42501"
            ? t("⚠ บัญชีนี้ไม่มีสิทธิ์บันทึกส่วนนี้ — การแก้ล่าสุดไม่ถูกบันทึก (รีเฟรชหน้าเพื่อดูค่าจริง แล้วติดต่อแอดมิน)","⚠ This account isn't allowed to save this part — your latest change was not saved (refresh to see the real data, then contact an admin)")
            : e?.code === "GONE"
            ? t("⚠ ข้อมูลส่วนนี้ถูกลบไปแล้ว (เช่นโครงการถูกลบ) — การแก้ล่าสุดไม่ถูกบันทึก กรุณารีเฟรชหน้า","⚠ This data was deleted (e.g. the project was removed) — your latest change was not saved. Please refresh")
            : t("⚠ บันทึกไม่สำเร็จ — ข้อมูลล่าสุดอาจยังไม่ถูกบันทึก กรุณาลองใหม่/ตรวจเน็ต","⚠ Save failed — your latest change may not be saved. Please retry / check your connection"));
        }));
  };
  const commit = useCallback((key, next, prev, setState, label) => {
    undoRef.current.push({ key, value: prev, setState, label });
    if (undoRef.current.length > 60) undoRef.current.shift();
    redoRef.current = []; // มีการแก้ใหม่ → ล้าง redo
    setState(next);
    persist(key, next, prev, setState);
    syncUndo();
  }, []);
  const undo = useCallback(() => {
    const e = undoRef.current.pop();
    if (!e) return;
    const cur = currentRef.current[e.key];
    redoRef.current.push({ key: e.key, value: cur, setState: e.setState, label: e.label });
    e.setState(e.value);
    persist(e.key, e.value, cur, e.setState);
    syncUndo();
  }, []);
  const redo = useCallback(() => {
    const e = redoRef.current.pop();
    if (!e) return;
    const cur = currentRef.current[e.key];
    undoRef.current.push({ key: e.key, value: cur, setState: e.setState, label: e.label });
    e.setState(e.value);
    persist(e.key, e.value, cur, e.setState);
    syncUndo();
  }, []);
  // เปลี่ยนโครงการ "หรือ" เปลี่ยนหน้า → ล้างประวัติ undo (กันย้อนข้ามโครงการ/ข้ามบริบท)
  useEffect(() => { undoRef.current = []; redoRef.current = []; syncUndo(); }, [activeId, screen]);
  // คีย์ลัด: Ctrl/Cmd+Z = ย้อนกลับ · Ctrl+Shift+Z หรือ Ctrl+Y = ทำซ้ำ
  // ไม่ดักถ้ากำลังพิมพ์อยู่ในช่องกรอก (ปล่อยให้ undo ของข้อความทำงานตามปกติ)
  useEffect(() => {
    const onKey = (e) => {
      const t = e.target;
      if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable)) return;
      if (!editModeRef.current) return; // ใช้ได้เฉพาะตอนอยู่ในโหมดแก้ไข
      if (!(e.ctrlKey || e.metaKey)) return;
      const k = (e.key || "").toLowerCase();
      if (k === "z" && !e.shiftKey) { e.preventDefault(); undo(); }
      else if ((k === "z" && e.shiftKey) || k === "y") { e.preventDefault(); redo(); }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [undo, redo]);

  // แถบสรุปแบบ Excel — ลากเป็น "กรอบสี่เหลี่ยม" คลุมตัวเลขในตาราง (marquee)
  // รวมเฉพาะตัวเลขที่อยู่ในกรอบ จึงลากลงคอลัมน์เดียวได้ตรง ๆ ไม่ติดเซลล์ข้าง ๆ
  // นับเฉพาะเลขที่มีจุดทศนิยม (ยอดเงิน) จึงไม่รวมรหัสบัญชี/ปี ที่เป็นจำนวนเต็ม
  useEffect(() => {
    const d = dragRef.current;
    const INTERACT = 'input,textarea,select,button,a,[contenteditable="true"]';
    // ยอดเงิน: 1,234.56 · -1,234.56 · (1,234.56) = ติดลบแบบบัญชี (ตัวแดงในวงเล็บ) · "–" / "-" = ศูนย์
    const NUM_RE = /^(-?\d[\d,]*\.\d+|\(\d[\d,]*\.\d+\))$/;
    const isZeroMark = (tv) => tv === "–" || tv === "-";
    const HL = "rgba(37,99,235,0.20)";
    const EDGE = 46, SPEED = 24;
    const clearHilite = () => { hiliteRef.current.forEach(({el,prev}) => { el.style.backgroundColor = prev; }); hiliteRef.current = []; };
    // getScroll: ตำแหน่ง/สเกลของตัวเลื่อน (กล่อง .mscroll ถ้ามี, ไม่งั้นใช้ทั้งหน้าต่าง)
    const getScroll = () => d.scrollEl
      ? (() => { const r = d.scrollEl.getBoundingClientRect(); return { x:d.scrollEl.scrollLeft, y:d.scrollEl.scrollTop, ox:r.left, oy:r.top }; })()
      : { x:window.scrollX, y:window.scrollY, ox:0, oy:0 };
    const compute = () => {
      const s = getScroll();
      // จุดปัจจุบันในพิกัด "เนื้อหา" (คงที่แม้เลื่อน) แล้วทำกรอบเทียบกับ anchor
      const cx = d.lastX - s.ox + s.x, cy = d.lastY - s.oy + s.y;
      const cb = { left:Math.min(d.ax,cx), top:Math.min(d.ay,cy), right:Math.max(d.ax,cx), bottom:Math.max(d.ay,cy) };
      // แปลงกลับเป็นพิกัดจอ (client) ตาม scroll ปัจจุบัน — anchor จึงยึดติดเซลล์เดิม
      const box = { left:cb.left - s.x + s.ox, top:cb.top - s.y + s.oy, right:cb.right - s.x + s.ox, bottom:cb.bottom - s.y + s.oy };
      setMarquee({ left:box.left, top:box.top, width:box.right-box.left, height:box.bottom-box.top });
      clearHilite();
      const nums = []; const cells = new Set(); const cellData = [];
      document.querySelectorAll("table").forEach((tbl) => {
        const walker = document.createTreeWalker(tbl, NodeFilter.SHOW_TEXT, {
          acceptNode(n){
            // ข้ามตัวเลข USD (บรรทัด ≈ $ ใต้ยอดบาท) ไม่ให้ถูกนับ/รวมซ้ำกับบาท
            if (n.parentElement && n.parentElement.closest(".usd-sub")) return NodeFilter.FILTER_REJECT;
            const tv = (n.nodeValue||"").trim();
            return (NUM_RE.test(tv) || isZeroMark(tv)) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
          },
        });
        let node;
        while ((node = walker.nextNode())) {
          const rng = document.createRange(); rng.selectNodeContents(node);
          const r = rng.getBoundingClientRect();
          if (r.width === 0 && r.height === 0) continue;
          if (r.right >= box.left && r.left <= box.right && r.bottom >= box.top && r.top <= box.bottom) {
            const txt = node.nodeValue.trim();
            const neg = /^\(.*\)$/.test(txt);                                   // (1,234.56) → −1,234.56
            const v = isZeroMark(txt) ? 0 : (neg ? -1 : 1) * parseFloat(txt.replace(/[(),]/g, ""));   // "–"/"-" = ยอด 0
            if (!isNaN(v)) {
              nums.push(v);
              cellData.push({ top: r.top, left: r.left, text: isZeroMark(txt) ? "0" : (neg ? "-" + txt.slice(1, -1) : txt) });   // คัดลอกไป Excel เป็นตัวเลข (ติดลบเป็น -)
              const td = node.parentElement && node.parentElement.closest("td"); if (td) cells.add(td);
            }
          }
        }
      });
      selCellsRef.current = cellData;
      cells.forEach((td) => { hiliteRef.current.push({ el:td, prev:td.style.backgroundColor }); td.style.backgroundColor = HL; });
      if (nums.length >= 2) {
        const sum = nums.reduce((a,b)=>a+b,0);
        setSelStats({ count:nums.length, sum, avg:sum/nums.length, min:Math.min(...nums), max:Math.max(...nums), vals:nums });
      } else { setSelStats(null); }
    };
    // เลื่อนตารางอัตโนมัติเมื่อลากชนขอบ (จะได้ลากทั้งแถวที่คอลัมน์เยอะได้)
    const autoScroll = () => {
      if (!d.active) { d.scrollRAF = 0; return; }
      let moved = false;
      const el = d.scrollEl;
      if (el) {
        const r = el.getBoundingClientRect();
        if (d.lastX > r.right - EDGE && el.scrollLeft + el.clientWidth < el.scrollWidth - 1) { el.scrollLeft += SPEED; moved = true; }
        else if (d.lastX < r.left + EDGE && el.scrollLeft > 0) { el.scrollLeft -= SPEED; moved = true; }
        if (d.lastY > r.bottom - EDGE && el.scrollTop + el.clientHeight < el.scrollHeight - 1) { el.scrollTop += SPEED; moved = true; }
        else if (d.lastY < r.top + EDGE && el.scrollTop > 0) { el.scrollTop -= SPEED; moved = true; }
      } else {
        if (d.lastY > window.innerHeight - EDGE) { window.scrollBy(0, SPEED); moved = true; }
        else if (d.lastY < EDGE) { window.scrollBy(0, -SPEED); moved = true; }
      }
      if (moved) { compute(); d.scrollRAF = requestAnimationFrame(autoScroll); }
      else d.scrollRAF = 0;
    };
    const nearEdge = () => {
      const el = d.scrollEl;
      if (el) { const r = el.getBoundingClientRect(); return d.lastX > r.right-EDGE || d.lastX < r.left+EDGE || d.lastY > r.bottom-EDGE || d.lastY < r.top+EDGE; }
      return d.lastY > window.innerHeight-EDGE || d.lastY < EDGE;
    };
    const onDown = (e) => {
      if (e.target instanceof Element && e.target.closest("[data-calc],[data-calc-toggle]")) return;   // ใช้เครื่องคิดเลข → คงการเลือกไว้ (ใส่ผลรวมที่เลือกได้)
      clearHilite(); setSelStats(null); setMarquee(null); selCellsRef.current = []; // คลิกที่ไหนก็ล้างไฮไลต์เดิม
      if (e.button !== 0) return;
      const t = e.target;
      if (!(t instanceof Element) || t.closest(INTERACT) || !t.closest("table")) return;
      // เริ่มลากเลือกสถิติเฉพาะเมื่อเริ่มบนเซลล์ที่เป็น "ยอดเงิน" (มีจุดทศนิยม)
      // ถ้าเริ่มบนเซลล์ข้อความ (รหัสบัญชี/ชื่อรายการ/หัวตาราง) ปล่อยให้เลือก-คัดลอกข้อความได้ตามปกติ
      const startCell = t.closest("td");
      const stTxt = (startCell && startCell.textContent || "").trim();
      if (!startCell || !(/\d[\d,]*\.\d/.test(stTxt) || stTxt === "–" || stTxt === "-")) return;   // "–" = ช่องยอดเงินที่เป็น 0
      d.scrollEl = t.closest(".mscroll") || t.closest(".hscroll") || t.closest(".fatscroll") || null;
      const s = getScroll();
      d.ax = e.clientX - s.ox + s.x; d.ay = e.clientY - s.oy + s.y; // anchor ในพิกัดเนื้อหา
      d.pending = true; d.active = false; d.lastX = e.clientX; d.lastY = e.clientY;
    };
    // Ctrl/Cmd+C = คัดลอกค่าที่เลือก (แบบตาราง) — ไม่ดักถ้ากำลังพิมพ์ในช่องกรอก
    const onCopy = (e) => {
      if (!(e.ctrlKey || e.metaKey) || (e.key||"").toLowerCase() !== "c") return;
      const t = e.target;
      if (t && (t.tagName==="INPUT" || t.tagName==="TEXTAREA" || t.tagName==="SELECT" || t.isContentEditable)) return;
      const tsv = buildTSV(selCellsRef.current);
      if (!tsv) return;
      e.preventDefault();
      if (navigator.clipboard?.writeText) navigator.clipboard.writeText(tsv).then(()=>{ setCopied(true); setTimeout(()=>setCopied(false), 1300); }).catch(()=>{});
    };
    const onMove = (e) => {
      if (!d.pending) return;
      if (!d.active) {
        // d.lastX/Y ยังเป็นตำแหน่งตอน mousedown — ขยับเกิน 5px ถึงเริ่มลากเลือกจริง
        if (Math.abs(e.clientX - d.lastX) + Math.abs(e.clientY - d.lastY) < 5) return;
        d.active = true; document.body.style.userSelect = "none";
      }
      d.lastX = e.clientX; d.lastY = e.clientY;
      e.preventDefault();
      if (!d.raf) d.raf = requestAnimationFrame(() => { d.raf = 0; compute(); });
      if (nearEdge() && !d.scrollRAF) d.scrollRAF = requestAnimationFrame(autoScroll);
    };
    const onUp = () => {
      if (d.raf) { cancelAnimationFrame(d.raf); d.raf = 0; }
      if (d.scrollRAF) { cancelAnimationFrame(d.scrollRAF); d.scrollRAF = 0; }
      if (d.active) d.suppressClick = true; // คงไฮไลต์ไว้ กันคลิกโดนแถว/เซลล์หลังปล่อยเมาส์
      d.pending = false; d.active = false;
      document.body.style.userSelect = "";
      setMarquee(null);
    };
    const onClickCap = (e) => { if (d.suppressClick) { e.stopPropagation(); e.preventDefault(); d.suppressClick = false; } };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
    document.addEventListener("click", onClickCap, true);
    document.addEventListener("keydown", onCopy);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
      document.removeEventListener("click", onClickCap, true);
      document.removeEventListener("keydown", onCopy);
      clearHilite();
    };
  }, []);

  const saveProjects = useCallback((list) => commit("tcs-projects", list, projects, setProjects, t("รายชื่อโครงการ","Project list")), [commit, projects]);
  const saveTenders  = useCallback((tc)   => commit(`tcs-tenders-${activeId}`, tc, tenderCosts, setTCosts, t("ราคาเดิม (Baseline)","Baseline")), [commit, activeId, tenderCosts]);
  const saveAdditions= useCallback((a)    => commit(`tcs-additions-${activeId}`, a, additions, setAdditions, t("ยอดเพิ่มรายเดือน","Monthly additions")), [commit, activeId, additions]);
  const saveExtraItems=useCallback((ex)   => commit(`tcs-extra-${activeId}`, ex, extraItems, setExtraItems, t("รายการ/แถว","Items/rows")), [commit, activeId, extraItems]);
  const saveHiddenAccounts=useCallback((h)=> commit(`tcs-hidden-${activeId}`, h, hiddenAccounts, setHiddenAccounts, t("การซ่อนหมวด","Hidden categories")), [commit, activeId, hiddenAccounts]);
  const saveIncomingPlan=useCallback((v)=> commit(`tcs-inplan-${activeId}`, v, incomingPlan, setIncomingPlan, t("แผนของเข้า","Incoming plan")), [commit, activeId, incomingPlan]);
  const savePO       = useCallback((po)   => commit(`tcs-po-${activeId}`, po, poEntries, setPO, t("PO / จัดซื้อ","PO / Procurement")), [commit, activeId, poEntries]);

  const openProject = (id) => leaveIfDirty(() => {
    setActiveId(id);
    if (session?.role === "admin") { setRole(null); setScreen("roleSelect"); }
    else { setRole(session?.role); setScreen("app"); }
  });
  const deleteProject = async (id) => {
    // ลบทั้งโครงการ = ลบข้อมูลของทุกแผนก → เฉพาะ Admin (กันทั้งปุ่มและในฟังก์ชัน)
    if (session?.role !== "admin") { uiAlert(t("ลบโครงการได้เฉพาะ Admin","Only Admin can delete projects")); return; }
    const proj = projects.find(p => p.id === id);
    const name = (proj?.name || "").trim();
    // ยืนยันแบบ "พิมพ์ชื่อโครงการให้ตรง" — กันเผลอลบ เพราะลบแล้วข้อมูลย่อยหายด้วย
    const typed = await uiPrompt(t(
      `⚠️ ลบโครงการ "${name}" ?\n\n` +
      `ข้อมูลทั้งหมดของโครงการนี้จะถูกลบด้วย:\n` +
      `• Tender Cost (ราคาเดิม)\n• PO / จัดซื้อ\n• ยอดเพิ่มรายเดือน · รายการเพิ่ม · หมวดที่ซ่อน\n\n` +
      `กู้คืนได้จาก Admin → กู้คืนข้อมูล (ได้ถึงสแนปช็อตล่าสุด 12:00/18:00)\n\n` +
      `ถ้าแน่ใจ พิมพ์ชื่อโครงการให้ตรงเพื่อยืนยัน:\n${name}`,
      `⚠️ Delete project "${name}"?\n\n` +
      `All of this project's data will be deleted too:\n` +
      `• Tender Cost (baseline)\n• PO / procurement\n• Monthly additions · extra items · hidden categories\n\n` +
      `Admin can restore it from Admin → Restore data (latest 12:00/18:00 snapshot)\n\n` +
      `If you're sure, type the project name exactly to confirm:\n${name}`
    ), { placeholder: name, match: name, okLabel: t("ลบโครงการ","Delete project"), danger: true });
    if (typed == null) return;                                   // กดยกเลิก
    if (typed.trim() !== name) { uiAlert(t("ชื่อโครงการไม่ตรง — ยกเลิกการลบแล้ว","Name doesn't match — deletion cancelled")); return; }
    // ไม่เข้า quick-undo เพราะการลบโครงการลบคีย์ย่อยด้วย — กู้ทั้งโครงการทำผ่านหน้า
    // Admin กู้คืนข้อมูล (kv_history เก็บไว้ให้ครบทุกคีย์)
    const next = projects.filter(p => p.id !== id);
    // ใช้ ssMerge (merge by id) แทน ss ธรรมดา — ตัดเฉพาะโครงการที่ลบออก โดยไม่ทับ
    // โครงการที่คนอื่นเพิ่งเพิ่มพร้อมกัน และแจ้ง error ถ้าบันทึกไม่สำเร็จ
    setProjects(next);
    lastWriteRef.current["tcs-projects"] = Date.now();
    // 1) เอาออกจากรายชื่อโครงการให้สำเร็จ "ก่อน" — ถ้าไม่สำเร็จ ห้ามลบข้อมูลย่อย (เดิมลบข้อมูลไปก่อน
    //    ถ้าเอาออกจากรายชื่อไม่ผ่าน โครงการจะกลับมาแบบข้อมูลว่างเปล่า)
    try {
      await ssMerge("tcs-projects", projects, next);
      lastWriteRef.current["tcs-projects"] = Date.now();
    } catch (e) {
      console.warn("ลบโครงการไม่สำเร็จ:", e);
      setProjects(projects);   // คืนรายการเดิม — ข้อมูลของโครงการยังอยู่ครบ
      setSyncError(t("⚠ ลบโครงการไม่สำเร็จ — ข้อมูลยังอยู่ครบ ตรวจเน็ตแล้วลองใหม่","⚠ Couldn't delete the project — its data is intact. Check your connection and try again"));
      return;
    }
    if (activeId === id) setActiveId(null);
    // 2) ลบข้อมูลย่อยทุกส่วน — ส่วนไหนพลาดแจ้งเตือน (ไม่ปล่อยเงียบ) ข้อมูลเก่ากู้ได้ที่ Admin → กู้คืนข้อมูล
    const keys = ["tenders","po","additions","extra","hidden","inplan"].map(k => `tcs-${k}-${id}`);
    const results = await Promise.allSettled(keys.map(k => sdOrThrow(k)));
    const failed = keys.filter((_, i) => results[i].status === "rejected");
    if (failed.length) {
      console.warn("ลบข้อมูลย่อยของโครงการไม่ครบ:", failed);
      setSyncError(t(`⚠ เอาโครงการออกแล้ว แต่ลบข้อมูลย่อยไม่ครบ ${failed.length} ส่วน — ไม่กระทบโครงการอื่น`,`⚠ Project removed, but ${failed.length} data part(s) could not be deleted — other projects are not affected`));
    } else {
      setSyncedAt(new Date()); setSyncError("");
    }
  };
  const activeProject = projects.find(p => p.id === activeId) || { name:"", area:"", panels:"" };
  const updateProject = (fields) => saveProjects(projects.map(p => p.id === activeId ? {...p,...fields} : p));

  // บันทึกว่าผู้ใช้คนนี้เปิดเวอร์ชันไหนอยู่ (Admin → "เวอร์ชันที่ใช้") — เงียบถ้ายังไม่ได้รัน SQL สิทธิ์
  useEffect(() => { if (session?.id) recordSeen({ name: session.name || "", role: session.role || "", ...appBuild() }); }, [session?.id]); // eslint-disable-line

  const updateBanner = newBuild && !updateLater ? (
    <div data-update-banner role="status" style={{position:"fixed",left:"50%",top:12,transform:"translateX(-50%)",zIndex:300,maxWidth:"94vw",
      background:"#1e3a8a",color:"#fff",borderRadius:12,padding:"10px 14px",boxShadow:"0 10px 30px rgba(15,23,42,0.3)",fontSize:13,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
      <span style={{flex:1,minWidth:200}}>{t("มีเวอร์ชันใหม่ของแอป — รีเฟรชเพื่อใช้เวอร์ชันเดียวกับทุกคน","A new version is available — refresh so everyone uses the same version")} <span style={{opacity:0.7,fontVariantNumeric:"tabular-nums"}}>(build {newBuild})</span></span>
      <button data-update-now onClick={() => leaveIfDirty(() => window.location.reload())} style={{background:"#fff",color:"#1e3a8a",border:"none",borderRadius:8,padding:"6px 14px",fontWeight:700,cursor:"pointer",fontSize:13}}>{t("รีเฟรชเลย","Refresh now")}</button>
      <button onClick={() => setUpdateLater(true)} style={{background:"none",color:"#c7d2fe",border:"none",cursor:"pointer",fontSize:12}}>{t("ภายหลัง","Later")}</button>
    </div>
  ) : null;

  if (!authReady) {
    return (
      <>
        <style>{GLOBAL_CSS}</style>
        <Loader />
      </>
    );
  }

  if (!session) {
    return (
      <>
        <style>{GLOBAL_CSS}</style>
        <LoginScreen onLogin={handleLogin} />
        {updateBanner}
        <DialogHost />
      </>
    );
  }

  if (!loaded) return <Loader />;

  // Non-admins can only ever act as the role tied to their account,
  // even if they somehow land on screen "roleSelect" or "app" with a stale role.
  const effectiveRole = session.role === "admin" ? role : session.role;

  const sharedProps = { project:activeProject, tenderCosts, poEntries, saveTenders, savePO,
    additions, saveAdditions, extraItems, saveExtraItems, hiddenAccounts, saveHiddenAccounts,
    incomingPlan, saveIncomingPlan,
    updateProject,
    onBack: () => leaveIfDirty(() => setScreen(session.role === "admin" ? "roleSelect" : "home")),
    onHome: () => leaveIfDirty(() => setScreen("home")),   // ปุ่ม Home → หน้าเลือกโครงการ (ทุกโรล)
    // ปุ่ม "เลือกแผนก" → หน้าแรกของแต่ละแผนก (เฉพาะ admin ที่สลับแผนกได้)
    onDept: session.role === "admin" ? () => leaveIfDirty(() => setScreen("roleSelect")) : null,
    syncedAt, syncing, session, onLogout: handleLogout, setEditMode };

  return (
    <>
      <style>{GLOBAL_CSS}</style>
      {syncError && (
        <div style={{position:"fixed",left:"50%",top:16,transform:"translateX(-50%)",zIndex:200,maxWidth:"92vw",
          background:"#fef2f2",color:"#991b1b",border:"1px solid #ef4444",borderRadius:12,padding:"10px 16px",
          boxShadow:"0 8px 28px rgba(15,23,42,0.18)",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:12}}>
          <span style={{flex:1}}>{syncError}</span>
          <button onClick={()=>setSyncError("")} style={{border:"none",background:"none",color:"#991b1b",cursor:"pointer",fontSize:16,fontWeight:650,lineHeight:1}}>×</button>
        </div>
      )}
      {exportMsg && (
        <div style={{position:"fixed",left:"50%",bottom:BOTTOM(22),transform:"translateX(-50%)",zIndex:200,
          background:"#0f172a",color:"#e2e8f0",borderRadius:10,padding:"10px 18px",boxShadow:"0 8px 28px rgba(15,23,42,0.28)",
          fontSize:13,fontWeight:600,whiteSpace:"nowrap"}}>{exportMsg}</div>
      )}
      {marquee && marquee.width > 2 && marquee.height > 2 && (
        <div style={{position:"fixed",left:marquee.left,top:marquee.top,width:marquee.width,height:marquee.height,
          background:"rgba(37,99,235,0.06)",border:"none",zIndex:97,pointerEvents:"none"}}/>
      )}
      {session && <><TableTopButton /><ScrollTopFab /><CalcFab /><CalculatorPopup selSum={selStats ? selStats.sum : null} /></>}
      {updateBanner}
      <DialogHost />
      {selStats && (
        <div style={{position:"fixed",right:FAB_GAP + FAB_SIZE + 12,bottom:BOTTOM(20),zIndex:96,maxWidth:`calc(100vw - ${FAB_GAP + FAB_SIZE + 24}px)`,display:"flex",alignItems:"center",gap:0,
          background:"#1e293b",color:"#e2e8f0",borderRadius:10,padding:"8px 4px",boxShadow:"0 8px 28px rgba(15,23,42,0.28)",
          fontSize:12,fontVariantNumeric:"tabular-nums",overflow:"hidden"}}>
          {(() => {
            const selRate = effRate(activeProject);   // อัตราแลกเปลี่ยน (0 = ปิด/ไม่โชว์ $)
            const segs = [
              {label:t("ผลรวม","Sum"), raw:selStats.sum, clr:"#34d399", money:true},
              {label:t("เฉลี่ย","Avg"), raw:selStats.avg, clr:"#93c5fd", money:true},
              {label:t("นับ","Count"),   text:String(selStats.count), clr:"#fcd34d", money:false},
              {label:t("ต่ำสุด","Min"), raw:selStats.min, clr:"#cbd5e1", money:true},
              {label:t("สูงสุด","Max"), raw:selStats.max, clr:"#cbd5e1", money:true},
            ];
            return segs.map((s,i)=>(
              <span key={s.label} style={{display:"flex",alignItems:"center",gap:6,padding:"0 12px",borderLeft:i?"1px solid #334155":"none"}}>
                <span style={{color:"#94a3b8",fontFamily:"system-ui,sans-serif",fontSize:11}}>{s.label}</span>
                <span style={{display:"flex",flexDirection:"column",alignItems:"flex-end",lineHeight:1.15}}>
                  <b style={{color:s.clr}}>{s.money ? fmt(s.raw) : s.text}</b>
                  {s.money && selRate>0 && <b style={{color:"#34d399",fontSize:11,fontWeight:650}}>${fmt(s.raw/selRate)}</b>}
                </span>
              </span>
            ));
          })()}
          <button onClick={()=>{ const tsv=buildTSV(selCellsRef.current); if(tsv&&navigator.clipboard?.writeText){ navigator.clipboard.writeText(tsv).then(()=>{setCopied(true); setTimeout(()=>setCopied(false),1300);}); } }}
            title={t("คัดลอกค่าที่เลือก (Ctrl+C)","Copy selected values (Ctrl+C)")}
            style={{marginLeft:6,marginRight:4,display:"flex",alignItems:"center",gap:5,border:"none",cursor:"pointer",borderRadius:8,padding:"6px 12px",
              fontFamily:"system-ui,sans-serif",fontSize:12,fontWeight:600,background:copied?"#065f46":"#334155",color:"#fff"}}>
            {copied ? t("✓ คัดลอกแล้ว","✓ Copied") : <><Ico name="copy" size={14} /> {t("คัดลอก","Copy")}</>}
          </button>
        </div>
      )}
      {editMode && (undoInfo.u > 0 || undoInfo.r > 0) && (
        <div style={{position:"fixed",left:20,bottom:BOTTOM(20),zIndex:95,display:"flex",gap:6,alignItems:"center",
          background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:12,padding:"7px 9px",boxShadow:"0 8px 28px rgba(15,23,42,0.16)"}}>
          <button onClick={undo} disabled={!undoInfo.u} title={t("ย้อนกลับ (Ctrl+Z)","Undo (Ctrl+Z)")}
            style={{display:"flex",alignItems:"center",gap:6,background:undoInfo.u?T.blue:"#e2e8f0",color:undoInfo.u?"#fff":"#94a3b8",
              border:"none",borderRadius:8,padding:"7px 12px",fontSize:13,fontWeight:600,cursor:undoInfo.u?"pointer":"default"}}>
            ↩︎ {t("ย้อนกลับ","Undo")}
          </button>
          <button onClick={redo} disabled={!undoInfo.r} title={t("ทำซ้ำ (Ctrl+Shift+Z)","Redo (Ctrl+Shift+Z)")}
            style={{background:undoInfo.r?T.blueLight:"transparent",color:undoInfo.r?T.blue:"#cbd5e1",
              border:`1px solid ${undoInfo.r?T.blue:T.cardBorder}`,borderRadius:8,padding:"7px 10px",fontSize:13,fontWeight:600,cursor:undoInfo.r?"pointer":"default"}}>
            ↪︎
          </button>
          {undoInfo.label && (
            <span style={{fontSize:11,color:T.textMuted,maxWidth:170,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",paddingRight:2}}>
              {t("ล่าสุด","Last")}: {undoInfo.label}
            </span>
          )}
        </div>
      )}
      <ErrorBoundary>
      {screen === "home" && (
        <HomeScreen projects={projects} loadErr={listLoadErr} onRetryLoad={() => { setLoaded(false); loadList(); }} saveProjects={saveProjects} openProject={openProject}
          deleteProject={deleteProject} newProjModal={newProjModal} setNewProjModal={setNewProjModal}
          syncedAt={syncedAt} syncing={syncing} session={session} onLogout={handleLogout}
          onOpenAdmin={() => setScreen("admin")} />
      )}
      {screen === "admin" && session.role === "admin" && (
        <AdminPanel onBack={() => setScreen("home")} onLogout={handleLogout} session={session} />
      )}
      {screen === "roleSelect" && session.role === "admin" && (
        <RoleSelect project={activeProject} updateProject={updateProject}
          onSelect={r=>{ setRole(r); setScreen("app"); }} onBack={()=>setScreen("home")} />
      )}
      {screen === "app" && projReadyId !== activeId && (
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:T.bg}}>
          <div style={{textAlign:"center",maxWidth:420,padding:24}}>
            {projLoadErr ? (<>
              <div style={{fontSize:32,marginBottom:10}}>⚠️</div>
              <div style={{fontSize:14,color:T.textPrimary,fontWeight:600,marginBottom:14}}>{projLoadErr}</div>
              <div style={{display:"flex",gap:8,justifyContent:"center"}}>
                <button className="btn-primary" onClick={()=>loadProject(activeId)}>{t("ลองใหม่","Retry")}</button>
                <button className="btn-ghost" onClick={()=>setScreen("home")}>{t("กลับหน้าโครงการ","Back to projects")}</button>
              </div>
            </>) : (<>
              <div style={{width:40,height:40,border:`3px solid ${T.blueMid}`,borderTopColor:T.blue,borderRadius:"50%",animation:"spin 0.7s linear infinite",margin:"0 auto 12px"}}/>
              <div style={{fontSize:13,color:T.textSecondary}}>{t("กำลังโหลดข้อมูลโครงการ","Loading project data")} {activeProject.name}…</div>
              <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
            </>)}
          </div>
        </div>
      )}
      {screen === "app" && projReadyId === activeId && effectiveRole === "qs"          && (
        <QSView {...sharedProps} runExportFn={runExport} onExport={() => runExport(() =>
          // ใช้ฟอร์มเดียวกันทั้งเปิด/ปิด USD — ปิด USD ก็แค่ไม่มีคอลัมน์ USD (ฟอร์มเหมือนกัน)
          Promise.resolve(exportQSExcel(activeProject, tenderCosts, additions, extraItems, hiddenAccounts))
        )} />
      )}
      {screen === "app" && projReadyId === activeId && effectiveRole === "procurement" && (
        <ProcurementView {...sharedProps} onExport={() => runExport(() =>
          // ใช้ฟอร์มเดียวกันทั้งเปิด/ปิด USD — ปิด USD ก็แค่ไม่มีคอลัมน์ USD (ฟอร์มเหมือนกัน)
          Promise.resolve(exportProcurementExcel(activeProject, poEntries, incomingPlan, tenderCosts, additions, extraItems, hiddenAccounts))
        )} />
      )}
      {screen === "app" && projReadyId === activeId && effectiveRole === "accounting"  && (
        <AccountingView {...sharedProps} onExport={() => runExport(() => exportAccountingExcel(activeProject, tenderCosts, additions, poEntries, extraItems, hiddenAccounts, incomingPlan))} />
      )}
      </ErrorBoundary>
    </>
  );
}
