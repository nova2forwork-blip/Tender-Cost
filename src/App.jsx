import { useState, useEffect, useLayoutEffect, useCallback, useRef, Fragment, Component } from "react";
import * as XLSX from "xlsx-js-style";
import { supabase, sg, sgOrThrow, sgMany, ssOrThrow, ssMerge, sdOrThrow, loadKvSnapshots, restoreKvSnapshot } from "./supabase.js";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend, CartesianGrid } from "recharts";
import {
  ROLE_LABELS, getSession, setSession, clearSession, verifyLogin,
  loadUsers, createUser, resetPassword, toggleActive, deleteUser, loadLogs,
} from "./auth.js";
// (saveUsers is used internally by auth.js helpers above, not needed directly here)

// ─── Master Data ──────────────────────────────────────────────────────────────
const ACCOUNTS = [
  { code:"511010", name:"Glass Purchases",                        group:"Materials"    },
  { code:"511015", name:"Screw & Fastener Purchases",             group:"Materials"    },
  { code:"511017", name:"Cast in Channel",                        group:"Materials"    },
  { code:"511020", name:"Gaskets Purchases",                      group:"Materials"    },
  { code:"511025", name:"Silicone & Sealant Purchases",           group:"Materials"    },
  { code:"511030", name:"Glazing Material Purchases",             group:"Materials"    },
  { code:"511035", name:"Miscellaneous Purchases",                group:"Materials"    },
  { code:"511037", name:"Hardware Purchases",                     group:"Materials"    },
  { code:"511040", name:"Tools & Consumables Purchases",          group:"Materials"    },
  { code:"511042", name:"Accessories Purchases",                  group:"Materials"    },
  { code:"511045", name:"Aluminium Extrusion Purchases",          group:"Aluminium"    },
  { code:"511050", name:"Aluminium Sheet Purchases",              group:"Aluminium"    },
  { code:"511051", name:"Extra Charge for Extrusion",             group:"Aluminium"    },
  { code:"511052", name:"Dies & Moulds Purchases",                group:"Aluminium"    },
  { code:"511053", name:"Aluminium Grates and Grids Purchases",   group:"Aluminium"    },
  { code:"511055", name:"Steel Purchases",                        group:"Steel"        },
  { code:"511060", name:"Steel Components Purchases",             group:"Steel"        },
  { code:"511062", name:"Iron Purchases",                         group:"Steel"        },
  { code:"511063", name:"Galvanized Purchases",                   group:"Steel"        },
  { code:"511065", name:"Stainless Steel Sheets Purchases",       group:"Steel"        },
  { code:"511070", name:"Composite Panel Purchases",              group:"Materials"    },
  { code:"511075", name:"Mechanical Components Purchases",        group:"Materials"    },
  { code:"511080", name:"Material for Protection Purchases",      group:"Materials"    },
  { code:"511085", name:"Insulation Material Purchases",          group:"Materials"    },
  { code:"511090", name:"Waterproofing Membranes Purchases",      group:"Materials"    },
  { code:"511093", name:"Extra Charge for Paint",                 group:"Finishing"    },
  { code:"511095", name:"PVF2 Expenses",                          group:"Finishing"    },
  { code:"511100", name:"Hot Dipped Galvanized (HDG)",            group:"Finishing"    },
  { code:"511105", name:"Powder Painting Expenses",               group:"Finishing"    },
  { code:"511110", name:"Anodising Expenses",                     group:"Finishing"    },
  { code:"511113", name:"Chromate Expenses",                      group:"Finishing"    },
  { code:"511115", name:"Varnishing Steel Expenses",              group:"Finishing"    },
  { code:"511120", name:"Sundry Chemical Treatments Expenses",    group:"Finishing"    },
  { code:"511125", name:"Packing Materials Expenses",             group:"Logistics"    },
  { code:"511128", name:"Installation Equipments Expenses",       group:"Installation" },
  { code:"511130", name:"Installation Expenses",                  group:"Installation" },
  { code:"511135", name:"Subcontractors",                         group:"Installation" },
  { code:"511140", name:"External Design Costs",                  group:"Design & Eng" },
  { code:"511145", name:"Other Design Costs",                     group:"Design & Eng" },
  { code:"511150", name:"External Engineering Costs",             group:"Design & Eng" },
  { code:"511155", name:"Other Engineering Costs",                group:"Design & Eng" },
  { code:"511160", name:"Health & Safety Costs",                  group:"Site"         },
  { code:"511165", name:"Skip and Rubbish Removal Costs",         group:"Site"         },
  { code:"511166", name:"Local Charge for Shipment",              group:"Logistics"    },
  { code:"511167", name:"Ocean Freight for Shipment",             group:"Logistics"    },
  { code:"511168", name:"U.S. Customs",                           group:"Logistics"    },
  { code:"511169", name:"Destination Charge for Shipment",        group:"Logistics"    },
  { code:"511170", name:"Transport Expenses on Purchases",        group:"Logistics"    },
  { code:"511173", name:"Transport Expenses on Sales",            group:"Logistics"    },
  { code:"511175", name:"Other Expenses on Transport Expenses",   group:"Logistics"    },
  { code:"511178", name:"Other Expenses on Transport Sales",      group:"Logistics"    },
  { code:"511180", name:"Customers Expenses on Projects",         group:"Site"         },
  { code:"511185", name:"PJM Travel and Accommodation Expenses",  group:"Site"         },
  { code:"511205", name:"Custom Duties and Operations",           group:"Logistics"    },
  { code:"511300", name:"Internal Production",                    group:"Production"   },
  { code:"511305", name:"External Production",                    group:"Production"   },
  { code:"511350", name:"Testing Expenses",                       group:"QA/QC"        },
  { code:"511353", name:"Mock Up Expenses",                       group:"QA/QC"        },
  { code:"511355", name:"Cost of NCR",                            group:"QA/QC"        },
  { code:"521005", name:"Minor Factory Equipment Purchases",      group:"Factory"      },
  { code:"521025", name:"Insurance: Shipment",                    group:"Logistics"    },
  { code:"521110", name:"Toll and Parking Expenses",              group:"Site"         },
  { code:"521250", name:"Other General Expenses",                 group:"Other"        },
];

const GROUPS      = [...new Set(ACCOUNTS.map(a => a.group))];
const PO_STATUS   = ["PO Issued","Delivered","Invoiced","Paid"];
// ชื่อสถานะ PO สำหรับแสดงผล (ค่าที่เก็บในข้อมูลยังเป็นภาษาอังกฤษเหมือนเดิม)
const PO_STATUS_TH = { "PO Issued":"ออก PO แล้ว", "Delivered":"ส่งของแล้ว", "Invoiced":"วางบิลแล้ว", "Paid":"จ่ายแล้ว", "All":"ทั้งหมด" };
const poStatusLabel = (s) => t(PO_STATUS_TH[s] || s, s);
const STATUS_CLR  = { Planning:"#94a3b8", Pending:"#94a3b8","PO Issued":"#3b82f6",Delivered:"#f59e0b",Invoiced:"#8b5cf6",Paid:"#10b981" };
const STATUS_BG   = { Planning:"#f1f5f9", Pending:"#f1f5f9","PO Issued":"#eff6ff",Delivered:"#fffbeb",Invoiced:"#f5f3ff",Paid:"#f0fdf4" };
const GRP_COLORS  = ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6","#06b6d4","#f97316","#ec4899"];

// ── รายการบัญชีแก้ได้ (แอดมิน) ─────────────────────────────────────────────────
//  เก็บใน kv 'tcs-accounts' (ใช้ร่วมทุกโครงการ). applyAccountList จะทับ ACCOUNTS/
//  GROUPS "ในที่" (mutate) เพื่อให้ทุกจุดที่อ่าน ACCOUNTS เห็นค่าล่าสุดหลัง re-render.
function applyAccountList(list) {
  if (!Array.isArray(list) || !list.length) return;
  const clean = list.filter(a => a && a.code).map(a => ({ code: String(a.code).trim(), name: a.name || "", group: a.group || "Other" }));
  if (!clean.length) return;
  ACCOUNTS.splice(0, ACCOUNTS.length, ...clean);
  GROUPS.splice(0, GROUPS.length, ...[...new Set(ACCOUNTS.map(a => a.group))]);
}
// รายการ "งานเพิ่ม" ของโครงการที่เปิดอยู่ (App อัปเดตทุกครั้งที่ render) — ให้ตัวช่วยระดับโมดูล
// เช่น poPayLines / ป๊อปอัพ PO / หน้าติดตาม หาชื่อของรหัสงานเพิ่มได้ ไม่ขึ้น "—"
let _EXTRA_ITEMS = [];
const setExtraRegistry = (items) => { _EXTRA_ITEMS = Array.isArray(items) ? items : []; };
const accountOf = (code) => {
  if (!code) return null;
  const a = ACCOUNTS.find(x => x.code === code);
  if (a) return a;
  const e = _EXTRA_ITEMS.find(x => x.code === code);
  return e ? { code: e.code, name: e.name || "", group: e.group || "Other", isExtra: true } : null;
};
// ย้ายข้อมูลเมื่อเปลี่ยนรหัสบัญชี (ข้ามทุกโครงการ) — renameMap = { oldCode: newCode }
async function migrateAccountCodes(renameMap) {
  const map = Object.fromEntries(Object.entries(renameMap || {}).filter(([o, n]) => o && n && o !== n));
  if (!Object.keys(map).length) return;
  const ren = (code) => map[code] || code;
  const renObjKeys = (obj) => {  // คีย์ที่เป็น "code" หรือ "code:col" ให้เปลี่ยนฐาน (ข้ามคีย์ meta $…)
    if (!obj || typeof obj !== "object") return obj;
    const out = {};
    for (const [k, v] of Object.entries(obj)) {
      if (k.startsWith("$")) { out[k] = v; continue; }
      const ci = k.indexOf(":"); const base = ci >= 0 ? k.slice(0, ci) : k; const nb = map[base];
      out[nb ? (ci >= 0 ? nb + k.slice(ci) : nb) : k] = v;
    }
    return out;
  };
  // ── ทำเป็น 2 เฟส กัน "ย้ายไม่ครบเงียบ ๆ" ──────────────────────────────────
  //  เฟส 1: อ่านทุกคีย์ของทุกโครงการ + สร้าง payload ที่เปลี่ยนรหัสแล้วเก็บไว้ในหน่วยความจำ
  //         ใช้ sgOrThrow — ถ้าอ่านพลาดแม้แต่คีย์เดียวจะ throw ทันที (ยังไม่เขียนอะไรเลย)
  //  เฟส 2: เขียนทุก payload ด้วย ssOrThrow — ถ้าพลาดจะ throw ให้ผู้เรียกรู้และหยุด
  //         ก่อนไปเขียน tcs-accounts ใหม่ (ไม่ปล่อยให้รหัสใหม่โผล่ทั้งที่ข้อมูลยังไม่ย้าย)
  const projects = (await sgOrThrow("tcs-projects")) || [];
  const writes = []; // { key, value }
  for (const proj of projects) {
    const id = proj?.id; if (!id) continue;
    const t = await sgOrThrow(`tcs-tenders-${id}`);   if (t && typeof t === "object") writes.push({ key:`tcs-tenders-${id}`, value: renObjKeys(t) });
    const ad = await sgOrThrow(`tcs-additions-${id}`); if (ad && typeof ad === "object") {
      const nad = {}; for (const [m, mo] of Object.entries(ad)) nad[m] = (m.startsWith("$") || typeof mo !== "object") ? mo : renObjKeys(mo);
      writes.push({ key:`tcs-additions-${id}`, value: nad });
    }
    const po = await sgOrThrow(`tcs-po-${id}`);  if (Array.isArray(po))  writes.push({ key:`tcs-po-${id}`,  value: po.map(p => ({ ...p, items: (p.items || []).map(it => ({ ...it, code: ren(it.code) })) })) });
    const ex = await sgOrThrow(`tcs-extra-${id}`); if (Array.isArray(ex)) writes.push({ key:`tcs-extra-${id}`, value: ex.map(e => ({ ...e, code: ren(e.code), ...(e.parentCode ? { parentCode: ren(e.parentCode) } : {}) })) });
    const hid = await sgOrThrow(`tcs-hidden-${id}`); if (Array.isArray(hid)) writes.push({ key:`tcs-hidden-${id}`, value: hid.map(ren) });
    const inp = await sgOrThrow(`tcs-inplan-${id}`); if (Array.isArray(inp)) writes.push({ key:`tcs-inplan-${id}`, value: inp.map(pl => ({ ...pl, items: (pl.items || []).map(it => ({ ...it, code: ren(it.code) })) })) });
  }
  // เฟส 2: เขียนจริง (อ่านครบทุกอย่างแล้วเท่านั้นถึงเริ่มเขียน)
  for (const w of writes) await ssOrThrow(w.key, w.value);
}

// ─── Incoming / Payment tracking status ────────────────────────────────────
// A PO's incoming status is derived from its planned/actual dates rather than
// stored directly, so it's always in sync with today's date.
// วันนี้ตาม "ปฏิทินท้องถิ่น" (ไม่ใช้ UTC เพื่อไม่ให้ข้ามวันตอนเช้ามืดในโซน UTC+7)
const todayStr = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
};

// ─── กันข้อมูลหาย: ธง "มีการแก้ไขที่ยังไม่บันทึก" (โมดูลเดียวทั้งแอป) ───────────────
// หน้าจอที่แก้แบบ draft (QS ราคาเดิม/รายเดือน) จะตั้งค่า .dirty ระหว่างพิมพ์ค้าง
// แล้วปุ่มออกจากหน้า/สลับโครงการ/ล็อกเอาต์/รีเฟรชเบราว์เซอร์ จะถามยืนยันก่อนทิ้ง
const UnsavedGuard = { dirty: false };
// ─── กล่องข้อความในแอป (แทน alert / confirm / prompt ของเบราว์เซอร์) ───────────
// หน้าตาเข้าชุดกับแอป อ่านง่ายบนมือถือ และไม่บล็อกทั้งหน้า · คืน Promise
//   uiAlert(msg)            → รอจนกด "ตกลง"
//   uiConfirm(msg, opts)    → true / false      opts: { title, okLabel, cancelLabel, danger }
//   uiPrompt(msg, opts)     → ข้อความที่พิมพ์ / null   opts: { placeholder, match }
const DialogStore = {
  q: [], subs: new Set(),
  push(d) { this.q = [...this.q, d]; this.subs.forEach(f => f(this.q)); },
  shift() { this.q = this.q.slice(1); this.subs.forEach(f => f(this.q)); },
};
const uiAlert   = (message, opts = {}) => new Promise(res => DialogStore.push({ kind: "alert",   message, ...opts, resolve: () => res() }));
const uiConfirm = (message, opts = {}) => new Promise(res => DialogStore.push({ kind: "confirm", message, ...opts, resolve: res }));
const uiPrompt  = (message, opts = {}) => new Promise(res => DialogStore.push({ kind: "prompt",  message, ...opts, resolve: res }));
// ออกจากหน้าตอนมีค่าที่ยังไม่บันทึก → ถามในแอปก่อน แล้วค่อยทำ fn
const leaveIfDirty = (fn) => {
  if (!UnsavedGuard.dirty) { fn(); return; }
  uiConfirm(t("มีการแก้ไขที่ยังไม่บันทึก — ออกจากหน้านี้โดยไม่บันทึก?","You have unsaved changes — leave without saving?"),
    { okLabel: t("ออกโดยไม่บันทึก","Leave without saving"), cancelLabel: t("อยู่ต่อ","Stay"), danger: true })
    .then(ok => { if (ok) { UnsavedGuard.dirty = false; fn(); } });
};
// "2026-07-02" + 30 -> "2026-08-01" — คำนวณด้วย UTC ล้วนทั้งไปและกลับ กัน bug timezone
// (ของเดิม parse เป็น local แต่อ่านกลับเป็น UTC ทำให้ในไทยคลาดไป 1 วันและตกเดือนผิด)
const addDays = (dateStr, days) => {
  if (!dateStr) return "";
  const [y,m,dd] = String(dateStr).split("-").map(Number);
  if (!y || !m || !dd) return "";
  const d = new Date(Date.UTC(y, m-1, dd));
  d.setUTCDate(d.getUTCDate()+days);
  return d.toISOString().slice(0,10);
};

// ─── Multi-code / multi-batch PO helpers ───────────────────────────────────
// A single PO can now be split across several Account Codes (each with its
// own amount that rolls up into the PO total) and can arrive in several
// delivery batches instead of a single date. Older records saved before this
// existed still carry a single `code`/`amount` and a single
// `incomingPlan`/`actualReceived` — these getters transparently upgrade them
// so both old and new records work everywhere without a one-off migration.
// ─── PO data model + migration ─────────────────────────────────────────────
// New shape: ONE supplier per PO. Each account-code line item carries its own
// `store` (qty already on hand, netted out of the % base) and its own list of
// delivery/payment `rounds`, so one material can arrive in several shipments.
// Payment is automatic — a round counts as paid once its due date (received
// date + credit term) has arrived. migratePO() upgrades every older record
// (multi-supplier, supplier-level rounds, item.supplierId, PO-level
// deliveries, paymentType "credit30") on read, so old + new records work
// everywhere with no destructive one-off migration.
const DEFAULT_CREDIT_DAYS = 30;

const isNewPO = (p) => !!(p && p.supplier && typeof p.supplier === "object" &&
  Array.isArray(p.items) && p.items.length > 0 && Array.isArray(p.items[0].rounds));

const migratePO = (p) => {
  if (!p) return p;
  if (isNewPO(p)) return { creditDays: DEFAULT_CREDIT_DAYS, ...p };
  // --- upgrade a legacy record ---
  const legacySuppliers = (p.suppliers && p.suppliers.length)
    ? p.suppliers
    : [{ name: (typeof p.supplier === "string" ? p.supplier : "") || "", poNumber: p.poNumber || "",
         rounds: (p.deliveries && p.deliveries.length
                   ? p.deliveries
                   : (p.incomingPlan || p.actualReceived ? [{ plan:p.incomingPlan||"", actual:p.actualReceived||"" }] : []))
                 .map(d => ({ amount:"", plan:d.plan||"", actual:d.actual||"" })) }];
  const first = legacySuppliers[0] || { name:"", poNumber:"", rounds:[] };
  const supplier = { name: first.name || (typeof p.supplier === "string" ? p.supplier : "") || "", poNumber: first.poNumber || p.poNumber || "" };
  const paymentType = p.paymentType === "credit30" ? "credit" : (p.paymentType || "");
  const creditDays  = p.creditDays || (p.paymentType === "credit30" ? 30 : DEFAULT_CREDIT_DAYS);
  const legacyRounds = legacySuppliers.flatMap(s => s.rounds || []);
  // Treat old data as fully received if it was ever marked delivered/paid or
  // carried an actual date, so received totals don't suddenly read as zero.
  const wasReceived = p.status === "Delivered" || p.status === "Paid" || legacyRounds.some(r => r.actual);
  const recvDate = legacyRounds.map(r=>r.actual).filter(Boolean).sort()[0] || p.date || "";
  const planDate = legacyRounds.map(r=>r.plan).filter(Boolean).sort()[0] || p.incomingPlan || "";
  const legacyItems = (p.items && p.items.length) ? p.items : [{ id:"legacy", code:p.code||"", amount:p.amount||"" }];
  const items = legacyItems.map(it => ({
    id: it.id && it.id !== "legacy" ? it.id : uid(),
    code: it.code || "", store: it.store || "", amount: it.amount || "",
    rounds: [{
      id: uid(), planDate, planAmount: it.amount || "",
      actualAmount: wasReceived ? (it.amount || "") : "",
      actualDate:  wasReceived ? recvDate : "",
    }],
  }));
  return { ...p, supplier, paymentType, creditDays, items };
};

const poItems = (p) => migratePO(p).items;
const poTotal = (p) => poItems(p).reduce((s,it) => s + (parseFloat(it.amount)||0), 0);
const poAmountForCode = (p, code) => poItems(p).filter(it => it.code===code).reduce((s,it) => s + (parseFloat(it.amount)||0), 0);

// ─── Supplier helpers (now exactly one supplier per PO) ─────────────────────
const poSupplier      = (p) => migratePO(p).supplier || { name:"", poNumber:"" };
const poSupplierName   = (p) => poSupplier(p).name || "—";
const poSupplierText   = (p) => poSupplier(p).name || "";
const poSupplierLabel  = (p) => poSupplier(p).name || "—";
const poNumbersLabel   = (p) => poSupplier(p).poNumber || "—";
const itemSupplierName = (p) => poSupplierName(p);
// Back-compat: some views still map over a suppliers[] array. There's now
// always exactly one supplier, so return it as a single-element list.
const poSuppliers = (p) => { const s = poSupplier(p); return [{ id:"main", name:s.name, poNumber:s.poNumber, rounds:[] }]; };

// Every round across every item, tagged with its item code. `plan`/`actual`/
// `amount` aliases are kept so older readers (tracking tab, exports) still work.
const poRounds = (p) => migratePO(p).items.flatMap(it =>
  (it.rounds && it.rounds.length ? it.rounds : []).map(r => ({
    ...r,
    plan: r.planDate, actual: r.actualDate, amount: r.planAmount,
    itemId: it.id, code: it.code,
  })));
const poDeliveries = poRounds;

// ─── Auto-pay: a round is paid once its due date has arrived ────────────────
// Cash pays on the received date; credit adds the PO's credit term (in days).
const roundPayDate  = (p, r) => {
  const P = migratePO(p);
  if (!r.actualDate) return "";
  if (P.paymentType === "cash") return r.actualDate;
  const d = parseInt(P.creditDays,10);
  return addDays(r.actualDate, isNaN(d) ? DEFAULT_CREDIT_DAYS : d);
};
// Received once the actual date has really arrived (a future date typed ahead
// of time doesn't count yet) and a quantity was recorded.
const roundReceived = (r) => !!r.actualDate && r.actualDate <= todayStr() && (parseFloat(r.actualAmount)||0) > 0;
// ตั้งสถานะ PO เป็น "Paid" เอง = ถือว่าจ่ายครบทุกงวดทันที (ไม่ต้องรอวันครบกำหนดเครดิต)
const roundPaid     = (p, r) => { if (migratePO(p).status === "Paid") return true; const d = roundPayDate(p,r); return !!d && d <= todayStr(); };
const itemOrdered   = (it) => parseFloat(it.amount)||0;
const itemReceived  = (it) => (it.rounds||[]).filter(roundReceived).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
const itemEntered   = (it) => (it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
// ปัดเป็นสตางค์ (2 ตำแหน่ง) กันเศษย่อยกว่าสตางค์ทำให้ "เหลือรับ 0.00" แต่ระบบยังคิดว่าไม่ครบ
const itemRemaining = (it) => Math.max(Math.round((itemOrdered(it) - itemEntered(it)) * 100) / 100, 0);

// ─── Edit history / audit log ──────────────────────────────────────────────
// Every PO keeps a short log of who changed what and when, so procurement
// can update a status in one click and everyone can still see the trail
// later (e.g. "ใครเปลี่ยนเป็น Delivered เมื่อไหร่"). Capped at 40 entries per
// PO so it never grows unbounded.
const HISTORY_ICON = { created:"🆕", status:"🔄", edited:"✏️" };
const historyEntry = (session, action, message) => ({
  id: uid(), at: new Date().toISOString(),
  user: session?.name || "—", role: session?.role ? (ROLE_LABELS[session.role] || session.role) : "",
  action, message,
});
const poHistory     = (p) => (p.history && p.history.length) ? p.history.slice().sort((a,b)=>(b.at||"").localeCompare(a.at||"")) : [];
const poLastUpdate  = (p) => poHistory(p)[0] || null;
const withHistory   = (po, entry) => ({ ...po, history: [entry, ...(po.history||[])].slice(0,40) });

// "2026-07-31T09:12:00Z" -> "2 วันที่แล้ว" — short, glanceable, always in Thai.
const relativeTime = (iso) => {
  if (!iso) return "";
  const diffMs = Date.now() - new Date(iso).getTime();
  const min = Math.floor(diffMs/60000);
  if (min < 1)  return t("เมื่อสักครู่","just now");
  if (min < 60) return `${min} ${t("นาทีที่แล้ว","min ago")}`;
  const hr = Math.floor(min/60);
  if (hr < 24)  return `${hr} ${t("ชม.ที่แล้ว","hr ago")}`;
  const day = Math.floor(hr/24);
  if (day < 30) return `${day} ${t("วันที่แล้ว","days ago")}`;
  return new Date(iso).toLocaleDateString(_LANG==="en"?"en-US":"th-TH",{day:"numeric",month:"short",year:"2-digit"});
};
const uiLocale = () => (_LANG === "en" ? "en-US" : "th-TH");   // วันที่/เวลาบนหน้าจอให้ตามภาษาที่เลือก
const formatDateTime = (iso) => iso ? new Date(iso).toLocaleString(uiLocale(),{day:"numeric",month:"short",year:"2-digit",hour:"2-digit",minute:"2-digit"}) : "—";

// ─── Actual received / paid dates ──────────────────────────────────────────
// Every round already carries its own "received" date; a PO's overall
// received date(s) are just the distinct actual dates across every round.
const poReceivedDates = (p) => [...new Set(poRounds(p).filter(roundReceived).map(r=>r.actualDate).filter(Boolean))].sort();
// Paid date = วันจ่ายที่กำหนดเอง (ถ้าตั้งสถานะ Paid เอง) มิฉะนั้นใช้วันครบกำหนดอัตโนมัติ
const poPaidDate = (p) => {
  const P = migratePO(p);
  if (P.status === "Paid" && P.paidDate) return P.paidDate;
  const paid = poRounds(P).filter(r=>roundPaid(P,r)).map(r=>roundPayDate(P,r)).filter(Boolean).sort();
  return paid.length ? paid[paid.length-1] : null;
};
// วันครบกำหนดจ่าย "แบบพยากรณ์" ของงวด — ใช้วันรับจริงถ้ามี ไม่มีก็ใช้วันแผน (+เทอมเครดิต)
// ต่างจาก roundPayDate (ที่ใช้วันจริงเท่านั้น เพื่อคุมตรรกะ "จ่ายแล้ว") — อันนี้ใช้ "แสดงผล"
// เท่านั้น (หน้าติดตาม/Export) ให้ตรงกับหน้าแผนจ่ายเงินที่คิดจากวันแผนด้วย
const roundDueForecast = (p, r) => {
  const P = migratePO(p);
  const incoming = r.actualDate || r.planDate;
  if (!incoming) return "";
  if (P.paymentType === "cash") return incoming;
  const d = parseInt(P.creditDays,10);
  return addDays(incoming, isNaN(d) ? DEFAULT_CREDIT_DAYS : d);
};
// Earliest upcoming/known payment due date across all rounds (for list/export display).
// Falls back to the plan date so a PO that only has a plan (no actual receipt yet)
// still shows its forecast due date, matching the Payment-plan page.
const poNextDueDate = (p) => {
  const due = poRounds(p).map(r=>roundDueForecast(p,r)).filter(Boolean).sort();
  return due.length ? due[0] : "";
};
// วันที่แบบอ่านง่ายตามภาษา: "18 ก.ย. 69" / "18 Sep 26" — ใช้แทน 2026-09-18 ในตาราง/ป๊อปอัพ
// (สร้างจากเวลาท้องถิ่น กันวันเลื่อน; ค่าที่ไม่ใช่วันที่คืนตามเดิม)
const fmtDate = (iso) => {
  if (!iso) return "";
  const [y, m, d] = String(iso).slice(0, 10).split("-").map(Number);
  if (!y || !m || !d) return String(iso);
  // เว้นวรรคแบบไม่ตัดบรรทัด — วันที่ไม่แตกเป็นหลายบรรทัดในคอลัมน์แคบ (มือถือ)
  return new Date(y, m - 1, d).toLocaleDateString(_LANG === "en" ? "en-GB" : "th-TH", { day: "numeric", month: "short", year: "2-digit" }).replace(/ /g, "\u00A0");
};
// ช่อง "วันจ่าย": จ่ายแล้ว = วันที่จ่าย (เขียว) · ยังไม่จ่าย = "ครบ <วันครบกำหนด>" (เดิมโชว์ "—" ทั้งที่รู้วันครบกำหนดแล้ว)
function PayDateText({ po }) {
  const paid = poPaidDate(po);
  if (paid) return <span style={{ color: T.green, fontWeight: 600 }}>{fmtDate(paid)}</span>;
  const due = poNextDueDate(po);
  if (!due) return <span style={{ color: T.textMuted }}>—</span>;
  return <span style={{ color: T.textSecondary }} title={t("วันครบกำหนดจ่าย (ยังไม่จ่าย)", "Payment due date (not paid yet)")}>{t("ครบ", "due")} {fmtDate(due)}</span>;
}

// ─── แผนจ่ายเงิน (Payment forecast lines) ───────────────────────────────────
// คืน "งวดจ่าย" ของ PO หนึ่งใบ สำหรับหน้าแผนจ่าย/Export. ปกติแตกตามงวดส่งของ
// (แต่ละงวดมียอด planAmount/actualAmount ของตัวเอง ซึ่งควรรวมกัน = ยอด item)
// แต่ถ้ายอดรวมของงวดไม่ตรงกับยอด item (เช่นมีงวดซ้ำยอดเต็ม) จะยุบเหลือ "หนึ่ง
// บรรทัดต่อ item" โดยยึดยอด item.amount เป็นหลัก เพื่อกันการนับซ้ำ. วันครบกำหนด
// จ่าย = วันรับของ (จริงถ้ามี ไม่มีใช้วันแผน) + เทอมเครดิต; เงินสดจ่ายวันรับของ.
const poPayLines = (p) => {
  const P = migratePO(p);
  // ตั้งสถานะ "Paid" เอง (มีหน้าต่างให้กรอกวันจ่าย) = จ่ายครบทั้งใบแล้ว แม้ของยังไม่เข้า (เช่นจ่ายล่วงหน้า)
  // → แผนจ่ายต้องไม่ขอให้จ่ายซ้ำ (เดิมยังโผล่ในรายการ "ต้องจ่าย" และยอด "ครบกำหนดเดือนนี้")
  const manualPaid = P.status === "Paid";
  const isCash = P.paymentType === "cash";
  const _cd = parseInt(P.creditDays,10);
  const term = isCash ? 0 : (isNaN(_cd) ? DEFAULT_CREDIT_DAYS : _cd); // นับ 0 วันเป็น 0 จริง (ให้ตรงกับ roundPayDate) — เฉพาะค่าว่าง/NaN ถึงใช้ค่าเริ่มต้น
  const method = isCash ? "เงินสด" : `เครดิต ${term} วัน`;
  const dueOf = (incoming) => incoming ? (isCash ? incoming : addDays(incoming, term)) : "";
  const roundAmt = (r) => (parseFloat(r.actualAmount)||0) || (parseFloat(r.planAmount)||0);
  const out = [];
  poItems(P).forEach((it, itemIdx) => {
    const itemAmt = parseFloat(it.amount)||0;
    const rounds  = it.rounds || [];
    const roundSum = rounds.reduce((s,r)=>s+roundAmt(r), 0);
    // งวดกระทบยอดตรงกับ item → เชื่อถือได้ ให้แตกเป็นรายงวดจริง
    const reconciled = rounds.length>0 && itemAmt>0 && Math.abs(roundSum - itemAmt) <= 0.5;
    if (reconciled) {
      rounds.forEach(r => {
        const amount = roundAmt(r);
        if (amount <= 0) return;
        const incoming = r.actualDate || r.planDate || "";
        // "จ่ายแล้ว" ต้อง (1) รับของจริงแล้ว และ (2) ถึงวันครบกำหนดจ่าย — ถ้ามีแต่
        // วันรับของแต่ยังไม่กรอกจำนวนที่รับจริง ถือว่ายังไม่รับ = ยังไม่จ่าย
        const received = roundReceived(r);
        const paid = manualPaid || (received && roundPaid(P, r));
        out.push({ itemIdx, code: it.code||"", incoming, incomingType: r.actualDate?"จริง":(r.planDate?"แผน":""),
          payDate: dueOf(incoming), amount, received, paid, paidAmount: paid ? amount : 0 });
      });
    } else if (itemAmt > 0) {
      // ยอดงวดไม่ตรง (หรือไม่มีงวด) → ยุบเหลือบรรทัดเดียว ใช้ยอด item เป็นหลัก
      // จ่ายบางส่วน: เก็บ paidAmount ไว้ให้ยอด "คงเหลือต้องจ่าย" หักออกถูกต้อง
      const actualDates = rounds.map(r=>r.actualDate).filter(Boolean).sort();
      const planDates   = rounds.map(r=>r.planDate).filter(Boolean).sort();
      const incoming = actualDates[0] || planDates[0] || "";
      const paidAmt  = manualPaid ? itemAmt : Math.min(rounds.filter(r=>roundReceived(r) && roundPaid(P,r)).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0), itemAmt);
      out.push({ itemIdx, code: it.code||"", incoming, incomingType: actualDates.length?"จริง":(planDates.length?"แผน":""),
        payDate: dueOf(incoming), amount: itemAmt, received: rounds.some(roundReceived), paid: paidAmt >= itemAmt-0.5, paidAmount: paidAmt });
    }
  });
  const today = todayStr();
  return out.map(l => ({
    ...l, isCash, method, term,
    supplier: poSupplierName(P), poNo: poNumbersLabel(P),
    accName: accountOf(l.code)?.name || "",
    month: l.payDate ? l.payDate.slice(0,7) : "",
    // "เกินกำหนดจ่าย" เฉพาะของที่รับแล้วแต่ยังไม่จ่ายและเลยกำหนด; ของที่ยังไม่รับ = "รอจ่าย"
    status: l.paid ? "paid" : (l.received && l.payDate && l.payDate < today ? "late" : "pending"),
  }));
};

// ยอด "จ่ายแล้ว" ของ PO — ใช้ตัวเดียวกันทุกที่ (การ์ดจัดซื้อ/บัญชี · แผนจ่าย · Excel · ป๊อปอัพ PO)
// = ผลรวม paidAmount ของแผนจ่าย: ไม่เกินยอดสั่งของแต่ละรายการ และสถานะ Paid เอง = จ่ายครบ
// (เดิมการ์ดรวม actualAmount ของทุกงวดที่ถึงกำหนด → งวดซ้ำ/กรอกเกินทำให้ "ชำระแล้ว" เกินยอด PO
//  และ PO ที่ตั้ง Paid เองแต่ยังไม่กรอกของเข้า นับเป็น 0 ขณะที่ป้ายบอกว่าจ่ายแล้ว)
const poPaidAmount = (p) => poPayLines(p).reduce((s, l) => s + (l.paidAmount || 0), 0);
// ต่อรายการ (ตามลำดับใน PO — PO แบบเก่าได้ id ใหม่ทุกครั้งที่อ่าน จึงอ้างด้วยลำดับ)
const itemPaidAmount = (p, idx) => poPayLines(p).filter(l => l.itemIdx === idx).reduce((s, l) => s + (l.paidAmount || 0), 0);

// ยอด "วางบิลแล้ว" ของ PO — ทีมไม่ได้เปลี่ยนสถานะเป็น "วางบิลแล้ว" ทีละใบ (จ่ายเงินคิดอัตโนมัติจากวันของเข้า)
// จึงนับจาก "ของที่รับแล้ว" (ถึงรอบวางบิล) ไม่เกินยอดสั่งของแต่ละรายการ · ถ้าตั้งสถานะ วางบิลแล้ว/จ่ายแล้ว เอง = ทั้งใบ
// และไม่น้อยกว่ายอดที่จ่ายแล้ว (จ่ายแล้วย่อมวางบิลแล้ว) — เดิมนับเฉพาะสถานะ จึงขึ้น 0 ทั้งที่จ่ายไปแล้วหลายสิบล้าน
const poBilledAmount = (p) => {
  const P = migratePO(p);
  if (P.status === "Invoiced" || P.status === "Paid") return poTotal(P);
  const recv = poItems(P).reduce((s, it) => s + Math.min(itemReceived(it), itemOrdered(it)), 0);
  return Math.max(recv, poPaidAmount(P));
};

// ─── Lock completed POs ─────────────────────────────────────────────────────
// Once a PO has been fully received AND fully paid, its numbers are final —
// only an admin can still edit or delete it, so the paper trail for a closed
// PO can't quietly change after the fact.
const isPOLocked = (p) => incomingStatus(p)==="received" && paymentStatus(p)==="paid";
const canEditPO  = (p, session) => !isPOLocked(p) || session?.role==="admin";

const deliveryStatus = (d) => {
  // Works on a round object (planDate/actualDate) or its aliases (plan/actual).
  const plan = d.planDate ?? d.plan, actual = d.actualDate ?? d.actual;
  if (actual) return actual <= todayStr() ? "received" : "pending";
  if (plan && plan < todayStr()) return "late";
  if (plan) return "pending";
  return "unset";
};
// PO-level incoming status by value received across all item rounds: fully
// received once received ≥ ordered; "partial" once some (but not all) is in.
const incomingStatus = (p) => {
  const rounds = poRounds(p);
  if (!rounds.length) return "unset";
  // เช็ค "รับครบ" ต่อ item — ถ้ารับเกินใน item หนึ่งจะได้ไม่ไปกลบ item ที่ยังรับไม่ครบ
  // (ก่อนหน้านี้เทียบยอดรวมกับ poTotal จึงล็อก PO เร็วเกินจริง)
  const items = poItems(p).filter(it => itemOrdered(it) > 0);
  const allReceived = items.length > 0 && items.every(it => Math.round(itemReceived(it)*100) >= Math.round(itemOrdered(it)*100)); // เทียบระดับสตางค์ ให้ตรงกับที่แสดง
  const anyReceived = rounds.some(roundReceived);
  // "ล่าช้า" เฉพาะงวดที่ยังไม่รับ และ "ยังไม่ได้ใส่วันรับ" และเลยวันแผนแล้ว
  // (ถ้าใส่วันรับไว้ล่วงหน้า = นัดไว้แล้ว ยังไม่ถือว่าล่าช้าจนกว่าจะเลยวันรับ)
  const anyLate  = rounds.some(r => !roundReceived(r) && !r.actualDate && r.planDate && r.planDate < todayStr());
  if (allReceived) return "received";
  if (anyLate) return "late";
  if (anyReceived) return "partial";
  // มีนัด (วันแผน) หรือมีวันรับล่วงหน้าที่ยังไม่ถึง = "รอของเข้า" (ให้ตรงกับป้ายระดับงวด)
  if (rounds.some(r=>r.planDate || r.actualDate)) return "pending";
  return "unset";
};
// Auto-pay: reaching a round's due date is what marks it paid, so payment is
// never "late" — it's "pending" until the due date, then "paid".
const paymentStatus = (p) => {
  if (migratePO(p).status === "Paid" && poTotal(p) > 0) return "paid";   // ตั้ง Paid เอง = จ่ายครบ (ให้ตรงกับแผนจ่าย/การ์ด)
  const rounds = poRounds(p);
  const recvRounds = rounds.filter(roundReceived);
  if (!recvRounds.length) return "unset";
  // จ่ายครบต่อ item: งวดที่ทั้งรับแล้วและถึงกำหนดจ่าย ต้องครอบคลุมยอด item ทุก item
  // (กันไม่ให้ "จ่ายแล้ว" เกิดขึ้นทั้งที่บาง item ยังจ่ายไม่ครบ)
  const items = poItems(p).filter(it => itemOrdered(it) > 0);
  const itemPaid = (it) => (it.rounds||[]).filter(r => roundReceived(r) && roundPaid(p,r)).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
  const allPaid = items.length > 0 && items.every(it => Math.round(itemPaid(it)*100) >= Math.round(itemOrdered(it)*100)); // เทียบระดับสตางค์
  if (allPaid) return "paid";
  return "pending";
};
const INCOMING_LABEL = { received:"รับแล้ว", partial:"รับบางส่วน", late:"ของเข้าล่าช้า", pending:"รอของเข้า", unset:"ยังไม่กำหนด" };
const INCOMING_CLR   = { received:"#10b981", partial:"#3b82f6", late:"#ef4444", pending:"#f59e0b", unset:"#94a3b8" };
const INCOMING_BG    = { received:"#f0fdf4", partial:"#eff6ff", late:"#fef2f2", pending:"#fffbeb", unset:"#f1f5f9" };
const PAYMENT_LABEL  = { paid:"จ่ายแล้ว", late:"เกินกำหนดจ่าย", pending:"รอจ่ายเงิน", unset:"ยังไม่กำหนด" };
const PAYMENT_CLR    = { paid:"#10b981", late:"#ef4444", pending:"#f59e0b", unset:"#94a3b8" };
const PAYMENT_BG     = { paid:"#f0fdf4", late:"#fef2f2", pending:"#fffbeb", unset:"#f1f5f9" };
// Payment method — cash pays right away, credit gives suppliers a 30-day term,
// so a credit PO's payment due date is auto-suggested as order date + 30 days.
const PAYMENT_TYPE_CLR   = { cash:"#10b981", credit:"#2563eb", credit30:"#2563eb" };
const PAYMENT_TYPE_BG    = { cash:"#f0fdf4", credit:"#eff6ff", credit30:"#eff6ff" };
// Label for a PO's payment method including its credit term, e.g. "เครดิต 45 วัน".
const creditTermDays = (P) => { const cd = parseInt(P.creditDays,10); return isNaN(cd) ? DEFAULT_CREDIT_DAYS : cd; }; // เทอมเครดิตแบบสอดคล้องกันทุกที่ (นับ 0 เป็น 0, เฉพาะว่าง/NaN ใช้ค่าเริ่มต้น)
// Bilingual runtime labels for the on-screen UI (the *_LABEL maps above stay Thai
// for Excel exports). These call t() at render time so they follow the toggle.
const INCOMING_LABEL_EN = { received:"Received", partial:"Partial", late:"Late", pending:"Awaiting", unset:"Unset" };
const PAYMENT_LABEL_EN  = { paid:"Paid", late:"Overdue", pending:"Awaiting pay", unset:"Unset" };
const incLabel = (k) => t(INCOMING_LABEL[k]||k, INCOMING_LABEL_EN[k]||k);
const payLabel = (k) => t(PAYMENT_LABEL[k]||k, PAYMENT_LABEL_EN[k]||k);
const payTypeLabelT = (p) => { const P = migratePO(p); if (P.paymentType==="cash") return t("เงินสด","Cash"); if (P.paymentType==="credit") { const n = creditTermDays(P); return t(`เครดิต ${n} วัน`,`Credit ${n}d`); } return "—"; };
const fmt  = n => new Intl.NumberFormat("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2}).format(n||0);
// ในตารางอ่านอย่างเดียว: ค่า 0 แสดงเป็น "–" (ลดตัวเลขรก ช่องที่มียอดจริงจะเด่นขึ้น)
const fmtZ = n => (Math.abs(parseFloat(n) || 0) < 0.005 ? "–" : fmt(n));
// บาทเต็ม (ไม่มีทศนิยม) — ใช้กับตัวเลขพาดหัวการ์ด/ยอดรวม ให้กวาดตาอ่านง่าย
const fmt0 = n => new Intl.NumberFormat("th-TH",{maximumFractionDigits:0}).format(Math.round(n||0));
// รองรับค่าติดลบ (งานลด) — เดิม -12,000 แสดงเป็น "-12000" แทน "-12K"
const fmtK = n => { const a = Math.abs(n||0), sg = (n||0) < 0 ? "-" : ""; return sg + (a>=1e6?`${(a/1e6).toFixed(1)}M`:a>=1e3?`${(a/1e3).toFixed(0)}K`:Math.round(a).toString()); };
// "2026-08" -> "ส.ค. 69" (TH, Buddhist year) or "Aug 25" (EN, Gregorian) — used
// wherever a month key needs a short label (QS Monthly tab's chips/headers, chart
// X-axis, and sub-item "เพิ่มเมื่อ ..." badges). Follows the current UI language.
// สร้างวันที่แบบเวลาท้องถิ่น (เดิม new Date("2026-08-01") = เที่ยงคืน UTC → ผู้ใช้ที่อยู่ฝั่งตะวันตกของ UTC เห็นเป็นเดือนก่อนหน้า)
const monthShortLabel = (m) => { const [y, mo] = String(m).split("-").map(Number); return new Date(y, (mo||1)-1, 1).toLocaleDateString(_LANG==="en"?"en-US":"th-TH",{month:"short",year:"2-digit"}); };
const uid  = () => Math.random().toString(36).slice(2,10);

// ─── Design Tokens ────────────────────────────────────────────────────────────
const T = {
  // Layout
  bg:        "#f0f4f8",
  sidebar:   "#1e293b",
  card:      "#ffffff",
  cardBorder:"#e2e8f0",
  // Text
  textPrimary:  "#0f172a",
  textSecondary:"#475569",   // ≥ 7:1 บนพื้นขาว
  textMuted:    "#5f6b7e",   // ≥ 4.5:1 (WCAG AA) ทั้งบนพื้นขาวและพื้นหน้า — เดิม #94a3b8 = 2.56:1 อ่านยาก
  // Brand blue
  blue:     "#2563eb",
  blueDark: "#1d4ed8",
  blueLight:"#eff6ff",
  blueMid:  "#dbeafe",
  // Accent
  green:    "#10b981",
  greenBg:  "#f0fdf4",
  amber:    "#f59e0b",
  amberBg:  "#fffbeb",
  purple:   "#8b5cf6",
  purpleBg: "#f5f3ff",
  red:      "#ef4444",
  redBg:    "#fef2f2",
  // Header gradient
  headerGrad: "linear-gradient(135deg, #1e40af 0%, #2563eb 50%, #3b82f6 100%)",
};

// ─── ภาษา (2 ภาษา: ไทย / English) — สลับได้ทั้งแอป ─────────────────────────────
//  ใช้ t("ไทย", "English") ทุกที่ที่แสดงข้อความ · เก็บภาษาที่เลือกไว้ใน localStorage
//  toggleLang() แจ้งทุก component ที่ subscribe (useLang) ให้ re-render ทันที
let _LANG = "th";
try { const s = localStorage.getItem("tcs-lang"); if (s === "en" || s === "th") _LANG = s; } catch { /* ignore */ }
const _langSubs = new Set();
const t = (th, en) => (_LANG === "en" ? (en ?? th) : th);
// ไฟล์ Excel ออกแบบเป็นภาษาไทยเสมอ — ระหว่างสร้างไฟล์ให้ตัวช่วยทุกตัว (ชื่อเดือน, วันที่, สถานะ PO, t())
// ทำงานเป็นไทย แล้วคืนภาษาเดิม (การสร้างไฟล์เป็น synchronous จึงไม่มีการวาดหน้าจอแทรกระหว่างนี้)
const inThai = (fn) => { const prev = _LANG; _LANG = "th"; try { return fn(); } finally { _LANG = prev; } };
const setLang = (l) => {
  if (l !== "en" && l !== "th") return;
  _LANG = l;
  try { localStorage.setItem("tcs-lang", l); } catch { /* ignore */ }
  _langSubs.forEach(fn => { try { fn(l); } catch { /* ignore */ } });
};
const toggleLang = () => setLang(_LANG === "th" ? "en" : "th");
// subscribe a component to language changes (re-render on toggle)
function useLang() {
  const [, force] = useState(0);
  useEffect(() => {
    const fn = () => force(x => x + 1);
    _langSubs.add(fn);
    return () => { _langSubs.delete(fn); };
  }, []);
  return _LANG;
}

// ─── Global CSS ───────────────────────────────────────────────────────────────
const FAB_SIZE = 54, FAB_GAP = 20;   // ปุ่มลอยมุมขวาล่าง (ใช้ใน CSS ด้านล่างด้วย จึงต้องประกาศก่อน)
const BNAV_H = 62;                    // แถบแท็บด้านล่าง (มือถือ) — ของลอยมุมล่างทั้งหมดยกขึ้นตามนี้
const bnavH = () => (typeof document !== "undefined" && document.body && document.body.classList.contains("has-bnav")) ? BNAV_H : 0;
const BOTTOM = (px) => `calc(var(--bnav, 0px) + var(--bnav-safe, 0px) + ${px}px)`;
const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300..800&family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', 'IBM Plex Sans Thai', 'Leelawadee UI', Tahoma, sans-serif; background: ${T.bg}; color: ${T.textPrimary}; font-variant-numeric: tabular-nums; }
  input, select, textarea, button { font-family: inherit; font-variant-numeric: tabular-nums; }
  input[type=number]::-webkit-inner-spin-button { opacity: 0.4; }
  ::-webkit-scrollbar { width: 16px; height: 16px; }
  ::-webkit-scrollbar-track { background: #eef2f7; border-radius: 10px; }
  ::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 10px; border: 3px solid #eef2f7; min-height: 48px; min-width: 48px; }
  ::-webkit-scrollbar-thumb:hover { background: #64748b; }
  * { scrollbar-color: #94a3b8 #eef2f7; scrollbar-width: auto; }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
  @keyframes fadeIn { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:translateY(0)} }
  .calc-pop button:active { transform: scale(0.96); }
  .fab-btn:hover { background: #1d4ed8 !important; }
  .fab-btn:active { transform: scale(0.94) !important; }
  @media print { .fab-btn, .calc-pop, [data-table-top] { display: none !important; } body.has-fab { padding: 0 !important; } }
  /* แถบหัวบนจอเล็ก: เหลือบรรทัดเดียว (ซ่อนข้อความรอง) */
  @media (max-width: 760px) {
    .shell-bar { padding: 8px 12px !important; gap: 8px !important; min-height: 52px !important; }
    .shell-bar .hdr-meta, .shell-bar .crumb-txt, .shell-bar .um-name, .shell-bar .crumb-sep, .shell-bar .hdr-lbl { display: none !important; }
  }
  @media (max-width: 760px) { .shell-bar .sync-badge { font-size: 0 !important; gap: 0 !important; } }
  /* เว้นพื้นที่ให้ปุ่มลอยมุมขวาล่าง — ไม่ให้ทับตาราง: จอกว้าง = แถบว่างด้านขวา · จอแคบ = เว้นท้ายหน้า */
  @media (min-width: 1024px) {
    body.has-fab { padding-right: ${FAB_SIZE + FAB_GAP * 2}px; }
    /* แถบหัวสียังยาวเต็มจอ (ยื่นเข้าไปในแถบว่าง) แต่ปุ่มในแถบหัวอยู่แนวเดียวกับเนื้อหา */
    body.has-fab .app-header { margin-right: -${FAB_SIZE + FAB_GAP * 2}px; padding-right: ${FAB_SIZE + FAB_GAP * 2 + 30}px !important; }
  }
  @media (max-width: 1023.98px) { body.has-fab { padding-bottom: ${(FAB_SIZE + 12) * 2 + FAB_GAP + 8}px; } }
  body.has-bnav { --bnav: ${BNAV_H}px; --bnav-safe: env(safe-area-inset-bottom, 0px); }
  body.has-fab.has-bnav { padding-bottom: calc(${(FAB_SIZE + 12) * 2 + FAB_GAP + 8 + BNAV_H}px + env(safe-area-inset-bottom, 0px)); }
  @media print { [data-bottom-nav] { display: none !important; } }
  .card-hover { transition: box-shadow 0.18s, transform 0.18s; }
  .icon-danger:hover { color: ${T.red} !important; background: ${T.redBg} !important; }
  .card-hover:hover { box-shadow: 0 8px 24px rgba(37,99,235,0.12); transform: translateY(-2px); }
  .btn-primary { background: ${T.blue}; color: #fff; border: none; border-radius: 10px; padding: 10px 22px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.15s, box-shadow 0.15s; }
  .btn-primary:hover { background: ${T.blueDark}; box-shadow: 0 4px 12px rgba(37,99,235,0.3); }
  .btn-ghost { background: transparent; color: ${T.textSecondary}; border: 1.5px solid ${T.cardBorder}; border-radius: 10px; padding: 9px 18px; font-size: 13px; font-weight: 500; cursor: pointer; transition: border-color 0.15s, color 0.15s; }
  .btn-ghost:hover { border-color: ${T.blue}; color: ${T.blue}; }
  /* แท็บแบบแถบเลือก (segmented): พื้นเทา แท็บที่เลือกเป็นการ์ดขาว — สีเรียบเหมือนกันทุกแผนก */
  /* การ์ดสรุปที่มีตัวหลัก: ตัวแรกกว้างกว่า (จอกว้าง) · เต็มแถว (มือถือ) */
  .stat-grid.has-lead { grid-template-columns: 1.6fr 1fr 1fr !important; }
  @media (max-width: 900px) { .stat-grid.has-lead { grid-template-columns: 1fr 1fr !important; } .stat-grid.has-lead .stat-lead, .stat-grid.has-lead .stat-card:last-child:nth-child(even) { grid-column: 1 / -1; } }
  /* ฟอร์ม PO บนมือถือ: ช่องหลักเรียงทีละช่อง · ช่องตัวเลขของแต่ละรายการ 2 คอลัมน์ (ช่องสุดท้ายเต็มแถว) */
  @media (max-width: 600px) {
    .po-form-grid { grid-template-columns: 1fr !important; }
    .po-item-grid { grid-template-columns: 1fr 1fr !important; }
    .po-item-grid > :last-child { grid-column: 1 / -1; }
  }
  .att-chip { display: inline-flex; align-items: center; gap: 6px; border: 1px solid #f5c2bd; background: #fff5f4; color: #b42318; border-radius: 999px; padding: 5px 12px; font-size: 13px; font-weight: 600; cursor: pointer; font-family: inherit; min-height: 32px; }
  .att-chip b { font-weight: 800; }
  .att-chip:hover { background: #feeae8; }
  .att-chip[aria-pressed="true"] { background: #b42318; color: #fff; border-color: #b42318; }
  .att-chip:focus-visible { outline: 2px solid #b42318; outline-offset: 1px; }
  .seg-tabs { display: inline-flex; gap: 2px; background: #e6ebf2; border-radius: 10px; padding: 3px; flex-wrap: wrap; }
  .seg-tab { background: transparent; border: none; border-radius: 8px; padding: 7px 16px; font-size: 13px; font-weight: 500; color: ${T.textSecondary}; cursor: pointer; white-space: nowrap; font-family: inherit; min-height: 34px; }
  .seg-tab:hover { color: ${T.textPrimary}; }
  .seg-tab.on { background: #fff; color: ${T.textPrimary}; font-weight: 650; box-shadow: 0 1px 2px rgba(15,23,42,0.08), 0 1px 6px rgba(15,23,42,0.06); }
  .seg-tab:focus-visible { outline: 2px solid ${T.blue}; outline-offset: 1px; }
  .date-th-text { font-size: 13px; }
  .po-row:hover { background: #fafbfd; }
  .po-grp:hover { background: #f1f5f9 !important; }
  .sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
  .input-base { background: ${T.bg}; border: 1.5px solid ${T.cardBorder}; border-radius: 10px; padding: 10px 13px; color: ${T.textPrimary}; font-size: 13px; outline: none; transition: border-color 0.15s, box-shadow 0.15s; width: 100%; }
  .input-base:focus { border-color: ${T.blue}; box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
  .tag { display: inline-flex; align-items: center; padding: 2px 9px; border-radius: 6px; font-size: 11px; font-weight: 600; }
  /* กล่องเลื่อนแนวนอน (ใช้กับตารางที่คอลัมน์เยอะ) — สกรอลบาร์เห็นชัดเสมอ */
  .hscroll { overflow-x: auto; overflow-y: hidden; }
  .hscroll::-webkit-scrollbar { height: 24px; }
  .hscroll::-webkit-scrollbar-track { background: #e2e8f0; border-radius: 12px; }
  .hscroll::-webkit-scrollbar-thumb { background: #64748b; border-radius: 12px; border: 4px solid #e2e8f0; min-width: 56px; }
  .hscroll::-webkit-scrollbar-thumb:hover { background: #475569; }
  .hscroll { scrollbar-color: #64748b #e2e8f0; scrollbar-width: auto; }
  /* ตารางรายเดือน: เลื่อนในกล่องเอง (สูงไม่เกิน 70vh) + ตรึงหัวตาราง + สกรอลบาร์เห็นชัด */
  .mscroll { overflow: auto; max-height: 70vh; }
  .mscroll::-webkit-scrollbar { height: 24px; width: 24px; }
  .mscroll::-webkit-scrollbar-track { background: #e2e8f0; border-radius: 12px; }
  .mscroll::-webkit-scrollbar-thumb { background: #64748b; border-radius: 12px; border: 4px solid #e2e8f0; min-width: 56px; min-height: 56px; }
  .mscroll::-webkit-scrollbar-thumb:hover { background: #475569; }
  .mscroll::-webkit-scrollbar-corner { background: #e2e8f0; }
  .mscroll { scrollbar-color: #64748b #e2e8f0; scrollbar-width: auto; }
  .mscroll thead th { position: sticky; background: #f8fafc; z-index: 4; box-shadow: inset 0 -1px 0 ${T.cardBorder}; }
  .mscroll thead tr:first-child th { top: 0; }
  .mscroll thead tr:nth-child(2) th { top: 33px; z-index: 4; }
  /* สกรอลบาร์แนวนอนแบบใหญ่ คลิก/ลากง่าย — ใช้กับตารางรายเดือน (กว้างมาก) */
  .fatscroll { overflow: auto; -webkit-overflow-scrolling: touch; scrollbar-color: #64748b #e2e8f0; scrollbar-width: auto; }
  .fatscroll::-webkit-scrollbar { height: 28px; width: 28px; }
  .fatscroll::-webkit-scrollbar-track { background: #dbe2ec; border-radius: 14px; }
  .fatscroll::-webkit-scrollbar-thumb { background: #556274; border-radius: 14px; border: 5px solid #dbe2ec; min-width: 64px; min-height: 64px; }
  .fatscroll::-webkit-scrollbar-thumb:hover { background: #3b4756; }
  .fatscroll::-webkit-scrollbar-corner { background: #e2e8f0; }
  /* เลื่อนลื่นบน iOS */
  .hscroll, .mscroll { -webkit-overflow-scrolling: touch; }
  /* ── มือถือ/จอแคบ: ปุ่มแตะง่ายขึ้น + ช่องกรอกไม่โดน iOS ซูมอัตโนมัติ (ต้อง ≥16px) ── */
  @media (max-width: 640px) {
    .btn-primary, .btn-ghost { min-height: 40px; padding-top: 10px; padding-bottom: 10px; }
    .input-base, .date-th-text { font-size: 16px !important; }
    .due-box { flex: 1 1 130px; min-width: 0 !important; }
    .chip-scroll { flex-wrap: nowrap !important; overflow-x: auto; flex-basis: 100% !important; padding-bottom: 2px; scrollbar-width: none; }
    .chip-scroll::-webkit-scrollbar { display: none; }
    .pay-h { flex-basis: 100%; }
    .hscroll::-webkit-scrollbar, .mscroll::-webkit-scrollbar { height: 18px; width: 18px; }
    .fatscroll::-webkit-scrollbar { height: 22px; width: 22px; }
    /* การ์ดสรุป: 2×2 แบบกะทัดรัด (เดิมเรียงลงมาทีละใบ ต้องเลื่อน ~600px กว่าจะถึงตาราง) */
    .stat-grid { grid-template-columns: 1fr 1fr !important; gap: 10px !important; margin-bottom: 14px !important; }
    .stat-card { padding: 12px !important; }
    .stat-card .stat-val { font-size: 16px !important; letter-spacing: -0.3px !important; word-break: break-all; }
    .stat-card .stat-icon { display: none !important; }
    /* หน้าต่างป๊อปอัพ PO: ช่องกรอกงวดเรียง 1 คอลัมน์ (เดิม 2 คอลัมน์แคบจนตัวเลขขาด) */
    .round-grid { grid-template-columns: 1fr !important; }
  }
`;

// ─── Excel Export ─────────────────────────────────────────────────────────────
// Every department gets its own styled workbook — xlsx-js-style (a SheetJS
// fork) lets us actually write cell colors/fonts/borders, which the plain
// community "xlsx" package silently drops on write.
// ยอดเพิ่มของ Acc. Code ในเดือน m = ค่าธรรมดา (code) ซึ่งคือ "ยอดรวมที่ roll-up
// ไว้แล้ว" ของเดือนนั้น (handleSave ตั้ง code = ผลรวมคอลัมน์ย่อย/รายการย่อยเสมอ)
// จึงอ่านตัวเดียว — ไม่บวกคีย์คอลัมน์ย่อย ":" ซ้ำ (กันนับซ้ำ) และคอลัมน์ที่ลบไป
// แล้วก็ไม่ถูกนับ เพราะยอด roll-up ถูกคำนวณใหม่โดยไม่รวมคอลัมน์นั้น
const monthAddValue = (additions, m, code) => parseFloat(additions?.[m]?.[code]) || 0;
// เผื่อเศษ/สูญเสีย: คิดจาก "ราคาเดิม" ของแต่ละ Acc. Code แล้วนับเป็นงบ (ทุน) จริง — บวกเข้าแต่ละรหัส
// ทำให้งบคงเหลือ / เกินงบ ของจัดซื้อและบัญชีรวมส่วนเผื่อนี้ด้วย (ตัดสินใจ 2026-09-26) · งานเพิ่มรายเดือนไม่ถูกบวก %
const WASTE_RATE = 0.03;
const WASTE_LBL  = "3%";
const wasteOf   = (v) => Math.round((parseFloat(v) || 0) * WASTE_RATE * 100) / 100;   // ปัดเป็นสตางค์ ให้จอ/Excel/เทียบ PO ตรงกัน
const withWaste = (v) => (parseFloat(v) || 0) + wasteOf(v);
// คำนวณ "ยอดแม่" ใหม่จากข้อมูลย่อย — ใช้หลังรวมการแก้จาก 2 เครื่อง (ยอดแม่เก็บเป็นค่าสรุป ถ้ารวมทีละช่อง
// ยอดแม่จะเป็นของเครื่องใดเครื่องหนึ่ง ไม่ตรงกับผลรวมลูก) · กติกาเดียวกับตอนกดบันทึกในหน้า QS
const rollupTenders = (tn, extra = []) => {
  if (!tn || typeof tn !== "object" || Array.isArray(tn)) return tn;
  const out = { ...tn }, kidsOf = {};
  (extra || []).forEach(e => { if (e && e.parentCode && !e.addedInMonth) (kidsOf[e.parentCode] = kidsOf[e.parentCode] || []).push(e.code); });
  Object.entries(kidsOf).forEach(([p, kids]) => {
    const v = kids.reduce((s, k) => s + (parseFloat(out[k]) || 0), 0);
    if (v > 0) out[p] = v; else delete out[p];
  });
  return out;
};
const rollupAdditions = (ad, extra = []) => {
  if (!ad || typeof ad !== "object" || Array.isArray(ad)) return ad;
  const out = { ...ad };
  const rows = [...ACCOUNTS.map(a => a.code), ...(extra || []).filter(e => e && !e.parentCode).map(e => e.code)];
  Object.keys(ad).forEach(m => {
    const mo0 = ad[m];
    if (m.startsWith("$") || !mo0 || typeof mo0 !== "object" || Array.isArray(mo0)) return;
    const mo = { ...mo0 }, cols = Array.isArray(mo.$columns) ? mo.$columns : Array.isArray(ad.$columns) ? ad.$columns : [];
    rows.forEach(code => {
      const kids = (extra || []).filter(e => e && e.parentCode === code && (!e.addedInMonth || e.addedInMonth <= m));
      let v;
      if (kids.length) v = kids.reduce((s, k) => s + (parseFloat(mo[k.code]) || 0), 0);
      else if (cols.length && cols.some(c => `${code}:${c.id}` in mo)) v = cols.reduce((s, c) => s + (parseFloat(mo[`${code}:${c.id}`]) || 0), 0);
      else return;
      if (v) mo[code] = v; else delete mo[code];
    });
    out[m] = mo;
  });
  return out;
};
const buildCombinedBudget = (tenderCosts, additions) => {
  const combined = {};
  Object.entries(tenderCosts || {}).forEach(([k, v]) => { combined[k] = k.startsWith("$") ? v : withWaste(v); });
  Object.entries(additions || {}).forEach(([mKey, monthObj]) => {
    if (mKey.startsWith("$")) return;
    Object.entries(monthObj || {}).forEach(([code, val]) => {
      // ข้ามคีย์ meta ($…) และคีย์คอลัมน์ย่อย (code:colId) — ค่าเหล่านี้ถูก roll-up
      // เข้าไปในค่าธรรมดา (code) แล้ว การบวกอีกจะนับซ้ำ
      if (code.startsWith("$") || code.includes(":")) return;
      combined[code] = (parseFloat(combined[code]) || 0) + (parseFloat(val) || 0);
    });
  });
  return combined;
};
// ยอดเพิ่มของ Acc. Code ในเดือน m แยกตามคอลัมน์ของเดือนนั้น สำหรับชีตรายเดือนใน Excel
// ยอดรวมยึด "ค่าหลัก" (roll-up ที่บันทึกไว้ = ตัวเดียวกับชีตสรุป/งบของทุกฝ่าย) — ส่วนที่ไม่อยู่ใน
// คอลัมน์ใด (เช่นแถวที่มีรายการย่อย ซึ่งรายการย่อยกรอกเป็นช่องเดียว) แยกไว้ใน other
// (เดิมอ่านเฉพาะ code:colId ทำให้แถวที่มีรายการย่อยได้ 0 แล้วหายจากชีต และ TOTAL ต่ำกว่าสรุป)
const monthRowBreakdown = (additions, m, code, cols) => {
  const obj = (additions && additions[m]) || {};
  const plainRaw = parseFloat(obj[code]);
  if (!cols.length) { const total = isNaN(plainRaw) ? 0 : plainRaw; return { vals: [total], other: 0, total }; }
  const vals = cols.map(c => parseFloat(obj[`${code}:${c.id}`]) || 0);
  const sumCols = vals.reduce((s, v) => s + v, 0);
  const total = isNaN(plainRaw) ? sumCols : plainRaw;
  const other = Math.round((total - sumCols) * 100) / 100;
  return { vals, other, total };
};
const OTHER_COL_LABEL = "อื่น ๆ / รายการย่อย";
const exportAccountList = (extraItems=[], hiddenAccounts=[], poEntries=null, plans=null) => {
  const list = [
    ...ACCOUNTS.filter(a => !hiddenAccounts.includes(a.code)),
    ...extraItems.filter(e => !e.parentCode).map(e => ({ code:e.code, name:e.name, group:e.group||"Other" })),
  ];
  if (!poEntries && !plans) return list;
  // รหัสที่มี PO/แผนผูกอยู่ แต่ไม่อยู่ในรายการแล้ว (เช่นลบรายการเพิ่มทิ้ง) — ต้องยังแสดงในหน้าบัญชี/Excel
  // ไม่งั้นยอด PO ของรหัสนั้นอยู่ในการ์ดรวม แต่ไม่มีแถวในตาราง → ผลรวมตารางไม่เท่าการ์ด
  const have = new Set(list.map(a => a.code)), orphans = [];
  [...(poEntries || []), ...(Array.isArray(plans) ? plans : [])].forEach(p => poItems(p).forEach(it => {
    if (!it.code || have.has(it.code) || !(parseFloat(it.amount) || 0)) return;
    have.add(it.code);
    const e = extraItems.find(x => x.code === it.code);
    orphans.push({ code: it.code, name: (e && e.name) || ORPHAN_NAME, group: "Other", orphan: true });
  }));
  return [...list, ...orphans];
};
const ORPHAN_NAME = "(รหัสนี้ไม่อยู่ในรายการแล้ว)";
// เซตของ Acc.Code ที่ยังมี PO ผูกอยู่ (ยอด ≠ 0)
const poCodeSet = (poEntries=[]) => { const s = new Set(); (poEntries||[]).forEach(p => poItems(p).forEach(it => { if (it.code && (parseFloat(it.amount)||0) !== 0) s.add(it.code); })); return s; };
// หน้าบัญชี/จัดซื้อ: ถ้า QS เผลอซ่อน Acc.Code ที่ยังมี PO อยู่ ต้องไม่ซ่อนในมุมมองเหล่านี้
// (ไม่งั้นยอด committed/งบของ code นั้นหายจากยอดรวมเงียบ ๆ) — คืนรายการ "ซ่อนได้จริง" คือที่ไม่มี PO
const hiddenSafeForPO = (hiddenAccounts=[], poEntries=[]) => { const withPO = poCodeSet(poEntries); return (hiddenAccounts||[]).filter(c => !withPO.has(c)); };
// ตัวเลขสรุปของโครงการสำหรับการ์ดหน้ารายการโครงการ — สูตรเดียวกับการ์ดสรุปหน้าบัญชี
// (งบรวม = เฉพาะรหัสระดับบน, ผูกพัน = มูลค่า PO ทั้งหมด, ต้องจ่ายเดือนนี้ = คงเหลือถึงเดือนนี้ รวมค้างจ่าย)
const projectSummary = ({ tenders = {}, additions = {}, extra = [], hidden = [], po = [] }) => {
  const combined = buildCombinedBudget(tenders || {}, additions || {});
  const effHidden = hiddenSafeForPO(hidden || [], po || []);
  const codes = [...ACCOUNTS.filter(a => !effHidden.includes(a.code)).map(a => a.code), ...(extra || []).filter(e => !e.parentCode).map(e => e.code)];
  const budget = codes.reduce((s, c) => s + (parseFloat(combined[c]) || 0), 0);
  const committed = (po || []).reduce((s, p) => s + poTotal(p), 0);
  const thisMonth = todayStr().slice(0, 7), byMonth = {};
  (po || []).flatMap(poPayLines).forEach(l => {
    const mk = l.month || "9999-99"; const b = byMonth[mk] || (byMonth[mk] = { sum: 0, paid: 0 });
    b.sum += l.amount; b.paid += (l.paidAmount || 0);
  });
  const dueNow = Object.entries(byMonth).filter(([mk]) => mk !== "9999-99" && mk <= thisMonth).reduce((s, [, b]) => s + Math.max(0, b.sum - b.paid), 0);
  const late = (po || []).filter(p => incomingStatus(p) === "late").length;
  const overCodes = codes.filter(c => { const b = parseFloat(combined[c]) || 0; const u = (po || []).reduce((s, p) => s + poAmountForCode(p, c), 0); return b > 0 && u > b + 0.005; }).length;
  return { budget, committed, pct: budget > 0 ? committed / budget * 100 : 0, dueNow, poCount: (po || []).length, late, overCodes };
};

// ─── Styling helper ─────────────────────────────────────────────────────────
// Lays down a colored title bar (merged across every column), optional gray
// info sub-rows, a bold colored header row with autofilter, zebra-striped
// bordered data rows with right-aligned money/% columns, and an optional
// bold total row — everything an aoa_to_sheet grid needs to read like a
// real report instead of a raw data dump.
const BORDER_THIN = (rgb) => ({ style:"thin", color:{rgb} });
// สีพิลล์ตามสถานะ (เขียว=เสร็จ/จ่ายแล้ว, เหลือง=กำลังทำ/รอ, แดง=ค้าง/เกินกำหนด)
// ใช้คีย์เวิร์ดจับ ครอบคลุมทั้งไทย/อังกฤษ สถานะอื่นเป็นพิลล์เทากลาง ๆ
const STATUS_PILL = [
  [/(completed|complete|เสร็จ|จ่ายแล้ว|รับของแล้ว|รับครบ|ปิดงาน|ปิด|อนุมัติ|approved|done|paid)/i, { bg:"D1FAE5", fg:"065F46" }],
  [/(in\s*progress|progress|กำลัง|ระหว่าง|บางส่วน|partial|สั่งซื้อ|สั่ง|รอรับ|รอจ่าย|pending|รอ)/i,        { bg:"FEF3C7", fg:"92400E" }],
  [/(ยังไม่กำหนด|unset|ไม่ระบุ)/i,                                                                  { bg:"E5E7EB", fg:"374151" }], // ยังไม่ตั้งค่า = เทา (ไม่ใช่แดงแบบเกินกำหนด)
  [/(to\s*do|todo|ร่าง|ยังไม่|ค้างจ่าย|ค้าง|เกินกำหนด|เกินงบ|ไม่มีงบ|overdue|ยกเลิก|cancel|reject)/i,               { bg:"FEE2E2", fg:"991B1B" }],
];
function statusPill(val) {
  const s = String(val == null ? "" : val);
  if (!s.trim() || s === "-") return null;
  for (const [re, st] of STATUS_PILL) if (re.test(s)) return st;
  return { bg:"E5E7EB", fg:"374151" };
}

// ผสมสีให้อ่อนลง (เข้าหาสีขาว) ratio 0..1 — ใช้ทำโทนพาสเทลนุ่ม ๆ
function lighten(hex, ratio) {
  const n = parseInt(hex, 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  const L = x => Math.round(x + (255 - x) * ratio);
  return ((L(r) << 16) | (L(g) << 8) | L(b)).toString(16).padStart(6, "0").toUpperCase();
}

// ─── Excel styling ─────────────────────────────────────────────────────────
function styleSheet(ws, { numCols, titleRow=0, subRows=[], headerRow, dataStart, dataEnd,
                           totalRow=null, moneyCols=[], usdCols=[], pctCols=[], centerCols=[], statusCols=[], theme,
                           // rowGroups: array aligned to dataStart..dataEnd holding a "group key" per
                           // row. When given, rows are shaded in solid blocks per group (instead of
                           // plain every-other-row zebra) and a heavier divider line marks where one
                           // group ends and the next begins — a long list then reads as clustered
                           // sections instead of a flat grid.
                           // groupDisplayCol: column index holding the group's label, bolded/tinted so
                           // the eye can track straight down that column.
                           // codeCol: column index whose text is tinted blue (เช่น Acc. Code) ให้ตรงกับหน้าจอ
                           rowGroups = null, groupDisplayCol = null, codeCol = null }) {
  ws["!rows"]   = ws["!rows"] || [];
  ws["!merges"] = ws["!merges"] || [];
  // ── ตารางสะอาด: เส้นตารางบาง ๆ สีเทาอ่อนทุกช่อง หัวตารางพื้นอ่อน (แบบรูปตัวอย่าง) ──
  const HFILL = lighten(theme.main, 0.88); // หัวตาราง พื้นอ่อน
  const GRID  = "CBD5E1";                  // เส้นตารางสีเทา (เข้มขึ้นให้เห็นชัด)
  const HRULE = "9AA7BA";                  // เส้นใต้หัวตาราง (เข้มกว่าเส้นทั่วไป)
  const BAND  = lighten(theme.main, 0.955);// แถบสลับสีจาง ๆ (โทนธีม)
  const TFILL = lighten(theme.main, 0.86); // แถวรวม พื้นอ่อน
  const TRULE = "9AA7BA";                  // เส้นเหนือแถวรวม
  const GLINE = lighten(theme.main, 0.60); // เส้นแบ่งกลุ่ม (ชัดขึ้น)
  const gridAll = { top:BORDER_THIN(GRID), bottom:BORDER_THIN(GRID), left:BORDER_THIN(GRID), right:BORDER_THIN(GRID) };

  ws["!merges"].push({ s:{r:titleRow,c:0}, e:{r:titleRow,c:numCols-1} });
  for (let c=0; c<numCols; c++) {
    const ref = XLSX.utils.encode_cell({r:titleRow,c});
    if (!ws[ref]) ws[ref] = { t:"s", v:"" };
    // หัวเรื่อง: ตัวหนาใหญ่ พื้นขาว ไม่มีแถบสี
    ws[ref].s = { font:{bold:true,sz:15,color:{rgb:theme.dark},name:"Tahoma"},
      alignment:{vertical:"center",horizontal:"left"} };
  }
  ws["!rows"][titleRow] = { hpx:34 };

  subRows.forEach(r => {
    for (let c=0; c<numCols; c++) {
      const ref = XLSX.utils.encode_cell({r,c});
      if (ws[ref]) ws[ref].s = { font:{sz:9.5,color:{rgb:"94A3B8"},name:"Tahoma"} };
    }
    ws["!rows"][r] = { hpx:18 };
  });

  for (let c=0; c<numCols; c++) {
    const ref = XLSX.utils.encode_cell({r:headerRow,c});
    if (!ws[ref]) ws[ref] = { t:"s", v:"" };
    const isMoney = moneyCols.includes(c), isPct = pctCols.includes(c), isCenter = centerCols.includes(c);
    ws[ref].s = { font:{bold:true,sz:10,color:{rgb:theme.dark},name:"Tahoma"},
      fill:{fgColor:{rgb:HFILL}},
      alignment:{vertical:"center",horizontal:isMoney||isPct?"right":isCenter?"center":"left",wrapText:true,indent:(isMoney||isPct||isCenter)?0:1},
      // เส้นตารางบางทุกด้าน + เส้นใต้หัวตารางเข้มขึ้นนิด
      border:{ ...gridAll, bottom:BORDER_THIN(HRULE) } };
  }
  ws["!rows"][headerRow] = { hpx:30 };
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s:{r:headerRow,c:0}, e:{r:headerRow,c:numCols-1} }) };
  // ตรึงทุกอย่างเหนือแถวข้อมูล (หัวข้อ+หัวตาราง) ให้ค้างไว้ตอนเลื่อน
  ws["!freeze"] = { xSplit:0, ySplit:headerRow+1, topLeftCell: XLSX.utils.encode_cell({ r:headerRow+1, c:0 }), activePane:"bottomLeft", state:"frozen" };

  // แถวสลับสีธรรมดาทีละแถว (ไม่ไล่สี/ไม่แบ่งเส้นตามกรุ๊ป) ให้ตารางเรียบเหมือนตัวอย่าง
  for (let r=dataStart; r<=dataEnd; r++) {
    const idx = r - dataStart;
    const zebra = idx % 2 === 1;
    const isGroupStart = false;
    for (let c=0; c<numCols; c++) {
      const ref = XLSX.utils.encode_cell({r,c});
      if (!ws[ref]) continue;
      const isMoney = moneyCols.includes(c), isPct = pctCols.includes(c), isCenter = centerCols.includes(c);
      const isGroupLabel = groupDisplayCol != null && c === groupDisplayCol;
      const isCode = codeCol != null && c === codeCol;
      const isText = !isMoney && !isPct && !isCenter;
      const s = { font: isGroupLabel
          ? {sz:10,name:"Tahoma",bold:true,color:{rgb:theme.dark}}
          : isCode
          ? {sz:10,name:"Tahoma",color:{rgb:theme.dark}}
          : {sz:10,name:"Tahoma",color:{rgb: isText ? "334155" : "475569"}},
        alignment:{ vertical:"center", horizontal:isMoney||isPct?"right":isCenter?"center":"left", wrapText:true, indent: isText?1:0 },
        // เส้นตารางบาง ๆ ทุกด้าน · ขึ้นกลุ่มใหม่ใช้เส้นบนเข้มขึ้นเป็นตัวแบ่ง
        border:{ ...gridAll, ...(isGroupStart?{ top:BORDER_THIN(GLINE) }:{}) } };
      if (zebra)   s.fill   = { fgColor:{rgb:BAND} };
      if (isMoney) s.numFmt = usdCols.includes(c) ? '"$"#,##0.00' : '"฿"#,##0';   // แยกสัญลักษณ์ $ / ฿
      if (isPct)   s.numFmt = "0.0%";
      if (statusCols.includes(c)) {
        const pill = statusPill(ws[ref].v);
        if (pill) {
          s.fill = { fgColor:{rgb:pill.bg} };
          s.font = { sz:10, name:"Tahoma", bold:true, color:{rgb:pill.fg} };
          s.alignment = { ...s.alignment, horizontal:"center", indent:0 };
        }
      }
      ws[ref].s = s;
    }
    ws["!rows"][r] = ws["!rows"][r] || { hpx:22 };
  }

  if (totalRow != null) {
    for (let c=0; c<numCols; c++) {
      const ref = XLSX.utils.encode_cell({r:totalRow,c});
      if (!ws[ref]) ws[ref] = { t:"s", v:"" };
      const isMoney = moneyCols.includes(c), isPct = pctCols.includes(c), isCenter = centerCols.includes(c);
      const isText = !isMoney && !isPct && !isCenter;
      ws[ref].s = { font:{bold:true,sz:10.5,color:{rgb:theme.dark},name:"Tahoma"}, fill:{fgColor:{rgb:TFILL}},
        alignment:{vertical:"center",horizontal:isMoney||isPct?"right":isCenter?"center":"left",indent:isText?1:0},
        // เส้นตารางบางทุกด้าน + เส้นเหนือแถวรวมเข้มขึ้นนิด
        border:{ ...gridAll, top:BORDER_THIN(TRULE) },
        numFmt: isMoney ? (usdCols.includes(c) ? '"$"#,##0.00' : '"฿"#,##0') : isPct?"0.0%":undefined };
    }
    ws["!rows"][totalRow] = { hpx:26 };
  }

  // จัดความกว้างคอลัมน์อัตโนมัติให้พอดีข้อความ (ดูจากหัวตาราง + ข้อมูล + แถวรวม)
  const cols = [];
  const scan = (r, c) => {
    if (r == null) return;
    const cell = ws[XLSX.utils.encode_cell({r,c})];
    if (!cell) return;
    let v = cell.v;
    let s = (typeof v === "number") ? Math.round(v).toLocaleString("en-US") : String(v == null ? "" : v);
    if (s.length > (cols[c]||0)) cols[c] = s.length;
  };
  for (let c=0; c<numCols; c++) {
    cols[c] = 0;
    scan(headerRow, c);
    for (let r=dataStart; r<=dataEnd; r++) scan(r, c);
    scan(totalRow, c);
  }
  ws["!cols"] = cols.map(w => ({ wch: Math.min(55, Math.max(8, w + 2)) }));
}

// ── ตัวช่วย USD สำหรับ Export ── ให้ไฟล์ Excel มีข้อมูลสอดคล้องกับหน้าจอ (บาท + ดอลลาร์)
// exportRate: อัตราแลกเปลี่ยนของโปรเจกต์ (0 = ปิด USD → export เป็นบาทล้วนตามเดิม)
function exportRate(project){ return (project?.showUsd !== false) ? (parseFloat(project?.usdRate)||0) : 0; }
// toUsd: แปลงยอดบาท→USD (คืน "" ถ้าไม่ได้เปิด USD หรือค่าไม่ใช่ตัวเลข)
function toUsd(thb, rate){ return (rate>0 && typeof thb==="number" && isFinite(thb)) ? Math.round((thb/rate)*100)/100 : ""; }

// ── ลิงก์ข้ามชีต (hyperlink ภายในไฟล์) — คลิกแล้วกระโดดไปชีตปลายทาง ช่วยไล่ว่าข้อมูลมาจากไหน ──
// ทำให้เซลล์ที่ ref เป็นลิงก์ไปยัง 'sheetName'  (สีน้ำเงินขีดเส้นใต้)
function xLinkCell(ws, ref, sheetName, tip){
  if (!ws[ref]) ws[ref] = { t:"s", v:"" };
  ws[ref].l = { Target:`#'${sheetName}'!A1`, Tooltip: tip || `ไปที่ชีต ${sheetName}` };
  const prev = ws[ref].s || {};
  ws[ref].s = { ...prev, font:{ ...(prev.font||{}), color:{rgb:"1D4ED8"}, underline:true } };
}
// วางแถวลิงก์ (links = [{text, sheet}]) ที่แถว r — ใช้ sheet_add_aoa เพื่อขยาย !ref ให้ด้วย
function xLinkRow(ws, r, links){
  if (!links || !links.length) return;
  XLSX.utils.sheet_add_aoa(ws, [links.map(l=>l.text)], { origin:{ r, c:0 } });
  links.forEach((l,i)=>{
    const ref = XLSX.utils.encode_cell({ r, c:i });
    ws[ref].l = { Target:`#'${l.sheet}'!A1`, Tooltip:`ไปที่ชีต ${l.sheet}` };
    ws[ref].s = { font:{ color:{rgb:"1D4ED8"}, underline:true, bold:true, name:"Tahoma", sz:10 }, alignment:{ vertical:"center" } };
  });
  ws["!rows"] = ws["!rows"] || []; ws["!rows"][r] = { hpx:20 };
}
// ลิงก์ "↑ กลับหน้าสรุป" ที่แถว r คอลัมน์ท้าย ๆ ของชีตรายละเอียด
function xBackLink(ws, r, c, backSheet){
  XLSX.utils.sheet_add_aoa(ws, [["↑ กลับหน้าสรุป"]], { origin:{ r, c } });
  const ref = XLSX.utils.encode_cell({ r, c });
  ws[ref].l = { Target:`#'${backSheet}'!A1`, Tooltip:`กลับไปชีต ${backSheet}` };
  ws[ref].s = { font:{ color:{rgb:"1D4ED8"}, underline:true, bold:true, name:"Tahoma", sz:10 }, alignment:{ horizontal:"right", vertical:"center" } };
}

// หน้า "สรุป (Dashboard)" — การ์ดตัวเลข + กราฟแท่งรายเดือน อยู่ในหน้าเดียวกัน
function addDashboardSheet(wb, sheetName, { title, subtitle, theme, cards = [], chartTitle, items = [], groups = null }) {
  items = items.filter(Boolean);
  groups = (groups || []).filter(Boolean);
  const n = items.length;
  const C = Math.max(1 + n, 8);              // อย่างน้อย 8 คอลัมน์
  const H = 10;                              // ความสูงกราฟ (แถว)
  const cardLabelRow = 3, cardValRow = 4, chartTitleRow = 6, valueRow = 7, chartTop = 8, labelRow = chartTop + H;
  const hasG = groups.length > 0;
  const gTitleRow = labelRow + 2, gHeadRow = gTitleRow + 1, gStart = gHeadRow + 1, gEnd = gStart + groups.length - 1, gTotalRow = gEnd + 1;
  const nRows = (hasG ? gTotalRow : labelRow) + 2;
  const aoa = Array.from({ length:nRows }, () => new Array(C).fill(""));
  aoa[0][0] = title; aoa[1][0] = subtitle || ""; aoa[chartTitleRow][0] = chartTitle || "";
  if (hasG) aoa[gTitleRow][0] = "สรุปตามกลุ่มวัสดุ (สัดส่วนงบรวม)";
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [
    { s:{r:0,c:0}, e:{r:0,c:C-1} },
    { s:{r:1,c:0}, e:{r:1,c:C-1} },
    { s:{r:chartTitleRow,c:0}, e:{r:chartTitleRow,c:C-1} },
  ];
  ws["!cols"] = [{ wch: hasG?18:4 }, ...Array.from({ length:C-1 }, () => ({ wch:12 }))];
  ws["!rows"] = [];
  ws["!rows"][0]={hpx:30}; ws["!rows"][1]={hpx:16};
  ws["!rows"][cardLabelRow]={hpx:18}; ws["!rows"][cardValRow]={hpx:32};
  ws["!rows"][chartTitleRow]={hpx:22};
  for (let r=chartTop; r<chartTop+H; r++) ws["!rows"][r]={hpx:14};
  ws["!rows"][labelRow]={hpx:22};
  const setS = (r, c, s, v, f) => {
    const ref = XLSX.utils.encode_cell({ r, c });
    if (v != null) ws[ref] = { t: typeof v === "number" ? "n" : "s", v };
    else if (!ws[ref]) ws[ref] = { t:"s", v:"" };
    if (f) { ws[ref].t = "n"; ws[ref].f = f; }
    ws[ref].s = s;
  };
  setS(0,0,{ font:{bold:true,sz:14,color:{rgb:"FFFFFF"},name:"Tahoma"}, fill:{fgColor:{rgb:theme.main}}, alignment:{vertical:"center",horizontal:"left",indent:1} });
  setS(1,0,{ font:{italic:true,sz:10,color:{rgb:"64748B"},name:"Tahoma"} });
  setS(chartTitleRow,0,{ font:{bold:true,sz:11,color:{rgb:theme.dark},name:"Tahoma"}, fill:{fgColor:{rgb:lighten(theme.main,0.85)}}, alignment:{vertical:"center",horizontal:"left",indent:1} });
  // การ์ดสรุป 4 ใบ (แต่ละใบกว้าง 2 คอลัมน์)
  // สีการ์ดมาตรฐาน; แต่ละการ์ดกำหนดสีเองได้ผ่าน cd.acc (["bg","fg"]) เช่น ให้ตรงกับหน้าจอ
  const ACC = [["DBEAFE","1D4ED8"],["D1FAE5","047857"],["FEF3C7","92400E"],["EDE9FE","6D28D9"]];
  cards.slice(0,4).forEach((cd, i) => {
    const c0 = i*2, c1 = c0+1, [bg,fg] = cd.acc || ACC[i%4];
    ws["!merges"].push({ s:{r:cardLabelRow,c:c0}, e:{r:cardLabelRow,c:c1} });
    ws["!merges"].push({ s:{r:cardValRow,c:c0}, e:{r:cardValRow,c:c1} });
    setS(cardLabelRow, c0, { font:{bold:true,sz:9.5,color:{rgb:fg},name:"Tahoma"}, fill:{fgColor:{rgb:bg}}, alignment:{horizontal:"center",vertical:"center"} }, cd.label);
    setS(cardLabelRow, c1, { fill:{fgColor:{rgb:bg}} });
    setS(cardValRow, c0, { font:{bold:true,sz:15,color:{rgb:fg},name:"Tahoma"}, fill:{fgColor:{rgb:bg}}, alignment:{horizontal:"center",vertical:"center"}, numFmt: cd.money?"#,##0":undefined }, cd.value, cd.f);
    setS(cardValRow, c1, { fill:{fgColor:{rgb:bg}} });
  });
  // กราฟแท่งรายเดือน
  const max = Math.max(...items.map(i=>i.value||0), 1);
  const barOn = theme.main, barOff = "F3F4F6";
  items.forEach((it, i) => {
    const c = 1+i;
    const filled = Math.max(0, Math.round(((it.value||0)/max)*H));
    setS(valueRow, c, { font:{bold:true,sz:8.5,color:{rgb:theme.dark},name:"Tahoma"}, alignment:{horizontal:"center"}, numFmt:"#,##0" }, it.value||0, it.f);
    for (let k=0; k<H; k++) { const r = chartTop+(H-1-k); setS(r, c, { fill:{fgColor:{rgb: k<filled?barOn:barOff }} }); }
    setS(labelRow, c, { font:{bold:true,sz:9,color:{rgb:"374151"},name:"Tahoma"}, alignment:{horizontal:"center",wrapText:true} }, it.label);
  });
  // ตารางสรุปตามกลุ่ม + แถบสัดส่วน
  if (hasG) {
    ws["!merges"].push({ s:{r:gTitleRow,c:0}, e:{r:gTitleRow,c:C-1} });
    ws["!rows"][gTitleRow] = {hpx:22};
    setS(gTitleRow, 0, { font:{bold:true,sz:11,color:{rgb:theme.dark},name:"Tahoma"}, fill:{fgColor:{rgb:lighten(theme.main,0.85)}}, alignment:{vertical:"center",horizontal:"left",indent:1} });
    const headFill = lighten(theme.main, 0.82);
    const gh = ["กลุ่ม",`ราคาเดิม + เผื่อ ${WASTE_LBL}`,"เพิ่มรายเดือน","งบรวม","สัดส่วน","กราฟสัดส่วน"];
    gh.forEach((h,c) => setS(gHeadRow, c, { font:{bold:true,sz:9.5,color:{rgb:theme.dark},name:"Tahoma"}, fill:{fgColor:{rgb:headFill}}, alignment:{horizontal:c===0?"left":c<5?"right":"left",vertical:"center",indent:c===0||c===5?1:0}, border:{bottom:BORDER_THIN(lighten(theme.main,0.45))} }, h));
    for (let c=6;c<C;c++) setS(gHeadRow, c, { fill:{fgColor:{rgb:headFill}}, border:{bottom:BORDER_THIN(lighten(theme.main,0.45))} });
    if (C-1 > 5) ws["!merges"].push({ s:{r:gHeadRow,c:5}, e:{r:gHeadRow,c:C-1} });
    ws["!rows"][gHeadRow] = {hpx:22};
    const zeb = lighten(theme.main, 0.95), gmax = Math.max(...groups.map(g=>g.total||0), 1);
    groups.forEach((g, i) => {
      const r = gStart + i, fillZ = i%2===1 ? { fill:{fgColor:{rgb:zeb}} } : {};
      const bd = { border:{bottom:BORDER_THIN("EEF0F2")} };
      setS(r, 0, { ...fillZ, ...bd, font:{bold:true,sz:9.5,color:{rgb:theme.dark},name:"Tahoma"}, alignment:{vertical:"center",horizontal:"left",indent:1} }, g.label);
      setS(r, 1, { ...fillZ, ...bd, font:{sz:9.5,color:{rgb:"1F2937"},name:"Tahoma"}, alignment:{vertical:"center",horizontal:"right"}, numFmt:"#,##0" }, g.base||0);
      setS(r, 2, { ...fillZ, ...bd, font:{sz:9.5,color:{rgb:"1F2937"},name:"Tahoma"}, alignment:{vertical:"center",horizontal:"right"}, numFmt:"#,##0" }, g.add||0);
      setS(r, 3, { ...fillZ, ...bd, font:{bold:true,sz:9.5,color:{rgb:"1F2937"},name:"Tahoma"}, alignment:{vertical:"center",horizontal:"right"}, numFmt:"#,##0" }, g.total||0);
      setS(r, 4, { ...fillZ, ...bd, font:{sz:9.5,color:{rgb:theme.dark},name:"Tahoma"}, alignment:{vertical:"center",horizontal:"center"}, numFmt:"0.0%" }, g.pct||0);
      setS(r, 5, { ...fillZ, ...bd, font:{sz:10,color:{rgb:theme.main},name:"Tahoma"}, alignment:{vertical:"center",horizontal:"left"} }, "█".repeat(Math.max(0, Math.round(((g.total||0)/gmax)*22))));
      for (let c=6;c<C;c++) setS(r, c, { ...fillZ, ...bd });
      if (C-1 > 5) ws["!merges"].push({ s:{r,c:5}, e:{r,c:C-1} });
      ws["!rows"][r] = {hpx:18};
    });
    const rT = gTotalRow, tb = { fill:{fgColor:{rgb:lighten(theme.main,0.88)}}, border:{top:BORDER_THIN(lighten(theme.main,0.40))} };
    const tf = (h) => ({ ...tb, font:{bold:true,sz:9.5,color:{rgb:theme.dark},name:"Tahoma"}, alignment:{vertical:"center",horizontal:h,indent:h==="left"?1:0} });
    setS(rT,0,tf("left"),"รวมทั้งหมด");
    setS(rT,1,{...tf("right"),numFmt:"#,##0"}, groups.reduce((s,g)=>s+(g.base||0),0));
    setS(rT,2,{...tf("right"),numFmt:"#,##0"}, groups.reduce((s,g)=>s+(g.add||0),0));
    setS(rT,3,{...tf("right"),numFmt:"#,##0"}, groups.reduce((s,g)=>s+(g.total||0),0));
    setS(rT,4,{...tf("center"),numFmt:"0.0%"}, 1);
    for (let c=5;c<C;c++) setS(rT,c,{...tb});
    ws["!rows"][rT] = {hpx:20};
  }
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return { ws, nextRow: nRows };   // คืน worksheet + แถวว่างถัดไป เผื่ออยากต่อตารางใต้ dashboard
}

// ─── QS: budget / tender-cost export ───────────────────────────────────────
function exportQSExcel(project, tenderCosts, additions, extraItems=[], hiddenAccounts=[]) {
  const wb = XLSX.utils.book_new();
  const theme = { main:"2563EB", dark:"1D4ED8" };
  const rate = exportRate(project); const U = rate > 0;   // U = ใส่คอลัมน์ USD ไหม
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);
  const accounts = exportAccountList(extraItems, hiddenAccounts);
  // รายการบัญชีที่มีค่า (ใช้ร่วมกันทั้ง 2 ชีต เพื่อให้ตำแหน่งแถวตรงกัน → ลิงก์สูตรได้)
  const dashList = accounts.filter(a => {
    const bs = parseFloat(tenderCosts[a.code]) || 0;
    const tt = parseFloat(combinedBudget[a.code]) || 0;
    return !(tt <= 0 && bs <= 0);
  });

  // หน้าแรก = สรุป (Dashboard): การ์ดตัวเลข + กราฟยอดเพิ่มรายเดือน (ในหน้าเดียว)
  const dashMonths = [...new Set(Object.keys(additions||{}).filter(k=>!k.startsWith("$")))].sort();
  const dashItems  = dashMonths.map(m => ({ label: monthShortLabel(m), value: accounts.reduce((s,a)=> s + monthAddValue(additions, m, a.code), 0) }));
  const dashBase   = accounts.reduce((s,a)=> s + (parseFloat(tenderCosts[a.code])||0), 0);
  const dashWaste  = accounts.reduce((s,a)=> s + wasteOf(tenderCosts[a.code]), 0);
  const dashAdded  = dashItems.reduce((s,i)=> s + i.value, 0);
  // แถว TOTAL (A1) ของชีต "งบประมาณ"/"รายเดือน (สรุป)" = 5 + จำนวนแถวข้อมูล
  // (หัวข้อ 3 แถว + หัวตารางแถว 4 → ข้อมูลเริ่มแถว 5 → TOTAL อยู่แถว 5+N)
  const TR = 5 + dashList.length;
  const dashItemsF = dashItems.map((it, i) => ({ ...it, f: `'รายเดือน (สรุป)'!${XLSX.utils.encode_col(4 + i)}${TR}` }));
  // สรุปตามกลุ่มวัสดุ (ไว้โชว์ตาราง+แถบสัดส่วนในหน้าสรุป)
  const byG = {};
  dashList.forEach(a => {
    const bs = withWaste(tenderCosts[a.code]), tt = parseFloat(combinedBudget[a.code]) || 0, g = a.group || "อื่น ๆ";   // base = ราคาเดิม + เผื่อ %
    (byG[g] = byG[g] || { base:0, total:0 }); byG[g].base += bs; byG[g].total += tt;
  });
  const grandTot = Object.values(byG).reduce((s,x)=>s+x.total,0) || 1;
  const groupData = Object.entries(byG)
    .map(([label,x]) => ({ label, base:x.base, add:x.total-x.base, total:x.total, pct:x.total/grandTot }))
    .sort((a,b)=> b.total - a.total);
  const dashQS = addDashboardSheet(wb, "สรุป", {
    title: `สรุปงบประมาณ — ${project.name}`,
    subtitle: `พื้นที่ ${project.area||"-"} ft² · แผง ${project.panels||"-"} · Export: ${new Date().toLocaleDateString("th-TH")}`,
    theme,
    // 4 การ์ดให้ตรงกับหน้า Baseline: ราคาเดิม · เผื่อเศษ 3% (รวมในงบ) · งานเพิ่ม · รวมทั้งหมด
    cards: [
      { label:"ราคาเดิม (Tender Cost)",                 value: dashBase,  money:true, f:`'งบประมาณ'!D${TR}`, acc:["DBEAFE","1D4ED8"] },
      { label:`เผื่อเศษ/สูญเสีย ${WASTE_LBL} (รวมในงบ)`, value: dashWaste, money:true, f:`'งบประมาณ'!E${TR}`, acc:["FEF3C7","92400E"] },
      { label:"งานเพิ่ม (รวมทุกเดือน)",                   value: dashAdded, money:true, f:`'งบประมาณ'!F${TR}`, acc:["EDE9FE","6D28D9"] },
      { label:"รวมทั้งหมด",                              value: dashBase + dashWaste + dashAdded, money:true, f:`'งบประมาณ'!G${TR}`, acc:["D1FAE5","047857"] },
    ],
    chartTitle: "กราฟ: ยอดเพิ่มรายเดือน (THB)",
    items: dashItemsF,
    groups: groupData,
  });
  const monthLinksQS = [];   // เก็บชื่อชีตรายเดือนไว้ทำลิงก์บนหน้าสรุป

  // Sheet 1 — Baseline + monthly additions rolled up per Acc. Code
  const rows1 = [[`งบประมาณ (Tender Cost) — ${project.name}`], [`พื้นที่ ${project.area||"-"} ft²  ·  แผง ${project.panels||"-"}  ·  Export: ${new Date().toLocaleDateString("th-TH")}`], []];
  rows1.push(["Acc. Code","Account Name","Group","ราคาเดิม",`เผื่อเศษ ${WASTE_LBL}`,"เพิ่มรายเดือน (รวม)","รวมทั้งหมด",...(U?["รวมทั้งหมด (USD)"]:[])]);
  const dataStart1 = rows1.length;
  const rowGroups1 = [];
  dashList.forEach(a => {
    const baseline = parseFloat(tenderCosts[a.code]) || 0;
    const waste    = wasteOf(baseline);
    const added    = Object.keys(additions||{}).reduce((s,m)=> m.startsWith("$") ? s : s + monthAddValue(additions, m, a.code), 0);   // รวมตรง ๆ (ไม่ลบกัน กันเศษทศนิยม -฿0)
    const total    = baseline + waste + added;
    rows1.push([a.code, a.name, a.group, baseline, waste, added, total, ...(U?[toUsd(total,rate)]:[])]);
    rowGroups1.push(a.group);
  });
  const dataEnd1 = rows1.length-1;
  // ใส่ยอดรวมจริงเป็นค่าของเซลล์ด้วย (นอกจากสูตร) — แอปพรีวิวที่ไม่คำนวณสูตร (มือถือ/อีเมล) จะไม่โชว์ ฿0
  const colSum1 = (c) => rows1.slice(dataStart1, dataEnd1 + 1).reduce((s, r) => s + (Number(r[c]) || 0), 0);
  rows1.push(["","TOTAL","",colSum1(3),colSum1(4),colSum1(5),colSum1(6), ...(U?[toUsd(dashBase+dashWaste+dashAdded,rate)]:[])]);
  const totalRow1 = rows1.length-1;
  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  // ลิงก์ด้วยสูตร: เผื่อเศษ = ราคาเดิม × % · งบรวม = ราคาเดิม + เผื่อเศษ + เพิ่ม (ต่อแถว) · TOTAL = ผลรวมทั้งคอลัมน์
  for (let r = dataStart1; r <= dataEnd1; r++) {
    const R = r + 1, refW = XLSX.utils.encode_cell({ r, c:4 }), ref = XLSX.utils.encode_cell({ r, c:6 });
    if (ws1[refW]) ws1[refW].f = `ROUND(D${R}*${WASTE_RATE},2)`;
    if (ws1[ref]) ws1[ref].f = `D${R}+E${R}+F${R}`;
  }
  ["D","E","F","G"].forEach((L, i) => {
    const ref = XLSX.utils.encode_cell({ r:totalRow1, c:3+i });
    if (ws1[ref]) ws1[ref].f = `SUM(${L}${dataStart1+1}:${L}${dataEnd1+1})`;
  });
  ws1["!cols"] = [{wch:12},{wch:40},{wch:16},{wch:18},{wch:14},{wch:18},{wch:18},...(U?[{wch:18}]:[])];
  styleSheet(ws1, { numCols:7+(U?1:0), subRows:[1], headerRow:3, dataStart:dataStart1, dataEnd:dataEnd1, totalRow:totalRow1,
    moneyCols:U?[3,4,5,6,7]:[3,4,5,6], usdCols:U?[7]:[], theme, rowGroups:rowGroups1, groupDisplayCol:2, codeCol:0 });
  xBackLink(ws1, 2, (7+(U?1:0))-1, "สรุป");
  XLSX.utils.book_append_sheet(wb, ws1, "งบประมาณ");

  // Sheet 2 — one column per month, so QS can see exactly how the budget grew
  const months = [...new Set(Object.keys(additions||{}).filter(k=>!k.startsWith("$")))].sort();
  const rows2 = [[`รายการเพิ่มรายเดือน — ${project.name}`], [`Export: ${new Date().toLocaleDateString("th-TH")}`], []];
  rows2.push(["Acc. Code","Account Name","ราคาเดิม",`เผื่อเศษ ${WASTE_LBL}`, ...months.map(monthShortLabel), "รวมทั้งหมด", ...(U?["รวม (USD)"]:[])]);
  const dataStart2 = rows2.length;
  const rowGroups2 = [];
  dashList.forEach(a => {
    const baseline  = parseFloat(tenderCosts[a.code]) || 0;
    const waste     = wasteOf(baseline);
    const monthVals = months.map(m => monthAddValue(additions, m, a.code));
    const total = baseline + waste + monthVals.reduce((s,v)=>s+v,0);
    rows2.push([a.code, a.name, baseline, waste, ...monthVals, total, ...(U?[toUsd(total,rate)]:[])]);
    rowGroups2.push(a.group);
  });
  const dataEnd2 = rows2.length-1;
  const M = months.length, totColC = 4 + M;
  const colSum2 = (c) => rows2.slice(dataStart2, dataEnd2 + 1).reduce((s, r) => s + (Number(r[c]) || 0), 0);
  rows2.push(["","TOTAL",colSum2(2),colSum2(3), ...months.map((_,i)=>colSum2(4+i)), colSum2(totColC), ...(U?[toUsd(dashBase+dashWaste+dashAdded,rate)]:[])]);
  const totalRow2 = rows2.length-1;
  const numCols2 = 5 + months.length + (U?1:0);
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  // ลิงก์ด้วยสูตร: เผื่อเศษ = ราคาเดิม × % · รวมทั้งหมด(ต่อแถว) = ราคาเดิม + เผื่อเศษ + ผลรวมทุกเดือน · TOTAL = ผลรวมคอลัมน์
  const lastMonthL = XLSX.utils.encode_col(3 + M);
  for (let r = dataStart2; r <= dataEnd2; r++) {
    const R = r + 1, refW = XLSX.utils.encode_cell({ r, c: 3 }), ref = XLSX.utils.encode_cell({ r, c: totColC });
    if (ws2[refW]) ws2[refW].f = `ROUND(C${R}*${WASTE_RATE},2)`;
    if (ws2[ref]) ws2[ref].f = M > 0 ? `C${R}+D${R}+SUM(E${R}:${lastMonthL}${R})` : `C${R}+D${R}`;
  }
  [2, 3, ...months.map((_,i)=>4+i), totColC].forEach(c => {
    const L = XLSX.utils.encode_col(c), ref = XLSX.utils.encode_cell({ r:totalRow2, c });
    if (ws2[ref]) ws2[ref].f = `SUM(${L}${dataStart2+1}:${L}${dataEnd2+1})`;
  });
  ws2["!cols"] = [{wch:12},{wch:34},{wch:14},{wch:12}, ...months.map(()=>({wch:12})), {wch:16}, ...(U?[{wch:16}]:[])];
  styleSheet(ws2, { numCols:numCols2, subRows:[1], headerRow:3, dataStart:dataStart2, dataEnd:dataEnd2, totalRow:totalRow2,
    moneyCols:[2, 3, ...months.map((_,i)=>4+i), 4+months.length, ...(U?[5+months.length]:[])], usdCols:U?[5+months.length]:[], theme, rowGroups:rowGroups2, codeCol:0 });
  xBackLink(ws2, 2, numCols2-1, "สรุป");
  XLSX.utils.book_append_sheet(wb, ws2, "รายเดือน (สรุป)");

  // Sheet 3+ — แยกรายเดือน โดย breakdown ตามคอลัมน์ (รายการย่อย) ของเดือนนั้น ๆ
  // คอลัมน์เก็บเป็นรายเดือน แต่ละเดือนอาจมีชุดคอลัมน์ต่างกัน → ทำหนึ่งชีตต่อเดือน
  const sheetName = (s) => String(s).replace(/[\\/?*[\]:]/g, "-").slice(0, 28);
  const usedNames = {};
  const monthSheetMap = {};   // เดือน → ชื่อชีต ไว้ทำลิงก์
  months.forEach((m) => {
    const cols = (additions[m] && additions[m].$columns) || additions.$columns || [];
    const hasCols = cols.length > 0;
    const brk = Object.fromEntries(accounts.map(a => [a.code, monthRowBreakdown(additions, m, a.code, cols)]));
    const hasOther = hasCols && accounts.some(a => brk[a.code].other !== 0);
    const valLabels = hasCols ? [...cols.map(c => c.name || "รายการ"), ...(hasOther ? [OTHER_COL_LABEL] : [])] : ["เพิ่มเดือนนี้"];
    const rows = [
      [`เพิ่มรายเดือน ${monthShortLabel(m)} — ${project.name}`],
      [hasCols ? `แยกตามรายการ ${cols.length} คอลัมน์  ·  Export: ${new Date().toLocaleDateString("th-TH")}`
               : `Export: ${new Date().toLocaleDateString("th-TH")}`],
      [],
      ["Acc. Code", "Account Name", "Group", ...valLabels, "รวมเดือนนี้", ...(U?["รวม (USD)"]:[])],
    ];
    const dataStart = rows.length;
    const colTotals = valLabels.map(() => 0);
    let grand = 0;
    const rowGroups = [];
    accounts.forEach(a => {
      const b = brk[a.code];
      const vals = hasOther ? [...b.vals, b.other] : b.vals;
      const rowTotal = b.total;
      if (rowTotal === 0 && !vals.some(v => v !== 0)) return; // เอาเฉพาะรายการที่มียอดในเดือนนี้
      rows.push([a.code, a.name, a.group, ...vals, rowTotal, ...(U?[toUsd(rowTotal,rate)]:[])]);
      rowGroups.push(a.group);
      vals.forEach((v, i) => { colTotals[i] += v; });
      grand += rowTotal;
    });
    if (rows.length === dataStart) return; // เดือนนี้ไม่มีข้อมูล ข้ามชีต
    const dataEnd = rows.length - 1;
    rows.push(["", "TOTAL", "", ...colTotals, grand, ...(U?[toUsd(grand,rate)]:[])]);
    const totalRow = rows.length - 1;
    const numCols = 4 + valLabels.length + (U?1:0);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 34 }, { wch: 14 }, ...valLabels.map(() => ({ wch: 15 })), { wch: 16 }, ...(U?[{ wch: 16 }]:[])];
    styleSheet(ws, {
      numCols, subRows: [1], headerRow: 3, dataStart, dataEnd, totalRow,
      moneyCols: [...valLabels.map((_, i) => 3 + i), 3 + valLabels.length, ...(U?[4 + valLabels.length]:[])], usdCols:U?[4 + valLabels.length]:[],
      theme, rowGroups, groupDisplayCol: 2, codeCol: 0,
    });
    let nm = sheetName(monthShortLabel(m));
    if (usedNames[nm]) { usedNames[nm] += 1; nm = sheetName(`${nm} ${usedNames[nm]}`); } else usedNames[nm] = 1;
    xBackLink(ws, 2, numCols-1, "สรุป");
    XLSX.utils.book_append_sheet(wb, ws, nm);
    monthSheetMap[m] = nm;
    monthLinksQS.push({ text: monthShortLabel(m), sheet: nm });
  });
  // ลิงก์หัวคอลัมน์เดือนในชีต "รายเดือน (สรุป)" → กระโดดไปชีตของเดือนนั้น (ไล่ที่มา)
  months.forEach((m,i)=>{ if (monthSheetMap[m]) xLinkCell(ws2, XLSX.utils.encode_cell({ r:3, c:4+i }), monthSheetMap[m], `ดูรายละเอียดเดือน ${monthShortLabel(m)}`); });
  // แถบลิงก์นำทางใต้ dashboard หน้าสรุป
  {
    const navR = dashQS.nextRow + 1;
    XLSX.utils.sheet_add_aoa(dashQS.ws, [["🔗 ไปที่ชีต:"]], { origin:{ r:navR, c:0 } });
    dashQS.ws[XLSX.utils.encode_cell({ r:navR, c:0 })].s = { font:{ bold:true, sz:10.5, color:{rgb:theme.dark}, name:"Tahoma" } };
    xLinkRow(dashQS.ws, navR+1, [{text:"📄 งบประมาณ (รายรหัส)", sheet:"งบประมาณ"}, {text:"📅 รายเดือน (สรุป)", sheet:"รายเดือน (สรุป)"}]);
    if (monthLinksQS.length) {
      XLSX.utils.sheet_add_aoa(dashQS.ws, [["🔗 รายละเอียดรายเดือน:"]], { origin:{ r:navR+2, c:0 } });
      dashQS.ws[XLSX.utils.encode_cell({ r:navR+2, c:0 })].s = { font:{ bold:true, sz:10.5, color:{rgb:theme.dark}, name:"Tahoma" } };
      xLinkRow(dashQS.ws, navR+3, monthLinksQS);
    }
  }

  const fname = `QS_Budget_${project.name.replace(/\s+/g,"_")}_${todayStr()}.xlsx`;
  XLSX.writeFile(wb, fname);
  return { wb, fname };
}

// ─── QS: export เฉพาะเดือนที่เลือก (แยกคอลัมน์ของเดือนนั้น + ยอดสะสมถึงเดือนนี้) ──
function exportQSMonthExcel(project, tenderCosts, additions, month, extraItems=[], hiddenAccounts=[]) {
  const wb = XLSX.utils.book_new();
  const theme = { main:"2563EB", dark:"1D4ED8" };
  const rate = exportRate(project); const U = rate > 0;   // U = ใส่คอลัมน์ USD ไหม
  const accounts = exportAccountList(extraItems, hiddenAccounts);
  const clean = (s) => String(s).replace(/[\\/?*[\]:]/g, "-").slice(0, 28);
  const allMonths = [...new Set(Object.keys(additions||{}).filter(k=>!k.startsWith("$")))].sort();
  const upto = allMonths.filter(m => m <= month);
  const cols = (additions[month] && additions[month].$columns) || additions.$columns || [];
  const hasCols = cols.length > 0;
  const brk = Object.fromEntries(accounts.map(a => [a.code, monthRowBreakdown(additions, month, a.code, cols)]));
  const hasOther = hasCols && accounts.some(a => brk[a.code].other !== 0);
  const valLabels = hasCols ? [...cols.map(c => c.name || "รายการ"), ...(hasOther ? [OTHER_COL_LABEL] : [])] : ["เพิ่มเดือนนี้"];

  const rows = [
    [`เพิ่มรายเดือน ${monthShortLabel(month)} — ${project.name}`],
    [hasCols ? `แยกตามรายการ ${cols.length} คอลัมน์  ·  Export: ${new Date().toLocaleDateString("th-TH")}`
             : `Export: ${new Date().toLocaleDateString("th-TH")}`],
    [],
    ["Acc. Code", "Account Name", "Group", "ราคาเดิม", `เผื่อเศษ ${WASTE_LBL}`, ...valLabels, "รวมเดือนนี้", "รวมสะสมถึงเดือนนี้", ...(U?["รวมสะสม (USD)"]:[])],
  ];
  const dataStart = rows.length;
  const colTotals = valLabels.map(() => 0);
  let gBase = 0, gWaste = 0, gMonth = 0, gCum = 0;
  const rowGroups = [];
  accounts.forEach(a => {
    const baseline = parseFloat(tenderCosts[a.code]) || 0;
    const b = brk[a.code];
    const vals = hasOther ? [...b.vals, b.other] : b.vals;
    const monthTot = b.total;   // = ค่าหลักที่บันทึกไว้ ตรงกับ "รวมสะสม" และชีตสรุป
    const waste = wasteOf(baseline);
    const cum = baseline + waste + upto.reduce((s, m) => s + monthAddValue(additions, m, a.code), 0);
    if (monthTot === 0 && baseline === 0 && cum === 0) return;
    rows.push([a.code, a.name, a.group, baseline, waste, ...vals, monthTot, cum, ...(U?[toUsd(cum,rate)]:[])]);
    rowGroups.push(a.group);
    vals.forEach((v, i) => { colTotals[i] += v; });
    gBase += baseline; gWaste += waste; gMonth += monthTot; gCum += cum;
  });
  const dataEnd = rows.length - 1;
  rows.push(["", "TOTAL", "", gBase, gWaste, ...colTotals, gMonth, gCum, ...(U?[toUsd(gCum,rate)]:[])]);
  const totalRow = rows.length - 1;
  const numCols = 7 + valLabels.length + (U?1:0);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch:12 }, { wch:34 }, { wch:14 }, { wch:16 }, { wch:13 }, ...valLabels.map(()=>({ wch:15 })), { wch:16 }, { wch:18 }, ...(U?[{ wch:18 }]:[])];
  styleSheet(ws, {
    numCols, subRows:[1], headerRow:3, dataStart, dataEnd, totalRow,
    moneyCols: [3, 4, ...valLabels.map((_, i) => 5 + i), 5 + valLabels.length, 6 + valLabels.length, ...(U?[7 + valLabels.length]:[])], usdCols:U?[7 + valLabels.length]:[],
    theme, rowGroups, groupDisplayCol: 2, codeCol: 0,
  });
  XLSX.utils.book_append_sheet(wb, ws, clean(monthShortLabel(month)));
  XLSX.writeFile(wb, `QS_${clean(monthShortLabel(month)).replace(/[^\dA-Za-zก-๙]/g,"")}_${project.name.replace(/\s+/g,"_")}_${todayStr()}.xlsx`);
}

// ─── ชีตรวม: "ของเข้ารายเดือน (แผน + PO จริง)" ───────────────────────────────
//  ใช้ร่วมทั้ง Export จัดซื้อ + บัญชี — ต่อ Acc. Code: ต้นทุน (Tender Cost / Balance
//  Pending PO / Stock / Balance Cost) + แต่ละเดือนแยก 3 ช่อง: จ่าย(เขียว)=จ่ายแล้ว ·
//  ปกติ(ดำ)=รับ/PO รอเข้า (ส้ม=ล่าช้า) · แผน(แดง)=ยังไม่เป็น PO + TOTAL แถว/คอลัมน์
function addIncomingMonthlySheet(wb, { project, poEntries, incomingPlan=[], tenderCosts={}, additions={}, extraItems=[], hiddenAccounts=[], theme, backSheet="สรุป" }) {
  const plansArr = Array.isArray(incomingPlan) ? incomingPlan : [];
  const acctList = exportAccountList(extraItems, hiddenSafeForPO(hiddenAccounts, poEntries), poEntries, plansArr);
  const nameOf   = (code) => acctList.find(a=>a.code===code)?.name || ACCOUNTS.find(a=>a.code===code)?.name || "";
  const combinedB = buildCombinedBudget(tenderCosts, additions);
  const today = todayStr();
  const lateOf = (r) => !!(r.planDate && r.planDate < today && !r.actualDate);
  const mCell = {};
  const bucket = (code, mk) => { const c=(mCell[code]=mCell[code]||{}); return (c[mk]=c[mk]||{paid:0,recv:0,po:0,poLate:false,plan:0,planLate:false}); };
  plansArr.forEach(pl => poItems(pl).forEach(it => (it.rounds||[]).forEach(r => {
    const a=parseFloat(r.planAmount)||0; if(!a) return;
    const cc=bucket(it.code,(r.planDate||pl.date||"").slice(0,7)); cc.plan+=a; if(lateOf(r)) cc.planLate=true;
  })));
  poEntries.forEach(p => poItems(p).forEach(it => (it.rounds||[]).forEach(r => {
    if (roundReceived(r)) {
      const a=parseFloat(r.actualAmount)||0; if(!a) return;
      const cc=bucket(it.code, r.actualDate.slice(0,7));
      if (roundPaid(p,r)) cc.paid+=a; else cc.recv+=a;
    } else {
      const a=parseFloat(r.actualAmount)||parseFloat(r.planAmount)||0; if(!a) return; // ยอดจริงที่กรอกไว้มาก่อน (ตรงกับหน้ารายละเอียด) ไม่มีค่อยใช้แผน
      const cc=bucket(it.code,(r.actualDate||r.planDate||p.date||"").slice(0,7)); cc.po+=a; if(lateOf(r)) cc.poLate=true;
    }
  })));
  const mCodes = Object.keys(mCell).sort();
  const mMonths = [...new Set(mCodes.flatMap(c=>Object.keys(mCell[c])))].filter(Boolean).sort();
  if (!mCodes.length || !mMonths.length) return;
  const cellOf = (code,mk) => mCell[code]?.[mk] || null;
  const cellTot = (c) => c ? (c.paid+c.recv+c.po+c.plan) : 0;
  const budgetOf    = (code) => parseFloat(combinedB[code])||0;
  const committedOf = (code) => poEntries.reduce((s,p)=>s+poAmountForCode(p,code),0);
  const plannedOf   = (code) => plansArr.reduce((s,pl)=>s+poAmountForCode(pl,code),0);
  const stockOf     = (code) => poEntries.reduce((s,p)=>s+poItems(p).filter(it=>it.code===code).reduce((ss,it)=>ss+(parseFloat(it.store)||0),0),0);
  const takeoffOf = (code) => [...poEntries, ...plansArr].reduce((s,p)=>s+poItems(p).filter(it=>it.code===code).reduce((ss,it)=>ss+(parseFloat(it.takeoff)||0),0),0);
  // เดือนละ 1 คอลัมน์ (ไม่แยก 3 ช่อง) — ในช่องใส่รายการแบบมีป้ายกำกับ จ่าย/รับ/รอเข้า/แผน
  // ช่วยให้ตารางไม่กว้างเกินเมื่อมีหลายเดือน (เช่น 2 ปี = 24 คอลัมน์ แทน 72)
  const cellLines = (c) => {
    if (!c) return [];
    const out = [];
    if (c.paid>0) out.push(`จ่าย ${fmt(c.paid)}`);
    if (c.recv>0) out.push(`รับ ${fmt(c.recv)}`);
    if (c.po>0)   out.push(`รอเข้า ${fmt(c.po)}${c.poLate?" ⚠":""}`);
    if (c.plan>0) out.push(`แผน ${fmt(c.plan)} *${c.planLate?" ⚠":""}`);
    return out;
  };
  // ถ้าช่องมีชนิดเดียวลงสีตามชนิด (จ่าย=เขียว/รับ,รอเข้า=ดำ/แผน=แดง/ล่าช้า=ส้ม); ถ้าปนกันใช้ดำ
  const cellColor = (c) => {
    if (!c) return null;
    const n = [c.paid>0, c.recv>0, c.po>0, c.plan>0].filter(Boolean).length;
    if (n !== 1) return (c.poLate || c.planLate) ? "D97706" : null;
    if (c.paid>0) return "10B981";
    if (c.plan>0) return c.planLate ? "D97706" : "EF4444";
    if (c.po>0)   return c.poLate ? "D97706" : "1F2937";
    return "1F2937";
  };
  // คอลัมน์ต้นทุน: Tender Cost · Take off · Stock · Issue PO · Pending PO + เดือน (1 ช่อง/เดือน) + TOTAL + Balance Cost
  const header = ["Acc. Code","Acc. Name",`Tender Cost (รวมเผื่อ ${WASTE_LBL})`,"Take off","Stock","Issue PO","Pending PO"];
  mMonths.forEach(mk => header.push(monthShortLabel(mk)));
  header.push("TOTAL","Balance Cost");
  const rows = [
    [`ของเข้ารายเดือน (แผน + PO จริง) — ${project.name}`],
    [`เดือนละ 1 ช่อง (มีป้ายกำกับ) — จ่าย=จ่ายแล้ว(เขียว) · รับ=รับของแล้ว · รอเข้า=PO ยังไม่รับ(⚠=ล่าช้า) · แผน=ยังไม่เป็น PO(แดง, มี *) · Issue PO = PO ที่ยื่นจริง · Pending PO = งบ−Stock−PO−แผน · Balance Cost = Tender Cost−Stock−Issue PO · Export: ${new Date().toLocaleDateString("th-TH")}`],
    [],
    header,
  ];
  const dataStart = rows.length, monthColStart = 7, totalCol = 7 + mMonths.length, balPOcol = totalCol + 1, numCols = balPOcol + 1;
  const lineCount = {}; // จำนวนบรรทัดสูงสุดต่อแถว → ใช้ตั้งความสูงแถว
  mCodes.forEach((code, ri) => {
    const budget=budgetOf(code), committed=committedOf(code), stock=stockOf(code), planned=plannedOf(code), takeoff=takeoffOf(code);
    const row = [code, nameOf(code), budget, takeoff, stock, committed, budget-stock-committed-planned];
    let maxLines = 1;
    mMonths.forEach(mk => { const l=cellLines(cellOf(code,mk)); maxLines=Math.max(maxLines, l.length||1); row.push(l.length?l.join("\n"):"-"); });
    row.push(mMonths.reduce((s,mk)=>s+cellTot(cellOf(code,mk)),0), budget-stock-committed); // Balance Cost = Tender Cost − Stock − Issue PO
    rows.push(row);
    lineCount[dataStart+ri] = maxLines;
  });
  const dataEnd = rows.length - 1;
  const sumOf = (fn) => mCodes.reduce((s,c)=>s+fn(c),0);
  const totalArr = ["","TOTAL", sumOf(budgetOf), sumOf(takeoffOf), sumOf(stockOf), sumOf(committedOf), sumOf(c=>budgetOf(c)-stockOf(c)-committedOf(c)-plannedOf(c))];
  // แถว TOTAL: รวมยอดทั้งเดือนเป็นตัวเลขเดียว (ไม่แยกจ่าย/รอเข้า/แผน)
  mMonths.forEach(mk => { const colT = sumOf(c=>cellTot(cellOf(c,mk))); totalArr.push(colT>0 ? fmt(colT) : "-"); });
  totalArr.push(mCodes.reduce((s,c)=> s + mMonths.reduce((ss,mk)=>ss+cellTot(cellOf(c,mk)),0), 0), sumOf(c=>budgetOf(c)-stockOf(c)-committedOf(c))); // Balance Cost = Tender − Stock − Issue PO
  rows.push(totalArr);
  const totalRow = rows.length - 1;
  lineCount[totalRow] = 1;
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:12},{wch:34},{wch:16},{wch:15},{wch:14},{wch:15},{wch:16}, ...mMonths.map(()=>({wch:17})), {wch:18},{wch:16}];
  const moneyCols = [2,3,4,5,6, totalCol, balPOcol]; // ช่องเดือนเป็นข้อความ ไม่ใช่ตัวเลข
  styleSheet(ws, { numCols, subRows:[1], headerRow:3, dataStart, dataEnd, totalRow, moneyCols, theme });
  // ลงสี + wrapText ช่องเดือน (ข้อความหลายบรรทัด) + ตั้งความสูงแถวตามจำนวนบรรทัด
  if (!ws["!rows"]) ws["!rows"] = [];
  const paintMonth = (r) => {
    mMonths.forEach((mk,mi)=>{
      const c = cellOf(mCodes[r-dataStart], mk);
      const ref = XLSX.utils.encode_cell({r, c:monthColStart+mi});
      if (!ws[ref]) return;
      const rgb = cellColor(c);
      ws[ref].s = { ...(ws[ref].s||{}),
        font: { ...((ws[ref].s||{}).font||{}), ...(rgb?{color:{rgb}}:{}) },
        alignment: { ...((ws[ref].s||{}).alignment||{}), horizontal:"right", vertical:"top", wrapText:true } };
    });
    const n = lineCount[r] || 1;
    if (n > 1) ws["!rows"][r] = { hpx: Math.max(19, n*14) };
  };
  for (let r=dataStart; r<=dataEnd; r++) {
    paintMonth(r);
    [6, balPOcol].forEach(cc => { const ref=XLSX.utils.encode_cell({r,c:cc}); if (ws[ref] && typeof ws[ref].v==="number" && ws[ref].v<0) ws[ref].s = { ...(ws[ref].s||{}), font:{ ...(ws[ref].s?.font||{}), color:{rgb:"DC2626"}, bold:true } }; });
  }
  // แถว TOTAL: ช่องเดือนเป็นยอดรวมเดียว — ตัวหนา ชิดขวา
  mMonths.forEach((mk,mi)=>{ const ref=XLSX.utils.encode_cell({r:totalRow,c:monthColStart+mi}); if(ws[ref]) ws[ref].s={ ...(ws[ref].s||{}), font:{ ...((ws[ref].s||{}).font||{}), bold:true }, alignment:{ ...((ws[ref].s||{}).alignment||{}), horizontal:"right", vertical:"center" } }; });
  xBackLink(ws, 2, numCols-1, backSheet);
  XLSX.utils.book_append_sheet(wb, ws, "ของเข้ารายเดือน");
}

// ─── ชีต "ตารางรวมเดือน" (แบบหน้าบัญชี) — mirror AccountingMatrixTab ──────────
//  ต้นทุน (Tender/Balance Pending PO/Stock/Balance Cost) + กลุ่มเดือน Incoming/Received
//  (รับจริง=ดำ · แผน/PO รอเข้า=แดง) + Payment Plan รายเดือน + สรุป PO (Total PO/PO Balance)
function addAccountingMatrixSheet(wb, { project, poEntries, incomingPlan=[], tenderCosts={}, additions={}, extraItems=[], hiddenAccounts=[], theme, backSheet="Summary" }) {
  const accounts = exportAccountList(extraItems, hiddenSafeForPO(hiddenAccounts, poEntries), poEntries, incomingPlan);
  const combined = buildCombinedBudget(tenderCosts, additions);
  const plansArr = Array.isArray(incomingPlan) ? incomingPlan : [];
  const committedByCode = {}, stockByCode = {}, plannedByCode = {}, actual = {}, incoming = {}, payplan = {};
  const bump = (obj, code, mk, amt) => { if (!mk || !amt) return; (obj[code]=obj[code]||{}); obj[code][mk]=(obj[code][mk]||0)+amt; };
  plansArr.forEach(pl => poItems(pl).forEach(it => {
    plannedByCode[it.code] = (plannedByCode[it.code]||0) + (parseFloat(it.amount)||0);
    (it.rounds||[]).forEach(r => { const a=parseFloat(r.planAmount)||0; if(a>0) bump(incoming, it.code, (r.planDate||pl.date||"").slice(0,7), a); });
  }));
  poEntries.forEach(p => { poItems(p).forEach(it => {
    committedByCode[it.code] = (committedByCode[it.code]||0)+(parseFloat(it.amount)||0);
    stockByCode[it.code] = (stockByCode[it.code]||0)+(parseFloat(it.store)||0);
    (it.rounds||[]).forEach(r => {
      if (roundReceived(r)) bump(actual, it.code, r.actualDate.slice(0,7), parseFloat(r.actualAmount)||0);
      else { const a=parseFloat(r.actualAmount)||parseFloat(r.planAmount)||0; if(a>0) bump(incoming, it.code, (r.actualDate||r.planDate||p.date||"").slice(0,7), a); }
    });
  }); poPayLines(p).forEach(l => bump(payplan, l.code, l.month, l.amount||0)); });
  const monthsOf = (obj) => [...new Set(Object.values(obj).flatMap(m=>Object.keys(m)))].sort();
  const mgM = [...new Set([...monthsOf(incoming), ...monthsOf(actual)])].sort();
  const payM = monthsOf(payplan);
  const rowsData = accounts.map(a => {
    const budget = parseFloat(combined[a.code])||0, committed = committedByCode[a.code]||0, stock = stockByCode[a.code]||0, planned = plannedByCode[a.code]||0;
    const mgRow = mgM.map(mk => { const av=actual[a.code]?.[mk]||0, pv=incoming[a.code]?.[mk]||0; return { eff: av+pv, real: pv===0, hasRecv: av>0 }; }); // รวมรับจริง+ยังไม่เข้า (ไม่ให้ตกหล่นเมื่อเดือนเดียวมีทั้งคู่) · เขียว=รับครบ, เหลือง=รับบางส่วน, แดง=ยังไม่เข้า
    const pyRow = payM.map(mk => payplan[a.code]?.[mk]||0);
    return { a, budget, committed, stock, planned, balPO:budget-committed, balCost:budget-stock-committed-planned, balPOout:budget-stock-committed,
      mgRow, pyRow, mgTot:mgRow.reduce((s,c)=>s+c.eff,0), pyTot:pyRow.reduce((s,x)=>s+x,0) };
  }).filter(r => r.budget||r.committed||r.stock||r.mgTot||r.pyTot);
  if (!rowsData.length) return;
  const header = ["Acc. Code","Acc. Name",`Tender Cost (รวมเผื่อ ${WASTE_LBL})`,"Balance Pending PO","Stock","Pending PO",
    ...mgM.map(mk=>`${monthShortLabel(mk)} (เข้า)`), "รวมเข้า",
    ...payM.map(mk=>`${monthShortLabel(mk)} (จ่าย)`), "รวมจ่าย", "Total PO", "Balance Cost"];
  const rows = [
    [`ตารางรวมเดือน — ${project.name}`],
    [`Incoming: รับครบ=เขียว · รับบางส่วน=เหลือง · แผน/PO รอเข้า=แดง · Pending PO = งบ − Stock − PO − แผน · Balance Cost = งบ − Stock − PO (ตรงกับหน้าจัดซื้อ) · Export: ${new Date().toLocaleDateString("th-TH")}`],
    [],
    header,
  ];
  const dataStart = rows.length;
  const mgStart = 6, mgTotCol = 6+mgM.length, payStart = mgTotCol+1, payTotCol = payStart+payM.length, poCol = payTotCol+1, poBalCol = poCol+1, numCols = poBalCol+1;
  const blank = (v) => v ? v : "-";
  rowsData.forEach(r => {
    rows.push([r.a.code, r.a.name, blank(r.budget), blank(r.balPO), blank(r.stock), blank(r.balCost),
      ...r.mgRow.map(c=>blank(c.eff)), blank(r.mgTot),
      ...r.pyRow.map(v=>blank(v)), blank(r.pyTot), blank(r.committed), blank(r.balPOout)]);
  });
  const dataEnd = rows.length - 1;
  const sumOf = (fn) => rowsData.reduce((s,r)=>s+fn(r),0);
  const totalArr = ["","TOTAL", sumOf(r=>r.budget), sumOf(r=>r.balPO), sumOf(r=>r.stock), sumOf(r=>r.balCost)];
  mgM.forEach((_,i)=>totalArr.push(sumOf(r=>r.mgRow[i]?.eff||0))); totalArr.push(sumOf(r=>r.mgTot));
  payM.forEach((_,i)=>totalArr.push(sumOf(r=>r.pyRow[i]||0)));    totalArr.push(sumOf(r=>r.pyTot));
  totalArr.push(sumOf(r=>r.committed), sumOf(r=>r.balPOout));
  rows.push(totalArr.map((v,i)=> i<2 ? v : blank(v)));
  const totalRow = rows.length - 1;
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:12},{wch:32},{wch:16},{wch:18},{wch:14},{wch:16},
    ...mgM.map(()=>({wch:14})), {wch:14}, ...payM.map(()=>({wch:14})), {wch:14}, {wch:16}, {wch:16}];
  const allMoney = [2,3,4,5, ...Array.from({length:numCols-6},(_,i)=>6+i)];
  styleSheet(ws, { numCols, subRows:[1], headerRow:3, dataStart, dataEnd, totalRow, moneyCols:allMoney, theme });
  const setColor = (r, col, rgb, bold) => { const ref=XLSX.utils.encode_cell({r,c:col}); if (ws[ref]) ws[ref].s = { ...(ws[ref].s||{}), font:{ ...((ws[ref].s||{}).font||{}), color:{rgb}, ...(bold?{bold:true}:{}) } }; };
  for (let r=dataStart; r<=dataEnd; r++) {
    const rd = rowsData[r-dataStart];
    rd.mgRow.forEach((c,i)=>{ if (c.eff) setColor(r, mgStart+i, c.real ? "059669" : (c.hasRecv ? "D97706" : "EF4444"), c.real || c.hasRecv); }); // เขียว=รับครบ · เหลือง=รับบางส่วน · แดง=ยังไม่เข้า (ให้ตรงกับสีในแอป)
    [5, poBalCol].forEach(cc => { const ref=XLSX.utils.encode_cell({r,c:cc}); if (ws[ref] && typeof ws[ref].v==="number" && ws[ref].v<0) setColor(r, cc, "DC2626", true); });
  }
  xBackLink(ws, 2, numCols-1, backSheet);
  XLSX.utils.book_append_sheet(wb, ws, "ตารางรวมเดือน");
}

// ─── Procurement: PO tracking export ───────────────────────────────────────
function exportProcurementExcel(project, poEntries, incomingPlan=[], tenderCosts={}, additions={}, extraItems=[], hiddenAccounts=[]) {
  const wb = XLSX.utils.book_new();
  const theme = { main:"F59E0B", dark:"B45309" };
  const rate = exportRate(project); const U = rate > 0;   // U = ใส่คอลัมน์ USD ไหม

  // หน้าแรก = สรุป (Dashboard): การ์ดตัวเลข + กราฟยอดสั่งซื้อรายเดือน
  const dPaid = poEntries.reduce((s,p)=> s + poPaidAmount(p), 0);
  const dTotal = poEntries.reduce((s,p)=> s + poTotal(p), 0);
  const dMonths = [...new Set(poEntries.map(p => (p.date||"").slice(0,7)).filter(Boolean))].sort();
  const dash = addDashboardSheet(wb, "สรุป", {
    title: `สรุปจัดซื้อ (PO) — ${project.name}`,
    subtitle: `Export: ${new Date().toLocaleDateString("th-TH")}`,
    theme,
    cards: [
      { label:"จำนวน PO",     value: poEntries.length },
      { label:"มูลค่ารวม",     value: dTotal, money:true },
      { label:"จ่ายแล้ว",      value: dPaid, money:true },
      { label:"คงค้างจ่าย",    value: Math.max(dTotal - dPaid, 0), money:true },
    ],
    chartTitle: "กราฟ: ยอดสั่งซื้อรายเดือน (ตามวันเปิด PO)",
    items: dMonths.map(m => ({ label: monthShortLabel(m), value: poEntries.filter(p=>(p.date||"").slice(0,7)===m).reduce((s,p)=>s+poTotal(p),0) })),
  });

  // Sheet 1 — every PO line → วางต่อท้ายหน้า "สรุป" (ชีตเดียวกัน)
  const rows1 = [[`รายการ PO — ${project.name}`], [`Export: ${new Date().toLocaleDateString("th-TH")}  ·  ทั้งหมด ${poEntries.length} PO${U?`  ·  อัตราแลกเปลี่ยน ${rate} บาท/USD`:""}`], []];
  rows1.push(["วันเปิด PO","Acc. Code","Account Name","Supplier","PO No.","มูลค่า (THB)",...(U?["มูลค่า (USD)"]:[]),"สถานะ PO","ของเข้า (แผน→จริง)","วันครบกำหนดจ่าย","สถานะจ่ายเงิน","หมายเหตุ"]);
  const dataStart1 = rows1.length;
  let grand1 = 0;
  const rowGroups1 = [];
  poEntries.slice().sort((a,b)=>(a.date||"").localeCompare(b.date||"")).forEach(p => {
    poItems(p).forEach(it => {
      // แต่ละบรรทัดแยกตาม Acc. Code → ของเข้า/วันครบกำหนด/สถานะจ่าย คิดเฉพาะงวดของ item นี้
      const pItem = { ...p, items:[it] };
      const pay = paymentStatus(pItem);
      const _rd = poRounds(pItem);
      const deliveryStr = _rd.map((r,i) => `${_rd.length>1?`งวด${i+1}: `:""}${r.plan||"—"} → ${r.actual? "รับ "+r.actual : "รอ"}`).join("\n") || "-";
      const acc = accountOf(it.code);
      const amount = parseFloat(it.amount) || 0;
      rows1.push([p.date, it.code, acc?.name||"", itemSupplierName(p), poNumbersLabel(p), amount, ...(U?[toUsd(amount,rate)]:[]), p.status, deliveryStr, poNextDueDate(pItem)||"-", PAYMENT_LABEL[pay], p.notes||""]);
      rowGroups1.push(acc?.group || "-");
      grand1 += amount;
    });
  });
  const dataEnd1 = rows1.length-1;
  rows1.push(["","","","","TOTAL", grand1, ...(U?[toUsd(grand1,rate)]:[]),"","","","",""]);
  const totalRow1 = rows1.length-1;
  // เขียนตาราง PO ต่อท้าย dashboard ในชีต "สรุป" (เว้น 1 บรรทัด) แล้วจัดสไตล์ตามออฟเซ็ตแถว
  const poStart = dash.nextRow + 1;
  XLSX.utils.sheet_add_aoa(dash.ws, rows1, { origin: { r: poStart, c: 0 } });
  styleSheet(dash.ws, { numCols:11+(U?1:0),
    titleRow: poStart, subRows:[poStart+1], headerRow: poStart+3,
    dataStart: poStart+dataStart1, dataEnd: poStart+dataEnd1, totalRow: poStart+totalRow1,
    moneyCols:U?[5,6]:[5], usdCols:U?[6]:[], centerCols:U?[7,10]:[6,9], statusCols:U?[7,10]:[6,9], theme, rowGroups:rowGroups1 });
  delete dash.ws["!freeze"];   // มี dashboard อยู่ด้านบน จึงไม่ freeze
  dash.ws["!cols"] = [{wch:12},{wch:10},{wch:34},{wch:22},{wch:16},{wch:16},...(U?[{wch:16}]:[]),{wch:12},{wch:30},{wch:16},{wch:16},{wch:28}];
  // ขยายความสูงแถวตามจำนวนงวดในคอลัมน์ "ของเข้า" (multiline) + จัดชิดบน
  { const dc = U?8:7; for (let r=dataStart1; r<=dataEnd1; r++){ const R=poStart+r; const ref=XLSX.utils.encode_cell({r:R,c:dc}); const v=dash.ws[ref]?.v; if(typeof v==="string"){ const n=v.split("\n").length; if(dash.ws[ref].s) dash.ws[ref].s.alignment={...(dash.ws[ref].s.alignment||{}),vertical:"top",wrapText:true}; if(n>1) dash.ws["!rows"][R]={hpx:Math.max(19,n*14)}; } } }

  // ชีต "ของเข้ารายเดือน (แผน + PO จริง)" — ต้นทุน + เดือนแยก 3 ช่อง จ่าย/ปกติ/แผน (สี)
  addIncomingMonthlySheet(wb, { project, poEntries, incomingPlan, tenderCosts, additions, extraItems, hiddenAccounts, theme, backSheet: "สรุป" });

  // แยกรายเดือนแบบละเอียด (หนึ่งชีตต่อเดือน) — เอา "สรุปสถานะ" และ "รายเดือน (สรุปกลุ่ม)" ออกแล้ว
  const poMonths = [...new Set(poEntries.map(p => (p.date||"").slice(0,7)).filter(Boolean))].sort();
  const monthLinks = [];   // เก็บชื่อชีตรายเดือนไว้ทำลิงก์บนหน้าสรุป
  if (poMonths.length) {
    // รายเดือนแบบละเอียด (Acc.Code / Supplier / PO No.) หนึ่งชีตต่อเดือน
    const clean = (s) => String(s).replace(/[\\/?*[\]:]/g, "-").slice(0, 28);
    const usedNames = {};
    poMonths.forEach(m => {
      const rows = [
        [`PO รายเดือน ${monthShortLabel(m)} — ${project.name}`],
        [`ตามวันเปิด PO · Export: ${new Date().toLocaleDateString("th-TH")}`],
        [],
        ["Acc. Code", "Account Name", "Group", "Supplier", "PO No.", "วันเปิด PO", "มูลค่า (THB)", ...(U?["มูลค่า (USD)"]:[]), "สถานะ PO", "สถานะจ่ายเงิน"],
      ];
      const dataStart = rows.length;
      const rowGroups = [];
      let grand = 0;
      poEntries.filter(p => (p.date||"").slice(0,7) === m)
        .sort((a,b)=>(a.date||"").localeCompare(b.date||""))
        .forEach(p => {
          const pay = paymentStatus(p);
          poItems(p).forEach(it => {
            const acc = accountOf(it.code);
            const amount = parseFloat(it.amount) || 0;
            rows.push([it.code, acc?.name||"", acc?.group||"-", itemSupplierName(p), poNumbersLabel(p), p.date, amount, ...(U?[toUsd(amount,rate)]:[]), p.status, PAYMENT_LABEL[pay]]);
            rowGroups.push(acc?.group||"-");
            grand += amount;
          });
        });
      if (rows.length === dataStart) return; // เดือนนี้ไม่มี PO
      const dataEnd = rows.length-1;
      rows.push(["", "", "", "", "", "TOTAL", grand, ...(U?[toUsd(grand,rate)]:[]), "", ""]);
      const totalRow = rows.length-1;
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = [{wch:10},{wch:32},{wch:14},{wch:22},{wch:16},{wch:12},{wch:16},...(U?[{wch:16}]:[]),{wch:12},{wch:16}];
      styleSheet(ws, { numCols:9+(U?1:0), subRows:[1], headerRow:3, dataStart, dataEnd, totalRow,
        moneyCols:U?[6,7]:[6], usdCols:U?[7]:[], centerCols:U?[8,9]:[7,8], statusCols:U?[8,9]:[7,8], theme, rowGroups, groupDisplayCol:2 });
      let nm = clean(monthShortLabel(m));
      if (usedNames[nm]) { usedNames[nm] += 1; nm = clean(`${nm} ${usedNames[nm]}`); } else usedNames[nm] = 1;
      xBackLink(ws, 2, (9+(U?1:0))-1, "สรุป");   // ลิงก์กลับหน้าสรุป
      XLSX.utils.book_append_sheet(wb, ws, nm);
      monthLinks.push({ text: monthShortLabel(m), sheet: nm });
    });
  }
  // ลิงก์ไปยังชีตรายเดือน วางไว้ใต้ตาราง PO ในหน้าสรุป — คลิกเพื่อไล่ที่มาของตัวเลข
  if (monthLinks.length) {
    const navR = poStart + rows1.length + 1;
    XLSX.utils.sheet_add_aoa(dash.ws, [["🔗 ไปดูรายละเอียดรายเดือน (คลิกเพื่อดูที่มา):"]], { origin:{ r:navR, c:0 } });
    dash.ws[XLSX.utils.encode_cell({ r:navR, c:0 })].s = { font:{ bold:true, sz:10.5, color:{rgb:theme.dark}, name:"Tahoma" } };
    xLinkRow(dash.ws, navR+1, monthLinks);
  }

  XLSX.writeFile(wb, `Procurement_PO_${project.name.replace(/\s+/g,"_")}_${todayStr()}.xlsx`);
}

// ─── Accounting: full financial export ─────────────────────────────────────
function exportAccountingExcel(project, tenderCosts, additions, poEntries, extraItems=[], hiddenAccounts=[], incomingPlan=[]) {
  const wb = XLSX.utils.book_new();
  const theme = { main:"10B981", dark:"047857" };
  const rate = exportRate(project); const U = rate > 0;   // U = ใส่คอลัมน์ USD ไหม
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);
  const accounts = exportAccountList(extraItems, hiddenSafeForPO(hiddenAccounts, poEntries), poEntries);

  // Sheet 1 — Budget vs Committed vs Variance per Acc. Code
  const rows1 = [[`สรุปงบประมาณ — ${project.name}`], [`พื้นที่ ${project.area||"-"} ft²  ·  แผง ${project.panels||"-"}  ·  Export: ${new Date().toLocaleDateString("th-TH")}`], []];
  rows1.push(U
    ? ["Acc. Code","Account Name","Group","งบประมาณ (Budget)","Budget (USD)","Committed (PO)","Committed (USD)","ส่วนต่าง","% ใช้ไป","สถานะ"]
    : ["Acc. Code","Account Name","Group","งบประมาณ (Budget)","Committed (PO)","ส่วนต่าง","% ใช้ไป","สถานะ"]);
  const dataStart1 = rows1.length;
  let gB=0, gC=0;
  const rowGroups1 = [];
  accounts.forEach(a => {
    const budget    = parseFloat(combinedBudget[a.code]) || 0;
    const committed = poEntries.reduce((s,p)=>s+poAmountForCode(p,a.code),0);
    if (budget<=0 && committed<=0) return;
    const variance = budget - committed;
    const pctUsed  = budget>0 ? committed/budget : (committed>0 ? 9.99 : 0);
    const status   = committed>0 && budget<=0 ? "ไม่มีงบ" : committed>budget && budget>0 ? "เกินงบ" : committed>0 ? "OK" : budget>0 ? "ยังไม่ PO" : "-";
    rows1.push(U
      ? [a.code, a.name, a.group, budget, toUsd(budget,rate), committed, toUsd(committed,rate), variance, pctUsed, status]
      : [a.code, a.name, a.group, budget, committed, variance, pctUsed, status]);
    rowGroups1.push(a.group);
    gB += budget; gC += committed;
  });
  const dataEnd1 = rows1.length-1;
  rows1.push(U
    ? ["","TOTAL","",gB,toUsd(gB,rate),gC,toUsd(gC,rate),gB-gC,gB>0?gC/gB:0,""]
    : ["","TOTAL","",gB,gC,gB-gC,gB>0?gC/gB:0,""]);
  const totalRow1 = rows1.length-1;
  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  ws1["!cols"] = U
    ? [{wch:12},{wch:38},{wch:16},{wch:16},{wch:16},{wch:16},{wch:16},{wch:14},{wch:10},{wch:12}]
    : [{wch:12},{wch:38},{wch:16},{wch:16},{wch:16},{wch:14},{wch:10},{wch:12}];
  styleSheet(ws1, { numCols:8+(U?2:0), subRows:[1], headerRow:3, dataStart:dataStart1, dataEnd:dataEnd1, totalRow:totalRow1,
    moneyCols:U?[3,4,5,6,7]:[3,4,5], usdCols:U?[4,6]:[], pctCols:U?[8]:[6], centerCols:U?[9]:[7], theme, rowGroups:rowGroups1, groupDisplayCol:2 });
  // Flag over-budget rows in red so they jump out without opening the app
  const varC1 = U?7:5, stC1 = U?9:7;
  for (let r=dataStart1; r<=dataEnd1; r++) {
    const varRef = XLSX.utils.encode_cell({r,c:varC1});
    const stRef  = XLSX.utils.encode_cell({r,c:stC1});
    if (ws1[varRef] && typeof ws1[varRef].v === "number" && ws1[varRef].v < 0) {
      ws1[varRef].s = { ...ws1[varRef].s, font:{...ws1[varRef].s.font, color:{rgb:"DC2626"}, bold:true} };
    }
    if (ws1[stRef] && ws1[stRef].v === "เกินงบ") {
      ws1[stRef].s = { ...ws1[stRef].s, font:{...ws1[stRef].s.font, color:{rgb:"DC2626"}, bold:true} };
    }
  }
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");

  // Sheet 2 — every PO line, full date + status detail
  const rows2 = [[`รายการ PO ทั้งหมด — ${project.name}`], [`ทั้งหมด ${poEntries.length} PO  ·  Export: ${new Date().toLocaleDateString("th-TH")}`], []];
  rows2.push(["วันเปิด PO","Acc. Code","Account Name","Group","Supplier","PO No.","มูลค่า (THB)",...(U?["มูลค่า (USD)"]:[]),"สถานะ","ของเข้า (แผน→จริง)","วันครบกำหนดจ่าย","สถานะจ่าย"]);
  const dataStart2 = rows2.length;
  let grand2 = 0;
  const rowGroups2 = [];
  poEntries.slice().sort((a,b)=>(a.date||"").localeCompare(b.date||"")).forEach(p => {
    poItems(p).forEach(it => {
      // แต่ละบรรทัดแยกตาม Acc. Code → ของเข้า/วันครบกำหนด/สถานะจ่าย คิดเฉพาะงวดของ item นี้
      const pItem = { ...p, items:[it] };
      const pay = paymentStatus(pItem);
      const _rd2 = poRounds(pItem);
      const deliveryStr = _rd2.map((r,i) => `${_rd2.length>1?`งวด${i+1}: `:""}${r.plan||"—"} → ${r.actual? "รับ "+r.actual : "รอ"}`).join("\n") || "-";
      const acc = accountOf(it.code);
      const amount = parseFloat(it.amount) || 0;
      rows2.push([p.date, it.code, acc?.name||"", acc?.group||"", itemSupplierName(p), poNumbersLabel(p), amount, ...(U?[toUsd(amount,rate)]:[]), p.status, deliveryStr, poNextDueDate(pItem)||"-", PAYMENT_LABEL[pay]]);
      rowGroups2.push(acc?.group || "-");
      grand2 += amount;
    });
  });
  const dataEnd2 = rows2.length-1;
  rows2.push(["","","","","","TOTAL", grand2, ...(U?[toUsd(grand2,rate)]:[]),"","","",""]);
  const totalRow2 = rows2.length-1;
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2["!cols"] = [{wch:12},{wch:10},{wch:34},{wch:14},{wch:22},{wch:16},{wch:16},...(U?[{wch:16}]:[]),{wch:12},{wch:30},{wch:16},{wch:16}];
  styleSheet(ws2, { numCols:11+(U?1:0), subRows:[1], headerRow:3, dataStart:dataStart2, dataEnd:dataEnd2, totalRow:totalRow2,
    moneyCols:U?[6,7]:[6], usdCols:U?[7]:[], centerCols:U?[8,11]:[7,10], theme, rowGroups:rowGroups2, groupDisplayCol:3 });
  // "ของเข้า (แผน→จริง)": หนึ่งงวดต่อบรรทัด — ตั้ง wrapText + ความสูงแถวตามจำนวนบรรทัด
  { const dc = U?9:8; if(!ws2["!rows"]) ws2["!rows"]=[]; for (let R=dataStart2; R<=dataEnd2; R++){ const ref=XLSX.utils.encode_cell({r:R,c:dc}); const v=ws2[ref]?.v; if(typeof v==="string"){ const n=v.split("\n").length; if(ws2[ref].s) ws2[ref].s.alignment={...(ws2[ref].s.alignment||{}),vertical:"top",wrapText:true}; if(n>1) ws2["!rows"][R]={hpx:Math.max(19,n*14)}; } } }
  XLSX.utils.book_append_sheet(wb, ws2, "PO Entries");

  // Sheet 3 — roll-up by Group
  const rows3 = [[`สรุปตามกลุ่ม — ${project.name}`], [], (U
    ? ["Group","Budget","Budget (USD)","Committed","Committed (USD)","ส่วนต่าง","% ใช้ไป"]
    : ["Group","Budget","Committed","ส่วนต่าง","% ใช้ไป"])];
  const dataStart3 = 3;
  let g3B=0, g3C=0;
  GROUPS.forEach(g => {
    const codes = accounts.filter(a=>a.group===g).map(a=>a.code);
    const b  = codes.reduce((s,c)=>s+(parseFloat(combinedBudget[c])||0),0);
    const c2 = poEntries.reduce((s,p)=>s+poItems(p).filter(it=>codes.includes(it.code)).reduce((s2,it)=>s2+(parseFloat(it.amount)||0),0),0);
    if (b<=0 && c2<=0) return;
    rows3.push(U ? [g,b,toUsd(b,rate),c2,toUsd(c2,rate),b-c2,b>0?c2/b:0] : [g,b,c2,b-c2,b>0?c2/b:0]);
    g3B += b; g3C += c2;
  });
  const dataEnd3 = rows3.length-1;
  rows3.push(U ? ["TOTAL",g3B,toUsd(g3B,rate),g3C,toUsd(g3C,rate),g3B-g3C,g3B>0?g3C/g3B:0] : ["TOTAL",g3B,g3C,g3B-g3C,g3B>0?g3C/g3B:0]);
  const totalRow3 = rows3.length-1;
  const ws3 = XLSX.utils.aoa_to_sheet(rows3);
  ws3["!cols"] = U ? [{wch:18},{wch:16},{wch:16},{wch:16},{wch:16},{wch:14},{wch:10}] : [{wch:18},{wch:16},{wch:16},{wch:14},{wch:10}];
  styleSheet(ws3, { numCols:5+(U?2:0), headerRow:2, dataStart:dataStart3, dataEnd:dataEnd3, totalRow:totalRow3, moneyCols:U?[1,2,3,4,5]:[1,2,3], usdCols:U?[2,4]:[], pctCols:U?[6]:[4], theme });
  XLSX.utils.book_append_sheet(wb, ws3, "By Group");

  // Sheet 4 — monthly cash-flow: how much budget was added and how much got
  // committed (PO'd) each month, plus the running cumulative totals, so
  // Accounting can see the trend over time rather than just a snapshot
  const additionMonths = Object.keys(additions||{}).filter(k=>!k.startsWith("$"));
  const poEntryMonths  = poEntries.map(p=>(p.date||"").slice(0,7)).filter(Boolean);
  const allMonths = [...new Set([...additionMonths, ...poEntryMonths])].sort();
  if (allMonths.length) {
    const rows4 = [[`รายเดือน — ${project.name}`], [`Export: ${new Date().toLocaleDateString("th-TH")}`], []];
    rows4.push(U
      ? ["เดือน","Budget เพิ่มเดือนนี้","งบสะสม","งบสะสม (USD)","Committed เดือนนี้","Committed สะสม","Committed สะสม (USD)","% ใช้ไปสะสม"]
      : ["เดือน","Budget เพิ่มเดือนนี้","งบสะสม","Committed เดือนนี้","Committed สะสม","% ใช้ไปสะสม"]);
    const dataStart4 = rows4.length;
    const baselineTotal = accounts.reduce((s,a)=>s+withWaste(tenderCosts[a.code]),0);   // ราคาเดิม + เผื่อเศษ (รวมในงบ)
    // "Committed" ต้องนิยามให้ตรงกับชีตอื่น: ผลรวมยอด item เฉพาะ code ที่อยู่ในผังบัญชี
    // (ไม่ใช้ poTotal ทั้งใบ เพราะ PO อาจมี item ที่ code ไม่อยู่ในผัง ทำให้ยอดสะสมไม่ตรงกับ Sheet อื่น)
    const acctCodeSet = new Set(accounts.map(a=>a.code));
    const poCommitted = (p) => poItems(p).filter(it=>acctCodeSet.has(it.code)).reduce((s,it)=>s+(parseFloat(it.amount)||0),0);
    let cumB = baselineTotal, cumC = 0;
    allMonths.forEach(m => {
      const addedThisMonth     = accounts.reduce((s,a)=>s+monthAddValue(additions, m, a.code),0);
      const committedThisMonth = poEntries.filter(p=>(p.date||"").slice(0,7)===m).reduce((s,p)=>s+poCommitted(p),0);
      cumB += addedThisMonth;
      cumC += committedThisMonth;
      rows4.push(U
        ? [monthShortLabel(m), addedThisMonth, cumB, toUsd(cumB,rate), committedThisMonth, cumC, toUsd(cumC,rate), cumB>0?cumC/cumB:0]
        : [monthShortLabel(m), addedThisMonth, cumB, committedThisMonth, cumC, cumB>0?cumC/cumB:0]);
    });
    const dataEnd4 = rows4.length-1;
    const ws4 = XLSX.utils.aoa_to_sheet(rows4);
    ws4["!cols"] = U ? [{wch:14},{wch:18},{wch:16},{wch:16},{wch:18},{wch:16},{wch:18},{wch:12}] : [{wch:14},{wch:18},{wch:16},{wch:18},{wch:16},{wch:12}];
    styleSheet(ws4, { numCols:6+(U?2:0), subRows:[1], headerRow:3, dataStart:dataStart4, dataEnd:dataEnd4, moneyCols:U?[1,2,3,4,5,6]:[1,2,3,4], usdCols:U?[3,6]:[], pctCols:U?[7]:[5], theme });
    XLSX.utils.book_append_sheet(wb, ws4, "รายเดือน");
  }

  // ─── Sheet 5 + 6 — แผนจ่ายเงินรายเดือน (Payment forecast) ──────────────────
  // สำหรับบัญชี: มองไปข้างหน้าว่าเดือนไหนต้องเตรียมเงินจ่ายเท่าไหร่ จ่ายอะไร และ
  // จ่ายแบบไหน (เงินสด/เครดิต). ใช้ตัวช่วย poPayLines() ตัวเดียวกับหน้าแอพ เพื่อ
  // ให้ตัวเลขตรงกันและกันการนับซ้ำเมื่อ PO มีงวดส่งของซ้ำ.
  const payLines = poEntries.flatMap(poPayLines);
  if (payLines.length) {
    const monthKey = (l) => l.month || "9999-99";

    // ── แผนจ่าย — รายละเอียดแต่ละงวด (เอาตารางสรุปรายเดือนด้านบนออกแล้ว) ──
    const rowsC = [
      [`แผนจ่ายเงิน — ${project.name}`],
      [`รายละเอียดแต่ละงวด · เรียงตามเดือนที่ต้องจ่าย · ${payLines.length} งวด${U?`  ·  อัตราแลกเปลี่ยน ${rate} บาท/USD`:""}  ·  Export: ${new Date().toLocaleDateString("th-TH")}`],
      [],
    ];
    rowsC.push(U
      ? ["เดือนที่ต้องจ่าย","วันครบกำหนดจ่าย","Supplier","PO No.","Acc. Code","Account Name","วิธีจ่าย","วันรับของ (แผน/จริง)","ยอดต้องจ่าย (THB)","ยอดต้องจ่าย (USD)","สถานะจ่าย"]
      : ["เดือนที่ต้องจ่าย","วันครบกำหนดจ่าย","Supplier","PO No.","Acc. Code","Account Name","วิธีจ่าย","วันรับของ (แผน/จริง)","ยอดต้องจ่าย (THB)","สถานะจ่าย"]);
    const detStart = rowsC.length;
    const sortedD = payLines.slice().sort((a,b)=>
      (monthKey(a).localeCompare(monthKey(b))) ||
      ((a.payDate||"9999").localeCompare(b.payDate||"9999")) ||
      a.supplier.localeCompare(b.supplier));
    const rowGroupsD = [];
    let grandD = 0;
    sortedD.forEach(l => {
      const mk = monthKey(l);
      const label = mk==="9999-99" ? "ยังไม่ระบุ" : monthShortLabel(mk);
      const incomingTxt = l.incoming ? `${l.incoming}${l.incomingType?` (${l.incomingType})`:""}` : "-";
      rowsC.push(U
        ? [label, l.payDate||"-", l.supplier, l.poNo, l.code, l.accName, l.method, incomingTxt, l.amount, toUsd(l.amount,rate), PAYMENT_LABEL[l.status]]
        : [label, l.payDate||"-", l.supplier, l.poNo, l.code, l.accName, l.method, incomingTxt, l.amount, PAYMENT_LABEL[l.status]]);
      rowGroupsD.push(mk);
      grandD += l.amount;
    });
    const detEnd = rowsC.length-1;
    rowsC.push(U
      ? ["","","","","","","","TOTAL", grandD, toUsd(grandD,rate), ""]
      : ["","","","","","","","TOTAL", grandD, ""]);
    const detTotal = rowsC.length-1;
    const wsC = XLSX.utils.aoa_to_sheet(rowsC);
    wsC["!cols"] = U
      ? [{wch:18},{wch:16},{wch:22},{wch:16},{wch:14},{wch:30},{wch:16},{wch:20},{wch:18},{wch:18},{wch:16}]
      : [{wch:18},{wch:16},{wch:22},{wch:16},{wch:14},{wch:30},{wch:16},{wch:20},{wch:18},{wch:16}];
    styleSheet(wsC, { numCols:10+(U?1:0), subRows:[1], headerRow:3, dataStart:detStart, dataEnd:detEnd, totalRow:detTotal,
      moneyCols:U?[8,9]:[8], usdCols:U?[9]:[], statusCols:U?[10]:[9], theme, rowGroups:rowGroupsD, groupDisplayCol:0 });
    XLSX.utils.book_append_sheet(wb, wsC, "แผนจ่าย");
  }

  // ชีตใหม่: ของเข้ารายเดือน (3 ช่อง จ่าย/ปกติ/แผน มีสี) + ตารางรวมเดือน (แบบหน้าบัญชี)
  addIncomingMonthlySheet(wb, { project, poEntries, incomingPlan, tenderCosts, additions, extraItems, hiddenAccounts, theme, backSheet: "Summary" });
  addAccountingMatrixSheet(wb, { project, poEntries, incomingPlan, tenderCosts, additions, extraItems, hiddenAccounts, theme, backSheet: "Summary" });
  const hasInSheet = wb.SheetNames.includes("ของเข้ารายเดือน");
  const hasMxSheet = wb.SheetNames.includes("ตารางรวมเดือน");

  // ลิงก์นำทางใต้ตารางหน้า Summary — คลิกเพื่อไปดูที่มาของตัวเลขในแต่ละชีต
  {
    const acctLinks = [
      { text:"📦 PO Entries (รายการ PO)", sheet:"PO Entries" },
      { text:"🏷 By Group (ตามกลุ่ม)", sheet:"By Group" },
      ...(allMonths.length ? [{ text:"📅 รายเดือน", sheet:"รายเดือน" }] : []),
      ...(payLines.length ? [{ text:"💰 แผนจ่าย", sheet:"แผนจ่าย" }] : []),
      ...(hasInSheet ? [{ text:"📥 ของเข้ารายเดือน", sheet:"ของเข้ารายเดือน" }] : []),
      ...(hasMxSheet ? [{ text:"📄 ตารางรวมเดือน", sheet:"ตารางรวมเดือน" }] : []),
    ];
    const navR = totalRow1 + 2;
    XLSX.utils.sheet_add_aoa(ws1, [["🔗 ไปที่ชีต (ไล่ที่มาของตัวเลข):"]], { origin:{ r:navR, c:0 } });
    ws1[XLSX.utils.encode_cell({ r:navR, c:0 })].s = { font:{ bold:true, sz:10.5, color:{rgb:theme.dark}, name:"Tahoma" } };
    xLinkRow(ws1, navR+1, acctLinks);
  }

  XLSX.writeFile(wb, `Accounting_${project.name.replace(/\s+/g,"_")}_${todayStr()}.xlsx`);
}

// ─── Root ─────────────────────────────────────────────────────────────────────
// สร้างข้อความแบบตาราง (TSV) จากเซลล์ที่เลือก — จัดกลุ่มเป็นแถวตามตำแหน่งแนวตั้ง
// แล้วเรียงในแถวตามแนวนอน เพื่อวางลง Excel/Sheets แล้วลงช่องตรงกัน
function buildTSV(cells) {
  if (!cells || !cells.length) return "";
  const arr = cells.slice().sort((a, b) => (a.top - b.top) || (a.left - b.left));
  const rows = []; let cur = []; let top0 = null;
  for (const c of arr) {
    if (top0 === null || Math.abs(c.top - top0) <= 6) { cur.push(c); if (top0 === null) top0 = c.top; }
    else { rows.push(cur); cur = [c]; top0 = c.top; }
  }
  if (cur.length) rows.push(cur);
  return rows.map(r => r.slice().sort((a, b) => a.left - b.left).map(c => c.text).join("\t")).join("\n");
}

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

export default function App() {
  useLang();                                            // re-render ทั้งแอปเมื่อสลับภาษา
  const [session,  setSessionState] = useState(null);   // โหลดแบบ async ด้านล่าง
  const [authReady, setAuthReady]   = useState(false);  // true เมื่อเช็ค session เสร็จ
  const [screen,   setScreen]   = useState("home");
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

// ─── Login Screen ─────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error,    setError]    = useState("");
  const [busy,     setBusy]     = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    if (!username.trim() || !password) { setError(t("กรอก Username และ Password ให้ครบ","Please enter both Username and Password")); return; }
    setBusy(true); setError("");
    try {
      const user = await verifyLogin(username, password);
      if (!user) { setError(t("Username หรือ Password ไม่ถูกต้อง หรือบัญชีถูกระงับ","Incorrect Username or Password, or the account is suspended")); setBusy(false); return; }
      onLogin(user);
    } catch (err) {
      setError(t("เกิดข้อผิดพลาด ลองใหม่อีกครั้ง","An error occurred, please try again"));
      setBusy(false);
    }
  };

  return (
    <div style={{minHeight:"100vh",background:T.headerGrad,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <form onSubmit={submit} style={{background:T.card,borderRadius:20,padding:36,width:400,maxWidth:"92vw",boxShadow:"0 24px 60px rgba(0,0,0,0.25)",position:"relative"}}>
        <div style={{position:"absolute",top:14,right:14}}><LangToggle dark={false}/></div>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{marginBottom:8,color:T.textMuted}}><Ico name="box" size={34} sw={1.5} /></div>
          <div style={{fontSize:11,letterSpacing:3,color:T.textMuted,textTransform:"uppercase",fontWeight:600}}>TENDER COST SYSTEM</div>
          <div style={{fontSize:19,fontWeight:700,color:T.textPrimary,marginTop:4}}>{t("เข้าสู่ระบบ","Sign in")}</div>
          <div style={{fontSize:12,color:T.textMuted,marginTop:4}}>{t("ล็อกอินตามแผนก: QS · จัดซื้อ · บัญชี · Admin","Login by department: QS · Procurement · Accounting · Admin")}</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <label style={{display:"flex",flexDirection:"column",gap:6}}>
            <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>Username</span>
            <input className="input-base" autoFocus value={username} onChange={e=>setUsername(e.target.value)} placeholder={t("เช่น qs, procurement, accounting, admin","e.g. qs, procurement, accounting, admin")} />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:6}}>
            <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>Password</span>
            <input className="input-base" type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" />
          </label>
          {error && <div style={{background:T.redBg,color:T.red,fontSize:12,padding:"9px 12px",borderRadius:8,fontWeight:500}}>{error}</div>}
          <button className="btn-primary" type="submit" disabled={busy} style={{marginTop:6,opacity:busy?0.7:1}}>
            {busy ? t("กำลังตรวจสอบ...","Signing in...") : t("เข้าสู่ระบบ","Sign in")}
          </button>
        </div>
      </form>
    </div>
  );
}

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
          {[["users",t("จัดการผู้ใช้","Manage users")],["accounts",t("รหัสบัญชี","Account codes")],["logs",t("Log การเข้าใช้งาน","Access log")],["restore",t("กู้คืนข้อมูล","Restore data")]].map(([id,label])=>(
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

// ─── Home Screen ──────────────────────────────────────────────────────────────
function HomeScreen({ projects, loadErr, onRetryLoad, saveProjects, openProject, deleteProject, newProjModal, setNewProjModal, syncedAt, syncing, session, onLogout, onOpenAdmin }) {
  const [draft, setDraft] = useState({ name:"", area:"", panels:"", client:"", currency:"THB", usdRate:"" });
  const [projSearch, setProjSearch] = useState("");
  // ตัวเลขสรุปของแต่ละโครงการ — โหลดเบื้องหลังในคำขอเดียว (ถ้าพลาดการ์ดแค่ไม่แสดงตัวเลข ไม่กระทบการใช้งาน)
  const [summaries, setSummaries] = useState({});
  const projIdsKey = projects.map(p => p.id).join("|");
  useEffect(() => {
    const ids = projIdsKey ? projIdsKey.split("|") : [];
    if (!ids.length) return;
    let alive = true;
    const parts = ["tenders","additions","extra","hidden","po"];
    // แบ่งเป็นชุดละ 40 โครงการ (กัน URL ยาวเกินเมื่อมีโครงการเยอะ)
    const chunks = [];
    for (let i = 0; i < ids.length; i += 40) chunks.push(ids.slice(i, i + 40));
    Promise.all(chunks.map(c => sgMany(c.flatMap(id => parts.map(k => `tcs-${k}-${id}`)))))
      .then(res => {
        if (!alive) return;
        const all = Object.assign({}, ...res), out = {};
        ids.forEach(id => {
          const d = {}; parts.forEach(k => { d[k] = all[`tcs-${k}-${id}`]; });
          try { out[id] = projectSummary(d); } catch (e) { console.warn("สรุปโครงการไม่ได้:", id, e); out[id] = "err"; }   // ข้อมูลเสียโครงการเดียวไม่ลากทั้งหน้า
        });
        setSummaries(out);
      })
      .catch(e => { console.warn("โหลดสรุปโครงการไม่สำเร็จ:", e); if (alive) setSummaries(Object.fromEntries(ids.map(id => [id, "err"]))); });
    return () => { alive = false; };
  }, [projIdsKey]);
  const shownProjects = projects.filter(p => {
    const q = projSearch.trim().toLowerCase();
    if (!q) return true;
    return (p.name||"").toLowerCase().includes(q) || (p.client||"").toLowerCase().includes(q);
  });

  const createProject = () => {
    if (!draft.name.trim()) return;
    const id = uid();
    saveProjects([...projects, { ...draft, id, createdAt: new Date().toISOString() }]);
    setNewProjModal(false);
    setDraft({ name:"", area:"", panels:"", client:"", currency:"THB", usdRate:"" });
  };

  return (
    <div style={{minHeight:"100vh",background:T.bg}}>
      <TopBar brand title={t("ระบบบริหารต้นทุนโครงการ","Project Cost Management")} sub={`QS · ${t("จัดซื้อ · บัญชี","Procurement · Accounting")} · Tender Cost`}>
        <SyncBadge light syncing={syncing} syncedAt={syncedAt}/>
        {session?.role === "admin" && (
          <button className="btn-ghost" onClick={onOpenAdmin} title="Admin" aria-label="Admin" style={{display:"inline-flex",alignItems:"center",gap:6}}>
            <Ico name="gear" size={16} /><span className="hdr-lbl">Admin</span>
          </button>
        )}
        {session?.role !== "accounting" && (
          <button className="btn-primary" onClick={()=>setNewProjModal(true)} title={t("โครงการใหม่","New project")} aria-label={t("โครงการใหม่","New project")} style={{display:"inline-flex",alignItems:"center",gap:6}}>
            <Ico name="plus" size={16} /><span className="hdr-lbl">{t("โครงการใหม่","New project")}</span>
          </button>
        )}
        <UserMenu session={session} onLogout={onLogout} />
      </TopBar>

      {/* Body */}
      <div style={{padding:"28px 32px"}}>
        {projects.length === 0 && loadErr ? (
          <div style={{textAlign:"center",padding:"80px 0",color:T.textMuted}}>
            <div style={{marginBottom:12,color:T.amber}}><Ico name="alert" size={40} sw={1.6} /></div>
            <div style={{fontSize:15,fontWeight:600,color:T.textPrimary,marginBottom:14}}>{t("โหลดรายการโครงการไม่สำเร็จ — ตรวจเน็ตแล้วกดลองใหม่","Couldn't load the project list — check your connection and retry")}</div>
            <button className="btn-primary" onClick={onRetryLoad}>{t("ลองใหม่","Retry")}</button>
          </div>
        ) : projects.length === 0 ? (
          <div style={{textAlign:"center",padding:"80px 0",color:T.textMuted}}>
            <div style={{marginBottom:14,color:T.textMuted}}><Ico name="box" size={44} sw={1.4} /></div>
            <div style={{fontSize:17,fontWeight:600,color:T.textSecondary,marginBottom:8}}>{t("ยังไม่มีโครงการ","No projects yet")}</div>
            <div style={{fontSize:13,marginBottom:20}}>{t('กด "โครงการใหม่" เพื่อเริ่มต้น','Press "New project" to start')}</div>
            {session?.role !== "accounting" && <button className="btn-primary" onClick={()=>setNewProjModal(true)}>+ {t("สร้างโครงการแรก","Create first project")}</button>}
          </div>
        ) : (
          <>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18,flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:220,maxWidth:360}}><SearchInput value={projSearch} onChange={setProjSearch} placeholder={t("ค้นหาโครงการ / ชื่อลูกค้า…","Search project / client…")} width="100%"/></div>
              <span style={{fontSize:12,color:T.textMuted,fontWeight:500}}>
                {projSearch.trim() ? `${t("พบ","Found")} ${shownProjects.length} ${t("จาก","of")} ${projects.length} ${t("โครงการ","projects")}` : `${projects.length} ${t("โครงการทั้งหมด","projects total")}`}
              </span>
            </div>
            {shownProjects.length === 0 ? (
              <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>{t("ไม่พบโครงการที่ตรงกับ","No projects match")} "{projSearch}"</div>
            ) : (
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:20}}>
                {shownProjects.map(p => <ProjectCard key={p.id} project={p} summary={summaries[p.id]} onOpen={()=>openProject(p.id)} onDelete={session?.role==="admin" ? ()=>deleteProject(p.id) : null} />)}
              </div>
            )}
          </>
        )}
      </div>

      {/* New Project Modal */}
      {newProjModal && (
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,backdropFilter:"blur(4px)"}}>
          <div style={{background:T.card,borderRadius:20,padding:32,width:500,maxWidth:"90vw",boxShadow:"0 24px 60px rgba(0,0,0,0.15)",animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
              <div>
                <div style={{fontSize:16,fontWeight:650,color:T.textPrimary}}>{t("สร้างโครงการใหม่","Create new project")}</div>
                <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>{t("กรอกข้อมูลโครงการเพื่อเริ่มต้น","Fill in project details to start")}</div>
              </div>
              <button onClick={()=>setNewProjModal(false)} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              {[
                [t("ชื่อโครงการ *","Project name *"),"name","text","1/-1"],
                [t("ลูกค้า / Client","Client"),"client","text","1/-1"],
                [t("พื้นที่รวม (ft²)","Total area (ft²)"),"area","number","auto"],
                [t("จำนวน Panels","Panels"),"panels","number","auto"],
                [t("สกุลเงิน","Currency"),"currency","text","auto"],
                [t("อัตราแลกเปลี่ยน (บาท/USD)","FX rate (THB/USD)"),"usdRate","number","auto"],
              ].map(([label,key,type,col]) => (
                <label key={key} style={{display:"flex",flexDirection:"column",gap:6,gridColumn:col}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{label}</span>
                  <input type={type} step={type==="number"?"any":undefined} value={draft[key]} onChange={e=>setDraft(d=>({...d,[key]:e.target.value}))} className="input-base"/>
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:24}}>
              <button onClick={createProject} disabled={!draft.name.trim()} className="btn-primary" style={{opacity:draft.name.trim()?1:0.5}}>{t("สร้างโครงการ","Create project")}</button>
              <button onClick={()=>setNewProjModal(false)} className="btn-ghost">{t("ยกเลิก","Cancel")}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ProjectCard({ project, summary: sm, onOpen, onDelete }) {
  const failed = sm === "err";                     // โหลดตัวเลขไม่ได้ → แสดง "—" (ไม่ค้าง "…")
  const summary = failed ? null : sm;
  const wait = failed ? "—" : "…";
  const ageRaw = Math.floor((Date.now() - new Date(project.createdAt)) / 86400000);
  const age = Number.isFinite(ageRaw) && ageRaw >= 0 ? ageRaw : null;
  return (
    <div className="card-hover" onClick={onOpen} title={t("เปิดโครงการ","Open project")} data-project-card={project.id}
      style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:22,cursor:"pointer",position:"relative"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:650,color:T.textPrimary,lineHeight:1.3,minWidth:0}}>{project.name}</div>
        {onDelete && (
          <button onClick={e=>{e.stopPropagation();onDelete();}} aria-label={t("ลบโครงการ","Delete project")} title={t("ลบโครงการ","Delete project")} className="icon-danger"
            style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",padding:6,borderRadius:8,display:"grid",placeItems:"center",flexShrink:0,marginTop:-4,marginRight:-6}}>
            <Ico name="trash" size={17} />
          </button>
        )}
      </div>
      {project.client ? <div style={{fontSize:12,color:T.textSecondary,marginBottom:12}}>{project.client}</div> : <div style={{height:8}}/>}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
        {project.area   && <span style={{background:"#f1f5f9",color:T.textSecondary,fontSize:12,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.area} ft²</span>}
        {project.panels && <span style={{background:"#f1f5f9",color:T.textSecondary,fontSize:12,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.panels} Panels</span>}
        {project.currency && <span style={{background:"#f1f5f9",color:T.textMuted,fontSize:12,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.currency}</span>}
      </div>
      <div className="proj-stats" style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,marginBottom:16,padding:"10px 12px",background:T.bg,borderRadius:10}}>
        {[
          [t("งบรวม","Budget"), summary ? "฿"+fmtK(summary.budget) : wait, T.textPrimary],
          [t("ออก PO แล้ว","PO issued"), summary ? (summary.budget > 0 ? `${summary.pct.toFixed(0)}%` : (summary.poCount ? "฿"+fmtK(summary.committed) : "0%")) : wait, summary && summary.pct > 100 ? T.red : T.blue],
          [t("ต้องจ่ายเดือนนี้","Due this month"), summary ? (summary.dueNow > 0 ? "฿"+fmtK(summary.dueNow) : "—") : wait, summary && summary.dueNow > 0 ? T.red : T.textMuted],
        ].map(([l, v, c]) => (
          <div key={l} style={{minWidth:0}}>
            <div style={{fontSize:11,color:T.textMuted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{l}</div>
            <div style={{fontSize:14,fontWeight:650,color:c,fontVariantNumeric:"tabular-nums",marginTop:2}}>{v}</div>
          </div>
        ))}
      </div>
      {/* ต้องรีบดู — เห็นทุกโครงการจากหน้าเดียว (ของเข้าล่าช้า / หมวดเกินงบ) */}
      {summary && (summary.late > 0 || summary.overCodes > 0) && (
        <div data-proj-attention style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",margin:"-4px 0 14px"}}>
          <Ico name="alert" size={15} color={T.red} />
          {summary.late > 0 && <span className="att-chip" style={{cursor:"inherit"}}><b>{summary.late}</b> {t("ของเข้าล่าช้า","late incoming")}</span>}
          {summary.overCodes > 0 && <span className="att-chip" style={{cursor:"inherit"}}><b>{summary.overCodes}</b> {t("หมวดเกินงบ","over budget")}</span>}
        </div>
      )}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{fontSize:12,color:T.textMuted}}>{age === null ? "—" : age === 0 ? t("สร้างวันนี้","Created today") : `${age} ${t("วันที่แล้ว","days ago")}`}</div>
        <button onClick={e=>{e.stopPropagation();onOpen();}} className="btn-primary" style={{padding:"8px 18px",fontSize:12}}>{t("เปิดโครงการ","Open")} →</button>
      </div>
    </div>
  );
}

// ─── Role Select ──────────────────────────────────────────────────────────────
function RoleSelect({ project, updateProject, onSelect, onBack }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(project);
  // จำค่าตอน "เริ่มแก้" ไว้ — ข้อมูลจากเครื่องอื่นที่เข้ามาระหว่างพิมพ์จะไม่ล้างฟอร์ม และตอนบันทึก
  // ส่งเฉพาะช่องที่เราแก้จริง (เดิมส่งทุกช่อง → ทับช่องที่คนอื่นเพิ่งแก้ด้วยค่าเก่า)
  const editBaseRef = useRef(project);
  const startEdit = () => { editBaseRef.current = project; setDraft(project); setEditing(true); };
  const saveEdit = () => {
    const base = editBaseRef.current || {}, changed = {};
    ["name","client","currency","usdRate","area","panels"].forEach(k => { if (String(draft[k] ?? "") !== String(base[k] ?? "")) changed[k] = draft[k]; });
    if (Object.keys(changed).length) updateProject(changed);
    setEditing(false);
  };

  const ROLES = [
    {id:"qs",label:"QS",sub:"Quantity Surveyor",desc:t("ลงราคา Tender Cost\nประมาณการต้นทุนโครงการ","Enter Tender Cost\nestimate project cost"),color:T.blue,bg:T.blueLight,icon:"📐"},
    {id:"procurement",label:t("จัดซื้อ","Procurement"),sub:"Procurement",desc:t("ลงราคาจริงที่ซื้อ + วันที่\nออก PO และติดตามสถานะ","Enter actual prices + dates\nissue POs and track status"),color:"#d97706",bg:"#fffbeb",icon:"📦"},
    {id:"accounting",label:t("บัญชี","Accounting"),sub:"Accounting",desc:t("Dashboard ต้นทุน\nBudget vs Actual + Export Excel","Cost dashboard\nBudget vs Actual + Export Excel"),color:T.green,bg:T.greenBg,icon:"📊"},
  ];

  return (
    <div style={{minHeight:"100vh",background:T.bg}}>
      <TopBar onBack={onBack} title={project.name} sub={t("เลือกแผนกที่จะทำงาน","Choose a department")}>
        {project.area && <span className="hdr-meta" style={{fontSize:12,color:T.textMuted,whiteSpace:"nowrap"}}>{project.area} ft² · {project.panels} Panels</span>}
        <LangToggle dark={false} />
      </TopBar>

      <div style={{padding:"32px"}}>
        {editing ? (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:24,marginBottom:28,maxWidth:640}}>
            <div style={{fontSize:14,fontWeight:600,color:T.textPrimary,marginBottom:16}}>{t("แก้ไขข้อมูลโครงการ","Edit project details")}</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
              {[[t("ชื่อโครงการ","Project name"),"name","text"],[t("ลูกค้า","Client"),"client","text"],[t("สกุลเงิน","Currency"),"currency","text"],[t("อัตราแลกเปลี่ยน (บาท/USD)","FX rate (THB/USD)"),"usdRate","number"],[t("พื้นที่ (ft²)","Area (ft²)"),"area","number"],["Panels","panels","number"]].map(([l,k,t]) => (
                <label key={k} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{l}</span>
                  <input type={t} step={t==="number"?"any":undefined} value={draft[k]||""} onChange={e=>setDraft(d=>({...d,[k]:e.target.value}))} className="input-base"/>
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:16}}>
              <button className="btn-primary" onClick={saveEdit}>{t("บันทึก","Save")}</button>
              <button className="btn-ghost" onClick={()=>setEditing(false)}>{t("ยกเลิก","Cancel")}</button>
            </div>
          </div>
        ) : (
          <button onClick={startEdit} className="btn-ghost" style={{marginBottom:24,fontSize:12}}><Ico name="edit" /> {t("แก้ไขข้อมูลโครงการ","Edit project details")}</button>
        )}

        <div style={{fontSize:13,color:T.textSecondary,marginBottom:20,fontWeight:500}}>{t("เลือกแผนกที่จะทำงาน","Choose your department")}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:20,maxWidth:800}}>
          {ROLES.map(r => (
            <button key={r.id} onClick={()=>onSelect(r.id)} className="card-hover"
              style={{background:T.card,border:`1.5px solid ${T.cardBorder}`,borderRadius:18,padding:"28px 24px",cursor:"pointer",textAlign:"left",display:"flex",flexDirection:"column",gap:12,position:"relative",overflow:"hidden"}}>
              <div style={{width:44,height:44,borderRadius:12,background:r.bg,color:r.color,display:"flex",alignItems:"center",justifyContent:"center"}}>{EMOJI_ICON[r.icon] ? <Ico name={EMOJI_ICON[r.icon]} size={22} /> : r.icon}</div>
              <div>
                <div style={{fontSize:20,fontWeight:650,color:r.color}}>{r.label}</div>
                <div style={{fontSize:11,color:T.textMuted,marginTop:2,letterSpacing:0.5}}>{r.sub}</div>
              </div>
              <p style={{margin:0,fontSize:12,color:T.textSecondary,lineHeight:1.7,whiteSpace:"pre-line"}}>{r.desc}</p>
              <div style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:r.color,fontWeight:600,marginTop:4}}>{t("เข้าใช้งาน","Enter")} <span>→</span></div>
            </button>
          ))}
        </div>
      </div>
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

// ─── QS View ─────────────────────────────────────────────────────────────────
function QSView({ project, updateProject, tenderCosts, saveTenders, additions, saveAdditions, extraItems, saveExtraItems, hiddenAccounts, saveHiddenAccounts, onBack, onHome, onDept, syncedAt, syncing, session, onLogout, onExport, runExportFn, setEditMode }) {
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
      {tab === "baseline"
        ? <QSBaselineTab project={project} tenderCosts={tenderCosts} saveTenders={saveTenders} extraItems={extraItems} additions={additions}
                         onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} onHideAccount={handleHideAccount} onRestoreAccount={handleRestoreAccount} setEditMode={setEditMode} />
        : <QSMonthlyTab tenderCosts={tenderCosts} additions={additions} saveAdditions={saveAdditions}
                         extraItems={extraItems} onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} setEditMode={setEditMode} project={project} registerMonthExport={registerMonthExport} />}
    </Shell>
  );
}

// ─── QS Tab 1: Baseline (original tender cost) ────────────────────────────────
function QSBaselineTab({ project, tenderCosts, saveTenders, extraItems, additions = {}, onAddExtra, onDeleteExtra, hiddenAccounts, onHideAccount, onRestoreAccount, setEditMode }) {
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
                    <td style={{padding:"10px 16px",color:a.isExtra?T.amber:T.blue,fontVariantNumeric:"tabular-nums",fontSize:13,fontWeight:500}}>
                      {hasKids && (
                        <span title={isCollapsed?t("ขยายรายการย่อย","Expand sub-items"):t("ย่อรายการย่อย","Collapse sub-items")}
                          style={{color:T.textMuted,fontSize:12,marginRight:6,verticalAlign:"middle",display:"inline-block"}}>
                          {isCollapsed?"▸":"▾"}
                        </span>
                      )}
                      {a.isExtra ? (a.code.startsWith("EX-") ? "—" : a.code) : a.code}
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
        <table style={{minWidth: isMultiCol ? "max-content" : "100%", width: isMultiCol ? "max-content" : "100%", borderCollapse:"collapse", fontSize:13}}>
          <thead>
            {isMultiCol ? (
              <>
                <tr style={{background:"#f8fafc"}}>
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
                      {r.code}
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
                        <td colSpan={columns.length} style={{padding:"8px 16px",textAlign:"right"}}>
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
                      })
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
                        <td style={{padding:"7px 16px",textAlign:"right"}}>
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
                      <td colSpan={(isMultiCol ? 8+columns.length : 9)-3} style={{padding:"7px 16px",display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button className="btn-primary" style={{padding:"5px 12px",fontSize:13}} onClick={()=>handleAddSub(r.code)}>+ {t("เพิ่ม","Add")}</button>
                        <button className="btn-ghost" style={{padding:"5px 12px",fontSize:13}} onClick={()=>setSubFor(null)}>{t("ยกเลิก","Cancel")}</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={isMultiCol ? 8+columns.length : 9} style={{padding:"28px 16px",textAlign:"center",color:T.textMuted,fontSize:13}}>{t("ไม่พบรายการที่ตรงกับการค้นหา","No items match your search")}</td></tr>
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
                  ); })
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
  return String(allowNegative ? Math.round(sum * 100) / 100 : Math.max(0, sum));
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

function PODetailModal({ po: rawPo, onClose, onEdit, onDelete, onStatusChange, onChangePO, session, usdRate=0 }) {
  const [historyOpen, setHistoryOpen] = useState(false);
  const [capWarn, setCapWarn] = useState(""); // เตือนเมื่อยอดของเข้าจริงรวมเกินยอดสั่ง
  const [confirmDel, setConfirmDel] = useState(false); // ยืนยันลบในแอป (กันกรณี window.confirm ถูกบล็อกใน webview)
  // กด Esc = ปิด/ยกเลิกหน้ารายละเอียด
  useEffect(() => {
    if (!rawPo) return;
    const onEsc = (e) => { if (e.key === "Escape") { e.preventDefault(); onClose?.(); } };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [rawPo, onClose]);
  if (!rawPo) return null;
  const po = migratePO(rawPo);
  const items = po.items;
  const supplier = po.supplier;
  const inc = incomingStatus(po), pay = paymentStatus(po);
  const history = poHistory(po);
  const lastUpd = poLastUpdate(po);
  const locked = !canEditPO(po, session);
  const receivedDates = poReceivedDates(po);
  const paidDate = poPaidDate(po);
  // ยอด "ของเข้าจริง" รวมทุกงวด ห้ามเกินยอดสั่ง — ใช้ปิดปุ่มบันทึก + เตือนค้างไว้
  // (ยอดแผนเกินมี overPlanned เตือนแบบไม่บล็อกอยู่แล้ว จะได้ไม่กันการบันทึกยอดที่รับจริง)
  const overCapItem = items.find(it => {
    const o = itemOrdered(it); if (!(o > 0)) return false;
    const recvSum = (it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
    return Math.round(recvSum*100) > Math.round(o*100);
  });

  // Record actual received / split remaining into a new round, then persist.
  // PO ที่ปิดแล้ว (รับครบ+จ่ายครบ) แก้งวด/ยอด/วันรับได้เฉพาะ Admin — กันการ "ปลดล็อกตัวเอง"
  // ด้วยการล้างวันรับจริง แล้วค่อยลบ/แก้ PO ได้ และทุกการแก้งวดบันทึกลงประวัติ PO
  const setItemRounds = (itemId, rounds) => {
    if (locked) { setCapWarn(t("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ยอด/วันรับได้เฉพาะ Admin","This PO is fully received & paid — only Admin can change amounts/dates")); return; }
    const code = po.items.find(it => it.id===itemId)?.code || "";
    const next = { ...po, items: po.items.map(it => it.id===itemId ? {...it, rounds} : it) };
    const msg  = t(`แก้งวดของเข้า ${code}`, `Edited delivery rounds ${code}`);
    // แก้ต่อเนื่องหลายช่องภายใน 2 นาที (คนเดิม รายการเดิม) = รวมเป็นรายการประวัติเดียว ไม่ให้ประวัติรก
    const last = (po.history||[])[0];
    const recent = last && last.message===msg && last.user===(session?.name||"—") && (Date.now()-new Date(last.at).getTime()) < 120000;
    onChangePO?.(recent
      ? { ...next, history: [{ ...last, at: new Date().toISOString() }, ...(po.history||[]).slice(1)] }
      : withHistory(next, historyEntry(session, "edited", msg)));
  };
  // ลงยอดของเข้าจริง — ห้ามให้ยอดรวมทุกงวดเกิน "ยอดสั่ง" ของ PO นั้น (บล็อก+เตือน)
  const setActualAmount = (itemId, roundId, val) => {
    const it = po.items.find(i=>i.id===itemId); if (!it) return;
    const ordered = itemOrdered(it);
    const newVal = parseFloat(val)||0;
    const otherReceived = (it.rounds||[]).filter(r=>r.id!==roundId).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
    if (ordered>0 && Math.round((otherReceived + newVal)*100) > Math.round(ordered*100)) {
      const maxAllow = Math.max(ordered - otherReceived, 0);
      setCapWarn(t(`⚠ ${it.code||"รายการนี้"}: ยอดของเข้ารวมห้ามเกินยอดสั่ง ${fmt(ordered)} — งวดนี้กรอกได้ไม่เกิน ${fmt(maxAllow)} (ระบบไม่บันทึกค่าที่เกิน)`, `⚠ ${it.code||"This item"}: total received can't exceed the order ${fmt(ordered)} — this round allows at most ${fmt(maxAllow)} (the excess was not saved)`));
      return; // บล็อก: ไม่บันทึกค่าที่เกินยอดสั่ง
    }
    setCapWarn("");
    updateRound(itemId, roundId, "actualAmount", val);
  };
  const updateRound = (itemId, roundId, key, val) => {
    const it = po.items.find(i=>i.id===itemId); if (!it) return;
    setItemRounds(itemId, it.rounds.map(r => r.id===roundId ? {...r,[key]:val} : r));
  };
  // เพิ่มงวดของเข้าใหม่ (งวดเปล่า) — ให้ผู้ใช้กรอกยอด/วันของเข้าเอง โดยยอดของเข้า
  // รวมทุกงวดถูกจำกัดไม่ให้เกินยอดสั่งอยู่แล้ว (setActualAmount) จึงไม่ตั้งยอดแผนซ้ำ
  const splitRound = (itemId) => {
    const it = po.items.find(i=>i.id===itemId); if (!it) return;
    setItemRounds(itemId, [...it.rounds, { id:uid(), planDate:"", planAmount:"", actualAmount:"", actualDate:"" }]);
  };
  // ลบงวดส่งของ — ต้องเหลืออย่างน้อย 1 งวดเสมอ (ใช้แก้กรณีมีงวดเกิน/ซ้ำ)
  const removeRound = async (itemId, roundId) => {
    const it = po.items.find(i=>i.id===itemId); if (!it) return;
    if ((it.rounds||[]).length <= 1) return;
    if (locked) { setItemRounds(itemId, it.rounds); return; } // แสดงคำเตือน 🔒 โดยไม่ถามยืนยันก่อน
    if (!(await uiConfirm(t("ลบงวดนี้? (ยอด/วันของเข้าที่กรอกในงวดนี้จะถูกลบ)","Delete this round? (its entered amount/date will be removed)"), { danger: true, okLabel: t("ลบงวด","Delete round") }))) return;
    setItemRounds(itemId, it.rounds.filter(r => r.id !== roundId));
  };
  const roundBadge = (r) => {
    if (!r.actualDate || !(parseFloat(r.actualAmount)||0)) return [t("รอของเข้า","Awaiting goods"), PAYMENT_BG.pending, PAYMENT_CLR.pending];
    // ถ้าวันของเข้าจริงยังมาไม่ถึง (วันในอนาคต) = ยังไม่ถือว่ารับของ แสดงเป็น "นัดรับ"
    if (r.actualDate > todayStr()) return [t(`นัดรับ ${r.actualDate} (ยังไม่ถึงวัน)`,`Due ${r.actualDate} (not yet)`), INCOMING_BG.pending, INCOMING_CLR.pending];
    return roundPaid(po,r) ? [t("ถึงกำหนดจ่ายแล้ว","Payment due"), PAYMENT_BG.paid, PAYMENT_CLR.paid]
                           : [t("ของเข้าแล้ว · รอครบกำหนด","Received · awaiting due"), INCOMING_BG.partial, INCOMING_CLR.partial];
  };

  const Row = ({ label, value, mono }) => (
    <div style={{display:"flex",justifyContent:"space-between",gap:16,padding:"10px 0",borderBottom:`1px solid #f1f5f9`}}>
      <span style={{fontSize:12,color:T.textMuted,fontWeight:500}}>{label}</span>
      <span style={{fontSize:13,color:T.textPrimary,fontWeight:600,textAlign:"right",fontVariantNumeric:mono?"tabular-nums":undefined}}>{value ?? "—"}</span>
    </div>
  );

  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,padding:20,animation:"fadeIn 0.15s ease"}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:16,padding:26,width:"100%",maxWidth:520,maxHeight:"88vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
          <div>
            <div style={{fontSize:16,fontWeight:650,color:T.textPrimary}}>{poSupplierLabel(po)}</div>
            <div style={{fontSize:12,color:T.textMuted,fontVariantNumeric:"tabular-nums",marginTop:2}}>{poNumbersLabel(po)}</div>
          </div>
          <button onClick={onClose} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted,flexShrink:0}}>×</button>
        </div>

        {locked && (
          <div style={{display:"flex",alignItems:"center",gap:6,background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"6px 10px",margin:"8px 0 2px",fontSize:11,color:"#92400e"}}>
            <Ico name="lock" size={14} /> {t("รับของและจ่ายเงินครบแล้ว — แก้ยอด/วันของเข้าจริงได้ (ลบ PO และแก้ผู้ขาย/หมวด/ยอดสั่ง เฉพาะ Admin)","Fully received & paid — actual amount/date still editable (delete PO and edit vendor/category/order: Admin only)")}
          </div>
        )}

        {/* Status is a live dropdown here too — the most natural place to
            update it right after reviewing everything else on the PO. */}
        <div style={{display:"flex",gap:6,margin:"12px 0 4px",flexWrap:"wrap",alignItems:"center"}}>
          <StatusPicker status={po.status} onChange={s=>onStatusChange?.(po,s)} disabled={locked}/>
          <span style={{background:INCOMING_BG[inc],color:INCOMING_CLR[inc],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{incLabel(inc)}</span>
          <span style={{background:PAYMENT_BG[pay],color:PAYMENT_CLR[pay],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{payLabel(pay)}</span>
          {po.paymentType && (
            <span style={{background:PAYMENT_TYPE_BG[po.paymentType],color:PAYMENT_TYPE_CLR[po.paymentType],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{payTypeLabelT(po)}</span>
          )}
        </div>
        {lastUpd && (
          <div style={{fontSize:11,color:T.textMuted,marginBottom:4}}>
            <Ico name="clock" size={13} /> {t("อัปเดตล่าสุด","Last updated")} {relativeTime(lastUpd.at)} {t("โดย","by")} <b style={{color:T.textSecondary}}>{lastUpd.user}</b>
          </div>
        )}

        {/* Supplier (one per PO) + top-line dates */}
        <div style={{marginTop:10,background:T.bg,borderRadius:10,padding:"10px 12px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline"}}>
            <div><span style={{fontSize:13,fontWeight:650,color:T.textPrimary}}>{supplier.name||"—"}</span>
              {supplier.poNumber && <span style={{fontSize:11,color:T.textMuted,fontVariantNumeric:"tabular-nums",marginLeft:8}}>{supplier.poNumber}</span>}</div>
            <span style={{fontSize:13,fontVariantNumeric:"tabular-nums",fontWeight:650,color:T.amber}}>{fmt(poTotal(po))}{usdRate>0 && <span className="usd-sub" style={{color:T.green,fontWeight:650,fontSize:12,marginLeft:6}}>≈ ${fmt(poTotal(po)/usdRate)}</span>}</span>
          </div>
        </div>

        <div style={{marginTop:4}}>
          <Row label={t("วันเปิด PO","PO date")} value={fmtDate(po.date)} />
          <Row label={t("วันรับของ","Received")} value={receivedDates.length ? receivedDates.map(fmtDate).join(", ") : t("ยังไม่ได้รับ","Not received")} />
          <Row label={paidDate ? t("วันจ่ายเงิน","Payment date") : t("ครบกำหนดจ่าย","Payment due")} value={paidDate ? fmtDate(paidDate) : (poNextDueDate(po) ? `${fmtDate(poNextDueDate(po))} · ${t("ยังไม่จ่าย","not paid yet")}` : t("ยังไม่กำหนด","Not set"))} />
          <Row label={t("วิธีจ่ายเงิน","Payment method")} value={po.paymentType ? payTypeLabelT(po) : "—"} />
        </div>

        {/* Per account-code: receiving in installments, with auto-pay + split */}
        <div style={{marginTop:12}}>
          <div style={{fontSize:11,fontWeight:650,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase",marginBottom:8}}>{t("ของเข้า / จ่ายเงิน (แบ่งงวดได้)","Incoming / payment (by rounds)")}</div>
          {items.map((it,ii)=>{
            const acc = accountOf(it.code);
            const ordered = itemOrdered(it), recv = itemReceived(it), remain = itemRemaining(it);
            const planned = (it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.planAmount)||0),0); // ยอดรวมที่วางแผนไว้ทุกงวด
            const planRemain = Math.max(ordered - planned, 0);   // ยอดที่ยัง "ไม่ถูกวางแผน" (ไว้แบ่งงวดเพิ่ม)
            const overPlanned = planned - ordered;               // >0 = รวมทุกงวดเกินยอดสั่ง (มีงวดเกิน/ซ้ำ)
            const paidAmt = itemPaidAmount(po, ii);
            return (
              <div key={it.id||ii} style={{background:T.bg,borderRadius:12,padding:"12px 14px",marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:8}}>
                  <div style={{minWidth:0}}>
                    <span style={{fontSize:11,color:T.blue,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{it.code||"—"}</span>
                    <span style={{fontSize:12,color:T.textSecondary,marginLeft:8}}>{acc?.name||"—"}</span>
                  </div>
                  <span style={{fontSize:12,color:T.textMuted}}>{t("สั่ง","Ordered")} <b style={{color:T.textPrimary,fontVariantNumeric:"tabular-nums"}}>{fmt(ordered)}</b></span>
                </div>

                {(it.rounds||[]).map((r,ri)=>{
                  const [label,bg,clr] = roundBadge(r);
                  const payDate = roundPayDate(po,r);
                  const late = r.actualDate && r.planDate && r.actualDate>r.planDate;
                  return (
                    <div key={r.id||ri} style={{border:`1px solid ${T.cardBorder}`,borderRadius:10,padding:10,marginBottom:6,background:T.card}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:8}}>
                        <span style={{fontSize:11,fontWeight:650,color:T.textSecondary}}>{t("งวดที่","Round")} {ri+1}{(parseFloat(r.planAmount)||0)>0 ? ` · ${t("แผน","plan")} ${fmt(r.planAmount)}` : ""}</span>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <span style={{background:bg,color:clr,fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:600}}>{label}</span>
                          {!locked && it.rounds.length>1 && (
                            <button type="button" onClick={()=>removeRound(it.id,r.id)} title={t("ลบงวดนี้","Delete round")}
                              aria-label={t("ลบงวดนี้","Delete this round")} className="icon-danger"
                              style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",padding:"6px 8px",minWidth:32,minHeight:32,borderRadius:8,lineHeight:1,display:"inline-grid",placeItems:"center"}}><Ico name="trash" size={16} /></button>
                          )}
                        </div>
                      </div>
                      <div className="round-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                        <label style={{display:"flex",flexDirection:"column",gap:3}}>
                          <span style={{fontSize:12,color:T.textSecondary}}>{t("ยอดของเข้าจริง (บาท)","Actual received (THB)")}</span>
                          <MoneyInput value={r.actualAmount} placeholder={t("บาท","THB")} disabled={locked}
                            onChange={v=>setActualAmount(it.id,r.id,v)}/>
                        </label>
                        <label style={{display:"flex",flexDirection:"column",gap:3}}>
                          <span style={{fontSize:12,color:T.textSecondary}}>{t("วันของเข้าจริง","Actual date")}</span>
                          <DateInput value={r.actualDate} disabled={locked}
                            onChange={e=>updateRound(it.id,r.id,"actualDate",e.target.value)}/>
                        </label>
                      </div>
                      <div style={{marginTop:6,fontSize:11,color:T.textSecondary}}>
                        <Ico name="wallet" size={13} /> {t("วันครบกำหนดจ่าย","Payment due")}: <span style={{fontVariantNumeric:"tabular-nums",color:T.textPrimary}}>{payDate ? fmtDate(payDate) : "—"}</span>
                        <span style={{color:T.textMuted}}> ({po.paymentType==="cash"?t("เงินสด","Cash"):po.paymentType==="credit"?t(`เครดิต ${po.creditDays} วัน`,`Credit ${po.creditDays}d`):t("ยังไม่ระบุวิธีจ่าย","No method")})</span>
                        {late && <span style={{color:T.red}}> · {t("ของมาช้า","late arrival")}</span>}
                      </div>
                    </div>
                  );
                })}

                {/* Progress + split */}
                <div style={{height:8,borderRadius:6,background:T.cardBorder,overflow:"hidden",marginTop:6}}>
                  <div style={{height:"100%",width:`${ordered>0?Math.min(recv/ordered*100,100):0}%`,background:remain>0?T.amber:T.green}}/>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",marginTop:5,fontSize:11,color:T.textSecondary}}>
                  <span>{t("ของเข้าแล้ว","Received")} <b style={{fontVariantNumeric:"tabular-nums",color:T.textPrimary}}>{fmt(recv)}</b> / {fmt(ordered)}</span>
                  <span>{t("จ่ายแล้ว","Paid")} <b style={{fontVariantNumeric:"tabular-nums",color:paidAmt>0?T.green:T.textMuted}}>{fmt(paidAmt)}</b></span>
                </div>
                {Math.round(overPlanned*100)>0 && (
                  <div style={{marginTop:8,fontSize:11,color:T.red,background:T.redBg,borderRadius:8,padding:"7px 10px",lineHeight:1.4}}>
                    ⚠ {t("ยอดรวมทุกงวด","Total all rounds")} <b style={{fontVariantNumeric:"tabular-nums"}}>{fmt(planned)}</b> {t("เกินยอดสั่ง","exceeds the order")} <b style={{fontVariantNumeric:"tabular-nums"}}>{fmt(ordered)}</b> {t("อยู่","by")} {fmt(overPlanned)} — {t("ลบงวดที่เกินออก (ปุ่มถังขยะ)","remove the extra round (trash button)")}
                  </div>
                )}
                {!locked && ordered>0 && remain>0.001 ? (
                  <button type="button" onClick={()=>splitRound(it.id)} className="btn-ghost"
                    style={{marginTop:8,padding:"6px 12px",fontSize:12,borderColor:T.amber,color:T.amber}}>
                    + {t("เพิ่มงวดของเข้า — เหลือรับอีก","Add round — remaining")} {fmt(remain)}
                  </button>
                ) : recv>0.001 && remain<=0.001 && ordered>0 ? (
                  <div style={{marginTop:8,fontSize:12,color:T.green}}>✓ {t("ของเข้าครบตามยอดสั่งแล้ว","Fully received")}</div>
                ) : null}
              </div>
            );
          })}
        </div>

        {po.notes && (
          <div style={{padding:"10px 0 0"}}>
            <div style={{fontSize:12,color:T.textMuted,fontWeight:500,marginBottom:4}}>{t("หมายเหตุ","Notes")}</div>
            <div style={{fontSize:13,color:T.textPrimary,lineHeight:1.5,whiteSpace:"pre-wrap"}}>{po.notes}</div>
          </div>
        )}

        {/* Edit history — a running log of who changed what, so status
            changes and edits are always traceable after the fact. */}
        {history.length > 0 && (
          <div style={{marginTop:14,borderTop:`1px solid ${T.cardBorder}`,paddingTop:10}}>
            <button onClick={()=>setHistoryOpen(v=>!v)}
              style={{background:"none",border:"none",padding:0,cursor:"pointer",display:"flex",alignItems:"center",gap:6,fontSize:11,fontWeight:650,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase"}}>
              <span style={{transition:"transform 0.15s",transform:historyOpen?"rotate(90deg)":"none",display:"inline-block"}}>▸</span>
              <Ico name="clock" size={14} /> {t("ประวัติการแก้ไข","Edit history")} ({history.length})
            </button>
            {historyOpen && (
              <div style={{marginTop:10,display:"flex",flexDirection:"column",gap:0}}>
                {history.map((h,i)=>(
                  <div key={h.id||i} style={{display:"flex",gap:10,padding:"7px 0",borderBottom:i<history.length-1?"1px solid #f1f5f9":"none"}}>
                    <span style={{fontSize:14,flexShrink:0}}>{HISTORY_ICON[h.action]||"•"}</span>
                    <div style={{minWidth:0,flex:1}}>
                      <div style={{fontSize:12,color:T.textPrimary,fontWeight:500}}>{h.message}</div>
                      <div style={{fontSize:11,color:T.textMuted,marginTop:1}}>{formatDateTime(h.at)} · {h.user}{h.role?` (${h.role})`:""}</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {(capWarn || overCapItem) && (
          <div style={{marginTop:14,fontSize:12,color:T.red,background:T.redBg,border:`1px solid ${T.red}`,borderRadius:10,padding:"9px 12px",lineHeight:1.5}}>
            {capWarn || t(`⚠ ${overCapItem.code||"รายการ"}: ยอดรวมทุกงวดเกินยอดสั่ง ${fmt(itemOrdered(overCapItem))} — แก้ให้ไม่เกินก่อน จึงจะกดบันทึกได้ (ปุ่มบันทึกถูกปิดไว้)`, `⚠ ${overCapItem.code||"Item"}: all rounds together exceed the order ${fmt(itemOrdered(overCapItem))} — fix it before saving (Save is disabled)`)}
          </div>
        )}
        {/* แถบปุ่มติดขอบล่างของป๊อปอัพ — เดิมอยู่ท้ายเนื้อหา ถ้า PO มีหลายงวดปุ่มจะถูกตัดจนต้องเลื่อนลงสุด */}
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center",position:"sticky",bottom:-26,margin:"16px -26px -26px",padding:"12px 26px 18px",background:T.card,borderTop:`1px solid ${T.cardBorder}`,boxShadow:"0 -8px 14px -10px rgba(15,23,42,0.25)",borderRadius:"0 0 16px 16px",zIndex:2}}>
          <button
            onClick={()=>{
              // ตรวจอีกครั้งก่อนปิด: ยอดของเข้าจริงรวมของทุกรายการห้ามเกินยอดสั่ง
              const bad = po.items.find(it => { const o=itemOrdered(it); const rc=(it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0); return o>0 && Math.round(rc*100) > Math.round(o*100); });
              if (bad) { setCapWarn(t(`⚠ ${bad.code||"รายการ"}: ยอดของเข้ารวมเกินยอดสั่ง ${fmt(itemOrdered(bad))} — แก้ให้ไม่เกินก่อนบันทึก`, `⚠ ${bad.code||"Item"}: total received exceeds the order ${fmt(itemOrdered(bad))} — fix it before saving`)); return; }
              setCapWarn(""); onClose();
            }}
            disabled={!!overCapItem} className="btn-primary"
            title={overCapItem?t(`${overCapItem.code||"รายการ"}: ยอดรวมทุกงวดเกินยอดสั่ง แก้ให้ไม่เกินก่อนบันทึก`,`${overCapItem.code||"item"}: total across rounds exceeds order — fix before saving`):undefined}
            style={overCapItem?{background:"#e2e8f0",color:"#94a3b8",cursor:"not-allowed"}:undefined}>{overCapItem && <Ico name="alert" size={15} />} {t("บันทึก","Save")}</button>
          {!locked && <button onClick={()=>onEdit(po)} className="btn-ghost" style={{fontSize:12}} title={t("แก้ผู้ขาย / หมวด / ยอดสั่ง","Edit vendor / category / order")}><Ico name="edit" size={14} /> {t("แก้ไข PO","Edit PO")}</button>}
          {confirmDel ? (
            <span style={{display:"flex",alignItems:"center",gap:6,background:T.redBg,border:`1px solid #fecaca`,borderRadius:10,padding:"4px 6px 4px 12px"}}>
              <span style={{fontSize:12,color:T.red,fontWeight:600,whiteSpace:"nowrap"}}>{t("ลบ PO นี้จริงไหม? ย้อนกลับไม่ได้","Delete this PO? Cannot be undone")}</span>
              <button onClick={()=>{ setConfirmDel(false); onDelete(po.id, true); }} className="btn-primary" style={{background:T.red,padding:"5px 12px",fontSize:12}}>{t("ลบเลย","Delete")}</button>
              <button onClick={()=>setConfirmDel(false)} className="btn-ghost" style={{padding:"5px 10px",fontSize:12}}>{t("ยกเลิก","Cancel")}</button>
            </span>
          ) : (
            <button onClick={()=>setConfirmDel(true)} disabled={locked} className="btn-ghost" style={{color:locked?"#cbd5e1":T.red,borderColor:locked?"#e2e8f0":T.red,cursor:locked?"not-allowed":"pointer",display:"inline-flex",alignItems:"center",gap:6}}><Ico name="trash" size={15} />{t("ลบ","Delete")}</button>
          )}
          <div style={{flex:1}}/>
          <button onClick={onClose} className="btn-ghost">{t("ปิด","Close")}</button>
        </div>
      </div>
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

// ─── จัดซื้อ: แผนของเข้าทั้งโปรเจค (ใช้ฟอร์มเดียวกับ PO) ────────────────────────
//  แผน = อ็อบเจ็กต์รูปเดียวกับ PO (isPlan:true) สร้าง/แก้ผ่านฟอร์ม PO โดยติ๊ก
//  "แผนของเข้า". แท็บนี้แค่แสดงลิสต์แผน + ปุ่มเรียกฟอร์ม. "→ ทำเป็น PO จริง" =
//  เปิดฟอร์มโดยเอาติ๊กออกให้ พอกดบันทึกก็กลายเป็น PO จริงและแผนถูกย้ายออก.
function IncomingPlanTab({ plans, poEntries = [], usdRate = 0, tenderCosts = {}, additions = {}, extraItems = [], hiddenAccounts = [], onNew, onEdit, onConvert, onDelete }) {
  const list = Array.isArray(plans) ? plans : [];
  const pos = Array.isArray(poEntries) ? poEntries : [];
  const acctList = exportAccountList(extraItems, hiddenAccounts);
  const nameOf = (code) => acctList.find(a => a.code === code)?.name || ACCOUNTS.find(a => a.code === code)?.name || "";
  const lbl = (d) => d ? new Date(d).toLocaleDateString(uiLocale(), { day: "numeric", month: "short", year: "2-digit" }) : "—";
  const planDates = (pl) => poRounds(pl).map(r => r.planDate).filter(Boolean).sort();
  const sorted = [...list].sort((a, b) => ((planDates(a)[0] || a.date || "")).localeCompare(planDates(b)[0] || b.date || ""));
  const monthLbl = (mk) => monthShortLabel(mk); // เดือนไทย + ปี พ.ศ. (เช่น "ส.ค. 69") ให้ตรงกับการ์ดแผน/Excel
  const today = todayStr();
  const [mSearch, setMSearch] = useState("");            // ค้นหาในตารางของเข้ารายเดือน (Acc. Code/ชื่อ)
  const [mSort, setMSort] = useState({ key: "code", dir: "asc" }); // เรียงตามหัวคอลัมน์

  // ── รวมรายการของเข้า แยกที่มา: จริง(รับแล้ว/PO) vs แผน — เดือนไหนมีทั้งคู่จะโชว์ 2 ค่า
  const entries = [];
  const lateOf = (r) => !!(r.planDate && r.planDate < today && !r.actualDate);
  list.forEach(pl => poItems(pl).forEach(it => (it.rounds || []).forEach(r => {
    const amt = parseFloat(r.planAmount) || 0; if (!amt) return;
    entries.push({ code: it.code, mk: (r.planDate || pl.date || "").slice(0, 7), amount: amt, src: "plan", late: lateOf(r), received: false });
  })));
  pos.forEach(p => poItems(p).forEach(it => (it.rounds || []).forEach(r => {
    if (roundReceived(r)) {
      const amt = parseFloat(r.actualAmount) || 0; if (!amt) return;
      entries.push({ code: it.code, mk: r.actualDate.slice(0, 7), amount: amt, src: "po", late: false, received: true });
    } else {
      // ยังไม่รับ = "PO รอเข้า" — ใช้ "ยอดของเข้าจริง" ถ้ากรอกไว้แล้ว (ตรงกับหน้ารายละเอียด) ไม่มีค่อยใช้ยอดแผน
      const amt = parseFloat(r.actualAmount) || parseFloat(r.planAmount) || 0; if (!amt) return;
      entries.push({ code: it.code, mk: (r.actualDate || r.planDate || p.date || "").slice(0, 7), amount: amt, src: "po", late: lateOf(r), received: false });
    }
  })));
  const months = [...new Set(entries.map(e => e.mk).filter(Boolean))].sort();

  // แต่ละช่องเก็บแยก: rec(รับแล้ว) / po(PO ยังไม่รับ) / plan(แผน) + ธง late
  const cellMap = {};
  entries.forEach(e => {
    const c = (cellMap[e.code] = cellMap[e.code] || {});
    const cell = (c[e.mk] = c[e.mk] || { rec: 0, po: 0, poLate: false, plan: 0, planLate: false });
    if (e.received) cell.rec += e.amount;
    else if (e.src === "po") { cell.po += e.amount; if (e.late) cell.poLate = true; }
    else { cell.plan += e.amount; if (e.late) cell.planLate = true; }
  });
  const codes = Object.keys(cellMap).sort();
  const cellOf = (code, mk) => cellMap[code]?.[mk] || null;
  const cellTot = (c) => c ? (c.rec + c.po + c.plan) : 0;
  const rowTot = (code) => months.reduce((s, mk) => s + cellTot(cellOf(code, mk)), 0);
  const colTot = (mk) => shownCodes.reduce((s, c) => s + cellTot(cellOf(c, mk)), 0);

  // ── คอลัมน์ต้นทุน: Tender Cost / Stock / Balance Cost ────────────────────────
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);
  const budgetOf = (code) => parseFloat(combinedBudget[code]) || 0;
  const committedOf = (code) => pos.reduce((s, p) => s + poAmountForCode(p, code), 0);
  const plannedOf = (code) => list.reduce((s, pl) => s + poAmountForCode(pl, code), 0);
  const stockOf = (code) => pos.reduce((s, p) => s + poItems(p).filter(it => it.code === code).reduce((ss, it) => ss + (parseFloat(it.store) || 0), 0), 0);
  const takeoffOf = (code) => [...pos, ...list].reduce((s, p) => s + poItems(p).filter(it => it.code === code).reduce((ss, it) => ss + (parseFloat(it.takeoff) || 0), 0), 0); // Take off (กรอกเอง)
  const issuePOof = (code) => committedOf(code);                                                   // Issue PO = ยอดรวม PO ที่ยื่นจริง
  const balCostOf = (code) => budgetOf(code) - stockOf(code) - committedOf(code) - plannedOf(code); // "Pending PO" = งบ − Stock − Issue PO − แผน (ยอดที่ยังต้องสั่ง)
  const balPOof   = (code) => budgetOf(code) - stockOf(code) - issuePOof(code);                      // "Balance Cost" = Tender Cost − Stock − Issue PO
  // ── ค้นหา + เรียงลำดับตามหัวคอลัมน์ ──────────────────────────────────────────
  const sortVal = (code, key) => {
    switch (key) {
      case "code":    return code;
      case "name":    return nameOf(code);
      case "tender":  return budgetOf(code);
      case "takeoff": return takeoffOf(code);
      case "stock":   return stockOf(code);
      case "issue":   return issuePOof(code);
      case "pending": return balCostOf(code);
      case "total":   return rowTot(code);
      case "balcost": return balPOof(code);
      default:        return key.startsWith("m:") ? cellTot(cellOf(code, key.slice(2))) : code;
    }
  };
  const mQ = mSearch.trim().toLowerCase();
  const shownCodes = codes
    .filter(c => !mQ || c.toLowerCase().includes(mQ) || nameOf(c).toLowerCase().includes(mQ))
    .sort((a, b) => {
      const va = sortVal(a, mSort.key), vb = sortVal(b, mSort.key);
      const d = (typeof va === "string" || typeof vb === "string")
        ? String(va).localeCompare(String(vb), "th")
        : (va - vb);
      return mSort.dir === "asc" ? d : -d;
    });
  const grand = shownCodes.reduce((s, c) => s + rowTot(c), 0);
  const toggleSort = (key) => setMSort(s => s.key === key
    ? { key, dir: s.dir === "asc" ? "desc" : "asc" }
    : { key, dir: (key === "code" || key === "name") ? "asc" : "desc" });
  const arrow = (key) => mSort.key === key ? (mSort.dir === "asc" ? " ▲" : " ▼") : "";

  const cM = { border: "1px solid #d9e0ea", padding: "8px 13px", fontSize:13, whiteSpace: "nowrap" };
  const nM = { ...cM, textAlign: "right", fontVariantNumeric: "tabular-nums" };
  const hM = (bg) => ({ ...cM, background: bg, fontWeight: 650, color: T.textSecondary, textAlign: "center", position: "sticky", top: 0 });
  const bCost = "#f4e9ef";
  // ตรึงคอลัมน์แรก 2 ช่อง (รหัส/ชื่อบัญชี) ให้ไม่เลื่อนหายตอนดูเดือนไกล ๆ
  const COL1_W = 86;
  const stickyBody0 = { position: "sticky", left: 0, background: "#fff", zIndex: 1 };
  const stickyBody1 = { position: "sticky", left: COL1_W, background: "#fff", zIndex: 1 };
  const stickyHead0 = { left: 0, zIndex: 3 };
  const stickyHead1 = { left: COL1_W, zIndex: 3 };
  const money = (n) => n ? (n < 0 ? `(${fmt(Math.abs(n))})` : fmt(n)) : "-";
  const sum = (fn) => shownCodes.reduce((s, c) => s + fn(c), 0);

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
        <button onClick={onNew} className="btn-primary" style={{ marginLeft: "auto" }}>+ {t("เพิ่ม PO","Add PO")}</button>
      </div>

      {/* รายการของเข้ารายเดือน — เดือนเป็นคอลัมน์ + ต้นทุน (แผน + PO จริง รวมกัน) */}
      {months.length > 0 && (
        <div style={{ marginBottom: 24 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: T.textPrimary, marginBottom: 10 }}>{t("รายการของเข้ารายเดือน (แผน + PO จริง)","Monthly incoming (plan + real PO)")}</div>
          <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
            {[[t("รับแล้ว","Received"), T.green, "#eafaf1"], [t("ล่าช้า ⚠","Late ⚠"), T.amber, "#fff6e6"], [t("PO รอเข้า","PO awaiting"), T.textPrimary, "#eef2f7"], [t("แผน (มี * ต่อท้าย)","Plan (with *)"), T.red, "#fdecec"]].map(([label, clr, bg]) => (
              <span key={label} style={{ display: "inline-flex", alignItems: "center", gap: 7, background: bg, border: `1.5px solid ${clr}`, borderRadius: 20, padding: "5px 12px", fontSize: 13, fontWeight: 700, color: clr }}>
                <span style={{ width: 14, height: 14, borderRadius: 4, background: clr, display: "inline-block" }}/>{label}
              </span>
            ))}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, flexWrap: "wrap" }}>
            <SearchInput value={mSearch} onChange={setMSearch} placeholder={t("ค้นหา Acc. Code / ชื่อบัญชี","Search Acc. Code / account name")} width={260} big/>
            <span style={{ fontSize: 11, color: T.textMuted }}>{t("คลิกหัวคอลัมน์เพื่อเรียงลำดับ · แสดง","Click a header to sort · showing")} {shownCodes.length}/{codes.length} {t("รายการ","items")}</span>
          </div>
          <div className="fatscroll" style={{ border: `1px solid ${T.cardBorder}`, borderRadius: 12 }}>
            <table style={{ borderCollapse: "collapse", width: "max-content", minWidth: "100%" }}>
              <thead>
                <tr>
                  <th onClick={()=>toggleSort("code")}    style={{ ...hM("#f1f5f9"), ...stickyHead0, textAlign: "left", minWidth: COL1_W, cursor:"pointer", userSelect:"none" }}>Acc. Code{arrow("code")}</th>
                  <th onClick={()=>toggleSort("name")}    style={{ ...hM("#f1f5f9"), ...stickyHead1, textAlign: "left", minWidth: 180, cursor:"pointer", userSelect:"none" }}>Acc. Name{arrow("name")}</th>
                  <th onClick={()=>toggleSort("tender")}  style={{ ...hM(bCost), minWidth: 120, cursor:"pointer", userSelect:"none" }} title={t(`งบ QS = ราคาเดิม + เผื่อเศษ ${WASTE_LBL} (ของราคาเดิม) + งานเพิ่ม`,`QS budget = baseline + ${WASTE_LBL} wastage (on baseline) + additions`)}>Tender Cost<span style={{fontSize:11,fontWeight:600,opacity:0.8,marginLeft:4}}>{t(`รวมเผื่อ ${WASTE_LBL}`,`incl. ${WASTE_LBL}`)}</span>{arrow("tender")}</th>
                  <th onClick={()=>toggleSort("takeoff")} style={{ ...hM(bCost), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Take off{arrow("takeoff")}</th>
                  <th onClick={()=>toggleSort("stock")}   style={{ ...hM(bCost), minWidth: 90, cursor:"pointer", userSelect:"none" }}>Stock{arrow("stock")}</th>
                  <th onClick={()=>toggleSort("issue")}   style={{ ...hM(bCost), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Issue PO{arrow("issue")}</th>
                  <th onClick={()=>toggleSort("pending")} style={{ ...hM(bCost), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Pending PO{arrow("pending")}</th>
                  {months.map(mk => <th key={mk} onClick={()=>toggleSort("m:"+mk)} style={{ ...hM("#eef3ee"), cursor:"pointer", userSelect:"none" }}>{monthLbl(mk)}{arrow("m:"+mk)}</th>)}
                  <th onClick={()=>toggleSort("total")}   style={{ ...hM("#eef3ee"), fontWeight: 700, cursor:"pointer", userSelect:"none" }}>TOTAL{arrow("total")}</th>
                  <th onClick={()=>toggleSort("balcost")} style={{ ...hM("#eaeef5"), minWidth: 110, cursor:"pointer", userSelect:"none" }}>Balance Cost{arrow("balcost")}</th>
                </tr>
              </thead>
              <tbody>
                {shownCodes.map(code => {
                  const bud = budgetOf(code), tko = takeoffOf(code), stk = stockOf(code), bc = balCostOf(code), iss = issuePOof(code), bpo = balPOof(code);
                  return (
                    <tr key={code}>
                      <td style={{ ...cM, ...stickyBody0, fontVariantNumeric: "tabular-nums", fontWeight: 600, color: T.blue }}>{code}</td>
                      <td style={{ ...cM, ...stickyBody1, color: T.textSecondary }}>{nameOf(code)}</td>
                      <td style={{ ...nM, background: bCost, fontWeight: 600, color: T.textPrimary }}>{money(bud)}{bud ? usdLine(bud, usdRate) : null}</td>
                      <td style={{ ...nM, background: bCost, fontWeight: 600, color: T.textPrimary }}>{money(tko)}{tko ? usdLine(tko, usdRate) : null}</td>
                      <td style={{ ...nM, background: bCost, fontWeight: 600, color: T.textPrimary }}>{money(stk)}{stk ? usdLine(stk, usdRate) : null}</td>
                      <td style={{ ...nM, background: bCost, fontWeight: 600, color: T.textPrimary }}>{money(iss)}{iss ? usdLine(iss, usdRate) : null}</td>
                      <td style={{ ...nM, background: bCost, fontWeight: 600, color: bc < 0 ? T.red : T.textPrimary }}>{money(bc)}{bc ? usdLine(Math.abs(bc), usdRate) : null}</td>
                      {months.map(mk => {
                        const c = cellOf(code, mk); const tot = cellTot(c);
                        return (
                          <td key={mk} style={{ ...nM, fontWeight: 600, color: tot ? T.textPrimary : T.textMuted }}>
                            {!tot ? "-" : (<>
                              {c.rec > 0 && <div style={{ color: T.green }}>{fmt(c.rec)}</div>}
                              {c.po > 0 && <div style={{ color: c.poLate ? T.amber : T.textPrimary }}>{fmt(c.po)}{c.poLate ? " ⚠" : ""}</div>}
                              {c.plan > 0 && <div style={{ color: c.planLate ? T.amber : T.red }}>{fmt(c.plan)} *{c.planLate ? "⚠" : ""}</div>}
                              {usdLine(tot, usdRate)}
                            </>)}
                          </td>
                        );
                      })}
                      <td style={{ ...nM, fontWeight: 600, background: "#f6faf6", color: T.textPrimary }}>{rowTot(code) ? fmt(rowTot(code)) : "-"}{rowTot(code) ? usdLine(rowTot(code), usdRate) : null}</td>
                      <td style={{ ...nM, fontWeight: 600, background: "#eaeef5", color: bpo < 0 ? T.red : T.textPrimary }}>{money(bpo)}{bpo ? usdLine(Math.abs(bpo), usdRate) : null}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td style={{ ...cM, ...stickyBody0, fontWeight: 700, background: "#f1f5f9" }} colSpan={2}>TOTAL</td>
                  <td style={{ ...nM, fontWeight: 700, background: "#eef2f7" }}>{money(sum(budgetOf))}{sum(budgetOf) ? usdLine(sum(budgetOf), usdRate) : null}</td>
                  <td style={{ ...nM, fontWeight: 700, background: "#eef2f7" }}>{money(sum(takeoffOf))}{sum(takeoffOf) ? usdLine(sum(takeoffOf), usdRate) : null}</td>
                  <td style={{ ...nM, fontWeight: 700, background: "#eef2f7" }}>{money(sum(stockOf))}{sum(stockOf) ? usdLine(sum(stockOf), usdRate) : null}</td>
                  <td style={{ ...nM, fontWeight: 700, background: "#eef2f7" }}>{money(sum(issuePOof))}{sum(issuePOof) ? usdLine(sum(issuePOof), usdRate) : null}</td>
                  <td style={{ ...nM, fontWeight: 700, background: "#eef2f7", color: sum(balCostOf) < 0 ? T.red : T.textPrimary }}>{money(sum(balCostOf))}{sum(balCostOf) ? usdLine(Math.abs(sum(balCostOf)), usdRate) : null}</td>
                  {months.map(mk => <td key={mk} style={{ ...nM, fontWeight: 650, background: "#e6ede6" }}>{colTot(mk) ? fmt(colTot(mk)) : "-"}{colTot(mk) ? usdLine(colTot(mk), usdRate) : null}</td>)}
                  <td style={{ ...nM, fontWeight: 700, background: "#e6ede6" }}>{grand ? fmt(grand) : "-"}{grand ? usdLine(grand, usdRate) : null}</td>
                  <td style={{ ...nM, fontWeight: 700, background: "#e2e8f2", color: sum(balPOof) < 0 ? T.red : T.textPrimary }}>{money(sum(balPOof))}{sum(balPOof) ? usdLine(Math.abs(sum(balPOof)), usdRate) : null}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      <div style={{ fontSize: 13, fontWeight: 650, color: T.textPrimary, marginBottom: 10 }}>{t("จัดการแผน","Manage plans")}</div>
      {sorted.length === 0 ? (
        <div style={{ textAlign: "center", padding: "52px 0", color: T.textMuted }}>
          <div style={{ marginBottom:10,color:T.textMuted }}><Ico name="calendar" size={32} sw={1.5} /></div>{t("ยังไม่มีแผนของเข้า — กด “+ เพิ่มแผน” เพื่อเริ่ม","No incoming plans — press “+ Add plan” to start")}
        </div>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {sorted.map(pl => {
            const items = poItems(pl);
            const total = items.reduce((s, it) => s + (parseFloat(it.amount) || 0), 0);
            const ds = planDates(pl);
            return (
              <div key={pl.id} style={{ background: T.card, border: `1px solid ${T.cardBorder}`, borderRadius: 14, padding: "14px 18px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
                  <span style={{ background: T.amberBg, color: T.amber, fontWeight: 650, fontSize: 12, padding: "4px 12px", borderRadius: 8 }}>{t("ของเข้า","Incoming")} {ds.length ? lbl(ds[0]) : lbl(pl.date)}{ds.length > 1 ? ` (+${ds.length - 1})` : ""}</span>
                  {pl.supplier?.name && <span style={{ fontSize: 12, color: T.textSecondary }}>· {pl.supplier.name}</span>}
                  <span style={{ fontSize: 12, color: T.textMuted }}>{items.length} {t("รายการ","items")}</span>
                  <span style={{ marginLeft: "auto", textAlign: "right" }}>
                    <span style={{ fontVariantNumeric: "tabular-nums", fontWeight: 650, color: T.textPrimary }}>฿{fmt(total)}</span>
                    {total ? usdLine(total, usdRate) : null}
                  </span>
                  <button onClick={() => onConvert(pl)} className="btn-primary" style={{ fontSize: 12, padding: "6px 12px" }}>→ {t("ทำเป็น PO จริง","Make real PO")}</button>
                  <button onClick={() => onEdit(pl)} className="btn-ghost" style={{ fontSize: 12, padding: "6px 10px" }}><Ico name="edit" /> {t("แก้ไข","Edit")}</button>
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {items.map(it => (
                    <span key={it.id} style={{ background: "#f8fafc", border: `1px solid ${T.cardBorder}`, borderRadius: 8, padding: "5px 10px", fontSize: 12 }}>
                      <b style={{ fontVariantNumeric: "tabular-nums", color: T.blue }}>{it.code}</b> {nameOf(it.code)} · <b style={{ fontVariantNumeric: "tabular-nums" }}>฿{fmt(parseFloat(it.amount) || 0)}</b>
                    </span>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── Procurement View ─────────────────────────────────────────────────────────
function ProcurementView({ project, updateProject, tenderCosts, additions, poEntries, savePO, onBack, onHome, onDept, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[], onExport, setEditMode, incomingPlan={}, saveIncomingPlan }) {
  const usdRate = effRate(project);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const [tab,    setTab]    = useState("list"); // "list" | "tracking"
  const [tabHist, setTabHist] = useState([]);   // ประวัติแท็บ — ปุ่มกลับย้อนทีละหน้า
  const goTab   = (id) => { if (id !== tab) { setTabHist(h => [...h, tab]); setTab(id); } };
  const [trackingOnlyIssues, setTrackingOnlyIssues] = useState(false); // lifted so the alert banner below can jump straight into "only late items"
  const [view,   setView]   = useState("browse"); // "browse" | "add"
  const blankItem = () => ({ id:uid(), code:"", takeoff:"", store:"", amount:"",
    rounds:[{ id:uid(), planDate:"", planAmount:"", actualAmount:"", actualDate:"" }] });
  const emptyForm = () => ({
    date:todayStr(), status:"PO Issued",
    supplier:{ name:"", poNumber:"" },
    paymentType:"", creditDays:DEFAULT_CREDIT_DAYS, notes:"",
    items:[ blankItem() ], isPlan:false,
  });
  const [form,   setForm]   = useState(emptyForm);
  // เตือน "ยังไม่กรอกเลข PO" เฉพาะหลังผู้ใช้ออกจากช่อง หรือกดบันทึกแล้ว (ไม่ขึ้นแดงตั้งแต่เปิดฟอร์ม)
  const [poNoTouched, setPoNoTouched] = useState(false);
  useEffect(() => { setPoNoTouched(false); }, [form.id, view]);
  const [editId, setEditId] = useState(null);
  const [editingPlan, setEditingPlan] = useState(false); // true = กำลังแก้ "แผนของเข้า" (มาจากลิสต์แผน)
  const [payModal, setPayModal] = useState(null);        // {po, date} — ตอนตั้งสถานะ Paid ให้กรอกวันจ่ายเอง
  const [filter, setFilter] = useState("All");
  // กรองเฉพาะ PO ที่มีปัญหา (กดจากชิป "ต้องรีบดู" หรือชิป "มีปัญหา") — ดูได้ในรายการ PO ทันที ไม่ต้องย้ายแท็บ
  const [issueFilter, setIssueFilter] = useState(null);   // null | "any" | "late-incoming" | "late-payment"
  const isPhone = useIsPhone();
  const [search, setSearch] = useState("");
  const [detailId, setDetailId] = useState(null);
  const [collapsed, setCollapsed] = useState({});
  const plans = Array.isArray(incomingPlan) ? incomingPlan : [];
  // อยู่ในโหมดแก้ไขเมื่อเปิดฟอร์มเพิ่ม/แก้ PO หรือเปิดหน้ารายละเอียด (บันทึกของเข้า/แบ่งงวด)
  useEffect(() => { setEditMode?.(view==="add" || detailId!=null); return () => setEditMode?.(false); }, [view, detailId, setEditMode]);

  const detailPO = poEntries.find(p => p.id === detailId) || null;
  const openDetail  = (p) => setDetailId(p.id);
  const closeDetail = () => setDetailId(null);
  const toggleGroup = (code) => setCollapsed(c => ({...c, [code]: !c[code]}));

  // Budget (QS) = baseline Tender Cost + every monthly addition (ค่าธรรมดา +
  // คอลัมน์ย่อย) combined per Acc. Code — ใช้ตัวช่วยกลางเดียวกับ Export ให้ตรงกัน
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);
  // Sum only top-level codes (accounts not hidden + standalone extras).
  // combinedBudget also carries an entry for every sub-item (EX-xxxx with a
  // parentCode) since those persist in tenderCosts/additions individually —
  // their total is already rolled into their parent's value, so summing
  // Object.values(combinedBudget) wholesale would double-count them.
  // ไม่ซ่อน Acc.Code ที่ยังมี PO อยู่ (กันงบของ code นั้นหายจากยอดรวม ทั้งที่ committed ยังนับ)
  const effHidden = hiddenSafeForPO(hiddenAccounts, poEntries);
  // ตัวเลือกในช่อง Account Code ของฟอร์ม PO = รหัสมาตรฐาน (ไม่รวมที่ QS ซ่อน) + "งานเพิ่ม" ของโครงการ
  const pickerOptions = exportAccountList(extraItems, effHidden);
  const topLevelCodes = [
    ...ACCOUNTS.filter(a => !effHidden.includes(a.code)).map(a => a.code),
    ...extraItems.filter(e => !e.parentCode).map(e => e.code),
  ];
  const tenderTotal = topLevelCodes.reduce((s,c) => s + (parseFloat(combinedBudget[c]) || 0), 0);
  const totalComm   = poEntries.reduce((s,p)=>s+poTotal(p),0);
  const totalPaid   = poEntries.reduce((s,p)=>s+poPaidAmount(p),0);
  const paidCount   = poEntries.filter(p=>paymentStatus(p)==="paid").length;

  // Late-item alert counts, shown as a banner regardless of which tab is
  // active so problems surface immediately instead of only inside "ติดตามของเข้า/จ่ายเงิน".
  const lateIncomingCount = poEntries.filter(p=>incomingStatus(p)==="late").length;
  const latePaymentCount  = poEntries.filter(p=>paymentStatus(p)==="late" && p.status!=="Paid").length;

  // Supplier (single) + item (account-code line) helpers
  const updateSupplierField = (key, val) => setForm(f=>({...f, supplier:{...f.supplier, [key]:val}}));
  const addItemRow    = () => setForm(f=>({...f, items:[...f.items, blankItem()]}));
  const removeItemRow = (id) => setForm(f=>({...f, items: f.items.length>1 ? f.items.filter(it=>it.id!==id) : f.items}));
  const updateItemRow = (id, key, val) => setForm(f=>({...f, items: f.items.map(it=>it.id===id?{...it,[key]:val}:it)}));
  // Update the first (order-time) round of an item — used for its แผนของเข้า date.
  const updateItemPlan = (id, key, val) => setForm(f=>({...f, items: f.items.map(it=>{
    if (it.id!==id) return it;
    const rounds = it.rounds && it.rounds.length ? it.rounds.slice() : [{id:uid(),planDate:"",planAmount:"",actualAmount:"",actualDate:""}];
    rounds[0] = {...rounds[0], [key]:val};
    return {...it, rounds};
  })}));

  // Budget / net / % helpers for the % field on each account-code line.
  const budgetForCode = (code) => parseFloat(combinedBudget[code])||0;
  // ยอด PO ที่สั่งไปแล้ว + แผนที่มีอยู่แล้วของ code นี้ (ยกเว้นรายการที่กำลังแก้อยู่)
  const otherCommitted = (code) => poEntries.reduce((s,p)=> (!editingPlan && p.id===editId) ? s : s + poAmountForCode(p, code), 0);
  const otherPlanned   = (code) => plans.reduce((s,pl)=> (editingPlan && pl.id===editId) ? s : s + poAmountForCode(pl, code), 0);
  // ของใน store ที่ "บันทึกไว้แล้ว" ในรายการ PO อื่นของ code นี้ (ยกเว้นรายการที่กำลังแก้อยู่)
  // — ต้องหักด้วย ไม่งั้น "ต้องสั่งสุทธิ" ในฟอร์มจะไม่ตรงกับ Balance Cost ในตาราง
  const otherStock = (code) => poEntries.reduce((s,p)=> (!editingPlan && p.id===editId) ? s : s + poItems(p).filter(it=>it.code===code).reduce((ss,it)=>ss+(parseFloat(it.store)||0),0), 0);
  // "ต้องสั่งสุทธิ" = ยอดที่เหลือต้องสั่งจริง = งบ − store(ในฟอร์ม) − store ที่บันทึกไว้แล้ว − PO ที่สั่งแล้ว − แผนที่มี
  // "ต้องสั่งสุทธิ" อ้างอิงจาก Take off (กรอกเอง) — ไม่ผูกกับงบโครงการ (QS) อีกต่อไป
  const itemNet = (it) => (parseFloat(it.takeoff)||0) - (parseFloat(it.store)||0) - otherStock(it.code) - otherCommitted(it.code) - otherPlanned(it.code);
  const setItemAmount = (id, val) => updateItemRow(id, "amount", val);
  // % ของยอดสั่ง = ช่องกรอกเอง (it.pct) ไม่ผูกกับมูลค่า PO อีกต่อไป

  const formTotal = form.items.reduce((s,it)=>s+(parseFloat(it.amount)||0),0);

  const submit = async () => {
    setPoNoTouched(true);
    // ชื่อ Supplier ไม่บังคับ — ใส่หรือไม่ใส่ก็ได้
    // กันมูลค่าติดลบ (ทำให้ยอดคงเหลือ/งบเพี้ยน)
    if (form.items.some(it=>it.code && (parseFloat(it.amount)||0) < 0)) { uiAlert(t("มูลค่า PO ต้องไม่ติดลบ กรุณาแก้ไขก่อนบันทึก","PO value cannot be negative — please fix before saving")); return; }
    // บรรทัดที่กรอกไม่ครบ — เดิมถูกทิ้งเงียบ ๆ ตอนบันทึก (ถ้ามีบรรทัดอื่นที่ครบ) ตอนนี้แจ้งก่อน
    const amtOf = (it) => parseFloat(it.amount) || 0;
    const lineNo = (it) => form.items.indexOf(it) + 1;
    const noAmt  = form.items.filter(it => it.code && !(amtOf(it) > 0));
    const noCode = form.items.filter(it => !it.code && amtOf(it) > 0);
    if (noAmt.length || noCode.length) {
      const msg = [
        ...noAmt.map(it => t(`บรรทัดที่ ${lineNo(it)} (${it.code}): ยังไม่กรอกมูลค่า หรือมูลค่าเป็น 0`, `Line ${lineNo(it)} (${it.code}): no value, or value is 0`)),
        ...noCode.map(it => t(`บรรทัดที่ ${lineNo(it)}: กรอกมูลค่าแล้วแต่ยังไม่เลือก Account Code`, `Line ${lineNo(it)}: has a value but no Account Code`)),
      ].join("\n");
      uiAlert(t("กรอกรายการไม่ครบ — แก้ไขหรือลบบรรทัดนั้นก่อนบันทึก:\n\n","Some lines are incomplete — fix or remove them before saving:\n\n") + msg);
      return;
    }
    // Acc. Code ซ้ำในใบเดียวกัน — ถามก่อน (ปกติควรรวมเป็นบรรทัดเดียว แล้วแบ่งงวดของเข้าแทน)
    const codes = form.items.filter(it => it.code).map(it => it.code);
    const dups = [...new Set(codes.filter((c, i) => codes.indexOf(c) !== i))];
    if (dups.length && !(await uiConfirm(t(`Acc. Code ซ้ำในใบเดียวกัน: ${dups.join(", ")}\nปกติควรรวมเป็นบรรทัดเดียว — ยืนยันบันทึกแบบนี้?`, `Duplicate Acc. Code in this PO: ${dups.join(", ")}\nUsually these should be one line — save anyway?`), { okLabel: t("บันทึกแบบนี้","Save as is") }))) return;
    const validItems = form.items.filter(it=>it.code && amtOf(it) > 0).map(it=>{
      const rs = (it.rounds && it.rounds.length ? it.rounds : [{id:uid()}]);
      return {
      id: it.id || uid(), code: it.code, takeoff: it.takeoff || "", store: it.store || "", pct: it.pct ?? "", amount: it.amount,
      rounds: rs.map((r,idx)=>({
        id: r.id || uid(),
        planDate: r.planDate || "",
        // มีงวดเดียว = ยอดแผนเท่ากับมูลค่า PO เสมอ (เดิมแก้มูลค่าแล้วยอดแผนค้างค่าเก่า → ตารางของเข้าโชว์ยอดผิด)
        planAmount: idx===0 ? (rs.length === 1 ? it.amount : (r.planAmount || it.amount)) : (r.planAmount || ""),
        actualAmount: r.actualAmount || "",
        actualDate: r.actualDate || "",
      })),
      };
    });
    if (!validItems.length) { uiAlert(t("กรุณาเลือก Account Code และกรอกมูลค่าอย่างน้อย 1 รายการ","Please select an Account Code and enter at least one value")); return; }
    // กันยอดของเข้าจริงรวมทุกงวดเกินยอดสั่งของแต่ละรายการ (แจ้งเตือน + บันทึกไม่ได้)
    const overItem = validItems.find(it => {
      const o = parseFloat(it.amount)||0; if (!(o>0)) return false;
      const rc = it.rounds.reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
      return Math.round(rc*100) > Math.round(o*100);
    });
    if (overItem) { uiAlert(t(`⚠ ${overItem.code}: ยอดของเข้าจริงรวมทุกงวด (${fmt(overItem.rounds.reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0))}) เกินยอดสั่ง ${fmt(overItem.amount)} — แก้ให้ไม่เกินก่อนบันทึก`,`⚠ ${overItem.code}: total received across rounds (${fmt(overItem.rounds.reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0))}) exceeds ordered ${fmt(overItem.amount)} — fix before saving`)); return; }
    // PO จริง (ไม่ใช่แผน) ต้องมีเลข PO เสมอ
    if (!form.isPlan && !(form.supplier.poNumber||"").trim()) { uiAlert(t("PO จริงต้องกรอก \"เลข PO\" ก่อนบันทึก","A real PO needs a PO number before saving")); return; }
    const payload = {
      date: form.date, status: form.status, notes: form.notes || "",
      supplier: { name: form.supplier.name.trim(), poNumber: (form.supplier.poNumber||"").trim() },
      paymentType: form.paymentType || "", creditDays: parseInt(form.creditDays,10) || DEFAULT_CREDIT_DAYS,
      items: validItems,
    };
    const toPlan = !!form.isPlan;   // ติ๊ก "แผนของเข้า" ไว้ไหม

    const prev = editId ? (editingPlan ? plans.find(x=>x.id===editId) : poEntries.find(x=>x.id===editId)) : null;
    const entries = [];
    if (prev && prev.status !== payload.status) entries.push(historyEntry(session, "status", `${t("เปลี่ยนสถานะ","Status change")}: ${prev.status} → ${payload.status}`));
    entries.push(historyEntry(session, prev ? "edited" : "created", toPlan ? (prev?t("แก้ไขแผนของเข้า","Edited incoming plan"):t("สร้างแผนของเข้า","Created incoming plan")) : (prev?t("แก้ไขข้อมูล PO","Edited PO"):t("สร้างรายการ PO","Created PO"))));
    const withLog = { ...payload, isPlan: toPlan, history: [...entries.reverse(), ...(prev?.history||[])].slice(0,40) };

    if (editingPlan) {
      // ต้นทางเป็น "แผน": ติ๊กแผนอยู่ = อัปเดตแผน · เอาติ๊กออก = แปลงแผน → PO จริง (ย้ายออก)
      if (toPlan) saveIncomingPlan(plans.map(pl=>pl.id===editId?{...withLog,id:editId}:pl));
      else { saveIncomingPlan(plans.filter(pl=>pl.id!==editId)); savePO([...poEntries,{...withLog,id:uid()}]); }
    } else if (toPlan) {
      // ต้นทางเป็น PO (หรือรายการใหม่) แต่ติ๊กเป็นแผน → บันทึกเป็นแผน (ถ้าเดิมเป็น PO ให้ย้ายออก)
      if (editId && poEntries.some(p=>p.id===editId)) savePO(poEntries.filter(p=>p.id!==editId));
      saveIncomingPlan([...plans,{...withLog,id:uid()}]);
    } else {
      // PO ปกติ (ใหม่/แก้ไข)
      savePO(editId ? poEntries.map(p=>p.id===editId?{...withLog,id:editId}:p) : [...poEntries,{...withLog,id:uid()}]);
    }
    setEditId(null); setEditingPlan(false);
    setForm(emptyForm());
    setView("browse");
  };

  // Persist an in-place update to a PO's items/rounds (used by the detail view
  // when recording actual goods received or splitting a round). Migrates the
  // record to the new shape on first touch so it's normalised going forward.
  const updatePO = (updated) => {
    // เช็คสิทธิ์กับ "PO ที่บันทึกอยู่จริง" (ไม่ใช่ค่าที่กำลังจะแก้) — PO ที่ปิดแล้วแก้ได้เฉพาะ Admin
    const stored = poEntries.find(x=>x.id===updated.id);
    if (stored && !canEditPO(stored, session)) { uiAlert(t("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ได้เฉพาะ Admin","This PO is fully received & paid — Admin only can edit")); return; }
    savePO(poEntries.map(x=>x.id===updated.id?updated:x));
  };

  // One-click status change — used by the StatusPicker wherever a PO is
  // listed, so procurement doesn't need to open the full edit form just to
  // move a PO from "PO Issued" to "Delivered". Still fully logged. Locked
  // once a PO is fully received + fully paid, unless the current user is admin.
  const applyStatus = (po, newStatus, paidDate) => {
    const patch = { ...po, status: newStatus };
    if (newStatus === "Paid") patch.paidDate = paidDate || todayStr();   // จำวันจ่ายที่กำหนดเอง
    const label = `${t("เปลี่ยนสถานะ","Status change")}: ${poStatusLabel(po.status)} → ${poStatusLabel(newStatus)}` + (newStatus === "Paid" && patch.paidDate ? ` (${t("จ่าย","paid")} ${patch.paidDate})` : "");
    const updated = withHistory(patch, historyEntry(session, "status", label));
    savePO(poEntries.map(x=>x.id===po.id?updated:x));
  };
  const changeStatus = async (po, newStatus) => {
    if (newStatus === po.status) return;
    if (!canEditPO(po, session)) { uiAlert(t("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ไขได้เฉพาะ Admin","This PO is fully received & paid — Admin only can edit")); return; }
    // ตั้งเป็น "Paid" → เตือนถ้ายังไม่มีการรับของเลย แล้วให้กรอกวันจ่ายเองก่อน
    if (newStatus === "Paid") {
      const anyReceived = poRounds(po).some(r => roundReceived(r));
      if (!anyReceived && !(await uiConfirm(t("PO นี้ยังไม่มีการรับของเลย — ยืนยันว่าจ่ายแล้วจริง?","This PO has no received goods yet — confirm it is really paid?"), { okLabel: t("ยืนยันจ่ายแล้ว","Yes, it is paid") }))) return;
      setPayModal({ po, date: po.paidDate || todayStr() });
      return;
    }
    applyStatus(po, newStatus);
  };

  const openEdit = (p) => {
    if (!canEditPO(p, session)) { uiAlert(t("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ไขได้เฉพาะ Admin","This PO is fully received & paid — Admin only can edit")); return; }
    const P = migratePO(p);
    setForm({
      date: P.date, status: P.status, notes: P.notes || "",
      supplier: { name: P.supplier.name || "", poNumber: P.supplier.poNumber || "" },
      paymentType: P.paymentType || "", creditDays: P.creditDays || DEFAULT_CREDIT_DAYS,
      items: P.items.map(it=>({
        id: it.id || uid(), code: it.code || "", takeoff: it.takeoff || "", store: it.store || "", pct: it.pct ?? "", amount: it.amount || "",
        rounds: (it.rounds && it.rounds.length ? it.rounds : [{id:uid(),planDate:"",planAmount:"",actualAmount:"",actualDate:""}])
          .map(r=>({ id:r.id||uid(), planDate:r.planDate||"", planAmount:r.planAmount||"", actualAmount:r.actualAmount||"", actualDate:r.actualDate||"" })),
      })), isPlan:false,
    });
    setEditId(p.id); setEditingPlan(false); setView("add"); setDetailId(null);
  };
  // โหลด PO/แผน (โครงสร้างเดียวกัน) เข้าฟอร์ม — ใช้ร่วมกันทั้งแก้แผนและแปลงเป็น PO
  const loadIntoForm = (p, asPlan) => {
    const P = migratePO(p);
    setForm({
      date: P.date || todayStr(), status: P.status || "PO Issued", notes: P.notes || "",
      supplier: { name: P.supplier?.name || "", poNumber: P.supplier?.poNumber || "" },
      paymentType: P.paymentType || "", creditDays: P.creditDays || DEFAULT_CREDIT_DAYS,
      items: (P.items||[]).map(it=>({
        id: it.id || uid(), code: it.code || "", takeoff: it.takeoff || "", store: it.store || "", pct: it.pct ?? "", amount: it.amount || "",
        rounds: (it.rounds && it.rounds.length ? it.rounds : [{id:uid(),planDate:"",planAmount:"",actualAmount:"",actualDate:""}])
          .map(r=>({ id:r.id||uid(), planDate:r.planDate||"", planAmount:r.planAmount||"", actualAmount:r.actualAmount||"", actualDate:r.actualDate||"" })),
      })), isPlan: asPlan,
    });
    setEditId(p.id); setEditingPlan(true); setDetailId(null); setView("add");
  };
  const openNewPO   = () => { setEditId(null); setEditingPlan(false); setForm({ ...emptyForm(), isPlan:false }); setDetailId(null); setView("add"); };
  const openEditPlan = (pl) => loadIntoForm(pl, true);   // แก้แผน (ติ๊กแผนอยู่)
  const startConvert = (pl) => loadIntoForm(pl, false);  // แปลงแผน → PO (เอาติ๊กออกให้แล้ว กดบันทึกก็เป็น PO)
  const deletePlan = async (id) => {
    const pl = (plans||[]).find(p=>p.id===id);
    const d = pl ? (poRounds(pl).map(r=>r.planDate).filter(Boolean).sort()[0] || pl.date || "") : "";
    const info = pl ? `${d||t("(ไม่มีวัน)","(no date)")}${pl.supplier?.name?` · ${pl.supplier.name}`:""} · ฿${fmt0(poItems(pl).reduce((s,it)=>s+(parseFloat(it.amount)||0),0))}` : "";
    if ((await uiConfirm(t(`ลบแผนของเข้านี้?${info?`\n\n${info}`:""}\n\n(ลบเฉพาะ "แผน" — ไม่กระทบ PO จริง)`,`Delete this incoming plan?${info?`\n\n${info}`:""}\n\n(deletes the "plan" only — real PO unaffected)`), { danger: true, okLabel: t("ลบแผน","Delete plan") }))) { saveIncomingPlan(plans.filter(pl=>pl.id!==id)); return true; }
    return false;   // กดยกเลิก — ให้ผู้เรียกรู้ (ไม่ปิดฟอร์มทิ้ง)
  };
  const deletePO = async (id, confirmed=false) => {
    const po = poEntries.find(x=>x.id===id);
    if (po && !canEditPO(po, session)) { uiAlert(t("PO นี้รับของและจ่ายเงินครบแล้ว — ลบได้เฉพาะ Admin","This PO is fully received & paid — Admin only can delete")); return; }
    // ถามยืนยันก่อนลบ (ลบแล้วย้อนกลับไม่ได้) — ถ้า confirmed=true แปลว่ายืนยันในแอปมาแล้ว
    const label = po ? `${poSupplierName(po)}${poNumbersLabel(po)!=="—"?` · ${poNumbersLabel(po)}`:""} · ฿${fmt(poTotal(po))}` : "";
    if (!confirmed && !(await uiConfirm(t(`ยืนยันการลบรายการ PO นี้?\n\n${label}\n\n⚠ ลบแล้วย้อนกลับไม่ได้`,`Confirm deleting this PO?\n\n${label}\n\n⚠ This cannot be undone`), { danger: true, okLabel: t("ลบ PO","Delete PO") }))) return;
    savePO(poEntries.filter(x=>x.id!==id)); setDetailId(null);
    if (editId === id) closeForm();   // ถ้าลบจากในฟอร์มแก้ไข ให้ปิดฟอร์มกลับหน้ารายการ
  };
  const closeForm = () => { setView("browse"); setEditId(null); setEditingPlan(false); setForm(emptyForm()); };
  // ปุ่ม "กลับ" — ย้อนทีละชั้น: ฟอร์ม → ปิดฟอร์ม, รายละเอียด → ปิด, สลับแท็บ → ย้อนแท็บ,
  // สุดทางแล้วค่อยออกไปหน้าก่อนหน้า (เลือกโครงการ/เลือกโรล)
  const backNav = () => {
    if (view === "add")       return closeForm();
    if (detailId != null)     return closeDetail();
    if (tabHist.length)       { const h = [...tabHist]; const p = h.pop(); setTabHist(h); setTab(p); return; }
    onBack();
  };
  // กด Esc ระหว่างเปิดฟอร์มเพิ่ม/แก้ PO = ยกเลิก (ปิดฟอร์มโดยไม่บันทึก)
  useEffect(() => {
    if (view !== "add") return;
    const onEsc = (e) => { if (e.key === "Escape" && !e.defaultPrevented) { e.preventDefault(); closeForm(); } };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [view]);

  const issueOf = (p) => ({ inc: incomingStatus(p)==="late", pay: paymentStatus(p)==="late" && p.status!=="Paid" });
  const matchesIssue = (p) => { if (!issueFilter) return true; const k = issueOf(p);
    return issueFilter==="late-incoming" ? k.inc : issueFilter==="late-payment" ? k.pay : (k.inc || k.pay); };
  const showIssues = (kind) => {
    const next = issueFilter === kind ? null : kind;      // กดซ้ำ = ยกเลิกตัวกรอง
    setIssueFilter(next); setFilter("All");
    if (tab !== "list") goTab("list");
    if (next) setTimeout(() => document.querySelector("[data-po-list-top]")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  };
  const filtered = poEntries.filter(p=>{
    const itemsText = poItems(p).map(it=>{ const acc=accountOf(it.code); return `${it.code} ${acc?.name||""}`; }).join(" ");
    return (filter==="All"||p.status===filter)&& matchesIssue(p) &&
      (search===""||[itemsText,poSupplierText(p),poNumbersLabel(p)].join(" ").toLowerCase().includes(search.toLowerCase()));
  });

  // Group the filtered POs by Account Code so long lists stay organised and
  // scannable — a PO split across several codes appears once per code, with
  // only that code's share of the amount counted in that group's subtotal.
  const groupedFiltered = {};
  filtered.forEach(p => {
    poItems(p).forEach(it => {
      if (!it.code) return;
      (groupedFiltered[it.code] = groupedFiltered[it.code] || []).push({ po:p, item:it });
    });
  });
  const groupTotals = Object.fromEntries(Object.entries(groupedFiltered).map(([c,rows])=>[c, rows.reduce((s,{item})=>s+(parseFloat(item.amount)||0),0)]));
  const sortedGroupCodes = Object.keys(groupedFiltered).sort();

  return (
    <Shell role="procurement" color={T.amber} project={project} onBack={backNav} onHome={onHome} onDept={onDept} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      {payModal && (
        <div onClick={()=>setPayModal(null)} style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.5)",zIndex:320,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:16,width:"min(380px,100%)",overflow:"hidden",boxShadow:"0 24px 60px rgba(15,23,42,0.3)"}}>
            <div style={{padding:"16px 20px 4px",color:T.textPrimary,fontWeight:650,fontSize:16,display:"flex",alignItems:"center",gap:8}}><Ico name="wallet" size={18} color={T.textSecondary} />{t("บันทึกการจ่ายเงิน","Record payment")}</div>
            <div style={{padding:20,display:"flex",flexDirection:"column",gap:12}}>
              <div style={{fontSize:12,color:T.textSecondary}}>{poSupplierName(payModal.po)} · <b>฿{fmt(poTotal(payModal.po))}</b></div>
              <label style={{fontSize:12,color:T.textSecondary,display:"flex",flexDirection:"column",gap:5}}>
                {t("วันที่จ่ายเงิน","Payment date")}
                <DateInput value={payModal.date} onChange={e=>setPayModal(m=>({...m,date:e.target.value}))} style={{padding:"8px 10px"}}/>
              </label>
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",gap:10,padding:"14px 20px",borderTop:`1px solid ${T.cardBorder}`}}>
              <button onClick={()=>setPayModal(null)} className="btn-ghost">{t("ยกเลิก","Cancel")}</button>
              <button onClick={()=>{ if(!payModal.date){uiAlert(t("เลือกวันที่จ่าย","Select a payment date"));return;} applyStatus(payModal.po,"Paid",payModal.date); setPayModal(null); }} className="btn-primary">{t("บันทึกจ่ายแล้ว","Save as paid")}</button>
            </div>
          </div>
        </div>
      )}
      {isPhone && view!=="add" && <BottomNav items={[
        { key:"list",   icon:"clipboard", label:t("รายการ PO","PO List"), on:tab==="list", onClick:()=>goTab("list") },
        { key:"inplan", icon:"calendar",  label:t("แผนของเข้า","Incoming plan"), on:tab==="inplan", onClick:()=>goTab("inplan") },
        { key:"add",    icon:"plus",      label:t("เพิ่ม PO","Add PO"), onClick:openNewPO },
        { key:"export", icon:"download",  label:"Export", onClick:onExport },
      ]} />}
      <div style={{padding:"24px 28px"}}>
        {view!=="add" && (
          <div style={{display:"flex",gap:8,marginBottom:20,alignItems:"center",flexWrap:"wrap"}}>
            {!isPhone && (
            <div className="seg-tabs" role="tablist">
            {[["list",t("รายการ PO","PO List")],["inplan",t("แผนของเข้า","Incoming plan")]].map(([id,label])=>(
              <button key={id} role="tab" aria-selected={tab===id} onClick={()=>goTab(id)} className={`seg-tab${tab===id?" on":""}`}>
                {label}
              </button>
            ))}
            </div>
            )}
            <div style={{marginLeft:"auto"}}><CurrencyControl project={project} updateProject={updateProject}/></div>
            {!isPhone && (
            <button onClick={onExport} className="btn-ghost" style={{display:"flex",alignItems:"center",gap:6}}>
              <Ico name="download" /> Export Excel
            </button>
            )}
          </div>
        )}
        <div className="stat-grid has-lead" style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
          <StatCard lead progress={tenderTotal>0 ? totalComm/tenderTotal*100 : null} label={t(`งบคงเหลือ (รวมเผื่อเศษ ${WASTE_LBL})`,`Budget remaining (incl. ${WASTE_LBL} wastage)`)} value={"฿"+fmt0(tenderTotal-totalComm)} thb={tenderTotal-totalComm} rate={usdRate} sub={tenderTotal>0?`${t("ใช้ไป","Used")} ${((totalComm/tenderTotal)*100).toFixed(1)}% ${t("ของงบ","of budget")} ฿${fmt0(tenderTotal)}`:"—"} color={tenderTotal-totalComm<0?T.red:T.textSecondary} icon={tenderTotal-totalComm<0?"⚠️":"💰"} accent={tenderTotal-totalComm<0?T.redBg:"#f8fafc"}/>
          <StatCard label={t("ผูกพันแล้ว (PO)","Committed (PO)")} value={"฿"+fmt0(totalComm)} thb={totalComm} rate={usdRate} sub={`${poEntries.length} ${t("รายการ","items")}`} color={T.amber} icon="📦" accent={T.amberBg}/>
          <StatCard label={t("ชำระแล้ว","Paid")} value={"฿"+fmt0(totalPaid)} thb={totalPaid} rate={usdRate} sub={`${paidCount} ${t("รายการ","items")} · ${t("จ่ายอัตโนมัติ","auto-paid")}`} color={T.green} icon="✅" accent={T.greenBg}/>
        </div>

        {/* เรื่องที่ต้องรีบดู = ชิปนับจำนวน (เดิมเป็นแถบแดงเต็มความกว้าง) — กดแล้วไปดูเฉพาะรายการที่มีปัญหา */}
        {view!=="add" && (lateIncomingCount>0 || latePaymentCount>0) && (
          <div role="group" aria-label={t("เรื่องที่ต้องรีบดู","Needs attention")} style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:16}}>
            <span style={{fontSize:13,color:T.textSecondary,fontWeight:600,display:"inline-flex",alignItems:"center",gap:6}}><Ico name="alert" size={16} color={T.red} />{t("ต้องรีบดู","Needs attention")}</span>
            {lateIncomingCount>0 && (
              <button className="att-chip" data-attention="late-incoming" aria-pressed={issueFilter==="late-incoming"} onClick={()=>showIssues("late-incoming")}>
                <b>{lateIncomingCount}</b> {t("ของเข้าล่าช้า","late incoming")}
              </button>
            )}
            {latePaymentCount>0 && (
              <button className="att-chip" data-attention="late-payment" aria-pressed={issueFilter==="late-payment"} onClick={()=>showIssues("late-payment")}>
                <b>{latePaymentCount}</b> {t("จ่ายเงินเกินกำหนด","overdue payments")}
              </button>
            )}
          </div>
        )}

        {view==="add" ? (
          <div style={{display:"flex",gap:20,alignItems:"flex-start",flexWrap:"wrap"}}>
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:28,flex:"1 1 520px",maxWidth:680,minWidth:0,animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
              <div>
                <div style={{fontSize:15,fontWeight:650,color:T.textPrimary}}>{(editId||editingPlan) ? (form.isPlan?t("แก้ไขแผนของเข้า","Edit incoming plan"):t("บันทึกเป็น PO จริง","Save as real PO")) : (form.isPlan?t("เพิ่มแผนของเข้า","Add incoming plan"):t("เพิ่ม PO ใหม่","Add new PO"))}</div>
                <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>{t("เลือกด้านล่างว่าจะบันทึกเป็น PO จริง หรือ แผนของเข้า (ฟอร์มเดียวกัน)","Choose below: save as a real PO or an incoming plan (same form)")}</div>
              </div>
              <button onClick={closeForm} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted}}>×</button>
            </div>
            {/* สวิตช์: PO จริง / แผนของเข้า — ติ๊กแผนจากลิสต์แผนแล้วเปลี่ยนเป็น PO = แปลงเป็น PO จริง */}
            <div style={{display:"flex",gap:6,marginBottom:18,background:T.bg,padding:4,borderRadius:10,width:"fit-content"}}>
              {[[t("PO จริง","Real PO"),false,T.blue],[t("แผนของเข้า","Incoming plan"),true,T.amber]].map(([label,val,clr])=>(
                <button key={label} onClick={()=>setForm(f=>({...f,isPlan:val}))}
                  style={{border:"none",borderRadius:8,padding:"7px 18px",fontSize:13,fontWeight:600,cursor:"pointer",
                    background: form.isPlan===val ? clr : "transparent", color: form.isPlan===val ? "#fff" : T.textSecondary}}>{label}</button>
              ))}
            </div>
            <div className="po-form-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <FormStep first n={1} title={form.isPlan ? t("ข้อมูลแผน","Plan details") : t("ข้อมูล PO","PO details")} />
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{t("วันที่สั่ง PO","PO order date")}</span>
                <DateInput value={form.date} onChange={e=>setForm(f=>({...f, date:e.target.value}))}/>
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{t("สถานะ","Status")}</span>
                <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} className="input-base">
                  {PO_STATUS.map(s=><option key={s} value={s}>{poStatusLabel(s)}</option>)}
                </select>
              </label>

              {/* Supplier — exactly one vendor per PO. */}
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>Supplier <span style={{color:T.textMuted,fontWeight:400}}>({t("ไม่บังคับ","optional")})</span></span>
                <input placeholder={t("ชื่อ Supplier","Supplier name")} value={form.supplier.name} onChange={e=>updateSupplierField("name",e.target.value)} className="input-base"/>
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{t("เลข PO","PO no.")} {form.isPlan ? <span style={{color:T.textMuted,fontWeight:400}}>({t("ไม่บังคับ","optional")})</span> : <span style={{color:T.red}}>*</span>}</span>
                <input placeholder={t("เช่น PO-2026-001","e.g. PO-2026-001")} value={form.supplier.poNumber} onChange={e=>updateSupplierField("poNumber",e.target.value)} onBlur={()=>setPoNoTouched(true)} className="input-base"
                  aria-required={!form.isPlan ? true : undefined}
                  aria-invalid={poNoTouched && !form.isPlan && !(form.supplier.poNumber||"").trim() ? true : undefined}
                  style={poNoTouched && !form.isPlan && !(form.supplier.poNumber||"").trim() ? {borderColor:T.red, background:T.redBg} : undefined}/>
                {poNoTouched && !form.isPlan && !(form.supplier.poNumber||"").trim() && <span role="alert" data-po-error style={{fontSize:12,color:T.red}}>{t("กรอกเลข PO ก่อนบันทึก","Enter the PO number before saving")}</span>}
              </label>

              {/* Account-code line items — each carries its own store amount and
                  its % of the net-to-purchase (budget − store). */}
              <FormStep n={2} title={<>{t("รายการ","Items")} <span style={{color:T.red}}>*</span></>} hint={t("เลือก Account Code แล้วกรอกมูลค่า · Take off / ของใน store / % ไม่บังคับ","Pick an Account Code and enter the value · Take off / store / % are optional")} />
              <div style={{gridColumn:"1/-1",display:"flex",flexDirection:"column",gap:12}}>
                {form.items.map((it)=>{
                  const budget = budgetForCode(it.code);
                  const net = itemNet(it);
                  // ยังไม่กรอก Take off → ยังคำนวณ "ต้องสั่งสุทธิ" ไม่ได้ (เดิมขึ้นแดง "เกิน" ทั้งที่ยังไม่ได้กรอก)
                  const hasTakeoff = String(it.takeoff ?? "").trim() !== "" && (parseFloat(it.takeoff)||0) !== 0;
                  const amt = parseFloat(it.amount)||0;
                  return (
                  <div key={it.id} style={{border:`1px solid ${T.cardBorder}`,borderRadius:12,padding:14,background:T.bg}}>
                    <div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:8,alignItems:"center",marginBottom:12}}>
                      <AccountPicker value={it.code} onChange={code=>updateItemRow(it.id,"code",code)} options={pickerOptions} />
                      <button type="button" onClick={()=>removeItemRow(it.id)} disabled={form.items.length===1}
                        style={{background:"none",border:"none",color:form.items.length===1?T.textMuted:T.red,cursor:form.items.length===1?"default":"pointer",padding:"4px 8px",fontSize:15,opacity:form.items.length===1?0.4:1,display:"inline-grid",placeItems:"center"}} aria-label={t("ลบบรรทัดนี้","Remove this line")}><Ico name="trash" size={17} /></button>
                    </div>
                    {/* แถวบน: Take off (กรอกเอง) · store · ต้องสั่งสุทธิ (อ่านอย่างเดียว) */}
                    <div className="po-item-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>Take off <span style={{color:T.textMuted,fontWeight:400}}>({t("กรอกเอง","manual")})</span></span>
                        <MoneyInput value={it.takeoff} onChange={v=>updateItemRow(it.id,"takeoff",v)}/>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>{t("มีใน store","In store")}</span>
                        <MoneyInput value={it.store} onChange={v=>updateItemRow(it.id,"store",v)}/>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        {hasTakeoff ? <>
                        <span style={{fontSize:11,color:net<0?T.red:T.amber,fontWeight:500}}>{t("ต้องสั่งสุทธิ","Net to order")} {net<0?t("(เกิน)","(over)"):""}</span>
                        <input className="input-base" readOnly tabIndex={-1} data-net value={net<0?`-${fmtMoneyInput(Math.abs(net))}`:fmtMoneyInput(net)}
                          style={{textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:600,background:net<0?T.redBg:T.amberBg,color:net<0?T.red:T.amber,borderColor:"transparent"}}/>
                        </> : <>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>{t("ต้องสั่งสุทธิ","Net to order")}</span>
                        <input className="input-base" readOnly tabIndex={-1} data-net value="–" title={t("กรอก Take off ก่อน จึงคำนวณได้","Enter Take off first")}
                          style={{textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:500,background:"#f1f5f9",color:T.textMuted,borderColor:"transparent"}}/>
                        </>}
                      </label>
                    </div>
                    {/* แถวล่าง: มูลค่า PO · % · แผนของเข้า — ความสูงเท่ากันหมด */}
                    <div className="po-item-grid" style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>{t("มูลค่า PO นี้ (THB)","This PO value (THB)")}</span>
                        <MoneyInput value={it.amount} onChange={v=>setItemAmount(it.id,v)}/>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>{t("% ของยอดสั่ง (กรอกเอง)","% of order (manual)")}</span>
                        <div style={{position:"relative",display:"flex",alignItems:"center"}}>
                          <input type="number" placeholder="0" value={it.pct ?? ""} onChange={e=>updateItemRow(it.id,"pct",e.target.value)}
                            className="input-base" style={{textAlign:"right",fontVariantNumeric:"tabular-nums",flex:1,paddingRight:26}}/>
                          <span style={{position:"absolute",right:11,fontSize:13,color:(it.pct??"")!==""?T.textPrimary:T.textMuted,fontWeight:600,pointerEvents:"none"}}>%</span>
                        </div>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>{t("แผนของเข้า (งวดแรก)","Incoming plan (1st round)")}</span>
                        <DateInput value={it.rounds?.[0]?.planDate||""} onChange={e=>updateItemPlan(it.id,"planDate",e.target.value)}/>
                      </label>
                    </div>
                    {it.code && ((it.pct??"")!=="" || amt>0) && (
                      <div style={{marginTop:10,fontSize:11,color:T.textSecondary}}>
                        {t("% ของยอดสั่ง PO นี้","% of this PO order")}: <b style={{color:(parseFloat(it.pct)||0)>100?T.red:T.textPrimary,fontSize:12}}>{(it.pct??"")!=="" ? `${it.pct}%` : "—"}</b> <span style={{color:T.textMuted}}>({t("ที่กรอกเอง","manual")})</span>
                        {amt>0 && <span style={{color:T.textMuted}}> · {t("ยอดจริง","actual")} {fmt(amt)} = {budget>0?Math.round(amt/budget*100):0}% {t("ของงบรวม","of total budget")}</span>}
                      </div>
                    )}
                  </div>
                  );
                })}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <button type="button" onClick={addItemRow} className="btn-ghost" style={{padding:"6px 12px",fontSize:12}}>+ {t("เพิ่ม Account Code","Add Account Code")}</button>
                  {form.items.length>1 && <span style={{fontSize:12,color:T.textSecondary}}>{t("รวม","Total")}: <b style={{color:T.amber,fontVariantNumeric:"tabular-nums"}}>{fmt(formTotal)}</b></span>}
                </div>
              </div>

              <FormStep n={3} title={t("การจ่ายเงิน","Payment")} />
              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{t("วิธีจ่ายเงิน","Payment method")}</span>
                <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                  <button type="button" onClick={()=>setForm(f=>({...f,paymentType:"cash"}))}
                    style={{flex:"1 1 160px",padding:"10px 14px",borderRadius:10,border:`1.5px solid ${form.paymentType==="cash"?T.green:T.cardBorder}`,background:form.paymentType==="cash"?T.greenBg:T.card,color:form.paymentType==="cash"?T.green:T.textSecondary,fontSize:13,fontWeight:600,cursor:"pointer",transition:"all 0.15s"}}>
                    {t("เงินสด","Cash")} <span style={{fontWeight:450,fontSize:11,opacity:0.8}}>({t("จ่ายวันของเข้า","pay on arrival")})</span>
                  </button>
                  <button type="button" onClick={()=>setForm(f=>({...f,paymentType:"credit",creditDays:f.creditDays||DEFAULT_CREDIT_DAYS}))}
                    style={{flex:"1 1 160px",padding:"10px 14px",borderRadius:10,border:`1.5px solid ${form.paymentType==="credit"?T.blue:T.cardBorder}`,background:form.paymentType==="credit"?T.blueLight:T.card,color:form.paymentType==="credit"?T.blue:T.textSecondary,fontSize:13,fontWeight:600,cursor:"pointer",transition:"all 0.15s"}}>
                    {t("เครดิต","Credit")}
                  </button>
                  {form.paymentType==="credit" && (
                    <span style={{display:"flex",alignItems:"center",gap:6}}>
                      <input type="number" value={form.creditDays} onChange={e=>setForm(f=>({...f,creditDays:e.target.value}))} className="input-base" style={{width:76}}/>
                      <span style={{fontSize:12,color:T.textSecondary}}>{t("วัน","days")}</span>
                    </span>
                  )}
                </div>
                {form.paymentType && (
                  <span style={{fontSize:11,color:T.blue}}>{t("วันครบกำหนดจ่ายคำนวณอัตโนมัติจาก \"วันของเข้าจริง\" ของแต่ละงวด","Due date is auto-calculated from each round's actual arrival date")}{form.paymentType==="credit"?` + ${form.creditDays||DEFAULT_CREDIT_DAYS} ${t("วัน","days")}`:""} — {t("จ่ายอัตโนมัติเมื่อถึงกำหนด","auto-paid when due")}</span>
                )}
              </label>

              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{t("หมายเหตุ","Notes")}</span>
                <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} className="input-base" style={{resize:"vertical"}}/>
              </label>
            </div>
            {/* เว้นที่ด้านล่างให้พ้นแถบ "ย้อนกลับ/ทำซ้ำ" ที่ลอยมุมซ้ายล่าง ไม่ให้ทับปุ่ม */}
            <div style={{display:"flex",gap:10,marginTop:20,marginBottom:76,flexWrap:"wrap",alignItems:"center"}}>
              <button onClick={submit} className="btn-primary">{editingPlan && !form.isPlan ? t("แปลงเป็น PO จริง","Convert to real PO") : form.isPlan ? t("บันทึกแผน","Save plan") : (editId?t("บันทึก","Save"):t("เพิ่ม PO","Add PO"))}</button>
              <button onClick={closeForm} className="btn-ghost">{t("ยกเลิก","Cancel")}</button>
              {editId && (
                <button onClick={()=>{ if (editingPlan) { deletePlan(editId).then(ok => { if (ok) closeForm(); }); } else deletePO(editId); }}
                  style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:6,background:T.redBg,border:`1px solid #fecaca`,color:T.red,borderRadius:10,padding:"9px 16px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
                  <Ico name="trash" size={15} /> {editingPlan ? t("ลบแผนนี้","Delete this plan") : t("ลบ PO นี้","Delete this PO")}
                </button>
              )}
            </div>
          </div>
          {/* สรุปงบข้างฟอร์ม — มีข้อมูลตั้งแต่เปิดฟอร์ม (ภาพรวมโครงการ) แล้วแตกเป็นรายรหัสเมื่อเลือก Account Code
              สมการแบบขั้นบันได: งบ − ที่ใช้ไปแล้ว − ใบนี้ = คงเหลือ พร้อมแถบสัดส่วน และวันครบจ่ายโดยประมาณ */}
          {(() => {
            const amtOf = (it) => parseFloat(it.amount) || 0;
            const codes = [...new Set(form.items.filter(it => it.code).map(it => it.code))];
            const money = (v) => <span style={{fontVariantNumeric:"tabular-nums"}}>{fmt(v)}</span>;
            const thisLbl = form.isPlan ? t("แผนนี้","This plan") : t("PO ใบนี้","This PO");
            const row = (label, v, opts={}) => (
              <div data-sum-row={opts.id} style={{display:"flex",justifyContent:"space-between",gap:10,fontSize:13,padding:"3px 0",color:opts.color||T.textSecondary,fontWeight:opts.bold?650:400}}>
                <span>{label}</span><span>{opts.minus && v ? "− " : ""}{money(v)}</span>
              </div>
            );
            // แถบ: ส่วนที่ใช้ไปแล้ว (เทา) + ใบนี้ (เหลือง) เทียบกับงบ — เกินงบเป็นสีแดง
            const bar = (budget, used, mine) => {
              const tot = Math.max(budget, used + mine, 1);
              const over = used + mine > budget;
              return (
                <div aria-hidden="true" style={{display:"flex",height:8,borderRadius:99,background:"#eef2f7",overflow:"hidden",margin:"8px 0 2px"}}>
                  <span style={{width:`${used/tot*100}%`,background:over?T.red:"#94a3b8"}}/>
                  <span style={{width:`${mine/tot*100}%`,background:over?"#f87171":T.amber}}/>
                </div>
              );
            };
            const result = (remain) => (
              <div style={{borderTop:`1px dashed ${T.cardBorder}`,marginTop:4,paddingTop:4}}>
                {row(remain < 0 ? t("เกินงบ","Over budget") : t("คงเหลือหลังใบนี้","Left after this"), Math.abs(remain), { id:"remain", color: remain < 0 ? T.red : T.green, bold:true })}
              </div>
            );
            const projOther = poEntries.reduce((s,p)=> (!editingPlan && p.id===editId) ? s : s + poTotal(p), 0);
            const projRemain = tenderTotal - projOther - formTotal;
            // วันครบจ่ายโดยประมาณ: วันของเข้าตามแผน (งวดแรก) + เครดิต
            const credit = form.paymentType==="credit" ? (parseInt(form.creditDays)||DEFAULT_CREDIT_DAYS) : 0;
            const dues = form.items.filter(it => it.rounds?.[0]?.planDate).map(it => ({ id: it.id, code: it.code, plan: it.rounds[0].planDate, due: addDays(it.rounds[0].planDate, credit) }));
            return (
              <aside data-po-summary style={{flex:"0 1 320px",minWidth:260,position:"sticky",top:16,background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:18}}>
                <div style={{fontSize:14,fontWeight:650,color:T.textPrimary,marginBottom:4,display:"flex",alignItems:"center",gap:8}}><Ico name="chart" size={16} color={T.textSecondary} />{t("สรุปงบ","Budget summary")}</div>
                <div style={{fontSize:12,color:T.textMuted,marginBottom:10}}>{t("คำนวณตามที่กรอกในฟอร์มนี้ (ยังไม่บันทึก)","Based on this form (not saved yet)")}</div>

                <div data-sum="project" style={{padding:"8px 0 10px"}}>
                  <div style={{fontSize:13,fontWeight:650,color:T.textPrimary,marginBottom:4}}>{t("ทั้งโครงการ","Whole project")}</div>
                  {row(t(`งบ QS (รวมเผื่อเศษ ${WASTE_LBL})`,`QS budget (incl. ${WASTE_LBL})`), tenderTotal, { id:"budget" })}
                  {row(t("PO อื่นที่สั่งแล้ว","Other POs ordered"), projOther, { id:"other", minus:true })}
                  {row(thisLbl, formTotal, { id:"this", minus:true, color: formTotal ? T.amber : T.textSecondary })}
                  {result(projRemain)}
                  {bar(tenderTotal, projOther, formTotal)}
                </div>

                {codes.length === 0 ? (
                  <div style={{fontSize:12,color:T.textMuted,borderTop:`1px solid ${T.cardBorder}`,padding:"10px 0 2px"}}>{t("เลือก Account Code ในส่วนที่ 2 เพื่อดูงบของรหัสนั้น","Pick an Account Code in section 2 to see that code's budget")}</div>
                ) : codes.map(code => {
                  const acc = accountOf(code);
                  const budget = parseFloat(combinedBudget[code]) || 0;
                  const lines = form.items.filter(it => it.code === code);
                  const thisAmt = lines.reduce((s, it) => s + amtOf(it), 0);
                  const stock = otherStock(code) + lines.reduce((s, it) => s + (parseFloat(it.store) || 0), 0);
                  const ordered = otherCommitted(code), planned = otherPlanned(code);
                  const remain = budget - stock - ordered - thisAmt;
                  return (
                    <div key={code} data-sum="code" style={{borderTop:`1px solid ${T.cardBorder}`,padding:"10px 0"}}>
                      <div style={{fontSize:13,marginBottom:6}}><b style={{color:T.blue,fontVariantNumeric:"tabular-nums"}}>{code}</b> <span style={{color:T.textSecondary}}>{acc?.name||""}</span></div>
                      {row(t("งบ (QS)","Budget (QS)"), budget)}
                      {row(t("Stock","Stock"), stock, { minus:true })}
                      {row(t("PO อื่นที่สั่งแล้ว","Other POs ordered"), ordered, { minus:true })}
                      {row(thisLbl, thisAmt, { color:T.amber, minus:true })}
                      {result(remain)}
                      {bar(budget, stock + ordered, thisAmt)}
                      {planned > 0 && <div style={{fontSize:12,color:T.textMuted,marginTop:4}}>{t("ยังมีแผนของเข้าที่ยังไม่เป็น PO","Incoming plans not yet PO")}: {fmt(planned)}</div>}
                    </div>
                  );
                })}

                <div data-sum="due" style={{borderTop:`1px solid ${T.cardBorder}`,paddingTop:10,marginTop:2}}>
                  <div style={{fontSize:13,fontWeight:650,color:T.textPrimary,marginBottom:4,display:"flex",alignItems:"center",gap:6}}><Ico name="calendar" size={15} color={T.textSecondary} />{t("วันครบจ่าย (ประมาณ)","Pay due (estimate)")}</div>
                  {dues.length === 0 ? (
                    <div style={{fontSize:12,color:T.textMuted}}>{t("ใส่ \"แผนของเข้า (งวดแรก)\" เพื่อดูวันครบจ่าย","Enter \"Incoming plan (1st round)\" to see the due date")}</div>
                  ) : dues.map(d => (
                    <div key={d.id} style={{fontSize:12,color:T.textSecondary,padding:"2px 0",display:"flex",justifyContent:"space-between",gap:8,flexWrap:"wrap"}}>
                      <span>{d.code ? <b style={{color:T.blue,fontWeight:600}}>{d.code}</b> : null} {t("ของเข้า","in")} {fmtDate(d.plan)}</span>
                      <span style={{color:form.paymentType?T.textPrimary:T.textMuted,fontWeight:600}}>{form.paymentType ? `${t("ครบจ่าย","due")} ${fmtDate(d.due)}` : t("เลือกวิธีจ่ายในส่วนที่ 3","choose payment in section 3")}</span>
                    </div>
                  ))}
                  {dues.length>0 && form.paymentType && <div style={{fontSize:11,color:T.textMuted,marginTop:2}}>{form.paymentType==="credit" ? t(`ของเข้า + เครดิต ${credit} วัน`,`arrival + ${credit}-day credit`) : t("เงินสด: จ่ายวันของเข้า","Cash: pay on arrival")}</div>}
                </div>
              </aside>
            );
          })()}
          </div>
        ) : tab==="inplan" ? (
          <>
            <IncomingPlanTab plans={plans} poEntries={poEntries} usdRate={usdRate} tenderCosts={tenderCosts} additions={additions} extraItems={extraItems} hiddenAccounts={hiddenAccounts} onNew={openNewPO} onEdit={openEditPlan} onConvert={startConvert} onDelete={deletePlan} />
            {/* ติดตามของเข้า/จ่ายเงิน — ย้ายมาไว้ใต้ "จัดการแผน" (เอาแท็บติดตามแยกออก) */}
            <div style={{marginTop:28,paddingTop:20,borderTop:`2px solid ${T.cardBorder}`}}>
              <div style={{fontSize:15,fontWeight:650,color:T.textPrimary,marginBottom:14,display:"flex",alignItems:"center",gap:8}}><Ico name="truck" size={18} color={T.textSecondary} />{t("ติดตามของเข้า / จ่ายเงิน","Track incoming / payments")}</div>
              <ProcurementTrackingTab poEntries={poEntries} onEdit={openEdit} onView={openDetail} onAddNew={openNewPO}
                onStatusChange={changeStatus} session={session} usdRate={usdRate}
                tenderCosts={tenderCosts} additions={additions} extraItems={extraItems} hiddenAccounts={hiddenAccounts}
                onlyIssues={trackingOnlyIssues} setOnlyIssues={setTrackingOnlyIssues} />
            </div>
          </>
        ) : (
          <>
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <SearchInput value={search} onChange={setSearch} placeholder={t("ค้นหา Account, supplier, PO...","Search Account, supplier, PO...")} width={isPhone?"100%":240}/>
              {(() => {
                // จำนวนต่อสถานะ (ตามคำค้นหา) — ชิปบอกตัวเลขในตัว ไม่ต้องกดดูก่อน
                const q = search.toLowerCase();
                const bySearch = poEntries.filter(p => q==="" || [poItems(p).map(it=>`${it.code} ${accountOf(it.code)?.name||""}`).join(" "),poSupplierText(p),poNumbersLabel(p)].join(" ").toLowerCase().includes(q));
                const nIssue = bySearch.filter(p => { const k = issueOf(p); return k.inc || k.pay; }).length;
                return (
                  <div className="chip-scroll" role="group" aria-label={t("กรองตามสถานะ","Filter by status")} style={{display:"flex",gap:5,flex:1,flexWrap:"wrap"}}>
                    {nIssue > 0 && (
                      <button onClick={()=>showIssues("any")} aria-pressed={!!issueFilter} data-po-filter="issues"
                        style={{background:issueFilter?T.red:T.redBg,border:`1.5px solid ${issueFilter?T.red:"#fecaca"}`,borderRadius:8,padding:"4px 11px",color:issueFilter?"#fff":T.red,fontSize:12,cursor:"pointer",fontWeight:600,minHeight:32,whiteSpace:"nowrap",flexShrink:0,display:"inline-flex",alignItems:"center",gap:5}}>
                        <Ico name="alert" size={13} /> {t("มีปัญหา","Needs attention")} <span style={{fontVariantNumeric:"tabular-nums"}}>{nIssue}</span>
                      </button>
                    )}
                    {["All",...PO_STATUS].map(st=>{ const n = st==="All" ? bySearch.length : bySearch.filter(p=>p.status===st).length; const on = filter===st; return (
                      <button key={st} onClick={()=>setFilter(st)} aria-pressed={on} data-po-filter={st}
                        style={{background:on?T.textPrimary:"transparent",border:`1.5px solid ${on?T.textPrimary:T.cardBorder}`,borderRadius:8,padding:"4px 11px",color:on?"#fff":T.textSecondary,fontSize:12,cursor:"pointer",fontWeight:500,transition:"all 0.15s",minHeight:32,whiteSpace:"nowrap",flexShrink:0}}>
                        {poStatusLabel(st)} <span style={{opacity:0.75,fontVariantNumeric:"tabular-nums"}}>{n}</span>
                      </button>
                    ); })}
                  </div>
                );
              })()}
              {filtered.length>0 && (
                <button onClick={()=>{
                    const allCollapsed = Object.keys(groupTotals).every(c=>collapsed[c]);
                    const next = {}; Object.keys(groupTotals).forEach(c=>{ next[c] = !allCollapsed; });
                    setCollapsed(next);
                  }}
                  className="btn-ghost" style={{padding:"7px 14px",fontSize:12,display:"inline-flex",alignItems:"center",gap:6}}>
                  {Object.keys(groupTotals).length>0 && Object.keys(groupTotals).every(c=>collapsed[c])
                    ? <><Ico name="chevrons" size={14} />{t("ขยายทั้งหมด","Expand all")}</>
                    : <><Ico name="chevrons" size={14} style={{transform:"rotate(180deg)"}} />{t("ย่อทั้งหมด","Collapse all")}</>}
                </button>
              )}
              {!isPhone && <button onClick={openNewPO} className="btn-primary">+ {t("เพิ่ม PO","Add PO")}</button>}{/* มือถือ: ใช้ปุ่ม "เพิ่ม PO" ในแถบล่าง */}
            </div>

            <div data-po-list-top style={{scrollMarginTop:12}} />
            {issueFilter && (
              <div data-issue-bar role="status" style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",background:T.redBg,border:"1px solid #fecaca",borderRadius:10,padding:"8px 12px",marginBottom:12,fontSize:13,color:T.red}}>
                <Ico name="filter" size={15} />
                <span style={{fontWeight:600}}>{t("กำลังแสดงเฉพาะ","Showing only")}: {issueFilter==="late-incoming" ? t("ของเข้าล่าช้า","late incoming") : issueFilter==="late-payment" ? t("จ่ายเงินเกินกำหนด","overdue payments") : t("PO ที่มีปัญหา (ของเข้าล่าช้า / จ่ายเกินกำหนด)","POs with problems (late incoming / overdue payment)")}</span>
                <span style={{color:T.textSecondary}}>· {filtered.length} {t("ใบ","POs")}</span>
                <button className="btn-ghost" onClick={()=>setIssueFilter(null)} style={{marginLeft:"auto",fontSize:12,padding:"4px 10px"}}>{t("แสดงทั้งหมด","Show all")}</button>
              </div>
            )}
            {filtered.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
                <div style={{marginBottom:12,color:T.textMuted}}><Ico name="clipboard" size={32} sw={1.5} /></div>
                <div style={{fontSize:14,fontWeight:500,color:T.textSecondary,marginBottom:6}}>{poEntries.length===0?t("ยังไม่มีรายการ","No items yet"):t("ไม่พบรายการที่ตรงเงื่อนไข","No items match")}</div>
                <div style={{fontSize:12}}>{poEntries.length===0?t('กด "+ เพิ่ม PO" เพื่อเริ่มต้น','Press "+ Add PO" to start'):t("ลองล้างตัวกรอง หรือคำค้นหา","Try clearing filters or search")}</div>
              </div>
            ) : (
              (() => {
                // ข้อมูลของแต่ละกลุ่ม (Acc. Code) — ใช้ทั้งตารางเดียว (จอกว้าง) และการ์ด (มือถือ)
                const groups = sortedGroupCodes.map(code => {
                  const rows = groupedFiltered[code].slice().sort((a,b)=> (b.po.date||"").localeCompare(a.po.date||""));
                  const grpBudget = parseFloat(combinedBudget[code]) || 0;
                  const grpCommitted = poEntries.reduce((s,p)=>s+poAmountForCode(p,code),0);
                  const grpStock = poEntries.reduce((s,p)=>s+poItems(p).filter(it=>it.code===code).reduce((ss,it)=>ss+(parseFloat(it.store)||0),0),0);
                  return { code, acc: accountOf(code), rows, isCollapsed: !!collapsed[code],
                    total: rows.reduce((s,{item})=>s+(parseFloat(item.amount)||0),0), budget: grpBudget, toOrder: grpBudget - grpStock - grpCommitted };
                });
                const grandTotal = filtered.reduce((s,p)=>s+poTotal(p),0);
                const chip = (bg, clr, txt, key) => <span key={key} style={{background:bg,color:clr,fontSize:12,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{txt}</span>;
                const recvText = (p) => { const d = poReceivedDates(p); return d.length===0 ? "" : d.length===1 ? fmtDate(d[0]) : `${fmtDate(d[0])} (+${d.length-1})`; };
                const budgetNote = (g) => (
                  <span style={{fontSize:12,color:T.textMuted,whiteSpace:"nowrap"}}>
                    {t("งบ","Budget")} <b style={{color:T.textSecondary,fontVariantNumeric:"tabular-nums",fontWeight:600}}>฿{fmt0(g.budget)}</b>
                    {" · "}
                    <span style={{color:g.toOrder<0?T.red:T.textMuted,fontWeight:g.toOrder<0?650:400}}>{g.toOrder<0 ? t("เกินงบ","Over budget") : t("ต้องสั่งเพิ่ม","To order")} <b style={{color:g.toOrder<0?T.red:T.amber,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(Math.abs(g.toOrder))}</b></span>
                  </span>
                );
                const editBtn = (p, locked) => (
                  <button onClick={e=>{ e.stopPropagation(); openEdit(p); }} disabled={locked} aria-label={t("แก้ไข","Edit")} title={locked?t("แก้ไขได้เฉพาะ Admin","Admin only"):t("แก้ไข (ลบได้ในหน้านี้)","Edit (delete available here)")}
                    style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textSecondary,cursor:locked?"not-allowed":"pointer",padding:"6px 8px",minWidth:36,minHeight:36,borderRadius:8,display:"inline-grid",placeItems:"center"}}><Ico name="edit" size={17} /></button>
                );

                // ── มือถือ: การ์ด 1 ใบต่อ PO (ตาราง 960px ต้องเลื่อนซ้ายขวา)
                if (isPhone) return (
                  <div data-po-cards style={{display:"flex",flexDirection:"column",gap:14}}>
                    {groups.map(g => (
                      <section key={g.code} data-po-group={g.code}>
                        <button onClick={()=>toggleGroup(g.code)} aria-expanded={!g.isCollapsed}
                          style={{width:"100%",border:"none",background:"none",padding:"2px 2px 8px",display:"flex",alignItems:"baseline",gap:8,cursor:"pointer",textAlign:"left",flexWrap:"wrap"}}>
                          <Ico name="chevrons" size={14} color={T.textMuted} style={{transform:g.isCollapsed?"rotate(-90deg)":"none",alignSelf:"center"}} />
                          <b style={{color:T.blue,fontSize:13,fontVariantNumeric:"tabular-nums"}}>{g.code}</b>
                          <span style={{fontSize:13,fontWeight:600,color:T.textPrimary,flex:1,minWidth:0}}>{g.acc?.name || "—"}</span>
                          <span style={{fontSize:13,fontWeight:650,color:T.textPrimary,fontVariantNumeric:"tabular-nums"}}>{fmt(g.total)}</span>
                          <span style={{flexBasis:"100%",paddingLeft:22}}>{budgetNote(g)} <span style={{fontSize:12,color:T.textMuted}}>· {g.rows.length} {t("รายการ","items")}</span></span>
                        </button>
                        {!g.isCollapsed && (
                          <div style={{display:"flex",flexDirection:"column",gap:8}}>
                            {g.rows.map(({po:p,item}) => {
                              const inc = incomingStatus(p), pay = paymentStatus(p), locked = !canEditPO(p, session);
                              const rcv = recvText(p);
                              return (
                                <div key={p.id+"-"+(item.id||item.code)} data-po-card role="button" tabIndex={0}
                                  onClick={()=>openDetail(p)} onKeyDown={e=>{ if (e.key==="Enter") openDetail(p); }}
                                  style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:12,padding:"12px 14px",cursor:"pointer"}}>
                                  <div style={{display:"flex",alignItems:"baseline",gap:8}}>
                                    <b style={{fontSize:14,color:T.textPrimary,flex:1,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{poNumbersLabel(p)}</b>
                                    <span style={{fontSize:15,fontWeight:700,color:T.textPrimary,fontVariantNumeric:"tabular-nums"}}>฿{fmt(item.amount)}</span>
                                  </div>
                                  <div style={{fontSize:12,color:T.textSecondary,marginTop:2}}>{itemSupplierName(p,item) || "—"} · {fmtDate(p.date)}{poItems(p).length>1 && <> · {t("จาก","from")} {poItems(p).length} {t("รหัส","codes")}</>}</div>
                                  {usdLine(parseFloat(item.amount)||0, usdRate)}
                                  <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:8}}>
                                    {chip(INCOMING_BG[inc], INCOMING_CLR[inc], rcv ? `${incLabel(inc)} ${rcv}` : incLabel(inc), "i")}
                                    {p.status!=="Paid" && chip(PAYMENT_BG[pay], PAYMENT_CLR[pay], payLabel(pay), "p")}
                                    {p.paymentType && chip(PAYMENT_TYPE_BG[p.paymentType], PAYMENT_TYPE_CLR[p.paymentType], payTypeLabelT(p), "m")}
                                  </div>
                                  <div style={{display:"flex",alignItems:"center",gap:8,marginTop:8,fontSize:12,color:T.textMuted}} onClick={e=>e.stopPropagation()}>
                                    <span style={{whiteSpace:"nowrap"}}>{t("จ่าย","Pay")}: <PayDateText po={p}/></span>
                                    <span style={{flex:1}}/>
                                    <StatusPicker status={p.status} onChange={st=>changeStatus(p,st)} disabled={locked} compact/>
                                    {locked && <span title={t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only")}><Ico name="lock" size={14} color={T.textMuted} /></span>}
                                    {editBtn(p, locked)}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </section>
                    ))}
                    <div style={{display:"flex",justifyContent:"space-between",gap:12,padding:"2px 4px",color:T.textMuted,fontSize:12}}>
                      <span>{filtered.length} {t("รายการทั้งหมด","items total")}</span>
                      <span style={{color:T.textPrimary,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{fmt(grandTotal)}</span>
                    </div>
                  </div>
                );

                // ── จอกว้าง: ตารางเดียว หัวคอลัมน์ชุดเดียว แถวหัวกลุ่มต่อ Acc. Code (เดิมเป็นกล่องแยก หัวคอลัมน์ซ้ำทุกกล่อง)
                const th = (align="left") => ({padding:"10px 14px",textAlign:align,color:T.textMuted,fontWeight:600,fontSize:12,borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap",background:"#fff",position:"sticky",top:0,zIndex:1});
                const td = {padding:"10px 14px",verticalAlign:"top"};
                return (
                  <div data-po-table style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                    <div className="hscroll"><table style={{width:"100%",minWidth:900,borderCollapse:"collapse",fontSize:13}}>
                      <thead><tr>
                        <th style={th()}>{t("PO / Supplier","PO / Supplier")}</th>
                        <th style={th()}>{t("วันเปิด PO","Open date")}</th>
                        <th style={th("right")}>{t("มูลค่า (THB)","Value (THB)")}</th>
                        <th style={th()}>{t("ของเข้า","Delivery")}</th>
                        <th style={th()}>{t("จ่ายเงิน","Payment")}</th>
                        <th style={th()}>{t("สถานะ","Status")}</th>
                        <th style={{...th(),width:52}}><span className="sr-only">{t("แก้ไข","Edit")}</span></th>
                      </tr></thead>
                      {groups.map(g => (
                        <tbody key={g.code} data-po-group={g.code}>
                          <tr className="po-grp" onClick={()=>toggleGroup(g.code)} aria-expanded={!g.isCollapsed} style={{background:"#f8fafc",cursor:"pointer",borderTop:`1px solid ${T.cardBorder}`,borderBottom:`1px solid ${T.cardBorder}`}}>
                            <td colSpan={2} style={{padding:"10px 14px"}}>
                              <span style={{display:"inline-flex",alignItems:"center",gap:8}}>
                                <Ico name="chevrons" size={14} color={T.textMuted} style={{transform:g.isCollapsed?"rotate(-90deg)":"none",transition:"transform 0.15s"}} />
                                <b style={{color:T.blue,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{g.code}</b>
                                <span style={{color:T.textPrimary,fontWeight:600}}>{g.acc?.name || "—"}</span>
                                <span style={{color:T.textMuted,fontSize:12}}>· {g.rows.length} {t("รายการ","items")}</span>
                              </span>
                            </td>
                            <td style={{padding:"10px 14px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:650,color:T.textPrimary,whiteSpace:"nowrap"}}>
                              {fmt(g.total)}{usdRate>0 && <div className="usd-sub" style={{color:T.green,fontWeight:600,fontSize:12}}>≈ ${fmt(g.total/usdRate)}</div>}
                            </td>
                            <td colSpan={4} style={{padding:"10px 14px"}}>{budgetNote(g)}</td>
                          </tr>
                          {!g.isCollapsed && g.rows.map(({po:p,item}) => {
                            const inc = incomingStatus(p), pay = paymentStatus(p), locked = !canEditPO(p, session);
                            const rcv = recvText(p);
                            return (
                              <tr key={p.id+"-"+(item.id||item.code)} onClick={()=>openDetail(p)} className="po-row"
                                style={{borderBottom:`1px solid #f1f5f9`,cursor:"pointer"}}>
                                <td style={td}>
                                  <div style={{fontWeight:600,color:T.textPrimary,fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap"}}>{poNumbersLabel(p)}</div>
                                  <div style={{fontSize:12,color:T.textSecondary}}>{itemSupplierName(p,item) || "—"}</div>
                                </td>
                                <td style={{...td,color:T.textSecondary,fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap"}}>{fmtDate(p.date)}</td>
                                <td style={{...td,textAlign:"right"}}>
                                  <div style={{color:T.textPrimary,fontVariantNumeric:"tabular-nums",fontWeight:600}}>{fmt(item.amount)}</div>
                                  {usdLine(parseFloat(item.amount)||0, usdRate)}
                                  {poItems(p).length>1 && <div style={{fontSize:12,color:T.textMuted}}>{t("จาก","from")} {poItems(p).length} {t("รหัส · รวม","codes · total")} {fmt(poTotal(p))}</div>}
                                </td>
                                <td style={td}>
                                  {chip(INCOMING_BG[inc], INCOMING_CLR[inc], incLabel(inc))}
                                  {rcv && <div style={{fontSize:12,color:T.textSecondary,marginTop:4,fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap"}}>{t("รับ","In")} {rcv}</div>}
                                </td>
                                <td style={td}>
                                  <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                                    {p.status!=="Paid" && chip(PAYMENT_BG[pay], PAYMENT_CLR[pay], payLabel(pay), "p")}
                                    {p.paymentType && chip(PAYMENT_TYPE_BG[p.paymentType], PAYMENT_TYPE_CLR[p.paymentType], payTypeLabelT(p), "m")}
                                  </div>
                                  <div style={{fontSize:12,marginTop:4,fontVariantNumeric:"tabular-nums",whiteSpace:"nowrap",color:poPaidDate(p)?T.green:T.textSecondary,fontWeight:poPaidDate(p)?600:400}}><PayDateText po={p}/></div>
                                </td>
                                <td style={td} onClick={e=>e.stopPropagation()}>
                                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                                    <StatusPicker status={p.status} onChange={st=>changeStatus(p,st)} disabled={locked} compact/>
                                    {locked && <span title={t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only")} style={{fontSize:12}}><Ico name="lock" size={14} color={T.textMuted} /></span>}
                                  </div>
                                  {poLastUpdate(p) && <div style={{fontSize:12,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>{t("อัปเดต","Updated")} {relativeTime(poLastUpdate(p).at)} · {poLastUpdate(p).user}</div>}
                                </td>
                                <td style={{...td,whiteSpace:"nowrap",textAlign:"right"}} onClick={e=>e.stopPropagation()}>{editBtn(p, locked)}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      ))}
                      <tfoot>
                        <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
                          <td colSpan={2} style={{padding:"12px 14px",color:T.textMuted,fontSize:13}}>{filtered.length} {t("รายการทั้งหมด","items total")}</td>
                          <td style={{padding:"12px 14px",textAlign:"right",fontVariantNumeric:"tabular-nums",fontWeight:700,color:T.textPrimary,whiteSpace:"nowrap"}}>
                            {fmt(grandTotal)}{usdRate>0 && <div className="usd-sub" style={{color:T.green,fontWeight:650,fontSize:12}}>≈ ${fmt(grandTotal/usdRate)}</div>}
                          </td>
                          <td colSpan={4}/>
                        </tr>
                      </tfoot>
                    </table></div>
                  </div>
                );
              })()
            )}
          </>
        )}
      </div>
      <PODetailModal key={detailPO?.id || "none"} po={detailPO} onClose={closeDetail} onEdit={openEdit} onDelete={deletePO} onStatusChange={changeStatus} onChangePO={updatePO} session={session} usdRate={usdRate} />
    </Shell>
  );
}

// ─── Procurement: Incoming / Payment Tracking tab ─────────────────────────────
// Groups every PO by its Account Code so the team can see, at a glance and per
// cost line, which deliveries and payments are on track vs. overdue.
function ProcurementTrackingTab({ poEntries, onEdit, onView, onAddNew, onlyIssues, setOnlyIssues, onStatusChange, session, usdRate=0, tenderCosts={}, additions={}, extraItems=[], hiddenAccounts=[] }) {
  const trkBudget = buildCombinedBudget(tenderCosts, additions);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all"); // กรองตามสถานะของเข้า/จ่าย
  const [groupBy, setGroupBy] = useState("code"); // "code" = จัดกลุ่มตาม Acc. Code · "po" = จัดกลุ่มตาม PO
  const STATUS_FILTERS = [["all","ทั้งหมด","All"],["pending","รอของเข้า","Awaiting"],["late","ล่าช้า","Late"],["received","รับแล้ว","Received"],["payPending","รอจ่าย","Awaiting pay"],["paid","จ่ายแล้ว","Paid"]];
  const matchesStatus = (p) => {
    const inc = incomingStatus(p), pay = paymentStatus(p);
    switch (statusFilter) {
      case "pending":    return inc === "pending";
      case "late":       return inc === "late" || pay === "late";
      case "received":   return inc === "received" || inc === "partial";
      case "payPending": return pay === "pending";
      case "paid":       return pay === "paid";
      default:           return true;
    }
  };
  // Which Account-Code groups are collapsed — lets a busy board with many
  // PO rows be tidied away group by group instead of scrolling forever.
  const [collapsed, setCollapsed] = useState(() => new Set());
  const toggleGroup = (code) => setCollapsed(prev => {
    const next = new Set(prev);
    next.has(code) ? next.delete(code) : next.add(code);
    return next;
  });


  const q = search.toLowerCase();
  const passesFilter = (p) => {
    const itemsText = poItems(p).map(it=>{ const acc=accountOf(it.code); return `${it.code} ${acc?.name||""}`; }).join(" ");
    const matchesSearch = q==="" || [itemsText,poSupplierText(p),poNumbersLabel(p)].join(" ").toLowerCase().includes(q);
    const hasIssue = incomingStatus(p)==="late" || paymentStatus(p)==="late";
    return matchesSearch && (!onlyIssues || hasIssue) && matchesStatus(p);
  };

  const filteredEntries = poEntries.filter(passesFilter);

  // Group by Account Code (via each PO's line items), sorted by code — a PO
  // split across several codes shows once per code, sharing the same
  // delivery-batch info since deliveries belong to the whole PO.
  const groups = {};
  filteredEntries.forEach(p => {
    poItems(p).forEach(it => {
      if (!it.code) return;
      (groups[it.code] = groups[it.code] || []).push({ po:p, item:it });
    });
  });
  const sortedCodes = Object.keys(groups).sort();
  // จัดกลุ่มตาม PO (ทางเลือก) — หนึ่งใบต่อกลุ่ม เรียงตามวันเปิด PO แล้วเลข PO
  const posSorted = filteredEntries.slice().sort((a,b)=> (a.date||"").localeCompare(b.date||"") || poNumbersLabel(a).localeCompare(poNumbersLabel(b)));
  // คีย์ที่ใช้ย่อ/ขยายทั้งหมด ตามโหมดที่เลือก
  const allGroupKeys = groupBy==="po" ? posSorted.map(p=>p.id) : sortedCodes;

  // committed/stock ต่อ code ต้องคิดจาก PO "ทั้งหมด" ไม่ใช่เฉพาะที่ผ่านตัวกรอง
  // (ไม่งั้น "ต้องสั่งเพิ่ม" จะเพี้ยน/เกินจริงเมื่อเปิดฟิลเตอร์สถานะหรือเฉพาะล่าช้า)
  const committedAll = {}, stockAll = {};
  poEntries.forEach(p => poItems(p).forEach(it => {
    if (!it.code) return;
    committedAll[it.code] = (committedAll[it.code] || 0) + (parseFloat(it.amount) || 0);
    stockAll[it.code]     = (stockAll[it.code]     || 0) + (parseFloat(it.store)  || 0);
  }));

  const DateCell = ({ value, lateTint }) => (
    <span style={{fontVariantNumeric:"tabular-nums",fontSize:13,color:value?(lateTint?T.red:T.textPrimary):T.textMuted,fontWeight:value&&lateTint?650:450}}>
      {value ? fmtDate(value) : "—"}
    </span>
  );
  const Badge = ({ text, clr, bg }) => (
    <span style={{background:bg,color:clr,fontSize:12,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{text}</span>
  );
  // Renders every delivery batch on a PO — one line per shipment, so a PO
  // that arrives in 2-3 batches shows each plan → actual date with its own status.
  const DeliveryList = ({ po }) => {
    const deliveries = poDeliveries(po);
    const multiSupplier = poSuppliers(po).length > 1;
    if (!deliveries.length) return <span style={{fontSize:13,color:T.textMuted}}>—</span>;
    return (
      <div style={{display:"flex",flexDirection:"column",gap:3}}>
        {deliveries.map((d,i)=>{
          const st = deliveryStatus(d);
          return (
            <div key={d.id||i} style={{display:"flex",alignItems:"center",gap:5}}>
              {deliveries.length>1 && <span style={{fontSize:12,color:T.textMuted,fontWeight:650,minWidth:14}}>#{i+1}</span>}
              {multiSupplier && <span style={{fontSize:12,color:T.textSecondary,fontWeight:600,whiteSpace:"nowrap"}}>{d.supplierName||"—"}:</span>}
              <DateCell value={d.plan} lateTint={st==="late"}/>
              <span style={{color:T.textMuted,fontSize:12}}>→</span>
              <span style={{fontVariantNumeric:"tabular-nums",fontSize:13,color:st==="received"?T.green:T.textMuted,fontWeight:st==="received"?600:450}}>{d.actual ? fmtDate(d.actual) : t("รอ","Pending")}</span>
              {(() => {
                const received = st === "received";   // รับจริงแล้ว (วันรับมาถึงแล้ว) → เขียว
                // แสดง "ยอดของเข้าจริง" ถ้ากรอกไว้แล้ว (ให้ตรงกับหน้ารายละเอียด) ไม่มีค่อยใช้ยอดแผน
                const amt = (parseFloat(d.actualAmount)||0) || (parseFloat(d.planAmount)||0);
                if (!amt) return <span style={{fontSize:12,color:T.textMuted,fontVariantNumeric:"tabular-nums"}}>(—)</span>;
                return <span style={{fontSize:12,color:received?T.green:T.textMuted,fontVariantNumeric:"tabular-nums",fontWeight:received?650:450}}>({fmt(amt)})</span>;
              })()}
            </div>
          );
        })}
      </div>
    );
  };

  // แถวเดียว (ต่อ item) ใช้ได้ทั้งโหมดจัดกลุ่มตาม Acc. Code และตาม PO
  // showAcc=true → โชว์ Acc. Code/ชื่อบัญชีแทนคอลัมน์ วันเปิด/Supplier/PO (ใช้ในโหมดจัดกลุ่มตาม PO)
  const renderRow = ({po:p,item},i,showAcc=false) => {
    // แถวนี้แยกตาม item → คิด/แสดงเฉพาะงวดของ item นี้ ไม่เอางวดของ code อื่นในใบเดียวกันมาปน
    const pItem = { ...p, items:[item] };
    const inc = incomingStatus(pItem), pay = paymentStatus(pItem);
    const splitAcrossCodes = poItems(p).length>1;
    const locked = !canEditPO(p, session);
    const receivedDates = poReceivedDates(pItem);
    const paidDate = poPaidDate(pItem);
    const acc = accountOf(item.code);
    return (
      <tr key={p.id+"-"+(item.id||item.code)} onClick={()=>onView?.(p)}
        style={{background:i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`,cursor:onView?"pointer":"default"}}
        onMouseEnter={e=>e.currentTarget.style.background="#fef9ec"}
        onMouseLeave={e=>e.currentTarget.style.background=i%2===0?T.card:"#fafbfd"}>
        {showAcc ? (
          <>
            <td style={{padding:"9px 16px",color:T.blue,fontSize:13,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{item.code||"—"}</td>
            <td style={{padding:"9px 16px",color:T.textSecondary,fontSize:13}} colSpan={2}>{acc?.name||"—"}</td>
          </>
        ) : (
          <>
            <td style={{padding:"9px 16px",color:T.textMuted,fontSize:13,fontVariantNumeric:"tabular-nums"}}>{fmtDate(p.date)}</td>
            <td style={{padding:"9px 16px",color:T.textPrimary,fontWeight:500}}>{itemSupplierName(p,item)}</td>
            <td style={{padding:"9px 16px",color:T.textMuted,fontVariantNumeric:"tabular-nums",fontSize:13,whiteSpace:"nowrap"}}>{poNumbersLabel(p)}</td>
          </>
        )}
        <td style={{padding:"9px 16px",textAlign:"right"}}>
          <div style={{color:T.textPrimary,fontVariantNumeric:"tabular-nums",fontWeight:600}}>{fmt(item.amount)}</div>
          {usdLine(parseFloat(item.amount)||0, usdRate)}
          {!showAcc && splitAcrossCodes && <div style={{fontSize:12,color:T.textMuted}}>{t("รวม","total")} {fmt(poTotal(p))}</div>}
        </td>
        <td style={{padding:"9px 16px",fontSize:13,fontVariantNumeric:"tabular-nums",color:receivedDates.length?T.textPrimary:T.textMuted}}>
          {receivedDates.length===0 ? "—" : receivedDates.length===1 ? fmtDate(receivedDates[0]) : `${fmtDate(receivedDates[0])} (+${receivedDates.length-1})`}
        </td>
        <td style={{padding:"9px 16px",fontSize:13,fontVariantNumeric:"tabular-nums",color:paidDate?T.green:T.textMuted,fontWeight:paidDate?600:450}}>
          <PayDateText po={pItem}/>
        </td>
        <td style={{padding:"9px 16px"}}><DeliveryList po={pItem}/></td>
        <td style={{padding:"9px 16px"}}>
          <DateCell value={poNextDueDate(pItem)} lateTint={false}/>
          {p.paymentType && (
            <div style={{marginTop:3}}>
              <Badge text={payTypeLabelT(p)} clr={PAYMENT_TYPE_CLR[p.paymentType]} bg={PAYMENT_TYPE_BG[p.paymentType]}/>
            </div>
          )}
        </td>
        <td style={{padding:"9px 16px"}}>
          <div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"flex-start"}}>
            <Badge text={incLabel(inc)} clr={INCOMING_CLR[inc]} bg={INCOMING_BG[inc]}/>
            <Badge text={payLabel(pay)} clr={PAYMENT_CLR[pay]} bg={PAYMENT_BG[pay]}/>
          </div>
        </td>
        <td style={{padding:"9px 16px"}} onClick={e=>e.stopPropagation()}>
          <div style={{display:"flex",alignItems:"center",gap:4}}>
            <StatusPicker status={p.status} onChange={s=>onStatusChange?.(p,s)} disabled={locked} compact/>
            {locked && <span title={t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only")} style={{fontSize:12}}><Ico name="lock" size={14} color={T.textMuted} /></span>}
          </div>
          {poLastUpdate(p) && <div style={{fontSize:12,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>{t("อัปเดต","Updated")} {relativeTime(poLastUpdate(p).at)}</div>}
        </td>
        <td style={{padding:"9px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
          <button onClick={()=>onEdit(p)} disabled={locked} title={locked?t("แก้ไขได้เฉพาะ Admin","Admin only"):t("แก้ไข","Edit")}
            style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textMuted,cursor:locked?"not-allowed":"pointer",padding:"6px 8px",minWidth:34,minHeight:34,borderRadius:8,display:"inline-grid",placeItems:"center"}} aria-label={t("แก้ไข","Edit")}><Ico name="edit" size={16} /></button>
        </td>
      </tr>
    );
  };
  const thStyle = (align) => ({padding:"9px 16px",textAlign:align,color:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`});
  const theadRow = (mode) => (
    <tr>
      {mode==="po" ? (<>
        <th style={thStyle("left")}>Acc. Code</th>
        <th colSpan={2} style={thStyle("left")}>Account Name</th>
      </>) : (<>
        <th style={thStyle("left")}>{t("วันเปิด PO","Open date")}</th>
        <th style={thStyle("left")}>Supplier</th>
        <th style={thStyle("left")}>PO No.</th>
      </>)}
      {[["มูลค่า (THB)","Value (THB)"],["วันรับของ","Received"],["วันจ่าย","Pay date"],["การส่งของ","Delivery"],["แผนจ่ายเงิน","Payment plan"],["ติดตาม","Track"],["สถานะ","Status"],["",""]].map(([h,he],ci)=>(
        <th key={ci} style={thStyle(h==="มูลค่า (THB)"?"right":"left")}>{t(h,he)}</th>
      ))}
    </tr>
  );

  return (
    <div>
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <SearchInput value={search} onChange={setSearch} placeholder={t("ค้นหา Acc. Code, supplier, PO...","Search Acc. Code, supplier, PO...")} width={240}/>
        <button onClick={()=>setOnlyIssues(v=>!v)}
          style={{background:onlyIssues?T.red:"transparent",border:`1.5px solid ${onlyIssues?T.red:T.cardBorder}`,borderRadius:8,padding:"7px 14px",color:onlyIssues?"#fff":T.textSecondary,fontSize:13,cursor:"pointer",fontWeight:600}}>
          <Ico name="alert" size={14} /> {t("แสดงเฉพาะรายการล่าช้า","Show late only")}
        </button>
        <select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}
          style={{padding:"7px 12px",border:`1.5px solid ${statusFilter!=="all"?T.amber:T.cardBorder}`,borderRadius:8,fontSize:13,fontWeight:600,color:statusFilter!=="all"?T.amber:T.textSecondary,background:"#fff",cursor:"pointer"}}>
          {STATUS_FILTERS.filter(([k])=>k!=="late").map(([k,l,e])=><option key={k} value={k}>{t(l,e)}</option>)}
        </select>
        <select value={groupBy} onChange={e=>{ setGroupBy(e.target.value); setCollapsed(new Set()); }}
          style={{padding:"7px 12px",border:`1.5px solid ${T.cardBorder}`,borderRadius:8,fontSize:13,fontWeight:600,color:T.textSecondary,background:"#fff",cursor:"pointer"}}>
          <option value="code">{t("จัดกลุ่ม: ตาม Acc. Code","Group: by Acc. Code")}</option>
          <option value="po">{t("จัดกลุ่ม: ตาม PO","Group: by PO")}</option>
        </select>
        {/* ปุ่มเดียวสลับ ย่อ/ขยายทั้งหมด (เดิมแยก 2 ปุ่ม) — แบบเดียวกับหน้ารายการ PO */}
        {allGroupKeys.length>0 && (() => {
          const allCollapsed = allGroupKeys.every(k => collapsed.has(k));
          return (
            <button data-toggle-all aria-expanded={!allCollapsed} onClick={()=>setCollapsed(allCollapsed ? new Set() : new Set(allGroupKeys))}
              className="btn-ghost" style={{padding:"7px 14px",fontSize:12,display:"inline-flex",alignItems:"center",gap:6}}>
              {allCollapsed
                ? <><Ico name="chevrons" size={14} />{t("ขยายทั้งหมด","Expand all")}</>
                : <><Ico name="chevrons" size={14} style={{transform:"rotate(180deg)"}} />{t("ย่อทั้งหมด","Collapse all")}</>}
            </button>
          );
        })()}
        <div style={{flex:1}}/>
      </div>

      {filteredEntries.length===0 ? (
        <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
          <div style={{marginBottom:12,color:T.textMuted}}><Ico name="truck" size={32} sw={1.5} /></div>
          <div style={{fontSize:14,fontWeight:500,color:T.textSecondary,marginBottom:6}}>{t("ไม่พบรายการที่ตรงเงื่อนไข","No items match")}</div>
          <div style={{fontSize:13}}>{t("ลองล้างตัวกรอง หรือคำค้นหา","Try clearing filters or search")}</div>
        </div>
      ) : groupBy==="po" ? (
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {posSorted.map(p => {
            const items = poItems(p);
            const isCollapsed = collapsed.has(p.id);
            const inc = incomingStatus(p), pay = paymentStatus(p);
            return (
              <div key={p.id} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                <div onClick={()=>toggleGroup(p.id)}
                  style={{padding:"12px 18px",background:"#f8fafc",borderBottom:isCollapsed?"none":`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:10,cursor:"pointer",userSelect:"none",flexWrap:"wrap"}}>
                  <span style={{fontSize:12,color:T.textMuted,transform:isCollapsed?"rotate(-90deg)":"none",transition:"transform 0.15s",display:"inline-block",width:12}}>▼</span>
                  <span style={{color:T.blue,fontSize:13,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{poNumbersLabel(p)}</span>
                  <span style={{color:T.textPrimary,fontSize:13,fontWeight:600}}>{poSupplierName(p)}</span>
                  <span style={{fontSize:12,color:T.textMuted,fontVariantNumeric:"tabular-nums"}}>{t("เปิด","Opened")} {fmtDate(p.date)}</span>
                  <span style={{flex:1}}/>
                  <span style={{fontSize:12,color:T.textMuted}}>{t("มูลค่า","Value")} <b style={{color:T.textSecondary,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(poTotal(p))}</b></span>
                  <span style={{color:T.textMuted,fontSize:12}}>{items.length} {t("รายการ","items")}</span>
                  <Badge text={incLabel(inc)} clr={INCOMING_CLR[inc]} bg={INCOMING_BG[inc]}/>
                  <Badge text={payLabel(pay)} clr={PAYMENT_CLR[pay]} bg={PAYMENT_BG[pay]}/>
                </div>
                {!isCollapsed && (
                <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
                  <thead>{theadRow("po")}</thead>
                  <tbody>{items.map((it,ii)=>renderRow({po:p,item:it},ii,true))}</tbody>
                </table></div>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {sortedCodes.map(code => {
            const acc = accountOf(code);
            const rows = groups[code];
            const lateCount = rows.filter(({po:p})=>incomingStatus(p)==="late"||paymentStatus(p)==="late").length;
            const isCollapsed = collapsed.has(code);
            // งบ + ยอดที่ต้องสั่งเพิ่ม (งบ − ของใน store − PO ที่สั่งแล้วของ code นี้)
            // ใช้ committed/stock จาก PO ทั้งหมด (committedAll/stockAll) ไม่ผูกกับตัวกรอง
            const grpBudget = parseFloat(trkBudget[code]) || 0;
            const grpCommitted = committedAll[code] || 0;
            const grpStock = stockAll[code] || 0;
            const grpToOrder = grpBudget - grpStock - grpCommitted;
            return (
              <div key={code} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                <div onClick={()=>toggleGroup(code)}
                  style={{padding:"12px 18px",background:"#f8fafc",borderBottom:isCollapsed?"none":`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:10,cursor:"pointer",userSelect:"none"}}>
                  <span style={{fontSize:12,color:T.textMuted,transform:isCollapsed?"rotate(-90deg)":"none",transition:"transform 0.15s",display:"inline-block",width:12}}>▼</span>
                  <span style={{color:T.blue,fontSize:13,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{code}</span>
                  <span style={{color:T.textPrimary,fontSize:13,fontWeight:600}}>{acc?.name || "—"}</span>
                  <span style={{flex:1}}/>
                  <span style={{fontSize:12,color:T.textMuted}}>{t("งบ","Budget")} <b style={{color:T.textSecondary,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(grpBudget)}</b></span>
                  <span style={{fontSize:12,color:grpToOrder<0?T.red:T.textMuted,fontWeight:grpToOrder<0?650:400}}>{grpToOrder<0 ? t("เกินงบ","Over budget") : t("ต้องสั่งเพิ่ม","To order")} <b style={{color:grpToOrder<0?T.red:T.amber,fontVariantNumeric:"tabular-nums"}}>฿{fmt0(Math.abs(grpToOrder))}</b></span>
                  <span style={{color:T.textMuted,fontSize:12}}>{rows.length} PO</span>
                  {lateCount>0 && <Badge text={`⚠️ ${lateCount} ${t("ล่าช้า","late")}`} clr={T.red} bg={T.redBg}/>}
                </div>
                {!isCollapsed && (
                <div className="hscroll"><table style={{width:"100%",minWidth:960,borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr>
                      {[["วันเปิด PO","Open date"],["Supplier","Supplier"],["PO No.","PO No."],["มูลค่า (THB)","Value (THB)"],["วันรับของ","Received"],["วันจ่าย","Pay date"],["การส่งของ","Delivery"],["แผนจ่ายเงิน","Payment plan"],["ติดตาม","Track"],["สถานะ","Status"],["",""]].map(([h,he])=>(
                        <th key={h||"x"} style={{padding:"9px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:12,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{t(h,he)}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(({po:p,item},i) => {
                      // แถวนี้แยกตาม Acc. Code → คิด/แสดงเฉพาะงวดของ item นี้ ไม่เอางวดของ code อื่นในใบเดียวกันมาปน
                      const pItem = { ...p, items:[item] };
                      const inc = incomingStatus(pItem), pay = paymentStatus(pItem);
                      const splitAcrossCodes = poItems(p).length>1;
                      const locked = !canEditPO(p, session);
                      const receivedDates = poReceivedDates(pItem);
                      const paidDate = poPaidDate(pItem);
                      return (
                        <tr key={p.id+"-"+(item.id||item.code)} onClick={()=>onView?.(p)}
                          style={{background:i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`,cursor:onView?"pointer":"default"}}
                          onMouseEnter={e=>e.currentTarget.style.background="#fef9ec"}
                          onMouseLeave={e=>e.currentTarget.style.background=i%2===0?T.card:"#fafbfd"}>
                          <td style={{padding:"9px 16px",color:T.textMuted,fontSize:13,fontVariantNumeric:"tabular-nums"}}>{fmtDate(p.date)}</td>
                          <td style={{padding:"9px 16px",color:T.textPrimary,fontWeight:500}}>{itemSupplierName(p,item)}</td>
                          <td style={{padding:"9px 16px",color:T.textMuted,fontVariantNumeric:"tabular-nums",fontSize:13,whiteSpace:"nowrap"}}>{poNumbersLabel(p)}</td>
                          <td style={{padding:"9px 16px",textAlign:"right"}}>
                            <div style={{color:T.textPrimary,fontVariantNumeric:"tabular-nums",fontWeight:600}}>{fmt(item.amount)}</div>
                            {usdLine(parseFloat(item.amount)||0, usdRate)}
                            {splitAcrossCodes && <div style={{fontSize:12,color:T.textMuted}}>{t("รวม","total")} {fmt(poTotal(p))}</div>}
                          </td>
                          <td style={{padding:"9px 16px",fontSize:13,fontVariantNumeric:"tabular-nums",color:receivedDates.length?T.textPrimary:T.textMuted}}>
                            {receivedDates.length===0 ? "—" : receivedDates.length===1 ? fmtDate(receivedDates[0]) : `${fmtDate(receivedDates[0])} (+${receivedDates.length-1})`}
                          </td>
                          <td style={{padding:"9px 16px",fontSize:13,fontVariantNumeric:"tabular-nums",color:paidDate?T.green:T.textMuted,fontWeight:paidDate?600:450}}>
                            <PayDateText po={pItem}/>
                          </td>
                          <td style={{padding:"9px 16px"}}><DeliveryList po={pItem}/></td>
                          <td style={{padding:"9px 16px"}}>
                            <DateCell value={poNextDueDate(pItem)} lateTint={false}/>
                            {p.paymentType && (
                              <div style={{marginTop:3}}>
                                <Badge text={payTypeLabelT(p)} clr={PAYMENT_TYPE_CLR[p.paymentType]} bg={PAYMENT_TYPE_BG[p.paymentType]}/>
                              </div>
                            )}
                          </td>
                          <td style={{padding:"9px 16px"}}>
                            <div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"flex-start"}}>
                              <Badge text={incLabel(inc)} clr={INCOMING_CLR[inc]} bg={INCOMING_BG[inc]}/>
                              <Badge text={payLabel(pay)} clr={PAYMENT_CLR[pay]} bg={PAYMENT_BG[pay]}/>
                            </div>
                          </td>
                          <td style={{padding:"9px 16px"}} onClick={e=>e.stopPropagation()}>
                            <div style={{display:"flex",alignItems:"center",gap:4}}>
                              <StatusPicker status={p.status} onChange={s=>onStatusChange?.(p,s)} disabled={locked} compact/>
                              {locked && <span title={t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only")} style={{fontSize:12}}><Ico name="lock" size={14} color={T.textMuted} /></span>}
                            </div>
                            {poLastUpdate(p) && <div style={{fontSize:12,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>{t("อัปเดต","Updated")} {relativeTime(poLastUpdate(p).at)}</div>}
                          </td>
                          <td style={{padding:"9px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                            <button onClick={()=>onEdit(p)} disabled={locked} title={locked?t("แก้ไขได้เฉพาะ Admin","Admin only"):t("แก้ไข","Edit")}
                              style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textMuted,cursor:locked?"not-allowed":"pointer",padding:"6px 8px",minWidth:34,minHeight:34,borderRadius:8,display:"inline-grid",placeItems:"center"}} aria-label={t("แก้ไข","Edit")}><Ico name="edit" size={16} /></button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table></div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

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
    ? [t("ของยังไม่เข้า (ตามแผน)","Goods not in yet (plan)"), T.textSecondary, "#f1f5f9"]
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
                      <td style={{padding:"10px 16px",textAlign:"right",fontVariantNumeric:"tabular-nums",color:variance<0?T.red:T.textSecondary,fontWeight:variance<0?650:500}}>
                        {noBudget ? <span data-word="nobudget" title={t("มี PO แต่ QS ยังไม่ได้ลงงบของรหัสนี้ — ทั้งยอดนับเป็นส่วนเกิน","Has POs but no QS budget for this code — the whole amount counts as over")} style={{background:T.redBg,color:T.red,fontSize:11,fontWeight:650,padding:"1px 7px",borderRadius:20,marginRight:8,verticalAlign:"1px"}}>{t("ไม่มีงบ","No budget")}</span>
                          : a.over ? <span data-word="over" style={{background:T.redBg,color:T.red,fontSize:11,fontWeight:650,padding:"1px 7px",borderRadius:20,marginRight:8,verticalAlign:"1px"}}>{t("เกินงบ","Over")}</span>
                          : (a.budget>0 && a.committed/a.budget>=0.8) ? <span data-word="near" style={{background:T.amberBg,color:T.amber,fontSize:11,fontWeight:650,padding:"1px 7px",borderRadius:20,marginRight:8,verticalAlign:"1px"}}>{t(`ใช้ ${Math.round(a.committed/a.budget*100)}%`,`${Math.round(a.committed/a.budget*100)}% used`)}</span> : null}
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
