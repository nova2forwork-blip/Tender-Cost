// Tender Cost — ข้อมูลหลัก + สูตรคำนวณทั้งหมด (งบ, 3%, PO, แผนจ่าย, สถานะ) + ภาษา/สี/CSS · ไม่มีหน้าจอ
import { useEffect, useState } from "react";
import { sg, sgOrThrow, ssOrThrow } from "./supabase.js";
import { ROLE_LABELS } from "./auth.js";

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

// ขั้นของ PO ที่ "ระบบคิดเอง" จากวันของเข้า/วันจ่าย (ไม่ต้องให้คนเปลี่ยนสถานะทีละใบ) — ใช้กับชิปกรองในหน้ารายการ PO
//   waiting = ยังไม่ได้รับของ (รวมที่ล่าช้า) · partial = รับบางส่วน · received = รับครบ รอจ่าย · paid = จ่ายแล้ว
const PO_STAGES = ["waiting","partial","received","paid"];
const poStage = (p) => {
  if (paymentStatus(p) === "paid") return "paid";
  const inc = incomingStatus(p);
  return inc === "received" ? "received" : inc === "partial" ? "partial" : "waiting";
};
const poStageLabel = (k) => ({
  All: t("ทั้งหมด","All"), waiting: t("รอของเข้า","Awaiting goods"), partial: t("รับบางส่วน","Partly received"),
  received: t("รับครบ · รอจ่าย","Received · to pay"), paid: t("จ่ายแล้ว","Paid"),
}[k] || k);

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

export { ACCOUNTS, GROUPS, PO_STATUS, PO_STATUS_TH, poStatusLabel, STATUS_CLR, STATUS_BG, GRP_COLORS, applyAccountList, _EXTRA_ITEMS, setExtraRegistry, accountOf, migrateAccountCodes, todayStr, UnsavedGuard, DialogStore, uiAlert, uiConfirm, uiPrompt, leaveIfDirty, addDays, DEFAULT_CREDIT_DAYS, isNewPO, migratePO, poItems, poTotal, poAmountForCode, poSupplier, poSupplierName, poSupplierText, poSupplierLabel, poNumbersLabel, itemSupplierName, poSuppliers, poRounds, poDeliveries, roundPayDate, roundReceived, roundPaid, itemOrdered, itemReceived, itemEntered, itemRemaining, HISTORY_ICON, historyEntry, poHistory, poLastUpdate, withHistory, relativeTime, uiLocale, formatDateTime, poReceivedDates, poPaidDate, roundDueForecast, poNextDueDate, fmtDate, PayDateText, poPayLines, poPaidAmount, itemPaidAmount, poBilledAmount, PO_STAGES, poStage, poStageLabel, isPOLocked, canEditPO, deliveryStatus, incomingStatus, paymentStatus, INCOMING_LABEL, INCOMING_CLR, INCOMING_BG, PAYMENT_LABEL, PAYMENT_CLR, PAYMENT_BG, PAYMENT_TYPE_CLR, PAYMENT_TYPE_BG, creditTermDays, INCOMING_LABEL_EN, PAYMENT_LABEL_EN, incLabel, payLabel, payTypeLabelT, fmt, fmtZ, fmt0, fmtK, monthShortLabel, uid, T, _LANG, _langSubs, t, inThai, setLang, toggleLang, useLang, FAB_SIZE, FAB_GAP, BNAV_H, bnavH, BOTTOM, GLOBAL_CSS, monthAddValue, WASTE_RATE, WASTE_LBL, wasteOf, withWaste, rollupTenders, rollupAdditions, buildCombinedBudget, monthRowBreakdown, OTHER_COL_LABEL, exportAccountList, ORPHAN_NAME, poCodeSet, hiddenSafeForPO, projectSummary };
