import { useState, useEffect, useCallback, useRef, Fragment, Component } from "react";
import * as XLSX from "xlsx-js-style";
import { supabase, sg, ss, ssMerge, sd, loadKvHistory, restoreKvVersion, loadKvSnapshots, restoreKvSnapshot } from "./supabase.js";
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
const PO_STATUS   = ["Pending","PO Issued","Delivered","Invoiced","Paid"];
const STATUS_CLR  = { Pending:"#94a3b8","PO Issued":"#3b82f6",Delivered:"#f59e0b",Invoiced:"#8b5cf6",Paid:"#10b981" };
const STATUS_BG   = { Pending:"#f1f5f9","PO Issued":"#eff6ff",Delivered:"#fffbeb",Invoiced:"#f5f3ff",Paid:"#f0fdf4" };
const GRP_COLORS  = ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6","#06b6d4","#f97316","#ec4899"];

// ─── Incoming / Payment tracking status ────────────────────────────────────
// A PO's incoming status is derived from its planned/actual dates rather than
// stored directly, so it's always in sync with today's date.
// วันนี้ตาม "ปฏิทินท้องถิ่น" (ไม่ใช้ UTC เพื่อไม่ให้ข้ามวันตอนเช้ามืดในโซน UTC+7)
const todayStr = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
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
const itemSupplier     = (p) => poSupplier(p);
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
const poRoundsAmount = (p) => poRounds(p).reduce((s,r)=>s+(parseFloat(r.planAmount)||0),0);

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
const roundPaid     = (p, r) => { const d = roundPayDate(p,r); return !!d && d <= todayStr(); };
const itemOrdered   = (it) => parseFloat(it.amount)||0;
const itemReceived  = (it) => (it.rounds||[]).filter(roundReceived).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
const itemEntered   = (it) => (it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
const itemRemaining = (it) => Math.max(itemOrdered(it) - itemEntered(it), 0);

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
  if (min < 1)  return "เมื่อสักครู่";
  if (min < 60) return `${min} นาทีที่แล้ว`;
  const hr = Math.floor(min/60);
  if (hr < 24)  return `${hr} ชม.ที่แล้ว`;
  const day = Math.floor(hr/24);
  if (day < 30) return `${day} วันที่แล้ว`;
  return new Date(iso).toLocaleDateString("th-TH",{day:"numeric",month:"short",year:"2-digit"});
};
const formatDateTime = (iso) => iso ? new Date(iso).toLocaleString("th-TH",{day:"numeric",month:"short",year:"2-digit",hour:"2-digit",minute:"2-digit"}) : "—";

// ─── Actual received / paid dates ──────────────────────────────────────────
// Every round already carries its own "received" date; a PO's overall
// received date(s) are just the distinct actual dates across every round.
const poReceivedDates = (p) => [...new Set(poRounds(p).filter(roundReceived).map(r=>r.actualDate).filter(Boolean))].sort();
// Paid date = the latest due date among rounds that have auto-paid.
const poPaidDate = (p) => {
  const paid = poRounds(p).filter(r=>roundPaid(p,r)).map(r=>roundPayDate(p,r)).filter(Boolean).sort();
  return paid.length ? paid[paid.length-1] : null;
};
// Earliest upcoming/known payment due date across all rounds (for list/export).
const poNextDueDate = (p) => {
  const due = poRounds(p).map(r=>roundPayDate(p,r)).filter(Boolean).sort();
  return due.length ? due[0] : "";
};

// ─── แผนจ่ายเงิน (Payment forecast lines) ───────────────────────────────────
// คืน "งวดจ่าย" ของ PO หนึ่งใบ สำหรับหน้าแผนจ่าย/Export. ปกติแตกตามงวดส่งของ
// (แต่ละงวดมียอด planAmount/actualAmount ของตัวเอง ซึ่งควรรวมกัน = ยอด item)
// แต่ถ้ายอดรวมของงวดไม่ตรงกับยอด item (เช่นมีงวดซ้ำยอดเต็ม) จะยุบเหลือ "หนึ่ง
// บรรทัดต่อ item" โดยยึดยอด item.amount เป็นหลัก เพื่อกันการนับซ้ำ. วันครบกำหนด
// จ่าย = วันรับของ (จริงถ้ามี ไม่มีใช้วันแผน) + เทอมเครดิต; เงินสดจ่ายวันรับของ.
const poPayLines = (p) => {
  const P = migratePO(p);
  const isCash = P.paymentType === "cash";
  const term = isCash ? 0 : (parseInt(P.creditDays,10) || DEFAULT_CREDIT_DAYS);
  const method = isCash ? "เงินสด" : `เครดิต ${term} วัน`;
  const dueOf = (incoming) => incoming ? (isCash ? incoming : addDays(incoming, term)) : "";
  const roundAmt = (r) => (parseFloat(r.actualAmount)||0) || (parseFloat(r.planAmount)||0);
  const out = [];
  poItems(P).forEach(it => {
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
        const paid = received && roundPaid(P, r);
        out.push({ code: it.code||"", incoming, incomingType: r.actualDate?"จริง":(r.planDate?"แผน":""),
          payDate: dueOf(incoming), amount, received, paid, paidAmount: paid ? amount : 0 });
      });
    } else if (itemAmt > 0) {
      // ยอดงวดไม่ตรง (หรือไม่มีงวด) → ยุบเหลือบรรทัดเดียว ใช้ยอด item เป็นหลัก
      // จ่ายบางส่วน: เก็บ paidAmount ไว้ให้ยอด "คงเหลือต้องจ่าย" หักออกถูกต้อง
      const actualDates = rounds.map(r=>r.actualDate).filter(Boolean).sort();
      const planDates   = rounds.map(r=>r.planDate).filter(Boolean).sort();
      const incoming = actualDates[0] || planDates[0] || "";
      const paidAmt  = Math.min(rounds.filter(r=>roundReceived(r) && roundPaid(P,r)).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0), itemAmt);
      out.push({ code: it.code||"", incoming, incomingType: actualDates.length?"จริง":(planDates.length?"แผน":""),
        payDate: dueOf(incoming), amount: itemAmt, received: rounds.some(roundReceived), paid: paidAmt >= itemAmt-0.5, paidAmount: paidAmt });
    }
  });
  const today = todayStr();
  return out.map(l => ({
    ...l, isCash, method,
    supplier: poSupplierName(P), poNo: poNumbersLabel(P),
    accName: (ACCOUNTS.find(a=>a.code===l.code)?.name)||"",
    month: l.payDate ? l.payDate.slice(0,7) : "",
    // "เกินกำหนดจ่าย" เฉพาะของที่รับแล้วแต่ยังไม่จ่ายและเลยกำหนด; ของที่ยังไม่รับ = "รอจ่าย"
    status: l.paid ? "paid" : (l.received && l.payDate && l.payDate < today ? "late" : "pending"),
  }));
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
  const allReceived = items.length > 0 && items.every(it => itemReceived(it) >= itemOrdered(it) - 0.001);
  const anyReceived = rounds.some(roundReceived);
  const anyLate  = rounds.some(r => !roundReceived(r) && r.planDate && r.planDate < todayStr());
  if (allReceived) return "received";
  if (anyLate) return "late";
  if (anyReceived) return "partial";
  if (rounds.some(r=>r.planDate)) return "pending";
  return "unset";
};
// Auto-pay: reaching a round's due date is what marks it paid, so payment is
// never "late" — it's "pending" until the due date, then "paid".
const paymentStatus = (p) => {
  const rounds = poRounds(p);
  const recvRounds = rounds.filter(roundReceived);
  if (!recvRounds.length) return "unset";
  // จ่ายครบต่อ item: งวดที่ทั้งรับแล้วและถึงกำหนดจ่าย ต้องครอบคลุมยอด item ทุก item
  // (กันไม่ให้ "จ่ายแล้ว" เกิดขึ้นทั้งที่บาง item ยังจ่ายไม่ครบ)
  const items = poItems(p).filter(it => itemOrdered(it) > 0);
  const itemPaid = (it) => (it.rounds||[]).filter(r => roundReceived(r) && roundPaid(p,r)).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
  const allPaid = items.length > 0 && items.every(it => itemPaid(it) >= itemOrdered(it) - 0.001);
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
const PAYMENT_TYPE_LABEL = { cash:"เงินสด", credit:"เครดิต", credit30:"เครดิต 30 วัน" };
const PAYMENT_TYPE_ICON  = { cash:"💵", credit:"💳", credit30:"💳" };
const PAYMENT_TYPE_CLR   = { cash:"#10b981", credit:"#2563eb", credit30:"#2563eb" };
const PAYMENT_TYPE_BG    = { cash:"#f0fdf4", credit:"#eff6ff", credit30:"#eff6ff" };
// Label for a PO's payment method including its credit term, e.g. "เครดิต 45 วัน".
const paymentTypeLabel = (p) => { const P = migratePO(p); if (P.paymentType==="cash") return "เงินสด"; if (P.paymentType==="credit") return `เครดิต ${P.creditDays||DEFAULT_CREDIT_DAYS} วัน`; return "—"; };
const fmt  = n => new Intl.NumberFormat("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2}).format(n||0);
// บาทเต็ม (ไม่มีทศนิยม) — ใช้กับตัวเลขพาดหัวการ์ด/ยอดรวม ให้กวาดตาอ่านง่าย
const fmt0 = n => new Intl.NumberFormat("th-TH",{maximumFractionDigits:0}).format(Math.round(n||0));
const fmtK = n => n>=1e6?`${(n/1e6).toFixed(1)}M`:n>=1e3?`${(n/1e3).toFixed(0)}K`:Math.round(n).toString();
// "2026-08" -> "ส.ค. 69" — used wherever a month key needs a short Thai label
// (QS Monthly tab's chips/headers, and sub-item "เพิ่มเมื่อ ..." badges).
const monthShortLabel = (m) => new Date(m+"-01").toLocaleDateString("th-TH",{month:"short",year:"2-digit"});
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
  textSecondary:"#64748b",
  textMuted:    "#94a3b8",
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

// ─── Global CSS ───────────────────────────────────────────────────────────────
const GLOBAL_CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Inter', sans-serif; background: ${T.bg}; color: ${T.textPrimary}; }
  input, select, textarea, button { font-family: 'Inter', sans-serif; }
  input[type=number]::-webkit-inner-spin-button { opacity: 0.4; }
  ::-webkit-scrollbar { width: 5px; height: 5px; }
  ::-webkit-scrollbar-track { background: #f1f5f9; }
  ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 99px; }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
  @keyframes fadeIn { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:translateY(0)} }
  .card-hover { transition: box-shadow 0.18s, transform 0.18s; }
  .card-hover:hover { box-shadow: 0 8px 24px rgba(37,99,235,0.12); transform: translateY(-2px); }
  .btn-primary { background: ${T.blue}; color: #fff; border: none; border-radius: 10px; padding: 10px 22px; font-size: 13px; font-weight: 600; cursor: pointer; transition: background 0.15s, box-shadow 0.15s; }
  .btn-primary:hover { background: ${T.blueDark}; box-shadow: 0 4px 12px rgba(37,99,235,0.3); }
  .btn-ghost { background: transparent; color: ${T.textSecondary}; border: 1.5px solid ${T.cardBorder}; border-radius: 10px; padding: 9px 18px; font-size: 13px; font-weight: 500; cursor: pointer; transition: border-color 0.15s, color 0.15s; }
  .btn-ghost:hover { border-color: ${T.blue}; color: ${T.blue}; }
  .input-base { background: ${T.bg}; border: 1.5px solid ${T.cardBorder}; border-radius: 10px; padding: 10px 13px; color: ${T.textPrimary}; font-size: 13px; outline: none; transition: border-color 0.15s, box-shadow 0.15s; width: 100%; }
  .input-base:focus { border-color: ${T.blue}; box-shadow: 0 0 0 3px rgba(37,99,235,0.1); }
  .tag { display: inline-flex; align-items: center; padding: 2px 9px; border-radius: 6px; font-size: 11px; font-weight: 600; }
  /* กล่องเลื่อนแนวนอน (ใช้กับตารางที่คอลัมน์เยอะ) — สกรอลบาร์เห็นชัดเสมอ */
  .hscroll { overflow-x: auto; overflow-y: hidden; }
  .hscroll::-webkit-scrollbar { height: 12px; }
  .hscroll::-webkit-scrollbar-track { background: #eef2f7; border-radius: 8px; }
  .hscroll::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 8px; border: 3px solid #eef2f7; }
  .hscroll::-webkit-scrollbar-thumb:hover { background: #64748b; }
  .hscroll { scrollbar-color: #94a3b8 #eef2f7; scrollbar-width: thin; }
  /* ตารางรายเดือน: เลื่อนในกล่องเอง (สูงไม่เกิน 70vh) + ตรึงหัวตาราง + สกรอลบาร์เห็นชัด */
  .mscroll { overflow: auto; max-height: 70vh; }
  .mscroll::-webkit-scrollbar { height: 13px; width: 13px; }
  .mscroll::-webkit-scrollbar-track { background: #eef2f7; }
  .mscroll::-webkit-scrollbar-thumb { background: #94a3b8; border-radius: 8px; border: 3px solid #eef2f7; }
  .mscroll::-webkit-scrollbar-thumb:hover { background: #64748b; }
  .mscroll::-webkit-scrollbar-corner { background: #eef2f7; }
  .mscroll { scrollbar-color: #94a3b8 #eef2f7; scrollbar-width: thin; }
  .mscroll thead th { position: sticky; background: #f8fafc; z-index: 2; box-shadow: inset 0 -1px 0 ${T.cardBorder}; }
  .mscroll thead tr:first-child th { top: 0; }
  .mscroll thead tr:nth-child(2) th { top: 33px; z-index: 2; }
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
const buildCombinedBudget = (tenderCosts, additions) => {
  const combined = {...tenderCosts};
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
const exportAccountList = (extraItems=[], hiddenAccounts=[]) => [
  ...ACCOUNTS.filter(a => !hiddenAccounts.includes(a.code)),
  ...extraItems.filter(e => !e.parentCode).map(e => ({ code:e.code, name:e.name, group:e.group||"Other" })),
];

// ─── Styling helper ─────────────────────────────────────────────────────────
// Lays down a colored title bar (merged across every column), optional gray
// info sub-rows, a bold colored header row with autofilter, zebra-striped
// bordered data rows with right-aligned money/% columns, and an optional
// bold total row — everything an aoa_to_sheet grid needs to read like a
// real report instead of a raw data dump.
const BORDER_THIN = (rgb) => ({ style:"thin", color:{rgb} });
const BORDER_MED  = (rgb) => ({ style:"medium", color:{rgb} });
// สีพิลล์ตามสถานะ (เขียว=เสร็จ/จ่ายแล้ว, เหลือง=กำลังทำ/รอ, แดง=ค้าง/เกินกำหนด)
// ใช้คีย์เวิร์ดจับ ครอบคลุมทั้งไทย/อังกฤษ สถานะอื่นเป็นพิลล์เทากลาง ๆ
const STATUS_PILL = [
  [/(completed|complete|เสร็จ|จ่ายแล้ว|รับของแล้ว|รับครบ|ปิดงาน|ปิด|อนุมัติ|approved|done|paid)/i, { bg:"D1FAE5", fg:"065F46" }],
  [/(in\s*progress|progress|กำลัง|ระหว่าง|บางส่วน|partial|สั่งซื้อ|สั่ง|รอรับ|รอจ่าย|pending|รอ)/i,        { bg:"FEF3C7", fg:"92400E" }],
  [/(to\s*do|todo|ร่าง|ยังไม่|ค้างจ่าย|ค้าง|เกินกำหนด|overdue|ยกเลิก|cancel|reject)/i,               { bg:"FEE2E2", fg:"991B1B" }],
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
                           rowGroups = null, groupDisplayCol = null }) {
  ws["!rows"]   = ws["!rows"] || [];
  ws["!merges"] = ws["!merges"] || [];
  // โทนพาสเทลนุ่ม ๆ ที่ได้จากสีธีม
  const HFILL = lighten(theme.main, 0.82); // หัวตาราง พื้นอ่อน
  const BAND  = lighten(theme.main, 0.95); // แถบสลับสีจาง ๆ
  const TFILL = lighten(theme.main, 0.75); // แถวรวม
  const GLINE = lighten(theme.main, 0.55); // เส้นแบ่งกลุ่ม

  ws["!merges"].push({ s:{r:titleRow,c:0}, e:{r:titleRow,c:numCols-1} });
  for (let c=0; c<numCols; c++) {
    const ref = XLSX.utils.encode_cell({r:titleRow,c});
    if (!ws[ref]) ws[ref] = { t:"s", v:"" };
    ws[ref].s = { font:{bold:true,sz:13,color:{rgb:theme.dark},name:"Tahoma"},
      fill:{fgColor:{rgb:lighten(theme.main,0.55)}}, alignment:{vertical:"center",horizontal:"left",indent:1} };
  }
  ws["!rows"][titleRow] = { hpx:30 };

  subRows.forEach(r => {
    for (let c=0; c<numCols; c++) {
      const ref = XLSX.utils.encode_cell({r,c});
      if (ws[ref]) ws[ref].s = { font:{italic:true,sz:9.5,color:{rgb:"64748B"},name:"Tahoma"} };
    }
    ws["!rows"][r] = { hpx:16 };
  });

  for (let c=0; c<numCols; c++) {
    const ref = XLSX.utils.encode_cell({r:headerRow,c});
    if (!ws[ref]) ws[ref] = { t:"s", v:"" };
    ws[ref].s = { font:{bold:true,sz:10.5,color:{rgb:theme.dark},name:"Tahoma"},
      fill:{fgColor:{rgb:HFILL}}, alignment:{vertical:"center",horizontal:"center",wrapText:true},
      border:{ top:BORDER_THIN(HFILL), bottom:BORDER_MED(theme.main), left:BORDER_THIN("FFFFFF"), right:BORDER_THIN("FFFFFF") } };
  }
  ws["!rows"][headerRow] = { hpx:28 };
  ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s:{r:headerRow,c:0}, e:{r:headerRow,c:numCols-1} }) };
  // ตรึงทุกอย่างเหนือแถวข้อมูล (หัวข้อ+หัวตาราง) ให้ค้างไว้ตอนเลื่อน
  ws["!freeze"] = { xSplit:0, ySplit:headerRow+1, topLeftCell: XLSX.utils.encode_cell({ r:headerRow+1, c:0 }), activePane:"bottomLeft", state:"frozen" };

  let band = 0, prevGroup;
  for (let r=dataStart; r<=dataEnd; r++) {
    const idx = r - dataStart;
    let zebra, isGroupStart = false;
    if (rowGroups) {
      const g = rowGroups[idx];
      isGroupStart = idx > 0 && g !== prevGroup;
      if (idx === 0 || isGroupStart) band = 1 - band;
      prevGroup = g;
      zebra = band === 1;
    } else {
      zebra = idx % 2 === 1;
    }
    for (let c=0; c<numCols; c++) {
      const ref = XLSX.utils.encode_cell({r,c});
      if (!ws[ref]) continue;
      const isMoney = moneyCols.includes(c), isPct = pctCols.includes(c), isCenter = centerCols.includes(c);
      const isGroupLabel = groupDisplayCol != null && c === groupDisplayCol;
      const s = { font: isGroupLabel
          ? {sz:10,name:"Tahoma",bold:true,color:{rgb:theme.dark}}
          : {sz:10,name:"Tahoma",color:{rgb:"1F2937"}},
        alignment:{ vertical:"center", horizontal:isMoney||isPct?"right":isCenter?"center":"left", wrapText:true },
        border:{ top: isGroupStart?BORDER_MED(GLINE):BORDER_THIN("EEF0F2"), bottom:BORDER_THIN("EEF0F2"),
                 left:BORDER_THIN("EEF0F2"), right:BORDER_THIN("EEF0F2") } };
      if (zebra)   s.fill   = { fgColor:{rgb:BAND} };
      if (isMoney) s.numFmt = usdCols.includes(c) ? '"$"#,##0.00' : '"฿"#,##0';   // แยกสัญลักษณ์ $ / ฿
      if (isPct)   s.numFmt = "0.0%";
      if (statusCols.includes(c)) {
        const pill = statusPill(ws[ref].v);
        if (pill) {
          s.fill = { fgColor:{rgb:pill.bg} };
          s.font = { sz:10, name:"Tahoma", bold:true, color:{rgb:pill.fg} };
          s.alignment = { ...s.alignment, horizontal:"center" };
        }
      }
      ws[ref].s = s;
    }
    ws["!rows"][r] = ws["!rows"][r] || { hpx:19 };
  }

  if (totalRow != null) {
    for (let c=0; c<numCols; c++) {
      const ref = XLSX.utils.encode_cell({r:totalRow,c});
      if (!ws[ref]) ws[ref] = { t:"s", v:"" };
      const isMoney = moneyCols.includes(c), isPct = pctCols.includes(c);
      ws[ref].s = { font:{bold:true,sz:10.5,color:{rgb:theme.dark},name:"Tahoma"}, fill:{fgColor:{rgb:TFILL}},
        alignment:{vertical:"center",horizontal:isMoney||isPct?"right":"left"},
        border:{ top:BORDER_MED(theme.main) },
        numFmt: isMoney ? (usdCols.includes(c) ? '"$"#,##0.00' : '"฿"#,##0') : isPct?"0.0%":undefined };
    }
    ws["!rows"][totalRow] = { hpx:24 };
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

// จัดความกว้างคอลัมน์ให้พอดีข้อความ สำหรับตาราง ExcelJS (จากหัว + แถวข้อมูล)
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

function fitExcelCols(ws, header, dataRows, { min=8, max=55 } = {}) {
  header.forEach((h, c) => {
    let m = String(h == null ? "" : h).length;
    dataRows.forEach(row => {
      const v = row[c];
      const s = (typeof v === "number") ? Math.round(v).toLocaleString("en-US") : String(v == null ? "" : v);
      if (s.length > m) m = s.length;
    });
    ws.getColumn(c+1).width = Math.min(max, Math.max(min, m + 2));
  });
}

// กราฟแท่งแนวตั้งที่ "วาดด้วยเซลล์" — ไลบรารีนี้ฝังกราฟจริง/รูปไม่ได้ จึงระบายสี
// เซลล์ไล่จากล่างขึ้นบนตามค่าให้ออกมาเป็นกราฟแท่งในชีต Excel
function addBarChartSheet(wb, sheetName, title, theme, items) {
  items = (items || []).filter(Boolean);
  if (!items.length) return;
  const H = 12;
  const max = Math.max(...items.map(i => i.value || 0), 1);
  const n = items.length;
  const LEFT = 1;                     // เว้นคอลัมน์แรกเป็นแกน
  const totalCols = LEFT + n;
  const valueRow = 2, chartTop = 3, labelRow = chartTop + H;
  const aoa = [[title], []];
  for (let r = 0; r < H + 2; r++) aoa.push(new Array(totalCols).fill(""));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!merges"] = [{ s:{r:0,c:0}, e:{r:0,c:totalCols-1} }];
  ws["!cols"] = [{ wch:4 }, ...items.map(()=>({ wch:11 }))];
  ws["!rows"] = [];
  ws["!rows"][0] = { hpx:26 };
  for (let r = chartTop; r < chartTop + H; r++) ws["!rows"][r] = { hpx:15 };
  ws["!rows"][labelRow] = { hpx:24 };
  const setS = (r, c, s, v) => {
    const ref = XLSX.utils.encode_cell({ r, c });
    if (v != null) ws[ref] = { t: typeof v === "number" ? "n" : "s", v };
    else if (!ws[ref]) ws[ref] = { t:"s", v:"" };
    ws[ref].s = s;
  };
  setS(0, 0, { font:{bold:true,sz:13,color:{rgb:"FFFFFF"},name:"Tahoma"}, fill:{fgColor:{rgb:theme.main}}, alignment:{vertical:"center",horizontal:"left",indent:1} });
  const barOn = theme.main, barOff = "F3F4F6";
  items.forEach((it, i) => {
    const c = LEFT + i;
    const filled = Math.max(0, Math.round(((it.value||0) / max) * H));
    setS(valueRow, c, { font:{bold:true,sz:9,color:{rgb:theme.dark},name:"Tahoma"}, alignment:{horizontal:"center"}, numFmt:"#,##0" }, it.value||0);
    for (let k = 0; k < H; k++) {
      const r = chartTop + (H - 1 - k); // k=0 = ล่างสุด
      setS(r, c, { fill:{fgColor:{rgb: k < filled ? barOn : barOff }} });
    }
    setS(labelRow, c, { font:{bold:true,sz:9,color:{rgb:"374151"},name:"Tahoma"}, alignment:{horizontal:"center",wrapText:true} }, it.label);
  });
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
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
  const ACC = [["DBEAFE","1D4ED8"],["D1FAE5","047857"],["FEF3C7","92400E"],["EDE9FE","6D28D9"]];
  cards.slice(0,4).forEach((cd, i) => {
    const c0 = i*2, c1 = c0+1, [bg,fg] = ACC[i%4];
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
    const gh = ["กลุ่ม","ราคาเดิม","เพิ่มรายเดือน","งบรวม","สัดส่วน","กราฟสัดส่วน"];
    gh.forEach((h,c) => setS(gHeadRow, c, { font:{bold:true,sz:9.5,color:{rgb:theme.dark},name:"Tahoma"}, fill:{fgColor:{rgb:headFill}}, alignment:{horizontal:c===0?"left":c<5?"right":"left",vertical:"center",indent:c===0||c===5?1:0}, border:{bottom:BORDER_MED(theme.main)} }, h));
    for (let c=6;c<C;c++) setS(gHeadRow, c, { fill:{fgColor:{rgb:headFill}}, border:{bottom:BORDER_MED(theme.main)} });
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
    const rT = gTotalRow, tb = { fill:{fgColor:{rgb:lighten(theme.main,0.75)}}, border:{top:BORDER_MED(theme.main)} };
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

// ── กราฟจริง (pie + bar) ในไฟล์ Excel: โหลด ExcelJS จาก CDN ตอนใช้งาน แล้ววาด
//    กราฟเป็นรูป PNG ฝังลงไฟล์ — ไม่ต้องเพิ่ม dependency / ไม่กระทบ build ──────────
function loadExcelJS() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  return new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js";
    s.onload = () => res(window.ExcelJS);
    s.onerror = () => rej(new Error("โหลด ExcelJS ไม่ได้"));
    document.head.appendChild(s);
  });
}

function chartPiePNG(items) {
  const W=520,H=300,dpr=2,cv=document.createElement("canvas"); cv.width=W*dpr; cv.height=H*dpr;
  const x=cv.getContext("2d"); x.scale(dpr,dpr); x.fillStyle="#fff"; x.fillRect(0,0,W,H);
  const cx=150,cy=155,r=115,tot=items.reduce((s,i)=>s+i.value,0)||1; let a=-Math.PI/2;
  items.forEach(it=>{ const f=it.value/tot,a2=a+f*Math.PI*2; x.beginPath(); x.moveTo(cx,cy); x.arc(cx,cy,r,a,a2); x.closePath(); x.fillStyle="#"+it.color; x.fill();
    if(f>0.04){ const m=(a+a2)/2; x.fillStyle="#fff"; x.font="bold 12px Tahoma"; x.textAlign="center"; x.fillText((f*100).toFixed(1)+"%",cx+Math.cos(m)*r*0.62,cy+Math.sin(m)*r*0.62+4);} a=a2; });
  x.beginPath(); x.arc(cx,cy,r*0.52,0,Math.PI*2); x.fillStyle="#fff"; x.fill();
  let ly=30; x.textAlign="left"; items.forEach(it=>{ x.fillStyle="#"+it.color; x.fillRect(300,ly,12,12); x.fillStyle="#334155"; x.font="12px Tahoma"; x.fillText(`${it.label} (${(it.value/tot*100).toFixed(1)}%)`,318,ly+11); ly+=23; });
  return cv.toDataURL("image/png").split(",")[1];
}
function chartBarPNG(items, color) {
  const W=520,H=300,dpr=2,pad=44,cv=document.createElement("canvas"); cv.width=W*dpr; cv.height=H*dpr;
  const x=cv.getContext("2d"); x.scale(dpr,dpr); x.fillStyle="#fff"; x.fillRect(0,0,W,H);
  const max=Math.max(...items.map(i=>i.value),1),n=items.length||1,pw=W-pad*2,ph=H-pad*2;
  x.strokeStyle="#e5e7eb"; x.beginPath(); x.moveTo(pad,H-pad); x.lineTo(W-pad,H-pad); x.stroke();
  items.forEach((it,i)=>{ const bw=pw/n*0.6,bh=(it.value/max)*ph,bx=pad+(pw/n)*i+(pw/n-bw)/2,by=H-pad-bh;
    x.fillStyle="#"+color; x.fillRect(bx,by,bw,bh);
    x.fillStyle="#6b7280"; x.font="9px Tahoma"; x.textAlign="center"; x.save(); x.translate(bx+bw/2,H-pad+4); x.rotate(-Math.PI/4); x.fillText(it.label,0,4); x.restore();
    if(it.value>0){ x.fillStyle="#334155"; x.font="bold 9px Tahoma"; x.textAlign="center"; x.fillText(fmtK(it.value),bx+bw/2,by-4);} });
  return cv.toDataURL("image/png").split(",")[1];
}
async function exportQSRich(project, tenderCosts, additions, extraItems=[], hiddenAccounts=[]) {
  const ExcelJS = await loadExcelJS();
  const F = "Tahoma";
  const combined = buildCombinedBudget(tenderCosts, additions);
  const accounts = exportAccountList(extraItems, hiddenAccounts);
  const list = accounts.filter(a => { const bs=parseFloat(tenderCosts[a.code])||0, tt=parseFloat(combined[a.code])||0; return !(tt<=0 && bs<=0); });
  const months = [...new Set(Object.keys(additions||{}).filter(k=>!k.startsWith("$")))].sort();
  const M = months.length;
  const base = list.reduce((s,a)=> s+(parseFloat(tenderCosts[a.code])||0),0);
  const added = list.reduce((s,a)=> s + months.reduce((ss,m)=> ss+monthAddValue(additions, m, a.code),0), 0);
  const HD = ["Acc. Code","Account Name","Group","ราคาเดิม", ...months.map(monthShortLabel), "รวมทั้งหมด"];
  const NC = HD.length;
  const soft = "FF"+lighten("2563EB",0.55), colL = c => XLSX.utils.encode_col(c);
  const wb = new ExcelJS.Workbook();
  const fillS = (a) => ({ type:"pattern", pattern:"solid", fgColor:{argb:a} });
  const ws = wb.addWorksheet("รายงานงบประมาณ", { views:[{ showGridLines:false, state:"frozen", ySplit:7 }] });
  ws.mergeCells(1,1,1,NC); const t=ws.getCell(1,1); t.value=`สรุปงบประมาณ — ${project.name}`; t.font={bold:true,size:15,color:{argb:"FF1D4ED8"},name:F}; t.fill=fillS(soft); t.alignment={vertical:"middle",indent:1}; ws.getRow(1).height=30;
  ws.mergeCells(2,1,2,NC); const stc=ws.getCell(2,1); stc.value=`พื้นที่ ${project.area||"-"} ft² · แผง ${project.panels||"-"} · Export: ${new Date().toLocaleDateString("th-TH")}`; stc.font={italic:true,size:10,color:{argb:"FF64748B"},name:F}; stc.alignment={indent:1};
  const cards=[["ราคาเดิม (Baseline)","FFDBEAFE","FF1D4ED8",base],["เพิ่มรายเดือนรวม","FFD1FAE5","FF047857",added],["งบรวมทั้งหมด","FFFEF3C7","FF92400E",base+added],["จำนวนเดือน","FFEDE9FE","FF6D28D9",M]];
  const span = Math.max(2, Math.floor(NC/4));
  cards.forEach((cd,i)=>{ const c0=1+i*span, c1=Math.min(NC, c0+span-1);
    ws.mergeCells(4,c0,4,c1); ws.mergeCells(5,c0,5,c1);
    const lc=ws.getCell(4,c0); lc.value=cd[0]; lc.font={bold:true,size:10,color:{argb:cd[2]},name:F}; lc.fill=fillS(cd[1]); lc.alignment={horizontal:"center",vertical:"middle"};
    const vc=ws.getCell(5,c0); vc.value=cd[3]; if(i<3) vc.numFmt="#,##0"; vc.font={bold:true,size:15,color:{argb:cd[2]},name:F}; vc.fill=fillS(cd[1]); vc.alignment={horizontal:"center",vertical:"middle"};
  }); ws.getRow(4).height=18; ws.getRow(5).height=30;
  const HR = 7;
  HD.forEach((h,i)=>{ const c=ws.getCell(HR,1+i); c.value=h; c.font={bold:true,size:9.5,color:{argb:"FF1D4ED8"},name:F}; c.fill=fillS("FFDCE6FB"); c.alignment={horizontal:i>2?"right":"left",vertical:"middle",wrapText:true}; c.border={bottom:{style:"medium",color:{argb:"FF2563EB"}}}; }); ws.getRow(HR).height=24;
  const drows = [];
  list.forEach((a,ri)=>{ const R=HR+1+ri;
    const mv = months.map(m=>monthAddValue(additions, m, a.code)), bs=parseFloat(tenderCosts[a.code])||0;
    drows.push([a.code, a.name, a.group, bs, ...mv, bs+mv.reduce((s,v)=>s+v,0)]);
    ws.getCell(R,1).value=a.code; ws.getCell(R,2).value=a.name; ws.getCell(R,3).value=a.group;
    const bc=ws.getCell(R,4); bc.value=bs; bc.numFmt="#,##0"; bc.alignment={horizontal:"right",vertical:"middle"}; bc.font={name:F,size:9.5};
    mv.forEach((v,mi)=>{ const c=ws.getCell(R,5+mi); c.value=v; c.numFmt="#,##0"; c.alignment={horizontal:"right",vertical:"middle"}; c.font={name:F,size:9.5}; });
    const lastM=colL(3+M); const tc=ws.getCell(R,NC); tc.value = M ? { formula:`D${R}+SUM(E${R}:${lastM}${R})`, result: bs+mv.reduce((s,v)=>s+v,0) } : { formula:`D${R}`, result: bs }; tc.numFmt="#,##0"; tc.font={bold:true,name:F,size:9.5}; tc.alignment={horizontal:"right",vertical:"middle"};
    [1,2,3].forEach(c=>{ ws.getCell(R,c).font={name:F,size:9.5}; ws.getCell(R,c).alignment={vertical:"middle"}; });
    if(ri%2) for(let c=1;c<=NC;c++){ const cell=ws.getCell(R,c); if(!cell.fill||!cell.fill.pattern) cell.fill=fillS("FFF4F7FE"); }
  });
  const tR = HR + 1 + list.length;
  for(let c=1;c<=NC;c++){ const cell=ws.getCell(tR,c); cell.fill=fillS("FFC9D8FA"); cell.border={top:{style:"medium",color:{argb:"FF2563EB"}}}; }
  const tl=ws.getCell(tR,2); tl.value="TOTAL"; tl.font={bold:true,color:{argb:"FF1D4ED8"},name:F};
  [4, ...months.map((_,i)=>5+i), NC].forEach(col=>{ const c=ws.getCell(tR,col), L=colL(col-1); c.value = list.length ? { formula:`SUM(${L}${HR+1}:${L}${HR+list.length})` } : 0; c.numFmt="#,##0"; c.font={bold:true,color:{argb:"FF1D4ED8"},name:F}; c.alignment={horizontal:"right",vertical:"middle"}; });
  ws.getCell(5,1).value = { formula:`D${tR}`, result: base };
  ws.getCell(5,1+span).value = { formula:`${colL(NC-1)}${tR}-D${tR}`, result: added };
  ws.getCell(5,1+span*2).value = { formula:`${colL(NC-1)}${tR}`, result: base+added };
  ws.autoFilter = `A${HR}:${colL(NC-1)}${HR}`;
  fitExcelCols(ws, HD, drows);

  // ชีตแยกแต่ละเดือน — breakdown ตามคอลัมน์ (รายการย่อย) ของเดือนนั้น
  const cleanNm = (s) => String(s).replace(/[\\/?*[\]:]/g,"-").slice(0,28);
  const usedNm = {};
  months.forEach(m => {
    const cols = (additions[m] && additions[m].$columns) || additions.$columns || [];
    const hasCols = cols.length > 0;
    const valLabels = hasCols ? cols.map(c => c.name || "รายการ") : ["เพิ่มเดือนนี้"];
    const V = valLabels.length, nc = 5 + V, lastValL = colL(3 + V);
    const HM = ["Acc. Code","Account Name","Group","ราคาเดิม", ...valLabels, "รวมเดือนนี้"];
    let nm = cleanNm(monthShortLabel(m)); if (usedNm[nm]) { usedNm[nm]++; nm = cleanNm(`${nm} ${usedNm[nm]}`); } else usedNm[nm] = 1;
    const wsm = wb.addWorksheet(nm, { views:[{ showGridLines:false, state:"frozen", ySplit:4 }] });
    wsm.mergeCells(1,1,1,nc); const mt=wsm.getCell(1,1); mt.value=`เพิ่มรายเดือน ${monthShortLabel(m)} — ${project.name}`; mt.font={bold:true,size:14,color:{argb:"FF1D4ED8"},name:F}; mt.fill=fillS(soft); mt.alignment={vertical:"middle",indent:1}; wsm.getRow(1).height=28;
    wsm.mergeCells(2,1,2,nc); const ms=wsm.getCell(2,1); ms.value = hasCols ? `แยกตามรายการ ${cols.length} คอลัมน์ · Export: ${new Date().toLocaleDateString("th-TH")}` : `Export: ${new Date().toLocaleDateString("th-TH")}`; ms.font={italic:true,size:10,color:{argb:"FF64748B"},name:F}; ms.alignment={indent:1};
    HM.forEach((h,i)=>{ const c=wsm.getCell(4,1+i); c.value=h; c.font={bold:true,size:9.5,color:{argb:"FF1D4ED8"},name:F}; c.fill=fillS("FFDCE6FB"); c.alignment={horizontal:i>2?"right":"left",vertical:"middle",wrapText:true}; c.border={bottom:{style:"medium",color:{argb:"FF2563EB"}}}; }); wsm.getRow(4).height=26;
    const mrows = [];
    list.forEach((a,ri)=>{ const R=5+ri;
      const bs = parseFloat(tenderCosts[a.code])||0;
      const cv = hasCols ? cols.map(c=>parseFloat((additions[m]||{})[`${a.code}:${c.id}`])||0) : [parseFloat((additions[m]||{})[a.code])||0];
      mrows.push([a.code, a.name, a.group, bs, ...cv, cv.reduce((s,v)=>s+v,0)]);
      wsm.getCell(R,1).value=a.code; wsm.getCell(R,2).value=a.name; wsm.getCell(R,3).value=a.group;
      const bc=wsm.getCell(R,4); bc.value=bs; bc.numFmt="#,##0"; bc.alignment={horizontal:"right",vertical:"middle"}; bc.font={name:F,size:9.5};
      cv.forEach((v,vi)=>{ const c=wsm.getCell(R,5+vi); c.value=v; c.numFmt="#,##0"; c.alignment={horizontal:"right",vertical:"middle"}; c.font={name:F,size:9.5}; });
      const tc=wsm.getCell(R,nc); tc.value = { formula:`SUM(E${R}:${lastValL}${R})`, result: cv.reduce((s,v)=>s+v,0) }; tc.numFmt="#,##0"; tc.font={bold:true,name:F,size:9.5}; tc.alignment={horizontal:"right",vertical:"middle"};
      [1,2,3].forEach(c=>{ wsm.getCell(R,c).font={name:F,size:9.5}; wsm.getCell(R,c).alignment={vertical:"middle"}; });
      if(ri%2) for(let c=1;c<=nc;c++){ const cell=wsm.getCell(R,c); if(!cell.fill||!cell.fill.pattern) cell.fill=fillS("FFF4F7FE"); }
    });
    const mtR = 5 + list.length;
    for(let c=1;c<=nc;c++){ const cell=wsm.getCell(mtR,c); cell.fill=fillS("FFC9D8FA"); cell.border={top:{style:"medium",color:{argb:"FF2563EB"}}}; }
    wsm.getCell(mtR,2).value="TOTAL"; wsm.getCell(mtR,2).font={bold:true,color:{argb:"FF1D4ED8"},name:F};
    [4, ...valLabels.map((_,i)=>5+i), nc].forEach(col=>{ const c=wsm.getCell(mtR,col), L=colL(col-1); c.value = list.length ? { formula:`SUM(${L}5:${L}${4+list.length})` } : 0; c.numFmt="#,##0"; c.font={bold:true,color:{argb:"FF1D4ED8"},name:F}; c.alignment={horizontal:"right",vertical:"middle"}; });
    wsm.autoFilter = `A4:${colL(nc-1)}4`;
    fitExcelCols(wsm, HM, mrows);
  });
  const buf=await wb.xlsx.writeBuffer(); const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}); const url=URL.createObjectURL(blob); const a2=document.createElement("a"); a2.href=url; a2.download=`QS_Budget_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`; document.body.appendChild(a2); a2.click(); a2.remove(); setTimeout(()=>URL.revokeObjectURL(url),1500);
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
  const dashAdded  = dashItems.reduce((s,i)=> s + i.value, 0);
  // แถว TOTAL (A1) ของชีต "งบประมาณ"/"รายเดือน (สรุป)" = 5 + จำนวนแถวข้อมูล
  // (หัวข้อ 3 แถว + หัวตารางแถว 4 → ข้อมูลเริ่มแถว 5 → TOTAL อยู่แถว 5+N)
  const TR = 5 + dashList.length;
  const dashItemsF = dashItems.map((it, i) => ({ ...it, f: `'รายเดือน (สรุป)'!${XLSX.utils.encode_col(3 + i)}${TR}` }));
  // สรุปตามกลุ่มวัสดุ (ไว้โชว์ตาราง+แถบสัดส่วนในหน้าสรุป)
  const byG = {};
  dashList.forEach(a => {
    const bs = parseFloat(tenderCosts[a.code]) || 0, tt = parseFloat(combinedBudget[a.code]) || 0, g = a.group || "อื่น ๆ";
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
    cards: [
      { label:"ราคาเดิม (Baseline)", value: dashBase, money:true, f:`'งบประมาณ'!D${TR}` },
      { label:"เพิ่มรายเดือนรวม",    value: dashAdded, money:true, f:`'งบประมาณ'!E${TR}` },
      { label:"งบรวมทั้งหมด",         value: dashBase + dashAdded, money:true, f:`'งบประมาณ'!F${TR}` },
      { label:"จำนวนเดือน",           value: dashMonths.length },
    ],
    chartTitle: "กราฟ: ยอดเพิ่มรายเดือน (THB)",
    items: dashItemsF,
    groups: groupData,
  });
  const monthLinksQS = [];   // เก็บชื่อชีตรายเดือนไว้ทำลิงก์บนหน้าสรุป

  // Sheet 1 — Baseline + monthly additions rolled up per Acc. Code
  const rows1 = [[`งบประมาณ (Tender Cost) — ${project.name}`], [`พื้นที่ ${project.area||"-"} ft²  ·  แผง ${project.panels||"-"}  ·  Export: ${new Date().toLocaleDateString("th-TH")}`], []];
  rows1.push(["Acc. Code","Account Name","Group","ราคาเดิม (Baseline)","เพิ่มรายเดือน (รวม)","งบรวมทั้งหมด",...(U?["งบรวม (USD)"]:[])]);
  const dataStart1 = rows1.length;
  const rowGroups1 = [];
  dashList.forEach(a => {
    const baseline = parseFloat(tenderCosts[a.code]) || 0;
    const total    = parseFloat(combinedBudget[a.code]) || 0;
    const added = total - baseline;
    rows1.push([a.code, a.name, a.group, baseline, added, total, ...(U?[toUsd(total,rate)]:[])]);
    rowGroups1.push(a.group);
  });
  const dataEnd1 = rows1.length-1;
  rows1.push(["","TOTAL","",0,0,0, ...(U?[toUsd(dashBase+dashAdded,rate)]:[])]);
  const totalRow1 = rows1.length-1;
  const ws1 = XLSX.utils.aoa_to_sheet(rows1);
  // ลิงก์ด้วยสูตร: งบรวม = ราคาเดิม + เพิ่ม (ต่อแถว) · TOTAL = ผลรวมทั้งคอลัมน์
  for (let r = dataStart1; r <= dataEnd1; r++) {
    const R = r + 1, ref = XLSX.utils.encode_cell({ r, c:5 });
    if (ws1[ref]) ws1[ref].f = `D${R}+E${R}`;
  }
  ["D","E","F"].forEach((L, i) => {
    const ref = XLSX.utils.encode_cell({ r:totalRow1, c:3+i });
    if (ws1[ref]) ws1[ref].f = `SUM(${L}${dataStart1+1}:${L}${dataEnd1+1})`;
  });
  ws1["!cols"] = [{wch:12},{wch:40},{wch:16},{wch:18},{wch:18},{wch:18},...(U?[{wch:18}]:[])];
  styleSheet(ws1, { numCols:6+(U?1:0), subRows:[1], headerRow:3, dataStart:dataStart1, dataEnd:dataEnd1, totalRow:totalRow1,
    moneyCols:U?[3,4,5,6]:[3,4,5], usdCols:U?[6]:[], theme, rowGroups:rowGroups1, groupDisplayCol:2 });
  xBackLink(ws1, 2, (6+(U?1:0))-1, "สรุป");
  XLSX.utils.book_append_sheet(wb, ws1, "งบประมาณ");

  // Sheet 2 — one column per month, so QS can see exactly how the budget grew
  const months = [...new Set(Object.keys(additions||{}).filter(k=>!k.startsWith("$")))].sort();
  const rows2 = [[`รายการเพิ่มรายเดือน — ${project.name}`], [`Export: ${new Date().toLocaleDateString("th-TH")}`], []];
  rows2.push(["Acc. Code","Account Name","ราคาเดิม", ...months.map(monthShortLabel), "รวมทั้งหมด", ...(U?["รวม (USD)"]:[])]);
  const dataStart2 = rows2.length;
  const rowGroups2 = [];
  dashList.forEach(a => {
    const baseline  = parseFloat(tenderCosts[a.code]) || 0;
    const monthVals = months.map(m => monthAddValue(additions, m, a.code));
    const total = baseline + monthVals.reduce((s,v)=>s+v,0);
    rows2.push([a.code, a.name, baseline, ...monthVals, total, ...(U?[toUsd(total,rate)]:[])]);
    rowGroups2.push(a.group);
  });
  const dataEnd2 = rows2.length-1;
  const M = months.length, totColC = 3 + M;
  rows2.push(["","TOTAL",0, ...months.map(()=>0), 0, ...(U?[toUsd(dashBase+dashAdded,rate)]:[])]);
  const totalRow2 = rows2.length-1;
  const numCols2 = 4 + months.length + (U?1:0);
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  // ลิงก์ด้วยสูตร: รวมทั้งหมด(ต่อแถว) = ราคาเดิม + ผลรวมทุกเดือน · TOTAL = ผลรวมคอลัมน์
  const lastMonthL = XLSX.utils.encode_col(2 + M);
  for (let r = dataStart2; r <= dataEnd2; r++) {
    const R = r + 1, ref = XLSX.utils.encode_cell({ r, c: totColC });
    if (ws2[ref]) ws2[ref].f = M > 0 ? `C${R}+SUM(D${R}:${lastMonthL}${R})` : `C${R}`;
  }
  [2, ...months.map((_,i)=>3+i), totColC].forEach(c => {
    const L = XLSX.utils.encode_col(c), ref = XLSX.utils.encode_cell({ r:totalRow2, c });
    if (ws2[ref]) ws2[ref].f = `SUM(${L}${dataStart2+1}:${L}${dataEnd2+1})`;
  });
  ws2["!cols"] = [{wch:12},{wch:34},{wch:14}, ...months.map(()=>({wch:12})), {wch:16}, ...(U?[{wch:16}]:[])];
  styleSheet(ws2, { numCols:numCols2, subRows:[1], headerRow:3, dataStart:dataStart2, dataEnd:dataEnd2, totalRow:totalRow2,
    moneyCols:[2, ...months.map((_,i)=>3+i), 3+months.length, ...(U?[4+months.length]:[])], usdCols:U?[4+months.length]:[], theme, rowGroups:rowGroups2 });
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
    const valLabels = hasCols ? cols.map(c => c.name || "รายการ") : ["เพิ่มเดือนนี้"];
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
      const vals = hasCols
        ? cols.map(c => parseFloat((additions[m] || {})[`${a.code}:${c.id}`]) || 0)
        : [parseFloat((additions[m] || {})[a.code]) || 0];
      const rowTotal = vals.reduce((s, v) => s + v, 0);
      if (rowTotal <= 0) return; // เอาเฉพาะรายการที่มียอดในเดือนนี้
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
      theme, rowGroups, groupDisplayCol: 2,
    });
    let nm = sheetName(monthShortLabel(m));
    if (usedNames[nm]) { usedNames[nm] += 1; nm = sheetName(`${nm} ${usedNames[nm]}`); } else usedNames[nm] = 1;
    xBackLink(ws, 2, numCols-1, "สรุป");
    XLSX.utils.book_append_sheet(wb, ws, nm);
    monthSheetMap[m] = nm;
    monthLinksQS.push({ text: monthShortLabel(m), sheet: nm });
  });
  // ลิงก์หัวคอลัมน์เดือนในชีต "รายเดือน (สรุป)" → กระโดดไปชีตของเดือนนั้น (ไล่ที่มา)
  months.forEach((m,i)=>{ if (monthSheetMap[m]) xLinkCell(ws2, XLSX.utils.encode_cell({ r:3, c:3+i }), monthSheetMap[m], `ดูรายละเอียดเดือน ${monthShortLabel(m)}`); });
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

  const fname = `QS_Budget_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`;
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
  const valLabels = hasCols ? cols.map(c => c.name || "รายการ") : ["เพิ่มเดือนนี้"];

  const rows = [
    [`เพิ่มรายเดือน ${monthShortLabel(month)} — ${project.name}`],
    [hasCols ? `แยกตามรายการ ${cols.length} คอลัมน์  ·  Export: ${new Date().toLocaleDateString("th-TH")}`
             : `Export: ${new Date().toLocaleDateString("th-TH")}`],
    [],
    ["Acc. Code", "Account Name", "Group", "ราคาเดิม", ...valLabels, "รวมเดือนนี้", "รวมสะสมถึงเดือนนี้", ...(U?["รวมสะสม (USD)"]:[])],
  ];
  const dataStart = rows.length;
  const colTotals = valLabels.map(() => 0);
  let gBase = 0, gMonth = 0, gCum = 0;
  const rowGroups = [];
  accounts.forEach(a => {
    const baseline = parseFloat(tenderCosts[a.code]) || 0;
    const vals = hasCols
      ? cols.map(c => parseFloat((additions[month] || {})[`${a.code}:${c.id}`]) || 0)
      : [parseFloat((additions[month] || {})[a.code]) || 0];
    const monthTot = vals.reduce((s, v) => s + v, 0);
    const cum = baseline + upto.reduce((s, m) => s + monthAddValue(additions, m, a.code), 0);
    if (monthTot <= 0 && baseline <= 0 && cum <= 0) return;
    rows.push([a.code, a.name, a.group, baseline, ...vals, monthTot, cum, ...(U?[toUsd(cum,rate)]:[])]);
    rowGroups.push(a.group);
    vals.forEach((v, i) => { colTotals[i] += v; });
    gBase += baseline; gMonth += monthTot; gCum += cum;
  });
  const dataEnd = rows.length - 1;
  rows.push(["", "TOTAL", "", gBase, ...colTotals, gMonth, gCum, ...(U?[toUsd(gCum,rate)]:[])]);
  const totalRow = rows.length - 1;
  const numCols = 6 + valLabels.length + (U?1:0);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch:12 }, { wch:34 }, { wch:14 }, { wch:16 }, ...valLabels.map(()=>({ wch:15 })), { wch:16 }, { wch:18 }, ...(U?[{ wch:18 }]:[])];
  styleSheet(ws, {
    numCols, subRows:[1], headerRow:3, dataStart, dataEnd, totalRow,
    moneyCols: [3, ...valLabels.map((_, i) => 4 + i), 4 + valLabels.length, 5 + valLabels.length, ...(U?[6 + valLabels.length]:[])], usdCols:U?[6 + valLabels.length]:[],
    theme, rowGroups, groupDisplayCol: 2,
  });
  XLSX.utils.book_append_sheet(wb, ws, clean(monthShortLabel(month)));
  XLSX.writeFile(wb, `QS_${clean(monthShortLabel(month)).replace(/[^\dA-Za-zก-๙]/g,"")}_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── Procurement: PO tracking export ───────────────────────────────────────
function exportProcurementExcel(project, poEntries) {
  const wb = XLSX.utils.book_new();
  const theme = { main:"F59E0B", dark:"B45309" };
  const rate = exportRate(project); const U = rate > 0;   // U = ใส่คอลัมน์ USD ไหม

  // หน้าแรก = สรุป (Dashboard): การ์ดตัวเลข + กราฟยอดสั่งซื้อรายเดือน
  const dPaid = poEntries.reduce((s,p)=> s + poRounds(p).filter(r=>roundPaid(p,r)).reduce((ss,r)=> ss + (parseFloat(r.actualAmount)||0), 0), 0);
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
    const pay = paymentStatus(p);
    const deliveryStr = poRounds(p).map(r => `${r.plan||"-"}→${r.actual||"รอ"}`).join(" | ") || "-";
    poItems(p).forEach(it => {
      const acc = ACCOUNTS.find(a=>a.code===it.code);
      const amount = parseFloat(it.amount) || 0;
      rows1.push([p.date, it.code, acc?.name||"", itemSupplierName(p), poNumbersLabel(p), amount, ...(U?[toUsd(amount,rate)]:[]), p.status, deliveryStr, poNextDueDate(p)||"-", PAYMENT_LABEL[pay], p.notes||""]);
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
  dash.ws["!cols"] = [{wch:12},{wch:10},{wch:34},{wch:22},{wch:16},{wch:16},...(U?[{wch:16}]:[]),{wch:12},{wch:26},{wch:16},{wch:16},{wch:28}];

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
            const acc = ACCOUNTS.find(a=>a.code===it.code);
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

  XLSX.writeFile(wb, `Procurement_PO_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// จัดซื้อ (PO) แบบ rich: หน้า "สรุป" หน้าเดียว มีการ์ด + pie (สถานะ) + bar (รายเดือน)
// + ตารางสรุปสถานะ และชีต "PO ทั้งหมด"
async function exportPORich(project, poEntries) {
  const ExcelJS = await loadExcelJS();
  const F = "Tahoma";
  const total = poEntries.reduce((s,p)=> s + poTotal(p), 0);
  const paid  = poEntries.reduce((s,p)=> s + poRounds(p).filter(r=>roundPaid(p,r)).reduce((ss,r)=> ss + (parseFloat(r.actualAmount)||0), 0), 0);
  const outstanding = Math.max(0, total - paid);
  const HD = ["วันเปิด PO","Acc. Code","Account Name","Supplier","PO No.","มูลค่า (THB)","สถานะ PO","ของเข้า (แผน→จริง)","วันที่รับของ","วันครบกำหนดจ่าย","สถานะจ่ายเงิน","หมายเหตุ"];
  const NC = HD.length;
  const rows = [];
  poEntries.slice().sort((a,b)=>(a.date||"").localeCompare(b.date||"")).forEach(p => {
    const pay = PAYMENT_LABEL[paymentStatus(p)];
    const delivery = poRounds(p).map(r => `${r.plan||"-"}→${r.actual||"รอ"}`).join(" | ") || "-";
    const received = poRounds(p).map(r => r.actual).filter(Boolean).join(", ") || "-";
    const due = poNextDueDate(p) || "-";
    poItems(p).forEach(it => {
      const acc = ACCOUNTS.find(a=>a.code===it.code);
      rows.push([ p.date||"", it.code, acc?.name||"", poSupplierName(p), poNumbersLabel(p), parseFloat(it.amount)||0, p.status||"-", delivery, received, due, pay, p.notes||"" ]);
    });
  });
  const soft = "FF"+lighten("F59E0B",0.55), colL = c => XLSX.utils.encode_col(c);
  const wb = new ExcelJS.Workbook();
  const fillS = (a) => ({ type:"pattern", pattern:"solid", fgColor:{argb:a} });
  const ws = wb.addWorksheet("รายงานจัดซื้อ", { views:[{ showGridLines:false, state:"frozen", ySplit:7 }] });
  ws.mergeCells(1,1,1,NC); const t=ws.getCell(1,1); t.value=`สรุปจัดซื้อ (PO) — ${project.name}`; t.font={bold:true,size:15,color:{argb:"FF92400E"},name:F}; t.fill=fillS(soft); t.alignment={vertical:"middle",indent:1}; ws.getRow(1).height=30;
  ws.mergeCells(2,1,2,NC); const stc=ws.getCell(2,1); stc.value=`Export: ${new Date().toLocaleDateString("th-TH")} · ทั้งหมด ${poEntries.length} PO`; stc.font={italic:true,size:10,color:{argb:"FF64748B"},name:F}; stc.alignment={indent:1};
  const cards=[["มูลค่า PO รวม","FFFEF3C7","FF92400E",total],["จ่ายแล้ว","FFD1FAE5","FF047857",paid],["ค้างจ่าย","FFFEE2E2","FF991B1B",outstanding],["จำนวน PO","FFDBEAFE","FF1D4ED8",poEntries.length]];
  const span = Math.max(2, Math.floor(NC/4));
  cards.forEach((cd,i)=>{ const c0=1+i*span, c1=Math.min(NC, c0+span-1);
    ws.mergeCells(4,c0,4,c1); ws.mergeCells(5,c0,5,c1);
    const lc=ws.getCell(4,c0); lc.value=cd[0]; lc.font={bold:true,size:10,color:{argb:cd[2]},name:F}; lc.fill=fillS(cd[1]); lc.alignment={horizontal:"center",vertical:"middle"};
    const vc=ws.getCell(5,c0); vc.value=cd[3]; if(i<3) vc.numFmt="#,##0"; vc.font={bold:true,size:15,color:{argb:cd[2]},name:F}; vc.fill=fillS(cd[1]); vc.alignment={horizontal:"center",vertical:"middle"};
  }); ws.getRow(4).height=18; ws.getRow(5).height=30;
  const HR = 7;
  HD.forEach((h,i)=>{ const c=ws.getCell(HR,1+i); c.value=h; c.font={bold:true,size:9.5,color:{argb:"FF92400E"},name:F}; c.fill=fillS("FFFDEED3"); c.alignment={horizontal:i===5?"right":"left",vertical:"middle",wrapText:true}; c.border={bottom:{style:"medium",color:{argb:"FFF59E0B"}}}; }); ws.getRow(HR).height=26;
  const pillOf = (s) => { const p=statusPill(s); return p ? { fill:fillS("FF"+p.bg), font:{bold:true,size:9.5,color:{argb:"FF"+p.fg},name:F} } : null; };
  rows.forEach((row,ri)=>{ const R=HR+1+ri;
    row.forEach((val,ci)=>{ const c=ws.getCell(R,1+ci); c.value=val;
      if(ci===5){ c.numFmt="#,##0"; c.alignment={horizontal:"right",vertical:"middle"}; c.font={name:F,size:9.5}; }
      else if(ci===6 || ci===10){ const pl=pillOf(val); c.alignment={horizontal:"center",vertical:"middle"}; if(pl){c.fill=pl.fill;c.font=pl.font;} else c.font={name:F,size:9.5}; }
      else { c.alignment={vertical:"middle"}; c.font={name:F,size:9.5}; }
    });
    if(ri%2) for(let c=1;c<=NC;c++){ const cell=ws.getCell(R,c); if(!cell.fill||!cell.fill.pattern) cell.fill=fillS("FFFFFAF3"); }
  });
  const tR = HR + 1 + rows.length;
  for(let c=1;c<=NC;c++){ const cell=ws.getCell(tR,c); cell.fill=fillS("FFFDE7C2"); cell.border={top:{style:"medium",color:{argb:"FFF59E0B"}}}; }
  const tl=ws.getCell(tR,5); tl.value="TOTAL"; tl.font={bold:true,color:{argb:"FF92400E"},name:F}; tl.alignment={horizontal:"right",vertical:"middle"};
  const tvc=ws.getCell(tR,6); tvc.value = rows.length ? { formula:`SUM(F${HR+1}:F${HR+rows.length})`, result: total } : 0; tvc.numFmt="#,##0"; tvc.font={bold:true,color:{argb:"FF92400E"},name:F}; tvc.alignment={horizontal:"right",vertical:"middle"};
  ws.getCell(5,1).value = { formula:`F${tR}`, result: total };
  ws.autoFilter = `A${HR}:${colL(NC-1)}${HR}`;
  fitExcelCols(ws, HD, rows);

  // ─── ชีต "รายเดือน" — แจกแจงราย PO ตามเดือนที่เปิด PO (วันเปิด PO) ─────────
  const moKeys = [...new Set(poEntries.map(p => (p.date||"").slice(0,7)).filter(Boolean))].sort();
  if (moKeys.length) {
    const MH  = ["เดือน","วันเปิด PO","PO No.","Supplier","มูลค่า PO (THB)","จ่ายแล้ว (THB)","ค้างจ่าย (THB)","% จ่ายแล้ว"];
    const MNC = MH.length;
    const poPaid = (p) => poRounds(p).filter(r=>roundPaid(p,r)).reduce((ss,r)=> ss + (parseFloat(r.actualAmount)||0), 0);
    const wsm = wb.addWorksheet("รายเดือน", { views:[{ showGridLines:false, state:"frozen", ySplit:4 }] });
    wsm.mergeCells(1,1,1,MNC); const mt=wsm.getCell(1,1); mt.value=`จัดซื้อรายเดือน (ตามวันเปิด PO) — ${project.name}`; mt.font={bold:true,size:14,color:{argb:"FF92400E"},name:F}; mt.fill=fillS(soft); mt.alignment={vertical:"middle",indent:1}; wsm.getRow(1).height=28;
    wsm.mergeCells(2,1,2,MNC); const mst=wsm.getCell(2,1); mst.value=`Export: ${new Date().toLocaleDateString("th-TH")} · ${moKeys.length} เดือน · ${poEntries.length} PO`; mst.font={italic:true,size:10,color:{argb:"FF64748B"},name:F}; mst.alignment={indent:1};
    const MHR = 4;
    MH.forEach((h,i)=>{ const c=wsm.getCell(MHR,1+i); c.value=h; c.font={bold:true,size:9.5,color:{argb:"FF92400E"},name:F}; c.fill=fillS("FFFDEED3"); c.alignment={horizontal:i>=4?"right":"left",vertical:"middle",wrapText:true}; c.border={bottom:{style:"medium",color:{argb:"FFF59E0B"}}}; }); wsm.getRow(MHR).height=24;
    // สร้างแถว: หนึ่งแถวต่อ PO จัดกลุ่มตามเดือน + แถว "รวมเดือน" ท้ายแต่ละกลุ่ม
    const bodyRows = [];
    moKeys.forEach(m => {
      const list = poEntries.filter(p => (p.date||"").slice(0,7) === m).sort((a,b)=>(a.date||"").localeCompare(b.date||""));
      let sT=0, sP=0;
      list.forEach((p,idx) => {
        const t=poTotal(p), pd=poPaid(p);
        sT+=t; sP+=pd;
        bodyRows.push({ type:"po", month: idx===0?monthShortLabel(m):"", date:p.date||"-", no:poNumbersLabel(p), sup:poSupplierName(p), total:t, paid:pd, out:Math.max(0,t-pd) });
      });
      bodyRows.push({ type:"sub", label:`รวม ${monthShortLabel(m)}`, total:sT, paid:sP, out:Math.max(0,sT-sP) });
    });
    let po_i = 0;
    bodyRows.forEach((r,ri) => { const R = MHR+1+ri;
      if (r.type === "sub") {
        for(let c=1;c<=MNC;c++){ const cell=wsm.getCell(R,c); cell.fill=fillS("FFFDEED3"); }
        const lc=wsm.getCell(R,1); lc.value=r.label; lc.font={bold:true,size:9.5,color:{argb:"FF92400E"},name:F}; lc.alignment={vertical:"middle",indent:1};
        [r.total, r.paid, r.out, r.total>0?r.paid/r.total:0].forEach((v,i)=>{ const c=wsm.getCell(R,5+i); c.value=v; c.numFmt = i===3 ? "0%" : "#,##0"; c.font={bold:true,size:9.5,color:{argb:"FF92400E"},name:F}; c.alignment={horizontal:"right",vertical:"middle"}; });
        return;
      }
      [r.month, r.date, r.no, r.sup, r.total, r.paid, r.out, r.total>0?r.paid/r.total:0].forEach((val,ci)=>{
        const c=wsm.getCell(R,1+ci); c.value=val; c.font={name:F,size:9.5};
        if(ci<=3){ c.alignment={vertical:"middle",indent:ci===0?1:0}; if(ci===0) c.font={name:F,size:9.5,bold:true,color:{argb:"FF92400E"}}; }
        else if(ci===7){ c.numFmt="0%"; c.alignment={horizontal:"right",vertical:"middle"}; }
        else { c.numFmt="#,##0"; c.alignment={horizontal:"right",vertical:"middle"}; }
      });
      if(po_i%2) for(let c=1;c<=MNC;c++){ const cell=wsm.getCell(R,c); if(!cell.fill||!cell.fill.pattern) cell.fill=fillS("FFFFFAF3"); }
      po_i++;
    });
    const mtR = MHR + 1 + bodyRows.length;
    for(let c=1;c<=MNC;c++){ const cell=wsm.getCell(mtR,c); cell.fill=fillS("FFFDE7C2"); cell.border={top:{style:"medium",color:{argb:"FFF59E0B"}}}; }
    const mtl=wsm.getCell(mtR,1); mtl.value="TOTAL"; mtl.font={bold:true,color:{argb:"FF92400E"},name:F}; mtl.alignment={vertical:"middle",indent:1};
    const gT=poEntries.reduce((s,p)=>s+poTotal(p),0), gP=poEntries.reduce((s,p)=>s+poPaid(p),0);
    [gT, gP, Math.max(0,gT-gP), gT>0?gP/gT:0].forEach((v,i)=>{ const c=wsm.getCell(mtR,5+i); c.value=v; c.numFmt = i===3 ? "0%" : "#,##0"; c.font={bold:true,color:{argb:"FF92400E"},name:F}; c.alignment={horizontal:"right",vertical:"middle"}; });
    wsm.getColumn(1).width=14; wsm.getColumn(2).width=14; wsm.getColumn(3).width=16; wsm.getColumn(4).width=26;
    for(let c=5;c<=MNC;c++) wsm.getColumn(c).width = c===MNC ? 12 : 16;
    wsm.autoFilter = `A${MHR}:${colL(MNC-1)}${MHR}`;
  }

  const buf=await wb.xlsx.writeBuffer(); const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}); const url=URL.createObjectURL(blob); const a2=document.createElement("a"); a2.href=url; a2.download=`Procurement_PO_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`; document.body.appendChild(a2); a2.click(); a2.remove(); setTimeout(()=>URL.revokeObjectURL(url),1500);
}


// ─── Accounting: full financial export ─────────────────────────────────────
function exportAccountingExcel(project, tenderCosts, additions, poEntries, extraItems=[], hiddenAccounts=[]) {
  const wb = XLSX.utils.book_new();
  const theme = { main:"10B981", dark:"047857" };
  const rate = exportRate(project); const U = rate > 0;   // U = ใส่คอลัมน์ USD ไหม
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);
  const accounts = exportAccountList(extraItems, hiddenAccounts);

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
    const status   = committed>budget && budget>0 ? "เกินงบ" : committed>0 ? "OK" : budget>0 ? "ยังไม่ PO" : "-";
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
    const pay = paymentStatus(p);
    const deliveryStr = poRounds(p).map(r => `${r.plan||"-"}→${r.actual||"รอ"}`).join(" | ") || "-";
    poItems(p).forEach(it => {
      const acc = ACCOUNTS.find(a=>a.code===it.code);
      const amount = parseFloat(it.amount) || 0;
      rows2.push([p.date, it.code, acc?.name||"", acc?.group||"", itemSupplierName(p), poNumbersLabel(p), amount, ...(U?[toUsd(amount,rate)]:[]), p.status, deliveryStr, poNextDueDate(p)||"-", PAYMENT_LABEL[pay]]);
      rowGroups2.push(acc?.group || "-");
      grand2 += amount;
    });
  });
  const dataEnd2 = rows2.length-1;
  rows2.push(["","","","","","TOTAL", grand2, ...(U?[toUsd(grand2,rate)]:[]),"","","",""]);
  const totalRow2 = rows2.length-1;
  const ws2 = XLSX.utils.aoa_to_sheet(rows2);
  ws2["!cols"] = [{wch:12},{wch:10},{wch:34},{wch:14},{wch:22},{wch:16},{wch:16},...(U?[{wch:16}]:[]),{wch:12},{wch:26},{wch:16},{wch:16}];
  styleSheet(ws2, { numCols:11+(U?1:0), subRows:[1], headerRow:3, dataStart:dataStart2, dataEnd:dataEnd2, totalRow:totalRow2,
    moneyCols:U?[6,7]:[6], usdCols:U?[7]:[], centerCols:U?[8,11]:[7,10], theme, rowGroups:rowGroups2, groupDisplayCol:3 });
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
    const baselineTotal = accounts.reduce((s,a)=>s+(parseFloat(tenderCosts[a.code])||0),0);
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
    const payMonths = [...new Set(payLines.map(monthKey))].sort();

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

  // ลิงก์นำทางใต้ตารางหน้า Summary — คลิกเพื่อไปดูที่มาของตัวเลขในแต่ละชีต
  {
    const acctLinks = [
      { text:"📦 PO Entries (รายการ PO)", sheet:"PO Entries" },
      { text:"🏷 By Group (ตามกลุ่ม)", sheet:"By Group" },
      ...(allMonths.length ? [{ text:"📅 รายเดือน", sheet:"รายเดือน" }] : []),
      ...(payLines.length ? [{ text:"💰 แผนจ่าย", sheet:"แผนจ่าย" }] : []),
    ];
    const navR = totalRow1 + 2;
    XLSX.utils.sheet_add_aoa(ws1, [["🔗 ไปที่ชีต (ไล่ที่มาของตัวเลข):"]], { origin:{ r:navR, c:0 } });
    ws1[XLSX.utils.encode_cell({ r:navR, c:0 })].s = { font:{ bold:true, sz:10.5, color:{rgb:theme.dark}, name:"Tahoma" } };
    xLinkRow(ws1, navR+1, acctLinks);
  }

  XLSX.writeFile(wb, `Accounting_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
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
          <div style={{fontSize:40}}>😵</div>
          <div style={{fontSize:16,fontWeight:700,color:"#0f172a"}}>เกิดข้อผิดพลาดในการแสดงผลหน้านี้</div>
          <div style={{fontSize:13,color:"#64748b",maxWidth:460}}>ข้อมูลของคุณยังปลอดภัย ลองกดปุ่มด้านล่างเพื่อโหลดใหม่ ถ้ายังเป็นอยู่ให้แจ้งผู้ดูแลระบบ</div>
          <button onClick={()=>{ this.setState({err:null}); if(typeof window!=="undefined") window.location.reload(); }}
            style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:10,padding:"9px 20px",fontSize:14,fontWeight:600,cursor:"pointer"}}>โหลดหน้าใหม่</button>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  const [session,  setSessionState] = useState(null);   // โหลดแบบ async ด้านล่าง
  const [authReady, setAuthReady]   = useState(false);  // true เมื่อเช็ค session เสร็จ
  const [screen,   setScreen]   = useState("home");
  const [projects, setProjects] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [role,     setRole]     = useState(null);
  const [tenderCosts, setTCosts]= useState({});
  const [additions,   setAdditions]  = useState({});
  const [extraItems,  setExtraItems] = useState([]);
  const [hiddenAccounts, setHiddenAccounts] = useState([]); // codes of fixed Acc. Codes QS has removed for this project
  const [poEntries,   setPO]    = useState([]);
  const [loaded,   setLoaded]   = useState(false);
  const [newProjModal, setNewProjModal] = useState(false);
  const [syncedAt,    setSyncedAt]    = useState(null);
  const [syncing,     setSyncing]     = useState(false);
  const [syncError,   setSyncError]   = useState("");   // ข้อความเตือนเมื่อบันทึก/โหลดพลาด
  const [exportMsg,   setExportMsg]   = useState("");   // สถานะตอนกด Export (กำลังสร้าง/เสร็จ/พลาด)
  const runExport = async (fn) => {
    setExportMsg("⏳ กำลังสร้างไฟล์ Excel…");
    try { await fn(); setExportMsg("✓ สร้างไฟล์เรียบร้อย — ดูที่โฟลเดอร์ดาวน์โหลด"); setTimeout(()=>setExportMsg(""), 3500); }
    catch (e) { console.warn("export failed:", e); setExportMsg("⚠ สร้างไฟล์ไม่สำเร็จ ลองใหม่อีกครั้ง"); setTimeout(()=>setExportMsg(""), 4500); }
  };

  const handleLogin = (user) => { setSession(user); setSessionState(user); };
  const handleLogout = () => {
    clearSession(); setSessionState(null);
    setScreen("home"); setRole(null); setActiveId(null);
  };

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
    const t  = await sg(`tcs-tenders-${id}`);
    const po = await sg(`tcs-po-${id}`);
    const ad = await sg(`tcs-additions-${id}`);
    const ex = await sg(`tcs-extra-${id}`);
    const hid = await sg(`tcs-hidden-${id}`);
    setTCosts(t || {});
    setPO(po || []);
    setAdditions(ad || {});
    setExtraItems(ex || []);
    setHiddenAccounts(hid || []);
  }, []);

  const fetchProjects = useCallback(async () => {
    const list = await sg("tcs-projects");
    if (list) setProjects(list);
  }, []);

  // โหลดรายการโครงการ "หลังล็อกอินเสร็จ" — สำคัญมากตอนใช้ RLS: ถ้าอ่านก่อน
  // Supabase แนบ token จะโดน DB ปฏิเสธแล้วขึ้น 0 โครงการ ทั้งที่มีสิทธิ์อ่าน
  // ผูกกับ session ไว้ พอล็อกอินเสร็จ (session มีค่า) จะดึงข้อมูลใหม่อัตโนมัติ
  useEffect(() => {
    if (!session) { setLoaded(true); return; }
    (async () => {
      try { await fetchProjects(); setSyncedAt(new Date()); setSyncError(""); }
      catch (e) { console.warn("โหลดรายการโครงการไม่สำเร็จ:", e); setSyncError("โหลดข้อมูลไม่สำเร็จ — ตรวจสอบเน็ตแล้วรีเฟรชหน้า"); }
      finally { setLoaded(true); }   // กันจอโหลดค้างเสมอ แม้ดึงข้อมูลพลาด
    })();
  }, [fetchProjects, session]);

  useEffect(() => {
    if (!activeId || !session) return;
    fetchProjectData(activeId).catch(e => { console.warn("โหลดข้อมูลโครงการไม่สำเร็จ:", e); setSyncError("โหลดข้อมูลโครงการไม่สำเร็จ — ลองเปิดใหม่อีกครั้ง"); });
  }, [activeId, fetchProjectData, session]);

  useEffect(() => {
    if (!session) return; // subscribe realtime หลังล็อกอิน เพื่อให้ RLS ยอมส่ง event
    const channel = supabase.channel("kv_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "kv_store" }, async (payload) => {
        const key = payload.new?.key || payload.old?.key || "";
        setSyncing(true);
        try {
          if (key === "tcs-projects") await fetchProjects();
          else if (activeId && (key === `tcs-tenders-${activeId}` || key === `tcs-po-${activeId}` || key === `tcs-additions-${activeId}` || key === `tcs-extra-${activeId}` || key === `tcs-hidden-${activeId}`)) await fetchProjectData(activeId);
          setSyncedAt(new Date());
        } catch (e) {
          console.warn("sync realtime ล้มเหลว:", e);
        } finally {
          setSyncing(false); // กันสปินเนอร์ค้างเมื่อ fetch ล้มเหลว
        }
      }).subscribe();
    return () => supabase.removeChannel(channel);
  }, [activeId, fetchProjects, fetchProjectData, session]);

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
  };
  const syncUndo = () => setUndoInfo({
    u: undoRef.current.length, r: redoRef.current.length,
    label: undoRef.current.length ? undoRef.current[undoRef.current.length - 1].label : "",
  });
  // เขียนลงเซิร์ฟเวอร์ พร้อมลองใหม่อัตโนมัติ 1 ครั้งเมื่อเน็ตสะดุดชั่วคราว ก่อนค่อย
  // แจ้งเตือน (กันเซฟหลุดเพราะ blip เล็ก ๆ). ถ้าส่ง prev มาด้วย จะใช้ ssMerge เพื่อ
  // "รวม" การแก้ของเราลงบนของล่าสุดบนเซิร์ฟเวอร์ (กันทับงานคนอื่นที่แก้พร้อมกัน).
  const persist = (key, value, prev) => {
    const attempt = () => (prev !== undefined ? ssMerge(key, prev, value) : ss(key, value));
    return attempt()
      .then(() => { setSyncedAt(new Date()); setSyncError(""); })
      .catch(() => new Promise(res => setTimeout(res, 900)).then(attempt)
        .then(() => { setSyncedAt(new Date()); setSyncError(""); })
        .catch(e => { console.warn("บันทึกไม่สำเร็จ (ลองใหม่แล้ว):", key, e); setSyncError("⚠ บันทึกไม่สำเร็จ — ข้อมูลล่าสุดอาจยังไม่ถูกบันทึก กรุณาลองใหม่/ตรวจเน็ต"); }));
  };
  const commit = useCallback((key, next, prev, setState, label) => {
    undoRef.current.push({ key, value: prev, setState, label });
    if (undoRef.current.length > 60) undoRef.current.shift();
    redoRef.current = []; // มีการแก้ใหม่ → ล้าง redo
    setState(next);
    persist(key, next, prev);
    syncUndo();
  }, []);
  const undo = useCallback(() => {
    const e = undoRef.current.pop();
    if (!e) return;
    const cur = currentRef.current[e.key];
    redoRef.current.push({ key: e.key, value: cur, setState: e.setState, label: e.label });
    e.setState(e.value);
    persist(e.key, e.value, cur);
    syncUndo();
  }, []);
  const redo = useCallback(() => {
    const e = redoRef.current.pop();
    if (!e) return;
    const cur = currentRef.current[e.key];
    undoRef.current.push({ key: e.key, value: cur, setState: e.setState, label: e.label });
    e.setState(e.value);
    persist(e.key, e.value, cur);
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
    const NUM_RE = /^-?\d[\d,]*\.\d+$/;
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
            return NUM_RE.test((n.nodeValue||"").trim()) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
          },
        });
        let node;
        while ((node = walker.nextNode())) {
          const rng = document.createRange(); rng.selectNodeContents(node);
          const r = rng.getBoundingClientRect();
          if (r.width === 0 && r.height === 0) continue;
          if (r.right >= box.left && r.left <= box.right && r.bottom >= box.top && r.top <= box.bottom) {
            const txt = node.nodeValue.trim();
            const v = parseFloat(txt.replace(/,/g, ""));
            if (!isNaN(v)) {
              nums.push(v);
              cellData.push({ top: r.top, left: r.left, text: txt });
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
      clearHilite(); setSelStats(null); setMarquee(null); selCellsRef.current = []; // คลิกที่ไหนก็ล้างไฮไลต์เดิม
      if (e.button !== 0) return;
      const t = e.target;
      if (!(t instanceof Element) || t.closest(INTERACT) || !t.closest("table")) return;
      // เริ่มลากเลือกสถิติเฉพาะเมื่อเริ่มบนเซลล์ที่เป็น "ยอดเงิน" (มีจุดทศนิยม)
      // ถ้าเริ่มบนเซลล์ข้อความ (รหัสบัญชี/ชื่อรายการ/หัวตาราง) ปล่อยให้เลือก-คัดลอกข้อความได้ตามปกติ
      const startCell = t.closest("td");
      if (!startCell || !/\d[\d,]*\.\d/.test(startCell.textContent || "")) return;
      d.scrollEl = t.closest(".mscroll") || t.closest(".hscroll") || null;
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

  const saveProjects = useCallback((list) => commit("tcs-projects", list, projects, setProjects, "รายชื่อโครงการ"), [commit, projects]);
  const saveTenders  = useCallback((t)    => commit(`tcs-tenders-${activeId}`, t, tenderCosts, setTCosts, "ราคาเดิม (Baseline)"), [commit, activeId, tenderCosts]);
  const saveAdditions= useCallback((a)    => commit(`tcs-additions-${activeId}`, a, additions, setAdditions, "ยอดเพิ่มรายเดือน"), [commit, activeId, additions]);
  const saveExtraItems=useCallback((ex)   => commit(`tcs-extra-${activeId}`, ex, extraItems, setExtraItems, "รายการ/แถว"), [commit, activeId, extraItems]);
  const saveHiddenAccounts=useCallback((h)=> commit(`tcs-hidden-${activeId}`, h, hiddenAccounts, setHiddenAccounts, "การซ่อนหมวด"), [commit, activeId, hiddenAccounts]);
  const savePO       = useCallback((po)   => commit(`tcs-po-${activeId}`, po, poEntries, setPO, "PO / จัดซื้อ"), [commit, activeId, poEntries]);

  const openProject = (id) => {
    setActiveId(id);
    if (session?.role === "admin") { setRole(null); setScreen("roleSelect"); }
    else { setRole(session?.role); setScreen("app"); }
  };
  const deleteProject = async (id) => {
    const proj = projects.find(p => p.id === id);
    const name = (proj?.name || "").trim();
    // ยืนยันแบบ "พิมพ์ชื่อโครงการให้ตรง" — กันเผลอลบ เพราะลบแล้วข้อมูลย่อยหายด้วย
    const typed = window.prompt(
      `⚠️ ลบโครงการ "${name}" ?\n\n` +
      `ข้อมูลทั้งหมดของโครงการนี้จะถูกลบด้วย:\n` +
      `• Tender Cost (ราคาเดิม)\n• PO / จัดซื้อ\n• ยอดเพิ่มรายเดือน · รายการเพิ่ม · หมวดที่ซ่อน\n\n` +
      `กู้คืนได้จาก Admin → กู้คืนข้อมูล (ได้ถึงสแนปช็อตล่าสุด 12:00/18:00)\n\n` +
      `ถ้าแน่ใจ พิมพ์ชื่อโครงการให้ตรงเพื่อยืนยัน:\n${name}`
    );
    if (typed == null) return;                                   // กดยกเลิก
    if (typed.trim() !== name) { alert("ชื่อโครงการไม่ตรง — ยกเลิกการลบแล้ว"); return; }
    // ไม่เข้า quick-undo เพราะการลบโครงการลบคีย์ย่อยด้วย — กู้ทั้งโครงการทำผ่านหน้า
    // Admin กู้คืนข้อมูล (kv_history เก็บไว้ให้ครบทุกคีย์)
    const next = projects.filter(p => p.id !== id);
    setProjects(next); ss("tcs-projects", next).then(()=>setSyncedAt(new Date()));
    await sd(`tcs-tenders-${id}`); await sd(`tcs-po-${id}`); await sd(`tcs-additions-${id}`); await sd(`tcs-extra-${id}`); await sd(`tcs-hidden-${id}`);
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
      </>
    );
  }

  if (!loaded) return <Loader />;

  // Non-admins can only ever act as the role tied to their account,
  // even if they somehow land on screen "roleSelect" or "app" with a stale role.
  const effectiveRole = session.role === "admin" ? role : session.role;

  const sharedProps = { project:activeProject, tenderCosts, poEntries, saveTenders, savePO,
    additions, saveAdditions, extraItems, saveExtraItems, hiddenAccounts, saveHiddenAccounts,
    updateProject, onBack: () => setScreen(session.role === "admin" ? "roleSelect" : "home"),
    onHome: () => setScreen("home"),   // ปุ่ม Home → หน้าเลือกโครงการ (ทุกโรล)
    syncedAt, syncing, session, onLogout: handleLogout, setEditMode };

  return (
    <>
      <style>{GLOBAL_CSS}</style>
      {syncError && (
        <div style={{position:"fixed",left:"50%",top:16,transform:"translateX(-50%)",zIndex:200,maxWidth:"92vw",
          background:"#fef2f2",color:"#991b1b",border:"1px solid #ef4444",borderRadius:12,padding:"10px 16px",
          boxShadow:"0 8px 28px rgba(15,23,42,0.18)",fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:12}}>
          <span style={{flex:1}}>{syncError}</span>
          <button onClick={()=>setSyncError("")} style={{border:"none",background:"none",color:"#991b1b",cursor:"pointer",fontSize:16,fontWeight:700,lineHeight:1}}>×</button>
        </div>
      )}
      {exportMsg && (
        <div style={{position:"fixed",left:"50%",bottom:22,transform:"translateX(-50%)",zIndex:200,
          background:"#0f172a",color:"#e2e8f0",borderRadius:10,padding:"10px 18px",boxShadow:"0 8px 28px rgba(15,23,42,0.28)",
          fontSize:13,fontWeight:600,whiteSpace:"nowrap"}}>{exportMsg}</div>
      )}
      {marquee && marquee.width > 2 && marquee.height > 2 && (
        <div style={{position:"fixed",left:marquee.left,top:marquee.top,width:marquee.width,height:marquee.height,
          background:"rgba(37,99,235,0.06)",border:"none",zIndex:97,pointerEvents:"none"}}/>
      )}
      {selStats && (
        <div style={{position:"fixed",right:20,bottom:20,zIndex:96,display:"flex",alignItems:"center",gap:0,
          background:"#1e293b",color:"#e2e8f0",borderRadius:10,padding:"8px 4px",boxShadow:"0 8px 28px rgba(15,23,42,0.28)",
          fontSize:12.5,fontFamily:"'JetBrains Mono',monospace",overflow:"hidden"}}>
          {(() => {
            const selRate = effRate(activeProject);   // อัตราแลกเปลี่ยน (0 = ปิด/ไม่โชว์ $)
            const segs = [
              {label:"ผลรวม", raw:selStats.sum, clr:"#34d399", money:true},
              {label:"เฉลี่ย", raw:selStats.avg, clr:"#93c5fd", money:true},
              {label:"นับ",   text:String(selStats.count), clr:"#fcd34d", money:false},
              {label:"ต่ำสุด", raw:selStats.min, clr:"#cbd5e1", money:true},
              {label:"สูงสุด", raw:selStats.max, clr:"#cbd5e1", money:true},
            ];
            return segs.map((s,i)=>(
              <span key={s.label} style={{display:"flex",alignItems:"center",gap:6,padding:"0 12px",borderLeft:i?"1px solid #334155":"none"}}>
                <span style={{color:"#94a3b8",fontFamily:"system-ui,sans-serif",fontSize:11}}>{s.label}</span>
                <span style={{display:"flex",flexDirection:"column",alignItems:"flex-end",lineHeight:1.15}}>
                  <b style={{color:s.clr}}>{s.money ? fmt(s.raw) : s.text}</b>
                  {s.money && selRate>0 && <b style={{color:"#34d399",fontSize:10.5,fontWeight:700}}>${fmt(s.raw/selRate)}</b>}
                </span>
              </span>
            ));
          })()}
          <button onClick={()=>{ const tsv=buildTSV(selCellsRef.current); if(tsv&&navigator.clipboard?.writeText){ navigator.clipboard.writeText(tsv).then(()=>{setCopied(true); setTimeout(()=>setCopied(false),1300);}); } }}
            title="คัดลอกค่าที่เลือก (Ctrl+C)"
            style={{marginLeft:6,marginRight:4,display:"flex",alignItems:"center",gap:5,border:"none",cursor:"pointer",borderRadius:8,padding:"6px 12px",
              fontFamily:"system-ui,sans-serif",fontSize:12,fontWeight:600,background:copied?"#065f46":"#334155",color:"#fff"}}>
            {copied ? "✓ คัดลอกแล้ว" : "⧉ คัดลอก"}
          </button>
        </div>
      )}
      {editMode && (undoInfo.u > 0 || undoInfo.r > 0) && (
        <div style={{position:"fixed",left:20,bottom:20,zIndex:95,display:"flex",gap:6,alignItems:"center",
          background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:12,padding:"7px 9px",boxShadow:"0 8px 28px rgba(15,23,42,0.16)"}}>
          <button onClick={undo} disabled={!undoInfo.u} title="ย้อนกลับ (Ctrl+Z)"
            style={{display:"flex",alignItems:"center",gap:6,background:undoInfo.u?T.blue:"#e2e8f0",color:undoInfo.u?"#fff":"#94a3b8",
              border:"none",borderRadius:8,padding:"7px 12px",fontSize:13,fontWeight:600,cursor:undoInfo.u?"pointer":"default"}}>
            ↩︎ ย้อนกลับ
          </button>
          <button onClick={redo} disabled={!undoInfo.r} title="ทำซ้ำ (Ctrl+Shift+Z)"
            style={{background:undoInfo.r?T.blueLight:"transparent",color:undoInfo.r?T.blue:"#cbd5e1",
              border:`1px solid ${undoInfo.r?T.blue:T.cardBorder}`,borderRadius:8,padding:"7px 10px",fontSize:13,fontWeight:600,cursor:undoInfo.r?"pointer":"default"}}>
            ↪︎
          </button>
          {undoInfo.label && (
            <span style={{fontSize:11,color:T.textMuted,maxWidth:170,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",paddingRight:2}}>
              ล่าสุด: {undoInfo.label}
            </span>
          )}
        </div>
      )}
      <ErrorBoundary>
      {screen === "home" && (
        <HomeScreen projects={projects} saveProjects={saveProjects} openProject={openProject}
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
      {screen === "app" && effectiveRole === "qs"          && (
        <QSView {...sharedProps} onExport={() => runExport(() =>
          // ใช้ฟอร์มเดียวกันทั้งเปิด/ปิด USD — ปิด USD ก็แค่ไม่มีคอลัมน์ USD (ฟอร์มเหมือนกัน)
          Promise.resolve(exportQSExcel(activeProject, tenderCosts, additions, extraItems, hiddenAccounts))
        )} />
      )}
      {screen === "app" && effectiveRole === "procurement" && (
        <ProcurementView {...sharedProps} onExport={() => runExport(() =>
          // ใช้ฟอร์มเดียวกันทั้งเปิด/ปิด USD — ปิด USD ก็แค่ไม่มีคอลัมน์ USD (ฟอร์มเหมือนกัน)
          Promise.resolve(exportProcurementExcel(activeProject, poEntries))
        )} />
      )}
      {screen === "app" && effectiveRole === "accounting"  && (
        <AccountingView {...sharedProps} onExport={() => runExport(() => exportAccountingExcel(activeProject, tenderCosts, additions, poEntries, extraItems, hiddenAccounts))} />
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
    if (!username.trim() || !password) { setError("กรอก Username และ Password ให้ครบ"); return; }
    setBusy(true); setError("");
    try {
      const user = await verifyLogin(username, password);
      if (!user) { setError("Username หรือ Password ไม่ถูกต้อง หรือบัญชีถูกระงับ"); setBusy(false); return; }
      onLogin(user);
    } catch (err) {
      setError("เกิดข้อผิดพลาด ลองใหม่อีกครั้ง");
      setBusy(false);
    }
  };

  return (
    <div style={{minHeight:"100vh",background:T.headerGrad,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <form onSubmit={submit} style={{background:T.card,borderRadius:20,padding:36,width:400,maxWidth:"92vw",boxShadow:"0 24px 60px rgba(0,0,0,0.25)"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{fontSize:34,marginBottom:8}}>🏗</div>
          <div style={{fontSize:11,letterSpacing:3,color:T.textMuted,textTransform:"uppercase",fontWeight:600}}>TENDER COST SYSTEM</div>
          <div style={{fontSize:19,fontWeight:800,color:T.textPrimary,marginTop:4}}>เข้าสู่ระบบ</div>
          <div style={{fontSize:12,color:T.textMuted,marginTop:4}}>ล็อกอินตามแผนก: QS · จัดซื้อ · บัญชี · Admin</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <label style={{display:"flex",flexDirection:"column",gap:6}}>
            <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>Username</span>
            <input className="input-base" autoFocus value={username} onChange={e=>setUsername(e.target.value)} placeholder="เช่น qs, procurement, accounting, admin" />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:6}}>
            <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>Password</span>
            <input className="input-base" type="password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" />
          </label>
          {error && <div style={{background:T.redBg,color:T.red,fontSize:12,padding:"9px 12px",borderRadius:8,fontWeight:500}}>{error}</div>}
          <button className="btn-primary" type="submit" disabled={busy} style={{marginTop:6,opacity:busy?0.7:1}}>
            {busy ? "กำลังตรวจสอบ..." : "เข้าสู่ระบบ"}
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
      <td style={{padding:"10px 14px",color:T.textPrimary,fontWeight:600}}>{u.username}{isSelf && <span style={{marginLeft:6,fontSize:10,color:T.textMuted}}>(คุณ)</span>}</td>
      <td style={{padding:"10px 14px",color:T.textSecondary}}>{u.name}</td>
      <td style={{padding:"10px 14px"}}>
        <span style={{background:T.blueLight,color:T.blue,fontSize:11,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{ROLE_LABELS[u.role]}</span>
      </td>
      <td style={{padding:"10px 14px"}}>
        <span style={{background:u.active?T.greenBg:T.redBg,color:u.active?T.green:T.red,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>
          {u.active ? "ใช้งานได้" : "ระงับแล้ว"}
        </span>
      </td>
      <td style={{padding:"10px 14px"}}>
        {resetting ? (
          <div style={{display:"flex",gap:6,alignItems:"center"}}>
            <input className="input-base" type="text" placeholder="รหัสผ่านใหม่" value={pw} onChange={e=>setPw(e.target.value)} style={{width:130,padding:"6px 10px"}} />
            <button className="btn-primary" style={{padding:"6px 12px"}} onClick={()=>{ if(pw){ onReset(u.id,pw); setPw(""); setResetting(false);} }}>บันทึก</button>
            <button className="btn-ghost" style={{padding:"6px 10px"}} onClick={()=>{setResetting(false);setPw("");}}>ยกเลิก</button>
          </div>
        ) : (
          <div style={{display:"flex",gap:8}}>
            <button className="btn-ghost" style={{padding:"6px 12px",fontSize:12}} onClick={()=>setResetting(true)}>รีเซ็ตรหัส</button>
            <button className="btn-ghost" style={{padding:"6px 12px",fontSize:12}} onClick={()=>onToggle(u.id)}>{u.active?"ระงับ":"เปิดใช้"}</button>
            {!isSelf && <button className="btn-ghost" style={{padding:"6px 12px",fontSize:12,color:T.red,borderColor:T.red}} onClick={()=>{if(confirm(`ลบผู้ใช้ ${u.username}?`)) onDelete(u.id);}}>ลบ</button>}
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
  const [selKey, setSelKey]   = useState(null);
  const [dept, setDept]       = useState("all"); // กรองตามแผนก
  const [loading, setLoading] = useState(true);
  const [busy, setBusy]       = useState(false);
  const [msg, setMsg]         = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    const [s, projs] = await Promise.all([loadKvSnapshots(), sg("tcs-projects")]);
    const map = {}; (projs || []).forEach(p => { map[p.id] = p.name; });
    setProjMap(map); setSnaps(s); setLoading(false);
  }, []);
  useEffect(() => { load(); }, [load]);

  const keyLabel = (key) => {
    if (key === "tcs-projects") return "📁 รายชื่อโครงการ";
    const m = key.match(/^tcs-(tenders|po|additions|extra|hidden)-(.+)$/);
    if (m) {
      const t = { tenders:"Tender Cost", po:"PO / จัดซื้อ", additions:"ยอดเพิ่มรายเดือน", extra:"รายการเพิ่ม", hidden:"หมวดที่ซ่อน" }[m[1]] || m[1];
      return `${t} — ${projMap[m[2]] || m[2]}`;
    }
    if (key === "tcs-users") return "ผู้ใช้ (คีย์เก่า)";
    if (key === "tcs-logs")  return "Log (คีย์เก่า)";
    return key;
  };
  const preview = (v) => {
    if (v == null) return "(ว่าง)";
    const s = String(v);
    return s.length > 90 ? s.slice(0, 90) + "…" : s;
  };
  const snapLabel = (r) => `${new Date(r.taken_at).toLocaleDateString("th-TH",{day:"numeric",month:"short"})} · ${r.slot}`;

  // จับคู่คีย์ข้อมูล → แผนกเจ้าของ
  const deptOf = (key) => {
    if (/^tcs-(tenders|additions|extra|hidden|columns)-/.test(key)) return "qs";
    if (/^tcs-po-/.test(key)) return "procurement";
    return "central"; // tcs-projects, tcs-users, tcs-logs, อื่น ๆ
  };
  const DEPTS = [["all","ทั้งหมด"],["qs","QS"],["procurement","จัดซื้อ"],["central","ส่วนกลาง"]];
  const deptTag = { qs:{label:"QS",color:T.blue,bg:T.blueLight}, procurement:{label:"จัดซื้อ",color:T.amber,bg:T.amberBg}, central:{label:"ส่วนกลาง",color:T.purple,bg:T.purpleBg} };

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
  const roundDateLabel = (r) => new Date(r.taken_at).toLocaleDateString("th-TH",{weekday:"short",day:"numeric",month:"short",year:"numeric"});
  const deptLabelOf = (id) => (DEPTS.find(([d])=>d===id)||[])[1] || "ข้อมูล";

  // กู้คืนทั้งชุดของแผนกที่เลือก กลับไปยังรอบเวลาที่กด — ย้อนทุกไฟล์พร้อมกัน
  const doRestoreRound = async (round) => {
    const rows = Object.values(round.byKey);
    const dl = deptLabelOf(dept);
    if (!window.confirm(`กู้คืน "${dl}" ทั้งชุด (${rows.length} รายการ)\nกลับเป็นสแนปช็อต ${roundDateLabel(round)} · ${round.slot}?\n\nข้อมูลปัจจุบันของทุกไฟล์ในชุดนี้จะถูกแทนที่ด้วยข้อมูลจากรอบที่เลือก`)) return;
    setBusy(true); setMsg("");
    let ok = 0, fail = 0;
    for (const row of rows) {
      try { await restoreKvSnapshot(row); ok++; }
      catch (e) { fail++; }
    }
    setMsg(fail === 0
      ? `✅ กู้คืน ${dl} สำเร็จ ${ok} รายการ — กลับไปหน้าหลักเพื่อดูข้อมูลที่กู้คืน`
      : `⚠️ กู้คืนสำเร็จ ${ok} รายการ · ไม่สำเร็จ ${fail} รายการ`);
    await load();
    setBusy(false);
  };

  if (loading) return <div style={{color:T.textMuted,fontSize:13}}>กำลังโหลดสแนปช็อต...</div>;

  if (!snaps.length) return (
    <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:24,fontSize:13,color:T.textSecondary,lineHeight:1.7}}>
      ยังไม่มีสแนปช็อต<br/>
      <span style={{color:T.textMuted}}>ระบบจะถ่ายสแนปช็อตอัตโนมัติวันละ 2 รอบ (12:00 และ 18:00) หลังจากรันไฟล์ <b>kv-snapshots.sql</b> ใน Supabase</span>
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
        สแนปช็อตอัตโนมัติวันละ 2 รอบ — 12:00 และ 18:00 · เลือกแผนก แล้วกด "กู้คืนทั้งชุด" กลับไปยังรอบเวลาที่ต้องการ — ทุกไฟล์ของแผนกนั้นจะย้อนกลับพร้อมกัน
      </div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap"}}>
        {DEPTS.map(([id,label])=>{
          const active = dept===id;
          return (
            <button key={id} onClick={()=>setDept(id)}
              style={{padding:"7px 16px",borderRadius:999,border:`1.5px solid ${active?T.blue:T.cardBorder}`,cursor:"pointer",fontSize:12.5,fontWeight:600,
                background:active?T.blue:T.card,color:active?"#fff":T.textSecondary}}>
              {label} <span style={{opacity:0.7,fontWeight:500}}>({deptCount[id]||0})</span>
            </button>
          );
        })}
      </div>

      {!rounds.length ? (
        <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:24,fontSize:13,color:T.textMuted}}>
          แผนก "{deptLabelOf(dept)}" ยังไม่มีสแนปช็อตให้กู้คืน
        </div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div style={{fontSize:12,fontWeight:700,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>
            รอบสแนปช็อตของ {deptLabelOf(dept)} ({rounds.length} รอบ)
          </div>
          {rounds.map(round => {
            const rows = Object.values(round.byKey);
            return (
              <div key={round.rk} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px"}}>
                <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                  <span style={{flexShrink:0,fontSize:11,fontWeight:700,padding:"3px 11px",borderRadius:8,background:T.blueLight,color:T.blue}}>{round.slot}</span>
                  <div style={{fontSize:14,fontWeight:700,color:T.textPrimary}}>{roundDateLabel(round)}</div>
                  <span style={{fontSize:12,color:T.textMuted}}>· {rows.length} ไฟล์ในชุดนี้</span>
                  <div style={{flex:1,minWidth:12}}/>
                  <button onClick={()=>doRestoreRound(round)} disabled={busy}
                    className="btn-primary" style={{flexShrink:0,padding:"8px 18px",fontSize:13,opacity:busy?0.5:1,cursor:busy?"default":"pointer"}}>
                    ↩︎ กู้คืนทั้งชุด ({rows.length})
                  </button>
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:6,marginTop:12}}>
                  {rows.map(r => {
                    const dt = deptTag[deptOf(r.key)];
                    return (
                      <span key={r.id} style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:11,padding:"3px 9px",borderRadius:7,background:T.bg,border:`1px solid ${T.cardBorder}`,color:T.textSecondary,maxWidth:260}}>
                        <span style={{flexShrink:0,fontSize:9,fontWeight:700,padding:"0 5px",borderRadius:4,background:dt.bg,color:dt.color}}>{dt.label}</span>
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

  const handleReset   = async (id, pw)  => setUsers(await resetPassword(users, id, pw));
  const handleToggle  = async (id)      => setUsers(await toggleActive(users, id));
  const handleDelete  = async (id)      => setUsers(await deleteUser(users, id));
  const handleCreate  = async () => {
    setErr("");
    if (!draft.username.trim() || !draft.password) { setErr("กรอก Username และ Password"); return; }
    try {
      const next = await createUser(draft);
      setUsers(next); setAddOpen(false); setDraft({ username:"", name:"", role:"qs", password:"" });
    } catch (e) { setErr(e.message || "สร้างผู้ใช้ไม่สำเร็จ"); }
  };

  return (
    <div style={{minHeight:"100vh",background:T.bg}}>
      <div style={{background:T.headerGrad,padding:"18px 32px",display:"flex",alignItems:"center",gap:16}}>
        <button onClick={onBack} title="กลับ" style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 14px",fontSize:15,fontWeight:600,display:"flex",alignItems:"center",gap:6}}>← กลับ</button>
        <div>
          <div style={{fontSize:10,letterSpacing:3,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",fontWeight:600}}>TENDER COST SYSTEM</div>
          <div style={{fontSize:16,fontWeight:700,color:"#fff",marginTop:2}}>Admin Panel</div>
        </div>
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:12}}>
          <span style={{fontSize:12,color:"rgba(255,255,255,0.8)"}}>👤 {session.name} ({ROLE_LABELS[session.role]})</span>
          <button onClick={onLogout} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"7px 14px",fontSize:12,fontWeight:600}}>ออกจากระบบ</button>
        </div>
      </div>

      <div style={{padding:"28px 32px"}}>
        <div style={{display:"flex",gap:8,marginBottom:22}}>
          {[["users","👥 จัดการผู้ใช้"],["logs","📜 Log การเข้าใช้งาน"],["restore","🕘 กู้คืนข้อมูล"]].map(([id,label])=>(
            <button key={id} onClick={()=>setTab(id)}
              style={{background:tab===id?T.blue:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.blue:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {label}
            </button>
          ))}
        </div>

        {!loaded ? (
          <div style={{color:T.textMuted,fontSize:13}}>กำลังโหลด...</div>
        ) : tab === "restore" ? (
          <AdminRestoreTab />
        ) : tab === "users" ? (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"16px 18px",borderBottom:`1px solid ${T.cardBorder}`}}>
              <div style={{fontSize:13,fontWeight:600,color:T.textPrimary}}>ผู้ใช้ทั้งหมด ({users.length})</div>
              <button className="btn-primary" onClick={()=>setAddOpen(v=>!v)}>+ เพิ่มผู้ใช้</button>
            </div>
            {addOpen && (
              <div style={{padding:18,borderBottom:`1px solid ${T.cardBorder}`,background:"#fafbfd"}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr auto",gap:10,alignItems:"end"}}>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>Username</span>
                    <input className="input-base" value={draft.username} onChange={e=>setDraft(d=>({...d,username:e.target.value}))} />
                  </label>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>ชื่อที่แสดง</span>
                    <input className="input-base" value={draft.name} onChange={e=>setDraft(d=>({...d,name:e.target.value}))} />
                  </label>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>แผนก</span>
                    <select className="input-base" value={draft.role} onChange={e=>setDraft(d=>({...d,role:e.target.value}))}>
                      <option value="qs">QS</option>
                      <option value="procurement">จัดซื้อ</option>
                      <option value="accounting">บัญชี</option>
                      <option value="admin">Admin</option>
                    </select>
                  </label>
                  <label style={{display:"flex",flexDirection:"column",gap:5}}>
                    <span style={{fontSize:11,color:T.textSecondary}}>Password</span>
                    <input className="input-base" type="text" value={draft.password} onChange={e=>setDraft(d=>({...d,password:e.target.value}))} />
                  </label>
                  <button className="btn-primary" onClick={handleCreate}>สร้าง</button>
                </div>
                {err && <div style={{color:T.red,fontSize:12,marginTop:8,fontWeight:500}}>{err}</div>}
              </div>
            )}
            <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {["Username","ชื่อที่แสดง","แผนก","สถานะ",""].map(h=>(
                    <th key={h} style={{padding:"10px 14px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
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
              ประวัติการเข้าใช้งานล่าสุด ({logs.length})
            </div>
            <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {["เวลา","Username","แผนก","ผลลัพธ์"].map(h=>(
                    <th key={h} style={{padding:"10px 14px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {logs.length===0 ? (
                  <tr><td colSpan={4} style={{padding:"30px",textAlign:"center",color:T.textMuted}}>ยังไม่มีข้อมูล</td></tr>
                ) : logs.map(l => (
                  <tr key={l.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                    <td style={{padding:"9px 14px",fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:T.textSecondary}}>{new Date(l.time).toLocaleString("th-TH")}</td>
                    <td style={{padding:"9px 14px",color:T.textPrimary,fontWeight:500}}>{l.username}</td>
                    <td style={{padding:"9px 14px",color:T.textSecondary}}>{ROLE_LABELS[l.role]||l.role}</td>
                    <td style={{padding:"9px 14px"}}>
                      <span style={{background:l.result==="success"?T.greenBg:T.redBg,color:l.result==="success"?T.green:T.red,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>
                        {l.result==="success"?"สำเร็จ":l.result==="inactive"?"บัญชีถูกระงับ":"ล้มเหลว"}
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
        <div style={{fontSize:13,color:T.textSecondary}}>กำลังโหลด...</div>
      </div>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

// ─── SyncBadge ────────────────────────────────────────────────────────────────
function SyncBadge({ syncing, syncedAt }) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:6,background:"rgba(255,255,255,0.15)",backdropFilter:"blur(8px)",borderRadius:8,padding:"5px 12px",fontSize:11,color:"rgba(255,255,255,0.85)"}}>
      <span style={{width:6,height:6,borderRadius:"50%",background:syncing?"#fbbf24":"#34d399",display:"inline-block",boxShadow:syncing?"0 0 6px #fbbf24":"0 0 6px #34d399",animation:syncing?"pulse 0.8s ease-in-out infinite":"none"}}/>
      {syncing ? "กำลัง sync..." : syncedAt ? `sync ${syncedAt.toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit",second:"2-digit"})}` : ""}
    </div>
  );
}

// ─── Search input with a clear (×) button ──────────────────────────────────
// Small wrapper around the standard .input-base search box used across QS,
// Procurement, and Accounting toolbars — shows an × to instantly clear the
// text once something has been typed, instead of having to select-and-delete.
function SearchInput({ value, onChange, placeholder, width = 240 }) {
  return (
    <div style={{position:"relative",width}}>
      <input value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
        className="input-base" style={{width:"100%",paddingRight:value?30:13}}/>
      {value && (
        <button type="button" onClick={()=>onChange("")} title="ล้างคำค้นหา"
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
  const label = !has ? "ทุกหมวด" : selected.length === 1 ? selected[0] : `${selected.length} หมวด`;
  const rowStyle = (on) => ({ display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",border:"none",
    background: on ? T.blueLight : "transparent", color: on ? color : T.textSecondary, cursor:"pointer",
    padding:"7px 10px", borderRadius:8, fontSize:12.5, fontWeight: on ? 700 : 500 });
  return (
    <div ref={ref} style={{position:"relative",flexShrink:0}}>
      <button onClick={()=>setOpen(o=>!o)} title="กรองตามหมวด (เลือกได้หลายหมวด)"
        style={{display:"flex",alignItems:"center",gap:7,padding:"6px 12px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap",
          border:`1.5px solid ${has?color:T.cardBorder}`, background: has?color:T.card, color: has?"#fff":T.textSecondary}}>
        🏷 {label}
        {has && <span style={{fontSize:11,opacity:0.85}}>({selected.length})</span>}
        <span style={{fontSize:9,opacity:0.8}}>▼</span>
      </button>
      {open && (
        <div style={{position:"absolute",top:"calc(100% + 6px)",left:0,zIndex:60,background:T.card,border:`1px solid ${T.cardBorder}`,
          borderRadius:12,boxShadow:"0 10px 32px rgba(15,23,42,0.18)",padding:6,minWidth:210,maxHeight:340,overflowY:"auto"}}>
          <button onClick={()=>{ onChange([]); }} style={rowStyle(!has)}>
            <span style={{fontSize:13}}>{!has?"◉":"◯"}</span> ทุกหมวด
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
                ✕ ล้างตัวเลือก
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
}

// ─── StatCard ─────────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, color, icon, accent, thb, rate }) {
  // ถ้าใส่ยอดบาท (thb) + อัตราแลกเปลี่ยน (rate = บาท/USD) จะโชว์ ≈ $ ควบคู่ให้
  const usd = (rate && rate > 0 && typeof thb === "number") ? thb / rate : null;
  return (
    <div style={{background:T.card,borderRadius:14,padding:"20px 22px",border:`1px solid ${T.cardBorder}`,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:color,borderRadius:"14px 14px 0 0"}}/>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
        <div style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{label}</div>
        {icon && <div style={{width:34,height:34,borderRadius:10,background:accent||T.blueLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>{icon}</div>}
      </div>
      <div style={{fontSize:22,fontWeight:700,color:T.textPrimary,letterSpacing:"-0.5px",fontFamily:"'JetBrains Mono',monospace"}}>{value}</div>
      {usd != null && <div style={{fontSize:14,color:T.green,fontWeight:700,fontFamily:"'JetBrains Mono',monospace",marginTop:3}}>≈ ${fmt0(usd)}</div>}
      {sub && <div style={{fontSize:11,color:T.textMuted,marginTop:5}}>{sub}</div>}
    </div>
  );
}

// แสดงบรรทัดเป็นดอลลาร์ ($) ใต้ยอดบาทในตาราง — ขนาดราวครึ่งหนึ่งของบาท, ทศนิยม 2 ตำแหน่ง
// คืน null ถ้าไม่ได้เปิดใช้อัตราแลกเปลี่ยน
function usdLine(thb, rate) {
  if (!rate || rate <= 0 || typeof thb !== "number" || !isFinite(thb)) return null;
  return <div className="usd-sub" style={{fontSize:11.5,color:T.green,fontWeight:600,fontFamily:"'JetBrains Mono',monospace",lineHeight:1.2,marginTop:2}}>≈ ${fmt(thb/rate)}</div>;
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
    <div style={{display:"flex",alignItems:"center",gap:8,padding:"5px 10px",border:`1px solid ${T.cardBorder}`,borderRadius:10,background:T.card}}>
      <button onClick={()=>updateProject({ showUsd: !on })} title="เปิด/ปิดการแสดงเป็นดอลลาร์ ($)"
        style={{display:"flex",alignItems:"center",gap:6,border:"none",background:"transparent",cursor:"pointer",padding:0}}>
        <span style={{width:34,height:18,borderRadius:99,background:on?T.green:"#cbd5e1",position:"relative",transition:"all .15s",display:"inline-block",flexShrink:0}}>
          <span style={{position:"absolute",top:2,left:on?18:2,width:14,height:14,borderRadius:99,background:"#fff",transition:"all .15s"}}/>
        </span>
        <span style={{fontSize:12,fontWeight:700,color:on?T.green:T.textMuted}}>USD</span>
      </button>
      <span style={{fontSize:11,color:T.textMuted,whiteSpace:"nowrap"}}>฿/$</span>
      <input type="number" step="any" min="0" value={txt} placeholder="อัตรา"
        onChange={e=>setTxt(e.target.value)} onBlur={commitRate}
        onKeyDown={e=>{ if(e.key==="Enter") e.currentTarget.blur(); }}
        style={{width:64,fontSize:12,padding:"4px 6px",border:`1px solid ${T.cardBorder}`,borderRadius:7,fontFamily:"'JetBrains Mono',monospace",textAlign:"right"}}/>
    </div>
  );
}

// ─── Home Screen ──────────────────────────────────────────────────────────────
function HomeScreen({ projects, saveProjects, openProject, deleteProject, newProjModal, setNewProjModal, syncedAt, syncing, session, onLogout, onOpenAdmin }) {
  const [draft, setDraft] = useState({ name:"", area:"", panels:"", client:"", currency:"THB", usdRate:"" });
  const [projSearch, setProjSearch] = useState("");
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
      {/* Header */}
      <div style={{background:T.headerGrad,padding:"0 32px"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"18px 0 20px",flexWrap:"wrap",gap:12}}>
          <div>
            <div style={{fontSize:11,letterSpacing:3,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",fontWeight:600,marginBottom:4}}>TENDER COST SYSTEM</div>
            <div style={{fontSize:22,fontWeight:800,color:"#fff",letterSpacing:"-0.5px"}}>ระบบบริหารต้นทุนโครงการ</div>
            <div style={{fontSize:12,color:"rgba(255,255,255,0.6)",marginTop:2}}>QS · จัดซื้อ · บัญชี — Real-time sync</div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <SyncBadge syncing={syncing} syncedAt={syncedAt}/>
            {session?.role === "admin" && (
              <button className="btn-primary" onClick={onOpenAdmin}
                style={{background:"rgba(255,255,255,0.2)",backdropFilter:"blur(8px)",border:"1.5px solid rgba(255,255,255,0.3)"}}>
                ⚙️ Admin
              </button>
            )}
            {session?.role !== "accounting" && (
              <button className="btn-primary" onClick={()=>setNewProjModal(true)}
                style={{background:"rgba(255,255,255,0.2)",backdropFilter:"blur(8px)",border:"1.5px solid rgba(255,255,255,0.3)",display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:16,lineHeight:1}}>+</span> โครงการใหม่
              </button>
            )}
            <div style={{width:1,alignSelf:"stretch",background:"rgba(255,255,255,0.2)"}}/>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:12,color:"#fff",fontWeight:600}}>{session?.name}</div>
              <div style={{fontSize:10,color:"rgba(255,255,255,0.6)"}}>{ROLE_LABELS[session?.role]}</div>
            </div>
            <button onClick={onLogout} title="ออกจากระบบ"
              style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"8px 12px",fontSize:12,fontWeight:600}}>
              ออกจากระบบ
            </button>
          </div>
        </div>
        {/* Summary row */}
        <div style={{display:"flex",gap:24,paddingBottom:20}}>
          {[
            {label:"โครงการทั้งหมด",value:projects.length,icon:"🏗"},
            {label:"Active Projects",value:projects.length,icon:"📊"},
          ].map(s=>(
            <div key={s.label} style={{display:"flex",alignItems:"center",gap:8,background:"rgba(255,255,255,0.12)",borderRadius:10,padding:"8px 16px"}}>
              <span style={{fontSize:16}}>{s.icon}</span>
              <span style={{fontSize:13,color:"rgba(255,255,255,0.85)",fontWeight:600}}>{s.value} {s.label}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Body */}
      <div style={{padding:"28px 32px"}}>
        {projects.length === 0 ? (
          <div style={{textAlign:"center",padding:"80px 0",color:T.textMuted}}>
            <div style={{fontSize:52,marginBottom:14}}>🏗</div>
            <div style={{fontSize:17,fontWeight:600,color:T.textSecondary,marginBottom:8}}>ยังไม่มีโครงการ</div>
            <div style={{fontSize:13,marginBottom:20}}>กด "โครงการใหม่" เพื่อเริ่มต้น</div>
            <button className="btn-primary" onClick={()=>setNewProjModal(true)}>+ สร้างโครงการแรก</button>
          </div>
        ) : (
          <>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18,flexWrap:"wrap"}}>
              <input value={projSearch} onChange={e=>setProjSearch(e.target.value)} placeholder="🔍 ค้นหาโครงการ / ชื่อลูกค้า…"
                style={{flex:1,minWidth:220,maxWidth:360,padding:"9px 14px",border:`1px solid ${T.cardBorder}`,borderRadius:10,fontSize:13,outline:"none"}}/>
              <span style={{fontSize:12,color:T.textMuted,fontWeight:500}}>
                {projSearch.trim() ? `พบ ${shownProjects.length} จาก ${projects.length} โครงการ` : `${projects.length} โครงการทั้งหมด`}
              </span>
            </div>
            {shownProjects.length === 0 ? (
              <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>ไม่พบโครงการที่ตรงกับ "{projSearch}"</div>
            ) : (
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:20}}>
                {shownProjects.map(p => <ProjectCard key={p.id} project={p} onOpen={()=>openProject(p.id)} onDelete={session?.role==="accounting" ? null : ()=>deleteProject(p.id)} />)}
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
                <div style={{fontSize:16,fontWeight:700,color:T.textPrimary}}>สร้างโครงการใหม่</div>
                <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>กรอกข้อมูลโครงการเพื่อเริ่มต้น</div>
              </div>
              <button onClick={()=>setNewProjModal(false)} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              {[
                ["ชื่อโครงการ *","name","text","1/-1"],
                ["ลูกค้า / Client","client","text","1/-1"],
                ["พื้นที่รวม (ft²)","area","number","auto"],
                ["จำนวน Panels","panels","number","auto"],
                ["สกุลเงิน","currency","text","auto"],
                ["อัตราแลกเปลี่ยน (บาท/USD)","usdRate","number","auto"],
              ].map(([label,key,type,col]) => (
                <label key={key} style={{display:"flex",flexDirection:"column",gap:6,gridColumn:col}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{label}</span>
                  <input type={type} step={type==="number"?"any":undefined} value={draft[key]} onChange={e=>setDraft(d=>({...d,[key]:e.target.value}))} className="input-base"/>
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:24}}>
              <button onClick={createProject} disabled={!draft.name.trim()} className="btn-primary" style={{opacity:draft.name.trim()?1:0.5}}>สร้างโครงการ</button>
              <button onClick={()=>setNewProjModal(false)} className="btn-ghost">ยกเลิก</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function ProjectCard({ project, onOpen, onDelete }) {
  const ageRaw = Math.floor((Date.now() - new Date(project.createdAt)) / 86400000);
  const age = Number.isFinite(ageRaw) && ageRaw >= 0 ? ageRaw : null;
  return (
    <div className="card-hover" onClick={onOpen} title="เปิดโครงการ"
      style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:24,cursor:"pointer",position:"relative"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:T.headerGrad,borderRadius:"16px 16px 0 0"}}/>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14,paddingTop:2}}>
        <span style={{fontSize:10,letterSpacing:2,color:T.blue,fontWeight:700,textTransform:"uppercase"}}>PROJECT</span>
        {onDelete && (
          <button onClick={e=>{e.stopPropagation();onDelete();}} style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",fontSize:14,padding:4,borderRadius:6,transition:"color 0.15s"}}
            onMouseEnter={e=>e.target.style.color="#ef4444"} onMouseLeave={e=>e.target.style.color=T.textMuted}>🗑</button>
        )}
      </div>
      <div style={{fontSize:18,fontWeight:700,color:T.textPrimary,marginBottom:4,lineHeight:1.3}}>{project.name}</div>
      {project.client && <div style={{fontSize:12,color:T.textSecondary,marginBottom:14}}>{project.client}</div>}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
        {project.area   && <span style={{background:T.blueLight,color:T.blue,fontSize:11,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.area} ft²</span>}
        {project.panels && <span style={{background:T.blueLight,color:T.blue,fontSize:11,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.panels} Panels</span>}
        {project.currency && <span style={{background:"#f8fafc",color:T.textMuted,fontSize:11,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.currency}</span>}
      </div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{fontSize:11,color:T.textMuted}}>{age === null ? "—" : age === 0 ? "สร้างวันนี้" : `${age} วันที่แล้ว`}</div>
        <button onClick={e=>{e.stopPropagation();onOpen();}} className="btn-primary" style={{padding:"8px 18px",fontSize:12}}>เปิดโครงการ →</button>
      </div>
    </div>
  );
}

// ─── Role Select ──────────────────────────────────────────────────────────────
function RoleSelect({ project, updateProject, onSelect, onBack }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(project);
  useEffect(() => setDraft(project), [project]);

  const ROLES = [
    {id:"qs",label:"QS",sub:"Quantity Surveyor",desc:"ลงราคา Tender Cost\nประมาณการต้นทุนโครงการ",color:T.blue,bg:T.blueLight,icon:"📐"},
    {id:"procurement",label:"จัดซื้อ",sub:"Procurement",desc:"ลงราคาจริงที่ซื้อ + วันที่\nออก PO และติดตามสถานะ",color:"#d97706",bg:"#fffbeb",icon:"📦"},
    {id:"accounting",label:"บัญชี",sub:"Accounting",desc:"Dashboard ต้นทุน\nBudget vs Actual + Export Excel",color:T.green,bg:T.greenBg,icon:"📊"},
  ];

  return (
    <div style={{minHeight:"100vh",background:T.bg}}>
      <div style={{background:T.headerGrad,padding:"18px 32px",display:"flex",alignItems:"center",gap:16}}>
        <button onClick={onBack} title="กลับ" style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 14px",fontSize:15,fontWeight:600,display:"flex",alignItems:"center",gap:6}}>← กลับ</button>
        <div>
          <div style={{fontSize:10,letterSpacing:3,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",fontWeight:600}}>TENDER COST SYSTEM</div>
          <div style={{fontSize:16,fontWeight:700,color:"#fff",marginTop:2}}>{project.name}</div>
        </div>
        {project.area && (
          <div style={{marginLeft:"auto",display:"flex",gap:8}}>
            {[`${project.area} ft²`,`${project.panels} Panels`].map(v=>(
              <span key={v} style={{background:"rgba(255,255,255,0.15)",color:"rgba(255,255,255,0.9)",fontSize:12,padding:"4px 12px",borderRadius:8,fontWeight:500}}>{v}</span>
            ))}
          </div>
        )}
      </div>

      <div style={{padding:"32px"}}>
        {editing ? (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:24,marginBottom:28,maxWidth:640}}>
            <div style={{fontSize:14,fontWeight:600,color:T.textPrimary,marginBottom:16}}>แก้ไขข้อมูลโครงการ</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
              {[["ชื่อโครงการ","name","text"],["ลูกค้า","client","text"],["สกุลเงิน","currency","text"],["อัตราแลกเปลี่ยน (บาท/USD)","usdRate","number"],["พื้นที่ (ft²)","area","number"],["Panels","panels","number"]].map(([l,k,t]) => (
                <label key={k} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{l}</span>
                  <input type={t} step={t==="number"?"any":undefined} value={draft[k]||""} onChange={e=>setDraft(d=>({...d,[k]:e.target.value}))} className="input-base"/>
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:16}}>
              <button className="btn-primary" onClick={()=>{updateProject(draft);setEditing(false);}}>บันทึก</button>
              <button className="btn-ghost" onClick={()=>setEditing(false)}>ยกเลิก</button>
            </div>
          </div>
        ) : (
          <button onClick={()=>setEditing(true)} className="btn-ghost" style={{marginBottom:24,fontSize:12}}>✏️ แก้ไขข้อมูลโครงการ</button>
        )}

        <div style={{fontSize:13,color:T.textSecondary,marginBottom:20,fontWeight:500}}>เลือก Role การทำงาน</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:20,maxWidth:800}}>
          {ROLES.map(r => (
            <button key={r.id} onClick={()=>onSelect(r.id)} className="card-hover"
              style={{background:T.card,border:`1.5px solid ${T.cardBorder}`,borderRadius:18,padding:"28px 24px",cursor:"pointer",textAlign:"left",display:"flex",flexDirection:"column",gap:12,position:"relative",overflow:"hidden"}}>
              <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:r.color}}/>
              <div style={{width:44,height:44,borderRadius:12,background:r.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>{r.icon}</div>
              <div>
                <div style={{fontSize:20,fontWeight:700,color:r.color}}>{r.label}</div>
                <div style={{fontSize:11,color:T.textMuted,marginTop:2,letterSpacing:0.5}}>{r.sub}</div>
              </div>
              <p style={{margin:0,fontSize:12,color:T.textSecondary,lineHeight:1.7,whiteSpace:"pre-line"}}>{r.desc}</p>
              <div style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:r.color,fontWeight:600,marginTop:4}}>เข้าใช้งาน <span>→</span></div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── Shell ────────────────────────────────────────────────────────────────────
function Shell({ role, color, project, onBack, onHome, children, syncedAt, syncing, session, onLogout }) {
  const labels = {qs:"QS · Quantity Surveyor",procurement:"จัดซื้อ · Procurement",accounting:"บัญชี · Accounting"};
  const gradients = {
    qs:          "linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)",
    procurement: "linear-gradient(135deg, #78350f 0%, #d97706 100%)",
    accounting:  "linear-gradient(135deg, #064e3b 0%, #10b981 100%)",
  };
  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",flexDirection:"column"}}>
      <div style={{background:gradients[role],padding:"14px 28px",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
        <button onClick={onBack} title="กลับหน้าก่อนหน้า" style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 14px",fontSize:15,fontWeight:600,display:"flex",alignItems:"center",gap:6}}>← กลับ</button>
        {onHome && (
          <button onClick={onHome} title="ไปหน้าเลือกโครงการ" style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 14px",fontSize:15,fontWeight:600,display:"flex",alignItems:"center",gap:6}}>🏠 หน้าโครงการ</button>
        )}
        <div style={{flex:1,minWidth:140}}>
          <div style={{fontSize:10,letterSpacing:3,color:"rgba(255,255,255,0.6)",textTransform:"uppercase",fontWeight:600}}>{labels[role]}</div>
          <div style={{fontSize:14,fontWeight:600,color:"#fff",marginTop:1}}>{project.name}</div>
        </div>
        <SyncBadge syncing={syncing} syncedAt={syncedAt}/>
        {project.area && (
          <div style={{display:"flex",gap:8}}>
            {[`${project.area} ft²`,`${project.panels} Panels`].map(v=>(
              <span key={v} style={{background:"rgba(255,255,255,0.15)",color:"rgba(255,255,255,0.85)",fontSize:11,padding:"3px 10px",borderRadius:6}}>{v}</span>
            ))}
          </div>
        )}
        {session && (
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:11,color:"rgba(255,255,255,0.8)"}}>👤 {session.name}</span>
            <button onClick={onLogout} title="ออกจากระบบ"
              style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 11px",fontSize:11,fontWeight:600}}>
              ออกจากระบบ
            </button>
          </div>
        )}
      </div>
      <div style={{flex:1,overflow:"auto"}}>{children}</div>
    </div>
  );
}

// ─── QS View ─────────────────────────────────────────────────────────────────
function QSView({ project, updateProject, tenderCosts, saveTenders, additions, saveAdditions, extraItems, saveExtraItems, hiddenAccounts, saveHiddenAccounts, onBack, onHome, syncedAt, syncing, session, onLogout, onExport, setEditMode }) {
  const [tab, setTab] = useState("baseline"); // "baseline" | "monthly"
  const [tabHist, setTabHist] = useState([]);  // ประวัติแท็บ — ปุ่มกลับย้อนทีละหน้า
  const goTab   = (id) => { if (id !== tab) { setTabHist(h => [...h, tab]); setTab(id); } };
  const backTab = () => { if (tabHist.length) { const h = [...tabHist]; const p = h.pop(); setTabHist(h); setTab(p); } else onBack(); };
  const usdRate = effRate(project);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
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
  const handleAddExtraItem = ({ name, group, parentCode, code, addedInMonth }) => {
    if (!name.trim()) return;
    const item = parentCode
      ? { code:`EX-${uid()}`, name:name.trim(), parentCode, ...(addedInMonth ? { addedInMonth } : {}) }
      : { code: code || `EX-${uid()}`, name:name.trim(), group };
    saveExtraItems([...extraItems, item]);
    return item.code;
  };

  const handleDeleteExtraItem = (code) => {
    if (!confirm("ลบรายการนี้? ยอดเงินทุกส่วนของรายการนี้ (ราคาเดิม + รายเดือนทุกเดือน) จะถูกลบด้วย")) return;
    saveExtraItems(extraItems.filter(e => e.code !== code));
    const nextTenders = { ...tenderCosts }; delete nextTenders[code];
    saveTenders(nextTenders);
    const nextAdd = {};
    Object.entries(additions).forEach(([m, obj]) => { const o = {...obj}; delete o[code]; nextAdd[m] = o; });
    saveAdditions(nextAdd);
  };

  // Hide / restore a fixed Acc. Code (511010 ... etc). Hiding doesn't erase its stored
  // numbers — it's reversible — it just removes it from the QS entry list and from
  // downstream totals, in case a project doesn't use that code at all.
  const handleHideAccount = (code) => {
    if (!confirm("นำ Acc. Code นี้ออกจากรายการหลัก? (กู้คืนได้ภายหลัง ตัวเลขที่เคยกรอกไว้จะยังไม่หาย)")) return;
    saveHiddenAccounts([...hiddenAccounts, code]);
  };
  const handleRestoreAccount = (code) => saveHiddenAccounts(hiddenAccounts.filter(c => c !== code));

  return (
    <Shell role="qs" color={T.blue} project={project} onBack={backTab} onHome={onHome} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"20px 28px 0"}}>
        <div style={{display:"flex",gap:8,marginBottom:20,alignItems:"center"}}>
          {[["baseline","📐 ราคาเดิม (Baseline)"],["monthly","📅 รายการเพิ่มรายเดือน"]].map(([id,label])=>(
            <button key={id} onClick={()=>goTab(id)}
              style={{background:tab===id?T.blue:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.blue:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {label}
            </button>
          ))}
          <div style={{marginLeft:"auto"}}><CurrencyControl project={project} updateProject={updateProject}/></div>
          {tab==="monthly" && (
            <button onClick={()=>monthlyExportRef.current && monthlyExportRef.current()} className="btn-ghost"
              style={{display:"flex",alignItems:"center",gap:6,borderColor:T.green,color:T.green}}>
              ⬇️ Export เดือนนี้
            </button>
          )}
          <button onClick={onExport} className="btn-ghost" style={{display:"flex",alignItems:"center",gap:6,borderColor:T.blue,color:T.blue}}>
            ⬇️ Export Excel
          </button>
        </div>
      </div>
      {tab === "baseline"
        ? <QSBaselineTab project={project} tenderCosts={tenderCosts} saveTenders={saveTenders} extraItems={extraItems}
                         onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} onHideAccount={handleHideAccount} onRestoreAccount={handleRestoreAccount} setEditMode={setEditMode} />
        : <QSMonthlyTab tenderCosts={tenderCosts} additions={additions} saveAdditions={saveAdditions}
                         extraItems={extraItems} onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} setEditMode={setEditMode} project={project} registerMonthExport={registerMonthExport} />}
    </Shell>
  );
}

// ─── QS Tab 1: Baseline (original tender cost) ────────────────────────────────
function QSBaselineTab({ project, tenderCosts, saveTenders, extraItems, onAddExtra, onDeleteExtra, hiddenAccounts, onHideAccount, onRestoreAccount, setEditMode }) {
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

  useEffect(() => setDraft({...tenderCosts}), [tenderCosts]);

  // The baseline counts as "saved" (and therefore locked, requiring "แก้ไข"
  // to unlock) once it carries the explicit $saved flag, or — for baselines
  // saved before this flag existed — once it already has any real value.
  const hasData = Object.entries(tenderCosts||{}).some(([k,v]) => !k.startsWith("$") && parseFloat(v));
  const baselineSaved = tenderCosts.$saved === true || hasData;
  const editingUnlocked = !baselineSaved || forceEdit;
  useEffect(() => { setEditMode?.(editingUnlocked); return () => setEditMode?.(false); }, [editingUnlocked, setEditMode]);

  // ยกเลิกการแก้ไข: ทิ้งค่าที่พิมพ์ค้าง คืนกลับเป็นค่าที่บันทึกไว้ล่าสุด แล้วล็อกกลับ
  const canCancel = baselineSaved; // มีค่าที่บันทึกไว้ให้ย้อนกลับได้
  const handleCancel = () => { setDraft({ ...tenderCosts }); setForceEdit(false); setAddOpen(false); setSubFor(null); };
  useEffect(() => {
    if (!editingUnlocked) return;
    const onEsc = (e) => { if (e.key === "Escape" && canCancel) { e.preventDefault(); handleCancel(); } };
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
  const adj3  = base * 0.03;
  const total = base + adj3;

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
      if (taken) { alert(`Acc. Code "${code}" มีอยู่แล้ว กรุณาใช้รหัสอื่น`); return; }
    }
    onAddExtra({ name:addDraft.name, group:addDraft.group, code: code || undefined });
    setAddDraft({ code:"", name:"", group:GROUPS[0] }); setAddOpen(false);
  };

  const handleAddSub = (parentCode) => {
    if (!subName.trim()) return;
    onAddExtra({ name:subName, parentCode });
    setCollapsed(c => ({...c, [parentCode]: false})); // reveal the newly-added sub-item
    setSubName(""); setSubFor(null);
  };

  const handleDeleteRow = (code) => {
    onDeleteExtra(code);
    setDraft(d => { const n = {...d}; delete n[code]; return n; });
  };

  return (
    <div style={{padding:"4px 28px 24px"}}>
      {/* Stats */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:24}}>
        <StatCard label="ราคาเดิมรวม (Tender Cost)" value={"฿"+fmt0(base)} thb={base} rate={usdRate} sub="ราคาเดิมทั้งหมด — ใช้เป็นงบตั้งต้นจริง" color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label="เผื่อเศษ/สูญเสีย 3%" value={"฿"+fmt0(adj3)} thb={adj3} rate={usdRate} sub="ตัวเลขอ้างอิงเท่านั้น (ไม่รวมในงบ)" color={T.amber} icon="⚙️" accent={T.amberBg}/>
        <StatCard label="รวมเผื่อ 3% (อ้างอิง)" value={"฿"+fmt0(total)} thb={total} rate={usdRate} sub="ประมาณการเผื่อเศษ — งบจริงใช้ราคาเดิม" color={T.green} icon="✅" accent={T.greenBg}/>
      </div>

      {/* Filters + Add row + Save */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <SearchInput value={search} onChange={setSearch} placeholder="🔍 ค้นหา Account Code / ชื่อ..." width={240}/>
        <button onClick={()=>setHideEmpty(v=>!v)}
          title="ซ่อน/แสดงแถวที่ไม่มีค่า (ราคาเดิม = 0)"
          style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:8,fontSize:11.5,fontWeight:600,cursor:"pointer",
            border:`1.5px solid ${hideEmpty?T.blue:T.cardBorder}`,background:hideEmpty?T.blue:T.card,color:hideEmpty?"#fff":T.textSecondary,whiteSpace:"nowrap"}}>
          {hideEmpty ? `✓ เฉพาะที่มีค่า${hiddenEmptyCount?` (ซ่อน ${hiddenEmptyCount})`:""}` : "⚡ เฉพาะที่มีค่า"}
        </button>
        <GroupFilter selected={filter} onChange={setFilter}/>
        <div style={{flex:1}}/>
        {hiddenList.length > 0 && (
          <button className="btn-ghost" onClick={()=>setShowHidden(v=>!v)} style={{color:T.textMuted}}>
            🗂 ที่ซ่อนไว้ ({hiddenList.length})
          </button>
        )}
        <button className="btn-ghost" onClick={()=>setAddOpen(v=>!v)} disabled={!editingUnlocked}
          style={!editingUnlocked?{opacity:0.4,cursor:"not-allowed"}:undefined}>+ เพิ่มรายการหลักใหม่</button>
        {!editingUnlocked && (
          <span style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:T.textMuted,background:"#f1f5f9",padding:"6px 12px",borderRadius:8,fontWeight:600}}>
            🔒 บันทึกแล้ว
          </span>
        )}
        {editingUnlocked ? (
          <>
            <button onClick={handleSave} className="btn-primary"
              style={{background:saved?T.green:T.blue,minWidth:140}}>
              {saved?"✓ บันทึกแล้ว":"บันทึก Tender Cost"}
            </button>
            {canCancel && (
              <button onClick={handleCancel} className="btn-ghost" title="ยกเลิกการแก้ไข (Esc)"
                style={{color:T.red,borderColor:T.red}}>✕ ยกเลิก</button>
            )}
          </>
        ) : (
          <button onClick={()=>setForceEdit(true)} className="btn-primary" style={{background:T.amber,minWidth:140}}>
            ✏️ แก้ไข Tender Cost
          </button>
        )}
      </div>

      {/* Hidden accounts panel */}
      {showHidden && hiddenList.length > 0 && (
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:14,marginBottom:16}}>
          <div style={{fontSize:11,color:T.textSecondary,marginBottom:8}}>Acc. Code ที่ซ่อนไว้ — ตัวเลขที่เคยกรอกยังอยู่ กด "กู้คืน" เพื่อนำกลับมาแสดง</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
            {hiddenList.map(a=>(
              <div key={a.code} style={{display:"flex",alignItems:"center",gap:8,background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:8,padding:"6px 10px"}}>
                <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:11,color:T.textMuted}}>{a.code}</span>
                <span style={{fontSize:12,color:T.textPrimary}}>{a.name}</span>
                <button onClick={()=>onRestoreAccount(a.code)} className="btn-ghost" style={{padding:"3px 9px",fontSize:11}}>↺ กู้คืน</button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Inline "add standalone row" form */}
      {addOpen && (
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"1fr 2fr 1fr auto",gap:10,alignItems:"end"}}>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>Acc. Code (ถ้ามี)</span>
            <input className="input-base" value={addDraft.code} onChange={e=>setAddDraft(d=>({...d,code:e.target.value}))}
              placeholder="เช่น 511099" style={{fontFamily:"'JetBrains Mono',monospace"}}
              onKeyDown={e=>e.key==="Enter"&&handleAddRow()} />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>ชื่อรายการใหม่ (งานที่ไม่มี Acc. Code เดิมรองรับ)</span>
            <input className="input-base" value={addDraft.name} onChange={e=>setAddDraft(d=>({...d,name:e.target.value}))}
              placeholder="พิมพ์ชื่อรายการที่ต้องการเพิ่ม" onKeyDown={e=>e.key==="Enter"&&handleAddRow()} autoFocus />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>Group</span>
            <select className="input-base" value={addDraft.group} onChange={e=>setAddDraft(d=>({...d,group:e.target.value}))}>
              {GROUPS.map(g => <option key={g} value={g}>{g}</option>)}
            </select>
          </label>
          <button className="btn-primary" onClick={handleAddRow}>+ เพิ่ม</button>
        </div>
      )}

      {/* Table */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
        <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
          <thead>
            <tr style={{background:"#f8fafc"}}>
              {[
                {label:"Acc. Code", key:"code"},
                {label:"Group", key:"group"},
                {label:"Account Name", key:"name"},
                {label:"Tender Cost (THB)", key:"value"},
                {label:"", key:null},
              ].map(({label,key})=>(
                <th key={label||"__actions"}
                  style={{padding:"11px 16px",textAlign:label.includes("Cost")?"right":"left",color:sortKey===key?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
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
                    <td style={{padding:"10px 16px",color:a.isExtra?T.amber:T.blue,fontFamily:"'JetBrains Mono',monospace",fontSize:12,fontWeight:500}}>
                      {hasKids && (
                        <span title={isCollapsed?"ขยายรายการย่อย":"ย่อรายการย่อย"}
                          style={{color:T.textMuted,fontSize:10,marginRight:6,verticalAlign:"middle",display:"inline-block"}}>
                          {isCollapsed?"▸":"▾"}
                        </span>
                      )}
                      {a.isExtra ? (a.code.startsWith("EX-") ? "—" : a.code) : a.code}
                    </td>
                    <td style={{padding:"10px 16px"}}>
                      <span style={{background:T.blueLight,color:T.blue,fontSize:10,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{a.group}</span>
                    </td>
                    <td style={{padding:"10px 16px",color:T.textPrimary}}>
                      {a.name}
                      {a.isExtra && <span style={{marginLeft:7,fontSize:10,background:T.amberBg,color:T.amber,padding:"1px 8px",borderRadius:6,fontWeight:600}}>รายการใหม่</span>}
                      {hasKids && <span style={{marginLeft:7,fontSize:10,background:T.greenBg,color:T.green,padding:"1px 8px",borderRadius:6,fontWeight:600}}>{kids.length} รายการย่อย</span>}
                      {!a.isExtra && editingUnlocked && (
                        <button onClick={(e)=>{e.stopPropagation(); setSubFor(subFor===a.code?null:a.code); setSubName(""); setCollapsed(c=>({...c,[a.code]:false}));}} title="เพิ่มรายการย่อยใต้ Acc. Code นี้"
                          style={{marginLeft:9,background:"none",border:`1px dashed ${T.cardBorder}`,borderRadius:6,color:T.textMuted,cursor:"pointer",fontSize:10,padding:"1px 7px"}}>
                          + รายการย่อย
                        </button>
                      )}
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right"}}>
                      {hasKids ? (
                        <div style={{width:160,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:T.blueLight,borderRadius:8,color:T.blue,fontWeight:700,fontSize:13}}>
                          {fmt(rowVal)}
                          {usdLine(rowVal, usdRate)}
                        </div>
                      ) : editingUnlocked ? (
                        <MoneyInput value={draft[a.code]??""} onChange={v=>setDraft(d=>({...d,[a.code]:v}))}
                          style={{width:160,background:(parseFloat(draft[a.code])||0)>0?T.blueLight:T.bg}}/>
                      ) : (
                        <div style={{width:160,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:13,color:draft[a.code]>0?T.textPrimary:T.textMuted}}>{fmt(rowVal)}{usdLine(rowVal, usdRate)}</div>
                      )}
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"center"}}>
                      {editingUnlocked && (a.isExtra
                        ? <button onClick={(e)=>{e.stopPropagation(); handleDeleteRow(a.code);}} title="ลบรายการนี้"
                            style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:14}}>✕</button>
                        : <button onClick={(e)=>{e.stopPropagation(); onHideAccount(a.code);}} title="นำ Acc. Code นี้ออกจากรายการหลัก (กู้คืนได้)"
                            style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",fontSize:14}}>✕</button>)}
                    </td>
                  </tr>

                  {/* Sub-items — roll up into the parent Acc. Code's total above */}
                  {!isCollapsed && kids.map((k,ki)=>(
                    <tr key={k.code} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(ki===kids.length-1 && subFor!==a.code)?"1px solid #f1f5f9":"none"}}>
                      <td style={{padding:"6px 16px 6px 30px",color:T.green,fontSize:12}}>↳</td>
                      <td/>
                      <td style={{padding:"6px 16px",color:T.green,fontSize:12,fontStyle:"italic"}}>
                        {k.name}
                        {k.addedInMonth && (
                          <span title="เพิ่มเข้ามาระหว่างทาง ไม่ได้มีมาตั้งแต่ต้น" style={{marginLeft:7,fontSize:9.5,background:T.amberBg,color:T.amber,padding:"1px 7px",borderRadius:6,fontWeight:600,fontStyle:"normal"}}>
                            เพิ่มเมื่อ {monthShortLabel(k.addedInMonth)}
                          </span>
                        )}
                      </td>
                      <td style={{padding:"6px 16px",textAlign:"right"}}>
                        {editingUnlocked ? (
                          <MoneyInput value={draft[k.code]??""} onChange={v=>setDraft(d=>({...d,[k.code]:v}))}
                            style={{width:160,fontSize:12,background:(parseFloat(draft[k.code])||0)>0?T.greenBg:T.bg}}/>
                        ) : (
                          <div style={{width:160,marginLeft:"auto",padding:"6px 8px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:draft[k.code]>0?T.textPrimary:T.textMuted}}>{fmt(parseFloat(draft[k.code])||0)}{usdLine(parseFloat(draft[k.code])||0, usdRate)}</div>
                        )}
                      </td>
                      <td style={{padding:"6px 16px",textAlign:"center"}}>
                        {editingUnlocked && (
                          <button onClick={()=>handleDeleteRow(k.code)} title="ลบรายการย่อยนี้"
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
                          placeholder="ชื่อรายการย่อย เช่น Silicone Structure" style={{width:"100%",fontSize:12}}
                          onKeyDown={e=>e.key==="Enter"&&handleAddSub(a.code)} autoFocus />
                      </td>
                      <td colSpan={2} style={{padding:"7px 16px",display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button className="btn-primary" style={{padding:"5px 12px",fontSize:12}} onClick={()=>handleAddSub(a.code)}>+ เพิ่ม</button>
                        <button className="btn-ghost" style={{padding:"5px 12px",fontSize:12}} onClick={()=>setSubFor(null)}>ยกเลิก</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
              <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:12}}>{filtered.length} รายการ</td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:14}}>
                {fmt(filtered.reduce((s,a)=>s+effectiveValue(a),0))}
                {usdLine(filtered.reduce((s,a)=>s+effectiveValue(a),0), usdRate)}
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
function QSMonthlyTab({ tenderCosts, additions, saveAdditions, extraItems, onAddExtra, onDeleteExtra, hiddenAccounts, setEditMode, project, registerMonthExport }) {
  const usdRate = effRate(project);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const thisMonth = new Date().toISOString().slice(0,7);
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

  useEffect(() => { setDraftAdd({...(additions[month]||{})}); }, [month, additions]);
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
    const onEsc = (e) => { if (e.key === "Escape") { e.preventDefault(); handleCancel(); } };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [editingUnlocked, monthEditMode, month, additions]);

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
  const baseTotal = allRows.reduce((s,r) => s + (parseFloat(tenderCosts[r.code]) || 0), 0);
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
    { label:"เริ่มต้น", cumulative: baseTotal, previous: baseTotal, added: 0 },
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
  const prevMonthLabel   = priorMonths.length ? monthShortLabel(priorMonths[priorMonths.length-1]) : "เริ่มต้น";
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
  const cumBeforeOf = (r) => months.filter(m=>m<month).reduce((s,m)=>s+rowMonthValue(r.code, m),0) + (parseFloat(tenderCosts[r.code])||0);
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

  const handleAddMonth = () => {
    if (!newMonth) return;
    if (months.includes(newMonth)) {
      // ห้ามซ้ำ — ถ้ามีเดือนนี้อยู่แล้ว แค่กระโดดไปที่เดือนนั้นแทนการสร้างซ้ำ
      alert(`มีเดือน ${monthShortLabel(newMonth)} อยู่แล้ว`);
      setMonth(newMonth); setNewMonth("");
      return;
    }
    // ถ้าเดือนล่าสุดมีคอลัมน์อยู่ ให้ถามก่อนว่าจะคัดลอกมาที่เดือนใหม่ไหม
    const prevMonth = months.length ? months[months.length - 1] : null;
    const prevCols = prevMonth ? columnsOf(prevMonth) : [];
    const monthObj = {};
    if (prevCols.length) {
      if (window.confirm(`คัดลอกคอลัมน์จากเดือน ${monthShortLabel(prevMonth)} มาที่เดือนใหม่ไหม?\n(${prevCols.map(c=>c.name).join(", ")})\n\nOK = คัดลอกคอลัมน์ (ยอดเริ่มที่ว่าง) · Cancel = เริ่มเดือนใหม่แบบไม่มีคอลัมน์`)) {
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
  const handleDeleteMonth = (m) => {
    if (!window.confirm(`ลบเดือน ${monthShortLabel(m)} และข้อมูลที่เพิ่มในเดือนนี้ทั้งหมด?\n(ราคาเดิม/Baseline ไม่ได้รับผลกระทบ)`)) return;
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
    saveAdditions({ ...additions, [month]: clean });
    setForceEdit(false); setAddExtraOpen(false); setSubFor(null); setAddColOpen(false);
    setSaved(true); setTimeout(()=>setSaved(false),2000);
  };

  const handleCreateExtra = () => {
    if (!extraDraft.name.trim()) return;
    const code = extraDraft.code.trim();
    if (code) {
      const taken = ACCOUNTS.some(a=>a.code===code) || extraItems.some(e=>e.code===code);
      if (taken) { alert(`Acc. Code "${code}" มีอยู่แล้ว กรุณาใช้รหัสอื่น`); return; }
    }
    onAddExtra({ name:extraDraft.name, group:extraDraft.group, code: code || undefined });
    setExtraDraft({ code:"", name:"", group:GROUPS[0] }); setAddExtraOpen(false);
  };

  const handleAddSub = (parentCode) => {
    if (!subName.trim()) return;
    onAddExtra({ name:subName, parentCode, addedInMonth: month });
    setRowCollapsed(c => ({...c, [parentCode]: false})); // reveal the newly-added sub-item
    setSubName(""); setSubFor(null);
  };

  const handleDeleteExtra = (code) => {
    onDeleteExtra(code);
    setDraftAdd(d => { const n = {...d}; delete n[code]; return n; });
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
      const seedCol = { id: "legacy", name: "รายการหลัก" };
      allRows.forEach(r => {
        const v = nextDraft[r.code];
        if (v !== undefined && v !== "" && parseFloat(v)) nextDraft[`${r.code}:legacy`] = v;
      });
      nextDraft.$columns = [seedCol, newCol];
    } else {
      nextDraft.$columns = [...columns, newCol];
    }
    setDraftAdd(nextDraft);
    saveAdditions({ ...additions, [month]: { ...(additions[month] || {}), ...nextDraft } });
    setNewColName(""); setAddColOpen(false);
  };

  // ลบคอลัมน์ — เฉพาะเดือนนี้ และเป็นแค่ "ร่าง" เท่านั้น จะมีผลจริงเมื่อกด "บันทึก"
  // ถ้ากด "ยกเลิก" คอลัมน์และค่าที่กรอกไว้จะกลับคืนมา (ไม่โดนลบ) และคอลัมน์ที่ลบ
  // ไปแล้วจะไม่ถูกนำไปคิดยอด (เพราะยอด roll-up ตอนบันทึกจะไม่รวมคอลัมน์นั้น)
  const handleRemoveColumn = (colId) => {
    if (!confirm("ลบคอลัมน์นี้เฉพาะเดือนนี้?\n\n• จะมีผลจริงเมื่อกด \"บันทึก\"\n• กด \"ยกเลิก\" เพื่อคืนคอลัมน์และค่าที่กรอกไว้")) return;
    const nextCols = columns.filter(c => c.id !== colId);
    const nextDraft = { ...draftAdd };
    Object.keys(nextDraft).forEach(k => { if (k.endsWith(`:${colId}`)) delete nextDraft[k]; });
    nextDraft.$columns = nextCols;
    setDraftAdd(nextDraft);
    // ไม่ saveAdditions ที่นี่ — รอกด "บันทึก" (handleSave) เท่านั้น เพื่อให้ยกเลิกได้
  };

  return (
    <div style={{padding:"4px 28px 24px"}}>
      {/* Trend chart — the whole project's cost growth over time, at a glance */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"18px 20px 8px",marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6,flexWrap:"wrap",gap:8}}>
          <span style={{fontSize:13,fontWeight:700,color:T.textPrimary}}>📈 แนวโน้มต้นทุนสะสม</span>
          <span style={{fontSize:12,color:T.textMuted}}>รวมล่าสุดทั้งโปรเจกต์: <b style={{color:T.green,fontFamily:"'JetBrains Mono',monospace",fontSize:15}}>฿{fmt0(grandTotal)}</b>{usdRate>0 && <b className="usd-sub" style={{color:T.green,fontFamily:"'JetBrains Mono',monospace",fontSize:12,marginLeft:6}}>≈ ${fmt(grandTotal/usdRate)}</b>}</span>
        </div>
        <div style={{display:"flex",gap:16,marginBottom:6,fontSize:11,color:T.textMuted,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{display:"inline-flex",alignItems:"center",gap:5}}><span style={{width:10,height:10,borderRadius:2,background:T.blue,display:"inline-block"}}/>ยอดก่อนหน้า (สะสม)</span>
          <span style={{display:"inline-flex",alignItems:"center",gap:5}}><span style={{width:10,height:10,borderRadius:2,background:T.amber,display:"inline-block"}}/>เพิ่มเดือนนี้</span>
          <span style={{color:T.textMuted,fontSize:10.5}}>· คลิกที่แท่งเพื่อเลือกเดือน (เดือนที่เลือกจะมีกรอบ)</span>
        </div>
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={chartData} margin={{top:14,right:8,left:-14,bottom:6}} barCategoryGap="22%"
            onClick={(st)=>{ const mk = st && st.activePayload && st.activePayload[0] && st.activePayload[0].payload && st.activePayload[0].payload.monthKey; if (mk) setMonth(mk); }}
            style={{cursor:"pointer"}}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f7"/>
            <XAxis dataKey="label" tick={<MonthAxisTick selectedLabel={selectedLabel}/>} height={34} axisLine={false} tickLine={false} interval={0}/>
            <YAxis tick={{fontSize:10,fill:T.textMuted}} axisLine={false} tickLine={false} tickFormatter={fmtK}/>
            <Tooltip cursor={{fill:"rgba(37,99,235,0.06)"}} formatter={(v,name)=>[`${fmt(v)} THB`,name]} labelStyle={{color:T.textPrimary,fontWeight:600,marginBottom:2}}
              contentStyle={{borderRadius:10,border:`1px solid ${T.cardBorder}`,fontSize:12,boxShadow:"0 4px 14px rgba(0,0,0,0.08)"}}/>
            <Bar dataKey="previous" stackId="cum" name="ยอดก่อนหน้า" radius={[0,0,0,0]}>
              {chartData.map((e,i)=>{ const sel = e.monthKey===month; return <Cell key={i} fill={T.blue} stroke={sel?"#0f172a":"none"} strokeWidth={sel?2.5:0} cursor="pointer"/>; })}
            </Bar>
            <Bar dataKey="added" stackId="cum" name="เพิ่มเดือนนี้" radius={[5,5,0,0]}>
              {chartData.map((e,i)=>{ const sel = e.monthKey===month; return <Cell key={i} fill={T.amber} stroke={sel?"#0f172a":"none"} strokeWidth={sel?2.5:0} cursor="pointer"/>; })}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Month picker — ปกติดูอย่างเดียว (คลิกสลับเดือน) · กด "จัดการเดือน" เพื่อเข้าโหมดเพิ่ม/ลบ */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16,flexWrap:"wrap"}}>
        <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:6,flex:1,minWidth:0}}>
          {sortedMonths.map(m=>{
            const active = m===month;
            const add = monthTotalLive(m);
            const exists = months.includes(m); // เดือนที่มีจริง (ไม่ใช่ default เปล่า) ถึงลบได้
            return (
              <div key={m} onClick={()=>setMonth(m)} ref={active?activeChipRef:null}
                style={{position:"relative",flexShrink:0,textAlign:"left",padding:"10px 16px",borderRadius:12,border:`1.5px solid ${active?T.blue:T.cardBorder}`,
                  background:active?T.blue:T.card,cursor:"pointer",minWidth:140,transition:"all 0.15s"}}>
                <div style={{fontSize:11,fontWeight:600,color:active?"#bfdbfe":T.textSecondary,marginBottom:3}}>{monthShortLabel(m)}</div>
                <div style={{fontSize:15,fontWeight:700,color:active?"#fff":T.textPrimary,fontFamily:"'JetBrains Mono',monospace"}}>{fmtK(cumulativeLive(m))}</div>
                <div style={{fontSize:10,color:active?"#dbeafe":T.textMuted,marginTop:2}}>{add>0?"+":""}{fmtK(add)} เดือนนี้</div>
                {monthEditMode && exists && (
                  <button onClick={(e)=>{e.stopPropagation(); handleDeleteMonth(m);}} title="ลบเดือนนี้"
                    style={{position:"absolute",top:5,right:5,width:20,height:20,borderRadius:6,border:"none",lineHeight:1,
                      background:active?"rgba(255,255,255,0.18)":T.redBg,color:active?"#fff":T.red,cursor:"pointer",fontSize:13,padding:0,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
                )}
              </div>
            );
          })}
          {monthEditMode && (
            <div style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,padding:"0 12px",borderRadius:12,border:`1.5px dashed ${T.blue}`,background:T.blueLight}}>
              <input type="month" value={newMonth} onChange={e=>setNewMonth(e.target.value)} className="input-base"
                style={{border:"none",background:"transparent",padding:"8px 4px",width:118,fontSize:12}}/>
              <button className="btn-primary" style={{padding:"6px 12px",fontSize:11,whiteSpace:"nowrap",background:T.blue}} onClick={handleAddMonth}>+ เพิ่มเดือน</button>
            </div>
          )}
        </div>
        <button onClick={()=>setMonthEditMode(v=>!v)}
          style={{flexShrink:0,alignSelf:"flex-start",padding:"9px 14px",borderRadius:10,border:`1.5px solid ${monthEditMode?T.blue:T.cardBorder}`,
            background:monthEditMode?T.blue:T.card,color:monthEditMode?"#fff":T.textSecondary,fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
          {monthEditMode?"✓ เสร็จ":"✏️ จัดการเดือน"}
        </button>
      </div>

      {/* Stats for selected month */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
        <StatCard label="ยอดยกมา (ก่อนเดือนนี้)" value={"฿"+fmt0(baselineForMonth)} thb={baselineForMonth} rate={usdRate} sub={`สะสมถึง ${prevMonthLabel}`} color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label="เพิ่มเดือนนี้" value={"฿"+fmt0(thisMonthAdd)} thb={thisMonthAdd} rate={usdRate} sub={new Date(month+"-01").toLocaleDateString("th-TH",{year:"numeric",month:"long"})} color={T.amber} icon="➕" accent={T.amberBg}/>
        <StatCard label="รวมสะสมถึงเดือนนี้" value={"฿"+fmt0(cumulativeSoFar)} thb={cumulativeSoFar} rate={usdRate} sub="เดิม + เพิ่มสะสมถึงเดือนที่เลือก" color={T.green} icon="✅" accent={T.greenBg}/>
        <StatCard label="รวมทั้งหมด" value={"฿"+fmt0(grandTotal)} thb={grandTotal} rate={usdRate} sub="เดิม + ทุกเดือนที่มีข้อมูล (ล่าสุด)" color={T.purple} icon="🧮" accent={T.purpleBg}/>
      </div>

      {/* Toolbar: search + group filter + actions */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <SearchInput value={search} onChange={setSearch} placeholder="🔍 ค้นหา Account Code / ชื่อ..." width={220}/>
        <button onClick={()=>setHideEmpty(v=>!v)}
          title="ซ่อน/แสดงแถวที่ไม่มีค่า (รวมสะสม = 0)"
          style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:8,fontSize:11.5,fontWeight:600,cursor:"pointer",
            border:`1.5px solid ${hideEmpty?T.blue:T.cardBorder}`,background:hideEmpty?T.blue:T.card,color:hideEmpty?"#fff":T.textSecondary,whiteSpace:"nowrap"}}>
          {hideEmpty ? `✓ เฉพาะที่มีค่า${hiddenEmptyCount?` (ซ่อน ${hiddenEmptyCount})`:""}` : "⚡ เฉพาะที่มีค่า"}
        </button>
        <GroupFilter selected={filter} onChange={setFilter}/>
        <div style={{flex:1}}/>
        <button className="btn-ghost" onClick={()=>setAddExtraOpen(v=>!v)} disabled={!editingUnlocked}
          style={!editingUnlocked?{opacity:0.4,cursor:"not-allowed"}:undefined}>+ งานพิเศษ</button>
        {!editingUnlocked && (
          <span style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:T.textMuted,background:"#f1f5f9",padding:"6px 12px",borderRadius:8,fontWeight:600}}>
            🔒 บันทึกแล้ว
          </span>
        )}
        {editingUnlocked ? (
          <>
            <button onClick={handleSave} className="btn-primary" style={{background:saved?T.green:T.blue,minWidth:170}}>
              {saved?"✓ บันทึกแล้ว":`บันทึกรายการเดือนนี้`}
            </button>
            {canCancel && (
              <button onClick={handleCancel} className="btn-ghost" title="ยกเลิกการแก้ไข (Esc)"
                style={{color:T.red,borderColor:T.red}}>✕ ยกเลิก</button>
            )}
          </>
        ) : (
          <button onClick={()=>setForceEdit(true)} className="btn-primary" style={{background:T.amber,minWidth:170}}>
            ✏️ แก้ไขเดือนนี้
          </button>
        )}
      </div>

      {addExtraOpen && (
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"1fr 2fr 1fr auto",gap:10,alignItems:"end"}}>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>Acc. Code (เว้นว่างให้สร้างอัตโนมัติ)</span>
            <input className="input-base" value={extraDraft.code} onChange={e=>setExtraDraft(d=>({...d,code:e.target.value}))} placeholder="เช่น 511099" />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>ชื่อรายการงานเพิ่ม</span>
            <input className="input-base" value={extraDraft.name} onChange={e=>setExtraDraft(d=>({...d,name:e.target.value}))} placeholder="เช่น งานเพิ่มกระจกโค้งพิเศษ" onKeyDown={e=>e.key==="Enter"&&handleCreateExtra()} />
          </label>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>Group</span>
            <select className="input-base" value={extraDraft.group} onChange={e=>setExtraDraft(d=>({...d,group:e.target.value}))}>
              {GROUPS.map(g => <option key={g} value={g}>{g}</option>)}
            </select>
          </label>
          <button className="btn-primary" onClick={handleCreateExtra}>+ สร้างรายการ</button>
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
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"left",color:sortKey==="code"?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                    <span onClick={()=>handleSort("code")} style={{cursor:"pointer",userSelect:"none"}}>Acc. Code{sortKey==="code"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"left",color:sortKey==="group"?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                    <span onClick={()=>handleSort("group")} style={{cursor:"pointer",userSelect:"none"}}>Group{sortKey==="group"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"left",color:sortKey==="name"?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                    <span onClick={()=>handleSort("name")} style={{cursor:"pointer",userSelect:"none"}}>Account Name{sortKey==="name"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"right",color:sortKey==="before"?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                    <span onClick={()=>handleSort("before")} style={{cursor:"pointer",userSelect:"none"}}>📐 ยอดก่อนหน้า{sortKey==="before"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"center",width:20,color:T.textMuted,borderBottom:`1px solid ${T.cardBorder}`}}>+</th>
                  <th colSpan={columns.length+1} style={{padding:"9px 16px",textAlign:"center",color:T.textMuted,fontWeight:700,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>
                    ➕ เพิ่มเดือนนี้ · {monthShortLabel(month)}
                  </th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"center",width:20,color:T.textMuted,borderBottom:`1px solid ${T.cardBorder}`}}>=</th>
                  <th rowSpan={2} style={{padding:"11px 16px",textAlign:"right",color:sortKey==="cum"?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                    <span onClick={()=>handleSort("cum")} style={{cursor:"pointer",userSelect:"none"}}>✅ รวมสะสม{sortKey==="cum"?(sortDir===1?" ▲":" ▼"):""}</span>
                  </th>
                  <th rowSpan={2} style={{width:20,borderBottom:`1px solid ${T.cardBorder}`}}></th>
                </tr>
                <tr style={{background:"#f8fafc"}}>
                  {columns.map(c=>(
                    <th key={c.id} style={{padding:"6px 18px",textAlign:"right",color:T.textMuted,fontWeight:600,fontSize:10.5,borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                      <div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:5}}>
                        <span>{c.name}</span>
                        {editingUnlocked && <button onClick={()=>handleRemoveColumn(c.id)} title="ลบรายการนี้ (เฉพาะเดือนนี้)" style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:11,padding:0}}>✕</button>}
                      </div>
                    </th>
                  ))}
                  <th style={{padding:"6px 10px",textAlign:"right",borderBottom:`1px solid ${T.cardBorder}`}}>
                    {editingUnlocked && (addColOpen ? (
                      <div style={{display:"flex",gap:4,alignItems:"center",justifyContent:"flex-end"}}>
                        <input autoFocus value={newColName} onChange={e=>setNewColName(e.target.value)} placeholder="ชื่อ เช่น CC#17"
                          className="input-base" style={{width:88,fontSize:11,padding:"4px 6px"}}
                          onKeyDown={e=>e.key==="Enter"&&handleAddColumn()} />
                        <button onClick={handleAddColumn} className="btn-primary" style={{padding:"4px 9px",fontSize:11}}>+</button>
                        <button onClick={()=>setAddColOpen(false)} className="btn-ghost" style={{padding:"4px 7px",fontSize:11}}>×</button>
                      </div>
                    ) : (
                      <button onClick={()=>setAddColOpen(true)} className="btn-ghost" style={{padding:"4px 10px",fontSize:11,whiteSpace:"nowrap"}}>+ เพิ่มรายการ</button>
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
                  {label:"📐 ยอดก่อนหน้า", key:"before", align:"right"},
                  {label:"+", key:null, align:"center", width:20},
                  {label:"➕ เพิ่มเดือนนี้", key:"add", align:"right"},
                  {label:"=", key:null, align:"center", width:20},
                  {label:"✅ รวมสะสม", key:"cum", align:"right"},
                  {label:"", key:null, width:20},
                ].map(({label,key,align,width},idx)=>(
                  <th key={idx}
                    style={{padding:"11px 16px",textAlign:align||"left",color:key&&sortKey===key?T.blue:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:label.length>2?0.8:0,textTransform:label.length>2?"uppercase":"none",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap",...(width?{width}:{})}}>
                    {key==="add" ? (
                      <div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:8}}>
                        <span onClick={()=>handleSort("add")} style={{cursor:"pointer",userSelect:"none"}}>{label}{sortKey==="add"?(sortDir===1?" ▲":" ▼"):""}</span>
                        {editingUnlocked && (addColOpen ? (
                          <div style={{display:"flex",gap:4,alignItems:"center"}} onClick={e=>e.stopPropagation()}>
                            <input autoFocus value={newColName} onChange={e=>setNewColName(e.target.value)} placeholder="ชื่อ เช่น CC#17"
                              className="input-base" style={{width:88,fontSize:11,padding:"4px 6px",textTransform:"none"}}
                              onKeyDown={e=>e.key==="Enter"&&handleAddColumn()} />
                            <button onClick={handleAddColumn} className="btn-primary" style={{padding:"4px 9px",fontSize:11}}>+</button>
                            <button onClick={()=>setAddColOpen(false)} className="btn-ghost" style={{padding:"4px 7px",fontSize:11}}>×</button>
                          </div>
                        ) : (
                          <button onClick={(e)=>{e.stopPropagation();setAddColOpen(true);}} className="btn-ghost" style={{padding:"3px 8px",fontSize:10,whiteSpace:"nowrap",textTransform:"none"}}>+ เพิ่มรายการ</button>
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
              return (
                <Fragment key={r.code}>
                  <tr onClick={()=>hasKids && setRowCollapsed(c=>({...c,[r.code]:!c[r.code]}))}
                      style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(hasKids&&!isCollapsed)||subFor===r.code?"none":"1px solid #f1f5f9",cursor:hasKids?"pointer":"default"}}>
                    <td style={{padding:"10px 16px",color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontSize:12,fontWeight:500}}>
                      {hasKids && (
                        <span title={isCollapsed?"ขยายรายการย่อย":"ย่อรายการย่อย"}
                          style={{color:T.textMuted,fontSize:10,marginRight:6,verticalAlign:"middle",display:"inline-block"}}>
                          {isCollapsed?"▸":"▾"}
                        </span>
                      )}
                      {r.code}
                    </td>
                    <td style={{padding:"10px 16px"}}>
                      <span style={{background:T.blueLight,color:T.blue,fontSize:10,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{r.group}</span>
                    </td>
                    <td style={{padding:"10px 16px",color:T.textPrimary}}>
                      {r.name}
                      {r.isExtra && <span style={{marginLeft:7,fontSize:10,background:T.amberBg,color:T.amber,padding:"1px 8px",borderRadius:6,fontWeight:600}}>งานเพิ่ม</span>}
                      {hasKids && <span style={{marginLeft:7,fontSize:10,background:T.greenBg,color:T.green,padding:"1px 8px",borderRadius:6,fontWeight:600}}>{kids.length} รายการย่อย</span>}
                      {editingUnlocked && (
                        <button onClick={(e)=>{e.stopPropagation(); setSubFor(subFor===r.code?null:r.code); setSubName(""); setRowCollapsed(c=>({...c,[r.code]:false}));}} title="เพิ่มรายการย่อยใต้ Acc. Code นี้"
                          style={{marginLeft:9,background:"none",border:`1px dashed ${T.cardBorder}`,borderRadius:6,color:T.textMuted,cursor:"pointer",fontSize:10,padding:"1px 7px"}}>
                          + รายการย่อย
                        </button>
                      )}
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right",color:cumBefore!==0?T.textPrimary:T.textMuted,fontFamily:"'JetBrains Mono',monospace"}} title="ราคาเดิม + ยอดเพิ่มของทุกเดือนก่อนหน้ารวมกัน">{fmt(cumBefore)}{usdLine(cumBefore, usdRate)}</td>
                    <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                    {isMultiCol ? (
                      hasKids ? (
                        <td colSpan={columns.length} style={{padding:"8px 16px",textAlign:"right"}}>
                          <div style={{width:"100%",padding:"7px 10px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:T.amberBg,borderRadius:8,color:T.amber,fontWeight:700,fontSize:13}}>
                            {fmt(thisVal)}
                            {usdLine(thisVal, usdRate)}
                          </div>
                        </td>
                      ) : columns.map(c=>{
                        const ck = `${r.code}:${c.id}`;
                        const cv = parseFloat(draftAdd[ck])||0;
                        return (
                          <td key={c.id} style={{padding:"8px 10px",textAlign:"right"}}>
                            {editingUnlocked ? (
                              <MoneyInput value={draftAdd[ck]??""} onChange={v=>setDraftAdd(d=>({...d,[ck]:v}))}
                                style={{width:104,fontSize:12,background:cv!==0?T.amberBg:T.bg}}/>
                            ) : (
                              <div style={{width:104,marginLeft:"auto",padding:"7px 8px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:cv!==0?T.textPrimary:T.textMuted}}>{fmt(cv)}{usdLine(cv, usdRate)}</div>
                            )}
                          </td>
                        );
                      })
                    ) : (
                      <td style={{padding:"8px 16px",textAlign:"right"}}>
                        {hasKids ? (
                          <div style={{width:130,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:T.amberBg,borderRadius:8,color:T.amber,fontWeight:700,fontSize:13}}>
                            {fmt(thisVal)}
                            {usdLine(thisVal, usdRate)}
                          </div>
                        ) : editingUnlocked ? (
                          <MoneyInput value={draftAdd[r.code]??""} onChange={v=>setDraftAdd(d=>({...d,[r.code]:v}))}
                            style={{width:130,background:thisVal!==0?T.amberBg:T.bg}}/>
                        ) : (
                          <div style={{width:130,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:13,color:thisVal!==0?T.textPrimary:T.textMuted}}>{fmt(thisVal)}{usdLine(thisVal, usdRate)}</div>
                        )}
                      </td>
                    )}
                    <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                    <td style={{padding:"8px 16px",textAlign:"right",color:cum!==0?T.textPrimary:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{fmt(cum)}{usdLine(cum, usdRate)}</td>
                    <td style={{padding:"8px 16px",textAlign:"center"}}>
                      {r.isExtra && editingUnlocked && (
                        <button onClick={(e)=>{e.stopPropagation(); handleDeleteExtra(r.code);}} title="ลบรายการงานเพิ่ม"
                          style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
                      )}
                    </td>
                  </tr>

                  {/* Sub-items — this month's value rolls up into the parent row above.
                      Only sub-items added on/before the currently-viewed month appear here
                      (kidsAsOf already filtered them), so a sub-item created in ก.ย. simply
                      doesn't exist in ส.ค. or earlier — no ghost "0.00" row. */}
                  {!isCollapsed && kids.map((k,ki) => {
                    const kBaseVal = parseFloat(tenderCosts[k.code]) || 0;
                    const kCumBefore = months.filter(m=>m<month).reduce((s,m)=>s+(parseFloat(additions[m]?.[k.code])||0),0) + kBaseVal;
                    const kThisVal = parseFloat(draftAdd[k.code]) || 0;
                    const kCum = kCumBefore + kThisVal;
                    const isNewThisMonth = k.addedInMonth === month;
                    return (
                      <tr key={k.code} style={{background:isNewThisMonth?T.greenBg:i%2===0?T.card:"#fafbfd",borderLeft:`3px solid ${isNewThisMonth?T.green:"#e2e8f0"}`,borderBottom:(ki===kids.length-1 && subFor!==r.code)?"1px solid #f1f5f9":"none",transition:"background 0.2s"}}>
                        <td style={{padding:"7px 16px 7px 27px",color:T.green,fontSize:13}}>↳</td>
                        <td/>
                        <td style={{padding:"7px 16px",color:T.green,fontSize:12.5,fontStyle:"italic"}}>
                          {k.name}
                          {k.addedInMonth && (
                            isNewThisMonth ? (
                              <span title="รายการนี้เพิ่งเพิ่มเข้ามาในเดือนนี้" style={{marginLeft:8,fontSize:10,background:T.green,color:"#fff",padding:"2px 8px",borderRadius:6,fontWeight:700,fontStyle:"normal",letterSpacing:0.2}}>
                                ✨ ใหม่เดือนนี้
                              </span>
                            ) : (
                              <span title="เพิ่มเข้ามาระหว่างทาง ไม่ได้มีมาตั้งแต่ต้น — เดือนก่อนหน้านั้นจะไม่แสดงรายการนี้" style={{marginLeft:8,fontSize:10,background:T.amberBg,color:T.amber,padding:"2px 8px",borderRadius:6,fontWeight:600,fontStyle:"normal"}}>
                                เพิ่มเมื่อ {monthShortLabel(k.addedInMonth)}
                              </span>
                            )
                          )}
                        </td>
                        <td style={{padding:"7px 16px",textAlign:"right",color:kCumBefore!==0?T.textPrimary:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12.5}}>{fmt(kCumBefore)}{usdLine(kCumBefore, usdRate)}</td>
                        <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                        <td style={{padding:"7px 16px",textAlign:"right"}}>
                          {editingUnlocked ? (
                            <MoneyInput value={draftAdd[k.code]??""} onChange={v=>setDraftAdd(d=>({...d,[k.code]:v}))}
                              style={{width:130,fontSize:12.5,background:kThisVal!==0?T.greenBg:T.bg}}/>
                          ) : (
                            <div style={{width:130,marginLeft:"auto",padding:"7px 8px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12.5,color:kThisVal!==0?T.textPrimary:T.textMuted}}>{fmt(kThisVal)}{usdLine(kThisVal, usdRate)}</div>
                          )}
                        </td>
                        <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                        <td style={{padding:"7px 16px",textAlign:"right",color:kCum!==0?T.textPrimary:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:12.5}}>{fmt(kCum)}{usdLine(kCum, usdRate)}</td>
                        <td style={{padding:"7px 16px",textAlign:"center"}}>
                          {editingUnlocked && (
                            <button onClick={()=>handleDeleteExtra(k.code)} title="ลบรายการย่อยนี้"
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
                      <td/><td/>
                      <td style={{padding:"7px 16px"}}>
                        <input className="input-base" value={subName} onChange={e=>setSubName(e.target.value)}
                          placeholder="ชื่อรายการย่อย เช่น Silicone Structure" style={{width:"100%",fontSize:12}}
                          onKeyDown={e=>e.key==="Enter"&&handleAddSub(r.code)} autoFocus />
                      </td>
                      <td colSpan={(isMultiCol ? 8+columns.length : 9)-3} style={{padding:"7px 16px",display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button className="btn-primary" style={{padding:"5px 12px",fontSize:12}} onClick={()=>handleAddSub(r.code)}>+ เพิ่ม</button>
                        <button className="btn-ghost" style={{padding:"5px 12px",fontSize:12}} onClick={()=>setSubFor(null)}>ยกเลิก</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={isMultiCol ? 8+columns.length : 9} style={{padding:"28px 16px",textAlign:"center",color:T.textMuted,fontSize:13}}>ไม่พบรายการที่ตรงกับการค้นหา</td></tr>
            )}
          </tbody>
          <tfoot>
            <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
              <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:12}}>{filtered.length} รายการ</td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.textPrimary,fontFamily:"'JetBrains Mono',monospace",fontWeight:600,fontSize:13}}>
                {fmt(filtered.reduce((s,r)=>s+cumBeforeOf(r),0))}
                {usdLine(filtered.reduce((s,r)=>s+cumBeforeOf(r),0), usdRate)}
              </td>
              <td/>
              {isMultiCol
                ? columns.map(c => { const ct = filtered.reduce((s,r)=> s + (parseFloat(draftAdd[`${r.code}:${c.id}`])||0), 0); return (
                    <td key={c.id} style={{padding:"12px 18px",textAlign:"right",color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:12.5,whiteSpace:"nowrap"}}>
                      {fmt(ct)}
                      {usdLine(ct, usdRate)}
                    </td>
                  ); })
                : (
                    <td style={{padding:"12px 16px",textAlign:"right",color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:13}}>
                      {fmt(filtered.reduce((s,r)=>s+rowMonthValue(r.code, month, draftAdd),0))}
                      {usdLine(filtered.reduce((s,r)=>s+rowMonthValue(r.code, month, draftAdd),0), usdRate)}
                    </td>
                  )
              }
              <td/>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:14}}>
                {fmt(filtered.reduce((s,r)=>s+cumBeforeOf(r)+rowMonthValue(r.code, month, draftAdd),0))}
                {usdLine(filtered.reduce((s,r)=>s+cumBeforeOf(r)+rowMonthValue(r.code, month, draftAdd),0), usdRate)}
              </td>
              <td/>
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
const evalMoney = (expr) => {
  const cleaned = String(expr ?? "").replace(/[,\s]/g, "");
  if (!cleaned) return "";
  const terms = cleaned.match(/[+-]?\d*\.?\d+/g);
  if (!terms) return "";
  const sum = terms.reduce((s, t) => s + (parseFloat(t) || 0), 0);
  if (isNaN(sum)) return "";
  // เงินติดลบไม่มีความหมาย (งบ/ยอด PO/ยอดรับ) — ยังพิมพ์สูตรลบได้ (เช่น 100-20=80)
  // แต่ถ้าผลรวมออกมาติดลบ ให้เป็น 0 กันข้อมูลเสียหาย
  return String(Math.max(0, sum));
};
const fmtMoneyInput = (v) => {
  if (v === "" || v == null || isNaN(Number(v))) return "";
  return Number(v).toLocaleString("en-US", { maximumFractionDigits: 2 });
};
function MoneyInput({ value, onChange, placeholder = "0", disabled, className = "input-base", style }) {
  const [focused, setFocused] = useState(false);
  const [text, setText] = useState("");
  const commit = () => { onChange(evalMoney(text)); setFocused(false); };
  return (
    <input
      type="text"
      inputMode="decimal"
      className={className}
      disabled={disabled}
      placeholder={placeholder}
      style={{ textAlign: "right", fontFamily: "'JetBrains Mono',monospace", ...(style || {}) }}
      value={focused ? text : fmtMoneyInput(value)}
      onFocus={() => { setFocused(true); setText(value != null && value !== "" ? String(value) : ""); }}
      onClick={(e) => e.stopPropagation()}
      onChange={(e) => setText(e.target.value.replace(/[^0-9.+\-,\s]/g, ""))}
      onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); commit(); e.currentTarget.blur(); } }}
      onBlur={commit}
      title="พิมพ์บวก/ลบได้ เช่น 20000+10000 แล้วกด Enter เพื่อรวมยอด"
    />
  );
}

// ─── Procurement: PO Detail Modal ──────────────────────────────────────────────
// Read-only detail view opened by clicking any PO row. Lets the user confirm
// exactly what was entered without hunting through a wide table, and offers
// Edit / Delete from the same place.
function PODetailModal({ po: rawPo, onClose, onEdit, onDelete, onStatusChange, onChangePO, session, usdRate=0 }) {
  const [historyOpen, setHistoryOpen] = useState(false);
  const [capWarn, setCapWarn] = useState(""); // เตือนเมื่อยอดของเข้าจริงรวมเกินยอดสั่ง
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

  // Record actual received / split remaining into a new round, then persist.
  const setItemRounds = (itemId, rounds) =>
    onChangePO?.({ ...po, items: po.items.map(it => it.id===itemId ? {...it, rounds} : it) });
  // ลงยอดของเข้าจริง — ห้ามให้ยอดรวมทุกงวดเกิน "ยอดสั่ง" ของ PO นั้น (บล็อก+เตือน)
  const setActualAmount = (itemId, roundId, val) => {
    const it = po.items.find(i=>i.id===itemId); if (!it) return;
    const ordered = itemOrdered(it);
    const newVal = parseFloat(val)||0;
    const otherReceived = (it.rounds||[]).filter(r=>r.id!==roundId).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
    if (ordered>0 && (otherReceived + newVal) > ordered + 0.001) {
      const maxAllow = Math.max(ordered - otherReceived, 0);
      setCapWarn(`⚠ ${it.code||"รายการนี้"}: ยอดของเข้ารวมห้ามเกินยอดสั่ง ${fmt(ordered)} — งวดนี้กรอกได้ไม่เกิน ${fmt(maxAllow)} (ระบบไม่บันทึกค่าที่เกิน)`);
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
  const removeRound = (itemId, roundId) => {
    const it = po.items.find(i=>i.id===itemId); if (!it) return;
    if ((it.rounds||[]).length <= 1) return;
    if (!confirm("ลบงวดนี้? (ยอด/วันของเข้าที่กรอกในงวดนี้จะถูกลบ)")) return;
    setItemRounds(itemId, it.rounds.filter(r => r.id !== roundId));
  };
  const roundBadge = (r) => {
    if (!r.actualDate || !(parseFloat(r.actualAmount)||0)) return ["รอของเข้า", PAYMENT_BG.pending, PAYMENT_CLR.pending];
    // ถ้าวันของเข้าจริงยังมาไม่ถึง (วันในอนาคต) = ยังไม่ถือว่ารับของ แสดงเป็น "นัดรับ"
    if (r.actualDate > todayStr()) return [`นัดรับ ${r.actualDate} (ยังไม่ถึงวัน)`, INCOMING_BG.pending, INCOMING_CLR.pending];
    return roundPaid(po,r) ? ["ถึงกำหนดจ่ายแล้ว", PAYMENT_BG.paid, PAYMENT_CLR.paid]
                           : ["ของเข้าแล้ว · รอครบกำหนด", INCOMING_BG.partial, INCOMING_CLR.partial];
  };

  const Row = ({ label, value, mono }) => (
    <div style={{display:"flex",justifyContent:"space-between",gap:16,padding:"10px 0",borderBottom:`1px solid #f1f5f9`}}>
      <span style={{fontSize:12,color:T.textMuted,fontWeight:500}}>{label}</span>
      <span style={{fontSize:13,color:T.textPrimary,fontWeight:600,textAlign:"right",fontFamily:mono?"'JetBrains Mono',monospace":undefined}}>{value ?? "—"}</span>
    </div>
  );

  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,padding:20,animation:"fadeIn 0.15s ease"}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:16,padding:26,width:"100%",maxWidth:520,maxHeight:"88vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.25)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
          <div>
            <div style={{fontSize:16,fontWeight:700,color:T.textPrimary}}>{poSupplierLabel(po)}</div>
            <div style={{fontSize:12,color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",marginTop:2}}>{poNumbersLabel(po)}</div>
          </div>
          <button onClick={onClose} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted,flexShrink:0}}>×</button>
        </div>

        {locked && (
          <div style={{display:"flex",alignItems:"center",gap:6,background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"6px 10px",margin:"8px 0 2px",fontSize:11,color:"#92400e"}}>
            🔒 รับของและจ่ายเงินครบแล้ว — แก้ไข/ลบได้เฉพาะ Admin
          </div>
        )}

        {/* Status is a live dropdown here too — the most natural place to
            update it right after reviewing everything else on the PO. */}
        <div style={{display:"flex",gap:6,margin:"12px 0 4px",flexWrap:"wrap",alignItems:"center"}}>
          <StatusPicker status={po.status} onChange={s=>onStatusChange?.(po,s)} disabled={locked}/>
          <span style={{background:INCOMING_BG[inc],color:INCOMING_CLR[inc],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{INCOMING_LABEL[inc]}</span>
          <span style={{background:PAYMENT_BG[pay],color:PAYMENT_CLR[pay],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{PAYMENT_LABEL[pay]}</span>
          {po.paymentType && (
            <span style={{background:PAYMENT_TYPE_BG[po.paymentType],color:PAYMENT_TYPE_CLR[po.paymentType],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{PAYMENT_TYPE_ICON[po.paymentType]} {paymentTypeLabel(po)}</span>
          )}
        </div>
        {lastUpd && (
          <div style={{fontSize:11,color:T.textMuted,marginBottom:4}}>
            🕓 อัปเดตล่าสุด {relativeTime(lastUpd.at)} โดย <b style={{color:T.textSecondary}}>{lastUpd.user}</b>
          </div>
        )}

        {/* Supplier (one per PO) + top-line dates */}
        <div style={{marginTop:10,background:T.bg,borderRadius:10,padding:"10px 12px"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline"}}>
            <div><span style={{fontSize:13,fontWeight:700,color:T.textPrimary}}>{supplier.name||"—"}</span>
              {supplier.poNumber && <span style={{fontSize:11,color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",marginLeft:8}}>{supplier.poNumber}</span>}</div>
            <span style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:T.amber}}>{fmt(poTotal(po))}{usdRate>0 && <span className="usd-sub" style={{color:T.green,fontWeight:700,fontSize:12,marginLeft:6}}>≈ ${fmt(poTotal(po)/usdRate)}</span>}</span>
          </div>
        </div>

        <div style={{marginTop:4}}>
          <Row label="วันเปิด PO" value={po.date} mono />
          <Row label="วันรับของ" value={receivedDates.length ? receivedDates.join(", ") : "ยังไม่ได้รับ"} mono={receivedDates.length>0} />
          <Row label="วันจ่ายเงิน" value={paidDate || "ยังไม่ถึงกำหนด"} mono={!!paidDate} />
          <Row label="วิธีจ่ายเงิน" value={po.paymentType ? `${PAYMENT_TYPE_ICON[po.paymentType]} ${paymentTypeLabel(po)}` : "—"} />
        </div>

        {/* Per account-code: receiving in installments, with auto-pay + split */}
        <div style={{marginTop:12}}>
          <div style={{fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase",marginBottom:8}}>📦 ของเข้า / จ่ายเงิน (แบ่งงวดได้)</div>
          {items.map((it,ii)=>{
            const acc = ACCOUNTS.find(a=>a.code===it.code);
            const ordered = itemOrdered(it), recv = itemReceived(it), remain = itemRemaining(it);
            const planned = (it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.planAmount)||0),0); // ยอดรวมที่วางแผนไว้ทุกงวด
            const planRemain = Math.max(ordered - planned, 0);   // ยอดที่ยัง "ไม่ถูกวางแผน" (ไว้แบ่งงวดเพิ่ม)
            const overPlanned = planned - ordered;               // >0 = รวมทุกงวดเกินยอดสั่ง (มีงวดเกิน/ซ้ำ)
            const paidAmt = (it.rounds||[]).filter(r=>roundPaid(po,r)).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0);
            return (
              <div key={it.id||ii} style={{background:T.bg,borderRadius:12,padding:"12px 14px",marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:8}}>
                  <div style={{minWidth:0}}>
                    <span style={{fontSize:11,color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{it.code||"—"}</span>
                    <span style={{fontSize:12,color:T.textSecondary,marginLeft:8}}>{acc?.name||"—"}</span>
                  </div>
                  <span style={{fontSize:12,color:T.textMuted}}>สั่ง <b style={{color:T.textPrimary,fontFamily:"'JetBrains Mono',monospace"}}>{fmt(ordered)}</b></span>
                </div>

                {(it.rounds||[]).map((r,ri)=>{
                  const [label,bg,clr] = roundBadge(r);
                  const payDate = roundPayDate(po,r);
                  const late = r.actualDate && r.planDate && r.actualDate>r.planDate;
                  return (
                    <div key={r.id||ri} style={{border:`1px solid ${T.cardBorder}`,borderRadius:10,padding:10,marginBottom:6,background:T.card}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,marginBottom:8}}>
                        <span style={{fontSize:11,fontWeight:700,color:T.textSecondary}}>งวดที่ {ri+1}{(parseFloat(r.planAmount)||0)>0 ? ` · แผน ${fmt(r.planAmount)}` : ""}</span>
                        <div style={{display:"flex",alignItems:"center",gap:8}}>
                          <span style={{background:bg,color:clr,fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600}}>{label}</span>
                          {!locked && it.rounds.length>1 && (
                            <button type="button" onClick={()=>removeRound(it.id,r.id)} title="ลบงวดนี้"
                              style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13,padding:"2px 4px",borderRadius:6,lineHeight:1}}>🗑</button>
                          )}
                        </div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                        <label style={{display:"flex",flexDirection:"column",gap:3}}>
                          <span style={{fontSize:10,color:T.textSecondary}}>ยอดของเข้าจริง (บาท)</span>
                          <MoneyInput value={r.actualAmount} disabled={locked} placeholder="บาท"
                            onChange={v=>setActualAmount(it.id,r.id,v)}/>
                        </label>
                        <label style={{display:"flex",flexDirection:"column",gap:3}}>
                          <span style={{fontSize:10,color:T.textSecondary}}>วันของเข้าจริง</span>
                          <input type="date" value={r.actualDate} disabled={locked}
                            onChange={e=>updateRound(it.id,r.id,"actualDate",e.target.value)} className="input-base"/>
                        </label>
                      </div>
                      <div style={{marginTop:6,fontSize:11,color:T.textSecondary}}>
                        💰 วันครบกำหนดจ่าย: <span style={{fontFamily:"'JetBrains Mono',monospace",color:T.textPrimary}}>{payDate||"—"}</span>
                        <span style={{color:T.textMuted}}> ({po.paymentType==="cash"?"เงินสด":po.paymentType==="credit"?`เครดิต ${po.creditDays} วัน`:"ยังไม่ระบุวิธีจ่าย"})</span>
                        {late && <span style={{color:T.red}}> · ของมาช้า</span>}
                      </div>
                    </div>
                  );
                })}

                {/* Progress + split */}
                <div style={{height:8,borderRadius:6,background:T.cardBorder,overflow:"hidden",marginTop:6}}>
                  <div style={{height:"100%",width:`${ordered>0?Math.min(recv/ordered*100,100):0}%`,background:remain>0?T.amber:T.green}}/>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",marginTop:5,fontSize:11,color:T.textSecondary}}>
                  <span>ของเข้าแล้ว <b style={{fontFamily:"'JetBrains Mono',monospace",color:T.textPrimary}}>{fmt(recv)}</b> / {fmt(ordered)}</span>
                  <span>จ่ายแล้ว <b style={{fontFamily:"'JetBrains Mono',monospace",color:paidAmt>0?T.green:T.textMuted}}>{fmt(paidAmt)}</b></span>
                </div>
                {overPlanned>0.001 && (
                  <div style={{marginTop:8,fontSize:11.5,color:T.red,background:T.redBg,borderRadius:8,padding:"7px 10px",lineHeight:1.4}}>
                    ⚠ ยอดรวมทุกงวด <b style={{fontFamily:"'JetBrains Mono',monospace"}}>{fmt(planned)}</b> เกินยอดสั่ง <b style={{fontFamily:"'JetBrains Mono',monospace"}}>{fmt(ordered)}</b> อยู่ {fmt(overPlanned)} — กด 🗑 ลบงวดที่เกินออก
                  </div>
                )}
                {!locked && ordered>0 && remain>0.001 ? (
                  <button type="button" onClick={()=>splitRound(it.id)} className="btn-ghost"
                    style={{marginTop:8,padding:"6px 12px",fontSize:12,borderColor:T.amber,color:T.amber}}>
                    ➕ เพิ่มงวดของเข้า — เหลือรับอีก {fmt(remain)}
                  </button>
                ) : recv>0.001 && remain<=0.001 && ordered>0 ? (
                  <div style={{marginTop:8,fontSize:12,color:T.green}}>✓ ของเข้าครบตามยอดสั่งแล้ว</div>
                ) : null}
              </div>
            );
          })}
        </div>

        {po.notes && (
          <div style={{padding:"10px 0 0"}}>
            <div style={{fontSize:12,color:T.textMuted,fontWeight:500,marginBottom:4}}>หมายเหตุ</div>
            <div style={{fontSize:13,color:T.textPrimary,lineHeight:1.5,whiteSpace:"pre-wrap"}}>{po.notes}</div>
          </div>
        )}

        {/* Edit history — a running log of who changed what, so status
            changes and edits are always traceable after the fact. */}
        {history.length > 0 && (
          <div style={{marginTop:14,borderTop:`1px solid ${T.cardBorder}`,paddingTop:10}}>
            <button onClick={()=>setHistoryOpen(v=>!v)}
              style={{background:"none",border:"none",padding:0,cursor:"pointer",display:"flex",alignItems:"center",gap:6,fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase"}}>
              <span style={{transition:"transform 0.15s",transform:historyOpen?"rotate(90deg)":"none",display:"inline-block"}}>▸</span>
              📜 ประวัติการแก้ไข ({history.length})
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

        {capWarn && (
          <div style={{marginTop:14,fontSize:12,color:T.red,background:T.redBg,border:`1px solid ${T.red}`,borderRadius:10,padding:"9px 12px",lineHeight:1.5}}>{capWarn}</div>
        )}
        <div style={{display:"flex",gap:10,marginTop:16,flexWrap:"wrap",alignItems:"center"}}>
          <button
            onClick={()=>{
              // ตรวจอีกครั้งก่อนปิด: ยอดของเข้าจริงรวมของทุกรายการห้ามเกินยอดสั่ง
              const bad = po.items.find(it => { const o=itemOrdered(it); const rc=(it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0); return o>0 && rc > o + 0.001; });
              if (bad) { setCapWarn(`⚠ ${bad.code||"รายการ"}: ยอดของเข้ารวมเกินยอดสั่ง ${fmt(itemOrdered(bad))} — แก้ให้ไม่เกินก่อนบันทึก`); return; }
              setCapWarn(""); onClose();
            }}
            disabled={locked} className="btn-primary"
            style={{background:locked?"#e2e8f0":T.green,color:locked?"#94a3b8":"#fff",cursor:locked?"not-allowed":"pointer"}}>{locked?"🔒":"💾"} บันทึก</button>
          {!locked && <button onClick={()=>onEdit(po)} className="btn-ghost" style={{fontSize:12}} title="แก้ผู้ขาย / หมวด / ยอดสั่ง">✏️ แก้ไข PO</button>}
          <button onClick={()=>onDelete(po.id)} disabled={locked} className="btn-ghost" style={{color:locked?"#cbd5e1":T.red,borderColor:locked?"#e2e8f0":T.red,cursor:locked?"not-allowed":"pointer"}}>🗑 ลบ</button>
          <div style={{flex:1}}/>
          <button onClick={onClose} className="btn-ghost">ปิด</button>
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
      style={{background:STATUS_BG[status],color:STATUS_CLR[status],fontSize:compact?11:12,padding:compact?"3px 8px":"5px 10px",
        borderRadius:20,fontWeight:600,border:`1px solid ${STATUS_CLR[status]}40`,cursor:disabled?"not-allowed":"pointer",outline:"none",
        opacity:disabled?0.65:1}} title={disabled?"รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin":undefined}>
      {PO_STATUS.map(s=><option key={s} value={s}>{s}</option>)}
    </select>
  );
}

// ─── Procurement View ─────────────────────────────────────────────────────────
function ProcurementView({ project, updateProject, tenderCosts, additions, poEntries, savePO, onBack, onHome, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[], onExport, setEditMode }) {
  const usdRate = effRate(project);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const [tab,    setTab]    = useState("list"); // "list" | "tracking"
  const [tabHist, setTabHist] = useState([]);   // ประวัติแท็บ — ปุ่มกลับย้อนทีละหน้า
  const goTab   = (id) => { if (id !== tab) { setTabHist(h => [...h, tab]); setTab(id); } };
  const [trackingOnlyIssues, setTrackingOnlyIssues] = useState(false); // lifted so the alert banner below can jump straight into "only late items"
  const [view,   setView]   = useState("browse"); // "browse" | "add"
  const blankItem = () => ({ id:uid(), code:"", store:"", amount:"",
    rounds:[{ id:uid(), planDate:"", planAmount:"", actualAmount:"", actualDate:"" }] });
  const emptyForm = () => ({
    date:new Date().toISOString().slice(0,10), status:"PO Issued",
    supplier:{ name:"", poNumber:"" },
    paymentType:"", creditDays:DEFAULT_CREDIT_DAYS, notes:"",
    items:[ blankItem() ],
  });
  const [form,   setForm]   = useState(emptyForm);
  const [editId, setEditId] = useState(null);
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [detailId, setDetailId] = useState(null);
  const [collapsed, setCollapsed] = useState({});
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
  const topLevelCodes = [
    ...ACCOUNTS.filter(a => !hiddenAccounts.includes(a.code)).map(a => a.code),
    ...extraItems.filter(e => !e.parentCode).map(e => e.code),
  ];
  const tenderTotal = topLevelCodes.reduce((s,c) => s + (parseFloat(combinedBudget[c]) || 0), 0);
  const totalComm   = poEntries.reduce((s,p)=>s+poTotal(p),0);
  const totalPaid   = poEntries.reduce((s,p)=>s+poRounds(p).filter(r=>roundPaid(p,r)).reduce((ss,r)=>ss+(parseFloat(r.actualAmount)||0),0),0);
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
  const itemNet = (it) => Math.max(budgetForCode(it.code) - (parseFloat(it.store)||0), 0);
  const setItemAmount = (id, val) => updateItemRow(id, "amount", val);
  // % ของยอดสั่ง = ช่องกรอกเอง (it.pct) ไม่ผูกกับมูลค่า PO อีกต่อไป

  const formTotal = form.items.reduce((s,it)=>s+(parseFloat(it.amount)||0),0);

  const submit = () => {
    // ชื่อ Supplier ไม่บังคับ — ใส่หรือไม่ใส่ก็ได้
    // กันมูลค่าติดลบ (ทำให้ยอดคงเหลือ/งบเพี้ยน)
    if (form.items.some(it=>it.code && (parseFloat(it.amount)||0) < 0)) { alert("มูลค่า PO ต้องไม่ติดลบ กรุณาแก้ไขก่อนบันทึก"); return; }
    const validItems = form.items.filter(it=>it.code && it.amount).map(it=>({
      id: it.id || uid(), code: it.code, store: it.store || "", amount: it.amount,
      rounds: (it.rounds && it.rounds.length ? it.rounds : [{id:uid()}]).map((r,idx)=>({
        id: r.id || uid(),
        planDate: r.planDate || "",
        planAmount: idx===0 ? (r.planAmount || it.amount) : (r.planAmount || ""),
        actualAmount: r.actualAmount || "",
        actualDate: r.actualDate || "",
      })),
    }));
    if (!validItems.length) { alert("กรุณาเลือก Account Code และกรอกมูลค่าอย่างน้อย 1 รายการ"); return; }
    const payload = {
      date: form.date, status: form.status, notes: form.notes || "",
      supplier: { name: form.supplier.name.trim(), poNumber: (form.supplier.poNumber||"").trim() },
      paymentType: form.paymentType || "", creditDays: parseInt(form.creditDays,10) || DEFAULT_CREDIT_DAYS,
      items: validItems,
    };

    const prev = editId ? poEntries.find(x=>x.id===editId) : null;
    const entries = [];
    if (prev && prev.status !== payload.status) entries.push(historyEntry(session, "status", `เปลี่ยนสถานะ: ${prev.status} → ${payload.status}`));
    entries.push(historyEntry(session, prev ? "edited" : "created", prev ? "แก้ไขข้อมูล PO" : "สร้างรายการ PO"));
    const withLog = { ...payload, history: [...entries.reverse(), ...(prev?.history||[])].slice(0,40) };

    savePO(editId ? poEntries.map(p=>p.id===editId?{...withLog,id:editId}:p) : [...poEntries,{...withLog,id:uid()}]);
    setEditId(null);
    setForm(emptyForm());
    setView("browse");
  };

  // Persist an in-place update to a PO's items/rounds (used by the detail view
  // when recording actual goods received or splitting a round). Migrates the
  // record to the new shape on first touch so it's normalised going forward.
  const updatePO = (updated) => savePO(poEntries.map(x=>x.id===updated.id?updated:x));

  // One-click status change — used by the StatusPicker wherever a PO is
  // listed, so procurement doesn't need to open the full edit form just to
  // move a PO from "PO Issued" to "Delivered". Still fully logged. Locked
  // once a PO is fully received + fully paid, unless the current user is admin.
  const changeStatus = (po, newStatus) => {
    if (newStatus === po.status) return;
    if (!canEditPO(po, session)) { alert("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ไขได้เฉพาะ Admin"); return; }
    const updated = withHistory({ ...po, status:newStatus }, historyEntry(session, "status", `เปลี่ยนสถานะ: ${po.status} → ${newStatus}`));
    savePO(poEntries.map(x=>x.id===po.id?updated:x));
  };

  const openEdit = (p) => {
    if (!canEditPO(p, session)) { alert("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ไขได้เฉพาะ Admin"); return; }
    const P = migratePO(p);
    setForm({
      date: P.date, status: P.status, notes: P.notes || "",
      supplier: { name: P.supplier.name || "", poNumber: P.supplier.poNumber || "" },
      paymentType: P.paymentType || "", creditDays: P.creditDays || DEFAULT_CREDIT_DAYS,
      items: P.items.map(it=>({
        id: it.id || uid(), code: it.code || "", store: it.store || "", amount: it.amount || "",
        rounds: (it.rounds && it.rounds.length ? it.rounds : [{id:uid(),planDate:"",planAmount:"",actualAmount:"",actualDate:""}])
          .map(r=>({ id:r.id||uid(), planDate:r.planDate||"", planAmount:r.planAmount||"", actualAmount:r.actualAmount||"", actualDate:r.actualDate||"" })),
      })),
    });
    setEditId(p.id); setView("add"); setDetailId(null);
  };
  const deletePO = (id) => {
    const po = poEntries.find(x=>x.id===id);
    if (po && !canEditPO(po, session)) { alert("PO นี้รับของและจ่ายเงินครบแล้ว — ลบได้เฉพาะ Admin"); return; }
    // ถามยืนยันทุกครั้งก่อนลบ (ลบแล้วย้อนกลับไม่ได้)
    const label = po ? `${poSupplierName(po)}${poNumbersLabel(po)!=="—"?` · ${poNumbersLabel(po)}`:""} · ${fmt(poTotal(po))} บาท` : "";
    if (!window.confirm(`ยืนยันการลบรายการ PO นี้?\n\n${label}\n\n⚠ ลบแล้วย้อนกลับไม่ได้`)) return;
    savePO(poEntries.filter(x=>x.id!==id)); setDetailId(null);
    if (editId === id) closeForm();   // ถ้าลบจากในฟอร์มแก้ไข ให้ปิดฟอร์มกลับหน้ารายการ
  };
  const closeForm = () => { setView("browse"); setEditId(null); setForm(emptyForm()); };
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
    const onEsc = (e) => { if (e.key === "Escape") { e.preventDefault(); closeForm(); } };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [view]);

  const filtered = poEntries.filter(p=>{
    const itemsText = poItems(p).map(it=>{ const acc=ACCOUNTS.find(a=>a.code===it.code); return `${it.code} ${acc?.name||""}`; }).join(" ");
    return (filter==="All"||p.status===filter)&&
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
    <Shell role="procurement" color={T.amber} project={project} onBack={backNav} onHome={onHome} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"24px 28px"}}>
        {view!=="add" && (
          <div style={{display:"flex",gap:8,marginBottom:20,alignItems:"center"}}>
            {[["list","📋 รายการ PO"],["tracking","🚚 ติดตามของเข้า/จ่ายเงิน"]].map(([id,label])=>(
              <button key={id} onClick={()=>goTab(id)}
                style={{background:tab===id?T.amber:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.amber:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
                {label}
              </button>
            ))}
            <div style={{marginLeft:"auto"}}><CurrencyControl project={project} updateProject={updateProject}/></div>
            <button onClick={onExport} className="btn-ghost" style={{display:"flex",alignItems:"center",gap:6,borderColor:T.amber,color:T.amber}}>
              ⬇️ Export Excel
            </button>
          </div>
        )}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
          <StatCard label="Budget (QS)" value={"฿"+fmt0(tenderTotal)} thb={tenderTotal} rate={usdRate} sub="เดิม + เพิ่มรายเดือนทุกเดือน" color={T.blue} icon="📋" accent={T.blueLight}/>
          <StatCard label="Committed (PO)" value={"฿"+fmt0(totalComm)} thb={totalComm} rate={usdRate} sub={`${poEntries.length} รายการ`} color={T.amber} icon="📦" accent={T.amberBg}/>
          <StatCard label="ชำระแล้ว" value={"฿"+fmt0(totalPaid)} thb={totalPaid} rate={usdRate} sub={`${paidCount} รายการ · จ่ายอัตโนมัติ`} color={T.green} icon="✅" accent={T.greenBg}/>
          <StatCard label="Budget คงเหลือ" value={"฿"+fmt0(tenderTotal-totalComm)} thb={tenderTotal-totalComm} rate={usdRate} sub={tenderTotal>0?`${((totalComm/tenderTotal)*100).toFixed(1)}% ใช้ไปแล้ว`:"—"} color={tenderTotal-totalComm<0?T.red:T.textSecondary} icon={tenderTotal-totalComm<0?"⚠️":"💰"} accent={tenderTotal-totalComm<0?T.redBg:"#f8fafc"}/>
        </div>

        {view!=="add" && (lateIncomingCount>0 || latePaymentCount>0) && (
          <div onClick={()=>{ goTab("tracking"); setTrackingOnlyIssues(true); }}
            style={{background:T.redBg,border:`1.5px solid #fecaca`,borderRadius:12,padding:"12px 18px",marginBottom:16,display:"flex",alignItems:"center",gap:12,cursor:"pointer"}}>
            <span style={{fontSize:20}}>⚠️</span>
            <div style={{flex:1}}>
              <div style={{fontSize:13,fontWeight:700,color:T.red}}>
                มีรายการที่ต้องรีบดู
                {lateIncomingCount>0 && <span> — ของเข้าล่าช้า {lateIncomingCount} รายการ</span>}
                {lateIncomingCount>0 && latePaymentCount>0 && <span>,</span>}
                {latePaymentCount>0 && <span> จ่ายเงินเกินกำหนด {latePaymentCount} รายการ</span>}
              </div>
              <div style={{fontSize:11.5,color:"#b91c1c",marginTop:1}}>คลิกเพื่อดูรายละเอียดทั้งหมด</div>
            </div>
            <span style={{fontSize:12,color:T.red,fontWeight:600,whiteSpace:"nowrap"}}>ดูรายการ →</span>
          </div>
        )}

        {view==="add" ? (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:28,maxWidth:680,animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:T.textPrimary}}>{editId?"แก้ไขรายการ PO":"เพิ่ม PO ใหม่"}</div>
                <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>กรอกข้อมูลคำสั่งซื้อ และแผนของเข้า/จ่ายเงิน</div>
              </div>
              <button onClick={closeForm} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>วันที่สั่ง PO</span>
                <input type="date" value={form.date} onChange={e=>setForm(f=>({...f, date:e.target.value}))} className="input-base"/>
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>สถานะ</span>
                <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} className="input-base">
                  {PO_STATUS.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
              </label>

              {/* Supplier — exactly one vendor per PO. */}
              <div style={{gridColumn:"1/-1",display:"flex",alignItems:"center",gap:8,marginTop:6,paddingTop:14,borderTop:`1px dashed ${T.cardBorder}`}}>
                <span style={{fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase"}}>🏢 Supplier (ไม่บังคับ · หนึ่งเจ้าต่อ PO)</span>
              </div>
              <div style={{gridColumn:"1/-1",display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                <input placeholder="ชื่อ Supplier (ถ้ามี)" value={form.supplier.name} onChange={e=>updateSupplierField("name",e.target.value)} className="input-base"/>
                <input placeholder="เลข PO" value={form.supplier.poNumber} onChange={e=>updateSupplierField("poNumber",e.target.value)} className="input-base"/>
              </div>

              {/* Account-code line items — each carries its own store amount and
                  its % of the net-to-purchase (budget − store). */}
              <div style={{gridColumn:"1/-1",display:"flex",alignItems:"center",gap:8,marginTop:6,paddingTop:14,borderTop:`1px dashed ${T.cardBorder}`}}>
                <span style={{fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase"}}>📐 หมวดต้นทุน * (กรอกของใน store และ % ของยอดสั่ง)</span>
              </div>
              <div style={{gridColumn:"1/-1",display:"flex",flexDirection:"column",gap:12}}>
                {form.items.map((it)=>{
                  const budget = budgetForCode(it.code);
                  const net = itemNet(it);
                  const amt = parseFloat(it.amount)||0;
                  const pct = net>0 ? Math.round(amt/net*100) : 0;
                  const prevOrdered = poEntries.reduce((s,p)=> p.id===editId ? s : s + poAmountForCode(p, it.code), 0);
                  const cumPct = net>0 ? Math.round((prevOrdered+amt)/net*100) : 0;
                  return (
                  <div key={it.id} style={{border:`1px solid ${T.cardBorder}`,borderRadius:12,padding:14,background:T.bg}}>
                    <div style={{display:"grid",gridTemplateColumns:"1fr auto",gap:8,alignItems:"center",marginBottom:12}}>
                      <select value={it.code} onChange={e=>updateItemRow(it.id,"code",e.target.value)} className="input-base">
                        <option value="">— เลือก Account Code —</option>
                        {ACCOUNTS.map(a=><option key={a.code} value={a.code}>{a.code} · {a.name}</option>)}
                      </select>
                      <button type="button" onClick={()=>removeItemRow(it.id)} disabled={form.items.length===1}
                        style={{background:"none",border:"none",color:form.items.length===1?T.textMuted:T.red,cursor:form.items.length===1?"default":"pointer",padding:"4px 8px",fontSize:15,opacity:form.items.length===1?0.4:1}}>🗑</button>
                    </div>
                    {/* แถวบน: งบ (อ่านอย่างเดียว) · store · ต้องสั่งสุทธิ (อ่านอย่างเดียว) */}
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:10}}>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>งบโครงการ</span>
                        <input className="input-base" readOnly tabIndex={-1} value={fmtMoneyInput(budget)}
                          style={{textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontWeight:600,background:T.card,color:T.textPrimary}}/>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>มีใน store</span>
                        <MoneyInput value={it.store} onChange={v=>updateItemRow(it.id,"store",v)}/>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.amber,fontWeight:500}}>ต้องสั่งสุทธิ</span>
                        <input className="input-base" readOnly tabIndex={-1} value={fmtMoneyInput(net)}
                          style={{textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontWeight:600,background:T.amberBg,color:T.amber,borderColor:"transparent"}}/>
                      </label>
                    </div>
                    {/* แถวล่าง: มูลค่า PO · % · แผนของเข้า — ความสูงเท่ากันหมด */}
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10}}>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>มูลค่า PO นี้ (THB)</span>
                        <MoneyInput value={it.amount} onChange={v=>setItemAmount(it.id,v)}/>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>% ของยอดสั่ง (กรอกเอง)</span>
                        <div style={{position:"relative",display:"flex",alignItems:"center"}}>
                          <input type="number" placeholder="0" value={it.pct ?? ""} onChange={e=>updateItemRow(it.id,"pct",e.target.value)}
                            className="input-base" style={{textAlign:"right",fontFamily:"'JetBrains Mono',monospace",flex:1,paddingRight:26}}/>
                          <span style={{position:"absolute",right:11,fontSize:13,color:(it.pct??"")!==""?T.textPrimary:T.textMuted,fontWeight:600,pointerEvents:"none"}}>%</span>
                        </div>
                      </label>
                      <label style={{display:"flex",flexDirection:"column",gap:5}}>
                        <span style={{fontSize:11,color:T.textSecondary,fontWeight:500}}>แผนของเข้า (งวดแรก)</span>
                        <input type="date" value={it.rounds?.[0]?.planDate||""} onChange={e=>updateItemPlan(it.id,"planDate",e.target.value)} className="input-base"/>
                      </label>
                    </div>
                    {it.code && ((it.pct??"")!=="" || amt>0) && (
                      <div style={{marginTop:10,fontSize:11,color:T.textSecondary}}>
                        % ของยอดสั่ง PO นี้: <b style={{color:(parseFloat(it.pct)||0)>100?T.red:T.textPrimary,fontSize:12}}>{(it.pct??"")!=="" ? `${it.pct}%` : "—"}</b> <span style={{color:T.textMuted}}>(ที่กรอกเอง)</span>
                        {amt>0 && <span style={{color:T.textMuted}}> · ยอดจริง {fmt(amt)} = {budget>0?Math.round(amt/budget*100):0}% ของงบรวม</span>}
                      </div>
                    )}
                  </div>
                  );
                })}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <button type="button" onClick={addItemRow} className="btn-ghost" style={{padding:"6px 12px",fontSize:12}}>+ เพิ่ม Account Code</button>
                  {form.items.length>1 && <span style={{fontSize:12,color:T.textSecondary}}>รวม: <b style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace"}}>{fmt(formTotal)}</b></span>}
                </div>
              </div>

              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>วิธีจ่ายเงิน</span>
                <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                  <button type="button" onClick={()=>setForm(f=>({...f,paymentType:"cash"}))}
                    style={{flex:"1 1 160px",padding:"10px 14px",borderRadius:10,border:`1.5px solid ${form.paymentType==="cash"?T.green:T.cardBorder}`,background:form.paymentType==="cash"?T.greenBg:T.card,color:form.paymentType==="cash"?T.green:T.textSecondary,fontSize:13,fontWeight:600,cursor:"pointer",transition:"all 0.15s"}}>
                    💵 เงินสด <span style={{fontWeight:400,fontSize:11,opacity:0.8}}>(จ่ายวันของเข้า)</span>
                  </button>
                  <button type="button" onClick={()=>setForm(f=>({...f,paymentType:"credit",creditDays:f.creditDays||DEFAULT_CREDIT_DAYS}))}
                    style={{flex:"1 1 160px",padding:"10px 14px",borderRadius:10,border:`1.5px solid ${form.paymentType==="credit"?T.blue:T.cardBorder}`,background:form.paymentType==="credit"?T.blueLight:T.card,color:form.paymentType==="credit"?T.blue:T.textSecondary,fontSize:13,fontWeight:600,cursor:"pointer",transition:"all 0.15s"}}>
                    💳 เครดิต
                  </button>
                  {form.paymentType==="credit" && (
                    <span style={{display:"flex",alignItems:"center",gap:6}}>
                      <input type="number" value={form.creditDays} onChange={e=>setForm(f=>({...f,creditDays:e.target.value}))} className="input-base" style={{width:76}}/>
                      <span style={{fontSize:12,color:T.textSecondary}}>วัน</span>
                    </span>
                  )}
                </div>
                {form.paymentType && (
                  <span style={{fontSize:11,color:T.blue}}>วันครบกำหนดจ่ายคำนวณอัตโนมัติจาก "วันของเข้าจริง" ของแต่ละงวด{form.paymentType==="credit"?` + ${form.creditDays||DEFAULT_CREDIT_DAYS} วัน`:""} — จ่ายอัตโนมัติเมื่อถึงกำหนด</span>
                )}
              </label>

              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>หมายเหตุ</span>
                <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} className="input-base" style={{resize:"vertical"}}/>
              </label>
            </div>
            {/* เว้นที่ด้านล่างให้พ้นแถบ "ย้อนกลับ/ทำซ้ำ" ที่ลอยมุมซ้ายล่าง ไม่ให้ทับปุ่ม */}
            <div style={{display:"flex",gap:10,marginTop:20,marginBottom:76,flexWrap:"wrap",alignItems:"center"}}>
              <button onClick={submit} className="btn-primary" style={{background:T.amber,color:"#fff"}}>{editId?"บันทึก":"เพิ่ม PO"}</button>
              <button onClick={closeForm} className="btn-ghost">ยกเลิก</button>
              {editId && (
                <button onClick={()=>{ deletePO(editId); }}
                  style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:6,background:T.redBg,border:`1px solid #fecaca`,color:T.red,borderRadius:10,padding:"9px 16px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
                  🗑 ลบ PO นี้
                </button>
              )}
            </div>
          </div>
        ) : tab==="tracking" ? (
          <ProcurementTrackingTab poEntries={poEntries} onEdit={openEdit} onView={openDetail} onAddNew={()=>setView("add")}
            onStatusChange={changeStatus} session={session} usdRate={usdRate}
            onlyIssues={trackingOnlyIssues} setOnlyIssues={setTrackingOnlyIssues} />
        ) : (
          <>
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <SearchInput value={search} onChange={setSearch} placeholder="🔍 ค้นหา Account, supplier, PO..." width={240}/>
              <div style={{display:"flex",gap:5,flex:1,flexWrap:"wrap"}}>
                {["All",...PO_STATUS].map(s=>(
                  <button key={s} onClick={()=>setFilter(s)}
                    style={{background:filter===s?T.amber:"transparent",border:`1.5px solid ${filter===s?T.amber:T.cardBorder}`,borderRadius:8,padding:"4px 11px",color:filter===s?"#fff":T.textSecondary,fontSize:11,cursor:"pointer",fontWeight:500,transition:"all 0.15s"}}>{s}</button>
                ))}
              </div>
              {filtered.length>0 && (
                <button onClick={()=>{
                    const allCollapsed = Object.keys(groupTotals).every(c=>collapsed[c]);
                    const next = {}; Object.keys(groupTotals).forEach(c=>{ next[c] = !allCollapsed; });
                    setCollapsed(next);
                  }}
                  className="btn-ghost" style={{padding:"7px 14px",fontSize:12}}>
                  {Object.keys(groupTotals).length>0 && Object.keys(groupTotals).every(c=>collapsed[c]) ? "⬇️ ขยายทั้งหมด" : "⬆️ ย่อทั้งหมด"}
                </button>
              )}
              <button onClick={()=>setView("add")} className="btn-primary" style={{background:T.amber}}>+ เพิ่ม PO</button>
            </div>

            {filtered.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
                <div style={{fontSize:32,marginBottom:12}}>📋</div>
                <div style={{fontSize:14,fontWeight:500,color:T.textSecondary,marginBottom:6}}>{poEntries.length===0?"ยังไม่มีรายการ":"ไม่พบรายการที่ตรงเงื่อนไข"}</div>
                <div style={{fontSize:12}}>{poEntries.length===0?'กด "+ เพิ่ม PO" เพื่อเริ่มต้น':"ลองล้างตัวกรอง หรือคำค้นหา"}</div>
              </div>
            ) : (
              <div style={{display:"flex",flexDirection:"column",gap:14}}>
                {sortedGroupCodes.map(code => {
                  const acc  = ACCOUNTS.find(a=>a.code===code);
                  const rows = groupedFiltered[code].slice().sort((a,b)=> (b.po.date||"").localeCompare(a.po.date||""));
                  const isCollapsed = !!collapsed[code];
                  const groupTotal = rows.reduce((s,{item})=>s+(parseFloat(item.amount)||0),0);
                  return (
                    <div key={code} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                      <div onClick={()=>toggleGroup(code)} style={{padding:"12px 18px",background:"#f8fafc",borderBottom:isCollapsed?"none":`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:10,cursor:"pointer"}}>
                        <span style={{color:T.textMuted,fontSize:11,transition:"transform 0.15s",transform:isCollapsed?"rotate(-90deg)":"none"}}>▾</span>
                        <span style={{color:T.blue,fontSize:12,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{code}</span>
                        <span style={{color:T.textPrimary,fontSize:13,fontWeight:600}}>{acc?.name || "—"}</span>
                        <span style={{flex:1}}/>
                        <span style={{color:T.textMuted,fontSize:11}}>{rows.length} รายการ</span>
                        <span style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:13}}>{fmt(groupTotal)}{usdRate>0 && <span className="usd-sub" style={{color:T.green,fontWeight:700,fontSize:12,marginLeft:6}}>≈ ${fmt(groupTotal/usdRate)}</span>}</span>
                      </div>
                      {!isCollapsed && (
                        <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
                          <thead>
                            <tr>
                              {["วันเปิด PO","Supplier","PO No.","มูลค่า (THB)","วันรับของ","วันจ่าย","การส่งของ / จ่ายเงิน","สถานะ",""].map(h=>(
                                <th key={h} style={{padding:"9px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:10,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {rows.map(({po:p,item},i)=>{
                              const splitAcrossCodes = poItems(p).length>1;
                              const inc = incomingStatus(p), pay = paymentStatus(p);
                              const locked = !canEditPO(p, session);
                              const receivedDates = poReceivedDates(p);
                              const paidDate = poPaidDate(p);
                              return (
                              <tr key={p.id+"-"+(item.id||item.code)} onClick={()=>openDetail(p)}
                                style={{background:i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`,cursor:"pointer"}}
                                onMouseEnter={e=>e.currentTarget.style.background="#fef9ec"}
                                onMouseLeave={e=>e.currentTarget.style.background=i%2===0?T.card:"#fafbfd"}>
                                <td style={{padding:"10px 16px",color:T.textMuted,fontSize:12,fontFamily:"'JetBrains Mono',monospace"}}>{p.date}</td>
                                <td style={{padding:"10px 16px",color:T.textPrimary,fontWeight:500}}>{itemSupplierName(p,item)}</td>
                                <td style={{padding:"10px 16px",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{poNumbersLabel(p)}</td>
                                <td style={{padding:"10px 16px",textAlign:"right"}}>
                                  <div style={{color:T.textPrimary,fontFamily:"'JetBrains Mono',monospace",fontWeight:600}}>{fmt(item.amount)}</div>
                                  {usdLine(parseFloat(item.amount)||0, usdRate)}
                                  {splitAcrossCodes && <div style={{fontSize:10,color:T.textMuted}}>จาก {poItems(p).length} รหัส · รวม {fmt(poTotal(p))}</div>}
                                </td>
                                <td style={{padding:"10px 16px",fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:receivedDates.length?T.textPrimary:T.textMuted}}>
                                  {receivedDates.length===0 ? "—" : receivedDates.length===1 ? receivedDates[0] : `${receivedDates[0]} (+${receivedDates.length-1})`}
                                </td>
                                <td style={{padding:"10px 16px",fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:paidDate?T.green:T.textMuted,fontWeight:paidDate?600:400}}>
                                  {paidDate || "—"}
                                </td>
                                <td style={{padding:"10px 16px"}}>
                                  <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                                    <span style={{background:INCOMING_BG[inc],color:INCOMING_CLR[inc],fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{INCOMING_LABEL[inc]}</span>
                                    {p.status!=="Paid" && (
                                      <span style={{background:PAYMENT_BG[pay],color:PAYMENT_CLR[pay],fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{PAYMENT_LABEL[pay]}</span>
                                    )}
                                    {p.paymentType && (
                                      <span style={{background:PAYMENT_TYPE_BG[p.paymentType],color:PAYMENT_TYPE_CLR[p.paymentType],fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{PAYMENT_TYPE_ICON[p.paymentType]} {paymentTypeLabel(p)}</span>
                                    )}
                                  </div>
                                </td>
                                <td style={{padding:"10px 16px"}}>
                                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                                    <StatusPicker status={p.status} onChange={s=>changeStatus(p,s)} disabled={locked} compact/>
                                    {locked && <span title="รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin" style={{fontSize:11}}>🔒</span>}
                                  </div>
                                  {poLastUpdate(p) && <div style={{fontSize:9,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>อัปเดต {relativeTime(poLastUpdate(p).at)} · {poLastUpdate(p).user}</div>}
                                </td>
                                <td style={{padding:"10px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                                  <button onClick={()=>openEdit(p)} disabled={locked} title={locked?"แก้ไขได้เฉพาะ Admin":"แก้ไข (ลบได้ในหน้านี้)"}
                                    style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textMuted,cursor:locked?"not-allowed":"pointer",padding:"2px 6px",borderRadius:6}}>✏️</button>
                                </td>
                              </tr>
                            );})}
                          </tbody>
                        </table></div>
                      )}
                    </div>
                  );
                })}
                <div style={{display:"flex",justifyContent:"flex-end",gap:16,padding:"4px 18px",color:T.textMuted,fontSize:12}}>
                  <span>{filtered.length} รายการทั้งหมด</span>
                  <span style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{fmt(filtered.reduce((s,p)=>s+poTotal(p),0))}{usdRate>0 && <span className="usd-sub" style={{color:T.green,fontWeight:700,fontSize:12,marginLeft:6}}>≈ ${fmt(filtered.reduce((s,p)=>s+poTotal(p),0)/usdRate)}</span>}</span>
                </div>
              </div>
            )}
          </>
        )}
      </div>
      <PODetailModal po={detailPO} onClose={closeDetail} onEdit={openEdit} onDelete={deletePO} onStatusChange={changeStatus} onChangePO={updatePO} session={session} usdRate={usdRate} />
    </Shell>
  );
}

// ─── Procurement: Incoming / Payment Tracking tab ─────────────────────────────
// Groups every PO by its Account Code so the team can see, at a glance and per
// cost line, which deliveries and payments are on track vs. overdue.
function ProcurementTrackingTab({ poEntries, onEdit, onView, onAddNew, onlyIssues, setOnlyIssues, onStatusChange, session, usdRate=0 }) {
  const [search, setSearch] = useState("");
  // Which Account-Code groups are collapsed — lets a busy board with many
  // PO rows be tidied away group by group instead of scrolling forever.
  const [collapsed, setCollapsed] = useState(() => new Set());
  const toggleGroup = (code) => setCollapsed(prev => {
    const next = new Set(prev);
    next.has(code) ? next.delete(code) : next.add(code);
    return next;
  });

  const counts = poEntries.reduce((acc,p) => {
    const inc = incomingStatus(p), pay = paymentStatus(p);
    if (inc==="pending") acc.incPending++;
    if (inc==="late")    acc.incLate++;
    if (pay==="pending") acc.payPending++;
    if (pay==="paid")    acc.payPaid++;
    return acc;
  }, { incPending:0, incLate:0, payPending:0, payPaid:0 });

  const q = search.toLowerCase();
  const passesFilter = (p) => {
    const itemsText = poItems(p).map(it=>{ const acc=ACCOUNTS.find(a=>a.code===it.code); return `${it.code} ${acc?.name||""}`; }).join(" ");
    const matchesSearch = q==="" || [itemsText,poSupplierText(p),poNumbersLabel(p)].join(" ").toLowerCase().includes(q);
    const hasIssue = incomingStatus(p)==="late" || paymentStatus(p)==="late";
    return matchesSearch && (!onlyIssues || hasIssue);
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

  const DateCell = ({ value, lateTint }) => (
    <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:value?(lateTint?T.red:T.textPrimary):T.textMuted,fontWeight:value&&lateTint?700:400}}>
      {value || "—"}
    </span>
  );
  const Badge = ({ text, clr, bg }) => (
    <span style={{background:bg,color:clr,fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{text}</span>
  );
  // Renders every delivery batch on a PO — one line per shipment, so a PO
  // that arrives in 2-3 batches shows each plan → actual date with its own status.
  const DeliveryList = ({ po }) => {
    const deliveries = poDeliveries(po);
    const multiSupplier = poSuppliers(po).length > 1;
    if (!deliveries.length) return <span style={{fontSize:12,color:T.textMuted}}>—</span>;
    return (
      <div style={{display:"flex",flexDirection:"column",gap:3}}>
        {deliveries.map((d,i)=>{
          const st = deliveryStatus(d);
          return (
            <div key={d.id||i} style={{display:"flex",alignItems:"center",gap:5}}>
              {deliveries.length>1 && <span style={{fontSize:10,color:T.textMuted,fontWeight:700,minWidth:14}}>#{i+1}</span>}
              {multiSupplier && <span style={{fontSize:10,color:T.textSecondary,fontWeight:600,whiteSpace:"nowrap"}}>{d.supplierName||"—"}:</span>}
              <DateCell value={d.plan} lateTint={st==="late"}/>
              <span style={{color:T.textMuted,fontSize:11}}>→</span>
              <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:st==="received"?T.green:T.textMuted,fontWeight:st==="received"?600:400}}>{d.actual||"รอ"}</span>
              {d.amount && <span style={{fontSize:10,color:T.textMuted,fontFamily:"'JetBrains Mono',monospace"}}>({fmt(d.amount)})</span>}
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div>
      {/* Quick-glance counts so problems surface without reading every row */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
        <StatCard label="รอของเข้า" value={counts.incPending} sub="ยังไม่ถึงวันที่นัด" color={T.amber} icon="⏳" accent={T.amberBg}/>
        <StatCard label="ของเข้าล่าช้า" value={counts.incLate} sub="เลยวันแผนของเข้าแล้ว" color={T.red} icon="⚠️" accent={T.redBg}/>
        <StatCard label="รอจ่ายเงิน" value={counts.payPending} sub="ของเข้าแล้ว รอถึงกำหนด" color={T.amber} icon="⏳" accent={T.amberBg}/>
        <StatCard label="ถึงกำหนดจ่าย" value={counts.payPaid} sub="ถึงวันครบกำหนดแล้ว (ระบบทำเครื่องหมายให้)" color={T.green} icon="✅" accent={T.greenBg}/>
      </div>

      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <SearchInput value={search} onChange={setSearch} placeholder="🔍 ค้นหา Acc. Code, supplier, PO..." width={240}/>
        <button onClick={()=>setOnlyIssues(v=>!v)}
          style={{background:onlyIssues?T.red:"transparent",border:`1.5px solid ${onlyIssues?T.red:T.cardBorder}`,borderRadius:8,padding:"7px 14px",color:onlyIssues?"#fff":T.textSecondary,fontSize:12,cursor:"pointer",fontWeight:600}}>
          ⚠️ แสดงเฉพาะรายการล่าช้า
        </button>
        <button onClick={()=>setCollapsed(new Set(sortedCodes))}
          style={{background:"transparent",border:`1.5px solid ${T.cardBorder}`,borderRadius:8,padding:"7px 14px",color:T.textSecondary,fontSize:12,cursor:"pointer",fontWeight:600}}>
          ▲ ย่อทั้งหมด
        </button>
        <button onClick={()=>setCollapsed(new Set())}
          style={{background:"transparent",border:`1.5px solid ${T.cardBorder}`,borderRadius:8,padding:"7px 14px",color:T.textSecondary,fontSize:12,cursor:"pointer",fontWeight:600}}>
          ▼ ขยายทั้งหมด
        </button>
        <div style={{flex:1}}/>
        <button onClick={onAddNew} className="btn-primary" style={{background:T.amber}}>+ เพิ่ม PO</button>
      </div>

      {sortedCodes.length===0 ? (
        <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
          <div style={{fontSize:32,marginBottom:12}}>🚚</div>
          <div style={{fontSize:14,fontWeight:500,color:T.textSecondary,marginBottom:6}}>ไม่พบรายการที่ตรงเงื่อนไข</div>
          <div style={{fontSize:12}}>ลองล้างตัวกรอง หรือคำค้นหา</div>
        </div>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {sortedCodes.map(code => {
            const acc = ACCOUNTS.find(a=>a.code===code);
            const rows = groups[code];
            const lateCount = rows.filter(({po:p})=>incomingStatus(p)==="late"||paymentStatus(p)==="late").length;
            const isCollapsed = collapsed.has(code);
            return (
              <div key={code} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                <div onClick={()=>toggleGroup(code)}
                  style={{padding:"12px 18px",background:"#f8fafc",borderBottom:isCollapsed?"none":`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:10,cursor:"pointer",userSelect:"none"}}>
                  <span style={{fontSize:11,color:T.textMuted,transform:isCollapsed?"rotate(-90deg)":"none",transition:"transform 0.15s",display:"inline-block",width:12}}>▼</span>
                  <span style={{color:T.blue,fontSize:12,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{code}</span>
                  <span style={{color:T.textPrimary,fontSize:13,fontWeight:600}}>{acc?.name || "—"}</span>
                  <span style={{flex:1}}/>
                  <span style={{color:T.textMuted,fontSize:11}}>{rows.length} PO</span>
                  {lateCount>0 && <Badge text={`⚠️ ${lateCount} ล่าช้า`} clr={T.red} bg={T.redBg}/>}
                </div>
                {!isCollapsed && (
                <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr>
                      {["วันเปิด PO","Supplier","PO No.","มูลค่า (THB)","วันรับของ","วันจ่าย","การส่งของ","แผนจ่ายเงิน","ติดตาม","สถานะ",""].map(h=>(
                        <th key={h} style={{padding:"9px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:10,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(({po:p,item},i) => {
                      const inc = incomingStatus(p), pay = paymentStatus(p);
                      const splitAcrossCodes = poItems(p).length>1;
                      const locked = !canEditPO(p, session);
                      const receivedDates = poReceivedDates(p);
                      const paidDate = poPaidDate(p);
                      return (
                        <tr key={p.id+"-"+(item.id||item.code)} onClick={()=>onView?.(p)}
                          style={{background:i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`,cursor:onView?"pointer":"default"}}
                          onMouseEnter={e=>e.currentTarget.style.background="#fef9ec"}
                          onMouseLeave={e=>e.currentTarget.style.background=i%2===0?T.card:"#fafbfd"}>
                          <td style={{padding:"9px 16px",color:T.textMuted,fontSize:12,fontFamily:"'JetBrains Mono',monospace"}}>{p.date}</td>
                          <td style={{padding:"9px 16px",color:T.textPrimary,fontWeight:500}}>{itemSupplierName(p,item)}</td>
                          <td style={{padding:"9px 16px",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{poNumbersLabel(p)}</td>
                          <td style={{padding:"9px 16px",textAlign:"right"}}>
                            <div style={{color:T.textPrimary,fontFamily:"'JetBrains Mono',monospace",fontWeight:600}}>{fmt(item.amount)}</div>
                            {usdLine(parseFloat(item.amount)||0, usdRate)}
                            {splitAcrossCodes && <div style={{fontSize:10,color:T.textMuted}}>รวม {fmt(poTotal(p))}</div>}
                          </td>
                          <td style={{padding:"9px 16px",fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:receivedDates.length?T.textPrimary:T.textMuted}}>
                            {receivedDates.length===0 ? "—" : receivedDates.length===1 ? receivedDates[0] : `${receivedDates[0]} (+${receivedDates.length-1})`}
                          </td>
                          <td style={{padding:"9px 16px",fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:paidDate?T.green:T.textMuted,fontWeight:paidDate?600:400}}>
                            {paidDate || "—"}
                          </td>
                          <td style={{padding:"9px 16px"}}><DeliveryList po={p}/></td>
                          <td style={{padding:"9px 16px"}}>
                            <DateCell value={poNextDueDate(p)} lateTint={false}/>
                            {p.paymentType && (
                              <div style={{marginTop:3}}>
                                <Badge text={`${PAYMENT_TYPE_ICON[p.paymentType]} ${paymentTypeLabel(p)}`} clr={PAYMENT_TYPE_CLR[p.paymentType]} bg={PAYMENT_TYPE_BG[p.paymentType]}/>
                              </div>
                            )}
                          </td>
                          <td style={{padding:"9px 16px"}}>
                            <div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"flex-start"}}>
                              <Badge text={INCOMING_LABEL[inc]} clr={INCOMING_CLR[inc]} bg={INCOMING_BG[inc]}/>
                              <Badge text={PAYMENT_LABEL[pay]} clr={PAYMENT_CLR[pay]} bg={PAYMENT_BG[pay]}/>
                            </div>
                          </td>
                          <td style={{padding:"9px 16px"}} onClick={e=>e.stopPropagation()}>
                            <div style={{display:"flex",alignItems:"center",gap:4}}>
                              <StatusPicker status={p.status} onChange={s=>onStatusChange?.(p,s)} disabled={locked} compact/>
                              {locked && <span title="รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin" style={{fontSize:11}}>🔒</span>}
                            </div>
                            {poLastUpdate(p) && <div style={{fontSize:9,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>อัปเดต {relativeTime(poLastUpdate(p).at)}</div>}
                          </td>
                          <td style={{padding:"9px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                            <button onClick={()=>onEdit(p)} disabled={locked} title={locked?"แก้ไขได้เฉพาะ Admin":"แก้ไข"}
                              style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textMuted,cursor:locked?"not-allowed":"pointer",padding:"2px 6px",borderRadius:6}}>✏️</button>
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

// ─── Accounting: ตารางรวมรายเดือน (Cost + Incoming / Actual / Payment) ──────────
//  ต่อ Acc. Code: Updated Tender Cost (งบ), Balance Pending PO (งบ − ผูกพัน),
//  Stock (ของที่รับแล้ว), Balance Cost (Balance Pending PO − Stock) แล้วตามด้วย 3
//  กลุ่มเดือน — Incoming Plan (วันแผนรับของ), Actual Received (วันรับจริง),
//  Payment Plan (วันครบกำหนดจ่าย) — โชว์เฉพาะเดือนที่มีข้อมูล + TOTAL แต่ละกลุ่ม.
const MATRIX_EN_MONTH = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
function AccountingMatrixTab({ tenderCosts, additions, poEntries, extraItems, hiddenAccounts }) {
  const accounts = exportAccountList(extraItems, hiddenAccounts);
  const combined = buildCombinedBudget(tenderCosts, additions);
  const committedByCode = {}, stockByCode = {};
  const incoming = {}, actual = {}, payplan = {};
  const bump = (obj, code, mk, amt) => { if (!mk || !amt) return; (obj[code] = obj[code] || {}); obj[code][mk] = (obj[code][mk] || 0) + amt; };
  poEntries.forEach(p => {
    poItems(p).forEach(it => {
      const code = it.code;
      committedByCode[code] = (committedByCode[code] || 0) + (parseFloat(it.amount) || 0);
      (it.rounds || []).forEach(r => {
        const planAmt = parseFloat(r.planAmount) || 0;
        if (r.planDate && planAmt) bump(incoming, code, r.planDate.slice(0, 7), planAmt);
        if (roundReceived(r)) {
          const amt = parseFloat(r.actualAmount) || 0;
          bump(actual, code, r.actualDate.slice(0, 7), amt);
          stockByCode[code] = (stockByCode[code] || 0) + amt;
        }
      });
    });
    poPayLines(p).forEach(l => bump(payplan, l.code, l.month, l.amount || 0));
  });
  const monthsOf = (obj) => [...new Set(Object.values(obj).flatMap(m => Object.keys(m)))].sort();
  const inM = monthsOf(incoming), acM = monthsOf(actual), payM = monthsOf(payplan);
  const lbl = (mk) => { const [y, m] = mk.split("-"); return `${MATRIX_EN_MONTH[(+m) - 1]} ${String(y).slice(2)}`; };
  const money = (n) => !n ? "-" : (n < 0 ? `(${fmt(Math.abs(n))})` : fmt(n));

  const rows = accounts.map(a => {
    const budget = parseFloat(combined[a.code]) || 0;
    const committed = committedByCode[a.code] || 0;
    const balPO = budget - committed;
    const stock = stockByCode[a.code] || 0;
    const balCost = balPO - stock;
    const inRow = inM.map(mk => incoming[a.code]?.[mk] || 0);
    const acRow = acM.map(mk => actual[a.code]?.[mk] || 0);
    const pyRow = payM.map(mk => payplan[a.code]?.[mk] || 0);
    const sum = (arr) => arr.reduce((s, x) => s + x, 0);
    return { a, budget, committed, balPO, stock, balCost, inRow, acRow, pyRow, inTot: sum(inRow), acTot: sum(acRow), pyTot: sum(pyRow) };
  }).filter(r => r.budget || r.committed || r.stock || r.inTot || r.acTot || r.pyTot);

  const colSum = (pick, i) => rows.reduce((s, r) => s + (pick(r)[i] || 0), 0);
  const totOf = (pick) => rows.reduce((s, r) => s + pick(r), 0);

  // styles
  const bCost = "#f4e9ef", bIn = "#eaf4ea", bAc = "#e8eff9", bPy = "#fdf1e2";
  const cell = { border: "1px solid #d9e0ea", padding: "5px 9px", fontSize: 11.5, whiteSpace: "nowrap" };
  const num  = { ...cell, textAlign: "right", fontFamily: "'JetBrains Mono',monospace" };
  const hCell = (bg) => ({ ...cell, background: bg, fontWeight: 700, color: T.textSecondary, textAlign: "center", position: "sticky", top: 0 });
  const numCell = (v, extraBg) => (
    <td style={{ ...num, background: extraBg, color: v < 0 ? T.red : (v ? T.textPrimary : T.textMuted), fontWeight: v ? 500 : 400 }}>{money(v)}</td>
  );

  return (
    <div>
      <div style={{ fontSize: 12.5, color: T.textMuted, marginBottom: 12 }}>
        โชว์เฉพาะเดือนที่มีข้อมูล · รวมยอดต่อ Acc. Code ต่อเดือน · Stock = ของที่รับแล้ว · Balance Cost = Balance Pending PO − Stock
      </div>
      <div style={{ overflow: "auto", border: `1px solid ${T.cardBorder}`, borderRadius: 12 }}>
        <table style={{ borderCollapse: "collapse", width: "max-content", minWidth: "100%" }}>
          <thead>
            {/* กลุ่ม */}
            <tr>
              <th style={hCell("#f8fafc")} colSpan={2}></th>
              <th style={hCell(bCost)}>A</th>
              <th style={hCell(bCost)}>A-B</th>
              <th style={hCell(bCost)}></th>
              <th style={hCell(bCost)}></th>
              <th style={hCell(bIn)} colSpan={inM.length + 1}>Incoming Plan</th>
              <th style={hCell(bAc)} colSpan={acM.length + 1}>Actual Received</th>
              <th style={hCell(bPy)} colSpan={payM.length + 1}>Payment Plan</th>
            </tr>
            {/* หัวคอลัมน์ */}
            <tr>
              <th style={{ ...hCell("#f1f5f9"), textAlign: "left", minWidth: 70 }}>Acc. Code</th>
              <th style={{ ...hCell("#f1f5f9"), textAlign: "left", minWidth: 190 }}>Acc. Name</th>
              <th style={{ ...hCell(bCost), minWidth: 130 }}>Updated Tender Cost</th>
              <th style={{ ...hCell(bCost), minWidth: 120 }}>Balance Pending PO</th>
              <th style={{ ...hCell(bCost), minWidth: 90 }}>Stock</th>
              <th style={{ ...hCell(bCost), minWidth: 100 }}>Balance Cost</th>
              {inM.map(mk => <th key={"i" + mk} style={hCell(bIn)}>{lbl(mk)}</th>)}
              <th style={{ ...hCell(bIn), fontWeight: 800 }}>TOTAL</th>
              {acM.map(mk => <th key={"a" + mk} style={hCell(bAc)}>{lbl(mk)}</th>)}
              <th style={{ ...hCell(bAc), fontWeight: 800 }}>TOTAL</th>
              {payM.map(mk => <th key={"p" + mk} style={hCell(bPy)}>{lbl(mk)}</th>)}
              <th style={{ ...hCell(bPy), fontWeight: 800 }}>TOTAL</th>
            </tr>
          </thead>
          <tbody>
            {rows.map(r => (
              <tr key={r.a.code}>
                <td style={{ ...cell, fontFamily: "'JetBrains Mono',monospace", fontWeight: 600 }}>{r.a.code}</td>
                <td style={{ ...cell }}>{r.a.name}</td>
                {numCell(r.budget, bCost)}
                {numCell(r.balPO, bCost)}
                {numCell(r.stock, bCost)}
                {numCell(r.balCost, bCost)}
                {r.inRow.map((v, i) => <Fragment key={"i" + i}>{numCell(v, bIn)}</Fragment>)}
                <td style={{ ...num, background: bIn, fontWeight: 700, color: r.inTot < 0 ? T.red : T.textPrimary }}>{money(r.inTot)}</td>
                {r.acRow.map((v, i) => <Fragment key={"a" + i}>{numCell(v, bAc)}</Fragment>)}
                <td style={{ ...num, background: bAc, fontWeight: 700, color: r.acTot < 0 ? T.red : T.textPrimary }}>{money(r.acTot)}</td>
                {r.pyRow.map((v, i) => <Fragment key={"p" + i}>{numCell(v, bPy)}</Fragment>)}
                <td style={{ ...num, background: bPy, fontWeight: 700, color: r.pyTot < 0 ? T.red : T.textPrimary }}>{money(r.pyTot)}</td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td style={{ ...cell, textAlign: "center", color: T.textMuted }} colSpan={6 + inM.length + acM.length + payM.length + 3}>— ยังไม่มีข้อมูล —</td></tr>
            )}
          </tbody>
          {rows.length > 0 && (
            <tfoot>
              <tr>
                <td style={{ ...cell, fontWeight: 800, background: "#f1f5f9" }} colSpan={2}>TOTAL</td>
                <td style={{ ...num, fontWeight: 800, background: "#eef2f7" }}>{money(totOf(r => r.budget))}</td>
                <td style={{ ...num, fontWeight: 800, background: "#eef2f7", color: totOf(r => r.balPO) < 0 ? T.red : T.textPrimary }}>{money(totOf(r => r.balPO))}</td>
                <td style={{ ...num, fontWeight: 800, background: "#eef2f7" }}>{money(totOf(r => r.stock))}</td>
                <td style={{ ...num, fontWeight: 800, background: "#eef2f7", color: totOf(r => r.balCost) < 0 ? T.red : T.textPrimary }}>{money(totOf(r => r.balCost))}</td>
                {inM.map((mk, i) => <td key={"ti" + mk} style={{ ...num, fontWeight: 700, background: "#e3efe3" }}>{money(colSum(r => r.inRow, i))}</td>)}
                <td style={{ ...num, fontWeight: 800, background: "#e3efe3" }}>{money(totOf(r => r.inTot))}</td>
                {acM.map((mk, i) => <td key={"ta" + mk} style={{ ...num, fontWeight: 700, background: "#e0eaf6" }}>{money(colSum(r => r.acRow, i))}</td>)}
                <td style={{ ...num, fontWeight: 800, background: "#e0eaf6" }}>{money(totOf(r => r.acTot))}</td>
                {payM.map((mk, i) => <td key={"tp" + mk} style={{ ...num, fontWeight: 700, background: "#fbe9d4" }}>{money(colSum(r => r.pyRow, i))}</td>)}
                <td style={{ ...num, fontWeight: 800, background: "#fbe9d4" }}>{money(totOf(r => r.pyTot))}</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </div>
  );
}

// ─── Accounting View ──────────────────────────────────────────────────────────
function AccountingView({ project, updateProject, tenderCosts, additions, poEntries, onBack, onHome, onExport, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[] }) {
  // บัญชี = อ่านอย่างเดียว (RLS ไม่ให้เขียน tcs-projects) → ปุ่มสกุลเงินจึงเป็นค่า
  // "ดูเฉพาะเครื่องนี้" ไม่บันทึกกลับไปที่โครงการร่วม กันไม่ให้บัญชีแก้ข้อมูลโครงการ
  const [curOverride, setCurOverride] = useState({});
  const curProject   = { ...project, ...curOverride };
  const setCurrency  = (fields) => setCurOverride(o => ({ ...o, ...fields }));
  const usdRate = effRate(curProject);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
  const [view, setView] = useState("dashboard");
  const [viewHist, setViewHist] = useState([]);   // ประวัติแท็บที่ดูมาก่อน — ปุ่มกลับจะย้อนทีละหน้า
  const goView = (v) => { if (v !== view) { setViewHist(h => [...h, view]); setView(v); } };
  const backView = () => { if (viewHist.length) { const h = [...viewHist]; const prev = h.pop(); setViewHist(h); setView(prev); } else onBack(); };
  const [sortKey, setSortKey] = useState(null);  // "code" | "name" | "group" | "budget" | "committed" | "pct" | null
  const [sortDir, setSortDir] = useState(1);
  // ค้นหา + ตัวกรอง "เฉพาะที่มี PO" บนแท็บ Cash Flow
  const [dateSearch, setDateSearch] = useState("");
  const [onlyWithPO, setOnlyWithPO] = useState(false);
  // Which Acc. Code groups are collapsed on the "วันที่ (Cash Flow)" tab.
  const [dateCollapsed, setDateCollapsed] = useState(() => new Set());
  const toggleDateGroup = (code) => setDateCollapsed(prev => {
    const next = new Set(prev);
    next.has(code) ? next.delete(code) : next.add(code);
    return next;
  });
  // Which months are collapsed on the "แผนจ่าย" (payment plan) tab.
  const [planCollapsed, setPlanCollapsed] = useState(() => new Set());
  const togglePlanMonth = (mk) => setPlanCollapsed(prev => {
    const next = new Set(prev);
    next.has(mk) ? next.delete(mk) : next.add(mk);
    return next;
  });

  // Budget = baseline Tender Cost + every monthly addition (ค่าธรรมดา + คอลัมน์
  // ย่อย) combined per Acc. Code — ใช้ตัวช่วยกลางเดียวกับ Export ให้ตัวเลขตรงกัน
  const combinedBudget = buildCombinedBudget(tenderCosts, additions);

  // Sum only top-level codes — see note in ProcurementView. Object.values()
  // over the whole combinedBudget double-counts sub-items (EX-xxxx rows),
  // since their value is already folded into their parent account's total.
  const topLevelCodes = [
    ...ACCOUNTS.filter(a => !hiddenAccounts.includes(a.code)).map(a => a.code),
    ...extraItems.filter(e => !e.parentCode).map(e => e.code),
  ];
  const tenderTotal   = topLevelCodes.reduce((s,c) => s + (parseFloat(combinedBudget[c]) || 0), 0);
  const totalComm     = poEntries.reduce((s,p)=>s+poTotal(p),0);
  // จ่ายแล้ว = ยอดที่ตัดจ่ายอัตโนมัติจริง (งวดที่รับของแล้ว + ถึงกำหนดจ่าย) ให้ตรงกับไฟล์ Excel
  const totalPaid     = poEntries.reduce((s,p)=> s + poRounds(p).filter(r=>roundReceived(r)&&roundPaid(p,r)).reduce((ss,r)=>ss+(parseFloat(r.actualAmount)||0),0), 0);
  const paidPOCount   = poEntries.filter(p=>paymentStatus(p)==="paid").length;
  const totalInvoiced = poEntries.filter(p=>["Invoiced","Paid"].includes(p.status)).reduce((s,p)=>s+poTotal(p),0);
  const pct           = tenderTotal>0?(totalComm/tenderTotal*100):0;

  // รวม "งานเพิ่ม" (standalone extra) เข้าไปในกราฟตามกลุ่มด้วย ไม่งั้นยอดในกราฟ
  // จะไม่ตรงกับการ์ดสรุป (ที่นับ topLevelCodes รวม extra) — และเคารพบัญชีที่ซ่อนไว้
  const chartGroups = [...new Set([...GROUPS, ...extraItems.filter(e=>!e.parentCode).map(e=>e.group||"อื่น ๆ")])];
  const groupData = chartGroups.map((g,i)=>{
    const codes=[
      ...ACCOUNTS.filter(a=>a.group===g && !hiddenAccounts.includes(a.code)).map(a=>a.code),
      ...extraItems.filter(e=>!e.parentCode && (e.group||"อื่น ๆ")===g).map(e=>e.code),
    ];
    const committed = poEntries.reduce((s,p)=>s+poItems(p).filter(it=>codes.includes(it.code)).reduce((s2,it)=>s2+(parseFloat(it.amount)||0),0),0);
    return {group:g,budget:codes.reduce((s,c)=>s+(parseFloat(combinedBudget[c])||0),0),committed,color:GRP_COLORS[i%GRP_COLORS.length]};
  }).filter(g=>g.budget>0||g.committed>0);

  // รวมบัญชีมาตรฐาน + "งานเพิ่ม" (standalone extra ที่ไม่ใช่รายการย่อย) ให้ยอดรวม
  // หน้าบัญชีตรงกับหน้า QS/ภาพรวม ที่นับ topLevelCodes เหมือนกัน
  const acctRows = [
    ...ACCOUNTS.filter(a=>!hiddenAccounts.includes(a.code)),
    ...extraItems.filter(e=>!e.parentCode).map(e=>({ code:e.code, name:e.name, group:e.group||"อื่น ๆ" })),
  ];
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

  // ─── Cash-flow-by-date view ──────────────────────────────────────────────
  // For each Acc. Code: budget vs. committed (+ variance / variance %), plus
  // every PO booked to it with its three key dates — when the PO was opened,
  // when goods are due in (plan → actual per delivery batch), and when
  // payment is due — so accounting can see cash timing at a glance.
  const dateGroups = accountData.map(a => {
    const rows = poEntries
      .filter(p => poItems(p).some(it => it.code === a.code))
      .map(p => ({ po: p, item: poItems(p).find(it => it.code === a.code) }))
      .sort((x, y) => (x.po.date || "").localeCompare(y.po.date || ""));
    const variance = a.budget - a.committed;
    const variancePct = a.budget > 0 ? (variance / a.budget) * 100 : (a.committed > 0 ? null : 0);
    // จ่ายแล้วจริง (งวดที่รับของ + ถึงกำหนดจ่าย) และยอดที่ยังต้องเก็บเงินไว้รอจ่าย
    // ของ Acc. Code นี้ — สิ่งที่บัญชีต้องรู้ว่าต้องกันเงินไว้เท่าไหร่
    const paid = poEntries.reduce((s,p)=> s + poItems(p).filter(it=>it.code===a.code)
      .reduce((ss,it)=> ss + (it.rounds||[]).filter(r=>roundReceived(r)&&roundPaid(p,r))
        .reduce((s3,r)=> s3 + (parseFloat(r.actualAmount)||0), 0), 0), 0);
    const toReserve = Math.max(a.committed - paid, 0);
    return { ...a, rows, variance, variancePct, paid, toReserve };
  }).sort((x, y) => x.code.localeCompare(y.code));
  // ตัวกรองแท็บ Cash Flow: ค้นหา (วันที่/Acc.Code/ชื่อรายการ/เลข PO) + เฉพาะที่มี PO
  const dq = dateSearch.trim().toLowerCase();
  const shownDateGroups = dateGroups.filter(a => {
    if (onlyWithPO && a.rows.length === 0) return false;
    if (!dq) return true;
    if (a.code.toLowerCase().includes(dq) || (a.name||"").toLowerCase().includes(dq)) return true;
    return a.rows.some(({po}) =>
      (poNumbersLabel(po)||"").toLowerCase().includes(dq) ||
      (po.date||"").includes(dq) ||
      (poNextDueDate(po)||"").includes(dq) ||
      poRounds(po).some(r => (r.actualDate||"").includes(dq) || (r.planDate||"").includes(dq))
    );
  });
  const withPOCount = dateGroups.filter(a => a.rows.length > 0).length;
  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => -d);
    else { setSortKey(key); setSortDir(1); }
  };
  const displayAccountData = (() => {
    if (!sortKey) return accountData;
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
    return arr;
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
    return { mk, label: mk==="9999-99"?"ยังไม่ระบุวันจ่าย":monthShortLabel(mk), lines, cash, credit, sum, paid:paidA, remain:Math.max(0,sum-paidA) };
  });
  const planTotal  = payLines.reduce((s,l)=>s+l.amount,0);
  const planPaid   = payLines.reduce((s,l)=>s+(l.paidAmount||0),0);
  const planRemain = Math.max(0, planTotal - planPaid);
  const thisMonthKey = payToday.slice(0,7);
  // "ครบกำหนดเดือนนี้" = คงเหลือของเดือนนี้ + ยอดที่เลยกำหนดจากเดือนก่อน ๆ ที่ยังไม่จ่าย
  const dueThisMonth = payByMonth.filter(m=>m.mk!=="9999-99" && m.mk<=thisMonthKey).reduce((s,m)=>s+m.remain,0);
  // เดือนถัดไป — สำหรับแจ้งเตือนให้บัญชีเตรียมเงินล่วงหน้า
  const nextMonthKey = (() => { const [y,m]=thisMonthKey.split("-").map(Number); const ny=m===12?y+1:y, nm=m===12?1:m+1; return `${ny}-${String(nm).padStart(2,"0")}`; })();
  const nextBucket   = payByMonth.find(m=>m.mk===nextMonthKey);
  const dueNextMonth = nextBucket?.remain || 0;
  const nextCash     = nextBucket?.cash || 0;
  const nextCredit   = nextBucket?.credit || 0;
  const nextCount    = nextBucket?.lines.length || 0;

  const pieData = PO_STATUS.map(s=>({name:s,value:poEntries.filter(p=>p.status===s).reduce((sum,p)=>sum+poTotal(p),0),color:STATUS_CLR[s]})).filter(d=>d.value>0);

  const CT = ({active,payload}) => {
    if (!active||!payload?.length) return null;
    return (
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:10,padding:"10px 14px",fontSize:12,boxShadow:"0 4px 16px rgba(0,0,0,0.1)"}}>
        <div style={{color:T.textMuted,marginBottom:4,fontWeight:600}}>{payload[0]?.payload?.group}</div>
        {payload.map(p=><div key={p.name} style={{color:p.fill||p.color,fontFamily:"'JetBrains Mono',monospace"}}>{p.name}: {fmt(p.value)}</div>)}
      </div>
    );
  };

  const DateCell = ({ value, lateTint }) => (
    <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:value?(lateTint?T.red:T.textPrimary):T.textMuted,fontWeight:value&&lateTint?700:400}}>
      {value || "—"}
    </span>
  );
  const Badge = ({ text, clr, bg }) => (
    <span style={{background:bg,color:clr,fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{text}</span>
  );
  // Every delivery batch of a PO, one line per shipment: plan → actual date
  // with its own on-time/late/received status, same as Procurement's view.
  const DeliveryDates = ({ po }) => {
    const deliveries = poDeliveries(po);
    if (!deliveries.length) return <span style={{fontSize:12,color:T.textMuted}}>—</span>;
    return (
      <div style={{display:"flex",flexDirection:"column",gap:3}}>
        {deliveries.map((d,i)=>{
          const st = deliveryStatus(d);
          return (
            <div key={d.id||i} style={{display:"flex",alignItems:"center",gap:5}}>
              {deliveries.length>1 && <span style={{fontSize:10,color:T.textMuted,fontWeight:700,minWidth:14}}>#{i+1}</span>}
              <DateCell value={d.plan} lateTint={st==="late"}/>
              <span style={{color:T.textMuted,fontSize:11}}>→</span>
              <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:st==="received"?T.green:T.textMuted,fontWeight:st==="received"?600:400}}>{d.actual||"รอ"}</span>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <Shell role="accounting" color={T.green} project={project} onBack={backView} onHome={onHome} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"24px 28px"}}>
        {/* Tabs + Export */}
        <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
          {[["dashboard","📊 ภาพรวมงบ","ภาพรวม: งบประมาณ vs ที่ผูกพันแล้ว (PO) ทั้งโครงการ"],["dates","📅 ราย Acc. Code","รายหมวด: PO + วันของเข้า/วันครบกำหนดจ่าย + ยอดที่ต้องเก็บไว้จ่าย"],["plan","💰 แผนจ่ายรายเดือน","รายเดือน: เงินที่ต้องเตรียมจ่ายแยกตามเดือนครบกำหนด"],["matrix","📄 ตารางรวมเดือน","ตารางรวม: ต้นทุน + Incoming Plan / Actual Received / Payment Plan รายเดือน (เฉพาะเดือนที่มีข้อมูล)"]].map(([v,l,tip])=>(
            <button key={v} onClick={()=>goView(v)} title={tip}
              style={{background:view===v?T.green:"transparent",border:`1.5px solid ${view===v?T.green:T.cardBorder}`,borderRadius:10,padding:"8px 20px",color:view===v?"#fff":T.textSecondary,fontSize:13,cursor:"pointer",fontWeight:view===v?600:500,transition:"all 0.15s"}}>{l}</button>
          ))}
          <div style={{marginLeft:"auto"}}><CurrencyControl project={curProject} updateProject={setCurrency}/></div>
          <button onClick={onExport} className="btn-ghost" style={{display:"flex",alignItems:"center",gap:6,borderColor:T.green,color:T.green}}>
            ⬇️ Export Excel
          </button>
        </div>
        {/* คำอธิบายสี (legend) */}
        <div style={{display:"flex",flexWrap:"wrap",gap:16,marginBottom:20,fontSize:11,color:T.textMuted,alignItems:"center"}}>
          <span style={{fontWeight:700,color:T.textSecondary}}>คำอธิบายสี:</span>
          {[[T.green,"ปกติ · ใช้งบ <80% · จ่ายแล้ว"],[T.amber,"เฝ้าระวัง · ใช้งบ 80–100% · รอจ่าย"],[T.red,"เกินงบ · เกินกำหนดจ่าย"]].map(([c,t])=>(
            <span key={t} style={{display:"inline-flex",alignItems:"center",gap:6}}>
              <span style={{width:11,height:11,borderRadius:3,background:c,display:"inline-block"}}/>{t}
            </span>
          ))}
        </div>

        {/* 🔔 แจ้งเตือนยอดต้องจ่ายเดือนหน้า — เห็นทุกแท็บ กดแล้วไปหน้าแผนจ่าย */}
        {(dueThisMonth>0 || dueNextMonth>0) && (
          <div onClick={()=>goView("plan")} title="ดูรายละเอียดในหน้าแผนจ่าย"
            style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap",cursor:"pointer",userSelect:"none",
              background:"linear-gradient(90deg,#fffbeb,#fff)",border:`1px solid ${T.amber}`,borderLeft:`5px solid ${T.amber}`,
              borderRadius:12,padding:"12px 16px",marginBottom:20}}>
            <span style={{fontSize:22,lineHeight:1}}>🔔</span>
            <span style={{fontSize:13,color:T.textSecondary,fontWeight:700}}>เตรียมเงินจ่าย</span>
            {/* เดือนนี้ */}
            <div style={{background:T.redBg,borderRadius:10,padding:"6px 12px",minWidth:150}}>
              <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>ครบกำหนดเดือนนี้ · {monthShortLabel(thisMonthKey)}</div>
              <div style={{fontSize:18,fontWeight:800,color:T.red,fontFamily:"'JetBrains Mono',monospace"}}>฿{fmt0(dueThisMonth)}</div>
              {usdLine(dueThisMonth, usdRate)}
            </div>
            {/* เดือนหน้า */}
            <div style={{background:T.amberBg,borderRadius:10,padding:"6px 12px",minWidth:150}}>
              <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>เตรียมเดือนหน้า · {monthShortLabel(nextMonthKey)}</div>
              <div style={{fontSize:18,fontWeight:800,color:T.amber,fontFamily:"'JetBrains Mono',monospace"}}>฿{fmt0(dueNextMonth)}</div>
              {usdLine(dueNextMonth, usdRate)}
            </div>
            <div style={{flex:1}}/>
            <span style={{fontSize:12,color:T.amber,fontWeight:700,whiteSpace:"nowrap"}}>ดูแผนจ่าย →</span>
          </div>
        )}

        {view==="dashboard" ? (
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:24}}>
              <StatCard label="งบประมาณ (QS)" value={"฿"+fmt0(tenderTotal)} thb={tenderTotal} rate={usdRate} sub="เดิม + เพิ่มรายเดือนทุกเดือน" color={T.blue} icon="📋" accent={T.blueLight}/>
              <StatCard label="ผูกพันแล้ว (PO)" value={"฿"+fmt0(totalComm)} thb={totalComm} rate={usdRate} sub={`${pct.toFixed(1)}% ของงบ`} color={T.amber} icon="📦" accent={T.amberBg}/>
              <StatCard label="วางบิลแล้ว" value={"฿"+fmt0(totalInvoiced)} thb={totalInvoiced} rate={usdRate} sub="รอจ่าย + จ่ายแล้ว" color={T.purple} icon="🧾" accent={T.purpleBg}/>
              <StatCard label="ชำระแล้ว" value={"฿"+fmt0(totalPaid)} thb={totalPaid} rate={usdRate} sub={`${paidPOCount} รายการ`} color={T.green} icon="✅" accent={T.greenBg}/>
            </div>

            {/* แถบเตือน "เกินงบ" (เรื่องเงินจ่ายย้ายไปรวมที่แถบ 🔔 ด้านบนแล้ว) */}
            {(() => {
              const overCount = accountData.filter(a=>a.over).length;
              if (!overCount) return null;
              return (
                <div style={{display:"flex",flexWrap:"wrap",gap:12,marginBottom:20}}>
                  <button onClick={()=>handleSort("variance")} title="เรียงตารางตามส่วนต่าง" style={{display:"flex",alignItems:"center",gap:8,background:T.redBg,color:T.red,border:`1px solid ${T.red}`,borderRadius:10,padding:"10px 16px",fontSize:13,fontWeight:700,cursor:"pointer"}}>
                    ⚠ {overCount} หมวดเกินงบ <span style={{fontSize:11,fontWeight:500,opacity:0.85}}>· กดเพื่อเรียงดู</span>
                  </button>
                </div>
              );
            })()}

            {/* Progress */}
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:22,marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                <span style={{fontSize:13,color:T.textPrimary,fontWeight:600}}>สัดส่วนการใช้งบ</span>
                <span style={{fontSize:13,color:tenderTotal-totalComm<0?T.red:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>
                  {tenderTotal-totalComm<0?"เกินงบ ":"คงเหลือ "}{fmt(Math.abs(tenderTotal-totalComm))}
                </span>
              </div>
              <div style={{background:"#f1f5f9",borderRadius:99,height:10,overflow:"hidden"}}>
                <div style={{width:`${Math.min(pct,100)}%`,background:pct>100?T.red:pct>80?T.amber:T.green,height:"100%",borderRadius:99,transition:"width 0.5s"}}/>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",marginTop:7,fontSize:11,color:T.textMuted,fontFamily:"'JetBrains Mono',monospace"}}>
                <span>0</span><span style={{fontWeight:600,color:pct>100?T.red:T.textSecondary}}>{pct.toFixed(1)}%</span><span>{fmt(tenderTotal)}</span>
              </div>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
              <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:22}}>
                <p style={{margin:"0 0 16px",fontSize:13,color:T.textPrimary,fontWeight:600}}>Budget vs Committed ตาม Group</p>
                {groupData.length===0
                  ? <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>QS ยังไม่ได้ลง Tender Cost</div>
                  : <ResponsiveContainer width="100%" height={260}>
                      <BarChart data={groupData} margin={{left:0,right:0,top:4,bottom:44}}>
                        <XAxis dataKey="group" tick={{fill:T.textMuted,fontSize:10}} angle={-30} textAnchor="end" interval={0}/>
                        <YAxis tick={{fill:T.textMuted,fontSize:10}} tickFormatter={fmtK} width={60}/>
                        <Tooltip content={<CT/>}/>
                        <Bar dataKey="budget" name="Budget" fill={T.blue} radius={[5,5,0,0]}/>
                        <Bar dataKey="committed" name="Committed" fill={T.amber} radius={[5,5,0,0]}/>
                      </BarChart>
                    </ResponsiveContainer>
                }
              </div>
              <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:22}}>
                <p style={{margin:"0 0 16px",fontSize:13,color:T.textPrimary,fontWeight:600}}>สถานะ PO</p>
                {pieData.length===0
                  ? <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>ยังไม่มี PO</div>
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
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden",marginTop:20}}>
            <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {[
                    {label:"Acc. Code", key:"code"},
                    {label:"Account Name", key:"name"},
                    {label:"Group", key:"group"},
                    {label:"Budget (QS)", key:"budget"},
                    {label:"Committed (PO)", key:"committed"},
                    {label:"ส่วนต่าง", key:"variance"},
                  ].map(({label,key})=>(
                    <th key={label||"__actions"}
                      style={{padding:"11px 16px",textAlign:["Budget (QS)","Committed (PO)","ส่วนต่าง"].includes(label)?"right":"left",color:sortKey===key?T.green:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                      <span onClick={()=>key&&handleSort(key)} style={{cursor:key?"pointer":"default",userSelect:"none"}}>{label}{key && sortKey===key ? (sortDir===1?" ▲":" ▼") : ""}</span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {displayAccountData.map((a,i)=>{
                  const variance = a.budget - a.committed;
                  return (
                    <tr key={a.code} style={{background:a.over?"#fff5f5":i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`}}>
                      <td style={{padding:"10px 16px",color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontSize:12,fontWeight:500}}>{a.code}</td>
                      <td style={{padding:"10px 16px",color:T.textPrimary}}>{a.name}</td>
                      <td style={{padding:"10px 16px"}}>
                        <span style={{background:T.blueLight,color:T.blue,fontSize:10,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{a.group}</span>
                      </td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:T.blue,fontWeight:500}}>{a.budget>0?fmt(a.budget):"—"}{a.budget>0&&usdLine(a.budget, usdRate)}</td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:a.over?T.red:T.amber,fontWeight:a.over?700:500}}>{a.committed>0?fmt(a.committed):"—"}{a.committed>0&&usdLine(a.committed, usdRate)}</td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:a.budget<=0&&a.committed>0?T.textMuted:variance<0?T.red:T.textSecondary,fontWeight:a.budget<=0&&a.committed>0?500:variance<0?700:500}}>
                        {a.budget<=0&&a.committed>0 ? "ไม่มีงบ" : (a.budget>0||a.committed>0?`${variance<0?"-":""}${fmt(Math.abs(variance))}`:"—")}
                        {(a.budget>0||a.committed>0)&&!(a.budget<=0&&a.committed>0)&&usdLine(Math.abs(variance), usdRate)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
                  <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:12}}>{accountData.length} รายการ</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:T.blue,fontWeight:700,fontSize:14}}>{fmt(accountData.reduce((s,a)=>s+a.budget,0))}{usdLine(accountData.reduce((s,a)=>s+a.budget,0), usdRate)}</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:T.amber,fontWeight:700,fontSize:14}}>{fmt(accountData.reduce((s,a)=>s+a.committed,0))}{usdLine(accountData.reduce((s,a)=>s+a.committed,0), usdRate)}</td>
                  {(() => {
                    const totalVariance = accountData.reduce((s,a)=>s+(a.budget-a.committed),0);
                    return (
                      <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:totalVariance<0?T.red:T.textSecondary,fontWeight:700,fontSize:14}}>
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
        ) : view==="dates" ? (
          <div>
            {/* Grand totals across every Acc. Code that has a budget or a PO */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
              <StatCard label="งบประมาณรวม" value={"฿"+fmt0(dateGroups.reduce((s,a)=>s+a.budget,0))} thb={dateGroups.reduce((s,a)=>s+a.budget,0)} rate={usdRate} sub={`${dateGroups.length} Acc. Code`} color={T.blue} icon="📋" accent={T.blueLight}/>
              <StatCard label="PO รวม" value={"฿"+fmt0(dateGroups.reduce((s,a)=>s+a.committed,0))} thb={dateGroups.reduce((s,a)=>s+a.committed,0)} rate={usdRate} sub={`${poEntries.length} PO`} color={T.amber} icon="📦" accent={T.amberBg}/>
              <StatCard label="ส่วนต่างรวม" value={"฿"+fmt0(Math.abs(dateGroups.reduce((s,a)=>s+a.variance,0)))} thb={Math.abs(dateGroups.reduce((s,a)=>s+a.variance,0))} rate={usdRate}
                sub={dateGroups.reduce((s,a)=>s+a.variance,0)<0?"เกินงบ":"คงเหลือ"}
                color={dateGroups.reduce((s,a)=>s+a.variance,0)<0?T.red:T.green}
                icon={dateGroups.reduce((s,a)=>s+a.variance,0)<0?"⚠️":"💰"}
                accent={dateGroups.reduce((s,a)=>s+a.variance,0)<0?T.redBg:T.greenBg}/>
              <StatCard label="ต้องเก็บไว้จ่ายรวม" value={"฿"+fmt0(dateGroups.reduce((s,a)=>s+a.toReserve,0))} thb={dateGroups.reduce((s,a)=>s+a.toReserve,0)} rate={usdRate}
                sub={`${poEntries.filter(p=>paymentStatus(p)==="pending"||paymentStatus(p)==="late").length} PO รอจ่าย`} color={T.red} icon="⏳" accent={T.redBg}/>
            </div>

            {/* ค้นหา + ตัวกรอง */}
            <div style={{display:"flex",gap:12,alignItems:"center",flexWrap:"wrap",marginBottom:16}}>
              <input value={dateSearch} onChange={e=>setDateSearch(e.target.value)}
                placeholder="🔍 ค้นหา: วันที่ / Acc. Code / ชื่อรายการ / เลข PO"
                style={{flex:1,minWidth:240,maxWidth:420,padding:"9px 14px",border:`1px solid ${T.cardBorder}`,borderRadius:10,fontSize:13,outline:"none"}}/>
              <button onClick={()=>setOnlyWithPO(v=>!v)} title="แสดงเฉพาะ Acc. Code ที่มี PO"
                style={{display:"flex",alignItems:"center",gap:8,padding:"9px 16px",borderRadius:10,fontSize:13,fontWeight:600,cursor:"pointer",
                  border:`1.5px solid ${onlyWithPO?T.green:T.cardBorder}`,background:onlyWithPO?T.green:"transparent",color:onlyWithPO?"#fff":T.textSecondary}}>
                <span style={{fontSize:14}}>{onlyWithPO?"☑":"☐"}</span> เฉพาะที่มี PO ({withPOCount})
              </button>
              <span style={{fontSize:12,color:T.textMuted}}>แสดง {shownDateGroups.length} / {dateGroups.length} หมวด</span>
            </div>

            {dateGroups.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
                <div style={{fontSize:32,marginBottom:12}}>📅</div>
                <div style={{fontSize:14,fontWeight:500,color:T.textSecondary}}>ยังไม่มีงบหรือ PO ให้แสดง</div>
              </div>
            ) : shownDateGroups.length===0 ? (
              <div style={{textAlign:"center",padding:"40px 0",color:T.textMuted,fontSize:13}}>
                ไม่พบหมวดที่ตรงกับเงื่อนไข {dateSearch.trim() && <>"{dateSearch}"</>} {onlyWithPO && "· (กรองเฉพาะที่มี PO)"}
              </div>
            ) : (
              <div style={{display:"flex",flexDirection:"column",gap:14}}>
                {shownDateGroups.map(a => {
                  const isCollapsed = dateCollapsed.has(a.code);
                  // แถบ = ความคืบหน้าการจ่ายของ PO ที่ผูกพันแล้ว (จ่ายแล้ว vs ต้องเก็บไว้จ่าย)
                  const paidPct   = a.committed>0 ? (a.paid/a.committed*100) : 0;
                  const barPct    = a.committed>0 ? Math.min(paidPct,100) : 0;
                  const statusClr = a.over?T.red:a.committed>0?T.green:T.textMuted;
                  const statusBg  = a.over?T.redBg:a.committed>0?T.greenBg:"#eef1f5";
                  const statusTxt = a.over?"⚠ เกินงบ":a.committed>0?"✅ OK":a.budget>0?"ยังไม่ PO":"—";
                  const varClr    = a.variancePct===null ? T.textMuted : a.variance<0 ? T.red : T.green;
                  return (
                    <div key={a.code} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderLeft:`4px solid ${statusClr}`,borderRadius:14,overflow:"hidden"}}>
                      <div onClick={()=>toggleDateGroup(a.code)}
                        style={{padding:"14px 18px",background:a.over?"#fff8f8":"#fbfcfe",borderBottom:isCollapsed?"none":`1px solid ${T.cardBorder}`,cursor:"pointer",userSelect:"none"}}>
                        {/* บรรทัด 1: ชื่อ + สถานะ */}
                        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
                          <span style={{fontSize:11,color:T.textMuted,transform:isCollapsed?"rotate(-90deg)":"none",transition:"transform 0.15s",display:"inline-block",width:12,flexShrink:0}}>▼</span>
                          <span style={{color:T.blue,fontSize:12,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,background:T.blueLight,padding:"2px 8px",borderRadius:6,flexShrink:0}}>{a.code}</span>
                          <span style={{color:T.textPrimary,fontSize:14,fontWeight:600,flex:1,minWidth:0}}>{a.name}</span>
                          <span style={{background:statusBg,color:statusClr,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:700,whiteSpace:"nowrap",flexShrink:0}}>{statusTxt}</span>
                        </div>
                        {/* บรรทัด 2: แถบความคืบหน้าการจ่าย + ยอดที่ต้องเก็บเงินไว้รอจ่าย */}
                        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}>
                          <div style={{flex:1,background:"#eef1f5",borderRadius:99,height:8,overflow:"hidden"}} title={`จ่ายแล้ว ${a.committed>0?paidPct.toFixed(0):0}% ของ PO`}>
                            <div style={{width:`${barPct}%`,background:T.green,height:"100%",borderRadius:99,transition:"width 0.5s"}}/>
                          </div>
                          <span style={{fontSize:12,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:a.toReserve>0?T.amber:a.committed>0?T.green:T.textMuted,textAlign:"right",whiteSpace:"nowrap"}}>
                            {a.committed>0 ? (a.toReserve>0 ? `เก็บไว้จ่าย ฿${fmt0(a.toReserve)}` : "จ่ายครบแล้ว") : "ยังไม่มี PO"}
                          </span>
                        </div>
                        {/* บรรทัด 3: ตัวเลขสรุป 3 ช่อง — มูลค่า PO · จ่ายแล้ว · ต้องเก็บไว้จ่าย */}
                        <div style={{display:"flex",gap:24,flexWrap:"wrap"}}>
                          <div style={{minWidth:96}}>
                            <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5,marginBottom:2}}>มูลค่า PO</div>
                            <div style={{fontSize:14,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:T.amber}}>{a.committed>0?fmt(a.committed):"—"}</div>
                            {a.committed>0&&usdLine(a.committed, usdRate)}
                          </div>
                          <div style={{minWidth:96}}>
                            <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5,marginBottom:2}}>จ่ายแล้ว</div>
                            <div style={{fontSize:14,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:T.green}}>{a.paid>0?fmt(a.paid):"—"}</div>
                            {a.paid>0&&usdLine(a.paid, usdRate)}
                          </div>
                          <div style={{minWidth:96}}>
                            <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5,marginBottom:2}}>ต้องเก็บไว้จ่าย</div>
                            <div style={{fontSize:14,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:a.toReserve>0?T.amber:T.green}}>{a.committed>0?fmt(a.toReserve):"—"}</div>
                            {a.committed>0&&usdLine(a.toReserve, usdRate)}
                          </div>
                        </div>
                        {/* งบประมาณ / ส่วนต่าง (ข้อมูลงบ ไว้ท้ายสุด) */}
                        <div style={{display:"flex",gap:24,flexWrap:"wrap",marginTop:8,paddingTop:8,borderTop:`1px dashed ${T.cardBorder}`}}>
                          <div style={{fontSize:11,color:T.textMuted}}>งบประมาณ: <b style={{color:T.blue,fontFamily:"'JetBrains Mono',monospace"}}>{a.budget>0?fmt(a.budget):"—"}</b></div>
                          <div style={{fontSize:11,color:T.textMuted}}>{a.variance<0?"เกินงบ":"งบคงเหลือ"}: <b style={{color:varClr,fontFamily:"'JetBrains Mono',monospace"}}>{a.variancePct===null ? "ไม่มีงบ" : `${a.variance<0?"-":""}${fmt(Math.abs(a.variance))}`}</b></div>
                        </div>
                      </div>
                      {!isCollapsed && (
                        a.rows.length===0 ? (
                          <div style={{padding:"14px 18px",fontSize:12,color:T.textMuted}}>ยังไม่มี PO ผูกกับ Acc. Code นี้</div>
                        ) : (
                        <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
                          <thead>
                            <tr>
                              {["วันเปิด PO (แพลน)","Supplier","PO No.","มูลค่า (THB)","ของเข้า (แผน→จริง)","ต้องจ่ายเงินวันไหน","สถานะจ่าย"].map(h=>(
                                <th key={h} style={{padding:"9px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:10,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {a.rows.map(({po:p,item},i)=>{
                              const pay = paymentStatus(p);
                              return (
                                <tr key={p.id+"-"+(item.id||item.code)}
                                  style={{background:i%2===0?T.card:"#fafbfd",borderBottom:"1px solid #f1f5f9"}}>
                                  <td style={{padding:"9px 16px"}}><DateCell value={p.date}/></td>
                                  <td style={{padding:"9px 16px",color:T.textPrimary,fontWeight:500}}>{itemSupplierName(p,item)}</td>
                                  <td style={{padding:"9px 16px",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{poNumbersLabel(p)}</td>
                                  <td style={{padding:"9px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontWeight:600,color:T.textPrimary}}>{fmt(item.amount)}{usdLine(parseFloat(item.amount)||0, usdRate)}</td>
                                  <td style={{padding:"9px 16px"}}><DeliveryDates po={p}/></td>
                                  <td style={{padding:"9px 16px"}}><DateCell value={poNextDueDate(p)} lateTint={false}/></td>
                                  <td style={{padding:"9px 16px"}}><Badge text={PAYMENT_LABEL[pay]} clr={PAYMENT_CLR[pay]} bg={PAYMENT_BG[pay]}/></td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table></div>
                        )
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        ) : view==="matrix" ? (
          <AccountingMatrixTab tenderCosts={tenderCosts} additions={additions} poEntries={poEntries} extraItems={extraItems} hiddenAccounts={hiddenAccounts} />
        ) : (
          <div>
            {/* สรุปยอดที่ต้องเตรียมจ่าย */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(190px,1fr))",gap:16,marginBottom:20}}>
              <StatCard label="ต้องจ่ายทั้งหมด" value={"฿"+fmt0(planTotal)} thb={planTotal} rate={usdRate} sub={`${payLines.length} งวด`} color={T.blue} icon="📋" accent={T.blueLight}/>
              <StatCard label="จ่ายแล้ว" value={"฿"+fmt0(planPaid)} thb={planPaid} rate={usdRate} sub="ครบกำหนด + ตัดจ่ายแล้ว" color={T.green} icon="✅" accent={T.greenBg}/>
              <StatCard label="คงเหลือต้องจ่าย" value={"฿"+fmt0(planRemain)} thb={planRemain} rate={usdRate} sub="ยอดที่ยังไม่จ่าย" color={T.amber} icon="⏳" accent={T.amberBg}/>
              <StatCard label={`ครบกำหนดเดือนนี้ (${monthShortLabel(thisMonthKey)})`} value={"฿"+fmt0(dueThisMonth)} thb={dueThisMonth} rate={usdRate} sub="เตรียมเงินเดือนนี้" color={T.red} icon="💰" accent={T.redBg}/>
              <StatCard label={`ครบกำหนดเดือนหน้า (${monthShortLabel(nextMonthKey)})`} value={"฿"+fmt0(dueNextMonth)} thb={dueNextMonth} rate={usdRate} sub={dueNextMonth>0?`${nextCount} งวด · เตรียมล่วงหน้า`:"ยังไม่มีที่ครบกำหนด"} color={T.amber} icon="🔔" accent={T.amberBg}/>
            </div>

            {payByMonth.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
                <div style={{fontSize:32,marginBottom:12}}>💰</div>
                <div style={{fontSize:14,fontWeight:500,color:T.textSecondary}}>ยังไม่มีงวดจ่ายให้แสดง</div>
                <div style={{fontSize:12,color:T.textMuted,marginTop:6}}>วันครบกำหนดจ่ายมาจากวันรับของ (แผน/จริง) + เทอมเครดิตของ PO</div>
              </div>
            ) : (
              <div style={{display:"flex",flexDirection:"column",gap:14}}>
                {payByMonth.map(m => {
                  const isCollapsed = planCollapsed.has(m.mk);
                  const isThis = m.mk===thisMonthKey;
                  return (
                    <div key={m.mk} style={{background:T.card,border:`1px solid ${isThis?T.amber:T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                      <div onClick={()=>togglePlanMonth(m.mk)}
                        style={{padding:"12px 18px",background:isThis?T.amberBg:"#f8fafc",borderBottom:isCollapsed?"none":`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:16,cursor:"pointer",userSelect:"none",flexWrap:"wrap"}}>
                        <span style={{fontSize:11,color:T.textMuted,transform:isCollapsed?"rotate(-90deg)":"none",transition:"transform 0.15s",display:"inline-block",width:12}}>▼</span>
                        <div style={{minWidth:150}}>
                          <span style={{color:T.textPrimary,fontSize:14,fontWeight:700}}>{m.label}</span>
                          {isThis && <span style={{marginLeft:8,background:T.amber,color:"#fff",fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600}}>เดือนนี้</span>}
                          <span style={{marginLeft:8,color:T.textMuted,fontSize:12}}>{m.lines.length} งวด</span>
                        </div>
                        <div style={{flex:1}}/>
                        <div style={{textAlign:"right",minWidth:88}}>
                          <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>เงินสด</div>
                          <div style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace",fontWeight:600,color:T.green}}>{m.cash>0?fmt(m.cash):"—"}</div>
                          {m.cash>0&&usdLine(m.cash, usdRate)}
                        </div>
                        <div style={{textAlign:"right",minWidth:88}}>
                          <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>เครดิต</div>
                          <div style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace",fontWeight:600,color:T.blue}}>{m.credit>0?fmt(m.credit):"—"}</div>
                          {m.credit>0&&usdLine(m.credit, usdRate)}
                        </div>
                        <div style={{textAlign:"right",minWidth:100}}>
                          <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>รวมต้องจ่าย</div>
                          <div style={{fontSize:14,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:T.textPrimary}}>{fmt(m.sum)}</div>
                          {usdLine(m.sum, usdRate)}
                        </div>
                        <div style={{textAlign:"right",minWidth:100}}>
                          <div style={{fontSize:9,color:T.textMuted,textTransform:"uppercase",letterSpacing:0.5}}>คงเหลือ</div>
                          <div style={{fontSize:14,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:m.remain>0?T.amber:T.green}}>{fmt(m.remain)}</div>
                          {usdLine(m.remain, usdRate)}
                        </div>
                      </div>
                      {!isCollapsed && (
                        <div className="hscroll"><table style={{width:"100%",minWidth:680,borderCollapse:"collapse",fontSize:13}}>
                          <thead>
                            <tr>
                              {["ครบกำหนดจ่าย","Supplier","PO No.","Acc. Code","วิธีจ่าย","วันรับของ","ยอดต้องจ่าย (THB)","สถานะ"].map(h=>(
                                <th key={h} style={{padding:"9px 16px",textAlign:h==="ยอดต้องจ่าย (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:10,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {m.lines.map((l,i)=>(
                              <tr key={i} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:"1px solid #f1f5f9"}}>
                                <td style={{padding:"9px 16px"}}><DateCell value={l.payDate} lateTint={l.status==="late"}/></td>
                                <td style={{padding:"9px 16px",color:T.textPrimary,fontWeight:500}}>{l.supplier}</td>
                                <td style={{padding:"9px 16px",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{l.poNo}</td>
                                <td style={{padding:"9px 16px",color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{l.code||"—"}</td>
                                <td style={{padding:"9px 16px"}}>
                                  <span style={{background:l.isCash?T.greenBg:T.blueLight,color:l.isCash?T.green:T.blue,fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{l.method}</span>
                                </td>
                                <td style={{padding:"9px 16px",whiteSpace:"nowrap"}}>
                                  <span style={{fontFamily:"'JetBrains Mono',monospace",fontSize:12,color:T.textSecondary}}>{l.incoming||"—"}</span>
                                  {l.incomingType && <span style={{marginLeft:5,fontSize:10,color:T.textMuted}}>({l.incomingType})</span>}
                                </td>
                                <td style={{padding:"9px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontWeight:600,color:T.textPrimary}}>{fmt(l.amount)}{usdLine(l.amount, usdRate)}</td>
                                <td style={{padding:"9px 16px"}}><Badge text={PAYMENT_LABEL[l.status]} clr={PAYMENT_CLR[l.status]} bg={PAYMENT_BG[l.status]}/></td>
                              </tr>
                            ))}
                          </tbody>
                        </table></div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </div>
    </Shell>
  );
}
