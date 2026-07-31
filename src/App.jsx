import { useState, useEffect, useCallback, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import { supabase, sg, ss, sd } from "./supabase.js";
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
const todayStr = () => new Date().toISOString().slice(0,10);

// ─── Multi-code / multi-batch PO helpers ───────────────────────────────────
// A single PO can now be split across several Account Codes (each with its
// own amount that rolls up into the PO total) and can arrive in several
// delivery batches instead of a single date. Older records saved before this
// existed still carry a single `code`/`amount` and a single
// `incomingPlan`/`actualReceived` — these getters transparently upgrade them
// so both old and new records work everywhere without a one-off migration.
const poItems = (p) => (p.items && p.items.length ? p.items : [{ id:"legacy", code:p.code||"", amount:p.amount||"" }]);
const poTotal = (p) => poItems(p).reduce((s,it) => s + (parseFloat(it.amount)||0), 0);
const poAmountForCode = (p, code) => poItems(p).filter(it => it.code===code).reduce((s,it) => s + (parseFloat(it.amount)||0), 0);

// Pre-multi-supplier records stored one `deliveries` array (or a single
// `incomingPlan`/`actualReceived` pair) directly on the PO. Kept only to
// upgrade those older records — new records live under `suppliers[].rounds`.
const legacyDeliveries = (p) => (p.deliveries && p.deliveries.length
  ? p.deliveries
  : (p.incomingPlan || p.actualReceived ? [{ id:"legacy", plan:p.incomingPlan||"", actual:p.actualReceived||"" }] : []));

// ─── Multi-supplier / multi-round payment helpers ──────────────────────────
// A single PO can involve several suppliers (e.g. different vendors for the
// same order), and each supplier can be paid/delivered in several rounds
// (installments) instead of one lump sum. Records saved before this existed
// carried one `supplier` + `poNumber` string and one `deliveries` array —
// this getter transparently upgrades them into one supplier with N rounds so
// old and new records work everywhere without a one-off migration.
const poSuppliers = (p) => (p.suppliers && p.suppliers.length
  ? p.suppliers
  : [{
      id: "legacy",
      name: p.supplier || "",
      poNumber: p.poNumber || "",
      rounds: legacyDeliveries(p).length
        ? legacyDeliveries(p).map(d => ({ id: d.id || "legacy-round", amount: "", plan: d.plan || "", actual: d.actual || "" }))
        : [{ id: "legacy-round", amount: "", plan: "", actual: "" }],
    }]);

// Flattened view of every payment/delivery round across every supplier on a
// PO — each round is tagged with which supplier it belongs to. Also serves
// as the drop-in replacement for the old `poDeliveries`, since every round
// still carries the same `plan`/`actual` fields that `deliveryStatus` reads.
const poRounds = (p) => poSuppliers(p).flatMap(s =>
  (s.rounds && s.rounds.length ? s.rounds : [{ id:"legacy-round", amount:"", plan:"", actual:"" }])
    .map(r => ({ ...r, supplierId: s.id, supplierName: s.name, poNumber: s.poNumber }))
);
const poDeliveries = poRounds; // backward-compatible alias used by tracking/export code

const poSupplierNames  = (p) => poSuppliers(p).map(s => s.name).filter(Boolean);
const poSupplierText   = (p) => poSupplierNames(p).join(" ");
const poSupplierLabel  = (p) => { const n = poSupplierNames(p); return n.length===0 ? "—" : n.length===1 ? n[0] : `${n[0]} +${n.length-1} เจ้า`; };
const poNumbersLabel   = (p) => { const nums = poSuppliers(p).map(s=>s.poNumber).filter(Boolean); return nums.length ? nums.join(", ") : "—"; };
const poRoundsAmount   = (p) => poRounds(p).reduce((s,r)=>s+(parseFloat(r.amount)||0),0);

// Which supplier a given account-code line item was ordered from. Items
// carry their own `supplierId` so, on a PO with several suppliers, each
// item can be traced to exactly one of them; items saved before this link
// existed (or with a stale/missing id) fall back to the PO's first supplier.
const itemSupplier     = (p, it) => { const sups = poSuppliers(p); return sups.find(s=>s.id===it.supplierId) || sups[0] || null; };
const itemSupplierName = (p, it) => itemSupplier(p, it)?.name || "—";

const deliveryStatus = (d) => {
  if (d.actual) return "received";
  if (d.plan && d.plan < todayStr()) return "late";
  if (d.plan) return "pending";
  return "unset";
};
// PO-level incoming status aggregates every delivery batch: fully received
// only once every batch has arrived; "partial" once some (but not all)
// batches are in, so a PO that comes in 2-3 shipments is tracked accurately.
const incomingStatus = (p) => {
  const deliveries = poDeliveries(p);
  if (!deliveries.length) return "unset";
  const sts = deliveries.map(deliveryStatus);
  const receivedCount = sts.filter(s=>s==="received").length;
  if (receivedCount === deliveries.length) return "received";
  if (sts.some(s=>s==="late")) return "late";
  if (receivedCount > 0) return "partial";
  if (sts.some(s=>s==="pending")) return "pending";
  return "unset";
};
const paymentStatus = (p) => {
  if (p.status === "Paid") return "paid";
  if (p.paymentPlan && p.paymentPlan < todayStr()) return "late";
  if (p.paymentPlan) return "pending";
  return "unset";
};
const INCOMING_LABEL = { received:"รับแล้ว", partial:"รับบางส่วน", late:"ของเข้าล่าช้า", pending:"รอของเข้า", unset:"ยังไม่กำหนด" };
const INCOMING_CLR   = { received:"#10b981", partial:"#3b82f6", late:"#ef4444", pending:"#f59e0b", unset:"#94a3b8" };
const INCOMING_BG    = { received:"#f0fdf4", partial:"#eff6ff", late:"#fef2f2", pending:"#fffbeb", unset:"#f1f5f9" };
const PAYMENT_LABEL  = { paid:"จ่ายแล้ว", late:"เกินกำหนดจ่าย", pending:"รอจ่ายเงิน", unset:"ยังไม่กำหนด" };
const PAYMENT_CLR    = { paid:"#10b981", late:"#ef4444", pending:"#f59e0b", unset:"#94a3b8" };
const PAYMENT_BG     = { paid:"#f0fdf4", late:"#fef2f2", pending:"#fffbeb", unset:"#f1f5f9" };
const fmt  = n => new Intl.NumberFormat("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2}).format(n||0);
const fmtK = n => n>=1e6?`${(n/1e6).toFixed(1)}M`:n>=1e3?`${(n/1e3).toFixed(0)}K`:Math.round(n).toString();
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
`;

// ─── Excel Export ─────────────────────────────────────────────────────────────
function exportToExcel(project, tenderCosts, additions, poEntries) {
  const wb = XLSX.utils.book_new();

  // Budget = baseline Tender Cost + every monthly addition entered so far, combined per Acc. Code
  const combinedBudget = {...tenderCosts};
  Object.values(additions || {}).forEach(monthObj => {
    Object.entries(monthObj || {}).forEach(([code, val]) => {
      combinedBudget[code] = (parseFloat(combinedBudget[code]) || 0) + (parseFloat(val) || 0);
    });
  });

  const summaryRows = [];
  summaryRows.push([`Project: ${project.name}`, "", "", "", "", ""]);
  summaryRows.push([`Area: ${project.area} ft²`, "", `Panels: ${project.panels}`, "", "", ""]);
  summaryRows.push([`Export Date: ${new Date().toLocaleDateString("th-TH")}`, "", "", "", "", ""]);
  summaryRows.push([]);
  summaryRows.push(["Acc. Code","Account Name","Group","Budget / Tender Cost","Committed (PO)","Remaining","% Used","Status"]);
  let grandBudget=0, grandCommitted=0;
  ACCOUNTS.forEach(a => {
    const budget    = parseFloat(combinedBudget[a.code]) || 0;
    const committed = poEntries.reduce((s,p)=>s+poAmountForCode(p,a.code),0);
    const remaining = budget - committed;
    const pctUsed   = budget > 0 ? committed/budget : (committed>0?999:0);
    const status    = committed > budget && budget > 0 ? "OVER BUDGET" : committed>0 ? "OK" : budget>0 ? "No PO" : "-";
    if (budget > 0 || committed > 0) {
      summaryRows.push([a.code, a.name, a.group, budget, committed, remaining, pctUsed, status]);
      grandBudget += budget; grandCommitted += committed;
    }
  });
  summaryRows.push([]);
  summaryRows.push(["","TOTAL","",grandBudget,grandCommitted,grandBudget-grandCommitted,grandBudget>0?grandCommitted/grandBudget:0,""]);

  const ws1 = XLSX.utils.aoa_to_sheet(summaryRows);
  ws1["!cols"] = [{wch:12},{wch:40},{wch:16},{wch:20},{wch:20},{wch:18},{wch:10},{wch:14}];
  for (let r=5; r<summaryRows.length; r++) {
    const cellRef = XLSX.utils.encode_cell({r, c:6});
    if (ws1[cellRef] && typeof ws1[cellRef].v === "number") ws1[cellRef].z = "0.0%";
    ["D","E","F"].forEach((_,i) => {
      const ref = XLSX.utils.encode_cell({r, c:3+i});
      if (ws1[ref] && typeof ws1[ref].v === "number") ws1[ref].z = '#,##0.00';
    });
  }
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");

  const poRows = [];
  poRows.push([`Project: ${project.name}`, "", "", "", "", "", "", "", ""]);
  poRows.push([]);
  poRows.push(["PO Date","Acc. Code","Account Name","Group","Supplier","PO Number","Amount","Status","Suppliers & Rounds","Notes"]);
  // One row per Account Code line — a PO split across several codes gets one
  // row per code, each with just that code's share of the amount. A PO can
  // now involve several suppliers, each paid/delivered across several
  // rounds — these are summarised into one "name [PO]: plan→actual (amount)"
  // list per row so the multi-supplier detail isn't lost in the export.
  poEntries.forEach(p => {
    const supplierStr = poSuppliers(p).map(s => {
      const roundsStr = (s.rounds||[]).map(r => `${r.plan||"-"}→${r.actual||"รอ"}${r.amount?` (${fmt(r.amount)})`:""}`).join(" & ") || "-";
      return `${s.name||"—"}${s.poNumber?` [PO:${s.poNumber}]`:""}: ${roundsStr}`;
    }).join(" | ") || "-";
    poItems(p).forEach(it => {
      const acc = ACCOUNTS.find(a=>a.code===it.code);
      poRows.push([p.date, it.code, acc?.name||"", acc?.group||"", itemSupplierName(p,it), poNumbersLabel(p), parseFloat(it.amount)||0, p.status, supplierStr, p.notes||""]);
    });
  });
  poRows.push([]);
  poRows.push(["","","","","","TOTAL", poEntries.reduce((s,p)=>s+poTotal(p),0), "","",""]);
  const ws2 = XLSX.utils.aoa_to_sheet(poRows);
  ws2["!cols"] = [{wch:12},{wch:10},{wch:38},{wch:14},{wch:24},{wch:16},{wch:18},{wch:12},{wch:34},{wch:30}];
  for (let r=2; r<poRows.length; r++) {
    const ref = XLSX.utils.encode_cell({r, c:6});
    if (ws2[ref] && typeof ws2[ref].v === "number") ws2[ref].z = '#,##0.00';
  }
  XLSX.utils.book_append_sheet(wb, ws2, "PO Entries");

  const grpRows = [["Group","Budget","Committed","Remaining","% Used"]];
  GROUPS.forEach(g => {
    const codes = ACCOUNTS.filter(a=>a.group===g).map(a=>a.code);
    const b = codes.reduce((s,c)=>s+(parseFloat(combinedBudget[c])||0),0);
    const c2 = poEntries.reduce((s,p)=>s+poItems(p).filter(it=>codes.includes(it.code)).reduce((s2,it)=>s2+(parseFloat(it.amount)||0),0),0);
    if (b>0||c2>0) grpRows.push([g,b,c2,b-c2,b>0?c2/b:0]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(grpRows), "By Group");

  XLSX.writeFile(wb, `TenderCost_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── Root ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [session,  setSessionState] = useState(() => getSession());
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

  const handleLogin = (user) => { setSession(user); setSessionState(user); };
  const handleLogout = () => {
    clearSession(); setSessionState(null);
    setScreen("home"); setRole(null); setActiveId(null);
  };

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

  useEffect(() => {
    (async () => { await fetchProjects(); setLoaded(true); setSyncedAt(new Date()); })();
  }, [fetchProjects]);

  useEffect(() => { if (!activeId) return; fetchProjectData(activeId); }, [activeId, fetchProjectData]);

  useEffect(() => {
    const channel = supabase.channel("kv_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "kv_store" }, async (payload) => {
        const key = payload.new?.key || payload.old?.key || "";
        setSyncing(true);
        if (key === "tcs-projects") await fetchProjects();
        else if (activeId && (key === `tcs-tenders-${activeId}` || key === `tcs-po-${activeId}` || key === `tcs-additions-${activeId}` || key === `tcs-extra-${activeId}` || key === `tcs-hidden-${activeId}`)) await fetchProjectData(activeId);
        setSyncedAt(new Date()); setSyncing(false);
      }).subscribe();
    return () => supabase.removeChannel(channel);
  }, [activeId, fetchProjects, fetchProjectData]);

  const saveProjects = useCallback((list) => { setProjects(list); ss("tcs-projects", list).then(()=>setSyncedAt(new Date())); }, []);
  const saveTenders  = useCallback((t)    => { setTCosts(t);      ss(`tcs-tenders-${activeId}`, t).then(()=>setSyncedAt(new Date())); }, [activeId]);
  const saveAdditions= useCallback((a)    => { setAdditions(a);   ss(`tcs-additions-${activeId}`, a).then(()=>setSyncedAt(new Date())); }, [activeId]);
  const saveExtraItems=useCallback((ex)   => { setExtraItems(ex); ss(`tcs-extra-${activeId}`, ex).then(()=>setSyncedAt(new Date())); }, [activeId]);
  const saveHiddenAccounts=useCallback((h)=> { setHiddenAccounts(h); ss(`tcs-hidden-${activeId}`, h).then(()=>setSyncedAt(new Date())); }, [activeId]);
  const savePO       = useCallback((po)   => { setPO(po);         ss(`tcs-po-${activeId}`, po).then(()=>setSyncedAt(new Date())); }, [activeId]);

  const openProject = (id) => {
    setActiveId(id);
    if (session?.role === "admin") { setRole(null); setScreen("roleSelect"); }
    else { setRole(session?.role); setScreen("app"); }
  };
  const deleteProject = async (id) => {
    if (!confirm("ลบโครงการนี้? ข้อมูลทั้งหมดจะหายถาวร")) return;
    saveProjects(projects.filter(p => p.id !== id));
    await sd(`tcs-tenders-${id}`); await sd(`tcs-po-${id}`); await sd(`tcs-additions-${id}`); await sd(`tcs-extra-${id}`); await sd(`tcs-hidden-${id}`);
  };
  const activeProject = projects.find(p => p.id === activeId) || { name:"", area:"", panels:"" };
  const updateProject = (fields) => saveProjects(projects.map(p => p.id === activeId ? {...p,...fields} : p));

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
    syncedAt, syncing, session, onLogout: handleLogout };

  return (
    <>
      <style>{GLOBAL_CSS}</style>
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
      {screen === "app" && effectiveRole === "qs"          && <QSView          {...sharedProps} />}
      {screen === "app" && effectiveRole === "procurement" && <ProcurementView {...sharedProps} />}
      {screen === "app" && effectiveRole === "accounting"  && (
        <AccountingView {...sharedProps} onExport={() => exportToExcel(activeProject, tenderCosts, additions, poEntries)} />
      )}
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
    const [u, l] = await Promise.all([loadUsers(), loadLogs()]);
    setUsers(u); setLogs(l); setLoadedU(true);
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
        <button onClick={onBack} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 12px",fontSize:18}}>←</button>
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
          {[["users","👥 จัดการผู้ใช้"],["logs","📜 Log การเข้าใช้งาน"]].map(([id,label])=>(
            <button key={id} onClick={()=>setTab(id)}
              style={{background:tab===id?T.blue:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.blue:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {label}
            </button>
          ))}
        </div>

        {!loaded ? (
          <div style={{color:T.textMuted,fontSize:13}}>กำลังโหลด...</div>
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
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
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
            </table>
          </div>
        ) : (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
            <div style={{padding:"16px 18px",borderBottom:`1px solid ${T.cardBorder}`,fontSize:13,fontWeight:600,color:T.textPrimary}}>
              ประวัติการเข้าใช้งานล่าสุด ({logs.length})
            </div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
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
            </table>
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

// ─── StatCard ─────────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, color, icon, accent }) {
  return (
    <div style={{background:T.card,borderRadius:14,padding:"20px 22px",border:`1px solid ${T.cardBorder}`,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:color,borderRadius:"14px 14px 0 0"}}/>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
        <div style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{label}</div>
        {icon && <div style={{width:34,height:34,borderRadius:10,background:accent||T.blueLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>{icon}</div>}
      </div>
      <div style={{fontSize:22,fontWeight:700,color:T.textPrimary,letterSpacing:"-0.5px",fontFamily:"'JetBrains Mono',monospace"}}>{value}</div>
      {sub && <div style={{fontSize:11,color:T.textMuted,marginTop:5}}>{sub}</div>}
    </div>
  );
}

// ─── Home Screen ──────────────────────────────────────────────────────────────
function HomeScreen({ projects, saveProjects, openProject, deleteProject, newProjModal, setNewProjModal, syncedAt, syncing, session, onLogout, onOpenAdmin }) {
  const [draft, setDraft] = useState({ name:"", area:"", panels:"", client:"", currency:"THB" });

  const createProject = () => {
    if (!draft.name.trim()) return;
    const id = uid();
    saveProjects([...projects, { ...draft, id, createdAt: new Date().toISOString() }]);
    setNewProjModal(false);
    setDraft({ name:"", area:"", panels:"", client:"", currency:"THB" });
  };

  return (
    <div style={{minHeight:"100vh",background:T.bg}}>
      {/* Header */}
      <div style={{background:T.headerGrad,padding:"0 32px"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"18px 0 20px"}}>
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
            <button className="btn-primary" onClick={()=>setNewProjModal(true)}
              style={{background:"rgba(255,255,255,0.2)",backdropFilter:"blur(8px)",border:"1.5px solid rgba(255,255,255,0.3)",display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:16,lineHeight:1}}>+</span> โครงการใหม่
            </button>
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
            <div style={{fontSize:12,color:T.textMuted,marginBottom:18,fontWeight:500}}>{projects.length} โครงการทั้งหมด</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:20}}>
              {projects.map(p => <ProjectCard key={p.id} project={p} onOpen={()=>openProject(p.id)} onDelete={()=>deleteProject(p.id)} />)}
            </div>
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
              ].map(([label,key,type,col]) => (
                <label key={key} style={{display:"flex",flexDirection:"column",gap:6,gridColumn:col}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{label}</span>
                  <input type={type} value={draft[key]} onChange={e=>setDraft(d=>({...d,[key]:e.target.value}))} className="input-base"/>
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
  const age = Math.floor((Date.now() - new Date(project.createdAt)) / 86400000);
  return (
    <div className="card-hover" style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:24,cursor:"pointer",position:"relative"}}>
      <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:T.headerGrad,borderRadius:"16px 16px 0 0"}}/>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14,paddingTop:2}}>
        <span style={{fontSize:10,letterSpacing:2,color:T.blue,fontWeight:700,textTransform:"uppercase"}}>PROJECT</span>
        <button onClick={e=>{e.stopPropagation();onDelete();}} style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",fontSize:14,padding:4,borderRadius:6,transition:"color 0.15s"}}
          onMouseEnter={e=>e.target.style.color="#ef4444"} onMouseLeave={e=>e.target.style.color=T.textMuted}>🗑</button>
      </div>
      <div style={{fontSize:18,fontWeight:700,color:T.textPrimary,marginBottom:4,lineHeight:1.3}}>{project.name}</div>
      {project.client && <div style={{fontSize:12,color:T.textSecondary,marginBottom:14}}>{project.client}</div>}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
        {project.area   && <span style={{background:T.blueLight,color:T.blue,fontSize:11,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.area} ft²</span>}
        {project.panels && <span style={{background:T.blueLight,color:T.blue,fontSize:11,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.panels} Panels</span>}
        {project.currency && <span style={{background:"#f8fafc",color:T.textMuted,fontSize:11,padding:"3px 10px",borderRadius:6,fontWeight:500}}>{project.currency}</span>}
      </div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div style={{fontSize:11,color:T.textMuted}}>{age === 0 ? "สร้างวันนี้" : `${age} วันที่แล้ว`}</div>
        <button onClick={onOpen} className="btn-primary" style={{padding:"8px 18px",fontSize:12}}>เปิดโครงการ →</button>
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
        <button onClick={onBack} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 12px",fontSize:18}}>←</button>
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
              {[["ชื่อโครงการ","name","text"],["ลูกค้า","client","text"],["สกุลเงิน","currency","text"],["พื้นที่ (ft²)","area","number"],["Panels","panels","number"]].map(([l,k,t]) => (
                <label key={k} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{l}</span>
                  <input type={t} value={draft[k]||""} onChange={e=>setDraft(d=>({...d,[k]:e.target.value}))} className="input-base"/>
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
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:20,maxWidth:800}}>
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
function Shell({ role, color, project, onBack, children, syncedAt, syncing, session, onLogout }) {
  const labels = {qs:"QS · Quantity Surveyor",procurement:"จัดซื้อ · Procurement",accounting:"บัญชี · Accounting"};
  const gradients = {
    qs:          "linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)",
    procurement: "linear-gradient(135deg, #78350f 0%, #d97706 100%)",
    accounting:  "linear-gradient(135deg, #064e3b 0%, #10b981 100%)",
  };
  return (
    <div style={{minHeight:"100vh",background:T.bg,display:"flex",flexDirection:"column"}}>
      <div style={{background:gradients[role],padding:"14px 28px",display:"flex",alignItems:"center",gap:14}}>
        <button onClick={onBack} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",cursor:"pointer",borderRadius:8,padding:"6px 12px",fontSize:18}}>←</button>
        <div style={{flex:1}}>
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
function QSView({ project, tenderCosts, saveTenders, additions, saveAdditions, extraItems, saveExtraItems, hiddenAccounts, saveHiddenAccounts, onBack, syncedAt, syncing, session, onLogout }) {
  const [tab, setTab] = useState("baseline"); // "baseline" | "monthly"

  // Shared "add / remove line item" logic — used by both Baseline and Monthly tabs,
  // and kept in sync with tenderCosts + every month's additions on delete.
  // Two kinds of extra item:
  //  - standalone (has `group`): a brand-new scope item with its own Acc-like code
  //  - sub-item   (has `parentCode`): a breakdown line that rolls up INTO an existing Acc. Code
  const handleAddExtraItem = ({ name, group, parentCode, code, scope }) => {
    if (!name.trim()) return;
    const item = parentCode
      ? { code:`EX-${uid()}`, name:name.trim(), parentCode, scope }
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
    <Shell role="qs" color={T.blue} project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"20px 28px 0"}}>
        <div style={{display:"flex",gap:8,marginBottom:20}}>
          {[["baseline","📐 ราคาเดิม (Baseline)"],["monthly","📅 รายการเพิ่มรายเดือน"]].map(([id,label])=>(
            <button key={id} onClick={()=>setTab(id)}
              style={{background:tab===id?T.blue:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.blue:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
              {label}
            </button>
          ))}
        </div>
      </div>
      {tab === "baseline"
        ? <QSBaselineTab tenderCosts={tenderCosts} saveTenders={saveTenders} extraItems={extraItems}
                         onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} onHideAccount={handleHideAccount} onRestoreAccount={handleRestoreAccount} />
        : <QSMonthlyTab tenderCosts={tenderCosts} additions={additions} saveAdditions={saveAdditions}
                         extraItems={extraItems} onAddExtra={handleAddExtraItem} onDeleteExtra={handleDeleteExtraItem}
                         hiddenAccounts={hiddenAccounts} />}
    </Shell>
  );
}

// ─── QS Tab 1: Baseline (original tender cost) ────────────────────────────────
function QSBaselineTab({ tenderCosts, saveTenders, extraItems, onAddExtra, onDeleteExtra, hiddenAccounts, onHideAccount, onRestoreAccount }) {
  const [draft,  setDraft]  = useState({...tenderCosts});
  const [filter, setFilter] = useState("All");
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

  useEffect(() => setDraft({...tenderCosts}), [tenderCosts]);

  // Sub-items (e.g. "Silicone Structure") roll up into an existing Acc. Code (e.g. 511025).
  // Standalone extras (no parentCode) are brand-new items with their own group, shown as their own row.
  // Only show sub-items created here on the Baseline tab — Monthly-tab sub-items
  // are a separate breakdown and are intentionally kept off this tab.
  const subItemsByParent = {};
  extraItems.forEach(e => {
    if (e.parentCode && e.scope !== "monthly") (subItemsByParent[e.parentCode] = subItemsByParent[e.parentCode] || []).push(e);
  });
  const standaloneExtras = extraItems.filter(e => !e.parentCode);
  const childrenOf = (code) => subItemsByParent[code] || [];

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
    if (filter!=="All" && a.group!==filter) return false;
    const selfMatch = a.name.toLowerCase().includes(q) || a.code.includes(search);
    const childMatch = !a.isExtra && childrenOf(a.code).some(k=>k.name.toLowerCase().includes(q));
    return selfMatch || childMatch;
  });

  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => -d);
    else { setSortKey(key); setSortDir(1); }
  };
  const displayRows = (() => {
    if (!sortKey) return filtered;
    const arr = [...filtered];
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
    Object.entries(merged).forEach(([k,v]) => { if(v!==""&&!isNaN(v)&&parseFloat(v)>0) clean[k]=parseFloat(v); });
    saveTenders(clean);
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
    onAddExtra({ name:subName, parentCode, scope:"baseline" });
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
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:24}}>
        <StatCard label="Tender Cost รวม" value={fmt(base)} sub="Base cost ทั้งหมด (ราคาเดิม)" color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label="Spare & Wastage 3%" value={fmt(adj3)} sub="เผื่อสูญหาย" color={T.amber} icon="⚙️" accent={T.amberBg}/>
        <StatCard label="Total Adjusted" value={fmt(total)} sub="ต้นทุนรวมสุทธิ" color={T.green} icon="✅" accent={T.greenBg}/>
      </div>

      {/* Filters + Add row + Save */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา Account Code / ชื่อ..."
          className="input-base" style={{width:240}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap",flex:1}}>
          {["All",...GROUPS].map(g=>(
            <button key={g} onClick={()=>setFilter(g)}
              style={{background:filter===g?T.blue:"transparent",border:`1.5px solid ${filter===g?T.blue:T.cardBorder}`,borderRadius:8,padding:"4px 11px",color:filter===g?"#fff":T.textSecondary,fontSize:11,cursor:"pointer",fontWeight:500,transition:"all 0.15s"}}>{g}</button>
          ))}
        </div>
        {hiddenList.length > 0 && (
          <button className="btn-ghost" onClick={()=>setShowHidden(v=>!v)} style={{color:T.textMuted}}>
            🗂 ที่ซ่อนไว้ ({hiddenList.length})
          </button>
        )}
        <button className="btn-ghost" onClick={()=>setAddOpen(v=>!v)}>+ เพิ่มรายการหลักใหม่</button>
        <button onClick={handleSave} className="btn-primary"
          style={{background:saved?T.green:T.blue,minWidth:140}}>
          {saved?"✓ บันทึกแล้ว":"บันทึก Tender Cost"}
        </button>
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
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
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
                      {!a.isExtra && (
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
                        </div>
                      ) : (
                        <input type="number" value={draft[a.code]??""} onChange={e=>setDraft(d=>({...d,[a.code]:e.target.value}))} onClick={e=>e.stopPropagation()}
                          placeholder="0.00" className="input-base" style={{width:160,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:draft[a.code]>0?T.blueLight:T.bg}}/>
                      )}
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"center"}}>
                      {a.isExtra
                        ? <button onClick={(e)=>{e.stopPropagation(); handleDeleteRow(a.code);}} title="ลบรายการนี้"
                            style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:14}}>✕</button>
                        : <button onClick={(e)=>{e.stopPropagation(); onHideAccount(a.code);}} title="นำ Acc. Code นี้ออกจากรายการหลัก (กู้คืนได้)"
                            style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",fontSize:14}}>✕</button>}
                    </td>
                  </tr>

                  {/* Sub-items — roll up into the parent Acc. Code's total above */}
                  {!isCollapsed && kids.map((k,ki)=>(
                    <tr key={k.code} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(ki===kids.length-1 && subFor!==a.code)?"1px solid #f1f5f9":"none"}}>
                      <td style={{padding:"6px 16px 6px 30px",color:T.green,fontSize:12}}>↳</td>
                      <td/>
                      <td style={{padding:"6px 16px",color:T.green,fontSize:12,fontStyle:"italic"}}>{k.name}</td>
                      <td style={{padding:"6px 16px",textAlign:"right"}}>
                        <input type="number" value={draft[k.code]??""} onChange={e=>setDraft(d=>({...d,[k.code]:e.target.value}))}
                          placeholder="0.00" className="input-base" style={{width:160,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12,background:draft[k.code]>0?T.greenBg:T.bg}}/>
                      </td>
                      <td style={{padding:"6px 16px",textAlign:"center"}}>
                        <button onClick={()=>handleDeleteRow(k.code)} title="ลบรายการย่อยนี้"
                          style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
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
              </td>
              <td/>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

// ─── QS Tab 2: Monthly additions (เดิม / เพิ่มเดือนนี้ / รวมสะสม) ─────────────
function QSMonthlyTab({ tenderCosts, additions, saveAdditions, extraItems, onAddExtra, onDeleteExtra, hiddenAccounts }) {
  const thisMonth = new Date().toISOString().slice(0,7);
  const months = Object.keys(additions).sort();
  const [month, setMonth] = useState(months.length ? months[months.length-1] : thisMonth);
  const [newMonth, setNewMonth] = useState("");
  const [filter, setFilter] = useState("All");
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

  useEffect(() => { setDraftAdd({...(additions[month]||{})}); }, [month, additions]);
  useEffect(() => { if (!months.includes(month) && months.length) setMonth(months[months.length-1]); }, [months]); // eslint-disable-line

  // All rows = original 70 account codes + standalone extra items.
  // Sub-items (parentCode set) can be added right here for a monthly
  // breakdown, or on the Baseline tab for a baseline breakdown — either way
  // they roll up into their parent row's figures and aren't listed on their own.
  const allRows = [...ACCOUNTS.filter(a=>!hiddenAccounts.includes(a.code)), ...extraItems.filter(e=>!e.parentCode).map(e => ({ code:e.code, name:e.name, group:e.group, isExtra:true }))];

  const subItemsByParent = {};
  extraItems.forEach(e => { if (e.parentCode && e.scope === "monthly") (subItemsByParent[e.parentCode] = subItemsByParent[e.parentCode] || []).push(e); });
  const childrenOf = (code) => subItemsByParent[code] || [];

  // A row's monthly figure is the sum of its sub-items' figures when it has
  // any (mirrors the Baseline tab), otherwise its own entered value.
  const rowMonthValue = (code, m, draft) => {
    const kids = childrenOf(code);
    if (kids.length) return kids.reduce((s,k)=>s+(parseFloat((draft||additions[m])?.[k.code])||0),0);
    return parseFloat((draft||additions[m])?.[code]) || 0;
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
  const monthShortLabel = (m) => new Date(m+"-01").toLocaleDateString("th-TH",{month:"short",year:"2-digit"});
  // Each bar = one stacked column: "previous" (running total up to the
  // month before) + "added" (that month's increment) in a different color,
  // so growth is visible within a single bar instead of a smooth area line.
  const chartData = [
    { label:"เริ่มต้น", cumulative: baseTotal, previous: baseTotal, added: 0 },
    ...sortedMonths.map(m => {
      const added = monthTotalLive(m);
      const cumulative = cumulativeLive(m);
      return { label: monthShortLabel(m), cumulative, previous: cumulative - added, added };
    }),
  ];

  // "ราคาเดิม (Baseline)" should reflect the running total as of the month
  // BEFORE the one currently selected — not the fixed original baseline —
  // so it moves forward as prior months get their additions saved.
  const priorMonths      = sortedMonths.filter(m => m < month);
  const prevMonthLabel   = priorMonths.length ? monthShortLabel(priorMonths[priorMonths.length-1]) : "เริ่มต้น";
  const baselineForMonth = baseTotal + priorMonths.reduce((s,m)=>s+monthTotalLive(m),0);

  const filtered = allRows.filter(r => {
    if (filter!=="All" && r.group!==filter) return false;
    const q = search.toLowerCase();
    const selfMatch = r.name.toLowerCase().includes(q) || r.code.includes(search);
    const childMatch = childrenOf(r.code).some(k=>k.name.toLowerCase().includes(q));
    return selfMatch || childMatch;
  });

  // Mirrors the per-row figures computed inline in the table body, so header
  // sorting can order rows by the same "ยอดก่อนหน้า / เพิ่มเดือนนี้ / รวมสะสม" values shown.
  const cumBeforeOf = (r) => months.filter(m=>m<month).reduce((s,m)=>s+(parseFloat(additions[m]?.[r.code])||0),0) + (parseFloat(tenderCosts[r.code])||0);
  const cumOf = (r) => cumBeforeOf(r) + rowMonthValue(r.code, month, draftAdd);

  const handleSort = (key) => {
    if (sortKey === key) setSortDir(d => -d);
    else { setSortKey(key); setSortDir(1); }
  };
  const displayRows = (() => {
    if (!sortKey) return filtered;
    const arr = [...filtered];
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

  const handleAddMonth = () => {
    if (!newMonth || months.includes(newMonth)) return;
    saveAdditions({ ...additions, [newMonth]: additions[newMonth] || {} });
    setMonth(newMonth); setNewMonth("");
  };

  const handleSave = () => {
    const merged = {...draftAdd};
    allRows.forEach(r => {
      const kids = childrenOf(r.code);
      if (kids.length) merged[r.code] = kids.reduce((s,k)=>s+(parseFloat(merged[k.code])||0),0);
    });
    const clean = {};
    Object.entries(merged).forEach(([k,v]) => { if(v!==""&&!isNaN(v)&&parseFloat(v)!==0) clean[k]=parseFloat(v); });
    saveAdditions({ ...additions, [month]: clean });
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
    onAddExtra({ name:subName, parentCode, scope:"monthly" });
    setRowCollapsed(c => ({...c, [parentCode]: false})); // reveal the newly-added sub-item
    setSubName(""); setSubFor(null);
  };

  const handleDeleteExtra = (code) => {
    onDeleteExtra(code);
    setDraftAdd(d => { const n = {...d}; delete n[code]; return n; });
  };

  return (
    <div style={{padding:"4px 28px 24px"}}>
      {/* Trend chart — the whole project's cost growth over time, at a glance */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"18px 20px 8px",marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:6,flexWrap:"wrap",gap:6}}>
          <span style={{fontSize:13,fontWeight:700,color:T.textPrimary}}>📈 แนวโน้มต้นทุนสะสม</span>
          <span style={{fontSize:12,color:T.textMuted}}>รวมล่าสุดทั้งโปรเจกต์: <b style={{color:T.green,fontFamily:"'JetBrains Mono',monospace",fontSize:15}}>{fmt(grandTotal)}</b> THB</span>
        </div>
        <ResponsiveContainer width="100%" height={170}>
          <BarChart data={chartData} margin={{top:8,right:8,left:-18,bottom:0}}>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f7"/>
            <XAxis dataKey="label" tick={{fontSize:11,fill:T.textMuted}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fontSize:10,fill:T.textMuted}} axisLine={false} tickLine={false} tickFormatter={fmtK}/>
            <Tooltip formatter={(v,name)=>[`${fmt(v)} THB`,name]} labelStyle={{color:T.textPrimary,fontWeight:600,marginBottom:2}}
              contentStyle={{borderRadius:10,border:`1px solid ${T.cardBorder}`,fontSize:12,boxShadow:"0 4px 14px rgba(0,0,0,0.08)"}}/>
            <Bar dataKey="previous" stackId="cum" name="ยอดก่อนหน้า" fill={T.blue} radius={[0,0,0,0]}/>
            <Bar dataKey="added" stackId="cum" name="เพิ่มงวดนี้" fill={T.amber} radius={[4,4,0,0]}/>
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* Month picker — horizontal chips, click any to switch, "+" chip to add a new month */}
      <div style={{display:"flex",gap:8,overflowX:"auto",paddingBottom:6,marginBottom:16}}>
        {sortedMonths.map(m=>{
          const active = m===month;
          const add = monthTotalLive(m);
          return (
            <button key={m} onClick={()=>setMonth(m)}
              style={{flexShrink:0,textAlign:"left",padding:"10px 16px",borderRadius:12,border:`1.5px solid ${active?T.blue:T.cardBorder}`,
                background:active?T.blue:T.card,cursor:"pointer",minWidth:140,transition:"all 0.15s"}}>
              <div style={{fontSize:11,fontWeight:600,color:active?"#bfdbfe":T.textSecondary,marginBottom:3}}>{monthShortLabel(m)}</div>
              <div style={{fontSize:15,fontWeight:700,color:active?"#fff":T.textPrimary,fontFamily:"'JetBrains Mono',monospace"}}>{fmtK(cumulativeLive(m))}</div>
              <div style={{fontSize:10,color:active?"#dbeafe":T.textMuted,marginTop:2}}>{add>0?"+":""}{fmtK(add)} เดือนนี้</div>
            </button>
          );
        })}
        <div style={{flexShrink:0,display:"flex",alignItems:"center",gap:6,padding:"0 12px",borderRadius:12,border:`1.5px dashed ${T.cardBorder}`}}>
          <input type="month" value={newMonth} onChange={e=>setNewMonth(e.target.value)} className="input-base"
            style={{border:"none",background:"transparent",padding:"8px 4px",width:118,fontSize:12}}/>
          <button className="btn-ghost" style={{padding:"6px 12px",fontSize:11,whiteSpace:"nowrap"}} onClick={handleAddMonth}>+ เพิ่มเดือน</button>
        </div>
      </div>

      {/* Stats for selected month */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:16,marginBottom:20}}>
        <StatCard label="ราคาเดิม (Baseline)" value={fmt(baselineForMonth)} sub={`สะสมถึง ${prevMonthLabel}`} color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label="เพิ่มเดือนนี้" value={fmt(thisMonthAdd)} sub={new Date(month+"-01").toLocaleDateString("th-TH",{year:"numeric",month:"long"})} color={T.amber} icon="➕" accent={T.amberBg}/>
        <StatCard label="รวมสะสมถึงเดือนนี้" value={fmt(cumulativeSoFar)} sub="เดิม + เพิ่มสะสมถึงเดือนที่เลือก" color={T.green} icon="✅" accent={T.greenBg}/>
        <StatCard label="รวมทั้งหมด" value={fmt(grandTotal)} sub="เดิม + ทุกเดือนที่มีข้อมูล (ล่าสุด)" color={T.purple} icon="🧮" accent={T.purpleBg}/>
      </div>

      {/* Toolbar: search + group filter + actions */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา Account Code / ชื่อ..."
          className="input-base" style={{width:220}}/>
        <div style={{display:"flex",gap:5,flexWrap:"wrap",flex:1}}>
          {["All",...GROUPS].map(g=>(
            <button key={g} onClick={()=>setFilter(g)}
              style={{background:filter===g?T.blue:"transparent",border:`1.5px solid ${filter===g?T.blue:T.cardBorder}`,borderRadius:8,padding:"4px 11px",color:filter===g?"#fff":T.textSecondary,fontSize:11,cursor:"pointer",fontWeight:500,transition:"all 0.15s"}}>{g}</button>
          ))}
        </div>
        <button className="btn-ghost" onClick={()=>setAddExtraOpen(v=>!v)}>+ งานพิเศษ</button>
        <button onClick={handleSave} className="btn-primary" style={{background:saved?T.green:T.blue,minWidth:170}}>
          {saved?"✓ บันทึกแล้ว":`บันทึกรายการเดือนนี้`}
        </button>
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
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead>
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
                  <span onClick={()=>key&&handleSort(key)} style={{cursor:key?"pointer":"default",userSelect:"none"}}>{label}{key && sortKey===key ? (sortDir===1?" ▲":" ▼") : ""}</span>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayRows.map((r,i) => {
              const kids = childrenOf(r.code);
              const hasKids = kids.length > 0;
              const isCollapsed = hasKids && rowCollapsed[r.code];
              const baseVal = parseFloat(tenderCosts[r.code]) || 0;
              const cumBefore = months.filter(m=>m<month).reduce((s,m)=>s+(parseFloat(additions[m]?.[r.code])||0),0) + baseVal;
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
                      <button onClick={(e)=>{e.stopPropagation(); setSubFor(subFor===r.code?null:r.code); setSubName(""); setRowCollapsed(c=>({...c,[r.code]:false}));}} title="เพิ่มรายการย่อยใต้ Acc. Code นี้"
                        style={{marginLeft:9,background:"none",border:`1px dashed ${T.cardBorder}`,borderRadius:6,color:T.textMuted,cursor:"pointer",fontSize:10,padding:"1px 7px"}}>
                        + รายการย่อย
                      </button>
                    </td>
                    <td style={{padding:"8px 16px",textAlign:"right",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace"}} title="ราคาเดิม + ยอดเพิ่มของทุกเดือนก่อนหน้ารวมกัน">{fmt(cumBefore)}</td>
                    <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                    <td style={{padding:"8px 16px",textAlign:"right"}}>
                      {hasKids ? (
                        <div style={{width:130,marginLeft:"auto",padding:"7px 10px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:T.amberBg,borderRadius:8,color:T.amber,fontWeight:700,fontSize:13}}>
                          {fmt(thisVal)}
                        </div>
                      ) : (
                        <input type="number" value={draftAdd[r.code]??""} onChange={e=>setDraftAdd(d=>({...d,[r.code]:e.target.value}))} onClick={e=>e.stopPropagation()}
                          placeholder="0.00" className="input-base" style={{width:130,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:thisVal!==0?T.amberBg:T.bg}}/>
                      )}
                    </td>
                    <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                    <td style={{padding:"8px 16px",textAlign:"right",color:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{fmt(cum)}</td>
                    <td style={{padding:"8px 16px",textAlign:"center"}}>
                      {r.isExtra && (
                        <button onClick={(e)=>{e.stopPropagation(); handleDeleteExtra(r.code);}} title="ลบรายการงานเพิ่ม"
                          style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
                      )}
                    </td>
                  </tr>

                  {/* Sub-items — this month's value rolls up into the parent row above */}
                  {!isCollapsed && kids.map((k,ki) => {
                    const kBaseVal = parseFloat(tenderCosts[k.code]) || 0;
                    const kCumBefore = months.filter(m=>m<month).reduce((s,m)=>s+(parseFloat(additions[m]?.[k.code])||0),0) + kBaseVal;
                    const kThisVal = parseFloat(draftAdd[k.code]) || 0;
                    const kCum = kCumBefore + kThisVal;
                    return (
                      <tr key={k.code} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:(ki===kids.length-1 && subFor!==r.code)?"1px solid #f1f5f9":"none"}}>
                        <td style={{padding:"6px 16px 6px 30px",color:T.green,fontSize:12}}>↳</td>
                        <td/>
                        <td style={{padding:"6px 16px",color:T.green,fontSize:12,fontStyle:"italic"}}>{k.name}</td>
                        <td style={{padding:"6px 16px",textAlign:"right",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{fmt(kCumBefore)}</td>
                        <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                        <td style={{padding:"6px 16px",textAlign:"right"}}>
                          <input type="number" value={draftAdd[k.code]??""} onChange={e=>setDraftAdd(d=>({...d,[k.code]:e.target.value}))}
                            placeholder="0.00" className="input-base" style={{width:130,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",fontSize:12,background:kThisVal!==0?T.greenBg:T.bg}}/>
                        </td>
                        <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                        <td style={{padding:"6px 16px",textAlign:"right",color:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:600,fontSize:12}}>{fmt(kCum)}</td>
                        <td style={{padding:"6px 16px",textAlign:"center"}}>
                          <button onClick={()=>handleDeleteExtra(k.code)} title="ลบรายการย่อยนี้"
                            style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
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
                      <td colSpan={5} style={{padding:"7px 16px",display:"flex",gap:6,justifyContent:"flex-end"}}>
                        <button className="btn-primary" style={{padding:"5px 12px",fontSize:12}} onClick={()=>handleAddSub(r.code)}>+ เพิ่ม</button>
                        <button className="btn-ghost" style={{padding:"5px 12px",fontSize:12}} onClick={()=>setSubFor(null)}>ยกเลิก</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={9} style={{padding:"28px 16px",textAlign:"center",color:T.textMuted,fontSize:13}}>ไม่พบรายการที่ตรงกับการค้นหา</td></tr>
            )}
          </tbody>
          <tfoot>
            <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
              <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:12}}>{filtered.length} รายการ</td>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontWeight:600,fontSize:13}}>
                {fmt(filtered.reduce((s,r)=>{
                  const baseVal = parseFloat(tenderCosts[r.code]) || 0;
                  const cumBefore = months.filter(m=>m<month).reduce((ss,m)=>ss+(parseFloat(additions[m]?.[r.code])||0),0)+baseVal;
                  return s + cumBefore;
                },0))}
              </td>
              <td/>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:13}}>
                {fmt(filtered.reduce((s,r)=>s+rowMonthValue(r.code, month, draftAdd),0))}
              </td>
              <td/>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:14}}>
                {fmt(filtered.reduce((s,r)=>{
                  const baseVal = parseFloat(tenderCosts[r.code]) || 0;
                  const cumBefore = months.filter(m=>m<month).reduce((ss,m)=>ss+(parseFloat(additions[m]?.[r.code])||0),0)+baseVal;
                  return s + cumBefore + rowMonthValue(r.code, month, draftAdd);
                },0))}
              </td>
              <td/>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

// ─── Procurement: PO Detail Modal ──────────────────────────────────────────────
// Read-only detail view opened by clicking any PO row. Lets the user confirm
// exactly what was entered without hunting through a wide table, and offers
// Edit / Delete from the same place.
function PODetailModal({ po, onClose, onEdit, onDelete }) {
  if (!po) return null;
  const items = poItems(po);
  const suppliers = poSuppliers(po);
  const inc = incomingStatus(po), pay = paymentStatus(po);

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

        <div style={{display:"flex",gap:6,margin:"12px 0 4px",flexWrap:"wrap"}}>
          <span style={{background:STATUS_BG[po.status],color:STATUS_CLR[po.status],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{po.status}</span>
          <span style={{background:INCOMING_BG[inc],color:INCOMING_CLR[inc],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{INCOMING_LABEL[inc]}</span>
          <span style={{background:PAYMENT_BG[pay],color:PAYMENT_CLR[pay],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{PAYMENT_LABEL[pay]}</span>
        </div>

        {/* Account-code line items — a PO can split its total across several codes */}
        <div style={{marginTop:14,background:T.bg,borderRadius:10,padding:"4px 12px"}}>
          {items.map((it,i)=>{
            const acc = ACCOUNTS.find(a=>a.code===it.code);
            return (
              <div key={it.id||i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,padding:"9px 0",borderBottom:i<items.length-1?`1px solid ${T.cardBorder}`:"none"}}>
                <div style={{minWidth:0}}>
                  <div style={{fontSize:11,color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{it.code||"—"}</div>
                  <div style={{fontSize:12,color:T.textSecondary,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{acc?.name||"—"}</div>
                  {suppliers.length>1 && <div style={{fontSize:10,color:T.textMuted,marginTop:1}}>🏢 {itemSupplierName(po,it)}</div>}
                </div>
                <div style={{fontSize:13,fontFamily:"'JetBrains Mono',monospace",fontWeight:600,color:T.textPrimary,flexShrink:0}}>{fmt(it.amount)}</div>
              </div>
            );
          })}
          {items.length>1 && (
            <div style={{display:"flex",justifyContent:"space-between",padding:"9px 0",fontSize:12,fontWeight:700}}>
              <span style={{color:T.textMuted}}>รวมทั้ง PO</span>
              <span style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace"}}>{fmt(poTotal(po))}</span>
            </div>
          )}
        </div>

        <div style={{marginTop:4}}>
          <Row label="วันที่สั่ง PO" value={po.date} mono />
          <Row label="แผนจ่ายเงิน" value={po.paymentPlan} mono />
        </div>

        {/* Suppliers — a PO can involve several vendors, each paid/delivered
            across several rounds (installments) instead of one lump sum. */}
        <div style={{marginTop:10}}>
          <div style={{fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase",marginBottom:8}}>🏢 Supplier {suppliers.length>1?`(${suppliers.length} เจ้า)`:""}</div>
          {suppliers.map((s,si)=>{
            const rounds = s.rounds && s.rounds.length ? s.rounds : [];
            const roundsTotal = rounds.reduce((sum,r)=>sum+(parseFloat(r.amount)||0),0);
            return (
              <div key={s.id||si} style={{background:T.bg,borderRadius:10,padding:"10px 12px",marginBottom:8}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:6}}>
                  <div>
                    <span style={{fontSize:13,fontWeight:700,color:T.textPrimary}}>{s.name||"—"}</span>
                    {s.poNumber && <span style={{fontSize:11,color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",marginLeft:8}}>{s.poNumber}</span>}
                  </div>
                  {roundsTotal>0 && <span style={{fontSize:12,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,color:T.amber}}>{fmt(roundsTotal)}</span>}
                </div>
                {rounds.length===0 ? (
                  <div style={{fontSize:12,color:T.textMuted}}>ยังไม่ได้กำหนดแผนของเข้า</div>
                ) : rounds.map((r,i)=>{
                  const st = deliveryStatus(r);
                  return (
                    <div key={r.id||i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,padding:"6px 0",borderBottom:i<rounds.length-1?"1px solid #eef2f7":"none"}}>
                      <span style={{fontSize:11,color:T.textMuted,fontWeight:600,minWidth:18}}>{rounds.length>1?`#${i+1}`:"—"}</span>
                      <span style={{flex:1,fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:T.textPrimary}}>{r.plan||"—"} → {r.actual||"รอ"}</span>
                      {r.amount && <span style={{fontSize:12,fontFamily:"'JetBrains Mono',monospace",color:T.textSecondary}}>{fmt(r.amount)}</span>}
                      <span style={{background:INCOMING_BG[st],color:INCOMING_CLR[st],fontSize:10,padding:"2px 8px",borderRadius:20,fontWeight:600,whiteSpace:"nowrap"}}>{INCOMING_LABEL[st]}</span>
                    </div>
                  );
                })}
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

        <div style={{display:"flex",gap:10,marginTop:20}}>
          <button onClick={()=>onEdit(po)} className="btn-primary" style={{background:T.amber,color:"#fff"}}>✏️ แก้ไข</button>
          <button onClick={()=>{ if(window.confirm("ลบรายการ PO นี้?")) onDelete(po.id); }} className="btn-ghost" style={{color:T.red,borderColor:T.red}}>🗑 ลบ</button>
          <div style={{flex:1}}/>
          <button onClick={onClose} className="btn-ghost">ปิด</button>
        </div>
      </div>
    </div>
  );
}

// ─── Procurement View ─────────────────────────────────────────────────────────
function ProcurementView({ project, tenderCosts, additions, poEntries, savePO, onBack, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[] }) {
  const [tab,    setTab]    = useState("list"); // "list" | "tracking"
  const [view,   setView]   = useState("browse"); // "browse" | "add"
  const emptyForm = () => {
    const supplierId = uid();
    return {
      date:new Date().toISOString().slice(0,10), status:"PO Issued",
      items:[{id:uid(),code:"",amount:"",supplierId}],
      suppliers:[{id:supplierId, name:"", poNumber:"", rounds:[{id:uid(),amount:"",plan:"",actual:""}]}],
      paymentPlan:"", notes:"",
    };
  };
  const [form,   setForm]   = useState(emptyForm);
  const [editId, setEditId] = useState(null);
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [detailId, setDetailId] = useState(null);
  const [collapsed, setCollapsed] = useState({});

  const detailPO = poEntries.find(p => p.id === detailId) || null;
  const openDetail  = (p) => setDetailId(p.id);
  const closeDetail = () => setDetailId(null);
  const toggleGroup = (code) => setCollapsed(c => ({...c, [code]: !c[code]}));

  // Budget (QS) = baseline Tender Cost + every monthly addition entered so far,
  // combined per Acc. Code — matches the "รวมทั้งหมด" total on the QS Monthly tab.
  const combinedBudget = {...tenderCosts};
  Object.values(additions || {}).forEach(monthObj => {
    Object.entries(monthObj || {}).forEach(([code, val]) => {
      combinedBudget[code] = (parseFloat(combinedBudget[code]) || 0) + (parseFloat(val) || 0);
    });
  });
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
  const totalPaid   = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+poTotal(p),0);

  // Item (account-code line) row helpers
  const addItemRow    = () => setForm(f=>({...f, items:[...f.items,{id:uid(),code:"",amount:"",supplierId:f.suppliers[0]?.id||""}]}));
  const removeItemRow = (id) => setForm(f=>({...f, items: f.items.length>1 ? f.items.filter(it=>it.id!==id) : f.items}));
  const updateItemRow = (id, key, val) => setForm(f=>({...f, items: f.items.map(it=>it.id===id?{...it,[key]:val}:it)}));

  // Supplier row helpers — a PO can involve several suppliers
  const addSupplier    = () => setForm(f=>({...f, suppliers:[...f.suppliers,{id:uid(), name:"", poNumber:"", rounds:[{id:uid(),amount:"",plan:"",actual:""}]}]}));
  const removeSupplier = (sid) => setForm(f=>{
    if (f.suppliers.length===1) return f;
    const suppliers = f.suppliers.filter(s=>s.id!==sid);
    // Items pointing at the removed supplier fall back to whichever supplier is now first
    const items = f.items.map(it=>it.supplierId===sid ? {...it, supplierId: suppliers[0]?.id||""} : it);
    return {...f, suppliers, items};
  });
  const updateSupplier = (sid, key, val) => setForm(f=>({...f, suppliers: f.suppliers.map(s=>s.id===sid?{...s,[key]:val}:s)}));

  // Payment/delivery round helpers — each supplier can be paid/delivered in
  // several rounds instead of one lump sum.
  const addRound    = (sid) => setForm(f=>({...f, suppliers: f.suppliers.map(s=>s.id===sid?{...s, rounds:[...s.rounds,{id:uid(),amount:"",plan:"",actual:""}]}:s)}));
  const removeRound = (sid, rid) => setForm(f=>({...f, suppliers: f.suppliers.map(s=>s.id===sid?{...s, rounds: s.rounds.length>1 ? s.rounds.filter(r=>r.id!==rid) : s.rounds}:s)}));
  const updateRound = (sid, rid, key, val) => setForm(f=>({...f, suppliers: f.suppliers.map(s=>s.id===sid?{...s, rounds: s.rounds.map(r=>r.id===rid?{...r,[key]:val}:r)}:s)}));

  const formTotal = form.items.reduce((s,it)=>s+(parseFloat(it.amount)||0),0);
  const formRoundsTotal = form.suppliers.reduce((s,sup)=>s+sup.rounds.reduce((ss,r)=>ss+(parseFloat(r.amount)||0),0),0);

  const submit = () => {
    const validSuppliers = form.suppliers
      .filter(s=>s.name.trim())
      .map(s=>({...s, rounds: s.rounds.filter(r=>r.amount||r.plan||r.actual)}));
    if (!validSuppliers.length) { alert("กรุณากรอกชื่อ Supplier อย่างน้อย 1 เจ้า"); return; }
    const validSupplierIds = new Set(validSuppliers.map(s=>s.id));
    const validItems = form.items.filter(it=>it.code&&it.amount).map(it =>
      validSupplierIds.has(it.supplierId) ? it : {...it, supplierId: validSuppliers[0].id}
    );
    if (!validItems.length) return;
    const payload = {...form, items:validItems, suppliers:validSuppliers};
    delete payload.supplier; delete payload.poNumber; delete payload.deliveries; // superseded by suppliers[]
    savePO(editId ? poEntries.map(p=>p.id===editId?{...payload,id:editId}:p) : [...poEntries,{...payload,id:uid()}]);
    setEditId(null);
    setForm(emptyForm());
    setView("browse");
  };

  const openEdit = (p) => {
    const newSuppliers = poSuppliers(p).map(s=>({
      ...s,
      id: s.id&&s.id!=="legacy" ? s.id : uid(),
      rounds: (s.rounds&&s.rounds.length ? s.rounds : [{amount:"",plan:"",actual:""}])
        .map(r=>({...r, id: r.id&&r.id!=="legacy-round" ? r.id : uid()})),
    }));
    // Old supplier id -> new supplier id, so each item stays linked to the
    // same supplier after ids get freshly minted for this edit session.
    const idMap = {};
    poSuppliers(p).forEach((s,i) => { idMap[s.id] = newSuppliers[i].id; });
    setForm({
      ...emptyForm(), date:p.date, status:p.status, paymentPlan:p.paymentPlan||"", notes:p.notes||"",
      items: poItems(p).map(it=>({
        ...it, id: it.id&&it.id!=="legacy" ? it.id : uid(),
        supplierId: idMap[it.supplierId] || newSuppliers[0]?.id || "",
      })),
      suppliers: newSuppliers,
    });
    setEditId(p.id); setView("add"); setDetailId(null);
  };
  const deletePO = (id) => { savePO(poEntries.filter(x=>x.id!==id)); setDetailId(null); };
  const closeForm = () => { setView("browse"); setEditId(null); setForm(emptyForm()); };

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
    <Shell role="procurement" color={T.amber} project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"24px 28px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:20}}>
          <StatCard label="Budget (QS)" value={fmt(tenderTotal)} sub="เดิม + เพิ่มรายเดือนทุกเดือน" color={T.blue} icon="📋" accent={T.blueLight}/>
          <StatCard label="Committed (PO)" value={fmt(totalComm)} sub={`${poEntries.length} รายการ`} color={T.amber} icon="📦" accent={T.amberBg}/>
          <StatCard label="ชำระแล้ว" value={fmt(totalPaid)} sub={`${poEntries.filter(p=>p.status==="Paid").length} รายการ`} color={T.green} icon="✅" accent={T.greenBg}/>
          <StatCard label="Budget คงเหลือ" value={fmt(tenderTotal-totalComm)} sub={tenderTotal>0?`${((totalComm/tenderTotal)*100).toFixed(1)}% ใช้ไปแล้ว`:"—"} color={tenderTotal-totalComm<0?T.red:T.textSecondary} icon={tenderTotal-totalComm<0?"⚠️":"💰"} accent={tenderTotal-totalComm<0?T.redBg:"#f8fafc"}/>
        </div>

        {view!=="add" && (
          <div style={{display:"flex",gap:8,marginBottom:16}}>
            {[["list","📋 รายการ PO"],["tracking","🚚 ติดตามของเข้า/จ่ายเงิน"]].map(([id,label])=>(
              <button key={id} onClick={()=>setTab(id)}
                style={{background:tab===id?T.amber:T.card,color:tab===id?"#fff":T.textSecondary,border:`1px solid ${tab===id?T.amber:T.cardBorder}`,borderRadius:10,padding:"9px 18px",fontSize:13,fontWeight:600,cursor:"pointer"}}>
                {label}
              </button>
            ))}
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
                <input type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))} className="input-base"/>
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>สถานะ</span>
                <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} className="input-base">
                  {PO_STATUS.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
              </label>

              {/* Suppliers — one PO can involve several vendors, and each
                  vendor can be paid/delivered across several rounds instead
                  of one lump sum. */}
              <div style={{gridColumn:"1/-1",display:"flex",alignItems:"center",gap:8,marginTop:6,paddingTop:14,borderTop:`1px dashed ${T.cardBorder}`}}>
                <span style={{fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase"}}>🏢 Supplier * (แบ่งจ่ายได้หลายเจ้า หลายรอบ ระบุก่อน แล้วค่อยผูกกับหมวดต้นทุนด้านล่าง)</span>
              </div>
              <div style={{gridColumn:"1/-1",display:"flex",flexDirection:"column",gap:12}}>
                {form.suppliers.map((s,si)=>(
                  <div key={s.id} style={{border:`1px solid ${T.cardBorder}`,borderRadius:10,padding:12,background:T.bg}}>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr auto",gap:8,alignItems:"center",marginBottom:10}}>
                      <input placeholder="ชื่อ Supplier *" value={s.name} onChange={e=>updateSupplier(s.id,"name",e.target.value)} className="input-base"/>
                      <input placeholder="เลข PO" value={s.poNumber} onChange={e=>updateSupplier(s.id,"poNumber",e.target.value)} className="input-base"/>
                      <button type="button" onClick={()=>removeSupplier(s.id)} disabled={form.suppliers.length===1}
                        style={{background:"none",border:"none",color:form.suppliers.length===1?T.textMuted:T.red,cursor:form.suppliers.length===1?"default":"pointer",padding:"4px 8px",fontSize:15,opacity:form.suppliers.length===1?0.4:1}}>🗑</button>
                    </div>
                    <div style={{fontSize:11,fontWeight:600,color:T.textMuted,marginBottom:6}}>รอบจ่ายเงิน / ของเข้า {s.rounds.length>1?`(${s.rounds.length} รอบ)`:""}</div>
                    <div style={{display:"flex",flexDirection:"column",gap:6}}>
                      {s.rounds.map((r,ri)=>(
                        <div key={r.id} style={{display:"grid",gridTemplateColumns:"18px 110px 1fr 1fr auto",gap:8,alignItems:"center"}}>
                          <span style={{fontSize:11,color:T.textMuted,fontWeight:600}}>{s.rounds.length>1?`#${ri+1}`:""}</span>
                          <input type="number" placeholder="มูลค่า" value={r.amount} onChange={e=>updateRound(s.id,r.id,"amount",e.target.value)} className="input-base"/>
                          <label style={{display:"flex",flexDirection:"column",gap:2}}>
                            {ri===0 && <span style={{fontSize:10,color:T.textSecondary}}>แผนของเข้า</span>}
                            <input type="date" value={r.plan} onChange={e=>updateRound(s.id,r.id,"plan",e.target.value)} className="input-base"/>
                          </label>
                          <label style={{display:"flex",flexDirection:"column",gap:2}}>
                            {ri===0 && <span style={{fontSize:10,color:T.textSecondary}}>รับจริง</span>}
                            <input type="date" value={r.actual} onChange={e=>updateRound(s.id,r.id,"actual",e.target.value)} className="input-base"/>
                          </label>
                          <button type="button" onClick={()=>removeRound(s.id,r.id)} disabled={s.rounds.length===1}
                            style={{background:"none",border:"none",color:s.rounds.length===1?T.textMuted:T.red,cursor:s.rounds.length===1?"default":"pointer",padding:"4px 8px",fontSize:15,opacity:s.rounds.length===1?0.4:1,alignSelf:ri===0?"end":"center",marginBottom:ri===0?1:0}}>🗑</button>
                        </div>
                      ))}
                    </div>
                    <button type="button" onClick={()=>addRound(s.id)} className="btn-ghost" style={{padding:"5px 10px",fontSize:11,marginTop:8}}>+ เพิ่มรอบจ่ายเงิน</button>
                  </div>
                ))}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <button type="button" onClick={addSupplier} className="btn-ghost" style={{padding:"6px 12px",fontSize:12}}>+ เพิ่ม Supplier</button>
                  {formRoundsTotal>0 && <span style={{fontSize:12,color:T.textSecondary}}>รวมทุกรอบ: <b style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace"}}>{fmt(formRoundsTotal)}</b></span>}
                </div>
              </div>

              {/* Account-code line items — one PO can be split across several
                  codes, each with its own amount and its own linked supplier;
                  the amounts sum to the PO total. */}
              <div style={{gridColumn:"1/-1",display:"flex",alignItems:"center",gap:8,marginTop:6,paddingTop:14,borderTop:`1px dashed ${T.cardBorder}`}}>
                <span style={{fontSize:11,fontWeight:700,color:T.textMuted,letterSpacing:0.6,textTransform:"uppercase"}}>📐 หมวดต้นทุน * (แยกได้หลายรหัส ผูกกับ Supplier ด้านบน)</span>
              </div>
              <div style={{gridColumn:"1/-1",display:"flex",flexDirection:"column",gap:8}}>
                {form.items.map((it,i)=>(
                  <div key={it.id} style={{display:"grid",gridTemplateColumns:"1fr 1fr 140px auto",gap:8,alignItems:"center"}}>
                    <select value={it.code} onChange={e=>updateItemRow(it.id,"code",e.target.value)} className="input-base">
                      <option value="">— เลือก Account Code —</option>
                      {ACCOUNTS.map(a=><option key={a.code} value={a.code}>{a.code} · {a.name}</option>)}
                    </select>
                    <select value={it.supplierId} onChange={e=>updateItemRow(it.id,"supplierId",e.target.value)} className="input-base">
                      {form.suppliers.map(s=><option key={s.id} value={s.id}>{s.name.trim() ? `🏢 ${s.name}` : "🏢 (ยังไม่ตั้งชื่อ)"}</option>)}
                    </select>
                    <input type="number" placeholder="มูลค่า (THB)" value={it.amount} onChange={e=>updateItemRow(it.id,"amount",e.target.value)} className="input-base"/>
                    <button type="button" onClick={()=>removeItemRow(it.id)} disabled={form.items.length===1}
                      style={{background:"none",border:"none",color:form.items.length===1?T.textMuted:T.red,cursor:form.items.length===1?"default":"pointer",padding:"4px 8px",fontSize:15,opacity:form.items.length===1?0.4:1}}>🗑</button>
                  </div>
                ))}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <button type="button" onClick={addItemRow} className="btn-ghost" style={{padding:"6px 12px",fontSize:12}}>+ เพิ่ม Account Code</button>
                  {form.items.length>1 && <span style={{fontSize:12,color:T.textSecondary}}>รวม: <b style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace"}}>{fmt(formTotal)}</b></span>}
                </div>
              </div>

              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>แผนจ่ายเงิน (Payment Plan)</span>
                <input type="date" value={form.paymentPlan} onChange={e=>setForm(f=>({...f,paymentPlan:e.target.value}))} className="input-base"/>
              </label>

              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>หมายเหตุ</span>
                <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} className="input-base" style={{resize:"vertical"}}/>
              </label>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button onClick={submit} className="btn-primary" style={{background:T.amber,color:"#fff"}}>{editId?"อัปเดต":"เพิ่ม PO"}</button>
              <button onClick={closeForm} className="btn-ghost">ยกเลิก</button>
            </div>
          </div>
        ) : tab==="tracking" ? (
          <ProcurementTrackingTab poEntries={poEntries} onEdit={openEdit} onView={openDetail} onAddNew={()=>setView("add")} />
        ) : (
          <>
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา Account, supplier, PO..."
                className="input-base" style={{width:240}}/>
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
                        <span style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:13}}>{fmt(groupTotal)}</span>
                      </div>
                      {!isCollapsed && (
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                          <thead>
                            <tr>
                              {["วันที่","Supplier","PO No.","มูลค่า (THB)","สถานะ",""].map(h=>(
                                <th key={h} style={{padding:"9px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:10,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {rows.map(({po:p,item},i)=>{
                              const splitAcrossCodes = poItems(p).length>1;
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
                                  {splitAcrossCodes && <div style={{fontSize:10,color:T.textMuted}}>จาก {poItems(p).length} รหัส · รวม {fmt(poTotal(p))}</div>}
                                </td>
                                <td style={{padding:"10px 16px"}}>
                                  <span style={{background:STATUS_BG[p.status],color:STATUS_CLR[p.status],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{p.status}</span>
                                </td>
                                <td style={{padding:"10px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                                  <button onClick={()=>openEdit(p)} style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",padding:"2px 6px",borderRadius:6,marginRight:4}}>✏️</button>
                                  <button onClick={()=>deletePO(p.id)} style={{background:"none",border:"none",color:T.red,cursor:"pointer",padding:"2px 6px",borderRadius:6}}>🗑</button>
                                </td>
                              </tr>
                            );})}
                          </tbody>
                        </table>
                      )}
                    </div>
                  );
                })}
                <div style={{display:"flex",justifyContent:"flex-end",gap:16,padding:"4px 18px",color:T.textMuted,fontSize:12}}>
                  <span>{filtered.length} รายการทั้งหมด</span>
                  <span style={{color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{fmt(filtered.reduce((s,p)=>s+poTotal(p),0))}</span>
                </div>
              </div>
            )}
          </>
        )}
      </div>
      <PODetailModal po={detailPO} onClose={closeDetail} onEdit={openEdit} onDelete={deletePO} />
    </Shell>
  );
}

// ─── Procurement: Incoming / Payment Tracking tab ─────────────────────────────
// Groups every PO by its Account Code so the team can see, at a glance and per
// cost line, which deliveries and payments are on track vs. overdue.
function ProcurementTrackingTab({ poEntries, onEdit, onView, onAddNew }) {
  const [search, setSearch] = useState("");
  const [onlyIssues, setOnlyIssues] = useState(false);

  const counts = poEntries.reduce((acc,p) => {
    const inc = incomingStatus(p), pay = paymentStatus(p);
    if (inc==="pending") acc.incPending++;
    if (inc==="late")    acc.incLate++;
    if (pay==="pending") acc.payPending++;
    if (pay==="late")    acc.payLate++;
    return acc;
  }, { incPending:0, incLate:0, payPending:0, payLate:0 });

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
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:20}}>
        <StatCard label="รอของเข้า" value={counts.incPending} sub="ยังไม่ถึงวันที่นัด" color={T.amber} icon="⏳" accent={T.amberBg}/>
        <StatCard label="ของเข้าล่าช้า" value={counts.incLate} sub="เลยวันแผนของเข้าแล้ว" color={T.red} icon="⚠️" accent={T.redBg}/>
        <StatCard label="รอจ่ายเงิน" value={counts.payPending} sub="ยังไม่ถึงวันครบกำหนด" color={T.amber} icon="⏳" accent={T.amberBg}/>
        <StatCard label="จ่ายเงินเกินกำหนด" value={counts.payLate} sub="เลยวันแผนจ่ายเงินแล้ว" color={T.red} icon="🔴" accent={T.redBg}/>
      </div>

      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา Acc. Code, supplier, PO..."
          className="input-base" style={{width:240}}/>
        <button onClick={()=>setOnlyIssues(v=>!v)}
          style={{background:onlyIssues?T.red:"transparent",border:`1.5px solid ${onlyIssues?T.red:T.cardBorder}`,borderRadius:8,padding:"7px 14px",color:onlyIssues?"#fff":T.textSecondary,fontSize:12,cursor:"pointer",fontWeight:600}}>
          ⚠️ แสดงเฉพาะรายการล่าช้า
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
            return (
              <div key={code} style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                <div style={{padding:"12px 18px",background:"#f8fafc",borderBottom:`1px solid ${T.cardBorder}`,display:"flex",alignItems:"center",gap:10}}>
                  <span style={{color:T.blue,fontSize:12,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{code}</span>
                  <span style={{color:T.textPrimary,fontSize:13,fontWeight:600}}>{acc?.name || "—"}</span>
                  <span style={{flex:1}}/>
                  <span style={{color:T.textMuted,fontSize:11}}>{rows.length} PO</span>
                  {lateCount>0 && <Badge text={`⚠️ ${lateCount} ล่าช้า`} clr={T.red} bg={T.redBg}/>}
                </div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr>
                      {["PO No.","Supplier","มูลค่า (THB)","การส่งของ","แผนจ่ายเงิน","สถานะ",""].map(h=>(
                        <th key={h} style={{padding:"9px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:10,letterSpacing:0.6,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(({po:p,item},i) => {
                      const inc = incomingStatus(p), pay = paymentStatus(p);
                      const splitAcrossCodes = poItems(p).length>1;
                      return (
                        <tr key={p.id+"-"+(item.id||item.code)} onClick={()=>onView?.(p)}
                          style={{background:i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`,cursor:onView?"pointer":"default"}}
                          onMouseEnter={e=>e.currentTarget.style.background="#fef9ec"}
                          onMouseLeave={e=>e.currentTarget.style.background=i%2===0?T.card:"#fafbfd"}>
                          <td style={{padding:"9px 16px",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{poNumbersLabel(p)}</td>
                          <td style={{padding:"9px 16px",color:T.textPrimary,fontWeight:500}}>{itemSupplierName(p,item)}</td>
                          <td style={{padding:"9px 16px",textAlign:"right"}}>
                            <div style={{color:T.textPrimary,fontFamily:"'JetBrains Mono',monospace",fontWeight:600}}>{fmt(item.amount)}</div>
                            {splitAcrossCodes && <div style={{fontSize:10,color:T.textMuted}}>รวม {fmt(poTotal(p))}</div>}
                          </td>
                          <td style={{padding:"9px 16px"}}><DeliveryList po={p}/></td>
                          <td style={{padding:"9px 16px"}}><DateCell value={p.paymentPlan} lateTint={pay==="late"}/></td>
                          <td style={{padding:"9px 16px"}}>
                            <div style={{display:"flex",flexDirection:"column",gap:3,alignItems:"flex-start"}}>
                              <Badge text={INCOMING_LABEL[inc]} clr={INCOMING_CLR[inc]} bg={INCOMING_BG[inc]}/>
                              <Badge text={PAYMENT_LABEL[pay]} clr={PAYMENT_CLR[pay]} bg={PAYMENT_BG[pay]}/>
                            </div>
                          </td>
                          <td style={{padding:"9px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                            <button onClick={()=>onEdit(p)} style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",padding:"2px 6px",borderRadius:6}}>✏️</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ─── Accounting View ──────────────────────────────────────────────────────────
function AccountingView({ project, tenderCosts, additions, poEntries, onBack, onExport, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[] }) {
  const [view, setView] = useState("dashboard");
  const [sortKey, setSortKey] = useState(null);  // "code" | "name" | "group" | "budget" | "committed" | "pct" | null
  const [sortDir, setSortDir] = useState(1);

  // Budget = baseline Tender Cost + every monthly addition entered so far,
  // combined per Acc. Code — matches the "รวมทั้งหมด" total on the QS Monthly tab.
  const combinedBudget = {...tenderCosts};
  Object.values(additions || {}).forEach(monthObj => {
    Object.entries(monthObj || {}).forEach(([code, val]) => {
      combinedBudget[code] = (parseFloat(combinedBudget[code]) || 0) + (parseFloat(val) || 0);
    });
  });

  // Sum only top-level codes — see note in ProcurementView. Object.values()
  // over the whole combinedBudget double-counts sub-items (EX-xxxx rows),
  // since their value is already folded into their parent account's total.
  const topLevelCodes = [
    ...ACCOUNTS.filter(a => !hiddenAccounts.includes(a.code)).map(a => a.code),
    ...extraItems.filter(e => !e.parentCode).map(e => e.code),
  ];
  const tenderTotal   = topLevelCodes.reduce((s,c) => s + (parseFloat(combinedBudget[c]) || 0), 0);
  const totalComm     = poEntries.reduce((s,p)=>s+poTotal(p),0);
  const totalPaid     = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+poTotal(p),0);
  const totalInvoiced = poEntries.filter(p=>["Invoiced","Paid"].includes(p.status)).reduce((s,p)=>s+poTotal(p),0);
  const pct           = tenderTotal>0?(totalComm/tenderTotal*100):0;

  const groupData = GROUPS.map((g,i)=>{
    const codes=ACCOUNTS.filter(a=>a.group===g).map(a=>a.code);
    const committed = poEntries.reduce((s,p)=>s+poItems(p).filter(it=>codes.includes(it.code)).reduce((s2,it)=>s2+(parseFloat(it.amount)||0),0),0);
    return {group:g,budget:codes.reduce((s,c)=>s+(parseFloat(combinedBudget[c])||0),0),committed,color:GRP_COLORS[i%GRP_COLORS.length]};
  }).filter(g=>g.budget>0||g.committed>0);

  const accountData = ACCOUNTS.map(a=>{
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
      else                              { av = pctUsedOf(a); bv = pctUsedOf(b); }
      if (typeof av === "string") return av.localeCompare(bv) * sortDir;
      return (av - bv) * sortDir;
    });
    return arr;
  })();

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

  return (
    <Shell role="accounting" color={T.green} project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"24px 28px"}}>
        {/* Tabs + Export */}
        <div style={{display:"flex",gap:8,marginBottom:24,alignItems:"center"}}>
          {[["dashboard","📊 Dashboard"],["detail","📋 รายละเอียด"]].map(([v,l])=>(
            <button key={v} onClick={()=>setView(v)}
              style={{background:view===v?T.green:"transparent",border:`1.5px solid ${view===v?T.green:T.cardBorder}`,borderRadius:10,padding:"8px 20px",color:view===v?"#fff":T.textSecondary,fontSize:13,cursor:"pointer",fontWeight:view===v?600:500,transition:"all 0.15s"}}>{l}</button>
          ))}
          <button onClick={onExport} className="btn-ghost" style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:6,borderColor:T.green,color:T.green}}>
            ⬇️ Export Excel
          </button>
        </div>

        {view==="dashboard" ? (
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:24}}>
              <StatCard label="งบประมาณ (QS)" value={fmt(tenderTotal)} sub="เดิม + เพิ่มรายเดือนทุกเดือน" color={T.blue} icon="📋" accent={T.blueLight}/>
              <StatCard label="Committed (PO)" value={fmt(totalComm)} sub={`${pct.toFixed(1)}% ของงบ`} color={T.amber} icon="📦" accent={T.amberBg}/>
              <StatCard label="Invoiced" value={fmt(totalInvoiced)} sub="รอจ่าย + จ่ายแล้ว" color={T.purple} icon="🧾" accent={T.purpleBg}/>
              <StatCard label="ชำระแล้ว (Paid)" value={fmt(totalPaid)} sub={`${poEntries.filter(p=>p.status==="Paid").length} รายการ`} color={T.green} icon="✅" accent={T.greenBg}/>
            </div>

            {/* Progress */}
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:22,marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                <span style={{fontSize:13,color:T.textPrimary,fontWeight:600}}>Budget Utilization</span>
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
          </>
        ) : (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#f8fafc"}}>
                  {[
                    {label:"Acc. Code", key:"code"},
                    {label:"Account Name", key:"name"},
                    {label:"Group", key:"group"},
                    {label:"Budget (QS)", key:"budget"},
                    {label:"Committed (PO)", key:"committed"},
                    {label:"% Used", key:"pct"},
                    {label:"", key:null},
                  ].map(({label,key})=>(
                    <th key={label||"__actions"}
                      style={{padding:"11px 16px",textAlign:["Budget (QS)","Committed (PO)","% Used"].includes(label)?"right":"left",color:sortKey===key?T.green:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`,whiteSpace:"nowrap"}}>
                      <span onClick={()=>key&&handleSort(key)} style={{cursor:key?"pointer":"default",userSelect:"none"}}>{label}{key && sortKey===key ? (sortDir===1?" ▲":" ▼") : ""}</span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {displayAccountData.map((a,i)=>{
                  const p2=a.budget>0?(a.committed/a.budget*100):a.committed>0?999:0;
                  return (
                    <tr key={a.code} style={{background:a.over?"#fff5f5":i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`}}>
                      <td style={{padding:"10px 16px",color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontSize:12,fontWeight:500}}>{a.code}</td>
                      <td style={{padding:"10px 16px",color:T.textPrimary}}>{a.name}</td>
                      <td style={{padding:"10px 16px"}}>
                        <span style={{background:T.blueLight,color:T.blue,fontSize:10,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{a.group}</span>
                      </td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:T.blue,fontWeight:500}}>{a.budget>0?fmt(a.budget):"—"}</td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:a.over?T.red:T.amber,fontWeight:a.over?700:500}}>{a.committed>0?fmt(a.committed):"—"}</td>
                      <td style={{padding:"10px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:p2>100?T.red:p2>80?T.amber:T.green,fontSize:12,fontWeight:600}}>
                        {a.budget>0?`${p2.toFixed(1)}%`:a.committed>0?"No Budget":"—"}
                      </td>
                      <td style={{padding:"10px 16px"}}>
                        {a.over
                          ? <span style={{background:T.redBg,color:T.red,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>⚠ เกินงบ</span>
                          : a.committed>0 ? <span style={{background:T.greenBg,color:T.green,fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>OK</span>
                          : a.budget>0 ? <span style={{background:"#f8fafc",color:T.textMuted,fontSize:11,padding:"3px 10px",borderRadius:20}}>ยังไม่ PO</span> : null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
                  <td colSpan={3} style={{padding:"12px 16px",color:T.textMuted,fontSize:12}}>{accountData.length} รายการ</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:T.blue,fontWeight:700,fontSize:14}}>{fmt(accountData.reduce((s,a)=>s+a.budget,0))}</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"'JetBrains Mono',monospace",color:T.amber,fontWeight:700,fontSize:14}}>{fmt(accountData.reduce((s,a)=>s+a.committed,0))}</td>
                  <td colSpan={2}/>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </div>
    </Shell>
  );
}
