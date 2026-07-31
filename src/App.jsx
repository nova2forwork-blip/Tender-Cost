import { useState, useEffect, useCallback, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import { supabase, sg, ss, sd } from "./supabase.js";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend, AreaChart, Area, CartesianGrid } from "recharts";
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
function exportToExcel(project, tenderCosts, poEntries) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [];
  summaryRows.push([`Project: ${project.name}`, "", "", "", "", ""]);
  summaryRows.push([`Area: ${project.area} ft²`, "", `Panels: ${project.panels}`, "", "", ""]);
  summaryRows.push([`Export Date: ${new Date().toLocaleDateString("th-TH")}`, "", "", "", "", ""]);
  summaryRows.push([]);
  summaryRows.push(["Acc. Code","Account Name","Group","Budget / Tender Cost","Committed (PO)","Remaining","% Used","Status"]);
  let grandBudget=0, grandCommitted=0;
  ACCOUNTS.forEach(a => {
    const budget    = parseFloat(tenderCosts[a.code]) || 0;
    const committed = poEntries.filter(p=>p.code===a.code).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
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
  poRows.push([`Project: ${project.name}`, "", "", "", "", "", "", ""]);
  poRows.push([]);
  poRows.push(["PO Date","Acc. Code","Account Name","Group","Supplier","PO Number","Amount","Status","Notes"]);
  poEntries.forEach(p => {
    const acc = ACCOUNTS.find(a=>a.code===p.code);
    poRows.push([p.date, p.code, acc?.name||"", acc?.group||"", p.supplier, p.poNumber, parseFloat(p.amount)||0, p.status, p.notes||""]);
  });
  poRows.push([]);
  poRows.push(["","","","","","TOTAL", poEntries.reduce((s,p)=>s+(parseFloat(p.amount)||0),0), "",""]);
  const ws2 = XLSX.utils.aoa_to_sheet(poRows);
  ws2["!cols"] = [{wch:12},{wch:10},{wch:38},{wch:14},{wch:24},{wch:16},{wch:18},{wch:12},{wch:30}];
  for (let r=2; r<poRows.length; r++) {
    const ref = XLSX.utils.encode_cell({r, c:6});
    if (ws2[ref] && typeof ws2[ref].v === "number") ws2[ref].z = '#,##0.00';
  }
  XLSX.utils.book_append_sheet(wb, ws2, "PO Entries");

  const grpRows = [["Group","Budget","Committed","Remaining","% Used"]];
  GROUPS.forEach(g => {
    const codes = ACCOUNTS.filter(a=>a.group===g).map(a=>a.code);
    const b = codes.reduce((s,c)=>s+(parseFloat(tenderCosts[c])||0),0);
    const c2 = poEntries.filter(p=>codes.includes(p.code)).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
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
        <AccountingView {...sharedProps} onExport={() => exportToExcel(activeProject, tenderCosts, poEntries)} />
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
  const handleAddExtraItem = ({ name, group, parentCode, code }) => {
    if (!name.trim()) return;
    const item = parentCode
      ? { code:`EX-${uid()}`, name:name.trim(), parentCode }
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
  const subItemsByParent = {};
  extraItems.forEach(e => {
    if (e.parentCode) (subItemsByParent[e.parentCode] = subItemsByParent[e.parentCode] || []).push(e);
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
              {["Acc. Code","Group","Account Name","Tender Cost (THB)",""].map(h=>(
                <th key={h} style={{padding:"11px 16px",textAlign:h.includes("Cost")?"right":"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.map((a,i)=>{
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
  const [draftAdd, setDraftAdd] = useState({...(additions[month]||{})});
  const [saved, setSaved] = useState(false);
  const [addExtraOpen, setAddExtraOpen] = useState(false);
  const [extraDraft, setExtraDraft] = useState({ name:"", group:GROUPS[0] });

  useEffect(() => { setDraftAdd({...(additions[month]||{})}); }, [month, additions]);
  useEffect(() => { if (!months.includes(month) && months.length) setMonth(months[months.length-1]); }, [months]); // eslint-disable-line

  // All rows = original 70 account codes + standalone extra items.
  // Sub-items that roll up into an existing Acc. Code (parentCode set) are
  // baseline-only breakdown lines — they're edited on the Baseline tab and
  // their sum already flows into tenderCosts[parentCode], so they don't need
  // a separate row here.
  const allRows = [...ACCOUNTS.filter(a=>!hiddenAccounts.includes(a.code)), ...extraItems.filter(e=>!e.parentCode).map(e => ({ code:e.code, name:e.name, group:e.group, isExtra:true }))];

  const monthTotal = (m) => allRows.reduce((s,r) => s + (parseFloat(additions[m]?.[r.code]) || 0), 0);

  const baseTotal = Object.values(tenderCosts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const thisMonthAdd = allRows.reduce((s,r) => s + (parseFloat(draftAdd[r.code]) || 0), 0);
  const cumulativeSoFar = months.filter(m=>m<month).reduce((s,m)=>s+monthTotal(m),0) + thisMonthAdd + baseTotal;

  // "Live" versions that use the currently-edited draft for the selected month
  // (instead of the last-saved value) so the top summary updates as you type.
  const monthTotalLive = (m) => m===month ? thisMonthAdd : monthTotal(m);
  const sortedMonths = months.length ? months : [thisMonth];
  const cumulativeLive = (uptoMonth) => baseTotal + sortedMonths.filter(m=>m<=uptoMonth).reduce((s,m)=>s+monthTotalLive(m),0);
  const grandTotal = cumulativeLive(sortedMonths[sortedMonths.length-1]);
  const monthShortLabel = (m) => new Date(m+"-01").toLocaleDateString("th-TH",{month:"short",year:"2-digit"});
  const chartData = [
    { label:"เริ่มต้น", cumulative: baseTotal },
    ...sortedMonths.map(m => ({ label: monthShortLabel(m), cumulative: cumulativeLive(m), added: monthTotalLive(m) })),
  ];

  const filtered = allRows.filter(r =>
    (filter==="All" || r.group===filter) &&
    (r.name.toLowerCase().includes(search.toLowerCase()) || r.code.includes(search))
  );

  const handleAddMonth = () => {
    if (!newMonth || months.includes(newMonth)) return;
    saveAdditions({ ...additions, [newMonth]: additions[newMonth] || {} });
    setMonth(newMonth); setNewMonth("");
  };

  const handleSave = () => {
    const clean = {};
    Object.entries(draftAdd).forEach(([k,v]) => { if(v!==""&&!isNaN(v)&&parseFloat(v)!==0) clean[k]=parseFloat(v); });
    saveAdditions({ ...additions, [month]: clean });
    setSaved(true); setTimeout(()=>setSaved(false),2000);
  };

  const handleCreateExtra = () => {
    if (!extraDraft.name.trim()) return;
    onAddExtra(extraDraft);
    setExtraDraft({ name:"", group:GROUPS[0] }); setAddExtraOpen(false);
  };

  const handleDeleteExtra = (code) => onDeleteExtra(code);

  return (
    <div style={{padding:"4px 28px 24px"}}>
      {/* Trend chart — the whole project's cost growth over time, at a glance */}
      <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"18px 20px 8px",marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:6,flexWrap:"wrap",gap:6}}>
          <span style={{fontSize:13,fontWeight:700,color:T.textPrimary}}>📈 แนวโน้มต้นทุนสะสม</span>
          <span style={{fontSize:12,color:T.textMuted}}>รวมล่าสุดทั้งโปรเจกต์: <b style={{color:T.green,fontFamily:"'JetBrains Mono',monospace",fontSize:15}}>{fmt(grandTotal)}</b> THB</span>
        </div>
        <ResponsiveContainer width="100%" height={170}>
          <AreaChart data={chartData} margin={{top:8,right:8,left:-18,bottom:0}}>
            <defs>
              <linearGradient id="cumGrad" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={T.blue} stopOpacity={0.3}/>
                <stop offset="100%" stopColor={T.blue} stopOpacity={0.02}/>
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eef2f7"/>
            <XAxis dataKey="label" tick={{fontSize:11,fill:T.textMuted}} axisLine={false} tickLine={false}/>
            <YAxis tick={{fontSize:10,fill:T.textMuted}} axisLine={false} tickLine={false} tickFormatter={fmtK}/>
            <Tooltip formatter={(v)=>[`${fmt(v)} THB`,"รวมสะสม"]} labelStyle={{color:T.textPrimary,fontWeight:600,marginBottom:2}}
              contentStyle={{borderRadius:10,border:`1px solid ${T.cardBorder}`,fontSize:12,boxShadow:"0 4px 14px rgba(0,0,0,0.08)"}}/>
            <Area type="monotone" dataKey="cumulative" stroke={T.blue} strokeWidth={2.5} fill="url(#cumGrad)"/>
          </AreaChart>
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
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:16,marginBottom:20}}>
        <StatCard label="ราคาเดิม (Baseline)" value={fmt(baseTotal)} sub="รวมทุก Account Code" color={T.blue} icon="📐" accent={T.blueLight}/>
        <StatCard label="เพิ่มเดือนนี้" value={fmt(thisMonthAdd)} sub={new Date(month+"-01").toLocaleDateString("th-TH",{year:"numeric",month:"long"})} color={T.amber} icon="➕" accent={T.amberBg}/>
        <StatCard label="รวมสะสมถึงเดือนนี้" value={fmt(cumulativeSoFar)} sub="เดิม + เพิ่มสะสมทุกเดือน" color={T.green} icon="✅" accent={T.greenBg}/>
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
        <div style={{background:"#fafbfd",border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:16,marginBottom:16,display:"grid",gridTemplateColumns:"2fr 1fr auto",gap:10,alignItems:"end"}}>
          <label style={{display:"flex",flexDirection:"column",gap:5}}>
            <span style={{fontSize:11,color:T.textSecondary}}>ชื่อรายการงานเพิ่ม</span>
            <input className="input-base" value={extraDraft.name} onChange={e=>setExtraDraft(d=>({...d,name:e.target.value}))} placeholder="เช่น งานเพิ่มกระจกโค้งพิเศษ" />
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
              <th style={{padding:"11px 16px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>Acc. Code</th>
              <th style={{padding:"11px 16px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>Group</th>
              <th style={{padding:"11px 16px",textAlign:"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>Account Name</th>
              <th style={{padding:"11px 16px",textAlign:"right",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>📐 เดิม</th>
              <th style={{padding:"11px 16px",textAlign:"center",color:T.textMuted,fontWeight:600,fontSize:11,borderBottom:`1px solid ${T.cardBorder}`,width:20}}>+</th>
              <th style={{padding:"11px 16px",textAlign:"right",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>➕ เพิ่มเดือนนี้</th>
              <th style={{padding:"11px 16px",textAlign:"center",color:T.textMuted,fontWeight:600,fontSize:11,borderBottom:`1px solid ${T.cardBorder}`,width:20}}>=</th>
              <th style={{padding:"11px 16px",textAlign:"right",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>✅ รวมสะสม</th>
              <th style={{padding:"11px 16px",borderBottom:`1px solid ${T.cardBorder}`,width:20}}></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((r,i) => {
              const baseVal = parseFloat(tenderCosts[r.code]) || 0;
              const cumBefore = months.filter(m=>m<month).reduce((s,m)=>s+(parseFloat(additions[m]?.[r.code])||0),0) + baseVal;
              const thisVal = parseFloat(draftAdd[r.code]) || 0;
              const cum = cumBefore + thisVal;
              return (
                <tr key={r.code} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:"1px solid #f1f5f9"}}>
                  <td style={{padding:"10px 16px",color:T.blue,fontFamily:"'JetBrains Mono',monospace",fontSize:12,fontWeight:500}}>{r.code}</td>
                  <td style={{padding:"10px 16px"}}>
                    <span style={{background:T.blueLight,color:T.blue,fontSize:10,padding:"2px 9px",borderRadius:6,fontWeight:600}}>{r.group}</span>
                  </td>
                  <td style={{padding:"10px 16px",color:T.textPrimary}}>
                    {r.name}
                    {r.isExtra && <span style={{marginLeft:7,fontSize:10,background:T.amberBg,color:T.amber,padding:"1px 8px",borderRadius:6,fontWeight:600}}>งานเพิ่ม</span>}
                  </td>
                  <td style={{padding:"8px 16px",textAlign:"right",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace"}}>{fmt(baseVal)}</td>
                  <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>+</td>
                  <td style={{padding:"8px 16px",textAlign:"right"}}>
                    <input type="number" value={draftAdd[r.code]??""} onChange={e=>setDraftAdd(d=>({...d,[r.code]:e.target.value}))}
                      placeholder="0.00" className="input-base" style={{width:130,textAlign:"right",fontFamily:"'JetBrains Mono',monospace",background:thisVal!==0?T.amberBg:T.bg}}/>
                  </td>
                  <td style={{textAlign:"center",color:T.cardBorder,fontSize:13}}>=</td>
                  <td style={{padding:"8px 16px",textAlign:"right",color:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:700}}>{fmt(cum)}</td>
                  <td style={{padding:"8px 16px",textAlign:"center"}}>
                    {r.isExtra && (
                      <button onClick={()=>handleDeleteExtra(r.code)} title="ลบรายการงานเพิ่ม"
                        style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:13}}>✕</button>
                    )}
                  </td>
                </tr>
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
                {fmt(filtered.reduce((s,r)=>s+(parseFloat(tenderCosts[r.code])||0),0))}
              </td>
              <td/>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:13}}>
                {fmt(filtered.reduce((s,r)=>s+(parseFloat(draftAdd[r.code])||0),0))}
              </td>
              <td/>
              <td style={{padding:"12px 16px",textAlign:"right",color:T.green,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:14}}>
                {fmt(filtered.reduce((s,r)=>{
                  const baseVal = parseFloat(tenderCosts[r.code]) || 0;
                  const cumBefore = months.filter(m=>m<month).reduce((ss,m)=>ss+(parseFloat(additions[m]?.[r.code])||0),0)+baseVal;
                  return s + cumBefore + (parseFloat(draftAdd[r.code])||0);
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

// ─── Procurement View ─────────────────────────────────────────────────────────
function ProcurementView({ project, tenderCosts, poEntries, savePO, onBack, syncedAt, syncing, session, onLogout }) {
  const [view,   setView]   = useState("list");
  const [form,   setForm]   = useState({code:"",supplier:"",poNumber:"",amount:"",date:new Date().toISOString().slice(0,10),status:"PO Issued",notes:""});
  const [editId, setEditId] = useState(null);
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");

  const tenderTotal = Object.values(tenderCosts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const totalComm   = poEntries.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalPaid   = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+(parseFloat(p.amount)||0),0);

  const submit = () => {
    if (!form.code||!form.amount) return;
    savePO(editId ? poEntries.map(p=>p.id===editId?{...form,id:editId}:p) : [...poEntries,{...form,id:uid()}]);
    setEditId(null);
    setForm({code:"",supplier:"",poNumber:"",amount:"",date:new Date().toISOString().slice(0,10),status:"PO Issued",notes:""});
    setView("list");
  };

  const filtered = poEntries.filter(p=>{
    const acc=ACCOUNTS.find(a=>a.code===p.code);
    return (filter==="All"||p.status===filter)&&
      (search===""||[acc?.name,p.supplier,p.poNumber].join(" ").toLowerCase().includes(search.toLowerCase()));
  });

  return (
    <Shell role="procurement" color={T.amber} project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout}>
      <div style={{padding:"24px 28px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:24}}>
          <StatCard label="Budget (QS)" value={fmt(tenderTotal)} sub="Tender Cost รวม" color={T.blue} icon="📋" accent={T.blueLight}/>
          <StatCard label="Committed (PO)" value={fmt(totalComm)} sub={`${poEntries.length} รายการ`} color={T.amber} icon="📦" accent={T.amberBg}/>
          <StatCard label="ชำระแล้ว" value={fmt(totalPaid)} sub={`${poEntries.filter(p=>p.status==="Paid").length} รายการ`} color={T.green} icon="✅" accent={T.greenBg}/>
          <StatCard label="Budget คงเหลือ" value={fmt(tenderTotal-totalComm)} sub={tenderTotal>0?`${((totalComm/tenderTotal)*100).toFixed(1)}% ใช้ไปแล้ว`:"—"} color={tenderTotal-totalComm<0?T.red:T.textSecondary} icon={tenderTotal-totalComm<0?"⚠️":"💰"} accent={tenderTotal-totalComm<0?T.redBg:"#f8fafc"}/>
        </div>

        {view==="add" ? (
          <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:16,padding:28,maxWidth:680,animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:T.textPrimary}}>{editId?"แก้ไขรายการ PO":"เพิ่ม PO ใหม่"}</div>
                <div style={{fontSize:12,color:T.textMuted,marginTop:2}}>กรอกข้อมูลคำสั่งซื้อ</div>
              </div>
              <button onClick={()=>{setView("list");setEditId(null);}} style={{background:T.bg,border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:T.textMuted}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>หมวดต้นทุน *</span>
                <select value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))} className="input-base">
                  <option value="">— เลือก Account Code —</option>
                  {ACCOUNTS.map(a=><option key={a.code} value={a.code}>{a.code} · {a.name}</option>)}
                </select>
              </label>
              {[["ชื่อ Supplier *","supplier","text"],["เลข PO","poNumber","text"],["มูลค่า (THB) *","amount","number"],["วันที่","date","date"]].map(([l,k,t])=>(
                <label key={k} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{l}</span>
                  <input type={t} value={form[k]} onChange={e=>setForm(f=>({...f,[k]:e.target.value}))} className="input-base"/>
                </label>
              ))}
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>สถานะ</span>
                <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} className="input-base">
                  {PO_STATUS.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>หมายเหตุ</span>
                <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} className="input-base" style={{resize:"vertical"}}/>
              </label>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button onClick={submit} className="btn-primary" style={{background:T.amber,color:"#fff"}}>{editId?"อัปเดต":"เพิ่ม PO"}</button>
              <button onClick={()=>{setView("list");setEditId(null);}} className="btn-ghost">ยกเลิก</button>
            </div>
          </div>
        ) : (
          <>
            <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,padding:"14px 18px",marginBottom:16,display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา supplier, PO..."
                className="input-base" style={{width:220}}/>
              <div style={{display:"flex",gap:5,flex:1,flexWrap:"wrap"}}>
                {["All",...PO_STATUS].map(s=>(
                  <button key={s} onClick={()=>setFilter(s)}
                    style={{background:filter===s?T.amber:"transparent",border:`1.5px solid ${filter===s?T.amber:T.cardBorder}`,borderRadius:8,padding:"4px 11px",color:filter===s?"#fff":T.textSecondary,fontSize:11,cursor:"pointer",fontWeight:500,transition:"all 0.15s"}}>{s}</button>
                ))}
              </div>
              <button onClick={()=>setView("add")} className="btn-primary" style={{background:T.amber}}>+ เพิ่ม PO</button>
            </div>

            {filtered.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
                <div style={{fontSize:32,marginBottom:12}}>📋</div>
                <div style={{fontSize:14,fontWeight:500,color:T.textSecondary,marginBottom:6}}>ยังไม่มีรายการ</div>
                <div style={{fontSize:12}}>กด "+ เพิ่ม PO" เพื่อเริ่มต้น</div>
              </div>
            ) : (
              <div style={{background:T.card,border:`1px solid ${T.cardBorder}`,borderRadius:14,overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#f8fafc"}}>
                      {["วันที่","Account","Supplier","PO No.","มูลค่า (THB)","สถานะ",""].map(h=>(
                        <th key={h} style={{padding:"11px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((p,i)=>{
                      const acc=ACCOUNTS.find(a=>a.code===p.code);
                      return (
                        <tr key={p.id} style={{background:i%2===0?T.card:"#fafbfd",borderBottom:`1px solid #f1f5f9`}}>
                          <td style={{padding:"10px 16px",color:T.textMuted,fontSize:12,fontFamily:"'JetBrains Mono',monospace"}}>{p.date}</td>
                          <td style={{padding:"10px 16px"}}>
                            <div style={{color:T.blue,fontSize:11,fontFamily:"'JetBrains Mono',monospace",fontWeight:500}}>{p.code}</div>
                            <div style={{color:T.textSecondary,fontSize:11,marginTop:2}}>{acc?.name}</div>
                          </td>
                          <td style={{padding:"10px 16px",color:T.textPrimary,fontWeight:500}}>{p.supplier}</td>
                          <td style={{padding:"10px 16px",color:T.textMuted,fontFamily:"'JetBrains Mono',monospace",fontSize:12}}>{p.poNumber||"—"}</td>
                          <td style={{padding:"10px 16px",textAlign:"right",color:T.textPrimary,fontFamily:"'JetBrains Mono',monospace",fontWeight:600}}>{fmt(p.amount)}</td>
                          <td style={{padding:"10px 16px"}}>
                            <span style={{background:STATUS_BG[p.status],color:STATUS_CLR[p.status],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{p.status}</span>
                          </td>
                          <td style={{padding:"10px 16px",whiteSpace:"nowrap"}}>
                            <button onClick={()=>{setForm({...p});setEditId(p.id);setView("add");}} style={{background:"none",border:"none",color:T.textMuted,cursor:"pointer",padding:"2px 6px",borderRadius:6,marginRight:4}}>✏️</button>
                            <button onClick={()=>savePO(poEntries.filter(x=>x.id!==p.id))} style={{background:"none",border:"none",color:T.red,cursor:"pointer",padding:"2px 6px",borderRadius:6}}>🗑</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{background:"#f8fafc",borderTop:`2px solid ${T.cardBorder}`}}>
                      <td colSpan={4} style={{padding:"12px 16px",color:T.textMuted,fontSize:12}}>{filtered.length} รายการ</td>
                      <td style={{padding:"12px 16px",textAlign:"right",color:T.amber,fontFamily:"'JetBrains Mono',monospace",fontWeight:700,fontSize:14}}>
                        {fmt(filtered.reduce((s,p)=>s+(parseFloat(p.amount)||0),0))}
                      </td>
                      <td colSpan={2}/>
                    </tr>
                  </tfoot>
                </table>
              </div>
            )}
          </>
        )}
      </div>
    </Shell>
  );
}

// ─── Accounting View ──────────────────────────────────────────────────────────
function AccountingView({ project, tenderCosts, poEntries, onBack, onExport, syncedAt, syncing, session, onLogout }) {
  const [view, setView] = useState("dashboard");

  const tenderTotal   = Object.values(tenderCosts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const totalComm     = poEntries.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalPaid     = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalInvoiced = poEntries.filter(p=>["Invoiced","Paid"].includes(p.status)).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const pct           = tenderTotal>0?(totalComm/tenderTotal*100):0;

  const groupData = GROUPS.map((g,i)=>{
    const codes=ACCOUNTS.filter(a=>a.group===g).map(a=>a.code);
    return {group:g,budget:codes.reduce((s,c)=>s+(parseFloat(tenderCosts[c])||0),0),committed:poEntries.filter(p=>codes.includes(p.code)).reduce((s,p)=>s+(parseFloat(p.amount)||0),0),color:GRP_COLORS[i%GRP_COLORS.length]};
  }).filter(g=>g.budget>0||g.committed>0);

  const accountData = ACCOUNTS.map(a=>{
    const budget=parseFloat(tenderCosts[a.code])||0;
    const pos=poEntries.filter(p=>p.code===a.code);
    return {...a,budget,committed:pos.reduce((s,p)=>s+(parseFloat(p.amount)||0),0),pos,over:pos.reduce((s,p)=>s+(parseFloat(p.amount)||0),0)>budget&&budget>0};
  }).filter(a=>a.budget>0||a.pos.length>0);

  const pieData = PO_STATUS.map(s=>({name:s,value:poEntries.filter(p=>p.status===s).reduce((sum,p)=>sum+(parseFloat(p.amount)||0),0),color:STATUS_CLR[s]})).filter(d=>d.value>0);

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
              <StatCard label="งบประมาณ (QS)" value={fmt(tenderTotal)} sub="Tender Cost" color={T.blue} icon="📋" accent={T.blueLight}/>
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
                  {["Acc. Code","Account Name","Group","Budget (QS)","Committed (PO)","% Used",""].map(h=>(
                    <th key={h} style={{padding:"11px 16px",textAlign:["Budget (QS)","Committed (PO)","% Used"].includes(h)?"right":"left",color:T.textMuted,fontWeight:600,fontSize:11,letterSpacing:0.8,textTransform:"uppercase",borderBottom:`1px solid ${T.cardBorder}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {accountData.map((a,i)=>{
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
