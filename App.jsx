import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { supabase, sg, ss, sd } from "./supabase.js";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie, Legend } from "recharts";

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
const GRP_COLORS  = ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6","#06b6d4","#f97316","#ec4899"];
const fmt  = n => new Intl.NumberFormat("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2}).format(n||0);
const fmtK = n => n>=1e6?`${fmt(n/1e6)}M`:n>=1e3?`${fmt(n/1e3)}K`:fmt(n);
const uid  = () => Math.random().toString(36).slice(2,10);

// ─── Storage helpers imported from ./supabase.js ────────────────────────────

// ─── Excel Export ─────────────────────────────────────────────────────────────
function exportToExcel(project, tenderCosts, poEntries) {
  const wb = XLSX.utils.book_new();
  const headerStyle = { font:{bold:true,color:{rgb:"FFFFFF"}}, fill:{fgColor:{rgb:"0F2040"}}, alignment:{horizontal:"center"} };
  const subHeaderStyle = { font:{bold:true}, fill:{fgColor:{rgb:"1E3A5F"}} };

  // ── Sheet 1: Summary ──────────────────────────────────────────────────────
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
  const pctCol = 6; // G = % Used
  for (let r=5; r<summaryRows.length; r++) {
    const cellRef = XLSX.utils.encode_cell({r, c:pctCol});
    if (ws1[cellRef] && typeof ws1[cellRef].v === "number") {
      ws1[cellRef].z = "0.0%";
    }
    ["D","E","F"].forEach((col,i) => {
      const ref = XLSX.utils.encode_cell({r, c:3+i});
      if (ws1[ref] && typeof ws1[ref].v === "number") ws1[ref].z = '#,##0.00';
    });
  }
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");

  // ── Sheet 2: PO Entries ───────────────────────────────────────────────────
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

  // ── Sheet 3: By Group ─────────────────────────────────────────────────────
  const grpRows = [];
  grpRows.push([`Project: ${project.name} — Group Summary`, "", "", "", ""]);
  grpRows.push([]);
  grpRows.push(["Group","Budget","Committed","Remaining","% Used"]);
  GROUPS.forEach(g => {
    const codes = ACCOUNTS.filter(a=>a.group===g).map(a=>a.code);
    const b = codes.reduce((s,c)=>s+(parseFloat(tenderCosts[c])||0),0);
    const c2 = poEntries.filter(p=>codes.includes(p.code)).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
    if (b>0||c2>0) grpRows.push([g,b,c2,b-c2,b>0?c2/b:0]);
  });
  const ws3 = XLSX.utils.aoa_to_sheet(grpRows);
  ws3["!cols"] = [{wch:18},{wch:18},{wch:18},{wch:18},{wch:10}];
  for (let r=2; r<grpRows.length; r++) {
    ["B","C","D"].forEach((col,i) => {
      const ref = XLSX.utils.encode_cell({r, c:1+i});
      if (ws3[ref] && typeof ws3[ref].v === "number") ws3[ref].z = '#,##0.00';
    });
    const pRef = XLSX.utils.encode_cell({r, c:4});
    if (ws3[pRef] && typeof ws3[pRef].v === "number") ws3[pRef].z = "0.0%";
  }
  XLSX.utils.book_append_sheet(wb, ws3, "By Group");

  // ── Sheet 4: Material Cost Plan (monthly) ─────────────────────────────────
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const planRows = [];
  planRows.push([`Project: ${project.name} — Monthly Cost Plan`, ...Array(15).fill("")]);
  planRows.push([]);
  planRows.push(["Acc. Code","Account Name","Budget",...months,"Total PO"]);
  ACCOUNTS.forEach(a => {
    const budget = parseFloat(tenderCosts[a.code])||0;
    const pos = poEntries.filter(p=>p.code===a.code);
    const monthly = months.map((_,mi) => pos.filter(p=>{ const d=new Date(p.date); return d.getMonth()===mi; }).reduce((s,p)=>s+(parseFloat(p.amount)||0),0));
    const totalPO = pos.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
    if (budget>0||totalPO>0) planRows.push([a.code,a.name,budget,...monthly,totalPO]);
  });
  const ws4 = XLSX.utils.aoa_to_sheet(planRows);
  ws4["!cols"] = [{wch:12},{wch:36},{wch:16},...months.map(()=>({wch:10})),{wch:12}];
  XLSX.utils.book_append_sheet(wb, ws4, "Monthly Plan");

  XLSX.writeFile(wb, `TenderCost_${project.name.replace(/\s+/g,"_")}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── Root ─────────────────────────────────────────────────────────────────────
export default function App() {
  const [screen,   setScreen]   = useState("home");   // home | roleSelect | app
  const [projects, setProjects] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [role,     setRole]     = useState(null);
  const [tenderCosts, setTCosts]= useState({});
  const [poEntries,   setPO]    = useState([]);
  const [loaded,   setLoaded]   = useState(false);
  const [newProjModal, setNewProjModal] = useState(false);
  const [syncedAt,    setSyncedAt]    = useState(null);   // last sync timestamp
  const [syncing,     setSyncing]     = useState(false);  // spinner flag

  // ── helpers ──────────────────────────────────────────────────────────────────
  const fetchProjectData = useCallback(async (id) => {
    const t  = await sg(`tcs-tenders-${id}`);
    const po = await sg(`tcs-po-${id}`);
    setTCosts(t || {});
    setPO(po || []);
  }, []);

  const fetchProjects = useCallback(async () => {
    const list = await sg("tcs-projects");
    if (list) setProjects(list);
  }, []);

  // Boot
  useEffect(() => {
    (async () => {
      await fetchProjects();
      setLoaded(true);
      setSyncedAt(new Date());
    })();
  }, [fetchProjects]);

  // Load project data whenever activeId changes
  useEffect(() => {
    if (!activeId) return;
    fetchProjectData(activeId);
  }, [activeId, fetchProjectData]);

  // ── Supabase Realtime — รับ update ทันทีเมื่อคนอื่นแก้ข้อมูล ─────────────────
  useEffect(() => {
    const channel = supabase
      .channel("kv_changes")
      .on("postgres_changes",
          { event: "*", schema: "public", table: "kv_store" },
          async (payload) => {
            const key = payload.new?.key || payload.old?.key || "";
            setSyncing(true);
            if (key === "tcs-projects") {
              await fetchProjects();
            } else if (activeId && (key === `tcs-tenders-${activeId}` || key === `tcs-po-${activeId}`)) {
              await fetchProjectData(activeId);
            }
            setSyncedAt(new Date());
            setSyncing(false);
          }
      )
      .subscribe();
    return () => supabase.removeChannel(channel);
  }, [activeId, fetchProjects, fetchProjectData]);

  const saveProjects = useCallback((list) => { setProjects(list); ss("tcs-projects", list).then(()=>setSyncedAt(new Date())); }, []);
  const saveTenders  = useCallback((t)    => { setTCosts(t);      ss(`tcs-tenders-${activeId}`, t).then(()=>setSyncedAt(new Date())); }, [activeId]);
  const savePO       = useCallback((po)   => { setPO(po);         ss(`tcs-po-${activeId}`, po).then(()=>setSyncedAt(new Date())); },   [activeId]);

  const openProject = (id) => { setActiveId(id); setRole(null); setScreen("roleSelect"); };

  const deleteProject = async (id) => {
    if (!confirm("ลบโครงการนี้? ข้อมูลทั้งหมดจะหายถาวร")) return;
    const updated = projects.filter(p => p.id !== id);
    saveProjects(updated);
    await sd(`tcs-tenders-${id}`); await sd(`tcs-po-${id}`);
  };

  const activeProject = projects.find(p => p.id === activeId) || { name:"", area:"", panels:"" };

  const updateProject = (fields) => {
    const updated = projects.map(p => p.id === activeId ? {...p, ...fields} : p);
    saveProjects(updated);
  };

  if (!loaded) return <Loader />;

  const sharedProps = { project:activeProject, tenderCosts, poEntries, saveTenders, savePO,
    updateProject, onBack:()=>setScreen("roleSelect"), syncedAt, syncing };

  return (
    <>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet"/>

      {screen === "home" && (
        <HomeScreen projects={projects} saveProjects={saveProjects} openProject={openProject}
          deleteProject={deleteProject} newProjModal={newProjModal} setNewProjModal={setNewProjModal}
          syncedAt={syncedAt} syncing={syncing} />
      )}
      {screen === "roleSelect" && (
        <RoleSelect project={activeProject} updateProject={updateProject}
          onSelect={r=>{ setRole(r); setScreen("app"); }} onBack={()=>setScreen("home")} />
      )}
      {screen === "app" && role === "qs"          && <QSView          {...sharedProps} />}
      {screen === "app" && role === "procurement" && <ProcurementView {...sharedProps} />}
      {screen === "app" && role === "accounting"  && (
        <AccountingView {...sharedProps} onExport={() => exportToExcel(activeProject, tenderCosts, poEntries)} />
      )}
    </>
  );
}

// ─── Loader ───────────────────────────────────────────────────────────────────
function Loader() {
  return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:"#0d1520",color:"#94a3b8",fontFamily:"'IBM Plex Sans',sans-serif",fontSize:14}}>กำลังโหลด...</div>;
}

// ─── Home Screen ──────────────────────────────────────────────────────────────
function HomeScreen({ projects, saveProjects, openProject, deleteProject, newProjModal, setNewProjModal, syncedAt, syncing }) {
  const [draft, setDraft] = useState({ name:"", area:"", panels:"", client:"", currency:"THB" });

  const createProject = () => {
    if (!draft.name.trim()) return;
    const id = uid();
    const newList = [...projects, { ...draft, id, createdAt: new Date().toISOString() }];
    saveProjects(newList);
    setNewProjModal(false);
    setDraft({ name:"", area:"", panels:"", client:"", currency:"THB" });
  };

  return (
    <div style={{minHeight:"100vh",background:"#0d1520",fontFamily:"'IBM Plex Sans',sans-serif"}}>
      {/* Top bar */}
      <div style={{background:"#060e1a",borderBottom:"1px solid #1e2d3d",padding:"16px 32px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <span style={{fontSize:10,letterSpacing:3,color:"#3b82f6",textTransform:"uppercase",fontWeight:600,fontFamily:"'IBM Plex Mono',monospace"}}>TENDER COST SYSTEM</span>
          <div style={{fontSize:18,fontWeight:700,color:"#f0f6ff",marginTop:2}}>ระบบบริหารต้นทุนโครงการ</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          {/* Shared-backend sync badge */}
          <div style={{display:"flex",alignItems:"center",gap:7,background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:8,padding:"6px 12px",fontSize:11,color:"#475569",fontFamily:"'IBM Plex Mono',monospace"}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:syncing?"#f59e0b":"#10b981",display:"inline-block",boxShadow:syncing?"0 0 6px #f59e0b":"0 0 6px #10b981",animation:syncing?"pulse 0.8s ease-in-out infinite":"none"}}/>
            <span style={{color:"#64748b"}}>Shared DB</span>
            {syncedAt && <span style={{color:"#334155"}}>· {syncedAt.toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit",second:"2-digit"})}</span>}
          </div>
          <button onClick={() => setNewProjModal(true)} style={{background:"#3b82f6",color:"#fff",border:"none",borderRadius:10,padding:"10px 22px",fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:18,lineHeight:1}}>+</span> โครงการใหม่
          </button>
        </div>
      </div>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.3}}`}</style>

      <div style={{padding:"32px"}}>
        {projects.length === 0 ? (
          <div style={{textAlign:"center",padding:"80px 0",color:"#475569"}}>
            <div style={{fontSize:48,marginBottom:16}}>🏗</div>
            <div style={{fontSize:16,fontWeight:600,color:"#64748b",marginBottom:8}}>ยังไม่มีโครงการ</div>
            <div style={{fontSize:13}}>กด "โครงการใหม่" เพื่อเริ่มต้น</div>
          </div>
        ) : (
          <>
            <div style={{fontSize:12,color:"#475569",marginBottom:20,letterSpacing:1}}>{projects.length} โครงการทั้งหมด</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:20}}>
              {projects.map(p => <ProjectCard key={p.id} project={p} onOpen={()=>openProject(p.id)} onDelete={()=>deleteProject(p.id)} />)}
            </div>
          </>
        )}
      </div>

      {/* New Project Modal */}
      {newProjModal && (
        <div style={{position:"fixed",inset:0,background:"#000000cc",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100}}>
          <div style={{background:"#0f1e30",border:"1px solid #1e3a5f",borderRadius:16,padding:32,width:480,maxWidth:"90vw"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
              <h3 style={{margin:0,color:"#f0f6ff",fontSize:16}}>สร้างโครงการใหม่</h3>
              <button onClick={()=>setNewProjModal(false)} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:22}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              {[
                ["ชื่อโครงการ *","name","text","1/-1"],
                ["ชื่อลูกค้า / Client","client","text","1/-1"],
                ["พื้นที่รวม (ft²)","area","number","auto"],
                ["จำนวน Panels","panels","number","auto"],
                ["สกุลเงิน","currency","text","auto"],
              ].map(([label,key,type,col]) => (
                <label key={key} style={{display:"flex",flexDirection:"column",gap:6,gridColumn:col}}>
                  <span style={{fontSize:11,color:"#64748b",letterSpacing:1}}>{label}</span>
                  <input type={type} value={draft[key]} onChange={e=>setDraft(d=>({...d,[key]:e.target.value}))}
                    style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none"}} />
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:24}}>
              <button onClick={createProject} disabled={!draft.name.trim()} style={{background:draft.name.trim()?"#3b82f6":"#1e3a5f",color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",fontSize:13,fontWeight:700,cursor:draft.name.trim()?"pointer":"not-allowed"}}>
                สร้างโครงการ
              </button>
              <button onClick={()=>setNewProjModal(false)} style={{background:"transparent",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 24px",color:"#94a3b8",fontSize:13,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>ยกเลิก</button>
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
    <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:14,padding:24,cursor:"pointer",transition:"border-color 0.15s,transform 0.15s",position:"relative"}}
      onMouseEnter={e=>{e.currentTarget.style.borderColor="#3b82f655";e.currentTarget.style.transform="translateY(-2px)";}}
      onMouseLeave={e=>{e.currentTarget.style.borderColor="#1e2d3d";e.currentTarget.style.transform="translateY(0)";}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div style={{fontSize:11,color:"#3b82f6",fontFamily:"'IBM Plex Mono',monospace",letterSpacing:1}}>PROJECT</div>
        <button onClick={e=>{e.stopPropagation();onDelete();}} style={{background:"none",border:"none",color:"#475569",cursor:"pointer",fontSize:14,padding:4}}>🗑</button>
      </div>
      <div style={{fontSize:17,fontWeight:700,color:"#f0f6ff",marginBottom:4,lineHeight:1.3}}>{project.name}</div>
      {project.client && <div style={{fontSize:12,color:"#64748b",marginBottom:12}}>{project.client}</div>}
      <div style={{display:"flex",gap:16,fontSize:11,color:"#475569",fontFamily:"'IBM Plex Mono',monospace",marginBottom:16}}>
        {project.area   && <span>{project.area} ft²</span>}
        {project.panels && <span>{project.panels} Panels</span>}
        {project.currency && <span>{project.currency}</span>}
      </div>
      <div style={{fontSize:11,color:"#475569",marginBottom:16}}>{age === 0 ? "สร้างวันนี้" : `${age} วันที่แล้ว`}</div>
      <button onClick={onOpen} style={{width:"100%",background:"#1e3a5f",color:"#3b82f6",border:"none",borderRadius:8,padding:"10px 0",fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>
        เปิดโครงการ →
      </button>
    </div>
  );
}

// ─── Role Select ──────────────────────────────────────────────────────────────
function RoleSelect({ project, updateProject, onSelect, onBack }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(project);
  useEffect(() => setDraft(project), [project]);

  return (
    <div style={{minHeight:"100vh",background:"#0d1520",fontFamily:"'IBM Plex Sans',sans-serif"}}>
      <div style={{background:"#060e1a",borderBottom:"1px solid #1e2d3d",padding:"14px 28px",display:"flex",alignItems:"center",gap:14}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:18,padding:4}}>←</button>
        <div style={{width:1,height:20,background:"#1e2d3d"}} />
        <div>
          <div style={{fontSize:10,letterSpacing:3,color:"#3b82f6",textTransform:"uppercase",fontWeight:600,fontFamily:"'IBM Plex Mono',monospace"}}>TENDER COST SYSTEM</div>
          <div style={{fontSize:14,color:"#94a3b8",marginTop:1}}>{project.name}</div>
        </div>
        {project.area && (
          <div style={{marginLeft:"auto",textAlign:"right",fontSize:11,color:"#475569",fontFamily:"'IBM Plex Mono',monospace"}}>
            <div>{project.area} ft²</div><div>{project.panels} Panels</div>
          </div>
        )}
      </div>

      <div style={{padding:"40px 32px",maxWidth:860}}>
        {editing ? (
          <div style={{background:"#0f1e30",border:"1px solid #1e3a5f",borderRadius:12,padding:24,marginBottom:36}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
              {[["ชื่อโครงการ","name","text"],["ลูกค้า","client","text"],["สกุลเงิน","currency","text"],["พื้นที่ (ft²)","area","number"],["Panels","panels","number"]].map(([l,k,t]) => (
                <label key={k} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <span style={{fontSize:11,color:"#64748b"}}>{l}</span>
                  <input type={t} value={draft[k]||""} onChange={e=>setDraft(d=>({...d,[k]:e.target.value}))}
                    style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:8,padding:"8px 12px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none"}} />
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:16}}>
              <button onClick={()=>{updateProject(draft);setEditing(false);}} style={{background:"#3b82f6",color:"#fff",border:"none",borderRadius:8,padding:"9px 20px",fontSize:13,fontWeight:600,cursor:"pointer"}}>บันทึก</button>
              <button onClick={()=>setEditing(false)} style={{background:"transparent",border:"1px solid #1e3a5f",borderRadius:8,padding:"9px 20px",color:"#94a3b8",fontSize:13,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>ยกเลิก</button>
            </div>
          </div>
        ) : (
          <button onClick={()=>setEditing(true)} style={{marginBottom:28,background:"transparent",border:"1px dashed #1e3a5f",borderRadius:8,padding:"7px 16px",color:"#64748b",fontSize:12,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>
            ✏️ แก้ไขข้อมูลโครงการ
          </button>
        )}

        <p style={{margin:"0 0 20px",fontSize:14,color:"#64748b"}}>เลือก Role</p>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:20}}>
          {[
            {id:"qs",label:"QS",sub:"Quantity Surveyor",desc:"ลงราคา Tender Cost\nประมาณการต้นทุนโครงการ",color:"#3b82f6",bg:"#0f2040",icon:"📐"},
            {id:"procurement",label:"จัดซื้อ",sub:"Procurement",desc:"ลงราคาจริงที่ซื้อ + วันที่\nออก PO และติดตามสถานะ",color:"#f59e0b",bg:"#1f1400",icon:"📦"},
            {id:"accounting",label:"บัญชี",sub:"Accounting",desc:"Dashboard ต้นทุน\nBudget vs Actual + Export Excel",color:"#10b981",bg:"#001f14",icon:"📊"},
          ].map(r => (
            <button key={r.id} onClick={()=>onSelect(r.id)} style={{background:r.bg,border:`1px solid ${r.color}33`,borderRadius:16,padding:"28px 24px",cursor:"pointer",textAlign:"left",fontFamily:"'IBM Plex Sans',sans-serif",display:"flex",flexDirection:"column",gap:10,transition:"all 0.15s"}}
              onMouseEnter={e=>{e.currentTarget.style.border=`1px solid ${r.color}88`;e.currentTarget.style.transform="translateY(-2px)";}}
              onMouseLeave={e=>{e.currentTarget.style.border=`1px solid ${r.color}33`;e.currentTarget.style.transform="translateY(0)";}}>
              <span style={{fontSize:28}}>{r.icon}</span>
              <div><div style={{fontSize:20,fontWeight:700,color:r.color}}>{r.label}</div><div style={{fontSize:11,color:"#64748b",marginTop:2,letterSpacing:1}}>{r.sub}</div></div>
              <p style={{margin:0,fontSize:12,color:"#94a3b8",lineHeight:1.7,whiteSpace:"pre-line"}}>{r.desc}</p>
              <div style={{fontSize:12,color:r.color,fontWeight:600}}>เข้าใช้งาน →</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── Shell ────────────────────────────────────────────────────────────────────
function Shell({ role, color, project, onBack, children, syncedAt, syncing }) {
  const labels = {qs:"QS · Quantity Surveyor",procurement:"จัดซื้อ · Procurement",accounting:"บัญชี · Accounting"};
  return (
    <div style={{minHeight:"100vh",background:"#0d1520",fontFamily:"'IBM Plex Sans',sans-serif",display:"flex",flexDirection:"column"}}>
      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.3}}`}</style>
      <div style={{background:"#060e1a",borderBottom:"1px solid #1e2d3d",padding:"13px 24px",display:"flex",alignItems:"center",gap:14}}>
        <button onClick={onBack} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:18,padding:4}}>←</button>
        <div style={{width:1,height:20,background:"#1e2d3d"}} />
        <div style={{flex:1}}>
          <div style={{fontSize:10,letterSpacing:3,color,textTransform:"uppercase",fontWeight:600,fontFamily:"'IBM Plex Mono',monospace"}}>
            {labels[role]}
          </div>
          <div style={{fontSize:13,color:"#94a3b8",marginTop:1}}>{project.name}</div>
        </div>
        {/* Sync badge */}
        <div style={{display:"flex",alignItems:"center",gap:6,fontSize:10,color:"#334155",fontFamily:"'IBM Plex Mono',monospace"}}>
          <span style={{width:6,height:6,borderRadius:"50%",background:syncing?"#f59e0b":"#10b981",display:"inline-block",boxShadow:syncing?"0 0 5px #f59e0b":"0 0 5px #10b981",animation:syncing?"pulse 0.8s ease-in-out infinite":"none"}}/>
          {syncing ? <span style={{color:"#f59e0b"}}>กำลัง sync...</span>
            : syncedAt ? <span>sync {syncedAt.toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit",second:"2-digit"})}</span> : null}
        </div>
        {project.area && (
          <div style={{textAlign:"right",fontSize:11,color:"#475569",fontFamily:"'IBM Plex Mono',monospace"}}>
            <div>{project.area} ft²</div><div>{project.panels} Panels</div>
          </div>
        )}
      </div>
      <div style={{flex:1,overflow:"auto"}}>{children}</div>
    </div>
  );
}

// ─── QS View ─────────────────────────────────────────────────────────────────
function QSView({ project, tenderCosts, saveTenders, onBack, syncedAt, syncing }) {
  const [draft,  setDraft]  = useState({...tenderCosts});
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");
  const [saved,  setSaved]  = useState(false);

  useEffect(() => setDraft({...tenderCosts}), [tenderCosts]);

  const base  = Object.values(draft).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const adj3  = base * 0.03;
  const total = base + adj3;

  const filtered = ACCOUNTS.filter(a =>
    (filter==="All" || a.group===filter) &&
    (a.name.toLowerCase().includes(search.toLowerCase()) || a.code.includes(search))
  );

  const handleSave = () => {
    const clean = {};
    Object.entries(draft).forEach(([k,v]) => { if(v!==""&&!isNaN(v)&&parseFloat(v)>0) clean[k]=parseFloat(v); });
    saveTenders(clean);
    setSaved(true); setTimeout(()=>setSaved(false),2000);
  };

  return (
    <Shell role="qs" color="#3b82f6" project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing}>
      <div style={{padding:"22px 26px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,marginBottom:20}}>
          {[
            {label:"Tender Cost รวม",value:fmt(base),color:"#3b82f6"},
            {label:"Spare & Wastage 3%",value:fmt(adj3),color:"#f59e0b"},
            {label:"Total Adjusted",value:fmt(total),color:"#10b981"},
          ].map(s=>(
            <div key={s.label} style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:10,padding:"14px 18px"}}>
              <div style={{fontSize:11,color:"#64748b",marginBottom:5}}>{s.label}</div>
              <div style={{fontSize:20,fontWeight:700,color:s.color,fontFamily:"'IBM Plex Mono',monospace"}}>{s.value}</div>
            </div>
          ))}
        </div>

        <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา..." style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:8,padding:"7px 13px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none",width:180}}/>
          <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
            {["All",...GROUPS].map(g=>(
              <button key={g} onClick={()=>setFilter(g)} style={{background:filter===g?"#1e3a5f":"transparent",border:"1px solid #1e3a5f",borderRadius:6,padding:"4px 10px",color:filter===g?"#3b82f6":"#64748b",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>{g}</button>
            ))}
          </div>
          <button onClick={handleSave} style={{marginLeft:"auto",background:saved?"#10b981":"#3b82f6",color:"#fff",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:700,cursor:"pointer",transition:"background 0.3s"}}>
            {saved?"✓ บันทึกแล้ว":"บันทึก Tender Cost"}
          </button>
        </div>

        <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:12,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead>
              <tr style={{background:"#0f2040"}}>
                {["Acc. Code","Group","Account Name","Tender Cost (THB)"].map(h=>(
                  <th key={h} style={{padding:"9px 14px",textAlign:h.includes("Cost")?"right":"left",color:"#64748b",fontWeight:600,fontSize:11,letterSpacing:1,textTransform:"uppercase",borderBottom:"1px solid #1e2d3d"}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((a,i)=>(
                <tr key={a.code} style={{background:i%2===0?"transparent":"#050c17",borderBottom:"1px solid #0f1e30"}}>
                  <td style={{padding:"9px 14px",color:"#3b82f6",fontFamily:"'IBM Plex Mono',monospace",fontSize:12}}>{a.code}</td>
                  <td style={{padding:"9px 14px"}}><span style={{background:"#1e2d3d",color:"#94a3b8",fontSize:10,padding:"2px 8px",borderRadius:4}}>{a.group}</span></td>
                  <td style={{padding:"9px 14px",color:"#e2e8f0"}}>{a.name}</td>
                  <td style={{padding:"7px 14px",textAlign:"right"}}>
                    <input type="number" value={draft[a.code]??""} onChange={e=>setDraft(d=>({...d,[a.code]:e.target.value}))}
                      placeholder="0.00" style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:6,padding:"5px 10px",color:"#f0f6ff",fontSize:13,fontFamily:"'IBM Plex Mono',monospace",outline:"none",width:148,textAlign:"right"}}/>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{background:"#0f2040"}}>
                <td colSpan={3} style={{padding:"9px 14px",color:"#64748b",fontSize:12}}>{filtered.length} รายการ</td>
                <td style={{padding:"9px 14px",textAlign:"right",color:"#3b82f6",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700}}>
                  {fmt(filtered.reduce((s,a)=>s+(parseFloat(draft[a.code])||0),0))}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </Shell>
  );
}

// ─── Procurement View ─────────────────────────────────────────────────────────
function ProcurementView({ project, tenderCosts, poEntries, savePO, onBack, syncedAt, syncing }) {
  const [view,   setView]   = useState("list");
  const [form,   setForm]   = useState({code:"",supplier:"",poNumber:"",amount:"",date:new Date().toISOString().slice(0,10),status:"PO Issued",notes:""});
  const [editId, setEditId] = useState(null);
  const [filter, setFilter] = useState("All");
  const [search, setSearch] = useState("");

  const tenderTotal   = Object.values(tenderCosts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const totalComm     = poEntries.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalPaid     = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+(parseFloat(p.amount)||0),0);

  const submit = () => {
    if (!form.code||!form.amount) return;
    const updated = editId
      ? poEntries.map(p=>p.id===editId?{...form,id:editId}:p)
      : [...poEntries,{...form,id:uid()}];
    savePO(updated); setEditId(null);
    setForm({code:"",supplier:"",poNumber:"",amount:"",date:new Date().toISOString().slice(0,10),status:"PO Issued",notes:""});
    setView("list");
  };

  const filtered = poEntries.filter(p=>{
    const acc=ACCOUNTS.find(a=>a.code===p.code);
    return (filter==="All"||p.status===filter)&&
      (search===""||[acc?.name,p.supplier,p.poNumber].join(" ").toLowerCase().includes(search.toLowerCase()));
  });

  return (
    <Shell role="procurement" color="#f59e0b" project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing}>
      <div style={{padding:"22px 26px"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
          {[
            {label:"Budget (QS)",value:fmt(tenderTotal),color:"#3b82f6",sub:"Tender Cost"},
            {label:"Committed (PO)",value:fmt(totalComm),color:"#f59e0b",sub:`${poEntries.length} รายการ`},
            {label:"ชำระแล้ว",value:fmt(totalPaid),color:"#10b981",sub:`${poEntries.filter(p=>p.status==="Paid").length} รายการ`},
            {label:"Budget คงเหลือ",value:fmt(tenderTotal-totalComm),color:tenderTotal-totalComm<0?"#ef4444":"#94a3b8",sub:tenderTotal>0?`${((totalComm/tenderTotal)*100).toFixed(1)}% ใช้ไปแล้ว`:"—"},
          ].map(s=>(
            <div key={s.label} style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:10,padding:"14px 18px"}}>
              <div style={{fontSize:11,color:"#64748b",marginBottom:5}}>{s.label}</div>
              <div style={{fontSize:18,fontWeight:700,color:s.color,fontFamily:"'IBM Plex Mono',monospace"}}>{s.value}</div>
              <div style={{fontSize:11,color:"#475569",marginTop:3}}>{s.sub}</div>
            </div>
          ))}
        </div>

        {view==="add" ? (
          <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:12,padding:26,maxWidth:680}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <h3 style={{margin:0,color:"#f0f6ff",fontSize:15}}>{editId?"แก้ไขรายการ PO":"เพิ่ม PO ใหม่"}</h3>
              <button onClick={()=>{setView("list");setEditId(null);}} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",fontSize:20}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:11,color:"#64748b"}}>หมวดต้นทุน *</span>
                <select value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))} style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none"}}>
                  <option value="">— เลือก Account Code —</option>
                  {ACCOUNTS.map(a=><option key={a.code} value={a.code}>{a.code} · {a.name}</option>)}
                </select>
              </label>
              {[["ชื่อ Supplier *","supplier","text"],["เลข PO","poNumber","text"],["มูลค่า (THB) *","amount","number"],["วันที่","date","date"]].map(([l,k,t])=>(
                <label key={k} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <span style={{fontSize:11,color:"#64748b"}}>{l}</span>
                  <input type={t} value={form[k]} onChange={e=>setForm(f=>({...f,[k]:e.target.value}))}
                    style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none"}}/>
                </label>
              ))}
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:11,color:"#64748b"}}>สถานะ</span>
                <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none"}}>
                  {PO_STATUS.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6,gridColumn:"1/-1"}}>
                <span style={{fontSize:11,color:"#64748b"}}>หมายเหตุ</span>
                <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} style={{background:"#0d1520",border:"1px solid #1e3a5f",borderRadius:8,padding:"9px 12px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none",resize:"vertical"}}/>
              </label>
            </div>
            <div style={{display:"flex",gap:10,marginTop:18}}>
              <button onClick={submit} style={{background:"#f59e0b",color:"#000",border:"none",borderRadius:8,padding:"10px 24px",fontSize:13,fontWeight:700,cursor:"pointer"}}>{editId?"อัปเดต":"เพิ่ม PO"}</button>
              <button onClick={()=>{setView("list");setEditId(null);}} style={{background:"transparent",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 24px",color:"#94a3b8",fontSize:13,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>ยกเลิก</button>
            </div>
          </div>
        ) : (
          <>
            <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 ค้นหา supplier, PO..." style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:8,padding:"7px 13px",color:"#e2e8f0",fontSize:13,fontFamily:"'IBM Plex Sans',sans-serif",outline:"none",width:200}}/>
              {["All",...PO_STATUS].map(s=>(
                <button key={s} onClick={()=>setFilter(s)} style={{background:filter===s?"#1e2d3d":"transparent",border:"1px solid #1e3a5f",borderRadius:6,padding:"4px 10px",color:filter===s?"#f59e0b":"#64748b",fontSize:11,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif"}}>{s}</button>
              ))}
              <button onClick={()=>setView("add")} style={{marginLeft:"auto",background:"#f59e0b",color:"#000",border:"none",borderRadius:8,padding:"8px 20px",fontSize:13,fontWeight:700,cursor:"pointer"}}>+ เพิ่ม PO</button>
            </div>

            {filtered.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:"#475569",fontSize:14}}>
                <div style={{fontSize:32,marginBottom:12}}>📋</div>
                ยังไม่มีรายการ — กด "+ เพิ่ม PO"
              </div>
            ) : (
              <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:12,overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#0f2040"}}>
                      {["วันที่","Account","Supplier","PO No.","มูลค่า (THB)","สถานะ",""].map(h=>(
                        <th key={h} style={{padding:"9px 13px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:"#64748b",fontWeight:600,fontSize:11,letterSpacing:1,borderBottom:"1px solid #1e2d3d"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((p,i)=>{
                      const acc=ACCOUNTS.find(a=>a.code===p.code);
                      return (
                        <tr key={p.id} style={{background:i%2===0?"transparent":"#050c17",borderBottom:"1px solid #0f1e30"}}>
                          <td style={{padding:"9px 13px",color:"#64748b",fontSize:12,fontFamily:"'IBM Plex Mono',monospace"}}>{p.date}</td>
                          <td style={{padding:"9px 13px"}}>
                            <div style={{color:"#3b82f6",fontSize:11,fontFamily:"'IBM Plex Mono',monospace"}}>{p.code}</div>
                            <div style={{color:"#94a3b8",fontSize:12,marginTop:1}}>{acc?.name}</div>
                          </td>
                          <td style={{padding:"9px 13px",color:"#e2e8f0"}}>{p.supplier}</td>
                          <td style={{padding:"9px 13px",color:"#64748b",fontFamily:"'IBM Plex Mono',monospace",fontSize:12}}>{p.poNumber||"—"}</td>
                          <td style={{padding:"9px 13px",textAlign:"right",color:"#f0f6ff",fontFamily:"'IBM Plex Mono',monospace",fontWeight:600}}>{fmt(p.amount)}</td>
                          <td style={{padding:"9px 13px"}}>
                            <span style={{background:STATUS_CLR[p.status]+"22",color:STATUS_CLR[p.status],fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:600}}>{p.status}</span>
                          </td>
                          <td style={{padding:"9px 13px",whiteSpace:"nowrap"}}>
                            <button onClick={()=>{setForm({...p});setEditId(p.id);setView("add");}} style={{background:"none",border:"none",color:"#64748b",cursor:"pointer",marginRight:6}}>✏️</button>
                            <button onClick={()=>savePO(poEntries.filter(x=>x.id!==p.id))} style={{background:"none",border:"none",color:"#ef4444",cursor:"pointer"}}>🗑</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{background:"#0f2040"}}>
                      <td colSpan={4} style={{padding:"9px 13px",color:"#64748b",fontSize:12}}>{filtered.length} รายการ</td>
                      <td style={{padding:"9px 13px",textAlign:"right",color:"#f59e0b",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700,fontSize:14}}>
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
function AccountingView({ project, tenderCosts, poEntries, onBack, onExport, syncedAt, syncing }) {
  const [view, setView] = useState("dashboard");

  const tenderTotal   = Object.values(tenderCosts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const totalComm     = poEntries.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalPaid     = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalInvoiced = poEntries.filter(p=>["Invoiced","Paid"].includes(p.status)).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const pct           = tenderTotal>0?(totalComm/tenderTotal*100):0;

  const groupData = GROUPS.map((g,i)=>{
    const codes=ACCOUNTS.filter(a=>a.group===g).map(a=>a.code);
    const budget=codes.reduce((s,c)=>s+(parseFloat(tenderCosts[c])||0),0);
    const committed=poEntries.filter(p=>codes.includes(p.code)).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
    return {group:g,budget,committed,color:GRP_COLORS[i%GRP_COLORS.length]};
  }).filter(g=>g.budget>0||g.committed>0);

  const accountData = ACCOUNTS.map(a=>{
    const budget=parseFloat(tenderCosts[a.code])||0;
    const pos=poEntries.filter(p=>p.code===a.code);
    const committed=pos.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
    return {...a,budget,committed,pos,over:committed>budget&&budget>0};
  }).filter(a=>a.budget>0||a.pos.length>0);

  const pieData = PO_STATUS.map(s=>({
    name:s, value:poEntries.filter(p=>p.status===s).reduce((sum,p)=>sum+(parseFloat(p.amount)||0),0), color:STATUS_CLR[s]
  })).filter(d=>d.value>0);

  const CT = ({active,payload}) => {
    if (!active||!payload?.length) return null;
    return (
      <div style={{background:"#0f1e30",border:"1px solid #1e3a5f",borderRadius:8,padding:"10px 14px",fontSize:12,fontFamily:"'IBM Plex Mono',monospace"}}>
        <div style={{color:"#94a3b8",marginBottom:4}}>{payload[0]?.payload?.group}</div>
        {payload.map(p=><div key={p.name} style={{color:p.fill||p.color}}>{p.name}: {fmt(p.value)}</div>)}
      </div>
    );
  };

  return (
    <Shell role="accounting" color="#10b981" project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing}>
      <div style={{padding:"22px 26px"}}>
        {/* Tabs + Export */}
        <div style={{display:"flex",gap:8,marginBottom:20,alignItems:"center"}}>
          {[["dashboard","📊 Dashboard"],["detail","📋 รายละเอียด"]].map(([v,l])=>(
            <button key={v} onClick={()=>setView(v)} style={{background:view===v?"#0f2a1a":"transparent",border:`1px solid ${view===v?"#10b981":"#1e3a5f"}`,borderRadius:8,padding:"8px 18px",color:view===v?"#10b981":"#64748b",fontSize:13,cursor:"pointer",fontFamily:"'IBM Plex Sans',sans-serif",fontWeight:view===v?600:400}}>{l}</button>
          ))}
          <button onClick={onExport} style={{marginLeft:"auto",background:"#0f2a1a",border:"1px solid #10b981",borderRadius:8,padding:"8px 18px",color:"#10b981",fontSize:13,fontWeight:600,cursor:"pointer",display:"flex",alignItems:"center",gap:6,fontFamily:"'IBM Plex Sans',sans-serif"}}>
            ⬇️ Export Excel
          </button>
        </div>

        {view==="dashboard" ? (
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
              {[
                {label:"งบประมาณ (QS)",value:fmt(tenderTotal),sub:"Tender Cost",color:"#3b82f6"},
                {label:"Committed (PO รวม)",value:fmt(totalComm),sub:`${pct.toFixed(1)}% ของงบ`,color:"#f59e0b"},
                {label:"Invoiced",value:fmt(totalInvoiced),sub:"รอจ่าย + จ่ายแล้ว",color:"#8b5cf6"},
                {label:"ชำระแล้ว (Paid)",value:fmt(totalPaid),sub:`${poEntries.filter(p=>p.status==="Paid").length} รายการ`,color:"#10b981"},
              ].map(s=>(
                <div key={s.label} style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:10,padding:"16px 18px"}}>
                  <div style={{fontSize:11,color:"#64748b",marginBottom:5}}>{s.label}</div>
                  <div style={{fontSize:19,fontWeight:700,color:s.color,fontFamily:"'IBM Plex Mono',monospace"}}>{s.value}</div>
                  <div style={{fontSize:11,color:"#475569",marginTop:3}}>{s.sub}</div>
                </div>
              ))}
            </div>

            {/* Progress bar */}
            <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:10,padding:20,marginBottom:20}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
                <span style={{fontSize:13,color:"#94a3b8",fontWeight:600}}>Budget Utilization</span>
                <span style={{fontSize:13,color:tenderTotal-totalComm<0?"#ef4444":"#10b981",fontFamily:"'IBM Plex Mono',monospace",fontWeight:700}}>
                  {tenderTotal-totalComm<0?"เกินงบ ":"คงเหลือ "}{fmt(Math.abs(tenderTotal-totalComm))}
                </span>
              </div>
              <div style={{background:"#1e2d3d",borderRadius:99,height:10,overflow:"hidden"}}>
                <div style={{width:`${Math.min(pct,100)}%`,background:pct>100?"#ef4444":pct>80?"#f59e0b":"#10b981",height:"100%",borderRadius:99,transition:"width 0.5s"}}/>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",marginTop:6,fontSize:11,color:"#475569",fontFamily:"'IBM Plex Mono',monospace"}}>
                <span>0</span><span>{pct.toFixed(1)}%</span><span>{fmt(tenderTotal)}</span>
              </div>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
              <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:10,padding:20}}>
                <p style={{margin:"0 0 14px",fontSize:13,color:"#94a3b8",fontWeight:600}}>Budget vs Committed (ตาม Group)</p>
                {groupData.length===0
                  ? <div style={{textAlign:"center",padding:"40px 0",color:"#475569",fontSize:13}}>QS ยังไม่ได้ลง Tender Cost</div>
                  : <ResponsiveContainer width="100%" height={260}>
                      <BarChart data={groupData} margin={{left:0,right:0,top:4,bottom:44}}>
                        <XAxis dataKey="group" tick={{fill:"#64748b",fontSize:10}} angle={-30} textAnchor="end" interval={0}/>
                        <YAxis tick={{fill:"#64748b",fontSize:10}} tickFormatter={fmtK} width={62}/>
                        <Tooltip content={<CT/>}/>
                        <Bar dataKey="budget" name="Budget" fill="#3b82f6" radius={[4,4,0,0]}/>
                        <Bar dataKey="committed" name="Committed" fill="#f59e0b" radius={[4,4,0,0]}/>
                      </BarChart>
                    </ResponsiveContainer>
                }
              </div>
              <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:10,padding:20}}>
                <p style={{margin:"0 0 14px",fontSize:13,color:"#94a3b8",fontWeight:600}}>สถานะ PO</p>
                {pieData.length===0
                  ? <div style={{textAlign:"center",padding:"40px 0",color:"#475569",fontSize:13}}>ยังไม่มี PO</div>
                  : <ResponsiveContainer width="100%" height={220}>
                      <PieChart>
                        <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={75} innerRadius={35}>
                          {pieData.map((d,i)=><Cell key={i} fill={d.color}/>)}
                        </Pie>
                        <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#0f1e30",border:"1px solid #1e3a5f",borderRadius:8,fontSize:11}}/>
                        <Legend iconType="circle" wrapperStyle={{fontSize:10,color:"#94a3b8"}}/>
                      </PieChart>
                    </ResponsiveContainer>
                }
              </div>
            </div>
          </>
        ) : (
          <div style={{background:"#0a1628",border:"1px solid #1e2d3d",borderRadius:12,overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead>
                <tr style={{background:"#0f2040"}}>
                  {["Acc. Code","Account Name","Group","Budget (QS)","Committed (PO)","% Used",""].map(h=>(
                    <th key={h} style={{padding:"9px 13px",textAlign:["Budget (QS)","Committed (PO)","% Used"].includes(h)?"right":"left",color:"#64748b",fontWeight:600,fontSize:11,letterSpacing:1,borderBottom:"1px solid #1e2d3d"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {accountData.map((a,i)=>{
                  const p2=a.budget>0?(a.committed/a.budget*100):a.committed>0?999:0;
                  return (
                    <tr key={a.code} style={{background:a.over?"#1f0808":i%2===0?"transparent":"#050c17",borderBottom:"1px solid #0f1e30"}}>
                      <td style={{padding:"9px 13px",color:"#3b82f6",fontFamily:"'IBM Plex Mono',monospace",fontSize:12}}>{a.code}</td>
                      <td style={{padding:"9px 13px",color:"#e2e8f0"}}>{a.name}</td>
                      <td style={{padding:"9px 13px"}}><span style={{background:"#1e2d3d",color:"#94a3b8",fontSize:10,padding:"2px 8px",borderRadius:4}}>{a.group}</span></td>
                      <td style={{padding:"9px 13px",textAlign:"right",fontFamily:"'IBM Plex Mono',monospace",color:"#3b82f6"}}>{a.budget>0?fmt(a.budget):"—"}</td>
                      <td style={{padding:"9px 13px",textAlign:"right",fontFamily:"'IBM Plex Mono',monospace",color:a.over?"#ef4444":"#f59e0b",fontWeight:a.over?700:400}}>{a.committed>0?fmt(a.committed):"—"}</td>
                      <td style={{padding:"9px 13px",textAlign:"right",fontFamily:"'IBM Plex Mono',monospace",color:p2>100?"#ef4444":p2>80?"#f59e0b":"#10b981",fontSize:12}}>
                        {a.budget>0?`${p2.toFixed(1)}%`:a.committed>0?"No Budget":"—"}
                      </td>
                      <td style={{padding:"9px 13px"}}>
                        {a.over?<span style={{background:"#ef444422",color:"#ef4444",fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:600}}>⚠ เกินงบ</span>:
                         a.committed>0?<span style={{background:"#10b98122",color:"#10b981",fontSize:11,padding:"2px 8px",borderRadius:20}}>OK</span>:
                         a.budget>0?<span style={{background:"#1e2d3d",color:"#64748b",fontSize:11,padding:"2px 8px",borderRadius:20}}>ยังไม่ PO</span>:null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"#0f2040"}}>
                  <td colSpan={3} style={{padding:"9px 13px",color:"#64748b",fontSize:12}}>{accountData.length} รายการ</td>
                  <td style={{padding:"9px 13px",textAlign:"right",fontFamily:"'IBM Plex Mono',monospace",color:"#3b82f6",fontWeight:700}}>{fmt(accountData.reduce((s,a)=>s+a.budget,0))}</td>
                  <td style={{padding:"9px 13px",textAlign:"right",fontFamily:"'IBM Plex Mono',monospace",color:"#f59e0b",fontWeight:700}}>{fmt(accountData.reduce((s,a)=>s+a.committed,0))}</td>
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
