import { useState, useEffect, useCallback } from "react";
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

const GROUPS     = [...new Set(ACCOUNTS.map(a => a.group))];
const PO_STATUS  = ["Pending","PO Issued","Delivered","Invoiced","Paid"];
const STATUS_CLR = { Pending:"#94a3b8","PO Issued":"#3b82f6",Delivered:"#f59e0b",Invoiced:"#a78bfa",Paid:"#22c55e" };
const STATUS_BG  = { Pending:"#1e293b","PO Issued":"#1e3a5f",Delivered:"#292001",Invoiced:"#2d1b69",Paid:"#052e16" };
const GRP_COLORS = ["#3b82f6","#22c55e","#f59e0b","#ef4444","#8b5cf6","#06b6d4","#f97316","#ec4899"];
const fmt  = n => new Intl.NumberFormat("th-TH",{minimumFractionDigits:2,maximumFractionDigits:2}).format(n||0);
const fmtK = n => n>=1e6?`${(n/1e6).toFixed(1)}M`:n>=1e3?`${(n/1e3).toFixed(0)}K`:fmt(n);
const uid  = () => Math.random().toString(36).slice(2,10);

// ─── Design tokens ────────────────────────────────────────────────────────────
const C = {
  bg:      "#0f172a",
  surface: "#1e293b",
  border:  "#334155",
  borderL: "#1e293b",
  text:    "#f1f5f9",
  textSub: "#94a3b8",
  textMut: "#64748b",
  blue:    "#3b82f6",
  green:   "#22c55e",
  amber:   "#f59e0b",
  red:     "#ef4444",
  purple:  "#a78bfa",
};

const G = {
  card: { background:C.surface, border:`1px solid ${C.border}`, borderRadius:16, padding:24 },
  input: { background:"#0f172a", border:`1px solid ${C.border}`, borderRadius:10, padding:"10px 14px",
           color:C.text, fontSize:14, fontFamily:"system-ui,sans-serif", outline:"none", width:"100%", boxSizing:"border-box" },
  label: { display:"flex", flexDirection:"column", gap:6 },
  labelText: { fontSize:12, fontWeight:600, color:C.textSub, textTransform:"uppercase", letterSpacing:"0.05em" },
};

// ─── CSS Global ───────────────────────────────────────────────────────────────
const CSS = `
  * { box-sizing: border-box; }
  body { margin:0; font-family: system-ui, -apple-system, sans-serif; background:${C.bg}; color:${C.text}; }
  input:focus, select:focus, textarea:focus { border-color: ${C.blue} !important; box-shadow: 0 0 0 3px ${C.blue}22; }
  input[type=number]::-webkit-inner-spin-button { opacity:0.4; }
  ::-webkit-scrollbar { width:6px; height:6px; }
  ::-webkit-scrollbar-track { background:${C.bg}; }
  ::-webkit-scrollbar-thumb { background:${C.border}; border-radius:3px; }
  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.3} }
  @keyframes fadeIn { from{opacity:0;transform:translateY(8px)} to{opacity:1;transform:none} }
  .card-hover:hover { border-color: ${C.blue}55 !important; transform:translateY(-2px); box-shadow: 0 8px 24px #00000040; }
  .row-hover:hover { background: ${C.surface} !important; }
  tr.row-hover:hover td { background: transparent; }
`;

// ─── Excel Export (unchanged) ─────────────────────────────────────────────────
function exportToExcel(project, tenderCosts, poEntries) {
  const wb = XLSX.utils.book_new();
  const summaryRows = [];
  summaryRows.push([`Project: ${project.name}`,"","","","",""]);
  summaryRows.push([`Area: ${project.area} ft²`,"",`Panels: ${project.panels}`,"","",""]);
  summaryRows.push([`Export Date: ${new Date().toLocaleDateString("th-TH")}`,"","","","",""]);
  summaryRows.push([]);
  summaryRows.push(["Acc. Code","Account Name","Group","Budget / Tender Cost","Committed (PO)","Remaining","% Used","Status"]);
  let grandBudget=0, grandCommitted=0;
  ACCOUNTS.forEach(a => {
    const budget    = parseFloat(tenderCosts[a.code]) || 0;
    const committed = poEntries.filter(p=>p.code===a.code).reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
    const remaining = budget - committed;
    const pct       = budget>0?(committed/budget*100).toFixed(1)+"%" : committed>0?"No Budget":"—";
    if (budget===0 && committed===0) return;
    grandBudget+=budget; grandCommitted+=committed;
    summaryRows.push([a.code,a.name,a.group,budget||"",committed||"",remaining,pct,committed>budget&&budget>0?"⚠ Over Budget":"OK"]);
  });
  summaryRows.push([]);
  summaryRows.push(["TOTAL","","",grandBudget,grandCommitted,grandBudget-grandCommitted,grandBudget>0?((grandCommitted/grandBudget)*100).toFixed(1)+"%":""]);
  const ws1 = XLSX.utils.aoa_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, ws1, "Cost Summary");

  const poRows = [["Date","Account Code","Account Name","Supplier","PO Number","Amount","Status","Notes"]];
  poEntries.forEach(p => {
    const acc = ACCOUNTS.find(a=>a.code===p.code);
    poRows.push([p.date,p.code,acc?.name||"",p.supplier,p.poNumber||"",parseFloat(p.amount)||0,p.status,p.notes||""]);
  });
  const ws2 = XLSX.utils.aoa_to_sheet(poRows);
  XLSX.utils.book_append_sheet(wb, ws2, "PO Entries");
  XLSX.writeFile(wb, `${project.name}_TenderCost.xlsx`);
}

// ─── Stat Card ────────────────────────────────────────────────────────────────
function StatCard({ label, value, sub, color, icon }) {
  return (
    <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:14,padding:"18px 20px",display:"flex",flexDirection:"column",gap:6}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontSize:12,fontWeight:600,color:C.textMut,textTransform:"uppercase",letterSpacing:"0.05em"}}>{label}</span>
        {icon && <span style={{fontSize:20}}>{icon}</span>}
      </div>
      <div style={{fontSize:22,fontWeight:700,color:color||C.text,letterSpacing:"-0.02em"}}>{value}</div>
      {sub && <div style={{fontSize:12,color:C.textSub}}>{sub}</div>}
    </div>
  );
}

// ─── Topbar ───────────────────────────────────────────────────────────────────
function Topbar({ left, center, right, syncedAt, syncing }) {
  return (
    <div style={{background:"#080f1e",borderBottom:`1px solid ${C.border}`,padding:"0 28px",height:60,display:"flex",alignItems:"center",gap:16,position:"sticky",top:0,zIndex:50}}>
      {left}
      {center && <div style={{flex:1,textAlign:"center"}}>{center}</div>}
      <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:12}}>
        {/* Sync status */}
        <div style={{display:"flex",alignItems:"center",gap:6,fontSize:11,color:C.textMut,background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:"5px 10px"}}>
          <span style={{width:7,height:7,borderRadius:"50%",background:syncing?C.amber:C.green,display:"inline-block",boxShadow:`0 0 6px ${syncing?C.amber:C.green}`,animation:syncing?"pulse 0.8s infinite":"none"}}/>
          <span style={{color:syncing?C.amber:C.textSub}}>{syncing?"กำลัง sync...":`sync ${syncedAt?.toLocaleTimeString("th-TH",{hour:"2-digit",minute:"2-digit",second:"2-digit"})||""}`}</span>
        </div>
        {right}
      </div>
    </div>
  );
}

// ─── Loader ───────────────────────────────────────────────────────────────────
function Loader() {
  return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg,flexDirection:"column",gap:16}}>
      <div style={{width:36,height:36,border:`3px solid ${C.border}`,borderTop:`3px solid ${C.blue}`,borderRadius:"50%",animation:"pulse 0.8s ease-in-out infinite"}}/>
      <div style={{fontSize:14,color:C.textSub}}>กำลังโหลด...</div>
    </div>
  );
}

// ─── Home Screen ──────────────────────────────────────────────────────────────
function HomeScreen({ projects, saveProjects, openProject, deleteProject, newProjModal, setNewProjModal, syncedAt, syncing }) {
  const [draft, setDraft] = useState({ name:"", area:"", panels:"", client:"", currency:"THB" });

  const createProject = () => {
    if (!draft.name.trim()) return;
    const id = uid();
    saveProjects([...projects, { ...draft, id, createdAt: new Date().toISOString() }]);
    setNewProjModal(false);
    setDraft({ name:"", area:"", panels:"", client:"", currency:"THB" });
  };

  return (
    <div style={{minHeight:"100vh",background:C.bg}}>
      <style>{CSS}</style>
      <Topbar syncedAt={syncedAt} syncing={syncing}
        left={
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <div style={{width:36,height:36,background:"linear-gradient(135deg,#3b82f6,#1d4ed8)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🏗</div>
            <div>
              <div style={{fontSize:15,fontWeight:700,color:C.text}}>Tender Cost System</div>
              <div style={{fontSize:11,color:C.textMut}}>ระบบบริหารต้นทุนโครงการ</div>
            </div>
          </div>
        }
        right={
          <button onClick={()=>setNewProjModal(true)}
            style={{background:C.blue,color:"#fff",border:"none",borderRadius:10,padding:"9px 20px",fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:16}}>＋</span> โครงการใหม่
          </button>
        }
      />

      <div style={{padding:"32px",maxWidth:1200,margin:"0 auto",animation:"fadeIn 0.3s ease"}}>
        {projects.length === 0 ? (
          <div style={{textAlign:"center",padding:"100px 0",color:C.textSub}}>
            <div style={{fontSize:64,marginBottom:20}}>🏗️</div>
            <div style={{fontSize:20,fontWeight:700,color:C.text,marginBottom:8}}>ยังไม่มีโครงการ</div>
            <div style={{fontSize:14,color:C.textSub,marginBottom:28}}>เริ่มต้นด้วยการสร้างโครงการใหม่</div>
            <button onClick={()=>setNewProjModal(true)}
              style={{background:C.blue,color:"#fff",border:"none",borderRadius:10,padding:"12px 28px",fontSize:14,fontWeight:700,cursor:"pointer"}}>
              + สร้างโครงการแรก
            </button>
          </div>
        ) : (
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
              <div style={{fontSize:24,fontWeight:700,color:C.text}}>โครงการทั้งหมด
                <span style={{marginLeft:10,fontSize:14,fontWeight:400,color:C.textMut,background:C.surface,border:`1px solid ${C.border}`,borderRadius:20,padding:"2px 10px"}}>{projects.length}</span>
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:20}}>
              {projects.map(p => <ProjectCard key={p.id} project={p} onOpen={()=>openProject(p.id)} onDelete={()=>deleteProject(p.id)} />)}
            </div>
          </>
        )}
      </div>

      {/* Modal */}
      {newProjModal && (
        <div style={{position:"fixed",inset:0,background:"#000c",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200,backdropFilter:"blur(4px)"}}
          onClick={e=>{if(e.target===e.currentTarget)setNewProjModal(false)}}>
          <div style={{background:"#131f35",border:`1px solid ${C.border}`,borderRadius:20,padding:32,width:500,maxWidth:"92vw",animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
              <h3 style={{margin:0,fontSize:18,fontWeight:700,color:C.text}}>สร้างโครงการใหม่</h3>
              <button onClick={()=>setNewProjModal(false)} style={{background:"none",border:"none",color:C.textMut,cursor:"pointer",fontSize:22,lineHeight:1}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
              {[
                ["ชื่อโครงการ *","name","text","1/-1"],
                ["ชื่อลูกค้า / Client","client","text","1/-1"],
                ["พื้นที่รวม (ft²)","area","number","auto"],
                ["จำนวน Panels","panels","number","auto"],
                ["สกุลเงิน","currency","text","auto"],
              ].map(([label,key,type,col]) => (
                <label key={key} style={{...G.label,gridColumn:col}}>
                  <span style={G.labelText}>{label}</span>
                  <input type={type} value={draft[key]} onChange={e=>setDraft(d=>({...d,[key]:e.target.value}))}
                    style={G.input} onKeyDown={e=>key==="currency"&&e.key==="Enter"&&createProject()} />
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10,marginTop:24}}>
              <button onClick={createProject} disabled={!draft.name.trim()}
                style={{flex:1,background:draft.name.trim()?C.blue:"#1e3a5f",color:"#fff",border:"none",borderRadius:10,padding:"12px 0",fontSize:14,fontWeight:700,cursor:draft.name.trim()?"pointer":"not-allowed",transition:"background 0.2s"}}>
                สร้างโครงการ
              </button>
              <button onClick={()=>setNewProjModal(false)}
                style={{padding:"12px 20px",background:"transparent",border:`1px solid ${C.border}`,borderRadius:10,color:C.textSub,fontSize:14,cursor:"pointer"}}>
                ยกเลิก
              </button>
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
    <div className="card-hover" style={{...G.card,cursor:"pointer",transition:"all 0.2s",position:"relative"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
        <div style={{width:42,height:42,borderRadius:12,background:"linear-gradient(135deg,#1d4ed8,#3b82f6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>🏛️</div>
        <button onClick={e=>{e.stopPropagation();if(confirm("ลบโครงการนี้?"))onDelete();}}
          style={{background:C.bg,border:`1px solid ${C.border}`,color:C.textMut,cursor:"pointer",fontSize:13,padding:"4px 10px",borderRadius:8,transition:"all 0.15s"}}
          onMouseEnter={e=>{e.currentTarget.style.borderColor=C.red;e.currentTarget.style.color=C.red;}}
          onMouseLeave={e=>{e.currentTarget.style.borderColor=C.border;e.currentTarget.style.color=C.textMut;}}>
          ลบ
        </button>
      </div>
      <div style={{fontSize:17,fontWeight:700,color:C.text,marginBottom:4,lineHeight:1.3}}>{project.name}</div>
      {project.client && <div style={{fontSize:13,color:C.textSub,marginBottom:12}}>{project.client}</div>}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
        {project.area   && <span style={{fontSize:11,background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 8px",color:C.textSub}}>{project.area} ft²</span>}
        {project.panels && <span style={{fontSize:11,background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 8px",color:C.textSub}}>{project.panels} Panels</span>}
        {project.currency && <span style={{fontSize:11,background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:"3px 8px",color:C.textSub}}>{project.currency}</span>}
      </div>
      <div style={{fontSize:11,color:C.textMut,marginBottom:16}}>{age === 0 ? "สร้างวันนี้" : `${age} วันที่แล้ว`}</div>
      <button onClick={onOpen}
        style={{width:"100%",background:C.blue,color:"#fff",border:"none",borderRadius:10,padding:"11px 0",fontSize:14,fontWeight:700,cursor:"pointer",transition:"opacity 0.15s"}}
        onMouseEnter={e=>e.currentTarget.style.opacity="0.9"}
        onMouseLeave={e=>e.currentTarget.style.opacity="1"}>
        เปิดโครงการ →
      </button>
    </div>
  );
}

// ─── Role Select ──────────────────────────────────────────────────────────────
function RoleSelect({ project, updateProject, onSelect, onBack, syncedAt, syncing }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(project);
  useEffect(()=>setDraft(project),[project]);

  const roles = [
    {id:"qs",       icon:"📐",label:"QS",     sub:"Quantity Surveyor",  desc:"ลงประมาณการ Tender Cost ตาม Account Code", color:C.blue,  grad:"linear-gradient(135deg,#1d4ed8,#3b82f6)"},
    {id:"procurement",icon:"📦",label:"จัดซื้อ",sub:"Procurement",       desc:"ออก PO ติดตามสถานะและมูลค่าจริงที่ซื้อ",   color:C.amber, grad:"linear-gradient(135deg,#92400e,#f59e0b)"},
    {id:"accounting",icon:"📊",label:"บัญชี",  sub:"Accounting",         desc:"Dashboard ต้นทุน Budget vs Actual และ Export Excel", color:C.green, grad:"linear-gradient(135deg,#14532d,#22c55e)"},
  ];

  return (
    <div style={{minHeight:"100vh",background:C.bg}}>
      <style>{CSS}</style>
      <Topbar syncedAt={syncedAt} syncing={syncing}
        left={
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <button onClick={onBack} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,color:C.textSub,cursor:"pointer",padding:"6px 12px",fontSize:13,display:"flex",alignItems:"center",gap:6}}>
              ← กลับ
            </button>
            <div style={{width:1,height:24,background:C.border}}/>
            <div>
              <div style={{fontSize:15,fontWeight:700,color:C.text}}>{project.name}</div>
              {project.client&&<div style={{fontSize:11,color:C.textMut}}>{project.client}</div>}
            </div>
          </div>
        }
      />

      <div style={{padding:"40px 32px",maxWidth:900,margin:"0 auto",animation:"fadeIn 0.3s ease"}}>
        {/* Project info bar */}
        <div style={{...G.card,marginBottom:28,display:"flex",alignItems:"center",gap:20,flexWrap:"wrap"}}>
          <div style={{flex:1}}>
            <div style={{fontSize:20,fontWeight:700,color:C.text,marginBottom:4}}>{project.name}</div>
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              {project.client&&<span style={{fontSize:13,color:C.textSub}}>👤 {project.client}</span>}
              {project.area&&<span style={{fontSize:13,color:C.textSub}}>📐 {project.area} ft²</span>}
              {project.panels&&<span style={{fontSize:13,color:C.textSub}}>🪟 {project.panels} Panels</span>}
              {project.currency&&<span style={{fontSize:13,color:C.textSub}}>💱 {project.currency}</span>}
            </div>
          </div>
          <button onClick={()=>setEditing(e=>!e)}
            style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:8,padding:"8px 16px",color:C.textSub,fontSize:13,cursor:"pointer"}}>
            {editing?"ยกเลิก":"✏️ แก้ไขข้อมูล"}
          </button>
        </div>

        {editing && (
          <div style={{...G.card,marginBottom:24,animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:16}}>
              {[["ชื่อโครงการ","name","text"],["ลูกค้า","client","text"],["สกุลเงิน","currency","text"],["พื้นที่ (ft²)","area","number"],["Panels","panels","number"]].map(([l,k,t])=>(
                <label key={k} style={G.label}>
                  <span style={G.labelText}>{l}</span>
                  <input type={t} value={draft[k]||""} onChange={e=>setDraft(d=>({...d,[k]:e.target.value}))} style={G.input}/>
                </label>
              ))}
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>{updateProject(draft);setEditing(false);}}
                style={{background:C.blue,color:"#fff",border:"none",borderRadius:8,padding:"10px 24px",fontSize:13,fontWeight:700,cursor:"pointer"}}>บันทึก</button>
              <button onClick={()=>setEditing(false)}
                style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:8,padding:"10px 20px",color:C.textSub,fontSize:13,cursor:"pointer"}}>ยกเลิก</button>
            </div>
          </div>
        )}

        <div style={{fontSize:13,color:C.textSub,marginBottom:16}}>เลือก Role เพื่อเข้าใช้งาน</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:20}}>
          {roles.map(r=>(
            <button key={r.id} onClick={()=>onSelect(r.id)} className="card-hover"
              style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:18,padding:"28px 24px",cursor:"pointer",textAlign:"left",fontFamily:"inherit",display:"flex",flexDirection:"column",gap:14,transition:"all 0.2s"}}>
              <div style={{width:52,height:52,borderRadius:14,background:r.grad,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>{r.icon}</div>
              <div>
                <div style={{fontSize:22,fontWeight:800,color:r.color,letterSpacing:"-0.02em"}}>{r.label}</div>
                <div style={{fontSize:12,color:C.textMut,marginTop:2,fontWeight:500}}>{r.sub}</div>
              </div>
              <p style={{margin:0,fontSize:13,color:C.textSub,lineHeight:1.6}}>{r.desc}</p>
              <div style={{fontSize:13,color:r.color,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>เข้าใช้งาน <span>→</span></div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── Shell (inner layout for QS/Procurement/Accounting) ──────────────────────
function Shell({ role, color, label, icon, project, onBack, children, syncedAt, syncing }) {
  return (
    <div style={{minHeight:"100vh",background:C.bg,display:"flex",flexDirection:"column"}}>
      <style>{CSS}</style>
      <Topbar syncedAt={syncedAt} syncing={syncing}
        left={
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <button onClick={onBack} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,color:C.textSub,cursor:"pointer",padding:"6px 12px",fontSize:13}}>← กลับ</button>
            <div style={{width:1,height:24,background:C.border}}/>
            <div style={{width:32,height:32,borderRadius:8,background:color+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>{icon}</div>
            <div>
              <div style={{fontSize:14,fontWeight:700,color:color}}>{label}</div>
              <div style={{fontSize:11,color:C.textMut}}>{project.name}</div>
            </div>
          </div>
        }
        right={
          project.area && (
            <div style={{fontSize:11,color:C.textMut,textAlign:"right"}}>
              <div>{project.area} ft²  ·  {project.panels} Panels</div>
            </div>
          )
        }
      />
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

  useEffect(()=>setDraft({...tenderCosts}),[tenderCosts]);

  const base  = Object.values(draft).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const adj3  = base * 0.03;
  const total = base + adj3;

  const filtered = ACCOUNTS.filter(a =>
    (filter==="All"||a.group===filter) &&
    (a.name.toLowerCase().includes(search.toLowerCase())||a.code.includes(search))
  );

  const handleSave = () => {
    const clean = {};
    Object.entries(draft).forEach(([k,v])=>{if(v!==""&&!isNaN(v)&&parseFloat(v)>0)clean[k]=parseFloat(v);});
    saveTenders(clean);
    setSaved(true); setTimeout(()=>setSaved(false),2500);
  };

  return (
    <Shell role="qs" color={C.blue} label="QS · Quantity Surveyor" icon="📐" project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing}>
      <div style={{padding:"28px 32px",maxWidth:1200,margin:"0 auto",animation:"fadeIn 0.3s ease"}}>
        {/* Stats */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:28}}>
          <StatCard label="Tender Cost รวม"   value={fmt(base)}  color={C.blue}  icon="📋" sub={`${Object.values(draft).filter(v=>parseFloat(v)>0).length} รายการที่มีค่า`}/>
          <StatCard label="Spare & Wastage 3%" value={fmt(adj3)}  color={C.amber} icon="📦"/>
          <StatCard label="Total Adjusted"     value={fmt(total)} color={C.green} icon="✅"/>
        </div>

        {/* Toolbar */}
        <div style={{...G.card,marginBottom:20,padding:"16px 20px",display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
          <div style={{position:"relative",flex:"0 0 220px"}}>
            <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:C.textMut,fontSize:14}}>🔍</span>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหา Account..."
              style={{...G.input,paddingLeft:36}}/>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",flex:1}}>
            {["All",...GROUPS].map(g=>(
              <button key={g} onClick={()=>setFilter(g)}
                style={{background:filter===g?C.blue+"22":"transparent",border:`1px solid ${filter===g?C.blue:C.border}`,
                        borderRadius:8,padding:"6px 12px",color:filter===g?C.blue:C.textSub,fontSize:12,cursor:"pointer",fontWeight:filter===g?600:400,transition:"all 0.15s"}}>
                {g}
              </button>
            ))}
          </div>
          <button onClick={handleSave}
            style={{background:saved?"#052e16":C.blue,color:saved?C.green:"#fff",border:`1px solid ${saved?C.green:C.blue}`,borderRadius:10,padding:"10px 24px",fontSize:14,fontWeight:700,cursor:"pointer",transition:"all 0.3s",whiteSpace:"nowrap"}}>
            {saved?"✓ บันทึกแล้ว":"💾 บันทึก"}
          </button>
        </div>

        {/* Table */}
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,overflow:"hidden"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead>
              <tr style={{background:"#0f1e35"}}>
                {["Acc. Code","Group","Account Name","Tender Cost (THB)"].map(h=>(
                  <th key={h} style={{padding:"14px 18px",textAlign:h.includes("Cost")?"right":"left",color:C.textMut,fontWeight:600,fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((a,i)=>(
                <tr key={a.code} className="row-hover" style={{borderBottom:`1px solid ${C.borderL}`,background:i%2===0?"transparent":"#16213a"}}>
                  <td style={{padding:"12px 18px",color:C.blue,fontFamily:"monospace",fontSize:12,fontWeight:600}}>{a.code}</td>
                  <td style={{padding:"12px 18px"}}>
                    <span style={{background:C.bg,border:`1px solid ${C.border}`,color:C.textSub,fontSize:11,padding:"3px 10px",borderRadius:20}}>{a.group}</span>
                  </td>
                  <td style={{padding:"12px 18px",color:C.text,fontSize:13}}>{a.name}</td>
                  <td style={{padding:"10px 18px",textAlign:"right"}}>
                    <input type="number" value={draft[a.code]??""} onChange={e=>setDraft(d=>({...d,[a.code]:e.target.value}))}
                      placeholder="0.00"
                      style={{background:C.bg,border:`1px solid ${parseFloat(draft[a.code])>0?C.blue+"55":C.border}`,borderRadius:8,padding:"8px 12px",
                              color:C.text,fontSize:13,fontFamily:"monospace",outline:"none",width:160,textAlign:"right",transition:"border-color 0.2s"}}/>
                  </td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr style={{background:"#0f1e35"}}>
                <td colSpan={3} style={{padding:"12px 18px",color:C.textMut,fontSize:12}}>{filtered.length} รายการ</td>
                <td style={{padding:"12px 18px",textAlign:"right",color:C.blue,fontFamily:"monospace",fontWeight:700,fontSize:15}}>
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

  const tenderTotal = Object.values(tenderCosts).reduce((s,v)=>s+(parseFloat(v)||0),0);
  const totalComm   = poEntries.reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const totalPaid   = poEntries.filter(p=>p.status==="Paid").reduce((s,p)=>s+(parseFloat(p.amount)||0),0);
  const remaining   = tenderTotal - totalComm;
  const pctUsed     = tenderTotal>0?(totalComm/tenderTotal*100):0;

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

  const inputStyle = {...G.input};

  return (
    <Shell role="procurement" color={C.amber} label="จัดซื้อ · Procurement" icon="📦" project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing}>
      <div style={{padding:"28px 32px",maxWidth:1200,margin:"0 auto",animation:"fadeIn 0.3s ease"}}>

        {/* Stats */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:28}}>
          <StatCard label="Budget (QS)"    value={fmt(tenderTotal)} color={C.blue}  icon="📋" sub="Tender Cost"/>
          <StatCard label="Committed (PO)" value={fmt(totalComm)}   color={C.amber} icon="📦" sub={`${poEntries.length} รายการ · ${pctUsed.toFixed(1)}% ของงบ`}/>
          <StatCard label="ชำระแล้ว"        value={fmt(totalPaid)}   color={C.green} icon="✅" sub={`${poEntries.filter(p=>p.status==="Paid").length} รายการ`}/>
          <StatCard label="คงเหลือ"         value={fmt(remaining)}   color={remaining<0?C.red:C.textSub} icon={remaining<0?"⚠️":"💰"} sub={remaining<0?"เกินงบ!":"ยังอยู่ในงบ"}/>
        </div>

        {/* Progress bar */}
        <div style={{...G.card,marginBottom:24,padding:"16px 20px"}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
            <span style={{fontSize:13,fontWeight:600,color:C.textSub}}>Budget Utilization</span>
            <span style={{fontSize:13,fontWeight:700,color:pctUsed>100?C.red:pctUsed>80?C.amber:C.green,fontFamily:"monospace"}}>{pctUsed.toFixed(1)}%</span>
          </div>
          <div style={{background:C.bg,borderRadius:99,height:10,overflow:"hidden"}}>
            <div style={{width:`${Math.min(pctUsed,100)}%`,background:pctUsed>100?C.red:pctUsed>80?C.amber:C.green,height:"100%",borderRadius:99,transition:"width 0.6s ease"}}/>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",marginTop:6,fontSize:11,color:C.textMut}}>
            <span>0</span><span>{fmt(tenderTotal)}</span>
          </div>
        </div>

        {/* Form */}
        {view==="add" ? (
          <div style={{...G.card,marginBottom:24,animation:"fadeIn 0.2s ease"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <h3 style={{margin:0,fontSize:16,fontWeight:700,color:C.text}}>{editId?"✏️ แก้ไขรายการ PO":"➕ เพิ่ม PO ใหม่"}</h3>
              <button onClick={()=>{setView("list");setEditId(null);}} style={{background:"none",border:"none",color:C.textMut,cursor:"pointer",fontSize:22}}>×</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
              <label style={{...G.label,gridColumn:"1/-1"}}>
                <span style={G.labelText}>หมวดต้นทุน *</span>
                <select value={form.code} onChange={e=>setForm(f=>({...f,code:e.target.value}))} style={inputStyle}>
                  <option value="">— เลือก Account Code —</option>
                  {ACCOUNTS.map(a=><option key={a.code} value={a.code}>{a.code} · {a.name}</option>)}
                </select>
              </label>
              {[["ชื่อ Supplier *","supplier","text"],["เลข PO","poNumber","text"],["มูลค่า (THB) *","amount","number"],["วันที่","date","date"]].map(([l,k,t])=>(
                <label key={k} style={G.label}>
                  <span style={G.labelText}>{l}</span>
                  <input type={t} value={form[k]} onChange={e=>setForm(f=>({...f,[k]:e.target.value}))} style={inputStyle}/>
                </label>
              ))}
              <label style={G.label}>
                <span style={G.labelText}>สถานะ</span>
                <select value={form.status} onChange={e=>setForm(f=>({...f,status:e.target.value}))} style={inputStyle}>
                  {PO_STATUS.map(s=><option key={s} value={s}>{s}</option>)}
                </select>
              </label>
              <label style={{...G.label,gridColumn:"1/-1"}}>
                <span style={G.labelText}>หมายเหตุ</span>
                <textarea value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} rows={2} style={{...inputStyle,resize:"vertical"}}/>
              </label>
            </div>
            <div style={{display:"flex",gap:10,marginTop:20}}>
              <button onClick={submit} disabled={!form.code||!form.amount}
                style={{background:form.code&&form.amount?C.amber:"#292001",color:form.code&&form.amount?"#000":C.textMut,border:"none",borderRadius:10,padding:"11px 28px",fontSize:14,fontWeight:700,cursor:form.code&&form.amount?"pointer":"not-allowed"}}>
                {editId?"อัปเดต PO":"➕ เพิ่ม PO"}
              </button>
              <button onClick={()=>{setView("list");setEditId(null);}}
                style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:10,padding:"11px 20px",color:C.textSub,fontSize:14,cursor:"pointer"}}>
                ยกเลิก
              </button>
            </div>
          </div>
        ) : (
          /* List toolbar */
          <div style={{...G.card,marginBottom:20,padding:"14px 18px",display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            <div style={{position:"relative",flex:"0 0 220px"}}>
              <span style={{position:"absolute",left:12,top:"50%",transform:"translateY(-50%)",color:C.textMut}}>🔍</span>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="ค้นหา supplier, PO..."
                style={{...G.input,paddingLeft:36}}/>
            </div>
            <div style={{display:"flex",gap:6,flexWrap:"wrap",flex:1}}>
              {["All",...PO_STATUS].map(s=>(
                <button key={s} onClick={()=>setFilter(s)}
                  style={{background:filter===s?(STATUS_CLR[s]||C.amber)+"22":"transparent",
                          border:`1px solid ${filter===s?(STATUS_CLR[s]||C.amber):C.border}`,
                          borderRadius:8,padding:"6px 12px",color:filter===s?(STATUS_CLR[s]||C.amber):C.textSub,
                          fontSize:12,cursor:"pointer",fontWeight:filter===s?600:400,transition:"all 0.15s"}}>
                  {s}
                </button>
              ))}
            </div>
            <button onClick={()=>setView("add")}
              style={{background:C.amber,color:"#000",border:"none",borderRadius:10,padding:"10px 20px",fontSize:13,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>
              ＋ เพิ่ม PO
            </button>
          </div>
        )}

        {/* PO Table */}
        {view==="list" && (
          filtered.length===0 ? (
            <div style={{...G.card,textAlign:"center",padding:"60px 0",color:C.textSub}}>
              <div style={{fontSize:40,marginBottom:12}}>📋</div>
              <div style={{fontSize:15,fontWeight:600,color:C.text,marginBottom:8}}>ยังไม่มีรายการ PO</div>
              <div style={{fontSize:13}}>กด "＋ เพิ่ม PO" เพื่อเริ่มต้น</div>
            </div>
          ) : (
            <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,overflow:"hidden"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead>
                  <tr style={{background:"#0f1e35"}}>
                    {["วันที่","Account","Supplier","PO No.","มูลค่า (THB)","สถานะ",""].map(h=>(
                      <th key={h} style={{padding:"13px 16px",textAlign:h==="มูลค่า (THB)"?"right":"left",color:C.textMut,fontWeight:600,fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.map((p,i)=>{
                    const acc=ACCOUNTS.find(a=>a.code===p.code);
                    return (
                      <tr key={p.id} className="row-hover" style={{borderBottom:`1px solid ${C.borderL}`,background:i%2===0?"transparent":"#16213a"}}>
                        <td style={{padding:"12px 16px",color:C.textMut,fontSize:12,fontFamily:"monospace"}}>{p.date}</td>
                        <td style={{padding:"12px 16px"}}>
                          <div style={{color:C.blue,fontSize:11,fontFamily:"monospace",fontWeight:600}}>{p.code}</div>
                          <div style={{color:C.textSub,fontSize:12,marginTop:2,maxWidth:180,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{acc?.name}</div>
                        </td>
                        <td style={{padding:"12px 16px",color:C.text,fontSize:13,fontWeight:500}}>{p.supplier}</td>
                        <td style={{padding:"12px 16px",color:C.textSub,fontFamily:"monospace",fontSize:12}}>{p.poNumber||"—"}</td>
                        <td style={{padding:"12px 16px",textAlign:"right",color:C.text,fontFamily:"monospace",fontWeight:700,fontSize:14}}>{fmt(p.amount)}</td>
                        <td style={{padding:"12px 16px"}}>
                          <span style={{background:STATUS_BG[p.status],color:STATUS_CLR[p.status],fontSize:11,padding:"4px 12px",borderRadius:20,fontWeight:600,border:`1px solid ${STATUS_CLR[p.status]}44`}}>
                            {p.status}
                          </span>
                        </td>
                        <td style={{padding:"12px 16px",whiteSpace:"nowrap"}}>
                          <button onClick={()=>{setForm({...p});setEditId(p.id);setView("add");}}
                            style={{background:"transparent",border:`1px solid ${C.border}`,borderRadius:6,color:C.textSub,cursor:"pointer",padding:"4px 10px",fontSize:12,marginRight:6}}>แก้ไข</button>
                          <button onClick={()=>{if(confirm("ลบรายการนี้?"))savePO(poEntries.filter(x=>x.id!==p.id));}}
                            style={{background:"transparent",border:`1px solid ${C.red}44`,borderRadius:6,color:C.red,cursor:"pointer",padding:"4px 10px",fontSize:12}}>ลบ</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{background:"#0f1e35"}}>
                    <td colSpan={4} style={{padding:"12px 16px",color:C.textMut,fontSize:12}}>{filtered.length} รายการ</td>
                    <td style={{padding:"12px 16px",textAlign:"right",color:C.amber,fontFamily:"monospace",fontWeight:700,fontSize:15}}>
                      {fmt(filtered.reduce((s,p)=>s+(parseFloat(p.amount)||0),0))}
                    </td>
                    <td colSpan={2}/>
                  </tr>
                </tfoot>
              </table>
            </div>
          )
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

  const CT = ({active,payload})=>{
    if(!active||!payload?.length) return null;
    return (
      <div style={{background:"#0f1e30",border:`1px solid ${C.border}`,borderRadius:10,padding:"10px 14px",fontSize:12}}>
        <div style={{color:C.textSub,marginBottom:6,fontWeight:600}}>{payload[0]?.payload?.group}</div>
        {payload.map(p=><div key={p.name} style={{color:p.fill||p.color,marginBottom:2}}>{p.name}: <b>{fmt(p.value)}</b></div>)}
      </div>
    );
  };

  return (
    <Shell role="accounting" color={C.green} label="บัญชี · Accounting" icon="📊" project={project} onBack={onBack} syncedAt={syncedAt} syncing={syncing}>
      <div style={{padding:"28px 32px",maxWidth:1400,margin:"0 auto",animation:"fadeIn 0.3s ease"}}>

        {/* Tab bar */}
        <div style={{display:"flex",gap:8,marginBottom:24,alignItems:"center"}}>
          {[["dashboard","📊 Dashboard"],["detail","📋 รายละเอียด"]].map(([v,l])=>(
            <button key={v} onClick={()=>setView(v)}
              style={{background:view===v?C.green+"22":"transparent",border:`1px solid ${view===v?C.green:C.border}`,
                      borderRadius:10,padding:"9px 20px",color:view===v?C.green:C.textSub,
                      fontSize:13,cursor:"pointer",fontWeight:view===v?700:400,transition:"all 0.15s"}}>
              {l}
            </button>
          ))}
          <button onClick={onExport}
            style={{marginLeft:"auto",background:"#052e16",border:`1px solid ${C.green}`,borderRadius:10,padding:"9px 20px",
                    color:C.green,fontSize:13,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
            ⬇️ Export Excel
          </button>
        </div>

        {view==="dashboard" ? (
          <>
            {/* KPI Cards */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:24}}>
              <StatCard label="งบประมาณ (QS)"      value={fmt(tenderTotal)}   color={C.blue}   icon="📋" sub="Tender Cost"/>
              <StatCard label="Committed (PO รวม)"  value={fmt(totalComm)}     color={C.amber}  icon="📦" sub={`${pct.toFixed(1)}% ของงบ`}/>
              <StatCard label="Invoiced"             value={fmt(totalInvoiced)} color={C.purple} icon="🧾" sub="รอจ่าย + จ่ายแล้ว"/>
              <StatCard label="ชำระแล้ว (Paid)"      value={fmt(totalPaid)}     color={C.green}  icon="✅" sub={`${poEntries.filter(p=>p.status==="Paid").length} รายการ`}/>
            </div>

            {/* Progress */}
            <div style={{...G.card,marginBottom:24,padding:"20px 24px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                <span style={{fontSize:14,fontWeight:700,color:C.text}}>Budget Utilization</span>
                <div style={{display:"flex",gap:16,fontSize:13}}>
                  <span style={{color:C.textSub}}>ใช้ไป <b style={{color:C.amber}}>{pct.toFixed(1)}%</b></span>
                  <span style={{color:tenderTotal-totalComm<0?C.red:C.green,fontWeight:700}}>
                    {tenderTotal-totalComm<0?"⚠ เกินงบ ":"✓ คงเหลือ "}{fmt(Math.abs(tenderTotal-totalComm))}
                  </span>
                </div>
              </div>
              <div style={{background:C.bg,borderRadius:99,height:14,overflow:"hidden"}}>
                <div style={{width:`${Math.min(pct,100)}%`,background:pct>100?C.red:pct>80?C.amber:C.green,
                             height:"100%",borderRadius:99,transition:"width 0.6s ease",
                             boxShadow:`0 0 8px ${pct>100?C.red:pct>80?C.amber:C.green}66`}}/>
              </div>
            </div>

            {/* Charts */}
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:20}}>
              <div style={{...G.card}}>
                <div style={{fontSize:14,fontWeight:700,color:C.text,marginBottom:20}}>📊 Budget vs Committed (ตาม Group)</div>
                {groupData.length===0
                  ? <div style={{textAlign:"center",padding:"50px 0",color:C.textSub}}>QS ยังไม่ได้ลง Tender Cost</div>
                  : <ResponsiveContainer width="100%" height={280}>
                      <BarChart data={groupData} margin={{left:0,right:0,top:4,bottom:48}}>
                        <XAxis dataKey="group" tick={{fill:C.textMut,fontSize:10}} angle={-30} textAnchor="end" interval={0}/>
                        <YAxis tick={{fill:C.textMut,fontSize:10}} tickFormatter={fmtK} width={68}/>
                        <Tooltip content={<CT/>}/>
                        <Bar dataKey="budget" name="Budget" fill={C.blue} radius={[6,6,0,0]}/>
                        <Bar dataKey="committed" name="Committed" fill={C.amber} radius={[6,6,0,0]}/>
                      </BarChart>
                    </ResponsiveContainer>
                }
              </div>
              <div style={{...G.card}}>
                <div style={{fontSize:14,fontWeight:700,color:C.text,marginBottom:20}}>🥧 สถานะ PO</div>
                {pieData.length===0
                  ? <div style={{textAlign:"center",padding:"50px 0",color:C.textSub}}>ยังไม่มี PO</div>
                  : <ResponsiveContainer width="100%" height={240}>
                      <PieChart>
                        <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="45%" outerRadius={80} innerRadius={40}>
                          {pieData.map((d,i)=><Cell key={i} fill={d.color}/>)}
                        </Pie>
                        <Tooltip formatter={v=>fmt(v)} contentStyle={{background:"#0f1e30",border:`1px solid ${C.border}`,borderRadius:10,fontSize:12}}/>
                        <Legend iconType="circle" wrapperStyle={{fontSize:11,color:C.textSub}}/>
                      </PieChart>
                    </ResponsiveContainer>
                }
              </div>
            </div>
          </>
        ) : (
          /* Detail Table */
          <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:16,overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead>
                <tr style={{background:"#0f1e35"}}>
                  {["Acc. Code","Account Name","Group","Budget (QS)","Committed (PO)","% Used","สถานะ"].map(h=>(
                    <th key={h} style={{padding:"13px 16px",textAlign:["Budget (QS)","Committed (PO)","% Used"].includes(h)?"right":"left",
                                        color:C.textMut,fontWeight:600,fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase",borderBottom:`1px solid ${C.border}`}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {accountData.map((a,i)=>{
                  const p2=a.budget>0?(a.committed/a.budget*100):a.committed>0?999:0;
                  return (
                    <tr key={a.code} className="row-hover" style={{background:a.over?"#1f0808":i%2===0?"transparent":"#16213a",borderBottom:`1px solid ${C.borderL}`}}>
                      <td style={{padding:"12px 16px",color:C.blue,fontFamily:"monospace",fontSize:12,fontWeight:600}}>{a.code}</td>
                      <td style={{padding:"12px 16px",color:C.text,fontSize:13}}>{a.name}</td>
                      <td style={{padding:"12px 16px"}}><span style={{background:C.bg,border:`1px solid ${C.border}`,color:C.textSub,fontSize:11,padding:"3px 10px",borderRadius:20}}>{a.group}</span></td>
                      <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"monospace",color:C.blue,fontWeight:600}}>{a.budget>0?fmt(a.budget):"—"}</td>
                      <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"monospace",color:a.over?C.red:C.amber,fontWeight:a.over?700:400}}>{a.committed>0?fmt(a.committed):"—"}</td>
                      <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"monospace",
                                  color:p2>100?C.red:p2>80?C.amber:C.green,fontSize:12,fontWeight:600}}>
                        {a.budget>0?`${p2.toFixed(1)}%`:a.committed>0?"No Budget":"—"}
                      </td>
                      <td style={{padding:"12px 16px"}}>
                        {a.over
                          ? <span style={{background:"#ef444422",color:C.red,fontSize:11,padding:"4px 12px",borderRadius:20,fontWeight:700,border:`1px solid ${C.red}44`}}>⚠ เกินงบ</span>
                          : a.committed>0
                            ? <span style={{background:"#22c55e22",color:C.green,fontSize:11,padding:"4px 12px",borderRadius:20,fontWeight:600,border:`1px solid ${C.green}44`}}>✓ OK</span>
                            : a.budget>0
                              ? <span style={{background:C.surface,color:C.textMut,fontSize:11,padding:"4px 12px",borderRadius:20,border:`1px solid ${C.border}`}}>ยังไม่ PO</span>
                              : null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr style={{background:"#0f1e35"}}>
                  <td colSpan={3} style={{padding:"12px 16px",color:C.textMut,fontSize:12}}>{accountData.length} รายการ</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"monospace",color:C.blue,fontWeight:700}}>{fmt(accountData.reduce((s,a)=>s+a.budget,0))}</td>
                  <td style={{padding:"12px 16px",textAlign:"right",fontFamily:"monospace",color:C.amber,fontWeight:700}}>{fmt(accountData.reduce((s,a)=>s+a.committed,0))}</td>
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

// ─── App Root ─────────────────────────────────────────────────────────────────
export default function App() {
  const [projects,   setProjects]  = useState([]);
  const [screen,     setScreen]    = useState("home");
  const [activeId,   setActiveId]  = useState(null);
  const [role,       setRole]      = useState(null);
  const [tenderCosts,setTCosts]    = useState({});
  const [poEntries,  setPO]        = useState([]);
  const [loaded,     setLoaded]    = useState(false);
  const [newProjModal,setNewProjModal] = useState(false);
  const [syncedAt,   setSyncedAt]  = useState(null);
  const [syncing,    setSyncing]   = useState(false);

  const fetchProjectData = useCallback(async id=>{
    const t=await sg(`tcs-tenders-${id}`); const po=await sg(`tcs-po-${id}`);
    setTCosts(t||{}); setPO(po||[]);
  },[]);
  const fetchProjects = useCallback(async()=>{
    const list=await sg("tcs-projects"); if(list)setProjects(list);
  },[]);

  useEffect(()=>{(async()=>{await fetchProjects();setLoaded(true);setSyncedAt(new Date());})();},[fetchProjects]);
  useEffect(()=>{if(!activeId)return;fetchProjectData(activeId);},[activeId,fetchProjectData]);

  useEffect(()=>{
    const channel=supabase.channel("kv_changes")
      .on("postgres_changes",{event:"*",schema:"public",table:"kv_store"},async payload=>{
        const key=payload.new?.key||payload.old?.key||"";
        setSyncing(true);
        if(key==="tcs-projects") await fetchProjects();
        else if(activeId&&(key===`tcs-tenders-${activeId}`||key===`tcs-po-${activeId}`)) await fetchProjectData(activeId);
        setSyncedAt(new Date()); setSyncing(false);
      }).subscribe();
    return ()=>supabase.removeChannel(channel);
  },[activeId,fetchProjects,fetchProjectData]);

  const saveProjects = useCallback(list=>{setProjects(list);ss("tcs-projects",list).then(()=>setSyncedAt(new Date()));},[]);
  const saveTenders  = useCallback(t   =>{setTCosts(t);    ss(`tcs-tenders-${activeId}`,t).then(()=>setSyncedAt(new Date()));},[activeId]);
  const savePO       = useCallback(po  =>{setPO(po);        ss(`tcs-po-${activeId}`,po).then(()=>setSyncedAt(new Date()));},[activeId]);

  const openProject  = id=>{setActiveId(id);setRole(null);setTCosts({});setPO([]);setScreen("roleSelect");};
  const updateProject= p =>{const list=projects.map(x=>x.id===p.id?p:x);saveProjects(list);};
  const deleteProject= async id=>{
    const list=projects.filter(p=>p.id!==id);
    saveProjects(list); await sd(`tcs-tenders-${id}`); await sd(`tcs-po-${id}`);
    if(activeId===id){setActiveId(null);setScreen("home");}
  };

  const activeProject = projects.find(p=>p.id===activeId)||{};
  const sharedProps   = {project:activeProject,tenderCosts,poEntries,saveTenders,savePO,updateProject,onBack:()=>setScreen("roleSelect"),syncedAt,syncing};

  if(!loaded) return <Loader/>;

  return (
    <>
      <style>{CSS}</style>
      {screen==="home" && (
        <HomeScreen projects={projects} saveProjects={saveProjects} openProject={openProject}
          deleteProject={deleteProject} newProjModal={newProjModal} setNewProjModal={setNewProjModal}
          syncedAt={syncedAt} syncing={syncing}/>
      )}
      {screen==="roleSelect" && (
        <RoleSelect project={activeProject} updateProject={updateProject}
          onSelect={r=>{setRole(r);setScreen("work");}} onBack={()=>setScreen("home")}
          syncedAt={syncedAt} syncing={syncing}/>
      )}
      {screen==="work" && role==="qs"          && <QSView          {...sharedProps} saveTenders={saveTenders}/>}
      {screen==="work" && role==="procurement" && <ProcurementView {...sharedProps}/>}
      {screen==="work" && role==="accounting"  && <AccountingView  {...sharedProps} onExport={()=>exportToExcel(activeProject,tenderCosts,poEntries)}/>}
    </>
  );
}
