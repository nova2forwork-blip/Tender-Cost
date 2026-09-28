// Tender Cost — หน้าเข้าสู่ระบบ · หน้ารายชื่อโครงการ · หน้าเลือกแผนก
import { useEffect, useRef, useState } from "react";
import { sgMany } from "./supabase.js";
import { verifyLogin } from "./auth.js";
import { T, appBuildLabel, fmtK, projectSummary, t, uid } from "./core.jsx";
import { EMOJI_ICON, Ico, LangToggle, SearchInput, SyncBadge, TopBar, UserMenu } from "./ui.jsx";

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
        <div data-login-version style={{marginTop:16,textAlign:"center",fontSize:11,color:T.textMuted,fontVariantNumeric:"tabular-nums"}}>{appBuildLabel()}</div>
      </form>
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

export { LoginScreen, HomeScreen, ProjectCard, RoleSelect };
