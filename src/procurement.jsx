// Tender Cost — แผนกจัดซื้อ (รายการ PO, ฟอร์ม PO, แผนของเข้า, ติดตาม)
import { useEffect, useLayoutEffect, useRef, useState } from "react";
import { ACCOUNTS, DEFAULT_CREDIT_DAYS, codeText, HISTORY_ICON, INCOMING_BG, INCOMING_CLR, PAYMENT_BG, PAYMENT_CLR, PAYMENT_TYPE_BG, PAYMENT_TYPE_CLR, PO_STAGES, PO_STATUS, PayDateText, T, WASTE_LBL, accountOf, addDays, buildCombinedBudget, canEditPO, deliveryStatus, exportAccountList, fmt, fmt0, fmtDate, formatDateTime, hiddenSafeForPO, historyEntry, farFromPODate, incLabel, incomingStatus, poDataContext, poDataIssues, similarSupplier, unusualAmountIssues, supplierCounts, itemOrdered, itemPaidAmount, itemReceived, itemRemaining, itemSupplierName, migratePO, monthShortLabel, payLabel, payTypeLabelT, paymentStatus, poAmountForCode, poDeliveries, poHistory, poItems, poLastUpdate, poNextDueDate, poNumbersLabel, poPaidAmount, poPaidDate, poReceivedDates, poRounds, poStage, poStageLabel, poStatusLabel, poSupplierLabel, poSupplierName, poSupplierText, poSuppliers, poTotal, relativeTime, roundPaid, roundPayDate, roundReceived, t, todayStr, uiAlert, uiConfirm, uiLocale, uid, withHistory } from "./core.jsx";
import { AccountPicker, BottomNav, CurrencyControl, DateInput, FormStep, Ico, MoneyInput, SearchInput, Shell, StatCard, StatusPicker, effRate, fmtMoneyInput, usdLine, useIsPhone } from "./ui.jsx";

function PODetailModal({ po: rawPo, issues = [], onClose, onEdit, onDelete, onStatusChange, onChangePO, session, usdRate=0, readOnly=false }) {
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
  const locked = readOnly || !canEditPO(po, session);   // readOnly = บัญชีเปิดดูข้อมูลจัดซื้อ (ดูอย่างเดียว)
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
    if (locked) { setCapWarn(readOnly ? t("ดูอย่างเดียว — แก้ไขไม่ได้","View only — can't edit") : t("PO นี้รับของและจ่ายเงินครบแล้ว — แก้ยอด/วันรับได้เฉพาะ Admin","This PO is fully received & paid — only Admin can change amounts/dates")); return; }
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

        {readOnly && (
          <div data-po-view-only style={{display:"flex",alignItems:"center",gap:6,background:"#f1f5f9",border:`1px solid ${T.cardBorder}`,borderRadius:8,padding:"6px 10px",margin:"8px 0 2px",fontSize:12,color:T.textSecondary}}>
            <Ico name="eye" size={14} /> {t("ดูอย่างเดียว — บัญชีดูข้อมูล PO ได้ แต่แก้ไขไม่ได้","View only — Accounting can view this PO but not edit it")}
          </div>
        )}
        {locked && !readOnly && (
          <div style={{display:"flex",alignItems:"center",gap:6,background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"6px 10px",margin:"8px 0 2px",fontSize:11,color:"#92400e"}}>
            <Ico name="lock" size={14} /> {t("รับของและจ่ายเงินครบแล้ว — แก้ยอด/วันของเข้าจริงได้ (ลบ PO และแก้ผู้ขาย/หมวด/ยอดสั่ง เฉพาะ Admin)","Fully received & paid — actual amount/date still editable (delete PO and edit vendor/category/order: Admin only)")}
          </div>
        )}

        {issues.length > 0 && (
          <div data-po-issues role="status" style={{background:T.amberBg,border:"1px solid #fde68a",borderRadius:10,padding:"8px 12px",margin:"8px 0 2px",fontSize:12,color:"#92400e",lineHeight:1.5}}>
            <div style={{fontWeight:650,display:"flex",alignItems:"center",gap:6}}><Ico name="alert" size={14} />{t("ข้อมูลที่ควรตรวจ","Worth checking")}</div>
            {issues.map((x,i) => <div key={i} data-issue-kind={x.kind}>• {x.msg}</div>)}
          </div>
        )}

        {/* Status is a live dropdown here too — the most natural place to
            update it right after reviewing everything else on the PO. */}
        <div style={{display:"flex",gap:6,margin:"12px 0 4px",flexWrap:"wrap",alignItems:"center"}}>
          <StatusPicker status={po.status} onChange={s=>onStatusChange?.(po,s)} disabled={locked} disabledTitle={readOnly ? t("ดูอย่างเดียว — แก้ไขไม่ได้","View only — cannot edit") : undefined}/>
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
                    <span style={{fontSize:11,color:T.blue,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{codeText(it.code)||"—"}</span>
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
                          {farFromPODate(po.date, r.actualDate) && <span data-date-warn style={{fontSize:11,color:T.red}}>{t(`ห่างจากวันเปิด PO (${fmtDate(po.date)}) มาก — ปีถูกไหม`,`Far from the PO date (${fmtDate(po.date)}) — right year?`)}</span>}
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
          {!readOnly && <button
            onClick={()=>{
              // ตรวจอีกครั้งก่อนปิด: ยอดของเข้าจริงรวมของทุกรายการห้ามเกินยอดสั่ง
              const bad = po.items.find(it => { const o=itemOrdered(it); const rc=(it.rounds||[]).reduce((s,r)=>s+(parseFloat(r.actualAmount)||0),0); return o>0 && Math.round(rc*100) > Math.round(o*100); });
              if (bad) { setCapWarn(t(`⚠ ${bad.code||"รายการ"}: ยอดของเข้ารวมเกินยอดสั่ง ${fmt(itemOrdered(bad))} — แก้ให้ไม่เกินก่อนบันทึก`, `⚠ ${bad.code||"Item"}: total received exceeds the order ${fmt(itemOrdered(bad))} — fix it before saving`)); return; }
              setCapWarn(""); onClose();
            }}
            disabled={!!overCapItem} className="btn-primary"
            title={overCapItem?t(`${overCapItem.code||"รายการ"}: ยอดรวมทุกงวดเกินยอดสั่ง แก้ให้ไม่เกินก่อนบันทึก`,`${overCapItem.code||"item"}: total across rounds exceeds order — fix before saving`):undefined}
            style={overCapItem?{background:"#e2e8f0",color:"#94a3b8",cursor:"not-allowed"}:undefined}>{overCapItem && <Ico name="alert" size={15} />} {t("บันทึก","Save")}</button>}
          {!locked && <button onClick={()=>onEdit(po)} className="btn-ghost" style={{fontSize:12}} title={t("แก้ผู้ขาย / หมวด / ยอดสั่ง","Edit vendor / category / order")}><Ico name="edit" size={14} /> {t("แก้ไข PO","Edit PO")}</button>}
          {readOnly ? null : confirmDel ? (
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

// ─── จัดซื้อ: แผนของเข้าทั้งโปรเจค (ใช้ฟอร์มเดียวกับ PO) ────────────────────────
//  แผน = อ็อบเจ็กต์รูปเดียวกับ PO (isPlan:true) สร้าง/แก้ผ่านฟอร์ม PO โดยติ๊ก
//  "แผนของเข้า". แท็บนี้แค่แสดงลิสต์แผน + ปุ่มเรียกฟอร์ม. "→ ทำเป็น PO จริง" =
//  เปิดฟอร์มโดยเอาติ๊กออกให้ พอกดบันทึกก็กลายเป็น PO จริงและแผนถูกย้ายออก.
function IncomingPlanTab({ plans, poEntries = [], usdRate = 0, tenderCosts = {}, additions = {}, extraItems = [], hiddenAccounts = [], onNew, onEdit, onConvert, onDelete }) {
  // ตารางของเข้ารายเดือน: ล็อกหัวตาราง — วัดความกว้างจริงของคอลัมน์ Acc. Code ให้คอลัมน์ชื่อบัญชีตรึงต่อพอดี
  const imTableRef = useRef(null);
  useLayoutEffect(() => {
    const tb = imTableRef.current; if (!tb) return;
    const set = () => { const c0 = tb.querySelector("thead th"); if (c0) tb.style.setProperty("--c1w", `${c0.getBoundingClientRect().width}px`); };
    set();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(set); ro.observe(tb);
    return () => ro.disconnect();
  });
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

  // ขอบแบบ separate (เส้นขวา+ล่างของแต่ละช่อง) — เส้นขอบติดไปกับช่องที่ตรึง ไม่มีร่องให้ข้อความลอดตอนเลื่อน
  const cM = { borderRight: "1px solid #d9e0ea", borderBottom: "1px solid #d9e0ea", padding: "8px 13px", fontSize:13, whiteSpace: "nowrap" };
  const nM = { ...cM, textAlign: "right", fontVariantNumeric: "tabular-nums" };
  const hM = (bg) => ({ ...cM, background: bg, fontWeight: 650, color: T.textSecondary, textAlign: "center", position: "sticky", top: 0 });
  const bCost = "#f4e9ef";
  // ตรึงคอลัมน์แรก 2 ช่อง (รหัส/ชื่อบัญชี) ให้ไม่เลื่อนหายตอนดูเดือนไกล ๆ
  const COL1_W = 86;
  const C1_LEFT = `var(--c1w, ${COL1_W}px)`;
  const stickyBody0 = { position: "sticky", left: 0, background: "#fff", zIndex: 1 };
  const stickyBody1 = { position: "sticky", left: C1_LEFT, background: "#fff", zIndex: 1 };
  const stickyHead0 = { left: 0, zIndex: 3 };
  const stickyHead1 = { left: C1_LEFT, zIndex: 3 };
  const money = (n) => n ? (n < 0 ? `(${fmt(Math.abs(n))})` : fmt(n)) : "-";
  const sum = (fn) => shownCodes.reduce((s, c) => s + fn(c), 0);

  return (
    <div>
      {onNew && (
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
          <button onClick={onNew} className="btn-primary" style={{ marginLeft: "auto" }}>+ {t("เพิ่ม PO","Add PO")}</button>
        </div>
      )}

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
          <div className="fatscroll lockhead" data-lockhead style={{ border: `1px solid ${T.cardBorder}`, borderRadius: 12 }}>
            <table ref={imTableRef} style={{ borderCollapse: "separate", borderSpacing: 0, width: "max-content", minWidth: "100%" }}>
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
                      <td style={{ ...cM, ...stickyBody0, fontVariantNumeric: "tabular-nums", fontWeight: 600, color: T.blue }}>{codeText(code)}</td>
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
                  {onConvert && <button onClick={() => onConvert(pl)} className="btn-primary" style={{ fontSize: 12, padding: "6px 12px" }}>→ {t("ทำเป็น PO จริง","Make real PO")}</button>}
                  {onEdit && <button onClick={() => onEdit(pl)} className="btn-ghost" style={{ fontSize: 12, padding: "6px 10px" }}><Ico name="edit" /> {t("แก้ไข","Edit")}</button>}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  {items.map(it => (
                    <span key={it.id} style={{ background: "#f8fafc", border: `1px solid ${T.cardBorder}`, borderRadius: 8, padding: "5px 10px", fontSize: 12 }}>
                      <b style={{ fontVariantNumeric: "tabular-nums", color: T.blue }}>{codeText(it.code)}</b> {nameOf(it.code)} · <b style={{ fontVariantNumeric: "tabular-nums" }}>฿{fmt(parseFloat(it.amount) || 0)}</b>
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
// readOnly = บัญชีเปิดดูข้อมูลจัดซื้อ: เห็นทุกอย่าง (รายการ PO / แผนของเข้า / ติดตาม / รายละเอียด / Export)
// แต่ไม่มีปุ่มเพิ่ม/แก้/ลบ/เปลี่ยนสถานะ และฟังก์ชันบันทึกทุกตัวถูกตัดทิ้ง (กันหลุดจากปุ่มที่ลืมซ่อน)
function ProcurementView({ project, updateProject: updateProjectProp, tenderCosts, additions, poEntries, savePO: savePOProp, onBack, onHome, onDept, syncedAt, syncing, session, onLogout, extraItems=[], hiddenAccounts=[], onExport, setEditMode, incomingPlan={}, saveIncomingPlan: saveIncomingPlanProp, readOnly=false, deptSwitch=null }) {
  const denyRO = () => { uiAlert(t("ดูอย่างเดียว — บัญชีดูข้อมูลจัดซื้อได้ แต่แก้ไขไม่ได้","View only — Accounting can view procurement data but not edit it")); };
  const savePO           = readOnly ? denyRO : savePOProp;
  const saveIncomingPlan = readOnly ? denyRO : saveIncomingPlanProp;
  // สกุลเงิน: โหมดดูอย่างเดียว = ตั้งค่าเฉพาะเครื่องนี้ (ไม่บันทึกกลับโครงการ) เหมือนหน้าบัญชี
  const [curOverride, setCurOverride] = useState({});
  const curProject    = readOnly ? { ...project, ...curOverride } : project;
  const updateProject = readOnly ? (fields) => setCurOverride(o => ({ ...o, ...fields })) : updateProjectProp;
  const usdRate = effRate(curProject);  // อัตราแลกเปลี่ยน บาท/USD (0 = ปิดแสดง $)
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
  useEffect(() => { setEditMode?.(!readOnly && (view==="add" || detailId!=null)); return () => setEditMode?.(false); }, [view, detailId, setEditMode, readOnly]);

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

  // ─── ตรวจข้อมูลที่น่าสงสัย (เตือน ไม่บังคับ) ───────────────────────────────
  // ทั้งโครงการ: ใช้กับชิป "ข้อมูลที่ควรตรวจ" และกล่องในหน้ารายละเอียด PO
  const dataCtx = poDataContext(poEntries);
  const dataIssuesOf = (p) => poDataIssues(p, dataCtx);
  const dataIssueCount = poEntries.filter(p => dataIssuesOf(p).length > 0).length;
  // ในฟอร์ม: เทียบกับ PO อื่น (ไม่รวมใบที่กำลังแก้)
  const otherPOs = poEntries.filter(p => editingPlan || p.id !== editId);
  const formCounts = supplierCounts(otherPOs);
  const supplierOptions = Object.entries(supplierCounts(poEntries)).sort((a,b)=>b[1]-a[1]).map(([n])=>n).filter(n => !/^[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]/.test(n));
  const supSuggest = similarSupplier(form.supplier.name, formCounts);
  const dupNoPOs = (() => { const k = (form.supplier.poNumber||"").trim().toLowerCase().replace(/[\s.,;:]+$/,""); return k ? otherPOs.filter(p => poNumbersLabel(p) !== "—" && poNumbersLabel(p).trim().toLowerCase().replace(/[\s.,;:]+$/,"") === k) : []; })();
  const FORM_KINDS = ["po-no-punct","po-no-month","same-amount","far-date","decimals","supplier-mark"];
  const formIssues = poDataIssues({ id:"__form", date: form.date, status: form.status, supplier: form.supplier,
    // ส่ง items ทั้งหมด (บรรทัดว่างไม่ทำให้เตือน) — ถ้า items ว่าง migratePO จะมองเป็น PO แบบเก่า
    items: form.items }, poDataContext(otherPOs)).filter(x => FORM_KINDS.includes(x.kind)).concat(unusualAmountIssues(form.items.filter(it => {   // แก้ PO เดิม: เตือนเฉพาะรายการที่ยอดเปลี่ยน (ยอดเดิมยืนยันไปแล้ว)
      const prevPO = editId ? (editingPlan ? plans : poEntries).find(p => p.id === editId) : null;
      const prevIt = prevPO && poItems(prevPO).find(x => x.id === it.id);
      return !prevIt || String(prevIt.amount) !== String(it.amount) || prevIt.code !== it.code;
    }), otherPOs));

  const submit = async () => {
    if (readOnly) return denyRO();
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
    // ข้อมูลน่าสงสัย → ถามก่อน 1 ครั้ง (บันทึกต่อได้)
    const warnList = [
      ...(dupNoPOs.length ? [t(`เลข PO ${form.supplier.poNumber.trim()} มีอยู่แล้ว (${dupNoPOs.length} ใบ)`, `PO no. ${form.supplier.poNumber.trim()} already exists (${dupNoPOs.length})`)] : []),
      ...(supSuggest ? [t(`ชื่อ Supplier ใกล้กับ "${supSuggest.name}" ที่ใช้อยู่แล้ว`, `Supplier looks like existing "${supSuggest.name}"`)] : []),
      ...formIssues.map(x => x.msg),
    ];
    if (warnList.length && !(await uiConfirm(t("ตรวจก่อนบันทึก:\n\n","Please check before saving:\n\n") + warnList.map(w => "• " + w).join("\n"),
      { title: t("ข้อมูลอาจผิด","Data may be wrong"), okLabel: t("บันทึกแบบนี้","Save as is"), cancelLabel: t("กลับไปแก้","Go back and fix") }))) return;
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
    if (readOnly) return denyRO();
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
    if (readOnly) return denyRO();
    const patch = { ...po, status: newStatus };
    if (newStatus === "Paid") patch.paidDate = paidDate || todayStr();   // จำวันจ่ายที่กำหนดเอง
    const label = `${t("เปลี่ยนสถานะ","Status change")}: ${poStatusLabel(po.status)} → ${poStatusLabel(newStatus)}` + (newStatus === "Paid" && patch.paidDate ? ` (${t("จ่าย","paid")} ${patch.paidDate})` : "");
    const updated = withHistory(patch, historyEntry(session, "status", label));
    savePO(poEntries.map(x=>x.id===po.id?updated:x));
  };
  const changeStatus = async (po, newStatus) => {
    if (newStatus === po.status) return;
    if (readOnly) return denyRO();
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
    if (readOnly) return denyRO();
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
    if (readOnly) return denyRO();
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
  const openNewPO   = () => { if (readOnly) return denyRO(); setEditId(null); setEditingPlan(false); setForm({ ...emptyForm(), isPlan:false }); setDetailId(null); setView("add"); };
  const openEditPlan = (pl) => loadIntoForm(pl, true);   // แก้แผน (ติ๊กแผนอยู่)
  const startConvert = (pl) => loadIntoForm(pl, false);  // แปลงแผน → PO (เอาติ๊กออกให้แล้ว กดบันทึกก็เป็น PO)
  const deletePlan = async (id) => {
    if (readOnly) { denyRO(); return false; }
    const pl = (plans||[]).find(p=>p.id===id);
    const d = pl ? (poRounds(pl).map(r=>r.planDate).filter(Boolean).sort()[0] || pl.date || "") : "";
    const info = pl ? `${d||t("(ไม่มีวัน)","(no date)")}${pl.supplier?.name?` · ${pl.supplier.name}`:""} · ฿${fmt0(poItems(pl).reduce((s,it)=>s+(parseFloat(it.amount)||0),0))}` : "";
    if ((await uiConfirm(t(`ลบแผนของเข้านี้?${info?`\n\n${info}`:""}\n\n(ลบเฉพาะ "แผน" — ไม่กระทบ PO จริง)`,`Delete this incoming plan?${info?`\n\n${info}`:""}\n\n(deletes the "plan" only — real PO unaffected)`), { danger: true, okLabel: t("ลบแผน","Delete plan") }))) { saveIncomingPlan(plans.filter(pl=>pl.id!==id)); return true; }
    return false;   // กดยกเลิก — ให้ผู้เรียกรู้ (ไม่ปิดฟอร์มทิ้ง)
  };
  const deletePO = async (id, confirmed=false) => {
    if (readOnly) return denyRO();
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
    return issueFilter==="data" ? dataIssuesOf(p).length > 0 : issueFilter==="late-incoming" ? k.inc : issueFilter==="late-payment" ? k.pay : (k.inc || k.pay); };
  const showIssues = (kind) => {
    const next = issueFilter === kind ? null : kind;      // กดซ้ำ = ยกเลิกตัวกรอง
    setIssueFilter(next); setFilter("All");
    if (tab !== "list") goTab("list");
    if (next) setTimeout(() => document.querySelector("[data-po-list-top]")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50);
  };
  const filtered = poEntries.filter(p=>{
    const itemsText = poItems(p).map(it=>{ const acc=accountOf(it.code); return `${it.code} ${acc?.name||""}`; }).join(" ");
    return (filter==="All"||poStage(p)===filter)&& matchesIssue(p) &&
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
    <Shell role="procurement" color={T.amber} project={project} onBack={backNav} onHome={onHome} onDept={onDept} syncedAt={syncedAt} syncing={syncing} session={session} onLogout={onLogout} deptSwitch={deptSwitch} viewOnly={readOnly}>
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
        ...(readOnly ? [] : [{ key:"add", icon:"plus", label:t("เพิ่ม PO","Add PO"), onClick:openNewPO }]),
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
            <div style={{marginLeft:"auto"}}><CurrencyControl project={curProject} updateProject={updateProject}/></div>
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
        {view!=="add" && (lateIncomingCount>0 || latePaymentCount>0 || dataIssueCount>0) && (
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
            {dataIssueCount>0 && (
              <button className="att-chip" data-attention="data-check" aria-pressed={issueFilter==="data"} onClick={()=>showIssues("data")}
                title={t("PO ที่ข้อมูลอาจผิด: เลข PO ซ้ำ · ยอดซ้ำ · วันที่ผิดปี · ยอดรับไม่ครบ · ชื่อ Supplier สะกดต่าง","POs whose data may be wrong: duplicate no. · same amount · wrong year · short received · supplier spelling")}>
                <b>{dataIssueCount}</b> {t("ข้อมูลที่ควรตรวจ","to check")}
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
                <input placeholder={t("ชื่อ Supplier","Supplier name")} value={form.supplier.name} onChange={e=>updateSupplierField("name",e.target.value)} className="input-base" list="po-supplier-list" autoComplete="off"/>
                <datalist id="po-supplier-list">{supplierOptions.map(n => <option key={n} value={n} />)}</datalist>
                {supSuggest && (
                  <span data-sup-suggest style={{fontSize:12,color:"#92400e",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                    {t(`ใกล้กับ "${supSuggest.name}" (ใช้ใน ${supSuggest.count} ใบ)`,`Looks like "${supSuggest.name}" (on ${supSuggest.count} POs)`)}
                    <button type="button" className="btn-ghost" onClick={e=>{ e.preventDefault(); updateSupplierField("name", supSuggest.name); }} style={{fontSize:12,padding:"2px 10px",minHeight:28}}>{t("ใช้ชื่อนี้","Use this name")}</button>
                  </span>
                )}
              </label>
              <label style={{display:"flex",flexDirection:"column",gap:6}}>
                <span style={{fontSize:12,color:T.textSecondary,fontWeight:500}}>{t("เลข PO","PO no.")} {form.isPlan ? <span style={{color:T.textMuted,fontWeight:400}}>({t("ไม่บังคับ","optional")})</span> : <span style={{color:T.red}}>*</span>}</span>
                <input placeholder={t("เช่น PO-2026-001","e.g. PO-2026-001")} value={form.supplier.poNumber} onChange={e=>updateSupplierField("poNumber",e.target.value)} onBlur={()=>setPoNoTouched(true)} className="input-base"
                  aria-required={!form.isPlan ? true : undefined}
                  aria-invalid={poNoTouched && !form.isPlan && !(form.supplier.poNumber||"").trim() ? true : undefined}
                  style={poNoTouched && !form.isPlan && !(form.supplier.poNumber||"").trim() ? {borderColor:T.red, background:T.redBg} : undefined}/>
                {poNoTouched && !form.isPlan && !(form.supplier.poNumber||"").trim() && <span role="alert" data-po-error style={{fontSize:12,color:T.red}}>{t("กรอกเลข PO ก่อนบันทึก","Enter the PO number before saving")}</span>}
                {dupNoPOs.length > 0 && (
                  <span data-po-warn="dup-no" style={{fontSize:12,color:"#92400e"}}>
                    {t("เลข PO นี้มีอยู่แล้ว","This PO no. already exists")}: {dupNoPOs.slice(0,3).map(p => `${poSupplierName(p)} · ฿${fmt(poTotal(p))} · ${fmtDate(p.date)}`).join(" / ")}
                  </span>
                )}
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
            {formIssues.length > 0 && (
              <div data-form-warnings role="status" style={{marginTop:16,background:T.amberBg,border:"1px solid #fde68a",borderRadius:10,padding:"10px 12px",fontSize:12,color:"#92400e",lineHeight:1.5}}>
                <div style={{fontWeight:650,marginBottom:4,display:"flex",alignItems:"center",gap:6}}><Ico name="alert" size={14} />{t("ตรวจก่อนบันทึก","Check before saving")}</div>
                {formIssues.map((x,i) => <div key={i} data-warn-kind={x.kind}>• {x.msg}</div>)}
              </div>
            )}
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
                      <div style={{fontSize:13,marginBottom:6}}><b style={{color:T.blue,fontVariantNumeric:"tabular-nums"}}>{codeText(code)}</b> <span style={{color:T.textSecondary}}>{acc?.name||""}</span></div>
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
            <IncomingPlanTab plans={plans} poEntries={poEntries} usdRate={usdRate} tenderCosts={tenderCosts} additions={additions} extraItems={extraItems} hiddenAccounts={hiddenAccounts} onNew={readOnly ? null : openNewPO} onEdit={readOnly ? null : openEditPlan} onConvert={readOnly ? null : startConvert} onDelete={readOnly ? null : deletePlan} />
            {/* ติดตามของเข้า/จ่ายเงิน — ย้ายมาไว้ใต้ "จัดการแผน" (เอาแท็บติดตามแยกออก) */}
            <div style={{marginTop:28,paddingTop:20,borderTop:`2px solid ${T.cardBorder}`}}>
              <div style={{fontSize:15,fontWeight:650,color:T.textPrimary,marginBottom:14,display:"flex",alignItems:"center",gap:8}}><Ico name="truck" size={18} color={T.textSecondary} />{t("ติดตามของเข้า / จ่ายเงิน","Track incoming / payments")}</div>
              <ProcurementTrackingTab poEntries={poEntries} onEdit={readOnly ? null : openEdit} onView={openDetail} onAddNew={readOnly ? null : openNewPO} readOnly={readOnly}
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
                    {["All",...PO_STAGES].map(st=>{ const n = st==="All" ? bySearch.length : bySearch.filter(p=>poStage(p)===st).length; const on = filter===st; return (
                      <button key={st} onClick={()=>setFilter(st)} aria-pressed={on} data-po-filter={st}
                        title={{ waiting: t("ยังไม่ได้รับของ (รวมที่ล่าช้า)","No goods received yet (incl. late)"), partial: t("รับของแล้วบางส่วน","Some goods received"), received: t("รับของครบแล้ว ยังไม่ถึงวันจ่าย/ยังไม่จ่าย","All goods received, not paid yet"), paid: t("จ่ายครบแล้ว (ถึงกำหนดจ่าย หรือตั้งสถานะจ่ายแล้ว)","Fully paid (due date reached or set to Paid)") }[st]}
                        style={{background:on?T.textPrimary:"transparent",border:`1.5px solid ${on?T.textPrimary:T.cardBorder}`,borderRadius:8,padding:"4px 11px",color:on?"#fff":T.textSecondary,fontSize:12,cursor:"pointer",fontWeight:500,transition:"all 0.15s",minHeight:32,whiteSpace:"nowrap",flexShrink:0}}>
                        {poStageLabel(st)} <span style={{opacity:0.75,fontVariantNumeric:"tabular-nums"}}>{n}</span>
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
              {!isPhone && !readOnly && <button onClick={openNewPO} className="btn-primary">+ {t("เพิ่ม PO","Add PO")}</button>}{/* มือถือ: ใช้ปุ่ม "เพิ่ม PO" ในแถบล่าง */}
            </div>

            <div data-po-list-top style={{scrollMarginTop:12}} />
            {issueFilter && (
              <div data-issue-bar role="status" style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",background:T.redBg,border:"1px solid #fecaca",borderRadius:10,padding:"8px 12px",marginBottom:12,fontSize:13,color:T.red}}>
                <Ico name="filter" size={15} />
                <span style={{fontWeight:600}}>{t("กำลังแสดงเฉพาะ","Showing only")}: {issueFilter==="data" ? t("PO ที่ข้อมูลควรตรวจ (เปิดดูรายละเอียดได้ในแต่ละใบ)","POs to check (open each one for details)") : issueFilter==="late-incoming" ? t("ของเข้าล่าช้า","late incoming") : issueFilter==="late-payment" ? t("จ่ายเงินเกินกำหนด","overdue payments") : t("PO ที่มีปัญหา (ของเข้าล่าช้า / จ่ายเกินกำหนด)","POs with problems (late incoming / overdue payment)")}</span>
                <span style={{color:T.textSecondary}}>· {filtered.length} {t("ใบ","POs")}</span>
                <button className="btn-ghost" onClick={()=>setIssueFilter(null)} style={{marginLeft:"auto",fontSize:12,padding:"4px 10px"}}>{t("แสดงทั้งหมด","Show all")}</button>
              </div>
            )}
            {filtered.length===0 ? (
              <div style={{textAlign:"center",padding:"60px 0",color:T.textMuted}}>
                <div style={{marginBottom:12,color:T.textMuted}}><Ico name="clipboard" size={32} sw={1.5} /></div>
                <div style={{fontSize:14,fontWeight:500,color:T.textSecondary,marginBottom:6}}>{poEntries.length===0?t("ยังไม่มีรายการ","No items yet"):t("ไม่พบรายการที่ตรงเงื่อนไข","No items match")}</div>
                <div style={{fontSize:12}}>{poEntries.length===0?(readOnly?t("จัดซื้อยังไม่ได้บันทึก PO","Procurement hasn't recorded any PO yet"):t('กด "+ เพิ่ม PO" เพื่อเริ่มต้น','Press "+ Add PO" to start')):t("ลองล้างตัวกรอง หรือคำค้นหา","Try clearing filters or search")}</div>
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
                const editBtn = (p, locked) => readOnly ? null : (
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
                          <b style={{color:T.blue,fontSize:13,fontVariantNumeric:"tabular-nums"}}>{codeText(g.code)}</b>
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
                                    <StatusPicker status={p.status} onChange={st=>changeStatus(p,st)} disabled={locked || readOnly} disabledTitle={readOnly ? t("ดูอย่างเดียว — แก้ไขไม่ได้","View only — cannot edit") : undefined} compact/>
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
                                <b style={{color:T.blue,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{codeText(g.code)}</b>
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
                                    <StatusPicker status={p.status} onChange={st=>changeStatus(p,st)} disabled={locked || readOnly} disabledTitle={readOnly ? t("ดูอย่างเดียว — แก้ไขไม่ได้","View only — cannot edit") : undefined} compact/>
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
      <PODetailModal key={detailPO?.id || "none"} po={detailPO} issues={detailPO ? dataIssuesOf(detailPO) : []} onClose={closeDetail} onEdit={openEdit} onDelete={deletePO} onStatusChange={changeStatus} onChangePO={updatePO} session={session} usdRate={usdRate} readOnly={readOnly} />
    </Shell>
  );
}

// ─── Procurement: Incoming / Payment Tracking tab ─────────────────────────────
// Groups every PO by its Account Code so the team can see, at a glance and per
// cost line, which deliveries and payments are on track vs. overdue.
function ProcurementTrackingTab({ poEntries, onEdit, onView, onAddNew, onlyIssues, setOnlyIssues, onStatusChange, session, usdRate=0, tenderCosts={}, additions={}, extraItems=[], hiddenAccounts=[], readOnly=false }) {
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
            <StatusPicker status={p.status} onChange={s=>onStatusChange?.(p,s)} disabled={locked || readOnly} disabledTitle={readOnly ? t("ดูอย่างเดียว — แก้ไขไม่ได้","View only — cannot edit") : undefined} compact/>
            {locked && <span title={t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only")} style={{fontSize:12}}><Ico name="lock" size={14} color={T.textMuted} /></span>}
          </div>
          {poLastUpdate(p) && <div style={{fontSize:12,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>{t("อัปเดต","Updated")} {relativeTime(poLastUpdate(p).at)}</div>}
        </td>
        <td style={{padding:"9px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
          {onEdit && <button onClick={()=>onEdit(p)} disabled={locked} title={locked?t("แก้ไขได้เฉพาะ Admin","Admin only"):t("แก้ไข","Edit")}
            style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textMuted,cursor:locked?"not-allowed":"pointer",padding:"6px 8px",minWidth:34,minHeight:34,borderRadius:8,display:"inline-grid",placeItems:"center"}} aria-label={t("แก้ไข","Edit")}><Ico name="edit" size={16} /></button>}
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
                  <span style={{color:T.blue,fontSize:13,fontVariantNumeric:"tabular-nums",fontWeight:650}}>{codeText(code)}</span>
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
                              <StatusPicker status={p.status} onChange={s=>onStatusChange?.(p,s)} disabled={locked || readOnly} disabledTitle={readOnly ? t("ดูอย่างเดียว — แก้ไขไม่ได้","View only — cannot edit") : undefined} compact/>
                              {locked && <span title={t("รับของและจ่ายเงินครบแล้ว แก้ไขได้เฉพาะ Admin","Fully received & paid — Admin only")} style={{fontSize:12}}><Ico name="lock" size={14} color={T.textMuted} /></span>}
                            </div>
                            {poLastUpdate(p) && <div style={{fontSize:12,color:T.textMuted,marginTop:3,whiteSpace:"nowrap"}}>{t("อัปเดต","Updated")} {relativeTime(poLastUpdate(p).at)}</div>}
                          </td>
                          <td style={{padding:"9px 16px",whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                            {onEdit && <button onClick={()=>onEdit(p)} disabled={locked} title={locked?t("แก้ไขได้เฉพาะ Admin","Admin only"):t("แก้ไข","Edit")}
                              style={{background:"none",border:"none",color:locked?"#cbd5e1":T.textMuted,cursor:locked?"not-allowed":"pointer",padding:"6px 8px",minWidth:34,minHeight:34,borderRadius:8,display:"inline-grid",placeItems:"center"}} aria-label={t("แก้ไข","Edit")}><Ico name="edit" size={16} /></button>}
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

export { PODetailModal, IncomingPlanTab, ProcurementView, ProcurementTrackingTab };
