// Tender Cost — สร้างไฟล์ Excel ทุกแผนก (ออกแบบเป็นภาษาไทยเสมอ)
import * as XLSX from "xlsx-js-style";
import { ACCOUNTS, GROUPS, OTHER_COL_LABEL, PAYMENT_LABEL, WASTE_LBL, WASTE_RATE, accountOf, buildCombinedBudget, exportAccountList, fmt, hiddenSafeForPO, itemSupplierName, monthAddValue, monthRowBreakdown, monthShortLabel, paymentStatus, poAmountForCode, poItems, poNextDueDate, poNumbersLabel, poPaidAmount, poPayLines, poRounds, poTotal, roundPaid, roundReceived, todayStr, wasteOf, withWaste } from "./core.jsx";

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
// รหัสที่เพิ่มเอง (EX-xxxx) เป็นรหัสภายในของระบบ → แสดง "(เพิ่มเอง)" แทน · ชื่ออยู่คอลัมน์ถัดไปอยู่แล้ว
// (รหัส EX ที่ไม่อยู่ในรายการแล้ว = รหัสเดียวที่ระบุตัวได้ → คงไว้)
const xCode = (code) => /^EX-/.test(String(code || "")) && accountOf(code) ? "(เพิ่มเอง)" : (code || "");
// สถานะของงวดจ่ายใน Excel — ส่วนที่ยอดรับของยังไม่ครบ (short) บอกตรง ๆ แทน "รอจ่ายเงิน"
const payLineLabel = (l) => l.short ? "รอรับของ (ยอดรับยังไม่ครบ)" : PAYMENT_LABEL[l.status];
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
    rows1.push([xCode(a.code), a.name, a.group, baseline, waste, added, total, ...(U?[toUsd(total,rate)]:[])]);
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
    rows2.push([xCode(a.code), a.name, baseline, waste, ...monthVals, total, ...(U?[toUsd(total,rate)]:[])]);
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
      rows.push([xCode(a.code), a.name, a.group, ...vals, rowTotal, ...(U?[toUsd(rowTotal,rate)]:[])]);
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
    rows.push([xCode(a.code), a.name, a.group, baseline, waste, ...vals, monthTot, cum, ...(U?[toUsd(cum,rate)]:[])]);
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
  const mMonths = [...new Set(Object.values(mCell).flatMap(c=>Object.keys(c)))].filter(Boolean).sort();
  const cellOf = (code,mk) => mCell[code]?.[mk] || null;
  const cellTot = (c) => c ? (c.paid+c.recv+c.po+c.plan) : 0;
  const budgetOf    = (code) => parseFloat(combinedB[code])||0;
  const committedOf = (code) => poEntries.reduce((s,p)=>s+poAmountForCode(p,code),0);
  const plannedOf   = (code) => plansArr.reduce((s,pl)=>s+poAmountForCode(pl,code),0);
  const stockOf     = (code) => poEntries.reduce((s,p)=>s+poItems(p).filter(it=>it.code===code).reduce((ss,it)=>ss+(parseFloat(it.store)||0),0),0);
  const takeoffOf = (code) => [...poEntries, ...plansArr].reduce((s,p)=>s+poItems(p).filter(it=>it.code===code).reduce((ss,it)=>ss+(parseFloat(it.takeoff)||0),0),0);
  // ทุกรหัสที่มีงบ (รวมงบติดลบ) / PO / Stock / แผน หรือมีของเข้า — เดิมใส่เฉพาะรหัสที่มีของเข้า รหัสที่มีงบแต่ยังไม่มี PO
  // หรือ PO ที่ยังไม่มียอดแผน/ยอดรับ จึงหาย และยอดรวม Tender Cost / Issue PO / Balance Cost ไม่ตรงหน้าสรุป
  const mCodes = [...new Set([...acctList.map(a => a.code).filter(c => Math.abs(budgetOf(c)) >= 0.005 || committedOf(c) || stockOf(c) || plannedOf(c)), ...Object.keys(mCell)])].sort();
  if (!mCodes.length || !mMonths.length) return;
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
  // คอลัมน์ต้นทุน: Tender Cost · Take off · % Take off · Stock · Issue PO · Pending PO + เดือน (1 ช่อง/เดือน) + TOTAL + Balance Cost
  const header = ["Acc. Code","Acc. Name",`Tender Cost (รวมเผื่อ ${WASTE_LBL})`,"Take off","% Take off","Stock","Issue PO","Pending PO"];
  const tkPct = (tk, bud) => bud > 0 && tk ? tk / bud : "-";   // สัดส่วน (Excel แสดงเป็น %) · ไม่มีงบ/ไม่มี Take off = "-"
  mMonths.forEach(mk => header.push(monthShortLabel(mk)));
  header.push("TOTAL","Balance Cost");
  const rows = [
    [`ของเข้ารายเดือน (แผน + PO จริง) — ${project.name}`],
    [`เดือนละ 1 ช่อง (มีป้ายกำกับ) — จ่าย=จ่ายแล้ว(เขียว) · รับ=รับของแล้ว · รอเข้า=PO ยังไม่รับ(⚠=ล่าช้า) · แผน=ยังไม่เป็น PO(แดง, มี *) · % Take off = Take off ÷ Tender Cost · Issue PO = PO ที่ยื่นจริง · Pending PO = งบ−Stock−PO−แผน · Balance Cost = Tender Cost−Stock−Issue PO · Export: ${new Date().toLocaleDateString("th-TH")}`],
    [],
    header,
  ];
  const dataStart = rows.length, monthColStart = 8, totalCol = 8 + mMonths.length, balPOcol = totalCol + 1, numCols = balPOcol + 1;
  const lineCount = {}; // จำนวนบรรทัดสูงสุดต่อแถว → ใช้ตั้งความสูงแถว
  mCodes.forEach((code, ri) => {
    const budget=budgetOf(code), committed=committedOf(code), stock=stockOf(code), planned=plannedOf(code), takeoff=takeoffOf(code);
    const row = [xCode(code), nameOf(code), budget, takeoff, tkPct(takeoff, budget), stock, committed, budget-stock-committed-planned];
    let maxLines = 1;
    mMonths.forEach(mk => { const l=cellLines(cellOf(code,mk)); maxLines=Math.max(maxLines, l.length||1); row.push(l.length?l.join("\n"):"-"); });
    row.push(mMonths.reduce((s,mk)=>s+cellTot(cellOf(code,mk)),0), budget-stock-committed); // Balance Cost = Tender Cost − Stock − Issue PO
    rows.push(row);
    lineCount[dataStart+ri] = maxLines;
  });
  const dataEnd = rows.length - 1;
  const sumOf = (fn) => mCodes.reduce((s,c)=>s+fn(c),0);
  const totalArr = ["","TOTAL", sumOf(budgetOf), sumOf(takeoffOf), tkPct(sumOf(takeoffOf), sumOf(budgetOf)), sumOf(stockOf), sumOf(committedOf), sumOf(c=>budgetOf(c)-stockOf(c)-committedOf(c)-plannedOf(c))];
  // แถว TOTAL: รวมยอดทั้งเดือนเป็นตัวเลขเดียว (ไม่แยกจ่าย/รอเข้า/แผน)
  mMonths.forEach(mk => { const colT = sumOf(c=>cellTot(cellOf(c,mk))); totalArr.push(colT>0 ? fmt(colT) : "-"); });
  totalArr.push(mCodes.reduce((s,c)=> s + mMonths.reduce((ss,mk)=>ss+cellTot(cellOf(c,mk)),0), 0), sumOf(c=>budgetOf(c)-stockOf(c)-committedOf(c))); // Balance Cost = Tender − Stock − Issue PO
  rows.push(totalArr);
  const totalRow = rows.length - 1;
  lineCount[totalRow] = 1;
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:12},{wch:34},{wch:16},{wch:15},{wch:11},{wch:14},{wch:15},{wch:16}, ...mMonths.map(()=>({wch:17})), {wch:18},{wch:16}];
  const moneyCols = [2,3,5,6,7, totalCol, balPOcol]; // ช่องเดือนเป็นข้อความ ไม่ใช่ตัวเลข
  styleSheet(ws, { numCols, subRows:[1], headerRow:3, dataStart, dataEnd, totalRow, moneyCols, pctCols:[4], theme });
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
    { const ref=XLSX.utils.encode_cell({r,c:4}); if (ws[ref] && typeof ws[ref].v==="number" && ws[ref].v>1) ws[ref].s = { ...(ws[ref].s||{}), font:{ ...(ws[ref].s?.font||{}), color:{rgb:"DC2626"}, bold:true } }; }   // Take off เกิน Tender
    [7, balPOcol].forEach(cc => { const ref=XLSX.utils.encode_cell({r,c:cc}); if (ws[ref] && typeof ws[ref].v==="number" && ws[ref].v<0) ws[ref].s = { ...(ws[ref].s||{}), font:{ ...(ws[ref].s?.font||{}), color:{rgb:"DC2626"}, bold:true } }; });
  }
  // แถว TOTAL: ช่องเดือนเป็นยอดรวมเดียว — ตัวหนา ชิดขวา
  mMonths.forEach((mk,mi)=>{ const ref=XLSX.utils.encode_cell({r:totalRow,c:monthColStart+mi}); if(ws[ref]) ws[ref].s={ ...(ws[ref].s||{}), font:{ ...((ws[ref].s||{}).font||{}), bold:true }, alignment:{ ...((ws[ref].s||{}).alignment||{}), horizontal:"right", vertical:"center" } }; });
  xBackLink(ws, 2, numCols-1, backSheet);
  XLSX.utils.book_append_sheet(wb, ws, "ของเข้ารายเดือน");
}

// ─── ชีต "ตารางรวมเดือน" (แบบหน้าบัญชี) — mirror AccountingMatrixTab ──────────
//  ต้นทุน (Tender Cost / Stock / Issue PO / Balance PO = Tender − Stock − Issue PO)
//  + กลุ่มเดือนของเข้า (รับครบ=เขียว · รับบางส่วน=เหลือง · แผน/PO รอเข้า=แดง) + แผนจ่ายรายเดือน
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
    return { a, budget, committed, stock, planned, balPO:budget-stock-committed,
      mgRow, pyRow, mgTot:mgRow.reduce((s,c)=>s+c.eff,0), pyTot:pyRow.reduce((s,x)=>s+x,0) };
  }).filter(r => r.budget||r.committed||r.stock||r.mgTot||r.pyTot);
  if (!rowsData.length) return;
  const header = ["Acc. Code","Acc. Name",`Tender Cost (รวมเผื่อ ${WASTE_LBL})`,"Stock","Issue PO","Balance PO",
    ...mgM.map(mk=>`${monthShortLabel(mk)} (เข้า)`), "รวมเข้า",
    ...payM.map(mk=>`${monthShortLabel(mk)} (จ่าย)`), "รวมจ่าย"];
  const rows = [
    [`ตารางรวมเดือน — ${project.name}`],
    [`Balance PO = Tender Cost − Stock − Issue PO (ติดลบ = เกินงบ) · ของเข้า: รับครบ=เขียว · รับบางส่วน=เหลือง · แผน/PO รอเข้า=แดง · Export: ${new Date().toLocaleDateString("th-TH")}`],
    [],
    header,
  ];
  const dataStart = rows.length;
  const mgStart = 6, mgTotCol = 6+mgM.length, payStart = mgTotCol+1, payTotCol = payStart+payM.length, numCols = payTotCol+1;
  const blank = (v) => v ? v : "-";
  rowsData.forEach(r => {
    rows.push([xCode(r.a.code), r.a.name, blank(r.budget), blank(r.stock), blank(r.committed), blank(r.balPO),
      ...r.mgRow.map(c=>blank(c.eff)), blank(r.mgTot),
      ...r.pyRow.map(v=>blank(v)), blank(r.pyTot)]);
  });
  const dataEnd = rows.length - 1;
  const sumOf = (fn) => rowsData.reduce((s,r)=>s+fn(r),0);
  const totalArr = ["","TOTAL", sumOf(r=>r.budget), sumOf(r=>r.stock), sumOf(r=>r.committed), sumOf(r=>r.balPO)];
  mgM.forEach((_,i)=>totalArr.push(sumOf(r=>r.mgRow[i]?.eff||0))); totalArr.push(sumOf(r=>r.mgTot));
  payM.forEach((_,i)=>totalArr.push(sumOf(r=>r.pyRow[i]||0)));    totalArr.push(sumOf(r=>r.pyTot));
  rows.push(totalArr.map((v,i)=> i<2 ? v : blank(v)));
  const totalRow = rows.length - 1;
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{wch:12},{wch:32},{wch:16},{wch:14},{wch:16},{wch:16},
    ...mgM.map(()=>({wch:14})), {wch:14}, ...payM.map(()=>({wch:14})), {wch:14}];
  const allMoney = [2,3,4,5, ...Array.from({length:numCols-6},(_,i)=>6+i)];
  styleSheet(ws, { numCols, subRows:[1], headerRow:3, dataStart, dataEnd, totalRow, moneyCols:allMoney, theme });
  const setColor = (r, col, rgb, bold) => { const ref=XLSX.utils.encode_cell({r,c:col}); if (ws[ref]) ws[ref].s = { ...(ws[ref].s||{}), font:{ ...((ws[ref].s||{}).font||{}), color:{rgb}, ...(bold?{bold:true}:{}) } }; };
  for (let r=dataStart; r<=dataEnd; r++) {
    const rd = rowsData[r-dataStart];
    rd.mgRow.forEach((c,i)=>{ if (c.eff) setColor(r, mgStart+i, c.real ? "059669" : (c.hasRecv ? "D97706" : "EF4444"), c.real || c.hasRecv); }); // เขียว=รับครบ · เหลือง=รับบางส่วน · แดง=ยังไม่เข้า (ให้ตรงกับสีในแอป)
    [5].forEach(cc => { const ref=XLSX.utils.encode_cell({r,c:cc}); if (ws[ref] && typeof ws[ref].v==="number" && ws[ref].v<0) setColor(r, cc, "DC2626", true); });
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
      rows1.push([p.date, xCode(it.code), acc?.name||"", itemSupplierName(p), poNumbersLabel(p), amount, ...(U?[toUsd(amount,rate)]:[]), p.status, deliveryStr, poNextDueDate(pItem)||"-", PAYMENT_LABEL[pay], p.notes||""]);
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
            rows.push([xCode(it.code), acc?.name||"", acc?.group||"-", itemSupplierName(p), poNumbersLabel(p), p.date, amount, ...(U?[toUsd(amount,rate)]:[]), p.status, PAYMENT_LABEL[pay]]);
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
    const pctUsed  = budget>0 ? committed/budget : "—";   // ไม่มีงบ → ไม่มี % (เดิมใส่ 9.99 = 999% ดูเหมือนเกินงบ 10 เท่า)
    const status   = committed>0 && budget<=0 ? "ไม่มีงบ" : committed>budget && budget>0 ? "เกินงบ" : committed>0 ? "OK" : budget>0 ? "ยังไม่ PO" : "-";
    rows1.push(U
      ? [xCode(a.code), a.name, a.group, budget, toUsd(budget,rate), committed, toUsd(committed,rate), variance, pctUsed, status]
      : [xCode(a.code), a.name, a.group, budget, committed, variance, pctUsed, status]);
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
      rows2.push([p.date, xCode(it.code), acc?.name||"", acc?.group||"", itemSupplierName(p), poNumbersLabel(p), amount, ...(U?[toUsd(amount,rate)]:[]), p.status, deliveryStr, poNextDueDate(pItem)||"-", PAYMENT_LABEL[pay]]);
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
        ? [label, l.payDate||"-", l.supplier, l.poNo, xCode(l.code), l.accName, l.method, incomingTxt, l.amount, toUsd(l.amount,rate), payLineLabel(l)]
        : [label, l.payDate||"-", l.supplier, l.poNo, xCode(l.code), l.accName, l.method, incomingTxt, l.amount, payLineLabel(l)]);
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

export { BORDER_THIN, STATUS_PILL, statusPill, lighten, styleSheet, exportRate, toUsd, xLinkCell, xLinkRow, xBackLink, addDashboardSheet, exportQSExcel, exportQSMonthExcel, addIncomingMonthlySheet, addAccountingMatrixSheet, exportProcurementExcel, exportAccountingExcel, buildTSV };
