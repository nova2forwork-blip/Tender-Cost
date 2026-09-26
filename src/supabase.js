import { createClient } from "@supabase/supabase-js";

// ── ค่าเชื่อมต่อดึงจาก Environment Variables เท่านั้น ─────────────────────────
//    Local:  ตั้งใน .env.local
//    Vercel: Project Settings → Environment Variables (แล้ว Redeploy ใหม่!)
//    ห้าม hardcode คีย์ลงในไฟล์นี้เด็ดขาด (มันจะถูก push ขึ้น Git)
const SUPABASE_URL  = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON = import.meta.env.VITE_SUPABASE_ANON_KEY;

// แสดงข้อความเต็มจอแทน "จอขาว" เวลาตั้งค่าไม่ครบ — ผู้ใช้จะรู้ทันทีว่าพลาดตรงไหน
function showFatal(message) {
  if (typeof document === "undefined") return;
  const el = document.createElement("div");
  el.style.cssText =
    "position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;" +
    "background:#0f172a;color:#e2e8f0;font-family:system-ui,'Segoe UI',sans-serif;padding:24px;";
  el.innerHTML =
    '<div style="max-width:520px;background:#1e293b;border:1px solid #334155;border-radius:16px;padding:28px;line-height:1.6">' +
    '<div style="font-size:20px;font-weight:700;margin-bottom:10px">⚙️ ยังตั้งค่าไม่ครบ</div>' +
    '<div style="font-size:14px;color:#cbd5e1;margin-bottom:16px">' + message + "</div>" +
    '<div style="font-size:12px;color:#94a3b8">Vercel → Settings → Environment Variables → เพิ่มค่า → ' +
    "แล้วไป Deployments กด <b>Redeploy</b> (ต้อง redeploy ค่าถึงจะมีผล)</div>" +
    "</div>";
  const mount = () => document.body && document.body.appendChild(el);
  if (document.body) mount();
  else document.addEventListener("DOMContentLoaded", mount);
}

// no-op client กันแอป crash ตอน import เมื่อยังไม่มีคีย์ — ทุก call จะคืนค่า
// ว่าง/error อย่างนุ่มนวล (sg/ss/sd/auth ห่อ try–catch อยู่แล้ว จึงไม่ล้ม)
function makeStub(reason) {
  const err = { message: reason };
  const asyncErr = async () => ({ data: null, error: err });
  const query = () =>
    new Proxy(
      {},
      {
        get(_t, prop) {
          if (prop === "maybeSingle" || prop === "single") return async () => ({ data: null, error: err });
          if (prop === "then") return undefined; // ไม่ทำตัวเป็น thenable เอง
          return () => query(); // select/eq/upsert/insert/update/delete/order/limit → ต่อ chain ได้
        },
      }
    );
  return {
    auth: new Proxy({}, { get: () => async () => ({ data: { session: null, user: null }, error: err }) }),
    functions: { invoke: asyncErr },
    from: () => query(),
    channel: () => ({ on() { return this; }, subscribe() { return { unsubscribe() {} }; } }),
    removeChannel: () => {},
    removeAllChannels: () => {},
  };
}

let supabase;
if (!SUPABASE_URL || !SUPABASE_ANON) {
  const missing = [!SUPABASE_URL && "VITE_SUPABASE_URL", !SUPABASE_ANON && "VITE_SUPABASE_ANON_KEY"]
    .filter(Boolean)
    .join(" และ ");
  console.error("❌ ยังไม่ได้ตั้งค่า Supabase env: " + missing);
  showFatal("ยังไม่ได้ตั้งค่า <b>" + missing + "</b> — ระบบเชื่อมต่อฐานข้อมูลไม่ได้");
  supabase = makeStub("Supabase env ยังไม่ได้ตั้งค่า");
} else {
  supabase = createClient(SUPABASE_URL, SUPABASE_ANON, {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
    },
  });
}

export { supabase };

// ── Storage helpers (ตาราง kv_store) ─────────────────────────────────────────
export const sg = async (key) => {
  try {
    const { data, error } = await supabase
      .from("kv_store")
      .select("value")
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    return data ? JSON.parse(data.value) : null;
  } catch (e) {
    console.warn("sg error", key, e);
    return null;
  }
};

export const ss = async (key, value) => {
  try {
    const { error } = await supabase
      .from("kv_store")
      .upsert({ key, value: JSON.stringify(value), updated_at: new Date().toISOString() });
    if (error) throw error;
  } catch (e) {
    console.warn("ss error (อาจเป็นเพราะสิทธิ์ไม่พอ)", key, e);
  }
};

// ── รุ่น "โยน error" (ไม่กลืน) สำหรับงานที่ต้องรู้ว่าพลาด เช่น ย้ายรหัสบัญชี ─────────
//  sg/ss ปกติจะกลืน error เพื่อไม่ให้ UI ล้ม แต่การ migrate ข้ามหลายคีย์ต้อง
//  "รู้ทันทีเมื่อพลาด" เพื่อหยุดก่อนเขียนทับ ไม่งั้นข้อมูลบางโครงการจะย้ายไม่ครบเงียบ ๆ
export const sgOrThrow = async (key) => {
  const { data, error } = await supabase
    .from("kv_store").select("value").eq("key", key).maybeSingle();
  if (error) throw error;
  return data ? JSON.parse(data.value) : null;
};
// อ่านหลายคีย์ในคำขอเดียว (หน้ารายการโครงการใช้สรุปตัวเลขของทุกโครงการ) — คืน { key: value }
export const sgMany = async (keys) => {
  if (!keys.length) return {};
  const { data, error } = await supabase.from("kv_store").select("key, value").in("key", keys);
  if (error) throw error;
  const out = {};
  for (const r of data || []) { try { out[r.key] = JSON.parse(r.value); } catch { /* ข้ามค่าที่อ่านไม่ได้ */ } }
  return out;
};
export const ssOrThrow = async (key, value) => {
  const { error } = await supabase
    .from("kv_store").upsert({ key, value: JSON.stringify(value), updated_at: new Date().toISOString() });
  if (error) throw error;
};

// ── เขียนแบบกันชนกันหลายคน (optimistic-concurrency + 3-way merge) ────────────
//  ปัญหาเดิม: ทั้ง PO/โครงการ เก็บเป็นก้อน JSON ก้อนเดียวต่อคีย์ → ถ้า 2 คนแก้
//  พร้อมกัน คนบันทึกทีหลังจะทับงานคนแรกหายทั้งก้อน (last-writer-wins).
//  วิธีแก้: ก่อนเขียน อ่านค่าล่าสุดบนเซิร์ฟเวอร์มาก่อน ถ้ามีคนแก้แทรกระหว่างทาง
//  ให้ "รวมการแก้ของเรา" ลงบนของล่าสุด (แทนการทับทั้งก้อน) แล้วเขียนแบบมี guard
//  ด้วย updated_at (ถ้ามีคนเขียนซ้อนอีกก็ลองใหม่). ชนิดข้อมูลที่รวมได้:
//    • array ที่ทุกตัวมี id (PO, โครงการ, extra) → รวมตาม id (รวมกรณีเริ่มจาก [] ด้วย)
//    • array ของสตริง (hidden) → รวมแบบเซต
//    • object ธรรมดา (tenders {code:val}, additions) → รวมตาม field (ลึกลงไปในแต่ละเดือน)
//  ชนิดอื่น → ของเราชนะ (last-writer-wins)
const _sgRaw = async (key) => {
  const { data, error } = await supabase
    .from("kv_store").select("value, updated_at").eq("key", key).maybeSingle();
  if (error) throw error;
  return data || null; // { value:string, updated_at } | null
};
const _isObj  = (v) => v != null && typeof v === "object" && !Array.isArray(v);
const _same   = (a, b) => a === b || JSON.stringify(a) === JSON.stringify(b);
// array ของ object ที่มี id ทุกตัว — นับ [] ว่าเข้าข่ายด้วย (เดิมต้องไม่ว่าง ทำให้
// "โครงการยังไม่มี PO แล้ว 2 เครื่องเพิ่ม PO พร้อมกัน" ตกไปทาง LWW → PO เครื่องแรกหาย)
const _idArr  = (a) => Array.isArray(a) && a.every(x => _isObj(x) && "id" in x);
const _codeArr = (a) => Array.isArray(a) && a.every(x => _isObj(x) && typeof x.code === "string" && !("id" in x)) && new Set(a.map(x => x.code)).size === a.length;
const _strArr = (a) => Array.isArray(a) && a.every(x => typeof x === "string" || typeof x === "number");

// f = ช่องที่ใช้ระบุตัว: "id" (PO/โครงการ/แผน) หรือ "code" (รายการเพิ่ม extra / รหัสบัญชี)
function _mergeById(base, prev, next, f = "id") {
  const nextMap = new Map(next.map(x => [x[f], x]));
  const prevMap = new Map(prev.map(x => [x[f], x]));
  const removed = new Set(prev.filter(x => !nextMap.has(x[f])).map(x => x[f])); // ตัวที่เราลบ
  const out = [], seen = new Set();
  for (const item of base) {
    if (removed.has(item[f])) continue;                          // เราลบ → เอาออก
    seen.add(item[f]);
    const mine = nextMap.get(item[f]);
    // ไม่ได้แตะ → คงของล่าสุดบนเซิร์ฟเวอร์ / เราแก้ → รวมลงไปทีละช่อง (ถ้ามีตั้งต้น) เช่น 2 คนแก้คนละงวด
    // ของ PO ใบเดียวกัน หรือคนละช่องของโครงการเดียวกัน → ได้ครบทั้งคู่; ช่องเดียวกัน → ของเราชนะ
    const base0 = prevMap.get(item[f]);
    out.push(!mine || _same(mine, base0) ? item : base0 ? mergeForWrite(item, base0, mine) : mine);
  }
  // ตัวที่เราเพิ่มใหม่: ถ้าเราเพิ่ม "ไว้หน้า" (เช่นประวัติ PO ล่าสุดอยู่บนสุด) → วางไว้หน้าเหมือนเดิม
  // ไม่งั้นต่อท้าย (เดิมต่อท้ายเสมอ → ประวัติใหม่ไปอยู่ท้ายแล้วโดนตัดทิ้งตอนตัดเหลือ 40)
  const isNew = (x) => !seen.has(x[f]) && !prevMap.has(x[f]);
  let lead = 0; while (lead < next.length && isNew(next[lead])) lead++;
  if (lead === next.length) lead = 0;   // ไม่มีตัวเดิมให้อ้างอิงตำแหน่ง (เช่นเริ่มจากว่าง) → ต่อท้าย
  const front = next.slice(0, lead), tail = next.slice(lead).filter(isNew);
  return [...front, ...out, ...tail];
}
// array ของสตริง (เช่น hidden codes) → รวมแบบเซต: ของล่าสุด − ที่เราเอาออก + ที่เราเพิ่ม
function _mergeSet(base, prev, next) {
  const P = new Set(prev), N = new Set(next);
  const out = base.filter(x => !(P.has(x) && !N.has(x)));
  for (const x of next) if (!P.has(x) && !out.includes(x)) out.push(x);
  return out;
}
// array ของ object ที่ไม่มี id แต่ของใหม่ถูก "ต่อหน้า" (เช่นประวัติการแก้ PO ล่าสุดอยู่บน)
// → เอารายการที่เราเพิ่มวางหน้า ของล่าสุดบนเซิร์ฟเวอร์ต่อท้าย (ประวัติของทั้ง 2 เครื่องไม่หาย)
function _mergeLog(base, prev, next) {
  const key = (x) => JSON.stringify(x);
  const P = new Set(prev.map(key));
  const added = next.filter(x => !P.has(key(x)));
  if (!added.length) return base;                                   // เราไม่ได้เพิ่มอะไร → ของล่าสุด
  if (!added.every((x, i) => key(x) === key(next[i]))) return next;   // ไม่ใช่แบบต่อหน้า → ของเราชนะ
  const A = new Set(added.map(key));
  return [...added, ...base.filter(x => !A.has(key(x)))];   // ความยาวถูกตัดเหลือ 40 ในแอปตอนบันทึกครั้งถัดไป
}
const _objArr = (a) => Array.isArray(a) && a.every(x => _isObj(x) && !("id" in x) && !("code" in x));   // แบบ "บันทึกประวัติ" (ไม่ใช่รายการที่มีรหัส)
// object → รวมตาม field; ถ้า field เป็น object ทั้ง 3 ฝั่ง (เช่นเดือนใน additions) รวมลึกลงไปอีกชั้น
// เพื่อให้ 2 คนแก้คนละ Account Code ในเดือนเดียวกันไม่ทับกัน
function _mergeByKey(base, prev, next) {
  const out = { ...base };
  for (const k of Object.keys(prev)) if (!(k in next)) delete out[k];          // field ที่เราลบ
  for (const k of Object.keys(next)) {
    if (k in prev && _same(next[k], prev[k])) continue;                          // ไม่ได้แตะ → คงของล่าสุด
    if (k in prev && k in base) { out[k] = mergeForWrite(base[k], prev[k], next[k]); continue; }
    // ทั้ง 2 เครื่องเพิ่ม field ใหม่ชื่อเดียวกันพร้อมกัน (เช่นเดือนใหม่ / $columns ชุดแรก) → รวมจาก "ว่าง"
    // (เดิมของเราทับของอีกเครื่องทั้งก้อน)
    if (!(k in prev) && k in base && base[k] && next[k] && typeof base[k] === "object" && Array.isArray(base[k]) === Array.isArray(next[k])) {
      out[k] = mergeForWrite(base[k], Array.isArray(next[k]) ? [] : {}, next[k]); continue;
    }
    out[k] = next[k];
  }
  return out;
}
// รวม "การแก้ของเรา (prev→next)" ลงบนค่าล่าสุดบนเซิร์ฟเวอร์ (server) — ฟังก์ชันล้วน ทดสอบได้
export function mergeForWrite(server, prev, next) {
  if (server == null || _same(server, prev)) return next;       // ไม่มีใครแก้แทรก → เขียนของเราได้เลย
  if (prev == null) return next;                                // ไม่มีจุดตั้งต้น → รวมไม่ได้ (LWW)
  if (_idArr(server) && _idArr(prev) && _idArr(next))  return _mergeById(server, prev, next);
  if (_codeArr(server) && _codeArr(prev) && _codeArr(next)) return _mergeById(server, prev, next, "code");
  if (_strArr(server) && _strArr(prev) && _strArr(next)) return _mergeSet(server, prev, next);
  if (_isObj(server) && _isObj(prev) && _isObj(next))  return _mergeByKey(server, prev, next);
  if (_objArr(server) && _objArr(prev) && _objArr(next)) return _mergeLog(server, prev, next);
  return next;                                                  // ชนิดไม่รู้จัก / ค่าเดี่ยว → ของเราชนะ
}

// opts.normalize(value) — ใช้เฉพาะตอน "มีการรวม" เกิดขึ้น เพื่อคำนวณค่าที่ได้จากค่าอื่นใหม่ (เช่นยอดแม่ = ผลรวมลูก)
// คืนค่าที่เขียนลงจริง — ให้หน้าจอแสดงงานของเครื่องอื่นที่ถูกรวมเข้ามาด้วย
export const ssMerge = async (key, prev, next, opts = {}) => {
  const tries = opts._tries || 0;
  const again = () => ssMerge(key, prev, next, { ...opts, _tries: tries + 1 });
  const combine = (row) => {
    if (!row) return next;
    let server = null;
    try { server = JSON.parse(row.value); } catch { server = null; }
    const m = mergeForWrite(server, prev, next);   // มีคนอื่นแก้ระหว่างทาง → รวมแทนการทับ
    return m !== next && opts.normalize ? opts.normalize(m) : m;
  };
  const cur = await _sgRaw(key);
  const toWrite = combine(cur);
  const payload = { key, value: JSON.stringify(toWrite), updated_at: new Date().toISOString() };
  if (cur) {
    // อัปเดตเฉพาะถ้า updated_at ยังไม่เปลี่ยนจากที่เพิ่งอ่าน (กันเขียนซ้อน)
    let q = supabase.from("kv_store").update(payload).eq("key", key);
    if (cur.updated_at) q = q.eq("updated_at", cur.updated_at);
    const { data, error } = await q.select("key");
    if (error) throw error;
    if (!data || data.length === 0) {
      // อัปเดตได้ 0 แถว มี 3 สาเหตุ: (ก) มีคนเขียนแทรก (updated_at เปลี่ยน) → ลองใหม่
      // (ข) RLS ไม่ให้สิทธิ์ — DB ไม่ส่ง error แต่คืน 0 แถวเงียบ ๆ → แจ้งว่าไม่มีสิทธิ์ทันที
      // (ค) แถวถูกลบไประหว่างทาง (เช่นแอดมินลบโครงการ) → ไม่สร้างกลับขึ้นมาใหม่
      const chk = await _sgRaw(key);
      if (!chk) {
        const e = new Error("ข้อมูลนี้ถูกลบไปแล้ว (row deleted: " + key + ")");
        e.code = "GONE";
        throw e;
      }
      if (!cur.updated_at || chk.updated_at === cur.updated_at) {
        const e = new Error("ไม่มีสิทธิ์บันทึกข้อมูลนี้ (permission denied: " + key + ")");
        e.code = "42501";
        throw e;
      }
      if (tries < 4) return again(); // มีคนเขียนแทรก → ลองใหม่
      // ยอมแพ้เรื่อง guard updated_at แต่ "ยังรวมไม่ทับ" — อ่านค่าล่าสุดอีกรอบแล้ว
      // merge การแก้ของเรา (prev→next) ลงบนของล่าสุด ก่อนเขียน กันงานคนอื่นหายเงียบ ๆ
      const merged = combine(await _sgRaw(key));
      const { error: e2 } = await supabase.from("kv_store")
        .upsert({ key, value: JSON.stringify(merged), updated_at: new Date().toISOString() });
      if (e2) throw e2;
      return merged;
    }
    return toWrite;
  }
  const { error } = await supabase.from("kv_store").insert(payload);
  if (error) {
    if (tries < 4) return again(); // อาจมีคนเพิ่งสร้าง → ลองใหม่
    throw error;
  }
  return toWrite;
};

export const sd = async (key) => {
  try {
    await sdOrThrow(key);
  } catch (e) {
    console.warn("sd error", key, e);
  }
};
// รุ่นโยน error — ใช้ตอนลบโครงการ เพื่อแจ้งผู้ใช้ได้ว่าลบข้อมูลย่อยไม่ครบ
export const sdOrThrow = async (key) => {
  const { error } = await supabase.from("kv_store").delete().eq("key", key);
  if (error) throw error;
};

// ── กู้คืนข้อมูล (ใช้ตาราง kv_history จาก kv-history.sql) ──────────────────────
// อ่านประวัติได้เฉพาะ admin (บังคับด้วย RLS ฝั่ง DB) และการเขียนคืนก็ต้องเป็น
// admin เท่านั้น (can_write_key อนุญาตทุกคีย์เฉพาะ role=admin)
export const loadKvHistory = async (key) => {
  try {
    let q = supabase
      .from("kv_history")
      .select("id,key,op,changed_by,changed_at,old_value")
      .order("changed_at", { ascending: false })
      .limit(400);
    if (key) q = q.eq("key", key);
    const { data, error } = await q;
    if (error) throw error;
    return data || [];
  } catch (e) {
    console.warn("loadKvHistory error", e);
    return [];
  }
};

// กู้ค่าจากประวัติแถวหนึ่ง กลับเข้า kv_store (เขียนค่าดิบกลับตรง ๆ ไม่ stringify ซ้ำ)
export const restoreKvVersion = async (row) => {
  if (!row || row.old_value == null) throw new Error("ไม่มีค่าเดิมให้กู้คืน");
  const { error } = await supabase
    .from("kv_store")
    .upsert({ key: row.key, value: row.old_value, updated_at: new Date().toISOString() });
  if (error) throw error;
};

// ── สแนปช็อตตามเวลา (จาก kv-snapshots.sql: 12:00 / 18:00 เก็บ 7 วัน) ──────────
export const loadKvSnapshots = async () => {
  try {
    const { data, error } = await supabase
      .from("kv_snapshots")
      .select("id,taken_at,slot,key,value")
      .order("taken_at", { ascending: false })
      .limit(3000);
    if (error) throw error;
    return data || [];
  } catch (e) {
    console.warn("loadKvSnapshots error", e);
    return [];
  }
};

// กู้ค่าจากสแนปช็อตแถวหนึ่งกลับเข้า kv_store
export const restoreKvSnapshot = async (row) => {
  if (!row || row.value == null) throw new Error("ไม่มีข้อมูลให้กู้คืน");
  const { error } = await supabase
    .from("kv_store")
    .upsert({ key: row.key, value: row.value, updated_at: new Date().toISOString() });
  if (error) throw error;
};
