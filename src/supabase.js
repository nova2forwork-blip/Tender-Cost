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
//    • array ที่ทุกตัวมี id (PO, โครงการ, extra) → รวมตาม id
//    • object ธรรมดา (tenders {code:val}, additions) → รวมตาม field
//  ชนิดอื่น (เช่น array ของสตริง hidden) → เขียนทับตามเดิม (ไม่แย่ลงกว่าเก่า)
const _sgRaw = async (key) => {
  const { data, error } = await supabase
    .from("kv_store").select("value, updated_at").eq("key", key).maybeSingle();
  if (error) throw error;
  return data || null; // { value:string, updated_at } | null
};
const _isObj  = (v) => v && typeof v === "object" && !Array.isArray(v);
const _hasIds = (a) => Array.isArray(a) && a.length > 0 && a.every(x => x && typeof x === "object" && "id" in x);

function _mergeById(base, prev, next) {
  const nextMap = new Map(next.map(x => [x.id, x]));
  const removed = new Set(prev.filter(x => !nextMap.has(x.id)).map(x => x.id)); // ตัวที่เราลบ
  const out = [], seen = new Set();
  for (const item of base) {
    if (removed.has(item.id)) continue;                 // เราลบ → เอาออก
    out.push(nextMap.has(item.id) ? nextMap.get(item.id) : item); // เราแก้→ของเรา / ไม่แตะ→คงไว้
    seen.add(item.id);
  }
  for (const item of next) if (!seen.has(item.id)) out.push(item); // ตัวที่เราเพิ่มใหม่
  return out;
}
function _mergeByKey(base, prev, next) {
  const out = { ...base };
  for (const k of Object.keys(prev)) if (!(k in next)) delete out[k];        // field ที่เราลบ
  for (const k of Object.keys(next)) if (!(k in prev) || next[k] !== prev[k]) out[k] = next[k]; // เพิ่ม/แก้
  return out;
}

export const ssMerge = async (key, prev, next, _tries = 0) => {
  const cur = await _sgRaw(key);
  let toWrite = next;
  if (cur) {
    let server = null;
    try { server = JSON.parse(cur.value); } catch { server = null; }
    if (server != null && JSON.stringify(server) !== JSON.stringify(prev)) {
      // มีคนอื่นแก้ระหว่างที่เรากำลังแก้ → รวมแทนการทับ
      if (_hasIds(server) && _hasIds(prev) && _hasIds(next))       toWrite = _mergeById(server, prev, next);
      else if (_isObj(server) && _isObj(prev) && _isObj(next))     toWrite = _mergeByKey(server, prev, next);
      else                                                          toWrite = next; // ชนิดไม่รู้จัก → LWW
    }
  }
  const payload = { key, value: JSON.stringify(toWrite), updated_at: new Date().toISOString() };
  if (cur) {
    // อัปเดตเฉพาะถ้า updated_at ยังไม่เปลี่ยนจากที่เพิ่งอ่าน (กันเขียนซ้อน)
    let q = supabase.from("kv_store").update(payload).eq("key", key);
    if (cur.updated_at) q = q.eq("updated_at", cur.updated_at);
    const { data, error } = await q.select("key");
    if (error) throw error;
    if (!data || data.length === 0) {
      if (_tries < 4) return ssMerge(key, prev, next, _tries + 1); // มีคนเขียนแทรก → ลองใหม่
      // ยอมแพ้เรื่อง guard updated_at แต่ "ยังรวมไม่ทับ" — อ่านค่าล่าสุดอีกรอบแล้ว
      // merge การแก้ของเรา (prev→next) ลงบนของล่าสุด ก่อนเขียน กันงานคนอื่นหายเงียบ ๆ
      const fresh = await _sgRaw(key);
      let merged = next;
      if (fresh) {
        let server = null; try { server = JSON.parse(fresh.value); } catch { server = null; }
        if (server != null && JSON.stringify(server) !== JSON.stringify(prev)) {
          if (_hasIds(server) && _hasIds(prev) && _hasIds(next))       merged = _mergeById(server, prev, next);
          else if (_isObj(server) && _isObj(prev) && _isObj(next))     merged = _mergeByKey(server, prev, next);
          else                                                          merged = next; // ชนิดไม่รู้จัก → LWW
        }
      }
      const { error: e2 } = await supabase.from("kv_store")
        .upsert({ key, value: JSON.stringify(merged), updated_at: new Date().toISOString() });
      if (e2) throw e2;
    }
  } else {
    const { error } = await supabase.from("kv_store").insert(payload);
    if (error) {
      if (_tries < 4) return ssMerge(key, prev, next, _tries + 1); // อาจมีคนเพิ่งสร้าง → ลองใหม่
      throw error;
    }
  }
};

export const sd = async (key) => {
  try {
    const { error } = await supabase.from("kv_store").delete().eq("key", key);
    if (error) throw error;
  } catch (e) {
    console.warn("sd error", key, e);
  }
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
