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
