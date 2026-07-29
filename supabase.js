import { createClient } from "@supabase/supabase-js";

// ── ใส่ค่าจาก Supabase Project Settings → API ────────────────────────────────
const SUPABASE_URL  = "https://rfvirqdqzzdjcbexewzl.supabase.co";
const SUPABASE_ANON = "sb_publishable_Il_2lYsiHsDqeNho5CNkcg_zpQMVtuz";

if (!SUPABASE_URL || !SUPABASE_ANON) {
  console.error(
    "❌  ยังไม่ได้ตั้งค่า Supabase!\n" +
    "    สร้างไฟล์ .env.local แล้วใส่:\n" +
    "    VITE_SUPABASE_URL=...\n" +
    "    VITE_SUPABASE_ANON_KEY=..."
  );
}

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);

// ── Storage helpers (แทน window.storage) ──────────────────────────────────────
// ใช้ตาราง kv_store (key TEXT primary key, value TEXT, updated_at TIMESTAMPTZ)

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
    console.warn("ss error", key, e);
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
