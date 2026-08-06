import { supabase } from "./supabase.js";

// ════════════════════════════════════════════════════════════════════════════
//  Authentication — ใช้ Supabase Auth ของจริง
//  ─ รหัสผ่านถูกเก็บ+แฮชโดย Supabase (bcrypt) เราไม่แตะ ไม่เห็น ไม่เก็บเอง
//  ─ session เป็น JWT ที่ปลอมไม่ได้ (เซ็นด้วย secret ฝั่งเซิร์ฟเวอร์)
//  ─ role อยู่ในตาราง profiles และถูกบังคับใช้จริงด้วย RLS ฝั่ง DB
//  ─ งาน admin (สร้าง/รีเซ็ต/ลบ user) วิ่งผ่าน Edge Function ที่ใช้ service_role
//    ซึ่งไม่มีวันหลุดมาฝั่ง browser
// ════════════════════════════════════════════════════════════════════════════

export const ROLES = ["admin", "qs", "procurement", "accounting"];
export const ROLE_LABELS = {
  admin: "ผู้ดูแลระบบ (Admin)",
  qs: "QS",
  procurement: "จัดซื้อ",
  accounting: "บัญชี",
};

// username ไม่ใช่อีเมล → map เป็นอีเมลสังเคราะห์ภายใต้โดเมนคงที่
// (ผู้ใช้ยังพิมพ์แค่ "qs", "admin" เหมือนเดิม)
const AUTH_DOMAIN = "tender.local";
const emailOf = (username) => `${String(username).trim().toLowerCase()}@${AUTH_DOMAIN}`;

// ─── Login ──────────────────────────────────────────────────────────────────
export async function verifyLogin(username, password) {
  const { data, error } = await supabase.auth.signInWithPassword({
    email: emailOf(username),
    password,
  });
  // ล็อกอินไม่ผ่าน — Supabase บันทึกความพยายามที่ล้มเหลวไว้ใน Auth Logs ให้แล้ว
  if (error || !data?.user) return null;

  const { data: prof } = await supabase
    .from("profiles")
    .select("username, name, role, active")
    .eq("id", data.user.id)
    .maybeSingle();

  // ไม่มี profile หรือถูกระงับ → ตัด session ทิ้งทันที
  if (!prof || prof.active === false) {
    await supabase.auth.signOut();
    return null;
  }

  const user = { id: data.user.id, username: prof.username, name: prof.name, role: prof.role };
  // บันทึก log การเข้าใช้งานที่สำเร็จ (ตอนนี้เราเป็น authenticated แล้ว จึงเขียนได้)
  await supabase.from("access_logs").insert({
    user_id: user.id, username: user.username, role: user.role, result: "success",
  });
  return user;
}

// ─── Session ─────────────────────────────────────────────────────────────────
// Supabase เป็นคนถือ session ให้ (ใน localStorage แต่เป็น JWT ที่ปลอมไม่ได้)
export async function getSession() {
  const { data } = await supabase.auth.getSession();
  const s = data?.session;
  if (!s) return null;

  const { data: prof } = await supabase
    .from("profiles")
    .select("username, name, role, active")
    .eq("id", s.user.id)
    .maybeSingle();

  if (!prof || prof.active === false) {
    await supabase.auth.signOut();
    return null;
  }
  return { id: s.user.id, username: prof.username, name: prof.name, role: prof.role };
}

// setSession เก็บไว้เพื่อความเข้ากันได้กับ App.jsx — Supabase จัดการ session ให้แล้ว
export function setSession() { /* no-op: session ถูกตั้งโดย signInWithPassword */ }
export async function clearSession() { await supabase.auth.signOut(); }

// ─── Access logs (admin อ่านได้) ─────────────────────────────────────────────
export async function loadLogs() {
  const { data, error } = await supabase
    .from("access_logs")
    .select("id, time, username, role, result")
    .order("time", { ascending: false })
    .limit(300);
  if (error) { console.warn("loadLogs", error); return []; }
  return data || [];
}

// ─── Admin user management (ผ่าน Edge Function เท่านั้น) ─────────────────────
async function callAdmin(payload) {
  const { data, error } = await supabase.functions.invoke("admin-users", { body: payload });
  if (error) {
    let msg = error.message;
    try { const ctx = await error.context?.json?.(); if (ctx?.error) msg = ctx.error; } catch { /* ignore */ }
    throw new Error(msg || "เรียก admin-users ไม่สำเร็จ");
  }
  if (data?.error) throw new Error(data.error);
  return data;
}

export async function loadUsers() {
  const data = await callAdmin({ action: "list" });
  return data.users || [];
}

export async function createUser({ username, name, role, password }) {
  const data = await callAdmin({ action: "create", username, name, role, password });
  return data.users || [];
}

export async function resetPassword(users, id, newPassword) {
  await callAdmin({ action: "reset", id, password: newPassword });
  return users; // ไม่ต้องเปลี่ยนหน้าจอ — รหัสผ่านไม่เคยแสดงอยู่แล้ว
}

export async function toggleActive(_users, id) {
  const data = await callAdmin({ action: "toggle", id });
  return data.users || [];
}

export async function deleteUser(_users, id) {
  const data = await callAdmin({ action: "delete", id });
  return data.users || [];
}
