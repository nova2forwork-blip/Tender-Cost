import { sg, ss } from "./supabase.js";

// ─── Password hashing ──────────────────────────────────────────────────────
// NOTE: this hashes passwords with SHA-256 before storing them, mainly to
// avoid plaintext passwords sitting in the table. It is NOT strong security:
// the Supabase anon key is bundled into the client app and the `kv_store`
// table's RLS policy allows public read/write, so anyone with the anon key
// can read the hashed values directly via the Supabase REST API. Treat this
// as a lightweight access gate for a small trusted team, not as protection
// against a determined attacker. For real security, migrate to Supabase Auth.
export async function hashPassword(pw) {
  const enc = new TextEncoder().encode(pw);
  const buf = await crypto.subtle.digest("SHA-256", enc);
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

export const ROLES = ["admin", "qs", "procurement", "accounting"];
export const ROLE_LABELS = {
  admin: "ผู้ดูแลระบบ (Admin)",
  qs: "QS",
  procurement: "จัดซื้อ",
  accounting: "บัญชี",
};

const DEFAULT_ACCOUNTS = [
  { username: "admin", name: "ผู้ดูแลระบบ", role: "admin", password: "admin123" },
  { username: "qs", name: "แผนก QS", role: "qs", password: "qs1234" },
  { username: "procurement", name: "แผนกจัดซื้อ", role: "procurement", password: "pr1234" },
  { username: "accounting", name: "แผนกบัญชี", role: "accounting", password: "ac1234" },
];

const uid = () => Math.random().toString(36).slice(2, 10);

// ─── Users ──────────────────────────────────────────────────────────────────
export async function loadUsers() {
  let users = await sg("tcs-users");
  if (!users || users.length === 0) {
    users = await Promise.all(
      DEFAULT_ACCOUNTS.map(async (u) => ({
        id: uid(),
        username: u.username,
        name: u.name,
        role: u.role,
        passwordHash: await hashPassword(u.password),
        active: true,
        createdAt: new Date().toISOString(),
      }))
    );
    await ss("tcs-users", users);
  }
  return users;
}

export const saveUsers = (list) => ss("tcs-users", list);

export async function createUser({ username, name, role, password }) {
  const users = await loadUsers();
  if (users.some((u) => u.username.toLowerCase() === username.trim().toLowerCase())) {
    throw new Error("มี username นี้อยู่แล้ว");
  }
  const user = {
    id: uid(),
    username: username.trim(),
    name: name?.trim() || username.trim(),
    role,
    passwordHash: await hashPassword(password),
    active: true,
    createdAt: new Date().toISOString(),
  };
  const next = [...users, user];
  await saveUsers(next);
  return next;
}

export async function resetPassword(users, id, newPassword) {
  const hash = await hashPassword(newPassword);
  const next = users.map((u) => (u.id === id ? { ...u, passwordHash: hash } : u));
  await saveUsers(next);
  return next;
}

export async function toggleActive(users, id) {
  const next = users.map((u) => (u.id === id ? { ...u, active: !u.active } : u));
  await saveUsers(next);
  return next;
}

export async function deleteUser(users, id) {
  const next = users.filter((u) => u.id !== id);
  await saveUsers(next);
  return next;
}

// ─── Access logs ────────────────────────────────────────────────────────────
export async function loadLogs() {
  return (await sg("tcs-logs")) || [];
}

export async function addLog(entry) {
  const logs = await loadLogs();
  const next = [{ id: uid(), time: new Date().toISOString(), ...entry }, ...logs].slice(0, 300);
  await ss("tcs-logs", next);
  return next;
}

// ─── Login ──────────────────────────────────────────────────────────────────
export async function verifyLogin(username, password) {
  const users = await loadUsers();
  const u = users.find((x) => x.username.toLowerCase() === username.trim().toLowerCase());
  if (!u || !u.active) {
    await addLog({ username: username.trim(), role: u?.role || "-", result: u ? "inactive" : "fail" });
    return null;
  }
  const hash = await hashPassword(password);
  if (hash !== u.passwordHash) {
    await addLog({ username: u.username, role: u.role, result: "fail" });
    return null;
  }
  await addLog({ username: u.username, role: u.role, result: "success" });
  return { id: u.id, username: u.username, name: u.name, role: u.role };
}

// ─── Session (kept in localStorage so a refresh doesn't log the user out) ──
const SESSION_KEY = "tcs-session";
export function getSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}
export function setSession(user) {
  localStorage.setItem(SESSION_KEY, JSON.stringify(user));
}
export function clearSession() {
  localStorage.removeItem(SESSION_KEY);
}
