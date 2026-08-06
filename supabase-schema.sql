-- ════════════════════════════════════════════════════════════════════════════
--  Tender Cost System — Secure Schema (Supabase Auth + Role-based RLS)
--  วิธีใช้: Supabase Dashboard → SQL Editor → New Query → วางทั้งหมด → Run
--  รันซ้ำได้ (idempotent) — ปลอดภัยถ้าเผลอ Run สองรอบ
-- ════════════════════════════════════════════════════════════════════════════

-- ── 0. ล้าง policy เก่าที่เปิดโล่ง (allow_all) ทิ้งก่อน ─────────────────────
drop policy if exists "allow_all" on public.kv_store;

-- ── 1. ตาราง kv_store (เก็บข้อมูลแอปเหมือนเดิม) ──────────────────────────────
create table if not exists public.kv_store (
  key        text        primary key,
  value      text        not null,
  updated_at timestamptz default now()
);
alter table public.kv_store enable row level security;

-- ── 2. ตาราง profiles (แหล่ง "role ของจริง" ผูกกับ auth.users) ───────────────
create table if not exists public.profiles (
  id         uuid        primary key references auth.users(id) on delete cascade,
  username   text        unique not null,
  name       text        not null,
  role       text        not null check (role in ('admin','qs','procurement','accounting')),
  active     boolean     not null default true,
  created_at timestamptz default now()
);
alter table public.profiles enable row level security;

-- ── 3. ตาราง access_logs (ประวัติเข้าใช้งาน — admin อ่าน) ────────────────────
create table if not exists public.access_logs (
  id       bigint generated always as identity primary key,
  user_id  uuid,
  username text,
  role     text,
  result   text,
  time     timestamptz default now()
);
alter table public.access_logs enable row level security;

-- ── 4. Helper: role ของผู้เรียกปัจจุบัน (null ถ้าไม่มี profile หรือถูกระงับ) ──
--    SECURITY DEFINER = อ่าน profiles ข้าม RLS ได้ → กันปัญหา recursion
create or replace function public.auth_role()
returns text
language sql
stable
security definer
set search_path = public
as $$
  select role from public.profiles where id = auth.uid() and active = true
$$;

-- ── 5. Helper: ผู้เรียกเขียนคีย์นี้ได้ไหม (บังคับ role ตามชนิดข้อมูล) ─────────
create or replace function public.can_write_key(k text)
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select case
    when public.auth_role() = 'admin' then true
    -- โครงการ: ทุก role ที่ล็อกอินสร้าง/แก้ได้
    when k = 'tcs-projects'
         then public.auth_role() in ('qs','procurement','accounting','admin')
    -- งาน QS: tender cost / additions / extra items / hidden accounts
    when k like 'tcs-tenders-%' or k like 'tcs-additions-%'
      or k like 'tcs-extra-%'   or k like 'tcs-hidden-%'
         then public.auth_role() in ('qs','admin')
    -- งานจัดซื้อ: PO
    when k like 'tcs-po-%'
         then public.auth_role() in ('procurement','admin')
    -- คีย์อื่น ๆ: admin เท่านั้น  (บัญชี = อ่านอย่างเดียว โดยไม่มีเงื่อนไขให้เขียน)
    else false
  end
$$;

-- ── 6. RLS policies: kv_store ────────────────────────────────────────────────
--    อ่าน = ผู้ใช้ที่ล็อกอินและมี role เท่านั้น (anon เข้าไม่ได้เลย)
drop policy if exists kv_select on public.kv_store;
create policy kv_select on public.kv_store
  for select to authenticated
  using (public.auth_role() is not null);

--    เขียน = ตามสิทธิ์ role ต่อชนิดคีย์
drop policy if exists kv_insert on public.kv_store;
create policy kv_insert on public.kv_store
  for insert to authenticated
  with check (public.can_write_key(key));

drop policy if exists kv_update on public.kv_store;
create policy kv_update on public.kv_store
  for update to authenticated
  using (public.can_write_key(key))
  with check (public.can_write_key(key));

drop policy if exists kv_delete on public.kv_store;
create policy kv_delete on public.kv_store
  for delete to authenticated
  using (public.can_write_key(key));

-- ── 7. RLS policies: profiles ────────────────────────────────────────────────
--    อ่าน profile ตัวเองได้ / admin อ่านได้ทุกคน
--    การ "เขียน" profiles ทำผ่าน Edge Function (service_role) เท่านั้น → ไม่มี policy เขียนให้ client
drop policy if exists profiles_read on public.profiles;
create policy profiles_read on public.profiles
  for select to authenticated
  using (id = auth.uid() or public.auth_role() = 'admin');

-- ── 8. RLS policies: access_logs ─────────────────────────────────────────────
--    ผู้ใช้ที่ล็อกอินเพิ่ม log ของตัวเองได้ / admin อ่านได้ทั้งหมด
drop policy if exists logs_insert on public.access_logs;
create policy logs_insert on public.access_logs
  for insert to authenticated
  with check (user_id = auth.uid());

drop policy if exists logs_read on public.access_logs;
create policy logs_read on public.access_logs
  for select to authenticated
  using (public.auth_role() = 'admin');

-- ── 9. Trigger: สร้าง profile อัตโนมัติเมื่อมี user ใหม่ใน Auth ──────────────
--    ดึง username/name/role จาก user_metadata ที่ Edge Function ส่งมาตอนสร้าง
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, username, name, role, active)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'username', split_part(new.email,'@',1)),
    coalesce(new.raw_user_meta_data->>'name',     split_part(new.email,'@',1)),
    coalesce(new.raw_user_meta_data->>'role',     'qs'),
    true
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- ── 10. เปิด Realtime สำหรับ kv_store ────────────────────────────────────────
--    (realtime จะเคารพ RLS ให้เอง — ผู้ใช้เห็นเฉพาะสิ่งที่ตัวเอง SELECT ได้)
do $$
begin
  alter publication supabase_realtime add table public.kv_store;
exception when duplicate_object then null;
end $$;

-- ════════════════════════════════════════════════════════════════════════════
--  เสร็จแล้ว! ขั้นตอนถัดไป (ทำครั้งเดียว):
--  1) Deploy Edge Function `admin-users` (ดู edge-functions/admin-users/)
--  2) สร้างบัญชี admin คนแรก — ดู MIGRATION.md ขั้นตอนที่ 3
--  ข้อมูลเก่าในคีย์ tcs-users / tcs-logs เลิกใช้แล้ว ลบทิ้งได้:
--    delete from public.kv_store where key in ('tcs-users','tcs-logs');
-- ════════════════════════════════════════════════════════════════════════════
