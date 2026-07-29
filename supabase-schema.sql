-- ════════════════════════════════════════════════════════════
-- Tender Cost System — Supabase Schema
-- วิธีใช้: ไปที่ Supabase Dashboard → SQL Editor → New Query
--          วาง SQL นี้แล้วกด Run
-- ════════════════════════════════════════════════════════════

-- 1. สร้างตาราง kv_store (key-value store สำหรับเก็บข้อมูลทั้งหมด)
create table if not exists public.kv_store (
  key         text        primary key,
  value       text        not null,
  updated_at  timestamptz default now()
);

-- 2. เปิด Row Level Security
alter table public.kv_store enable row level security;

-- 3. Policy: อนุญาตทุกคน read/write (ใช้ anon key เป็น shared access)
--    ถ้าต้องการเพิ่ม auth ในอนาคต แก้ policy ตรงนี้
create policy "allow_all" on public.kv_store
  for all
  using (true)
  with check (true);

-- 4. เปิด Realtime สำหรับตาราง kv_store
--    (ต้องทำใน Dashboard: Database → Replication → kv_store ✓)
alter publication supabase_realtime add table public.kv_store;
