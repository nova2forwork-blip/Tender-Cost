# Tender Cost System — คู่มือ Deploy

ระบบบริหารต้นทุนโครงการ (QS · จัดซื้อ · บัญชี)  
Backend: Supabase | Hosting: Vercel | เวลาติดตั้ง: ~20 นาที

---

## ขั้นตอนที่ 1 — สร้าง Supabase Project (ฟรี)

1. ไปที่ [supabase.com](https://supabase.com) → **Start your project**
2. Sign up / Login ด้วย GitHub
3. กด **New Project** → ตั้งชื่อ เช่น `tender-cost` → ตั้ง Database Password → **Create**
4. รอ ~2 นาที ให้ project พร้อม

### สร้างตารางใน Supabase

5. ไปที่ **SQL Editor** (แถบซ้าย) → **New Query**
6. วาง SQL จากไฟล์ `supabase-schema.sql` ทั้งหมด → กด **Run**
7. ตรวจสอบ: ไปที่ **Table Editor** จะเห็นตาราง `kv_store`

### เปิด Realtime

8. ไปที่ **Database → Replication**
9. ช่อง **Source** กด toggle ✓ หน้า `kv_store`

### คัดลอก API Keys

10. ไปที่ **Project Settings → API**
11. คัดลอก:
    - **Project URL** → `https://xxxx.supabase.co`
    - **anon / public key** → `eyJhbGci...`

---

## ขั้นตอนที่ 2 — อัปโหลดโค้ดขึ้น GitHub

```bash
# ใน terminal ที่โฟลเดอร์ tender-web
git init
git add .
git commit -m "first commit"
```

จากนั้นไปสร้าง repo ใหม่ที่ [github.com/new](https://github.com/new) แล้ว:

```bash
git remote add origin https://github.com/YOUR_USERNAME/tender-cost.git
git push -u origin main
```

---

## ขั้นตอนที่ 3 — Deploy บน Vercel (ฟรี)

1. ไปที่ [vercel.com](https://vercel.com) → **Continue with GitHub**
2. กด **Add New Project** → เลือก repo `tender-cost`
3. หน้า Configure Project → **Environment Variables** ใส่:

   | Name | Value |
   |------|-------|
   | `VITE_SUPABASE_URL` | `https://xxxx.supabase.co` |
   | `VITE_SUPABASE_ANON_KEY` | `eyJhbGci...` |

4. กด **Deploy** → รอ ~1 นาที
5. ✅ ได้ URL เช่น `https://tender-cost-abc.vercel.app`

**แชร์ URL นี้ให้ทีม** — ทุกคนเปิดได้ทันที ข้อมูล sync real-time

---

## การใช้งาน Local (Development)

```bash
# clone repo หรืออยู่ในโฟลเดอร์ tender-web แล้ว
npm install

# สร้างไฟล์ .env.local
cp .env.local.example .env.local
# แก้ไขใส่ค่า Supabase จริง

npm run dev
# เปิด http://localhost:5173
```

---

## โครงสร้างไฟล์

```
tender-web/
├── src/
│   ├── main.jsx          ← entry point
│   ├── App.jsx           ← แอปหลักทั้งหมด
│   └── supabase.js       ← Supabase client + storage helpers
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
├── supabase-schema.sql   ← SQL สำหรับ setup database
└── .env.local.example    ← ตัวอย่างไฟล์ environment
```

---

## Features

- 📐 **QS** — ลง Tender Cost ตาม Account Code 70 รายการ
- 📦 **จัดซื้อ** — ลง PO พร้อม Supplier, เลข PO, สถานะ
- 📊 **บัญชี** — Dashboard Budget vs Committed + Export Excel
- ⚡ **Real-time sync** — ทีมเห็นข้อมูลเดียวกันทันที
- 🌐 **Multi-user** — ทุกคนใช้ URL เดียวกันได้เลย ไม่ต้อง login
