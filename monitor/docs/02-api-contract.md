# 02 — สัญญา JSON ระหว่าง Apps Script กับจอ

## 1. Endpoint

```
GET {WEB_APP_URL}/exec?k=<token>&b=<branch>[&h=<hash เดิม>]
```

| พารามิเตอร์ | ค่า | ความหมาย |
| :-- | :-- | :-- |
| `k` | บังคับ | token ที่ตั้งไว้ใน Script Property `MONITOR_TOKEN` ไม่ตรง → `{"e":"auth"}` |
| `b` | `TPS-03` / `ALL` | สาขาที่จอนี้แสดง (ค่าเริ่มต้น `ALL`) |
| `h` | ไม่บังคับ | hash ที่จอถืออยู่ ถ้าตรงกับของใหม่ server ตอบ `{"nc":1,"ts":"..."}` |

> **ทำไมไม่ใช้ ETag/304** — Apps Script Web App กำหนด HTTP header เองไม่ได้
> จึงย้ายกลไก "ไม่เปลี่ยนก็ไม่ต้องอ่านใหม่" มาไว้ใน body ผ่าน `h`/`nc` แทน
> `h` คำนวณจาก payload ทั้งก้อน **ยกเว้น `ts`** เวลาที่เดินไปเรื่อย ๆ จึงไม่ทำให้ hash เปลี่ยน

Deploy: **Execute as = Me**, **Who has access = Anyone** (URL + token คือการป้องกันชั้นเดียว)
ห้ามใส่ `Manager_Email`, `Recorded_By`, `LINE_Group_ID` ลงใน payload

## 2. รูปแบบ payload

```jsonc
{
  "v": 1,                          // เวอร์ชันสัญญา เปลี่ยนเมื่อ breaking change
  "b": "TPS-03",
  "bn": "โรบินสันฉลอง",
  "nm": 24,                        // จำนวนตู้ Active ของสาขานี้
  "d": { "r": 12450, "t": 3033, "p": 410.5, "n": 12 },
  "c": {
    "l": [ { "t": "09:40", "a": 3200, "n": 6 } ],   // รอบเก็บเงิน สูงสุด 6 รอบ
    "cash": 9800, "df": -2650, "fl": 1              // fl=1 เมื่อส่วนต่างเกิน 3%
  },
  "w": { "v": [9800,11200,8600,13400,15100,12900,12450], "s": 83450, "pv": 76200, "ch": 9.5 },
  "m": {
    "pd": "2569-09", "r": 186500, "t": 91000, "p": 204.9,
    "f": 349687, "fs": "G", "ms": 110, "sh": 0,
    "cash": 180000, "ed": 16, "dm": 30
  },
  "s": { "hdr": ["จ","อ","พ","พฤ","ศ","ส","อา"], "rows": [["ไอริน","D","D","OFF","D","D","N","N"]] },
  "wn": ["first-reading:2"],
  "h": "1a2b3c4d",
  "ts": "2026-09-16 14:32"
}
```

| คีย์ | ความหมาย |
| :-- | :-- |
| `d.r / d.t / d.p / d.n` | ยอดวันนี้ / เป้าวัน / % ของเป้า (`null` = ไม่มีเป้า) / จำนวน log |
| `c.l[].t / a / n` | เวลารอบ / เงินสดรวมรอบนั้น / จำนวนตู้ในรอบ |
| `w.v[]` | ยอด 7 วันเรียงเก่า→ใหม่ ตัวสุดท้าย = วันนี้ |
| `m.fs` | `G` ≥100% · `Y` 90–99.9% · `R` <90% · `N` ไม่มีเป้า |
| `m.sh` | 1 = สาขานี้ใช้เป้ารวมกับสาขาคู่ (TSP-06 / TSP-57) |
| `wn[]` | `code:count` — `first-reading`, `meter-reset`, `no-coin-price`, `bad-date`, `no-target`, `target-shared` |

ขนาด payload ปกติ ~1.2–1.8 KB (ไม่รวมตารางกะ)

## 3. สูตรคำนวณ

| ค่า | สูตร | ฟังก์ชัน |
| :-- | :-- | :-- |
| รายได้ต่อแถว | `Revenue` ถ้า > 0, ไม่งั้น `Meter_Diff × Coin_Price` | `TSP_rowRevenue_()` |
| กัน first reading | `Meter_Previous = 0` และไม่เคยมี log ของตู้นั้น → 0 | `TSP_rowRevenue_()` |
| มิเตอร์รีเซ็ต | `Meter_Diff < 0` → 0 + warning | `TSP_rowRevenue_()` |
| รายได้รายวัน | `Σ` รายได้ต่อแถวของสาขานั้นในวันนั้น | `TSP_aggregate_()` |
| รอบเช็คพอยต์ | แถว `Is_Collection = TRUE` ที่ `Timestamp` ห่างกัน ≤ 45 นาที = รอบเดียวกัน | `TSP_clusterCheckpoints_()` |
| ส่วนต่างเงินสด | `Σ Cash_Collected − รายได้วันนั้น` ; ธงแดงเมื่อ `|ส่วนต่าง| / รายได้ > 3%` | `TSP_buildPayload_()` |
| เป้าเดือน | แท็บ Target งวด `พ.ศ.-MM` (ยึดแถวท้ายสุด) ; fallback `R0 × seasonal_index` | `TSP_findTarget_()` |
| เป้ารายวัน | `เป้าเดือน ÷ จำนวนวันในเดือน` | `TSP_buildPayload_()` |
| % ของเป้า | `ยอด ÷ เป้า × 100` (เป้า ≤ 0 → `null`) | `TSP_pct_()` |
| คาดการณ์ปิดเดือน | `MTD ÷ วันที่ผ่านไป × จำนวนวันในเดือน` | `TSP_pace_()` |
| milestone | ค่ามากสุดใน `MILESTONES` ที่ % ผ่านแล้ว | `TSP_milestone_()` |
| hash | FNV-1a 32-bit ของ payload (ไม่รวม `ts`) | `TSP_hash_()` |

## 4. ติดตั้ง

1. เปิด `TSP_Master_Database Ver.2` → Extensions → Apps Script (หรือสร้างโปรเจกต์ standalone)
2. วางไฟล์ `TSP_Monitor_Calc.gs` และ `TSP_Monitor_API.gs`
3. แก้ token ในฟังก์ชัน `TSP_setup()` แล้วรัน 1 ครั้ง (จะเขียนค่าลง Script Properties)
4. รัน `TSP_test_sheets()` → ต้องเจอครบ Daily_Logs / Machines / Branches / Target / Branch / SeasonalIndex
5. รัน `TSP_test_today()` → ดู payload ใน Execution log แล้วเทียบยอดกับชีตด้วยมือ
6. Deploy → New deployment → Web app → Execute as **Me** / Access **Anyone** → คัดลอก URL
7. เปิด `monitor/apps-script/preview.html` วาง URL + `?k=...&b=TPS-03` กดโหลด → ตรวจ layout ทั้ง 5 หน้า

## 5. เปลี่ยนสัญญาเมื่อไหร่ต้องขยับอะไร

แก้คีย์ใน payload → ต้องแก้ 3 จุดพร้อมกัน:
`TSP_Monitor_Calc.gs` (สร้าง) · `preview.html` (ตัวอย่าง + การวาด) · `net_client.cpp` (parse)
เพิ่มคีย์ใหม่อย่างเดียวไม่ต้องขยับ `v` — ลบหรือเปลี่ยนความหมายให้เพิ่ม `v` เป็น 2
