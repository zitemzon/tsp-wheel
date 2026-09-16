# 01 — ประเมินความสามารถระบบมอนิเตอร์รายได้

> ตรวจจากข้อมูลจริงในชีต + โค้ดจริงของระบบเดิม (16 ก.ย. 2026)
> `TSP_Master_Database Ver.2` — `1FU1iosNzzepz8AVAFOz-ZRwzPzi_O5wgiOV3TYwcomk`
> `TSP_TargetBot_Config` — `1MmioNBkITy7HiZTHXJcJOQFzRW0eTUwam55jovaxnJ4`
> โค้ดเดิม: `TSP_Unified_Report.gs` v3.1 · `TSP_LINE_Flex.gs` v1.1 · `Code.gs`

## 1. ข้อมูลที่มีอยู่จริง

| แท็บ | แถว | คอลัมน์ที่จอใช้ |
| :-- | --: | :-- |
| Branches | 47 (Active 8) | `Branch_ID`, `Branch_Name`, `Status` |
| Machines | 206 | `Machine_ID`, `Branch_ID`, `Status`, **`Revenue_Method`** |
| **Daily_Logs** | 79 | `Log_Date`, `Machine_ID`, `Branch_ID`, `Revenue`, `Meter_Diff`, `Cash_Collected`, `Is_Collection`, `Timestamp` |
| Employees | 13 | `Emp_Name`, `Branch_ID`, `Status` |
| Cash_Deposits / Expenses / Inventory_Logs / Settings | 0 | ยังว่าง |

ประเภทตู้ตาม `Revenue_Method`: **METER 185 ตู้ · CAPSULE 21 ตู้** (`Coin_Price` = 10 ทุกตู้)
ประเภทตามป้าย: ตู้คีบ 142 · ตู้กาชาปอง 34 · เครื่องเล่น 28 · ตู้โมเดลกล่องสุ่ม 2

## 2. สรุปว่าทำอะไรได้/ไม่ได้

| หน้าที่ต้องการ | สถานะ | หมายเหตุ |
| :-- | :-: | :-- |
| ยอดวันนี้ + นาฬิกา | ✅ | ผลรวม `Revenue` ตรงกับการ์ด LINE |
| จุดเช็คพอยต์รอบเก็บเงิน | ✅ | `Cash_Collected` + กติกา CAPSULE (ดูข้อ 3.1) |
| ยอด 7 วัน + เทียบสัปดาห์ก่อน | ✅ | แหล่งเดียวกับ `lnTrend_()` |
| ยอดเดือน + % เป้า + milestone | ✅ | Target งวด 2569-09 พร้อมใช้ |
| คาดการณ์ยอดปิดเดือน (pace) | ✅ | MTD ÷ วันที่ผ่าน × วันในเดือน |
| ตารางงานรายสัปดาห์ | ⚠️ | ไม่มีแท็บกะใน Master DB — ต้องต่อ `SCHEDULE_URL` |
| กำไร/ค่าใช้จ่ายต่อสาขา | ❌ | Expenses/Cash_Deposits ยังว่าง |
| ยอดรายตู้บนจอ | ❌ (ตั้งใจ) | จอ 320×240 อ่านไม่ไหว 206 ตู้ — ดูใน AppSheet/PDF แทน |

## 3. กติกาที่ระบบเดิมใช้จริง (จอต้องทำตามเป๊ะ)

### 3.1 แถวที่ "ต้องมีเงินสด" — ห้ามใช้ `Is_Collection` เดี่ยว ๆ

`Code.gs > shouldHaveCash_()` ระบุว่าต้องตรงกับ Show_If / Require_If ของ `Cash_Collected` ใน AppSheet:

```
OR([Is_Collection], [Machine_ID].[Revenue_Method] <> "METER")
```

ตู้ CAPSULE ถูกซ่อนช่องติ๊กไว้ `Is_Collection` จึงเป็น FALSE เสมอทั้ง 21 ตู้
ถ้ากรองด้วย `Is_Collection` อย่างเดียว **เงินสดจากตู้ไข่จะหายทั้งหมด**
→ ระบบมอนิเตอร์ใช้ `TSP_shouldHaveCash_()` ที่คัดลอกกติกานี้มา และมีเทสต์กำกับไว้

### 3.2 รายได้ = ผลรวมคอลัมน์ `Revenue` ล้วน

`buildBranchData_()` และ `sumRevenue_()` ใช้ `rev += o.rev` ไม่มีการคูณ `Meter_Diff × Coin_Price`
จอจึงต้องไม่คำนวณเอง ไม่งั้นตัวเลขจะมากกว่าการ์ด LINE และ PDF
(`Code.gs` บันทึกไว้ว่า median รายได้ต่อแถว = 0 และ 57% ของแถวเป็น 0 — เป็นเรื่องปกติของธุรกิจนี้)

แถวที่ `Revenue = 0` ทั้งที่ `Meter_Diff > 0` จะถูกนับเป็น warning `rev-missing` ขึ้นบนจอ
เพื่อให้เห็นว่ามีแถวที่ AppSheet ไม่ได้คำนวณให้ — แต่ **ไม่เอาไปบวกในยอด**

### 3.3 ปี พ.ศ. ที่หลุดลงชีต

`Code.gs` มี `fixBuddhistYear_()` + `repairBuddhistDates()` เพราะ AppSheet เคยเขียนปี **2569**
ลงชีตจริง (บันทึกไว้ว่าเจอ 30 ก.ค. 2026) ถ้าไม่แปลง แถวนั้นจะหลุดออกนอกทุกช่วงเวลาและยอดหายเงียบ ๆ
→ `TSP_toDate_()` / `TSP_toDateTime_()` ของมอนิเตอร์แปลงให้ทั้งกรณี Date object, สตริง และ serial number

### 3.4 การตีความวันที่แบบสตริงยังไม่ตรงกันในระบบเดิม

| ที่ | สตริง `d/m/Y` ถูกอ่านเป็น |
| :-- | :-- |
| `TSP_Unified_Report.gs > toDate_()` | **DMY** |
| `Code.gs > toDate_()` | ปล่อยให้ JS ตีความ = **MDY** |
| มอนิเตอร์ | ตัวไหนเกิน 12 = วัน ถ้าชี้ขาดไม่ได้ใช้ `DATE_ORDER` (ค่าเริ่มต้น MDY) |

ในทางปฏิบัติจะไม่ต่างกัน **ถ้าเซลล์เป็น Date จริง** (ซึ่งเป็นกรณีปกติ — ชีตตั้ง locale แบบ en_US จึงแสดง `07/28/2026`)
มอนิเตอร์ยังส่ง `Log_Date` ผ่าน `toDate_()` ของเอนจินเดิมก่อนเสมอ เพื่อไม่ให้ตีความต่างจากรายงาน

## 4. ช่องโหว่ข้อมูลที่ระบบจัดการให้แล้ว

| # | ปัญหา | ทางที่ใช้ |
| :-: | :-- | :-- |
| 1 | รหัสสาขาคนละสคีมา `TPS-xx` ↔ `TSP-xx` | `TSP_normBranch_()` |
| 2 | TargetBot ยุบหน่วย `TPS-06A+06B → TSP-06`, `TPS-05+TPS-07 → TSP-57` | `TSP_targetUnit_()` + ธง `sh:1` บนจอ |
| 3 | แท็บ Target มีแถวซ้ำ (งวด 2569-08 ซ้ำ 3 ชุด) | ยึดแถวท้ายสุดของคู่ (period, branch_code) |
| 4 | ไม่มีเป้ารายวัน | v1 กระจายเท่า ๆ `เป้าเดือน ÷ จำนวนวัน` |
| 5 | แถวเก็บเงินที่ยังไม่กรอกจำนวน | warning `cash-missing` |

## 5. เรื่องที่ต้องตัดสินใจ (ไม่ใช่ปัญหาโค้ด)

1. **ยอดไม่ตรงที่ TargetBot เตือนไว้เอง** — เซ็นทรัลภูเก็ต ส.ค. 2569 รายสัปดาห์รวม 74,020
   แต่ยอดรวมเดือนเขียนไว้ 71,000 (ต่าง 3,020 บาท) ต้องแก้ที่ไฟล์ต้นทาง ระบบไม่เดาให้
2. **แท็บ Settings / Config ใน Master DB ว่าง** — ควรใช้เก็บ config ระดับร้าน
3. **Daily_Logs เริ่ม 28/07/2026** — กราฟ 7 วันจะเต็มจริงเมื่อคีย์ครบทุกวัน

## 6. ข้อจำกัดฝั่งเทคนิค

**Apps Script**
- ตั้ง HTTP header เองไม่ได้ → ใช้ ETag/304 ไม่ได้ เปลี่ยนมาใช้ `?h=<hash>` แทน
- **มี `doGet` ได้ตัวเดียวต่อโปรเจกต์** — โปรเจกต์ Master DB มีอยู่แล้ว (redirect ไป Dashboard)
  จึงต้องรวมเป็น router เดียว (ดู `02-api-contract.md`)
- โควตาบัญชี @gmail ≈ 90 นาที execution/วัน และแชร์กับบอท LINE
  → poll 60 วิ + `CacheService` 90 วิ + กลไก `?h=` (ปกติไม่แตะชีตเลยถ้ายอดไม่ขยับ)
- `TSP_monLogs_()` อ่าน Daily_Logs ทั้งแท็บเข้าหน่วยความจำ — ที่ 79 แถวไม่มีปัญหา
  ถ้าโตเกิน ~50,000 แถวค่อยเปลี่ยนไปกรองด้วย `dlIndex_().byBranch`

**ฮาร์ดแวร์**
- ESP32-S3-N16R8: flash 16MB / PSRAM 8MB — sprite 3 ชั้น (320×240×2B) ใช้ ~450KB
- สาย IDC 8 เส้น (`SDA SCL CS RS RST BLK VCC GND`) — ไม่มี MISO จึงอ่าน driver ID จากจอไม่ได้
- fade ทำที่ framebuffer (blend RGB565) ไม่ใช่หรี่ backlight — BLK สงวนไว้หรี่จอกลางคืน
- ฟอนต์ไทยไม่มีใน LovyanGFX → ค่าเริ่มต้นใช้ป้ายอังกฤษ (`HAS_THAI_FONT 0`) ตัวเลขอ่านได้ครบปกติ
