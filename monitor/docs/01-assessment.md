# 01 — ประเมินความสามารถระบบมอนิเตอร์รายได้

> ตรวจจากข้อมูลจริงในชีตเมื่อ 16 ก.ย. 2026
> `TSP_Master_Database Ver.2` — `1FU1iosNzzepz8AVAFOz-ZRwzPzi_O5wgiOV3TYwcomk`
> `TSP_TargetBot_Config` — `1MmioNBkITy7HiZTHXJcJOQFzRW0eTUwam55jovaxnJ4`

## 1. ข้อมูลที่มีอยู่จริง

| แท็บ (ระบุด้วยหัวคอลัมน์) | แถว | คอลัมน์ที่ระบบมอนิเตอร์ใช้ |
| :-- | --: | :-- |
| Branches | 47 (Active 8) | `Branch_ID`, `Branch_Name`, `Machine_Prefix`, `Status` |
| Machines | 207 | `Machine_ID`, `Branch_ID`, `Coin_Price`, `Revenue_Method`, `Status` |
| **Daily_Logs** | 79 | `Log_Date`, `Machine_ID`, `Branch_ID`, `Meter_Current/Previous/Diff`, `Revenue`, `Reading_Status`, `Cash_Collected`, `Is_Collection`, `Timestamp` |
| Employees | 13 | `Emp_Name`, `Branch_ID`, `Status` |
| Deposits / Expenses / Inventory / Settings | 0 | — ยังว่าง |

สาขาที่ Active: TPS-01 โกรเซอรี่เมือง · TPS-02 ไลม์ไลท์ · TPS-03 โรบินสันฉลอง ·
TPS-04 เซ็นทรัลเชิงทะเล · TPS-05 จังซีลอน · TPS-06A/06B เซ็นทรัลภูเก็ต · TPS-07 เซ็นทรัลป่าตอง

ฝั่ง TargetBot มีครบสำหรับ Progress Bar: `R0` รายสาขา, `seasonal_index` 12 เดือน,
แท็บ Target งวด `2569-09` คำนวณไว้แล้ว, `MILESTONES = 50,80,100,110`, `weeks_in_month`

## 2. สรุปว่าทำอะไรได้/ไม่ได้

| หน้าที่ต้องการ | สถานะ | หมายเหตุ |
| :-- | :-: | :-- |
| ยอดวันนี้ + นาฬิกา | ✅ | จาก `Revenue` รายแถว |
| จุดเช็คพอยต์รอบเก็บเงิน | ✅ | `Is_Collection` + `Cash_Collected` + `Timestamp` |
| ยอด 7 วัน + เทียบสัปดาห์ก่อน | ✅ | group by `Log_Date` |
| ยอดเดือน + % เทียบเป้า + milestone | ✅ | Target งวด 2569-09 พร้อมใช้ |
| คาดการณ์ยอดปิดเดือน (pace) | ✅ | MTD ÷ วันที่ผ่าน × วันในเดือน |
| ตารางงานรายสัปดาห์ | ⚠️ | ไม่มีแท็บกะใน Master DB — ต้องต่อกับ web app เดิมผ่าน `SCHEDULE_URL` |
| กำไร/ค่าใช้จ่ายต่อสาขา | ❌ | แท็บ Expenses/Deposits ว่าง — ไม่อยู่ในรอบนี้ |
| ยอดรายตู้บนจอ | ❌ (ตั้งใจ) | จอ 320×240 อ่านไม่ไหว 207 ตู้ — ดูใน AppSheet แทน |

## 3. ช่องโหว่ข้อมูลที่ระบบจัดการให้แล้ว

| # | ปัญหา | ผลถ้าไม่แก้ | ทางที่ใช้ |
| :-: | :-- | :-- | :-- |
| 1 | รหัสสาขาคนละสคีมา: Master DB = `TPS-xx`, TargetBot = `TSP-xx` | join ไม่ติด เป้าไม่ขึ้น | `TSP_normBranch_()` แปลงเป็นรูปเดียว |
| 2 | TargetBot ยุบหน่วย: `TPS-06A+06B → TSP-06`, `TPS-05+TPS-07 → TSP-57` | เทียบเป้าผิดสาขา | `TSP_targetUnit_()` + ธง `sh:1` บนจอว่า "เป้ารวมหน่วย" |
| 3 | แถวอ่านมิเตอร์ครั้งแรก `Meter_Previous = 0` | ยอดพุ่งผิดปกติ | `TSP_rowRevenue_()` ตัดเป็น 0 + warning `first-reading` |
| 4 | `Meter_Diff` ติดลบ (รีเซ็ต/เปลี่ยนบอร์ด) | ยอดติดลบ | ตัดเป็น 0 + warning `meter-reset` |
| 5 | ตู้ไม่มี `Coin_Price` | fallback คำนวณไม่ได้ | ข้าม + warning `no-coin-price` |
| 6 | แท็บ Target มีแถวซ้ำ (งวด 2569-08 ซ้ำ 3 ชุด) | เป้าถูกนับซ้ำ | ยึดแถวท้ายสุดของคู่ (period, branch_code) |
| 7 | ไม่มีเป้ารายวัน | จอหน้า 1 ไม่มี Progress Bar | v1 กระจายเท่า ๆ `Target_month ÷ days_in_month` |
| 8 | `Log_Date` เป็นสตริง `07/28/2026` (MDY) | วันที่เพี้ยน | `TSP_toDate_()` + สวิตช์ `DATE_ORDER` |
| 9 | ชื่อแท็บอาจถูกแก้ | สคริปต์พัง | ค้นแท็บจาก "ลายเซ็นหัวคอลัมน์" ไม่ใช่ชื่อแท็บ |

## 4. เรื่องที่ยังต้องตัดสินใจ (ไม่ใช่ปัญหาโค้ด)

1. **ยอดไม่ตรงที่ TargetBot เตือนไว้เอง** — เซ็นทรัลภูเก็ต ส.ค. 2569 รายสัปดาห์รวม 74,020
   แต่ยอดรวมเดือนเขียนไว้ 71,000 (ต่าง 3,020 บาท) ต้องแก้ที่ไฟล์ต้นทาง ระบบไม่เดาให้
2. **แท็บ Settings (Key/Value) ใน Master DB ว่าง** — ควรใช้เก็บ config ระดับร้าน
3. **Daily_Logs มี 79 แถว เริ่ม 28/07/2026** — ยอดย้อนหลังยังสั้น กราฟ 7 วันจะเต็มจริงเมื่อคีย์ครบทุกวัน

## 5. ข้อจำกัดฝั่งเทคนิค

**Apps Script**
- ตั้ง HTTP header เองไม่ได้ → ใช้ ETag/304 ไม่ได้ เปลี่ยนมาใช้กลไก `?h=<hash>` แทน (ดู `02-api-contract.md`)
- โควตาบัญชี @gmail ≈ 90 นาที execution/วัน → poll 60 วิ = 1,440 ครั้ง/วัน ต้องพึ่ง `CacheService` (TTL 90 วิ)
- ถ้าเปลี่ยนเป็น poll 30 วิ ให้ลด TTL ตามกัน ไม่งั้นจอค้างข้อมูลเก่า

**ฮาร์ดแวร์**
- ESP32-S3-N16R8: flash 16MB / PSRAM 8MB — sprite 3 ชั้น (320×240×2B = 150KB/ชั้น) ใช้ ~450KB สบาย ๆ
- สาย IDC ไปจอมี 8 เส้น (`SDA SCL CS RS RST BLK VCC GND`) — **ไม่มี MISO** จึงอ่าน driver ID จากจอไม่ได้
  ต้องตั้ง `PANEL_ILI9341` ใน `config.h` ให้ตรงรุ่นเอง
- fade ทำที่ framebuffer (blend RGB565) ไม่ใช่การหรี่ backlight — BLK สงวนไว้หรี่จอกลางคืน
- ฟอนต์ไทยไม่มีใน LovyanGFX → ค่าเริ่มต้นใช้ป้ายภาษาอังกฤษ (`HAS_THAI_FONT 0`) ตัวเลขทั้งหมดอ่านได้ปกติ
