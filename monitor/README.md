# TSP Revenue Monitor

จอมอนิเตอร์รายได้ประจำวันของ Toy Station Plus+ — ESP32-S3 + TFT 320×240
ดึงสรุปยอดจาก Apps Script Web App ที่อ่าน `TSP_Master_Database Ver.2` (ที่พนักงานคีย์ผ่าน AppSheet)
แล้ววนแสดง 5 หน้าแบบ crossfade

```
AppSheet ──► TSP_Master_Database Ver.2 ─┐
TSP_TargetBot_Config ───────────────────┼─► TSP_Monitor_API.gs ──HTTPS──► ESP32-S3 ──► TFT
Schedule Web App (ของเดิม) ─────────────┘      (JSON + cache 90s)         5 หน้า fade
```

| หน้า | แสดงอะไร |
| :-: | :-- |
| 1 | ยอดวันนี้ (ตัวเลขใหญ่) + Progress Bar เป้าวัน + นาฬิกา |
| 2 | จุดเช็คพอยต์รอบเก็บเงิน + เงินสดที่นับได้ + ส่วนต่าง |
| 3 | แท่ง 7 วันล่าสุด + เทียบสัปดาห์ก่อน |
| 4 | ยอดสะสมเดือน + % ของเป้า + milestone + คาดการณ์ปิดเดือน |
| 5 | ตารางกะสัปดาห์นี้ (รอเชื่อม `SCHEDULE_URL`) |

## โครงไฟล์

```
docs/01-assessment.md     ประเมินความสามารถ + ช่องโหว่ข้อมูลที่เจอในชีตจริง
docs/02-api-contract.md   สัญญา JSON + สูตรคำนวณทุกตัว + ขั้นตอน deploy
docs/03-hardware.md       ผังขา, driver IC, งบหน่วยความจำ, ฟอนต์ไทย
apps-script/TSP_Monitor_Calc.gs   สูตรล้วน (ทดสอบด้วย node ได้)
apps-script/TSP_Monitor_Route.gs  doGet router + cache + token + ตารางกะ
apps-script/preview.html          จำลองหน้าจอ 320×240 ในเบราว์เซอร์ = สเปกภาพของเฟิร์มแวร์
firmware/                         PlatformIO (LovyanGFX + ArduinoJson)
tests/calc.test.js                ทดสอบสูตรฝั่ง Apps Script  (node)
tests/ui_util_test.cpp            ทดสอบฟังก์ชันคำนวณฝั่งจอ  (g++)
```

## รันทดสอบ

```bash
node monitor/tests/calc.test.js
g++ -std=c++17 -I monitor/firmware/src monitor/tests/ui_util_test.cpp -o /tmp/t && /tmp/t
```

## เริ่มใช้งาน

1. วาง `.gs` 2 ไฟล์ใน **โปรเจกต์เดียวกับ `TSP_Unified_Report.gs`** (ผูกกับ Master DB)
   แล้วเปลี่ยนชื่อ `doGet` เดิมในไฟล์นั้นเป็น `TSP_dashboardRedirect_()`
2. `TSP_monSetup()` → `TSP_test_columns()` → **`TSP_test_vs_report()` ต้องขึ้น ✅ ตรงกันทุกสาขา**
3. Deploy แบบ **New version** (คง URL เดิมไว้ ไม่ให้ LINE webhook พัง)
4. เปิด `preview.html` ตรวจหน้าตา → `cp firmware/src/config.h.example firmware/src/config.h` → `pio run -t upload`

ยอดบนจอถูกออกแบบให้ **เท่ากับการ์ด LINE และรายงาน PDF เสมอ** — ใช้ตัวอ่านชีตและกติกาชุดเดียวกัน
(`Revenue` ล้วน · กติกาเงินสดของตู้ CAPSULE · การแปลงปี พ.ศ.)

รายละเอียดทั้งหมดอยู่ใน `docs/`
