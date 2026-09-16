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
apps-script/TSP_Monitor_API.gs    doGet + cache + token + ตารางกะ
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

1. วาง `.gs` 2 ไฟล์ในโปรเจกต์ Apps Script → รัน `TSP_setup()` → `TSP_test_sheets()` → `TSP_test_today()`
2. Deploy เป็น Web app แล้วเปิด `preview.html` ตรวจหน้าตาก่อน
3. `cp firmware/src/config.h.example firmware/src/config.h` แก้ค่า → `pio run -t upload`

รายละเอียดทั้งหมดอยู่ใน `docs/`
