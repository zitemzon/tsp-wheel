/**
 * tsp_monitor.ino — TSP Revenue Monitor (ESP32-S3 + TFT 320x240)
 * ==================================================================
 * ไฟล์นี้เป็น "ป้ายชื่อสเก็ตช์" ให้ Arduino IDE เท่านั้น ไม่มีโค้ดอยู่ในนี้
 * โค้ดจริงอยู่ในไฟล์ .cpp / .h ในโฟลเดอร์เดียวกัน ซึ่ง IDE คอมไพล์ให้อัตโนมัติ:
 *
 *   main.cpp        setup() / loop()  — ลูปหลัก ดึงข้อมูล + สลับหน้า
 *   net_client.cpp  WiFi + HTTPS + แปลง JSON
 *   ui_pages.cpp    วาดจอ 5 หน้า + crossfade
 *   ui_util.h       ฟังก์ชันคำนวณล้วน (มีเทสต์ฝั่ง PC)
 *   config.h        ค่าของคุณเอง — ก๊อปจาก config.h.example แล้วแก้
 *
 * ── ก่อน Upload ครั้งแรก ─────────────────────────────────────────
 * 1) Library Manager: ติดตั้ง LovyanGFX และ ArduinoJson
 * 2) ก๊อป config.h.example -> config.h แล้วแก้ WiFi / URL / token / รุ่นจอ
 * 3) Tools ตั้งให้ตรงกับ ESP32-S3-N16R8:
 *      Board              : ESP32S3 Dev Module
 *      PSRAM              : OPI PSRAM          <-- สำคัญ ผิดข้อนี้ sprite จองไม่ผ่าน
 *      Flash Size         : 16MB (128Mb)
 *      Partition Scheme   : 16M Flash (3MB APP/9.9MB FATFS)
 *      USB CDC On Boot    : Enabled            <-- ไม่งั้นไม่เห็น Serial Monitor
 *      Upload Speed       : 921600
 *
 * ยังไม่รู้ว่าจอเป็น ILI9341 หรือ ST7789 ?
 *   เปิด ../probe/panel_probe/panel_probe.ino แล้ว Upload ก่อน
 *
 * รายละเอียดทั้งหมด: monitor/docs/03-hardware.md
 * ==================================================================
 */
