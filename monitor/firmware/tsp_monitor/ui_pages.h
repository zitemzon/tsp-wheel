#pragma once
#include "monitor_data.h"

#define PAGE_COUNT 5

void ui_init();                                   // เปิดจอ + สร้าง sprite + ตรวจ byte order
void ui_splash(const char *line1, const char *line2);
void ui_render(int page, const MonitorData &d);   // วาดทันที ไม่มี fade (ใช้ตอนอัปเดตนาฬิกา)
void ui_fadeTo(int page, const MonitorData &d);   // crossfade จากภาพที่ค้างอยู่ไปหน้าใหม่
void ui_backlightForHour(int hour);               // หรี่จอตามเวลา
