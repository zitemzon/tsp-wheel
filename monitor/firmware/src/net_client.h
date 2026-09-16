#pragma once
#include "monitor_data.h"

enum NetResult { NET_UPDATED = 0, NET_NO_CHANGE = 1, NET_ERR_WIFI = -1, NET_ERR_HTTP = -2, NET_ERR_PARSE = -3 };

bool  net_begin();                       // ต่อ WiFi (บล็อกสูงสุด 20 วิ)
bool  net_isConnected();
bool  net_syncTime();                    // NTP + ตั้ง timezone ไทย
int   net_fetch(MonitorData &out);       // ดึง/อัปเดตข้อมูล -> NetResult
int   net_lastHttpCode();
