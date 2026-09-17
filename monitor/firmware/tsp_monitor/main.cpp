/**
 * TSP Revenue Monitor — ESP32-S3 + TFT 320x240
 * ดึงสรุปยอดจาก Apps Script Web App แล้ววนแสดง 5 หน้าแบบ fade
 */
#include <Arduino.h>
#include <time.h>
#include "config.h"
#include "net_client.h"
#include "ui_pages.h"

static MonitorData gData;
static int      gPage       = 0;
static uint32_t tFetch      = 0;
static uint32_t tPage       = 0;
static uint32_t tClock      = 0;
static int      failStreak  = 0;

static void fetchNow() {
  int r = net_fetch(gData);
  tFetch = millis();
  switch (r) {
    case NET_UPDATED:
      failStreak = 0;
      Serial.printf("[net] อัปเดตแล้ว hash=%s วันนี้=%ld เดือน=%ld\n", gData.hash, gData.dayRev, gData.mRev);
      ui_render(gPage, gData);           // วาดทับทันทีด้วยตัวเลขใหม่ ไม่ต้องรอสลับหน้า
      break;
    case NET_NO_CHANGE:
      failStreak = 0;
      Serial.println("[net] ข้อมูลไม่เปลี่ยน");
      break;
    default:
      failStreak++;
      Serial.printf("[net] ล้มเหลว r=%d http=%d ครั้งที่ %d\n", r, net_lastHttpCode(), failStreak);
      if (failStreak == 1 && !gData.valid) ui_splash("cannot reach API", "retrying...");
      if (failStreak >= 5) { Serial.println("[net] ต่อไม่ได้ 5 ครั้ง -> รีสตาร์ต"); delay(500); ESP.restart(); }
      break;
  }
}

void setup() {
  Serial.begin(115200);
  delay(200);
  ui_init();
  ui_splash("connecting wifi...", WIFI_SSID);

  if (!net_begin()) {
    ui_splash("wifi failed", "check config.h");
    delay(3000);
    ESP.restart();
  }
  ui_splash("syncing clock...", nullptr);
  net_syncTime();

  ui_splash("loading data...", API_BRANCH);
  fetchNow();

  if (gData.valid) ui_render(0, gData);
  tPage = tClock = millis();
}

void loop() {
  uint32_t now = millis();

  if (now - tFetch >= (uint32_t)POLL_SEC * 1000UL) fetchNow();

  if (gData.valid && now - tPage >= PAGE_HOLD_MS) {
    gPage = (gPage + 1) % PAGE_COUNT;
    ui_fadeTo(gPage, gData);
    tPage = millis();
  }

  if (now - tClock >= CLOCK_REFRESH_MS) {
    tClock = now;
    struct tm tm;
    if (getLocalTime(&tm, 5)) {
      ui_backlightForHour(tm.tm_hour);
      // เวลาบนหัวจอมาจาก payload — วาดซ้ำเพื่อให้ตัวเลขอื่นสดตามรอบ fetch
      if (gData.valid) ui_render(gPage, gData);
    }
  }
  delay(20);
}
