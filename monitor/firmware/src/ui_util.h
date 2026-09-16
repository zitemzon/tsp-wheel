/**
 * ui_util.h — ฟังก์ชันคำนวณล้วนของฝั่งจอ (ไม่พึ่ง Arduino/LovyanGFX)
 * แยกออกมาเพื่อทดสอบบนเครื่อง PC ได้: monitor/tests/ui_util_test.cpp
 */
#pragma once
#include <stdint.h>
#include <stdio.h>
#include <string.h>

/** 1234567 -> "1,234,567" (รองรับค่าติดลบ) */
static inline void tsp_fmtBaht(long v, char *buf, size_t cap) {
  if (!buf || cap == 0) return;
  char tmp[24];
  bool neg = v < 0;
  unsigned long uv = neg ? (unsigned long)(-v) : (unsigned long)v;
  int n = snprintf(tmp, sizeof(tmp), "%lu", uv);
  int commas = (n - 1) / 3;
  int out = n + commas + (neg ? 1 : 0);
  if ((size_t)out >= cap) { snprintf(buf, cap, "%s%s", neg ? "-" : "", tmp); return; }
  buf[out] = '\0';
  int j = out - 1, cnt = 0;
  for (int i = n - 1; i >= 0; i--) {
    buf[j--] = tmp[i];
    if (++cnt % 3 == 0 && i > 0) buf[j--] = ',';
  }
  if (neg) buf[0] = '-';
}

/** ผสมสอง RGB565 ด้วยน้ำหนัก f (0 = สีแรกล้วน, 32 = สีที่สองล้วน) */
static inline uint16_t tsp_blend565(uint16_t a, uint16_t b, uint8_t f) {
  uint16_t ar = (a >> 11) & 0x1F, ag = (a >> 5) & 0x3F, ab = a & 0x1F;
  uint16_t br = (b >> 11) & 0x1F, bg = (b >> 5) & 0x3F, bb = b & 0x1F;
  uint16_t r = (uint16_t)((ar * (32 - f) + br * f) >> 5);
  uint16_t g = (uint16_t)((ag * (32 - f) + bg * f) >> 5);
  uint16_t l = (uint16_t)((ab * (32 - f) + bb * f) >> 5);
  return (uint16_t)((r << 11) | (g << 5) | l);
}

/** ตอนนี้อยู่ช่วงกลางคืนหรือยัง (รองรับช่วงข้ามเที่ยงคืน เช่น 22 -> 8) */
static inline bool tsp_isNight(int hour, int startHour, int endHour) {
  if (startHour > endHour) return (hour >= startHour || hour < endHour);
  return (hour >= startHour && hour < endHour);
}

/** ตำแหน่ง x ของจุดเช็คพอยต์ที่ i เมื่อมีทั้งหมด n จุด บนเส้น x0..x1 */
static inline int tsp_cpX(int i, int n, int x0, int x1) {
  if (n <= 1) return x0 + (x1 - x0) / 2;
  return x0 + (x1 - x0) * i / (n - 1);
}
