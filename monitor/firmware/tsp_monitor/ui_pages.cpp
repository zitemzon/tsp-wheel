/**
 * ui_pages.cpp — 5 หน้าจอของ TSP Revenue Monitor
 * พิกัดทั้งหมดตรงกับ monitor/apps-script/preview.html (canvas 320x240)
 * แก้ layout ที่ preview ก่อน แล้วค่อยย้ายตัวเลขมาที่นี่
 */
#include "config.h"
#include "ui_pages.h"
#include <LovyanGFX.hpp>
#include <strings.h>
#include <math.h>
#include "ui_util.h"

/* ---------------- นิยามจอ ---------------- */
#if PANEL_ILI9341
  using PanelT = lgfx::Panel_ILI9341;
#else
  using PanelT = lgfx::Panel_ST7789;
#endif

class LGFX : public lgfx::LGFX_Device {
  PanelT          _panel;
  lgfx::Bus_SPI   _bus;
  lgfx::Light_PWM _light;
public:
  LGFX() {
    { auto c = _bus.config();
      c.spi_host   = SPI2_HOST;  c.spi_mode = 0;
      c.freq_write = TFT_SPI_HZ; c.freq_read = 16000000;
      c.spi_3wire  = false;      c.use_lock  = true;
      c.dma_channel= SPI_DMA_CH_AUTO;
      c.pin_sclk   = PIN_TFT_SCLK; c.pin_mosi = PIN_TFT_MOSI;
      c.pin_miso   = -1;           c.pin_dc   = PIN_TFT_DC;   // สาย 8 เส้นไม่มี MISO
      _bus.config(c); _panel.setBus(&_bus); }
    { auto c = _panel.config();
      c.pin_cs = PIN_TFT_CS; c.pin_rst = PIN_TFT_RST; c.pin_busy = -1;
      c.panel_width = 240; c.panel_height = 320;
      c.offset_x = 0; c.offset_y = 0; c.offset_rotation = 0;
      c.readable = false;                 // อ่านกลับจากจอไม่ได้
      c.invert   = TFT_INVERT; c.rgb_order = false;
      c.dlen_16bit = false;   c.bus_shared = false;
      _panel.config(c); }
    { auto c = _light.config();
      c.pin_bl = PIN_TFT_BL; c.invert = (BL_ON_HIGH ? false : true);
      c.freq = 12000; c.pwm_channel = 7;
      _light.config(c); _panel.setLight(&_light); }
    setPanel(&_panel);
  }
};

static LGFX lcd;
static LGFX_Sprite spCur(&lcd), spNext(&lcd), spMix(&lcd);
static bool  s_swapped = false;     // sprite เก็บ RGB565 แบบสลับไบต์หรือไม่
static int   s_lastPage = -1;

static const int W = 320, H = 240;

/* ---------------- สี (ตรงกับ preview) ---------------- */
static uint16_t C_BG, C_HEAD, C_LINE, C_INK, C_INK2, C_INK3,
                C_DATA, C_DATA_HI, C_GOOD, C_WARN, C_CRIT, C_TRACK;
static void initColors() {
  C_BG      = lcd.color565(0x07,0x13,0x20);
  C_HEAD    = lcd.color565(0x0F,0x27,0x40);
  C_LINE    = lcd.color565(0x1E,0x3B,0x58);
  C_INK     = lcd.color565(0xE8,0xF0,0xF8);
  C_INK2    = lcd.color565(0x9F,0xB4,0xC8);
  C_INK3    = lcd.color565(0x5D,0x75,0x90);
  C_DATA    = lcd.color565(0x4F,0xA3,0xE3);
  C_DATA_HI = lcd.color565(0x8A,0xCD,0xF7);
  C_GOOD    = lcd.color565(0x2F,0xBF,0x71);
  C_WARN    = lcd.color565(0xF2,0xB0,0x1E);
  C_CRIT    = lcd.color565(0xE5,0x48,0x4D);
  C_TRACK   = lcd.color565(0x16,0x32,0x4C);
}

/* ---------------- ฟอนต์ ---------------- */
#define F_HERO  &fonts::FreeSansBold24pt7b
#define F_BIG   &fonts::FreeSansBold18pt7b
#define F_MED   &fonts::FreeSansBold12pt7b
#define F_BODY  &fonts::FreeSans9pt7b
#define F_TINY  &fonts::Font0

/* ป้ายข้อความ — สลับไทย/อังกฤษด้วย HAS_THAI_FONT (ฟอนต์ ASCII ในตัวไม่มีสระไทย) */
#if HAS_THAI_FONT
  #define T(th, en) th
#else
  #define T(th, en) en
#endif

/* ---------------- ตัวช่วย ---------------- */
static inline void fmtBaht(long v, char *buf, size_t cap) { tsp_fmtBaht(v, buf, cap); }

static void roundBar(LGFX_Sprite &s, int x, int y, int w, int h, uint16_t col) {
  if (w < h) w = h;
  s.fillRoundRect(x, y, w, h, h / 2, col);
}

static void progressBar(LGFX_Sprite &s, int x, int y, int w, int h, float pct, bool milestones) {
  roundBar(s, x, y, w, h, C_TRACK);
  float scale = milestones ? 120.0f : 100.0f;
  if (pct > 0) {
    float f = pct / scale; if (f > 1) f = 1;
    roundBar(s, x, y, (int)(w * f), h, C_DATA);
  }
  if (milestones) {
    const int ms[4] = {50, 80, 100, 110};
    for (int i = 0; i < 4; i++) {
      int mx = x + (int)(w * ms[i] / 120.0f);
      s.fillRect(mx, y + 2, 1, h - 4, (pct >= ms[i]) ? C_BG : C_INK3);
    }
  }
}

static void statusChip(LGFX_Sprite &s, int x, int y, char st) {
  uint16_t col = C_INK3; const char *txt = T("ไม่มีเป้า", "NO TARGET"); const char *ic = "-";
  if (st == 'G') { col = C_GOOD; txt = T("ตามเป้า",   "ON TRACK"); ic = "+"; }
  if (st == 'Y') { col = C_WARN; txt = T("เฉียดเป้า", "CLOSE");    ic = "!"; }
  if (st == 'R') { col = C_CRIT; txt = T("ต่ำกว่าเป้า","BEHIND");   ic = "v"; }
  s.setFont(F_BODY); s.setTextDatum(middle_left);
  char line[28]; snprintf(line, sizeof(line), "%s %s", ic, txt);
  int w = s.textWidth(line) + 16;
  s.fillRoundRect(x, y, w, 18, 9, col);
  s.setTextColor(C_BG, col);
  s.drawString(line, x + 8, y + 9);
}

static void pageDots(LGFX_Sprite &s, int idx) {
  for (int i = 0; i < PAGE_COUNT; i++) {
    int cx = W / 2 - (PAGE_COUNT - 1) * 5 + i * 10;
    s.fillCircle(cx, H - 7, i == idx ? 3 : 2, i == idx ? C_DATA_HI : C_LINE);
  }
}

static void header(LGFX_Sprite &s, const MonitorData &d, const char *title) {
  s.fillRect(0, 0, W, 24, C_HEAD);
  s.fillRect(0, 24, W, 1, C_LINE);
  s.setTextDatum(middle_left);
  s.setFont(F_BODY);
  s.setTextColor(C_DATA_HI, C_HEAD);
  s.drawString(d.branch, 8, 12);
  int x = 12 + s.textWidth(d.branch);
#if HAS_THAI_FONT
  s.setTextColor(C_INK, C_HEAD);
  s.drawString(d.branchName, x, 12);
  x += s.textWidth(d.branchName) + 8;
#endif
  s.setTextColor(C_INK3, C_HEAD);
  s.drawString(title, x, 12);
  s.setTextDatum(middle_right);
  s.setTextColor(C_INK, C_HEAD);
  const char *clock = strlen(d.ts) >= 5 ? d.ts + strlen(d.ts) - 5 : "--:--";
  s.drawString(clock, W - 8, 12);
}

/* ---------------- หน้า 1 : วันนี้ ---------------- */
static void page1(LGFX_Sprite &s, const MonitorData &d) {
  s.fillSprite(C_BG);
  header(s, d, T("วันนี้", "TODAY"));
  char buf[24];

  s.setTextDatum(top_left); s.setFont(F_BODY); s.setTextColor(C_INK2, C_BG);
  s.drawString(T("ยอดขายวันนี้", "REVENUE TODAY"), 14, 40);

  fmtBaht(d.dayRev, buf, sizeof(buf));
  s.setFont(F_HERO); s.setTextColor(C_INK, C_BG);
  s.setTextDatum(bottom_left);
  s.drawString(buf, 14, 106);
  int heroW = s.textWidth(buf);
  s.setFont(F_BODY); s.setTextColor(C_INK2, C_BG);
  s.drawString(T("บาท", "THB"), 14 + heroW + 8, 106);

  progressBar(s, 14, 128, 292, 16, d.dayPct, false);

  s.setTextDatum(top_left); s.setFont(F_BODY); s.setTextColor(C_INK2, C_BG);
  fmtBaht(d.dayTgt, buf, sizeof(buf));
  char line[40]; snprintf(line, sizeof(line), "%s %s", T("เป้าวัน", "DAY TARGET"), buf);
  s.drawString(line, 14, 152);
  s.setTextDatum(top_right); s.setTextColor(C_INK, C_BG);
  if (d.dayPct < 0) s.drawString(T("ไม่มีเป้า", "NO TARGET"), 306, 152);
  else { snprintf(line, sizeof(line), "%d%%", (int)(d.dayPct + 0.5f)); s.drawString(line, 306, 152); }

  s.setTextDatum(top_left); s.setTextColor(C_INK3, C_BG);
  snprintf(line, sizeof(line), "%d %s . %d %s", d.dayLogs, T("รายการ", "logs"), d.machines, T("ตู้", "machines"));
  s.drawString(line, 14, 176);
  if (d.warn[0]) {
    s.setTextColor(C_WARN, C_BG);
    snprintf(line, sizeof(line), "! %s", d.warn);
    s.drawString(line, 14, 198);
  }
  pageDots(s, 0);
}

/* ---------------- หน้า 2 : รอบเก็บเงิน ---------------- */
static void page2(LGFX_Sprite &s, const MonitorData &d) {
  s.fillSprite(C_BG);
  header(s, d, T("รอบเก็บเงิน", "COLLECTION"));
  char buf[24], line[48];

  s.setTextDatum(top_left); s.setFont(F_BODY); s.setTextColor(C_INK2, C_BG);
  snprintf(line, sizeof(line), "%s %d", T("เช็คพอยต์วันนี้", "CHECKPOINTS TODAY"), d.cpCount);
  s.drawString(line, 14, 40);

  const int x0 = 26, x1 = 294, y = 92;
  s.fillRect(x0, y, x1 - x0, 2, C_LINE);
  if (d.cpCount > 1) s.fillRect(x0, y, x1 - x0, 2, C_DATA);

  if (d.cpCount == 0) {
    s.setTextDatum(middle_center); s.setTextColor(C_INK3, C_BG);
    s.drawString(T("ยังไม่มีรอบเก็บเงินวันนี้", "NO COLLECTION YET"), W / 2, y + 8);
  }
  for (int i = 0; i < d.cpCount; i++) {
    int cx = tsp_cpX(i, d.cpCount, x0, x1);
    s.fillCircle(cx, y + 1, 7, C_BG);
    s.fillCircle(cx, y + 1, 5, C_DATA);
    s.setTextDatum(bottom_center); s.setTextColor(C_INK, C_BG);
    s.drawString(d.cp[i].t, cx, y - 8);
    fmtBaht(d.cp[i].amount, buf, sizeof(buf));
    s.setTextDatum(top_center); s.setTextColor(C_INK2, C_BG);
    s.drawString(buf, cx, y + 14);
  }

  const char *labels[3] = { T("เงินสดที่นับได้", "CASH COUNTED"),
                            T("รายได้จากมิเตอร์", "METER REVENUE"),
                            T("ส่วนต่าง", "DIFFERENCE") };
  long vals[3] = { d.cash, d.dayRev, d.cashDiff };
  for (int i = 0; i < 3; i++) {
    int yy = 140 + i * 22;
    s.setTextDatum(top_left); s.setTextColor(C_INK2, C_BG);
    s.drawString(labels[i], 14, yy);
    fmtBaht(vals[i], buf, sizeof(buf));
    if (i == 2 && vals[2] > 0) { snprintf(line, sizeof(line), "+%s", buf); strlcpy(buf, line, sizeof(buf)); }
    s.setTextDatum(top_right);
    s.setTextColor((i == 2 && d.cashFlag) ? C_CRIT : C_INK, C_BG);
    s.drawString(buf, 306, yy);
  }
  if (d.cashFlag) {
    s.setTextDatum(top_left); s.setTextColor(C_CRIT, C_BG);
    s.drawString(T("! ส่วนต่างเกิน 3% - ตรวจการนับเงิน", "! CASH GAP > 3% - RECHECK"), 14, 210);
  }
  pageDots(s, 1);
}

/* ---------------- หน้า 3 : 7 วันล่าสุด ---------------- */
static void page3(LGFX_Sprite &s, const MonitorData &d) {
  s.fillSprite(C_BG);
  header(s, d, T("7 วันล่าสุด", "LAST 7 DAYS"));
  char buf[24], line[40];

  long mx = 1;
  for (int i = 0; i < 7; i++) if (d.week[i] > mx) mx = d.week[i];

  const int x0 = 16, bw = 38, gap = 4, base = 158, hMax = 84;
  static const char *dow[7] = { "MO", "TU", "WE", "TH", "FR", "SA", "SU" };
  struct tm tm; bool haveTime = getLocalTime(&tm, 5);
  int todayDow = haveTime ? (tm.tm_wday + 6) % 7 : 6;      // 0 = จันทร์

  for (int i = 0; i < 7; i++) {
    int h = (int)(hMax * (float)d.week[i] / mx); if (h < 2) h = 2;
    int x = x0 + i * (bw + gap);
    bool isToday = (i == 6);
    s.fillRoundRect(x, base - h, bw, h, 4, isToday ? C_DATA_HI : C_DATA);
    s.setTextDatum(top_center); s.setFont(F_TINY); s.setTextColor(C_INK3, C_BG);
    s.drawString(dow[(todayDow - (6 - i) + 7) % 7], x + bw / 2, base + 5);
    if (isToday || d.week[i] == mx) {
      s.setFont(F_TINY); s.setTextColor(C_INK, C_BG);
      s.setTextDatum(bottom_center);
      snprintf(line, sizeof(line), "%ldk", (long)((d.week[i] + 500) / 1000));
      s.drawString(line, x + bw / 2, base - h - 3);
    }
  }
  s.fillRect(x0, base, W - 2 * x0, 1, C_LINE);

  s.setFont(F_BODY); s.setTextDatum(top_left); s.setTextColor(C_INK2, C_BG);
  s.drawString(T("รวมสัปดาห์นี้", "THIS WEEK"), 16, 182);
  fmtBaht(d.weekSum, buf, sizeof(buf));
  s.setFont(F_BIG); s.setTextColor(C_INK, C_BG);
  s.drawString(buf, 16, 200);

  if (d.hasWeekChg) {
    bool up = d.weekChg >= 0;
    s.setFont(F_MED); s.setTextDatum(top_right); s.setTextColor(up ? C_GOOD : C_CRIT, C_BG);
    snprintf(line, sizeof(line), "%s%.1f%%", up ? "+" : "-", fabsf(d.weekChg));
    s.drawString(line, 306, 204);
    fmtBaht(d.weekPrev, buf, sizeof(buf));
    s.setFont(F_TINY); s.setTextColor(C_INK3, C_BG);
    snprintf(line, sizeof(line), "%s %s", T("สัปดาห์ก่อน", "PREV"), buf);
    s.drawString(line, 306, 186);
  }
  pageDots(s, 2);
}

/* ---------------- หน้า 4 : เดือน + คาดการณ์ ---------------- */
static void page4(LGFX_Sprite &s, const MonitorData &d) {
  s.fillSprite(C_BG);
  char title[24]; snprintf(title, sizeof(title), "%s %s", T("เดือน", "MONTH"), d.period);
  header(s, d, title);
  char buf[24], line[52];

  s.setTextDatum(top_left); s.setFont(F_BODY); s.setTextColor(C_INK2, C_BG);
  s.drawString(T("ยอดสะสม", "MONTH TO DATE"), 14, 38);

  fmtBaht(d.mRev, buf, sizeof(buf));
  s.setFont(F_BIG); s.setTextColor(C_INK, C_BG);
  s.setTextDatum(bottom_left);
  s.drawString(buf, 14, 88);

  fmtBaht(d.mTgt, buf, sizeof(buf));
  s.setFont(F_TINY); s.setTextDatum(top_left); s.setTextColor(C_INK3, C_BG);
  snprintf(line, sizeof(line), "%s %s%s", T("เป้า", "TARGET"), buf, d.mShared ? T(" (เป้ารวมหน่วย)", " (SHARED)") : "");
  s.drawString(line, 14, 98);

  progressBar(s, 14, 116, 292, 20, d.mPct, true);

  s.setFont(F_MED); s.setTextDatum(top_right); s.setTextColor(C_INK, C_BG);
  if (d.mPct < 0) s.drawString(T("ไม่มีเป้า", "NO TARGET"), 306, 144);
  else { snprintf(line, sizeof(line), "%.1f%%", d.mPct); s.drawString(line, 306, 144); }

  s.setFont(F_TINY); s.setTextDatum(top_left); s.setTextColor(C_INK3, C_BG);
  snprintf(line, sizeof(line), "%s %d/%d . milestone %d%%", T("ผ่านมา", "DAY"), d.mElapsed, d.mDays, d.mMilestone);
  s.drawString(line, 14, 150);

  s.setFont(F_BODY); s.setTextColor(C_INK2, C_BG);
  s.drawString(T("คาดการณ์ปิดเดือน", "FORECAST"), 14, 176);
  fmtBaht(d.mFore, buf, sizeof(buf));
  s.setFont(F_BIG); s.setTextColor(C_INK, C_BG);
  s.drawString(buf, 14, 194);
  statusChip(s, 196, 196, d.mStat);
  pageDots(s, 3);
}

/* ---------------- หน้า 5 : ตารางงานสัปดาห์นี้ ---------------- */
static void page5(LGFX_Sprite &s, const MonitorData &d) {
  s.fillSprite(C_BG);
  header(s, d, T("ตารางงานสัปดาห์นี้", "THIS WEEK SHIFTS"));

  if (d.schedRows == 0) {
    s.setFont(F_BODY); s.setTextDatum(middle_center); s.setTextColor(C_INK3, C_BG);
    s.drawString(T("ยังไม่ได้เชื่อมข้อมูลตารางกะ", "SHIFT DATA NOT LINKED"), W / 2, 112);
    s.setFont(F_TINY);
    s.drawString("set SCHEDULE_URL in Script Properties", W / 2, 136);
    pageDots(s, 4);
    return;
  }

  const int nameW = 78, avail = H - 42 - 26;
  const int cw = (W - nameW - 16) / 7;
  int rh = avail / d.schedRows; if (rh > 30) rh = 30;
  int top = 42 + (avail - rh * d.schedRows) / 2;
  static const char *dow[7] = { "MO", "TU", "WE", "TH", "FR", "SA", "SU" };

  s.setFont(F_TINY); s.setTextDatum(bottom_center); s.setTextColor(C_INK2, C_BG);
  for (int i = 0; i < 7; i++) s.drawString(dow[i], 8 + nameW + cw * i + cw / 2, top - 4);
  s.fillRect(8, top - 2, W - 16, 1, C_LINE);

  for (int r = 0; r < d.schedRows; r++) {
    int y = top + r * rh;
    s.setFont(F_BODY); s.setTextDatum(middle_left); s.setTextColor(C_INK, C_BG);
    s.drawString(d.sched[r].name, 10, y + rh / 2);
    for (int c = 0; c < 7; c++) {
      const char *cell = d.sched[r].cell[c];
      bool on = cell[0] && strcasecmp(cell, "OFF") != 0 && strcmp(cell, "-") != 0;
      int x = 8 + nameW + cw * c;
      s.fillRoundRect(x + 2, y + 3, cw - 4, rh - 8, 4, on ? C_DATA : C_TRACK);
      s.setFont(F_TINY); s.setTextDatum(middle_center);
      s.setTextColor(on ? C_BG : C_INK3, on ? C_DATA : C_TRACK);
      s.drawString(cell[0] ? cell : "-", x + cw / 2, y + rh / 2);
    }
  }
  pageDots(s, 4);
}

/* ---------------- fade ---------------- */
static inline uint16_t rd(uint16_t v) { return s_swapped ? __builtin_bswap16(v) : v; }
static inline uint16_t wr(uint16_t v) { return s_swapped ? __builtin_bswap16(v) : v; }

static inline uint16_t blend565(uint16_t a, uint16_t b, uint8_t f) { return tsp_blend565(a, b, f); }

static void drawPageTo(LGFX_Sprite &s, int page, const MonitorData &d) {
  switch (page) {
    case 0: page1(s, d); break;
    case 1: page2(s, d); break;
    case 2: page3(s, d); break;
    case 3: page4(s, d); break;
    default: page5(s, d); break;
  }
}

void ui_render(int page, const MonitorData &d) {
  drawPageTo(spCur, page, d);
  spCur.pushSprite(0, 0);
  s_lastPage = page;
}

void ui_fadeTo(int page, const MonitorData &d) {
  if (s_lastPage < 0) { ui_render(page, d); return; }
  drawPageTo(spNext, page, d);

  uint16_t *a = (uint16_t *)spCur.getBuffer();
  uint16_t *b = (uint16_t *)spNext.getBuffer();
  uint16_t *o = (uint16_t *)spMix.getBuffer();
  const size_t n = (size_t)W * H;

  if (a && b && o) {
    for (int step = 1; step <= FADE_STEPS; step++) {
      uint8_t f = (uint8_t)(32 * step / FADE_STEPS);
      for (size_t i = 0; i < n; i++) o[i] = wr(blend565(rd(a[i]), rd(b[i]), f));
      spMix.pushSprite(0, 0);
    }
  }
  spNext.pushSprite(0, 0);
  // เก็บภาพล่าสุดไว้เป็นต้นทางของ fade รอบหน้า (คัดลอกบัฟเฟอร์ เร็วกว่าวาดใหม่)
  if (a && b) memcpy(a, b, n * sizeof(uint16_t));
  else drawPageTo(spCur, page, d);
  s_lastPage = page;
}

void ui_splash(const char *line1, const char *line2) {
  spCur.fillSprite(C_BG);
  spCur.setTextDatum(middle_center);
  spCur.setFont(F_MED); spCur.setTextColor(C_DATA_HI, C_BG);
  spCur.drawString("TOY STATION PLUS+", W / 2, 96);
  spCur.setFont(F_BODY); spCur.setTextColor(C_INK2, C_BG);
  if (line1) spCur.drawString(line1, W / 2, 130);
  spCur.setFont(F_TINY); spCur.setTextColor(C_INK3, C_BG);
  if (line2) spCur.drawString(line2, W / 2, 154);
  spCur.pushSprite(0, 0);
  s_lastPage = -1;
}

void ui_backlightForHour(int hour) {
  lcd.setBrightness(tsp_isNight(hour, NIGHT_START_HOUR, NIGHT_END_HOUR) ? BL_NIGHT : BL_DAY);
}

void ui_init() {
  lcd.init();
  lcd.setRotation(TFT_ROTATION);
  lcd.setBrightness(BL_DAY);
  initColors();
  lcd.fillScreen(C_BG);

  LGFX_Sprite *sps[3] = { &spCur, &spNext, &spMix };
  for (int i = 0; i < 3; i++) {
    sps[i]->setPsram(true);
    sps[i]->setColorDepth(16);
    if (!sps[i]->createSprite(W, H)) {
      lcd.setTextColor(C_CRIT, C_BG);
      lcd.drawString("sprite alloc failed - enable PSRAM", 10, 110);
    }
    sps[i]->fillSprite(C_BG);
  }
  // ตรวจว่า sprite เก็บ RGB565 สลับไบต์หรือไม่ (ต่างกันตามเวอร์ชัน LovyanGFX)
  spMix.drawPixel(0, 0, lcd.color565(0xFF, 0x00, 0x00));     // แดงล้วน = 0xF800
  uint16_t *probe = (uint16_t *)spMix.getBuffer();
  s_swapped = probe && (probe[0] == 0x00F8);
  spMix.fillSprite(C_BG);
}
