/**
 * panel_probe.cpp — หาว่าจอที่ต่ออยู่เป็น driver ตัวไหน
 * ------------------------------------------------------------------
 * แฟลชครั้งเดียวแล้วนั่งดู จอจะวนแสดง 4 โหมด โหมดละ 6 วินาที
 *
 *   1  ILI9341  inv:OFF          3  ST7789   inv:OFF
 *   2  ILI9341  inv:ON           4  ST7789   inv:ON
 *
 * รอบไหนที่ "กรอบขาวเต็มขอบ · แถบสีเรียง R G B ถูกสี · พื้นหลังดำสนิท"
 * = ค่านั้นคือค่าที่ต้องใส่ใน config.h
 *
 *   pio run -e panel-probe -t upload && pio device monitor
 *
 * ขาที่ใช้: ถ้ามี src/config.h อยู่แล้วจะอ่านจากไฟล์นั้น
 * ถ้ายังไม่มี จะใช้ค่าของบอร์ด Router_V2.0 ที่ยืนยันแล้วด้านล่าง
 * ------------------------------------------------------------------
 */
#include <Arduino.h>
#include <LovyanGFX.hpp>

#if __has_include("config.h")
  #include "config.h"
#endif

#ifndef PIN_TFT_MOSI
  #define PIN_TFT_MOSI  6
  #define PIN_TFT_SCLK  7
  #define PIN_TFT_CS    8
  #define PIN_TFT_DC    12
  #define PIN_TFT_RST   5
  #define PIN_TFT_BL    9
#endif
#ifndef TFT_ROTATION
  #define TFT_ROTATION  1        /* 1 = แนวนอน 320x240 เหมือนหน้าจอจริง */
#endif
#ifndef TFT_SPI_HZ
  #define TFT_SPI_HZ    40000000
#endif
#ifndef BL_ON_HIGH
  #define BL_ON_HIGH    1
#endif

static const uint32_t HOLD_MS = 6000;

static lgfx::Bus_SPI        bus;
static lgfx::Light_PWM      light;
static lgfx::Panel_ILI9341  panelIli;
static lgfx::Panel_ST7789   panelSt;
static lgfx::LGFX_Device    lcd;

struct Mode { const char *name; bool invert; };
static const Mode MODES[4] = {
  { "ILI9341", false }, { "ILI9341", true },
  { "ST7789",  false }, { "ST7789",  true }
};

static void setupBus() {
  auto c = bus.config();
  c.spi_host    = SPI2_HOST;
  c.spi_mode    = 0;
  c.freq_write  = TFT_SPI_HZ;
  c.freq_read   = 16000000;
  c.spi_3wire   = false;
  c.use_lock    = true;
  c.dma_channel = SPI_DMA_CH_AUTO;
  c.pin_sclk    = PIN_TFT_SCLK;
  c.pin_mosi    = PIN_TFT_MOSI;
  c.pin_miso    = -1;                 // สาย 8 เส้นไม่มี MISO
  c.pin_dc      = PIN_TFT_DC;
  bus.config(c);

  auto l = light.config();
  l.pin_bl      = PIN_TFT_BL;
  l.invert      = (BL_ON_HIGH ? false : true);
  l.freq        = 12000;
  l.pwm_channel = 7;
  light.config(l);
}

/** ผูกพาเนลตัวที่เลือกเข้ากับบัส แล้ว init ใหม่ */
static void applyMode(int i) {
  lgfx::Panel_Device *p = (i < 2) ? static_cast<lgfx::Panel_Device *>(&panelIli)
                                  : static_cast<lgfx::Panel_Device *>(&panelSt);
  auto c = p->config();
  c.pin_cs       = PIN_TFT_CS;
  c.pin_rst      = PIN_TFT_RST;
  c.pin_busy     = -1;
  c.panel_width  = 240;
  c.panel_height = 320;
  c.offset_x     = 0;
  c.offset_y     = 0;
  c.offset_rotation = 0;
  c.readable     = false;
  c.invert       = MODES[i].invert;
  c.rgb_order    = false;
  c.dlen_16bit   = false;
  c.bus_shared   = false;
  p->config(c);
  p->setBus(&bus);
  p->setLight(&light);

  lcd.setPanel(p);
  lcd.init();
  lcd.setRotation(TFT_ROTATION);
  lcd.setBrightness(255);
}

static void drawTestCard(int i) {
  const int W = lcd.width(), H = lcd.height();
  lcd.fillScreen(TFT_BLACK);

  /* กรอบขาว 2px ชิดขอบจอพอดี — ถ้าพาเนลผิดตัว กรอบจะขาด เลื่อน หรือไม่ขึ้นเลย */
  lcd.drawRect(0, 0, W, H, TFT_WHITE);
  lcd.drawRect(1, 1, W - 2, H - 2, TFT_WHITE);

  /* หมุดมุม 4 มุม — เช็คว่าภาพไม่ถูกครอบตัด */
  const int m = 10;
  lcd.fillRect(2, 2, m, m, TFT_YELLOW);
  lcd.fillRect(W - m - 2, 2, m, m, TFT_YELLOW);
  lcd.fillRect(2, H - m - 2, m, m, TFT_YELLOW);
  lcd.fillRect(W - m - 2, H - m - 2, m, m, TFT_YELLOW);

  /* เลขโหมดตัวใหญ่ + ชื่อ driver */
  lcd.setTextDatum(top_left);
  lcd.setFont(&fonts::FreeSansBold24pt7b);
  lcd.setTextColor(TFT_WHITE, TFT_BLACK);
  char num[2] = { (char)('1' + i), 0 };
  lcd.drawString(num, 16, 18);

  lcd.setFont(&fonts::FreeSansBold12pt7b);
  lcd.drawString(MODES[i].name, 56, 24);
  lcd.setFont(&fonts::FreeSans9pt7b);
  lcd.setTextColor(TFT_CYAN, TFT_BLACK);
  lcd.drawString(MODES[i].invert ? "invert ON" : "invert OFF", 56, 52);

  /* แถบสี 5 ช่อง — ตรวจทั้งความถูกของสีและลำดับ RGB/BGR */
  const uint16_t cols[5] = { TFT_RED, TFT_GREEN, TFT_BLUE, TFT_WHITE, TFT_BLACK };
  const char *labels[5]  = { "R", "G", "B", "W", "K" };
  const int barY = 90, barH = H - barY - 34, barW = (W - 24) / 5;
  for (int b = 0; b < 5; b++) {
    int x = 12 + b * barW;
    lcd.fillRect(x, barY, barW - 2, barH, cols[b]);
    if (b == 4) lcd.drawRect(x, barY, barW - 2, barH, TFT_WHITE);   // ช่องดำต้องมีขอบถึงจะเห็น
    lcd.setTextDatum(top_center);
    lcd.setTextColor(TFT_WHITE, TFT_BLACK);
    lcd.drawString(labels[b], x + barW / 2, barY + barH + 6);
  }

  lcd.setTextDatum(bottom_left);
  lcd.setFont(&fonts::Font0);
  lcd.setTextColor(TFT_DARKGREY, TFT_BLACK);
  lcd.drawString("frame full + R G B correct + K black = this mode is right", 12, H - 4);
}

void setup() {
  Serial.begin(115200);
  delay(300);
  Serial.println();
  Serial.println("=== TSP panel probe ===");
  Serial.printf("pins: MOSI=%d SCLK=%d CS=%d DC=%d RST=%d BL=%d  rotation=%d\n",
                PIN_TFT_MOSI, PIN_TFT_SCLK, PIN_TFT_CS, PIN_TFT_DC, PIN_TFT_RST, PIN_TFT_BL, TFT_ROTATION);
  Serial.println("วนโหมดละ 6 วินาที — จำเลขโหมดที่ภาพถูกต้องไว้");
  setupBus();
}

void loop() {
  for (int i = 0; i < 4; i++) {
    Serial.printf("[%d/4] %s  invert=%s\n", i + 1, MODES[i].name, MODES[i].invert ? "ON" : "OFF");
    applyMode(i);
    drawTestCard(i);
    delay(HOLD_MS);
  }
}
