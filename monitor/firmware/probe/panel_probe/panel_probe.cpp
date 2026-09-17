/**
 * panel_probe.cpp — วินิจฉัยจอ TFT: สายไฟหรือไลบรารีกันแน่
 * ------------------------------------------------------------------
 * แฟลชครั้งเดียวแล้วนั่งดู จะวน 2 ช่วงสลับกันไปเรื่อย ๆ
 *
 *   STAGE 1  RAW SPI — คุมขาเองทั้งหมด ไม่ผ่าน LovyanGFX เลย
 *            ถมสีแดง/เขียวเต็มจอ ด้วยชุด init ของ ILI9341 แล้วของ ST7789
 *
 *   STAGE 2  LovyanGFX 4 โหมด — ILI9341/ST7789 x invert ON/OFF
 *            วาด test card: กรอบขาว · มุมเหลือง · เลขโหมด · แถบสี R G B W K
 *
 * ── อ่านผลยังไง ──────────────────────────────────────────────
 *
 *   STAGE 1 ถมสีติด + STAGE 2 ไม่ขึ้น  ->  สายไฟถูกหมด ปัญหาอยู่ที่ LovyanGFX
 *   ทั้งสอง stage ไม่ขึ้นอะไรเลย        ->  ปัญหาฮาร์ดแวร์/สายไฟ ดู checklist ใน
 *                                          monitor/docs/03-hardware.md
 *   STAGE 2 ขึ้นภาพถูกต้อง              ->  จำเลขโหมดไปใส่ config.h ได้เลย
 *
 * ขาที่ใช้: ถ้ามี config.h อยู่ในโฟลเดอร์เดียวกันจะอ่านจากไฟล์นั้น
 * ถ้ายังไม่มี จะใช้ค่าของบอร์ด Router_V2.0 ด้านล่าง
 * ------------------------------------------------------------------
 */
#include <Arduino.h>
#include <SPI.h>
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
#ifndef BL_ON_HIGH
  #define BL_ON_HIGH    1
#endif

/* ทั้งสอง stage ใช้ความถี่เดียวกัน จะได้ไม่มีตัวแปรอื่นมาปน
   10 MHz ต่ำพอที่สายแพยาว ๆ ยังส่งได้ชัวร์ */
#define PROBE_SPI_HZ  10000000

static const uint32_t HOLD_MS     = 6000;   // LovyanGFX โหมดละกี่ ms
static const uint32_t RAW_HOLD_MS = 4000;   // RAW ถมสีละกี่ ms
static const int      PANEL_W     = 240;    // ขนาดจริงของพาเนล (ก่อนหมุน)
static const int      PANEL_H     = 320;

/*===================================================================
 * STAGE 1 : RAW SPI — ไม่พึ่ง LovyanGFX เลยแม้แต่บรรทัดเดียว
 *==================================================================*/
static SPIClass rawSpi(FSPI);
static bool rawBusOpen = false;

static void rawBegin() {
  pinMode(PIN_TFT_CS,  OUTPUT); digitalWrite(PIN_TFT_CS,  HIGH);
  pinMode(PIN_TFT_DC,  OUTPUT); digitalWrite(PIN_TFT_DC,  HIGH);
  pinMode(PIN_TFT_RST, OUTPUT); digitalWrite(PIN_TFT_RST, HIGH);
  rawSpi.begin(PIN_TFT_SCLK, -1, PIN_TFT_MOSI, -1);
  rawBusOpen = true;
}

static void rawEnd() {
  if (!rawBusOpen) return;
  rawSpi.end();
  rawBusOpen = false;
}

static void rawHwReset() {
  digitalWrite(PIN_TFT_RST, HIGH); delay(10);
  digitalWrite(PIN_TFT_RST, LOW);  delay(20);
  digitalWrite(PIN_TFT_RST, HIGH); delay(150);
}

static void rawCmd(uint8_t c) {
  rawSpi.beginTransaction(SPISettings(PROBE_SPI_HZ, MSBFIRST, SPI_MODE0));
  digitalWrite(PIN_TFT_DC, LOW);
  digitalWrite(PIN_TFT_CS, LOW);
  rawSpi.transfer(c);
  digitalWrite(PIN_TFT_CS, HIGH);
  rawSpi.endTransaction();
}

static void rawData(const uint8_t *d, size_t n) {
  rawSpi.beginTransaction(SPISettings(PROBE_SPI_HZ, MSBFIRST, SPI_MODE0));
  digitalWrite(PIN_TFT_DC, HIGH);
  digitalWrite(PIN_TFT_CS, LOW);
  rawSpi.writeBytes(d, n);
  digitalWrite(PIN_TFT_CS, HIGH);
  rawSpi.endTransaction();
}

static void rawCmdData(uint8_t c, uint8_t d) { rawCmd(c); rawData(&d, 1); }

/** ชุด init ขั้นต่ำที่ทั้ง ILI9341 และ ST7789 รับได้ ต่างกันแค่ INVOFF/INVON */
static void rawInitPanel(bool inversionOn) {
  rawHwReset();
  rawCmd(0x01); delay(150);            // SWRESET
  rawCmd(0x11); delay(120);            // SLPOUT
  rawCmdData(0x3A, 0x55);              // COLMOD = RGB565
  rawCmdData(0x36, 0x00);              // MADCTL = ปกติ
  rawCmd(inversionOn ? 0x21 : 0x20);   // INVON / INVOFF
  rawCmd(0x13); delay(10);             // NORON
  rawCmd(0x29); delay(120);            // DISPON
}

static void rawSetWindow(uint16_t x0, uint16_t y0, uint16_t x1, uint16_t y1) {
  uint8_t b[4];
  b[0] = x0 >> 8; b[1] = x0 & 0xFF; b[2] = x1 >> 8; b[3] = x1 & 0xFF;
  rawCmd(0x2A); rawData(b, 4);         // CASET
  b[0] = y0 >> 8; b[1] = y0 & 0xFF; b[2] = y1 >> 8; b[3] = y1 & 0xFF;
  rawCmd(0x2B); rawData(b, 4);         // RASET
  rawCmd(0x2C);                        // RAMWR
}

static void rawFill(uint16_t color) {
  rawSetWindow(0, 0, PANEL_W - 1, PANEL_H - 1);

  uint8_t buf[512];
  for (size_t i = 0; i < sizeof(buf); i += 2) {
    buf[i]     = (uint8_t)(color >> 8);
    buf[i + 1] = (uint8_t)(color & 0xFF);
  }
  const uint32_t total = (uint32_t)PANEL_W * PANEL_H * 2;
  const uint32_t rounds = total / sizeof(buf);

  rawSpi.beginTransaction(SPISettings(PROBE_SPI_HZ, MSBFIRST, SPI_MODE0));
  digitalWrite(PIN_TFT_DC, HIGH);
  digitalWrite(PIN_TFT_CS, LOW);
  for (uint32_t i = 0; i < rounds; i++) rawSpi.writeBytes(buf, sizeof(buf));
  digitalWrite(PIN_TFT_CS, HIGH);
  rawSpi.endTransaction();
}

static void stageRaw() {
  Serial.println();
  Serial.println("=== STAGE 1 : RAW SPI (ไม่ผ่าน LovyanGFX) ===");
  rawBegin();

  const struct { const char *name; bool inv; } inits[2] = {
    { "ILI9341-style (INVOFF)", false },
    { "ST7789-style  (INVON)",  true  }
  };
  const struct { const char *name; uint16_t color; } fills[2] = {
    { "แดง",  0xF800 },
    { "เขียว", 0x07E0 }
  };

  for (int i = 0; i < 2; i++) {
    Serial.printf("  init %s\n", inits[i].name);
    rawInitPanel(inits[i].inv);
    for (int f = 0; f < 2; f++) {
      Serial.printf("    ถมสี%s ... จอเปลี่ยนสีไหม\n", fills[f].name);
      rawFill(fills[f].color);
      delay(RAW_HOLD_MS);
    }
  }
  rawEnd();
  Serial.println("=== จบ STAGE 1 ===");
}

/*===================================================================
 * STAGE 2 : LovyanGFX
 *==================================================================*/
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
  c.freq_write  = PROBE_SPI_HZ;
  c.freq_read   = 8000000;
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
  c.panel_width  = PANEL_W;
  c.panel_height = PANEL_H;
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

static void stageLovyan() {
  Serial.println();
  Serial.println("=== STAGE 2 : LovyanGFX 4 โหมด ===");
  for (int i = 0; i < 4; i++) {
    Serial.printf("  [%d/4] %s  invert=%s\n", i + 1, MODES[i].name, MODES[i].invert ? "ON" : "OFF");
    applyMode(i);
    drawTestCard(i);
    delay(HOLD_MS);
  }
  bus.release();                      // คืนบัสให้ STAGE 1 ใช้ต่อในรอบหน้า
  Serial.println("=== จบ STAGE 2 ===");
}

/*===================================================================
 * STAGE 0 : รายงานสภาพแวดล้อม + ลูปหลัก
 *==================================================================*/
void setup() {
  Serial.begin(115200);
  delay(400);
  Serial.println();
  Serial.println("======================================");
  Serial.println("  TSP panel probe v2");
  Serial.println("======================================");
  Serial.printf("ขา : MOSI=%d SCLK=%d CS=%d DC=%d RST=%d BL=%d\n",
                PIN_TFT_MOSI, PIN_TFT_SCLK, PIN_TFT_CS, PIN_TFT_DC, PIN_TFT_RST, PIN_TFT_BL);
  Serial.printf("SPI: %d Hz (ทั้งสอง stage ใช้เท่ากัน) · rotation=%d\n", PROBE_SPI_HZ, TFT_ROTATION);
  Serial.printf("IDF: %s\n", esp_get_idf_version());
#ifdef ESP_ARDUINO_VERSION_MAJOR
  Serial.printf("Arduino core: %d.%d.%d\n",
                ESP_ARDUINO_VERSION_MAJOR, ESP_ARDUINO_VERSION_MINOR, ESP_ARDUINO_VERSION_PATCH);
#endif
#ifdef LGFX_VERSION_MAJOR
  Serial.printf("LovyanGFX: %d.%d.%d\n", LGFX_VERSION_MAJOR, LGFX_VERSION_MINOR, LGFX_VERSION_PATCH);
#endif
  Serial.printf("PSRAM: %u ไบต์  %s\n", (unsigned)ESP.getPsramSize(),
                ESP.getPsramSize() ? "" : "<-- เป็น 0 = ยังไม่ได้เปิด Tools > PSRAM = OPI PSRAM");
  Serial.println("--------------------------------------");

  setupBus();
  light.init(255);                    // เปิด backlight ไว้ตลอด ทั้งสอง stage
}

void loop() {
  stageRaw();
  stageLovyan();
}
