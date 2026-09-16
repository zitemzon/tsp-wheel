/**
 * ทดสอบฟังก์ชันคำนวณของฝั่งจอบนเครื่อง PC (ไม่ต้องมีบอร์ด)
 *   g++ -std=c++17 -I../firmware/src ui_util_test.cpp -o /tmp/ui_util_test && /tmp/ui_util_test
 */
#include "ui_util.h"
#include <cstdio>
#include <cstring>

static int pass = 0, fail = 0;
static void eqs(const char *label, const char *got, const char *want) {
  if (strcmp(got, want) == 0) { pass++; printf("  [ok] %s\n", label); }
  else { fail++; printf("  [NG] %s -> \"%s\" (ควรได้ \"%s\")\n", label, got, want); }
}
static void eqi(const char *label, long got, long want) {
  if (got == want) { pass++; printf("  [ok] %s\n", label); }
  else { fail++; printf("  [NG] %s -> %ld (ควรได้ %ld)\n", label, got, want); }
}

int main() {
  char b[24];
  printf("\n[1] จัดรูปตัวเลขเงิน\n");
  tsp_fmtBaht(0, b, sizeof(b));        eqs("0", b, "0");
  tsp_fmtBaht(999, b, sizeof(b));      eqs("999", b, "999");
  tsp_fmtBaht(1000, b, sizeof(b));     eqs("1000", b, "1,000");
  tsp_fmtBaht(12450, b, sizeof(b));    eqs("12450", b, "12,450");
  tsp_fmtBaht(186500, b, sizeof(b));   eqs("186500", b, "186,500");
  tsp_fmtBaht(1234567, b, sizeof(b));  eqs("1234567", b, "1,234,567");
  tsp_fmtBaht(-2650, b, sizeof(b));    eqs("-2650", b, "-2,650");
  char small[5];
  tsp_fmtBaht(1234567, small, sizeof(small)); /* ต้องไม่ล้นบัฟเฟอร์ */
  eqi("บัฟเฟอร์เล็กไม่ overflow", (long)strlen(small), 4);

  printf("\n[2] ผสมสี RGB565\n");
  const uint16_t RED = 0xF800, BLUE = 0x001F, WHITE = 0xFFFF, BLACK = 0x0000;
  eqi("f=0 ได้สีต้นทาง", tsp_blend565(RED, BLUE, 0), RED);
  eqi("f=32 ได้สีปลายทาง", tsp_blend565(RED, BLUE, 32), BLUE);
  eqi("ขาว->ขาว คงที่", tsp_blend565(WHITE, WHITE, 16), WHITE);
  eqi("ดำ->ดำ คงที่", tsp_blend565(BLACK, BLACK, 16), BLACK);
  uint16_t mid = tsp_blend565(BLACK, WHITE, 16);
  eqi("ดำ->ขาว ครึ่งทาง = เทากลาง", mid, (uint16_t)((15 << 11) | (31 << 5) | 15));
  bool mono = true;
  uint16_t prev = 0;
  for (int f = 0; f <= 32; f++) { uint16_t v = tsp_blend565(BLACK, WHITE, f) >> 11; if (f && v < prev) mono = false; prev = v; }
  eqi("ไล่เฉดขึ้นเรื่อย ๆ ไม่ย้อนกลับ", mono ? 1 : 0, 1);

  printf("\n[3] ช่วงกลางคืน (22:00-08:00)\n");
  eqi("23 น. = กลางคืน", tsp_isNight(23, 22, 8), 1);
  eqi("02 น. = กลางคืน", tsp_isNight(2, 22, 8), 1);
  eqi("08 น. = กลางวัน", tsp_isNight(8, 22, 8), 0);
  eqi("14 น. = กลางวัน", tsp_isNight(14, 22, 8), 0);
  eqi("ช่วงไม่ข้ามเที่ยงคืน 1-5: 3 น.", tsp_isNight(3, 1, 5), 1);

  printf("\n[4] ตำแหน่งจุดเช็คพอยต์ (เส้น 26..294)\n");
  eqi("1 จุด อยู่กลางเส้น", tsp_cpX(0, 1, 26, 294), 160);
  eqi("3 จุด: จุดแรกชิดซ้าย", tsp_cpX(0, 3, 26, 294), 26);
  eqi("3 จุด: จุดกลาง", tsp_cpX(1, 3, 26, 294), 160);
  eqi("3 จุด: จุดท้ายชิดขวา", tsp_cpX(2, 3, 26, 294), 294);

  printf("\n────────────────────────\nผ่าน %d / ไม่ผ่าน %d\n", pass, fail);
  return fail ? 1 : 0;
}
