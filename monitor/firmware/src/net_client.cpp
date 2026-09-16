#include "config.h"
#include "net_client.h"
#include <WiFi.h>
#include <WiFiClientSecure.h>
#include <HTTPClient.h>
#include <ArduinoJson.h>
#include <time.h>

static int s_httpCode = 0;
int net_lastHttpCode() { return s_httpCode; }

bool net_isConnected() { return WiFi.status() == WL_CONNECTED; }

bool net_begin() {
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  uint32_t t0 = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - t0 < 20000) delay(250);
  return net_isConnected();
}

bool net_syncTime() {
  // ICT-7 = UTC+7 ไม่มี DST
  configTzTime("ICT-7", "pool.ntp.org", "time.google.com");
  struct tm tm;
  for (int i = 0; i < 20; i++) {
    if (getLocalTime(&tm, 500) && tm.tm_year > 120) return true;
    delay(200);
  }
  return false;
}

/** แปลงรหัส warning จาก API เป็นข้อความสั้นที่จอแสดงได้ */
static void appendWarn(char *dst, size_t cap, const char *code) {
  const char *txt = code;
  if (strncmp(code, "first-reading", 13) == 0)      txt = HAS_THAI_FONT ? "อ่านมิเตอร์ครั้งแรก" : "first meter read";
  else if (strncmp(code, "meter-reset", 11) == 0)   txt = HAS_THAI_FONT ? "มิเตอร์รีเซ็ต" : "meter reset";
  else if (strncmp(code, "no-coin-price", 13) == 0) txt = HAS_THAI_FONT ? "ไม่มีราคาเหรียญ" : "no coin price";
  else if (strncmp(code, "no-target", 9) == 0)      txt = HAS_THAI_FONT ? "ยังไม่ตั้งเป้า" : "no target set";
  else if (strncmp(code, "target-shared", 13) == 0) txt = HAS_THAI_FONT ? "ใช้เป้ารวมกับสาขาคู่" : "shared target";
  else if (strncmp(code, "bad-date", 8) == 0)       txt = HAS_THAI_FONT ? "วันที่ผิดรูปแบบ" : "bad date rows";
  size_t len = strlen(dst);
  if (len && len + 3 < cap) { strlcat(dst, " . ", cap); len = strlen(dst); }
  strlcat(dst, txt, cap);
}

int net_fetch(MonitorData &out) {
  if (!net_isConnected() && !net_begin()) return NET_ERR_WIFI;

  String url = String(API_BASE) + "?k=" + API_TOKEN + "&b=" + API_BRANCH;
  if (out.valid && out.hash[0]) url += "&h=" + String(out.hash);

  WiFiClientSecure client;
  client.setInsecure();              // Apps Script เปลี่ยน CA เป็นระยะ — ใช้ pinning ไม่คุ้มกับจอ read-only
  client.setTimeout(10);

  HTTPClient http;
  http.setFollowRedirects(HTTPC_STRICT_FOLLOW_REDIRECTS);   // exec -> script.googleusercontent.com
  http.setTimeout(10000);
  if (!http.begin(client, url)) return NET_ERR_HTTP;

  s_httpCode = http.GET();
  if (s_httpCode != 200) { http.end(); return NET_ERR_HTTP; }

  JsonDocument doc;
  DeserializationError err = deserializeJson(doc, http.getStream());
  http.end();
  if (err) return NET_ERR_PARSE;

  if (doc["nc"].is<int>()) {                 // {"nc":1} = ข้อมูลไม่เปลี่ยน
    strlcpy(out.ts, doc["ts"] | out.ts, sizeof(out.ts));
    return NET_NO_CHANGE;
  }
  if (doc["e"].is<const char*>()) return NET_ERR_HTTP;   // {"e":"auth"} เป็นต้น

  MonitorData d;
  strlcpy(d.hash,       doc["h"]  | "", sizeof(d.hash));
  strlcpy(d.ts,         doc["ts"] | "", sizeof(d.ts));
  strlcpy(d.branch,     doc["b"]  | "", sizeof(d.branch));
  strlcpy(d.branchName, doc["bn"] | "", sizeof(d.branchName));
  d.machines = doc["nm"] | 0;

  JsonObject day = doc["d"];
  d.dayRev  = day["r"] | 0;
  d.dayTgt  = day["t"] | 0;
  d.dayPct  = day["p"].isNull() ? -1.0f : (float)(day["p"] | 0.0);
  d.dayLogs = day["n"] | 0;

  JsonObject c = doc["c"];
  d.cash     = c["cash"] | 0;
  d.cashDiff = c["df"]   | 0;
  d.cashFlag = (c["fl"] | 0) != 0;
  for (JsonObject p : c["l"].as<JsonArray>()) {
    if (d.cpCount >= 6) break;
    strlcpy(d.cp[d.cpCount].t, p["t"] | "--:--", sizeof(d.cp[0].t));
    d.cp[d.cpCount].amount   = p["a"] | 0;
    d.cp[d.cpCount].machines = p["n"] | 0;
    d.cpCount++;
  }

  JsonObject w = doc["w"];
  int i = 0;
  for (JsonVariant v : w["v"].as<JsonArray>()) { if (i < 7) d.week[i++] = v.as<long>(); }
  d.weekSum  = w["s"]  | 0;
  d.weekPrev = w["pv"] | 0;
  d.hasWeekChg = !w["ch"].isNull();
  d.weekChg  = d.hasWeekChg ? (float)(w["ch"] | 0.0) : 0.0f;

  JsonObject m = doc["m"];
  strlcpy(d.period, m["pd"] | "", sizeof(d.period));
  d.mRev = m["r"] | 0;  d.mTgt = m["t"] | 0;  d.mFore = m["f"] | 0;
  d.mPct = m["p"].isNull() ? -1.0f : (float)(m["p"] | 0.0);
  const char *st = m["fs"] | "N";
  d.mStat = st[0];
  d.mMilestone = m["ms"] | 0;
  d.mElapsed   = m["ed"] | 0;
  d.mDays      = m["dm"] | 30;
  d.mShared    = (m["sh"] | 0) != 0;

  if (doc["s"].is<JsonObject>()) {
    for (JsonArray row : doc["s"]["rows"].as<JsonArray>()) {
      if (d.schedRows >= 6) break;
      SchedRow &sr = d.sched[d.schedRows];
      int col = 0;
      for (JsonVariant v : row) {
        const char *sv = v | "";
        if (col == 0) strlcpy(sr.name, sv, sizeof(sr.name));
        else if (col <= 7) strlcpy(sr.cell[col - 1], sv, sizeof(sr.cell[0]));
        col++;
      }
      d.schedRows++;
    }
  }

  for (JsonVariant v : doc["wn"].as<JsonArray>()) appendWarn(d.warn, sizeof(d.warn), v | "");

  d.valid = true;
  out = d;
  return NET_UPDATED;
}
