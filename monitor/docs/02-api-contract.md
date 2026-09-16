# 02 — สัญญา JSON ระหว่าง Apps Script กับจอ

## 1. Endpoint

```
GET {WEB_APP_URL}/exec?route=monitor&k=<token>&b=<branch>[&h=<hash เดิม>]
```

| พารามิเตอร์ | ค่า | ความหมาย |
| :-- | :-- | :-- |
| `route` | `monitor` | บังคับ — ไม่ใส่จะได้พฤติกรรมเดิม (redirect ไป Dashboard) |
| `k` | บังคับ | token ใน Script Property `MONITOR_TOKEN` ไม่ตรง → `{"e":"auth"}` |
| `b` | `TPS-03` / `ALL` | สาขาที่จอนี้แสดง |
| `h` | ไม่บังคับ | hash ที่จอถืออยู่ ถ้าตรงกับของใหม่ ตอบ `{"nc":1,"ts":"..."}` |

> **ทำไมไม่ใช้ ETag/304** — Apps Script Web App ตั้ง HTTP header เองไม่ได้
> `h` คำนวณจาก payload ทั้งก้อน **ยกเว้น `ts`** เวลาที่เดินไปเรื่อย ๆ จึงไม่ทำให้ hash เปลี่ยน

**ห้ามใส่ลง payload:** `Manager_Email`, `Recorded_By`, `LINE_Group_ID`

## 2. รูปแบบ payload

```jsonc
{
  "v": 1,
  "b": "TPS-03",
  "bn": "โรบินสันฉลอง",
  "nm": 24,                        // จำนวนตู้ Active ของสาขานี้
  "d": { "r": 12450, "t": 3033, "p": 410.5, "n": 12 },
  "c": {
    "l": [ { "t": "09:40", "a": 3200, "n": 6 } ],   // รอบเก็บเงิน สูงสุด 6 รอบ
    "cash": 9800, "df": -2650, "fl": 1              // fl=1 เมื่อส่วนต่างเกิน 3%
  },
  "w": { "v": [9800,11200,8600,13400,15100,12900,12450], "s": 83450, "pv": 76200, "ch": 9.5 },
  "m": {
    "pd": "2569-09", "r": 186500, "t": 91000, "p": 204.9,
    "f": 349687, "fs": "G", "ms": 110, "sh": 0,
    "cash": 180000, "ed": 16, "dm": 30
  },
  "s": { "hdr": ["จ","อ","พ","พฤ","ศ","ส","อา"], "rows": [["ไอริน","D","D","OFF","D","D","N","N"]] },
  "wn": ["rev-missing:2"],
  "h": "1a2b3c4d",
  "ts": "2026-09-16 14:32"
}
```

| คีย์ | ความหมาย |
| :-- | :-- |
| `d.r / d.t / d.p / d.n` | ยอดวันนี้ / เป้าวัน / % ของเป้า (`null` = ไม่มีเป้า) / จำนวน log |
| `c.l[].t / a / n` | เวลารอบ / เงินสดรวมรอบนั้น / จำนวนตู้ในรอบ |
| `w.v[]` | ยอด 7 วันเรียงเก่า→ใหม่ ตัวสุดท้าย = วันนี้ |
| `m.fs` | `G` ≥100% · `Y` 90–99.9% · `R` <90% · `N` ไม่มีเป้า |
| `m.sh` | 1 = สาขานี้ใช้เป้ารวมกับสาขาคู่ (TSP-06 / TSP-57) |
| `wn[]` | `code:count` — `rev-missing`, `cash-missing`, `bad-date`, `no-target`, `target-shared` |

ขนาด payload ปกติ ~1.2–1.8 KB (ไม่รวมตารางกะ)

## 3. สูตรคำนวณ — ยึดตามเอนจินเดิมทุกข้อ

| ค่า | สูตร | อ้างอิงของเดิม |
| :-- | :-- | :-- |
| รายได้ต่อแถว | คอลัมน์ `Revenue` ล้วน (ไม่คูณ `Meter_Diff × Coin_Price`) | `buildBranchData_` |
| รายได้รายวัน | `Σ Revenue` ของสาขานั้นในวันนั้น | `sumRevenue_` |
| แถวที่ต้องมีเงินสด | `Revenue_Method <> 'METER'` **หรือ** `Is_Collection = TRUE` | `Code.gs > shouldHaveCash_` |
| รอบเช็คพอยต์ | แถวที่มีเงินสด และ `Timestamp` ห่างกัน ≤ 45 นาที = รอบเดียวกัน | — (ของใหม่) |
| ส่วนต่างเงินสด | `Σ Cash_Collected − รายได้วันนั้น` ; ธงแดงเมื่อ `|ส่วนต่าง| / รายได้ > 3%` | — (ของใหม่) |
| ปี พ.ศ. | ปี > 2400 → ลบ 543 | `Code.gs > fixBuddhistYear_` |
| เป้าเดือน | แท็บ Target งวด `พ.ศ.-MM` (ยึดแถวท้ายสุด) ; fallback `R0 × seasonal_index` | TargetBot |
| เป้ารายวัน | `เป้าเดือน ÷ จำนวนวันในเดือน` | — (ของใหม่) |
| % ของเป้า | `ยอด ÷ เป้า × 100` (เป้า ≤ 0 → `null`) | — |
| คาดการณ์ปิดเดือน | `MTD ÷ วันที่ผ่านไป × จำนวนวันในเดือน` | — |
| milestone | ค่ามากสุดใน `MILESTONES` ที่ % ผ่านแล้ว | TargetBot `MILESTONES` |
| hash | FNV-1a 32-bit ของ payload (ไม่รวม `ts`) | — |

## 4. ติดตั้ง

ไฟล์ทั้งสองต้องอยู่ใน **โปรเจกต์ Apps Script เดียวกับ `TSP_Unified_Report.gs`**
(ผูกกับ TSP_Master_Database Ver.2) เพราะใช้ `table_()` `toDate_()` `branchNames_()` `SH` `C` `CFG` ร่วมกัน

1. วาง `TSP_Monitor_Calc.gs` + `TSP_Monitor_Route.gs` ในโปรเจกต์นั้น
2. ใน `TSP_Unified_Report.gs` เปลี่ยนชื่อ `doGet` เดิม:
   ```diff
   -function doGet() {
   +function TSP_dashboardRedirect_() {
      return HtmlService.createHtmlOutput( ... );   // เนื้อในไม่ต้องแก้
    }
   ```
   *ถ้าลืมข้อนี้ระบบยังทำงานได้ — router จะสร้าง redirect จาก `DASHBOARD_URL` ให้เอง*
3. แก้ token ใน `TSP_monSetup()` แล้วรัน 1 ครั้ง
4. รัน `TSP_test_columns()` → ต้องเจอครบ 8 คอลัมน์ และเห็นจำนวน METER/CAPSULE
5. รัน **`TSP_test_vs_report()`** → ต้องขึ้น `✅ ตรงกันทุกสาขา`
   (เทียบยอดจอกับ `sumRevenue_()` และ `lnTrend_()` — ถ้าไม่ตรง ห้าม deploy)
6. Deploy: **Manage deployments → ✏️ → Version: New version → Deploy**
   ⚠ ห้ามกด "New deployment" เพราะจะได้ URL ใหม่ แล้ว LINE webhook เดิมจะชี้ผิดที่
7. ทดสอบหลัง deploy 3 อย่าง:
   - เปิด `{WEB_APP_URL}/exec` เปล่า ๆ → ต้อง redirect ไป Dashboard เหมือนเดิม
   - ส่ง "ยอดขาย" เข้ากลุ่ม LINE → การ์ดต้องขึ้นปกติ
   - เปิด `?route=monitor&k=...&b=TPS-03` → ต้องได้ JSON
8. เอา URL ไปวางใน `monitor/apps-script/preview.html` ตรวจ layout ทั้ง 5 หน้า

## 5. เปลี่ยนสัญญาเมื่อไหร่ต้องขยับอะไร

แก้คีย์ใน payload → ต้องแก้ 3 จุดพร้อมกัน:
`TSP_Monitor_Calc.gs` (สร้าง) · `preview.html` (ตัวอย่าง + การวาด) · `net_client.cpp` (parse)
เพิ่มคีย์ใหม่ไม่ต้องขยับ `v` — ลบหรือเปลี่ยนความหมายให้เพิ่ม `v` เป็น 2

ถ้ามีคนแก้ `dlIndex_()` / `buildBranchData_()` / กติกาเงินสดในอนาคต **ให้รัน `TSP_test_vs_report()` ทุกครั้ง**
