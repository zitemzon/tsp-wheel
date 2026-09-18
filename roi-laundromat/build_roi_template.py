# -*- coding: utf-8 -*-
"""
สร้างเทมเพลต Excel: "การคำนวณ ROI ร้านสะดวกซัก Samsung Commercial"
โครงสร้างอ้างอิงจากฟอร์ม LG WM Franchise / Laundry Crew ROI Model และขยายให้ครบถ้วน
Run: python3 build_roi_template.py
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

# ---------------------------------------------------------------- style tokens
FONT   = "Tahoma"
NAVY   = "1428A0"   # Samsung blue
DARK   = "0B1A4A"
GREY   = "F2F2F2"
BAND   = "111111"
YELLOW = "FFFF00"
KPIFIL = "D9E1F2"
OKFILL = "E2EFDA"
WARN   = "FCE4D6"

C_IN   = "0000FF"   # hardcoded input
C_CALC = "000000"   # formula
C_LINK = "008000"   # cross-sheet link
C_WHT  = "FFFFFF"

FMT_THB  = '#,##0;(#,##0);"-"'
FMT_THB2 = '#,##0.00;(#,##0.00);"-"'
FMT_PCT  = '0.0%'
FMT_PCT2 = '0.00%'
FMT_NUM  = '#,##0.0;(#,##0.0);"-"'
FMT_INT  = '#,##0;(#,##0);"-"'
FMT_X    = '0.0"x"'
FMT_DATE = 'dd-mm-yyyy'

thin = Side(style="thin", color="BFBFBF")
med  = Side(style="medium", color="808080")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

REF = {}   # (sheet, key) -> "'Sheet'!$B$12"

def put(ws, cell, value, *, fmt=None, bold=False, color=C_CALC, fill=None,
        size=10, align=None, wrap=False, border=True, italic=False):
    c = ws[cell]
    c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, color=color, italic=italic)
    if fmt:  c.number_format = fmt
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    if border: c.border = BOX
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return c

def band(ws, row, text, last_col="E", fill=BAND, color=C_WHT, size=11):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    put(ws, f"A{row}", text, bold=True, color=color, fill=fill, size=size,
        align="left", border=False)
    for col in range(1, ws.max_column + 1):
        pass
    for ch in [get_column_letter(i) for i in range(1, 8)]:
        ws[f"{ch}{row}"].fill = PatternFill("solid", fgColor=fill)
    ws.row_dimensions[row].height = 20

class SheetBuilder:
    """ผู้ช่วยเขียนแถวแบบ label(A) | value(B) | alt(C) | type(D) | note(E)"""
    def __init__(self, wb, name, tab_color=NAVY):
        self.ws = wb.create_sheet(name)
        self.ws.sheet_properties.tabColor = tab_color
        self.name = name
        self.r = 1
    def _key(self, key, col="B"):
        if key:
            REF[(self.name, key)] = f"'{self.name}'!${col}${self.r}"
    def skip(self, n=1):
        self.r += n
    def band(self, text, last_col="E", fill=BAND):
        band(self.ws, self.r, text, last_col, fill)
        self.r += 1
    def sub(self, text, last_col="E"):
        band(self.ws, self.r, text, last_col, fill=DARK)
        self.r += 1
    def hdr(self, cells, last_col="E"):
        for col, txt in cells.items():
            put(self.ws, f"{col}{self.r}", txt, bold=True, color=C_WHT,
                fill=NAVY, align="center", wrap=True, size=9)
        self.ws.row_dimensions[self.r].height = 30
        self.r += 1
    def inp(self, key, label, value, fmt=FMT_THB, unit="Input", note="",
            dv=None, alt=None, comment=None):
        put(self.ws, f"A{self.r}", label, size=10)
        c = put(self.ws, f"B{self.r}", value, fmt=fmt, color=C_IN,
                fill=YELLOW, bold=True, align="right")
        if comment:
            c.comment = Comment(comment, "ROI Template")
        if alt is not None:
            put(self.ws, f"C{self.r}", alt, fmt=fmt, color="808080", align="right")
        put(self.ws, f"D{self.r}", unit, size=9, color="808080", align="center")
        put(self.ws, f"E{self.r}", note, size=9, color="595959", wrap=True)
        if dv is not None:
            dv.add(c)
        self._key(key)
        self.r += 1
        return c
    def calc(self, key, label, formula, fmt=FMT_THB, unit="Formula", note="",
             bold=False, fill=None, color=C_CALC, alt=None):
        put(self.ws, f"A{self.r}", label, size=10, bold=bold)
        put(self.ws, f"B{self.r}", formula, fmt=fmt, color=color, bold=bold,
            fill=fill, align="right")
        if alt is not None:
            put(self.ws, f"C{self.r}", alt, fmt=fmt, color="808080", align="right")
        put(self.ws, f"D{self.r}", unit, size=9, color="808080", align="center")
        put(self.ws, f"E{self.r}", note, size=9, color="595959", wrap=True)
        self._key(key)
        self.r += 1
    def text(self, label, note="", bold=False, italic=False, color="595959"):
        put(self.ws, f"A{self.r}", label, size=10, bold=bold, italic=italic,
            border=False)
        if note:
            put(self.ws, f"E{self.r}", note, size=9, color=color, wrap=True,
                border=False)
        self.r += 1

def R(sheet, key):
    return REF[(sheet, key)]
def IN(key):  return R("A_INPUT", key)
def CA(key):  return R("C_CALC", key)
def PK(key):  return R("B_PACKAGES", key)
def SV(key):  return R("F_SERVICE", key)

wb = Workbook()
wb.remove(wb.active)
TODAY = datetime.date(2026, 9, 18)

# ============================================================ Z_LISTS (dropdown)
zl = wb.create_sheet("Z_LISTS")
zl.sheet_properties.tabColor = "A6A6A6"
LISTS = {
    "A": ("PackageCode", ["S", "M", "L", "XL", "CUSTOM"]),
    "B": ("RevenueMode", ["A: รวมรอบ (ซัก+อบ ราคาเดียว)", "B: แยกเครื่องซัก / เครื่องอบ"]),
    "C": ("UtilityModel", ["1: % ของยอดขาย", "2: คำนวณจากหน่วยจริง (kWh/ลิตร)"]),
    "D": ("YesNo", ["Yes / ใช่", "No / ไม่ใช่"]),
    "E": ("Scenario", ["Base / ฐาน", "Best / ดีที่สุด", "Worst / แย่ที่สุด"]),
    "F": ("CapturePct", [0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50]),
    "G": ("VatRate", [0.0, 0.07]),
    "H": ("ServicePack", ["Package A (Minor x3 + Major x1)", "Package B (Major x2)", "ไม่ซื้อ / No package"]),
    "I": ("Warranty", ["2+0 ปี (มาตรฐาน)", "2+1 ปี", "2+2 ปี", "2+3 ปี"]),
    "J": ("Owner", ["Krittanan", "Sales 2", "Sales 3", "อื่นๆ"]),
    "K": ("LocType", ["หน้าหอพัก/คอนโด", "ในหมู่บ้านจัดสรร", "ริมถนนหลัก/ปั๊มน้ำมัน", "ใกล้ตลาด/ชุมชน", "ในห้างฯ/คอมมูนิตี้มอลล์"]),
    "L": ("Score", [1, 2, 3, 4, 5]),
}
for col, (name, vals) in LISTS.items():
    put(zl, f"{col}1", name, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
    for i, v in enumerate(vals, start=2):
        put(zl, f"{col}{i}", v, size=9,
            fmt=(FMT_PCT if name in ("CapturePct", "VatRate") else None))
    zl.column_dimensions[col].width = 26
    REF[("Z_LISTS", name)] = f"Z_LISTS!${col}$2:${col}${len(vals)+1}"
zl.sheet_state = "hidden"

def DV(name, sheet):
    dv = DataValidation(type="list", formula1="=" + REF[("Z_LISTS", name)],
                        allow_blank=True, showDropDown=False)
    sheet.add_data_validation(dv)
    return dv

# ============================================================ B_PACKAGES
p = SheetBuilder(wb, "B_PACKAGES", "2E75B6")
ws = p.ws
for c, w in zip("ABCDEFGHIJ", [30, 34, 14, 12, 12, 13, 13, 13, 13, 13]):
    ws.column_dimensions[c].width = w
ws.merge_cells("A1:J1")
put(ws, "A1", "B) MACHINE & PACKAGE MASTER / ฐานข้อมูลเครื่องและแพ็กเกจ",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
ws.row_dimensions[1].height = 26
put(ws, "A2", "ตัวเลขทั้งหมดในหน้านี้เป็น 'ค่าตั้งต้นตัวอย่าง' — ต้องแทนที่ด้วยราคาจริงจากใบเสนอราคา Samsung Commercial ก่อนใช้งานจริง",
    italic=True, size=9, color="C00000", border=False)
p.r = 4

# ---- 1) MACHINE MASTER
p.band("1) MACHINE MASTER / รายการเครื่องและอุปกรณ์", "J")
p.hdr({"A": "รหัส / Code", "B": "รุ่น / Model", "C": "ประเภท / Type",
       "D": "ความจุ (kg)", "E": "เวลา/รอบ (นาที)", "F": "ราคา/เครื่อง (บาท)",
       "G": "ราคาขาย/รอบ (บาท)", "H": "ไฟฟ้า kWh/รอบ", "I": "น้ำ ลิตร/รอบ",
       "J": "อายุใช้งาน (ปี)"}, "J")
MACHINES = [
    ("W14", "Samsung Commercial Washer 14 kg", "เครื่องซัก", 14, 32, 95000,  50, 0.9,  90, 8),
    ("W21", "Samsung Commercial Washer 21 kg", "เครื่องซัก", 21, 36, 135000, 70, 1.3, 130, 8),
    ("W28", "Samsung Commercial Washer 28 kg", "เครื่องซัก", 28, 40, 185000, 100, 1.7, 170, 8),
    ("D15", "Samsung Commercial Dryer 15 kg",  "เครื่องอบ",  15, 30, 90000,  40, 3.2,   0, 8),
    ("D25", "Samsung Commercial Dryer 25 kg",  "เครื่องอบ",  25, 35, 140000, 60, 5.0,   0, 8),
    ("VND", "ตู้จำหน่ายน้ำยา / Vending",        "อุปกรณ์เสริม", 0, 0, 35000,   0, 0.1,   0, 5),
    ("KSK", "ตู้แลกเหรียญ + POS / Kiosk",       "อุปกรณ์เสริม", 0, 0, 60000,   0, 0.1,   0, 5),
]
M_START = p.r
for m in MACHINES:
    rr = p.r
    put(ws, f"A{rr}", m[0], bold=True, align="center", size=9)
    put(ws, f"B{rr}", m[1], size=9)
    put(ws, f"C{rr}", m[2], size=9, align="center")
    for col, val, fmt in zip("DEFGHIJ", m[3:],
                             [FMT_INT, FMT_INT, FMT_THB, FMT_THB, '0.0', FMT_INT, FMT_INT]):
        put(ws, f"{col}{rr}", val, fmt=fmt, color=C_IN, fill=YELLOW,
            align="right", bold=True, size=9)
    p.r += 1
M_END = p.r - 1
REF[("B_PACKAGES", "m_code")]  = f"'B_PACKAGES'!$A${M_START}:$A${M_END}"
REF[("B_PACKAGES", "m_price")] = f"'B_PACKAGES'!$F${M_START}:$F${M_END}"
REF[("B_PACKAGES", "m_cycle")] = f"'B_PACKAGES'!$E${M_START}:$E${M_END}"
REF[("B_PACKAGES", "m_kwh")]   = f"'B_PACKAGES'!$H${M_START}:$H${M_END}"
REF[("B_PACKAGES", "m_water")] = f"'B_PACKAGES'!$I${M_START}:$I${M_END}"
REF[("B_PACKAGES", "m_sell")]  = f"'B_PACKAGES'!$G${M_START}:$G${M_END}"
W_ROWS = (M_START, M_START + 2)        # แถวเครื่องซัก
D_ROWS = (M_START + 3, M_START + 4)    # แถวเครื่องอบ
put(ws, f"A{p.r}", "หมายเหตุ: ราคา/เครื่อง = ราคาหลังหักส่วนลด ยังไม่รวม VAT | เวลา/รอบ = เวลาซักจริง + เวลาโหลด-ปลดผ้า",
    italic=True, size=9, color="595959", border=False)
p.r += 2

# ---- 2) PACKAGE MATRIX
p.band("2) PACKAGE MATRIX / จำนวนเครื่องในแต่ละแพ็กเกจ (แก้ช่องสีเหลืองเพื่อปรับแพ็กเกจ)", "J")
put(ws, f"A{p.r}", "รหัสคอลัมน์ต้องตรงกับตัวเลือกใน A_INPUT เป๊ะ:  S = เล็ก | M = กลาง | L = ใหญ่ | XL = แฟล็กชิป | CUSTOM = กำหนดเอง",
    italic=True, size=9, color="595959", border=False)
p.r += 1
p.hdr({"A": "รหัส / Code", "B": "รุ่น / Model", "C": "ราคา/เครื่อง (บาท)",
       "D": "S", "E": "M", "F": "L", "G": "XL", "H": "CUSTOM"}, "H")
QTY = {   # code: (S, M, L, XL, CUSTOM)
    "W14": (3, 4, 5,  6,  0),
    "W21": (1, 2, 3,  4,  0),
    "W28": (0, 1, 2,  3,  0),
    "D15": (2, 3, 4,  5,  0),
    "D25": (0, 1, 2,  3,  0),
    "VND": (1, 1, 1,  2,  0),
    "KSK": (1, 1, 1,  1,  0),
}
Q_START = p.r
for i, m in enumerate(MACHINES):
    rr = p.r
    mrow = M_START + i
    put(ws, f"A{rr}", m[0], bold=True, align="center", size=9)
    put(ws, f"B{rr}", f"='B_PACKAGES'!$B${mrow}", size=9, color=C_LINK)
    put(ws, f"C{rr}", f"='B_PACKAGES'!$F${mrow}", fmt=FMT_THB, color=C_LINK, align="right", size=9)
    for col, q in zip("DEFGH", QTY[m[0]]):
        put(ws, f"{col}{rr}", q, fmt=FMT_INT, color=C_IN, fill=YELLOW,
            align="right", bold=True, size=10)
    p.r += 1
Q_END = p.r - 1
PKG_COLS = {"S": "D", "M": "E", "L": "F", "XL": "G", "CUSTOM": "H"}
REF[("B_PACKAGES", "pkg_hdr")] = f"'B_PACKAGES'!$D${Q_START-1}:$H${Q_START-1}"

def qrange(col):  return f"'B_PACKAGES'!${col}${Q_START}:${col}${Q_END}"
def qrange_w(col): return f"'B_PACKAGES'!${col}${W_ROWS[0]-M_START+Q_START}:${col}${W_ROWS[1]-M_START+Q_START}"
def qrange_d(col): return f"'B_PACKAGES'!${col}${D_ROWS[0]-M_START+Q_START}:${col}${D_ROWS[1]-M_START+Q_START}"
WM_PRICE = f"'B_PACKAGES'!$E${W_ROWS[0]}:$E${W_ROWS[1]}"   # cycle time washers
DM_PRICE = f"'B_PACKAGES'!$E${D_ROWS[0]}:$E${D_ROWS[1]}"
WM_SELL  = f"'B_PACKAGES'!$G${W_ROWS[0]}:$G${W_ROWS[1]}"
DM_SELL  = f"'B_PACKAGES'!$G${D_ROWS[0]}:$G${D_ROWS[1]}"
WM_KWH   = f"'B_PACKAGES'!$H${W_ROWS[0]}:$H${W_ROWS[1]}"
DM_KWH   = f"'B_PACKAGES'!$H${D_ROWS[0]}:$H${D_ROWS[1]}"
WM_WATER = f"'B_PACKAGES'!$I${W_ROWS[0]}:$I${W_ROWS[1]}"

SUMROWS = {}
def prow(label, builder, fmt=FMT_THB, bold=False, note="", fill=None):
    rr = p.r
    ws.merge_cells(f"A{rr}:B{rr}")
    put(ws, f"A{rr}", label, size=10, bold=bold, fill=fill)
    if fill:
        ws[f"B{rr}"].fill = PatternFill("solid", fgColor=fill)
    ws[f"B{rr}"].border = BOX
    put(ws, f"C{rr}", note, size=8, color="595959", wrap=True, fill=fill)
    for col in "DEFGH":
        put(ws, f"{col}{rr}", builder(col), fmt=fmt, bold=bold, align="right",
            size=10, fill=fill)
    p.r += 1
    return rr

SUMROWS["w_units"] = prow("รวมเครื่องซัก / Washer units",
    lambda c: f"=SUM({qrange_w(c)})", FMT_INT, note="เครื่อง")
SUMROWS["d_units"] = prow("รวมเครื่องอบ / Dryer units",
    lambda c: f"=SUM({qrange_d(c)})", FMT_INT, note="เครื่อง")
SUMROWS["t_units"] = prow("รวมทั้งหมด / Total units",
    lambda c: f"=SUM({qrange(c)})", FMT_INT, bold=True, note="เครื่อง+อุปกรณ์")
SUMROWS["capex_m"] = prow("CAPEX เครื่อง+อุปกรณ์ / Machine CAPEX",
    lambda c: f"=SUMPRODUCT({qrange(c)},{REF[('B_PACKAGES','m_price')]})",
    bold=True, note="บาท (ไม่รวม VAT)")
SUMROWS["avg_w_cycle"] = prow("เวลาเฉลี่ย/รอบ เครื่องซัก",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_w(c)},{WM_PRICE})/SUM({qrange_w(c)}),0)",
    FMT_NUM, note="นาที (ถ่วงน้ำหนักตามจำนวนเครื่อง)")
SUMROWS["avg_d_cycle"] = prow("เวลาเฉลี่ย/รอบ เครื่องอบ",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_d(c)},{DM_PRICE})/SUM({qrange_d(c)}),0)",
    FMT_NUM, note="นาที")
SUMROWS["avg_w_price"] = prow("ราคาเฉลี่ย/รอบ ซัก",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_w(c)},{WM_SELL})/SUM({qrange_w(c)}),0)",
    FMT_NUM, note="บาท/รอบ")
SUMROWS["avg_d_price"] = prow("ราคาเฉลี่ย/รอบ อบ",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_d(c)},{DM_SELL})/SUM({qrange_d(c)}),0)",
    FMT_NUM, note="บาท/รอบ")
SUMROWS["kwh_w"] = prow("ไฟฟ้าเฉลี่ย kWh/รอบซัก",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_w(c)},{WM_KWH})/SUM({qrange_w(c)}),0)",
    '0.00', note="kWh")
SUMROWS["kwh_d"] = prow("ไฟฟ้าเฉลี่ย kWh/รอบอบ",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_d(c)},{DM_KWH})/SUM({qrange_d(c)}),0)",
    '0.00', note="kWh")
SUMROWS["water_w"] = prow("น้ำเฉลี่ย ลิตร/รอบซัก",
    lambda c: f"=IFERROR(SUMPRODUCT({qrange_w(c)},{WM_WATER})/SUM({qrange_w(c)}),0)",
    FMT_NUM, note="ลิตร")

p.r += 1
p.band("3) CAPEX อื่นๆ ต่อแพ็กเกจ / Fit-out & Set-up cost", "H")
FITOUT = {
    "fit_out":  ("ค่าตกแต่งร้าน + งานระบบไฟ/น้ำ / Fit-out", (250000, 350000, 480000, 650000, 0)),
    "signage":  ("ป้ายหน้าร้าน + ไฟฟ้า 3 เฟส / Signage & Power", (80000, 100000, 130000, 160000, 0)),
    "sec":      ("CCTV + ระบบความปลอดภัย / Security", (25000, 30000, 40000, 50000, 0)),
    "furn":     ("เฟอร์นิเจอร์ + แอร์ + TV / Furniture", (45000, 60000, 90000, 120000, 0)),
    "install":  ("ค่าขนส่ง+ติดตั้งเครื่อง / Delivery & Install", (25000, 35000, 50000, 70000, 0)),
    "opening":  ("การตลาดเปิดร้าน / Opening marketing", (20000, 30000, 40000, 60000, 0)),
}
for key, (label, vals) in FITOUT.items():
    rr = p.r
    ws.merge_cells(f"A{rr}:B{rr}")
    put(ws, f"A{rr}", label, size=10)
    ws[f"B{rr}"].border = BOX
    put(ws, f"C{rr}", "บาท", size=8, color="595959")
    for col, v in zip("DEFGH", vals):
        put(ws, f"{col}{rr}", v, fmt=FMT_THB, color=C_IN, fill=YELLOW,
            align="right", size=10)
    SUMROWS[key] = rr
    p.r += 1
SUMROWS["capex_fit"] = prow("รวม CAPEX ตกแต่ง/ติดตั้ง / Total fit-out",
    lambda c: "=SUM({0}{1}:{0}{2})".format(c, SUMROWS["fit_out"], SUMROWS["opening"]),
    bold=True, fill=KPIFIL, note="บาท")
SUMROWS["capex_all"] = prow("CAPEX รวมทั้งโครงการ (ก่อน VAT) / Total CAPEX",
    lambda c: "={0}{1}+{0}{2}".format(c, SUMROWS["capex_m"], SUMROWS["capex_fit"]),
    bold=True, fill=OKFILL, note="บาท")
SUMROWS["area"] = prow("พื้นที่แนะนำ / Recommended area",
    lambda c: {"D": 40, "E": 55, "F": 75, "G": 100, "H": 0}[c], FMT_INT, note="ตร.ม.")
SUMROWS["pop"] = prow("ประชากรในรัศมี 1 กม. ที่แนะนำ",
    lambda c: {"D": 2000, "E": 3500, "F": 5500, "G": 8000, "H": 0}[c], FMT_INT, note="คน")

def pk_lookup(rowkey):
    """INDEX/MATCH ตามรหัสแพ็กเกจที่เลือกใน A_INPUT"""
    return (f"INDEX('B_PACKAGES'!$D${SUMROWS[rowkey]}:$H${SUMROWS[rowkey]},"
            f"MATCH({IN('pkg')},{REF[('B_PACKAGES','pkg_hdr')]},0))")

# ============================================================ A_INPUT
a = SheetBuilder(wb, "A_INPUT", "C00000")
ws = a.ws
for c, w in zip("ABCDE", [56, 20, 18, 14, 54]):
    ws.column_dimensions[c].width = w
ws.merge_cells("A1:E1")
put(ws, "A1", "การคำนวณ ROI ร้านสะดวกซัก  Samsung Commercial  /  LAUNDROMAT ROI MODEL",
    bold=True, size=15, color=C_WHT, fill=NAVY, align="center", border=False)
ws.row_dimensions[1].height = 30
ws.merge_cells("A2:E2")
put(ws, "A2", "กรอกเฉพาะช่องสีเหลือง (Input) → ผลลัพธ์อัปเดตอัตโนมัติ | Fill yellow cells only → results auto-update   "
              "•  ช่อง override ใส่ 0 = ใช้ค่าจากแพ็กเกจ",
    italic=True, size=9, color="595959", align="center", border=False, fill=GREY)
a.r = 4

dv_pkg   = DV("PackageCode", ws)
dv_mode  = DV("RevenueMode", ws)
dv_util  = DV("UtilityModel", ws)
dv_yn    = DV("YesNo", ws)
dv_cap   = DV("CapturePct", ws)
dv_vat   = DV("VatRate", ws)
dv_owner = DV("Owner", ws)
dv_loc   = DV("LocType", ws)

# ---- 1
a.band("1) PROJECT INFO / ข้อมูลโครงการ")
a.inp("proj",  "ชื่อโครงการ / Project name", "Laundromat_สาขาตัวอย่าง", None, "Input", "ใช้แสดงบนหน้า H_PRINT_VIEW")
a.inp("cust",  "ลูกค้า / Customer", "K.ตัวอย่าง", None, "Input")
a.inp("loc",   "ทำเล / Location", "ระบุชื่อทำเล", None, "Input")
a.inp("loctype","ประเภททำเล / Location type", "หน้าหอพัก/คอนโด", None, "เลือก", "มีผลกับสมมติฐานความถี่การใช้บริการ", dv=dv_loc)
a.inp("owner", "ผู้รับผิดชอบ / Sales owner", "Krittanan", None, "เลือก", dv=dv_owner)
a.inp("date",  "วันที่ / Date", TODAY, FMT_DATE, "Input")
a.skip()

# ---- 2
a.band("2) MARKET SIZE & POPULATION / ขนาดตลาดและประชากร")
a.inp("pop",     "จำนวนประชากรในรัศมีบริการ (1 กม.) / Population", 3000, FMT_INT, "Input",
      "นับจากจำนวนห้องพัก x คนเฉลี่ย/ห้อง หรือข้อมูลทะเบียนราษฎร์",
      comment="แนะนำ: นับจำนวนห้องหอพัก/คอนโดในรัศมี 1 กม. x 1.5 คน/ห้อง")
a.inp("capture", "% คาดว่าจะใช้บริการ / Capture rate", 0.35, FMT_PCT, "เลือก/แก้ได้",
      "ทำเลหอพัก 30-40% | หมู่บ้านจัดสรร 15-25%", dv=dv_cap)
a.inp("comp",    "จำนวนคู่แข่งในรัศมี 500 ม. / Competitors", 1, FMT_INT, "Input",
      "0 = ไม่มีคู่แข่ง")
a.calc("share",  "ตัวคูณส่วนแบ่งตลาด / Market share factor",
       f"=IFERROR(1/(1+{IN('comp')}*0.6),1)", FMT_PCT, "Formula",
       "สูตร 1/(1+คู่แข่ง x 0.6) — ปรับตัวเลข 0.6 ได้ถ้ามีข้อมูลจริง")
a.calc("cust_n", "จำนวนลูกค้าคาดการณ์ / Estimated customers",
       f"={IN('pop')}*{IN('capture')}*{R('A_INPUT','share')}", FMT_INT, "Formula", "คน/เดือน")
a.inp("freq",    "ความถี่ใช้บริการ / Usage frequency", 3, FMT_NUM, "Input", "ครั้ง/คน/เดือน")
a.inp("cyc_visit","รอบซักต่อการมา 1 ครั้ง / Wash cycles per visit", 1.2, FMT_NUM, "Input",
      "ลูกค้าครอบครัวมักซัก 1.5-2 รอบ/ครั้ง")
a.inp("dry_att", "% ที่ใช้เครื่องอบด้วย / Dryer attach rate", 0.7, FMT_PCT, "Input",
      "หน้าฝน/คอนโดสูงถึง 85%")
a.calc("dem_w",  "ความต้องการรอบซัก / Wash cycles demand",
       f"={R('A_INPUT','cust_n')}*{IN('freq')}*{IN('cyc_visit')}", FMT_INT, "Formula", "รอบ/เดือน")
a.calc("dem_d",  "ความต้องการรอบอบ / Dry cycles demand",
       f"={R('A_INPUT','dem_w')}*{IN('dry_att')}", FMT_INT, "Formula", "รอบ/เดือน")
a.skip()

# ---- 3
a.band("3) PACKAGE & PRICING / แพ็กเกจและราคา")
a.inp("pkg",  "เลือกแพ็กเกจ / Package code", "M", None, "เลือก",
      "S / M / L / XL / CUSTOM (CUSTOM = กรอกจำนวนเครื่องเองที่ชีต B_PACKAGES)", dv=dv_pkg)
a.inp("mode", "รูปแบบการคิดรายได้ / Revenue model", "B: แยกเครื่องซัก / เครื่องอบ", None, "เลือก",
      "โหมด A = คิดราคาเดียวต่อรอบ (ซัก+อบ) แบบฟอร์ม LG | โหมด B = แยกราคาซัก/อบ (แม่นกว่า)", dv=dv_mode)
a.calc("w_pkg", "เครื่องซักตามแพ็กเกจ / Washers in package", f"={pk_lookup('w_units')}", FMT_INT, "จาก B_PACKAGES", "เครื่อง", color=C_LINK)
a.calc("d_pkg", "เครื่องอบตามแพ็กเกจ / Dryers in package",  f"={pk_lookup('d_units')}", FMT_INT, "จาก B_PACKAGES", "เครื่อง", color=C_LINK)
a.inp("w_ov",  "override จำนวนเครื่องซัก (0 = ใช้แพ็กเกจ)", 0, FMT_INT, "Input")
a.inp("d_ov",  "override จำนวนเครื่องอบ (0 = ใช้แพ็กเกจ)", 0, FMT_INT, "Input")
a.inp("price_w", "ราคา/รอบ เครื่องซัก (0 = ใช้ค่าเฉลี่ยแพ็กเกจ)", 0, FMT_THB, "Input", "บาท/รอบ")
a.inp("price_d", "ราคา/รอบ เครื่องอบ (0 = ใช้ค่าเฉลี่ยแพ็กเกจ)", 0, FMT_THB, "Input", "บาท/รอบ")
a.inp("price_c", "ราคา/รอบ รวมซัก+อบ (ใช้เฉพาะโหมด A)", 80, FMT_THB, "Input", "บาท/รอบ")
a.inp("cyc_w", "เวลา/รอบซัก (0 = ใช้ค่าแพ็กเกจ)", 0, FMT_NUM, "Input", "นาที")
a.inp("cyc_d", "เวลา/รอบอบ (0 = ใช้ค่าแพ็กเกจ)", 0, FMT_NUM, "Input", "นาที")
a.inp("cyc_c", "เวลา/รอบ รวมซัก+อบ (ใช้เฉพาะโหมด A)", 76, FMT_NUM, "Input", "นาที")
a.inp("chg",   "เวลาโหลด-ปลดผ้า+รอเครื่องว่าง / Changeover", 6, FMT_NUM, "Input",
      "นาที/รอบ — ทำให้ capacity สมจริงขึ้น (ฟอร์ม LG ไม่คิดส่วนนี้)")
a.skip()

# ---- 4
a.band("4) CAPEX / เงินลงทุน")
a.calc("capex_m_p", "CAPEX เครื่อง+อุปกรณ์ (จากแพ็กเกจ)", f"={pk_lookup('capex_m')}", FMT_THB, "จาก B_PACKAGES", color=C_LINK)
a.calc("capex_f_p", "CAPEX ตกแต่ง/ติดตั้ง (จากแพ็กเกจ)", f"={pk_lookup('capex_fit')}", FMT_THB, "จาก B_PACKAGES", color=C_LINK)
a.inp("capex_m_ov", "override CAPEX เครื่อง (0 = ใช้แพ็กเกจ)", 0, FMT_THB, "Input", "ใส่ตัวเลขจากใบเสนอราคาจริง")
a.inp("capex_f_ov", "override CAPEX ตกแต่ง (0 = ใช้แพ็กเกจ)", 0, FMT_THB, "Input")
a.inp("capex_other","CAPEX อื่นๆ / Other CAPEX", 0, FMT_THB, "Input", "เช่น ค่าเซ้ง ค่าโอนมิเตอร์ ค่าที่ปรึกษา")
a.inp("dep_month",  "เงินประกันค่าเช่า / Rent deposit", 3, FMT_NUM, "Input", "จำนวนเดือน (คืนเมื่อเลิกสัญญา)")
a.inp("wcap",       "เงินทุนหมุนเวียน / Working capital", 100000, FMT_THB, "Input", "เงินสดสำรอง 2-3 เดือนของ OPEX")
a.inp("vat",        "VAT %", 0.07, FMT_PCT, "เลือก", dv=dv_vat)
a.inp("vat_reg",    "จดทะเบียน VAT? / VAT registered", "No / ไม่ใช่", None, "เลือก",
      "ถ้า No → VAT ซื้อเครื่องถือเป็นต้นทุน (บวกเข้า CAPEX)", dv=dv_yn)
a.skip()

# ---- 5
a.band("5) OPERATION / การดำเนินงาน")
a.inp("open_h", "ชั่วโมงเปิด/วัน / Open hours per day", 24, FMT_NUM, "Input", "24 ชม. = ตู้หยอดเหรียญอัตโนมัติ")
a.inp("wd_util","% ใช้งานวันธรรมดา / Weekday utilization", 0.35, FMT_PCT, "Input", "ค่าอ้างอิงฟอร์ม LG = 35%")
a.inp("we_util","% ใช้งานวันหยุด / Weekend utilization", 0.65, FMT_PCT, "Input", "ค่าอ้างอิงฟอร์ม LG = 65%")
a.inp("wd_days","จำนวนวันธรรมดา/เดือน", 22, FMT_INT, "Input")
a.inp("we_days","จำนวนวันหยุด/เดือน", 8, FMT_INT, "Input")
a.calc("days",  "รวมวันเปิดบริการ/เดือน", f"={IN('wd_days')}+{IN('we_days')}", FMT_INT, "Formula", "วัน")
a.inp("o2o_on", "เปิดบริการรับ-ส่งผ้า O2O?", "No / ไม่ใช่", None, "เลือก", dv=dv_yn)
a.inp("o2o_ord","ออเดอร์ O2O ต่อวัน / O2O orders per day", 0, FMT_NUM, "Input", "บิล/วัน")
a.inp("o2o_cyc","รอบต่อ 1 ออเดอร์ O2O / Cycles per order", 2, FMT_NUM, "Input", "ซัก 1 + อบ 1")
a.inp("o2o_tk", "ยอดเฉลี่ยต่อบิล O2O / Avg ticket", 0, FMT_THB, "Input", "บาท/บิล")
a.inp("o2o_fee","ค่าธรรมเนียมแพลตฟอร์ม / Platform fee", 0.0, FMT_PCT, "Input", "% ของยอด O2O")
a.inp("o2o_rd", "ค่าส่ง (ไรเดอร์) / Rider cost", 0.0, FMT_PCT, "Input", "% ของยอด O2O")
a.skip()

# ---- 6
a.band("6) OTHER REVENUE / รายได้เสริม")
a.inp("vend_att","% ลูกค้าที่ซื้อน้ำยา-ของใช้ / Vending attach", 0.25, FMT_PCT, "Input")
a.inp("vend_sp", "ยอดซื้อเฉลี่ย/คน / Avg vending spend", 25, FMT_THB, "Input", "บาท/ครั้ง")
a.inp("vend_gp", "%กำไรขั้นต้นสินค้า vending / GP%", 0.45, FMT_PCT, "Input")
a.inp("extra",   "รายได้อื่น (ตู้กดน้ำ/ตู้เกม/ตู้คีบ) / Other income", 0, FMT_THB, "Input",
      "บาท/เดือน — เหมาะกับการวางตู้เสริมหน้าร้านสะดวกซัก")
a.inp("extra_gp","%กำไรขั้นต้นรายได้อื่น / GP%", 0.7, FMT_PCT, "Input")
a.skip()

# ---- 7
a.band("7) MONTHLY COSTS (OPEX) / ค่าใช้จ่ายต่อเดือน")
a.sub("7.1) FIXED COST / ค่าใช้จ่ายคงที่")
a.inp("rent",    "ค่าเช่า / Rent", 20000, FMT_THB, "Input", "บาท/เดือน")
a.inp("rent_esc","อัตราปรับค่าเช่าต่อปี / Rent escalation", 0.05, FMT_PCT, "Input", "ปกติสัญญา 3 ปี ปรับ 5-10%")
a.inp("staff",   "ค่าจ้างพนักงาน / Staff", 15000, FMT_THB, "Input", "บาท/เดือน (รวมประกันสังคมส่วนนายจ้าง)")
a.inp("ins",     "ประกันภัย / Insurance", 1200, FMT_THB, "Input", "ประกันทรัพย์สิน+บุคคลที่ 3")
a.inp("clean",   "ทำความสะอาด + เก็บเหรียญ / Cleaning & Coin collection", 3000, FMT_THB, "Input")
a.calc("maint",  "ค่าบำรุงรักษาตามสัญญา / Maintenance (contract)",
       0, FMT_THB, "จาก F_SERVICE", "เฉลี่ยต่อเดือนจากแพ็กเกจบริการ (ลิงก์อัตโนมัติ)", color=C_LINK)
a.inp("mkt",     "การตลาด / Marketing", 3000, FMT_THB, "Input", "ป้าย โปรโมชั่น ค่าโฆษณาออนไลน์")
a.inp("pos",     "เน็ต + ระบบ POS/แอป / Internet & Software", 1500, FMT_THB, "Input")
a.inp("other_f", "ค่าใช้จ่ายคงที่อื่นๆ / Other fixed", 0, FMT_THB, "Input")
a.calc("fix_tot","รวมค่าใช้จ่ายคงที่ / TOTAL FIXED COST",
       f"=SUM({IN('rent')},{IN('staff')},{IN('ins')},{IN('clean')},{R('A_INPUT','maint')},{IN('mkt')},{IN('pos')},{IN('other_f')})",
       FMT_THB, "Formula", "บาท/เดือน (ยังไม่รวมค่าน้ำ-ไฟ)", bold=True, fill=KPIFIL)
a.sub("7.2) VARIABLE COST / ค่าใช้จ่ายผันแปร")
a.inp("util_mode","โมเดลค่าน้ำ-ไฟ / Utilities model", "1: % ของยอดขาย", None, "เลือก",
      "โมเดล 1 = ประมาณเป็น % ของยอดขาย | โมเดล 2 = คำนวณจาก kWh และลิตรต่อรอบจริง", dv=dv_util)
a.inp("elec_p",  "ค่าไฟ (% ของยอดขาย) / Electricity %", 0.12, FMT_PCT, "โมเดล 1", "ค่าอ้างอิงฟอร์ม LG = 12%")
a.inp("water_p", "ค่าน้ำ (% ของยอดขาย) / Water %", 0.02, FMT_PCT, "โมเดล 1", "ค่าอ้างอิงฟอร์ม LG = 2%")
a.inp("elec_r",  "ค่าไฟต่อหน่วย / Electricity tariff", 4.8, FMT_THB2, "โมเดล 2", "บาท/kWh (รวม Ft + VAT)")
a.inp("water_r", "ค่าน้ำต่อหน่วย / Water tariff", 18, FMT_THB2, "โมเดล 2", "บาท/ลูกบาศก์เมตร")
a.inp("base_kwh","ค่าไฟส่วนกลาง (แอร์/ไฟ/ป้าย) / Base electricity", 3500, FMT_THB, "Input", "บาท/เดือน คงที่")
a.inp("cons_cyc","วัสดุสิ้นเปลืองต่อรอบ / Consumable per cycle", 2.5, FMT_THB2, "Input", "บาท/รอบ (น้ำยา ถุง ฯลฯ)")
a.inp("epay_p",  "ค่าธรรมเนียมรับชำระเงิน / e-Payment fee", 0.015, FMT_PCT, "Input", "% ของยอดที่จ่ายผ่าน QR/บัตร")
a.inp("epay_sh", "สัดส่วนยอดที่จ่ายแบบไม่ใช้เหรียญ / Cashless share", 0.4, FMT_PCT, "Input")
a.inp("repair_p","ค่าซ่อมผันแปร / Variable repair", 0.02, FMT_PCT, "Input", "% ของยอดขาย (นอกประกัน)")
a.inp("royal_p", "ค่าสิทธิ์/แฟรนไชส์ / Royalty fee", 0.0, FMT_PCT, "Input", "% ของยอดขาย")
a.skip()

# ---- 8
a.band("8) FINANCE & TAX / แหล่งเงินทุนและภาษี")
a.inp("loan_p",  "สัดส่วนเงินกู้ / Loan % of CAPEX", 0.0, FMT_PCT, "Input", "0% = ลงทุนด้วยเงินสดทั้งหมด")
a.inp("loan_r",  "ดอกเบี้ยเงินกู้ต่อปี / Interest rate", 0.075, FMT_PCT, "Input", "สินเชื่อ SME 6.5-9%")
a.inp("loan_y",  "ระยะเวลาผ่อน / Loan term", 5, FMT_INT, "Input", "ปี")
a.inp("dep_y_m", "อายุตัดค่าเสื่อมเครื่อง / Machine depreciation", 8, FMT_INT, "Input", "ปี (เส้นตรง)")
a.inp("dep_y_f", "อายุตัดค่าเสื่อมตกแต่ง / Fit-out depreciation", 5, FMT_INT, "Input", "ปี (เส้นตรง)")
a.inp("tax_r",   "อัตราภาษีเงินได้ / Corporate tax", 0.2, FMT_PCT, "Input",
      "นิติบุคคลทั่วไป 20% | SME กำไร<300k ยกเว้น, 300k-3M = 15%")
a.inp("disc_r",  "อัตราคิดลด / Discount rate (WACC)", 0.1, FMT_PCT, "Input", "ใช้คำนวณ NPV/IRR")
a.inp("growth",  "อัตราเติบโตยอดขายต่อปี / Revenue growth", 0.05, FMT_PCT, "Input", "ปีที่ 2 เป็นต้นไป")
a.inp("horizon", "ระยะเวลาประเมิน / Evaluation horizon", 5, FMT_INT, "Input", "ปี (สูงสุด 5)")
ws.freeze_panes = "A4"
ws.sheet_view.showGridLines = False

# ============================================================ F_SERVICE
s = SheetBuilder(wb, "F_SERVICE", "7030A0")
ws = s.ws
for c, w in zip("ABCDEFGH", [36, 16, 14, 12, 14, 14, 14, 40]):
    ws.column_dimensions[c].width = w
ws.merge_cells("A1:H1")
put(ws, "A1", "F) SERVICE, WARRANTY & MAINTENANCE PROGRAM / แพ็กเกจรับประกันและบำรุงรักษา",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
ws.row_dimensions[1].height = 26
put(ws, "A2", "โครงสร้างอ้างอิงฟอร์ม Franchise Maintenance Pgm. (Extended Warranty + Maintenance Package + Extra Visit Fee) — แก้ราคาตามเงื่อนไขผู้ผลิตจริง",
    italic=True, size=9, color="C00000", border=False)
s.r = 4
dv_sp = DV("ServicePack", ws)
dv_wr = DV("Warranty", ws)

s.band("1) EXTENDED WARRANTY / ประกันขยายเวลา", "H")
s.hdr({"A": "รุ่น / Model", "B": "ระยะประกัน / Period", "C": "ราคา/เครื่อง (EA)",
       "D": "จำนวน / Qty", "E": "รวม / Sub total", "F": "จากแพ็กเกจ",
       "G": "", "H": "หมายเหตุ"}, "H")
s.inp("wr_period", "ระยะประกันที่เลือก / Warranty period", "2+2 ปี", None, "เลือก",
      "2+2 = ประกันมาตรฐาน 2 ปี + ขยายอีก 2 ปี", dv=dv_wr)
W_START = s.r
WR_PRICE = {"W14": 4100, "W21": 4600, "W28": 5200, "D15": 2700, "D25": 3100}
for i, m in enumerate(MACHINES[:5]):
    rr = s.r
    mrow = M_START + i
    qrow = Q_START + i
    put(ws, f"A{rr}", f"='B_PACKAGES'!$B${mrow}", size=9, color=C_LINK)
    put(ws, f"B{rr}", f"={R('F_SERVICE','wr_period')}", size=9, align="center", color=C_LINK)
    put(ws, f"C{rr}", WR_PRICE[m[0]], fmt=FMT_THB, color=C_IN, fill=YELLOW, align="right", size=9)
    put(ws, f"D{rr}", f"=F{rr}", fmt=FMT_INT, align="right", size=9)
    put(ws, f"E{rr}", f"=C{rr}*D{rr}", fmt=FMT_THB, align="right", size=9)
    put(ws, f"F{rr}", f"=INDEX('B_PACKAGES'!$D${qrow}:$H${qrow},MATCH({IN('pkg')},{REF[('B_PACKAGES','pkg_hdr')]},0))",
        fmt=FMT_INT, align="right", size=9, color=C_LINK)
    put(ws, f"G{rr}", "", size=9)
    put(ws, f"H{rr}", "ราคาต่อเครื่อง ต่อรอบสัญญา", size=8, color="595959")
    s.r += 1
W_END = s.r - 1
rr = s.r
put(ws, f"A{rr}", "รวมค่าประกันขยายเวลา / Total extended warranty", bold=True, size=10)
for col in "BCD":
    put(ws, f"{col}{rr}", None, fill=KPIFIL)
put(ws, f"E{rr}", f"=SUM(E{W_START}:E{W_END})", fmt=FMT_THB, bold=True, fill=KPIFIL, align="right")
put(ws, f"F{rr}", None, fill=KPIFIL); put(ws, f"G{rr}", None, fill=KPIFIL)
put(ws, f"H{rr}", "จ่ายครั้งเดียว — ตัดเฉลี่ยตามจำนวนปีที่ขยาย", size=8, color="595959", fill=KPIFIL)
REF[("F_SERVICE", "wr_total")] = f"'F_SERVICE'!$E${rr}"
s.r += 2

s.band("2) MAINTENANCE PACKAGE / แพ็กเกจบำรุงรักษา", "H")
s.text("Package A : Minor x3 (เดือนที่ 3, 6, 9) + Major x1 (เดือนที่ 12)", bold=True)
s.text("Package B : Major x2 (เดือนที่ 6, 12)", bold=True)
s.inp("sv_pack", "เลือกแพ็กเกจบริการ / Service package", "Package A (Minor x3 + Major x1)",
      None, "เลือก", "เลือก 'ไม่ซื้อ' ถ้าลูกค้าดูแลเอง", dv=dv_sp)
s.inp("pa_ea", "Package A: ราคาต่อชุด / EA", 11495, FMT_THB, "Input", "บาท/เครื่อง/ปี")
s.inp("pb_ea", "Package B: ราคาต่อชุด / EA", 13680, FMT_THB, "Input", "บาท/เครื่อง/ปี")
s.calc("sv_qty", "จำนวนเครื่องที่เข้าโปรแกรม / Qty",
       f"={pk_lookup('w_units')}+{pk_lookup('d_units')}", FMT_INT, "จาก B_PACKAGES", "เครื่อง", color=C_LINK)
s.calc("sv_year", "ค่าบำรุงรักษาต่อปี / Maintenance per year",
       f'=IF({R("F_SERVICE","sv_pack")}="Package A (Minor x3 + Major x1)",{R("F_SERVICE","pa_ea")},'
       f'IF({R("F_SERVICE","sv_pack")}="Package B (Major x2)",{R("F_SERVICE","pb_ea")},0))*{R("F_SERVICE","sv_qty")}',
       FMT_THB, "Formula", "บาท/ปี", bold=True)
s.skip()

s.band("3) EXTRA VISIT FEE / ค่าเดินทางนอกพื้นที่", "H")
s.inp("km_in",   "ระยะทางจากศูนย์บริการ / Distance", 60, FMT_NUM, "Input", "กิโลเมตร (เที่ยวเดียว) — ฟอร์มอ้างอิงเดิมใช้ 112 กม.")
s.inp("km_free", "ระยะฟรี / Free radius", 40, FMT_NUM, "Input", "กิโลเมตร")
s.inp("km_rate", "ค่าเดินทาง/กม. / Rate per km", 160, FMT_THB2, "Input", "บาท/กม. (ไป-กลับรวมแล้ว)")
s.calc("km_ex",  "ระยะส่วนเกิน / Extra km", f"=MAX(0,{R('F_SERVICE','km_in')}-{R('F_SERVICE','km_free')})",
       FMT_NUM, "Formula", "กม.")
s.calc("visit_fee", "ค่าเดินทางต่อเที่ยว / Cost per visit",
       f"={R('F_SERVICE','km_ex')}*{R('F_SERVICE','km_rate')}", FMT_THB, "Formula", "บาท/เที่ยว")
s.calc("visits", "จำนวนเที่ยวต่อปี / Visits per year",
       f'=IF({R("F_SERVICE","sv_pack")}="Package A (Minor x3 + Major x1)",4,'
       f'IF({R("F_SERVICE","sv_pack")}="Package B (Major x2)",2,0))', FMT_INT, "Formula", "เที่ยว/ปี")
s.calc("km_year", "ค่าเดินทางรวมต่อปี / Extra visit fee per year",
       f"={R('F_SERVICE','visit_fee')}*{R('F_SERVICE','visits')}", FMT_THB, "Formula", "บาท/ปี", bold=True)
s.skip()

s.band("4) SUMMARY / สรุปค่าบริการ", "H")
s.calc("sv_total_y", "ค่าบำรุงรักษา + เดินทาง ต่อปี",
       f"={R('F_SERVICE','sv_year')}+{R('F_SERVICE','km_year')}", FMT_THB, "Formula", "บาท/ปี", bold=True)
s.inp("wr_years", "จำนวนปีที่เฉลี่ยค่าประกันขยาย / Amortize warranty over", 4, FMT_INT, "Input", "ปี")
s.calc("per_month", "ค่าบริการเฉลี่ยต่อเดือน (ส่งไป A_INPUT)",
       f"=({R('F_SERVICE','sv_total_y')}/12)+IFERROR({REF[('F_SERVICE','wr_total')]}/{R('F_SERVICE','wr_years')}/12,0)",
       FMT_THB, "Formula", "บาท/เดือน — ลิงก์เข้าช่อง 'ค่าบำรุงรักษาตามสัญญา' ใน A_INPUT",
       bold=True, fill=OKFILL)
ws.sheet_view.showGridLines = False

# patch A_INPUT maintenance link
_m = REF[("A_INPUT", "maint")].split("!")[1].replace("$", "")
a.ws[_m] = f"={SV('per_month')}"

# ============================================================ C_CALC
c = SheetBuilder(wb, "C_CALC", "548235")
ws = c.ws
for col, w in zip("ABCDE", [52, 20, 18, 16, 56]):
    ws.column_dimensions[col].width = w
ws.merge_cells("A1:E1")
put(ws, "A1", "C) CALCULATION ENGINE / เครื่องคำนวณ (ห้ามแก้ไข — เป็นสูตรทั้งหมด)",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
ws.row_dimensions[1].height = 26
c.r = 3
MODE_A = f'LEFT({IN("mode")},1)="A"'
YES    = lambda ref: f'LEFT({ref},1)="Y"'

c.band("1) CAPACITY / กำลังผลิต")
c.calc("w_units", "เครื่องซักที่ใช้จริง / Washer units",
       f"=IF({IN('w_ov')}>0,{IN('w_ov')},{IN('w_pkg')})", FMT_INT, "เครื่อง")
c.calc("d_units", "เครื่องอบที่ใช้จริง / Dryer units",
       f"=IF({IN('d_ov')}>0,{IN('d_ov')},{IN('d_pkg')})", FMT_INT, "เครื่อง")
c.calc("t_units", "รวมเครื่องซัก+อบ / Total machines",
       f"={CA('w_units')}+{CA('d_units')}", FMT_INT, "เครื่อง", bold=True)
c.calc("cyc_w", "เวลา/รอบซักที่ใช้ / Wash cycle time",
       f"=IF({IN('cyc_w')}>0,{IN('cyc_w')},{pk_lookup('avg_w_cycle')})", FMT_NUM, "นาที")
c.calc("cyc_d", "เวลา/รอบอบที่ใช้ / Dry cycle time",
       f"=IF({IN('cyc_d')}>0,{IN('cyc_d')},{pk_lookup('avg_d_cycle')})", FMT_NUM, "นาที")
c.calc("cpd_w", "รอบ/วัน ต่อเครื่องซัก / Cycles per washer-day",
       f"=IFERROR({IN('open_h')}*60/({CA('cyc_w')}+{IN('chg')}),0)", FMT_NUM, "รอบ",
       note="สูตร: ชั่วโมงเปิด x 60 / (เวลารอบ + เวลาโหลด-ปลดผ้า)")
c.calc("cpd_d", "รอบ/วัน ต่อเครื่องอบ / Cycles per dryer-day",
       f"=IFERROR({IN('open_h')}*60/({CA('cyc_d')}+{IN('chg')}),0)", FMT_NUM, "รอบ")
c.calc("cpd_c", "รอบ/วัน ต่อเครื่อง (โหมด A) / Combined",
       f"=IFERROR({IN('open_h')}*60/({IN('cyc_c')}+{IN('chg')}),0)", FMT_NUM, "รอบ")
c.calc("eff_days", "วันทำการถ่วงน้ำหนักการใช้งาน / Effective days",
       f"={IN('wd_days')}*{IN('wd_util')}+{IN('we_days')}*{IN('we_util')}", FMT_NUM, "วัน-เทียบเท่า",
       note="สูตร: วันธรรมดา x %ใช้งาน + วันหยุด x %ใช้งาน")
c.calc("cap_w", "กำลังผลิตเครื่องซัก / Washer capacity",
       f"={CA('w_units')}*{CA('cpd_w')}*{CA('eff_days')}", FMT_INT, "รอบ/เดือน")
c.calc("cap_d", "กำลังผลิตเครื่องอบ / Dryer capacity",
       f"={CA('d_units')}*{CA('cpd_d')}*{CA('eff_days')}", FMT_INT, "รอบ/เดือน")
c.calc("cap_c", "กำลังผลิตรวม (โหมด A) / Combined capacity",
       f"={CA('t_units')}*{CA('cpd_c')}*{CA('eff_days')}", FMT_INT, "รอบ/เดือน")
c.calc("cap_tot", "กำลังผลิตที่ใช้ตามโหมด / Capacity (active mode)",
       f"=IF({MODE_A},{CA('cap_c')},{CA('cap_w')}+{CA('cap_d')})", FMT_INT, "รอบ/เดือน", bold=True)
c.calc("cap_max", "กำลังผลิตสูงสุด 100% / Theoretical max",
       f"=IF({MODE_A},{CA('t_units')}*{CA('cpd_c')}*{IN('days')},"
       f"({CA('w_units')}*{CA('cpd_w')}+{CA('d_units')}*{CA('cpd_d')})*{IN('days')})",
       FMT_INT, "รอบ/เดือน", note="ใช้ดูเพดานสูงสุดถ้าเครื่องเดินเต็ม 24 ชม.")
c.skip()

c.band("2) PRICE / ราคาต่อรอบ")
c.calc("p_w", "ราคา/รอบซักที่ใช้", f"=IF({IN('price_w')}>0,{IN('price_w')},{pk_lookup('avg_w_price')})", FMT_THB2, "บาท")
c.calc("p_d", "ราคา/รอบอบที่ใช้", f"=IF({IN('price_d')}>0,{IN('price_d')},{pk_lookup('avg_d_price')})", FMT_THB2, "บาท")
c.calc("p_c", "ราคา/รอบ รวม (โหมด A)", f"={IN('price_c')}", FMT_THB2, "บาท")
c.skip()

c.band("3) DEMAND & BILLED CYCLES / ความต้องการและรอบที่ขายได้")
c.calc("o2o_w", "รอบซักจาก O2O", f"=IF({YES(IN('o2o_on'))},{IN('o2o_ord')}*{IN('days')},0)", FMT_INT, "รอบ/เดือน")
c.calc("o2o_d", "รอบอบจาก O2O",
       f"=IF({YES(IN('o2o_on'))},{IN('o2o_ord')}*{IN('days')}*MAX({IN('o2o_cyc')}-1,0),0)", FMT_INT, "รอบ/เดือน")
c.calc("dem_tot", "ความต้องการรวม / Total demand",
       f"=IF({MODE_A},{IN('dem_w')}+{CA('o2o_w')},{IN('dem_w')}+{IN('dem_d')}+{CA('o2o_w')}+{CA('o2o_d')})",
       FMT_INT, "รอบ/เดือน", bold=True)
c.calc("cap_chk", "Capacity check (ความต้องการ / กำลังผลิต)",
       f"=IFERROR({CA('dem_tot')}/{CA('cap_tot')},0)", FMT_PCT, "%",
       note=">100% = เครื่องไม่พอ ควรเพิ่มเครื่องหรือขึ้นราคา | <50% = ลงทุนเกินตัว", bold=True)
c.calc("fill", "อัตราที่รองรับได้จริง / Fill rate",
       f"=IFERROR(MIN(1,{CA('cap_tot')}/{CA('dem_tot')}),1)", FMT_PCT, "%",
       note="ถ้าความต้องการเกินกำลังผลิต รายได้จะถูกจำกัดตามสัดส่วนนี้")
c.calc("bil_w", "รอบซักที่ขายได้ (walk-in)", f"={IN('dem_w')}*{CA('fill')}", FMT_INT, "รอบ/เดือน")
c.calc("bil_d", "รอบอบที่ขายได้ (walk-in)", f"=IF({MODE_A},0,{IN('dem_d')}*{CA('fill')})", FMT_INT, "รอบ/เดือน")
c.calc("bil_tot", "รอบที่ขายได้รวม (รวม O2O)",
       f"={CA('bil_w')}+{CA('bil_d')}+({CA('o2o_w')}+{CA('o2o_d')})*{CA('fill')}", FMT_INT, "รอบ/เดือน", bold=True)
c.calc("cyc_day", "รอบเฉลี่ยต่อวัน / Cycles per day",
       f"=IFERROR({CA('bil_tot')}/{IN('days')},0)", FMT_NUM, "รอบ/วัน")
c.skip()

c.band("4) REVENUE / รายได้ต่อเดือน")
c.calc("rev_w", "รายได้เครื่องซัก / Wash revenue",
       f"=IF({MODE_A},{CA('bil_w')}*{CA('p_c')},{CA('bil_w')}*{CA('p_w')})", FMT_THB, "บาท")
c.calc("rev_d", "รายได้เครื่องอบ / Dry revenue", f"={CA('bil_d')}*{CA('p_d')}", FMT_THB, "บาท")
c.calc("rev_o2o", "รายได้ O2O / O2O revenue",
       f"=IF({YES(IN('o2o_on'))},{IN('o2o_ord')}*{IN('days')}*{IN('o2o_tk')}*{CA('fill')},0)", FMT_THB, "บาท")
c.calc("rev_vd", "รายได้สินค้า vending / Vending revenue",
       f"={R('A_INPUT','cust_n')}*{IN('vend_att')}*{IN('vend_sp')}*{CA('fill')}", FMT_THB, "บาท")
c.calc("rev_ex", "รายได้อื่น / Other income", f"={IN('extra')}", FMT_THB, "บาท")
c.calc("rev_tot", "รายได้รวม / TOTAL REVENUE",
       f"=SUM({CA('rev_w')},{CA('rev_d')},{CA('rev_o2o')},{CA('rev_vd')},{CA('rev_ex')})",
       FMT_THB, "บาท/เดือน", bold=True, fill=KPIFIL)
c.calc("rev_mach", "รายได้ต่อเครื่องต่อวัน / Revenue per machine-day",
       f"=IFERROR(({CA('rev_w')}+{CA('rev_d')})/{CA('t_units')}/{IN('days')},0)", FMT_THB2, "บาท",
       note="ตัวชี้วัดเทียบสาขา: ต่ำกว่า 200 บาท/เครื่อง/วัน ถือว่าน่ากังวล")
c.calc("p_avg", "ราคาเฉลี่ยต่อรอบ / Blended price per cycle",
       f"=IFERROR(({CA('rev_w')}+{CA('rev_d')}+{CA('rev_o2o')})/{CA('bil_tot')},0)", FMT_THB2, "บาท")
c.skip()

c.band("5) VARIABLE COST / ต้นทุนผันแปร")
c.calc("cost_util", "ค่าน้ำ-ไฟ / Utilities",
       f"=IF(LEFT({IN('util_mode')},1)=\"1\",({IN('elec_p')}+{IN('water_p')})*{CA('rev_tot')},"
       f"({CA('bil_w')}+{CA('o2o_w')}*{CA('fill')})*({pk_lookup('kwh_w')}*{IN('elec_r')}+{pk_lookup('water_w')}/1000*{IN('water_r')})"
       f"+({CA('bil_d')}+{CA('o2o_d')}*{CA('fill')})*{pk_lookup('kwh_d')}*{IN('elec_r')})+{IN('base_kwh')}",
       FMT_THB, "บาท", note="โมเดล 1 = % ของยอดขาย | โมเดล 2 = kWh/ลิตร จริง + ค่าไฟส่วนกลาง")
c.calc("cost_cons", "วัสดุสิ้นเปลือง / Consumables", f"={CA('bil_tot')}*{IN('cons_cyc')}", FMT_THB, "บาท")
c.calc("cost_o2o", "ค่าธรรมเนียม + ค่าส่ง O2O",
       f"={CA('rev_o2o')}*({IN('o2o_fee')}+{IN('o2o_rd')})", FMT_THB, "บาท")
c.calc("cost_epay", "ค่าธรรมเนียมรับชำระเงิน / e-Payment fee",
       f"={CA('rev_tot')}*{IN('epay_sh')}*{IN('epay_p')}", FMT_THB, "บาท")
c.calc("cost_rep", "ค่าซ่อมผันแปร / Variable repair", f"={CA('rev_tot')}*{IN('repair_p')}", FMT_THB, "บาท")
c.calc("cost_roy", "ค่าสิทธิ์/แฟรนไชส์ / Royalty", f"={CA('rev_tot')}*{IN('royal_p')}", FMT_THB, "บาท")
c.calc("cost_cogs", "ต้นทุนสินค้า vending + รายได้อื่น",
       f"={CA('rev_vd')}*(1-{IN('vend_gp')})+{CA('rev_ex')}*(1-{IN('extra_gp')})", FMT_THB, "บาท")
c.calc("var_tot", "รวมต้นทุนผันแปร / TOTAL VARIABLE COST",
       f"=SUM({CA('cost_util')},{CA('cost_cons')},{CA('cost_o2o')},{CA('cost_epay')},"
       f"{CA('cost_rep')},{CA('cost_roy')},{CA('cost_cogs')})", FMT_THB, "บาท/เดือน", bold=True, fill=KPIFIL)
c.calc("var_pct", "ต้นทุนผันแปร % ของยอดขาย", f"=IFERROR({CA('var_tot')}/{CA('rev_tot')},0)", FMT_PCT, "%")
c.calc("cm", "กำไรส่วนเกิน / Contribution margin",
       f"={CA('rev_tot')}-{CA('var_tot')}", FMT_THB, "บาท/เดือน")
c.calc("cm_pct", "อัตรากำไรส่วนเกิน / CM %", f"=IFERROR({CA('cm')}/{CA('rev_tot')},0)", FMT_PCT, "%", bold=True)
c.skip()

c.band("6) CAPEX & FUNDING / เงินลงทุนและแหล่งเงิน")
c.calc("capex_m", "CAPEX เครื่อง+อุปกรณ์",
       f"=IF({IN('capex_m_ov')}>0,{IN('capex_m_ov')},{pk_lookup('capex_m')})", FMT_THB, "บาท")
c.calc("capex_f", "CAPEX ตกแต่ง/ติดตั้ง",
       f"=IF({IN('capex_f_ov')}>0,{IN('capex_f_ov')},{pk_lookup('capex_fit')})", FMT_THB, "บาท")
c.calc("capex_vat", "VAT ที่เป็นต้นทุน (กรณีไม่จด VAT)",
       f"=IF({YES(IN('vat_reg'))},0,({CA('capex_m')}+{CA('capex_f')})*{IN('vat')})", FMT_THB, "บาท",
       note="ถ้าจด VAT จะขอคืนได้ → ไม่ถือเป็นต้นทุน")
c.calc("deposit", "เงินประกันค่าเช่า", f"={IN('rent')}*{IN('dep_month')}", FMT_THB, "บาท")
c.calc("capex_tot", "เงินลงทุนรวม / TOTAL CAPEX",
       f"=SUM({CA('capex_m')},{CA('capex_f')},{IN('capex_other')},{CA('capex_vat')},{CA('deposit')},{IN('wcap')})",
       FMT_THB, "บาท", bold=True, fill=OKFILL)
c.calc("loan", "เงินกู้ / Loan amount", f"={CA('capex_tot')}*{IN('loan_p')}", FMT_THB, "บาท")
c.calc("equity", "เงินลงทุนของเจ้าของ / Equity", f"={CA('capex_tot')}-{CA('loan')}", FMT_THB, "บาท", bold=True)
c.calc("pmt", "ค่างวดเงินกู้/เดือน / Loan payment",
       f"=IFERROR(IF({CA('loan')}>0,PMT({IN('loan_r')}/12,{IN('loan_y')}*12,-{CA('loan')}),0),0)",
       FMT_THB, "บาท/เดือน")
c.skip()

c.band("7) MONTHLY P&L / งบกำไรขาดทุนต่อเดือน (ปีที่ 1)")
c.calc("pl_rev", "รายได้ / Revenue", f"={CA('rev_tot')}", FMT_THB, "บาท", bold=True)
c.calc("pl_var", "หัก ต้นทุนผันแปร", f"=-{CA('var_tot')}", FMT_THB, "บาท")
c.calc("pl_fix", "หัก ค่าใช้จ่ายคงที่", f"=-{IN('fix_tot')}", FMT_THB, "บาท")
c.calc("ebitda", "EBITDA / กำไรก่อนดอกเบี้ย ภาษี ค่าเสื่อม",
       f"={CA('pl_rev')}+{CA('pl_var')}+{CA('pl_fix')}", FMT_THB, "บาท", bold=True, fill=KPIFIL)
c.calc("dep", "ค่าเสื่อมราคา / Depreciation",
       f"=IFERROR({CA('capex_m')}/{IN('dep_y_m')}/12,0)+IFERROR({CA('capex_f')}/{IN('dep_y_f')}/12,0)",
       FMT_THB, "บาท")
c.calc("ebit", "EBIT", f"={CA('ebitda')}-{CA('dep')}", FMT_THB, "บาท")
c.calc("int", "ดอกเบี้ยจ่าย (เฉลี่ยปีแรก)", f"={CA('loan')}*{IN('loan_r')}/12", FMT_THB, "บาท")
c.calc("ebt", "กำไรก่อนภาษี / EBT", f"={CA('ebit')}-{CA('int')}", FMT_THB, "บาท")
c.calc("tax", "ภาษีเงินได้ / Income tax", f"=MAX(0,{CA('ebt')})*{IN('tax_r')}", FMT_THB, "บาท")
c.calc("np", "กำไรสุทธิ / NET PROFIT", f"={CA('ebt')}-{CA('tax')}", FMT_THB, "บาท/เดือน",
       bold=True, fill=OKFILL)
c.calc("np_m", "อัตรากำไรสุทธิ / Net margin", f"=IFERROR({CA('np')}/{CA('rev_tot')},0)", FMT_PCT, "%", bold=True)
c.calc("ebitda_m", "อัตรา EBITDA / EBITDA margin", f"=IFERROR({CA('ebitda')}/{CA('rev_tot')},0)", FMT_PCT, "%")
c.calc("cf", "กระแสเงินสดสุทธิ/เดือน / Net cash flow",
       f"={CA('np')}+{CA('dep')}", FMT_THB, "บาท/เดือน", bold=True,
       note="กำไรสุทธิ + ค่าเสื่อม (ค่าเสื่อมไม่ใช่เงินสดออกจริง)")
c.skip()

c.band("8) KEY RESULTS / ตัวชี้วัดหลัก")
c.calc("payback", "ระยะคืนทุน / Payback period",
       f"=IFERROR(IF({CA('cf')}<=0,999,{CA('capex_tot')}/{CA('cf')}),999)", FMT_NUM, "เดือน", bold=True)
c.calc("payback_y", "ระยะคืนทุน (ปี)", f"={CA('payback')}/12", FMT_NUM, "ปี")
c.calc("roi", "ROI ต่อปี (ต่อเงินลงทุนรวม)",
       f"=IFERROR({CA('np')}*12/{CA('capex_tot')},0)", FMT_PCT, "%", bold=True)
c.calc("roe", "ROE ต่อปี (ต่อเงินของเจ้าของ)",
       f"=IFERROR({CA('np')}*12/{CA('equity')},0)", FMT_PCT, "%")
c.calc("bep_rev", "จุดคุ้มทุน (ยอดขาย) / Break-even revenue",
       f"=IFERROR(({IN('fix_tot')}+{CA('int')})/{CA('cm_pct')},0)", FMT_THB, "บาท/เดือน",
       note="ฐานเงินสด (ไม่รวมค่าเสื่อม)")
c.calc("bep_cyc", "จุดคุ้มทุน (รอบ/วัน) / Break-even cycles",
       f"=IFERROR({CA('bep_rev')}/{CA('p_avg')}/{IN('days')},0)", FMT_NUM, "รอบ/วัน", bold=True)
c.calc("bep_pct", "จุดคุ้มทุนคิดเป็น % ของยอดขายปัจจุบัน",
       f"=IFERROR({CA('bep_rev')}/{CA('rev_tot')},0)", FMT_PCT, "%",
       note="ยิ่งต่ำยิ่งปลอดภัย — เกิน 80% ถือว่าเสี่ยงสูง")
c.calc("dscr", "DSCR (ความสามารถชำระหนี้)",
       f"=IFERROR(IF({CA('pmt')}=0,\"N/A\",{CA('ebitda')}/{CA('pmt')}),\"N/A\")", FMT_X, "เท่า",
       note="ธนาคารต้องการ > 1.25 เท่า")
c.skip()

c.band("9) HEALTH CHECK / ตรวจสุขภาพโมเดล")
def check(key, label, formula, note):
    c.calc(key, label, formula, None, "ตรวจสอบ", note)
check("chk_cap", "กำลังผลิต / Capacity",
      f'=IF({CA("cap_chk")}>1,"เครื่องไม่พอ — ความต้องการเกินกำลังผลิต ควรเพิ่มเครื่อง/ขึ้นราคา",'
      f'IF({CA("cap_chk")}<0.4,"ลงทุนเกินความต้องการ — พิจารณาลดจำนวนเครื่อง","OK"))',
      "เป้าหมาย 60-95%")
check("chk_rent", "สัดส่วนค่าเช่า / Rent ratio",
      f'=IF(IFERROR({IN("rent")}/{CA("rev_tot")},1)>0.15,"ค่าเช่าสูงเกิน 15% ของยอดขาย — ต่อรองใหม่","OK")',
      "ค่าเช่าไม่ควรเกิน 12-15% ของยอดขาย")
check("chk_util", "สัดส่วนค่าน้ำ-ไฟ / Utility ratio",
      f'=IF(IFERROR({CA("cost_util")}/{CA("rev_tot")},1)>0.22,"ค่าน้ำ-ไฟสูงเกิน 22% — ตรวจสอบอัตราค่าไฟ/ขนาดเครื่อง","OK")',
      "ปกติ 14-20% ของยอดขาย")
check("chk_pb", "ระยะคืนทุน / Payback",
      f'=IF({CA("payback")}>36,"คืนทุนช้ากว่า 36 เดือน — ทบทวนทำเล/ค่าเช่า/ราคา",'
      f'IF({CA("payback")}>24,"คืนทุน 24-36 เดือน — รับได้แต่ควรปรับปรุง","ดี: คืนทุนภายใน 24 เดือน"))',
      "เกณฑ์ธุรกิจร้านสะดวกซัก: 18-30 เดือน")
check("chk_bep", "ความปลอดภัยจุดคุ้มทุน / BEP safety",
      f'=IF({CA("bep_pct")}>0.8,"เสี่ยงสูง — ยอดขายต้องถึง 80%+ ของประมาณการจึงจะคุ้มทุน","OK")',
      "Margin of safety = 1 - ค่านี้")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

# ============================================================ D_DASHBOARD
d = wb.create_sheet("D_DASHBOARD")
d.sheet_properties.tabColor = "FFC000"
d.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGH", [30, 18, 4, 30, 18, 4, 26, 18]):
    d.column_dimensions[col].width = w
d.merge_cells("A1:H1")
put(d, "A1", "D) DASHBOARD / สรุปผลการลงทุน", bold=True, size=16, color=C_WHT,
    fill=NAVY, align="center", border=False)
d.row_dimensions[1].height = 32
d.merge_cells("A2:H2")
put(d, "A2", f"=\"โครงการ: \"&{IN('proj')}&\"   |   ลูกค้า: \"&{IN('cust')}&\"   |   ทำเล: \"&{IN('loc')}&"
             f"\"   |   แพ็กเกจ: \"&{IN('pkg')}&\"   |   ผู้รับผิดชอบ: \"&{IN('owner')}",
    size=10, align="center", border=False, fill=GREY)

KPI = [
    ("เงินลงทุนรวม / Total CAPEX", CA('capex_tot'), FMT_THB, "บาท"),
    ("รายได้ต่อเดือน / Revenue", CA('rev_tot'), FMT_THB, "บาท/เดือน"),
    ("EBITDA ต่อเดือน", CA('ebitda'), FMT_THB, "บาท/เดือน"),
    ("กำไรสุทธิต่อเดือน / Net profit", CA('np'), FMT_THB, "บาท/เดือน"),
    ("อัตรากำไรสุทธิ / Net margin", CA('np_m'), FMT_PCT, "%"),
    ("ระยะคืนทุน / Payback", CA('payback'), FMT_NUM, "เดือน"),
    ("ROI ต่อปี", CA('roi'), FMT_PCT, "%/ปี"),
    ("จุดคุ้มทุน / BEP", CA('bep_cyc'), FMT_NUM, "รอบ/วัน"),
    ("Capacity check", CA('cap_chk'), FMT_PCT, "%"),
]
r0 = 4
put(d, f"A{r0}", "KEY RESULTS / ผลลัพธ์หลัก", bold=True, size=12, color=C_WHT, fill=BAND)
for ch in "BCDEFGH":
    put(d, f"{ch}{r0}", None, fill=BAND)
rr = r0 + 1
for i, (lab, ref, fmt, unit) in enumerate(KPI):
    col = ["A", "D", "G"][i % 3]
    vcol = ["B", "E", "H"][i % 3]
    row = rr + (i // 3) * 2
    put(d, f"{col}{row}", lab, size=10, bold=True, fill=KPIFIL, wrap=True)
    put(d, f"{vcol}{row}", f"={ref}", fmt=fmt, size=14, bold=True, align="right",
        fill=KPIFIL, color=DARK)
    put(d, f"{col}{row+1}", unit, size=8, color="808080")
    put(d, f"{vcol}{row+1}", None)
    d.row_dimensions[row].height = 26
rr = rr + ((len(KPI) - 1) // 3 + 1) * 2 + 1

# --- monthly P&L
put(d, f"A{rr}", "งบกำไรขาดทุนต่อเดือน (ปีที่ 1) / MONTHLY P&L", bold=True, size=12,
    color=C_WHT, fill=BAND)
for ch in "BCDEFGH":
    put(d, f"{ch}{rr}", None, fill=BAND)
rr += 1
put(d, f"A{rr}", "รายการ / Item", bold=True, color=C_WHT, fill=NAVY, align="center")
put(d, f"B{rr}", "บาท/เดือน", bold=True, color=C_WHT, fill=NAVY, align="center")
put(d, f"D{rr}", "% ของยอดขาย", bold=True, color=C_WHT, fill=NAVY, align="center")
put(d, f"E{rr}", "บาท/ปี", bold=True, color=C_WHT, fill=NAVY, align="center")
rr += 1
PL = [
    ("รายได้เครื่องซัก", CA('rev_w'), False),
    ("รายได้เครื่องอบ", CA('rev_d'), False),
    ("รายได้ O2O", CA('rev_o2o'), False),
    ("รายได้ vending + อื่นๆ", f"{CA('rev_vd')}+{CA('rev_ex')}", False),
    ("รวมรายได้ / Total revenue", CA('rev_tot'), True),
    ("ค่าน้ำ-ไฟ", f"-{CA('cost_util')}", False),
    ("วัสดุสิ้นเปลือง", f"-{CA('cost_cons')}", False),
    ("ค่าธรรมเนียม O2O + รับชำระเงิน", f"-{CA('cost_o2o')}-{CA('cost_epay')}", False),
    ("ค่าซ่อม + ค่าสิทธิ์ + ต้นทุนสินค้า", f"-{CA('cost_rep')}-{CA('cost_roy')}-{CA('cost_cogs')}", False),
    ("รวมต้นทุนผันแปร", f"-{CA('var_tot')}", True),
    ("กำไรส่วนเกิน / Contribution margin", CA('cm'), True),
    ("ค่าเช่า", f"-{IN('rent')}", False),
    ("ค่าจ้างพนักงาน", f"-{IN('staff')}", False),
    ("ค่าบำรุงรักษาตามสัญญา", f"-{R('A_INPUT','maint')}", False),
    ("ค่าใช้จ่ายคงที่อื่นๆ", f"-({IN('fix_tot')}-{IN('rent')}-{IN('staff')}-{R('A_INPUT','maint')})", False),
    ("รวมค่าใช้จ่ายคงที่", f"-{IN('fix_tot')}", True),
    ("EBITDA", CA('ebitda'), True),
    ("ค่าเสื่อมราคา", f"-{CA('dep')}", False),
    ("ดอกเบี้ยจ่าย", f"-{CA('int')}", False),
    ("ภาษีเงินได้", f"-{CA('tax')}", False),
    ("กำไรสุทธิ / NET PROFIT", CA('np'), True),
]
for lab, ref, bold in PL:
    fill = KPIFIL if bold else None
    put(d, f"A{rr}", lab, size=10, bold=bold, fill=fill)
    put(d, f"B{rr}", f"={ref}", fmt=FMT_THB, bold=bold, align="right", fill=fill)
    put(d, f"D{rr}", f"=IFERROR(B{rr}/{CA('rev_tot')},0)", fmt=FMT_PCT, bold=bold, align="right", fill=fill)
    put(d, f"E{rr}", f"=B{rr}*12", fmt=FMT_THB, bold=bold, align="right", fill=fill)
    rr += 1
REF[("D_DASHBOARD", "pl_end")] = rr
rr += 1

put(d, f"A{rr}", "HEALTH CHECK / ตรวจสุขภาพโมเดล", bold=True, size=12, color=C_WHT, fill=BAND)
for ch in "BCDEFGH":
    put(d, f"{ch}{rr}", None, fill=BAND)
rr += 1
for key, lab in [("chk_cap", "กำลังผลิต"), ("chk_rent", "ค่าเช่า"), ("chk_util", "ค่าน้ำ-ไฟ"),
                 ("chk_pb", "ระยะคืนทุน"), ("chk_bep", "จุดคุ้มทุน")]:
    put(d, f"A{rr}", lab, size=10, bold=True)
    d.merge_cells(f"B{rr}:H{rr}")
    put(d, f"B{rr}", f"={CA(key)}", size=10, align="left")
    for ch in "CDEFGH":
        d[f"{ch}{rr}"].border = BOX
    rr += 1
from openpyxl.formatting.rule import FormulaRule
_hc0 = REF[("D_DASHBOARD", "pl_end")] + 2
_hc_rng = f"A{_hc0}:H{_hc0+4}"
d.conditional_formatting.add(_hc_rng, FormulaRule(
    formula=[f'AND($B{_hc0}<>"",ISERROR(SEARCH("OK",$B{_hc0})),ISERROR(SEARCH("ดี",$B{_hc0})))'],
    fill=PatternFill("solid", fgColor=WARN), stopIfTrue=False))
d.conditional_formatting.add(_hc_rng, FormulaRule(
    formula=[f'OR(NOT(ISERROR(SEARCH("OK",$B{_hc0}))),NOT(ISERROR(SEARCH("ดี",$B{_hc0}))))'],
    fill=PatternFill("solid", fgColor=OKFILL), stopIfTrue=False))
d.freeze_panes = "A4"

# ============================================================ E_CASHFLOW
e = wb.create_sheet("E_CASHFLOW")
e.sheet_properties.tabColor = "00B0F0"
e.sheet_view.showGridLines = False
HEAD = ["เดือน", "ปี", "รายได้", "ต้นทุนผันแปร", "ค่าใช้จ่ายคงที่", "EBITDA", "ค่าเสื่อม",
        "เงินกู้ต้นงวด", "ดอกเบี้ย", "เงินต้น", "เงินกู้ปลายงวด", "กำไรก่อนภาษี", "ภาษี",
        "กำไรสุทธิ", "มูลค่าคงเหลือ", "FCF โครงการ", "FCF สะสม", "CF เจ้าของ",
        "Discount factor", "DCF", "DCF สะสม"]
for i, h in enumerate(HEAD):
    col = get_column_letter(i + 1)
    e.column_dimensions[col].width = 15 if i > 1 else 9
e.merge_cells("A1:U1")
put(e, "A1", "E) CASH FLOW 60 เดือน / MONTHLY CASH FLOW & IRR", bold=True, size=14,
    color=C_WHT, fill=NAVY, align="left", border=False)
e.row_dimensions[1].height = 26

SUM_R = 3
put(e, "A3", "NPV (บาท)", bold=True, fill=KPIFIL)
put(e, "C3", "IRR ต่อปี (โครงการ)", bold=True, fill=KPIFIL)
put(e, "E3", "IRR ต่อปี (เจ้าของ)", bold=True, fill=KPIFIL)
put(e, "G3", "คืนทุน (เดือน) แบบเงินสด", bold=True, fill=KPIFIL, wrap=True)
put(e, "I3", "คืนทุน (เดือน) แบบคิดลด", bold=True, fill=KPIFIL, wrap=True)
put(e, "K3", "กำไรสะสม 5 ปี", bold=True, fill=KPIFIL)

HDR_ROW = 6
for i, h in enumerate(HEAD):
    put(e, f"{get_column_letter(i+1)}{HDR_ROW}", h, bold=True, color=C_WHT, fill=NAVY,
        align="center", wrap=True, size=9)
e.row_dimensions[HDR_ROW].height = 34
R0 = HDR_ROW + 1          # month 0 row
MONTHS = 60
LAST = R0 + MONTHS

# month 0
put(e, f"A{R0}", 0, fmt=FMT_INT, align="center", size=9)
put(e, f"B{R0}", 0, fmt=FMT_INT, align="center", size=9)
for col in "CDEFGHIJKLMNO":
    put(e, f"{col}{R0}", 0, fmt=FMT_THB, align="right", size=9)
put(e, f"H{R0}", f"={CA('loan')}", fmt=FMT_THB, align="right", size=9)
put(e, f"K{R0}", f"={CA('loan')}", fmt=FMT_THB, align="right", size=9)
put(e, f"P{R0}", f"=-{CA('capex_tot')}", fmt=FMT_THB, align="right", size=9, bold=True)
put(e, f"Q{R0}", f"=P{R0}", fmt=FMT_THB, align="right", size=9)
put(e, f"R{R0}", f"=-{CA('equity')}", fmt=FMT_THB, align="right", size=9, bold=True)
put(e, f"S{R0}", 1, fmt='0.0000', align="right", size=9)
put(e, f"T{R0}", f"=P{R0}*S{R0}", fmt=FMT_THB, align="right", size=9)
put(e, f"U{R0}", f"=T{R0}", fmt=FMT_THB, align="right", size=9)

H12 = f"{IN('horizon')}*12"
for m in range(1, MONTHS + 1):
    r = R0 + m
    pr = r - 1
    alt = OKFILL if m % 12 == 0 else None
    put(e, f"A{r}", m, fmt=FMT_INT, align="center", size=9, fill=alt)
    put(e, f"B{r}", f"=ROUNDUP(A{r}/12,0)", fmt=FMT_INT, align="center", size=9, fill=alt)
    put(e, f"C{r}", f"=IF(A{r}>{H12},0,{CA('rev_tot')}*(1+{IN('growth')})^(B{r}-1))",
        fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"D{r}", f"=-C{r}*{CA('var_pct')}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"E{r}", f"=IF(A{r}>{H12},0,-(({IN('fix_tot')}-{IN('rent')})+{IN('rent')}*(1+{IN('rent_esc')})^(B{r}-1)))",
        fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"F{r}", f"=C{r}+D{r}+E{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"G{r}", f"=IF(A{r}>{H12},0,IF(A{r}<={IN('dep_y_m')}*12,IFERROR({CA('capex_m')}/{IN('dep_y_m')}/12,0),0)"
                    f"+IF(A{r}<={IN('dep_y_f')}*12,IFERROR({CA('capex_f')}/{IN('dep_y_f')}/12,0),0))",
        fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"H{r}", f"=K{pr}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"I{r}", f"=H{r}*{IN('loan_r')}/12", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"J{r}", f"=IF(H{r}<=0,0,MIN(H{r},{CA('pmt')}-I{r}))", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"K{r}", f"=MAX(0,H{r}-J{r})", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"L{r}", f"=F{r}-G{r}-I{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"M{r}", f"=MAX(0,L{r})*{IN('tax_r')}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"N{r}", f"=L{r}-M{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"O{r}", f"=IF(A{r}={H12},MAX(0,{CA('capex_m')}*(1-A{r}/12/{IN('dep_y_m')}))"
                    f"+MAX(0,{CA('capex_f')}*(1-A{r}/12/{IN('dep_y_f')}))+{CA('deposit')}+{IN('wcap')},0)",
        fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"P{r}", f"=F{r}-M{r}+O{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"Q{r}", f"=Q{pr}+P{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"R{r}", f"=N{r}+G{r}-J{r}+O{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"S{r}", f"=1/(1+{IN('disc_r')}/12)^A{r}", fmt='0.0000', align="right", size=9, fill=alt)
    put(e, f"T{r}", f"=P{r}*S{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"U{r}", f"=U{pr}+T{r}", fmt=FMT_THB, align="right", size=9, fill=alt)
    put(e, f"W{r}", f"=IF(AND(Q{r}>=0,Q{pr}<0),A{r},0)", fmt=FMT_INT, size=8, color="BFBFBF")
    put(e, f"X{r}", f"=IF(AND(U{r}>=0,U{pr}<0),A{r},0)", fmt=FMT_INT, size=8, color="BFBFBF")
put(e, f"W{HDR_ROW}", "flag คืนทุนเงินสด", size=8, color="BFBFBF", bold=True)
put(e, f"X{HDR_ROW}", "flag คืนทุนคิดลด", size=8, color="BFBFBF", bold=True)
e.column_dimensions["W"].width = 12
e.column_dimensions["X"].width = 12

put(e, "B3", f"=U{LAST}", fmt=FMT_THB, bold=True, align="right", fill=KPIFIL, size=12)
put(e, "D3", f"=IFERROR((1+IRR(P{R0}:P{LAST},0.01))^12-1,\"N/A\")", fmt=FMT_PCT, bold=True,
    align="right", fill=KPIFIL, size=12)
put(e, "F3", f"=IFERROR((1+IRR(R{R0}:R{LAST},0.01))^12-1,\"N/A\")", fmt=FMT_PCT, bold=True,
    align="right", fill=KPIFIL, size=12)
put(e, "H3", f"=IF(SUM(W{R0}:W{LAST})=0,\"เกิน 60 ด.\",SUM(W{R0}:W{LAST}))", fmt=FMT_NUM,
    bold=True, align="right", fill=KPIFIL, size=12)
put(e, "J3", f"=IF(SUM(X{R0}:X{LAST})=0,\"เกิน 60 ด.\",SUM(X{R0}:X{LAST}))", fmt=FMT_NUM,
    bold=True, align="right", fill=KPIFIL, size=12)
put(e, "L3", f"=SUM(N{R0}:N{LAST})", fmt=FMT_THB, bold=True, align="right", fill=KPIFIL, size=12)
put(e, "A4", "NPV คิดจาก FCF โครงการ ด้วยอัตราคิดลดใน A_INPUT | IRR เจ้าของคิดจากกระแสเงินสดหลังชำระหนี้ | "
             "เดือนสุดท้ายของ horizon รวมมูลค่าคงเหลือของเครื่อง + เงินประกันค่าเช่า + เงินทุนหมุนเวียนคืน",
    italic=True, size=9, color="595959", border=False)
e.freeze_panes = f"C{HDR_ROW+1}"
REF[("E_CASHFLOW", "npv")]  = "'E_CASHFLOW'!$B$3"
REF[("E_CASHFLOW", "irr")]  = "'E_CASHFLOW'!$D$3"
REF[("E_CASHFLOW", "irre")] = "'E_CASHFLOW'!$F$3"
REF[("E_CASHFLOW", "pb")]   = "'E_CASHFLOW'!$H$3"
REF[("E_CASHFLOW", "dpb")]  = "'E_CASHFLOW'!$J$3"
REF[("E_CASHFLOW", "cum")]  = "'E_CASHFLOW'!$L$3"

# ============================================================ G_SCENARIO
g = wb.create_sheet("G_SCENARIO")
g.sheet_properties.tabColor = "ED7D31"
g.sheet_view.showGridLines = False
for col, w in zip("ABCDEFGHI", [26, 14, 14, 14, 16, 16, 16, 14, 14]):
    g.column_dimensions[col].width = w
g.merge_cells("A1:I1")
put(g, "A1", "G) SCENARIO & SENSITIVITY / วิเคราะห์สถานการณ์และความอ่อนไหว",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
g.row_dimensions[1].height = 26

def scen(capm, prcm, rentm):
    """สร้างสูตรชุดผลลัพธ์ตามตัวคูณที่กำหนด (อ้างอิงฐานจาก C_CALC)"""
    dem  = f"{CA('dem_tot')}*{capm}"
    bil  = f"MIN({dem},{CA('cap_tot')})"
    rev  = f"({bil}*{CA('p_avg')}*{prcm}+{CA('rev_vd')}*{capm}+{CA('rev_ex')})"
    fix  = f"({IN('fix_tot')}+{IN('rent')}*({rentm}-1))"
    ebd  = f"({rev}*(1-{CA('var_pct')})-{fix})"
    ebt  = f"({ebd}-{CA('dep')}-{CA('int')})"
    npf  = f"({ebt}-MAX(0,{ebt})*{IN('tax_r')})"
    return rev, ebd, npf

put(g, "A3", "1) SCENARIO ANALYSIS / สถานการณ์จำลอง", bold=True, size=12, color=C_WHT, fill=BAND)
for ch in "BCDEFGHI":
    put(g, f"{ch}3", None, fill=BAND)
SC_HDR = ["สถานการณ์", "ตัวคูณลูกค้า", "ตัวคูณราคา", "ตัวคูณค่าเช่า",
          "รายได้/เดือน", "EBITDA/เดือน", "กำไรสุทธิ/เดือน", "คืนทุน (เดือน)", "ROI/ปี"]
for i, h in enumerate(SC_HDR):
    put(g, f"{get_column_letter(i+1)}4", h, bold=True, color=C_WHT, fill=NAVY, align="center", wrap=True, size=9)
g.row_dimensions[4].height = 30
SCEN = [("Worst / แย่ที่สุด", 0.70, 0.90, 1.10, WARN),
        ("Base / ฐาน",        1.00, 1.00, 1.00, KPIFIL),
        ("Best / ดีที่สุด",    1.30, 1.10, 1.00, OKFILL)]
for i, (name, cm_, pm_, rm_, fl) in enumerate(SCEN):
    r = 5 + i
    put(g, f"A{r}", name, bold=True, fill=fl)
    put(g, f"B{r}", cm_, fmt=FMT_PCT, color=C_IN, fill=YELLOW, align="right")
    put(g, f"C{r}", pm_, fmt=FMT_PCT, color=C_IN, fill=YELLOW, align="right")
    put(g, f"D{r}", rm_, fmt=FMT_PCT, color=C_IN, fill=YELLOW, align="right")
    rev, ebd, npf = scen(f"$B{r}", f"$C{r}", f"$D{r}")
    put(g, f"E{r}", f"={rev}", fmt=FMT_THB, align="right", fill=fl)
    put(g, f"F{r}", f"={ebd}", fmt=FMT_THB, align="right", fill=fl)
    put(g, f"G{r}", f"={npf}", fmt=FMT_THB, align="right", bold=True, fill=fl)
    put(g, f"H{r}", f"=IFERROR(IF(G{r}+{CA('dep')}<=0,999,{CA('capex_tot')}/(G{r}+{CA('dep')})),999)",
        fmt=FMT_NUM, align="right", bold=True, fill=fl)
    put(g, f"I{r}", f"=IFERROR(G{r}*12/{CA('capex_tot')},0)", fmt=FMT_PCT, align="right", fill=fl)
put(g, "A9", "ตัวคูณ 100% = ค่าฐานจาก A_INPUT | Worst = ลูกค้าเหลือ 70% ราคาลด 10% ค่าเช่าขึ้น 10%",
    italic=True, size=9, color="595959", border=False)

# --- 2-way sensitivity: payback
CAPS  = [0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50]
PRICE = [0.80, 0.90, 1.00, 1.10, 1.20]
def grid(start_row, title, out, fmt, note):
    put(g, f"A{start_row}", title, bold=True, size=12, color=C_WHT, fill=BAND)
    for ch in "BCDEFGHI":
        put(g, f"{ch}{start_row}", None, fill=BAND)
    hr = start_row + 1
    put(g, f"A{hr}", "ราคา \\ % ลูกค้าที่ใช้บริการ", bold=True, color=C_WHT, fill=NAVY,
        align="center", wrap=True, size=9)
    for j, cp in enumerate(CAPS):
        put(g, f"{get_column_letter(j+2)}{hr}", cp, fmt=FMT_PCT, bold=True,
            color=C_WHT, fill=NAVY, align="center")
    for i, pm_ in enumerate(PRICE):
        r = hr + 1 + i
        put(g, f"A{r}", pm_, fmt=FMT_PCT, bold=True, color=C_WHT, fill=NAVY, align="center")
        for j, cp in enumerate(CAPS):
            col = get_column_letter(j + 2)
            capm = f"({col}${hr}/{IN('capture')})"
            rev, ebd, npf = scen(capm, f"$A{r}", "1")
            put(g, f"{col}{r}", f"={out(npf)}", fmt=fmt, align="right",
                fill=(OKFILL if (pm_ == 1.0 and abs(cp - 0.35) < 1e-9) else None))
    put(g, f"A{hr+len(PRICE)+1}", note, italic=True, size=9, color="595959", border=False)
    return hr + len(PRICE) + 3

nr = grid(11, "2) SENSITIVITY — ระยะคืนทุน (เดือน) / Payback months",
          lambda npf: f"IFERROR(IF({npf}+{CA('dep')}<=0,999,{CA('capex_tot')}/({npf}+{CA('dep')})),999)",
          FMT_NUM, "ช่องเขียว = สมมติฐานฐานปัจจุบัน | 999 = ขาดทุน คืนทุนไม่ได้ | เกณฑ์ที่ยอมรับได้ ≤ 30 เดือน")
nr = grid(nr, "3) SENSITIVITY — กำไรสุทธิต่อเดือน (บาท) / Net profit",
          lambda npf: npf, FMT_THB, "ใช้ประกอบการต่อรองค่าเช่าและตั้งราคา/รอบ")

# ============================================================ H_LOCATION_SCORE
h = wb.create_sheet("H_LOCATION_SCORE")
h.sheet_properties.tabColor = "70AD47"
h.sheet_view.showGridLines = False
for col, w in zip("ABCDE", [44, 12, 12, 14, 52]):
    h.column_dimensions[col].width = w
h.merge_cells("A1:E1")
put(h, "A1", "H) LOCATION SCORE / แบบประเมินทำเล (ทำก่อนตัดสินใจลงทุน)",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
h.row_dimensions[1].height = 26
put(h, "A2", "ให้คะแนน 1 = แย่มาก ถึง 5 = ดีมาก ในช่องสีเหลือง | น้ำหนักรวมต้องเท่ากับ 100%",
    italic=True, size=9, color="595959", border=False)
dv_sc = DV("Score", h)
for i, t in enumerate(["เกณฑ์ / Criteria", "น้ำหนัก", "คะแนน (1-5)", "คะแนนถ่วงน้ำหนัก", "แนวทางให้คะแนน"]):
    put(h, f"{get_column_letter(i+1)}4", t, bold=True, color=C_WHT, fill=NAVY, align="center", wrap=True, size=9)
h.row_dimensions[4].height = 28
CRIT = [
    ("ประชากร/จำนวนห้องพักในรัศมี 1 กม.", 0.15, 4, "5 = >5,000 คน | 3 = 3,000 คน | 1 = <1,500 คน"),
    ("สัดส่วนหอพัก/คอนโด/ห้องเช่า", 0.12, 4, "5 = >70% ของที่อยู่อาศัยโดยรอบ"),
    ("จำนวนคู่แข่งในรัศมี 500 ม.", 0.12, 3, "5 = ไม่มีคู่แข่ง | 3 = 1 ราย | 1 = 3 รายขึ้นไป"),
    ("การมองเห็นจากถนน / ป้ายหน้าร้าน", 0.10, 4, "5 = ติดถนนหลัก เห็นชัดทั้งกลางวัน-กลางคืน"),
    ("ที่จอดรถ", 0.10, 3, "5 = จอดได้ >6 คัน หน้าร้าน"),
    ("ค่าเช่าต่อ ตร.ม. เทียบตลาด", 0.10, 3, "5 = ต่ำกว่าตลาด >20%"),
    ("ระบบไฟฟ้า 3 เฟส + โหลดเพียงพอ", 0.08, 5, "1 = ต้องขอเพิ่มมิเตอร์ (ค่าใช้จ่าย/เวลาสูง)"),
    ("แรงดันน้ำ/ปริมาณน้ำประปา", 0.07, 4, "5 = แรงดัน >2 บาร์ ไม่ต้องติดปั๊ม+แท็งก์"),
    ("การระบายอากาศ/ท่อระบายเครื่องอบ", 0.05, 4, "5 = ทำท่อออกภายนอกได้สั้น ตรง"),
    ("ความปลอดภัย + CCTV + แสงสว่างกลางคืน", 0.05, 4, "ร้าน 24 ชม. ต้องปลอดภัยสำหรับลูกค้าหญิง"),
    ("การเข้าถึง/ขนส่งสาธารณะ/ทางเดินเท้า", 0.03, 3, "5 = อยู่บนเส้นทางเดินประจำวันของลูกค้า"),
    ("เงื่อนไขสัญญาเช่า (อายุ/การต่อสัญญา)", 0.03, 4, "5 = สัญญา ≥3 ปี + สิทธิ์ต่ออายุ"),
]
CR0 = 5
for i, (name, wgt, sc, note) in enumerate(CRIT):
    r = CR0 + i
    put(h, f"A{r}", name, size=10)
    put(h, f"B{r}", wgt, fmt=FMT_PCT, color=C_IN, fill=YELLOW, align="right")
    cell = put(h, f"C{r}", sc, fmt=FMT_INT, color=C_IN, fill=YELLOW, align="center", bold=True)
    dv_sc.add(cell)
    put(h, f"D{r}", f"=B{r}*C{r}", fmt='0.00', align="right")
    put(h, f"E{r}", note, size=9, color="595959", wrap=True)
CRN = CR0 + len(CRIT)
put(h, f"A{CRN}", "รวม / TOTAL", bold=True, fill=KPIFIL)
put(h, f"B{CRN}", f"=SUM(B{CR0}:B{CRN-1})", fmt=FMT_PCT, bold=True, align="right", fill=KPIFIL)
put(h, f"C{CRN}", None, fill=KPIFIL)
put(h, f"D{CRN}", f"=SUM(D{CR0}:D{CRN-1})", fmt='0.00', bold=True, align="right", fill=KPIFIL)
put(h, f"E{CRN}", "คะแนนเต็ม 5.00", size=9, color="595959", fill=KPIFIL)
put(h, f"A{CRN+2}", "ผลประเมิน / VERDICT", bold=True, size=12, fill=BAND, color=C_WHT)
for ch in "BCDE":
    put(h, f"{ch}{CRN+2}", None, fill=BAND)
h.merge_cells(f"A{CRN+3}:E{CRN+3}")
put(h, f"A{CRN+3}",
    f'=IF(ABS(B{CRN}-1)>0.001,"น้ำหนักรวมไม่เท่ากับ 100% — แก้ไขคอลัมน์น้ำหนักก่อน",'
    f'IF(D{CRN}>=4,"GO — ทำเลดีมาก เดินหน้าเจรจาสัญญาเช่า",'
    f'IF(D{CRN}>=3.5,"GO (มีเงื่อนไข) — ต่อรองค่าเช่า/แก้จุดอ่อนก่อนเซ็นสัญญา",'
    f'IF(D{CRN}>=3,"HOLD — หาทำเลเปรียบเทียบอีก 2-3 แห่ง","NO-GO — ความเสี่ยงสูงเกินไป"))))',
    bold=True, size=12, align="center", fill=OKFILL)
h.row_dimensions[CRN+3].height = 26
REF[("H_LOCATION_SCORE", "score")] = f"'H_LOCATION_SCORE'!$D${CRN}"
REF[("H_LOCATION_SCORE", "verdict")] = f"'H_LOCATION_SCORE'!$A${CRN+3}"

# ============================================================ I_PRINT_VIEW
from openpyxl.worksheet.properties import PageSetupProperties
pv = wb.create_sheet("I_PRINT_VIEW")
pv.sheet_properties.tabColor = "C00000"
pv.sheet_view.showGridLines = False
for col, w in zip("ABCDEF", [26, 22, 6, 28, 20, 6]):
    pv.column_dimensions[col].width = w

pv.merge_cells("A1:B3")
put(pv, "A1", "SAMSUNG\nCOMMERCIAL", bold=True, size=16, color=NAVY, align="center",
    wrap=True, border=False)
pv.merge_cells("D1:E3")
put(pv, "D1", "ชื่อบริษัท / ที่อยู่ / โทรศัพท์\n(แก้ไขข้อความนี้ให้เป็นข้อมูลตัวแทนจำหน่ายของท่าน)",
    size=9, color="595959", align="right", wrap=True, border=False)
pv.merge_cells("A5:E5")
put(pv, "A5", "INVESTMENT SUMMARY / สรุปการลงทุน ร้านสะดวกซัก", bold=True, size=15,
    color=C_WHT, fill="C00000", align="center", border=False)
pv.row_dimensions[5].height = 30

put(pv, "A7", "PROJECT / โครงการ", bold=True, size=11, color=C_WHT, fill=BAND)
put(pv, "B7", None, fill=BAND)
put(pv, "D7", "KEY RESULTS / ผลลัพธ์หลัก", bold=True, size=11, color=C_WHT, fill=BAND)
put(pv, "E7", None, fill=BAND)
LEFT = [("ชื่อโครงการ / Project", IN('proj'), None),
        ("ลูกค้า / Customer", IN('cust'), None),
        ("ทำเล / Location", IN('loc'), None),
        ("ประเภททำเล", IN('loctype'), None),
        ("ผู้รับผิดชอบ / Sales owner", IN('owner'), None),
        ("วันที่ / Date", IN('date'), FMT_DATE),
        ("แพ็กเกจ / Package", IN('pkg'), None),
        ("จำนวนเครื่อง / Units", CA('t_units'), FMT_INT),
        ("คะแนนทำเล / Location score", REF[("H_LOCATION_SCORE", "score")], '0.00')]
RIGHT = [("เงินลงทุนรวม / Total CAPEX", CA('capex_tot'), FMT_THB),
         ("เงินลงทุนเจ้าของ / Equity", CA('equity'), FMT_THB),
         ("รายได้/เดือน / Revenue", CA('rev_tot'), FMT_THB),
         ("EBITDA/เดือน", CA('ebitda'), FMT_THB),
         ("กำไรสุทธิ/เดือน / Net profit", CA('np'), FMT_THB),
         ("อัตรากำไรสุทธิ / Net margin", CA('np_m'), FMT_PCT),
         ("ระยะคืนทุน / Payback (เดือน)", CA('payback'), FMT_NUM),
         ("ROI ต่อปี", CA('roi'), FMT_PCT),
         ("IRR 5 ปี (โครงการ)", REF[("E_CASHFLOW", "irr")], FMT_PCT)]
for i, ((ll, lr, lf), (rl, rr_, rf)) in enumerate(zip(LEFT, RIGHT)):
    r = 8 + i
    put(pv, f"A{r}", ll, size=10)
    put(pv, f"B{r}", f"={lr}", size=10, fmt=lf, align="right")
    put(pv, f"D{r}", rl, size=10, bold=True)
    put(pv, f"E{r}", f"={rr_}", size=11, fmt=rf, align="right", bold=True, fill=KPIFIL)

r = 18
put(pv, f"A{r}", "ASSUMPTIONS / สมมติฐาน (ใช้คุยกับลูกค้า)", bold=True, size=11, color=C_WHT, fill=BAND)
for ch in "BDE":
    put(pv, f"{ch}{r}", None, fill=BAND)
ASM = [("ประชากรในพื้นที่ (คน)", IN('pop'), FMT_INT, "% ที่ใช้บริการ", IN('capture'), FMT_PCT),
       ("ความถี่ (ครั้ง/คน/เดือน)", IN('freq'), FMT_NUM, "ลูกค้าคาดการณ์ (คน/เดือน)", R('A_INPUT','cust_n'), FMT_INT),
       ("ราคา/รอบซัก (บาท)", CA('p_w'), FMT_THB2, "ราคา/รอบอบ (บาท)", CA('p_d'), FMT_THB2),
       ("เวลา/รอบซัก (นาที)", CA('cyc_w'), FMT_NUM, "เวลา/รอบอบ (นาที)", CA('cyc_d'), FMT_NUM),
       ("ชั่วโมงเปิด/วัน", IN('open_h'), FMT_NUM, "วันเปิด/เดือน", IN('days'), FMT_INT),
       ("% ใช้งานวันธรรมดา", IN('wd_util'), FMT_PCT, "% ใช้งานวันหยุด", IN('we_util'), FMT_PCT),
       ("ค่าเช่า (บาท/เดือน)", IN('rent'), FMT_THB, "ค่าจ้างพนักงาน (บาท/เดือน)", IN('staff'), FMT_THB),
       ("รอบที่ขายได้ (รอบ/เดือน)", CA('bil_tot'), FMT_INT, "Capacity check", CA('cap_chk'), FMT_PCT),
       ("จุดคุ้มทุน (รอบ/วัน)", CA('bep_cyc'), FMT_NUM, "จุดคุ้มทุน (บาท/เดือน)", CA('bep_rev'), FMT_THB)]
for i, (l1, v1, f1, l2, v2, f2) in enumerate(ASM):
    rr_ = r + 1 + i
    put(pv, f"A{rr_}", l1, size=9)
    put(pv, f"B{rr_}", f"={v1}", size=9, fmt=f1, align="right")
    put(pv, f"D{rr_}", l2, size=9)
    put(pv, f"E{rr_}", f"={v2}", size=9, fmt=f2, align="right")
r = r + len(ASM) + 2

put(pv, f"A{r}", "MONTHLY P&L / งบกำไรขาดทุนต่อเดือน", bold=True, size=11, color=C_WHT, fill=BAND)
for ch in "BDE":
    put(pv, f"{ch}{r}", None, fill=BAND)
SHORT = [("รายได้รวม / Revenue", CA('rev_tot'), True),
         ("ต้นทุนผันแปร / Variable cost", f"-{CA('var_tot')}", False),
         ("ค่าใช้จ่ายคงที่ / Fixed cost", f"-{IN('fix_tot')}", False),
         ("EBITDA", CA('ebitda'), True),
         ("ค่าเสื่อม + ดอกเบี้ย + ภาษี", f"-{CA('dep')}-{CA('int')}-{CA('tax')}", False),
         ("กำไรสุทธิ / NET PROFIT", CA('np'), True)]
for i, (lab, ref, bold) in enumerate(SHORT):
    rr_ = r + 1 + i
    fill = KPIFIL if bold else None
    put(pv, f"A{rr_}", lab, size=10, bold=bold, fill=fill)
    put(pv, f"B{rr_}", f"={ref}", size=10, fmt=FMT_THB, align="right", bold=bold, fill=fill)
    put(pv, f"D{rr_}", "% ของยอดขาย", size=9, fill=fill)
    put(pv, f"E{rr_}", f"=IFERROR(B{rr_}/{CA('rev_tot')},0)", size=10, fmt=FMT_PCT,
        align="right", bold=bold, fill=fill)
r = r + len(SHORT) + 2
pv.merge_cells(f"A{r}:E{r}")
put(pv, f"A{r}", f"={CA('chk_pb')}&\"  |  \"&{CA('chk_cap')}", bold=True, size=10,
    align="center", fill=OKFILL, wrap=True)
r += 2
put(pv, f"A{r}", "ผู้เสนอ / Proposed by", size=9, border=False)
put(pv, f"D{r}", "ผู้อนุมัติ / Approved by", size=9, border=False)
r += 3
for col in ("A", "D"):
    put(pv, f"{col}{r}", "____________________________", size=9, border=False)
    put(pv, f"{col}{r+1}", "วันที่ ____ / ____ / ______", size=9, border=False)
put(pv, f"A{r+3}", "เอกสารนี้เป็นการประมาณการเพื่อประกอบการตัดสินใจ ผลลัพธ์จริงขึ้นกับทำเล พฤติกรรมผู้บริโภค และการบริหารจัดการ",
    italic=True, size=8, color="808080", border=False)
pv.page_setup.orientation = "portrait"
pv.page_setup.paperSize = 9
pv.page_setup.fitToWidth = 1
pv.page_setup.fitToHeight = 1
pv.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
pv.print_area = f"A1:E{r+3}"
pv.page_margins.left = pv.page_margins.right = 0.4
pv.page_margins.top = pv.page_margins.bottom = 0.5

# ============================================================ 0_README
rd = wb.create_sheet("0_README")
rd.sheet_properties.tabColor = "404040"
rd.sheet_view.showGridLines = False
rd.column_dimensions["A"].width = 30
rd.column_dimensions["B"].width = 110
rd.merge_cells("A1:B1")
put(rd, "A1", "การคำนวณ ROI ร้านสะดวกซัก  Samsung Commercial", bold=True, size=16,
    color=C_WHT, fill=NAVY, align="center", border=False)
rd.row_dimensions[1].height = 34
rd.merge_cells("A2:B2")
put(rd, "A2", "LAUNDROMAT ROI MODEL — Version 1.0  |  จัดทำ 18-09-2026", size=10,
    color="595959", align="center", border=False, fill=GREY)

def rsec(r, title):
    rd.merge_cells(f"A{r}:B{r}")
    put(rd, f"A{r}", title, bold=True, size=12, color=C_WHT, fill=BAND, border=False)
    rd.row_dimensions[r].height = 22
    return r + 1
def rrow(r, k, v, color="000000", bold=False):
    put(rd, f"A{r}", k, size=10, bold=True, wrap=True)
    put(rd, f"B{r}", v, size=10, wrap=True, color=color, bold=bold)
    rd.row_dimensions[r].height = max(16, 14 * (len(v) // 95 + 1))
    return r + 1

r = 4
r = rsec(r, "1) วิธีใช้งาน 5 ขั้นตอน / HOW TO USE")
for k, v in [
    ("ขั้นที่ 1", "ไปที่ชีต H_LOCATION_SCORE ให้คะแนนทำเลก่อน — ถ้าได้ NO-GO ไม่ต้องคำนวณต่อ"),
    ("ขั้นที่ 2", "ไปที่ชีต B_PACKAGES ใส่ราคาเครื่องจริงจากใบเสนอราคา และกำหนดจำนวนเครื่องในแต่ละแพ็กเกจ"),
    ("ขั้นที่ 3", "ไปที่ชีต A_INPUT กรอกเฉพาะช่องสีเหลือง (ประชากร ค่าเช่า ราคา/รอบ ฯลฯ)"),
    ("ขั้นที่ 4", "ดูผลที่ชีต D_DASHBOARD และ E_CASHFLOW — ตรวจ HEALTH CHECK ว่ามีคำเตือนหรือไม่"),
    ("ขั้นที่ 5", "ทดสอบความเสี่ยงที่ชีต G_SCENARIO แล้วพิมพ์ชีต I_PRINT_VIEW (A4 1 หน้า) เสนอลูกค้า"),
]:
    r = rrow(r, k, v)
r += 1

r = rsec(r, "2) ความหมายของสี / COLOR LEGEND")
for k, v, col in [
    ("พื้นเหลือง + ตัวเลขสีน้ำเงิน", "ช่องกรอกข้อมูล (Input) — แก้ได้เฉพาะช่องเหล่านี้", C_IN),
    ("ตัวเลขสีดำ", "สูตรคำนวณ — ห้ามพิมพ์ทับ", "000000"),
    ("ตัวเลขสีเขียว", "ลิงก์มาจากชีตอื่น — ห้ามพิมพ์ทับ", C_LINK),
    ("พื้นฟ้าอ่อน / เขียวอ่อน", "ผลลัพธ์สำคัญ (KPI)", "000000"),
]:
    r = rrow(r, k, v, color=col, bold=True)
r += 1

r = rsec(r, "3) แผนผังชีต / SHEET MAP")
for k, v in [
    ("0_README", "หน้านี้ — คู่มือ ข้อสมมติ และสิ่งที่ต้องเติมข้อมูล"),
    ("A_INPUT", "ช่องกรอกข้อมูลทั้งหมด 8 หมวด: โครงการ / ตลาด / แพ็กเกจ / CAPEX / การดำเนินงาน / รายได้เสริม / OPEX / การเงิน-ภาษี"),
    ("B_PACKAGES", "ฐานข้อมูลเครื่อง ราคา เวลาต่อรอบ ค่าไฟ-ค่าน้ำต่อรอบ และนิยามแพ็กเกจ S/M/L/XL/CUSTOM"),
    ("C_CALC", "เครื่องคำนวณ 9 ส่วน (capacity → ราคา → รอบที่ขายได้ → รายได้ → ต้นทุน → CAPEX → P&L → KPI → health check)"),
    ("D_DASHBOARD", "สรุป KPI + งบกำไรขาดทุนต่อเดือน + คำเตือนอัตโนมัติ"),
    ("E_CASHFLOW", "กระแสเงินสด 60 เดือน + NPV + IRR + คืนทุนแบบเงินสดและแบบคิดลด + ตารางผ่อนชำระเงินกู้"),
    ("F_SERVICE", "ประกันขยายเวลา + แพ็กเกจบำรุงรักษา + ค่าเดินทางนอกพื้นที่ (ส่งค่าเฉลี่ยต่อเดือนเข้า A_INPUT)"),
    ("G_SCENARIO", "Worst / Base / Best + ตารางความอ่อนไหว 2 ทาง (ราคา x % ลูกค้า)"),
    ("H_LOCATION_SCORE", "แบบให้คะแนนทำเล 12 เกณฑ์ พร้อมข้อสรุป GO / HOLD / NO-GO"),
    ("I_PRINT_VIEW", "หน้าสรุปการลงทุน A4 1 หน้า สำหรับพิมพ์เสนอลูกค้า"),
    ("Z_LISTS", "รายการตัวเลือกของ dropdown (ซ่อนไว้ — เปิดได้ด้วยคลิกขวาที่แท็บ > Unhide)"),
]:
    r = rrow(r, k, v)
r += 1

r = rsec(r, "4) ที่มาของตัวเลขตั้งต้น / SOURCE OF DEFAULT VALUES")
for k, v, col in [
    ("โครงสร้างโมเดล", "อ้างอิงโครงสร้างฟอร์ม Franchise Maintenance Pgm. และ Laundry Crew ROI Model ที่ผู้ใช้ให้มา "
                       "(หมวด Market size / Package & Pricing / Monthly costs / Investment summary)", "000000"),
    ("% ใช้งานวันธรรมดา 35% และวันหยุด 65%", "ค่าตั้งต้นจากฟอร์มต้นฉบับที่ผู้ใช้ให้มา", "000000"),
    ("ค่าไฟ 12% ค่าน้ำ 2% ของยอดขาย", "ค่าตั้งต้นจากฟอร์มต้นฉบับที่ผู้ใช้ให้มา (โมเดล 1)", "000000"),
    ("ราคาเครื่อง / ราคาประกัน / ราคาแพ็กเกจบริการ",
     "เป็นตัวเลขสมมติเพื่อให้สูตรทำงานได้ ยังไม่ใช่ราคาจริงของ Samsung — ต้องแทนที่ด้วยใบเสนอราคาจริงก่อนใช้กับลูกค้า", "C00000"),
    ("ค่าไฟ 4.80 บาท/kWh, ค่าน้ำ 18 บาท/ลบ.ม.", "ค่าประมาณอัตราธุรกิจขนาดเล็กในไทย — ตรวจสอบกับบิลจริงของพื้นที่", "C00000"),
    ("อัตราภาษีนิติบุคคล 20%", "อัตราทั่วไป — กรณี SME (ทุนจดทะเบียน ≤5 ลบ. รายได้ ≤30 ลบ.) กำไร 300k แรกยกเว้น, 300k-3M = 15%", "000000"),
]:
    r = rrow(r, k, v, color=col)
r += 1

r = rsec(r, "5) สิ่งที่ต้องเติมก่อนใช้งานจริง / TO-DO BEFORE LIVE USE")
for k, v in [
    ("1. ราคาเครื่อง", "ใส่ราคาเครื่องซัก/อบ Samsung Commercial แต่ละรุ่นจริง ที่ชีต B_PACKAGES คอลัมน์ F"),
    ("2. สเปกเครื่อง", "เวลาต่อรอบ / kWh ต่อรอบ / ลิตรต่อรอบ จากเอกสารสเปกผู้ผลิต (คอลัมน์ E, H, I)"),
    ("3. เงื่อนไขบริการ", "ราคาประกันขยายเวลา + แพ็กเกจบำรุงรักษา + ค่าเดินทาง/กม. ที่ชีต F_SERVICE"),
    ("4. โครงสร้างแพ็กเกจ", "จำนวนเครื่องในแพ็กเกจ S/M/L/XL ให้ตรงกับที่บริษัทขายจริง"),
    ("5. ค่าตกแต่งร้าน", "ราคางานระบบไฟ 3 เฟส งานน้ำ ท่อระบายเครื่องอบ จากผู้รับเหมาจริง"),
    ("6. ข้อมูลอ้างอิงจากสาขาที่เปิดแล้ว", "% การใช้งานจริง รอบ/วัน/เครื่อง เพื่อ calibrate ค่า utilization ให้แม่นขึ้น"),
]:
    r = rrow(r, k, v)
r += 1

r = rsec(r, "6) ข้อควรระวังเชิงตัวเลข / MODEL NOTES")
for k, v in [
    ("การจำกัดด้วยกำลังผลิต", "ถ้าความต้องการเกินกำลังผลิต โมเดลจะจำกัดรายได้ด้วย Fill rate อัตโนมัติ (ดู C_CALC ส่วนที่ 3)"),
    ("โหมดรายได้", "โหมด A คิดราคาเดียวต่อรอบ (ซัก+อบ) ตามฟอร์มเดิม | โหมด B แยกเครื่องซัก-อบ ซึ่งสะท้อนกำลังผลิตจริงมากกว่า"),
    ("ระยะคืนทุน", "D_DASHBOARD ใช้สูตรง่าย (เงินลงทุน / กระแสเงินสดต่อเดือน) ส่วน E_CASHFLOW คำนวณจากกระแสเงินสดสะสมจริง "
                   "ซึ่งรวมค่าเช่าที่ปรับขึ้นและการผ่อนชำระหนี้ — ตัวเลขทั้งสองจะต่างกันเล็กน้อยเป็นเรื่องปกติ"),
    ("VAT", "ถ้าไม่จดทะเบียน VAT ระบบจะบวก VAT ค่าเครื่องเข้าเป็นต้นทุน CAPEX อัตโนมัติ"),
    ("มูลค่าคงเหลือ", "เดือนสุดท้ายของ horizon จะรวมมูลค่าตามบัญชีที่เหลือของเครื่อง + เงินประกันค่าเช่า + เงินทุนหมุนเวียนคืน"),
    ("ตัวคูณส่วนแบ่งตลาด", "สูตร 1/(1+จำนวนคู่แข่ง x 0.6) เป็นค่าประมาณ ปรับตัวคูณ 0.6 ได้ที่ชีต A_INPUT ถ้ามีข้อมูลจริง"),
]:
    r = rrow(r, k, v)

# ---------------------------------------------------------------- order & save
ORDER = ["0_README", "A_INPUT", "B_PACKAGES", "C_CALC", "D_DASHBOARD", "E_CASHFLOW",
         "F_SERVICE", "G_SCENARIO", "H_LOCATION_SCORE", "I_PRINT_VIEW", "Z_LISTS"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
for name in ORDER:
    sh = wb[name]
    if name not in ("I_PRINT_VIEW",):
        sh.page_setup.orientation = "landscape" if name in ("E_CASHFLOW", "B_PACKAGES", "G_SCENARIO") else "portrait"
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

OUT = "การคำนวณROI_ร้านสะดวกซัก_Samsung_Commercial_v1.0.xlsx"
wb.save(OUT)
print("saved:", OUT)
