# -*- coding: utf-8 -*-
"""
สร้างเทมเพลต Excel: "การคำนวณ ROI ร้านสะดวกซัก Samsung Commercial" v2.0
อิงราคาและสเปกจริงจากเอกสาร Samsung Commercial Franchise (ก.ย. 2026)
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
TODAY = datetime.date(2026, 9, 19)

# ============================================================ Z_LISTS (dropdown)
zl = wb.create_sheet("Z_LISTS")
zl.sheet_properties.tabColor = "A6A6A6"
PAY_SCAN = "เปิดสแกน PromptPay 3% (ฟรีค่าบริการระบบ)"
PAY_NONE = "ไม่เปิดสแกน (จ่ายค่าบริการระบบ)"
LISTS = {
    "A": ("PackageCode", ["S", "M", "L", "CUSTOM"]),
    "B": ("RevenueMode", ["A: รวมรอบ (ซัก+อบ ราคาเดียว)", "B: แยกเครื่องซัก / เครื่องอบ"]),
    "C": ("UtilityModel", ["1: % ของยอดขาย", "2: คำนวณจากหน่วยจริง (kWh/ลิตร)"]),
    "D": ("YesNo", ["Yes / ใช่", "No / ไม่ใช่"]),
    "E": ("PayMode", [PAY_SCAN, PAY_NONE]),
    "F": ("CapturePct", [0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50]),
    "G": ("VatRate", [0.0, 0.07]),
    "H": ("Owner", ["Krittanan", "Sales 2", "Sales 3", "อื่นๆ"]),
    "I": ("LocType", ["หน้าหอพัก/คอนโด", "ในหมู่บ้านจัดสรร", "ริมถนนหลัก/ปั๊มน้ำมัน",
                      "ใกล้ตลาด/ชุมชน", "ในห้างฯ/คอมมูนิตี้มอลล์"]),
    "J": ("Score", [1, 2, 3, 4, 5]),
}
for col, (name, vals) in LISTS.items():
    put(zl, f"{col}1", name, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
    for i, v in enumerate(vals, start=2):
        put(zl, f"{col}{i}", v, size=9,
            fmt=(FMT_PCT if name in ("CapturePct", "VatRate") else None))
    zl.column_dimensions[col].width = 30
    REF[("Z_LISTS", name)] = f"Z_LISTS!${col}$2:${col}${len(vals)+1}"
zl.sheet_state = "hidden"

def DV(name, sheet):
    dv = DataValidation(type="list", formula1="=" + REF[("Z_LISTS", name)],
                        allow_blank=True, showDropDown=False)
    sheet.add_data_validation(dv)
    return dv

SRC = "ที่มา: เอกสาร Samsung Commercial Franchise / ตารางสรุปราคาอุปกรณ์ (ก.ย. 2026)"

# ============================================================ B_PACKAGES
p = SheetBuilder(wb, "B_PACKAGES", "2E75B6")
ws = p.ws
for c, w in zip("ABCDEFGH", [46, 16, 18, 16, 16, 16, 16, 46]):
    ws.column_dimensions[c].width = w
ws.merge_cells("A1:H1")
put(ws, "A1", "B) MACHINE & PACKAGE MASTER / ฐานข้อมูลเครื่องและแพ็กเกจ  (ราคาจริง Samsung Commercial)",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
ws.row_dimensions[1].height = 26
put(ws, "A2", SRC + "  |  ราคาแพ็กเกจเป็นราคารวม VAT แล้ว  |  ช่องสีเหลือง = แก้ไขได้",
    italic=True, size=9, color="C00000", border=False)
p.r = 4

# ---- 1) MACHINE MASTER
p.band("1) MACHINE MASTER / สเปกเครื่อง Samsung Commercial", "H")
p.hdr({"A": "รหัส / รุ่น", "B": "ประเภท", "C": "ความจุ (kg)", "D": "เวลา/รอบ (นาที)",
       "E": "ไฟฟ้า kWh/รอบ", "F": "น้ำ ลิตร/รอบ", "G": "แก๊ส LPG กก./รอบ",
       "H": "อายุใช้งาน (รอบ)"}, "H")
MACHINES = [
    ("WM18 — เครื่องซัก 18 kg",      "เครื่องซัก", 18, 36, 0.55, 110, 0.00, 30000),
    ("DR14 — เครื่องอบแก๊ส 14 kg",   "เครื่องอบ",  14, 45, 0.35,   0, 0.55, 30000),
]
M_START = p.r
for m in MACHINES:
    rr = p.r
    put(ws, f"A{rr}", m[0], bold=True, size=10)
    put(ws, f"B{rr}", m[1], size=10, align="center")
    for col, val, fmt in zip("CDEFGH", m[2:], [FMT_INT, FMT_INT, '0.00', FMT_INT, '0.00', FMT_INT]):
        put(ws, f"{col}{rr}", val, fmt=fmt, color=C_IN, fill=YELLOW, align="right", bold=True, size=10)
    p.r += 1
M_END = p.r - 1
W_ROW, D_ROW = M_START, M_START + 1
REF[("B_PACKAGES", "cyc_w")]  = f"'B_PACKAGES'!$D${W_ROW}"
REF[("B_PACKAGES", "cyc_d")]  = f"'B_PACKAGES'!$D${D_ROW}"
REF[("B_PACKAGES", "kwh_w")]  = f"'B_PACKAGES'!$E${W_ROW}"
REF[("B_PACKAGES", "kwh_d")]  = f"'B_PACKAGES'!$E${D_ROW}"
REF[("B_PACKAGES", "wat_w")]  = f"'B_PACKAGES'!$F${W_ROW}"
REF[("B_PACKAGES", "gas_d")]  = f"'B_PACKAGES'!$G${D_ROW}"
REF[("B_PACKAGES", "life_w")] = f"'B_PACKAGES'!$H${W_ROW}"
put(ws, f"A{p.r}", "เวลา/รอบ และอายุ 30,000 รอบ = ค่าจากเอกสาร Samsung | kWh, ลิตร, กก.แก๊ส ต่อรอบ = ค่าประมาณ "
                   "(เครื่องอบเป็นระบบแก๊ส LPG) — ควรวัดจากบิลจริงของสาขาที่เปิดแล้ว",
    italic=True, size=9, color="C00000", border=False)
p.r += 2

# ---- 2) PACKAGE MASTER
p.band("2) PACKAGE MASTER / แพ็กเกจ Franchise (ราคารวม VAT)", "G")
put(ws, f"A{p.r}", "รหัสคอลัมน์ต้องตรงกับตัวเลือกใน A_INPUT เป๊ะ:  S = 3 คู่ | M = 4 คู่ | L = 5 คู่ (Best seller) | CUSTOM = กำหนดเอง",
    italic=True, size=9, color="595959", border=False)
p.r += 1
p.hdr({"A": "รายการ / Item", "B": "หน่วย", "C": "หมายเหตุ",
       "D": "S", "E": "M", "F": "L", "G": "CUSTOM"}, "G")
PKG_HDR_ROW = p.r - 1
REF[("B_PACKAGES", "pkg_hdr")] = f"'B_PACKAGES'!$D${PKG_HDR_ROW}:$G${PKG_HDR_ROW}"
SUMROWS = {}

def prow(key, label, unit, note, vals=None, formula=None, fmt=FMT_THB,
         bold=False, fill=None, inp=False):
    rr = p.r
    put(ws, f"A{rr}", label, size=10, bold=bold, fill=fill)
    put(ws, f"B{rr}", unit, size=9, color="595959", align="center", fill=fill)
    put(ws, f"C{rr}", note, size=8, color="595959", wrap=True, fill=fill)
    for i, col in enumerate("DEFG"):
        v = formula(col) if formula else vals[i]
        put(ws, f"{col}{rr}", v, fmt=fmt, bold=bold, align="right", size=10,
            fill=(YELLOW if inp else fill), color=(C_IN if inp else C_CALC))
    SUMROWS[key] = rr
    p.r += 1
    return rr

prow("pairs", "จำนวนคู่ (ซัก+อบ) / Pairs", "คู่", "1 คู่ = ซัก 1 + อบ 1 (วางซ้อน Stack)",
     vals=[3, 4, 5, 4], fmt=FMT_INT, inp=True)
prow("w_units", "เครื่องซัก 18 kg / Washers", "เครื่อง", "",
     formula=lambda c: f"={c}{SUMROWS['pairs']}", fmt=FMT_INT)
prow("d_units", "เครื่องอบแก๊ส 14 kg / Dryers", "เครื่อง", "",
     formula=lambda c: f"={c}{SUMROWS['pairs']}", fmt=FMT_INT)
prow("t_units", "รวมเครื่อง (ระบบนับ 1 Stack = 2 เครื่อง)", "เครื่อง",
     "ใช้คิดค่าบริการรายเดือน 100 บาท/เครื่อง",
     formula=lambda c: f"={c}{SUMROWS['w_units']}+{c}{SUMROWS['d_units']}", fmt=FMT_INT, bold=True)
prow("price_pkg", "ราคาแพ็กเกจ Franchise (รวม VAT)", "บาท",
     "S 699,000 | M 799,000 | L 899,000 — รวมเครื่อง ติดตั้ง งานไฟ-น้ำ-แก๊ส งานตกแต่ง Built-in "
     "อุปกรณ์แถม และระบบ I'M CONTROL 1 ปี",
     vals=[699000, 799000, 899000, 799000], bold=True, fill=KPIFIL, inp=True)
prow("maint_first", "งานบำรุงรักษาที่รวมในแพ็กเกจ (ปีแรก)", "ครั้ง/ปี",
     "S-Built-in = 1 ครั้ง | M&L-Built-in = 2 ครั้ง", vals=[1, 2, 2, 2], fmt=FMT_INT, inp=True)
prow("warranty_y", "ประกันเครื่องซัก-อบ", "ปี", "กล่องหยอดเหรียญรับประกัน 1 ปี",
     vals=[3, 3, 3, 3], fmt=FMT_INT, inp=True)
prow("imc_free", "ระบบ I'M CONTROL + ค่าบริการ ฟรี", "ปี", "หลังหมดปีแรกดูชีต F_SERVICE",
     vals=[1, 1, 1, 1], fmt=FMT_INT, inp=True)
prow("area", "พื้นที่แนะนำ / Recommended area", "ตร.ม.",
     "งานกรุผนังในแพ็กเกจครอบคลุมไม่เกิน 25 ตร.ม.", vals=[30, 35, 45, 35], fmt=FMT_INT, inp=True)
prow("pop_rec", "ประชากรในรัศมี 1 กม. ที่แนะนำ", "คน", "",
     vals=[2500, 3500, 4500, 3500], fmt=FMT_INT, inp=True)
p.r += 1

# ---- 3) สิ่งที่ลูกค้าต้องเตรียมเอง
p.band("3) CAPEX นอกแพ็กเกจ — สิ่งที่ลูกค้าต้องเตรียมเอง (หมายเหตุข้อ 2 ของเอกสาร Samsung)", "G")
PREP = [
    ("prep_elec",  "งานเดินสายเมนไฟฟ้า + ขอ/เพิ่มขนาดมิเตอร์ไฟฟ้า", "ไม่รวมในแพ็กเกจ", [60000, 70000, 80000, 70000]),
    ("prep_water", "มิเตอร์ประปา + งานประปาหลักก่อนถึงจุดติดตั้ง", "ไม่รวมในแพ็กเกจ", [15000, 18000, 20000, 18000]),
    ("prep_base",  "แท่นวางเครื่องซัก-อบ (ฐานปูน)", "ไม่รวมในแพ็กเกจ", [25000, 30000, 35000, 30000]),
    ("prep_floor", "ฝ้าอาคาร + พื้นกระเบื้อง", "ไม่รวมในแพ็กเกจ", [40000, 50000, 60000, 50000]),
    ("prep_gas",   "ถังแก๊ส LPG 2 ถัง + เงินมัดจำถัง", "ระบบแก๊สสลับอัตโนมัติ 2 ข้าง", [8000, 8000, 8000, 8000]),
    ("prep_net",   "อินเทอร์เน็ต + ค่าติดตั้ง", "จำเป็นสำหรับ I'M CONTROL", [5000, 5000, 5000, 5000]),
    ("prep_permit","ค่าขออนุญาต / ป้าย / อื่นๆ", "", [10000, 12000, 15000, 12000]),
    ("prep_hot",   "ระบบน้ำร้อน STIEBEL ELTRON + เดินท่อน้ำร้อน", "หมายเหตุข้อ 5 — ไม่รวมในแพ็กเกจ (ใส่ 0 ถ้าไม่ติดตั้ง)", [0, 0, 0, 0]),
    ("prep_ship",  "ค่าขนส่ง + ค่าเดินทางติดตั้ง", "หมายเหตุข้อ 1 — ตามเงื่อนไขบริษัทฯ", [15000, 15000, 15000, 15000]),
]
for key, label, note, vals in PREP:
    prow(key, label, "บาท", note, vals=vals, inp=True)
prow("prep_tot", "รวม CAPEX นอกแพ็กเกจ / Total prep cost", "บาท", "",
     formula=lambda c: "=SUM({0}{1}:{0}{2})".format(c, SUMROWS["prep_elec"], SUMROWS["prep_ship"]),
     bold=True, fill=KPIFIL)
p.r += 1

# ---- 4) อุปกรณ์เสริม (ราคาจริงจากตารางสรุป)
p.band("4) อุปกรณ์เสริมซื้อเพิ่ม / Optional add-ons (ราคาขายจริง)", "H")
p.hdr({"A": "รายการอุปกรณ์", "B": "ราคาขาย (บาท)", "C": "จำนวน", "D": "รวม (บาท)",
       "E": "", "F": "", "G": "", "H": "หมายเหตุ / เงื่อนไข"}, "H")
ADDON = [
    ("กล่องไซด์บาร์หน้าจอสั่งงาน + ชุดรับเหรียญ (1 ชิ้น)", 13000, 0, "คอนโทรลแบบ All-in-One + รับเหรียญ"),
    ("กล่องไซด์บาร์ + ชุดรับเหรียญและแบงก์ (1 ชิ้น)", 15000, 0, "รับเหรียญและแบงก์"),
    ("กล่องไซด์บาร์ + ชุดรับเหรียญ (สำหรับ Stack บน-ล่าง)", 15500, 0, "สำหรับเครื่องวางซ้อน"),
    ("กล่องไซด์บาร์ + ชุดรับเหรียญและแบงก์ (Stack บน-ล่าง)", 17500, 0, "สำหรับเครื่องวางซ้อน"),
    ("เชื่อมต่อระบบสแกน & ระบบปฏิบัติการ iAm Control", 4990, 0, "ต้องมีกล่องไซด์บาร์แล้ว (ต้นทุนคู่ค้า 3,990)"),
    ("กล่องสแกนออนไลน์สำหรับอุปกรณ์หยอดเหรียญ", 5690, 0, "ติดตั้งกับอุปกรณ์หยอดเหรียญภายในร้าน"),
    ("เครื่องแลกเหรียญ รุ่นมินิ", 16990, 0, "แพ็กเกจ Franchise แถมมาแล้ว 1 เครื่อง"),
    ("เครื่องแลกเหรียญ + เซ็นทรัลเพย์เมนต์ (จอ 10 นิ้ว)", 42990, 0, "ฟรีค่าบริการเมื่อเปิดสแกน 3%"),
    ("เครื่องแลกเหรียญ + เซ็นทรัลเพย์เมนต์ (จอ 22 นิ้ว)", 49000, 0, "ฟรีค่าบริการเมื่อเปิดสแกน 3%"),
    ("เครื่องจำหน่ายสินค้า 4 ช่อง (พร้อมระบบออนไลน์)", 29990, 0, "แพ็กเกจแถมเครื่องจำหน่ายสินค้ามินิมาแล้ว"),
    ("เครื่องจำหน่ายสินค้า 8 ช่อง (พร้อมระบบออนไลน์)", 36990, 0, "จ่ายเฉพาะค่าธรรมเนียม 3% ของยอดสแกน"),
]
AD_START = p.r
for label, price, qty, note in ADDON:
    rr = p.r
    put(ws, f"A{rr}", label, size=9)
    put(ws, f"B{rr}", price, fmt=FMT_THB, color=C_IN, fill=YELLOW, align="right", size=9)
    put(ws, f"C{rr}", qty, fmt=FMT_INT, color=C_IN, fill=YELLOW, align="right", size=9, bold=True)
    put(ws, f"D{rr}", f"=B{rr}*C{rr}", fmt=FMT_THB, align="right", size=9)
    for col in "EFG":
        put(ws, f"{col}{rr}", None, size=9)
    put(ws, f"H{rr}", note, size=8, color="595959", wrap=True)
    p.r += 1
AD_END = p.r - 1
rr = p.r
put(ws, f"A{rr}", "รวมอุปกรณ์เสริม / Total add-ons", bold=True, fill=KPIFIL)
put(ws, f"B{rr}", None, fill=KPIFIL); put(ws, f"C{rr}", None, fill=KPIFIL)
put(ws, f"D{rr}", f"=SUM(D{AD_START}:D{AD_END})", fmt=FMT_THB, bold=True, align="right", fill=KPIFIL)
for col in "EFG":
    put(ws, f"{col}{rr}", None, fill=KPIFIL)
put(ws, f"H{rr}", "ส่งเข้า A_INPUT อัตโนมัติ", size=8, color="595959", fill=KPIFIL)
REF[("B_PACKAGES", "addon_tot")] = f"'B_PACKAGES'!$D${rr}"
p.r += 2

# ---- 5) สิ่งที่รวมอยู่ในแพ็กเกจแล้ว (checklist ใช้คุยกับลูกค้า)
p.band("5) สิ่งที่รวมอยู่ในแพ็กเกจแล้ว / What's included (ใช้เป็น checklist ตอนเสนอลูกค้า)", "H")
INCL = [
    ("ระบบบริหารจัดการร้านออนไลน์", "I'M CONTROL + QR CODE 1 ปี | ลงทะเบียน QR API 1 บัญชี | ติดตั้งและสอนใช้งาน 1 ครั้ง"),
    ("งานติดตั้งเครื่อง / บำรุงรักษา", "ติดตั้งเครื่องซัก-อบ 1 งาน | บำรุงรักษา S = 1 ครั้ง, M&L = 2 ครั้ง"),
    ("งานไฟฟ้า 1 เฟส", "ตู้ควบคุมไฟฟ้า + เบรกเกอร์เมน 63A | เดินสายเข้าเครื่อง ≤8 ม. | เดินสายเข้าปั๊มน้ำ | เดินปลั๊ก "
                       "(ไม่รวมสายเมนและงานมิเตอร์)"),
    ("งานระบบท่อน้ำเข้า-น้ำทิ้ง", "ถังเก็บน้ำ 1,000 ลิตร | ปั๊มน้ำ 150 วัตต์ | ท่อน้ำเมน 1/2\" ≤8 ม. | "
                                  "ท่อน้ำทิ้ง PVC 2 1/2\" ≤8 ม. | อ่างล้างมือ"),
    ("งานระบบแก๊ส + ท่อระบายความร้อน", "ท่อลมร้อน 4\" แบบ 1:1 ≤2.5 ม./เครื่อง | ระบบแก๊สสลับใช้อัตโนมัติ 2 ข้าง ≤4 ม."),
    ("งานตกแต่ง", "กรุผนังกันชื้น + ทาสี ≤25 ตร.ม. (พร้อมแบบ 3D) | ป้ายกล่องไฟ 60x60 ซม. | ป้ายขั้นตอนใช้งาน 3 ป้าย | "
                  "ไฟส่องสว่าง 6 + 2 ดวง"),
    ("อุปกรณ์พิเศษในร้าน", "เครื่องแลกเหรียญ Go Center มินิ 1 | เครื่องจำหน่ายสินค้า Go Center มินิ 1 | "
                           "กล้องวงจรปิด Wi-Fi 1 | โต๊ะ+เก้าอี้ 2 ชุด | พัดลม 1"),
    ("การรับประกัน", "เครื่องซัก-อบ 3 ปี | กล่องหยอดเหรียญ 1 ปี (ตามเงื่อนไขการใช้งานปกติ)"),
]
for title, detail in INCL:
    rr = p.r
    put(ws, f"A{rr}", title, size=10, bold=True)
    ws.merge_cells(f"B{rr}:H{rr}")
    put(ws, f"B{rr}", detail, size=9, color="404040", wrap=True)
    for col in "CDEFGH":
        ws[f"{col}{rr}"].border = BOX
    ws.row_dimensions[rr].height = max(16, 13 * (len(detail) // 100 + 1))
    p.r += 1
ws.sheet_view.showGridLines = False

def pk(rowkey):
    """INDEX/MATCH ตามรหัสแพ็กเกจที่เลือกใน A_INPUT"""
    return (f"INDEX('B_PACKAGES'!$D${SUMROWS[rowkey]}:$G${SUMROWS[rowkey]},"
            f"MATCH({IN('pkg')},{REF[('B_PACKAGES','pkg_hdr')]},0))")

# ============================================================ A_INPUT
a = SheetBuilder(wb, "A_INPUT", "C00000")
ws = a.ws
for c, w in zip("ABCDE", [58, 20, 18, 16, 56]):
    ws.column_dimensions[c].width = w
ws.merge_cells("A1:E1")
put(ws, "A1", "การคำนวณ ROI ร้านสะดวกซัก  Samsung Commercial  /  LAUNDROMAT ROI MODEL  v2.0",
    bold=True, size=15, color=C_WHT, fill=NAVY, align="center", border=False)
ws.row_dimensions[1].height = 30
ws.merge_cells("A2:E2")
put(ws, "A2", "กรอกเฉพาะช่องสีเหลือง (Input) → ผลลัพธ์อัปเดตอัตโนมัติ  •  ช่อง override ใส่ 0 = ใช้ค่าจากแพ็กเกจ  •  "
              "ราคาแพ็กเกจ Samsung เป็นราคารวม VAT และรวมงานติดตั้ง/ตกแต่งแล้ว",
    italic=True, size=9, color="595959", align="center", border=False, fill=GREY)
a.r = 4
dv_pkg, dv_mode, dv_util = DV("PackageCode", ws), DV("RevenueMode", ws), DV("UtilityModel", ws)
dv_yn, dv_cap, dv_vat = DV("YesNo", ws), DV("CapturePct", ws), DV("VatRate", ws)
dv_owner, dv_loc, dv_pay = DV("Owner", ws), DV("LocType", ws), DV("PayMode", ws)

a.band("1) PROJECT INFO / ข้อมูลโครงการ")
a.inp("proj",   "ชื่อโครงการ / Project name", "Laundromat_สาขาตัวอย่าง", None, "Input", "ใช้แสดงบนหน้า I_PRINT_VIEW")
a.inp("cust",   "ลูกค้า / Customer", "K.ตัวอย่าง", None, "Input")
a.inp("loc",    "ทำเล / Location", "ระบุชื่อทำเล", None, "Input")
a.inp("loctype","ประเภททำเล / Location type", "หน้าหอพัก/คอนโด", None, "เลือก",
      "มีผลกับสมมติฐานความถี่การใช้บริการ", dv=dv_loc)
a.inp("owner",  "ผู้รับผิดชอบ / Sales owner", "Krittanan", None, "เลือก", dv=dv_owner)
a.inp("date",   "วันที่ / Date", TODAY, FMT_DATE, "Input")
a.skip()

a.band("2) MARKET SIZE & POPULATION / ขนาดตลาดและประชากร")
a.inp("pop",     "จำนวนประชากรในรัศมีบริการ (1 กม.)", 3000, FMT_INT, "Input",
      "นับจากจำนวนห้องพัก x คนเฉลี่ย/ห้อง หรือข้อมูลทะเบียนราษฎร์",
      comment="แนะนำ: นับห้องหอพัก/คอนโดในรัศมี 1 กม. x 1.5 คน/ห้อง")
a.inp("capture", "% คาดว่าจะใช้บริการ / Capture rate", 0.35, FMT_PCT, "เลือก/แก้ได้",
      "ทำเลหอพัก 30-40% | หมู่บ้านจัดสรร 15-25%", dv=dv_cap)
a.inp("comp",    "จำนวนคู่แข่งในรัศมี 500 ม.", 1, FMT_INT, "Input", "0 = ไม่มีคู่แข่ง")
a.calc("share",  "ตัวคูณส่วนแบ่งตลาด / Market share factor",
       f"=IFERROR(1/(1+{IN('comp')}*0.6),1)", FMT_PCT, "Formula",
       "สูตร 1/(1+คู่แข่ง x 0.6) — ปรับตัวเลข 0.6 ได้ถ้ามีข้อมูลจริง")
a.calc("cust_n", "จำนวนลูกค้าคาดการณ์ / Estimated customers",
       f"={IN('pop')}*{IN('capture')}*{R('A_INPUT','share')}", FMT_INT, "Formula", "คน/เดือน")
a.inp("freq",    "ความถี่ใช้บริการ / Usage frequency", 3, FMT_NUM, "Input", "ครั้ง/คน/เดือน")
a.inp("cyc_visit","รอบซักต่อการมา 1 ครั้ง", 1.2, FMT_NUM, "Input", "ลูกค้าครอบครัวมักซัก 1.5-2 รอบ/ครั้ง")
a.inp("dry_att", "% ที่ใช้เครื่องอบด้วย / Dryer attach rate", 0.75, FMT_PCT, "Input",
      "เครื่องอบแก๊ส 45 นาที — หน้าฝน/คอนโดสูงถึง 85%")
a.calc("dem_w",  "ความต้องการรอบซัก / Wash cycles demand",
       f"={R('A_INPUT','cust_n')}*{IN('freq')}*{IN('cyc_visit')}", FMT_INT, "Formula", "รอบ/เดือน")
a.calc("dem_d",  "ความต้องการรอบอบ / Dry cycles demand",
       f"={R('A_INPUT','dem_w')}*{IN('dry_att')}", FMT_INT, "Formula", "รอบ/เดือน")
a.skip()

a.band("3) PACKAGE & PRICING / แพ็กเกจและราคา")
a.inp("pkg",  "เลือกแพ็กเกจ / Package code", "L", None, "เลือก",
      "S = 3 คู่ 699,000 | M = 4 คู่ 799,000 | L = 5 คู่ 899,000 (Best seller) | CUSTOM = กำหนดเองที่ B_PACKAGES", dv=dv_pkg)
a.inp("mode", "รูปแบบการคิดรายได้ / Revenue model", "B: แยกเครื่องซัก / เครื่องอบ", None, "เลือก",
      "โหมด A = ราคาเดียวต่อรอบ (ซัก+อบ) | โหมด B = แยกราคาซัก/อบ (แนะนำ เพราะเวลารอบต่างกัน 36 vs 45 นาที)", dv=dv_mode)
a.calc("w_pkg", "เครื่องซักตามแพ็กเกจ", f"={pk('w_units')}", FMT_INT, "จาก B_PACKAGES", "เครื่อง", color=C_LINK)
a.calc("d_pkg", "เครื่องอบตามแพ็กเกจ",  f"={pk('d_units')}", FMT_INT, "จาก B_PACKAGES", "เครื่อง", color=C_LINK)
a.inp("w_ov",  "override จำนวนเครื่องซัก (0 = ใช้แพ็กเกจ)", 0, FMT_INT, "Input")
a.inp("d_ov",  "override จำนวนเครื่องอบ (0 = ใช้แพ็กเกจ)", 0, FMT_INT, "Input")
a.calc("w_act","เครื่องซักที่ใช้จริง", f"=IF({IN('w_ov')}>0,{IN('w_ov')},{IN('w_pkg')})", FMT_INT, "Formula", "เครื่อง")
a.calc("d_act","เครื่องอบที่ใช้จริง", f"=IF({IN('d_ov')}>0,{IN('d_ov')},{IN('d_pkg')})", FMT_INT, "Formula", "เครื่อง")
a.calc("t_act","รวมเครื่องทั้งหมด", f"={R('A_INPUT','w_act')}+{R('A_INPUT','d_act')}", FMT_INT, "Formula",
       "ใช้คิดค่าบริการรายเดือน 100 บาท/เครื่อง (1 Stack = 2 เครื่อง)", bold=True)
a.inp("price_w", "ราคา/รอบ เครื่องซัก 18 kg", 50, FMT_THB, "Input", "บาท/รอบ — ราคาตลาด 40-60 บาท")
a.inp("price_d", "ราคา/รอบ เครื่องอบ 14 kg", 50, FMT_THB, "Input", "บาท/รอบ — อบแก๊ส 45 นาที ตลาด 40-60 บาท")
a.inp("price_c", "ราคา/รอบ รวมซัก+อบ (ใช้เฉพาะโหมด A)", 100, FMT_THB, "Input", "บาท/รอบ")
a.inp("cyc_w", "เวลา/รอบซัก (0 = ใช้ค่าจาก B_PACKAGES = 36 นาที)", 0, FMT_NUM, "Input", "นาที")
a.inp("cyc_d", "เวลา/รอบอบ (0 = ใช้ค่าจาก B_PACKAGES = 45 นาที)", 0, FMT_NUM, "Input", "นาที")
a.inp("cyc_c", "เวลา/รอบ รวมซัก+อบ (ใช้เฉพาะโหมด A)", 81, FMT_NUM, "Input", "นาที (36 + 45)")
a.inp("chg",   "เวลาโหลด-ปลดผ้า + รอเครื่องว่าง", 6, FMT_NUM, "Input",
      "นาที/รอบ — ทำให้ capacity สมจริงขึ้น")
a.skip()

a.band("4) CAPEX / เงินลงทุน")
a.calc("capex_pkg_r", "ราคาแพ็กเกจ Franchise (รวม VAT)", f"={pk('price_pkg')}", FMT_THB,
       "จาก B_PACKAGES", "รวมเครื่อง ติดตั้ง งานระบบ งานตกแต่ง อุปกรณ์แถม และระบบ 1 ปี", color=C_LINK)
a.inp("capex_pkg_ov", "override ราคาแพ็กเกจ (0 = ใช้ราคาตาราง)", 0, FMT_THB, "Input", "ใส่ราคาหลังต่อรอง/ส่วนลดจริง")
a.calc("capex_prep_r","CAPEX นอกแพ็กเกจ (ลูกค้าเตรียมเอง)", f"={pk('prep_tot')}", FMT_THB,
       "จาก B_PACKAGES", "งานไฟเมน มิเตอร์ ประปา แท่นเครื่อง ฝ้า-พื้น ถังแก๊ส เน็ต ขนส่ง", color=C_LINK)
a.calc("capex_addon_r","อุปกรณ์เสริมซื้อเพิ่ม", f"={REF[('B_PACKAGES','addon_tot')]}", FMT_THB,
       "จาก B_PACKAGES", "กล่องไซด์บาร์ / เครื่องแลกเหรียญ / ตู้จำหน่ายสินค้า", color=C_LINK)
a.inp("capex_other", "CAPEX อื่นๆ / Other CAPEX", 0, FMT_THB, "Input", "เช่น ค่าเซ้งร้าน ค่าที่ปรึกษา")
a.inp("mach_share", "% ของราคาแพ็กเกจที่เป็นตัวเครื่อง", 0.7, FMT_PCT, "Input",
      "ใช้แยกตัดค่าเสื่อม: ส่วนเครื่อง 8 ปี / ส่วนงานตกแต่ง-ติดตั้ง 5 ปี")
a.inp("dep_month",  "เงินประกันค่าเช่า / Rent deposit", 3, FMT_NUM, "Input", "จำนวนเดือน (ได้คืนเมื่อเลิกสัญญา)")
a.inp("wcap",       "เงินทุนหมุนเวียน / Working capital", 150000, FMT_THB, "Input",
      "เงินสดสำรอง 2-3 เดือนของ OPEX + เงินทอน/เหรียญในเครื่องแลกเหรียญ")
a.inp("vat",        "VAT %", 0.07, FMT_PCT, "เลือก", dv=dv_vat)
a.inp("vat_reg",    "จดทะเบียน VAT? / VAT registered", "No / ไม่ใช่", None, "เลือก",
      "ถ้า Yes → ขอคืน VAT ได้ ระบบจะถอด VAT ออกจาก CAPEX (หาร 1.07)", dv=dv_yn)
a.skip()

a.band("5) OPERATION / การดำเนินงาน")
a.inp("open_h", "ชั่วโมงเปิด/วัน / Open hours per day", 24, FMT_NUM, "Input", "24 ชม. = ระบบหยอดเหรียญ/สแกนอัตโนมัติ")
a.inp("wd_util","% ใช้งานวันธรรมดา / Weekday utilization", 0.35, FMT_PCT, "Input", "ค่าอ้างอิงฟอร์มเดิม = 35%")
a.inp("we_util","% ใช้งานวันหยุด / Weekend utilization", 0.65, FMT_PCT, "Input", "ค่าอ้างอิงฟอร์มเดิม = 65%")
a.inp("wd_days","จำนวนวันธรรมดา/เดือน", 22, FMT_INT, "Input")
a.inp("we_days","จำนวนวันหยุด/เดือน", 8, FMT_INT, "Input")
a.calc("days",  "รวมวันเปิดบริการ/เดือน", f"={IN('wd_days')}+{IN('we_days')}", FMT_INT, "Formula", "วัน")
a.inp("o2o_on", "เปิดบริการรับ-ส่งผ้า O2O?", "No / ไม่ใช่", None, "เลือก", dv=dv_yn)
a.inp("o2o_ord","ออเดอร์ O2O ต่อวัน", 0, FMT_NUM, "Input", "บิล/วัน")
a.inp("o2o_cyc","รอบต่อ 1 ออเดอร์ O2O", 2, FMT_NUM, "Input", "ซัก 1 + อบ 1")
a.inp("o2o_tk", "ยอดเฉลี่ยต่อบิล O2O", 0, FMT_THB, "Input", "บาท/บิล")
a.inp("o2o_fee","ค่าธรรมเนียมแพลตฟอร์ม", 0.0, FMT_PCT, "Input", "% ของยอด O2O")
a.inp("o2o_rd", "ค่าส่ง (ไรเดอร์)", 0.0, FMT_PCT, "Input", "% ของยอด O2O")
a.skip()

a.band("6) OTHER REVENUE / รายได้เสริม")
a.inp("vend_att","% ลูกค้าที่ซื้อน้ำยา-ของใช้ (ตู้จำหน่ายสินค้า)", 0.25, FMT_PCT, "Input")
a.inp("vend_sp", "ยอดซื้อเฉลี่ย/คน", 25, FMT_THB, "Input", "บาท/ครั้ง")
a.inp("vend_gp", "%กำไรขั้นต้นสินค้า vending", 0.45, FMT_PCT, "Input")
a.inp("extra",   "รายได้อื่น (ตู้กดน้ำ/ตู้เกม/ตู้คีบ)", 0, FMT_THB, "Input",
      "บาท/เดือน — วางตู้เสริมหน้าร้านเพิ่มรายได้ต่อพื้นที่")
a.inp("extra_gp","%กำไรขั้นต้นรายได้อื่น", 0.7, FMT_PCT, "Input")
a.skip()

a.band("7) MONTHLY COSTS (OPEX) / ค่าใช้จ่ายต่อเดือน")
a.sub("7.1) FIXED COST / ค่าใช้จ่ายคงที่")
a.inp("rent",    "ค่าเช่า / Rent", 25000, FMT_THB, "Input", "บาท/เดือน")
a.inp("rent_esc","อัตราปรับค่าเช่าต่อปี", 0.05, FMT_PCT, "Input", "ปกติสัญญา 3 ปี ปรับ 5-10%")
a.inp("staff",   "ค่าจ้างพนักงาน / Staff", 12000, FMT_THB, "Input",
      "บาท/เดือน — ร้าน built-in ใช้พนักงาน part-time ดูแลความสะอาด")
a.inp("ins",     "ประกันภัย / Insurance", 1200, FMT_THB, "Input", "ประกันทรัพย์สิน + บุคคลที่ 3")
a.inp("clean",   "ทำความสะอาด + เก็บเหรียญ", 3000, FMT_THB, "Input")
a.calc("maint",  "ค่าบำรุงรักษาเครื่อง (หลังปีแรก)", 0, FMT_THB, "จาก F_SERVICE",
       "ลิงก์อัตโนมัติจากชีต F_SERVICE", color=C_LINK)
a.calc("sysfee", "ค่าบริการระบบ I'M CONTROL (หลังปีแรก)", 0, FMT_THB, "จาก F_SERVICE",
       "ลิงก์อัตโนมัติ — เป็น 0 ถ้าเลือกเปิดสแกน 3%", color=C_LINK)
a.inp("mkt",     "การตลาด / Marketing", 3000, FMT_THB, "Input", "โปรโมชั่น ป้าย โฆษณาออนไลน์")
a.inp("pos",     "ค่าอินเทอร์เน็ตรายเดือน", 800, FMT_THB, "Input", "จำเป็นต่อระบบ I'M CONTROL")
a.inp("other_f", "ค่าใช้จ่ายคงที่อื่นๆ", 0, FMT_THB, "Input")
a.calc("fix_tot","รวมค่าใช้จ่ายคงที่ / TOTAL FIXED COST",
       f"=SUM({IN('rent')},{IN('staff')},{IN('ins')},{IN('clean')},{R('A_INPUT','maint')},"
       f"{R('A_INPUT','sysfee')},{IN('mkt')},{IN('pos')},{IN('other_f')})",
       FMT_THB, "Formula", "บาท/เดือน (ยังไม่รวมค่าน้ำ-ไฟ-แก๊ส)", bold=True, fill=KPIFIL)
a.sub("7.2) VARIABLE COST / ค่าใช้จ่ายผันแปร")
a.inp("util_mode","โมเดลค่าน้ำ-ไฟ / Utilities model", "2: คำนวณจากหน่วยจริง (kWh/ลิตร)", None, "เลือก",
      "โมเดล 1 = ประมาณเป็น % ของยอดขาย | โมเดล 2 = คำนวณจาก kWh/ลิตรต่อรอบ (แม่นกว่า) — "
      "ค่าแก๊สคิดแยกทุกกรณี", dv=dv_util)
a.inp("elec_p",  "ค่าไฟ (% ของยอดขาย)", 0.12, FMT_PCT, "โมเดล 1", "ค่าอ้างอิงฟอร์มเดิม = 12%")
a.inp("water_p", "ค่าน้ำ (% ของยอดขาย)", 0.02, FMT_PCT, "โมเดล 1", "ค่าอ้างอิงฟอร์มเดิม = 2%")
a.inp("elec_r",  "ค่าไฟต่อหน่วย", 4.8, FMT_THB2, "โมเดล 2", "บาท/kWh (รวม Ft + VAT)")
a.inp("water_r", "ค่าน้ำต่อหน่วย", 18, FMT_THB2, "โมเดล 2", "บาท/ลูกบาศก์เมตร")
a.inp("gas_r",   "ค่าแก๊ส LPG ต่อกิโลกรัม", 25, FMT_THB2, "Input",
      "บาท/กก. — เครื่องอบ Samsung เป็นระบบแก๊ส คิดแยกจากค่าไฟทุกโมเดล")
a.inp("base_kwh","ค่าไฟส่วนกลาง (แอร์/ไฟ/ป้าย/CCTV)", 3500, FMT_THB, "Input", "บาท/เดือน คงที่")
a.inp("cons_cyc","วัสดุสิ้นเปลืองต่อรอบ", 2.5, FMT_THB2, "Input", "บาท/รอบ (น้ำยา ถุง ฯลฯ)")
a.inp("pay_mode","ระบบรับชำระเงิน / Payment mode", PAY_SCAN, None, "เลือก",
      "เปิดสแกน 3% → ฟรีค่าบริการรายเดือน/สาขา/ปี | ไม่เปิดสแกน → จ่ายค่าบริการระบบ (ดู F_SERVICE)", dv=dv_pay)
a.inp("scan_fee","ค่าธรรมเนียมสแกน PromptPay", 0.03, FMT_PCT, "Input", "3% คิดเฉพาะยอดที่สแกนเท่านั้น")
a.inp("scan_share","% ยอดขายที่ชำระผ่านสแกน", 0.6, FMT_PCT, "Input",
      "ส่วนที่เหลือเป็นเหรียญ/แบงก์ (ไม่เสียค่าธรรมเนียม)")
a.inp("repair_p","ค่าซ่อมผันแปร (นอกประกัน)", 0.01, FMT_PCT, "Input",
      "% ของยอดขาย — ต่ำใน 3 ปีแรกเพราะมีประกันเครื่อง")
a.inp("royal_p", "ค่าสิทธิ์/แฟรนไชส์รายเดือน", 0.0, FMT_PCT, "Input", "% ของยอดขาย (Samsung ไม่เก็บ royalty)")
a.skip()

a.band("8) FINANCE & TAX / แหล่งเงินทุนและภาษี")
a.inp("loan_p",  "สัดส่วนเงินกู้ / Loan % of CAPEX", 0.0, FMT_PCT, "Input", "0% = ลงทุนด้วยเงินสดทั้งหมด")
a.inp("loan_r",  "ดอกเบี้ยเงินกู้ต่อปี", 0.075, FMT_PCT, "Input", "สินเชื่อ SME 6.5-9%")
a.inp("loan_y",  "ระยะเวลาผ่อน", 5, FMT_INT, "Input", "ปี")
a.inp("dep_y_m", "อายุตัดค่าเสื่อมเครื่อง", 8, FMT_INT, "Input", "ปี (เส้นตรง) — ดู health check อายุตามจำนวนรอบด้วย")
a.inp("dep_y_f", "อายุตัดค่าเสื่อมงานตกแต่ง/ติดตั้ง", 5, FMT_INT, "Input", "ปี (เส้นตรง)")
a.inp("tax_r",   "อัตราภาษีเงินได้", 0.2, FMT_PCT, "Input",
      "นิติบุคคลทั่วไป 20% | SME กำไร<300k ยกเว้น, 300k-3M = 15%")
a.inp("disc_r",  "อัตราคิดลด / Discount rate (WACC)", 0.1, FMT_PCT, "Input", "ใช้คำนวณ NPV/IRR")
a.inp("growth",  "อัตราเติบโตยอดขายต่อปี", 0.05, FMT_PCT, "Input", "ปีที่ 2 เป็นต้นไป")
a.inp("horizon", "ระยะเวลาประเมิน / Evaluation horizon", 5, FMT_INT, "Input", "ปี (สูงสุด 5)")
ws.freeze_panes = "A4"
ws.sheet_view.showGridLines = False

# ============================================================ F_SERVICE
s = SheetBuilder(wb, "F_SERVICE", "7030A0")
ws = s.ws
for c, w in zip("ABCDE", [52, 20, 18, 16, 56]):
    ws.column_dimensions[c].width = w
ws.merge_cells("A1:E1")
put(ws, "A1", "F) WARRANTY, MAINTENANCE & SYSTEM FEE / ประกัน บำรุงรักษา และค่าบริการระบบ",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
ws.row_dimensions[1].height = 26
put(ws, "A2", SRC + "  |  ปีแรกทุกอย่างรวมอยู่ในราคาแพ็กเกจแล้ว — ชีตนี้คิดค่าใช้จ่ายตั้งแต่ปีที่ 2 เป็นต้นไป",
    italic=True, size=9, color="C00000", border=False)
s.r = 4

s.band("1) สิ่งที่รวมอยู่ในแพ็กเกจแล้ว / Included in package")
s.calc("wr_mach", "ประกันเครื่องซัก-อบ", f"={pk('warranty_y')}", FMT_INT, "จาก B_PACKAGES",
       "ปี — สูงสุดในไทย ตามเงื่อนไขการใช้งานปกติ", color=C_LINK)
s.inp("wr_coin",  "ประกันกล่องหยอดเหรียญ", 1, FMT_INT, "Input", "ปี")
s.calc("mt_first","งานบำรุงรักษาที่รวมในปีแรก", f"={pk('maint_first')}", FMT_INT, "จาก B_PACKAGES",
       "ครั้ง — S = 1 ครั้ง, M&L = 2 ครั้ง", color=C_LINK)
s.calc("imc_free","ระบบ I'M CONTROL + ค่าบริการ ฟรี", f"={pk('imc_free')}", FMT_INT, "จาก B_PACKAGES",
       "ปี — รวมระบบจ่ายเงิน QR CODE และ QR API 1 บัญชี", color=C_LINK)
s.skip()

s.band("2) ค่าบำรุงรักษาหลังปีแรก / Maintenance from year 2")
s.inp("mt_rate",  "ค่าบำรุงรักษาต่อครั้ง (ทั้งร้าน)", 3500, FMT_THB, "Input",
      "บาท/ครั้ง — ยังไม่ใช่ราคาทางการ ต้องยืนยันกับ Samsung")
s.inp("mt_times", "จำนวนครั้งต่อปี", 2, FMT_INT, "Input", "ครั้ง/ปี (เท่ากับที่แพ็กเกจ M&L ให้)")
s.calc("mt_year", "ค่าบำรุงรักษาต่อปี", f"={R('F_SERVICE','mt_rate')}*{R('F_SERVICE','mt_times')}",
       FMT_THB, "Formula", "บาท/ปี")
s.calc("mt_month","ค่าบำรุงรักษาเฉลี่ยต่อเดือน (ส่งไป A_INPUT)",
       f"={R('F_SERVICE','mt_year')}/12", FMT_THB, "Formula", "บาท/เดือน", bold=True, fill=OKFILL)
s.skip()

s.band("3) ค่าบริการระบบ I'M CONTROL หลังปีแรก — เปรียบเทียบ 2 ทางเลือก")
s.text("ทางเลือก A : เปิดใช้งานระบบสแกนชำระเงิน PromptPay 3%", bold=True, color="000000")
s.calc("a_fee", "ค่าบริการรายเดือน/รายสาขา/รายปี", 0, FMT_THB, "Formula",
       "ได้รับยกเว้นทั้งหมดทันทีเมื่อเปิดสแกน — ต้นทุนไปอยู่ที่ค่าธรรมเนียม 3% ของยอดสแกน (คิดในต้นทุนผันแปร)")
s.text("ทางเลือก B : ไม่เปิดสแกน (รับเหรียญ/แบงก์อย่างเดียว)", bold=True, color="000000")
s.inp("b_mach", "ค่าบริการรายเดือน ต่อเครื่อง", 100, FMT_THB, "Input", "บาท/เครื่อง/เดือน (1 Stack นับ 2 เครื่อง)")
s.inp("b_brch", "ค่าบริการรายเดือน ต่อสาขา", 700, FMT_THB, "Input", "บาท/สาขา/เดือน — เฉพาะสาขาที่มี 7 เครื่องขึ้นไป")
s.inp("b_year", "ค่าบริการรายปี", 8000, FMT_THB, "Input", "บาท/ปี — ค่าดูแลรักษาและบำรุงรักษาระบบรายปี")
s.calc("b_units","จำนวนเครื่องที่ระบบนับ", f"={IN('t_act')}", FMT_INT, "จาก A_INPUT", "เครื่อง", color=C_LINK)
s.calc("b_month","ค่าบริการระบบรวมต่อเดือน (ทางเลือก B)",
       f"={R('F_SERVICE','b_units')}*{R('F_SERVICE','b_mach')}"
       f"+IF({R('F_SERVICE','b_units')}>=7,{R('F_SERVICE','b_brch')},0)"
       f"+{R('F_SERVICE','b_year')}/12", FMT_THB, "Formula", "บาท/เดือน", bold=True)
s.calc("b_break","ยอดสแกนที่ทำให้สองทางเลือกเท่ากัน",
       f"=IFERROR({R('F_SERVICE','b_month')}/{IN('scan_fee')},0)", FMT_THB, "Formula",
       "บาท/เดือน — ถ้ายอดที่ลูกค้าสแกนจริงสูงกว่านี้ ทางเลือก B (ไม่เปิดสแกน) จะถูกกว่าในแง่ค่าธรรมเนียม")
s.calc("scan_act", "ยอดสแกนจริงต่อเดือน (ประมาณการ)", 0, FMT_THB, "จาก C_CALC",
       "รายได้รวม x %ที่ชำระผ่านสแกน (ลิงก์อัตโนมัติ)", color=C_LINK)
s.calc("fee_act", "ค่าธรรมเนียม 3% ที่ต้องจ่าย (ทางเลือก A)",
       f"={R('F_SERVICE','scan_act')}*{IN('scan_fee')}", FMT_THB, "Formula", "บาท/เดือน")
s.calc("advice", "คำแนะนำ / Recommendation",
       f'=IF({R("F_SERVICE","fee_act")}>{R("F_SERVICE","b_month")},'
       f'"ทางเลือก B ถูกกว่า "&TEXT({R("F_SERVICE","fee_act")}-{R("F_SERVICE","b_month")},"#,##0")'
       f'&" บาท/เดือน — แต่ถ้าไม่เปิดสแกน ลูกค้าจ่ายได้เฉพาะเหรียญ/แบงก์ ซึ่งมักทำให้ยอดขายลดลงมากกว่าที่ประหยัดได้",'
       f'"ทางเลือก A (เปิดสแกน 3%) ถูกกว่า "&TEXT({R("F_SERVICE","b_month")}-{R("F_SERVICE","fee_act")},"#,##0")'
       f'&" บาท/เดือน และสะดวกกับลูกค้ามากกว่า")', None, "Formula",
       "เทียบค่าธรรมเนียมจริงกับค่าบริการระบบรายเดือน")
s.calc("sys_month","ค่าบริการระบบที่ใช้จริง (ส่งไป A_INPUT)",
       f'=IF({IN("pay_mode")}="{PAY_SCAN}",{R("F_SERVICE","a_fee")},{R("F_SERVICE","b_month")})',
       FMT_THB, "Formula", "บาท/เดือน", bold=True, fill=OKFILL)
s.skip()

s.band("4) SUMMARY / สรุป")
s.calc("free_y1", "ส่วนลดปีแรก (รวมอยู่ในแพ็กเกจแล้ว)",
       f"={R('F_SERVICE','mt_month')}+{R('F_SERVICE','sys_month')}", FMT_THB, "Formula",
       "บาท/เดือน — ชีต E_CASHFLOW จะหักค่านี้ออกจากค่าใช้จ่ายคงที่ของปีที่ 1 ให้อัตโนมัติ", bold=True)
s.calc("total_y", "ค่าบำรุงรักษา + ค่าบริการระบบ ต่อปี (ปีที่ 2 เป็นต้นไป)",
       f"=({R('F_SERVICE','mt_month')}+{R('F_SERVICE','sys_month')})*12", FMT_THB, "Formula", "บาท/ปี")
s.text("หมายเหตุเพิ่มเติม", "", bold=True)
for t in ["ค่าธรรมเนียมสแกน PromptPay 3% คิดเฉพาะยอดที่สแกนเท่านั้น ไม่เกี่ยวกับเงินสด — ตัดรอบ 7 วัน (จันทร์-อาทิตย์) โอนทุกวันพุธผ่าน K BIZ",
          "ช่องทาง TrueMoney Payment คิดค่าธรรมเนียม 3.5% ต่อรายการ (เป็นตัวเลือกเสริม เปิด/ปิดได้)",
          "การนับเครื่อง: เครื่องวางซ้อนบน-ล่าง 1 Stack ระบบนับเป็น 2 เครื่อง",
          "หลังหมดประกัน 3 ปี ควรตั้งงบซ่อมเพิ่มที่ช่อง 'ค่าซ่อมผันแปร' ใน A_INPUT เป็น 2-3%"]:
    s.text("•  " + t, "")
ws.sheet_view.showGridLines = False
# patch A_INPUT links
for k, srckey in (("maint", "mt_month"), ("sysfee", "sys_month")):
    _c = REF[("A_INPUT", k)].split("!")[1].replace("$", "")
    a.ws[_c] = f"={R('F_SERVICE', srckey)}"

# ============================================================ C_CALC
c = SheetBuilder(wb, "C_CALC", "548235")
ws = c.ws
for col, w in zip("ABCDE", [52, 20, 18, 16, 58]):
    ws.column_dimensions[col].width = w
ws.merge_cells("A1:E1")
put(ws, "A1", "C) CALCULATION ENGINE / เครื่องคำนวณ (ห้ามแก้ไข — เป็นสูตรทั้งหมด)",
    bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
ws.row_dimensions[1].height = 26
c.r = 3
MODE_A = f'LEFT({IN("mode")},1)="A"'
YES = lambda ref: f'LEFT({ref},1)="Y"'

c.band("1) CAPACITY / กำลังผลิต")
c.calc("w_units", "เครื่องซักที่ใช้จริง", f"={IN('w_act')}", FMT_INT, "จาก A_INPUT", "เครื่อง", color=C_LINK)
c.calc("d_units", "เครื่องอบที่ใช้จริง", f"={IN('d_act')}", FMT_INT, "จาก A_INPUT", "เครื่อง", color=C_LINK)
c.calc("t_units", "รวมเครื่องซัก+อบ", f"={IN('t_act')}", FMT_INT, "จาก A_INPUT", "เครื่อง", bold=True, color=C_LINK)
c.calc("cyc_w", "เวลา/รอบซักที่ใช้",
       f"=IF({IN('cyc_w')}>0,{IN('cyc_w')},{REF[('B_PACKAGES','cyc_w')]})", FMT_NUM, "นาที")
c.calc("cyc_d", "เวลา/รอบอบที่ใช้",
       f"=IF({IN('cyc_d')}>0,{IN('cyc_d')},{REF[('B_PACKAGES','cyc_d')]})", FMT_NUM, "นาที")
c.calc("cpd_w", "รอบ/วัน ต่อเครื่องซัก",
       f"=IFERROR({IN('open_h')}*60/({CA('cyc_w')}+{IN('chg')}),0)", FMT_NUM, "รอบ",
       note="สูตร: ชั่วโมงเปิด x 60 / (เวลารอบ + เวลาโหลด-ปลดผ้า)")
c.calc("cpd_d", "รอบ/วัน ต่อเครื่องอบ",
       f"=IFERROR({IN('open_h')}*60/({CA('cyc_d')}+{IN('chg')}),0)", FMT_NUM, "รอบ",
       note="เครื่องอบ 45 นาที = คอขวดของร้าน")
c.calc("cpd_c", "รอบ/วัน ต่อ 1 คู่ (โหมด A)",
       f"=IFERROR({IN('open_h')}*60/({IN('cyc_c')}+{IN('chg')}),0)", FMT_NUM, "รอบ")
c.calc("eff_days", "วันทำการถ่วงน้ำหนักการใช้งาน",
       f"={IN('wd_days')}*{IN('wd_util')}+{IN('we_days')}*{IN('we_util')}", FMT_NUM, "วัน-เทียบเท่า",
       note="สูตร: วันธรรมดา x %ใช้งาน + วันหยุด x %ใช้งาน")
c.calc("cap_w", "กำลังผลิตเครื่องซัก", f"={CA('w_units')}*{CA('cpd_w')}*{CA('eff_days')}", FMT_INT, "รอบ/เดือน")
c.calc("cap_d", "กำลังผลิตเครื่องอบ", f"={CA('d_units')}*{CA('cpd_d')}*{CA('eff_days')}", FMT_INT, "รอบ/เดือน")
c.calc("cap_c", "กำลังผลิตรวม (โหมด A)", f"={CA('w_units')}*{CA('cpd_c')}*{CA('eff_days')}", FMT_INT, "รอบ/เดือน",
       note="โหมด A: 1 รอบ = ใช้เครื่องซัก 1 + เครื่องอบ 1 จึงจำกัดด้วยจำนวนคู่")
c.calc("cap_tot", "กำลังผลิตที่ใช้ตามโหมด",
       f"=IF({MODE_A},{CA('cap_c')},{CA('cap_w')}+{CA('cap_d')})", FMT_INT, "รอบ/เดือน", bold=True)
c.calc("cap_max", "กำลังผลิตสูงสุด 100%",
       f"=IF({MODE_A},{CA('w_units')}*{CA('cpd_c')}*{IN('days')},"
       f"({CA('w_units')}*{CA('cpd_w')}+{CA('d_units')}*{CA('cpd_d')})*{IN('days')})",
       FMT_INT, "รอบ/เดือน", note="เพดานถ้าเครื่องเดินเต็ม 24 ชม. ทุกวัน")
c.skip()

c.band("2) PRICE / ราคาต่อรอบ")
c.calc("p_w", "ราคา/รอบซัก", f"={IN('price_w')}", FMT_THB2, "บาท")
c.calc("p_d", "ราคา/รอบอบ", f"={IN('price_d')}", FMT_THB2, "บาท")
c.calc("p_c", "ราคา/รอบ รวม (โหมด A)", f"={IN('price_c')}", FMT_THB2, "บาท")
c.skip()

c.band("3) DEMAND & BILLED CYCLES / ความต้องการและรอบที่ขายได้")
c.calc("o2o_w", "รอบซักจาก O2O", f"=IF({YES(IN('o2o_on'))},{IN('o2o_ord')}*{IN('days')},0)", FMT_INT, "รอบ/เดือน")
c.calc("o2o_d", "รอบอบจาก O2O",
       f"=IF({YES(IN('o2o_on'))},{IN('o2o_ord')}*{IN('days')}*MAX({IN('o2o_cyc')}-1,0),0)", FMT_INT, "รอบ/เดือน")
c.calc("dem_tot", "ความต้องการรวม",
       f"=IF({MODE_A},{IN('dem_w')}+{CA('o2o_w')},{IN('dem_w')}+{IN('dem_d')}+{CA('o2o_w')}+{CA('o2o_d')})",
       FMT_INT, "รอบ/เดือน", bold=True)
c.calc("cap_chk", "Capacity check (ความต้องการ / กำลังผลิต)",
       f"=IFERROR({CA('dem_tot')}/{CA('cap_tot')},0)", FMT_PCT, "%",
       note=">100% = เครื่องไม่พอ ควรอัปแพ็กเกจหรือขึ้นราคา | <50% = ลงทุนเกินตัว", bold=True)
c.calc("fill", "อัตราที่รองรับได้จริง / Fill rate",
       f"=IFERROR(MIN(1,{CA('cap_tot')}/{CA('dem_tot')}),1)", FMT_PCT, "%",
       note="ถ้าความต้องการเกินกำลังผลิต รายได้จะถูกจำกัดตามสัดส่วนนี้")
c.calc("bil_w", "รอบซักที่ขายได้ (walk-in)", f"={IN('dem_w')}*{CA('fill')}", FMT_INT, "รอบ/เดือน")
c.calc("bil_d", "รอบอบที่ขายได้ (walk-in)", f"=IF({MODE_A},0,{IN('dem_d')}*{CA('fill')})", FMT_INT, "รอบ/เดือน")
c.calc("bil_tot", "รอบที่ขายได้รวม (รวม O2O)",
       f"={CA('bil_w')}+{CA('bil_d')}+({CA('o2o_w')}+{CA('o2o_d')})*{CA('fill')}", FMT_INT, "รอบ/เดือน", bold=True)
c.calc("run_w", "จำนวนรอบที่เครื่องซักเดินจริง",
       f"={CA('bil_w')}+{CA('o2o_w')}*{CA('fill')}", FMT_INT, "รอบ/เดือน",
       note="ใช้คิดค่าไฟ-ค่าน้ำ และอายุเครื่องตามจำนวนรอบ")
c.calc("run_d", "จำนวนรอบที่เครื่องอบเดินจริง",
       f"=IF({MODE_A},{CA('bil_w')}+{CA('o2o_w')}*{CA('fill')},{CA('bil_d')}+{CA('o2o_d')}*{CA('fill')})",
       FMT_INT, "รอบ/เดือน", note="โหมด A: 1 รอบรวม = เดินเครื่องอบ 1 รอบด้วย")
c.calc("cyc_day", "รอบเฉลี่ยต่อวัน", f"=IFERROR({CA('bil_tot')}/{IN('days')},0)", FMT_NUM, "รอบ/วัน")
c.skip()

c.band("4) REVENUE / รายได้ต่อเดือน")
c.calc("rev_w", "รายได้เครื่องซัก",
       f"=IF({MODE_A},{CA('bil_w')}*{CA('p_c')},{CA('bil_w')}*{CA('p_w')})", FMT_THB, "บาท")
c.calc("rev_d", "รายได้เครื่องอบ", f"={CA('bil_d')}*{CA('p_d')}", FMT_THB, "บาท")
c.calc("rev_o2o", "รายได้ O2O",
       f"=IF({YES(IN('o2o_on'))},{IN('o2o_ord')}*{IN('days')}*{IN('o2o_tk')}*{CA('fill')},0)", FMT_THB, "บาท")
c.calc("rev_vd", "รายได้ตู้จำหน่ายสินค้า",
       f"={R('A_INPUT','cust_n')}*{IN('vend_att')}*{IN('vend_sp')}*{CA('fill')}", FMT_THB, "บาท")
c.calc("rev_ex", "รายได้อื่น", f"={IN('extra')}", FMT_THB, "บาท")
c.calc("rev_tot", "รายได้รวม / TOTAL REVENUE",
       f"=SUM({CA('rev_w')},{CA('rev_d')},{CA('rev_o2o')},{CA('rev_vd')},{CA('rev_ex')})",
       FMT_THB, "บาท/เดือน", bold=True, fill=KPIFIL)
c.calc("rev_mach", "รายได้ต่อเครื่องต่อวัน",
       f"=IFERROR(({CA('rev_w')}+{CA('rev_d')})/{CA('t_units')}/{IN('days')},0)", FMT_THB2, "บาท",
       note="ตัวชี้วัดเทียบสาขา: ต่ำกว่า 200 บาท/เครื่อง/วัน ถือว่าน่ากังวล")
c.calc("p_avg", "ราคาเฉลี่ยต่อรอบ",
       f"=IFERROR(({CA('rev_w')}+{CA('rev_d')}+{CA('rev_o2o')})/{CA('bil_tot')},0)", FMT_THB2, "บาท")
c.skip()

c.band("5) VARIABLE COST / ต้นทุนผันแปร")
c.calc("cost_util", "ค่าไฟ + ค่าน้ำ / Electricity & Water",
       f"=IF(LEFT({IN('util_mode')},1)=\"1\",({IN('elec_p')}+{IN('water_p')})*{CA('rev_tot')},"
       f"{CA('run_w')}*({REF[('B_PACKAGES','kwh_w')]}*{IN('elec_r')}+{REF[('B_PACKAGES','wat_w')]}/1000*{IN('water_r')})"
       f"+{CA('run_d')}*{REF[('B_PACKAGES','kwh_d')]}*{IN('elec_r')})+{IN('base_kwh')}",
       FMT_THB, "บาท", note="โมเดล 1 = % ของยอดขาย | โมเดล 2 = kWh/ลิตรจริง + ค่าไฟส่วนกลาง")
c.calc("cost_gas", "ค่าแก๊ส LPG เครื่องอบ / Gas",
       f"={CA('run_d')}*{REF[('B_PACKAGES','gas_d')]}*{IN('gas_r')}", FMT_THB, "บาท",
       note="เครื่องอบ Samsung เป็นระบบแก๊ส — คิดแยกจากค่าไฟทุกโมเดล")
c.calc("cost_cons", "วัสดุสิ้นเปลือง", f"={CA('bil_tot')}*{IN('cons_cyc')}", FMT_THB, "บาท")
c.calc("cost_o2o", "ค่าธรรมเนียม + ค่าส่ง O2O",
       f"={CA('rev_o2o')}*({IN('o2o_fee')}+{IN('o2o_rd')})", FMT_THB, "บาท")
c.calc("cost_scan", "ค่าธรรมเนียมสแกน PromptPay",
       f'=IF({IN("pay_mode")}="{PAY_SCAN}",{CA("rev_tot")}*{IN("scan_share")}*{IN("scan_fee")},0)',
       FMT_THB, "บาท", note="3% ของยอดที่สแกนเท่านั้น — แลกกับการยกเว้นค่าบริการระบบทั้งหมด")
c.calc("cost_rep", "ค่าซ่อมผันแปร", f"={CA('rev_tot')}*{IN('repair_p')}", FMT_THB, "บาท")
c.calc("cost_roy", "ค่าสิทธิ์/แฟรนไชส์", f"={CA('rev_tot')}*{IN('royal_p')}", FMT_THB, "บาท")
c.calc("cost_cogs", "ต้นทุนสินค้า vending + รายได้อื่น",
       f"={CA('rev_vd')}*(1-{IN('vend_gp')})+{CA('rev_ex')}*(1-{IN('extra_gp')})", FMT_THB, "บาท")
c.calc("var_tot", "รวมต้นทุนผันแปร / TOTAL VARIABLE COST",
       f"=SUM({CA('cost_util')},{CA('cost_gas')},{CA('cost_cons')},{CA('cost_o2o')},{CA('cost_scan')},"
       f"{CA('cost_rep')},{CA('cost_roy')},{CA('cost_cogs')})", FMT_THB, "บาท/เดือน", bold=True, fill=KPIFIL)
c.calc("var_pct", "ต้นทุนผันแปร % ของยอดขาย", f"=IFERROR({CA('var_tot')}/{CA('rev_tot')},0)", FMT_PCT, "%")
c.calc("cm", "กำไรส่วนเกิน / Contribution margin", f"={CA('rev_tot')}-{CA('var_tot')}", FMT_THB, "บาท/เดือน")
c.calc("cm_pct", "อัตรากำไรส่วนเกิน / CM %", f"=IFERROR({CA('cm')}/{CA('rev_tot')},0)", FMT_PCT, "%", bold=True)
c.skip()

c.band("6) CAPEX & FUNDING / เงินลงทุนและแหล่งเงิน")
c.calc("vat_f", "ตัวคูณ VAT", f"=IF({YES(IN('vat_reg'))},1/(1+{IN('vat')}),1)", '0.0000', "Formula",
       note="จด VAT = ขอคืนภาษีซื้อได้ → ถอด VAT ออกจาก CAPEX | ไม่จด = VAT เป็นต้นทุน")
c.calc("capex_pkg", "ราคาแพ็กเกจ (หลังปรับ VAT)",
       f"=IF({IN('capex_pkg_ov')}>0,{IN('capex_pkg_ov')},{IN('capex_pkg_r')})*{CA('vat_f')}", FMT_THB, "บาท")
c.calc("capex_prep", "CAPEX นอกแพ็กเกจ", f"={IN('capex_prep_r')}*{CA('vat_f')}", FMT_THB, "บาท")
c.calc("capex_add", "อุปกรณ์เสริม", f"={IN('capex_addon_r')}*{CA('vat_f')}", FMT_THB, "บาท")
c.calc("capex_m", "ส่วนที่เป็นตัวเครื่อง (ตัดค่าเสื่อม 8 ปี)",
       f"={CA('capex_pkg')}*{IN('mach_share')}+{CA('capex_add')}", FMT_THB, "บาท")
c.calc("capex_f", "ส่วนงานติดตั้ง/ตกแต่ง/งานเตรียม (ตัดค่าเสื่อม 5 ปี)",
       f"={CA('capex_pkg')}*(1-{IN('mach_share')})+{CA('capex_prep')}+{IN('capex_other')}", FMT_THB, "บาท")
c.calc("deposit", "เงินประกันค่าเช่า", f"={IN('rent')}*{IN('dep_month')}", FMT_THB, "บาท")
c.calc("capex_tot", "เงินลงทุนรวม / TOTAL CAPEX",
       f"={CA('capex_m')}+{CA('capex_f')}+{CA('deposit')}+{IN('wcap')}", FMT_THB, "บาท",
       bold=True, fill=OKFILL)
c.calc("loan", "เงินกู้ / Loan amount", f"={CA('capex_tot')}*{IN('loan_p')}", FMT_THB, "บาท")
c.calc("equity", "เงินลงทุนของเจ้าของ / Equity", f"={CA('capex_tot')}-{CA('loan')}", FMT_THB, "บาท", bold=True)
c.calc("pmt", "ค่างวดเงินกู้/เดือน",
       f"=IFERROR(IF({CA('loan')}>0,PMT({IN('loan_r')}/12,{IN('loan_y')}*12,-{CA('loan')}),0),0)",
       FMT_THB, "บาท/เดือน")
c.skip()

c.band("7) MONTHLY P&L / งบกำไรขาดทุนต่อเดือน (สภาวะปกติ ปีที่ 2 เป็นต้นไป)")
c.calc("pl_rev", "รายได้ / Revenue", f"={CA('rev_tot')}", FMT_THB, "บาท", bold=True)
c.calc("pl_var", "หัก ต้นทุนผันแปร", f"=-{CA('var_tot')}", FMT_THB, "บาท")
c.calc("pl_fix", "หัก ค่าใช้จ่ายคงที่", f"=-{IN('fix_tot')}", FMT_THB, "บาท")
c.calc("ebitda", "EBITDA", f"={CA('pl_rev')}+{CA('pl_var')}+{CA('pl_fix')}", FMT_THB, "บาท",
       bold=True, fill=KPIFIL)
c.calc("dep", "ค่าเสื่อมราคา",
       f"=IFERROR({CA('capex_m')}/{IN('dep_y_m')}/12,0)+IFERROR({CA('capex_f')}/{IN('dep_y_f')}/12,0)",
       FMT_THB, "บาท")
c.calc("ebit", "EBIT", f"={CA('ebitda')}-{CA('dep')}", FMT_THB, "บาท")
c.calc("int", "ดอกเบี้ยจ่าย (เฉลี่ยปีแรก)", f"={CA('loan')}*{IN('loan_r')}/12", FMT_THB, "บาท")
c.calc("ebt", "กำไรก่อนภาษี / EBT", f"={CA('ebit')}-{CA('int')}", FMT_THB, "บาท")
c.calc("tax", "ภาษีเงินได้", f"=MAX(0,{CA('ebt')})*{IN('tax_r')}", FMT_THB, "บาท")
c.calc("np", "กำไรสุทธิ / NET PROFIT", f"={CA('ebt')}-{CA('tax')}", FMT_THB, "บาท/เดือน",
       bold=True, fill=OKFILL)
c.calc("np_m", "อัตรากำไรสุทธิ", f"=IFERROR({CA('np')}/{CA('rev_tot')},0)", FMT_PCT, "%", bold=True)
c.calc("ebitda_m", "อัตรา EBITDA", f"=IFERROR({CA('ebitda')}/{CA('rev_tot')},0)", FMT_PCT, "%")
c.calc("cf", "กระแสเงินสดสุทธิ/เดือน", f"={CA('np')}+{CA('dep')}", FMT_THB, "บาท/เดือน", bold=True,
       note="กำไรสุทธิ + ค่าเสื่อม (ค่าเสื่อมไม่ใช่เงินสดออกจริง)")
c.skip()

c.band("8) KEY RESULTS / ตัวชี้วัดหลัก")
c.calc("payback", "ระยะคืนทุน / Payback period",
       f"=IFERROR(IF({CA('cf')}<=0,999,{CA('capex_tot')}/{CA('cf')}),999)", FMT_NUM, "เดือน", bold=True)
c.calc("payback_y", "ระยะคืนทุน (ปี)", f"={CA('payback')}/12", FMT_NUM, "ปี")
c.calc("roi", "ROI ต่อปี", f"=IFERROR({CA('np')}*12/{CA('capex_tot')},0)", FMT_PCT, "%", bold=True)
c.calc("roe", "ROE ต่อปี", f"=IFERROR({CA('np')}*12/{CA('equity')},0)", FMT_PCT, "%")
c.calc("bep_rev", "จุดคุ้มทุน (ยอดขาย)",
       f"=IFERROR(({IN('fix_tot')}+{CA('int')})/{CA('cm_pct')},0)", FMT_THB, "บาท/เดือน",
       note="ฐานเงินสด (ไม่รวมค่าเสื่อม)")
c.calc("bep_cyc", "จุดคุ้มทุน (รอบ/วัน)",
       f"=IFERROR({CA('bep_rev')}/{CA('p_avg')}/{IN('days')},0)", FMT_NUM, "รอบ/วัน", bold=True)
c.calc("bep_pct", "จุดคุ้มทุนคิดเป็น % ของยอดขายปัจจุบัน",
       f"=IFERROR({CA('bep_rev')}/{CA('rev_tot')},0)", FMT_PCT, "%",
       note="ยิ่งต่ำยิ่งปลอดภัย — เกิน 80% ถือว่าเสี่ยงสูง")
c.calc("dscr", "DSCR (ความสามารถชำระหนี้)",
       f"=IFERROR(IF({CA('pmt')}=0,\"N/A\",{CA('ebitda')}/{CA('pmt')}),\"N/A\")", FMT_X, "เท่า",
       note="ธนาคารต้องการ > 1.25 เท่า")
c.calc("life_y", "อายุเครื่องซักตามจำนวนรอบ",
       f"=IFERROR({REF[('B_PACKAGES','life_w')]}/({CA('run_w')}*12/{CA('w_units')}),99)", FMT_NUM, "ปี",
       note="สเปก 30,000 รอบ หารด้วยรอบที่เดินจริงต่อเครื่องต่อปี")
c.skip()

c.band("9) HEALTH CHECK / ตรวจสุขภาพโมเดล")
def check(key, label, formula, note):
    c.calc(key, label, formula, None, "ตรวจสอบ", note)
check("chk_cap", "กำลังผลิต / Capacity",
      f'=IF({CA("cap_chk")}>1,"เครื่องไม่พอ — ความต้องการเกินกำลังผลิต ควรอัปแพ็กเกจหรือขึ้นราคา",'
      f'IF({CA("cap_chk")}<0.4,"ลงทุนเกินความต้องการ — พิจารณาลดขนาดแพ็กเกจ","OK"))',
      "เป้าหมาย 60-95%")
check("chk_rent", "สัดส่วนค่าเช่า / Rent ratio",
      f'=IF(IFERROR({IN("rent")}/{CA("rev_tot")},1)>0.15,"ค่าเช่าสูงเกิน 15% ของยอดขาย — ต่อรองใหม่","OK")',
      "ค่าเช่าไม่ควรเกิน 12-15% ของยอดขาย")
check("chk_util", "สัดส่วนค่าน้ำ-ไฟ-แก๊ส / Utility ratio",
      f'=IF(IFERROR(({CA("cost_util")}+{CA("cost_gas")})/{CA("rev_tot")},1)>0.25,'
      f'"ค่าน้ำ-ไฟ-แก๊สสูงเกิน 25% — ตรวจสอบอัตราค่าไฟ/ค่าแก๊ส และ %การใช้งาน","OK")',
      "ปกติ 16-24% ของยอดขาย (รวมค่าแก๊สเครื่องอบ)")
check("chk_pb", "ระยะคืนทุน / Payback",
      f'=IF({CA("payback")}>36,"คืนทุนช้ากว่า 36 เดือน — ทบทวนทำเล/ค่าเช่า/ราคา",'
      f'IF({CA("payback")}>24,"คืนทุน 24-36 เดือน — รับได้แต่ควรปรับปรุง","ดี: คืนทุนภายใน 24 เดือน"))',
      "เกณฑ์ธุรกิจร้านสะดวกซัก: 18-30 เดือน")
check("chk_bep", "ความปลอดภัยจุดคุ้มทุน / BEP safety",
      f'=IF({CA("bep_pct")}>0.8,"เสี่ยงสูง — ยอดขายต้องถึง 80%+ ของประมาณการจึงจะคุ้มทุน","OK")',
      "Margin of safety = 1 - ค่านี้")
check("chk_life", "อายุเครื่องตามจำนวนรอบ",
      f'=IF({CA("life_y")}<{IN("dep_y_m")},"เครื่องจะครบ 30,000 รอบก่อนตัดค่าเสื่อมหมด — ตั้งงบเปลี่ยนเครื่องเพิ่ม","OK")',
      "สเปก Samsung: ทนทานรับ 30,000 รอบ")
check("chk_pay", "ทางเลือกระบบรับชำระเงิน",
      f'=IF({IN("pay_mode")}="{PAY_SCAN}",IF({CA("rev_tot")}*{IN("scan_share")}*{IN("scan_fee")}>{R("F_SERVICE","b_month")},'
      f'"ยอดสแกนสูง — ค่าธรรมเนียม 3% แพงกว่าค่าบริการระบบแบบรายเดือน ลองเทียบ F_SERVICE",'
      f'"OK — เปิดสแกนคุ้มกว่า"),"OK — ไม่เปิดสแกน")',
      "เปรียบเทียบรายละเอียดที่ชีต F_SERVICE")
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A3"

# link F_SERVICE!scan_act back to C_CALC (ไม่เกิด circular เพราะ rev_tot ไม่ได้ขึ้นกับค่าบริการ)
_sa = REF[("F_SERVICE", "scan_act")].split("!")[1].replace("$", "")
s.ws[_sa] = f"={CA('rev_tot')}*{IN('scan_share')}"
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
    ("ค่าไฟ + ค่าน้ำ", f"-{CA('cost_util')}", False),
    ("ค่าแก๊ส LPG (เครื่องอบ)", f"-{CA('cost_gas')}", False),
    ("วัสดุสิ้นเปลือง", f"-{CA('cost_cons')}", False),
    ("ค่าธรรมเนียมสแกน 3% + O2O", f"-{CA('cost_scan')}-{CA('cost_o2o')}", False),
    ("ค่าซ่อม + ค่าสิทธิ์ + ต้นทุนสินค้า", f"-{CA('cost_rep')}-{CA('cost_roy')}-{CA('cost_cogs')}", False),
    ("รวมต้นทุนผันแปร", f"-{CA('var_tot')}", True),
    ("กำไรส่วนเกิน / Contribution margin", CA('cm'), True),
    ("ค่าเช่า", f"-{IN('rent')}", False),
    ("ค่าจ้างพนักงาน", f"-{IN('staff')}", False),
    ("ค่าบำรุงรักษา + ค่าบริการระบบ", f"-{R('A_INPUT','maint')}-{R('A_INPUT','sysfee')}", False),
    ("ค่าใช้จ่ายคงที่อื่นๆ",
     f"-({IN('fix_tot')}-{IN('rent')}-{IN('staff')}-{R('A_INPUT','maint')}-{R('A_INPUT','sysfee')})", False),
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
for key, lab in [("chk_cap", "กำลังผลิต"), ("chk_rent", "ค่าเช่า"), ("chk_util", "ค่าน้ำ-ไฟ-แก๊ส"),
                 ("chk_pb", "ระยะคืนทุน"), ("chk_bep", "จุดคุ้มทุน"), ("chk_life", "อายุเครื่อง"),
                 ("chk_pay", "ระบบรับชำระเงิน")]:
    put(d, f"A{rr}", lab, size=10, bold=True)
    d.merge_cells(f"B{rr}:H{rr}")
    put(d, f"B{rr}", f"={CA(key)}", size=10, align="left")
    for ch in "CDEFGH":
        d[f"{ch}{rr}"].border = BOX
    rr += 1
from openpyxl.formatting.rule import FormulaRule
_hc0 = REF[("D_DASHBOARD", "pl_end")] + 2
_hc_rng = f"A{_hc0}:H{_hc0+6}"
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
    put(e, f"E{r}", f"=IF(A{r}>{H12},0,-(({IN('fix_tot')}-{IN('rent')})-IF(B{r}=1,{R('F_SERVICE','free_y1')},0)"
                    f"+{IN('rent')}*(1+{IN('rent_esc')})^(B{r}-1)))",
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
             "ปีที่ 1 ระบบหักค่าบำรุงรักษาและค่าบริการระบบออกให้ (รวมในแพ็กเกจแล้ว) | เดือนสุดท้ายของ horizon รวมมูลค่าคงเหลือของเครื่อง + เงินประกันค่าเช่า + เงินทุนหมุนเวียนคืน",
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
       ("ค่าไฟ+น้ำ (บาท/เดือน)", CA('cost_util'), FMT_THB, "ค่าแก๊ส LPG (บาท/เดือน)", CA('cost_gas'), FMT_THB),
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
rd.column_dimensions["A"].width = 32
rd.column_dimensions["B"].width = 112
rd.merge_cells("A1:B1")
put(rd, "A1", "การคำนวณ ROI ร้านสะดวกซัก  Samsung Commercial", bold=True, size=16,
    color=C_WHT, fill=NAVY, align="center", border=False)
rd.row_dimensions[1].height = 34
rd.merge_cells("A2:B2")
put(rd, "A2", "LAUNDROMAT ROI MODEL — Version 2.0  |  อัปเดต 19-09-2026  |  อิงราคาและสเปกจริงจากเอกสาร Samsung Commercial Franchise",
    size=10, color="595959", align="center", border=False, fill=GREY)

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
    ("ขั้นที่ 1", "ชีต H_LOCATION_SCORE — ให้คะแนนทำเลก่อน ถ้าได้ NO-GO ไม่ต้องคำนวณต่อ"),
    ("ขั้นที่ 2", "ชีต B_PACKAGES — ตรวจราคาแพ็กเกจ (S/M/L) และกรอก 'CAPEX นอกแพ็กเกจ' ที่ลูกค้าต้องเตรียมเอง "
                  "พร้อมเลือกอุปกรณ์เสริมที่จะซื้อเพิ่ม"),
    ("ขั้นที่ 3", "ชีต A_INPUT — กรอกช่องสีเหลือง: ประชากร ค่าเช่า ราคา/รอบ และเลือกระบบรับชำระเงิน"),
    ("ขั้นที่ 4", "ชีต F_SERVICE — เทียบ 2 ทางเลือกค่าบริการระบบ (เปิดสแกน 3% vs จ่ายรายเดือน)"),
    ("ขั้นที่ 5", "ดูผลที่ D_DASHBOARD / E_CASHFLOW → ทดสอบความเสี่ยงที่ G_SCENARIO → พิมพ์ I_PRINT_VIEW เสนอลูกค้า"),
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
    ("0_README", "หน้านี้ — คู่มือ ที่มาตัวเลข ข้อควรระวัง และสิ่งที่ต้องยืนยันกับ Samsung"),
    ("A_INPUT", "ช่องกรอก 8 หมวด: โครงการ / ตลาด / แพ็กเกจ / CAPEX / การดำเนินงาน / รายได้เสริม / OPEX / การเงิน-ภาษี"),
    ("B_PACKAGES", "สเปกเครื่อง + แพ็กเกจ S/M/L/CUSTOM + CAPEX นอกแพ็กเกจ + อุปกรณ์เสริม + checklist สิ่งที่รวมในแพ็กเกจ"),
    ("C_CALC", "เครื่องคำนวณ 9 ส่วน (capacity → รอบที่ขายได้ → รายได้ → ต้นทุน → CAPEX → P&L → KPI → health check)"),
    ("D_DASHBOARD", "สรุป KPI + งบกำไรขาดทุนต่อเดือน + คำเตือนอัตโนมัติ 7 ข้อ"),
    ("E_CASHFLOW", "กระแสเงินสด 60 เดือน + NPV + IRR + คืนทุนแบบเงินสด/คิดลด + ตารางผ่อนเงินกู้"),
    ("F_SERVICE", "ประกันที่รวมในแพ็กเกจ + ค่าบำรุงรักษาหลังปีแรก + เปรียบเทียบค่าบริการระบบ 2 ทางเลือก"),
    ("G_SCENARIO", "Worst / Base / Best + ตารางความอ่อนไหว 2 ทาง (ราคา x % ลูกค้า)"),
    ("H_LOCATION_SCORE", "ให้คะแนนทำเล 12 เกณฑ์ พร้อมข้อสรุป GO / HOLD / NO-GO"),
    ("I_PRINT_VIEW", "สรุปการลงทุน A4 1 หน้า สำหรับพิมพ์เสนอลูกค้า"),
    ("Z_LISTS", "รายการตัวเลือก dropdown (ซ่อนไว้ — คลิกขวาที่แท็บ > Unhide)"),
]:
    r = rrow(r, k, v)
r += 1

r = rsec(r, "4) ตัวเลขที่ยืนยันแล้วจากเอกสาร Samsung / CONFIRMED DATA")
for k, v in [
    ("ราคาแพ็กเกจ (รวม VAT)", "S = 3 คู่ 699,000 บาท | M = 4 คู่ 799,000 บาท | L = 5 คู่ 899,000 บาท (Best seller)"),
    ("สเปกเครื่อง", "เครื่องซัก 18 kg รอบละ 36 นาที | เครื่องอบแก๊ส 14 kg รอบละ 45 นาที | ทนทานรับ 30,000 รอบ"),
    ("การรับประกัน", "เครื่องซัก-อบ 3 ปี | กล่องหยอดเหรียญ 1 ปี | ระบบ I'M CONTROL + ค่าบริการ ฟรี 1 ปี"),
    ("งานบำรุงรักษาในแพ็กเกจ", "S-Built-in = 1 ครั้ง | M&L-Built-in = 2 ครั้ง"),
    ("ค่าบริการระบบ (หลังปีแรก)", "100 บาท/เครื่อง/เดือน (1 Stack = 2 เครื่อง) + 700 บาท/สาขา/เดือน (สาขา 7 เครื่องขึ้นไป) "
                                  "+ 8,000 บาท/ปี — ยกเว้นทั้งหมดถ้าเปิดสแกน PromptPay 3%"),
    ("ค่าธรรมเนียมชำระเงิน", "PromptPay 3% ของยอดสแกน (ตัดรอบ 7 วัน โอนทุกวันพุธผ่าน K BIZ) | TrueMoney 3.5% ต่อรายการ"),
    ("อุปกรณ์เสริม", "กล่องไซด์บาร์ 13,000-17,500 | iAm Control 4,990 | กล่องสแกนออนไลน์ 5,690 | "
                     "เครื่องแลกเหรียญ 16,990-49,000 | ตู้จำหน่ายสินค้า 29,990-36,990"),
    ("สิ่งที่รวมในแพ็กเกจ", "งานไฟฟ้า 1 เฟส (ตู้ควบคุม + เบรกเกอร์ 63A) | ถังน้ำ 1,000 ล. + ปั๊ม 150 W | "
                            "ระบบแก๊สสลับอัตโนมัติ 2 ข้าง | ท่อลมร้อน 4\" | กรุผนัง ≤25 ตร.ม. | ป้ายกล่องไฟ | "
                            "เครื่องแลกเหรียญ+ตู้ขายของมินิ | CCTV | โต๊ะเก้าอี้ | พัดลม | อ่างล้างมือ"),
]:
    r = rrow(r, k, v)
r += 1

r = rsec(r, "5) ตัวเลขที่ยังเป็นค่าประมาณ — ต้องยืนยันก่อนใช้จริง / TO CONFIRM")
for k, v in [
    ("1. ค่าไฟ/ค่าน้ำ/ค่าแก๊ส ต่อรอบ", "ใส่ไว้ที่ B_PACKAGES: ซัก 0.55 kWh + 110 ลิตร | อบ 0.35 kWh + 0.55 กก.LPG — "
                                        "ต้องยืนยันจากสเปกผู้ผลิตหรือวัดจากบิลจริงของสาขาที่เปิดแล้ว"),
    ("2. ราคาขายต่อรอบ", "ตั้งไว้ซัก 50 / อบ 50 บาท ตามราคาตลาด — ต้องสำรวจราคาคู่แข่งในทำเลจริง"),
    ("3. ค่าบำรุงรักษาหลังปีแรก", "ตั้งไว้ 3,500 บาท/ครั้ง x 2 ครั้ง/ปี ที่ชีต F_SERVICE — ยังไม่ใช่ราคาทางการ"),
    ("4. CAPEX นอกแพ็กเกจ", "งานไฟเมน มิเตอร์ ประปา แท่นเครื่อง ฝ้า-พื้น (รวม ~178,000-238,000 บาท) — "
                            "ต้องขอใบเสนอราคาจากผู้รับเหมาจริง เพราะแตกต่างกันมากตามสภาพอาคาร"),
    ("5. % การใช้งาน (utilization)", "ตั้งไว้ 35% วันธรรมดา / 65% วันหยุด ตามฟอร์มอ้างอิงเดิม — "
                                     "ควร calibrate จากข้อมูล I'M CONTROL ของสาขาที่เปิดแล้ว"),
    ("6. ค่าแก๊ส LPG 25 บาท/กก.", "ราคาอ้างอิงตลาด — ตรวจสอบราคารับถังในพื้นที่จริง"),
]:
    r = rrow(r, k, v, color="C00000")
r += 1

r = rsec(r, "6) ข้อควรระวังเชิงตัวเลข / MODEL NOTES")
for k, v in [
    ("เครื่องอบคือคอขวด", "เครื่องอบใช้ 45 นาที เทียบกับเครื่องซัก 36 นาที ในแพ็กเกจที่ซัก:อบ = 1:1 "
                          "กำลังผลิตฝั่งอบจะเต็มก่อนเสมอ — ถ้า Capacity check สูง ให้พิจารณาเพิ่มเครื่องอบก่อน"),
    ("ค่าแก๊สแยกจากค่าไฟ", "เครื่องอบเป็นระบบแก๊ส LPG ระบบจึงคิดค่าแก๊สแยกทุกโมเดล "
                           "(โมเดลค่าน้ำ-ไฟแบบ % ของยอดขายครอบคลุมเฉพาะไฟกับน้ำ)"),
    ("ราคาแพ็กเกจรวม VAT แล้ว", "ถ้าเลือก 'จดทะเบียน VAT = Yes' ระบบจะถอด VAT ออกจาก CAPEX (หาร 1.07) เพราะขอคืนภาษีซื้อได้"),
    ("การแยกค่าเสื่อม", "ราคาแพ็กเกจเป็นก้อนเดียว จึงใช้ช่อง '% ของราคาแพ็กเกจที่เป็นตัวเครื่อง' (ตั้งไว้ 70%) "
                        "แยกตัดค่าเสื่อม: ส่วนเครื่อง 8 ปี / ส่วนงานตกแต่ง-ติดตั้ง 5 ปี"),
    ("ปีแรกไม่มีค่าบริการระบบ", "P&L รายเดือนแสดงสภาวะปกติ (ปีที่ 2 เป็นต้นไป) ส่วน E_CASHFLOW หักส่วนลดปีแรกให้อัตโนมัติ"),
    ("การจำกัดด้วยกำลังผลิต", "ถ้าความต้องการเกินกำลังผลิต โมเดลจะจำกัดรายได้ด้วย Fill rate (ดู C_CALC ส่วนที่ 3)"),
    ("ระยะคืนทุน 2 ตัวเลข", "D_DASHBOARD ใช้สูตรง่าย (เงินลงทุน / กระแสเงินสดต่อเดือน) ส่วน E_CASHFLOW คิดจากกระแสเงินสดสะสมจริง "
                            "รวมค่าเช่าที่ปรับขึ้นและส่วนลดปีแรก — ต่างกันเล็กน้อยเป็นเรื่องปกติ"),
    ("ถ้าไม่เปิดสแกน", "ลูกค้าจะจ่ายได้เฉพาะเหรียญ/แบงก์ ซึ่งในทางปฏิบัติมักทำให้ยอดขายลดลง "
                       "โมเดลยังไม่ได้หักผลกระทบส่วนนี้ ให้ปรับ % ลูกค้าลงเองถ้าเลือกทางนี้"),
]:
    r = rrow(r, k, v)

# ---------------------------------------------------------------- order & save
ORDER = ["0_README", "A_INPUT", "B_PACKAGES", "C_CALC", "D_DASHBOARD", "E_CASHFLOW",
         "F_SERVICE", "G_SCENARIO", "H_LOCATION_SCORE", "I_PRINT_VIEW", "Z_LISTS"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
for name in ORDER:
    sh = wb[name]
    if name != "I_PRINT_VIEW":
        sh.page_setup.orientation = "landscape" if name in ("E_CASHFLOW", "B_PACKAGES", "G_SCENARIO") else "portrait"
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

OUT = "การคำนวณROI_ร้านสะดวกซัก_Samsung_Commercial_v2.0.xlsx"
wb.save(OUT)
print("saved:", OUT)
