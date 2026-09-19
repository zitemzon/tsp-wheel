# -*- coding: utf-8 -*-
"""
PM Pricing Tool v2.0 — เครื่องมือตั้งราคางานบำรุงรักษาเชิงป้องกัน Samsung Commercial
ต่อยอดจากไฟล์ 4_PM_Pricing_Tool.xlsx ของผู้ใช้ พร้อมแก้สูตรต้นทุนที่คลาดเคลื่อน
Run: python3 build_pm_pricing_tool.py
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.properties import PageSetupProperties

FONT   = "Tahoma"
NAVY   = "1F3864"
BAND   = "111111"
DARK   = "2F3B52"
GREY   = "F2F2F2"
YELLOW = "FFFF00"
KPIFIL = "D9E1F2"
OKFILL = "E2EFDA"
WARN   = "FCE4D6"
C_IN, C_CALC, C_LINK, C_WHT = "0000FF", "000000", "008000", "FFFFFF"
FMT_THB  = '#,##0;(#,##0);"-"'
FMT_THB2 = '#,##0.00;(#,##0.00);"-"'
FMT_PCT  = '0.0%'
FMT_INT  = '#,##0;(#,##0);"-"'
FMT_NUM  = '#,##0.0;(#,##0.0);"-"'
FMT_DATE = 'yyyy-mm-dd'
thin = Side(style="thin", color="BFBFBF")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

def put(ws, cell, value, *, fmt=None, bold=False, color=C_CALC, fill=None, size=10,
        align=None, wrap=False, border=True, italic=False):
    c = ws[cell]
    c.value = value
    c.font = Font(name=FONT, size=size, bold=bold, color=color, italic=italic)
    if fmt:  c.number_format = fmt
    if fill: c.fill = PatternFill("solid", fgColor=fill)
    if border: c.border = BOX
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return c

def band(ws, row, text, last_col, fill=BAND, size=11):
    ws.merge_cells(f"A{row}:{last_col}{row}")
    put(ws, f"A{row}", text, bold=True, color=C_WHT, fill=fill, size=size, align="left", border=False)
    for i in range(1, ws.max_column + 2):
        ws[f"{get_column_letter(i)}{row}"].fill = PatternFill("solid", fgColor=fill)
    ws.row_dimensions[row].height = 20

def inp(ws, cell, value, fmt=None, align="right"):
    return put(ws, cell, value, fmt=fmt, color=C_IN, fill=YELLOW, bold=True, align=align)

wb = Workbook(); wb.remove(wb.active)
TODAY = datetime.date(2026, 9, 19)

# ---------------------------------------------------------------- Z_LISTS
zl = wb.create_sheet("Z_LISTS")
LISTS = {
    "A": ("CostBasis", ["ค่าเฉลี่ย (แนะนำ)", "งานกระจุกจุดเดียว", "งาน Route หลายจุด"]),
    "B": ("Pack", ["Care", "Care+", "Care Pro"]),
    "C": ("YesNo", ["ใช่", "ไม่ใช่"]),
}
REF = {}
for col, (name, vals) in LISTS.items():
    put(zl, f"{col}1", name, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
    for i, v in enumerate(vals, start=2):
        put(zl, f"{col}{i}", v, size=9)
    zl.column_dimensions[col].width = 24
    REF[name] = f"Z_LISTS!${col}$2:${col}${len(vals)+1}"
zl.sheet_state = "hidden"
def DV(name, sheet):
    dv = DataValidation(type="list", formula1="=" + REF[name], allow_blank=True, showDropDown=False)
    sheet.add_data_validation(dv)
    return dv

# ============================================================ 1_สมมติฐานต้นทุน
s1 = wb.create_sheet("1_สมมติฐานต้นทุน")
s1.sheet_view.showGridLines = False
s1.sheet_properties.tabColor = NAVY
for c, w in zip("ABCD", [52, 16, 18, 70]):
    s1.column_dimensions[c].width = w
s1.merge_cells("A1:D1")
put(s1, "A1", "สมมติฐานต้นทุนงาน PM — แก้ตัวเลขในช่องสีเหลืองเท่านั้น", bold=True, size=14,
    color=C_WHT, fill=NAVY, align="left", border=False)
s1.row_dimensions[1].height = 26
put(s1, "A2", "ช่องสีเหลือง = ช่องกรอกข้อมูล (ตัวอักษรสีน้ำเงิน) | ช่องอื่นเป็นสูตรคำนวณ ห้ามแก้",
    italic=True, size=9, color="595959", border=False)
for col, t in zip("ABCD", ["รายการ", "ค่า", "หน่วย", "หมายเหตุ / ที่มาของตัวเลข"]):
    put(s1, f"{col}4", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)

A1ROWS = {}
def a1(key, label, value, unit, note, fmt=FMT_THB, row=None, new=False, dv=None):
    r = row
    put(s1, f"A{r}", ("★ " if new else "") + label, size=10, bold=new)
    c = inp(s1, f"B{r}", value, fmt)
    if dv: dv.add(c)
    put(s1, f"C{r}", unit, size=9, color="595959", align="center")
    put(s1, f"D{r}", note, size=9, color=("C00000" if new else "595959"), wrap=True)
    A1ROWS[key] = r
    return r

a1("lab1",  "ค่าแรงช่างหลัก ต่อวัน", 900, "บาท/วัน", "ตัวเลขจากผู้ใช้ — ปรับตามค่าจ้างจริงในพื้นที่", row=5)
a1("lab2",  "ค่าแรงผู้ช่วยช่าง ต่อวัน", 600, "บาท/วัน", "ตัวเลขจากผู้ใช้", row=6)
a1("car",   "ค่ารถ + น้ำมัน + ทางด่วน ต่อวัน", 500, "บาท/วัน", "ตัวเลขจากผู้ใช้ — งานในเขตเมือง", row=7)
a1("mat",   "วัสดุสิ้นเปลือง ต่อเครื่อง", 100, "บาท/เครื่อง", "น้ำยาล้าง ผ้า จารบี ซิลิโคน เทป foil ถุงมือ", row=8)
a1("cap1",  "จำนวนเครื่องที่ทำได้ ต่อวัน (งานกระจุกจุดเดียว)", 10, "เครื่อง/วัน",
   "เครื่องซัก 45-60 นาที / เครื่องอบ 60-75 นาที", FMT_INT, row=9)
a1("cap2",  "จำนวนเครื่องที่ทำได้ ต่อวัน (งาน Route หลายจุด)", 8, "เครื่อง/วัน",
   "เสียเวลาเดินทางระหว่างจุด", FMT_INT, row=10)
a1("kmsell","ค่าเดินทางส่วนเกิน (ราคาที่เรียกเก็บ)", 12, "บาท/กม.",
   "คิดเฉพาะขาไป | เบนช์มาร์กตลาด SPINZ = 10 บาท/กม.", FMT_THB2, row=11)
a1("hotel", "ค่าที่พักกรณีค้างคืน", 1200, "บาท/คืน/ทีม", "ตัวเลขจากผู้ใช้", row=12)
a1("noshow","ค่าเดินทางเสียเที่ยว (ลูกค้าไม่พร้อมให้เข้างาน)", 800, "บาท/ครั้ง",
   "ตามเงื่อนไขในใบเสนอราคาข้อ 5", row=13)
a1("gmfloor", "★ GM ขั้นต่ำที่ยอมรับได้ (Floor)", 0.35, "%",
   "ต่ำกว่านี้ไม่ควรรับงาน หรือต้องขออนุมัติพิเศษ — อ้างอิงราคาตลาด CODE CLEAN 600 บาท/เครื่อง/ครั้ง "
   "ซึ่งให้ GM ราว 38% บนโครงสร้างต้นทุนชุดนี้ (ดูแถวล่างสุดของชีต)", FMT_PCT, row=14, new=True)
a1("gmtgt", "★ GM เป้าหมายที่ต้องการ (Target)", 0.45, "%",
   "ระดับที่ราคาป้ายควรอยู่ เพื่อเหลือรองรับ overhead และกำไร — เดิมตั้งไว้ 50% ซึ่งสูงกว่าที่ตลาดจ่ายจริง",
   FMT_PCT, row=15, new=True)

band(s1, 17, "★ สมมติฐานที่เพิ่มในเวอร์ชัน 2 — จำเป็นต่อการคิดต้นทุนให้ครบ", "D", fill=DARK)
a1("dryf", "ตัวคูณต้นทุนเครื่องอบ (เทียบเครื่องซัก)", 1.27, "เท่า",
   "เดิมค่านี้ถูกฝังไว้ในสูตร F12 เท่านั้น ทำให้ต้นทุนแพ็กเกจคิดเครื่องอบเท่าเครื่องซัก — "
   "งานอบหนักกว่าเพราะต้องถอดล้างท่อลมร้อนและตรวจระบบแก๊ส", '0.00', row=18, new=True)
a1("corrc","ต้นทุนงานซ่อมด่วน (Corrective) ต่อครั้ง", 800, "บาท/ครั้ง",
   "ประมาณ 1/3 วันทำงาน + วัสดุ — ใช้คิดต้นทุนของ Care Pro ที่ให้ฟรีไม่จำกัด", row=19, new=True)
a1("corrn","จำนวนครั้ง Corrective เฉลี่ย ต่อปี ต่อชุดคู่", 2, "ครั้ง/ปี",
   "ปรับตามสถิติจริง — ตัวเลขนี้คือความเสี่ยงหลักของแพ็กเกจ Care Pro", FMT_NUM, row=20, new=True)
a1("flex", "ต้นทุนท่อ Flex ลมร้อน 4 นิ้ว ต่อเส้น", 600, "บาท/เส้น",
   "Care Pro ให้ฟรีปีละ 1 เส้นต่อเครื่องอบ 1 เครื่อง", row=21, new=True)
a1("kmcost","ต้นทุนเดินทางส่วนเกินจริง", 8, "บาท/กม.",
   "ค่าน้ำมัน + สึกหรอจริง — เดิมไม่ถูกนับเป็นต้นทุนในชีต 3 ทำให้ GM สูงเกินจริง", FMT_THB2, row=22, new=True)
dv_basis = DV("CostBasis", s1)
a1("basis","ฐานต้นทุนที่ใช้ตั้งราคา", "ค่าเฉลี่ย (แนะนำ)", "เลือก",
   "งานกระจุกจุดเดียว = ร้านสะดวกซักที่มีเครื่องหลายตัวในจุดเดียว | Route = งานกระจายหลายจุด",
   None, row=23, new=True, dv=dv_basis)
s1["B23"].alignment = Alignment(horizontal="center", vertical="center")

band(s1, 25, "ต้นทุนที่คำนวณได้ / Calculated cost", "D")
for col, t in zip("ABCD", ["รายการ", "ค่า", "หน่วย", "สูตร"]):
    put(s1, f"{col}26", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
def a1c(key, label, formula, unit, note, fmt=FMT_THB, row=None, bold=False, fill=None):
    put(s1, f"A{row}", label, size=10, bold=bold, fill=fill)
    put(s1, f"B{row}", formula, fmt=fmt, align="right", bold=bold, fill=fill)
    put(s1, f"C{row}", unit, size=9, color="595959", align="center", fill=fill)
    put(s1, f"D{row}", note, size=9, color="595959", wrap=True, fill=fill)
    A1ROWS[key] = row
B = lambda k: f"$B${A1ROWS[k]}"
a1c("perday", "ต้นทุนค่าแรง+รถ ต่อวัน", f"={B('lab1')}+{B('lab2')}+{B('car')}", "บาท/วัน",
    "ช่างหลัก + ผู้ช่วย + ค่ารถ", row=27)
a1c("cst1", "ต้นทุนต่อเครื่อง — งานกระจุกจุดเดียว",
    f"=ROUND({B('perday')}/{B('cap1')}+{B('mat')},0)", "บาท/เครื่อง", "ต้นทุนต่อวัน ÷ เครื่องต่อวัน + วัสดุ", row=28)
a1c("cst2", "ต้นทุนต่อเครื่อง — งาน Route หลายจุด",
    f"=ROUND({B('perday')}/{B('cap2')}+{B('mat')},0)", "บาท/เครื่อง", "ต้นทุนต่อวัน ÷ เครื่องต่อวัน + วัสดุ", row=29)
a1c("cstavg", "ต้นทุนเฉลี่ยของสองกรณี", f"=ROUND(AVERAGE({B('cst1')},{B('cst2')}),0)",
    "บาท/เครื่อง", "ค่าเฉลี่ย", row=30)
a1c("cw", "ต้นทุนต่อเครื่องซัก (ที่ใช้ตั้งราคา)",
    f'=IF({B("basis")}="งานกระจุกจุดเดียว",{B("cst1")},IF({B("basis")}="งาน Route หลายจุด",{B("cst2")},{B("cstavg")}))',
    "บาท/เครื่อง/รอบ", "ตามฐานต้นทุนที่เลือกไว้ด้านบน", row=31, bold=True, fill=KPIFIL)
a1c("cd", "★ ต้นทุนต่อเครื่องอบ (ที่ใช้ตั้งราคา)", f"=ROUND({B('cw')}*{B('dryf')},0)",
    "บาท/เครื่อง/รอบ", "ต้นทุนเครื่องซัก x ตัวคูณเครื่องอบ", row=32, bold=True, fill=KPIFIL)
a1c("cpair", "★ ต้นทุนต่อ 1 ชุดคู่ ต่อรอบ PM", f"={B('cw')}+{B('cd')}",
    "บาท/ชุดคู่/รอบ", "ตัวเลขหลักที่ชีต 2 และ 3 ใช้คิดต้นทุนทั้งหมด", row=33, bold=True, fill=OKFILL)

band(s1, 35, "ราคาตลาดอ้างอิง (เก็บไว้เทียบ ไม่เชื่อมกับสูตร)", "D")
for col, t in zip("ABCD", ["ผู้ให้บริการ", "ราคา", "หน่วย", "ที่มา"]):
    put(s1, f"{col}36", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
MKT = [("CODE CLEAN — PM ตามรอบ", 600, "บาท/เครื่อง/ครั้ง", "code-clean.com/services/preventive-maintenance (ขั้นต่ำ 4,800 บาท/ครั้ง)"),
       ("CODE CLEAN — ซ่อมเมื่อเสีย เครื่องแรก", 1500, "บาท/ครั้ง", "ไม่รวมอะไหล่"),
       ("CODE CLEAN — ซ่อมเมื่อเสีย เครื่องที่ 2 ขึ้นไป", 1000, "บาท/ครั้ง", "ไม่รวมอะไหล่"),
       ("CODE CLEAN — ค่าเดินทางส่วนเกิน", 10, "บาท/กม.", "ฟรี 50 กม.แรก คิดเฉพาะขาไป จาก hub 4 จังหวัด")]
for i, (n, p, u, src) in enumerate(MKT):
    r = 37 + i
    put(s1, f"A{r}", n, size=9)
    inp(s1, f"B{r}", p, FMT_THB)
    put(s1, f"C{r}", u, size=9, color="595959", align="center")
    put(s1, f"D{r}", src, size=8, color="595959", wrap=True)
put(s1, "A42", "ข้อควรระวัง: ราคาตลาดอ้างอิงเป็นข้อมูล ณ วันที่จัดทำ ควรตรวจสอบใหม่ก่อนใช้อ้างอิงกับลูกค้า",
    italic=True, size=9, color="C00000", border=False)
a1c("mktgm", "★ GM ที่ได้ถ้าตั้งราคาเท่าราคาตลาด (CODE CLEAN)",
    f"=IFERROR((B37*2-{B('cpair')})/(B37*2),0)", "%",
    "ราคาตลาด 600 บาท/เครื่อง/ครั้ง = 1,200 บาท/ชุดคู่/รอบ เทียบต้นทุนต่อชุดคู่ของเรา "
    "— ใช้เป็นหลักฐานว่าเป้า GM ที่ตั้งไว้สมเหตุสมผลกับตลาดหรือไม่",
    FMT_PCT, row=44, bold=True, fill=KPIFIL)
S1 = lambda k: f"'1_สมมติฐานต้นทุน'!{B(k)}"

# ============================================================ 2_ราคาแพ็คเกจ
s2 = wb.create_sheet("2_ราคาแพ็คเกจ")
s2.sheet_view.showGridLines = False
s2.sheet_properties.tabColor = "2E75B6"
for c, w in zip("ABCDEFGHIJK", [30, 14, 12, 14, 18, 14, 12, 26, 16, 16, 44]):
    s2.column_dimensions[c].width = w
s2.merge_cells("A1:K1")
put(s2, "A1", "โครงสร้างราคาแพ็คเกจ และการตรวจสอบ Gross Margin", bold=True, size=14,
    color=C_WHT, fill=NAVY, align="left", border=False)
s2.row_dimensions[1].height = 26
put(s2, "A2", "แก้ราคาขายในคอลัมน์ B (สีเหลือง) แล้วดู GM ในคอลัมน์ G — คอลัมน์ I บอกราคาขั้นต่ำที่ยังได้ GM ตามเป้า",
    italic=True, size=9, color="595959", border=False)

HDR2 = ["แพ็คเกจ (ต่อ 1 ชุดคู่ ต่อปี)", "ราคาขาย/ปี", "รอบ PM/ปี", "ต้นทุนงาน PM",
        "★ ต้นทุนบริการที่แถมฟรี", "ต้นทุนรวม", "Gross Margin", "สถานะเทียบเป้า",
        "★ ราคาขั้นต่ำ (Floor)", "★ ราคาเป้าหมาย (Target)", "หมายเหตุ"]
for i, t in enumerate(HDR2):
    put(s2, f"{get_column_letter(i+1)}4", t, bold=True, color=C_WHT, fill=NAVY,
        align="center", wrap=True, size=9)
s2.row_dimensions[4].height = 32
PACKS = [
    ("Care — 2 รอบ/ปี",     2600, 2, "0",
     "PM ทุก 6 เดือน · ถอดล้างท่อลมร้อนปีละ 1 · ตอบสนอง 72 ชม. · ค่าแรงซ่อมด่วนเก็บ 1,500/ครั้ง"),
    ("Care+ — 4 รอบ/ปี",    4900, 4, "0",
     "PM ทุก 3 เดือน · ถอดล้างท่อปีละ 2 · ตอบสนอง 48 ชม. · ค่าแรงซ่อมด่วนลด 50% เหลือ 750/ครั้ง"),
    ("Care Pro — 4 รอบ/ปี", 7900, 4, "FLEXCORR",
     "ตอบสนอง 24 ชม. · ถอดล้างทุกรอบ · ★ ฟรีค่าแรงซ่อมไม่จำกัด + ท่อ Flex ปีละ 1 เส้น = ต้นทุนแฝงที่ต้องตั้งเผื่อ"),
]
P_START = 5
for i, (name, price, rounds, extra, note) in enumerate(PACKS):
    r = P_START + i
    put(s2, f"A{r}", name, bold=True, size=10)
    inp(s2, f"B{r}", price, FMT_THB)
    inp(s2, f"C{r}", rounds, FMT_INT, align="center")
    put(s2, f"D{r}", f"=C{r}*{S1('cpair')}", fmt=FMT_THB, align="right")
    ex = "0" if extra == "0" else f"{S1('flex')}+{S1('corrn')}*{S1('corrc')}"
    put(s2, f"E{r}", f"={ex}", fmt=FMT_THB, align="right", color=(C_CALC if extra == "0" else "C00000"))
    put(s2, f"F{r}", f"=D{r}+E{r}", fmt=FMT_THB, align="right", bold=True)
    put(s2, f"G{r}", f"=IFERROR((B{r}-F{r})/B{r},0)", fmt=FMT_PCT, align="right", bold=True)
    put(s2, f"H{r}", f'=IF(G{r}>={S1("gmtgt")},"ผ่านเป้าหมาย",'
                     f'IF(G{r}>={S1("gmfloor")},"รับได้ (ต่ำกว่าเป้า "&TEXT({S1("gmtgt")}-G{r},"0.0%")&")",'
                     f'"ต่ำกว่าขั้นต่ำ — ทบทวนราคา"))', size=9, align="center", wrap=True)
    put(s2, f"I{r}", f'=IFERROR(CEILING(F{r}/(1-{S1("gmfloor")}),50),0)', fmt=FMT_THB, align="right",
        bold=True, fill=WARN)
    put(s2, f"J{r}", f'=IFERROR(CEILING(F{r}/(1-{S1("gmtgt")}),50),0)', fmt=FMT_THB, align="right",
        bold=True, fill=OKFILL)
    put(s2, f"K{r}", note, size=8, color="595959", wrap=True)
P_END = P_START + len(PACKS) - 1
PRICE_RNG = f"'2_ราคาแพ็คเกจ'!$B${P_START}:$B${P_END}"
ROUND_RNG = f"'2_ราคาแพ็คเกจ'!$C${P_START}:$C${P_END}"
EXTRA_RNG = f"'2_ราคาแพ็คเกจ'!$E${P_START}:$E${P_END}"
put(s2, "A9", "ราคาขั้นต่ำปัดขึ้นทีละ 50 บาท | ต้นทุนงาน PM = รอบต่อปี x ต้นทุนต่อชุดคู่ต่อรอบ (ชีต 1 แถว 33)",
    italic=True, size=9, color="595959", border=False)

band(s2, 11, "★ ตรวจสอบงานซ่อมด่วน (Corrective) — คิดแยกจากแพ็คเกจ", "K", fill=DARK)
for i, t in enumerate(["รายการ", "เรียกเก็บ/ครั้ง", "ต้นทุน/ครั้ง", "กำไร/ครั้ง", "GM", "", "", "", "", "หมายเหตุ"]):
    if t: put(s2, f"{get_column_letter(i+1)}12", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
CORR = [("Care — เก็บเต็ม", 1500, "ราคาตามใบเสนอราคา"),
        ("Care+ — ลด 50%", 750, "★ ราคาลด 50% ต่ำกว่าต้นทุนต่อครั้ง — ยิ่งลูกค้าเรียกมาก ยิ่งขาดทุน"),
        ("Care Pro — ฟรีไม่จำกัด", 0, "ต้นทุนถูกตั้งเผื่อไว้ในคอลัมน์ E ของตารางด้านบนแล้ว")]
for i, (n, p, note) in enumerate(CORR):
    r = 13 + i
    put(s2, f"A{r}", n, size=10)
    inp(s2, f"B{r}", p, FMT_THB)
    put(s2, f"C{r}", f"={S1('corrc')}", fmt=FMT_THB, align="right", color=C_LINK)
    put(s2, f"D{r}", f"=B{r}-C{r}", fmt=FMT_THB, align="right", bold=True)
    put(s2, f"E{r}", f"=IFERROR(D{r}/B{r},-1)", fmt=FMT_PCT, align="right")
    for col in "FGHI":
        put(s2, f"{col}{r}", None)
    put(s2, f"J{r}", note, size=8, color=("C00000" if i == 1 else "595959"), wrap=True)
CORR_SELL = f"'2_ราคาแพ็คเกจ'!$B$13:$B$15"

band(s2, 17, "ราคางาน PM แบบครั้งเดียว (Single PM)", "K")
for i, t in enumerate(["รายการ", "ราคาขาย", "", "ต้นทุน", "", "", "Gross Margin", "สถานะ", "ราคาขั้นต่ำ", "หมายเหตุ"]):
    if t: put(s2, f"{get_column_letter(i+1)}18", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
SINGLE = [("เครื่องซัก 18 kg", 650, "cw", "งานมาตรฐาน"),
          ("เครื่องอบ 14 kg", 850, "cd", "หนักกว่าเพราะต้องถอดล้างท่อลมร้อนและตรวจระบบแก๊ส")]
SG_START = 19
for i, (n, p, ck, note) in enumerate(SINGLE):
    r = SG_START + i
    put(s2, f"A{r}", n, size=10)
    inp(s2, f"B{r}", p, FMT_THB)
    put(s2, f"D{r}", f"={S1(ck)}", fmt=FMT_THB, align="right", color=C_LINK)
    put(s2, f"G{r}", f"=IFERROR((B{r}-D{r})/B{r},0)", fmt=FMT_PCT, align="right", bold=True)
    put(s2, f"H{r}", f'=IF(G{r}>={S1("gmtgt")},"ผ่านเป้าหมาย",IF(G{r}>={S1("gmfloor")},"รับได้","ต่ำกว่าขั้นต่ำ"))',
        size=9, align="center")
    put(s2, f"I{r}", f'=IFERROR(CEILING(D{r}/(1-{S1("gmfloor")}),50),0)', fmt=FMT_THB, align="right", fill=WARN)
    put(s2, f"J{r}", f'=IFERROR(CEILING(D{r}/(1-{S1("gmtgt")}),50),0)', fmt=FMT_THB, align="right", fill=OKFILL)
    put(s2, f"K{r}", note, size=8, color="595959", wrap=True)
    for col in "CEF":
        put(s2, f"{col}{r}", None)
SGL_W = f"'2_ราคาแพ็คเกจ'!$B${SG_START}"
SGL_D = f"'2_ราคาแพ็คเกจ'!$B${SG_START+1}"

band(s2, 22, "★ ส่วนลดตามจำนวนชุดคู่ — พร้อมต้นทุนและ GM ของทุกขั้น (เวอร์ชัน 1 ไม่มีการตรวจส่วนนี้)", "K", fill=DARK)
TH = ["จำนวนชุดคู่", "ส่วนลด", "รอบฟรีเพิ่ม", "Care ราคา/ชุด", "Care GM",
      "Care+ ราคา/ชุด", "Care+ GM", "Care Pro ราคา/ชุด", "Care Pro GM", "สถานะขั้นนี้"]
for i, t in enumerate(TH):
    put(s2, f"{get_column_letter(i+1)}23", t, bold=True, color=C_WHT, fill=NAVY,
        align="center", wrap=True, size=9)
s2.row_dimensions[23].height = 30
TIERS = [(1, 0.00, 0), (2, 0.10, 0), (3, 0.10, 0), (4, 0.20, 0), (5, 0.20, 0), (6, 0.25, 1)]
T_START = 24
for i, (pairs, disc, free) in enumerate(TIERS):
    r = T_START + i
    put(s2, f"A{r}", pairs, fmt=FMT_INT, align="center", bold=True)
    inp(s2, f"B{r}", disc, FMT_PCT)
    inp(s2, f"C{r}", free, FMT_INT, align="center")
    for j, prow in enumerate([P_START, P_START + 1, P_START + 2]):
        pc = get_column_letter(4 + j * 2)      # D, F, H
        gc = get_column_letter(5 + j * 2)      # E, G, I
        put(s2, f"{pc}{r}", f"=ROUND($B${prow}*(1-$B{r}),0)", fmt=FMT_THB, align="right")
        put(s2, f"{gc}{r}", f"=IFERROR(({pc}{r}-(($C${prow}+$C{r})*{S1('cpair')}+$E${prow}))/{pc}{r},0)",
            fmt=FMT_PCT, align="right", bold=True)
    put(s2, f"J{r}", f'=IF(MIN(E{r},G{r},I{r})<0,"ขาดทุน — ต้องแก้ทันที",'
                     f'IF(MIN(E{r},G{r},I{r})<{S1("gmfloor")},"ต่ำกว่าขั้นต่ำ",'
                     f'IF(MIN(E{r},G{r},I{r})<{S1("gmtgt")},"รับได้","ผ่านเป้าหมายทุกแพ็คเกจ")))',
        size=9, align="center", wrap=True)
T_END = T_START + len(TIERS) - 1
DISC_PAIRS = f"'2_ราคาแพ็คเกจ'!$A${T_START}:$A${T_END}"
DISC_RATE  = f"'2_ราคาแพ็คเกจ'!$B${T_START}:$B${T_END}"
DISC_FREE  = f"'2_ราคาแพ็คเกจ'!$C${T_START}:$C${T_END}"
put(s2, f"A{T_END+1}", "หมายเหตุ: ขั้น 6 ชุดคู่ขึ้นไปในเงื่อนไขเดิมให้ทั้งส่วนลด 25% และรอบบริการฟรีอีก 1 รอบพร้อมกัน "
                       "ซึ่งเป็นจุดที่ GM ติดลบ — ทางเลือกคือตัดรอบฟรีออก หรือลดส่วนลดเหลือ 15%",
    italic=True, size=9, color="C00000", border=False)
for rng in [f"G{P_START}:G{P_END}", f"E{T_START}:E{T_END}", f"G{T_START}:G{T_END}", f"I{T_START}:I{T_END}",
            f"G{SG_START}:G{SG_START+1}"]:
    first = rng.split(":")[0]
    col, row = first[0], first[1:]
    s2.conditional_formatting.add(rng, FormulaRule(
        formula=[f"${col}{row}<0"], fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(name=FONT, bold=True, color="9C0006"), stopIfTrue=True))
    s2.conditional_formatting.add(rng, FormulaRule(
        formula=[f"${col}{row}<{S1('gmfloor')}"], fill=PatternFill("solid", fgColor="F8CBAD"),
        font=Font(name=FONT, bold=True, color="833C00"), stopIfTrue=True))
    s2.conditional_formatting.add(rng, FormulaRule(
        formula=[f"${col}{row}<{S1('gmtgt')}"], fill=PatternFill("solid", fgColor="FFF2CC"), stopIfTrue=True))
    s2.conditional_formatting.add(rng, FormulaRule(
        formula=[f"${col}{row}>={S1('gmtgt')}"], fill=PatternFill("solid", fgColor=OKFILL), stopIfTrue=True))

# ---- เพดานส่วนลด และบันไดส่วนลดที่แนะนำ (เพิ่มใน v2.1)
CAP0 = T_END + 3
band(s2, CAP0, "★ เพดานส่วนลดที่ยังผ่าน GM ขั้นต่ำ — ใช้ตอบทันทีเวลาลูกค้าขอลดราคา", "K", fill=DARK)
CAPH = ["แพ็คเกจ", "ราคาป้าย", "ต้นทุน (รอบปกติ)", "เพดานส่วนลด", "ต้นทุน (+รอบฟรี 1 รอบ)",
        "เพดานส่วนลดถ้าแถมรอบฟรี", "", "", "", "", "หมายเหตุ"]
for i, t in enumerate(CAPH):
    if t: put(s2, f"{get_column_letter(i+1)}{CAP0+1}", t, bold=True, color=C_WHT, fill=NAVY,
              align="center", wrap=True, size=9)
s2.row_dimensions[CAP0 + 1].height = 30
for i, prow in enumerate([P_START, P_START + 1, P_START + 2]):
    r = CAP0 + 2 + i
    put(s2, f"A{r}", f"=LEFT($A${prow},FIND(\" —\",$A${prow})-1)", size=10, bold=True)
    put(s2, f"B{r}", f"=$B${prow}", fmt=FMT_THB, align="right", color=C_LINK)
    put(s2, f"C{r}", f"=$F${prow}", fmt=FMT_THB, align="right", color=C_LINK)
    put(s2, f"D{r}", f'=IFERROR(MAX(0,1-(C{r}/(1-{S1("gmfloor")}))/B{r}),0)', fmt=FMT_PCT,
        align="right", bold=True)
    put(s2, f"E{r}", f"=$C${prow}*{S1('cpair')}+{S1('cpair')}+$E${prow}", fmt=FMT_THB, align="right")
    put(s2, f"F{r}", f'=IFERROR(MAX(0,1-(E{r}/(1-{S1("gmfloor")}))/B{r}),0)', fmt=FMT_PCT,
        align="right", bold=True)
    for col in "GHIJ":
        put(s2, f"{col}{r}", None)
    put(s2, f"K{r}", "0% = ลดไม่ได้เลยที่ GM ขั้นต่ำปัจจุบัน", size=8, color="595959", wrap=True)
put(s2, f"A{CAP0+5}", "เพดานนี้คิดจากฐานต้นทุนที่เลือกไว้ในชีต 1 — ลูกค้าหลายชุดคู่ในจุดเดียวควรเลือกฐาน "
                      "'งานกระจุกจุดเดียว' ซึ่งต้นทุนต่ำกว่า ทำให้ลดราคาได้มากขึ้นอย่างมีเหตุผล",
    italic=True, size=9, color="595959", border=False)

REC0 = CAP0 + 7
band(s2, REC0, "★ บันไดส่วนลดที่แนะนำ (ทางเลือกแทนบันไดเดิมที่ทำให้ขาดทุน)", "K", fill=DARK)
RECH = ["จำนวนชุดคู่", "ส่วนลดที่แนะนำ", "รอบฟรีเพิ่ม", "Care GM", "Care+ GM", "Care Pro GM",
        "", "", "", "", "สถานะ"]
for i, t in enumerate(RECH):
    if t: put(s2, f"{get_column_letter(i+1)}{REC0+1}", t, bold=True, color=C_WHT, fill=NAVY,
              align="center", wrap=True, size=9)
s2.row_dimensions[REC0 + 1].height = 28
RECT = [("1", 0.00, 0), ("2-3", 0.05, 0), ("4-5", 0.10, 0), ("6 ขึ้นไป", 0.12, 0)]
for i, (label, disc, free) in enumerate(RECT):
    r = REC0 + 2 + i
    put(s2, f"A{r}", label, align="center", bold=True)
    inp(s2, f"B{r}", disc, FMT_PCT)
    inp(s2, f"C{r}", free, FMT_INT, align="center")
    for j, prow in enumerate([P_START, P_START + 1, P_START + 2]):
        gc = get_column_letter(4 + j)
        put(s2, f"{gc}{r}", f"=IFERROR((ROUND($B${prow}*(1-$B{r}),0)-(($C${prow}+$C{r})*{S1('cpair')}"
                            f"+$E${prow}))/ROUND($B${prow}*(1-$B{r}),0),0)", fmt=FMT_PCT,
            align="right", bold=True)
    for col in "GHIJ":
        put(s2, f"{col}{r}", None)
    put(s2, f"K{r}", f'=IF(MIN(D{r},E{r},F{r})>={S1("gmfloor")},"ผ่านขั้นต่ำทุกแพ็คเกจ",'
                     f'IF(MIN(D{r},E{r})>={S1("gmfloor")},"Care/Care+ ผ่าน · Care Pro ต่ำกว่าขั้นต่ำตั้งแต่ราคาป้าย",'
                     f'"ยังต่ำกว่าขั้นต่ำ"))', size=9, align="center", wrap=True)
put(s2, f"A{REC0+6}", "ข้อเสนอนี้ตัดรอบบริการฟรีออกทั้งหมด และจำกัดส่วนลดสูงสุดที่ 12% "
                      "— ถ้าต้องการให้ของแถมกับลูกค้ารายใหญ่ ควรเป็นสิ่งที่ต้นทุนต่ำ เช่น ยกระดับเวลาตอบสนอง "
                      "หรือฟรี Single PM ให้เครื่องเสริม แทนการเพิ่มรอบ PM ซึ่งเป็นต้นทุนเต็มจำนวน",
    italic=True, size=9, color="C00000", border=False)
put(s2, f"A{REC0+7}", "ตัวเลข GM ด้านบนคิดบนฐานต้นทุนที่เลือกไว้ในชีต 1 — ถ้าเปลี่ยนเป็นฐาน 'งานกระจุกจุดเดียว' "
                      "(ซึ่งตรงกับลักษณะงานร้านสะดวกซักที่มีเครื่องหลายตัวในจุดเดียว) ทุกขั้นของบันไดนี้จะผ่าน GM ขั้นต่ำ "
                      "สำหรับ Care และ Care+",
    italic=True, size=9, color="595959", border=False)
for rng in [f"D{CAP0+2}:D{CAP0+4}", f"F{CAP0+2}:F{CAP0+4}"]:
    first = rng.split(":")[0]
    s2.conditional_formatting.add(rng, FormulaRule(formula=[f"${first[0]}{first[1:]}=0"],
        fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONT, bold=True, color="833C00"),
        stopIfTrue=True))
for j in range(3):
    gc = get_column_letter(4 + j)
    rng = f"{gc}{REC0+2}:{gc}{REC0+5}"
    s2.conditional_formatting.add(rng, FormulaRule(formula=[f"${gc}{REC0+2}<{S1('gmfloor')}"],
        fill=PatternFill("solid", fgColor="F8CBAD"), stopIfTrue=True))
    s2.conditional_formatting.add(rng, FormulaRule(formula=[f"${gc}{REC0+2}>={S1('gmfloor')}"],
        fill=PatternFill("solid", fgColor=OKFILL), stopIfTrue=True))

# ============================================================ 3_คำนวณใบเสนอราคา
s3 = wb.create_sheet("3_คำนวณใบเสนอราคา")
s3.sheet_view.showGridLines = False
s3.sheet_properties.tabColor = "C00000"
for c, w in zip("ABCDE", [46, 14, 16, 18, 62]):
    s3.column_dimensions[c].width = w
s3.merge_cells("A1:E1")
put(s3, "A1", "ตัวคำนวณใบเสนอราคา PM รายลูกค้า", bold=True, size=14, color=C_WHT,
    fill=NAVY, align="left", border=False)
s3.row_dimensions[1].height = 26
put(s3, "A2", "กรอกเฉพาะช่องสีเหลือง — ผลลัพธ์ด้านล่างนำไปกรอกในใบเสนอราคา PDF ได้ทันที",
    italic=True, size=9, color="595959", border=False)
dv_pack3, dv_yn3 = DV("Pack", s3), DV("YesNo", s3)
band(s3, 3, "ข้อมูลนำเข้า", "E")
IN3 = {}
def i3(key, label, value, row, fmt=None, note="", dv=None, align="right"):
    put(s3, f"A{row}", label, size=10)
    c = inp(s3, f"B{row}", value, fmt, align=align)
    if dv: dv.add(c)
    put(s3, f"E{row}", note, size=9, color="595959", wrap=True)
    IN3[key] = f"$B${row}"
i3("cust",  "ชื่อลูกค้า / อาคาร", "ลูกค้าตัวอย่าง", 4, align="left")
i3("pairs", "จำนวนชุดคู่ (ซัก+อบ)", 2, 5, FMT_INT)
i3("sw",    "จำนวนเครื่องซักเดี่ยว (ไม่จับคู่)", 0, 6, FMT_INT)
i3("sd",    "จำนวนเครื่องอบเดี่ยว (ไม่จับคู่)", 0, 7, FMT_INT)
i3("pack",  "แพ็คเกจที่เลือก", "Care+", 8, None, "เลือกจากรายการ", dv=dv_pack3, align="center")
i3("km",    "ระยะทางจากฐานทีมช่าง (กม.)", 45, 9, FMT_NUM, "ฟรีในรัศมี 30 กม.")
i3("night", "จำนวนคืนที่ต้องค้างแรม", 0, 10, FMT_INT)
i3("route", "อยู่ในรอบ Route ประจำโซน?", "ไม่ใช่", 11, None, "ถ้าใช่ = ยกเว้นค่าเดินทางส่วนเกิน",
    dv=dv_yn3, align="center")
i3("pre3",  "ชำระล่วงหน้า 3 ปี?", "ไม่ใช่", 12, None, "ลด 15% ตามข้อ 5.2 ของสัญญา",
    dv=dv_yn3, align="center")
i3("corrn", "★ จำนวนครั้งเรียกซ่อมด่วนที่คาดว่าจะเกิด/ปี", 2, 13, FMT_NUM,
    "ใช้คิดทั้งรายได้ (Care/Care+) และต้นทุน (ทุกแพ็คเกจ)")
K = lambda k: IN3[k]
PIDX  = f"IFERROR(MATCH({K('pack')},{REF['Pack']},0),2)"
PRICE = f"INDEX({PRICE_RNG},{PIDX})"
ROUNDS= f"INDEX({ROUND_RNG},{PIDX})"
CSELL = f"INDEX({CORR_SELL},{PIDX})"
TIDX  = f"IFERROR(MATCH(MIN(MAX({K('pairs')},1),6),{DISC_PAIRS},0),1)"
DISC  = f"INDEX({DISC_RATE},{TIDX})"
FREE  = f"IF({K('pairs')}=0,0,INDEX({DISC_FREE},{TIDX}))"
EXKM  = f'IF({K("route")}="ใช่",0,MAX(0,{K("km")}-30))'

band(s3, 15, "รายได้ของดีลนี้ (ต่อปี)", "E")
for col, t in zip("ABCDE", ["รายการ", "จำนวน", "ราคา/หน่วย", "จำนวนเงิน", "หมายเหตุ"]):
    put(s3, f"{col}16", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
def r3(row, label, qty, price, amount, note, bold=False, fill=None, fmt=FMT_THB):
    put(s3, f"A{row}", label, size=10, bold=bold, fill=fill)
    put(s3, f"B{row}", qty, fmt=FMT_NUM, align="right", fill=fill)
    put(s3, f"C{row}", price, fmt=FMT_THB, align="right", fill=fill)
    put(s3, f"D{row}", amount, fmt=fmt, align="right", bold=bold, fill=fill)
    put(s3, f"E{row}", note, size=8, color="595959", wrap=True, fill=fill)
r3(17, "ค่าบริการรายปี — ชุดคู่", f"={K('pairs')}", f"={PRICE}", "=B17*C17",
   "ราคาต่อชุดคู่ต่อปี ดึงจากชีต 2 ตามแพ็คเกจที่เลือก")
r3(18, "ค่าบริการรายปี — เครื่องซักเดี่ยว", f"={K('sw')}", f"={SGL_W}*{ROUNDS}", "=B18*C18",
   "ราคา Single PM เครื่องซัก x จำนวนรอบต่อปี")
r3(19, "ค่าบริการรายปี — เครื่องอบเดี่ยว", f"={K('sd')}", f"={SGL_D}*{ROUNDS}", "=B19*C19",
   "ราคา Single PM เครื่องอบ x จำนวนรอบต่อปี")
r3(20, "รวมค่าบริการก่อนส่วนลด", None, None, "=SUM(D17:D19)", "", bold=True, fill=KPIFIL)
r3(21, "ส่วนลดตามจำนวนชุดคู่", None, f"={DISC}", "=-ROUND(D20*C21,0)",
   "ดึงอัตราส่วนลดจากตารางขั้นบันไดในชีต 2")
s3["C21"].number_format = FMT_PCT
r3(22, "ค่าเดินทางส่วนเกิน (เรียกเก็บ)", f"={EXKM}", f"={S1('kmsell')}",
   f"=ROUND(B22*C22*({ROUNDS}+{FREE}),0)", "กม.ส่วนเกิน x อัตรา x จำนวนรอบ (รวมรอบฟรี)")
r3(23, "ค่าที่พักค้างคืน (เรียกเก็บ)", f"={K('night')}", f"={S1('hotel')}", "=B23*C23", "")
r3(24, "★ รายได้ค่าแรงซ่อมด่วน", f"={K('corrn')}", f"={CSELL}", "=B24*C24",
   "Care เก็บ 1,500 · Care+ เก็บ 750 · Care Pro ฟรี (= 0)")
r3(25, "ราคาสุทธิก่อน VAT (ต่อปี)", None, None, "=D20+D21+D22+D23+D24", "", bold=True, fill=KPIFIL)
r3(26, "ภาษีมูลค่าเพิ่ม 7%", None, None, "=ROUND(D25*0.07,0)", "")
r3(27, "รวมทั้งสิ้น (ต่อปี)", None, None, "=D25+D26", "นำตัวเลขนี้ไปกรอกในใบเสนอราคา", bold=True, fill=OKFILL)

band(s3, 29, "★ ต้นทุนของดีลนี้ (ต่อปี) — เวอร์ชัน 1 คิดเฉพาะงาน PM และคิดเครื่องอบเท่าเครื่องซัก", "E", fill=DARK)
for col, t in zip("ABCDE", ["รายการ", "จำนวน", "ต้นทุน/หน่วย", "จำนวนเงิน", "หมายเหตุ"]):
    put(s3, f"{col}30", t, bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
r3(31, "ต้นทุนงาน PM — ชุดคู่", f"={K('pairs')}*({ROUNDS}+{FREE})", f"={S1('cpair')}", "=B31*C31",
   "จำนวนชุดคู่ x (รอบตามแพ็คเกจ + รอบฟรี) x ต้นทุนต่อชุดคู่ต่อรอบ")
r3(32, "ต้นทุนงาน PM — เครื่องเดี่ยว", f"=({K('sw')}+{K('sd')})*({ROUNDS}+{FREE})",
   f"=IFERROR(({K('sw')}*{S1('cw')}+{K('sd')}*{S1('cd')})/MAX({K('sw')}+{K('sd')},1),0)",
   f"=({K('sw')}*{S1('cw')}+{K('sd')}*{S1('cd')})*({ROUNDS}+{FREE})",
   "แยกต้นทุนเครื่องซักและเครื่องอบตามจริง | รวมรอบฟรีที่แถมให้ (มีต้นทุนแต่ไม่มีรายได้)")
r3(33, "★ ต้นทุนเดินทางส่วนเกิน (จริง)", f"={EXKM}", f"={S1('kmcost')}",
   f"=ROUND(B33*C33*({ROUNDS}+{FREE}),0)", "ค่าน้ำมัน+สึกหรอจริง — เวอร์ชัน 1 ไม่ได้นับส่วนนี้")
r3(34, "ต้นทุนที่พักค้างคืน", f"={K('night')}", f"={S1('hotel')}", "=B34*C34", "ส่งผ่านเท่าราคาที่เรียกเก็บ")
r3(35, "★ ต้นทุนงานซ่อมด่วน", f"={K('corrn')}", f"={S1('corrc')}", "=B35*C35",
   "เกิดทุกแพ็คเกจ — Care Pro ไม่มีรายได้มาชดเชย")
r3(36, "★ ต้นทุนท่อ Flex (เฉพาะ Care Pro)", f'=IF({K("pack")}="Care Pro",{K("pairs")},0)',
   f"={S1('flex')}", "=B36*C36", "Care Pro ให้ฟรีปีละ 1 เส้นต่อเครื่องอบ 1 เครื่อง")
r3(37, "รวมต้นทุนทั้งหมด", None, None, "=SUM(D31:D36)", "", bold=True, fill=KPIFIL)

band(s3, 39, "ผลลัพธ์และเพดานการต่อรอง", "E")
def o3(row, label, formula, fmt, note, bold=True, fill=None):
    put(s3, f"A{row}", label, size=10, bold=bold, fill=fill)
    s3.merge_cells(f"B{row}:C{row}")
    put(s3, f"B{row}", None, fill=fill)
    s3[f"C{row}"].border = BOX
    if fill: s3[f"C{row}"].fill = PatternFill("solid", fgColor=fill)
    put(s3, f"D{row}", formula, fmt=fmt, align="right", bold=bold, fill=fill)
    put(s3, f"E{row}", note, size=8, color="595959", wrap=True, fill=fill)
o3(40, "กำไรขั้นต้นเป็นตัวเงิน (ต่อปี)", "=D25-D37", FMT_THB, "ราคาสุทธิก่อน VAT − ต้นทุนรวม")
o3(41, "Gross Margin ของดีลนี้", "=IFERROR((D25-D37)/D25,0)", FMT_PCT, "", fill=KPIFIL)
o3(42, "สถานะเทียบเป้าหมาย",
   f'=IF(D41>={S1("gmtgt")},"ผ่านเป้าหมาย",IF(D41>={S1("gmfloor")},'
   f'"รับได้ — ต่ำกว่าเป้า "&TEXT({S1("gmtgt")}-D41,"0.0%"),'
   f'"ต่ำกว่าขั้นต่ำ "&TEXT({S1("gmfloor")}-D41,"0.0%")&" — ทบทวนส่วนลดหรือปฏิเสธงาน"))',
   None, "")
o3(43, "★ ราคาขั้นต่ำที่ยังผ่าน GM ขั้นต่ำ (ก่อน VAT)",
   f'=IFERROR(CEILING(D37/(1-{S1("gmfloor")}),100),0)', FMT_THB,
   "เพดานล่างสุดของการต่อรอง — ห้ามเสนอต่ำกว่านี้", fill=WARN)
o3(44, "★ ราคาที่ควรเสนอเพื่อให้ถึง GM เป้าหมาย (ก่อน VAT)",
   f'=IFERROR(CEILING(D37/(1-{S1("gmtgt")}),100),0)', FMT_THB,
   "ราคาตั้งต้นที่ควรยื่นก่อนเริ่มต่อรอง", fill=OKFILL)
o3(45, "★ ส่วนลดสูงสุดที่ให้ได้โดยไม่หลุดขั้นต่ำ",
   "=IFERROR(MAX(0,1-(D43-D22-D23-D24)/D20),0)", FMT_PCT,
   "เทียบกับส่วนลดที่ให้อยู่ในแถว 21", fill=OKFILL)
o3(46, "คำแนะนำฐานต้นทุนสำหรับดีลนี้",
   f'=IF({K("pairs")}*2+{K("sw")}+{K("sd")}>=6,'
   f'"เครื่อง "&({K("pairs")}*2+{K("sw")}+{K("sd")})&" ตัวในจุดเดียว — ควรเลือกฐาน \'งานกระจุกจุดเดียว\' ที่ชีต 1 แถว 23",'
   f'"เครื่องน้อย — ใช้ค่าเฉลี่ยหรือฐาน Route ตามลักษณะงานจริง")', None, "", bold=False)

band(s3, 48, "ทางเลือก: ชำระล่วงหน้า 3 ปี (ลด 15% และตรึงราคาตลอดสัญญา)", "E")
o3(49, "ราคา 3 ปี หลังส่วนลด 15% (ก่อน VAT)", "=ROUND(D25*3*0.85,0)", FMT_THB,
   "ใช้เมื่อลูกค้าตอบรับข้อเสนอตามข้อ 5.2 ของสัญญา", bold=False)
o3(50, "ส่วนที่ลูกค้าประหยัดได้เทียบจ่ายรายปี", "=D25*3-D49", FMT_THB, "", bold=False)
o3(51, "★ GM ของดีล 3 ปี (ต้นทุนคงที่)", "=IFERROR((D49-D37*3)/D49,0)", FMT_PCT,
   "ยังไม่รวมเงินเฟ้อค่าแรง — ถ้าต้นทุนขึ้นปีละ 5% GM จริงจะต่ำกว่านี้อีกราว 3-4 จุด", fill=KPIFIL)
for rng in ["D41", "D51"]:
    s3.conditional_formatting.add(rng, FormulaRule(formula=[f"{rng}<0"],
        fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, bold=True, color="9C0006"), stopIfTrue=True))
    s3.conditional_formatting.add(rng, FormulaRule(formula=[f"{rng}<{S1('gmfloor')}"],
        fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONT, bold=True, color="833C00"), stopIfTrue=True))
    s3.conditional_formatting.add(rng, FormulaRule(formula=[f"{rng}<{S1('gmtgt')}"],
        fill=PatternFill("solid", fgColor="FFF2CC"), stopIfTrue=True))
    s3.conditional_formatting.add(rng, FormulaRule(formula=[f"{rng}>={S1('gmtgt')}"],
        fill=PatternFill("solid", fgColor=OKFILL), stopIfTrue=True))

# ============================================================ 4_ทะเบียนลูกค้าและรอบPM
s4 = wb.create_sheet("4_ทะเบียนลูกค้าและรอบPM")
s4.sheet_view.showGridLines = False
s4.sheet_properties.tabColor = "70AD47"
H4 = ["เลขที่สัญญา", "ชื่อลูกค้า/อาคาร", "จังหวัด/โซน", "แพ็คเกจ", "จำนวนชุดคู่", "จำนวนเครื่อง",
      "รอบตามสัญญา", "วันเริ่มสัญญา", "วันสิ้นสุดสัญญา", "รอบ 1", "รอบ 2", "รอบ 3", "รอบ 4",
      "★ รอบ 5 (ฟรี)", "รอบที่ทำแล้ว", "รอบคงเหลือ", "กำหนดรอบถัดไป", "★ วันคงเหลือ", "★ สถานะ"]
W4 = [16, 30, 14, 12, 12, 12, 13, 14, 14, 12, 12, 12, 12, 12, 12, 12, 15, 12, 22]
for i, w in enumerate(W4):
    s4.column_dimensions[get_column_letter(i + 1)].width = w
s4.merge_cells("A1:S1")
put(s4, "A1", "ทะเบียนสัญญา PM และตารางติดตามรอบเข้าปฏิบัติงาน", bold=True, size=14,
    color=C_WHT, fill=NAVY, align="left", border=False)
s4.row_dimensions[1].height = 26
put(s4, "A2", "1 บรรทัด = 1 สัญญา | กรอกวันที่เข้าจริงในคอลัมน์ J-N ระบบจะคำนวณรอบคงเหลือ วันครบกำหนดถัดไป และสถานะให้",
    italic=True, size=9, color="595959", border=False)
for i, t in enumerate(H4):
    put(s4, f"{get_column_letter(i+1)}4", t, bold=True, color=C_WHT, fill=NAVY,
        align="center", wrap=True, size=9)
s4.row_dimensions[4].height = 34
dv_pack4 = DV("Pack", s4)
R4_START, R4_END = 5, 44
for r in range(R4_START, R4_END + 1):
    demo = (r == R4_START)
    for col, val in zip("ABC", (["PM-2569/001", "หอพักตัวอย่าง (แถวตัวอย่าง — ลบทิ้งได้)", "ภูเก็ต"] if demo else [None, None, None])):
        inp(s4, f"{col}{r}", val, None, align="left")
    c = inp(s4, f"D{r}", "Care+" if demo else None, None, align="center"); dv_pack4.add(c)
    inp(s4, f"E{r}", 2 if demo else None, FMT_INT)
    put(s4, f"F{r}", f"=IF(E{r}=\"\",\"\",E{r}*2)", fmt=FMT_INT, align="right")
    put(s4, f"G{r}", f'=IF(D{r}="","",IF(D{r}="Care",2,4)+IF(N(E{r})>=6,1,0))', fmt=FMT_INT, align="center")
    inp(s4, f"H{r}", datetime.date(2026, 8, 1) if demo else None, FMT_DATE, align="center")
    put(s4, f"I{r}", f'=IF(H{r}="","",H{r}+364)', fmt=FMT_DATE, align="center")
    for col, val in zip("JKLMN", ([datetime.date(2026, 8, 5), None, None, None, None] if demo else [None]*5)):
        inp(s4, f"{col}{r}", val, FMT_DATE, align="center")
    put(s4, f"O{r}", f"=COUNT(J{r}:N{r})", fmt=FMT_INT, align="center")
    put(s4, f"P{r}", f'=IF(D{r}="","",G{r}-O{r})', fmt=FMT_INT, align="center")
    put(s4, f"Q{r}", f'=IF(OR(H{r}="",N(P{r})<=0),"",IF(O{r}=0,H{r}+7,MAX(J{r}:N{r})+IF(D{r}="Care",182,91)))',
        fmt=FMT_DATE, align="center")
    put(s4, f"R{r}", f'=IF(Q{r}="","",Q{r}-TODAY())', fmt=FMT_INT, align="center")
    put(s4, f"S{r}", f'=IF(A{r}="","",IF(Q{r}="","ครบรอบตามสัญญาแล้ว",'
                     f'IF(R{r}<0,"เกินกำหนด "&ABS(R{r})&" วัน",IF(R{r}<=14,"ใกล้ถึงกำหนด","ปกติ"))))',
        size=9, align="center")
S_RNG = f"S{R4_START}:S{R4_END}"
s4.conditional_formatting.add(S_RNG, FormulaRule(formula=[f'ISNUMBER(SEARCH("เกินกำหนด",$S{R4_START}))'],
    fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=FONT, bold=True, color="9C0006"), stopIfTrue=True))
s4.conditional_formatting.add(S_RNG, FormulaRule(formula=[f'ISNUMBER(SEARCH("ใกล้ถึง",$S{R4_START}))'],
    fill=PatternFill("solid", fgColor=WARN), stopIfTrue=True))
s4.conditional_formatting.add(S_RNG, FormulaRule(formula=[f'$S{R4_START}="ปกติ"'],
    fill=PatternFill("solid", fgColor=OKFILL), stopIfTrue=True))
s4.freeze_panes = "D5"

# ============================================================ 5_วิธีใช้
s5 = wb.create_sheet("5_วิธีใช้")
s5.sheet_view.showGridLines = False
s5.sheet_properties.tabColor = "808080"
s5.column_dimensions["A"].width = 30
s5.column_dimensions["B"].width = 112
s5.merge_cells("A1:B1")
put(s5, "A1", "วิธีใช้ไฟล์นี้", bold=True, size=14, color=C_WHT, fill=NAVY, align="left", border=False)
s5.row_dimensions[1].height = 26
HOW = [
    ("ชีต 1_สมมติฐานต้นทุน", "แก้ค่าแรงช่าง ค่ารถ วัสดุ และจำนวนเครื่องที่ทำได้ต่อวัน ให้ตรงกับทีมช่างจริง "
                             "แล้วเลือกฐานต้นทุนที่ใช้ตั้งราคา ทุกชีตที่เหลือดึงต้นทุนจากที่นี่"),
    ("ชีต 2_ราคาแพ็คเกจ", "ปรับราคาขายในคอลัมน์ B แล้วดู GM ในคอลัมน์ G | คอลัมน์ I = ราคาขั้นต่ำ (Floor) "
                          "คอลัมน์ J = ราคาเป้าหมาย (Target) | ด้านล่างมีตารางส่วนลดเดิม เพดานส่วนลด "
                          "และบันไดส่วนลดที่แนะนำ ให้เทียบกันได้"),
    ("ชีต 3_คำนวณใบเสนอราคา", "กรอกข้อมูลลูกค้ารายราย ระบบคำนวณรายได้ ต้นทุน GM ราคาขั้นต่ำ และส่วนลดสูงสุดที่ให้ได้"),
    ("ชีต 4_ทะเบียนลูกค้าและรอบPM", "บันทึกสัญญาและวันที่เข้าปฏิบัติงานจริง คอลัมน์สถานะจะเปลี่ยนสีเมื่อใกล้ถึงกำหนดหรือเกินกำหนด"),
    ("สีของเซลล์", "ช่องสีเหลือง + ตัวอักษรสีน้ำเงิน = ช่องที่ต้องกรอกเอง | ตัวอักษรสีดำ = สูตรคำนวณ ห้ามพิมพ์ทับ | "
                   "★ = รายการที่เพิ่ม/แก้ไขในเวอร์ชัน 2"),
    ("การเพิ่มบรรทัดในชีต 4", "คัดลอกบรรทัดสุดท้ายลงมา เพื่อให้สูตรในคอลัมน์ F, G, I, O, P, Q, R, S ติดไปด้วย"),
    ("ข้อควรระวัง", "ราคาตลาดอ้างอิงในชีต 1 เป็นข้อมูล ณ วันที่จัดทำ ควรตรวจสอบใหม่ก่อนใช้อ้างอิงกับลูกค้า"),
]
r = 3
for k, v in HOW:
    put(s5, f"A{r}", k, bold=True, size=10, wrap=True)
    put(s5, f"B{r}", v, size=10, wrap=True)
    s5.row_dimensions[r].height = max(18, 14 * (len(v) // 100 + 1))
    r += 1

# ============================================================ 0_สิ่งที่แก้ในเวอร์ชัน2
s0 = wb.create_sheet("0_สิ่งที่แก้ในเวอร์ชัน2")
s0.sheet_view.showGridLines = False
s0.sheet_properties.tabColor = "404040"
for c, w in zip("ABC", [30, 60, 58]):
    s0.column_dimensions[c].width = w
s0.merge_cells("A1:C1")
put(s0, "A1", "PM Pricing Tool  v2.1 — สรุปสิ่งที่แก้จากเวอร์ชัน 1", bold=True, size=15,
    color=C_WHT, fill=NAVY, align="center", border=False)
s0.row_dimensions[1].height = 30
s0.merge_cells("A2:C2")
put(s0, "A2", "จัดทำ 19-09-2026 | ไฟล์เดิมยังใช้งานได้ ไม่ถูกแก้ไข — ไฟล์นี้เป็นเวอร์ชันแยก",
    italic=True, size=9, color="595959", align="center", border=False, fill=GREY)
def s0band(r, t):
    s0.merge_cells(f"A{r}:C{r}")
    put(s0, f"A{r}", t, bold=True, size=12, color=C_WHT, fill=BAND, border=False)
    s0.row_dimensions[r].height = 22
    return r + 1
def s0row(r, a, b, c, color="000000"):
    put(s0, f"A{r}", a, bold=True, size=10, wrap=True)
    put(s0, f"B{r}", b, size=10, wrap=True, color=color)
    put(s0, f"C{r}", c, size=10, wrap=True, color="006100")
    s0.row_dimensions[r] = s0.row_dimensions[r]
    s0.row_dimensions[r].height = max(30, 13 * (max(len(b), len(c)) // 55 + 1))
    return r + 1
r = 4
r = s0band(r, "1) ข้อผิดพลาดที่แก้ / BUG FIXES")
put(s0, f"A{r}", "หัวข้อ", bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
put(s0, f"B{r}", "ปัญหาในเวอร์ชัน 1", bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
put(s0, f"C{r}", "วิธีแก้ในเวอร์ชัน 2", bold=True, color=C_WHT, fill=NAVY, align="center", size=9)
r += 1
FIXES = [
 ("ต้นทุนเครื่องอบไม่ถูกใช้",
  "สูตรต้นทุนแพ็คเกจ (ชีต 2 คอลัมน์ F) ใช้ 'ต้นทุนเฉลี่ยต่อเครื่อง' 325 บาทกับทั้งเครื่องซักและเครื่องอบ "
  "ทั้งที่ไฟล์เองระบุไว้ว่าเครื่องอบหนักกว่า 1.27 เท่า (413 บาท) → GM สูงเกินจริง 4-7 จุดทุกแพ็คเกจ",
  "แยกต้นทุนเครื่องซัก/เครื่องอบเป็นสองบรรทัด และสร้าง 'ต้นทุนต่อ 1 ชุดคู่ ต่อรอบ' (ชีต 1 แถว 33) "
  "ให้ทุกสูตรดึงไปใช้จุดเดียว"),
 ("GM ของดีลในชีต 3 สูงเกินจริง",
  "สูตร D32 นับทุกเครื่องเป็นเครื่องซัก และไม่นับต้นทุนเดินทางส่วนเกิน ทั้งที่เก็บเงินค่าเดินทางเป็นรายได้",
  "แยกบล็อกต้นทุนออกมา 6 บรรทัด ครบทั้งงาน PM เครื่องเดี่ยว เดินทาง ที่พัก งานซ่อมด่วน และท่อ Flex"),
 ("ขั้นส่วนลด 6 ชุดคู่ขึ้นไปขาดทุน",
  "ลด 25% พร้อมแถมรอบบริการฟรีอีก 1 รอบพร้อมกัน ทำให้ Care+ เหลือ 3,675 บาท เทียบต้นทุน 5 รอบ = 3,690 บาท "
  "และเวอร์ชัน 1 ไม่มีคอลัมน์ต้นทุน/GM ในตารางส่วนลดเลย จึงไม่มีอะไรเตือน",
  "เพิ่มคอลัมน์ GM ของทั้ง 3 แพ็คเกจในทุกขั้นส่วนลด พร้อมแถบสีแดงเมื่อติดลบ และคอลัมน์สถานะสรุปรายขั้น"),
 ("ค่าแรงซ่อมด่วนไม่มีต้นทุน",
  "Care Pro ให้ฟรีไม่จำกัด แต่ตั้งเผื่อไว้เพียง 1,000 บาทซึ่งระบุว่าเป็นค่าท่อ Flex | "
  "Care+ เก็บ 750 บาท/ครั้ง ซึ่งต่ำกว่าต้นทุนจริงต่อครั้ง",
  "เพิ่มสมมติฐาน 'ต้นทุนงานซ่อมด่วนต่อครั้ง' และ 'จำนวนครั้งเฉลี่ยต่อปี' แยกจากค่าท่อ Flex "
  "พร้อมตารางตรวจกำไรต่อครั้งของทั้ง 3 แพ็คเกจ (ชีต 2 แถว 13-15)"),
 ("ชีต 4 ไม่มีคอลัมน์สถานะ",
  "คู่มือในชีต 5 เขียนว่า 'คอลัมน์สถานะจะเตือนเมื่อใกล้ถึงกำหนดหรือเกินกำหนด' แต่ในไฟล์ไม่มีคอลัมน์นั้น "
  "และสูตรรอบคงเหลือไม่ได้นับรอบฟรีของลูกค้า 6 ชุดคู่ขึ้นไป",
  "เพิ่มคอลัมน์ รอบตามสัญญา / รอบ 5 (ฟรี) / วันคงเหลือ / สถานะ พร้อมแถบสีแดง-เหลือง-เขียวอัตโนมัติ"),
]
for a, b, c in FIXES:
    r = s0row(r, a, b, c, color="C00000")
r += 1
r = s0band(r, "2) ความสามารถที่เพิ่มเข้ามา / NEW")
for a, b, c in [
 ("ราคาขั้นต่ำอัตโนมัติ", "เวอร์ชัน 1 บอกได้แค่ผ่าน/ไม่ผ่านเกณฑ์",
  "ชีต 2 คอลัมน์ I และชีต 3 แถว 43 คำนวณราคาขั้นต่ำที่ยังได้ GM ตามเป้าให้ทันที (ปัดขึ้นทีละ 50/100 บาท)"),
 ("เพดานส่วนลดขณะต่อรอง", "ต้องลองปรับตัวเลขไปมาเอง",
  "ชีต 3 แถว 44 บอกเปอร์เซ็นต์ส่วนลดสูงสุดที่ให้ได้โดยไม่หลุดเป้า GM"),
 ("เลือกฐานต้นทุนได้", "ใช้ค่าเฉลี่ยของสองกรณีเสมอ",
  "ชีต 1 แถว 23 เลือกได้ว่าจะใช้ค่าเฉลี่ย / งานกระจุกจุดเดียว / งาน Route หลายจุด — "
  "ร้านสะดวกซักที่มีหลายเครื่องในจุดเดียวควรใช้ 'งานกระจุกจุดเดียว'"),
 ("ตรวจ GM ของดีล 3 ปี", "คำนวณแค่ราคาและส่วนที่ประหยัด",
  "ชีต 3 แถว 49 แสดง GM ของดีลชำระล่วงหน้า 3 ปี พร้อมเตือนเรื่องต้นทุนที่จะขึ้นระหว่างสัญญา"),
 ("เลือกแพ็คเกจด้วยชื่อ", "ต้องพิมพ์เลข 1/2/3",
  "ชีต 3 และชีต 4 เลือกจาก dropdown เป็นชื่อแพ็คเกจโดยตรง"),
]:
    r = s0row(r, a, b, c)
r += 1
r = s0band(r, "3) การปรับเป้า Gross Margin ให้ตรงกับตลาด (v2.1)")
for a, b, c in [
 ("เดิมใช้เป้าเดียว 50%",
  "บอกได้แค่ผ่าน/ไม่ผ่าน ใช้ต่อรองจริงไม่ได้ และไม่มีแพ็กเกจไหนผ่านเลยแม้ลูกค้ารายเดียว "
  "ทำให้ทุกดีลขึ้นสถานะแดงจนคนใช้เลิกสนใจสัญญาณเตือน",
  "แยกเป็น 2 ระดับ: GM ขั้นต่ำที่ยอมรับได้ (Floor) 35% และ GM เป้าหมาย (Target) 45% "
  "สถานะจึงมี 3 ระดับ — ผ่านเป้าหมาย / รับได้ / ต่ำกว่าขั้นต่ำ"),
 ("ที่มาของตัวเลข 35-45%",
  "ต้องมีหลักฐานอ้างอิง ไม่ใช่ตั้งลอยๆ",
  "ราคาตลาด CODE CLEAN 600 บาท/เครื่อง/ครั้ง = 1,200 บาท/ชุดคู่/รอบ เทียบต้นทุนของเราต่อชุดคู่ "
  "ให้ GM ประมาณ 38% (ดูแถวล่างสุดของชีต 1) จึงตั้ง Floor ต่ำกว่าเล็กน้อยที่ 35% "
  "และ Target สูงกว่าตลาดพอสมควรที่ 45%"),
 ("เพิ่มราคา 2 ระดับทุกที่",
  "เวอร์ชัน 2.0 มีแค่ราคาขั้นต่ำเดียว",
  "ชีต 2 มีทั้ง 'ราคาขั้นต่ำ (Floor)' และ 'ราคาเป้าหมาย (Target)' | "
  "ชีต 3 บอกทั้งราคาที่ควรยื่นก่อนต่อรอง และเพดานล่างสุดที่ห้ามต่ำกว่า"),
 ("เพดานส่วนลดต่อแพ็กเกจ",
  "ตอบลูกค้าไม่ได้ทันทีว่าลดได้กี่เปอร์เซ็นต์",
  "ชีต 2 เพิ่มตาราง 'เพดานส่วนลดที่ยังผ่าน GM ขั้นต่ำ' ทั้งกรณีมีและไม่มีรอบบริการฟรี"),
 ("บันไดส่วนลดที่แนะนำ",
  "บันไดเดิม 10/20/25% + รอบฟรี ทำให้ขาดทุนตั้งแต่ขั้น 2 เป็นต้นไป",
  "เสนอบันไดใหม่ 0 / 5 / 10 / 12% และตัดรอบบริการฟรีออกทั้งหมด "
  "พร้อมตาราง GM ให้เทียบกับบันไดเดิมแบบเคียงข้างกัน"),
]:
    r = s0row(r, a, b, c)
r += 1
r = s0band(r, "4) สิ่งที่ยังต้องตัดสินใจ / TO DECIDE")
for a, b, c in [
 ("ขั้นส่วนลด 6 ชุดคู่ขึ้นไป", "ยังติดลบอยู่ตามตัวเลขปัจจุบัน",
  "ทางเลือก: (ก) ตัดรอบบริการฟรีออก (ข) ลดส่วนลดเหลือ 15% (ค) ขึ้นราคาป้าย Care+ "
  "— แก้ได้ที่ชีต 2 ตารางแถว 24-29"),
  ("ฐานต้นทุนที่ใช้กับลูกค้ารายใหญ่", "ตั้งไว้ที่ 'ค่าเฉลี่ย' ซึ่งเป็นค่ากลางระหว่างงานกระจุกกับงาน Route",
  "ร้านสะดวกซักที่มีเครื่อง 8-10 ตัวในจุดเดียวควรใช้ฐาน 'งานกระจุกจุดเดียว' (ต้นทุนชุดคู่ 681 แทน 738) "
  "ซึ่งทำให้ลดราคาได้มากขึ้นอย่างมีเหตุผล — ชีต 3 มีบรรทัดแนะนำให้อัตโนมัติตามจำนวนเครื่องที่กรอก"),
 ("จำนวนครั้งซ่อมด่วนเฉลี่ย", "ตั้งไว้ 2 ครั้ง/ปี/ชุดคู่ เป็นค่าประมาณ",
  "เก็บสถิติจริงจากงานที่ผ่านมา แล้วแก้ที่ชีต 1 แถว 20 — ตัวเลขนี้กระทบ GM ของ Care Pro มากที่สุด"),
]:
    r = s0row(r, a, b, c)

# ---------------------------------------------------------------- order & save
ORDER = ["0_สิ่งที่แก้ในเวอร์ชัน2", "1_สมมติฐานต้นทุน", "2_ราคาแพ็คเกจ", "3_คำนวณใบเสนอราคา",
         "4_ทะเบียนลูกค้าและรอบPM", "5_วิธีใช้", "Z_LISTS"]
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
for n in ORDER:
    sh = wb[n]
    sh.page_setup.orientation = "landscape" if n in ("4_ทะเบียนลูกค้าและรอบPM", "2_ราคาแพ็คเกจ") else "portrait"
    sh.page_setup.fitToWidth = 1
    sh.page_setup.fitToHeight = 0
    sh.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
OUT = "4_PM_Pricing_Tool_v2.1.xlsx"
wb.save(OUT)
print("saved:", OUT)
