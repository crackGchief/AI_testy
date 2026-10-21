"""Сборка презентации «История автомобилестроения» (первые 2 слайда).
Стиль: технический/современный, металл, хром, ретро-фотографии."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

DARK   = RGBColor(0x14, 0x16, 0x1C)
CHROME = RGBColor(0xC9, 0xD2, 0xE0)
STEEL  = RGBColor(0x5A, 0x64, 0x78)
ACCENT = RGBColor(0xE8, 0x7A, 0x2B)   # «ретро-оранжевый» акцент
WHITE  = RGBColor(0xF2, 0xF4, 0xF8)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def bg(slide, color=DARK):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color

def bar(slide, x, y, w, h, color):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = color; shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def text(slide, x, y, w, h, s, size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT, font="Arial"):
    tb = slide.shapes.add_textbox(x, y, w, h); tf = tb.text_frame
    tf.word_wrap = True
    lines = s.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ln; p.alignment = align
        for r in p.runs:
            r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = font
    return tb

# ---------------- SLIDE 1: титульный ----------------
s1 = prs.slides.add_slide(BLANK); bg(s1)
bar(s1, 0, 0, prs.slide_width, Inches(0.12), CHROME)
bar(s1, 0, Inches(7.38), prs.slide_width, Inches(0.12), CHROME)
s1.shapes.add_picture("результаты_презентации/images/slide1_benz_1886.png",
                      Inches(6.7), Inches(1.35), height=Inches(4.6))
text(s1, Inches(0.6), Inches(1.1), Inches(6.0), Inches(1.6),
     "ИСТОРИЯ АВТОМОБИЛЕСТРОЕНИЯ", size=40, bold=True, color=CHROME)
text(s1, Inches(0.6), Inches(2.35), Inches(5.9), Inches(0.8),
     "от первых двигателей до электрокаров", size=26, color=ACCENT)
text(s1, Inches(0.6), Inches(3.4), Inches(5.8), Inches(2.6),
     "1886 — Карл Бенц запатентовал Benz Patent-Motorwagen Nr.1 — первый в мире автомобиль с ДВС\n"
     "Мощность 0,75 л.с. · 3 колеса · скорость 16 км/ч", size=17, color=WHITE)
text(s1, Inches(0.6), Inches(6.35), Inches(6.0), Inches(0.5),
     "Дисциплина: НИР / УИР · стиль: metal & chrome · ретро-фото", size=13, color=STEEL)

# ---------------- SLIDE 2: хронология этапов ----------------
s2 = prs.slides.add_slide(BLANK); bg(s2)
bar(s2, 0, 0, prs.slide_width, Inches(0.12), CHROME)
text(s2, Inches(0.6), Inches(0.35), Inches(12.0), Inches(0.8),
     "ОСНОВНЫЕ ЭТАПЫ РАЗВИТИЯ АВТОМОБИЛЬНОЙ ИНДУСТРИИ", size=28, bold=True, color=CHROME)
s2.shapes.add_picture("результаты_презентации/images/slide2_timeline_metal.png",
                      Inches(0.6), Inches(1.35), width=Inches(12.1))
cols = [
 ("1886–1900","Эра изобретения: Бенц, Даймлер; первые ДВС и ручная сборка."),
 ("1908–1960","Конвейер Форда; массовые заводы XX века; старт АвтоВАЗа (Lada, 1966)."),
 ("1966–2000","Японская экспансия Toyota; инжекторы, безопасность, экономичность."),
 ("2000–н.в.","Электрокары и автономные машины: Tesla, Toyota bZ, Lada e-Largus."),
]
x = Inches(0.6)
for i,(t_,d_) in enumerate(cols):
    cx = Emu(int(x) + i*int(Inches(3.1)))
    bar(s2, cx, Inches(5.6), Inches(2.9), Inches(0.06), ACCENT)
    text(s2, cx, Inches(5.7), Inches(2.9), Inches(0.4), t_, size=15, bold=True, color=CHROME)
    text(s2, cx, Inches(6.1), Inches(2.9), Inches(1.2), d_, size=12, color=WHITE)

out = "результаты_презентации/prezentaciya_istoriya_avtomobilestroyeniya.pptx"
prs.save(out)
print("saved:", out)
