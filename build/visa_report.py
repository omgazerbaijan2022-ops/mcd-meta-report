"""Visa x McDonald's (October 2026) Meta report.

data/visa_report.xlsx + data/visa_template.pptx -> visa/McDonalds_AZ_Visa_Report_Oct_2026.pptx (+ .pdf)

Run from build/:  python3 visa_report.py
The PDF is produced with LibreOffice (soffice). For the exact brand font
(Speedee), open the .pptx in PowerPoint and use File > Export > PDF.
"""
import copy
import os
import subprocess
import sys

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA = os.path.join(ROOT, "data", "visa_report.xlsx")
TEMPLATE = os.path.join(ROOT, "data", "visa_template.pptx")
CREATIVE = os.path.join(ROOT, "assets", "visa_post_creative.jpg")
OUT_DIR = os.path.join(ROOT, "visa")
OUT = os.path.join(OUT_DIR, "McDonalds_AZ_Visa_Report_Oct_2026.pptx")

# Brand colours (template theme). Chart series colours validated with the
# dataviz palette checker (#DB0007 / #E59A00 pass CVD + lightness; labels shown on every bar).
RED = RGBColor(0xDB, 0x00, 0x07)
GOLD = RGBColor(0xE5, 0x9A, 0x00)
YELLOW = RGBColor(0xFF, 0xBC, 0x0D)
INK = RGBColor(0x29, 0x29, 0x29)
MUTED = RGBColor(0x53, 0x56, 0x5A)
LIGHT = RGBColor(0xF4, 0xF1, 0xEC)
LINE = RGBColor(0xD9, 0xD9, 0xD6)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

L_COVER, L_CONTENT, L_CLOSE = 0, 45, 64


# ---------------------------------------------------------------- data
def load():
    df = pd.read_excel(DATA)
    df = df.set_index("Ad name")
    num = lambda v: 0.0 if (v == "" or pd.isna(v)) else float(v)
    rows = {}
    for name, r in df.iterrows():
        rows[name] = {k: (num(v) if k not in ("Result type", "Reporting starts", "Reporting ends", "Results (initial)") else v)
                      for k, v in r.items()}
    return rows


def money(v, d=2):
    return f"${v:,.{d}f}"


def n(v):
    return f"{v:,.0f}"


def pct(v, d=1):
    return f"{v * 100:.{d}f}%"


# ---------------------------------------------------------------- helpers
def remove_all_slides(prs):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)


def text(slide, x, y, w, h, runs, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.0):
    """runs: str, or list of paragraphs; each paragraph a str or list of (text, opts) tuples."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, 0)
    paras = runs if isinstance(runs, list) else [runs]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        for t, o in ([(para, {})] if isinstance(para, str) else para):
            r = p.add_run()
            r.text = t
            f = r.font
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.color.rgb = o.get("color", color)
        if isinstance(para, list) and para and "space_after" in para[0][1]:
            p.space_after = Pt(para[0][1]["space_after"])
    return tb


def rect(slide, x, y, w, h, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s


def content_slide(prs, title, notes=None):
    s = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    s.shapes.title.text = title
    for ph in list(s.placeholders):
        if ph.placeholder_format.idx == 1:
            ph._element.getparent().remove(ph._element)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def footnote(slide, t):
    text(slide, 0.37, 6.55, 12.0, 0.3, t, size=9, color=MUTED)


def style_chart_text(chart, size=11):
    chart.font.size = Pt(size)
    chart.font.color.rgb = INK


# ---------------------------------------------------------------- slides
def build():
    d = load()
    p1, p2 = d["visa_02.10"], d["visa_20.10"]
    K = ["Amount spent (USD)", "Impressions", "Reach", "Post engagements", "Clicks (all)", "Link clicks",
         "Post reactions", "Post comments", "Post shares", "Post saves"]
    tot = {k: p1[k] + p2[k] for k in K}
    # Meta de-duplicates reach across ads; the export's total row shows 1,267,297 (not the sum).
    tot["Reach (dedup)"] = 1_267_297
    tot["CPM"] = tot["Amount spent (USD)"] / tot["Impressions"] * 1000
    tot["CPE"] = tot["Amount spent (USD)"] / tot["Post engagements"]
    tot["Frequency"] = tot["Impressions"] / tot["Reach (dedup)"]
    er = lambda r: r["Post engagements"] / r["Impressions"]
    ctr = lambda r: r["Clicks (all)"] / r["Impressions"]

    prs = Presentation(TEMPLATE)
    remove_all_slides(prs)

    # 1. Cover ---------------------------------------------------------
    s = prs.slides.add_slide(prs.slide_layouts[L_COVER])
    s.shapes.title.text = "Visa × McDonald's"
    sub = s.placeholders[1]
    sub.text_frame.text = "NFC ödənişlərə 10% endirim kampaniyası"
    pp = sub.text_frame.add_paragraph()
    pp.text = "Meta reklam hesabatı · 1–31 oktyabr 2026"

    # 2. Campaign overview ---------------------------------------------
    s = content_slide(prs, "Kampaniya haqqında")
    items = [
        ("Təklif", "Visa kartı ilə Apple Pay və ya Google Pay vasitəsilə 20 AZN-dən yuxarı sifarişlərə 10% endirim"),
        ("Müddət", "1–31 oktyabr 2026"),
        ("Filiallar", "8 ticarət mərkəzi: Gənclik Mall, Park Bulvar, 28 Mall, Dəniz Mall, Crescent Mall, "
                      "Metropark, Sevinc Mall, Gəncə Mall"),
        ("Paylaşım", "mcdonalds.az × visaaz birgə Instagram postu (1 oktyabr)"),
        ("Reklam", "Meta (Facebook + Instagram) üzərində 2 reklam: visa_02.10 – əhatə (Reach) kampaniyası, "
                   "visa_20.10 – qarşılıqlı əlaqə (Engagement) kampaniyası"),
    ]
    y = 1.45
    for label, body in items:
        rect(s, 0.37, y + 0.05, 0.08, 0.72, YELLOW)
        text(s, 0.6, y, 7.3, 0.3, label.upper(), size=11, bold=True, color=RED)
        text(s, 0.6, y + 0.3, 7.3, 0.6, body, size=15, color=INK)
        y += 1.0
    # creative (screenshot of the post)
    h = 5.3
    w = h * 802 / 1004
    pic = s.shapes.add_picture(CREATIVE, Inches(12.95 - w), Inches(1.3), Inches(w), Inches(h))
    pic.line.color.rgb = LINE
    pic.line.width = Pt(0.75)

    # 3. KPI summary ---------------------------------------------------
    s = content_slide(prs, "Əsas göstəricilər")
    tiles = [
        ("2", "Visa paylaşımı reklam edildi", True),
        (money(tot["Amount spent (USD)"]), "Ümumi xərc", False),
        (n(tot["Impressions"]), "Göstərilmə (Impressions)", False),
        (n(tot["Reach (dedup)"]), "Unikal əhatə (Reach)", False),
        (n(tot["Post engagements"]), "Qarşılıqlı əlaqə (Engagements)", False),
        (n(tot["Clicks (all)"]), "Kliklər (hamısı)", False),
    ]
    tw, th, gx, gy = 4.0, 1.75, 0.3, 0.3
    for i, (val, lab, hero) in enumerate(tiles):
        x = 0.37 + (i % 3) * (tw + gx)
        y = 1.4 + (i // 3) * (th + gy)
        rect(s, x, y, tw, th, YELLOW if hero else LIGHT)
        text(s, x + 0.3, y + 0.25, tw - 0.6, 0.85, val, size=40, bold=True, color=INK)
        text(s, x + 0.3, y + 1.15, tw - 0.6, 0.4, lab, size=14, color=INK)
    y = 1.4 + 2 * (th + gy) + 0.05
    small = [
        (money(tot["CPM"]), "CPM (1000 göstərilmənin dəyəri)"),
        (money(tot["CPE"]), "1 qarşılıqlı əlaqənin dəyəri (CPE)"),
        (f"{tot['Frequency']:.2f}", "Tezlik (bir nəfərə orta göstərilmə)"),
    ]
    for i, (val, lab) in enumerate(small):
        x = 0.37 + i * (tw + gx)
        text(s, x, y, tw, 0.5, [[(val + "  ", {"size": 22, "bold": True}), (lab, {"size": 13, "color": MUTED})]])
    footnote(s, "Mənbə: Meta Ads Manager, 1–31.10.2026, adında “visa” olan reklamlar. "
                "Əhatə Meta tərəfindən təkrarsız hesablanır, ona görə iki reklamın cəmindən (1,286,603) azdır.")

    # 4. Per-post table -----------------------------------------------
    s = content_slide(prs, "Paylaşımlar üzrə statistika")
    rows = [
        ("Kampaniya məqsədi", "Əhatə (Reach)", "Qarşılıqlı əlaqə", "—"),
        ("Xərc", money(p1["Amount spent (USD)"]), money(p2["Amount spent (USD)"]), money(tot["Amount spent (USD)"])),
        ("Göstərilmə", n(p1["Impressions"]), n(p2["Impressions"]), n(tot["Impressions"])),
        ("Əhatə", n(p1["Reach"]), n(p2["Reach"]), n(tot["Reach (dedup)"]) + "*"),
        ("Tezlik", f"{p1['Frequency']:.2f}", f"{p2['Frequency']:.2f}", f"{tot['Frequency']:.2f}"),
        ("CPM", money(p1["CPM (cost per 1,000 impressions)"]), money(p2["CPM (cost per 1,000 impressions)"]),
         money(tot["CPM"])),
        ("Nəticənin dəyəri", money(p1["Cost per result"]) + " / 1000 nəfər",
         money(p2["Cost per result"]) + " / əlaqə", "—"),
        ("Qarşılıqlı əlaqə", n(p1["Post engagements"]), n(p2["Post engagements"]), n(tot["Post engagements"])),
        ("1 əlaqənin dəyəri", money(p1["Cost per post engagement"]), money(p2["Cost per post engagement"]),
         money(tot["CPE"])),
        ("Engagement rate", pct(er(p1), 2), pct(er(p2), 2), pct(tot["Post engagements"] / tot["Impressions"], 2)),
        ("Reaksiyalar", n(p1["Post reactions"]), n(p2["Post reactions"]), n(tot["Post reactions"])),
        ("Şərhlər", n(p1["Post comments"]), n(p2["Post comments"]), n(tot["Post comments"])),
        ("Paylaşmalar", n(p1["Post shares"]), n(p2["Post shares"]), n(tot["Post shares"])),
        ("Yadda saxlamalar", n(p1["Post saves"]), n(p2["Post saves"]), n(tot["Post saves"])),
        ("Kliklər (hamısı)", n(p1["Clicks (all)"]), n(p2["Clicks (all)"]), n(tot["Clicks (all)"])),
        ("Link klikləri", n(p1["Link clicks"]), "—", n(tot["Link clicks"])),
    ]
    header = ("Göstərici", "visa_02.10", "visa_20.10", "Cəmi")
    nr, nc = len(rows) + 1, 4
    tbl_h = 5.05
    gt = s.shapes.add_table(nr, nc, Inches(0.37), Inches(1.3), Inches(12.6), Inches(tbl_h))
    tbl = gt.table
    tblPr = tbl._tbl.tblPr
    # drop the template's table style banding so our fills are the only fills
    tblPr.set("bandRow", "0")
    tblPr.set("firstRow", "0")
    widths = [3.6, 3.0, 3.0, 3.0]
    for i, wv in enumerate(widths):
        tbl.columns[i].width = Inches(wv)
    for r in range(nr):
        tbl.rows[r].height = Inches(tbl_h / nr)
        vals = header if r == 0 else rows[r - 1]
        for c in range(nc):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = RED if c in (1, 2) else INK
            elif c == 3:
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xF1, 0xCC)
            else:
                cell.fill.fore_color.rgb = LIGHT if r % 2 == 0 else WHITE
            tf = cell.text_frame
            tf.text = vals[c]
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            f = para.runs[0].font
            f.size = Pt(12)
            f.bold = (r == 0) or (c == 3)
            f.color.rgb = WHITE if r == 0 else INK
    footnote(s, "* Ümumi əhatə təkrarsızdır (eyni insan iki reklamı görübsə, bir dəfə sayılır). "
                "Engagement rate = qarşılıqlı əlaqə ÷ göstərilmə.")

    # 5. Share of budget vs share of results -------------------------
    reach_sum = p1["Reach"] + p2["Reach"]
    cats = ["Büdcə", "Göstərilmə", "Əhatə", "Kliklər", "Qarşılıqlı əlaqə"]
    a = [p1["Amount spent (USD)"] / tot["Amount spent (USD)"], p1["Impressions"] / tot["Impressions"],
         p1["Reach"] / reach_sum, p1["Clicks (all)"] / tot["Clicks (all)"],
         p1["Post engagements"] / tot["Post engagements"]]
    b = [1 - v for v in a]
    s = content_slide(
        prs, "Hər reklam öz işini gördü",
        notes="Payların hesablanması: hər göstəricidə reklamın payı = reklamın dəyəri ÷ iki reklamın cəmi. "
              "Əhatə payı iki reklamın əhatəsinin cəmindən (1,286,603) hesablanıb.")
    cd = CategoryChartData()
    cd.categories = list(reversed(cats))
    cd.add_series("visa_02.10 (Reach)", list(reversed(a)))
    cd.add_series("visa_20.10 (Engagement)", list(reversed(b)))
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, Inches(0.37), Inches(1.3), Inches(8.4), Inches(5.1), cd)
    ch = gf.chart
    style_chart_text(ch, 12)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(12)
    plot = ch.plots[0]
    plot.gap_width = 55
    plot.overlap = 100
    for ser, col in zip(plot.series, (RED, GOLD)):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = col
        ser.format.line.color.rgb = WHITE
        ser.format.line.width = Pt(1.5)
    for ser in plot.series:
        dl = ser.data_labels
        dl.show_value = True
        dl.number_format = "0%"
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.CENTER
        dl.font.size = Pt(12)
        dl.font.bold = True
        dl.font.color.rgb = WHITE
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    ca = ch.category_axis
    ca.format.line.fill.background()
    ca.tick_labels.font.size = Pt(13)
    # hide labels on slivers too thin to hold them (< 6%): per-point <c:dLbl><c:delete/>
    from pptx.oxml.ns import qn
    from lxml import etree
    for si, vals in enumerate((list(reversed(a)), list(reversed(b)))):
        ser = plot.series[si]._element
        dLbls = ser.find(qn("c:dLbls"))
        for pi, v in enumerate(vals):
            if v < 0.06:
                d = etree.SubElement(dLbls, qn("c:dLbl"))
                etree.SubElement(d, qn("c:idx")).set("val", str(pi))
                etree.SubElement(d, qn("c:delete")).set("val", "1")
                dLbls.remove(d)
                dLbls.insert(0, d)
    text(s, 9.15, 1.5, 3.8, 4.8, [
        [("visa_02.10", {"bold": True, "color": RED, "size": 16})],
        [(f"Büdcənin {pct(a[0], 0)}-i ilə göstərilmələrin {pct(a[1], 0)}-ni və əhatənin {pct(a[2], 0)}-ni verdi. "
          "Məqsəd geniş auditoriyaya təklifi çatdırmaq idi.", {"size": 14, "space_after": 18})],
        [("visa_20.10", {"bold": True, "color": GOLD, "size": 16})],
        [(f"Büdcənin cəmi {pct(b[0], 0)}-i ilə bütün qarşılıqlı əlaqələrin {pct(b[4], 0)}-ni topladı "
          f"({n(p2['Post engagements'])} əlaqə, onlardan {n(p2['Post reactions'])} reaksiya).", {"size": 14})],
    ], spacing=1.1)

    # 6. Efficiency: two small charts (different scales -> separate) --
    s = content_slide(prs, "Səmərəlilik: CPM və CPE")

    def mini(x, title, v1, v2, sub):
        text(s, x, 1.35, 6.0, 0.4, title, size=16, bold=True)
        text(s, x, 1.75, 6.0, 0.35, sub, size=12, color=MUTED)
        cd = CategoryChartData()
        cd.categories = ["visa_02.10\n(Reach)", "visa_20.10\n(Engagement)"]
        cd.add_series(title, [v1, v2])
        gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(2.15), Inches(5.9), Inches(3.55), cd)
        ch = gf.chart
        style_chart_text(ch, 12)
        ch.has_title = False
        ch.has_legend = False
        pl = ch.plots[0]
        pl.gap_width = 90
        for i, col in enumerate((RED, GOLD)):
            pt = pl.series[0].points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = col
        pl.has_data_labels = True
        pl.data_labels.number_format = '"$"0.00'
        pl.data_labels.number_format_is_linked = False
        pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        pl.data_labels.font.size = Pt(14)
        pl.data_labels.font.bold = True
        va = ch.value_axis
        va.visible = False
        va.has_major_gridlines = False
        va.minimum_scale = 0
        va.maximum_scale = max(v1, v2) * 1.2
        ch.category_axis.format.line.color.rgb = LINE
        ch.category_axis.tick_labels.font.size = Pt(12)

    cpm1, cpm2 = p1["CPM (cost per 1,000 impressions)"], p2["CPM (cost per 1,000 impressions)"]
    cpe1, cpe2 = p1["Cost per post engagement"], p2["Cost per post engagement"]
    mini(0.37, "CPM – 1000 göstərilmənin dəyəri", cpm1, cpm2, "Aşağı = auditoriyaya çatmaq daha ucuzdur")
    mini(6.95, "CPE – 1 qarşılıqlı əlaqənin dəyəri", cpe1, cpe2, "Aşağı = qarşılıqlı əlaqə daha ucuzdur")
    text(s, 0.37, 5.85, 12.6, 0.6, [[
        ("Nəticə: ", {"bold": True}),
        (f"Reach reklamı auditoriyaya çatmağı {cpm2 / cpm1:.0f} dəfə ucuz etdi ({money(cpm1)} vs {money(cpm2)}), "
         f"Engagement reklamı isə qarşılıqlı əlaqəni {cpe1 / cpe2:.0f} dəfə ucuz etdi ({money(cpe2)} vs {money(cpe1)}). "
         "Bu, iki reklamın fərqli məqsədlərlə qurulmasının düzgün olduğunu göstərir.", {}),
    ]], size=13)

    # 7. Conclusions --------------------------------------------------
    s = content_slide(prs, "Nəticələr və tövsiyələr")
    rect(s, 0.37, 1.35, 6.1, 5.1, LIGHT)
    rect(s, 6.87, 1.35, 6.1, 5.1, LIGHT)
    rect(s, 0.37, 1.35, 6.1, 0.08, RED)
    rect(s, 6.87, 1.35, 6.1, 0.08, YELLOW)
    concl = [
        f"Oktyabrda Visa təklifi üçün 2 paylaşım reklam edildi, ümumi xərc {money(tot['Amount spent (USD)'])}.",
        f"Kampaniya {tot['Reach (dedup)'] / 1e6:.2f} mln unikal istifadəçiyə çatdı və "
        f"{tot['Impressions'] / 1e6:.2f} mln göstərilmə qazandı (CPM {money(tot['CPM'])}).",
        f"Ümumilikdə {n(tot['Post engagements'])} qarşılıqlı əlaqə: {n(tot['Post reactions'])} reaksiya, "
        f"{n(tot['Post shares'])} paylaşma, {n(tot['Post saves'])} yadda saxlama, {n(tot['Post comments'])} şərh.",
        "Reach + Engagement birləşməsi effektiv işlədi: biri geniş əhatə, digəri ucuz qarşılıqlı əlaqə verdi.",
    ]
    recs = [
        "Növbəti partnyor kampaniyalarında eyni iki mərhələli modeli saxlamaq: əvvəl geniş əhatə, sonra xatırlatma/engagement.",
        "Təklif yalnız 8 ticarət mərkəzində keçərli olduğundan, həmin ünvanlar ətrafında lokasiya hədəfləməsini sınamaq.",
        "Kampaniyanın satışa təsirini ölçmək üçün Visa tərəfindən endirimli NFC tranzaksiyalarının sayını almaq "
        "və reklam göstəriciləri ilə müqayisə etmək.",
    ]

    def bullets(x, head, items, col):
        paras = [[(head, {"bold": True, "size": 18, "color": col, "space_after": 10})]]
        for i, it in enumerate(items, 1):
            paras.append([(f"{i}.  ", {"bold": True, "space_after": 14}), (it, {})])
        text(s, x + 0.3, 1.7, 5.5, 4.6, paras, size=16, spacing=1.1)

    bullets(0.37, "Nəticələr", concl, RED)
    bullets(6.87, "Tövsiyələr", recs, INK)

    # 8. Close ----------------------------------------------------------
    s = prs.slides.add_slide(prs.slide_layouts[L_CLOSE])
    s.shapes.title.text = "Təşəkkür edirik!"

    os.makedirs(OUT_DIR, exist_ok=True)
    prs.save(OUT)
    print("wrote", OUT)
    if "--no-pdf" not in sys.argv:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", OUT_DIR, OUT],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300)
        print("wrote", OUT[:-5] + ".pdf")


if __name__ == "__main__":
    build()
