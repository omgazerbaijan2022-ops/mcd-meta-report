"""Visa x McDonald's (October 2026) report: Meta + YouTube (Google Ads).

data/visa_report.xlsx (Meta export) + data/visa_youtube.csv (Google Ads campaign row)
+ data/visa_template.pptx -> visa/McDonalds_AZ_Visa_Report_Oct_2026.pptx (+ .pdf)

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
YT_DATA = os.path.join(ROOT, "data", "visa_youtube.csv")
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
TOTAL_FILL = RGBColor(0xFF, 0xF1, 0xCC)
# Platform colours (validated pair, distinct from the per-ad red/gold): Meta blue, YouTube brown.
META_C = RGBColor(0x2F, 0x7F, 0xB5)
YT_C = RGBColor(0xB0, 0x7A, 0x2A)

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


# ---------------------------------------------------------------- reusable blocks
def data_table(slide, header, rows, widths, top=1.3, height=5.05, size=12, head_fills=None, total_col=None):
    nr, nc = len(rows) + 1, len(header)
    gt = slide.shapes.add_table(nr, nc, Inches(0.37), Inches(top), Inches(sum(widths)), Inches(height))
    tbl = gt.table
    tblPr = tbl._tbl.tblPr
    # drop the template's table style banding so our fills are the only fills
    tblPr.set("bandRow", "0")
    tblPr.set("firstRow", "0")
    for i, wv in enumerate(widths):
        tbl.columns[i].width = Inches(wv)
    for r in range(nr):
        tbl.rows[r].height = Inches(height / nr)
        vals = header if r == 0 else rows[r - 1]
        for c in range(nc):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = (head_fills or {}).get(c, INK)
            elif c == total_col:
                cell.fill.fore_color.rgb = TOTAL_FILL
            else:
                cell.fill.fore_color.rgb = LIGHT if r % 2 == 0 else WHITE
            tf = cell.text_frame
            tf.text = vals[c]
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            f = para.runs[0].font
            f.size = Pt(size)
            f.bold = (r == 0) or (c == total_col)
            f.color.rgb = WHITE if r == 0 else INK


def share_chart(slide, cats, a, names, colors, x=0.37, y=1.3, w=8.4, h=5.1):
    """100% stacked horizontal bars: share of series A vs B for each category (a = A's share)."""
    from pptx.oxml.ns import qn
    from lxml import etree
    b = [1 - v for v in a]
    cd = CategoryChartData()
    cd.categories = list(reversed(cats))
    cd.add_series(names[0], list(reversed(a)))
    cd.add_series(names[1], list(reversed(b)))
    ch = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED_100, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    style_chart_text(ch, 12)
    ch.has_title = False
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(12)
    plot = ch.plots[0]
    plot.gap_width = 55
    plot.overlap = 100
    for ser, col in zip(plot.series, colors):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = col
        ser.format.line.color.rgb = WHITE
        ser.format.line.width = Pt(1.5)
        dl = ser.data_labels
        dl.show_value = True
        dl.number_format = "0%"
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.CENTER
        dl.font.size = Pt(12)
        dl.font.bold = True
        dl.font.color.rgb = WHITE
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.fill.background()
    ch.category_axis.tick_labels.font.size = Pt(13)
    # hide labels on slivers too thin to hold them (< 6%): per-point <c:dLbl><c:delete/>
    for si, vals in enumerate((list(reversed(a)), list(reversed(b)))):
        dLbls = plot.series[si]._element.find(qn("c:dLbls"))
        for pi, v in enumerate(vals):
            if v < 0.06:
                d = etree.Element(qn("c:dLbl"))
                etree.SubElement(d, qn("c:idx")).set("val", str(pi))
                etree.SubElement(d, qn("c:delete")).set("val", "1")
                dLbls.insert(0, d)
    return b


def tiles(slide, items, cols=3, top=1.4, tw=4.0, th=1.75, gx=0.3, gy=0.3, val_size=40):
    for i, (val, lab, hero) in enumerate(items):
        x = 0.37 + (i % cols) * (tw + gx)
        y = top + (i // cols) * (th + gy)
        rect(slide, x, y, tw, th, YELLOW if hero else LIGHT)
        text(slide, x + 0.3, y + 0.22, tw - 0.6, th * 0.5, val, size=val_size, bold=True, color=INK)
        text(slide, x + 0.3, y + th - 0.6, tw - 0.6, 0.4, lab, size=14, color=INK)


def stat_row(slide, items, y, tw=4.0, gx=0.3):
    for i, (val, lab) in enumerate(items):
        x = 0.37 + i * (tw + gx)
        text(slide, x, y, tw, 0.5, [[(val + "  ", {"size": 22, "bold": True}), (lab, {"size": 13, "color": MUTED})]])


def load_youtube():
    r = pd.read_csv(YT_DATA).iloc[0].to_dict()
    r["Engagements"] = r["Shorts likes"] + r["Shorts comments"] + r["Shorts shares"]
    return r


# ---------------------------------------------------------------- slides
def build():
    d = load()
    p1, p2 = d["visa_02.10"], d["visa_20.10"]
    yt = load_youtube()
    K = ["Amount spent (USD)", "Impressions", "Reach", "Post engagements", "Clicks (all)", "Link clicks",
         "Post reactions", "Post comments", "Post shares", "Post saves"]
    tot = {k: p1[k] + p2[k] for k in K}
    # Meta de-duplicates reach across ads; the export's total row shows 1,267,297 (not the sum).
    tot["Reach (dedup)"] = 1_267_297
    tot["CPM"] = tot["Amount spent (USD)"] / tot["Impressions"] * 1000
    tot["CPE"] = tot["Amount spent (USD)"] / tot["Post engagements"]
    tot["Frequency"] = tot["Impressions"] / tot["Reach (dedup)"]
    er = lambda r: r["Post engagements"] / r["Impressions"]

    yt["Frequency"] = yt["Impressions"] / yt["Unique users"]
    yt["CPM"] = yt["Cost (USD)"] / yt["Impressions"] * 1000
    yt["CPE"] = yt["Cost (USD)"] / yt["Engagements"]
    yt["CPV"] = yt["Cost (USD)"] / yt["TrueView views"]

    # both platforms (reach is summed: platforms cannot de-duplicate each other's audiences)
    all_ = {
        "spend": tot["Amount spent (USD)"] + yt["Cost (USD)"],
        "impr": tot["Impressions"] + yt["Impressions"],
        "reach": tot["Reach (dedup)"] + yt["Unique users"],
        "eng": tot["Post engagements"] + yt["Engagements"],
        "clicks": tot["Clicks (all)"] + yt["Clicks"],
        "likes": tot["Post reactions"] + yt["Shorts likes"],
        "comments": tot["Post comments"] + yt["Shorts comments"],
        "shares": tot["Post shares"] + yt["Shorts shares"],
    }
    all_["CPM"] = all_["spend"] / all_["impr"] * 1000
    all_["CPE"] = all_["spend"] / all_["eng"]

    prs = Presentation(TEMPLATE)
    remove_all_slides(prs)

    # 1. Cover ---------------------------------------------------------
    s = prs.slides.add_slide(prs.slide_layouts[L_COVER])
    s.shapes.title.text = "Visa × McDonald's"
    sub = s.placeholders[1]
    sub.text_frame.text = "NFC ödənişlərə 10% endirim kampaniyası"
    pp = sub.text_frame.add_paragraph()
    pp.text = "Meta + YouTube hesabatı · oktyabr 2026"

    # 2. Campaign overview ---------------------------------------------
    s = content_slide(prs, "Kampaniya haqqında")
    items = [
        ("Təklif", "Visa kartı ilə Apple Pay və ya Google Pay vasitəsilə 20 AZN-dən yuxarı sifarişlərə 10% endirim"),
        ("Müddət", "1–31 oktyabr 2026"),
        ("Filiallar", "8 ticarət mərkəzi: Gənclik Mall, Park Bulvar, 28 Mall, Dəniz Mall, Crescent Mall, "
                      "Metropark, Sevinc Mall, Gəncə Mall"),
        ("Meta (Facebook + Instagram)", "mcdonalds.az × visaaz birgə postu, 2 reklam: visa_02.10 – əhatə (Reach), "
                                        "visa_20.10 – qarşılıqlı əlaqə (Engagement)"),
        ("YouTube (Google Ads)", "Visa_Reach_Oct'26 – video kampaniya (Efficient reach, Target CPM), "
                                 "25 oktyabrdan başlayıb"),
    ]
    y = 1.45
    for label, body in items:
        rect(s, 0.37, y + 0.05, 0.08, 0.72, YELLOW)
        text(s, 0.6, y, 7.3, 0.3, label.upper(), size=11, bold=True, color=RED)
        text(s, 0.6, y + 0.3, 7.3, 0.6, body, size=15, color=INK)
        y += 1.0
    h = 5.3
    w = h * 802 / 1004
    pic = s.shapes.add_picture(CREATIVE, Inches(12.95 - w), Inches(1.3), Inches(w), Inches(h))
    pic.line.color.rgb = LINE
    pic.line.width = Pt(0.75)

    # 3. KPI summary (both platforms) ------------------------------------
    s = content_slide(prs, "Əsas göstəricilər: Meta + YouTube")
    tiles(s, [
        ("3", "Reklam: 2 Meta + 1 YouTube", True),
        (money(all_["spend"]), "Ümumi xərc", False),
        (n(all_["impr"]), "Göstərilmə (Impressions)", False),
        (n(all_["reach"]), "Əhatə (Reach / Unique users)*", False),
        (n(all_["eng"]), "Qarşılıqlı əlaqə**", False),
        (n(yt["TrueView views"]), "YouTube video baxışları", False),
    ])
    stat_row(s, [
        (money(all_["CPM"]), "CPM (1000 göstərilmənin dəyəri)"),
        (money(all_["CPE"]), "1 qarşılıqlı əlaqənin dəyəri"),
        (n(all_["clicks"]), "Kliklər (Meta)"),
    ], y=1.4 + 2 * 2.05 + 0.05)
    footnote(s, "* Platformalar arası təkrarsız deyil: eyni insan həm Meta-da, həm YouTube-da sayıla bilər. "
                "** Meta: post engagements; YouTube: Shorts bəyənmə + şərh + paylaşma. "
                "Mənbə: Meta Ads Manager (1–31.10.2026), Google Ads (1.10–4.11.2026).")

    # 4. Platform comparison table ---------------------------------------
    s = content_slide(prs, "Platformalar üzrə müqayisə")
    dash = "—"
    rows = [
        ("Reklam sayı", "2", "1", "3"),
        ("Kampaniya növü", "Reach + Engagement", "Video, Efficient reach", dash),
        ("Xərc", money(tot["Amount spent (USD)"]), money(yt["Cost (USD)"]), money(all_["spend"])),
        ("Göstərilmə", n(tot["Impressions"]), n(yt["Impressions"]), n(all_["impr"])),
        ("Əhatə / unikal istifadəçi", n(tot["Reach (dedup)"]), n(yt["Unique users"]), n(all_["reach"]) + "*"),
        ("Tezlik", f"{tot['Frequency']:.2f}", f"{yt['Frequency']:.2f}", dash),
        ("CPM", money(tot["CPM"]), money(yt["CPM"]), money(all_["CPM"])),
        ("Video baxışları (TrueView)", dash, n(yt["TrueView views"]), n(yt["TrueView views"])),
        ("1 baxışın dəyəri (CPV)", dash, money(yt["CPV"], 3), dash),
        ("Qarşılıqlı əlaqə**", n(tot["Post engagements"]), n(yt["Engagements"]), n(all_["eng"])),
        ("1 əlaqənin dəyəri", money(tot["CPE"]), money(yt["CPE"]), money(all_["CPE"])),
        ("Reaksiyalar / bəyənmələr", n(tot["Post reactions"]), n(yt["Shorts likes"]), n(all_["likes"])),
        ("Şərhlər", n(tot["Post comments"]), n(yt["Shorts comments"]), n(all_["comments"])),
        ("Paylaşmalar", n(tot["Post shares"]), n(yt["Shorts shares"]), n(all_["shares"])),
        ("Yadda saxlamalar", n(tot["Post saves"]), dash, n(tot["Post saves"])),
        ("Kliklər", n(tot["Clicks (all)"]), n(yt["Clicks"]), n(all_["clicks"])),
    ]
    data_table(s, ("Göstərici", "Meta", "YouTube", "Cəmi"), rows, [3.6, 3.0, 3.0, 3.0],
               head_fills={1: META_C, 2: YT_C}, total_col=3)
    footnote(s, "* Platformalar arası cəm (təkrarsız deyil). ** Meta: post engagements (reaksiya, şərh, paylaşma, "
                "yadda saxlama, klik və s.); YouTube: Shorts bəyənmə + şərh + paylaşma.")

    # 5. Platform shares chart --------------------------------------------
    s = content_slide(prs, "Büdcə və nəticələrin platformalar üzrə payı")
    cats = ["Büdcə", "Göstərilmə", "Əhatə", "Qarşılıqlı əlaqə", "Kliklər"]
    a = [tot["Amount spent (USD)"] / all_["spend"], tot["Impressions"] / all_["impr"],
         tot["Reach (dedup)"] / all_["reach"], tot["Post engagements"] / all_["eng"],
         tot["Clicks (all)"] / all_["clicks"]]
    b = share_chart(s, cats, a, ("Meta", "YouTube"), (META_C, YT_C))
    text(s, 9.15, 1.5, 3.8, 4.8, [
        [("Meta", {"bold": True, "color": META_C, "size": 16})],
        [(f"Büdcənin {pct(a[0], 0)}-i ilə göstərilmələrin {pct(a[1], 0)}-ni verdi; "
          f"bütün kliklər ({n(tot['Clicks (all)'])}) Meta-dan gəldi.", {"size": 14, "space_after": 18})],
        [("YouTube", {"bold": True, "color": YT_C, "size": 16})],
        [(f"Büdcənin {pct(b[0], 0)}-i ilə əhatənin {pct(b[2], 0)}-ni və qarşılıqlı əlaqələrin "
          f"{pct(b[3], 0)}-ni topladı ({n(yt['Shorts likes'])} bəyənmə, {n(yt['Shorts comments'])} şərh).",
          {"size": 14})],
    ], spacing=1.1)

    # 6. Meta per-post table ----------------------------------------------
    s = content_slide(prs, "Meta: paylaşımlar üzrə statistika")
    rows = [
        ("Kampaniya məqsədi", "Əhatə (Reach)", "Qarşılıqlı əlaqə", dash),
        ("Xərc", money(p1["Amount spent (USD)"]), money(p2["Amount spent (USD)"]), money(tot["Amount spent (USD)"])),
        ("Göstərilmə", n(p1["Impressions"]), n(p2["Impressions"]), n(tot["Impressions"])),
        ("Əhatə", n(p1["Reach"]), n(p2["Reach"]), n(tot["Reach (dedup)"]) + "*"),
        ("Tezlik", f"{p1['Frequency']:.2f}", f"{p2['Frequency']:.2f}", f"{tot['Frequency']:.2f}"),
        ("CPM", money(p1["CPM (cost per 1,000 impressions)"]), money(p2["CPM (cost per 1,000 impressions)"]),
         money(tot["CPM"])),
        ("Nəticənin dəyəri", money(p1["Cost per result"]) + " / 1000 nəfər",
         money(p2["Cost per result"]) + " / əlaqə", dash),
        ("Qarşılıqlı əlaqə", n(p1["Post engagements"]), n(p2["Post engagements"]), n(tot["Post engagements"])),
        ("1 əlaqənin dəyəri", money(p1["Cost per post engagement"]), money(p2["Cost per post engagement"]),
         money(tot["CPE"])),
        ("Engagement rate", pct(er(p1), 2), pct(er(p2), 2), pct(tot["Post engagements"] / tot["Impressions"], 2)),
        ("Reaksiyalar", n(p1["Post reactions"]), n(p2["Post reactions"]), n(tot["Post reactions"])),
        ("Şərhlər", n(p1["Post comments"]), n(p2["Post comments"]), n(tot["Post comments"])),
        ("Paylaşmalar", n(p1["Post shares"]), n(p2["Post shares"]), n(tot["Post shares"])),
        ("Yadda saxlamalar", n(p1["Post saves"]), n(p2["Post saves"]), n(tot["Post saves"])),
        ("Kliklər (hamısı)", n(p1["Clicks (all)"]), n(p2["Clicks (all)"]), n(tot["Clicks (all)"])),
        ("Link klikləri", n(p1["Link clicks"]), dash, n(tot["Link clicks"])),
    ]
    data_table(s, ("Göstərici", "visa_02.10", "visa_20.10", "Cəmi"), rows, [3.6, 3.0, 3.0, 3.0],
               head_fills={1: RED, 2: RED}, total_col=3)
    footnote(s, "* Ümumi əhatə təkrarsızdır (eyni insan iki reklamı görübsə, bir dəfə sayılır). "
                "Engagement rate = qarşılıqlı əlaqə ÷ göstərilmə.")

    # 7. Meta: share of budget vs results per ad ------------------------
    reach_sum = p1["Reach"] + p2["Reach"]
    a = [p1["Amount spent (USD)"] / tot["Amount spent (USD)"], p1["Impressions"] / tot["Impressions"],
         p1["Reach"] / reach_sum, p1["Clicks (all)"] / tot["Clicks (all)"],
         p1["Post engagements"] / tot["Post engagements"]]
    s = content_slide(
        prs, "Meta: hər reklam öz işini gördü",
        notes="Hər göstəricidə reklamın payı = reklamın dəyəri ÷ iki reklamın cəmi. "
              "Əhatə payı iki reklamın əhatəsinin cəmindən (1,286,603) hesablanıb.")
    b = share_chart(s, ["Büdcə", "Göstərilmə", "Əhatə", "Kliklər", "Qarşılıqlı əlaqə"], a,
                    ("visa_02.10 (Reach)", "visa_20.10 (Engagement)"), (RED, GOLD))
    text(s, 9.15, 1.5, 3.8, 4.8, [
        [("visa_02.10", {"bold": True, "color": RED, "size": 16})],
        [(f"Büdcənin {pct(a[0], 0)}-i ilə göstərilmələrin {pct(a[1], 0)}-ni və əhatənin {pct(a[2], 0)}-ni verdi. "
          "Məqsəd geniş auditoriyaya təklifi çatdırmaq idi.", {"size": 14, "space_after": 18})],
        [("visa_20.10", {"bold": True, "color": GOLD, "size": 16})],
        [(f"Büdcənin cəmi {pct(b[0], 0)}-i ilə bütün qarşılıqlı əlaqələrin {pct(b[4], 0)}-ni topladı "
          f"({n(p2['Post engagements'])} əlaqə, onlardan {n(p2['Post reactions'])} reaksiya).", {"size": 14})],
    ], spacing=1.1)

    # 8. Meta efficiency: two small charts (different scales -> separate) --
    s = content_slide(prs, "Meta: səmərəlilik (CPM və CPE)")

    def mini(x, title, v1, v2, sub):
        text(s, x, 1.35, 6.0, 0.4, title, size=16, bold=True)
        text(s, x, 1.75, 6.0, 0.35, sub, size=12, color=MUTED)
        cd = CategoryChartData()
        cd.categories = ["visa_02.10\n(Reach)", "visa_20.10\n(Engagement)"]
        cd.add_series(title, [v1, v2])
        ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(2.15), Inches(5.9), Inches(3.55),
                                cd).chart
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

    # 9. YouTube ---------------------------------------------------------
    s = content_slide(prs, "YouTube: Visa_Reach_Oct'26")
    tiles(s, [
        (money(yt["Cost (USD)"]), "Xərc (büdcə $1,400)", True),
        (n(yt["Impressions"]), "Göstərilmə", False),
        (n(yt["Unique users"]), "Unikal istifadəçi", False),
        (n(yt["TrueView views"]), "Video baxışları (TrueView)", False),
        (n(yt["Shorts likes"]), "Shorts bəyənmələri", False),
        (n(yt["Shorts comments"] + yt["Shorts shares"]),
         f"Shorts şərh ({n(yt['Shorts comments'])}) + paylaşma ({n(yt['Shorts shares'])})", False),
    ])
    stat_row(s, [
        (money(yt["CPM"]), "CPM (1000 göstərilmənin dəyəri)"),
        (money(yt["CPV"], 3), "1 baxışın dəyəri (CPV)"),
        (f"{yt['Frequency']:.2f}", "Tezlik (bir istifadəçiyə orta göstərilmə)"),
    ], y=1.4 + 2 * 2.05 + 0.05)
    footnote(s, "Video kampaniya, Efficient reach, Target CPM; 25 oktyabrdan başlayıb, status: başa çatıb. "
                "Məqsəd əhatə olduğundan kliklər hədəflənməyib (0 klik). Mənbə: Google Ads, 1.10–4.11.2026.")

    # 10. Conclusions -------------------------------------------------
    s = content_slide(prs, "Nəticələr və tövsiyələr")
    rect(s, 0.37, 1.35, 6.1, 5.1, LIGHT)
    rect(s, 6.87, 1.35, 6.1, 5.1, LIGHT)
    rect(s, 0.37, 1.35, 6.1, 0.08, RED)
    rect(s, 6.87, 1.35, 6.1, 0.08, YELLOW)
    concl = [
        f"Visa təklifi üçün 3 reklam yayımlandı (2 Meta, 1 YouTube), ümumi xərc {money(all_['spend'])}.",
        f"Kampaniya {all_['impr'] / 1e6:.2f} mln göstərilmə və {all_['reach'] / 1e6:.2f} mln əhatə qazandı "
        f"(CPM {money(all_['CPM'])}).",
        f"{n(all_['eng'])} qarşılıqlı əlaqə: {pct(yt['Engagements'] / all_['eng'], 0)}-i YouTube Shorts-dan "
        f"({n(yt['Shorts likes'])} bəyənmə, {n(yt['Shorts comments'])} şərh).",
        "Meta-da Reach + Engagement birləşməsi effektiv işlədi: biri geniş əhatə, digəri ucuz qarşılıqlı əlaqə verdi.",
    ]
    recs = [
        "Növbəti partnyor kampaniyalarında eyni modeli saxlamaq: geniş əhatə (Meta + YouTube), sonra engagement.",
        "Təklif yalnız 8 ticarət mərkəzində keçərli olduğundan, həmin ünvanlar ətrafında lokasiya hədəfləməsini sınamaq.",
        "Satışa təsiri ölçmək üçün Visa tərəfindən endirimli NFC tranzaksiyalarının sayını alıb "
        "reklam göstəriciləri ilə müqayisə etmək.",
    ]

    def bullets(x, head, items, col):
        paras = [[(head, {"bold": True, "size": 18, "color": col, "space_after": 10})]]
        for i, it in enumerate(items, 1):
            paras.append([(f"{i}.  ", {"bold": True, "space_after": 14}), (it, {})])
        text(s, x + 0.3, 1.7, 5.5, 4.6, paras, size=16, spacing=1.1)

    bullets(0.37, "Nəticələr", concl, RED)
    bullets(6.87, "Tövsiyələr", recs, INK)

    # 11. Close ---------------------------------------------------------
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
