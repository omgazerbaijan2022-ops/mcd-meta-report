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


# ---------------------------------------------------------------- reusable blocks
SECTION_FILL = RGBColor(0xFF, 0xE3, 0x8F)
POST_URL = "https://www.instagram.com/p/Dd88Z4fNag3/"


def data_table(slide, header, rows, widths, x, top, row_h=0.3, size=11, head_fills=None):
    """rows: tuple of cells, or a str for a full-width section row. A cell may be (text, url)."""
    nr, nc = len(rows) + 1, len(header)
    tbl = slide.shapes.add_table(nr, nc, Inches(x), Inches(top), Inches(sum(widths)), Inches(row_h * nr)).table
    tblPr = tbl._tbl.tblPr
    # drop the template's table style banding so our fills are the only fills
    tblPr.set("bandRow", "0")
    tblPr.set("firstRow", "0")
    for i, wv in enumerate(widths):
        tbl.columns[i].width = Inches(wv)
    band = 0
    for r in range(nr):
        tbl.rows[r].height = Inches(row_h)
        vals = header if r == 0 else rows[r - 1]
        section = isinstance(vals, str)
        if section:
            tbl.cell(r, 0).merge(tbl.cell(r, nc - 1))
            vals = (vals,) + ("",) * (nc - 1)
            band = 0
        else:
            band += 1
        for c in range(nc):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = (head_fills or {}).get(c, INK)
            elif section:
                cell.fill.fore_color.rgb = SECTION_FILL
            else:
                cell.fill.fore_color.rgb = LIGHT if band % 2 == 0 else WHITE
            if section and c > 0:
                continue
            v = vals[c]
            url = None
            if isinstance(v, tuple):
                v, url = v
            tf = cell.text_frame
            tf.text = v
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if (c == 0 or section) else PP_ALIGN.RIGHT
            run = para.runs[0]
            f = run.font
            f.size = Pt(size - 1 if section else size)
            f.bold = (r == 0) or section
            f.color.rgb = WHITE if r == 0 else INK
            if url:
                run.hyperlink.address = url
                f.color.rgb = RGBColor(0x1F, 0x5F, 0xA8)
                f.underline = True


def insight(slide, x, y, w, runs):
    rect(slide, x, y, 0.08, 0.62, RED)
    text(slide, x + 0.22, y, w - 0.22, 0.65, [runs], size=13, anchor=MSO_ANCHOR.MIDDLE, spacing=1.05)


def load_youtube():
    return pd.read_csv(YT_DATA).iloc[0].to_dict()


# ---------------------------------------------------------------- slides
# Short format: cover, one page per platform, close. Rows follow the client's metric list:
# common rows for every ad, + Frequency/CPM for reach ads, + engagement rows for engagement ads.
RX, RW = 4.75, 8.22          # right-hand column (left column holds the creative / campaign card)
COMMON = "Hər iki reklam üçün"
REACH_X = "Reach reklamı üçün"
ENG_X = "Engagement reklamı üçün"


def build():
    d = load()
    p1, p2 = d["visa_02.10"], d["visa_20.10"]
    yt = load_youtube()
    yt["Frequency"] = yt["Impressions"] / yt["Unique users"]
    yt["CPM"] = yt["Cost (USD)"] / yt["Impressions"] * 1000
    dash = "—"
    # Pending from the client team: exact ad dates and the links of the 2nd Meta post and the YouTube video.
    TBC = "dəqiqləşdirilir"

    prs = Presentation(TEMPLATE)
    remove_all_slides(prs)

    # 1. Cover ---------------------------------------------------------
    s = prs.slides.add_slide(prs.slide_layouts[L_COVER])
    s.shapes.title.text = "Visa × McDonald's"
    sub = s.placeholders[1]
    sub.text_frame.text = "NFC ödənişlərə 10% endirim kampaniyası"
    pp = sub.text_frame.add_paragraph()
    pp.text = "Meta + YouTube hesabatı · oktyabr 2026"

    # 2. Meta ----------------------------------------------------------
    s = content_slide(prs, "Meta (Facebook + Instagram)")
    h = 5.15
    w = h * 802 / 1004
    pic = s.shapes.add_picture(CREATIVE, Inches(0.37), Inches(1.3), Inches(w), Inches(h))
    pic.line.color.rgb = LINE
    pic.line.width = Pt(0.75)
    rows = [
        COMMON,
        ("Postun linki", ("instagram.com/p/Dd88Z4fNag3", POST_URL), TBC),
        ("Reklam məqsədi", "Reach", "Engagement"),
        ("Reklam tarixləri", TBC, TBC),
        ("Xərc (Spend)", money(p1["Amount spent (USD)"]), money(p2["Amount spent (USD)"])),
        ("Əhatə (Reach)", n(p1["Reach"]), n(p2["Reach"])),
        ("Göstərilmə (Impressions)", n(p1["Impressions"]), n(p2["Impressions"])),
        REACH_X,
        ("Tezlik (Frequency)", f"{p1['Frequency']:.2f}", dash),
        ("CPM", money(p1["CPM (cost per 1,000 impressions)"]), dash),
        ENG_X,
        ("Post engagements", dash, n(p2["Post engagements"])),
        ("Cost per post engagement", dash, money(p2["Cost per post engagement"])),
        ("Reaksiyalar (Reactions)", dash, n(p2["Post reactions"])),
        ("Şərhlər (Comments)", dash, n(p2["Post comments"])),
        ("Paylaşmalar (Shares)", dash, n(p2["Post shares"])),
        ("Yadda saxlamalar (Saves)", dash, n(p2["Post saves"])),
    ]
    data_table(s, ("Göstərici", "visa_02.10 · Reach", "visa_20.10 · Engagement"), rows, [3.0, 2.6, 2.62],
               RX, 1.3, row_h=0.25, head_fills={1: RED, 2: RED})
    insight(s, RX, 5.86, RW, [
        ("Cəmi: ", {"bold": True}),
        (f"{money(p1['Amount spent (USD)'] + p2['Amount spent (USD)'])} xərc, "
         f"{n(p1['Impressions'] + p2['Impressions'])} göstərilmə, 1,267,297 təkrarsız əhatə.", {}),
    ])
    footnote(s, "Mənbə: Meta Ads Manager, 1–31.10.2026, adında “visa” olan reklamlar. "
                "Cəmi əhatə Meta tərəfindən təkrarsız hesablanır.")

    # 3. YouTube -------------------------------------------------------
    s = content_slide(prs, "YouTube (Google Ads)")
    rect(s, 0.37, 1.3, w, h, LIGHT)
    rect(s, 0.37, 1.3, w, 0.08, RED)
    card = [
        ("Kampaniya", "Visa_Reach_Oct'26"),
        ("Format", "Video reklam (YouTube + Shorts)"),
        ("Strategiya", "Efficient reach, Target CPM"),
        ("Büdcə", money(yt["Budget (USD)"])),
        ("Status", "Başa çatıb"),
    ]
    paras = []
    for k, v in card:
        paras.append([(k.upper(), {"size": 10, "bold": True, "color": RED})])
        paras.append([(v, {"size": 15, "space_after": 12})])
    text(s, 0.62, 1.6, w - 0.5, h - 0.5, paras, spacing=1.05)
    rows = [
        COMMON,
        ("Videonun linki", TBC),
        ("Reklam məqsədi", "Reach"),
        ("Reklam tarixləri", "25.10.2026 – " + TBC),
        ("Xərc (Spend)", money(yt["Cost (USD)"])),
        ("Əhatə (Unique users)", n(yt["Unique users"])),
        ("Göstərilmə (Impressions)", n(yt["Impressions"])),
        REACH_X,
        ("Tezlik (Frequency)", f"{yt['Frequency']:.2f}"),
        ("CPM", money(yt["CPM"])),
    ]
    data_table(s, ("Göstərici", "Visa_Reach_Oct'26 · Reach"), rows, [4.6, 3.62],
               RX, 1.3, row_h=0.4, head_fills={1: RED})
    insight(s, RX, 5.85, RW, [
        ("Nəticə: ", {"bold": True}),
        (f"{money(yt['Cost (USD)'])} xərclə {n(yt['Unique users'])} unikal istifadəçiyə çatıldı, "
         f"{n(yt['Impressions'])} göstərilmə, CPM {money(yt['CPM'])}.", {}),
    ])
    footnote(s, "Mənbə: Google Ads, 1.10–4.11.2026. Əhatə = Unique users (Google Ads).")

    # 4. Close ---------------------------------------------------------
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
