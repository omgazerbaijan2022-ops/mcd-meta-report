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


# ---------------------------------------------------------------- metric pages (red McDonald's style)
WATERMARK = os.path.join(ROOT, "assets", "arch_watermark.png")   # template arches, tinted a shade off the red
L_BLANK = 61
SOFT = RGBColor(0xFF, 0xD6, 0xD7)        # secondary text on red


def load_youtube():
    return pd.read_csv(YT_DATA).iloc[0].to_dict()


def line(slide, x1, y, x2, color, width=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


X0, GW = 0.65, 12.03          # grid left edge and width (4 columns)
COLW = GW / 4


def metric_page(prs, platform, subtitle, rows, source):
    """rows: list of (groups, cells) where groups = [(tag, n_cols)] and cells = [(title, desc, value)] x4."""
    s = prs.slides.add_slide(prs.slide_layouts[L_BLANK])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RED
    # faint golden-arches watermark, bottom right, bleeding off the page
    s.shapes.add_picture(WATERMARK, Inches(7.4), Inches(3.2), Inches(6.6), Inches(6.6 * 2101 / 2401))

    # header: platform ——— period
    text(s, X0, 0.42, 3.2, 0.45, platform, size=22, bold=True, color=WHITE)
    text(s, X0, 0.9, 6.0, 0.3, subtitle, size=11, color=SOFT)
    line(s, X0 + 3.0, 0.66, 9.9, WHITE, 0.75)
    text(s, 10.1, 0.42, 2.58, 0.45, "Oktyabr 2026", size=22, bold=True, color=YELLOW, align=PP_ALIGN.RIGHT)

    y = 1.55
    for groups, cells in rows:
        cw = GW / len(cells)                  # 3 wide headline columns or 4 detail columns
        big = len(cells) == 3
        col = 0
        for tag, span in groups:
            gx = X0 + col * cw
            text(s, gx, y, span * cw - 0.3, 0.25, tag.upper(), size=10, bold=True, color=YELLOW)
            line(s, gx, y + 0.32, gx + span * cw - 0.3, YELLOW, 1.0)
            col += span
        for i, (title, desc, val) in enumerate(cells):
            cx = X0 + i * cw
            text(s, cx, y + 0.48, cw - 0.3, 0.3, title, size=14 if big else 13, bold=True, color=WHITE)
            text(s, cx, y + 0.78, cw - 0.3, 0.3, desc, size=9.5, color=SOFT)
            text(s, cx, y + 1.02, cw - 0.3, 0.7, val, size=38 if big else 32, bold=True, color=YELLOW)
        y += 1.85
    text(s, X0, 7.02, 11.0, 0.25, source, size=8.5, color=SOFT)
    return s


# ---------------------------------------------------------------- slides
# Short format: cover, one metric page per platform, close. Metrics follow the client's list:
# spend / reach / impressions, + frequency & CPM for reach ads,
# + engagements, cost per engagement, reactions, comments, shares, saves for engagement ads.
def build():
    d = load()
    p1, p2 = d["visa_02.10"], d["visa_20.10"]
    yt = load_youtube()
    yt["Frequency"] = yt["Impressions"] / yt["Unique users"]
    yt["CPM"] = yt["Cost (USD)"] / yt["Impressions"] * 1000
    meta_spend = p1["Amount spent (USD)"] + p2["Amount spent (USD)"]
    meta_impr = p1["Impressions"] + p2["Impressions"]
    META_REACH = 1_267_297   # de-duplicated across both ads (total row of the Meta export)

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
    metric_page(prs, "Meta Ads", "Facebook + Instagram · Visa kampaniyası", [
        ([("Ümumi nəticə · 2 reklam", 3)], [
            ("Spend", "Reklamlara xərclənən büdcə", money(meta_spend)),
            ("Reach", "Reklamı görən unikal insanlar (təkrarsız)", n(META_REACH)),
            ("Impressions", "Reklamın göstərilmə sayı", n(meta_impr)),
        ]),
        ([("Reach reklamı · visa_02.10", 2), ("Engagement reklamı · visa_20.10", 2)], [
            ("Frequency", "Bir nəfərə orta göstərilmə", f"{p1['Frequency']:.2f}"),
            ("CPM", "1000 göstərilmənin dəyəri", money(p1["CPM (cost per 1,000 impressions)"])),
            ("Post engagements", "Posta edilən bütün əlaqələr", n(p2["Post engagements"])),
            ("Cost per engagement", "1 engagement-in dəyəri", money(p2["Cost per post engagement"])),
        ]),
        ([("Engagement reklamı · posta reaksiyalar", 4)], [
            ("Reactions", "Like və digər reaksiyalar", n(p2["Post reactions"])),
            ("Comments", "Şərhlər", n(p2["Post comments"])),
            ("Shares", "Paylaşmalar", n(p2["Post shares"])),
            ("Saves", "Yadda saxlamalar", n(p2["Post saves"])),
        ]),
    ], "Mənbə: Meta Ads Manager, 1–31.10.2026. Reach iki reklam üzrə təkrarsız hesablanıb.")

    # 3. YouTube -------------------------------------------------------
    metric_page(prs, "YouTube Ads", "Google Ads · video kampaniya Visa_Reach_Oct'26 (25 oktyabrdan)", [
        ([("Ümumi nəticə", 3)], [
            ("Spend", f"Reklama xərclənən büdcə (limit {money(yt['Budget (USD)'])})", money(yt["Cost (USD)"])),
            ("Reach", "Reklamı görən unikal istifadəçilər", n(yt["Unique users"])),
            ("Impressions", "Reklamın göstərilmə sayı", n(yt["Impressions"])),
        ]),
        ([("Reach reklamı · Visa_Reach_Oct'26", 2), ("Video", 2)], [
            ("Frequency", "Bir istifadəçiyə orta göstərilmə", f"{yt['Frequency']:.2f}"),
            ("CPM", "1000 göstərilmənin dəyəri", money(yt["CPM"])),
            ("Views", "Video baxışları (TrueView)", n(yt["TrueView views"])),
            ("CPV", "1 baxışın dəyəri", money(yt["TrueView avg. CPV (USD)"], 3)),
        ]),
        ([("YouTube Shorts", 4)], [
            ("Likes", "Bəyənmələr", n(yt["Shorts likes"])),
            ("Comments", "Şərhlər", n(yt["Shorts comments"])),
            ("Shares", "Paylaşmalar", n(yt["Shorts shares"])),
            ("Clicks", "Məqsəd reach idi, klik hədəflənməyib", n(yt["Clicks"])),
        ]),
    ], "Mənbə: Google Ads, 1.10–4.11.2026.")

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
