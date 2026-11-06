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
L_BLANK = 61
SOFT = RGBColor(0xFF, 0xD6, 0xD7)        # secondary text on red
WATERMARK = os.path.join(ROOT, "assets", "arch_watermark.png")   # template arches, a shade off the red
YT_COVER = os.path.join(ROOT, "assets", "visa_youtube_cover.jpg")  # Shorts cover (i.ytimg.com oardefault)
YT_URL = "https://www.youtube.com/shorts/pDVNL1nRB1A"
YT_VIDEO = os.path.join(ROOT, "assets", "visa_youtube_video.mp4")  # the Shorts ad itself, embedded in the .pptx


def load_youtube():
    return pd.read_csv(YT_DATA).iloc[0].to_dict()


def line(slide, x1, y, x2, color, width=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y), Inches(x2), Inches(y))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    return c


X0 = 0.65                     # grid left edge
FULL_W = 12.03                # grid width without a visual
VIS_W, VIS_H, VIS_Y = 4.0, 5.0, 1.55   # right-hand visual slot (creative / video cover)
VIS_X = X0 + FULL_W - VIS_W


def metric_page(prs, platform, subtitle, rows, source, visual=None):
    """rows: list of (groups, cells): groups = [(tag, n_cols)], cells = [(title, desc, value)].
    visual: optional (image_path, aspect w/h, caption, url, video_path) shown on the right; the grid narrows.
    With a video, the .pptx plays it on the slide and the image is its poster (the PDF shows the poster)."""
    s = prs.slides.add_slide(prs.slide_layouts[L_BLANK])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RED

    # header: platform ——— period
    text(s, X0, 0.42, 3.2, 0.45, platform, size=22, bold=True, color=WHITE)
    text(s, X0, 0.9, 7.5, 0.3, subtitle, size=11, color=SOFT)
    line(s, X0 + 3.0, 0.66, 9.9, WHITE, 0.75)
    text(s, 10.1, 0.42, 2.58, 0.45, "Oktyabr 2026", size=22, bold=True, color=YELLOW, align=PP_ALIGN.RIGHT)

    if visual:
        path, aspect, caption, url, video = visual
        w = VIS_H * aspect
        x = VIS_X + (VIS_W - w) / 2
        if video:
            pic = s.shapes.add_movie(video, Inches(x), Inches(VIS_Y), Inches(w), Inches(VIS_H),
                                     poster_frame_image=path, mime_type="video/mp4")
        else:
            pic = s.shapes.add_picture(path, Inches(x), Inches(VIS_Y), Inches(w), Inches(VIS_H))
        pic.line.color.rgb = YELLOW
        pic.line.width = Pt(2)
        cap = text(s, VIS_X, VIS_Y + VIS_H + 0.1, VIS_W, 0.3, caption, size=10, color=SOFT, align=PP_ALIGN.CENTER)
        if url:
            run = cap.text_frame.paragraphs[0].runs[0]
            run.hyperlink.address = url
            run.font.color.rgb = SOFT
            run.font.underline = True
        gw = VIS_X - 0.45 - X0
    else:
        # faint golden-arches watermark, bottom right, bleeding off the page
        s.shapes.add_picture(WATERMARK, Inches(7.4), Inches(3.2), Inches(6.6), Inches(6.6 * 2101 / 2401))
        gw = FULL_W
    compact = visual is not None

    y = 1.55
    for groups, cells in rows:
        cw = gw / len(cells)                  # 3 wide headline columns or 4 detail columns
        big = len(cells) == 3
        col = 0
        for tag, span in groups:
            gx = X0 + col * cw
            text(s, gx, y, span * cw - 0.25, 0.25, tag.upper(), size=10, bold=True, color=YELLOW)
            line(s, gx, y + 0.32, gx + span * cw - 0.25, YELLOW, 1.0)
            col += span
        for i, (title, desc, val) in enumerate(cells):
            cx = X0 + i * cw
            tsize = (13 if big else 12) if compact else (14 if big else 13)
            vsize = (32 if big else 26) if compact else (36 if big else 32)
            text(s, cx, y + 0.48, cw - 0.2, 0.3, title, size=tsize, bold=True, color=WHITE)
            text(s, cx, y + 0.78, cw - 0.2, 0.3, desc, size=9 if compact else 9.5, color=SOFT)
            text(s, cx, y + 1.04, cw - 0.2, 0.7, val, size=vsize, bold=True, color=YELLOW)
        y += 1.8 if compact else 1.85
    text(s, X0, 7.02, 11.0, 0.25, source, size=8.5, color=SOFT)
    return s


# ---------------------------------------------------------------- slides
# Short format: cover, overall, one metric page per platform, close. Metrics follow the client's list:
# spend / reach / impressions, + frequency & CPM for reach ads,
# + engagements, cost per engagement, reactions, comments, shares, saves for engagement ads.
def build():
    d = load()
    p1, p2 = d["visa_02.10"], d["visa_20.10"]
    yt = load_youtube()
    yt["Frequency"] = yt["Impressions"] / yt["Unique users"]
    yt["CPM"] = yt["Cost (USD)"] / yt["Impressions"] * 1000
    meta = {k: p1[k] + p2[k] for k in ("Amount spent (USD)", "Impressions", "Post engagements",
                                       "Post reactions", "Post comments", "Post shares", "Post saves")}
    meta["Reach"] = 1_267_297   # de-duplicated across both ads (total row of the Meta export)
    yt_eng = yt["Shorts likes"] + yt["Shorts comments"] + yt["Shorts shares"]
    tot = {
        "spend": meta["Amount spent (USD)"] + yt["Cost (USD)"],
        "reach": meta["Reach"] + yt["Unique users"],          # platforms cannot de-duplicate each other
        "impr": meta["Impressions"] + yt["Impressions"],
        "eng": meta["Post engagements"] + yt_eng,
        "likes": meta["Post reactions"] + yt["Shorts likes"],
        "comments": meta["Post comments"] + yt["Shorts comments"],
        "shares": meta["Post shares"] + yt["Shorts shares"],
        "saves": meta["Post saves"],
    }
    tot["freq"] = tot["impr"] / tot["reach"]
    tot["cpm"] = tot["spend"] / tot["impr"] * 1000
    tot["cpe"] = tot["spend"] / tot["eng"]

    prs = Presentation(TEMPLATE)
    remove_all_slides(prs)

    # 1. Cover ---------------------------------------------------------
    s = prs.slides.add_slide(prs.slide_layouts[L_COVER])
    s.shapes.title.text = "Visa × McDonald's"
    sub = s.placeholders[1]
    sub.text_frame.text = "NFC ödənişlərə 10% endirim kampaniyası"
    pp = sub.text_frame.add_paragraph()
    pp.text = "Meta + YouTube hesabatı · oktyabr 2026"

    # 2. Overall (numbers only) ------------------------------------------
    metric_page(prs, "Overall", "Meta + YouTube · Visa kampaniyasının ümumi nəticəsi", [
        ([("Ümumi nəticə · Meta + YouTube", 3)], [
            ("Spend", "Reklamlara xərclənən ümumi büdcə", money(tot["spend"])),
            ("Reach", "Reklamı görən istifadəçilər*", n(tot["reach"])),
            ("Impressions", "Reklamın ümumi göstərilmə sayı", n(tot["impr"])),
        ]),
        ([("Səmərəlilik", 4)], [
            ("Frequency", "Bir istifadəçiyə orta göstərilmə", f"{tot['freq']:.2f}"),
            ("CPM", "1000 göstərilmənin dəyəri", money(tot["cpm"])),
            ("Engagements", "Bütün əlaqələr**", n(tot["eng"])),
            ("Cost per engagement", "1 engagement-in dəyəri", money(tot["cpe"])),
        ]),
        ([("Reaksiyalar · Meta + YouTube", 4)], [
            ("Reactions / Likes", "Like və reaksiyalar", n(tot["likes"])),
            ("Comments", "Şərhlər", n(tot["comments"])),
            ("Shares", "Paylaşmalar", n(tot["shares"])),
            ("Saves", "Yadda saxlamalar (Meta)", n(tot["saves"])),
        ]),
    ], "* Meta və YouTube əhatəsinin cəmi (platformalar arası təkrarsız deyil).  ** Meta post engagements + "
       "YouTube Shorts like, şərh və paylaşmaları.  Mənbə: Meta Ads Manager, Google Ads.")

    # 3. Meta ----------------------------------------------------------
    metric_page(prs, "Meta Ads", "Facebook + Instagram · Visa kampaniyası", [
        ([("Ümumi nəticə · 2 reklam", 3)], [
            ("Spend", "Reklamlara xərclənən büdcə", money(meta["Amount spent (USD)"])),
            ("Reach", "Unikal insanlar (təkrarsız)", n(meta["Reach"])),
            ("Impressions", "Göstərilmə sayı", n(meta["Impressions"])),
        ]),
        ([("Reach · visa_02.10", 2), ("Engagement · visa_20.10", 2)], [
            ("Frequency", "Bir nəfərə orta göstərilmə", f"{p1['Frequency']:.2f}"),
            ("CPM", "1000 göstərilmənin dəyəri", money(p1["CPM (cost per 1,000 impressions)"])),
            ("Engagements", "Posta edilən əlaqələr", n(p2["Post engagements"])),
            ("Cost / engagement", "1 engagement-in dəyəri", money(p2["Cost per post engagement"])),
        ]),
        ([("Engagement reklamı · posta reaksiyalar", 4)], [
            ("Reactions", "Like və reaksiyalar", n(p2["Post reactions"])),
            ("Comments", "Şərhlər", n(p2["Post comments"])),
            ("Shares", "Paylaşmalar", n(p2["Post shares"])),
            ("Saves", "Yadda saxlamalar", n(p2["Post saves"])),
        ]),
    ], "Mənbə: Meta Ads Manager, 1–31.10.2026. Reach iki reklam üzrə təkrarsız hesablanıb.",
        visual=(CREATIVE, 802 / 1004, "Instagram post · mcdonalds.az × visaaz", None, None))

    # 4. YouTube -------------------------------------------------------
    metric_page(prs, "YouTube Ads", "Google Ads · Visa_Reach_Oct'26 (25 oktyabrdan)", [
        ([("Ümumi nəticə", 3)], [
            ("Spend", f"Xərclənən büdcə (limit {money(yt['Budget (USD)'])})", money(yt["Cost (USD)"])),
            ("Reach", "Unikal istifadəçilər", n(yt["Unique users"])),
            ("Impressions", "Göstərilmə sayı", n(yt["Impressions"])),
        ]),
        ([("Reach · Visa_Reach_Oct'26", 2), ("Video", 2)], [
            ("Frequency", "Bir nəfərə orta göstərilmə", f"{yt['Frequency']:.2f}"),
            ("CPM", "1000 göstərilmənin dəyəri", money(yt["CPM"])),
            ("Views", "Video baxışları", n(yt["TrueView views"])),
            ("CPV", "1 baxışın dəyəri", money(yt["TrueView avg. CPV (USD)"], 3)),
        ]),
        ([("YouTube Shorts", 4)], [
            ("Likes", "Bəyənmələr", n(yt["Shorts likes"])),
            ("Comments", "Şərhlər", n(yt["Shorts comments"])),
            ("Shares", "Paylaşmalar", n(yt["Shorts shares"])),
            ("Clicks", "Klik hədəflənməyib", n(yt["Clicks"])),
        ]),
    ], "Mənbə: Google Ads, 1.10–4.11.2026.",
        visual=(YT_COVER, 9 / 16, "YouTube Shorts · videoya bax", YT_URL, YT_VIDEO))

    # 5. Close ---------------------------------------------------------
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
