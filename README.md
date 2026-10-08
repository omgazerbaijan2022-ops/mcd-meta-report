# McDonald's Azerbaijan: Meta and TikTok performance reviews (Q2 vs Q3 2026)

A client-ready comparison of Meta Ads performance:

- **BEFORE:** 1 Apr – 30 Jun 2026 (previous agency)
- **AFTER:** 1 Jul – 30 Sep 2026 (current agency)

## Deliverables

### Meta

| File | What it is |
|---|---|
| `McDonalds_AZ_Meta_Performance_Review_Q2_vs_Q3_2026.pptx` | 16-slide editable deck with native charts and speaker notes giving the exact figures |
| `Analysis_Summary.md` | The 10 most important findings, plus data caveats |
| `analysis/analysis_tables.xlsx` | Audit trail: KPI summary and breakdowns by objective, result type, campaign, ad set, ad and initiative |
| `analysis/metrics.json` | Every number the deck uses |

### TikTok (`tiktok/`)

| File | What it is |
|---|---|
| `tiktok/McDonalds_AZ_TikTok_Performance_Review_Q2_vs_Q3_2026.pptx` | 16-slide editable deck (15 slides plus an appendix) with native charts and speaker notes |
| `tiktok/McDonalds_AZ_TikTok_Analysis_Q2_vs_Q3_2026.xlsx` | Analysis workbook with live formulas: KPI summary, objectives, like-for-like initiatives, campaigns, ad groups, ad-level data, validation against the export totals, and excluded rows |
| `tiktok/TikTok_Analysis_Summary.md` | The 9 most important findings, plus data caveats |
| `analysis/tiktok/metrics.json` | Every number the TikTok deck uses |

## Rebuild

```bash
cd build
npm install            # pptxgenjs
python3 analysis.py    # data/*.xlsx -> analysis/metrics.json + analysis_tables.xlsx
node deck.js           # -> ../McDonalds_AZ_Meta_Performance_Review_Q2_vs_Q3_2026.pptx

python3 tiktok_analysis.py   # data/tiktok-*.xlsx -> analysis/tiktok/metrics.json
node tiktok_deck.js          # -> ../tiktok/McDonalds_AZ_TikTok_Performance_Review_Q2_vs_Q3_2026.pptx
python3 tiktok_workbook.py   # -> ../tiktok/McDonalds_AZ_TikTok_Analysis_Q2_vs_Q3_2026.xlsx
```

The workbook is written with formulas only. Open it in Excel, or recalculate it with LibreOffice, to see values.

Requires Python 3 with pandas and openpyxl, and Node 18 or later.

## Brand assets

The deck uses only official assets supplied by the client. Add these files to `assets/` and rebuild, and they are placed automatically:

- `assets/mcd_logo.png`: the official Golden Arches logo. It goes on the cover and replaces the small arch motif in the top-right corner of every content slide. Both decks use it.
- `assets/product_hero.jpg` (or `.png`): an official product image for the cover panel.
- `assets/decor_arches.png` (optional): a decorative abstract visual for the cover, used only if no product image is supplied.

Until these files are added, the cover shows clearly marked logo placeholders and a vector arch illustration.
