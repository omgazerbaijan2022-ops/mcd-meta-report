# McDonald's Azerbaijan: Meta performance review (Q2 vs Q3 2026)

A client-ready comparison of Meta Ads performance:

- **BEFORE:** 1 Apr – 30 Jun 2026 (previous agency)
- **AFTER:** 1 Jul – 30 Sep 2026 (current agency)

## Deliverables

| File | What it is |
|---|---|
| `McDonalds_AZ_Meta_Performance_Review_Q2_vs_Q3_2026.pptx` | 16-slide editable deck with native charts and speaker notes giving the exact figures |
| `Analysis_Summary.md` | The 10 most important findings, plus data caveats |
| `analysis/analysis_tables.xlsx` | Audit trail: KPI summary and breakdowns by objective, result type, campaign, ad set, ad and initiative |
| `analysis/metrics.json` | Every number the deck uses |

## Rebuild

```bash
cd build
npm install            # pptxgenjs
python3 analysis.py    # data/*.xlsx -> analysis/metrics.json + analysis_tables.xlsx
node deck.js           # -> ../McDonalds_AZ_Meta_Performance_Review_Q2_vs_Q3_2026.pptx
```

Requires Python 3 with pandas and openpyxl, and Node 18 or later.

## Brand assets

The deck uses only official assets supplied by the client. Add these files to `assets/` and rebuild, and they are placed automatically:

- `assets/mcd_logo.png`: the official Golden Arches logo. It goes on the cover and replaces the small arch motif in the top-right corner of every content slide.
- `assets/product_hero.jpg` (or `.png`): an official product image for the cover panel.
- `assets/decor_arches.png` (optional): a decorative abstract visual for the cover, used only if no product image is supplied.

Until these files are added, the cover shows clearly marked logo placeholders and a vector arch illustration.
