# Scoreboard Data Extraction Utility
**Author:** Rish Kumaria  
**Contact:** rishabhpk@hotmail.com

## What this does

Takes the clinic's weekly performance spreadsheet (`Scoreboard Test.xlsx`) and turns it into clean, queryable JSON. The tricky part: this isn't a normal spreadsheet. It's got 7 rows of headers, merged cells everywhere, hidden columns, and metrics that share the same name across different sections. So I couldn't just pandas.read_excel() it and call it a day.

## Running it

```bash
pip install -r requirements.txt

# flat output — one row per observation, ready for Tableau/Power BI
python convert.py --input "Scoreboard Test.xlsx" --output "output.json" --tidy

# nested output — grouped by week, if you prefer that structure
python convert.py --input "Scoreboard Test.xlsx" --output "output.json"
```

## How the output is structured

I went with an unpivoted "tidy data" layout for the `--tidy` flag. Each record looks like:

```json
{ "date": "2026-02-16", "metric_id": "collection_ar_90_days", "value": 0.0, "formula": null }
```

The full JSON also ships a `metrics` dictionary with all the metadata (section, focus area, source system, responsible person, target values, cell comments, hyperlinks) so you can join on `metric_id` without duplicating that info across every row.

There's also a `rules` array — I pulled the conditional formatting logic out of Excel so a dashboard could recreate the red/green color coding without guessing at thresholds.

## Problems I ran into and how I solved them

**Name collisions.** "Utilization" appears in like 4 different sections (PT, OT, Chiro, Pelvic Health). Same with "PVA (4 wk avg)". If you just use the metric name as a key, you silently lose data. I fixed this by building IDs from the section + metric name, and if that still collides, I append the Excel column letter as a tiebreaker. Caught 11 collisions this way.

**The dual-pass thing.** openpyxl can give you computed values OR formulas, but not both at once. So I load the file twice — once with `data_only=True` for the numbers, once with `data_only=False` for the formulas, comments, and hyperlinks. It's a bit slow but it means nothing gets dropped.

**Merged cells.** Row 1 has section banners (like "PHONE PERFORMANCE") that span a bunch of columns. I wrote a resolver that walks the merged ranges and propagates the value to every column underneath.

**Orphan data.** Column 21 has actual numbers in it but no header in Row 2. Instead of silently skipping it, the script picks it up as "Unlabeled Metric (U)" so you can investigate it yourself.

**HYPERLINK formulas.** Two columns use `=HYPERLINK("url", "label")` instead of normal cell hyperlinks. I added a regex fallback to catch those too.

## What I'd do next

Build a small HTML page that acts as a data dictionary — loop through the metrics index and render each one with its definition, target, and source link. Would make it way easier for someone non-technical to understand what they're looking at.

---
*Rish Kumaria*
