# Roster Convert

Upload a roster spreadsheet (`.xlsx`, `.xls`, `.csv`) and download a clean PDF. Everything runs in the browser, so the file is never uploaded anywhere.

## What it does

- **Crew roster exports** (the "Roster → Download Excel" file, with `Date`, `Pairing/Activity`, `Duty Report`, `Item`, `Dep Stn`, `Arr. Stn`, `Flying Hrs`, `Duty Hrs` and similar columns) become a **daily roster**:
  - the whole roster month is listed (1st to last day), including blank days, even if the export starts later or ends earlier
  - flights that land after midnight are split onto the arrival day
  - empty days at an outstation between two flights are marked `LO` (layover)
  - day-off and leave codes (`DO`, `AL`, `EL`, `MCH`…) are highlighted, and weekend dates are coloured
  - colour is used for the text only, with no background highlights: duty is blue, off/leave green, layover purple, weekend dates orange
  - public-holiday dates are pink. The built-in Malaysia national and KL list for 2026–2027 can be edited in the app
  - block hours, duty hours, sectors, duty days, and days off at base or away are totalled
  - the crew name and details are taken from the `NAME | ID | FLEET | BASE | RANK` cell at the top of the file
  - the home base is the most common departure station
- **Any other spreadsheet** is converted to a PDF table as-is (you can pick the sheet).
- **Times: Local (LT) or UTC**. Times are converted per station using each airport's time zone, including daylight saving, and duties move to their UTC date. Off and leave days stay on their date. If a station isn't in the built-in list, the app warns you and uses the home base's zone.
- Exports: **Download PDF** (A4 or Letter, portrait, landscape or auto), **Print**, **CSV**. The PDF and Print shrink to fit **one page**: text gets smaller step by step, but never below a readable size. Statistics, code descriptions and public holidays sit side by side at the bottom. Print opens that same one-page PDF.
- Downloaded files are named **`Name, Staff number, Month Year`**, e.g. `Shafiq Fakhri Othman, 2106537, November 2026.pdf`. Exports often cut names short, so type the full name in the title box once. It is remembered for that staff number.
- Columns can be hidden. The code legend can be edited. Both are remembered on your device.

## Run it

It is a static site with no build step:

```bash
python3 -m http.server 8000
```

Then open http://localhost:8000.

### Publish on GitHub Pages

Repo **Settings → Pages → Deploy from a branch**, then choose the branch and `/ (root)`.

## Files

| File | Purpose |
|---|---|
| `index.html` | Page layout |
| `style.css` | Dark theme, plus light styles for printing |
| `app.js` | Parsing, rendering, PDF/CSV export |
| `vendor/` | SheetJS 0.18.5, jsPDF 2.5.1, jspdf-autotable 3.8.2 (bundled here so the app works offline and does not depend on a CDN) |
