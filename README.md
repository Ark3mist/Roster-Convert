# Roster Convert

Upload a roster as a spreadsheet (`.xlsx`, `.xls`, `.csv`), the Roster Report **PDF** or saved web page (**`.html`**), or the roster calendar **`.ics`** file, and download a clean PDF. Everything runs in the browser, so the file is never uploaded anywhere.

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
- **Roster Report PDF:** the text is read with pdf.js, including rotated pages and wrapped cells, and rebuilt into the same table as the Excel export, so the result is identical.
- **Roster Report web page (.html)**, saved from the crew portal: the table is read as plain data (no scripts run), with colspan/rowspan and multi-page tables handled.
- **File times: Auto-detect / Home base / UTC** lets you override the detection described in the next point, e.g. for a UTC file with no days off.
- **UTC exports are detected automatically.** Some versions (such as the portal's web page) give times in UTC without saying so. Days off then run 16:00–15:59 instead of 00:00–23:59 (for KUL), and the app converts the times accordingly. Days off are always shown as home-base calendar days.
- **Roster calendar (.ics):** duty periods ("Report time …"), flights ("MH 020 KUL - CDG") and all-day leave/off events are turned into the same table. Calendar files have no aircraft, rank, work type, duty code or crew name. Block hours are calculated from the scheduled flight times, so augmented sectors show full flight time, not credited hours.
- **Any other spreadsheet** is converted to a PDF table as-is (you can pick the sheet).
- **Times: Station local, UTC, or Home base.** The crew Roster Report gives every station's times in home-base time (London departures in KUL time, for example). The app converts them:
  - **Station local (LT):** each time in that airport's own local time. Report times use the departure station and debrief times the arrival station.
  - **UTC:** one universal clock for every station.
  - **Home base (as exported):** the times unchanged, as in the export.

  Daylight saving is handled automatically. Duties move to the date they fall on in the chosen clock. Off and leave days stay on their date. If a station isn't in the built-in list, the app warns you and uses the home base's zone.
- Exports: **Download PDF** (A4 or Letter, portrait, landscape or auto), **Print**, **CSV**. The PDF and Print shrink to fit **one page**: text gets smaller step by step, but never below a readable size. Statistics, code descriptions and public holidays sit side by side at the bottom. Print opens that same one-page PDF.
- Downloaded files are named **`Name, Staff number, Month Year`**, e.g. `Shafiq Fakhri Othman, 2106537, November 2026.pdf`. Exports often cut names short, so type the full name in the title box once. It is remembered for that staff number.
- Columns can be hidden. The code legend can be edited. Both are remembered on your device.

## Privacy

- The roster file is read and converted **in the browser only**. It is never uploaded, sent or stored on a server.
- A Content-Security-Policy (`connect-src 'none'`, `form-action 'none'`, `default-src 'self'`) makes the browser block any attempt by the page to send data. All fonts and libraries are served from this site; there are no third-party requests.
- The roster is held in memory only and cleared when the page is refreshed or closed.
- Only settings are kept in the browser's `localStorage`: hidden columns, code descriptions, holiday edits, time mode, page size, and names typed per staff number. **Clear saved settings** on the page wipes them.

## Run it

It is a static site with no build step:

```bash
python3 -m http.server 8000
```

Then open http://localhost:8000.

### Publish on GitHub Pages

When releasing a change, bump `APP_VERSION` in `importers.js` and the `?v=` on the asset URLs in `index.html` and `style.css`. Browsers then load the new files instead of cached ones. The version is shown in the page footer.


Repo **Settings → Pages → Deploy from a branch**, then choose the branch and `/ (root)`.

## Files

| File | Purpose |
|---|---|
| `index.html` | Page layout |
| `style.css` | Dark theme, plus light styles for printing |
| `importers.js` | PDF, saved web page (HTML) and iCalendar readers that rebuild the Roster Report table |
| `app.js` | Parsing, rendering, PDF/CSV export |
| `vendor/fonts/` | DM Mono font (SIL OFL): woff2 for the page, base64 TTF for the PDF |
| `vendor/` | SheetJS 0.18.5, jsPDF 2.5.1, jspdf-autotable 3.8.2, pdf.js 3.11.174 (loaded only when a PDF is opened) (bundled here so the app works offline and does not depend on a CDN) |
