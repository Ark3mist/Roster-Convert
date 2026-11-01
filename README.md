# Roster Convert

Upload a roster as a spreadsheet (`.xlsx`, `.xls`, `.csv`), the Roster Report **PDF** or saved web page (**`.html`**), or the roster calendar **`.ics`** file, and download a clean PDF. Everything runs in the browser, so the file is never uploaded anywhere.

## What it does

- **Crew roster exports** (the "Roster → Download Excel" file, with `Date`, `Pairing/Activity`, `Duty Report`, `Item`, `Dep Stn`, `Arr. Stn`, `Flying Hrs`, `Duty Hrs` and similar columns) become a **daily roster**:
  - the whole roster month is listed (1st to last day), including blank days, even if the export starts later or ends earlier
  - a flight that lands after midnight keeps its arrival on the same line, marked `(+1)`. The day after shows `LO` if you are still away.
  - empty days at an outstation between two flights are marked `LO` (layover)
  - any other day with no code is marked `AD` (available day), with the last station from the previous duty (e.g. `AD  KUL`). Available days are counted in the statistics.
  - day-off and leave codes (`DO`, `DO1`, `D`, `AL`, `MC1`…) are highlighted and shown as the code only, without the 00:00–23:59 times. The station is shown only when the day off is away from base. Weekend dates are coloured.
  - colour is used for the text only: duty is blue, off/leave green, layover purple, weekend dates orange. The one exception is the **report time**, which is set in bold black on a yellow highlighter mark so it can't be missed.
  - public-holiday dates are pink. The built-in Malaysia national and KL list for 2026–2027 can be edited in the app
  - block hours, duty hours, sectors, duty days, and days off at base or away are totalled
  - the crew name and details are taken from the `NAME | ID | FLEET | BASE | RANK` cell at the top of the file
  - the home base is the most common departure station
- **Roster Report PDF:** the text is read with pdf.js, including rotated pages and wrapped cells, and rebuilt into the same table as the Excel export, so the result is identical.
- **Roster Report web page (.html)**, saved from the crew portal: the table is read as plain data (no scripts run), with colspan/rowspan and multi-page tables handled.
- **The file's clock is always auto-detected.** The note under the upload box says which clock it found:
  - **Station local:** on flights between different time zones, arrival minus departure differs from the block time by exactly the zone difference (e.g. MH727 KUL 22:35 → CGK 23:45 is 1:10 for a 2:10 block, because CGK is UTC+7).
  - **UTC:** days off run 16:00–15:59 (for KUL) instead of 00:00–23:59. The portal's web page uses UTC.
  - **Home base:** otherwise, as in the Excel and PDF Roster Report. A "(LT)" header is only used as a tie-breaker.
  A note on the page says what was detected. Days off are always shown as home-base calendar days.
- **Roster calendar (.ics):** duty periods ("Report time …"), flights ("MH 020 KUL - CDG") and all-day leave/off events are turned into the same table. Calendar files have no aircraft, rank, work type, duty code or crew name. Block hours are calculated from the scheduled flight times, so augmented sectors show full flight time, not credited hours.
- **Any other spreadsheet** is converted to a PDF table as-is (you can pick the sheet).
- **Time menu (Station local / UTC / Home base KUL (UTC+8))** chooses the clock the roster is shown in:
  - **Station local (LT):** each time in that airport's own local time. Report times use the departure station and debrief times the arrival station.
  - **UTC:** one universal clock for every station.
  - **Home base:** every time in home-base time (KUL, always UTC+8).

  Daylight saving is handled automatically. Duties move to the date they fall on in the chosen clock. Off and leave days stay on their date. If a station isn't in the built-in list, the app warns you and uses the home base's zone.
- Exports: **Download PDF** (A4 or Letter, portrait, landscape or auto), **Print**, **CSV**. The PDF and Print shrink to fit **one page**: text gets smaller step by step, but never below a readable size. A compact block at the bottom puts Statistics (in two columns), **Code** (code descriptions) and **PH** (public holidays) side by side. Print opens that same one-page PDF.
- **PDF footer:** `Generated <date> / <name> (<device type · browser>) ~by LipasTerbang`. Browsers can't read a device's own name, so the device type is worked out locally from the browser.
- Downloaded files are named **`Name, Staff number, Month Year`**, e.g. `Shafiq Fakhri Othman, 2106537, November 2026.pdf`. Exports often cut names short, so type the full name in the title box once. It is remembered for that staff number.
- **Themes:** Default, Black & White, Dark, Minecraft, Roblox, **F1** (Titillium Web, black/red with timing-screen purple, green and yellow) and **MotoGP** (Barlow Condensed, asphalt grey with a MotoGP-red rule and flag colours). They apply to the PDF, Print and the on-screen roster. Minecraft uses the pixel font *Pixelify Sans* and Roblox uses *Montserrat*, both open-licensed (SIL OFL) look-alikes, since the official game fonts can't be redistributed. These fonts load only when their theme is chosen. The F1, MotoGP, Minecraft and Roblox themes also show an **original pixel-art badge** before the name (chequered flag, helmet, grass block, blocky smiley head). The official logos are trademarks and are not used.
- **Row lines** checkbox: a thin line under every row (on by default), or no lines with extra space between days. It applies to the PDF, Print and the screen, and is remembered.
- Columns can be hidden. **Duty Code** is hidden by default until you tick it. **Hotel** is opt-in: it stays hidden for everyone, including users who saved column choices before it was added, until it is ticked. It shows on the arrival flight and layover days, and long hotel names wrap so the table never runs off the page. Calendar (.ics) files have no hotel, so the option is greyed out for them. The code legend can be edited. All of these are remembered on your device.

## Privacy

- The privacy notice under the upload box pulses gently until a roster is loaded. It doesn't pulse for users who have reduced motion switched on.
- The roster file is read and converted **in the browser only**. It is never uploaded, sent or stored on a server.
- A strict Content-Security-Policy (`connect-src 'none'`, `form-action 'none'`, `default-src 'self'`, `style-src 'self'` with no `'unsafe-inline'`) makes the browser block any attempt by the page to send data, and any inline script or style. All fonts and libraries are served from this site; there are no third-party requests. The importer strips a saved web page's own CSS before reading it, because only the cell text is used.
- The roster is held in memory only and cleared when the page is refreshed or closed.
- Only settings are kept in the browser's `localStorage`: hidden columns, code descriptions, holiday edits, time mode, page size, and names typed per staff number. **Clear saved settings** on the page wipes them.

## Run it

It is a static site with no build step:

```bash
python3 -m http.server 8000
```

Then open http://localhost:8000.

### Publish on GitHub Pages

When releasing a change, run:

```bash
./release.sh
```

It sets the next version (today's date plus a counter, e.g. `2026.11.01-28`) in `APP_VERSION` in `importers.js` and in every `?v=` asset URL in `index.html` and `style.css`. It then checks that every reference matches and that no local asset is missing a `?v=`. Browsers then load the new files instead of cached ones. The version is shown in the page footer.

- `./release.sh 2026.12.01-01` sets an exact version.
- `./release.sh --check` only verifies, and exits with an error if anything is out of step.


Repo **Settings → Pages → Deploy from a branch**, then choose the branch and `/ (root)`.

## Files

| File | Purpose |
|---|---|
| `index.html` | Page layout |
| `style.css` | Dark theme, plus light styles for printing |
| `importers.js` | PDF, saved web page (HTML) and iCalendar readers that rebuild the Roster Report table |
| `app.js` | Parsing, rendering, PDF/CSV export |
| `release.sh` | Bumps and verifies the version on every asset URL |
| `vendor/fonts/` | DM Mono, Pixelify Sans, Montserrat, Titillium Web and Barlow Condensed (all SIL OFL): woff2 for the page, base64 TTF for the PDF |
| `vendor/` | SheetJS 0.18.5, jsPDF 2.5.1, jspdf-autotable 3.8.2, pdf.js 3.11.174 (loaded only when a PDF is opened) (bundled here so the app works offline and does not depend on a CDN) |
