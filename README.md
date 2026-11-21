# Goodfield Printing

Static site (plain HTML, CSS, JavaScript, no build step): portfolio, price list, order form, staff print log, RU/EN, light/dark theme.

## Files (upload ALL of them to the GitHub repository)
- `index.html` - the whole site
- `manifest.webmanifest`, `icon-180.png`, `icon-192.png`, `icon-512.png` - phone home-screen icon and full-screen mode
- `og-image.jpg` - preview picture when the link is shared in messengers

## Where to edit
- Prices (TJS): top of the `<script>` in index.html (`PRICE`, `CARD_PER_SHEET`, `CUT_A5`, `BINDING`)
- `SCRIPT_URL`: Google Apps Script Web app URL (orders -> email + sheet "Orders", print log -> sheet "Print Log")
- Translations: the `ROWS` list in the script (English, Russian)

## Publish (free)
GitHub Pages: upload the files to the repository root and commit.

## Google Apps Script
`apps-script.gs` is the full script (paste it into Extensions -> Apps Script of the Google Sheet). Prices are edited in the "Prices" tab; order status in the "Orders" tab; staff report needs the REPORT_CODE script property.
