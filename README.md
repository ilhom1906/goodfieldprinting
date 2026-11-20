# Goodfield Printing

Static site (plain HTML, CSS, JavaScript, no build step): portfolio, price list, order form, staff print log, RU/EN.

## Files (upload ALL of them to the GitHub repository)
- `index.html` - the whole site
- `manifest.webmanifest`, `icon-180.png`, `icon-192.png`, `icon-512.png` - phone home-screen icon and full-screen mode

## Where to edit
- Prices (TJS): top of the `<script>` in index.html (`PRICE`, `CARD_PER_SHEET`, `CUT_A5`, `BINDING`)
- `SCRIPT_URL`: Google Apps Script Web app URL (orders -> email + sheet "Orders", print log -> sheet "Print Log")
- Translations: the `ROWS` list in the script (English, Russian)

## Publish (free)
GitHub Pages: upload the files to the repository root and commit.
