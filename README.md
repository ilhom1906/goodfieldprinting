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

## Orders with several parts and payments
- An order is a list of parts (cover, inside pages, ...). Each part has its own paper (paper / card stock), colour, sides and page count; binding, cut to A5 and number of copies apply to the whole order.
- Payments mode (staff code): customer orders can be marked paid (full or partial); staff debts (print log, grouped by "Work by") can be marked paid per month or per record.
- Sheet columns are found by header name. New columns are added automatically: Orders -> Payment, Paid TJS, Paid at; Print Log -> Payment, Paid at, ID.

## Files, comments and totals
- Customers can attach up to 3 files (500 MB each, sent in 5 MB pieces); they are saved in the Drive folder "Goodfield Orders" and the links go to the Orders sheet (Files) and the email. The comment goes to the Comment column.
- Orders and Print Log end with a TOTAL row (SUM formulas) that the script keeps at the bottom after every new row.
