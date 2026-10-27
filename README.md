# Pesto Parcel Scanner

Phone web app for Pesto Sea Group: photograph a courier label (DHL, UPS, TNT, GLS, BRT, Amazon…), the phone reads the barcode and the text, fills in courier, tracking number, sender and recipient, and sends it to the office.

- Everything is read on the phone (ZXing barcodes + Tesseract.js text recognition, Italian + English). No paid AI.
- The office inbox is a Google Sheet behind a small Apps Script web app: see `office-sheet/SETUP-GOOGLE-SHEET.md`.
- Put the web app URL in `config.js`. The team code is never stored in this repo; each person types it once in the app.
- Install on a phone: open the site, then Share → Add to Home Screen.
