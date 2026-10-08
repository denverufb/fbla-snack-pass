# FBLA Snack Pass

A student snack-point POS system using **GitHub Pages + Google Sheets + Google Apps Script**.

## What it does
- Associate a student's existing ID barcode with their account.
- Load points onto the student.
- Register snacks by their UPC barcode.
- Scan student ID, then snack barcode, to deduct points automatically.
- Store all students, balances, reloads, inventory and transactions in one shared Google Sheet.

## Google Sheets setup
1. Create a blank Google Sheet.
2. Open **Extensions → Apps Script**.
3. Replace `Code.gs` with `apps-script/Code.gs`.
4. Create an HTML file named `Index` and paste in `apps-script/Index.html`.
5. Deploy as a **Web app**.
6. Execute as **Me**.
7. Use the most restrictive access setting that works for your school.
8. Copy the deployed `/exec` URL.

The script automatically creates these tabs:
- Students
- Snacks
- Reloads
- Transactions

## GitHub Pages
The root `index.html` is the GitHub Pages front door. Enable Pages from the repository's **main branch / root**.

When you first open the Pages site, paste the Apps Script `/exec` URL. The browser remembers it.

## Scanner workflow
Most USB barcode scanners work like keyboards and send Enter after a scan.

Checkout:
1. Scan student ID.
2. Student name and balance appear.
3. Scan snack UPC.
4. Points are deducted.
5. The sale is logged to Google Sheets.

## Privacy
Do not put student names, student IDs, balances, or transaction exports in this public repository. Student data should stay in the Google Sheet.