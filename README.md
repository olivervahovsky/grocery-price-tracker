# Grocery Price Tracker

Grocery Price Tracker is an offline, single-user Windows 11 desktop app for
recording grocery prices and reviewing their history. It uses a dark Tkinter
interface and stores purchases locally in SQLite.

## What you can do

- **Record a purchase.** Enter a product, price, shop, and purchase date.
  Choose price per kilogram, price per piece, or price for a specific amount.
  The amount is free text, so you can enter values such as `500g`, `250ml`,
  `6ks`, or `285g/157g`.
- **Find and manage purchases.** Search by part of a product name, ignoring
  case. Sort the records by clicking a column header. Delete individual
  records after confirming the selection.
- **Keep price history.** Updating a purchase adds a new record with the
  new price, shop, and date. It copies the selected record’s price type and
  amount and keeps the earlier record in history.
- **Review statistics.** Search products and see their minimum, maximum,
  average, and record count. Double-click a product to view its ten most
  recent prices and purchase dates. Statistics use recorded prices across
  all price types without unit conversion.
- **Exchange data with Excel.** Import multiple prices per product row or
  export the full history to `PriceHistory.xlsx`.

## Build and run on Windows 11

1. Install Python 3.12 with Tcl/Tk and the Windows Python launcher. If you use
   Python Install Manager, run `py install 3.12`.
2. Clone or download this repository and open Command Prompt in its folder.
3. Run `build-windows.bat`. It creates a local virtual environment, installs
   the dependencies, and builds the application with PyInstaller.
4. Open `dist\GroceryTracker.exe`. The executable bundles its dependencies
   and runs offline. The initial build needs internet access to install packages.

After running the build script, you can also run the app from source with
`.venv\Scripts\python.exe GroceryTracker.py`.

## Excel import and export

Import reads the first worksheet of an `.xlsx` file. An optional header can
begin with **Shop**, **Grocery** or **Product**, and **Type** or **Amount**.

| Shop | Product | Amount | Price1 | Price2 |
| --- | --- | --- | ---: | ---: |
| Lidl | Bananas | 1kg | 24.90 | 25.90 |
| Tesco | Milk | 1ks | 28.90 | |
| Albert | Cheese | 250g | 89.90 | 99.90 |

Column C treats `1kg` as price per kilogram, `1ks` as price per piece, and
any other nonblank value as a specific amount. Every nonblank price cell
from column D onward creates a separate record dated today. Paste formula
results as values before importing.

The app asks whether to keep a possible duplicate when product, price, shop,
and date match. This also applies to records accepted earlier in the same
workbook. Choosing **No** skips that price and continues the import.

Export includes **Product**, **Price**, **Type**, **Amount**, **Shop**, and
**Date**. Import expects the layout shown above.

## Local data and backups

The database is created on first launch at
`%LOCALAPPDATA%\GroceryTracker\GroceryTracker.sqlite3`. Existing databases
from the earlier version are upgraded automatically. Saved weight and unit
values become amount text, such as `250 g`.

To back up the history, close the app and copy the database file. To restore it,
close the app and replace the database file with your backup.
