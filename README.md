# Grocery Price Tracker

Grocery Price Tracker is an offline, single-user Windows 11 desktop app for
recording grocery prices and reviewing their history. It uses a dark Tkinter
interface and stores purchases locally in SQLite.

## What you can do

- **Record a shopping trip.** In **New Purchase**, enter the shop and date
  once. Click **Add Item** to enter more groceries and review or remove
  rows before clicking **Done**. Each item is saved as a price record
  with the same shop and date. You can choose price per kilogram, price
  per piece, or a free-text amount such as `500g`, `250ml`, or `6ks`.
- **Reuse saved values.** The app suggests matches as you type in
  product, shop, amount, price, and date fields. Each field suggests
  values from its own part of the purchase history. Click a suggestion
  or use the arrow keys and Enter. You can also enter a new value.
- **Find and manage purchases.** Search Purchases shows the full history when
  the search box is empty. Filter by part of a product name, sort by clicking
  a column header, or delete a selected record after confirming.
- **Correct a record.** Select a record in Search Purchases and click **Edit**
  to change its fields in place. Use **Update** to add a new price to history.
- **Keep price history.** Updating a purchase adds a new record with the
  new price, shop, and date. It copies the selected record’s price type and
  amount and keeps the earlier record in history.
- **Review statistics.** Search products and see their minimum, maximum,
  average, and record count. Double-click a product to view its ten most
  recent prices and purchase dates. Statistics use recorded prices across
  all price types without unit conversion. Select a product and click **Edit**
  to change its name or apply shared values to its purchases. Blank price,
  date, and shop fields retain each record’s values. Choose **Keep existing
  price type** to retain each record’s price type and amount.
- **Exchange data with Excel.** Import multiple prices per product row or
  export the full history to `PriceHistory.xlsx`.

## Download or build on Windows 11

See [GitHub Releases](https://github.com/olivervahovsky/grocery-price-tracker/releases)
for the latest Windows executable. The release also includes a source
archive if you prefer to build it yourself.

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
and date match. This also applies to entries accepted earlier in the same
shopping trip or Excel workbook. Choosing **No** skips that item.

Export includes **Product**, **Price**, **Type**, **Amount**, **Shop**, and
**Date**. Import expects the layout shown above.

## Local data and backups

The database is created on first launch at
`%LOCALAPPDATA%\GroceryTracker\GroceryTracker.sqlite3`. Existing databases
from the earlier version are upgraded automatically. Saved weight and unit
values become amount text, such as `250 g`.

To back up the history, close the app and copy the database file. To restore it,
close the app and replace the database file with your backup.

Use **Import Database** to bring purchase history from a previous
Grocery Price Tracker `.sqlite3` file into the current app. It adds purchases
while retaining the history already in the current database and reading the
source without modifying it. It supports the original weight-and-unit format
and the newer amount format, and preserves the original dates, shops, and
prices. Already-present entries are counted using all purchase fields, so
re-importing an unchanged database does not add another copy of its history.
