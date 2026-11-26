"""SQLite storage and Excel exchange for Grocery Price Tracker."""

import math
import sqlite3
from collections import Counter
from contextlib import closing
from dataclasses import dataclass, replace
from datetime import date
from itertools import chain
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

PRICE_TYPES = {
    "1kg": "Price per kg",
    "specific amount": "Price for specific amount",
    "1ks": "Price per piece",
}
LEGACY_AMOUNT = """CASE WHEN weight_value IS NOT NULL
    THEN trim(printf('%.15g', weight_value) || ' ' || coalesce(weight_unit, '')) END"""
EXPORT_HEADERS = ("Product", "Price", "Type", "Amount", "Shop", "Date")


def number(value, label):
    try:
        if isinstance(value, bool):
            raise ValueError
        parsed = float(str(value).strip().replace(",", "."))
        if not math.isfinite(parsed) or parsed < 0:
            raise ValueError
        return parsed
    except (ValueError, OverflowError):
        raise ValueError(f"{label} must be a finite, non-negative number.") from None


@dataclass(frozen=True)
class Purchase:
    product: str
    price: float
    price_type: str
    amount: str | None
    shop: str
    date: date

    @property
    def duplicate_key(self):
        return self.product.casefold(), self.price, self.shop.casefold(), self.date.isoformat()


class Database:
    def __init__(self, path):
        self.connection = sqlite3.connect(path)
        self.connection.row_factory = sqlite3.Row
        self.connection.create_function("casefold", 1, str.casefold, deterministic=True)
        self.connection.executescript("""
            PRAGMA foreign_keys = ON;
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE
            );
            CREATE TABLE IF NOT EXISTS price_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                price REAL NOT NULL,
                price_type TEXT NOT NULL,
                amount TEXT,
                purchase_date DATE NOT NULL,
                shop_name TEXT NOT NULL,
                FOREIGN KEY(product_id) REFERENCES products(id)
            );
            CREATE INDEX IF NOT EXISTS history_product ON price_history(product_id);
        """)

        self.upgrade_legacy_amounts()

    def upgrade_legacy_amounts(self):
        with self.connection:
            self.connection.execute("BEGIN IMMEDIATE")
            columns = {row["name"] for row in self.connection.execute(
                "PRAGMA table_info(price_history)"
            )}
            if "amount" not in columns:
                self.connection.execute("ALTER TABLE price_history ADD COLUMN amount TEXT")
                self.connection.execute(f"""
                    UPDATE price_history
                    SET price_type = 'specific amount',
                        amount = {LEGACY_AMOUNT}
                    WHERE price_type = 'specific weight'
                """)

    def records(self, search=""):
        return self.connection.execute("""
            SELECT h.id, p.name, h.price, h.price_type, h.amount,
                   h.shop_name, h.purchase_date
            FROM price_history h JOIN products p ON h.product_id = p.id
            WHERE casefold(p.name) LIKE ? ESCAPE '!'
            ORDER BY h.purchase_date DESC, h.id DESC
        """, (search_pattern(search),)).fetchall()

    def suggestions(self, field, text):
        fields = {
            "product": "p.name",
            "shop": "h.shop_name",
            "amount": "h.amount",
            "price": "CAST(h.price AS TEXT)",
            "date": "h.purchase_date",
        }
        column = fields[field]
        query = text.strip()
        if not query:
            return []
        if field == "price":
            query = query.replace(",", ".")
        return [row[0] for row in self.connection.execute(f"""
            SELECT MIN({column}) AS value
            FROM price_history h JOIN products p ON h.product_id = p.id
            WHERE {column} IS NOT NULL AND casefold({column}) LIKE ? ESCAPE '!'
            GROUP BY casefold({column})
            ORDER BY instr(casefold(value), ?), casefold(value)
            LIMIT 8
        """, (search_pattern(query), query.casefold()))]

    def is_duplicate(self, purchase, exclude_id=None):
        return self.connection.execute("""
            SELECT 1 FROM price_history h JOIN products p ON h.product_id = p.id
            WHERE casefold(p.name) = ? AND h.price = ? AND casefold(h.shop_name) = ?
                  AND h.purchase_date = ? AND h.id IS NOT ? LIMIT 1
        """, (*purchase.duplicate_key, exclude_id)).fetchone() is not None

    def duplicate_keys(self):
        return {tuple(row) for row in self.connection.execute("""
            SELECT casefold(p.name), h.price, casefold(h.shop_name), h.purchase_date
            FROM price_history h JOIN products p ON p.id = h.product_id
        """)}

    def product_id(self, name):
        product = self.connection.execute("""
            SELECT id FROM products WHERE casefold(name) = ? ORDER BY id LIMIT 1
        """, (name.casefold(),)).fetchone()
        if product:
            return product["id"]
        return self.connection.execute(
            "INSERT INTO products(name) VALUES (?)", (name,)
        ).lastrowid

    def write_purchase(self, purchase, record_id=None, product_id=None):
        if product_id is None:
            product_id = self.product_id(purchase.product)
        values = (product_id, purchase.price, purchase.price_type,
                  purchase.amount, purchase.date.isoformat(), purchase.shop)
        if record_id is None:
            self.connection.execute("""
                INSERT INTO price_history (
                    product_id, price, price_type, amount, purchase_date, shop_name
                ) VALUES (?, ?, ?, ?, ?, ?)
            """, values)
        else:
            self.connection.execute("""
                UPDATE price_history
                SET product_id = ?, price = ?, price_type = ?, amount = ?,
                    purchase_date = ?, shop_name = ?
                WHERE id = ?
            """, (*values, record_id))

    def save(self, purchases):
        with self.connection:
            self.connection.execute("BEGIN IMMEDIATE")
            for purchase in purchases:
                self.write_purchase(purchase)

    def edit(self, record_id, purchase):
        with self.connection:
            self.connection.execute("BEGIN IMMEDIATE")
            self.write_purchase(purchase, record_id)

    def edit_product_purchases(self, product_id, changes):
        with self.connection:
            self.connection.execute("BEGIN IMMEDIATE")
            rows = self.connection.execute("""
                SELECT h.id, p.name, h.price, h.price_type, h.amount,
                       h.shop_name, h.purchase_date
                FROM price_history h JOIN products p ON h.product_id = p.id
                WHERE p.id = ?
            """, (product_id,)).fetchall()
            if not rows:
                return 0
            new_name = changes["product"]
            target_id = product_id
            if new_name != rows[0]["name"]:
                if new_name.casefold() == rows[0]["name"].casefold():
                    self.connection.execute("UPDATE products SET name = ? WHERE id = ?",
                                            (new_name, product_id))
                else:
                    target_id = self.product_id(new_name)
            for row in rows:
                purchase = replace(purchase_from_row(row), **changes)
                self.write_purchase(purchase, row["id"], target_id)
            if target_id != product_id:
                self.connection.execute("DELETE FROM products WHERE id = ?", (product_id,))
            return len(rows)

    def import_database(self, path):
        purchases = read_database(path)
        with self.connection:
            self.connection.execute("BEGIN IMMEDIATE")
            remaining = Counter(import_key(purchase_from_row(row)) for row in self.records())
            imported = 0
            for purchase in purchases:
                key = import_key(purchase)
                if remaining[key]:
                    remaining[key] -= 1
                else:
                    self.write_purchase(purchase)
                    imported += 1
        return imported, len(purchases) - imported

    def delete(self, record_id):
        with self.connection:
            self.connection.execute("DELETE FROM price_history WHERE id = ?", (record_id,))

    def statistics(self, search=""):
        return self.connection.execute("""
            SELECT p.id, p.name, MIN(h.price), MAX(h.price), AVG(h.price), COUNT(h.id)
            FROM price_history h JOIN products p ON h.product_id = p.id
            WHERE casefold(p.name) LIKE ? ESCAPE '!'
            GROUP BY p.id ORDER BY casefold(p.name)
        """, (search_pattern(search),)).fetchall()

    def recent_prices(self, product_id):
        return self.connection.execute("""
            SELECT price, purchase_date FROM price_history
            WHERE product_id = ?
            ORDER BY purchase_date DESC, id DESC LIMIT 10
        """, (product_id,)).fetchall()

    def close(self):
        self.connection.close()


def purchase_from_row(row):
    return Purchase(row["name"], row["price"], row["price_type"], row["amount"],
                    row["shop_name"], date.fromisoformat(row["purchase_date"]))


def import_key(purchase):
    return purchase.duplicate_key, purchase.price_type, purchase.amount


def read_database(path):
    uri = Path(path).resolve().as_uri() + "?mode=ro"
    with closing(sqlite3.connect(uri, uri=True)) as source:
        source.row_factory = sqlite3.Row
        source.execute("PRAGMA trusted_schema = OFF")
        source.execute("BEGIN")
        tables = {row[0] for row in source.execute(
            "SELECT name FROM sqlite_schema WHERE type = 'table'"
        )}
        if not {"products", "price_history"}.issubset(tables):
            raise ValueError("Select a Grocery Price Tracker SQLite database.")
        columns = {row["name"] for row in source.execute("PRAGMA table_info(price_history)")}
        if "amount" in columns:
            kind, amount = "h.price_type", "h.amount"
        elif {"weight_value", "weight_unit"}.issubset(columns):
            kind = "CASE h.price_type WHEN 'specific weight' THEN 'specific amount' ELSE h.price_type END"
            amount = f"CASE WHEN h.price_type = 'specific weight' THEN {LEGACY_AMOUNT} END"
        else:
            raise ValueError("This purchase history database format is unsupported.")
        rows = source.execute(f"""
            SELECT p.name, h.price, {kind} AS price_type, {amount} AS amount,
                   h.shop_name, h.purchase_date
            FROM price_history h LEFT JOIN products p ON p.id = h.product_id
            ORDER BY h.id
        """)
        purchases = []
        for row in rows:
            if (not isinstance(row["name"], str) or not row["name"].strip()
                    or not isinstance(row["shop_name"], str) or not row["shop_name"].strip()
                    or row["price_type"] not in PRICE_TYPES
                    or (row["amount"] is not None and not isinstance(row["amount"], str))
                    or not isinstance(row["purchase_date"], str)):
                raise ValueError("The database contains an invalid purchase record.")
            purchase = purchase_from_row(row)
            purchases.append(replace(purchase, price=number(purchase.price, "Stored price")))
        return purchases


def search_pattern(search):
    escaped = search.casefold().replace("!", "!!").replace("%", "!%").replace("_", "!_")
    return f"%{escaped}%"


def cell_text(value):
    return "" if value is None else str(value).strip()


def read_excel(path, purchase_date):
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        rows = workbook.worksheets[0].iter_rows(values_only=True)
        first_row = next(rows, ())
        labels = tuple(cell_text(value).casefold() for value in first_row[:3])
        has_header = labels[:2] in {("shop", "grocery"), ("shop", "product")} and (
            labels[2:] in {("type",), ("amount",)}
        )
        if not has_header:
            rows = chain((first_row,), rows)
        purchases = []
        for row_number, row in enumerate(rows, start=2 if has_header else 1):
            prices = [(column, value) for column, value in enumerate(row[3:], start=4)
                      if cell_text(value)]
            if not prices:
                continue
            shop, product, amount = map(cell_text, row[:3])
            if not shop or not product or not amount:
                raise ValueError(f"Row {row_number} needs a shop, a product, and an amount or type.")
            code = amount.casefold()
            if code in {"1kg", "1ks"}:
                price_type, amount = code, None
            else:
                price_type = "specific amount"
            for column, value in prices:
                price = number(value, f"Price in {get_column_letter(column)}{row_number}")
                purchases.append(Purchase(product, price, price_type, amount, shop, purchase_date))
        return purchases
    finally:
        workbook.close()


def export_excel(path, records):
    path = Path(path)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Price History"
    sheet.append(EXPORT_HEADERS)
    for row in records:
        sheet.append((row["name"], row["price"], row["price_type"], row["amount"],
                      row["shop_name"],
                      date.fromisoformat(row["purchase_date"])))
        for column in (1, 3, 4, 5):
            cell = sheet.cell(sheet.max_row, column)
            if cell.value is not None:
                cell.data_type = "s"
        sheet.cell(sheet.max_row, 2).number_format = "0.00"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in zip("ABCDEF", (28, 12, 20, 22, 22, 15)):
        sheet.column_dimensions[column].width = width
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        workbook.save(temporary_path)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)
        workbook.close()
