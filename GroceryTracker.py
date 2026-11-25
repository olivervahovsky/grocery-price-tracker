"""Offline desktop application for recording grocery purchases."""

import calendar
import os
import sqlite3
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException
from xml.etree.ElementTree import ParseError

from storage import Database, EXPORT_HEADERS, PRICE_TYPES, Purchase, export_excel, number, read_excel

RECORD_FIELDS = ("name", "price", "price_type", "amount",
                 "shop_name", "purchase_date")
RECORD_HEADERS = ("Product", "Price", "Price Type", *EXPORT_HEADERS[3:])


BACKGROUND = "#181b21"
SURFACE = "#232831"
BORDER = "#434c5a"
TEXT = "#eef1f6"
MUTED = "#9da8b7"
ACCENT = "#345b8c"
HOVER = "#303a49"


def apply_dark_theme(root):
    root.configure(background=BACKGROUND)
    root.option_add("*Listbox.background", SURFACE)
    root.option_add("*Listbox.foreground", TEXT)
    root.option_add("*Listbox.selectBackground", ACCENT)
    root.option_add("*Listbox.selectForeground", TEXT)
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=BACKGROUND, foreground=TEXT, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER, troughcolor=SURFACE,
                    focuscolor=MUTED, selectbackground=ACCENT, selectforeground=TEXT)
    style.configure("TButton", background=SURFACE, padding=(8, 5))
    style.map("TButton", background=[("pressed", ACCENT), ("active", HOVER)],
              foreground=[("disabled", MUTED)])
    for widget in ("TEntry", "TCombobox"):
        style.configure(widget, fieldbackground=SURFACE, foreground=TEXT,
                        insertcolor=TEXT, arrowcolor=TEXT, padding=4)
        style.map(widget, fieldbackground=[("disabled", BACKGROUND), ("readonly", SURFACE)],
                  foreground=[("disabled", MUTED), ("readonly", TEXT)])
    style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE,
                    foreground=TEXT, rowheight=28)
    style.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", TEXT)])
    style.configure("Treeview.Heading", background=BACKGROUND, padding=(6, 5))
    style.map("Treeview.Heading", background=[("active", HOVER)])
    style.configure("TScrollbar", background=BORDER, arrowcolor=TEXT)
    style.map("TScrollbar", background=[("active", ACCENT)])



def database_path():
    folder = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".local" / "share")
    folder = folder / "GroceryTracker"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "GroceryTracker.sqlite3"


class CalendarPicker(tk.Toplevel):
    def __init__(self, parent, variable):
        super().__init__(parent)
        self.configure(background=BACKGROUND)
        self.title("Choose purchase date")
        self.transient(parent)
        self.resizable(False, False)
        self.variable = variable
        try:
            initial = date.fromisoformat(variable.get())
        except ValueError:
            initial = date.today()
        self.month = initial.replace(day=1)
        self.body = ttk.Frame(self, padding=12)
        self.body.grid()
        self.bind("<Escape>", lambda event: self.dismiss())
        self.render()
        self.grab_set()

    def dismiss(self):
        self.destroy()
        self.master.grab_set()

    def move(self, offset):
        year, month = divmod(self.month.year * 12 + self.month.month - 1 + offset, 12)
        self.month = date(year, month + 1, 1)
        self.render()

    def select(self, selected):
        self.variable.set(selected.isoformat())
        self.dismiss()

    def render(self):
        for child in self.body.winfo_children():
            child.destroy()
        self.protocol("WM_DELETE_WINDOW", self.dismiss)
        previous = ttk.Button(self.body, text="‹", width=4, command=lambda: self.move(-1))
        previous.grid(row=0, column=0)
        following = ttk.Button(self.body, text="›", width=4, command=lambda: self.move(1))
        following.grid(row=0, column=6)
        if self.month == date.min:
            previous.state(["disabled"])
        if self.month == date(9999, 12, 1):
            following.state(["disabled"])
        ttk.Label(self.body, text=self.month.strftime("%B %Y")).grid(row=0, column=1,
                                                                      columnspan=5)
        for column, day in enumerate(calendar.day_abbr):
            ttk.Label(self.body, text=day).grid(row=1, column=column, pady=4)
        for row, week in enumerate(calendar.Calendar().monthdayscalendar(
                self.month.year, self.month.month), start=2):
            for column, day in enumerate(week):
                if day:
                    selected = self.month.replace(day=day)
                    ttk.Button(self.body, text=str(day), width=4,
                               command=lambda value=selected: self.select(value)).grid(
                                   row=row, column=column, padx=1, pady=1)
        ttk.Button(self.body, text="Today", command=lambda: self.select(date.today())).grid(
            row=8, column=0, columnspan=7, pady=(8, 0))


class PurchaseForm(tk.Toplevel):
    def __init__(self, app, original=None):
        super().__init__(app.root)
        self.configure(background=BACKGROUND)
        self.app = app
        self.original = original
        self.title("Update Purchase" if original else "New Purchase")
        self.transient(app.root)
        self.resizable(False, False)
        self.bind("<Escape>", lambda event: self.destroy())
        body = ttk.Frame(self, padding=20)
        body.grid(sticky="nsew")
        self.product = tk.StringVar(value=original["name"] if original else "")
        self.price = tk.StringVar()
        self.kind = tk.StringVar(value=PRICE_TYPES["1kg"])
        self.amount = tk.StringVar()
        self.date = tk.StringVar(value=date.today().isoformat())
        self.shop = tk.StringVar(value=original["shop_name"] if original else "")

        product_field = ttk.Entry(body, textvariable=self.product, width=32)
        self.field(body, 0, "Product", product_field)
        if original:
            product_field.state(["readonly"])
        price_field = ttk.Entry(body, textvariable=self.price)
        self.field(body, 1, "New Price" if original else "Price", price_field)
        if original is None:
            kinds = ttk.Combobox(body, textvariable=self.kind, values=tuple(PRICE_TYPES.values()),
                                 state="readonly")
            self.field(body, 2, "Price Type", kinds)
            self.amount_field = ttk.Entry(body, textvariable=self.amount)
            self.field(body, 3, "Amount", self.amount_field)
            kinds.bind("<<ComboboxSelected>>", self.toggle_amount)
            self.toggle_amount()
        date_field = ttk.Frame(body)
        ttk.Entry(date_field, textvariable=self.date, width=14).pack(side="left")
        ttk.Button(date_field, text="Calendar…", command=lambda: CalendarPicker(
            self, self.date)).pack(side="left", padx=(6, 0))
        self.field(body, 5, "Date (YYYY-MM-DD)", date_field)
        self.field(body, 6, "Shop", ttk.Entry(body, textvariable=self.shop))

        actions = ttk.Frame(body)
        actions.grid(row=7, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="left", padx=6)
        ttk.Button(actions, text="Save", command=self.save).pack(side="left")
        self.bind("<Return>", lambda event: self.save())
        self.grab_set()
        (price_field if original else product_field).focus_set()

    @staticmethod
    def field(parent, row, label, widget):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 16), pady=5)
        widget.grid(row=row, column=1, sticky="ew", pady=5)

    def toggle_amount(self, event=None):
        enabled = self.kind.get() == PRICE_TYPES["specific amount"]
        self.amount_field.configure(state="normal" if enabled else "disabled")

    def purchase(self):
        product, shop = self.product.get().strip(), self.shop.get().strip()
        if not product or not shop:
            raise ValueError("Enter both a product and a shop.")
        price = number(self.price.get(), "Price")
        raw_date = self.date.get().strip()
        try:
            purchase_date = date.fromisoformat(raw_date)
            if purchase_date.isoformat() != raw_date:
                raise ValueError
        except ValueError:
            raise ValueError("Enter a valid date in YYYY-MM-DD format.") from None
        amount = None
        if self.original:
            kind = self.original["price_type"]
            amount = self.original["amount"]
        else:
            kind = next(code for code, label in PRICE_TYPES.items() if label == self.kind.get())
            if kind == "specific amount":
                amount = self.amount.get()
                if not amount.strip():
                    raise ValueError("Enter an amount.")
        return Purchase(product, price, kind, amount, shop, purchase_date)

    def save(self):
        try:
            purchase = self.purchase()
            if self.app.db.is_duplicate(purchase) and not messagebox.askyesno(
                "Possible duplicate detected", "Possible duplicate detected.\nSave anyway?",
                parent=self, default="no"
            ):
                return
            self.app.db.save([purchase])
        except (ValueError, sqlite3.Error) as error:
            messagebox.showerror("Could not save purchase", str(error), parent=self)
            return
        self.destroy()
        self.app.refresh()
        self.app.status.set("Purchase saved.")


class Tracker:
    def __init__(self, root, db):
        apply_dark_theme(root)
        self.root = root
        self.db = db
        self.status = tk.StringVar()
        self.refresh = lambda: None
        root.title("Grocery Price Tracker")
        root.geometry("1000x600")
        root.minsize(760, 450)
        root.protocol("WM_DELETE_WINDOW", self.close)

        header = ttk.Frame(root, padding=(18, 14))
        header.pack(fill="x")
        ttk.Label(header, text="Grocery Price Tracker", font=("TkDefaultFont", 18, "bold")).pack(
            anchor="w", pady=(0, 10))
        for label, command in (
            ("New Purchase", lambda: PurchaseForm(self)),
            ("Search Purchases", lambda: self.show_records(search=True)),
            ("Display All Records", self.show_records),
            ("Import Excel", self.import_workbook),
            ("Export Excel", self.export_workbook),
            ("Statistics", self.show_statistics),
            ("Exit", self.close),
        ):
            ttk.Button(header, text=label, command=command).pack(side="left", padx=(0, 5))

        ttk.Separator(root).pack(fill="x", padx=18)
        self.body = ttk.Frame(root, padding=(18, 12))
        self.body.pack(fill="both", expand=True)
        ttk.Label(root, textvariable=self.status, padding=(18, 4)).pack(anchor="w")
        self.show_records()

    def clear(self):
        for child in self.body.winfo_children():
            child.destroy()

    @staticmethod
    def table(parent, headers, widths):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, pady=(8, 0))
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=headers, show="headings", selectmode="browse")
        tree.grid(row=0, column=0, sticky="nsew")
        for orientation, scroll_option, callback, row, column, sticky in (
            ("vertical", "yscrollcommand", tree.yview, 0, 1, "ns"),
            ("horizontal", "xscrollcommand", tree.xview, 1, 0, "ew"),
        ):
            scrollbar = ttk.Scrollbar(frame, orient=orientation, command=callback)
            scrollbar.grid(row=row, column=column, sticky=sticky)
            tree.configure(**{scroll_option: scrollbar.set})
        for header, width in zip(headers, widths):
            tree.heading(header, text=header)
            tree.column(header, width=width, minwidth=50, anchor="w")
        return tree

    def search_bar(self):
        controls = ttk.Frame(self.body)
        controls.pack(fill="x")
        query = tk.StringVar()
        ttk.Label(controls, text="Product").pack(side="left")
        field = ttk.Entry(controls, textvariable=query, width=26)
        field.pack(side="left", padx=8)
        return controls, query, field

    def show_records(self, search=False):
        self.clear()
        self.status.set("")
        ttk.Label(self.body, text="Search Purchases" if search else "All Records",
                  font=("TkDefaultFont", 13, "bold")).pack(anchor="w", pady=(0, 8))
        controls, query, field = self.search_bar()
        tree = self.table(self.body, RECORD_HEADERS, (195, 85, 190, 135, 135, 105))
        records = {}
        sort_index = RECORD_FIELDS.index("purchase_date")
        reverse = True

        def sort(index):
            nonlocal sort_index, reverse
            reverse = not reverse if index == sort_index else False
            sort_index = index
            populate()

        def populate():
            selected = tree.selection()
            rows = self.db.records(query.get().strip())
            key = RECORD_FIELDS[sort_index]

            def sort_value(row):
                value = row[key]
                if key == "price_type":
                    value = PRICE_TYPES.get(value, value)
                return value.casefold() if isinstance(value, str) else value

            present = [row for row in rows if row[key] is not None]
            missing = [row for row in rows if row[key] is None]
            ordered = sorted(present, key=sort_value, reverse=reverse) + missing
            for index, header in enumerate(RECORD_HEADERS):
                arrow = (" ▼" if reverse else " ▲") if index == sort_index else ""
                tree.heading(header, text=header + arrow, command=lambda i=index: sort(i))
            tree.delete(*tree.get_children())
            records.clear()
            for row in ordered:
                record_id = str(row["id"])
                records[record_id] = row
                tree.insert("", "end", iid=record_id, values=(
                    row["name"], f'{row["price"]:.2f}',
                    PRICE_TYPES.get(row["price_type"], row["price_type"]),
                    row["amount"] or "", row["shop_name"], row["purchase_date"],
                ))
            if selected and selected[0] in records:
                tree.selection_set(selected[0])
                tree.focus(selected[0])
                tree.see(selected[0])
            self.status.set(f"{len(rows):,} record(s)")

        def update():
            selected = tree.selection()
            if selected:
                PurchaseForm(self, records[selected[0]])

        def delete():
            selected = tree.selection()
            if selected and messagebox.askyesno("Delete purchase", "Delete this entry?",
                                                parent=self.root, default="no"):
                try:
                    self.db.delete(int(selected[0]))
                except sqlite3.Error as error:
                    messagebox.showerror("Could not delete purchase", str(error), parent=self.root)
                    return
                populate()

        delete_button = ttk.Button(controls, text="Delete Selected Record", command=delete,
                                   state="disabled")
        delete_button.pack(side="right")
        buttons = [delete_button]
        if search:
            update_button = ttk.Button(controls, text="Update", command=update, state="disabled")
            update_button.pack(side="right", padx=6)
            buttons.append(update_button)

        def selection_changed(event):
            state = "normal" if tree.selection() else "disabled"
            for button in buttons:
                button.configure(state=state)

        tree.bind("<<TreeviewSelect>>", selection_changed)
        ttk.Button(controls, text="Refresh", command=populate).pack(side="left")
        query.trace_add("write", lambda *_: populate())
        self.refresh = populate
        populate()
        if search:
            field.focus_set()

    def show_statistics(self):
        self.clear()
        ttk.Label(self.body, text="Statistics", font=("TkDefaultFont", 13, "bold")).pack(
            anchor="w")
        ttk.Label(self.body, text="Double-click a product to see its ten most recent prices.").pack(
            anchor="w", pady=(4, 8))
        _, query, field = self.search_bar()
        tree = self.table(self.body, ("Product", "Lowest Price", "Highest Price",
                                     "Average Price", "Number of Records"),
                          (230, 135, 135, 135, 150))

        def populate():
            tree.delete(*tree.get_children())
            rows = self.db.statistics(query.get().strip())
            for product_id, name, low, high, average, count in rows:
                tree.insert("", "end", iid=str(product_id), values=(
                    name, f"{low:.2f}", f"{high:.2f}", f"{average:.2f}", count,
                ))
            self.status.set(f"{len(rows):,} product(s)")

        def open_recent(event):
            if event.type == tk.EventType.ButtonPress:
                item = tree.identify_row(event.y)
            else:
                item = tree.focus()
            if item:
                self.show_recent_prices(int(item), tree.set(item, "Product"))

        tree.bind("<Double-1>", open_recent)
        tree.bind("<Return>", open_recent)
        query.trace_add("write", lambda *_: populate())
        self.refresh = populate
        populate()
        field.focus_set()

    def show_recent_prices(self, product_id, product_name):
        rows = self.db.recent_prices(product_id)
        window = tk.Toplevel(self.root, background=BACKGROUND)
        window.title(f"{product_name} · Recent Prices")
        window.transient(self.root)
        window.minsize(320, 200)
        body = ttk.Frame(window, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"Recent prices for {product_name}", wraplength=440).pack(anchor="w")
        tree = self.table(body, ("Date", "Price"), (155, 105))
        tree.configure(height=max(1, len(rows)))
        for price, purchase_date in rows:
            tree.insert("", "end", values=(purchase_date, f"{price:.2f}"))
        ttk.Button(body, text="Close", command=window.destroy).pack(anchor="e", pady=(10, 0))
        window.bind("<Escape>", lambda event: window.destroy())
        window.grab_set()
        tree.focus_set()

    def import_workbook(self):
        path = filedialog.askopenfilename(parent=self.root, title="Import Excel",
                                          filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            purchases = read_excel(path, date.today())
            known = self.db.duplicate_keys()
            accepted = []
            skipped = 0
            for purchase in purchases:
                key = purchase.duplicate_key
                if key in known and not messagebox.askyesno(
                    "Possible duplicate detected",
                    f"Possible duplicate detected.\n\n{purchase.product} · "
                    f"{purchase.price:.2f} · {purchase.shop} · {purchase.date}\n\nSave anyway?",
                    parent=self.root, default="no",
                ):
                    skipped += 1
                    continue
                accepted.append(purchase)
                known.add(key)
            if accepted:
                self.db.save(accepted)
        except (ValueError, OSError, InvalidFileException, BadZipFile, ParseError,
                KeyError, IndexError, sqlite3.Error) as error:
            messagebox.showerror("Import failed", f"No purchases were imported.\n\n{error}",
                                 parent=self.root)
            return
        self.refresh()
        messagebox.showinfo("Import complete",
                            f"Imported {len(accepted):,} record(s).\n"
                            f"Skipped {skipped:,} duplicate(s).", parent=self.root)

    def export_workbook(self):
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Export Excel", initialfile="PriceHistory.xlsx",
            defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not path:
            return
        try:
            export_excel(path, self.db.records())
        except (OSError, ValueError, sqlite3.Error) as error:
            messagebox.showerror("Export failed", str(error), parent=self.root)
            return
        self.status.set(f"Exported {Path(path).name}.")

    def close(self):
        self.db.close()
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        db = Database(database_path())
    except (OSError, sqlite3.Error) as error:
        root.withdraw()
        messagebox.showerror("Could not open database", str(error), parent=root)
        root.destroy()
        return
    Tracker(root, db)
    root.mainloop()


if __name__ == "__main__":
    main()
