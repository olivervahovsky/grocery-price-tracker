"""Editable Tk entry with database-backed suggestions."""

import tkinter as tk
from tkinter import ttk


class SuggestionEntry(ttk.Entry):
    def __init__(self, parent, variable, lookup, colors, **kwargs):
        super().__init__(parent, textvariable=variable, **kwargs)
        self.variable = variable
        self.lookup = lookup
        self.colors = colors
        self.popup = None
        self.listbox = None
        self.accepting = False
        self.trace = variable.trace_add("write", self.refresh)
        self.bind("<FocusOut>", self.hide, add=True)
        self.bind("<Unmap>", self.hide, add=True)
        self.bind("<Destroy>", self.cleanup, add=True)
        self.bind("<Down>", lambda event: self.move_selection(1))
        self.bind("<Up>", lambda event: self.move_selection(-1))
        self.bind("<Return>", self.accept_selection)
        self.bind("<Escape>", self.hide_on_escape)

    def refresh(self, *_):
        if self.accepting or self.focus_get() != self or self.instate(("disabled",)):
            self.hide()
            return
        query = self.variable.get()
        choices = [value for value in self.lookup(query)
                   if value.casefold() != query.strip().casefold()]
        if not choices:
            self.hide()
            return
        if self.popup is None:
            self.popup = tk.Toplevel(self)
            self.popup.withdraw()
            self.popup.overrideredirect(True)
            self.popup.transient(self.winfo_toplevel())
            background, foreground, selection, border = self.colors
            self.listbox = tk.Listbox(
                self.popup, exportselection=False, takefocus=False,
                background=background, foreground=foreground, selectbackground=selection,
                selectforeground=foreground, highlightbackground=border,
                highlightthickness=1, relief="flat",
            )
            self.listbox.pack(fill="both", expand=True)
            self.listbox.bind("<ButtonPress-1>", self.click)
        self.listbox.delete(0, "end")
        self.listbox.insert("end", *choices)
        self.listbox.configure(height=len(choices))
        self.popup.update_idletasks()
        if self.popup is None:
            return
        width = max(self.winfo_width(), 220)
        height = self.popup.winfo_reqheight()
        x = min(self.winfo_rootx(), self.winfo_screenwidth() - width)
        y = self.winfo_rooty() + self.winfo_height()
        if y + height > self.winfo_screenheight():
            y = self.winfo_rooty() - height
        self.popup.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")
        self.popup.deiconify()
        self.popup.lift()

    def move_selection(self, offset):
        if self.listbox is None:
            self.refresh()
        if self.listbox is None:
            return None
        selected = self.listbox.curselection()
        index = selected[0] + offset if selected else (0 if offset > 0 else self.listbox.size() - 1)
        index = max(0, min(index, self.listbox.size() - 1))
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(index)
        self.listbox.see(index)
        return "break"

    def click(self, event):
        self.accept(self.listbox.nearest(event.y))
        return "break"

    def accept_selection(self, event):
        if self.listbox is not None and (selected := self.listbox.curselection()):
            self.accept(selected[0])
            return "break"
        return None

    def accept(self, index):
        value = self.listbox.get(index)
        self.hide()
        self.accepting = True
        try:
            self.variable.set(value)
            self.icursor("end")
        finally:
            self.accepting = False

    def hide_on_escape(self, event):
        if self.popup is not None:
            self.hide()
            return "break"
        return None

    def hide(self, event=None):
        if self.popup is not None:
            popup = self.popup
            self.popup = self.listbox = None
            popup.destroy()

    def cleanup(self, event):
        self.variable.trace_remove("write", self.trace)
        self.hide()
