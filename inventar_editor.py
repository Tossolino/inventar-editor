"""Inventar-Editor

Eigenständiges Programm zum Bearbeiten der geöffneten Excel-Prüfliste.
Es verbindet sich per COM mit der laufenden Excel-Instanz; die Excel-Datei
selbst enthält dafür keinerlei Makros oder Add-ins.
"""

import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import urllib.error
import urllib.request
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk

from version import VERSION

try:
    import pywintypes
    import win32com.client

    HAVE_WIN32 = True
except ImportError:  # z.B. beim Entwickeln unter Linux
    HAVE_WIN32 = False

# --------------------------------------------------------------------------
# Einstellungen zur Excel-Vorlage
# --------------------------------------------------------------------------
GITHUB_OWNER = "Tossolino"
GITHUB_REPO = "inventar-editor"
ASSET_NAME = "InventarEditor.exe"

REF_SHEET = "akt. Med.-Produkte"   # Referenzblatt
BED_SHEET = "PRO Pflegebett"       # Blatt für "Daten Bett"
HEADER_ROW = 6                     # Überschriften
FIRST_ROW = 7                      # erste Datenzeile
LAST_COL = "S"                     # Zeile wird von A bis S markiert
ID_CELL = "A2"                     # hierhin wird die aktuelle Inv.-Nr. geschrieben
AREA_CELL = "B2"                   # Bereich
FLOOR_CELL = "C2"                  # Etage
DONE_COL = "L"                     # bekommt beim Übernehmen ein "x"
DONE_MARK = "x"

SELECTABLE = ["B", "C", "D", "E", "F", "G", "H", "J", "O", "P", "Q", "R", "S"]
AUTO_COLS = ["Q", "R", "S"]        # werden bei gelb/rot automatisch eingeblendet

FALLBACK_HEADERS = {
    "A": "Inv.-Nr. Geräte-ID", "B": "ID (optional)", "C": "Geräteart",
    "D": "Hersteller", "E": "Typ/Modell", "F": "Seriennummer", "G": "BJ",
    "H": "CE/GS", "I": "Etage", "J": "Zimmer", "K": "Wohnbereich",
    "L": "letzte STK", "M": "Prüfintervall", "N": "Nächste STK",
    "O": "Eigentümer", "P": "Reg.Nr.", "Q": "Mängel / Bemerkung",
    "R": "Mängelklasse", "S": "Stillgelegt",
}
FIELD_WIDTH = {
    "B": 11, "C": 14, "D": 20, "E": 16, "F": 18, "G": 6, "H": 6,
    "J": 8, "O": 10, "P": 10, "Q": 36, "R": 7, "S": 8,
}

STATUS = {
    None: ("#b5b5b5", "–"),
    "green": ("#2e9e44", "ohne Mängel"),
    "yellow": ("#f2c200", "Mangel ohne Klasse"),
    "red": ("#d93025", "Mangel mit Klasse"),
}

COLOR_CHANGED = "#fff1b8"
COLOR_FORMULA = "#ececec"
COLOR_NORMAL = "#ffffff"

XL_UP = -4162
BUSY_CODES = {-2147418111, -2147417846, -2146777998}   # Excel beschäftigt
NO_EXCEL_CODES = {-2147221021}                          # kein Excel offen


def col_index(letter):
    return ord(letter) - ord("A") + 1


def all_cols():
    return [chr(c) for c in range(ord("A"), ord(LAST_COL) + 1)]


def norm_id(value):
    """Vergleichbare Textform einer Inventarnummer (16, 16.0 und "16" sind gleich)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().lower()


def compute_status(q, r, s):
    """Grün: Q, R, S leer – Gelb: nur Q – Rot: Q und R (mit oder ohne S)."""
    q, r, s = (bool(v.strip()) for v in (q, r, s))
    if r:
        return "red"
    if q or s:
        return "yellow"
    return "green"


def parse_version(text):
    m = re.fullmatch(r"[Vv]?(\d+)\.(\d+)\.(\d+)", (text or "").strip())
    if not m:
        return None
    major, counter, year = (int(x) for x in m.groups())
    return (major, year, counter)


def is_newer(remote, local):
    r, l = parse_version(remote), parse_version(local)
    return r is not None and l is not None and r > l


def clean_header(text):
    return re.sub(r"\s+", " ", str(text or "")).replace("- ", "").strip()


def config_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "InventarEditor", "config.json")


def load_config():
    cfg = {"shown": ["C", "D", "E", "F", "J", "O", "Q"], "topmost": True}
    try:
        with open(config_path(), encoding="utf-8") as f:
            cfg.update(json.load(f))
    except (OSError, ValueError):
        pass
    return cfg


def save_config(cfg):
    try:
        os.makedirs(os.path.dirname(config_path()), exist_ok=True)
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except OSError:
        pass


def work_area(root):
    """Arbeitsbereich des Bildschirms (ohne Taskleiste)."""
    try:
        import ctypes
        from ctypes import wintypes

        rect = wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0)
        return rect.left, rect.top, rect.right, rect.bottom
    except Exception:
        return 0, 0, root.winfo_screenwidth(), root.winfo_screenheight()


def flow_place(container, widgets, width, gap_x=6, gap_y=4):
    """Ordnet Widgets nebeneinander an und bricht um, wenn die Breite nicht reicht."""
    for w in container.place_slaves():
        w.place_forget()
    x = y = line_h = 0
    for w in widgets:
        ww, wh = w.winfo_reqwidth(), w.winfo_reqheight()
        if x > 0 and x + ww > width:
            x, y, line_h = 0, y + line_h + gap_y, 0
        w.place(x=x, y=y)
        x += ww + gap_x
        line_h = max(line_h, wh)
    container.configure(height=max(1, y + line_h))


class ExcelError(Exception):
    pass


def com_message(err):
    code = err.args[0] if err.args else None
    if code in BUSY_CODES:
        return ("Excel ist gerade beschäftigt – bitte eine laufende Zelleingabe in Excel "
                "mit Enter oder Esc beenden und es erneut versuchen.")
    if code in NO_EXCEL_CODES:
        return "Excel ist nicht geöffnet."
    detail = ""
    try:
        detail = err.args[2][2] or ""
    except Exception:
        pass
    return f"Excel-Fehler: {detail or err}"


def excel_action(fn):
    """Fängt Excel-Fehler ab und zeigt sie in der Statuszeile an."""
    def wrapper(self, *args, **kwargs):
        try:
            return fn(self, *args, **kwargs)
        except ExcelError as e:
            self.message(str(e), error=True)
        except pywintypes.com_error as e:
            self.message(com_message(e), error=True)
        except Exception as e:  # noqa: BLE001
            self.message(f"Fehler: {e}", error=True)
        return "break"
    return wrapper


# --------------------------------------------------------------------------
# Zugriff auf Excel
# --------------------------------------------------------------------------
def get_book():
    try:
        app = win32com.client.GetActiveObject("Excel.Application")
    except pywintypes.com_error:
        raise ExcelError("Excel ist nicht geöffnet.")
    books = []
    try:
        active = app.ActiveWorkbook
        if active is not None:
            books.append(active)
    except pywintypes.com_error:
        pass
    for i in range(1, app.Workbooks.Count + 1):
        books.append(app.Workbooks(i))
    for wb in books:
        try:
            return app, wb, wb.Worksheets(REF_SHEET)
        except pywintypes.com_error:
            continue
    raise ExcelError(f"Keine geöffnete Liste mit dem Blatt „{REF_SHEET}“ gefunden.")


def find_rows(ws, wanted):
    last = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
    if last < FIRST_ROW:
        return []
    values = ws.Range(f"A{FIRST_ROW}:A{last}").Value2
    if not isinstance(values, tuple):
        values = ((values,),)
    return [FIRST_ROW + i for i, (v,) in enumerate(values) if norm_id(v) == wanted]


def read_row(ws, row):
    """Liefert {Spalte: (Inhalt wie in Excel eingetippt, ist_Formel, angezeigter Text)}."""
    raw = ws.Range(f"A{row}:{LAST_COL}{row}").FormulaLocal[0]
    result = {}
    for col, value in zip(all_cols(), raw):
        text = "" if value is None else str(value)
        is_formula = text.startswith("=")
        shown = ws.Range(f"{col}{row}").Text if is_formula else text
        result[col] = (text, is_formula, str(shown or ""))
    return result


# --------------------------------------------------------------------------
# Programmfenster
# --------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.headers = dict(FALLBACK_HEADERS)
        self.headers_loaded = False

        self.row = None          # Excel-Zeile der geladenen Inv.-Nr.
        self.id_norm = ""
        self.id_label = ""
        self.orig = {c: "" for c in SELECTABLE}
        self.formula = {c: False for c in SELECTABLE}
        self.vars = {c: tk.StringVar() for c in SELECTABLE}
        self.show_vars = {c: tk.BooleanVar(value=c in self.cfg["shown"]) for c in SELECTABLE}
        self.auto_shown = set()
        self.undo_stack = []
        self.loading = False
        self.status = None

        self.topmost = tk.BooleanVar(value=bool(self.cfg.get("topmost", True)))
        self.bed = tk.BooleanVar(value=False)
        self.update_result = None
        self.update_info = None
        self._layout_job = None
        self._last_width = 0

        self.build_ui()
        for c in SELECTABLE:
            self.vars[c].trace_add("write", lambda *_a, c=c: self.on_value_change(c))

        root.title(f"Inventar-Editor {VERSION}")
        root.attributes("-topmost", self.topmost.get())
        self.place_initial()
        root.bind("<Configure>", self.on_configure)
        root.bind_all("<Control-s>", lambda e: self.save_book())

        self.set_status(None)
        self.update_undo_button()
        self.id_entry.focus_set()
        self.poll()
        self.check_updates(silent=True)

    # ---------------------------------------------------------------- Aufbau
    def build_ui(self):
        root = self.root
        base = ("Segoe UI", 10)
        root.option_add("*Font", base)
        style = ttk.Style()
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Update.TButton", foreground="#b3261e", font=("Segoe UI", 9, "bold"))

        outer = tk.Frame(root, padx=8, pady=5)
        outer.pack(fill="both", expand=True)

        # Zeile 1: Dateiname | Meldung | Vordergrund, Version, Update
        top = tk.Frame(outer)
        top.pack(fill="x")
        self.file_label = tk.Label(top, text="– keine Liste geöffnet –", font=("Segoe UI", 11, "bold"))
        self.file_label.pack(side="left")
        right = tk.Frame(top)
        right.pack(side="right")
        ttk.Checkbutton(right, text="Immer im Vordergrund", variable=self.topmost,
                        command=self.toggle_topmost).pack(side="left", padx=(0, 12))
        tk.Label(right, text=VERSION, fg="#666").pack(side="left", padx=(0, 6))
        self.update_btn = ttk.Button(right, text="Nach Updates suchen",
                                     command=self.on_update_button)
        self.update_btn.pack(side="left")
        self.msg_label = tk.Label(top, text="", anchor="w", fg="#444")
        self.msg_label.pack(side="left", fill="x", expand=True, padx=16)

        # Zeile 2: Inv.-Nr., Status, Aktionen | Bereich, Etage
        line = tk.Frame(outer)
        line.pack(fill="x", pady=(5, 3))
        tk.Label(line, text="Inventar-Nr.:", font=("Segoe UI", 11)).pack(side="left")
        self.id_entry = tk.Entry(line, width=10, font=("Segoe UI", 14, "bold"), justify="center")
        self.id_entry.pack(side="left", padx=(6, 8))
        self.id_entry.bind("<Return>", self.load_id)
        self.id_entry.bind("<KP_Enter>", self.load_id)
        self.dot = tk.Canvas(line, width=22, height=22, highlightthickness=0)
        self.dot_item = self.dot.create_oval(2, 2, 20, 20, outline="#777")
        self.dot.pack(side="left")
        self.status_label = tk.Label(line, text="–", width=18, anchor="w")
        self.status_label.pack(side="left", padx=(4, 10))

        ttk.Button(line, text="Übernehmen", style="Accent.TButton",
                   command=self.commit).pack(side="left", padx=2)
        self.undo_btn = ttk.Button(line, text="Rückgängig", command=self.undo)
        self.undo_btn.pack(side="left", padx=2)
        ttk.Button(line, text="Speichern", command=self.save_book).pack(side="left", padx=2)
        ttk.Checkbutton(line, text="Daten Bett", style="Toolbutton", variable=self.bed,
                        command=self.toggle_bed).pack(side="left", padx=(10, 2))

        info = tk.Frame(line)
        info.pack(side="right")
        tk.Label(info, text="Bereich:").pack(side="left")
        self.area_label = tk.Label(info, text="–", font=("Segoe UI", 12, "bold"), width=6, anchor="w")
        self.area_label.pack(side="left", padx=(3, 12))
        tk.Label(info, text="Etage:").pack(side="left")
        self.floor_label = tk.Label(info, text="–", font=("Segoe UI", 12, "bold"), width=6, anchor="w")
        self.floor_label.pack(side="left", padx=(3, 0))

        ttk.Separator(outer).pack(fill="x", pady=3)

        # Zeile 3: Auswahlschiene
        self.rail = tk.Frame(outer)
        self.rail.pack(fill="x")
        self.rail_widgets = [tk.Label(self.rail, text="Anzeigen:", fg="#666")]
        self.check_widgets = {}
        for c in SELECTABLE:
            cb = ttk.Checkbutton(self.rail, variable=self.show_vars[c],
                                 command=lambda c=c: self.on_check(c))
            self.check_widgets[c] = cb
            self.rail_widgets.append(cb)

        ttk.Separator(outer).pack(fill="x", pady=3)

        # Zeile 4: Felder
        self.fields = tk.Frame(outer)
        self.fields.pack(fill="x")
        self.field_frames, self.field_labels, self.entries = {}, {}, {}
        for c in SELECTABLE:
            fr = tk.Frame(self.fields)
            lb = tk.Label(fr, anchor="w", fg="#555", font=("Segoe UI", 8))
            lb.pack(fill="x")
            en = tk.Entry(fr, textvariable=self.vars[c], width=FIELD_WIDTH.get(c, 12),
                          bg=COLOR_NORMAL, readonlybackground=COLOR_FORMULA, relief="solid", bd=1)
            en.pack(fill="x", ipady=2)
            en.bind("<Return>", self.commit)
            en.bind("<KP_Enter>", self.commit)
            en.bind("<Escape>", lambda e: self.focus_id())
            self.field_frames[c], self.field_labels[c], self.entries[c] = fr, lb, en
        self.apply_headers()

    def apply_headers(self):
        for c in SELECTABLE:
            name = self.headers.get(c, c)
            self.check_widgets[c].configure(text=f"{c} {name}")
            suffix = "  (Formel)" if self.formula[c] else ""
            self.field_labels[c].configure(text=f"{c} · {name}{suffix}")

    # ---------------------------------------------------------- Fenstergröße
    def place_initial(self):
        self.root.update_idletasks()
        left, top, right, bottom = work_area(self.root)
        width = right - left
        self.relayout(width - 16, fit=False)
        self.root.update_idletasks()
        h = self.root.winfo_reqheight()
        self.root.geometry(f"{width}x{h}+{left}+{bottom - h - 40}")
        self.root.update_idletasks()
        self.root.after(50, lambda: self.fit_height(anchor_bottom=True))

    def on_configure(self, event):
        if event.widget is not self.root:
            return
        if event.width != self._last_width:
            self._last_width = event.width
            self.schedule_layout()

    def schedule_layout(self):
        if self._layout_job:
            self.root.after_cancel(self._layout_job)
        self._layout_job = self.root.after(60, self._do_layout)

    def _do_layout(self):
        self._layout_job = None
        self.relayout(self.root.winfo_width() - 16)

    def relayout(self, width, fit=True):
        width = max(width, 300)
        self.root.update_idletasks()
        flow_place(self.rail, self.rail_widgets, width, gap_x=10)
        shown = [self.field_frames[c] for c in SELECTABLE if self.show_vars[c].get()]
        flow_place(self.fields, shown, width)
        if fit:
            self.root.update_idletasks()
            self.fit_height()

    def fit_height(self, anchor_bottom=False):
        root = self.root
        root.update_idletasks()
        h = root.winfo_reqheight()
        m = re.match(r"(\d+)x(\d+)([+-]-?\d+)([+-]-?\d+)", root.geometry())
        if not m:
            return
        w, cur_h, x, y = (int(v) for v in m.groups())
        if cur_h == h and not anchor_bottom:
            return
        _, top, _, bottom = work_area(root)
        border = max(0, root.winfo_rootx() - x)
        title = max(0, root.winfo_rooty() - y)
        outer_bottom = y + title + cur_h + border
        if anchor_bottom or outer_bottom >= bottom - 4 or y + title + h + border > bottom:
            y = max(top, bottom - h - title - border)
        root.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------- Felder & Status
    def message(self, text, error=False):
        self.msg_label.configure(text=text, fg="#b3261e" if error else "#2b5d2b")
        if error:
            self.root.bell()

    def focus_id(self):
        self.id_entry.focus_set()
        self.id_entry.select_range(0, "end")
        return "break"

    def toggle_topmost(self):
        self.root.attributes("-topmost", self.topmost.get())
        self.cfg["topmost"] = self.topmost.get()
        save_config(self.cfg)

    def on_check(self, c):
        self.auto_shown.discard(c)   # Manuelle Auswahl bleibt bestehen
        self.cfg["shown"] = [x for x in SELECTABLE
                             if self.show_vars[x].get() and x not in self.auto_shown]
        save_config(self.cfg)
        self.schedule_layout()

    def reset_auto(self):
        if self.auto_shown:
            for c in self.auto_shown:
                self.show_vars[c].set(False)
            self.auto_shown.clear()
            self.schedule_layout()

    def set_status(self, status):
        self.status = status
        color, text = STATUS[status]
        self.dot.itemconfigure(self.dot_item, fill=color)
        self.status_label.configure(text=text)
        if status in ("yellow", "red"):
            added = False
            for c in AUTO_COLS:
                if not self.show_vars[c].get():
                    self.show_vars[c].set(True)
                    self.auto_shown.add(c)
                    added = True
            if added:
                self.schedule_layout()

    def on_value_change(self, c):
        if self.loading:
            return
        self.paint_entry(c)
        if c in AUTO_COLS and self.row is not None:
            self.set_status(compute_status(*(self.vars[x].get() for x in AUTO_COLS)))

    def paint_entry(self, c):
        en = self.entries[c]
        if self.formula[c]:
            en.configure(state="readonly")
        else:
            en.configure(state="normal")
            changed = self.row is not None and self.vars[c].get() != self.orig[c]
            en.configure(bg=COLOR_CHANGED if changed else COLOR_NORMAL)

    def fill(self, data):
        self.loading = True
        try:
            for c in SELECTABLE:
                text, is_formula, shown = data[c] if data else ("", False, "")
                self.orig[c] = text
                self.formula[c] = is_formula
                self.entries[c].configure(state="normal")
                self.vars[c].set(shown if is_formula else text)
        finally:
            self.loading = False
        for c in SELECTABLE:
            self.paint_entry(c)
        self.apply_headers()
        if data:
            self.set_status(compute_status(*(self.vars[x].get() for x in AUTO_COLS)))
        else:
            self.set_status(None)
        self.schedule_layout()

    def changed_cols(self):
        if self.row is None:
            return []
        return [c for c in SELECTABLE
                if not self.formula[c] and self.vars[c].get() != self.orig[c]]

    # ---------------------------------------------------------- Excel-Aktionen
    def show_row(self, app, wb, ws, row):
        """Zeile A–S auswählen und als 2. Zeile des scrollbaren Bereichs anzeigen."""
        if app.ActiveWorkbook is None or app.ActiveWorkbook.Name != wb.Name:
            wb.Activate()
        if wb.ActiveSheet.Name != ws.Name:
            return   # z.B. "Daten Bett" aktiv – beim Zurückwechseln wird markiert
        win = app.ActiveWindow
        ws.Range(f"A{row}:{LAST_COL}{row}").Select()
        pane = win.Panes(win.Panes.Count) if win.FreezePanes else win
        pane.ScrollRow = max(FIRST_ROW, row - 1)
        pane.ScrollColumn = 1

    def locate(self, ws):
        if self.row and norm_id(ws.Cells(self.row, 1).Value2) == self.id_norm:
            return self.row
        rows = find_rows(ws, self.id_norm)
        if not rows:
            raise ExcelError(f"Inv.-Nr. {self.id_label} ist in der Liste nicht mehr vorhanden.")
        self.row = rows[0]
        return self.row

    def load_headers(self, ws):
        values = ws.Range(f"A{HEADER_ROW}:{LAST_COL}{HEADER_ROW}").Value2[0]
        for col, v in zip(all_cols(), values):
            if v not in (None, ""):
                self.headers[col] = clean_header(v)
        self.headers_loaded = True
        self.apply_headers()
        self.schedule_layout()

    @excel_action
    def load_id(self, event=None):
        raw = self.id_entry.get().strip()
        if not raw:
            return "break"
        pending = self.changed_cols()
        if pending and not messagebox.askyesno(
                "Ungespeicherte Änderungen",
                f"Die Änderungen bei Inv.-Nr. {self.id_label} wurden noch nicht übernommen.\n\n"
                "Verwerfen und neue Inventarnummer laden?", parent=self.root):
            return "break"

        app, wb, ws = get_book()
        if not self.headers_loaded:
            self.load_headers(ws)
        self.reset_auto()
        rows = find_rows(ws, norm_id(raw))
        if not rows:
            self.row = None
            self.id_norm = self.id_label = ""
            self.fill(None)
            self.message(f"Inv.-Nr. {raw} wurde nicht gefunden.", error=True)
            return self.focus_id()

        row = rows[0]
        id_value = ws.Cells(row, 1).Value2
        ws.Range(ID_CELL).Value2 = id_value          # gleicher Typ wie in Spalte A
        self.row, self.id_norm, self.id_label = row, norm_id(id_value), raw
        self.fill(read_row(ws, row))
        self.show_row(app, wb, ws, row)
        self.update_header_info(wb, ws)

        if len(rows) > 1:
            others = ", ".join(str(r) for r in rows)
            self.message(f"Achtung: Inv.-Nr. {raw} kommt {len(rows)}× vor (Zeilen {others}) "
                         f"– bearbeitet wird Zeile {row}.", error=True)
        else:
            self.message(f"Inv.-Nr. {raw} geladen (Zeile {row}).")
        self.root.after(30, self.focus_first_field)
        return "break"

    def focus_first_field(self):
        self.root.focus_force()
        for c in SELECTABLE:
            if self.show_vars[c].get() and not self.formula[c]:
                self.entries[c].focus_set()
                self.entries[c].icursor("end")
                return
        self.focus_id()

    @excel_action
    def commit(self, event=None):
        if self.row is None:
            self.message("Bitte zuerst eine Inventarnummer eingeben.", error=True)
            return self.focus_id()
        app, wb, ws = get_book()
        row = self.locate(ws)
        current = read_row(ws, row)
        changes = self.changed_cols()

        conflicts = [c for c in changes if current[c][0] != self.orig[c]]
        if conflicts:
            names = ", ".join(f"{c} ({self.headers.get(c, c)})" for c in conflicts)
            if not messagebox.askyesno(
                    "Zwischenzeitlich geändert",
                    f"Diese Felder wurden inzwischen in Excel geändert:\n{names}\n\n"
                    "Trotzdem mit den Werten aus dem Programm überschreiben?", parent=self.root):
                return "break"

        undo = {"row": row, "id_norm": self.id_norm, "label": self.id_label, "cells": []}
        for c in changes:
            undo["cells"].append((c, current[c][0]))
            ws.Range(f"{c}{row}").FormulaLocal = self.vars[c].get()
        undo["cells"].append((DONE_COL, current[DONE_COL][0]))
        ws.Range(f"{DONE_COL}{row}").Value2 = DONE_MARK
        self.undo_stack.append(undo)
        del self.undo_stack[:-50]
        self.update_undo_button()

        self.fill(read_row(ws, row))
        n = len(changes)
        self.message(f"Inv.-Nr. {self.id_label} übernommen: {n} Feld{'er' if n != 1 else ''} "
                     f"geändert, {DONE_COL} = „{DONE_MARK}“.")
        return self.focus_id()

    @excel_action
    def undo(self):
        if not self.undo_stack:
            self.message("Nichts zum Rückgängigmachen.")
            return
        u = self.undo_stack[-1]
        if not messagebox.askyesno(
                "Rückgängig",
                f"Letzte Übernahme für Inv.-Nr. {u['label']} rückgängig machen?\n"
                f"({len(u['cells'])} Zellen werden auf den alten Stand gesetzt.)", parent=self.root):
            return
        app, wb, ws = get_book()
        row = u["row"]
        if norm_id(ws.Cells(row, 1).Value2) != u["id_norm"]:
            rows = find_rows(ws, u["id_norm"])
            if not rows:
                raise ExcelError(f"Inv.-Nr. {u['label']} ist nicht mehr in der Liste.")
            row = rows[0]
        for c, old in u["cells"]:
            ws.Range(f"{c}{row}").FormulaLocal = old
        self.undo_stack.pop()
        self.update_undo_button()
        if self.id_norm == u["id_norm"]:
            self.row = row
            self.fill(read_row(ws, row))
        self.message(f"Übernahme für Inv.-Nr. {u['label']} rückgängig gemacht.")
        self.focus_id()

    def update_undo_button(self):
        self.undo_btn.configure(state="normal" if self.undo_stack else "disabled")

    @excel_action
    def save_book(self):
        app, wb, ws = get_book()
        wb.Save()
        self.message(f"„{wb.Name}“ gespeichert.")
        return "break"

    def toggle_bed(self):
        wanted = self.bed.get()
        try:
            app, wb, ws = get_book()
            if app.ActiveWorkbook is None or app.ActiveWorkbook.Name != wb.Name:
                wb.Activate()
            if wanted:
                wb.Worksheets(BED_SHEET).Activate()
                self.message(f"Blatt „{BED_SHEET}“ angezeigt.")
            else:
                ws.Activate()
                if self.row is not None:
                    self.show_row(app, wb, ws, self.locate(ws))
                self.message(f"Zurück zu „{REF_SHEET}“.")
        except ExcelError as e:
            self.bed.set(not wanted)
            self.message(str(e), error=True)
        except pywintypes.com_error as e:
            self.bed.set(not wanted)
            self.message(com_message(e), error=True)
        self.focus_id()

    def update_header_info(self, wb, ws):
        self.file_label.configure(text=wb.Name)
        self.area_label.configure(text=str(ws.Range(AREA_CELL).Text or "–"))
        self.floor_label.configure(text=str(ws.Range(FLOOR_CELL).Text or "–"))

    def poll(self):
        """Dateiname, Bereich und Etage regelmäßig aktualisieren (nur lesend)."""
        try:
            app, wb, ws = get_book()
            if not self.headers_loaded:
                self.load_headers(ws)
            self.update_header_info(wb, ws)
            try:
                self.bed.set(wb.ActiveSheet.Name == BED_SHEET)
            except pywintypes.com_error:
                pass
        except ExcelError:
            self.file_label.configure(text="– keine Liste geöffnet –")
        except Exception:
            pass   # Excel beschäftigt – beim nächsten Mal wieder versuchen
        self.root.after(2000, self.poll)

    # ---------------------------------------------------------------- Updates
    def check_updates(self, silent):
        self.update_btn.configure(state="disabled", text="Suche Updates …")
        self.update_result = None

        def worker():
            try:
                url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
                req = urllib.request.Request(url, headers={
                    "Accept": "application/vnd.github+json", "User-Agent": "InventarEditor"})
                with urllib.request.urlopen(req, timeout=10) as r:
                    data = json.load(r)
                asset = next((a for a in data.get("assets", []) if a.get("name") == ASSET_NAME), None)
                self.update_result = ("ok", {
                    "tag": data.get("tag_name", ""), "notes": data.get("body") or "",
                    "url": asset and asset.get("browser_download_url"),
                    "page": data.get("html_url")})
            except urllib.error.HTTPError as e:
                self.update_result = ("none", None) if e.code == 404 else ("error", str(e))
            except Exception as e:  # noqa: BLE001
                self.update_result = ("error", str(e))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(200, lambda: self.wait_update_check(silent))

    def wait_update_check(self, silent):
        if self.update_result is None:
            self.root.after(200, lambda: self.wait_update_check(silent))
            return
        kind, info = self.update_result
        self.update_btn.configure(state="normal")
        if kind == "ok" and is_newer(info["tag"], VERSION) and info["url"]:
            self.update_info = info
            self.update_btn.configure(text=f"Update {info['tag']} verfügbar!", style="Update.TButton")
            if not silent:
                self.ask_update()
            return
        self.update_info = None
        self.update_btn.configure(text="Nach Updates suchen", style="TButton")
        if not silent:
            if kind == "error":
                messagebox.showwarning("Update", f"Update-Prüfung fehlgeschlagen:\n{info}",
                                       parent=self.root)
            else:
                messagebox.showinfo("Update", f"Du hast die neueste Version ({VERSION}).",
                                    parent=self.root)

    def on_update_button(self):
        if self.update_info:
            self.ask_update()
        else:
            self.check_updates(silent=False)

    def ask_update(self):
        info = self.update_info
        notes = info["notes"].strip()
        text = (f"Installiert: {VERSION}\nNeu: {info['tag']}\n\n"
                + (f"Änderungen:\n{notes[:1200]}\n\n" if notes else "")
                + "Jetzt aktualisieren? Das Programm wird dazu kurz neu gestartet.\n"
                  "Die Excel-Liste bleibt geöffnet.")
        if self.changed_cols():
            text += "\n\nAchtung: Nicht übernommene Änderungen im Programm gehen verloren."
        if not messagebox.askyesno("Update verfügbar", text, parent=self.root):
            return
        if not getattr(sys, "frozen", False):
            webbrowser.open(info["page"])
            return
        self.download_update(info)

    def download_update(self, info):
        exe = sys.executable
        new = exe + ".new"
        self.update_btn.configure(state="disabled", text="Lade Update …")
        self.download_result = None

        def worker():
            try:
                req = urllib.request.Request(info["url"], headers={"User-Agent": "InventarEditor"})
                with urllib.request.urlopen(req, timeout=120) as r, open(new, "wb") as f:
                    while True:
                        chunk = r.read(1 << 16)
                        if not chunk:
                            break
                        f.write(chunk)
                if os.path.getsize(new) < 1_000_000:
                    raise RuntimeError("Download unvollständig.")
                self.download_result = ("ok", None)
            except Exception as e:  # noqa: BLE001
                try:
                    os.remove(new)
                except OSError:
                    pass
                self.download_result = ("error", str(e))

        def wait():
            if self.download_result is None:
                self.root.after(200, wait)
                return
            kind, err = self.download_result
            if kind != "ok":
                self.update_btn.configure(state="normal", text=f"Update {info['tag']} verfügbar!")
                messagebox.showerror(
                    "Update", f"Update konnte nicht geladen werden:\n{err}\n\n"
                    "Liegt das Programm in einem Ordner, in den du schreiben darfst "
                    "(z.B. Dokumente oder Desktop)?", parent=self.root)
                return
            self.restart_with(exe, new)

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(200, wait)

    def restart_with(self, exe, new):
        bat = os.path.join(tempfile.gettempdir(), "inventar_editor_update.bat")
        script = (
            "@chcp 65001 >nul\r\n"
            "@echo off\r\n"
            f'set "EXE={exe}"\r\n'
            f'set "NEW={new}"\r\n'
            "set /a n=0\r\n"
            ":tausch\r\n"
            'move /Y "%NEW%" "%EXE%" >nul 2>&1\r\n'
            "if not errorlevel 1 goto start\r\n"
            "set /a n+=1\r\n"
            "if %n% geq 40 goto start\r\n"
            "ping -n 2 127.0.0.1 >nul\r\n"
            "goto tausch\r\n"
            ":start\r\n"
            'start "" "%EXE%"\r\n'
            'del "%~f0"\r\n'
        )
        with open(bat, "w", encoding="utf-8", newline="") as f:
            f.write(script)
        subprocess.Popen(["cmd", "/c", bat], creationflags=0x08000000 | 0x00000200,
                         close_fds=True)
        self.root.destroy()
        sys.exit(0)


def main():
    try:   # scharfe Schrift auf hochauflösenden Bildschirmen
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    if not HAVE_WIN32:
        root.withdraw()
        messagebox.showerror("Inventar-Editor",
                             "Dieses Programm läuft nur unter Windows mit Microsoft Excel.")
        return
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
