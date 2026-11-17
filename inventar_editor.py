"""Inventar-Editor

Eigenständiges Programm zum Bearbeiten der geöffneten Excel-Prüfliste.
Es verbindet sich per COM mit der laufenden Excel-Instanz; die Excel-Datei
selbst enthält dafür keinerlei Makros oder Add-ins.
"""

import datetime
import json
import math
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
import tkinter.font as tkfont
from tkinter import messagebox, ttk

from logo_data import LOGO_PNG_B64
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
FIELD_TITLE = {"R": "MK"}          # abweichende Überschrift im Eingabefeld
FIELD_WIDTH = {
    "B": 11, "C": 14, "D": 20, "E": 16, "F": 18, "G": 6, "H": 6,
    "J": 8, "O": 10, "P": 10, "Q": 36, "R": 5, "S": 8,
}

STATUS = {
    None: ("#b5b5b5", "–"),
    "green": ("#2e9e44", "ohne Mängel"),
    "yellow": ("#f2c200", "Mangel ohne Klasse"),
    "red": ("#d93025", "Mängelklasse"),
}

COLOR_CHANGED = "#fff1b8"
COLOR_FORMULA = "#ececec"
COLOR_NORMAL = "#ffffff"
FONT = "Segoe UI"

XL_UP = -4162
BUSY_CODES = {-2147418111, -2147417846, -2146777998}   # Excel beschäftigt
NO_EXCEL_CODES = {-2147221021}                          # kein Excel offen

NUMBER_RE = re.compile(r"-?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?")
# Texte, die Excel sonst als Zahl, Datum oder Formel deuten würde
RISKY_RE = re.compile(r"[=+\-@]|[\d\s.,:/%\-]+$")


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


def parse_number(text):
    """Deutsche Zahl ("1.234,5", "461") -> float, sonst None."""
    t = text.strip()
    if not NUMBER_RE.fullmatch(t):
        return None
    return float(t.replace(".", "").replace(",", "."))


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
        return "Excel ist beschäftigt – Zelleingabe in Excel mit Enter oder Esc beenden."
    if code in NO_EXCEL_CODES:
        return "Excel ist nicht geöffnet."
    detail = ""
    try:
        detail = err.args[2][2] or ""
    except Exception:
        pass
    return f"Excel-Fehler: {detail or err}"


def excel_action(fn):
    """Fängt Excel-Fehler ab und zeigt sie in der Meldungszeile an."""
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
def excel_app():
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except pywintypes.com_error:
        raise ExcelError("Excel ist nicht geöffnet.")


def list_books(app):
    """Alle offenen Mappen mit dem Referenzblatt: [(Name, Mappe, Blatt)]."""
    result = []
    for i in range(1, app.Workbooks.Count + 1):
        wb = app.Workbooks(i)
        try:
            result.append((wb.Name, wb, wb.Worksheets(REF_SHEET)))
        except pywintypes.com_error:
            continue
    return result


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


def snapshot(rng):
    """Inhalt und Zahlenformat einer Zelle für "Rückgängig"."""
    return rng.Value2, rng.NumberFormat


def restore_number_format(rng, fmt):
    if fmt is not None:
        try:
            if rng.NumberFormat != fmt:
                rng.NumberFormat = fmt
        except pywintypes.com_error:
            pass


def parse_date(text):
    """"17.11.2026" / "17.11.26" -> Excel-Datumszahl, sonst None."""
    m = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})", text.strip())
    if not m:
        return None
    d, mo, y = (int(x) for x in m.groups())
    if y < 100:
        y += 2000 if y < 50 else 1900
    try:
        return float((datetime.date(y, mo, d) - datetime.date(1899, 12, 30)).days)
    except ValueError:
        return None


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def put_value(rng, value, fmt):
    """Schreibt Zahl oder Text unverändert und stellt das Zellformat wieder her.

    Excel speichert in einer als „Text“ formatierten Zelle jede Eingabe als Text.
    Für Zahlen wird das Format deshalb kurz auf „Standard“ gesetzt und danach
    zurückgestellt – die Zahl bleibt eine Zahl, das Format bleibt wie vorher.
    """
    if value is None or value == "":
        rng.ClearContents()
    elif is_number(value):
        if fmt == "@":
            rng.NumberFormat = "General"
        rng.Value2 = float(value)
    elif fmt == "@":
        rng.Value2 = value
    elif RISKY_RE.match(value):
        rng.Value2 = "'" + value                        # als Text erzwingen
    else:
        rng.Value2 = value
    restore_number_format(rng, fmt)


def column_holds_numbers(ws, col, row):
    """Für leere Zellen: Was steht in den Nachbarzeilen – Zahlen oder Text?"""
    for r in (row - 1, row + 1, row - 2, row + 2):
        if r < FIRST_ROW:
            continue
        v = ws.Range(f"{col}{r}").Value2
        if v not in (None, ""):
            return is_number(v)
    return None


def write_cell(ws, col, row, text):
    """Schreibt einen Wert so, dass Datentyp und Zellformat erhalten bleiben.

    - War vorher eine Zahl/ein Datum in der Zelle, wird wieder eine Zahl/ein
      Datum geschrieben – unabhängig vom Zellformat.
    - War vorher Text in der Zelle, bleibt es Text.
    - Ist die Zelle leer, richtet sich der Typ nach den Nachbarzeilen.
    - Das Zellformat wird danach immer wiederhergestellt.
    """
    rng = ws.Range(f"{col}{row}")
    old, fmt = snapshot(rng)
    t = text.strip()
    if not t:
        put_value(rng, None, fmt)
        return
    if old in (None, ""):
        want_number = column_holds_numbers(ws, col, row)
        if want_number is None:
            want_number = fmt != "@"
    else:
        want_number = is_number(old)
    if want_number:
        number = parse_date(t)
        if number is None:
            number = parse_number(t)
        if number is not None:
            put_value(rng, number, fmt)
            return
    put_value(rng, text, fmt)


def restore_cell(rng, saved):
    value, fmt = saved
    put_value(rng, value, fmt)


# --------------------------------------------------------------------------
# Zweizeiliges Eingabefeld
# --------------------------------------------------------------------------
class Field:
    """Zweizeiliges Textfeld: einzeiliger Text steht vertikal mittig,
    längerer Text wird automatisch umgebrochen."""

    def __init__(self, parent, var, width, on_return, on_escape):
        self.var = var
        self.font = tkfont.Font(family=FONT, size=10)
        self.text = tk.Text(parent, height=2, width=width, wrap="word", font=self.font,
                            relief="solid", bd=1, padx=4, pady=1, undo=False,
                            highlightthickness=1, highlightcolor="#3c7fd1",
                            highlightbackground="#c8c8c8")
        self.line_h = self.font.metrics("linespace")
        self._busy = False
        var.trace_add("write", self._from_var)
        self.text.bind("<<Modified>>", self._from_text)
        self.text.bind("<Configure>", lambda e: self.valign())
        for key in ("<Return>", "<KP_Enter>"):
            self.text.bind(key, on_return)
        self.text.bind("<Escape>", lambda e: on_escape())
        self.text.bind("<Tab>", self._next)
        self.text.bind("<Shift-Tab>", self._prev)
        self.text.bind("<ISO_Left_Tab>", self._prev)

    def _next(self, event):
        event.widget.tk_focusNext().focus_set()
        return "break"

    def _prev(self, event):
        event.widget.tk_focusPrev().focus_set()
        return "break"

    def _from_var(self, *_):
        if self._busy:
            return
        self._busy = True
        state = self.text.cget("state")
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", self.var.get())
        self.text.configure(state=state)
        self.text.edit_modified(False)
        self._busy = False
        self.valign()

    def _from_text(self, _event=None):
        if not self.text.edit_modified():
            return
        self.text.edit_modified(False)
        if self._busy:
            return
        content = self.text.get("1.0", "end-1c").replace("\r", "").replace("\n", " ")
        self._busy = True
        self.var.set(content)
        self._busy = False
        self.valign()

    def valign(self):
        try:
            n = self.text.count("1.0", "end-1c", "displaylines")
        except tk.TclError:
            return
        if isinstance(n, tuple):
            n = n[0]
        pad = self.line_h // 2 if not n else 0
        self.text.tag_add("v", "1.0", "end")
        self.text.tag_configure("v", spacing1=pad)

    def set_look(self, readonly, bg):
        self.text.configure(state="disabled" if readonly else "normal", bg=bg,
                            takefocus=0 if readonly else 1)

    def focus(self):
        self.text.focus_set()
        self.text.mark_set("insert", "end-1c")


# --------------------------------------------------------------------------
# Hellblauer Rahmen über der Excel-Zeile
# --------------------------------------------------------------------------
class RowMarker:
    """Rahmenloses, durchklickbares Fenster, das nur einen hellblauen Rahmen zeigt.
    Es liegt über Excel – an der Tabelle selbst wird nichts verändert."""

    KEY = "#ff00fe"      # wird durchsichtig
    COLOR = "#4da3ff"
    THICK = 3

    def __init__(self, root):
        self.ok = False
        self.shown = None
        try:
            self.win = tk.Toplevel(root)
            self.win.withdraw()
            self.win.overrideredirect(True)
            self.win.configure(bg=self.KEY)
            self.win.attributes("-transparentcolor", self.KEY)
            self.win.attributes("-topmost", True)
            self.canvas = tk.Canvas(self.win, bg=self.KEY, highlightthickness=0, bd=0)
            self.canvas.pack(fill="both", expand=True)
            self.ok = True
            root.after(100, self._click_through)
        except tk.TclError:
            pass   # nur unter Windows verfügbar

    def _click_through(self):
        try:
            import ctypes
            self.win.update_idletasks()
            user32 = ctypes.windll.user32
            hwnd = user32.GetParent(self.win.winfo_id()) or self.win.winfo_id()
            ex = user32.GetWindowLongW(hwnd, -20)
            # LAYERED | TRANSPARENT | TOOLWINDOW | NOACTIVATE
            user32.SetWindowLongW(hwnd, -20, ex | 0x80000 | 0x20 | 0x80 | 0x08000000)
        except Exception:
            pass

    def show(self, x, y, w, h):
        if not self.ok:
            return
        t = self.THICK
        rect = (x - t, y - t, w + 2 * t, h + 2 * t)
        if rect != self.shown:
            self.shown = rect
            self.win.geometry(f"{rect[2]}x{rect[3]}+{rect[0]}+{rect[1]}")
            self.canvas.delete("all")
            self.canvas.create_rectangle(t / 2, t / 2, rect[2] - t / 2, rect[3] - t / 2,
                                         outline=self.COLOR, width=t)
            self.win.deiconify()
            self.win.attributes("-topmost", True)

    def hide(self):
        if self.ok and self.shown is not None:
            self.shown = None
            self.win.withdraw()


def foreground_window():
    try:
        import ctypes
        return ctypes.windll.user32.GetForegroundWindow()
    except Exception:
        return None


# --------------------------------------------------------------------------
# Programmfenster
# --------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.headers = dict(FALLBACK_HEADERS)
        self.headers_loaded = False
        self.selected_book = None

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
        self.book_var = tk.StringVar()
        self.update_result = None
        self.update_info = None
        self._layout_job = None
        self._last_width = 0

        self.logo_full = tk.PhotoImage(data=LOGO_PNG_B64)
        root.iconphoto(True, self.logo_full.subsample(4))

        self.marker = RowMarker(root)
        self._ctx = None         # (app, wb, ws) der zuletzt benutzten Liste

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
        self.track_marker()
        self.check_updates(silent=True)

    # ---------------------------------------------------------------- Aufbau
    def build_ui(self):
        root = self.root
        root.option_add("*Font", (FONT, 10))
        style = ttk.Style()
        style.configure("Accent.TButton", font=(FONT, 10, "bold"))
        style.configure("Update.TButton", foreground="#b3261e", font=(FONT, 9, "bold"))
        style.configure("Bold.TCheckbutton", font=(FONT, 10, "bold"))

        outer = tk.Frame(root, padx=8, pady=5)
        outer.pack(fill="both", expand=True)

        # Kopf: Logo links, daneben zwei Zeilen
        head = tk.Frame(outer)
        head.pack(fill="x")
        self.logo_label = tk.Label(head, bd=0)
        self.logo_label.pack(side="left", padx=(0, 10))
        rows = tk.Frame(head)
        rows.pack(side="left", fill="x", expand=True)
        self.head_rows = rows

        # Zeile 1: Liste | Meldung | Bereich, Etage | Vordergrund, Version, Update
        top = tk.Frame(rows)
        top.pack(fill="x")
        self.book_combo = ttk.Combobox(top, textvariable=self.book_var, state="readonly",
                                       width=34, font=(FONT, 11, "bold"))
        self.book_combo.pack(side="left")
        self.book_combo.bind("<<ComboboxSelected>>", self.on_book_selected)
        right = tk.Frame(top)
        right.pack(side="right")
        ttk.Checkbutton(right, text="Immer im Vordergrund", variable=self.topmost,
                        command=self.toggle_topmost).pack(side="left", padx=(0, 12))
        tk.Label(right, text=VERSION, fg="#666").pack(side="left", padx=(0, 6))
        self.update_btn = ttk.Button(right, text="Nach Updates suchen",
                                     command=self.on_update_button)
        self.update_btn.pack(side="left")
        self.msg_label = tk.Label(top, text="", anchor="w", fg="#444", width=58)
        self.msg_label.pack(side="left", padx=(14, 8))
        info = tk.Frame(top)
        info.pack(side="left")
        tk.Label(info, text="Bereich:").pack(side="left")
        self.area_label = tk.Label(info, text="–", font=(FONT, 12, "bold"), width=5, anchor="w")
        self.area_label.pack(side="left", padx=(3, 10))
        tk.Label(info, text="Etage:").pack(side="left")
        self.floor_label = tk.Label(info, text="–", font=(FONT, 12, "bold"), width=5, anchor="w")
        self.floor_label.pack(side="left", padx=(3, 0))

        # Zeile 2: Inv.-Nr., Status, Aktionen
        line = tk.Frame(rows)
        line.pack(fill="x", pady=(5, 0))
        tk.Label(line, text="Inventar-Nr.:", font=(FONT, 11, "bold")).pack(side="left")
        self.id_entry = tk.Entry(line, width=10, font=(FONT, 14, "bold"), justify="center")
        self.id_entry.pack(side="left", padx=(6, 8))
        self.id_entry.bind("<Return>", self.load_id)
        self.id_entry.bind("<KP_Enter>", self.load_id)
        self.dot = tk.Canvas(line, width=22, height=22, highlightthickness=0)
        self.dot_item = self.dot.create_oval(2, 2, 20, 20, outline="#777")
        self.dot.pack(side="left")
        self.status_label = tk.Label(line, text="–", width=18, anchor="w")
        self.status_label.pack(side="left", padx=(4, 10))
        group = tk.Frame(line, highlightthickness=1, highlightbackground="#9a9a9a",
                         highlightcolor="#9a9a9a", padx=4, pady=3)
        group.pack(side="left")
        ttk.Button(group, text="Übernehmen", style="Accent.TButton",
                   command=self.commit).pack(side="left", padx=2)
        self.undo_btn = ttk.Button(group, text="Rückgängig", command=self.undo)
        self.undo_btn.pack(side="left", padx=2)
        ttk.Button(group, text="Speichern", command=self.save_book).pack(side="left", padx=2)

        self.bed_btn = tk.Button(line, text=BED_SHEET, font=(FONT, 10), relief="raised", bd=1,
                                 padx=10, pady=3, cursor="hand2", command=self.on_bed_button)
        self.bed_btn.pack(side="left", padx=(12, 2))
        self.bed_default = (self.bed_btn.cget("bg"), self.bed_btn.cget("fg"),
                            self.bed_btn.cget("activebackground"))
        self.bed.trace_add("write", lambda *_: self.paint_bed_button())

        ttk.Separator(outer).pack(fill="x", pady=4)

        # Zeile 3: Auswahlschiene
        self.rail = tk.Frame(outer)
        self.rail.pack(fill="x")
        self.rail_widgets = [tk.Label(self.rail, text="Anzeigen:", fg="#666")]
        self.check_names = {}
        for c in SELECTABLE:
            item = tk.Frame(self.rail)
            cb = ttk.Checkbutton(item, variable=self.show_vars[c],
                                 command=lambda c=c: self.on_check(c))
            cb.pack(side="left")
            letter = tk.Label(item, text=c, font=(FONT, 10, "bold"), padx=0)
            letter.pack(side="left")
            name = tk.Label(item, font=(FONT, 10), padx=0)
            name.pack(side="left")
            for w in (letter, name):   # Klick auf den Text schaltet den Haken
                w.bind("<Button-1>", lambda e, cb=cb: cb.invoke())
            self.check_names[c] = name
            self.rail_widgets.append(item)

        ttk.Separator(outer).pack(fill="x", pady=4)

        # Zeile 4: Felder
        self.fields = tk.Frame(outer)
        self.fields.pack(fill="x")
        self.field_frames, self.field_labels, self.entries = {}, {}, {}
        for c in SELECTABLE:
            fr = tk.Frame(self.fields)
            title = tk.Frame(fr)
            title.pack(fill="x")
            tk.Label(title, text=c, fg="#333", font=(FONT, 9, "bold"), padx=0).pack(side="left")
            lb = tk.Label(title, anchor="w", fg="#333", font=(FONT, 9), padx=0)
            lb.pack(side="left")
            field = Field(fr, self.vars[c], FIELD_WIDTH.get(c, 12), self.commit, self.focus_id)
            field.text.pack(fill="x")
            self.field_frames[c], self.field_labels[c], self.entries[c] = fr, lb, field
        self.apply_headers()

    def apply_headers(self):
        for c in SELECTABLE:
            name = self.headers.get(c, c)
            self.check_names[c].configure(text=f" · {name}")
            suffix = "  (Formel)" if self.formula[c] else ""
            self.field_labels[c].configure(text=f" · {FIELD_TITLE.get(c, name)}{suffix}")

    def size_logo(self):
        """Logo so groß wie die beiden Kopfzeilen."""
        self.root.update_idletasks()
        h = max(24, self.head_rows.winfo_reqheight())
        factor = max(1, math.ceil(self.logo_full.height() / h))
        self.logo_small = self.logo_full.subsample(factor)
        self.logo_label.configure(image=self.logo_small)

    # ---------------------------------------------------------- Fenstergröße
    def place_initial(self):
        self.size_logo()
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
        flow_place(self.rail, self.rail_widgets, width, gap_x=12)
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
        if status == "red":
            text = f"{text} {self.vars['R'].get().strip()}"
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
        if self.formula[c]:
            self.entries[c].set_look(True, COLOR_FORMULA)
        else:
            changed = self.row is not None and self.vars[c].get() != self.orig[c]
            self.entries[c].set_look(False, COLOR_CHANGED if changed else COLOR_NORMAL)

    def fill(self, data):
        self.loading = True
        try:
            for c in SELECTABLE:
                text, is_formula, shown = data[c] if data else ("", False, "")
                self.orig[c] = text
                self.formula[c] = is_formula
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

    def clear_loaded(self):
        self.row = None
        self.id_norm = self.id_label = ""
        self.reset_auto()
        self.fill(None)

    # ---------------------------------------------------------- Excel-Aktionen
    def get_book(self):
        """Die im Auswahlmenü gewählte Liste (oder die aktive / erste offene)."""
        app = excel_app()
        books = list_books(app)
        if not books:
            raise ExcelError(f"Keine geöffnete Liste mit dem Blatt „{REF_SHEET}“ gefunden.")
        names = [n for n, _, _ in books]
        if self.selected_book not in names:
            active = None
            try:
                active = app.ActiveWorkbook.Name if app.ActiveWorkbook is not None else None
            except pywintypes.com_error:
                pass
            self.selected_book = active if active in names else names[0]
            self.headers_loaded = False
        if list(self.book_combo["values"]) != names:
            self.book_combo["values"] = names
        if self.book_var.get() != self.selected_book:
            self.book_var.set(self.selected_book)
        _, wb, ws = books[names.index(self.selected_book)]
        if not self.headers_loaded:
            self.load_headers(ws)
        self._ctx = (app, wb, ws)
        return app, wb, ws

    @excel_action
    def on_book_selected(self, event=None):
        name = self.book_var.get()
        if name == self.selected_book:
            return self.focus_id()
        if self.changed_cols() and not messagebox.askyesno(
                "Ungespeicherte Änderungen",
                f"Die Änderungen bei Inv.-Nr. {self.id_label} wurden noch nicht übernommen.\n\n"
                "Verwerfen und Liste wechseln?", parent=self.root):
            self.book_var.set(self.selected_book)
            return self.focus_id()
        self.selected_book = name
        self.headers_loaded = False
        self.clear_loaded()
        self.undo_stack.clear()
        self.update_undo_button()
        app, wb, ws = self.get_book()
        wb.Activate()
        self.bed.set(False)
        self.update_header_info(wb, ws)
        self.message(f"Liste „{name}“ ausgewählt.")
        return self.focus_id()

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
        self.headers = dict(FALLBACK_HEADERS)
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
        if self.changed_cols() and not messagebox.askyesno(
                "Ungespeicherte Änderungen",
                f"Die Änderungen bei Inv.-Nr. {self.id_label} wurden noch nicht übernommen.\n\n"
                "Verwerfen und neue Inventarnummer laden?", parent=self.root):
            return "break"

        app, wb, ws = self.get_book()
        self.reset_auto()
        rows = find_rows(ws, norm_id(raw))
        if not rows:
            self.clear_loaded()
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
            self.message(f"Inv.-Nr. {raw} kommt {len(rows)}× vor (Zeilen {others}) – "
                         f"bearbeitet: Zeile {row}.", error=True)
        else:
            self.message(f"Inv.-Nr. {raw} geladen (Zeile {row}).")
        self.root.after(30, self.focus_first_field)
        return "break"

    def focus_first_field(self):
        self.root.focus_force()
        for c in SELECTABLE:
            if self.show_vars[c].get() and not self.formula[c]:
                self.entries[c].focus()
                return
        self.focus_id()

    @excel_action
    def commit(self, event=None):
        if self.row is None:
            self.message("Bitte zuerst eine Inventarnummer eingeben.", error=True)
            return self.focus_id()
        app, wb, ws = self.get_book()
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
        for c in changes + [DONE_COL]:
            rng = ws.Range(f"{c}{row}")
            undo["cells"].append((c, snapshot(rng)))
            write_cell(ws, c, row, DONE_MARK if c == DONE_COL else self.vars[c].get())
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
        app, wb, ws = self.get_book()
        row = u["row"]
        if norm_id(ws.Cells(row, 1).Value2) != u["id_norm"]:
            rows = find_rows(ws, u["id_norm"])
            if not rows:
                raise ExcelError(f"Inv.-Nr. {u['label']} ist nicht mehr in der Liste.")
            row = rows[0]
        for c, saved in u["cells"]:
            restore_cell(ws.Range(f"{c}{row}"), saved)
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
        app, wb, ws = self.get_book()
        wb.Save()
        self.message(f"„{wb.Name}“ gespeichert.")
        return "break"

    def on_bed_button(self):
        self.bed.set(not self.bed.get())
        self.toggle_bed()

    def paint_bed_button(self):
        if self.bed.get():
            self.bed_btn.configure(bg="#2e9e44", fg="white", activebackground="#25803a",
                                   activeforeground="white", font=(FONT, 10, "bold"),
                                   relief="sunken")
        else:
            bg, fg, abg = self.bed_default
            self.bed_btn.configure(bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                                   font=(FONT, 10), relief="raised")

    def toggle_bed(self):
        wanted = self.bed.get()
        try:
            app, wb, ws = self.get_book()
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
        self.area_label.configure(text=str(ws.Range(AREA_CELL).Text or "–"))
        self.floor_label.configure(text=str(ws.Range(FLOOR_CELL).Text or "–"))

    def row_rect(self):
        """Bildschirmposition der Zeile A–S, wenn sie in Excel sichtbar ist."""
        if self.row is None or self._ctx is None:
            return None
        app, wb, ws = self._ctx
        if app.ActiveWorkbook is None or app.ActiveWorkbook.Name != wb.Name:
            return None
        if wb.ActiveSheet.Name != ws.Name:
            return None
        win = app.ActiveWindow
        if win.WindowState == -4140:          # minimiert
            return None
        fg = foreground_window()
        if fg is not None:
            allowed = {int(self.root.wm_frame(), 16), app.Hwnd}
            try:
                allowed.add(win.Hwnd)
            except Exception:
                pass
            if fg not in allowed:
                return None
        pane = win.Panes(win.Panes.Count) if win.FreezePanes else win.ActivePane
        visible = pane.VisibleRange
        first = visible.Row
        last = first + visible.Rows.Count - 1
        if not first <= self.row <= last - 1:   # letzte Zeile ist meist nur angeschnitten
            return None
        rng = ws.Range(f"A{self.row}:{LAST_COL}{self.row}")
        x1 = pane.PointsToScreenPixelsX(rng.Left)
        x2 = pane.PointsToScreenPixelsX(rng.Left + rng.Width)
        y1 = pane.PointsToScreenPixelsY(rng.Top)
        y2 = pane.PointsToScreenPixelsY(rng.Top + rng.Height)
        if x2 <= x1 or y2 <= y1:
            return None
        return x1, y1, x2 - x1, y2 - y1

    def track_marker(self):
        """Rahmen laufend nachführen (Scrollen, Zoomen, Fenster verschieben)."""
        if self.marker.ok:
            try:
                rect = self.row_rect()
            except Exception:
                rect = None
            if rect:
                self.marker.show(*rect)
            else:
                self.marker.hide()
        self.root.after(250, self.track_marker)

    def poll(self):
        """Listenauswahl, Bereich und Etage regelmäßig aktualisieren (nur lesend)."""
        try:
            app, wb, ws = self.get_book()
            self.update_header_info(wb, ws)
            try:
                self.bed.set(wb.ActiveSheet.Name == BED_SHEET)
            except pywintypes.com_error:
                pass
        except ExcelError:
            self.book_combo["values"] = []
            self.book_var.set("– keine Liste geöffnet –")
            self.selected_book = None
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
        # Saubere Umgebung: Die neue Instanz darf nicht die Einstellungen der alten
        # (temporärer Entpack-Ordner _MEI…) erben, sonst „Failed to load Python DLL“.
        env = {k: v for k, v in os.environ.items()
               if not k.upper().startswith(("_PYI", "_MEIPASS"))
               and k.upper() not in ("TCL_LIBRARY", "TK_LIBRARY")}
        env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            env["PATH"] = os.pathsep.join(
                p for p in env.get("PATH", "").split(os.pathsep)
                if not os.path.normcase(p).startswith(os.path.normcase(meipass)))
        subprocess.Popen(["cmd", "/c", bat], creationflags=0x08000000 | 0x00000200,
                         close_fds=True, env=env, cwd=os.path.dirname(exe))
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
