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
from tkinter import filedialog, messagebox, ttk

from logo_data import LOGO_PNG_B64, PRINT_PNG_B64, SAVE_PNG_B64
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
BED_SHEET = "PRO Pflegebett"       # Prüfprotokoll Pflegebett (Typenliste, Haken)
# Kacheln für Prüfprotokolle: (Beschriftung, Tabellenblatt)
# Druckersymbol links = als PDF speichern, Text = zum Blatt wechseln
SHEET_TILES = [
    ("Pflegebett", "PRO Pflegebett"),
    ("Med.Produkte", "PRO Med.-Produkte"),
    ("O2", "PRO O2-Konz"),
]
# Auswahl "Geräteart" im Typ-Fenster -> C (Geräteart), dazu H = "CE"
DEVICE_TYPES = ["Pflegebett", "O2-Konzentrator", "Absauggerät", "O2-Armatur",
                "Aufstehhilfe", "Patientenlifter", "Sitzwaage"]
HEADER_ROW = 6                     # Überschriften
FIRST_ROW = 7                      # erste Datenzeile
LAST_COL = "S"                     # Zeile wird von A bis S markiert
ID_CELL = "A2"                     # hierhin wird die aktuelle Inv.-Nr. geschrieben
AREA_CELL = "B2"                   # Bereich
FLOOR_CELL = "C2"                  # Etage
DONE_COL = "L"                     # bekommt beim Übernehmen ein "x"
DONE_MARK = "x"

# Neues Bett: Typenliste aus "PRO Pflegebett" (Typ in P, Hersteller in O, ab Zeile 7)
BED_TYPE_COL = "P"
BED_MAKER_COL = "O"
BED_LIST_FIRST_ROW = 7
TYPE_COL = "E"             # ist E leer, erscheint die Typauswahl
NEW_BED_VALUES = {"C": "Pflegebett", "H": "CE", "O": "Heim"}   # D und E aus der Auswahl

# Optionen auf "PRO Pflegebett": (Name, Zelle, voreingestellt, Wert wenn aus, rot wenn aus)
# Angehakt wird immer "Ja" eingetragen.
BED_OPTIONS = [
    ("Aufrichter", "C1", False, "-", False),
    ("Bettleuchte", "C2", False, "-", False),
    ("Seitengitter vorh.", "C3", True, "Nein", False),
    ("B-Maß", "C4", True, "Nein", True),
    ("Bedienteil", "K16", True, "Nein", True),
    ("Netzstecker", "K19", True, "Nein", True),
    ("Furnierschaden", "G46", True, "Nein", True),
]
OPTION_ON = "Ja"
# "Etage abschließen": alle Zeilen mit "x" in L ausfüllen
COVER_SHEET, COVER_DATE_CELL = "Deckblatt", "F23"   # Prüfdatum
FLOOR_TARGET_COL = "I"         # Etage  (aus C2)
AREA_TARGET_COL = "K"          # Wohnbereich (aus B2)
CHECK_DATE_COL = "L"           # letzte STK  = Prüfdatum
NEXT_DATE_COL = "N"            # nächste STK = Prüfdatum + 1 Jahr
# Fehlerhaftes Format aus V1.013.26 (wurde von deutschem Excel als „00/ yy“ gelesen)
BROKEN_DATE_FORMAT = "[$-407]mmm/ yy;@"
DATE_SERIAL_RANGE = (36526, 73051)  # 2000 … 2099: Zahlen in diesem Bereich sind Datumswerte
MONTHS = ["Jan.", "Feb.", "März", "Apr.", "Mai", "Juni", "Juli", "Aug.", "Sep.", "Okt.",
          "Nov.", "Dez."]
OPTIONS_TOP_ROW = 3        # obere Reihe: die ersten 3, darunter der Rest
# Stichworte in Q (Mängel) -> Haken wird beim Laden abgewählt (Groß-/Kleinschreibung egal)
OPTION_KEYWORDS = {
    "B-Maß": ["b-maß", "b-mass"],
    "Bedienteil": ["handschalter", "bedienteil"],
    "Netzstecker": ["netzstecker"],
    "Furnierschaden": ["furnier"],          # trifft auch "Furnierschaden"
}

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
# Kurzbezeichnungen (Auswahlschiene und Eingabefelder)
SHORT_TITLE = {"E": "Typ", "F": "Ser.Nr.", "Q": "Mängel"}
FIELD_TITLE = dict(SHORT_TITLE, R="MK")            # Überschrift im Eingabefeld
RAIL_TITLE = dict(SHORT_TITLE, R="Mängelklasse")   # Bezeichnung in der Auswahlschiene
CENTERED = {"G", "H", "J", "R", "S"}   # Eingabefelder mit zentriertem Text
OWNER_COL, REGNR_COL = "O", "P"
OWNER_WITH_REGNR = "RT Beuthel"     # leeres O -> "RT Beuthel": P (Reg.Nr.) einblenden
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
    text = str(value).strip().lower()
    if text.isdigit():                 # "001", "01" und "1" sind dieselbe Nummer
        text = str(int(text))
    return text


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
    """Ordnet Widgets nebeneinander an und bricht um, wenn die Breite nicht reicht.
    Liefert (belegte Breite, Anzahl Zeilen)."""
    for w in container.place_slaves():
        w.place_forget()
    x = y = line_h = used = 0
    lines = 1 if widgets else 0
    for w in widgets:
        ww, wh = w.winfo_reqwidth(), w.winfo_reqheight()
        if x > 0 and x + ww > width:
            x, y, line_h = 0, y + line_h + gap_y, 0
            lines += 1
        w.place(x=x, y=y)
        used = max(used, x + ww)
        x += ww + gap_x
        line_h = max(line_h, wh)
    container.configure(height=max(1, y + line_h))
    return used, lines


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


def excel_date(value):
    """Excel-Datumszahl (oder Text "17.11.2026") -> datetime.date, sonst None."""
    if is_number(value) and value > 0:
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(value))
    if isinstance(value, str):
        serial = parse_date(value)
        if serial is not None:
            return excel_date(serial)
    return None


def column_date_format(ws, col, last):
    """Häufigstes Zahlenformat der Datumszellen einer Spalte (so, wie Excel es liefert)."""
    if last < FIRST_ROW:
        return None
    values = ws.Range(f"{col}{FIRST_ROW}:{col}{last}").Value2
    if not isinstance(values, tuple):
        values = ((values,),)
    counts = {}
    for i, (v,) in enumerate(values):
        if is_number(v) and DATE_SERIAL_RANGE[0] <= v <= DATE_SERIAL_RANGE[1]:
            fmt = ws.Range(f"{col}{FIRST_ROW + i}").NumberFormat
            if fmt and fmt != BROKEN_DATE_FORMAT and fmt not in ("General", "Standard", "@"):
                counts[fmt] = counts.get(fmt, 0) + 1
            if sum(counts.values()) >= 25:      # Stichprobe genügt
                break
    return max(counts, key=counts.get) if counts else None


def has_year_token(fmt):
    """Enthält ein Zahlenformat ein Jahr (engl. "y" oder dt. "J")? Text in Anführungs-
    zeichen, mit \\ geschützte Zeichen und [...]-Angaben zählen nicht."""
    core = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", str(fmt or "")).lower()
    return "y" in core or "j" in core


def repair_date_formats(ws, last, date_fmt, undo):
    """Zellen mit dem fehlerhaften Format aus V1.013.26 wieder richtig formatieren."""
    repaired = 0
    for col, good in date_fmt.items():
        if not good or last < FIRST_ROW:
            continue
        for row in range(FIRST_ROW, last + 1):
            rng = ws.Range(f"{col}{row}")
            fmt = rng.NumberFormat
            value = rng.Value2
            is_date_value = is_number(value) and DATE_SERIAL_RANGE[0] <= value <= DATE_SERIAL_RANGE[1]
            if fmt == BROKEN_DATE_FORMAT or (is_date_value and fmt != good
                                             and not has_year_token(fmt)
                                             and fmt not in ("General", "Standard", "@")):
                undo["cells"].append((f"{col}{row}", snapshot(rng)))
                try:
                    rng.NumberFormat = good
                    repaired += 1
                except pywintypes.com_error:
                    pass
    return repaired


def add_year(day):
    try:
        return day.replace(year=day.year + 1)
    except ValueError:            # 29. Februar
        return day.replace(year=day.year + 1, day=28)


def to_serial(day):
    return float((day - datetime.date(1899, 12, 30)).days)


def month_text(day):
    return f"{MONTHS[day.month - 1]} {day.year % 100:02d}"


def pdf_name(id_norm):
    """Dateiname aus der Inventarnummer: Zahlen dreistellig (1 -> 001.pdf)."""
    text = (id_norm or "").strip()
    if text.isdigit():
        text = f"{int(text):03d}"
    text = re.sub(r'[\\/:*?"<>|]', "_", text) or "ohne_Nummer"
    return f"{text}.pdf"


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def put_value(rng, value, fmt):
    """Schreibt Zahl oder Text unverändert und stellt das Zellformat wieder her.

    Zahlen werden als echte Zahl übergeben – so bleiben sie auch in einer als
    „Text“ formatierten Zelle eine Zahl, ohne dass das Format angefasst wird.
    """
    if value is None or value == "":
        rng.ClearContents()
    elif is_number(value):
        rng.Value2 = float(value)
        # Falls Excel die Zahl in einer Text-Zelle doch als Text abgelegt hat:
        # Format kurz auf "Standard" setzen. Ist das nicht erlaubt (z.B. Blatt-
        # schutz), bleibt es ohne Fehlermeldung beim geschriebenen Wert.
        if fmt == "@" and isinstance(rng.Value2, str):
            try:
                rng.NumberFormat = "General"
                rng.Value2 = float(value)
            except pywintypes.com_error:
                pass
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

    def __init__(self, parent, var, width, on_return, on_escape, justify="left"):
        self.var = var
        self.justify = justify
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
        self.text.tag_configure("v", spacing1=pad, justify=self.justify)

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


class Tooltip:
    """Kleiner Hinweistext beim Darüberfahren mit der Maus."""

    def __init__(self, widget, text):
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _event=None):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + 10
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.overrideredirect(True)
        self.tip.attributes("-topmost", True)
        tk.Label(self.tip, text=self.text, bg="#ffffe0", relief="solid", bd=1,
                 padx=4, pady=1, font=(FONT, 9)).pack()
        self.tip.geometry(f"+{x}+{y}")

    def hide(self, _event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# --------------------------------------------------------------------------
# Auswahl des Bett-Typs
# --------------------------------------------------------------------------
class TypeDialog:
    """Pop-up mit der Typenliste aus "PRO Pflegebett" (mit Suchfeld)."""

    def __init__(self, root, types, id_label):
        self.types = types
        self.result = None
        top = self.top = tk.Toplevel(root)
        top.title(f"Neues Bett – Inv.-Nr. {id_label}")
        top.transient(root)
        top.attributes("-topmost", True)
        top.resizable(False, True)

        body = tk.Frame(top, padx=12, pady=10)
        body.pack(fill="both", expand=True)
        tk.Label(body, text="Bettentyp", font=(FONT, 14, "bold"), anchor="w").pack(fill="x")
        tk.Label(body, text=f"Für Inv.-Nr. {id_label} ist noch kein Typ (E) eingetragen.\n"
                            "Bitte den Bett-Typ auswählen – oder unten eine andere "
                            "Geräteart wählen:",
                 justify="left", anchor="w").pack(fill="x")
        search_line = tk.Frame(body)
        search_line.pack(fill="x", pady=(8, 4))
        tk.Label(search_line, text="Suchen:").pack(side="left")
        self.search = tk.StringVar()
        entry = tk.Entry(search_line, textvariable=self.search)
        entry.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.search.trace_add("write", lambda *_: self.refresh())

        table = tk.Frame(body)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, columns=("typ", "hersteller"), show="headings",
                                 height=16, selectmode="browse")
        self.tree.heading("typ", text="Typ")
        self.tree.heading("hersteller", text="Hersteller")
        self.tree.column("typ", width=200)
        self.tree.column("hersteller", width=180)
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self.ok())

        buttons = tk.Frame(body)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Abbrechen", command=self.cancel).pack(side="right")
        ttk.Button(buttons, text="OK", style="Accent.TButton",
                   command=self.ok).pack(side="right", padx=(0, 6))
        self.device_btn = ttk.Button(buttons, text="Geräteart ▾", command=self.show_devices)
        self.device_btn.pack(side="right", padx=(0, 6))
        self.device_menu = tk.Menu(top, tearoff=False, font=(FONT, 11))
        for name in DEVICE_TYPES:
            self.device_menu.add_command(label=name, command=lambda n=name: self.choose_device(n))

        for w in (top, entry, self.tree):
            w.bind("<Return>", lambda e: self.ok())
            w.bind("<KP_Enter>", lambda e: self.ok())
            w.bind("<Escape>", lambda e: self.cancel())
        entry.bind("<Down>", lambda e: self.tree.focus_set())
        top.protocol("WM_DELETE_WINDOW", self.cancel)

        self.refresh()
        self.place(root)
        top.grab_set()
        entry.focus_set()
        root.wait_window(top)

    def refresh(self):
        needle = self.search.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        for i, (typ, maker) in enumerate(self.types):
            if needle in typ.lower() or needle in maker.lower():
                self.tree.insert("", "end", iid=str(i), values=(typ, maker))
        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])

    def place(self, root):
        """Über dem Programmfenster, waagerecht mittig."""
        self.top.update_idletasks()
        w, h = self.top.winfo_reqwidth(), self.top.winfo_reqheight()
        x = root.winfo_rootx() + (root.winfo_width() - w) // 2
        y = max(0, root.winfo_rooty() - h - 40)
        self.top.geometry(f"+{max(0, x)}+{y}")

    def ok(self):
        sel = self.tree.selection()
        if not sel:
            return "break"
        self.result = ("type", self.types[int(sel[0])])
        self.top.destroy()
        return "break"

    def show_devices(self):
        b = self.device_btn
        self.device_menu.tk_popup(b.winfo_rootx(), b.winfo_rooty() + b.winfo_height())

    def choose_device(self, name):
        self.result = ("device", name)
        self.top.destroy()

    def cancel(self):
        self.result = None
        self.top.destroy()
        return "break"


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
        self.sheet_var = tk.StringVar(value="")   # aktives Protokoll-Blatt ("" = Liste)
        self._print_locked = False
        self.book_var = tk.StringVar()
        self.update_result = None
        self.update_info = None
        self._layout_job = None
        self._last_width = 0

        self.logo_full = tk.PhotoImage(data=LOGO_PNG_B64)
        root.iconphoto(True, self.logo_full.subsample(4))

        self.marker = RowMarker(root)
        self.temp_mark = None    # vorläufiges "x" in L: {book, row, id_norm, saved}
        self.pending = False     # Inventarnummer geladen, aber noch nicht übernommen
        self.repaired_books = set()
        self._ctx = None         # (app, wb, ws) der zuletzt benutzten Liste

        self.build_ui()
        for c in SELECTABLE:
            self.vars[c].trace_add("write", lambda *_a, c=c: self.on_value_change(c))

        root.title(f"Inventar-Editor {VERSION}")
        root.attributes("-topmost", self.topmost.get())
        self.place_initial()
        root.bind("<Configure>", self.on_configure)
        root.bind_all("<Control-s>", lambda e: self.save_book())
        root.protocol("WM_DELETE_WINDOW", self.on_close)

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

        # Raster: Spalte 0 = Liste / Inv.-Nr. | 1 = Bereich+Etage / Knopfgruppe
        #         2 = Optionen (zweizeilig) | 3 = Meldung | 4 = Update / Status (rechts)
        rows.columnconfigure(3, weight=1)

        self.book_combo = ttk.Combobox(rows, textvariable=self.book_var, state="readonly",
                                       width=26, font=(FONT, 11, "bold"))
        self.book_combo.grid(row=0, column=0, sticky="w")
        self.book_combo.bind("<<ComboboxSelected>>", self.on_book_selected)

        id_box = tk.Frame(rows)
        id_box.grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))
        tk.Label(id_box, text="Inventar-Nr.:", font=(FONT, 11, "bold")).pack(side="left")
        self.id_entry = tk.Entry(id_box, width=4, font=(FONT, 14, "bold"), justify="center")
        self.id_entry.pack(side="left", padx=(6, 0))
        self.id_entry.bind("<Return>", self.load_id)
        self.id_entry.bind("<KP_Enter>", self.load_id)

        # Bereich / Etage
        info = tk.Frame(rows)
        info.grid(row=0, column=1, sticky="w", padx=(14, 0))
        tk.Label(info, text="Bereich:").pack(side="left")
        self.area_entry = tk.Entry(info, font=(FONT, 12, "bold"), width=4, justify="center")
        self.area_entry.pack(side="left", padx=(3, 10))
        tk.Label(info, text="Etage:").pack(side="left")
        self.floor_entry = tk.Entry(info, font=(FONT, 12, "bold"), width=4, justify="center")
        self.floor_entry.pack(side="left", padx=(3, 0))
        self.header_cells = {AREA_CELL: self.area_entry, FLOOR_CELL: self.floor_entry}
        self.header_shown = {AREA_CELL: "", FLOOR_CELL: ""}
        for cell, entry in self.header_cells.items():
            for key in ("<Return>", "<KP_Enter>", "<FocusOut>"):
                entry.bind(key, lambda e, c=cell: self.write_header_cell(c))
        tk.Button(info, text="Etage abschließen", font=(FONT, 9, "bold"), relief="raised", bd=1,
                  padx=8, pady=0, cursor="hand2",
                  command=self.close_floor).pack(side="left", padx=(16, 0))

        # Knopfgruppe direkt rechts neben dem Nummernfeld (gleiche Zeile wie id_box)
        actions = tk.Frame(id_box)
        actions.pack(side="left", padx=(18, 0))
        group = tk.Frame(actions, highlightthickness=1, highlightbackground="#9a9a9a",
                         highlightcolor="#9a9a9a", padx=4, pady=3)
        group.pack(side="left")
        self.commit_btn = tk.Button(group, text="Übernehmen", font=(FONT, 10, "bold"),
                                    relief="raised", bd=1, padx=8, pady=1, cursor="hand2",
                                    command=self.commit)
        self.commit_btn.pack(side="left", padx=2)
        self.commit_default = (self.commit_btn.cget("bg"), self.commit_btn.cget("fg"),
                               self.commit_btn.cget("activebackground"))
        self.undo_btn = ttk.Button(group, text="Rückgängig", command=self.undo)
        self.undo_btn.pack(side="left", padx=2)
        self.save_icon = tk.PhotoImage(data=SAVE_PNG_B64)
        save_btn = ttk.Button(group, image=self.save_icon, command=self.save_book, width=3)
        save_btn.pack(side="left", padx=2)
        Tooltip(save_btn, "Liste speichern (Strg+S)")

        # Kacheln: [Drucker | Text] je Prüfprotokoll
        self.print_icon = tk.PhotoImage(data=PRINT_PNG_B64)
        self.tiles = {}
        for i, (label, sheet) in enumerate(SHEET_TILES):
            tile = tk.Frame(actions, relief="raised", bd=1)
            tile.pack(side="left", padx=(12 if i == 0 else 6, 0))
            icon = tk.Label(tile, image=self.print_icon, padx=6, pady=4, cursor="hand2")
            icon.pack(side="left")
            sep = tk.Frame(tile, width=1, bg="#a0a0a0")
            sep.pack(side="left", fill="y", pady=4)
            text = tk.Label(tile, text=label, font=(FONT, 10), padx=10, pady=4, cursor="hand2")
            text.pack(side="left")
            icon.bind("<Button-1>", lambda e, s=sheet: self.quick_print(s))
            text.bind("<Button-1>", lambda e, s=sheet: self.on_sheet_tile(s))
            Tooltip(icon, f"„{sheet}“ als PDF speichern")
            Tooltip(text, f"Zum Blatt „{sheet}“ wechseln")
            self.tiles[sheet] = (tile, icon, text)
        self.tile_default = (actions.cget("bg"), "black")
        self.sheet_var.trace_add("write", lambda *_: self.paint_tiles())

        # Optionen für "PRO Pflegebett" (zweizeilig)
        opts = tk.Frame(rows)
        opts.grid(row=0, column=2, rowspan=2, sticky="w", padx=(16, 0))
        self.option_vars, self.option_widgets = {}, {}
        for i, (name, cell, default, _off, _red) in enumerate(BED_OPTIONS):
            r, col = (0, i) if i < OPTIONS_TOP_ROW else (1, i - OPTIONS_TOP_ROW)
            var = tk.BooleanVar(value=default)
            cb = tk.Checkbutton(opts, text=name, variable=var, anchor="w", padx=2,
                                command=lambda n=name: self.on_option(n))
            cb.grid(row=r, column=col, sticky="w", padx=(0, 2), pady=1)
            self.option_vars[name], self.option_widgets[name] = var, cb
        self.option_default_look = (opts.cget("bg"), "black")
        for name, *_ in BED_OPTIONS:
            self.paint_option(name)


        # Meldungen – rechtsbündig unter "Nach Updates suchen"
        self.msg_label = tk.Label(rows, text="", anchor="e", justify="right", fg="#444", width=1)
        self.msg_label.grid(row=1, column=3, columnspan=2, sticky="we", padx=(14, 28), pady=(5, 0))

        right = tk.Frame(rows)
        right.grid(row=0, column=4, sticky="e", padx=(0, 28))
        ttk.Checkbutton(right, text="Immer im Vordergrund", variable=self.topmost,
                        command=self.toggle_topmost).pack(side="left", padx=(0, 12))
        tk.Label(right, text=VERSION, fg="#666").pack(side="left", padx=(0, 6))
        self.update_btn = ttk.Button(right, text="Nach Updates suchen",
                                     command=self.on_update_button)
        self.update_btn.pack(side="left")


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
        field_row = tk.Frame(outer)
        field_row.pack(fill="x")
        self.fields = tk.Frame(field_row)
        self.fields.pack(side="left", fill="x", expand=True)
        self.field_frames, self.field_labels, self.entries = {}, {}, {}
        for c in SELECTABLE:
            fr = tk.Frame(self.fields)
            title = tk.Frame(fr)
            title.pack(fill="x")
            tk.Label(title, text=c, fg="#333", font=(FONT, 9, "bold"), padx=0).pack(side="left")
            lb = tk.Label(title, anchor="w", fg="#333", font=(FONT, 9), padx=0)
            lb.pack(side="left")
            field = Field(fr, self.vars[c], FIELD_WIDTH.get(c, 12), self.commit,
                          lambda c=c: self.escape_field(c),
                          justify="center" if c in CENTERED else "left")
            field.text.pack(fill="x")
            self.field_frames[c], self.field_labels[c], self.entries[c] = fr, lb, field
        self.apply_headers()

    def apply_headers(self):
        for c in SELECTABLE:
            name = self.headers.get(c, c)
            self.check_names[c].configure(text=f" · {RAIL_TITLE.get(c, name)}")
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
        flow_place(self.fields, shown, width, gap_x=6)
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
        if status in ("yellow", "red"):
            added = False
            for c in AUTO_COLS:
                if not self.show_vars[c].get():
                    self.show_vars[c].set(True)
                    self.auto_shown.add(c)
                    added = True
            if added:
                self.schedule_layout()

    def paint_commit_button(self):
        """Orange ab dem Laden einer Inventarnummer bis zum Übernehmen
        (und immer, solange es nicht übernommene Änderungen gibt)."""
        if (getattr(self, "row", None) is not None and getattr(self, "pending", False)) \
                or self.changed_cols():
            self.commit_btn.configure(bg="#f28c18", fg="white", activebackground="#d97706",
                                      activeforeground="white")
        else:
            bg, fg, abg = self.commit_default
            self.commit_btn.configure(bg=bg, fg=fg, activebackground=abg, activeforeground=fg)

    def on_value_change(self, c):
        if self.loading:
            return
        self.paint_entry(c)
        self.paint_commit_button()
        if (c == OWNER_COL and self.row is not None and not self.orig[c].strip()
                and self.vars[c].get().strip().lower() == OWNER_WITH_REGNR.lower()
                and not self.show_vars[REGNR_COL].get()):
            self.show_vars[REGNR_COL].set(True)
            self.auto_shown.add(REGNR_COL)
            self.schedule_layout()
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
        self.paint_commit_button()
        self.schedule_layout()

    def changed_cols(self):
        if getattr(self, "row", None) is None:
            return []
        return [c for c in SELECTABLE
                if not self.formula[c] and self.vars[c].get() != self.orig[c]]

    def clear_loaded(self):
        self.pending = False
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
        if wb.Name not in self.repaired_books:
            self.repaired_books.add(wb.Name)
            self.repair_once(wb, ws)
        return app, wb, ws

    def repair_once(self, wb, ws):
        """Einmal pro Liste: fehlerhafte Datumsformate aus V1.013.26 reparieren."""
        try:
            last = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
            cover_fmt = wb.Worksheets(COVER_SHEET).Range(COVER_DATE_CELL).NumberFormat
            date_fmt = {col: column_date_format(ws, col, last) or cover_fmt
                        for col in (CHECK_DATE_COL, NEXT_DATE_COL)}
            undo = {"kind": "bulk", "label": "Reparatur der Datumsformate", "cells": []}
            repaired = repair_date_formats(ws, last, date_fmt, undo)
        except (pywintypes.com_error, ExcelError):
            return
        if repaired:
            self.undo_stack.append(undo)
            self.update_undo_button()
            self.message(f"{repaired} fehlerhafte Datumsformate in {CHECK_DATE_COL}/"
                         f"{NEXT_DATE_COL} repariert.")

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
        self.release_temp_mark()
        self.selected_book = name
        self.headers_loaded = False
        self.clear_loaded()
        self.undo_stack.clear()
        self.update_undo_button()
        app, wb, ws = self.get_book()
        wb.Activate()
        self.sheet_var.set("")
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
        self.release_temp_mark()
        self.reset_options(wb)
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
        self.pending = True                  # bis "Übernehmen": Knopf orange
        self.set_temp_mark(wb, ws, row)
        data = read_row(ws, row)
        self.fill(data)
        keyword_hits = self.options_from_defects(wb, data["Q"][0])
        if app.ActiveWorkbook is None or app.ActiveWorkbook.Name != wb.Name:
            wb.Activate()
        if wb.ActiveSheet.Name != ws.Name:       # z.B. von "PRO Pflegebett" zurück
            ws.Activate()
        self.sheet_var.set("")
        self.show_row(app, wb, ws, row)
        self.update_header_info(wb, ws)
        no_type = not data[TYPE_COL][0].strip()        # E (Typ/Modell) leer

        if len(rows) > 1:
            others = ", ".join(str(r) for r in rows)
            self.message(f"Inv.-Nr. {raw} kommt {len(rows)}× vor (Zeilen {others}) – "
                         f"bearbeitet: Zeile {row}.", error=True)
        else:
            hint = f" – abgewählt: {', '.join(keyword_hits)}" if keyword_hits else ""
            self.message(f"Inv.-Nr. {raw} geladen (Zeile {row}){hint}.")
        if no_type:
            self.root.after(50, lambda: self.offer_bed_types(wb))
        else:
            self.root.after(30, self.focus_first_field)
        return "break"

    # ------------------------------------------------------------ Neues Bett
    def read_bed_types(self, wb):
        """[(Typ, Hersteller)] aus "PRO Pflegebett" ab Zeile 7."""
        bed = wb.Worksheets(BED_SHEET)
        col = col_index(BED_TYPE_COL)
        last = bed.Cells(bed.Rows.Count, col).End(XL_UP).Row
        if last < BED_LIST_FIRST_ROW:
            return []
        first_col, last_col = sorted((BED_MAKER_COL, BED_TYPE_COL))
        values = bed.Range(f"{first_col}{BED_LIST_FIRST_ROW}:{last_col}{last}").Value2
        if not isinstance(values[0], tuple):
            values = (values,)
        i_type = col_index(BED_TYPE_COL) - col_index(first_col)
        i_maker = col_index(BED_MAKER_COL) - col_index(first_col)

        def text(v):
            if v is None:
                return ""
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            return str(v).strip()

        result = []
        for line in values:
            typ, maker = text(line[i_type]), text(line[i_maker])
            if typ:
                result.append((typ, maker))
        return result

    @excel_action
    def offer_bed_types(self, wb):
        types = self.read_bed_types(wb)
        if not types:
            self.message(f"Keine Typen in „{BED_SHEET}“ ab {BED_TYPE_COL}{BED_LIST_FIRST_ROW} "
                         "gefunden.", error=True)
            return self.focus_first_field()
        choice = TypeDialog(self.root, types, self.id_label).result
        if choice is None:
            self.message(f"Inv.-Nr. {self.id_label}: kein Typ ausgewählt.")
            return self.focus_first_field()
        kind, picked = choice
        if kind == "device":                       # andere Geräteart: nur C und H
            values = {"C": picked, "H": "CE"}
            info = f"Geräteart „{picked}“ eingetragen"
        else:
            typ, maker = picked
            # D und E kommen immer aus der Auswahl; C, H und O nur, wenn sie noch leer sind
            values = {c: v for c, v in NEW_BED_VALUES.items() if not self.vars[c].get().strip()}
            values.update(D=maker, E=typ)
            info = f"Neues Bett: {maker} {typ} eingetragen"
        for c in SELECTABLE:
            if c in values:
                self.vars[c].set(values[c])
                if not self.show_vars[c].get():
                    self.show_vars[c].set(True)
                    self.auto_shown.add(c)
        self.schedule_layout()
        self.message(f"{info} – mit „Übernehmen“ speichern.")
        self.root.after(30, self.focus_first_field)
        return "break"

    def escape_field(self, c):
        """Esc im Feld: Änderung verwerfen, wieder den Eintrag aus der Liste zeigen.
        Ist das Feld unverändert, springt Esc wie bisher ins Nummernfeld."""
        if self.row is None or self.formula[c]:
            return self.focus_id()
        current = self.orig[c]
        try:                                   # aktuellen Eintrag aus Excel holen
            app, wb, ws = self._ctx or self.get_book()
            text = ws.Range(f"{c}{self.row}").FormulaLocal
            while isinstance(text, tuple):     # manche Excel-Versionen liefern ((wert,),)
                text = text[0] if text else None
            text = "" if text is None else str(text)
            if not text.startswith("="):
                current = text
        except Exception:
            pass
        if self.vars[c].get() == current and self.orig[c] == current:
            return self.focus_id()
        self.orig[c] = current
        self.vars[c].set(current)
        self.paint_entry(c)
        self.entries[c].focus()
        name = FIELD_TITLE.get(c, self.headers.get(c, c))
        self.message(f"{c} · {name}: Änderung verworfen.")
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
        self.do_commit()
        return self.focus_id()

    def do_commit(self, reset_options=True):
        """Änderungen und "x" in L schreiben. Liefert True, wenn übernommen wurde."""
        if self.row is None:
            self.message("Bitte zuerst eine Inventarnummer eingeben.", error=True)
            return False
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
                return False

        undo = {"row": row, "id_norm": self.id_norm, "label": self.id_label, "cells": []}
        for c in changes + [DONE_COL]:
            rng = ws.Range(f"{c}{row}")
            saved = snapshot(rng)
            if c == DONE_COL and self.temp_mark and self.temp_mark["id_norm"] == self.id_norm:
                saved = self.temp_mark["saved"]      # Stand vor dem vorläufigen "x"
            undo["cells"].append((c, saved))
            write_cell(ws, c, row, DONE_MARK if c == DONE_COL else self.vars[c].get())
        self.temp_mark = None                        # "x" bleibt jetzt stehen
        self.undo_stack.append(undo)
        del self.undo_stack[:-50]
        self.update_undo_button()
        if reset_options:
            self.reset_options(wb)

        self.pending = False
        self.fill(read_row(ws, row))
        n = len(changes)
        self.message(f"Inv.-Nr. {self.id_label} übernommen: {n} Feld{'er' if n != 1 else ''} "
                     f"geändert, {DONE_COL} = „{DONE_MARK}“.")
        return True

    @excel_action
    def undo(self):
        if not self.undo_stack:
            self.message("Nichts zum Rückgängigmachen.")
            return
        u = self.undo_stack[-1]
        if u.get("kind") == "bulk":
            if not messagebox.askyesno(
                    "Rückgängig", f"„{u['label']}“ rückgängig machen?\n"
                    f"({len(u['cells'])} Zellen werden auf den alten Stand gesetzt.)",
                    parent=self.root):
                return
            app, wb, ws = self.get_book()
            for addr, saved in reversed(u["cells"]):
                restore_cell(ws.Range(addr), saved)
            self.undo_stack.pop()
            self.update_undo_button()
            self.clear_loaded()
            self.message(f"„{u['label']}“ rückgängig gemacht.")
            return self.focus_id()
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

    # ------------------------------------------------- Optionen PRO Pflegebett
    @staticmethod
    def option_spec(name):
        return next(o for o in BED_OPTIONS if o[0] == name)

    def paint_option(self, name):
        _n, _cell, _default, _off, red = self.option_spec(name)
        cb = self.option_widgets[name]
        bg, fg = self.option_default_look
        if red and not self.option_vars[name].get():
            cb.configure(bg="#f6c9c5", fg="#b3261e", activebackground="#f6c9c5",
                         activeforeground="#b3261e", selectcolor="white",
                         font=(FONT, 9, "bold"))
        else:
            cb.configure(bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
                         selectcolor="white", font=(FONT, 9))

    def write_option(self, wb, name, on):
        _n, cell, _default, off, _red = self.option_spec(name)
        rng = wb.Worksheets(BED_SHEET).Range(cell)
        value = OPTION_ON if on else off
        put_value(rng, value, rng.NumberFormat)
        return cell, value

    @excel_action
    def on_option(self, name):
        on = self.option_vars[name].get()
        try:
            app, wb, ws = self.get_book()
            cell, value = self.write_option(wb, name, on)
        except Exception:
            self.option_vars[name].set(not on)
            self.paint_option(name)
            raise
        self.paint_option(name)
        self.message(f"{BED_SHEET}: {name} → {cell} = „{value}“")

    # ------------------------------------------ vorläufiges "x" in Spalte L
    def set_temp_mark(self, wb, ws, row):
        """Beim Laden sofort "x" in L setzen; der alte Inhalt wird gemerkt."""
        rng = ws.Range(f"{DONE_COL}{row}")
        self.temp_mark = {"book": wb.Name, "row": row, "id_norm": self.id_norm,
                          "saved": snapshot(rng)}
        write_cell(ws, DONE_COL, row, DONE_MARK)

    def release_temp_mark(self):
        """Ohne "Übernehmen": den vorherigen Inhalt von L wiederherstellen."""
        mark, self.temp_mark = self.temp_mark, None
        if not mark:
            return
        try:
            app = excel_app()
            book = next((wb for name, wb, _ in list_books(app) if name == mark["book"]), None)
            if book is None:
                return
            ws = book.Worksheets(REF_SHEET)
            row = mark["row"]
            if norm_id(ws.Cells(row, 1).Value2) != mark["id_norm"]:
                rows = find_rows(ws, mark["id_norm"])
                if not rows:
                    return
                row = rows[0]
            restore_cell(ws.Range(f"{DONE_COL}{row}"), mark["saved"])
        except (ExcelError, pywintypes.com_error):
            pass

    def on_close(self):
        self.release_temp_mark()
        self.root.destroy()

    def options_from_defects(self, wb, defects):
        """Haken abwählen, deren Stichwort in Q (Mängel / Bemerkung) vorkommt."""
        text = (defects or "").lower()
        hits = []
        for name, words in OPTION_KEYWORDS.items():
            if any(w in text for w in words) and self.option_vars[name].get():
                self.write_option(wb, name, False)
                self.option_vars[name].set(False)
                self.paint_option(name)
                hits.append(name)
        return hits

    def quick_print(self, sheet):
        """Blatt als PDF speichern (Druckbereich des Blatts), Dateiname = Inv.-Nr. dreistellig."""
        if self._print_locked:                          # Doppelklick-Schutz
            return
        self._print_locked = True
        self.root.after(3000, lambda: setattr(self, "_print_locked", False))
        self._quick_print(sheet)

    @excel_action
    def _quick_print(self, sheet):
        """Erst übernehmen, dann drucken, danach die Haken zurücksetzen
        (sonst stünden auf dem Ausdruck schon die Voreinstellungen)."""
        app, wb, ws = self.get_book()
        if self.row is None:
            raise ExcelError("Bitte zuerst eine Inventarnummer eingeben.")
        # 1. Dateiname abfragen (vorausgefüllt), erst danach etwas verändern
        name = pdf_name(self.id_norm)
        folder = self.cfg.get("pdf_folder") or os.path.expanduser("~/Documents")
        if not os.path.isdir(folder):
            folder = os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            parent=self.root, title=f"„{sheet}“ als PDF speichern",
            initialdir=folder, initialfile=name, defaultextension=".pdf",
            filetypes=[("PDF-Datei", "*.pdf")])
        if not path:
            self.message("PDF nicht gespeichert (abgebrochen).")
            return self.focus_id()
        path = os.path.normpath(path)
        self.cfg["pdf_folder"] = os.path.dirname(path)
        save_config(self.cfg)
        # 2. übernehmen, 3. PDF erzeugen, 4. Haken zurücksetzen
        #    (sonst stünden im PDF schon die Voreinstellungen)
        committed = False
        if self.pending or self.changed_cols():
            if not self.do_commit(reset_options=False):
                return self.focus_id()              # z.B. Rückfrage abgelehnt
            committed = True
        wb.Worksheets(sheet).ExportAsFixedFormat(0, path)      # 0 = PDF
        self.reset_options(wb)
        done = "übernommen, " if committed else ""
        self.message(f"Inv.-Nr. {self.id_label} {done}gespeichert: {os.path.basename(path)}")
        return self.focus_id()

    @excel_action
    def close_floor(self):
        """Alle Zeilen mit "x" in L: Etage, Bereich, Prüfdatum und nächstes Datum eintragen."""
        if self.changed_cols() and not messagebox.askyesno(
                "Ungespeicherte Änderungen",
                f"Die Änderungen bei Inv.-Nr. {self.id_label} wurden noch nicht übernommen.\n\n"
                "Verwerfen und Etage abschließen?", parent=self.root):
            return
        app, wb, ws = self.get_book()
        self.release_temp_mark()      # vorläufiges "x" zählt nicht mit

        floor = ws.Range(FLOOR_CELL).Value2
        area = ws.Range(AREA_CELL).Value2
        if floor in (None, "") or area in (None, ""):
            raise ExcelError("Bitte zuerst Bereich und Etage eintragen.")
        check = excel_date(wb.Worksheets(COVER_SHEET).Range(COVER_DATE_CELL).Value2)
        if check is None:
            raise ExcelError(f"Kein gültiges Datum in „{COVER_SHEET}“ {COVER_DATE_CELL}.")
        next_check = add_year(check)

        last = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        marks = ws.Range(f"{DONE_COL}{FIRST_ROW}:{DONE_COL}{max(last, FIRST_ROW)}").Value2
        if not isinstance(marks, tuple):
            marks = ((marks,),)
        rows = [FIRST_ROW + i for i, (v,) in enumerate(marks)
                if isinstance(v, str) and v.strip().lower() == DONE_MARK]
        if not rows:
            self.clear_loaded()
            self.message(f"Keine Zeile mit „{DONE_MARK}“ in Spalte {DONE_COL} gefunden.", error=True)
            return self.focus_id()

        show = lambda v: norm_id(v) or "–"
        if not messagebox.askyesno(
                "Etage abschließen",
                f"{len(rows)} Zeile(n) mit „{DONE_MARK}“ in Spalte {DONE_COL}.\n\n"
                f"{FLOOR_TARGET_COL} (Etage) = {show(floor)}\n"
                f"{AREA_TARGET_COL} (Wohnbereich) = {show(area)}\n"
                f"{CHECK_DATE_COL} (letzte STK) = {month_text(check)}\n"
                f"{NEXT_DATE_COL} (nächste STK) = {month_text(next_check)}\n\n"
                "Jetzt eintragen?", parent=self.root):
            return

        # Datumsformat nicht selbst vorgeben (Excel deutet Formatcodes je nach Sprache
        # unterschiedlich), sondern von vorhandenen Datumszellen der Spalte übernehmen.
        cover_fmt = wb.Worksheets(COVER_SHEET).Range(COVER_DATE_CELL).NumberFormat
        date_fmt = {col: column_date_format(ws, col, last) or cover_fmt
                    for col in (CHECK_DATE_COL, NEXT_DATE_COL)}
        undo = {"kind": "bulk", "label": f"Etage abschließen ({len(rows)} Zeilen)", "cells": []}
        values = [(FLOOR_TARGET_COL, floor, False), (AREA_TARGET_COL, area, False),
                  (CHECK_DATE_COL, to_serial(check), True),
                  (NEXT_DATE_COL, to_serial(next_check), True)]
        try:
            app.ScreenUpdating = False
        except pywintypes.com_error:
            pass
        try:
            for row in rows:
                for col, value, is_date in values:
                    rng = ws.Range(f"{col}{row}")
                    saved = snapshot(rng)
                    undo["cells"].append((f"{col}{row}", saved))
                    put_value(rng, value, saved[1])
                    if is_date and date_fmt[col] and rng.NumberFormat != date_fmt[col]:
                        try:
                            rng.NumberFormat = date_fmt[col]
                        except pywintypes.com_error:
                            pass
            repaired = repair_date_formats(ws, last, date_fmt, undo)
        finally:
            try:
                app.ScreenUpdating = True
            except pywintypes.com_error:
                pass
        self.undo_stack.append(undo)
        del self.undo_stack[:-50]
        self.update_undo_button()
        self.clear_loaded()
        fixed = f" ({repaired} Datumsformat(e) repariert)" if repaired else ""
        self.message(f"Etage abgeschlossen: {len(rows)} Zeile(n) – Etage {show(floor)}, "
                     f"Bereich {show(area)}, {month_text(check)} / {month_text(next_check)}{fixed}.")
        return self.focus_id()

    def reset_options(self, wb):
        """Geänderte Optionen auf die Voreinstellung zurücksetzen (auch in Excel)."""
        count = 0
        for name, _cell, default, _off, _red in BED_OPTIONS:
            if self.option_vars[name].get() != default:
                self.write_option(wb, name, default)
                self.option_vars[name].set(default)
                self.paint_option(name)
                count += 1
        return count

    def sync_options(self, wb):
        """Haken so setzen, wie es gerade in "PRO Pflegebett" steht."""
        bed = wb.Worksheets(BED_SHEET)
        for name, cell, _default, _off, _red in BED_OPTIONS:
            value = bed.Range(cell).Value2
            on = str(value or "").strip().lower() == OPTION_ON.lower()
            if self.option_vars[name].get() != on:
                self.option_vars[name].set(on)
                self.paint_option(name)

    def paint_tiles(self):
        active = self.sheet_var.get()
        bg, fg = self.tile_default
        for sheet, (tile, icon, text) in self.tiles.items():
            on = sheet == active
            color = "#2e9e44" if on else bg
            tile.configure(bg=color, relief="sunken" if on else "raised")
            icon.configure(bg=color)
            text.configure(bg=color, fg="white" if on else fg,
                           font=(FONT, 10, "bold") if on else (FONT, 10))

    def on_sheet_tile(self, sheet):
        """Text der Kachel: zum Blatt wechseln – ist es schon aktiv (grün), zurück zur Liste."""
        back = self.sheet_var.get() == sheet
        try:
            app, wb, ws = self.get_book()
            if app.ActiveWorkbook is None or app.ActiveWorkbook.Name != wb.Name:
                wb.Activate()
            if back:
                ws.Activate()
                self.sheet_var.set("")
                if self.row is not None:
                    self.show_row(app, wb, ws, self.locate(ws))
                self.message(f"Zurück zu „{REF_SHEET}“.")
            else:
                wb.Worksheets(sheet).Activate()
                self.sheet_var.set(sheet)
                self.message(f"Blatt „{sheet}“ angezeigt.")
        except ExcelError as e:
            self.message(str(e), error=True)
        except pywintypes.com_error as e:
            self.message(com_message(e), error=True)
        self.focus_id()

    def update_header_info(self, wb, ws):
        """Bereich (B2) und Etage (C2) anzeigen – nicht, während dort getippt wird."""
        focus = self.root.focus_get()
        for cell, entry in self.header_cells.items():
            if focus is entry:
                continue
            text = str(ws.Range(cell).Text or "")
            if entry.get() != text:
                entry.delete(0, "end")
                entry.insert(0, text)
            self.header_shown[cell] = text

    @excel_action
    def write_header_cell(self, cell):
        entry = self.header_cells[cell]
        text = entry.get().strip()
        if text == self.header_shown[cell]:
            return
        app, wb, ws = self.get_book()
        col, row = cell[0], int(cell[1:])
        write_cell(ws, col, row, text)
        self.header_shown[cell] = text
        name = "Bereich" if cell == AREA_CELL else "Etage"
        self.message(f"{name} → {REF_SHEET}!{cell} = „{text}“")

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
                self.sync_options(wb)
            except pywintypes.com_error:
                pass
            try:
                active = wb.ActiveSheet.Name
                wanted = active if active in self.tiles else ""
                if self.sheet_var.get() != wanted:
                    self.sheet_var.set(wanted)
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
        self.release_temp_mark()
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
