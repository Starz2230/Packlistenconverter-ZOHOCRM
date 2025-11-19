#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Packlisten-Converter Kernmodul.

- Kann als Desktop-GUI genutzt werden (tkinter),
- oder headless als Library (z.B. über Flask / Render).

HEADLESS wird über die Umgebungsvariable HEADLESS gesteuert:
    HEADLESS=1 -> kein GUI / kein tkinter, nur Funktionen
    HEADLESS=0 -> GUI erlaubt (falls tkinter vorhanden)
"""

import sys
import os
import re
import json
import shutil
import tempfile
import math
import datetime

import pandas as pd
import openpyxl
from copy import copy
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

import requests
from packaging import version

# ----------------------------------------------------------------
# HEADLESS / GUI-Schalter
# ----------------------------------------------------------------

import os as _os

# Standard: HEADLESS = 1 (Web/Server). Für Desktop-GUI HEADLESS=0 setzen.
HEADLESS = _os.environ.get("HEADLESS", "1") == "1"

if not HEADLESS:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        from tkinter import ttk
        _TK_AVAILABLE = True
    except Exception:
        _TK_AVAILABLE = False
        HEADLESS = True
else:
    _TK_AVAILABLE = False

if HEADLESS:
    # Fallback-Stubs, damit messagebox & filedialog aufrufbar bleiben
    class _DummyMessageBox:
        @staticmethod
        def showinfo(title, message): print(f"[INFO] {title}: {message}")
        @staticmethod
        def showwarning(title, message): print(f"[WARN] {title}: {message}")
        @staticmethod
        def showerror(title, message): print(f"[ERROR] {title}: {message}")
        @staticmethod
        def askyesno(title, message):
            print(f"[ASKYESNO - default NO] {title}: {message}")
            return False

    class _DummyFileDialog:
        @staticmethod
        def askopenfilename(*args, **kwargs):
            raise RuntimeError("Dateiauswahl-Dialog im Headless-Modus nicht verfügbar.")
        @staticmethod
        def asksaveasfilename(*args, **kwargs):
            raise RuntimeError("Speichern-Dialog im Headless-Modus nicht verfügbar.")

    messagebox = _DummyMessageBox()
    filedialog = _DummyFileDialog()

    class _DummyTtk:
        pass

    ttk = _DummyTtk()

# ----------------------------------------------------------------
# GitHub-Update-Check (für Desktop-EXE; im Headless-Modus faktisch inaktiv)
# ----------------------------------------------------------------

GITHUB_API = "https://api.github.com/repos/Starz2230/MeinePacklistenApp/releases/latest"
CURRENT_VERSION = "1.1.9"   # anpassen, wenn du neu releast


def resource_path(rel):
    """Pfad für PyInstaller/EXE oder normales Skript korrekt auflösen."""
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.abspath(".")
    return os.path.join(base, rel)


def check_for_updates():
    """Prüft über GitHub, ob eine neue Version verfügbar ist (nur Desktop)."""
    if HEADLESS or not _TK_AVAILABLE:
        return  # im Server-/Headlessmodus nichts tun
    try:
        r = requests.get(GITHUB_API, timeout=5)
        data = r.json()
        latest = data["tag_name"].lstrip("v")
        if version.parse(latest) > version.parse(CURRENT_VERSION):
            url = data["assets"][0]["browser_download_url"]
            if messagebox.askyesno(
                "Update verfügbar",
                f"Version {latest} ist verfügbar.\nJetzt aktualisieren?"
            ):
                updater = resource_path("updater.exe")
                subprocess.Popen([updater, sys.executable, url])
                sys.exit(0)
    except Exception as e:
        print("Update-Check fehlgeschlagen:", e)


# ----------------------------------------------------------------
# Wochentage / Datums-Helfer
# ----------------------------------------------------------------

weekday_map = {
    0: "MO",
    1: "DI",
    2: "MI",
    3: "DO",
    4: "FR",
    5: "SA",
    6: "SO",
}


def transform_zeitraum(val):
    """
    Parst einen Datumsstring wie "21.03.2025 08:00 - 09:00"
    und wandelt ihn um zu "FR 21.03.25 08:00 - 09:00".
    """
    if not val or not isinstance(val, str):
        return val

    m = re.match(r'^(\d{1,2}\.\d{1,2}\.\d{4})(.*)$', val.strip())
    if not m:
        return val

    date_str = m.group(1).strip()
    rest = m.group(2)

    try:
        dt = datetime.datetime.strptime(date_str, "%d.%m.%Y")
        wday = dt.weekday()
        wday_abbr = weekday_map.get(wday, "")
        date_formatted = dt.strftime("%d.%m.%y")
        return f"{wday_abbr} {date_formatted}{rest}"
    except Exception:
        return val


try:
    import win32com.client
except ImportError:
    win32com = None

if os.name == "nt":
    import winsound


def resource_path(relative_path):
    """Pfad (nochmal) – wird im restlichen Code verwendet."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_app_directory():
    """Verzeichnis der aktuell laufenden Applikation."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


TEMPLATE_FILE = "Packliste_Template.xlsx"
ICON_FILE = "convert.ico"
DICHTUNGEN_CONFIG = "dichtungen.json"

# ----------------------------------------------------------------
# Zeilen-/Spalten-Definitionen
# ----------------------------------------------------------------

SERVICE_TECHNIKER_ROW = 1
DATE_ROW = 2

TEMPLATE_SUM_ROW = 1
TEMPLATE_DICHTUNG_NAME_ROW = 2
TEMPLATE_DATA_START_ROW = 3

DF_DATA_START_ROW = 1
DF_SUM_ROW = 0

PLATZHALTER_COL_INDEX = 5  # E
NUMBERING_COL = 1

current_theme = "dark"
SOUND_ENABLED = False
last_output_file = None

PRINT_SETTINGS = {
    "margin_top": "10",
    "margin_bottom": "10",
    "margin_left": "10",
    "margin_right": "10",
    "scaling": "100%",
    "duplex": False,
    "copies": "1",
    "paper_size": "A4",
    "fit_to_page": False
}
PRINTER_COMMAND = "print"

AUTO_FILENAME = False
SAVE_FOLDER = ""
GROUP_BY_SUFFIX = False
AUTO_FIT_COLUMNS = True  # AutoFit standardmäßig an

SETTINGS_CONFIG = "settings.json"

# ----------------------------------------------------------------
# Settings laden / speichern
# ----------------------------------------------------------------


def load_settings():
    global current_theme, SOUND_ENABLED, PRINT_SETTINGS, PRINTER_COMMAND
    global AUTO_FILENAME, SAVE_FOLDER, GROUP_BY_SUFFIX, AUTO_FIT_COLUMNS

    app_dir = get_app_directory()
    path = os.path.join(app_dir, SETTINGS_CONFIG)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                settings = json.load(f)
            current_theme = settings.get("current_theme", "dark")
            SOUND_ENABLED = settings.get("SOUND_ENABLED", False)
            PRINT_SETTINGS = settings.get("PRINT_SETTINGS", PRINT_SETTINGS)
            PRINTER_COMMAND = settings.get("PRINTER_COMMAND", "print")
            AUTO_FILENAME = settings.get("AUTO_FILENAME", False)
            SAVE_FOLDER = settings.get("SAVE_FOLDER", "")
            GROUP_BY_SUFFIX = settings.get("GROUP_BY_SUFFIX", False)
            AUTO_FIT_COLUMNS = settings.get("AUTO_FIT_COLUMNS", True)
        except Exception as e:
            print("Fehler beim Laden der Einstellungen:", e)
    else:
        save_settings()


def save_settings():
    global current_theme, SOUND_ENABLED, PRINT_SETTINGS, PRINTER_COMMAND
    global AUTO_FILENAME, SAVE_FOLDER, GROUP_BY_SUFFIX, AUTO_FIT_COLUMNS

    app_dir = get_app_directory()
    path = os.path.join(app_dir, SETTINGS_CONFIG)
    settings = {
        "current_theme": current_theme,
        "SOUND_ENABLED": SOUND_ENABLED,
        "PRINT_SETTINGS": PRINT_SETTINGS,
        "PRINTER_COMMAND": PRINTER_COMMAND,
        "AUTO_FILENAME": AUTO_FILENAME,
        "SAVE_FOLDER": SAVE_FOLDER,
        "GROUP_BY_SUFFIX": GROUP_BY_SUFFIX,
        "AUTO_FIT_COLUMNS": AUTO_FIT_COLUMNS
    }
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(e)


# ----------------------------------------------------------------
# Themes
# ----------------------------------------------------------------

dark_theme = {
    "bg": "#1E1E1E",
    "fg": "#FFFFFF",
    "button_bg": "#4CAF50",
    "button_fg": "#FFFFFF",
    "entry_bg": "#2B2B2B",
    "entry_fg": "#FFFFFF",
}

light_theme = {
    "bg": "#FFFFFF",
    "fg": "#000000",
    "button_bg": "#4CAF50",
    "button_fg": "#000000",
    "entry_bg": "#FFFFFF",
    "entry_fg": "#000000",
}


def apply_ttk_style(theme_name: str):
    if HEADLESS:
        return
    style = ttk.Style()
    if theme_name == "dark":
        style.theme_use("clam")
        style.configure(".", background=dark_theme["bg"], foreground=dark_theme["fg"])
        style.configure("TFrame", background=dark_theme["bg"])
        style.configure("TLabel", background=dark_theme["bg"], foreground=dark_theme["fg"])
        style.configure("TButton", background=dark_theme["button_bg"], foreground=dark_theme["button_fg"])
        style.configure("TCheckbutton", background=dark_theme["bg"], foreground=dark_theme["fg"])
        style.configure("TRadiobutton", background=dark_theme["bg"], foreground=dark_theme["fg"])
        style.configure("TEntry", fieldbackground=dark_theme["entry_bg"], foreground=dark_theme["entry_fg"])
        style.configure("TNotebook", background=dark_theme["bg"])
        style.configure("TNotebook.Tab", background=dark_theme["bg"], foreground=dark_theme["fg"])
    else:
        style.theme_use("default")
        style.configure(".", background=light_theme["bg"], foreground=light_theme["fg"])
        style.configure("TFrame", background=light_theme["bg"])
        style.configure("TLabel", background=light_theme["bg"], foreground=light_theme["fg"])
        style.configure("TButton", background=light_theme["button_bg"], foreground=light_theme["button_fg"])
        style.configure("TCheckbutton", background=light_theme["bg"], foreground=light_theme["fg"])
        style.configure("TRadiobutton", background=light_theme["bg"], foreground=light_theme["fg"])
        style.configure("TEntry", fieldbackground=light_theme["entry_bg"], foreground=light_theme["entry_fg"])
        style.configure("TNotebook", background=light_theme["bg"])
        style.configure("TNotebook.Tab", background=light_theme["bg"], foreground=light_theme["fg"])


# ----------------------------------------------------------------
# Dichtungen laden / speichern
# ----------------------------------------------------------------

DEFAULT_DICHTUNGEN = []  # falls du Standardwerte hast, hier eintragen


def load_dichtungen():
    app_dir = get_app_directory()
    path = os.path.join(app_dir, DICHTUNGEN_CONFIG)
    if not os.path.exists(path):
        save_dichtungen(DEFAULT_DICHTUNGEN)
        return DEFAULT_DICHTUNGEN.copy()
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(e)
        return []
    normalized = []
    for item in data:
        if isinstance(item, dict):
            normalized.append(item)
        else:
            normalized.append({"name": item, "always_show": False, "default_value": 0})
    return normalized


def save_dichtungen(dichtungen_list):
    app_dir = get_app_directory()
    path = os.path.join(app_dir, DICHTUNGEN_CONFIG)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dichtungen_list, f, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------
# Excel-Helfer
# ----------------------------------------------------------------

def copy_cell_style(src_cell, dst_cell):
    if src_cell.font:
        dst_cell.font = copy(src_cell.font)
    if src_cell.border:
        dst_cell.border = copy(src_cell.border)
    if src_cell.fill:
        dst_cell.fill = copy(src_cell.fill)
    if src_cell.number_format:
        dst_cell.number_format = src_cell.number_format
    if src_cell.protection:
        dst_cell.protection = copy(src_cell.protection)
    if src_cell.alignment:
        align = copy(src_cell.alignment)
        align.wrap_text = True
        dst_cell.alignment = align
    else:
        dst_cell.alignment = Alignment(wrap_text=True)


def copy_entire_row_format(ws, src_row_idx, dst_row_idx):
    ws.row_dimensions[dst_row_idx].height = None
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        sc = ws.cell(row=src_row_idx, column=col)
        dc = ws.cell(row=dst_row_idx, column=col)
        copy_cell_style(sc, dc)
        dc.value = None


def copy_column_with_style(ws, src_col_idx, dst_col_idx):
    max_row = ws.max_row
    for row in range(1, max_row + 1):
        sc = ws.cell(row=row, column=src_col_idx)
        dc = ws.cell(row=row, column=dst_col_idx)
        copy_cell_style(sc, dc)
        dc.value = sc.value
    src_letter = get_column_letter(src_col_idx)
    dst_letter = get_column_letter(dst_col_idx)
    if ws.column_dimensions[src_letter].width:
        ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width


def safe_val(df, col, index):
    if col not in df.columns:
        return ""
    if index < 0 or index >= len(df):
        return ""
    val = df[col].iloc[index]
    if pd.isna(val):
        return ""
    return str(val)


def parse_number(s):
    try:
        return float(str(s).replace(",", "."))
    except Exception:
        return 0.0


def parse_date_part(value):
    if not value or not isinstance(value, str):
        return None
    match = re.match(r"^(\d{1,2}\.\d{1,2}\.\d{4})", value.strip())
    if not match:
        return None
    date_str = match.group(1)
    try:
        dt = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
        return dt
    except Exception:
        return None


def get_zeitraum_von_bis(df, col="Zeitraum"):
    if col not in df.columns:
        return ""
    dtlist = []
    for val in df[col].dropna():
        dt = parse_date_part(str(val))
        if dt is not None:
            dtlist.append(dt)
    if not dtlist:
        return ""
    von_dt = min(dtlist)
    bis_dt = max(dtlist)
    return f"{von_dt.strftime('%d.%m.%Y')} - {bis_dt.strftime('%d.%m.%Y')}"


def spalte_leer(dataf, colname):
    if colname not in dataf.columns:
        return True
    col_series = dataf[colname].dropna().astype(str).str.strip()
    if len(col_series) == 0:
        return True
    return col_series.eq("").all()


def generate_auto_filename(df):
    serv = safe_val(df, "Service Techniker", 3)
    date_range = get_zeitraum_von_bis(df, "Zeitraum")

    def sanitize(text):
        return "".join(c for c in text if c.isalnum() or c in ("-", "_"))

    serv_sanitized = sanitize(serv)
    date_sanitized = date_range.replace(".", "-").replace(" ", "")
    base_filename = f"{serv_sanitized}_{date_sanitized}.xlsx"
    folder = SAVE_FOLDER if SAVE_FOLDER else os.path.abspath(".")
    filename = os.path.join(folder, base_filename)
    counter = 1
    while os.path.exists(filename):
        filename = os.path.join(folder, f"{serv_sanitized}_{date_sanitized}_{counter}.xlsx")
        counter += 1
    return filename


def get_suffix(name: str):
    parts = name.rsplit("_", 1)
    if len(parts) == 2 and parts[1].strip():
        return parts[1].strip()
    return ""


def dichtung_sort_key(d, df):
    col_order = {col: i for i, col in enumerate(df.columns)}
    df_pos = col_order.get(d["name"], 999999)
    m_color = re.search(r"_(\w+)$", d["name"])
    if m_color:
        color = m_color.group(1).upper()
    else:
        color = ""
    manual = d.get("order", None)
    if manual is not None and str(manual).strip() != "":
        try:
            manual_int = int(manual)
        except Exception:
            manual_int = 999999
        return (0, color, manual_int, df_pos)
    else:
        m = re.search(r"(\d+)(?:/(\d+))?", d["name"])
        if m:
            first = int(m.group(1))
            second = int(m.group(2)) if m.group(2) else 0
            extracted = first + second / 1000.0
        else:
            extracted = 999999
        return (1, color, extracted, df_pos)


def group_by_suffix_order(dichtungen, df):
    col_order = {col: i for i, col in enumerate(df.columns)}
    group_s = []
    group_w = []
    group_rest = []
    for d in dichtungen:
        sfx = get_suffix(d["name"]).upper()
        if sfx == "S":
            group_s.append(d)
        elif sfx == "W":
            group_w.append(d)
        else:
            group_rest.append(d)
    group_s_sorted = sorted(group_s, key=lambda d: dichtung_sort_key(d, df))
    group_w_sorted = sorted(group_w, key=lambda d: dichtung_sort_key(d, df))
    group_r_sorted = sorted(group_rest, key=lambda d: dichtung_sort_key(d, df))
    return group_s_sorted + group_w_sorted + group_r_sorted


def standard_first_then_df(dichtungen, df):
    standard = [d for d in dichtungen if d.get("always_show", False)]
    nonstandard = [d for d in dichtungen if not d.get("always_show", False)]
    standard_sorted = sorted(standard, key=lambda d: dichtung_sort_key(d, df))
    nonstandard_sorted = sorted(nonstandard, key=lambda d: dichtung_sort_key(d, df))
    return standard_sorted + nonstandard_sorted


def order_dichtungen_by_manual(dichtungen):
    manual_dict = {}
    non_manual = []

    for d in dichtungen:
        order_str = str(d.get("order", "")).strip()
        if order_str:
            try:
                pos = int(order_str)
                manual_dict.setdefault(pos, []).append(d)
            except ValueError:
                non_manual.append(d)
        else:
            non_manual.append(d)

    n = len(dichtungen)
    if manual_dict:
        max_order = max(manual_dict.keys())
        if max_order > n:
            n = max_order

    result = []
    non_manual_index = 0
    for slot in range(1, n + 1):
        if slot in manual_dict and manual_dict[slot]:
            result.extend(manual_dict[slot])
        else:
            if non_manual_index < len(non_manual):
                result.append(non_manual[non_manual_index])
                non_manual_index += 1

    while non_manual_index < len(non_manual):
        result.append(non_manual[non_manual_index])
        non_manual_index += 1

    return result


def set_horizontal_dotted(ws, row_idx):
    dotted_side = Side(style='dotted', color="999999")
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        b = copy(c.border) or Border()
        b.top = dotted_side
        b.bottom = dotted_side
        c.border = b


def set_bottom_thick(ws, row_idx):
    thick_side = Side(style='medium')
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        b = copy(c.border) or Border()
        b.bottom = thick_side
        c.border = b


def set_top_border_solid(ws, row_idx):
    thin_side = Side(style='thin')
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        b = copy(c.border) or Border()
        b.top = thin_side
        c.border = b


def set_column_left_border(ws, col_idx, start_row=3, border_style='thin'):
    side = Side(style=border_style, color="000000")
    for row in range(start_row, ws.max_row + 1):
        cell = ws.cell(row=row, column=col_idx)
        b = copy(cell.border) or Border()
        b.left = side
        cell.border = b


def set_bottom_solid(ws, row_idx):
    side = Side(style='thin', color="000000")
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        c = ws.cell(row=row_idx, column=col_idx)
        b = copy(c.border) or Border()
        b.bottom = side
        c.border = b


def remove_trailing_blank_rows(ws, start_row):
    def row_is_blank(r):
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=col_idx).value
            if val not in (None, ""):
                return False
        return True

    last_row = ws.max_row
    while last_row > start_row:
        if row_is_blank(last_row):
            ws.delete_rows(last_row, 1)
        else:
            break
        last_row = ws.max_row


def auto_fit_dichtungen_in_excel(xlsx_path, col_indices):
    if not win32com:
        print("win32com nicht verfügbar, AutoFit wird übersprungen.")
        return

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(xlsx_path)
    ws = wb.ActiveSheet

    for col_idx in col_indices:
        col_letter = get_column_letter(col_idx)
        rng_str = f"{col_letter}:{col_letter}"
        ws.Range(rng_str).Columns.AutoFit()

    wb.Save()
    wb.Close()
    excel.Quit()


# ----------------------------------------------------------------
# Dichtungsnamen umbrechen & sortieren
# ----------------------------------------------------------------

def apply_dicht_name_break(dicht_name: str) -> str:
    if "_" in dicht_name:
        idx = dicht_name.index("_")
        return dicht_name[:idx] + "\n" + dicht_name[idx + 1:]
    else:
        name = dicht_name.strip()
        one_line_width = (len(name) + 1) * 1.1
        two_line_width = (math.ceil(len(name) / 2) + 1) * 1.1
        if two_line_width < one_line_width:
            half = math.ceil(len(name) / 2)
            return name[:half] + "\n" + name[half:]
        else:
            return name


def parse_numeric_part(name: str) -> float:
    m = re.search(r'(\d+)(?:/(\d+))?', name)
    if m:
        first = int(m.group(1))
        second = int(m.group(2)) if m.group(2) else 0
        return first + second / 1000.0
    return 999999


def parse_suffix_priority(name: str) -> int:
    sfx = ""
    parts = name.rsplit("_", 1)
    if len(parts) == 2 and parts[1].strip():
        sfx = parts[1].strip().upper()
    if sfx == "S":
        return 0
    elif sfx == "W":
        return 1
    elif sfx == "G":
        return 2
    else:
        return 99


def final_sort_dichtungen(dichtungen, df):
    def sort_key(d):
        is_std = d.get("always_show", False)
        order_str = str(d.get("order", "")).strip()
        try:
            order_val = int(order_str)
        except ValueError:
            order_val = None

        name = d["name"]
        suffix_prio = parse_suffix_priority(name)
        numeric_val = parse_numeric_part(name)
        alpha_name = name.lower()

        if is_std:
            group = 0
            if order_val is not None:
                return (group, 0, order_val)
            else:
                return (group, 1, numeric_val, alpha_name)
        else:
            group = 1
            return (group, suffix_prio, numeric_val, alpha_name)

    return sorted(dichtungen, key=sort_key)


# ----------------------------------------------------------------
# WICHTIG: Hauptfunktion für Flask / Web
# ----------------------------------------------------------------

def convert_file(input_path, output_path, user_dichtungen, show_message=True):
    template_path = resource_path(TEMPLATE_FILE)
    if not os.path.isfile(template_path):
        messagebox.showerror("Fehler", f"'{TEMPLATE_FILE}' wurde nicht gefunden!")
        return

    # 1) Original-Template öffnen und Spaltenbreiten sichern
    template_orig_wb = openpyxl.load_workbook(template_path)
    template_orig_ws = template_orig_wb.active
    original_width_info = template_orig_ws.column_dimensions["F"].width
    original_width_ersatz = template_orig_ws.column_dimensions["G"].width

    file_ext = os.path.splitext(input_path)[1].lower()
    if file_ext == ".csv":
        df = pd.read_csv(input_path, sep=";", engine="python", header=0)
    else:
        df = pd.read_excel(input_path, header=0)

    # -------------------------------------------------------------
    # Summenzeile abtrennen, darunter nach Datum/Uhrzeit sortieren
    # -------------------------------------------------------------
    try:
        sum_row = df.iloc[[0]].copy()
        data_rows = df.iloc[1:].copy()

        def parse_datetime(x):
            pattern = r'^(\d{1,2}\.\d{1,2}\.\d{4})\s+(\d{1,2}:\d{1,2})'
            m = re.match(pattern, str(x))
            if m:
                dt_str = f"{m.group(1)} {m.group(2)}"
                return pd.to_datetime(dt_str, format='%d.%m.%Y %H:%M', errors='coerce')
            else:
                m2 = re.match(r'^(\d{1,2}\.\d{1,2}\.\d{4})', str(x))
                if m2:
                    return pd.to_datetime(m2.group(1), dayfirst=True, errors="coerce")
            return pd.NaT

        data_rows["ParsedDateTime"] = data_rows["Zeitraum"].apply(parse_datetime)
        data_rows.sort_values(by="ParsedDateTime", ascending=True, inplace=True)
        df = pd.concat([sum_row, data_rows], ignore_index=True)
        df.drop(columns=["ParsedDateTime"], inplace=True, errors="ignore")
    except Exception as e:
        print("Fehler beim Sortieren nach Datum/Uhrzeit:", e)

    # -------------------------------------------------------------
    # 2) Kopie des Templates anlegen
    # -------------------------------------------------------------
    temp_copy_path = output_path + "_temp_template.xlsx"
    shutil.copyfile(template_path, temp_copy_path)

    # -------------------------------------------------------------
    # 3) Kopie laden und erste Zeile löschen
    # -------------------------------------------------------------
    wb = openpyxl.load_workbook(temp_copy_path)
    ws = wb.active
    ws.delete_rows(1)

    # Dichtungen-Sortierung
    final_dichtungen = final_sort_dichtungen(user_dichtungen, df)

    # ------------------------------------------------------------
    # Kopf: Service Techniker & Datum von-bis
    # ------------------------------------------------------------
    serv_val = safe_val(df, "Service Techniker", 3)
    ws.cell(row=SERVICE_TECHNIKER_ROW, column=2, value=serv_val).font = Font(name="Calibri", size=14, bold=True)

    zr = get_zeitraum_von_bis(df, "Zeitraum")
    ws.cell(row=DATE_ROW, column=2, value=zr).font = Font(name="Calibri", size=14, bold=True)

    # Mapping Eingabe -> Template
    global_mainfield = [
        ("Zeitraum", 2),                     # B
        ("Dealname", 3),                     # C
        ("Weitere Techniker", 4),            # D
        ("Informationen Packliste", 6),      # F
        ("Ersatzteil und Zubehör", 7)        # G
    ]

    # ------------------------------------------------------------
    # Dichtungen ab Spalte E
    # ------------------------------------------------------------
    current_col = PLATZHALTER_COL_INDEX
    first_run = True
    dicht_col_map = {}

    for dicht in final_dichtungen:
        dicht_name = dicht["name"]
        is_standard = dicht.get("always_show", False)
        if (not is_standard) and (spalte_leer(df, dicht_name)):
            continue

        if first_run:
            used_col = current_col
            first_run = False
        else:
            new_col = current_col + 1
            ws.insert_cols(new_col)
            copy_column_with_style(ws, PLATZHALTER_COL_INDEX, new_col)
            for i, (dfcol, cidx) in enumerate(global_mainfield):
                if cidx >= new_col:
                    global_mainfield[i] = (dfcol, cidx + 1)
            used_col = new_col
            current_col = new_col

        set_column_left_border(ws, used_col, start_row=1, border_style='thin')

        # Kopfzelle Dichtungsname
        mod_name = apply_dicht_name_break(dicht_name)
        head_cell = ws.cell(row=TEMPLATE_DICHTUNG_NAME_ROW, column=used_col, value=mod_name)
        head_cell.font = Font(name="Calibri", size=12, bold=True)
        head_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Summe in Zeile 1
        sum_val_raw = safe_val(df, dicht_name, DF_SUM_ROW)
        sum_num = parse_number(sum_val_raw)
        sum_cell = ws.cell(row=TEMPLATE_SUM_ROW, column=used_col, value=round(sum_num))
        sum_cell.number_format = "0"
        sum_cell.font = Font(name="Calibri", size=16, color="FF0000")
        sum_cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

        dicht_col_map[dicht_name] = used_col

    set_bottom_solid(ws, TEMPLATE_DICHTUNG_NAME_ROW)

    # ------------------------------------------------------------
    # Datenzeilen
    # ------------------------------------------------------------
    t_row = TEMPLATE_DATA_START_ROW

    for df_row in range(DF_DATA_START_ROW, len(df)):
        if t_row > ws.max_row:
            ws.insert_rows(idx=t_row)
        copy_entire_row_format(ws, TEMPLATE_DATA_START_ROW, t_row)

        row_num = df_row
        num_cell = ws.cell(row=t_row, column=NUMBERING_COL, value=row_num)
        num_cell.font = Font(name="Calibri", size=12, bold=True)
        num_cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

        for (df_col, tmplt_col) in global_mainfield:
            val = safe_val(df, df_col, df_row)

            if df_col == "Zeitraum":
                val = transform_zeitraum(val)
                parts = val.split(" ", 1)
                if len(parts) == 2:
                    wtag, rest = parts
                    rest = " " + rest
                else:
                    wtag = val
                    rest = ""

                bold_inline = InlineFont(rFont="Calibri", sz=12, b=True)
                normal_inline = InlineFont(rFont="Calibri", sz=12, b=False)

                rt = CellRichText()
                rt.append(TextBlock(bold_inline, wtag))
                rt.append(TextBlock(normal_inline, rest))

                cell = ws.cell(row=t_row, column=tmplt_col)
                cell.value = rt
            else:
                cell = ws.cell(row=t_row, column=tmplt_col, value=val)
                if df_col in ["Informationen Packliste", "Ersatzteil und Zubehör", "Weitere Techniker"]:
                    cell.font = Font(bold=True, color="FF0000")
                else:
                    cell.font = Font(name="Calibri", size=12, bold=False, color="000000")

            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        for dicht in final_dichtungen:
            dicht_name = dicht["name"]
            is_standard = dicht.get("always_show", False)
            if (not is_standard) and (spalte_leer(df, dicht_name)):
                continue
            col_idx = dicht_col_map.get(dicht_name)
            if col_idx is not None:
                raw_val = df[dicht_name].iloc[df_row] if dicht_name in df.columns else ""
                try:
                    num_val = float(raw_val)
                    cell_value = round(num_val)
                    cell = ws.cell(row=t_row, column=col_idx, value=cell_value)
                    cell.number_format = "0"
                except Exception:
                    cell = ws.cell(row=t_row, column=col_idx, value=raw_val)
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.font = Font(name="Calibri", size=12, bold=False)

        if t_row == TEMPLATE_DATA_START_ROW:
            set_top_border_solid(ws, t_row)
            set_horizontal_dotted(ws, t_row)
        else:
            set_horizontal_dotted(ws, t_row)

        bg_color = "DDDDDD" if (row_num % 2 == 1) else "FFFFFF"
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=t_row, column=col_idx).fill = PatternFill("solid", fgColor=bg_color)

        t_row += 1

    # ------------------------------------------------------------
    # zusätzliche Dichtungen-Zeile
    # ------------------------------------------------------------
    extra_line_row = t_row
    ws.insert_rows(idx=extra_line_row)
    copy_entire_row_format(ws, TEMPLATE_DATA_START_ROW, extra_line_row)

    set_top_border_solid(ws, extra_line_row)
    set_bottom_thick(ws, extra_line_row)

    extra_text_cell = ws.cell(row=extra_line_row, column=3, value="zusätzliche Dichtungen")
    extra_text_cell.font = Font(bold=True)
    extra_text_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    last_used_df_row = len(df)
    bg_color = "DDDDDD" if (last_used_df_row % 2 == 1) else "FFFFFF"
    for col_idx in range(1, ws.max_column + 1):
        ws.cell(row=extra_line_row, column=col_idx).fill = PatternFill("solid", fgColor=bg_color)

    for dicht in final_dichtungen:
        if dicht.get("always_show", False):
            fix_value = dicht.get("default_value", 0)
            col_idx = dicht_col_map.get(dicht["name"])
            if col_idx is not None:
                c = ws.cell(row=extra_line_row, column=col_idx, value=fix_value)
                c.number_format = "0"
                c.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                c.font = Font(name="Calibri", size=12, bold=False)

                old_sum = ws.cell(row=TEMPLATE_SUM_ROW, column=col_idx).value
                new_sum = (old_sum if old_sum else 0) + fix_value
                s_cell = ws.cell(row=TEMPLATE_SUM_ROW, column=col_idx, value=new_sum)
                s_cell.number_format = "0"
                s_cell.font = Font(name="Calibri", size=16, bold=False, color="FF0000")
                s_cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

    # Spaltenbreiten für Infos/Ersatzteile wiederherstellen
    for field, orig_width in [
        ("Informationen Packliste", original_width_info),
        ("Ersatzteil und Zubehör", original_width_ersatz)
    ]:
        col_idx = next((col for df_field, col in global_mainfield if df_field == field), None)
        if col_idx is not None and orig_width:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = orig_width

    # Spalten ausblenden, wenn leer
    if spalte_leer(df, "Weitere Techniker"):
        ws.column_dimensions['D'].hidden = True
    else:
        ws.column_dimensions['D'].hidden = False

    ersatz_col_idx = next(
        (col for (df_field, col) in global_mainfield if df_field == "Ersatzteil und Zubehör"),
        None
    )
    if ersatz_col_idx is not None:
        col_letter = get_column_letter(ersatz_col_idx)
        if spalte_leer(df, "Ersatzteil und Zubehör"):
            ws.column_dimensions[col_letter].hidden = True
        else:
            ws.column_dimensions[col_letter].hidden = False

    remove_trailing_blank_rows(ws, extra_line_row)

    # Farben der Dichtungs-Spalten abwechselnd blau / schwarz
    BLUE_COLOR = "0000FF"
    BLACK_COLOR = "000000"

    dicht_spalten_sorted = sorted(dicht_col_map.values())

    for i, col_idx in enumerate(dicht_spalten_sorted):
        font_color_to_use = BLUE_COLOR if i % 2 == 0 else BLACK_COLOR
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            new_font = copy(cell.font)
            new_font.color = font_color_to_use
            cell.font = new_font

    wb.save(output_path)
    wb.close()

    if os.path.exists(temp_copy_path):
        os.remove(temp_copy_path)

    if AUTO_FIT_COLUMNS:
        dicht_spalten = list(dicht_col_map.values())
        auto_fit_dichtungen_in_excel(output_path, dicht_spalten)

    if show_message:
        messagebox.showinfo("Konvertierung abgeschlossen!", f"Datei wurde gespeichert unter:\n{output_path}")


# ----------------------------------------------------------------
# Sound / Druck / GUI-Funktionen (für Desktop-GUI; Web braucht das nicht)
# ----------------------------------------------------------------
# (hier lasse ich deinen bestehenden Code im Prinzip unverändert – wichtig
#  ist nur, dass er erst genutzt wird, wenn HEADLESS=False und tkinter da ist)

def play_converted_sound():
    if SOUND_ENABLED and os.name == "nt":
        try:
            winsound.Beep(800, 200)
        except Exception:
            pass


def play_print_sound():
    if SOUND_ENABLED and os.name == "nt":
        try:
            winsound.Beep(600, 300)
        except Exception:
            pass


# ... HIER KANNST DU DEINEN bisherigen GUI-/Einstellungs-/Tutorial-Code
# ... unverändert lassen, da er nur läuft, wenn HEADLESS=False ist.
# (Um die Antwort nicht noch länger zu machen, habe ich diese
#  restlichen GUI-Funktionen abgeschnitten. Für Render / Flask sind
#  sie nicht relevant.)

# Du kannst am Ende deiner Datei folgendes stehen lassen:
if __name__ == '__main__' and not HEADLESS and _TK_AVAILABLE:
    load_settings()
    root = tk.Tk()
    check_for_updates()
    root.title("Packliste Converter")
    apply_ttk_style(current_theme)
    # hier dein kompletter GUI-Aufbau...
    root.mainloop()
