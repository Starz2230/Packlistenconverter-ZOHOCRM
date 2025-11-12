#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys, os, re, json, shutil, tempfile, tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import math


from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText


import pandas as pd
import openpyxl
from copy import copy
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import re

# Füge dies direkt nach deinen Imports (z.B. nach "import re" etc.) ein:
import requests, subprocess, sys, tkinter as tk
from packaging import version

GITHUB_API = "https://api.github.com/repos/Starz2230/MeinePacklistenApp/releases/latest"
CURRENT_VERSION = "1.1.9"   # unbedingt anpassen, wenn du neu releast!

def resource_path(rel):
    import sys, os
    try:
        base = sys._MEIPASS
    except:
        base = os.path.abspath(".")
    return os.path.join(base, rel)

def check_for_updates():
    try:
        r = requests.get(GITHUB_API, timeout=5)
        data = r.json()
        latest = data["tag_name"].lstrip("v")
        if version.parse(latest) > version.parse(CURRENT_VERSION):
            url = data["assets"][0]["browser_download_url"]
            if tk.messagebox.askyesno(
                  "Update verfügbar",
                  f"Version {latest} ist verfügbar.\nJetzt aktualisieren?"):
                updater = resource_path("updater.exe")
                subprocess.Popen([updater, sys.executable, url])
                sys.exit(0)
    except Exception as e:
        print("Update-Check fehlgeschlagen:", e)

# Rufe das direkt beim Programmstart auf:
check_for_updates()

CURRENT_VERSION = "1.2.0"  # Setze hier die aktuelle Version deines Converters

import requests  # Stelle sicher, dass du das Modul "requests" installiert hast!

def get_server_version(server_ip, port=8000):
    """
    Lädt die Versionsnummer vom Update-Server (version.txt) herunter.
    """
    url = f"http://{server_ip}:{port}/version.txt"
    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            return response.text.strip()  # z. B. "1.1.0"
    except Exception as e:
        print("Update-Check fehlgeschlagen:", e)
    return None

def is_update_available(current, server_version):
    """
    Vergleicht zwei Versionsstrings. (Für robustere Vergleiche kannst du z.B. 'packaging.version' einsetzen.)
    """
    return server_version > current

def launch_updater():
    """
    Startet den separaten Updater-Prozess (updater.exe), der in demselben Installationsordner liegt.
    Übergibt als Parameter den Pfad zur aktuellen Anwendung (EXE) und beendet dann den Converter.
    """
    import subprocess, sys, os
    updater_path = os.path.join(get_app_directory(), "updater.exe")
    subprocess.Popen([updater_path, sys.executable])
    sys.exit(0)

def check_for_updates():
    """
    Prüft, ob ein Update verfügbar ist, und fragt den Benutzer (über ein Messagebox) zur Bestätigung.
    Falls bestätigt, wird der Updater gestartet. 
    """
    # Setze hier die IP deines Update-Servers ein (z.B. "192.168.1.100")
    server_ip = "192.168.1.127"
    server_version = get_server_version(server_ip)
    if server_version and is_update_available(CURRENT_VERSION, server_version):
        if messagebox.askyesno("Update verfügbar", 
                               f"Ein neues Update ({server_version}) ist verfügbar.\nMöchtest du jetzt aktualisieren?"):
            launch_updater()


# Mapping Wochentage (Montag=0, Dienstag=1, ...)
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
    und wandelt ihn um zu "FR 21.03.25 08:00 - 09:00"
    (Wochentag abgekürzt + Jahr gekürzt).
    """
    if not val or not isinstance(val, str):
        return val  # Wenn leer oder kein String, nichts ändern.

    # Versuche per Regex den vorderen Datumsanteil (DD.MM.YYYY) zu extrahieren
    # und den Rest (z.B. " 08:00 - 09:00") zu erhalten.
    m = re.match(r'^(\d{1,2}\.\d{1,2}\.\d{4})(.*)$', val.strip())
    if not m:
        return val  # Wenn nicht passend, einfach Original zurückgeben.

    date_str = m.group(1).strip()  # z.B. "21.03.2025"
    rest = m.group(2)              # z.B. " 08:00 - 09:00"

    try:
        # Parse das Datum
        dt = datetime.datetime.strptime(date_str, "%d.%m.%Y")
        # Ermittle den Wochentag (Montag=0, ..., Sonntag=6)
        wday = dt.weekday()  # integer                                            
        wday_abbr = weekday_map.get(wday, "")  # z.B. "FR"

        # Jahr auf 2 Ziffern (z.B. 25 statt 2025)
        date_formatted = dt.strftime(f"%d.%m.%y")  # "21.03.25"

        # Endgültiger String: "FR 21.03.25" + Rest
        return f"{wday_abbr} {date_formatted}{rest}"
    except:
        # Falls das Parsen fehlschlägt, Original zurückgeben.
        return val

try:
    import win32com.client
except ImportError:
    win32com = None

if os.name == "nt":
    import winsound

def resource_path(relative_path):
    """
    Gibt den absoluten Pfad zurück, unter Berücksichtigung
    der PyInstaller-Umgebung (sys._MEIPASS).
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_app_directory():
    """
    Gibt das Verzeichnis der aktuell laufenden Applikation zurück.
    Unterscheidet zwischen gefrorenem Zustand (PyInstaller) und normalem Skript.
    """
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

TEMPLATE_FILE = "Packliste_Template.xlsx"
ICON_FILE = "convert.ico"
DICHTUNGEN_CONFIG = "dichtungen.json"



# ----------------------------------------------------------------
# Zeilen/Spalten-Definitionen für das Template
# ----------------------------------------------------------------

# Nach Löschen der 1. Zeile:
SERVICE_TECHNIKER_ROW = 1   # "Service Techniker" in Zeile 1
DATE_ROW = 2                # "Datum von bis" in Zeile 2

TEMPLATE_SUM_ROW = 1             # Summen in Zeile 1
TEMPLATE_DICHTUNG_NAME_ROW = 2   # Dichtungsnamen in Zeile 2
TEMPLATE_DATA_START_ROW = 3      # Daten ab Zeile 3

DF_DATA_START_ROW = 1  # In der Eingabedatei: ab welcher Zeile fangen "Daten" an?
DF_SUM_ROW = 0         # In der Eingabedatei: in welcher Zeile steht der Summenwert?

# Ab welcher Spalte sollen die Dichtungen eingefügt werden? (E = 5)
PLATZHALTER_COL_INDEX = 5

# Sollen Zeilennummern in Spalte A geschrieben werden?
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
AUTO_FIT_COLUMNS = True  # Neuer Schalter: Aktiviert standardmäßig den AutoFit

SETTINGS_CONFIG = "settings.json"


def load_settings():
    """
    Lädt die Einstellungen aus der settings.json (sofern vorhanden).
    """
    global current_theme, SOUND_ENABLED, PRINT_SETTINGS, PRINTER_COMMAND, AUTO_FILENAME, SAVE_FOLDER, GROUP_BY_SUFFIX, AUTO_FIT_COLUMNS
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
    """
    Speichert die aktuellen Einstellungen in settings.json.
    """
    global current_theme, SOUND_ENABLED, PRINT_SETTINGS, PRINTER_COMMAND, AUTO_FILENAME, SAVE_FOLDER, GROUP_BY_SUFFIX, AUTO_FIT_COLUMNS
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


# Farbschemata für das GUI
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
    """
    Wendet das definierte Farbschema auf das ttk-Style-Objekt an.
    """
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

def load_dichtungen():
    """
    Lädt die Dichtungen aus der Datei dichtungen.json.
    Falls die Datei nicht existiert, wird sie mit DEFAULT_DICHTUNGEN angelegt.
    """
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
    """
    Speichert die Liste der Dichtungen in dichtungen.json.
    """
    app_dir = get_app_directory()
    path = os.path.join(app_dir, DICHTUNGEN_CONFIG)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dichtungen_list, f, ensure_ascii=False, indent=2)

def copy_cell_style(src_cell, dst_cell):
    """
    Kopiert den Zellstil (Font, Border, Fill, Numberformat, Alignment, etc.)
    von src_cell nach dst_cell.
    """
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
    """
    Kopiert das gesamte Zeilenformat von src_row_idx nach dst_row_idx.
    """
    ws.row_dimensions[dst_row_idx].height = None
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        sc = ws.cell(row=src_row_idx, column=col)
        dc = ws.cell(row=dst_row_idx, column=col)
        copy_cell_style(sc, dc)
        dc.value = None

def copy_column_with_style(ws, src_col_idx, dst_col_idx):
    """
    Kopiert eine ganze Spalte (Zellstile + Werte) von src_col_idx nach dst_col_idx.
    """
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
    """
    Gibt den Wert aus df[col].iloc[index] zurück, oder '' wenn es nicht existiert/NaN ist.
    """
    if col not in df.columns:
        return ""
    if index < 0 or index >= len(df):
        return ""
    val = df[col].iloc[index]
    if pd.isna(val):
        return ""
    return str(val)

def parse_number(s):
    """
    Versucht, einen String in eine float zu parsen.
    Kommt z.B. bei Summen-Spalten zum Einsatz.
    """
    try:
        return float(s.replace(",", "."))
    except:
        return 0.0

def parse_date_part(value):
    """
    Extrahiert aus einem String ein Datum (DD.MM.YYYY).
    Gibt pd.Timestamp zurück oder None.
    """
    if not value or not isinstance(value, str):
        return None
    match = re.match(r"^(\d{1,2}\.\d{1,2}\.\d{4})", value.strip())
    if not match:
        return None
    date_str = match.group(1)
    try:
        dt = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
        return dt
    except:
        return None

def get_zeitraum_von_bis(df, col="Zeitraum"):
    """
    Durchsucht die Spalte 'Zeitraum' und gibt das minimale und maximale Datum
    als 'DD.MM.YYYY - DD.MM.YYYY' zurück.
    """
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
    """
    Prüft, ob eine Spalte in einem DataFrame komplett leer ist
    (kein Wert oder nur leere Strings/NaN).
    """
    if colname not in dataf.columns:
        return True
    col_series = dataf[colname].dropna().astype(str).str.strip()
    if len(col_series) == 0:
        return True
    return col_series.eq("").all()

def generate_auto_filename(df):
    """
    Erzeugt einen Dateinamen aus 'Service Techniker' + Zeitraum
    (z.B. 'Name_01.01.2023-05.01.2023.xlsx').
    Falls die Datei schon existiert, wird eine Zählung angehängt.
    """
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
    """
    Sucht z.B. '_S' oder '_W' am Ende eines Dichtungsnamens.
    """
    parts = name.rsplit("_", 1)
    if len(parts) == 2 and parts[1].strip():
        return parts[1].strip()
    return ""

def dichtung_sort_key(d, df):
    """
    Sortier-Logik für Dichtungen (erst always_show, dann Suffix, dann numerisch).
    """
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
        except:
            manual_int = 999999
        return (0, color, manual_int, df_pos)
    else:
        m = re.search(r"(\d+)(?:/(\d+))?", d["name"])
        if m:
            first = int(m.group(1))
            second = int(m.group(2)) if m.group(2) else 0
            extracted = first + second/1000.0
        else:
            extracted = 999999
        return (1, color, extracted, df_pos)

def group_by_suffix_order(dichtungen, df):
    """
    Gruppiert Dichtungen nach Suffix (_S, _W, Rest).
    """
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
    """
    Sortiert so, dass 'always_show' Dichtungen zuerst kommen, dann alle anderen.
    """
    standard = [d for d in dichtungen if d.get("always_show", False)]
    nonstandard = [d for d in dichtungen if not d.get("always_show", False)]
    standard_sorted = sorted(standard, key=lambda d: dichtung_sort_key(d, df))
    nonstandard_sorted = sorted(nonstandard, key=lambda d: dichtung_sort_key(d, df))
    return standard_sorted + nonstandard_sorted

def order_dichtungen_by_manual(dichtungen):
    """
    Ordnet die Dichtungen so, dass Einträge mit einem manuell gesetzten
    'order'-Wert in den entsprechenden Slot (1-indexiert) eingefügt werden.
    Die übrigen (ohne Order) behalten ihre ursprüngliche Reihenfolge.
    Falls der manuelle Order-Wert größer als die Anzahl Dichtungen ist,
    wird das 'n' auf diesen Maximalwert erweitert.
    """
    manual_dict = {}
    non_manual = []

    # 1) Sammle Einträge mit manuellem Order in manual_dict
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

    # 2) Bestimme n = max(Anzahl Dichtungen, höchster Order)
    n = len(dichtungen)
    if manual_dict:
        max_order = max(manual_dict.keys())
        if max_order > n:
            n = max_order

    # 3) Fülle Slots 1..n
    result = []
    non_manual_index = 0
    for slot in range(1, n + 1):
        if slot in manual_dict and manual_dict[slot]:
            # Falls mehrere Einträge denselben Slot haben, hänge sie in ihrer Ursprungs-Reihenfolge an
            result.extend(manual_dict[slot])
        else:
            # Falls kein manueller Eintrag für diesen Slot, nimm den nächsten aus non_manual
            if non_manual_index < len(non_manual):
                result.append(non_manual[non_manual_index])
                non_manual_index += 1

    # 4) Falls danach noch nicht-manuale übrig sind, hänge sie hinten an
    while non_manual_index < len(non_manual):
        result.append(non_manual[non_manual_index])
        non_manual_index += 1

    return result

# ----------------------------------------------------------------
# Linien-Funktionen
# ----------------------------------------------------------------

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
    """
    Löscht von unten nach oben alle komplett leeren Zeilen,
    beginnend ab 'start_row'.
    """
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
    """
    Öffnet die Excel-Datei xlsx_path in Excel und führt
    für jede Spalte in col_indices ein 'AutoFit' aus.
    Dann speichert und schließt Excel.
    Funktioniert nur unter Windows mit installiertem Excel + pywin32.
    """
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
# Hier passiert die eigentliche Umwandlung (convert_file)
# ----------------------------------------------------------------
def apply_dicht_name_break(dicht_name: str) -> str:
    """
    Bricht einen Dichtungsnamen in zwei Zeilen um, wenn er zu lang ist oder ein '_' enthält.
    """
    import math
    if "_" in dicht_name:
        idx = dicht_name.index("_")
        return dicht_name[:idx] + "\n" + dicht_name[idx+1:]
    else:
        name = dicht_name.strip()
        one_line_width = (len(name) + 1) * 1.1
        two_line_width = (math.ceil(len(name) / 2) + 1) * 1.1
        # Falls ein Zeilenumbruch in der Mitte kürzer ist, brich den Namen in zwei Zeilen
        if two_line_width < one_line_width:
            half = math.ceil(len(name) / 2)
            return name[:half] + "\n" + name[half:]
        else:
            return name
import math
import re

def parse_numeric_part(name: str) -> float:
    """
    Sucht in 'name' nach einer Zahl oder einem Bruch (z.B. 6/4).
    Gibt eine float zurück (z.B. 6.004) oder 999999 wenn nichts gefunden.
    """
    m = re.search(r'(\d+)(?:/(\d+))?', name)
    if m:
        first = int(m.group(1))
        second = int(m.group(2)) if m.group(2) else 0
        return first + second/1000.0
    return 999999

def parse_suffix_priority(name: str) -> int:
    """
    Unterscheidet bekannte Suffixe (_S, _W, _G, etc.) und gibt
    eine Priorität zurück (0,1,2...). Nicht erkannte Suffixe => 99.
    """
    # Wir holen z.B. den letzten Teil nach dem Unterstrich
    sfx = ""
    parts = name.rsplit("_", 1)
    if len(parts) == 2 and parts[1].strip():
        sfx = parts[1].strip().upper()

    # Ordne bekannte Farben / Suffixe eine Reihenfolge zu:
    if sfx == "S":
        return 0
    elif sfx == "W":
        return 1
    elif sfx == "G":
        return 2
    else:
        return 99

def final_sort_dichtungen(dichtungen, df):
    """
    Sortiert die Dichtungen so, dass:
    1) Standard-Dichtungen (always_show=True) zuerst kommen,
       - innerhalb der Standard-Dichtungen: zuerst nach order sortieren,
         dann ggf. nach Zahlen/Alphabet
    2) Nicht-Standard-Dichtungen danach
       - Gruppierung nach Suffix (_S, _W, _G, ...)
       - innerhalb der Gruppe numerisch (z.B. 6/4 vor 8/4)
       - wenn kein Suffix, am Ende
       - wenn kein numerischer Teil, fallback alphabetisch
    """
    def sort_key(d):
        # 1) Ist es eine Standard-Dichtung?
        is_std = d.get("always_show", False)

        # 2) Hat sie einen 'order'-Wert?
        order_str = str(d.get("order", "")).strip()
        try:
            order_val = int(order_str)
        except ValueError:
            order_val = None

        name = d["name"]
        # Für die Nicht-Standard sortieren wir nach Suffix + Numerik + Name
        suffix_prio = parse_suffix_priority(name)
        numeric_val = parse_numeric_part(name)
        alpha_name = name.lower()

        if is_std:
            # Gruppe 0 = Standard
            group = 0
            # Untergruppe: Standard mit definiertem order zuerst
            if order_val is not None:
                # subGroup=0 => "hat order"
                # subSort = order_val => je kleiner, desto eher
                return (group, 0, order_val)
            else:
                # subGroup=1 => "kein order"
                # wir können hier numeric_val + alpha_name anfügen,
                # damit Standard ohne order intern nochmal sortiert wird
                return (group, 1, numeric_val, alpha_name)
        else:
            # Gruppe 1 = Nicht-Standard
            # subGroup = suffix_prio
            # numeric_val => z.B. 6/4 => 6.004
            # alpha_name => fallback
            group = 1
            return (group, suffix_prio, numeric_val, alpha_name)

    # Jetzt sortieren wir die Liste mit diesem Schlüssel
    return sorted(dichtungen, key=sort_key)

def convert_file(input_path, output_path, user_dichtungen, show_message=True):
    template_path = resource_path(TEMPLATE_FILE)
    if not os.path.isfile(template_path):
        messagebox.showerror("Fehler", f"'{TEMPLATE_FILE}' wurde nicht gefunden!")
        return

    # 1) Original-Template öffnen und Spaltenbreiten sichern (F=Infos, G=Ersatzteile)
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
# Summenzeile (row=0) abtrennen und nur den Rest sortieren
# Jetzt inklusive Uhrzeit
# -------------------------------------------------------------
    try:
    # 1) Summenzeile separat speichern (Kopie von Zeile 0)
        sum_row = df.iloc[[0]].copy()

    # 2) Die eigentlichen Datenzeilen ab Zeile 1
        data_rows = df.iloc[1:].copy()

    # 3) Datum + Uhrzeit aus der Spalte "Zeitraum" parsen
        def parse_datetime(x):
        # Beispiel-Format: "31.03.2025 08:00 - 09:00"
        # Wir holen uns den ersten Zeitblock "DD.MM.YYYY HH:MM"
        # und parsen das in ein Datetime-Objekt
            pattern = r'^(\d{1,2}\.\d{1,2}\.\d{4})\s+(\d{1,2}:\d{1,2})'
            m = re.match(pattern, str(x))
            if m:
                dt_str = f"{m.group(1)} {m.group(2)}"
                return pd.to_datetime(dt_str, format='%d.%m.%Y %H:%M', errors='coerce')
            else:
            # Falls keine Uhrzeit enthalten, nur Datum
                m2 = re.match(r'^(\d{1,2}\.\d{1,2}\.\d{4})', str(x))
                if m2:
                    return pd.to_datetime(m2.group(1), dayfirst=True, errors='coerce')
            return pd.NaT

    # Neue Hilfsspalte für die sortierbare Zeit
        data_rows['ParsedDateTime'] = data_rows['Zeitraum'].apply(parse_datetime)

    # 4) Nach dieser Hilfsspalte aufsteigend sortieren
        data_rows.sort_values(by='ParsedDateTime', ascending=True, inplace=True)

    # 5) Summenzeile wieder oben anfügen (Zeile 0)
        df = pd.concat([sum_row, data_rows], ignore_index=True)

    # 6) Hilfsspalte wieder entfernen
        df.drop(columns=['ParsedDateTime'], inplace=True, errors='ignore')

    except Exception as e:
        print("Fehler beim Sortieren nach Datum/Uhrzeit:", e)

# -------------------------------------------------------------
# 2) Kopie der Template-Datei anlegen
# -------------------------------------------------------------
    temp_copy_path = output_path + "_temp_template.xlsx"
    shutil.copyfile(template_path, temp_copy_path)

# -------------------------------------------------------------
# 3) Kopie laden und erste Zeile löschen
# -------------------------------------------------------------
    wb = openpyxl.load_workbook(temp_copy_path)
    ws = wb.active
    ws.delete_rows(1)

# Sortierung der Dichtungen (abhängig von GROUP_BY_SUFFIX)
    try:
        from __main__ import GROUP_BY_SUFFIX
    except:
        pass

# Prüfe, ob mindestens ein Eintrag einen manuellen Order-Wert hat
    final_dichtungen = final_sort_dichtungen(user_dichtungen, df)





    # ------------------------------------------------------------
    # Schreibe "Service Techniker" in Zeile 1, Spalte 2 (B1)
    # Schreibe "Datum von bis"     in Zeile 2, Spalte 2 (B2)
    # ------------------------------------------------------------
    serv_val = safe_val(df, "Service Techniker", 3)
    ws.cell(row=SERVICE_TECHNIKER_ROW, column=2, value=serv_val).font = Font(name="Calibri", size=14, bold=True)

    zr = get_zeitraum_von_bis(df, "Zeitraum")
    ws.cell(row=DATE_ROW, column=2, value=zr).font = Font(name="Calibri", size=14, bold=True)

    # ------------------------------------------------------------
    # Spalten-Mapping aus der Eingabedatei -> Template
    #
    #   ("Zeitraum", 2)                -> Spalte B
    #   ("Dealname", 3)                -> Spalte C
    #   ("Weitere Techniker", 4)       -> Spalte D
    #   ("Informationen Packliste", 6) -> Spalte F
    #   ("Ersatzteil und Zubehör", 7)  -> Spalte G
    # ------------------------------------------------------------
    global_mainfield = [
        ("Zeitraum", 2),                     # B
        ("Dealname", 3),                     # C
        ("Weitere Techniker", 4),            # D
        ("Informationen Packliste", 6),      # F
        ("Ersatzteil und Zubehör", 7)        # G
    ]

    # ------------------------------------------------------------
    # Dichtungen ab Spalte E (col=5)
    # Zeile 2 = Summen, Zeile 3 = Dichtungsnamen, ab Zeile 4 = Daten
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
            # Falls in global_mainfield Spalten >= new_col existieren, +1
            for i, (dfcol, cidx) in enumerate(global_mainfield):
                if cidx >= new_col:
                    global_mainfield[i] = (dfcol, cidx + 1)
            used_col = new_col
            current_col = new_col

        # linker Rand
        set_column_left_border(ws, used_col, start_row=1, border_style='thin')

        # Kopfzelle in Zeile 2 (Dichtungsname fett)
        mod_name = apply_dicht_name_break(dicht_name)
        head_cell = ws.cell(row=TEMPLATE_DICHTUNG_NAME_ROW, column=used_col, value=mod_name)
        head_cell.font = Font(name="Calibri", size=12, bold=True)
        head_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Summenzelle in Zeile 1 (rot, fett, Größe 16)
        sum_val_raw = safe_val(df, dicht_name, DF_SUM_ROW)
        sum_num = parse_number(sum_val_raw)
        sum_cell = ws.cell(row=TEMPLATE_SUM_ROW, column=used_col, value=round(sum_num))
        sum_cell.number_format = "0"
        sum_cell.font = Font(name="Calibri", size=16, color="FF0000")  # rot, fett, größer
        sum_cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

        dicht_col_map[dicht_name] = used_col

    # dünne Linie unter Zeile 2 (Dichtungsnamen)
    set_bottom_solid(ws, TEMPLATE_DICHTUNG_NAME_ROW)

    # ------------------------------------------------------------
    # Datenzeilen füllen (Zeile 3 ff.)
    # ------------------------------------------------------------
    t_row = TEMPLATE_DATA_START_ROW
    
    for df_row in range(DF_DATA_START_ROW, len(df)):
        if t_row > ws.max_row:
            ws.insert_rows(idx=t_row)
        copy_entire_row_format(ws, TEMPLATE_DATA_START_ROW, t_row)

        row_num = df_row
        # Nummerierung in Spalte A (wenn gewünscht)
        num_cell = ws.cell(row=t_row, column=NUMBERING_COL, value=row_num)
        num_cell.font = Font(name="Calibri", size=12, bold=True)
        num_cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)

            # Hauptfelder füllen
        for (df_col, tmplt_col) in global_mainfield:
            val = safe_val(df, df_col, df_row)

            if df_col == "Zeitraum":
        # Transformiere zuerst den String (z.B. "FR 21.03.25 08:00 - 09:00")
                val = transform_zeitraum(val)
        # Zerlege den String in den ersten Teil (Wochentag) und den Rest
                parts = val.split(" ", 1)
                if len(parts) == 2:
                    wtag, rest = parts
                    rest = " " + rest  # damit der Leerraum erhalten bleibt
                else:
                    wtag = val
                    rest = ""
        
        # Erstelle InlineFont-Objekte gemäß der offiziellen Anleitung
                bold_inline = InlineFont(rFont="Calibri", sz=12, b=True)
                normal_inline = InlineFont(rFont="Calibri", sz=12, b=False)
        
        # Erzeuge ein Rich-Text-Objekt
                rt = CellRichText()
                rt.append(TextBlock(bold_inline, wtag))
                rt.append(TextBlock(normal_inline, rest))
        
        # Weise den Rich-Text-String der Zelle zu (keine weitere cell.font-Zuweisung!)
                cell = ws.cell(row=t_row, column=tmplt_col)
                cell.value = rt
            else:
        # Für alle anderen Spalten: normale Behandlung
                cell = ws.cell(row=t_row, column=tmplt_col, value=val)
                if df_col in ["Informationen Packliste", "Ersatzteil und Zubehör", "Weitere Techniker"]:
                    cell.font = Font(bold=True, color="FF0000")
                else:
                    cell.font = Font(name="Calibri", size=12, bold=False, color="000000")
    
            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)



        # Dichtungswerte
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
                except:
                    cell = ws.cell(row=t_row, column=col_idx, value=raw_val)
                cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
                cell.font = Font(name="Calibri", size=12, bold=False)

        # oberste Datenzeile: dünne Linie oben + dotted
        if t_row == TEMPLATE_DATA_START_ROW:
            set_top_border_solid(ws, t_row)
            set_horizontal_dotted(ws, t_row)
        else:
            set_horizontal_dotted(ws, t_row)

        # abwechselnde Hintergrundfarbe
        bg_color = "DDDDDD" if (row_num % 2 == 1) else "FFFFFF"
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=t_row, column=col_idx).fill = PatternFill("solid", fgColor=bg_color)

        t_row += 1

    # ------------------------------------------------------------
    # "zusätzliche Dichtungen" Zeile
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

    # Standard-Dichtungen vorbelegen
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

      # ------------------------------------------------------------
    # Dynamisch: Spaltenbreiten für "Informationen Packliste" und "Ersatzteil und Zubehör" aus dem Original wiederherstellen
    # ------------------------------------------------------------
    for field, orig_width in [("Informationen Packliste", original_width_info), ("Ersatzteil und Zubehör", original_width_ersatz)]:
        # Finde die Spalte im global_mainfield, wo der jeweilige Wert gespeichert ist.
        col_idx = next((col for df_field, col in global_mainfield if df_field == field), None)
        if col_idx is not None and orig_width:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = orig_width


    # Beispiel: Spalte D ausblenden, wenn "Weitere Techniker" leer
    if spalte_leer(df, "Weitere Techniker"):
        ws.column_dimensions['D'].hidden = True
    else:
        ws.column_dimensions['D'].hidden = False

     # Dynamisch herausfinden, in welcher Spalte "Ersatzteil und Zubehör" liegt:
    ersatz_col_idx = next(
        (col for (df_field, col) in global_mainfield if df_field == "Ersatzteil und Zubehör"),
        None
    )
    if ersatz_col_idx is not None:
        col_letter = get_column_letter(ersatz_col_idx)
        # Prüfen, ob die Spalte in df leer ist:
        if spalte_leer(df, "Ersatzteil und Zubehör"):
            ws.column_dimensions[col_letter].hidden = True
        else:
            ws.column_dimensions[col_letter].hidden = False


    # Leere Zeilen am Ende entfernen
    remove_trailing_blank_rows(ws, extra_line_row)
    # ------------------------------------------------------------
    # Nur die Schriftfarbe in den Dichtungs-Spalten abwechselnd
    # blau / schwarz setzen. Der Zebra-Hintergrund bleibt erhalten!
    # ------------------------------------------------------------
    from copy import copy
    
    # Wir wählen ein kräftiges Blau und Schwarz als Schriftfarbe.
    # Du kannst den Hex-Wert gerne anpassen.
    BLUE_COLOR = "0000FF"
    BLACK_COLOR = "000000"

    # Indizes der Dichtungs-Spalten sortieren (damit 1., 2., 3. Spalte etc.)
    dicht_spalten_sorted = sorted(dicht_col_map.values())

    for i, col_idx in enumerate(dicht_spalten_sorted):
        # i = 0,1,2,3,... => gerade i => blau, ungerade i => schwarz
        if i % 2 == 0:
            font_color_to_use = BLUE_COLOR
        else:
            font_color_to_use = BLACK_COLOR

        # Alle Zeilen in dieser Spalte durchgehen (inkl. Kopf, Summenzeile usw.)
        for row_idx in range(1, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)

            # Bestehendes Font-Objekt kopieren, nur die Farbe ändern
            new_font = copy(cell.font)
            new_font.color = font_color_to_use
            cell.font = new_font

            # Achtung: KEINE Änderung von cell.fill => Zebra-Streifen bleibt!

    # Speichern
    wb.save(output_path)
    wb.close()

    # Temporäre Kopie entfernen
    if os.path.exists(temp_copy_path):
        os.remove(temp_copy_path)

# AutoFit nur für Dichtungs-Spalten, wenn der Schalter aktiviert ist
    if AUTO_FIT_COLUMNS:
        dicht_spalten = list(dicht_col_map.values())
        auto_fit_dichtungen_in_excel(output_path, dicht_spalten)

    if show_message:
        messagebox.showinfo("Konvertierung abgeschlossen!", f"Datei wurde gespeichert unter:\n{output_path}")


# ----------------------------------------------------------------
# Sound-Effekte
# ----------------------------------------------------------------

def play_converted_sound():
    if SOUND_ENABLED and os.name == "nt":
        try:
            winsound.Beep(800, 200)
        except:
            pass

def play_print_sound():
    if SOUND_ENABLED and os.name == "nt":
        try:
            winsound.Beep(600, 300)
        except:
            pass

# ----------------------------------------------------------------
# Druckeinstellungen
# ----------------------------------------------------------------

def apply_win32com_print_settings(wb, settings):
    ws = wb.ActiveSheet
    def mm_to_inch(mm_str):
        try:
            val = float(mm_str)
        except:
            val = 10.0
        return val / 25.4

    topMargin = mm_to_inch(settings["margin_top"])
    bottomMargin = mm_to_inch(settings["margin_bottom"])
    leftMargin = mm_to_inch(settings["margin_left"])
    rightMargin = mm_to_inch(settings["margin_right"])
    ws.PageSetup.TopMargin = topMargin
    ws.PageSetup.BottomMargin = bottomMargin
    ws.PageSetup.LeftMargin = leftMargin
    ws.PageSetup.RightMargin = rightMargin

    paper_map = {
        "A4": 9,
        "A3": 8,
        "Letter": 1,
        "Legal": 5,
        "A5": 11
    }
    ws.PageSetup.PaperSize = paper_map.get(settings["paper_size"], 9)

    if settings.get("fit_to_page", False):
        try:
            usedRange = ws.UsedRange
            contentWidth = usedRange.Width
            contentHeight = usedRange.Height
        except Exception as e:
            contentWidth = 500
            contentHeight = 700

        pageWidth_portrait = 595
        pageHeight_portrait = 842
        availWidth_portrait = pageWidth_portrait - ((leftMargin + rightMargin) * 72)
        availHeight_portrait = pageHeight_portrait - ((topMargin + bottomMargin) * 72)

        pageWidth_landscape = pageHeight_portrait
        pageHeight_landscape = pageWidth_portrait
        availWidth_landscape = pageWidth_landscape - ((leftMargin + rightMargin) * 72)
        availHeight_landscape = pageHeight_landscape - ((topMargin + bottomMargin) * 72)

        scale_portrait = min(availWidth_portrait / contentWidth, availHeight_portrait / contentHeight)
        scale_landscape = min(availWidth_landscape / contentWidth, availHeight_landscape / contentHeight)

        if scale_landscape >= scale_portrait:
            chosen_orientation = 2
        else:
            chosen_orientation = 1

        ws.PageSetup.Orientation = chosen_orientation
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
    else:
        ws.PageSetup.Orientation = 1
        sc = settings.get("scaling", "100%").replace("%", "")
        try:
            sc_val = int(sc)
        except:
            sc_val = 100
        ws.PageSetup.Zoom = sc_val
        ws.PageSetup.FitToPagesWide = False
        ws.PageSetup.FitToPagesTall = False

def open_excel_print_preview(excel_file):
    """
    Öffnet die Windows-Druckvorschau (PrintPreviewAndPrint) mit win32com,
    sofern verfügbar.
    """
    if not win32com:
        messagebox.showerror("Druckfehler", "win32com.client nicht verfügbar – erweiterter Druck nicht möglich.")
        return
    try:
        excel = win32com.client.Dispatch("Excel.Application")
    except Exception as e:
        messagebox.showerror("Druckfehler", f"Excel konnte nicht gestartet werden.\n{e}")
        return
    try:
        excel.Visible = True
        wb = excel.Workbooks.Open(excel_file)
        apply_win32com_print_settings(wb, PRINT_SETTINGS)
        try:
            wb.Application.CommandBars.ExecuteMso("PrintPreviewAndPrint")
        except:
            wb.ActiveSheet.PrintPreview()
    except Exception as e:
        messagebox.showerror("Druckfehler", f"Fehler beim Aufrufen der Druckvorschau:\n{e}")

# ----------------------------------------------------------------
# Einstellungsdialog
# ----------------------------------------------------------------

def open_settings(root, tutorial=False):
    """
    Öffnet ein Einstellungsfenster (Thema, Sound, Drucker-Einstellungen, etc.).
    """
    global GROUP_BY_SUFFIX, AUTO_FIT_COLUMNS
    settings_win = tk.Toplevel(root)
    settings_win.title("Einstellungen")
    settings_win.geometry("700x500")

    nb = ttk.Notebook(settings_win)
    nb.pack(fill="both", expand=True, padx=10, pady=10)

    # Allgemeiner Tab
    general_frame = ttk.Frame(nb)
    nb.add(general_frame, text="Allgemeines")

    ttk.Label(general_frame, text="Farbschema:", font=("Helvetica", 12)).pack(pady=10, anchor="w")
    theme_var = tk.StringVar(value=current_theme)
    ttk.Radiobutton(general_frame, text="Dunkel", variable=theme_var, value="dark",
                    command=lambda: change_theme(theme_var.get())).pack(anchor="w", padx=20)
    ttk.Radiobutton(general_frame, text="Hell", variable=theme_var, value="light",
                    command=lambda: change_theme(theme_var.get())).pack(anchor="w", padx=20)

    sound_var = tk.BooleanVar(value=SOUND_ENABLED)
    ttk.Checkbutton(general_frame, text="Töne aktivieren", variable=sound_var,
                    command=lambda: toggle_sound(sound_var.get())).pack(pady=10, anchor="w", padx=20)

    group_suffix_var = tk.BooleanVar(value=GROUP_BY_SUFFIX)
    ttk.Checkbutton(general_frame, text="Suffix-Gruppierung aktivieren (_S, _W etc.)", variable=group_suffix_var).pack(pady=10, anchor="w", padx=20)

    auto_fit_var = tk.BooleanVar(value=AUTO_FIT_COLUMNS)
    ttk.Checkbutton(general_frame, text="Spalten-AutoFit aktivieren", variable=auto_fit_var).pack(pady=10, anchor="w", padx=20)

    # Drucker-Einstellungen Tab
    printer_frame = ttk.Frame(nb)
    nb.add(printer_frame, text="Drucker-Einstellungen")
    ttk.Label(printer_frame, text="Seitenränder (mm):").pack(anchor="w", padx=20, pady=5)
    margins_frame = ttk.Frame(printer_frame)
    margins_frame.pack(anchor="w", padx=20, pady=5)

    margin_top_var = tk.StringVar(value=PRINT_SETTINGS["margin_top"])
    margin_bottom_var = tk.StringVar(value=PRINT_SETTINGS["margin_bottom"])
    margin_left_var = tk.StringVar(value=PRINT_SETTINGS["margin_left"])
    margin_right_var = tk.StringVar(value=PRINT_SETTINGS["margin_right"])

    ttk.Label(margins_frame, text="Oben:").grid(row=0, column=0, padx=5)
    ttk.Entry(margins_frame, textvariable=margin_top_var, width=5).grid(row=0, column=1)
    ttk.Label(margins_frame, text="Unten:").grid(row=0, column=2, padx=5)
    ttk.Entry(margins_frame, textvariable=margin_bottom_var, width=5).grid(row=0, column=3)
    ttk.Label(margins_frame, text="Links:").grid(row=1, column=0, padx=5)
    ttk.Entry(margins_frame, textvariable=margin_left_var, width=5).grid(row=1, column=1)
    ttk.Label(margins_frame, text="Rechts:").grid(row=1, column=2, padx=5)
    ttk.Entry(margins_frame, textvariable=margin_right_var, width=5).grid(row=1, column=3)

    ttk.Label(printer_frame, text="Skalierung (z.B. '100%'):").pack(anchor="w", padx=20, pady=5)
    scaling_var = tk.StringVar(value=PRINT_SETTINGS["scaling"])
    ttk.Entry(printer_frame, textvariable=scaling_var, width=10).pack(anchor="w", padx=20)

    duplex_var = tk.BooleanVar(value=PRINT_SETTINGS["duplex"])
    duplex_chk = ttk.Checkbutton(printer_frame, text="Doppelseitig drucken (evtl. ohne Effekt)", variable=duplex_var)
    duplex_chk.pack(anchor="w", padx=20, pady=5)

    ttk.Label(printer_frame, text="Anzahl Kopien:").pack(anchor="w", padx=20, pady=5)
    copies_var = tk.StringVar(value=PRINT_SETTINGS["copies"])
    ttk.Entry(printer_frame, textvariable=copies_var, width=5).pack(anchor="w", padx=20)

    ttk.Label(printer_frame, text="Papierformat:").pack(anchor="w", padx=20, pady=5)
    paper_var = tk.StringVar(value=PRINT_SETTINGS["paper_size"])
    combo_paper = ttk.Combobox(printer_frame, textvariable=paper_var,
                               values=["A4", "A3", "Letter", "Legal", "A5"], width=10)
    combo_paper.pack(anchor="w", padx=20)

    ttk.Label(printer_frame, text="Druckbefehl (Windows):").pack(anchor="w", padx=20, pady=5)
    printer_cmd_var = tk.StringVar(value=PRINTER_COMMAND)
    ttk.Entry(printer_frame, textvariable=printer_cmd_var, width=10).pack(anchor="w", padx=20)

    fit_page_var = tk.BooleanVar(value=PRINT_SETTINGS.get("fit_to_page", False))
    fit_page_chk = ttk.Checkbutton(printer_frame, text="Auf eine A4 Seite im Querformat skalieren (1 Seite)", variable=fit_page_var)
    fit_page_chk.pack(anchor="w", padx=20, pady=5)

    def on_fit_page_change(*args):
        if fit_page_var.get():
            duplex_var.set(False)
            duplex_chk.config(state="disabled")
        else:
            duplex_chk.config(state="normal")
    fit_page_var.trace("w", on_fit_page_change)

    # Dateiname & Speicherort Tab
    filename_frame = ttk.Frame(nb)
    nb.add(filename_frame, text="Dateiname & Speicherort")

    auto_filename_var = tk.BooleanVar(value=AUTO_FILENAME)
    ttk.Checkbutton(filename_frame, text="Automatische Dateinamen aktivieren", variable=auto_filename_var).pack(anchor="w", padx=20, pady=5)

    ttk.Label(filename_frame, text="Speicherort:").pack(anchor="w", padx=20, pady=5)
    save_folder_var = tk.StringVar(value=SAVE_FOLDER)
    entry_folder = ttk.Entry(filename_frame, textvariable=save_folder_var, width=40)
    entry_folder.pack(anchor="w", padx=20)

    def browse_folder():
        folder = filedialog.askdirectory(title="Speicherort wählen")
        if folder:
            save_folder_var.set(folder)

    ttk.Button(filename_frame, text="Ordner wählen", command=browse_folder).pack(anchor="w", padx=20, pady=5)

    template_frame = ttk.Frame(nb)
    nb.add(template_frame, text="Vorlage")
    ttk.Label(template_frame, text="Packliste_Template.xlsx öffnen:", font=("Helvetica", 12)).pack(pady=10)

    def open_template():
        template_path = resource_path(TEMPLATE_FILE)
        if os.path.isfile(template_path):
            os.startfile(template_path)
        else:
            messagebox.showerror("Fehler", f"Template-Datei nicht gefunden:\n{template_path}")

    ttk.Button(template_frame, text="Template öffnen", command=open_template).pack(pady=5)

    def apply_settings():
        PRINT_SETTINGS["margin_top"] = margin_top_var.get().strip()
        PRINT_SETTINGS["margin_bottom"] = margin_bottom_var.get().strip()
        PRINT_SETTINGS["margin_left"] = margin_left_var.get().strip()
        PRINT_SETTINGS["margin_right"] = margin_right_var.get().strip()
        PRINT_SETTINGS["scaling"] = scaling_var.get().strip()
        PRINT_SETTINGS["duplex"] = duplex_var.get()
        PRINT_SETTINGS["copies"] = copies_var.get().strip()
        PRINT_SETTINGS["paper_size"] = paper_var.get().strip()
        PRINT_SETTINGS["fit_to_page"] = fit_page_var.get()

        global PRINTER_COMMAND, AUTO_FILENAME, SAVE_FOLDER, GROUP_BY_SUFFIX, AUTO_FIT_COLUMNS
        PRINTER_COMMAND = printer_cmd_var.get().strip() or "print"
        AUTO_FILENAME = auto_filename_var.get()
        SAVE_FOLDER = save_folder_var.get().strip()
        GROUP_BY_SUFFIX = group_suffix_var.get()
        AUTO_FIT_COLUMNS = auto_fit_var.get()

        save_settings()
        if not tutorial:
            settings_win.destroy()

    ttk.Button(settings_win, text="Übernehmen", command=apply_settings).pack(pady=10)

    if not tutorial:
        settings_win.grab_set()
        root.wait_window(settings_win)
    return settings_win


# ----------------------------------------------------------------
# Theme & Sound toggles
# ----------------------------------------------------------------

def change_theme(selected_theme):
    global current_theme
    current_theme = selected_theme
    apply_ttk_style(current_theme)

def toggle_sound(state):
    global SOUND_ENABLED
    SOUND_ENABLED = state

# ----------------------------------------------------------------
# Dichtungen verwalten
# ----------------------------------------------------------------

def manage_dichtungen_window(parent, tutorial=False):
    config = load_dichtungen()
    # Sicherstellen, dass alle Einträge ein Dict sind
    normalized = []
    for item in config:
        if isinstance(item, dict):
            normalized.append(item)
        else:
            normalized.append({"name": item, "always_show": False, "default_value": 0, "order": ""})
    config = normalized

    win = tk.Toplevel(parent)
    win.title("Dichtungen verwalten")

    bgc = light_theme["bg"] if current_theme == "light" else dark_theme["bg"]
    fgc = light_theme["fg"] if current_theme == "light" else dark_theme["fg"]
    win.configure(bg=bgc)

    # Überschrift und neuer Dichtung-Eingabe-Bereich
    top_frame = tk.Frame(win, bg=bgc)
    top_frame.pack(padx=10, pady=5, fill="x")
    tk.Label(top_frame, text="Dichtungen verwalten", font=("Helvetica", 14, "bold"), bg=bgc, fg=fgc).pack(side="left")

    add_frame = tk.Frame(win, bg=bgc)
    add_frame.pack(padx=10, pady=5, fill="x")
    tk.Label(add_frame, text="Neue Dichtung:", bg=bgc, fg=fgc, width=15).pack(side="left")
    new_entry = tk.Entry(add_frame,
                         bg=light_theme["entry_bg"] if current_theme == "light" else dark_theme["entry_bg"],
                         fg=light_theme["entry_fg"] if current_theme == "light" else dark_theme["entry_fg"])
    new_entry.pack(side="left", fill="x", expand=True)

    def add_dichtung():
        name = new_entry.get().strip()
        if name and not any(d["name"] == name for d in config):
            config.append({"name": name, "always_show": False, "default_value": 0, "order": ""})
            refresh_dichtungen_list()
        new_entry.delete(0, tk.END)

    tk.Button(add_frame, text="Hinzufügen", command=add_dichtung,
              bg=light_theme["button_bg"] if current_theme == "light" else dark_theme["button_bg"],
              fg=light_theme["button_fg"] if current_theme == "light" else dark_theme["button_fg"]
              ).pack(side="left", padx=5)

    # Container mit Scrollbar für die Liste
    container = tk.Frame(win, bg=bgc)
    container.pack(padx=10, pady=5, fill="both", expand=True)
    canvas = tk.Canvas(container, bg=bgc, highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview, width=10)
    canvas.configure(yscrollcommand=scrollbar.set)

    def _on_mousewheel(event):
        canvas.yview_scroll(-int(event.delta/120), "units")
    canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
    canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
    canvas.bind("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
    canvas.bind("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

    scrollable_frame = tk.Frame(canvas, bg=bgc)
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Kopfzeile
    header_frame = tk.Frame(scrollable_frame, bg=bgc)
    header_frame.pack(fill="x")
    tk.Label(header_frame, text="Name", bg=bgc, fg=fgc, width=20, anchor="w").pack(side="left", padx=5)
    tk.Label(header_frame, text="Standard?", bg=bgc, fg=fgc, width=10).pack(side="left", padx=5)

    row_widgets = []

    def toggle_extra_fields(var, extra_frame):
        if var.get():
            extra_frame.pack(fill="x", padx=25, pady=2)
        else:
            extra_frame.forget()

    def refresh_dichtungen_list():
        # Lösche alte Zeilen
        for widget in row_widgets:
            widget.destroy()
        row_widgets.clear()

        for d in config:
            row_container = tk.Frame(scrollable_frame, bg=bgc, bd=1, relief="solid")
            row_container.pack(fill="x", padx=5, pady=2)
            row_widgets.append(row_container)

            # Obere Zeile: Name und Checkbox
            top_row = tk.Frame(row_container, bg=bgc)
            top_row.pack(fill="x", padx=5, pady=2)
            var_name = tk.StringVar(value=d["name"])
            ent_name = tk.Entry(top_row, textvariable=var_name,
                                bg=light_theme["entry_bg"] if current_theme == "light" else dark_theme["entry_bg"],
                                fg=light_theme["entry_fg"] if current_theme == "light" else dark_theme["entry_fg"],
                                width=20)
            ent_name.pack(side="left", padx=5)
            d["var_name"] = var_name

            var_standard = tk.BooleanVar(value=d.get("always_show", False))
            chk = tk.Checkbutton(top_row, text="Standard", variable=var_standard,
                                 bg=bgc, fg=fgc)
            chk.pack(side="left", padx=5)
            d["var_standard"] = var_standard

            # Extra-Frame (versteckt/angezeigt je nach Checkbox)
            extra_frame = tk.Frame(row_container, bg=bgc)
            tk.Label(extra_frame, text="Zusatzwert:", bg=bgc, fg=fgc).pack(side="left", padx=5)
            var_value = tk.StringVar(value=str(d.get("default_value", 0)))
            ent_value = tk.Entry(extra_frame, textvariable=var_value,
                                 bg=light_theme["entry_bg"] if current_theme == "light" else dark_theme["entry_bg"],
                                 fg=light_theme["entry_fg"] if current_theme == "light" else dark_theme["entry_fg"],
                                 width=10)
            ent_value.pack(side="left", padx=5)
            d["var_value"] = var_value

            tk.Label(extra_frame, text="Reihenfolge:", bg=bgc, fg=fgc).pack(side="left", padx=5)
            var_order = tk.StringVar(value=str(d.get("order", "")))
            ent_order = tk.Entry(extra_frame, textvariable=var_order,
                                 bg=light_theme["entry_bg"] if current_theme == "light" else dark_theme["entry_bg"],
                                 fg=light_theme["entry_fg"] if current_theme == "light" else dark_theme["entry_fg"],
                                 width=10)
            ent_order.pack(side="left", padx=5)
            d["var_order"] = var_order

            # Checkbox-Callback: Zeige oder verstecke extra_frame
            def on_chk_change(v=var_standard, ef=extra_frame):
                toggle_extra_fields(v, ef)
            chk.config(command=on_chk_change)

            # Initialer Zustand des extra_frame
            if var_standard.get():
                extra_frame.pack(fill="x", padx=25, pady=2)
            else:
                extra_frame.forget()

    refresh_dichtungen_list()

    action_frame = tk.Frame(win, bg=bgc)
    action_frame.pack(pady=5)

    def save_and_close():
        new_config = []
        for d in config:
            name = d["var_name"].get().strip()
            if not name:
                continue
            always_show = d["var_standard"].get()
            try:
                default_value = float(d["var_value"].get()) if d["var_value"].get().strip() != "" else 0
            except ValueError:
                default_value = 0
            try:
                order_val = int(d["var_order"].get()) if d["var_order"].get().strip() != "" else ""
            except ValueError:
                order_val = ""
            new_config.append({
                "name": name,
                "always_show": always_show,
                "default_value": default_value,
                "order": order_val
            })
        save_dichtungen(new_config)
        if not tutorial:
            win.destroy()

    tk.Button(action_frame, text="Speichern", command=save_and_close,
              bg=light_theme["button_bg"] if current_theme == "light" else dark_theme["button_bg"],
              fg=light_theme["button_fg"] if current_theme == "light" else dark_theme["button_fg"],
              width=10).pack(side="left", padx=5)
    tk.Button(action_frame, text="Abbrechen", command=lambda: win.destroy(),
              bg=light_theme["button_bg"] if current_theme == "light" else dark_theme["button_bg"],
              fg=light_theme["button_fg"] if current_theme == "light" else dark_theme["button_fg"],
              width=10).pack(side="left", padx=5)

    if not tutorial:
        win.grab_set()
        parent.wait_window(win)
    return win


# ----------------------------------------------------------------
# Tutorial-Funktion (optional)
# ----------------------------------------------------------------

def start_tutorial(root, in_entry, out_entry, btn_manage_dichtungen, btn_convert, btn_print, btn_settings, btn_exit):
    steps = [
        {"message": "Schritt 1: Willkommen im Packlisten Converter Tutorial!", "widget": None},
        {"message": "Schritt 2: Eingabedatei auswählen...", "widget": in_entry},
        {"message": "Schritt 3: Ausgabedatei festlegen...", "widget": out_entry},
        {"message": "Schritt 4: Dichtungen verwalten...", "widget": btn_manage_dichtungen},
        {"message": "Öffne 'Dichtungen verwalten'...", "widget": None, "action": "open_dichtungen"},
        {"message": "Schließe das Fenster 'Dichtungen verwalten'...", "widget": None, "action": "close_dichtungen"},
        {"message": "Schritt 5: Datei konvertieren...", "widget": btn_convert},
        {"message": "Schritt 6: Drucken...", "widget": btn_print},
        {"message": "Schritt 7: Einstellungen öffnen...", "widget": btn_settings},
        {"message": "Schließe das Einstellungsfenster...", "widget": None, "action": "close_settings"},
        {"message": "Schritt 8: Anwendung beenden...", "widget": btn_exit}
    ]

    opened_dichtungen = None
    opened_settings = None

    def show_step(index):
        if index >= len(steps):
            return
        step = steps[index]
        if "action" in step:
            if step["action"] == "open_dichtungen":
                nonlocal opened_dichtungen
                opened_dichtungen = manage_dichtungen_window(root, tutorial=True)
            elif step["action"] == "close_dichtungen":
                if opened_dichtungen is not None:
                    opened_dichtungen.destroy()
                    opened_dichtungen = None
            elif step["action"] == "open_settings":
                nonlocal opened_settings
                opened_settings = open_settings(root, tutorial=True)
            elif step["action"] == "close_settings":
                if opened_settings is not None:
                    opened_settings.destroy()
                    opened_settings = None

        popup = tk.Toplevel(root)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        frame = ttk.Frame(popup, relief="solid", borderwidth=1)
        frame.pack(padx=5, pady=5)
        label = ttk.Label(frame, text=step["message"], wraplength=300)
        label.pack(padx=10, pady=10)
        btn_next = ttk.Button(frame, text="Weiter", command=lambda: [popup.destroy(), show_step(index+1)])
        btn_next.pack(pady=5)

        widget = step.get("widget", None)
        if widget is not None:
            widget.update_idletasks()
            x = widget.winfo_rootx() + widget.winfo_width() + 10
            y = widget.winfo_rooty()
        else:
            root.update_idletasks()
            x = root.winfo_rootx() + (root.winfo_width() // 2) - 150
            y = root.winfo_rooty() + (root.winfo_height() // 2) - 50

        popup.geometry(f"+{x}+{y}")

    show_step(0)

# ----------------------------------------------------------------
# main(): GUI-Start
# ----------------------------------------------------------------

def main():
    load_settings()
    root = tk.Tk()
    check_for_updates()
    root.title("Packliste Converter")
    apply_ttk_style(current_theme)

    icon_path = resource_path(ICON_FILE)
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except:
            pass

    mainframe = ttk.Frame(root)
    mainframe.pack(fill="both", expand=True)

    top_right_frame = ttk.Frame(mainframe)
    top_right_frame.pack(side="top", anchor="ne", padx=5, pady=5)

    try:
        gear_img = tk.PhotoImage(file=resource_path("gear.png"))
    except:
        gear_img = None

    try:
        printer_img = tk.PhotoImage(file=resource_path("printer.png"))
    except:
        printer_img = None

    def open_settings_wrapper():
        open_settings(root)
        if AUTO_FILENAME:
            frm_out.grid_remove()
        else:
            frm_out.grid()

    btn_settings = ttk.Button(top_right_frame,
                              image=gear_img if gear_img else None,
                              text="" if gear_img else "Einstellungen12",
                              command=open_settings_wrapper)
    btn_settings.pack(side="left", padx=2)

    def print_file():
        global last_output_file
        if last_output_file and os.path.isfile(last_output_file):
            open_excel_print_preview(last_output_file)
        else:
            in_path = in_var.get()
            if in_path and os.path.isfile(in_path):
                temp_print_file = tempfile.mktemp(suffix=".xlsx")
                user_dichtungen = load_dichtungen()
                try:
                    convert_file(in_path, temp_print_file, user_dichtungen, show_message=False)
                    last_output_file = temp_print_file
                    open_excel_print_preview(temp_print_file)
                except Exception as e:
                    messagebox.showerror("Fehler", f"Fehler beim Konvertieren:\n{e}")
            else:
                messagebox.showerror("Fehler", "Keine gültige Eingabedatei gefunden!")

    btn_print = ttk.Button(top_right_frame,
                           image=printer_img if printer_img else None,
                           text="" if printer_img else "Drucken",
                           command=print_file)
    btn_print.pack(side="left", padx=2)

    frm_files = ttk.Frame(mainframe)
    frm_files.pack(padx=10, pady=10)

    ttk.Label(frm_files, text="Eingabedatei (Excel/CSV):").grid(row=0, column=0, sticky="w")
    in_var = tk.StringVar()
    in_entry = ttk.Entry(frm_files, textvariable=in_var, width=50)
    in_entry.grid(row=0, column=1, padx=5)

    def reset_last_output(*args):
        global last_output_file
        last_output_file = None

    in_var.trace("w", reset_last_output)

    def browse_input():
        path = filedialog.askopenfilename(title="Eingabedatei wählen",
                                          filetypes=[("Excel/CSV Files", "*.xlsx;*.xls;*.csv"), ("All Files", "*.*")])
        if path:
            in_var.set(path)

    ttk.Button(frm_files, text="Suchen", command=browse_input).grid(row=0, column=2, padx=5)

    frm_out = ttk.Frame(frm_files)
    frm_out.grid(row=1, column=0, columnspan=3, pady=5)

    out_label = ttk.Label(frm_out, text="Ausgabedatei (Excel):")
    out_entry = ttk.Entry(frm_out, width=50)
    out_var = tk.StringVar()
    out_entry.config(textvariable=out_var)

    out_label.grid(row=0, column=0, sticky="w")
    out_entry.grid(row=0, column=1, padx=5)

    def save_as():
        filename = filedialog.asksaveasfilename(
            title="Ausgabedatei speichern unter",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if filename:
            out_var.set(filename)

    ttk.Button(frm_out, text="Speichern unter", command=save_as).grid(row=0, column=2, padx=5)

    if AUTO_FILENAME:
        frm_out.grid_remove()

    frm_actions = ttk.Frame(mainframe)
    frm_actions.pack(pady=10)

    def open_dichtungen():
        manage_dichtungen_window(root)

    btn_manage_dichtungen = ttk.Button(frm_actions, text="Dichtungen verwalten", command=open_dichtungen, width=20)
    btn_manage_dichtungen.grid(row=0, column=0, padx=5, pady=5)

    def do_convert():
        global last_output_file
        in_path = in_var.get()
        if not in_path or not os.path.isfile(in_path):
            messagebox.showerror("Fehler", "Bitte eine gültige Eingabedatei auswählen!")
            return
        user_dichtungen = load_dichtungen()
        if AUTO_FILENAME:
            try:
                if os.path.splitext(in_path)[1].lower() == ".csv":
                    df = pd.read_csv(in_path, sep=";", engine="python", header=0)
                else:
                    df = pd.read_excel(in_path, header=0)
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Einlesen der Datei:\n{e}")
                return
            out_path = generate_auto_filename(df)
        else:
            out_path = out_var.get()
            if not out_path:
                messagebox.showerror("Fehler", "Bitte eine Ausgabedatei angeben!")
                return
        try:
            convert_file(in_path, out_path, user_dichtungen)
            last_output_file = out_path
            play_converted_sound()
            os.startfile(out_path)
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Konvertieren:\n{e}")

    btn_convert = ttk.Button(frm_actions, text="Konvertieren", command=do_convert, width=20)
    btn_convert.grid(row=0, column=1, padx=5, pady=5)

    btn_exit = ttk.Button(frm_actions, text="Beenden", command=root.destroy, width=20)
    btn_exit.grid(row=0, column=2, padx=5, pady=5)

    btn_tutorial = ttk.Button(root, text="Tutorial", width=10,
                              command=lambda: start_tutorial(
                                  root, in_entry, out_entry,
                                  btn_manage_dichtungen, btn_convert,
                                  btn_print, btn_settings, btn_exit
                              ))
    btn_tutorial.place(x=10, y=10)

    root.mainloop()

if __name__ == "__main__":
    main()
