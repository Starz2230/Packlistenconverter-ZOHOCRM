#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    flash,
    redirect,
    url_for,
    jsonify,          # NEU
)
from werkzeug.utils import secure_filename
import tempfile

from packliste_core import (
    convert_file,
    load_dichtungen,  # NEU
    save_dichtungen,  # NEU
)

# -------------------------------------------------------
# Flask-App
# -------------------------------------------------------
app = Flask(__name__)

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# -------------------------------------------------------
# Auto-Dateinamen wie im EXE-Tool
# Service Techniker + Zeitraum -> Dateiname
# -------------------------------------------------------
def suggest_auto_stem(input_path: str) -> str | None:
    """
    Liest die Eingabedatei und erzeugt einen Dateinamen-Stamm wie
    'DanielOberrauner_24-11-2025-28-11-2025'.
    Gibt None zurück, wenn etwas schiefgeht.
    """
    try:
        ext = os.path.splitext(input_path)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(input_path, sep=";", engine="python", header=0)
        else:
            df = pd.read_excel(input_path, header=0)
    except Exception:
        return None

    def safe_val(df_, col, index):
        if col not in df_.columns or index < 0 or index >= len(df_):
            return ""
        val = df_[col].iloc[index]
        if pd.isna(val):
            return ""
        return str(val)

    def parse_date_part(value):
        import re as _re

        if not isinstance(value, str):
            value = str(value)
        m = _re.match(r"^(\d{1,2}\.\d{1,2}\.\d{4})", value.strip())
        if not m:
            return None
        ds = m.group(1)
        try:
            return pd.to_datetime(ds, dayfirst=True, errors="coerce")
        except Exception:
            return None

    def get_zeitraum_von_bis(df_, col="Zeitraum"):
        if col not in df_.columns:
            return ""
        dtlist = []
        for val in df_[col].dropna():
            dt = parse_date_part(val)
            if dt is not None:
                dtlist.append(dt)
        if not dtlist:
            return ""
        von_dt = min(dtlist)
        bis_dt = max(dtlist)
        # mit Bindestrich statt " - ", damit es im Dateinamen sauber ist
        return f"{von_dt.strftime('%d.%m.%Y')}-{bis_dt.strftime('%d.%m.%Y')}"

    serv = safe_val(df, "Service Techniker", 3)
    date_range = get_zeitraum_von_bis(df, "Zeitraum")

    if not serv and not date_range:
        return None

    def sanitize(text: str) -> str:
        # nur Buchstaben, Zahlen, Unterstrich und Minus
        return "".join(c for c in text if c.isalnum() or c in ("_", "-"))

    serv_sanitized = sanitize(serv) or "Packliste"
    date_sanitized = sanitize(date_range.replace(" ", "")) if date_range else ""

    if date_sanitized:
        stem = f"{serv_sanitized}_{date_sanitized}"
    else:
        stem = serv_sanitized

    return stem or None


# -------------------------------------------------------
# Startseite: Upload + Konvertierung
# -------------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        error = None

        upload = request.files.get("input_file")
        desired_stem = request.form.get("desired_name", "").strip()

        if not upload or upload.filename == "":
            error = "Bitte eine Packlisten-Datei auswählen (.xlsx / .xls / .csv)."
            return render_template("index.html", error=error)

        if not allowed(upload.filename):
            error = "Ungültiges Dateiformat. Erlaubt sind: .xlsx, .xls, .csv"
            return render_template("index.html", error=error)

        # Temporäres Arbeitsverzeichnis
        tmpdir = Path(tempfile.mkdtemp(prefix="packliste_"))
        input_path = tmpdir / secure_filename(upload.filename)
        upload.save(input_path)

        # Dateinamen-Stamm:
        # 1. Wenn der User etwas eingibt -> das verwenden
        # 2. Sonst automatisch aus Service Techniker + Zeitraum
        # 3. Fallback: Packliste_YYYYMMDD
        if desired_stem:
            stem = desired_stem
        else:
            stem = suggest_auto_stem(str(input_path))
            if not stem:
                stem = f"Packliste_{date.today():%Y%m%d}"

        output_path = tmpdir / f"{stem}.xlsx"

        try:
            user_dichtungen = load_dichtungen()
            # packliste_core kümmert sich um alles – wir wollen keine Messageboxen
            convert_file(str(input_path), str(output_path), user_dichtungen, show_message=False)

            if not output_path.exists():
                raise RuntimeError("Konvertierung hat keine neue Excel-Datei erzeugt.")
        except Exception as e:
            print("Fehler bei der Konvertierung:", e)
            error = f"Unerwarteter Fehler bei der Konvertierung: {e}"
            return render_template("index.html", error=error)

        # Erfolgreich -> Datei direkt zum Download schicken
        return send_file(
            output_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{stem}.xlsx",
        )

    # GET-Aufruf
    default_stem = f"Packliste_{date.today():%Y%m%d}"
    return render_template("index.html", error=None, default_stem=default_stem)


# -------------------------------------------------------
# Dichtungen-Verwaltung (wird vom /dichtungen-Frontend genutzt)
# -------------------------------------------------------
@app.route("/dichtungen", methods=["GET", "POST"])
def manage_dichtungen():
    if request.method == "POST":
        data = request.get_json() or {}
        dichtungen = data.get("dichtungen", [])
        save_dichtungen(dichtungen)
        return jsonify({"ok": True})
    else:
        user_dichtungen = load_dichtungen()
        return render_template("dichtungen.html", dichtungen=user_dichtungen)

# -------------------------------------------------------
# Dichtungs-Einstellungen (Seite)
# -------------------------------------------------------

@app.route("/dichtungen", methods=["GET"])
def dichtungen_page():
    """
    Zeigt die Dichtungsverwaltungs-Seite.
    Das Template heißt 'dichtungen.html'.
    """
    return render_template("dichtungen.html")


# -------------------------------------------------------
# Dichtungs-API (für das Frontend-JS)
# -------------------------------------------------------

@app.route("/api/dichtungen", methods=["GET", "POST"])
def api_dichtungen():
    # --- aktuelle Konfiguration laden ---
    if request.method == "GET":
        items = load_dichtungen()
        return jsonify({"success": True, "items": items})

    # --- neue Konfiguration speichern ---
    data = request.get_json(silent=True) or {}
    raw_items = data.get("items")

    if not isinstance(raw_items, list):
        return jsonify({"success": False, "error": "invalid_payload"}), 400

    cleaned = []
    for item in raw_items:
        if not isinstance(item, dict):
            continue

        name = str(item.get("name", "")).strip()
        if not name:
            continue

        # Standard-Haken
        always_show = bool(item.get("always_show"))

        # Standardwert (Zahl, sonst 0)
        default_value = item.get("default_value", 0)
        try:
            default_value = float(default_value)
        except Exception:
            default_value = 0.0

        # Reihenfolge (optional int oder "")
        order_raw = str(item.get("order", "")).strip()
        if order_raw:
            try:
                order = int(order_raw)
            except Exception:
                order = ""
        else:
            order = ""

        cleaned.append(
            {
                "name": name,
                "always_show": always_show,
                "default_value": default_value,
                "order": order,
            }
        )

    # In dichtungen.json schreiben
    save_dichtungen(cleaned)

    # Zurückgeben, was jetzt gespeichert ist
    return jsonify({"success": True, "items": load_dichtungen()})


if __name__ == "__main__":
    # Für Render ist das egal, lokal aber praktisch.
    app.run(debug=True, host="0.0.0.0", port=8000)
