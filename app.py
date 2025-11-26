#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from pathlib import Path
import tempfile

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    redirect,
    url_for,
    flash,
    jsonify,
)
from werkzeug.utils import secure_filename

from packliste_core import convert_file, load_dichtungen, save_dichtungen

# -------------------------------------------------------
# Grund-Setup
# -------------------------------------------------------

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")

BASE_DIR = Path(__file__).resolve().parent


# -------------------------------------------------------
# Hilfsfunktion: Konvertierung anstoßen
# -------------------------------------------------------

def run_conversion(upload_path: Path, out_dir: Path, desired_stem: str) -> Path:
    """
    Ruft packliste_core.convert_file(...) auf und gibt den Pfad zur fertigen
    Excel-Datei zurück.
    """
    user_dichtungen = load_dichtungen()

    stem = (desired_stem or "").strip() or "Packliste"
    output_path = out_dir / f"{stem}.xlsx"

    # packliste_core kümmert sich um alles; show_message=False,
    # weil wir im Web keine Messagebox wollen.
    convert_file(
        str(upload_path),
        str(output_path),
        user_dichtungen,
        show_message=False,
    )

    if not output_path.exists():
        raise RuntimeError(
            f"Konvertierung hat keine neue Excel-Datei erzeugt "
            f"(output_path={output_path})"
        )

    return output_path


# -------------------------------------------------------
# Startseite + Upload / Download
# -------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        # Nur Formular anzeigen
        return render_template("index.html")

    # --- POST: Datei wurde hochgeladen ---
    if "input_file" not in request.files or request.files["input_file"].filename == "":
        flash("Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).", "error")
        return redirect(url_for("index"))

    file = request.files["input_file"]

    if not allowed(file.filename):
        flash("Ungültiger Dateityp. Erlaubt sind .xlsx, .xls, .csv.", "error")
        return redirect(url_for("index"))

    desired_name = (request.form.get("desired_name") or "").strip()
    if not desired_name:
        desired_name = "Packliste"

    # Temporäres Verzeichnis für Upload & Konvertierung
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Upload speichern
        input_path = tmpdir_path / secure_filename(file.filename)
        file.save(input_path)

        # Konvertieren
        try:
            converted_path = run_conversion(input_path, tmpdir_path, desired_name)
        except Exception as e:
            app.logger.exception("Fehler bei der Konvertierung")
            return render_template(
                "error.html",
                message="Fehler beim Konvertieren",
                details=str(e),
            ), 500

        # Erfolgreich → Datei zurückgeben
        return send_file(
            converted_path,
            as_attachment=True,
            download_name=f"{desired_name}.xlsx",
        )


# -------------------------------------------------------
# Dichtungsverwaltung – Seite
# -------------------------------------------------------

@app.route("/dichtungen", methods=["GET"])
def dichtungen_page():
    """
    Zeigt die hübsche Dichtungsverwaltungs-Seite.
    Die Daten selbst holt sich das Frontend per /api/dichtungen.
    """
    return render_template("dichtungen.html")


# -------------------------------------------------------
# Dichtungsverwaltung – JSON-API (für das Frontend)
# -------------------------------------------------------

@app.route("/api/dichtungen", methods=["GET", "POST"])
def dichtungen_api():
    # ---- GET: aktuelle Konfiguration laden ----
    if request.method == "GET":
        items = load_dichtungen()
        return jsonify({"success": True, "items": items})

    # ---- POST: neue Konfiguration speichern ----
    data = request.get_json(silent=True, force=True) or {}
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

        # Standardwert (Zahl)
        default_value = item.get("default_value", 0)
        try:
            default_value = float(default_value)
        except Exception:
            default_value = 0.0

        # Reihenfolge (optional ganzzahlig)
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

    # In die dichtungen.json schreiben (packliste_core kümmert sich um den Pfad)
    save_dichtungen(cleaned)

    # Direkt das zurückgeben, was wirklich gespeichert ist
    return jsonify({"success": True, "items": load_dichtungen()})


# -------------------------------------------------------
# Debug-Start (lokal)
# -------------------------------------------------------

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8000)
