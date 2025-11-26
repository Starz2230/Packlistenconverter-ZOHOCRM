#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import tempfile
from pathlib import Path

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    flash,
    redirect,
    url_for,
    jsonify,
)
from werkzeug.utils import secure_filename

# kommt aus deiner packliste_core.py
from packliste_core import convert_file, load_dichtungen, save_dichtungen

# ---------------------------------------------------
# Grund-Konfiguration
# ---------------------------------------------------

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename: str) -> bool:
    """Prüft, ob die Dateiendung erlaubt ist."""
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


# ---------------------------------------------------
# Wrapper um convert_file (Konvertierung)
# ---------------------------------------------------

def run_conversion(input_path: Path, output_dir: Path, desired_stem: str | None) -> Path:
    """
    Führt die eigentliche Packlisten-Konvertierung aus.

    input_path  = Pfad zur hochgeladenen Datei
    output_dir  = temporäres Verzeichnis
    desired_stem = gewünschter Dateiname (ohne .xlsx)
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    if desired_stem:
        out_name = f"{desired_stem}.xlsx"
    else:
        out_name = "Packliste_konvertiert.xlsx"

    out_path = output_dir / out_name

    # Dichtungen aus dichtungen.json laden
    user_dichtungen = load_dichtungen()

    # packliste_core.convert_file schreibt an out_path
    convert_file(
        str(input_path),
        str(out_path),
        user_dichtungen,
        show_message=False,  # keine GUI-Meldung im Web-Modus
    )

    if not out_path.exists():
        # Wenn hier was schiefgeht, sehen wir den Fehler im Log
        raise RuntimeError(
            f"Konvertierung hat keine neue Excel-Datei erzeugt (output_dir={output_dir})"
        )

    return out_path


# ---------------------------------------------------
# Startseite: Upload & Konvertierung
# ---------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # 1) Datei vorhanden?
        if (
            "input_file" not in request.files
            or request.files["input_file"].filename == ""
        ):
            flash(
                "Bitte eine Packlisten-Datei hochladen (.xlsx / .xls / .csv).",
                "error",
            )
            return redirect(url_for("index"))

        file = request.files["input_file"]
        filename = secure_filename(file.filename)

        # 2) Dateiendung prüfen
        if not allowed_file(filename):
            flash(
                "Nur .xlsx, .xls oder .csv Dateien sind erlaubt.",
                "error",
            )
            return redirect(url_for("index"))

        # 3) Optionaler Wunschnamen (ohne .xlsx)
        desired_name = (request.form.get("output_name") or "").strip()
        if desired_name.lower().endswith(".xlsx"):
            desired_name = desired_name[:-5]

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)

                # Eingabedatei speichern
                input_path = tmpdir_path / filename
                file.save(input_path)

                # Konvertierung ausführen
                converted_path = run_conversion(
                    input_path=input_path,
                    output_dir=tmpdir_path,
                    desired_stem=desired_name or None,
                )

                # Download an Browser senden
                return send_file(
                    converted_path,
                    as_attachment=True,
                    download_name=converted_path.name,
                    mimetype=(
                        "application/"
                        "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )

        except Exception as e:
            # Fehler ins Render-Log schreiben
            app.logger.exception("Fehler bei der Konvertierung")
            flash(f"Unerwarteter Fehler im Server: {e}", "error")
            return render_template("index.html")

    # GET → nur Formular anzeigen
    return render_template("index.html")


# ---------------------------------------------------
# Dichtungs-Verwaltung: Seite
# ---------------------------------------------------

@app.route("/dichtungen", methods=["GET"])
def dichtungen_page():
    """
    Seite mit dem UI zur Dichtungs-Verwaltung.
    Template: templates/dichtungen.html
    """
    return render_template("dichtungen.html")


# ---------------------------------------------------
# Dichtungs-API: GET = laden, POST = speichern
# ---------------------------------------------------

@app.route("/api/dichtungen", methods=["GET", "POST"])
def api_dichtungen():
    # ----------------- GET -----------------
    if request.method == "GET":
        items = load_dichtungen()
        return jsonify({"success": True, "items": items})

    # ----------------- POST ----------------
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

        always_show = bool(item.get("always_show"))

        # Standardwert → Zahl
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

    # In dichtungen.json speichern
    save_dichtungen(cleaned)

    # und direkt das, was jetzt gilt, zurückgeben
    return jsonify({"success": True, "items": load_dichtungen()})


# ---------------------------------------------------
# Lokaler Start (Render nutzt gunicorn)
# ---------------------------------------------------

if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
    )
