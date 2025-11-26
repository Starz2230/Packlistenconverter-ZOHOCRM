#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import tempfile
from pathlib import Path

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    redirect,
    url_for,
    flash,
)
from werkzeug.utils import secure_filename

from packliste_core import convert_file, load_dichtungen, save_dichtungen

# ------------------------------------------------------------
# Flask-Setup
# ------------------------------------------------------------

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_stem_from_filename(filename: str) -> str:
    """Dateinamen-Stem „sauber“ machen (nur Buchstaben/Zahlen/_/-)."""
    stem = Path(filename).stem
    stem = re.sub(r"[^A-Za-z0-9_-]+", "_", stem)
    return stem or "Packliste"


def run_conversion(input_path: Path, output_dir: Path, desired_stem: str | None) -> Path:
    """
    Führt die eigentliche Konvertierung aus:
    - lädt die Dichtungen aus dichtungen.json
    - ruft convert_file(...) aus packliste_core auf
    - gibt Pfad zur erzeugten .xlsx zurück
    """
    user_dichtungen = load_dichtungen()

    if desired_stem:
        desired_stem = desired_stem.strip()
    if not desired_stem:
        safe_stem = safe_stem_from_filename(input_path.name)
    else:
        safe_stem = re.sub(r"[^A-Za-z0-9_-]+", "_", desired_stem) or "Packliste"

    output_path = output_dir / f"{safe_stem}.xlsx"

    convert_file(str(input_path), str(output_path), user_dichtungen, show_message=False)

    if not output_path.exists():
        raise RuntimeError(
            f"Konvertierung hat keine neue Excel-Datei erzeugt (output_dir={output_dir})"
        )

    return output_path


# ------------------------------------------------------------
# Startseite: Upload + Download
# ------------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    last_desired_name = ""
    last_filename_hint = ""

    if request.method == "POST":
        try:
            # 1) Datei vorhanden?
            if "input_file" not in request.files:
                error = "Es wurde keine Datei hochgeladen."
                raise ValueError(error)

            file = request.files["input_file"]
            if file.filename == "":
                error = "Bitte eine Packlisten-Datei auswählen."
                raise ValueError(error)

            if not allowed_file(file.filename):
                error = "Nur .xlsx, .xls oder .csv sind erlaubt."
                raise ValueError(error)

            desired_name = request.form.get("desired_name", "").strip()
            last_desired_name = desired_name
            last_filename_hint = file.filename

            # 2) Temporäre Dateien/Ordner
            tmpdir = Path(tempfile.mkdtemp(prefix="packliste_input_"))
            input_path = tmpdir / secure_filename(file.filename)
            file.save(input_path)

            # 3) Konvertierung
            outdir = Path(tempfile.mkdtemp(prefix="packliste_output_"))
            converted_path = run_conversion(input_path, outdir, desired_name)

            download_name = f"{converted_path.stem}.xlsx"

            # 4) Fertige Datei zurückgeben
            return send_file(
                converted_path,
                as_attachment=True,
                download_name=download_name,
            )

        except Exception as e:
            # Fehlertext fürs Template (und für Debug in Render-Logs)
            error = f"Fehler beim Konvertieren: {e}"
            print(error)

    # GET oder Fehlerfall -> normales Template
    return render_template(
        "index.html",
        error=error,
        last_desired_name=last_desired_name,
        last_filename_hint=last_filename_hint,
    )


# ------------------------------------------------------------
# Dichtungen verwalten (Web-Frontend für dichtungen.json)
# ------------------------------------------------------------

@app.route("/dichtungen", methods=["GET", "POST"])
def manage_dichtungen():
    if request.method == "POST":
        row_ids = request.form.getlist("row_ids")
        new_config = []

        for row_id in row_ids:
            row_id = row_id.strip()
            if not row_id:
                continue

            name = request.form.get(f"name_{row_id}", "").strip()
            if not name:
                # Leere Namen ignorieren
                continue

            delete_flag = request.form.get(f"delete_{row_id}") == "on"
            if delete_flag:
                # Zeile ist als "löschen" markiert -> überspringen
                continue

            always_show = request.form.get(f"always_show_{row_id}") == "on"

            default_raw = request.form.get(f"default_value_{row_id}", "").strip()
            try:
                default_value = float(default_raw) if default_raw != "" else 0.0
            except ValueError:
                default_value = 0.0

            order_raw = request.form.get(f"order_{row_id}", "").strip()
            try:
                order_val = int(order_raw) if order_raw != "" else ""
            except ValueError:
                order_val = ""

            new_config.append(
                {
                    "name": name,
                    "always_show": always_show,
                    "default_value": default_value,
                    "order": order_val,
                }
            )

        # Neue Konfiguration speichern
        save_dichtungen(new_config)
        flash("Dichtungen wurden gespeichert.", "success")
        return redirect(url_for("manage_dichtungen"))

    # GET: aktuelle Dichtungen laden und anzeigen
    config = load_dichtungen() or []

    # Standardwerte sicherstellen + Row-IDs vergeben
    for idx, d in enumerate(config):
        d.setdefault("always_show", False)
        d.setdefault("default_value", 0)
        d.setdefault("order", "")
        d["row_id"] = idx

    return render_template("dichtungen.html", dichtungen=config)


# ------------------------------------------------------------
# Main-Entry (für lokalen Start)
# ------------------------------------------------------------

if __name__ == "__main__":
    # Lokaler Test: http://127.0.0.1:5000
    app.run(debug=True, host="0.0.0.0", port=5000)
