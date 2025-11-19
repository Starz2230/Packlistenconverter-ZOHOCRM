import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import tempfile

# Wichtig: Standardmäßig HEADLESS=1 setzen, damit auf dem Server
# kein GUI-/tkinter-Code ausgeführt wird.
os.environ.setdefault("HEADLESS", "1")

from packliste_core import convert_file  # erst nach HEADLESS setzen importieren

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # 1) Datei vorhanden?
        if "input_file" not in request.files or request.files["input_file"].filename == "":
            flash("Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).", "error")
            return redirect(url_for("index"))

        f = request.files["input_file"]

        # 2) Endung checken
        if not allowed(f.filename):
            flash("Ungültiges Dateiformat. Erlaubt sind .xlsx, .xls, .csv.", "error")
            return redirect(url_for("index"))

        # 3) optionale Dichtungen aus Textfeld (JSON)
        user_dichtungen = []
        try:
            raw = request.form.get("user_dichtungen")
            if raw:
                import json
                user_dichtungen = json.loads(raw)
        except Exception as e:
            flash(f"Dichtungen-JSON ungültig: {e}", "error")
            return redirect(url_for("index"))

        # 4) temporäre Dateien anlegen
        suffix = os.path.splitext(f.filename)[1]
        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        f.save(tmp_in.name)
        tmp_in.close()

        tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_out.close()

        # 5) Konvertierung
        try:
            convert_file(tmp_in.name, tmp_out.name, user_dichtungen, show_message=False)
        except Exception as e:
            flash(f"Fehler bei der Konvertierung: {e}", "error")
            # Aufräumen
            try:
                os.remove(tmp_in.name)
            except Exception:
                pass
            try:
                os.remove(tmp_out.name)
            except Exception:
                pass
            return redirect(url_for("index"))

        # Input-Temp löschen, Output zum Download anbieten
        try:
            os.remove(tmp_in.name)
        except Exception:
            pass

        download_name = secure_filename(os.path.splitext(f.filename)[0] + "_konvertiert.xlsx")
        return send_file(tmp_out.name, as_attachment=True, download_name=download_name)

    # GET → Formular anzeigen
    return render_template("index.html")


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # debug=True nur lokal, auf Render besser False (oder weglassen)
    app.run(host="0.0.0.0", port=port, debug=True)
