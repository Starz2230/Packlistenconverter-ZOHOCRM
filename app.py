import os
import json
import tempfile

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    flash,
    redirect,
    url_for,
)
from werkzeug.utils import secure_filename

# WICHTIG: Headless, damit kein tkinter/GUI auf dem Server losläuft
os.environ.setdefault("HEADLESS", "1")

# Jetzt dein Original-Modul importieren
from packliste_core import convert_file, load_dichtungen

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


@app.context_processor
def inject_defaults():
    """
    Stellt Standard-Dichtungen als JSON für das Template bereit.
    (Das sind dieselben, die deine EXE über dichtungen.json nutzt.)
    """
    try:
        default_dichtungen = load_dichtungen()
    except Exception:
        default_dichtungen = []
    return {
        "default_dichtungen_json": json.dumps(
            default_dichtungen, ensure_ascii=False, indent=2
        )
    }


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # 1) Datei vorhanden?
        if (
            "input_file" not in request.files
            or request.files["input_file"].filename.strip() == ""
        ):
            flash(
                "Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).", "error"
            )
            return redirect(url_for("index"))

        f = request.files["input_file"]

        # 2) Dateiendung prüfen
        if not allowed(f.filename):
            flash(
                "Ungültiges Dateiformat. Erlaubt sind .xlsx, .xls, .csv.",
                "error",
            )
            return redirect(url_for("index"))

        # 3) Dichtungen bestimmen
        #    a) Wenn im Formular eigenes JSON angegeben wurde → das verwenden
        #    b) sonst: die dichtungen.json vom Server laden (wie in der EXE)
        raw = (request.form.get("user_dichtungen") or "").strip()
        try:
            if raw:
                user_dichtungen = json.loads(raw)
            else:
                user_dichtungen = load_dichtungen()
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

        # 5) Konvertierung (EXAKT dieselbe Logik wie in deiner EXE)
        try:
            convert_file(tmp_in.name, tmp_out.name, user_dichtungen, show_message=False)
        except Exception as e:
            # Wenn in packliste_core irgendwas schiefgeht → klare Fehlermeldung im Browser
            flash(f"Fehler bei der Konvertierung: {e}", "error")
            try:
                os.remove(tmp_in.name)
            except Exception:
                pass
            try:
                os.remove(tmp_out.name)
            except Exception:
                pass
            return redirect(url_for("index"))

        # Input-Temp löschen
        try:
            os.remove(tmp_in.name)
        except Exception:
            pass

        # 6) konvertierte Datei zum Download schicken
        download_name = secure_filename(
            os.path.splitext(f.filename)[0] + "_konvertiert.xlsx"
        )
        return send_file(tmp_out.name, as_attachment=True, download_name=download_name)

    # GET → Formular anzeigen
    return render_template("index.html")


@app.route("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # debug nur lokal
    app.run(host="0.0.0.0", port=port, debug=True)
