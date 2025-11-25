# app.py
import os
import tempfile
import inspect

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

# >>> WICHTIG: packliste_core.py bleibt dein großer Original-Code! <<<
from packliste_core import convert_file


ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed(filename: str) -> bool:
    """Prüft, ob die Dateiendung erlaubt ist."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def run_conversion_with_unknown_signature(input_path: str, workdir: str) -> str:
    """
    Ruft convert_file() aus packliste_core auf, ohne die genaue Signatur kennen zu müssen.

    Strategie:
    - vorher: Liste der Dateien im workdir merken
    - convert_file() mit 1 oder 2 Parametern ausprobieren
    - wenn convert_file() einen Dateipfad zurückgibt und der existiert → nutzen
    - sonst: im workdir nach neuen .xlsx/.xls-Dateien suchen und die nehmen
    """
    before_files = set(os.listdir(workdir))

    # Sicherstellen, dass relativ erzeugte Dateien im workdir landen
    old_cwd = os.getcwd()
    os.chdir(workdir)
    try:
        sig = inspect.signature(convert_file)
        param_count = len(sig.parameters)

        result = None
        if param_count == 1:
            # z.B. def convert_file(input_path)
            result = convert_file(input_path)
        elif param_count == 2:
            # z.B. def convert_file(input_path, output_dir_or_path)
            # Wir geben hier das workdir mit – der Core kann selbst entscheiden,
            # ob er dort einen Dateinamen erzeugt oder direkt einen Pfad verwendet.
            result = convert_file(input_path, workdir)
        else:
            # Fallback: einfach wie im 1-Argument-Fall versuchen
            result = convert_file(input_path)
    finally:
        os.chdir(old_cwd)

    # 1) Falls convert_file einen Pfad zurückgibt und der existiert → nutzen
    if isinstance(result, str):
        candidate = result
        if not os.path.isabs(candidate):
            candidate = os.path.join(workdir, candidate)
        if os.path.exists(candidate):
            return candidate

    # 2) Andernfalls: neue Excel-Datei(en) im workdir suchen
    after_files = set(os.listdir(workdir))
    new_files = [
        f for f in (after_files - before_files)
        if f.lower().endswith((".xlsx", ".xls"))
    ]

    if not new_files:
        raise RuntimeError(
            "Der Converter hat keine Ausgabedatei erzeugt "
            "(keine neue .xlsx/.xls Datei gefunden)."
        )

    # wenn mehrere neu sind, nehmen wir einfach die zuletzt sortierte
    new_files.sort()
    out_path = os.path.join(workdir, new_files[-1])
    return out_path


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("input_file")

        # 1) Datei da?
        if not file or file.filename == "":
            flash(
                "Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).",
                "error",
            )
            return redirect(url_for("index"))

        # 2) Endung ok?
        if not allowed(file.filename):
            flash(
                "Ungültiger Dateityp. Erlaubt sind .xlsx, .xls oder .csv.",
                "error",
            )
            return redirect(url_for("index"))

        # 3) Temporäres Arbeitsverzeichnis
        tmpdir = tempfile.mkdtemp(prefix="packliste_")

        # 4) Upload speichern
        safe_name = secure_filename(file.filename)
        input_path = os.path.join(tmpdir, safe_name)
        file.save(input_path)

        # 5) Konvertierung
        try:
            output_path = run_conversion_with_unknown_signature(input_path, tmpdir)
        except Exception as e:
            # Für Debugging auf Render im Log sichtbar machen
            print("Fehler bei convert_file:", repr(e), flush=True)
            flash(
                "Beim Konvertieren ist ein Fehler aufgetreten. "
                "Details siehe Server-Log.",
                "error",
            )
            return redirect(url_for("index"))

        if not os.path.exists(output_path):
            flash(
                "Konvertierung fehlgeschlagen: Ausgabedatei wurde nicht gefunden.",
                "error",
            )
            return redirect(url_for("index"))

        # 6) Datei zurückgeben (Download wie in der EXE, nur eben per Browser)
        download_name = os.path.basename(output_path)
        return send_file(
            output_path,
            as_attachment=True,
            download_name=download_name,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    # GET: Formular anzeigen
    return render_template("index.html")
    

if __name__ == "__main__":
    # für lokalen Test: python app.py
    app.run(host="0.0.0.0", port=5000, debug=True)
