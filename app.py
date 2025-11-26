import os
import tempfile
import traceback
from pathlib import Path

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

# Hier verwenden wir deine ursprüngliche Konvertierungslogik
from packliste_core import convert_file

# Erlaubte Dateiendungen für Upload
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename: str) -> bool:
    """Prüft, ob die Datei eine der erlaubten Endungen hat."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


def run_conversion(input_path: Path, output_path: Path) -> Path:
    """
    Wrapper um deine ursprüngliche convert_file-Funktion aus packliste_core.

    Entspricht funktional der EXE-Variante:
    - input_path  -> hochgeladene Datei
    - output_path -> Ziel-Datei
    - user_dichtungen -> aktuell leeres Dict (kann später aus Datei/DB geladen werden)
    """
    user_dichtungen = {}  # hier könntest du später deine gespeicherten Dichtungen laden

    # packliste_core arbeitet mit String-Pfaden
    convert_file(str(input_path), str(output_path), user_dichtungen)

    # convert_file schreibt die Datei nach output_path
    return output_path


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html")

    # POST: Datei-Upload
    if "input_file" not in request.files:
        flash("Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).", "error")
        return redirect(url_for("index"))

    file = request.files["input_file"]

    if file.filename == "":
        flash("Bitte eine Packlisten-Datei auswählen.", "error")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Ungültiger Dateityp. Erlaubt sind .xlsx, .xls und .csv.", "error")
        return redirect(url_for("index"))

    # Optionaler Ausgabename aus dem Formular
    custom_name = request.form.get("custom_output_name", "").strip()

    # Temporärer Arbeitsordner (wird nach der Anfrage automatisch gelöscht)
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Eingabedatei sicher speichern
        safe_input_name = secure_filename(file.filename)
        input_path = tmpdir_path / safe_input_name
        file.save(input_path)

        # Ausgabedatei-Namen bestimmen
        if custom_name:
            safe_output_stem = secure_filename(custom_name)
        else:
            safe_output_stem = Path(safe_input_name).stem + "_konvertiert"

        output_name = safe_output_stem + ".xlsx"
        output_path = tmpdir_path / output_name

        # Konvertierung ausführen
        try:
            converted_path = run_conversion(input_path, output_path)
        except Exception as exc:
            app.logger.exception("Fehler beim Konvertieren")
            return (
                render_template(
                    "error.html",
                    error=str(exc),
                    traceback=traceback.format_exc(),
                ),
                500,
            )

        # Fertige Excel-Datei an den Browser schicken
        return send_file(
            converted_path,
            as_attachment=True,
            download_name=output_name,
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
