import os
import tempfile
import traceback
import html
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

# -------------------------------------------------------
# Versuch, die Konvertierfunktion aus deinem Core zu holen
# -------------------------------------------------------
try:
    from packliste_core import convert_file
except ImportError:
    convert_file = None

# Welche Dateitypen erlaubt sind
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
# max. 32 MB Upload (kannst du anpassen)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


# -------------------------------------------------------------------
# Wrapper, der versucht, deine convert_file-Funktion flexibel aufzurufen
# -------------------------------------------------------------------
def run_conversion(input_path: Path, output_path: Path) -> Path:
    """
    Ruft convert_file aus packliste_core auf.

    Unterstützte Varianten:
      - convert_file(input_path, output_path)
      - convert_file(input_path)  → liefert entweder Pfad zurück
                                   oder überschreibt input in-place
    """
    if convert_file is None:
        raise RuntimeError(
            "Konnte convert_file aus packliste_core.py nicht importieren. "
            "Bitte prüfe, ob dort eine Funktion 'convert_file' definiert ist."
        )

    # 1) Versuche: convert_file(input_path, output_path)
    try:
        result = convert_file(str(input_path), str(output_path))

        # Wenn die Funktion einen Pfad zurückgibt, nimm den
        if isinstance(result, (str, Path)) and Path(result).is_file():
            return Path(result)

        # Wenn sie selbst in output_path geschrieben hat
        if output_path.is_file():
            return output_path

    except TypeError:
        # Signatur passt nicht, dann nächste Variante
        pass

    # 2) Versuche: convert_file(input_path)
    result = convert_file(str(input_path))

    if isinstance(result, (str, Path)):
        result_path = Path(result)
        if result_path.is_file():
            return result_path

    # Wenn nichts zurückgegeben wurde, vielleicht in-place überschrieben
    if input_path.is_file():
        return input_path

    raise RuntimeError(
        "convert_file hat keine gültige Ausgabedatei erzeugt. "
        "Bitte prüfe die Implementierung in packliste_core.py."
    )


# ---------------------------------
# Startseite + Upload / Download
# ---------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # --- Datei vorhanden? ---
            if (
                "input_file" not in request.files
                or request.files["input_file"].filename == ""
            ):
                flash(
                    "Bitte eine Packlisten-Datei hochladen (.xlsx / .xls / .csv).",
                    "error",
                )
                return redirect(url_for("index"))

            upload_file = request.files["input_file"]
            filename = secure_filename(upload_file.filename)

            # --- Typ prüfen ---
            if not allowed_file(filename):
                flash(
                    "Ungültiger Dateityp. Erlaubt sind: .xlsx, .xls, .csv.",
                    "error",
                )
                return redirect(url_for("index"))

            # --- Temp-Ordner & Pfade ---
            tmpdir = Path(tempfile.mkdtemp(prefix="packliste_"))
            input_path = tmpdir / filename
            upload_file.save(input_path)

            # Ziel-Dateiname: gleicher Name + "_konvertiert.xlsx"
            base_name = input_path.stem
            output_filename = f"{base_name}_konvertiert.xlsx"
            output_path = tmpdir / output_filename

            # --- Konvertierung ausführen ---
            converted_path = run_conversion(input_path, output_path)

            if not converted_path.is_file():
                raise RuntimeError(
                    f"Konvertierung hat keine Ausgabedatei erzeugt: {converted_path}"
                )

            # --- Download an Browser schicken ---
            return send_file(
                converted_path,
                as_attachment=True,
                download_name=converted_path.name,
            )

        except Exception:
            # Vollständigen Traceback im Browser anzeigen,
            # damit du auf Render genau siehst, was schiefgeht.
            tb = traceback.format_exc()
            app.logger.error("Fehler bei der Konvertierung:\n%s", tb)

            return (
                "<h1>Fehler beim Konvertieren</h1>"
                "<p>Details siehe unten. "
                "Die gleiche Fehlermeldung findest du auch in den Render-Logs.</p>"
                f"<pre>{html.escape(tb)}</pre>",
                500,
                {"Content-Type": "text/html; charset=utf-8"},
            )

    # GET → Formular anzeigen
    return render_template("index.html")


# Lokales Testen
if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
    )
