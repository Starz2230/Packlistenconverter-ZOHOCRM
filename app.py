import os
import tempfile
import html
from pathlib import Path
import importlib
import traceback

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
# Konfiguration
# -------------------------------------------------------

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename: str) -> bool:
    """Prüfe, ob die Dateiendung erlaubt ist."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # max. 32 MB Upload
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


# -------------------------------------------------------
# Hilfsfunktion: packliste_core dynamisch laden + ausführen
# -------------------------------------------------------

def run_conversion(input_path: Path, output_path: Path) -> Path:
    """
    Lädt packliste_core.py dynamisch und ruft dort convert_file(...) auf.

    Unterstützte Varianten in packliste_core:

      def convert_file(input_path, output_path):
          # schreibt nach output_path
          return output_path   (optional)

      oder

      def convert_file(input_path):
          # schreibt z.B. eine Datei und gibt den Pfad zurück
          return pfad_zur_ausgabedatei

      oder (zur Not)

      def convert_file(input_path):
          # überschreibt input_path in-place
          return None
    """
    # 1) packliste_core importieren
    try:
        core = importlib.import_module("packliste_core")
    except Exception as e:
        raise RuntimeError(
            "Fehler beim Import von 'packliste_core.py'. "
            "Stelle sicher, dass die Datei im gleichen Verzeichnis wie app.py liegt."
        ) from e

    convert_fn = getattr(core, "convert_file", None)
    if convert_fn is None or not callable(convert_fn):
        raise RuntimeError(
            "In 'packliste_core.py' wurde keine Funktion 'convert_file' gefunden.\n\n"
            "Bitte füge dort z.B. folgendes hinzu (als Wrapper um deine bestehende Logik):\n"
            "    def convert_file(input_path, output_path):\n"
            "        # deine bisherige Konvertierung hier aufrufen\n"
            "        ...\n"
        )

    # 2) Versuch: convert_file(input_path, output_path)
    try:
        result = convert_fn(str(input_path), str(output_path))

        # Wenn ein Pfad zurückgegeben wird → benutzen
        if isinstance(result, (str, Path)):
            result_path = Path(result)
            if result_path.is_file():
                return result_path

        # Wenn nichts zurückgegeben, aber output_path beschrieben wurde
        if output_path.is_file():
            return output_path

    except TypeError:
        # Signatur akzeptiert offenbar keine 2 Argumente → nächste Variante
        pass

    # 3) Versuch: convert_file(input_path)
    result = convert_fn(str(input_path))

    if isinstance(result, (str, Path)):
        result_path = Path(result)
        if result_path.is_file():
            return result_path

    # Fallback: vielleicht wurde input in-place überschrieben
    if input_path.is_file():
        return input_path

    raise RuntimeError(
        "convert_file aus 'packliste_core.py' hat keine Ausgabedatei erzeugt.\n"
        "Bitte prüfe die Implementierung."
    )


# -------------------------------------------------------
# Globale Fehlerbehandlung → IMMER Trace im Browser
# -------------------------------------------------------

def _format_exception(e: Exception) -> str:
    return "".join(traceback.format_exception(type(e), e, e.__traceback__))


@app.errorhandler(Exception)
def handle_unhandled_exception(e: Exception):
    tb = _format_exception(e)
    app.logger.error("Unerwarteter Fehler:\n%s", tb)

    return (
        "<h1>Unerwarteter Fehler im Server</h1>"
        "<p>Details siehe unten. Die gleiche Meldung findest du in den Render-Logs.</p>"
        f"<pre>{html.escape(tb)}</pre>",
        500,
        {"Content-Type": "text/html; charset=utf-8"},
    )


# -------------------------------------------------------
# Startseite + Upload / Download
# -------------------------------------------------------

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

            # --- Konvertierung ausführen (nutzt deinen packliste_core) ---
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

        except Exception as e:
            # Explicit: Fehlerseite mit Traceback
            tb = _format_exception(e)
            app.logger.error("Fehler bei der Konvertierung:\n%s", tb)

            return (
                "<h1>Fehler beim Konvertieren</h1>"
                "<p>Details siehe unten. Die gleiche Meldung findest du in den Render-Logs.</p>"
                f"<pre>{html.escape(tb)}</pre>",
                500,
                {"Content-Type": "text/html; charset=utf-8"},
            )

    # GET → Formular anzeigen
    # Wichtig: In templates/index.html muss das File-Feld name="input_file" haben
    return render_template("index.html")


# -------------------------------------------------------
# Lokaler Start (z.B. python app.py)
# -------------------------------------------------------

if __name__ == "__main__":
    app.run(
        debug=True,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
    )
