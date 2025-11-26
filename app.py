import os
import tempfile
import traceback
from pathlib import Path
from typing import Optional

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

# Deine ursprüngliche Konvertierungslogik
from packliste_core import convert_file

# Erlaubte Dateiendungen
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename: str) -> bool:
    """Prüft, ob die Datei eine der erlaubten Endungen hat."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


def run_conversion(
    input_path: Path,
    output_dir: Path,
    desired_stem: Optional[str],
) -> Path:
    """
    Führt die Konvertierung aus.

    Wichtig: Hier wird davon ausgegangen, dass convert_file so arbeitet wie in
    deiner EXE:
      convert_file(eingabedatei, AUSGABE_ORDNER, user_dichtungen)

    Danach suchen wir im Ausgabeordner nach neu entstandenen .xlsx-Dateien.
    """
    user_dichtungen = {}  # später kannst du hier echte Dichtungen laden

    # Vorher merken, welche .xlsx-Dateien im Ordner existieren
    before = set(output_dir.rglob("*.xlsx"))

    # Konvertierung ausführen – 2. Parameter = AUSGABE-ORDNER
    convert_file(str(input_path), str(output_dir), user_dichtungen)

    # Nachher schauen, was neu dazu gekommen ist
    after = set(output_dir.rglob("*.xlsx"))
    new_files = [p for p in after if p not in before]

    if not new_files:
        raise RuntimeError(
            "Konvertierung hat keine neue Excel-Datei erzeugt "
            f"(output_dir={output_dir})."
        )

    # Falls mehrere neu sind, nehmen wir die jüngste
    result = max(new_files, key=lambda p: p.stat().st_mtime)

    # Optional umbenennen auf gewünschten Namen
    if desired_stem:
        target = output_dir / f"{desired_stem}.xlsx"
        try:
            if target.exists():
                target.unlink()
            result.rename(target)
            result = target
        except Exception:
            # Wenn das Umbenennen scheitert, loggen wir es, nutzen aber trotzdem result
            app.logger.exception(
                "Konnte Ausgabedatei nicht in '%s' umbenennen – verwende Originalnamen.",
                target,
            )

    return result


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        # Startseite mit Formular
        return render_template("index.html")

    # POST: Datei wurde hochgeladen
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

    # Optionaler gewünschter Dateiname (ohne .xlsx)
    custom_name_raw = request.form.get("custom_output_name", "").strip()
    desired_stem: Optional[str] = (
        secure_filename(custom_name_raw) if custom_name_raw else None
    )

    # Temporäres Arbeitsverzeichnis
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Eingabedatei speichern
        safe_input_name = secure_filename(file.filename)
        input_path = tmpdir_path / safe_input_name
        file.save(input_path)

        try:
            # Konvertierung ausführen – output_dir = tmpdir_path
            converted_path = run_conversion(input_path, tmpdir_path, desired_stem)
        except Exception:
            # Fehler loggen und direkt als HTML ausgeben
            app.logger.exception("Fehler beim Konvertieren")
            return (
                "Unerwarteter Fehler im Server<br><br>"
                "<pre>"
                + traceback.format_exc()
                + "</pre>",
                500,
            )

        # Sicherheitscheck: existiert die Datei wirklich?
        if not converted_path.exists():
            app.logger.error(
                "Ausgabedatei existiert nicht: %s", converted_path
            )
            return (
                "Fehler: Die konvertierte Datei konnte nicht gefunden werden.",
                500,
            )

        # Fertige Excel an den Browser schicken
        return send_file(
            converted_path,
            as_attachment=True,
            download_name=converted_path.name,
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
