import os
import sys
import shutil
import tempfile
import time
import traceback
from pathlib import Path
from datetime import datetime

from flask import (
    Flask,
    render_template,
    request,
    send_file,
    redirect,
    url_for,
    flash,
)

# -------------------------------------------------
# Basis-Konfiguration
# -------------------------------------------------
ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}

BASE_DIR = Path(__file__).resolve().parent

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")


def allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_user_dichtungen() -> dict:
    """
    Optional: user_dichtungen.json einlesen (wenn vorhanden).
    Falls es die Datei nicht gibt oder fehlerhaft ist → einfach leeres Dict.
    """
    cfg_path = BASE_DIR / "user_dichtungen.json"
    if not cfg_path.is_file():
        return {}

    try:
        import json

        with cfg_path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:  # nur loggen, nicht abbrechen
        print(f"Warnung: user_dichtungen.json konnte nicht gelesen werden: {exc}", file=sys.stderr)
        return {}


def run_conversion(input_path: Path, tmpdir_path: Path, desired_stem: str | None) -> Path:
    """
    Wrapper um packliste_core.convert_file.

    Egal, ob convert_file den 2. Parameter als Datei-Pfad oder Ordner interpretiert:
    Wir rufen es auf und suchen danach nach der neu/aktualisierten .xlsx-Datei.

    Rückgabe: Pfad zur endgültigen Excel-Datei im tmpdir_path.
    """
    from packliste_core import convert_file

    user_dichtungen = load_user_dichtungen()

    tmpdir_path.mkdir(parents=True, exist_ok=True)

    # Basisname für die Ausgabedatei wählen
    if desired_stem:
        base_name = desired_stem.strip()
    else:
        base_name = input_path.stem

    if not base_name:
        base_name = "Packliste"

    # Primärer, bevorzugter Zielpfad
    primary_output = tmpdir_path / f"{base_name}.xlsx"

    # Kandidaten-Verzeichnisse, in denen wir nach neuer/aktualisierter Excel suchen
    candidate_dirs = {
        tmpdir_path,
        input_path.parent,
        BASE_DIR,  # Projekt-Root (dort liegt auch Packliste_Template.xlsx)
    }

    # Vorher-Zeitstempel und bekannte Dateien merken
    baseline_mtimes: dict[Path, float] = {}
    for d in candidate_dirs:
        if not d.is_dir():
            continue
        for p in d.glob("*.xlsx"):
            try:
                baseline_mtimes[p.resolve()] = p.stat().st_mtime
            except FileNotFoundError:
                continue

    start_time = time.time()

    # ----------------- Konverter ausführen -----------------
    # 2. Parameter als Dateipfad übergeben (falls convert_file das so erwartet)
    convert_file(str(input_path), str(primary_output), user_dichtungen)

    # 1) Ideal: convert_file hat genau diese Datei geschrieben
    if primary_output.exists():
        return primary_output

    # 2) Sonst suchen wir nach der "neuesten" geänderten/erzeugten .xlsx
    newest_path: Path | None = None
    newest_mtime: float = 0.0

    for d in candidate_dirs:
        if not d.is_dir():
            continue
        for p in d.glob("*.xlsx"):
            try:
                rp = p.resolve()
                mtime = p.stat().st_mtime
            except FileNotFoundError:
                continue

            old_mtime = baseline_mtimes.get(rp)

            # "Neu" = gab es vorher nicht ODER wurde nach dem Start geändert
            if (old_mtime is None or mtime > old_mtime + 1e-6) and mtime >= start_time - 1.0:
                if mtime > newest_mtime:
                    newest_mtime = mtime
                    newest_path = rp

    if newest_path is None:
        # Nichts gefunden → Fehler hochreichen
        raise RuntimeError(f"Konvertierung hat keine neue Excel-Datei erzeugt (output_dir={tmpdir_path})")

    # Datei ins tmpdir unter unserem gewünschten Namen kopieren
    final_path = tmpdir_path / f"{base_name}.xlsx"
    if newest_path != final_path:
        try:
            shutil.copy2(newest_path, final_path)
        except Exception:
            shutil.move(newest_path, final_path)

    return final_path


# -------------------------------------------------
# Routes
# -------------------------------------------------


@app.route("/", methods=["GET", "POST"])
def index():
    # GET → leeres Formular anzeigen
    if request.method == "GET":
        default_name = f"Packliste_{datetime.now().strftime('%Y%m%d')}"
        return render_template(
            "index.html",
            error_message=None,
            traceback_str=None,
            default_name=default_name,
        )

    # POST → Datei verarbeiten
    if "input_file" not in request.files or request.files["input_file"].filename == "":
        flash("Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).", "error")
        return redirect(url_for("index"))

    file = request.files["input_file"]

    if not allowed(file.filename):
        flash("Ungültiger Dateityp. Erlaubt sind: .xlsx, .xls, .csv.", "error")
        return redirect(url_for("index"))

    desired_name_raw = request.form.get("desired_name", "").strip()
    desired_stem = desired_name_raw.replace(".xlsx", "").strip() or None

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)

            input_path = tmpdir_path / file.filename
            file.save(input_path)

            converted_path = run_conversion(input_path, tmpdir_path, desired_stem)

            download_name = converted_path.name

            return send_file(
                converted_path,
                as_attachment=True,
                download_name=download_name,
            )

    except Exception as exc:
        tb = traceback.format_exc()
        print(tb, file=sys.stderr)

        # Fallback für den Feld-Vorschlag
        try:
            default_name = (
                desired_stem
                or (input_path.stem if "input_path" in locals() else f"Packliste_{datetime.now().strftime('%Y%m%d')}")
            )
        except Exception:
            default_name = f"Packliste_{datetime.now().strftime('%Y%m%d')}"

        error_msg = f"Unerwarteter Fehler im Server: {exc}"

        return (
            render_template(
                "index.html",
                error_message=error_msg,
                traceback_str=tb,
                default_name=default_name,
            ),
            500,
        )


if __name__ == "__main__":
    # Lokal testen
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
