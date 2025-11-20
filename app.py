import os
import tempfile

from flask import (
    Flask, render_template, request, send_file,
    redirect, url_for, flash, after_this_request
)

# Immer Headless
os.environ.setdefault("HEADLESS", "1")

from packliste_core import convert_file, load_dichtungen, save_dichtungen  # noqa: E402

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")


def allowed_file(filename: str) -> bool:
    return "." in filename and \
        filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "input_file" not in request.files:
            flash("Keine Datei hochgeladen.", "error")
            return redirect(request.url)

        file = request.files["input_file"]
        if file.filename == "":
            flash("Keine Datei ausgewählt.", "error")
            return redirect(request.url)

        if not allowed_file(file.filename):
            flash("Ungültiger Dateityp. Erlaubt sind .xlsx, .xls, .csv.",
                  "error")
            return redirect(request.url)

        # Temporäre Dateien
        in_suffix = "." + file.filename.rsplit(".", 1)[1].lower()
        tmp_in = tempfile.NamedTemporaryFile(
            delete=False, suffix=in_suffix
        )
        file.save(tmp_in.name)
        tmp_in.close()

        tmp_out = tempfile.NamedTemporaryFile(
            delete=False, suffix=".xlsx"
        )
        out_path = tmp_out.name
        tmp_out.close()

        try:
            user_dichtungen = load_dichtungen()
            convert_file(tmp_in.name, out_path, user_dichtungen)
        except Exception as e:
            flash(f"Fehler bei der Konvertierung: {e}", "error")
            try:
                os.remove(out_path)
            except Exception:
                pass
            return redirect(request.url)
        finally:
            try:
                os.remove(tmp_in.name)
            except Exception:
                pass

        download_name = "Packliste_konvertiert.xlsx"

        @after_this_request
        def cleanup(response):
            try:
                os.remove(out_path)
            except Exception:
                pass
            return response

        return send_file(out_path,
                         as_attachment=True,
                         download_name=download_name)

    return render_template("index.html")


@app.route("/dichtungen", methods=["GET", "POST"])
def manage_dichtungen():
    if request.method == "POST":
        try:
            max_index = int(request.form.get("max_index", "0"))
        except ValueError:
            max_index = 0

        new_cfg = []
        for i in range(max_index + 1):
            name = request.form.get(f"name_{i}", "").strip()
            if not name:
                continue
            always = request.form.get(f"always_{i}") == "on"
            value_str = request.form.get(f"value_{i}", "").strip()
            order_str = request.form.get(f"order_{i}", "").strip()

            try:
                default_value = float(
                    value_str.replace(",", ".")
                ) if value_str else 0.0
            except Exception:
                default_value = 0.0

            new_cfg.append({
                "name": name,
                "always_show": always,
                "default_value": default_value,
                "order": order_str,
            })

        save_dichtungen(new_cfg)
        flash("Dichtungen gespeichert.", "success")
        return redirect(url_for("manage_dichtungen"))

    # GET
    dichtungen = load_dichtungen()
    # Ensure keys for template
    for d in dichtungen:
        d.setdefault("always_show", False)
        d.setdefault("default_value", 0)
        d.setdefault("order", "")

    extra_rows = 5
    total_rows = len(dichtungen) + extra_rows
    return render_template("dichtungen.html",
                           dichtungen=dichtungen,
                           total_rows=total_rows)


@app.route("/health")
def health():
    return {"status": "ok"}
