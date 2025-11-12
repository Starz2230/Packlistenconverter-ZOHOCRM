
import os
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import tempfile
from packliste_core import convert_file

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}

def allowed(filename):
    return "." in filename and filename.rsplit(".",1)[1].lower() in ALLOWED_EXTENSIONS

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY","dev-key")

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        if "input_file" not in request.files or request.files["input_file"].filename == "":
            flash("Bitte eine Packlisten-Datei hochladen (.xlsx/.xls/.csv).","error")
            return redirect(url_for("index"))
        f = request.files["input_file"]
        if not allowed(f.filename):
            flash("Ungültiges Dateiformat.","error")
            return redirect(url_for("index"))

        # optional: JSON mit Dichtungen aus einem Textfeld
        user_dichtungen = []
        try:
            if request.form.get("user_dichtungen"):
                import json
                user_dichtungen = json.loads(request.form.get("user_dichtungen"))
        except Exception as e:
            flash(f"Dichtungen-JSON ungültig: {e}","error")
            return redirect(url_for("index"))

        tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.filename)[1])
        f.save(tmp_in.name)

        tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_in.close(); tmp_out.close()

        try:
            # call the original converter
            convert_file(tmp_in.name, tmp_out.name, user_dichtungen, show_message=False)
        except Exception as e:
            flash(f"Fehler bei der Konvertierung: {e}","error")
            return redirect(url_for("index"))

        download_name = secure_filename(os.path.splitext(f.filename)[0] + "_konvertiert.xlsx")
        return send_file(tmp_out.name, as_attachment=True, download_name=download_name)

    return render_template("index.html")

@app.route("/health")
def health():
    return {"status":"ok"}

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
