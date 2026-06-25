import os, io, subprocess, tempfile
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

@app.route("/")
def index():
    return "PDF conversion service is running."

@app.route("/pdf", methods=["POST"])
def convert_pdf():
    if "file" not in request.files:
        return jsonify(error="missing file"), 400
    docx_bytes = request.files["file"].read()
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "input.docx")
        pdf_path  = os.path.join(tmpdir, "input.pdf")
        with open(docx_path, "wb") as fp:
            fp.write(docx_bytes)
        try:
            r = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, docx_path],
                capture_output=True, text=True, timeout=60)
        except FileNotFoundError:
            return jsonify(error="libreoffice not installed"), 500
        except subprocess.TimeoutExpired:
            return jsonify(error="conversion timeout"), 500
        if r.returncode != 0 or not os.path.exists(pdf_path):
            return jsonify(error=r.stderr.strip() or "conversion failed"), 500
        with open(pdf_path, "rb") as fp:
            pdf_bytes = fp.read()
    return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf", as_attachment=False)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
