from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No se encontró el archivo", 400
    file = request.files["file"]
    if file.filename == "":
        return "No se seleccionó ningún archivo", 400

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Procesar el archivo
    df = pd.read_excel(file_path)
    df["Nueva Columna"] = "Procesado"  # Modificación de ejemplo

    processed_file_path = os.path.join(PROCESSED_FOLDER, "procesado.xlsx")
    df.to_excel(processed_file_path, index=False)

    return send_file(processed_file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
