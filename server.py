from flask import Flask, request, send_file, jsonify
import os
import subprocess
import fitz  # PyMuPDF
from pdf2docx import Converter
import pytesseract

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Ensure Tesseract and Pandoc paths are set
TESSERACT_PATH = "/usr/bin/tesseract"
PANDOC_PATH = "/usr/bin/pandoc"
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    output_filename = os.path.splitext(file.filename)[0] + ".docx"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    try:
        # Convert PDF to DOCX using pdf2docx
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()

        # Additional formatting using Pandoc
        subprocess.run([PANDOC_PATH, output_path, "-o", output_path])

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000, debug=True)
