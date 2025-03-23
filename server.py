import os
import subprocess
import fitz  # PyMuPDF
import pytesseract
from pdf2docx import Converter
from flask import Flask, request, send_file, jsonify
from werkzeug.utils import secure_filename

# Create necessary folders
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

# Path to Tesseract OCR (ensure it's installed on server)
TESSERACT_PATH = "/usr/bin/tesseract"  # Update if necessary
pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


@app.route("/", methods=["GET"])
def home():
    return "Welcome to the PDF to DOCX Converter API"


@app.route("/convert", methods=["POST"])
def convert_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    output_docx = os.path.join(app.config["OUTPUT_FOLDER"], filename.replace(".pdf", ".docx"))

    file.save(input_path)

    try:
        # Use pdf2docx as the primary converter
        convert_using_pdf2docx(input_path, output_docx)

        # Use Pandoc as a secondary method (if installed)
        pandoc_output = output_docx.replace(".docx", "_pandoc.docx")
        convert_using_pandoc(input_path, pandoc_output)

        return send_file(output_docx, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


def convert_using_pdf2docx(pdf_path, output_path):
    """ Converts PDF to DOCX using pdf2docx (better formatting) """
    cv = Converter(pdf_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()


def convert_using_pandoc(pdf_path, output_path):
    """ Converts PDF to DOCX using Pandoc """
    try:
        subprocess.run(["pandoc", pdf_path, "-o", output_path], check=True)
    except Exception as e:
        print(f"Pandoc conversion failed: {e}")


def extract_text_using_tesseract(pdf_path):
    """ Uses Tesseract OCR for scanned PDFs """
    doc = fitz.open(pdf_path)
    text = ""

    for page in doc:
        image = page.get_pixmap()
        text += pytesseract.image_to_string(image.tobytes(), lang="eng") + "\n"

    return text


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
