from flask import Flask, request, send_file, jsonify
import os
import fitz  # PyMuPDF for extracting text
from pdf2docx import Converter  # For better DOCX conversion
import subprocess  # For Pandoc execution
import pytesseract  # OCR for scanned PDFs
from PIL import Image

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Set Tesseract OCR Path (Change this path as per your system)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_text_from_pdf(pdf_path):
    """ Extracts text from a PDF using PyMuPDF """
    doc = fitz.open(pdf_path)
    extracted_text = ""
    for page in doc:
        extracted_text += page.get_text("text") + "\n"
    return extracted_text.strip()

def convert_pdf_to_docx(pdf_path, docx_path):
    """ Converts PDF to DOCX using pdf2docx """
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

def convert_pdf_with_pandoc(pdf_path, docx_path):
    """ Uses Pandoc to convert PDF to DOCX for better formatting """
    try:
        command = ["pandoc", pdf_path, "-o", docx_path]
        subprocess.run(command, check=True)
    except Exception as e:
        print(f"Error using Pandoc: {e}")

def extract_text_from_images(pdf_path):
    """ Extracts text from scanned PDFs using Tesseract OCR """
    doc = fitz.open(pdf_path)
    extracted_text = ""
    
    for page_number in range(len(doc)):
        img = doc[page_number].get_pixmap()
        img_path = f"output/page_{page_number}.png"
        img.save(img_path)

        text = pytesseract.image_to_string(Image.open(img_path))
        extracted_text += text + "\n"

    return extracted_text.strip()

@app.route("/convert", methods=["POST"])
def convert_file():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    output_filename = os.path.splitext(file.filename)[0] + ".docx"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    try:
        # Convert using pdf2docx
        convert_pdf_to_docx(input_path, output_path)

        # If pdf2docx creates text boxes, use Pandoc for better formatting
        convert_pdf_with_pandoc(input_path, output_path)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
