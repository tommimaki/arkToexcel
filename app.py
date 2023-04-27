from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import re
import pdfplumber
import pytesseract
from wand.image import Image
from openpyxl import Workbook
import tempfile
from PIL import Image as PILImage
from werkzeug.utils import secure_filename
import uuid


def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            # Extract selectable text from the current page
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

            # Convert the page to an image and extract text using OCR with Tesseract
            #  needs resolution fix prolly
            img = page.to_image(resolution=300)

            # Convert to black and white to improve OCR accuracy
            with tempfile.NamedTemporaryFile(suffix=".png") as temp_img:
                img.save(temp_img.name, format="PNG")
                with Image(filename=temp_img.name) as wand_img:
                    wand_img.type = 'bilevel'
                    # Save wand image as PNG and open it with PIL
                    with tempfile.NamedTemporaryFile(suffix=".png") as pil_temp_img:
                        wand_img.save(filename=pil_temp_img.name)
                        pil_img = PILImage.open(pil_temp_img.name)
                        ocr_text = pytesseract.image_to_string(pil_img)
                text += ocr_text + "\n"
    return text


# Function to extract relevant data from the text using regular expressions


def extract_data(text):
    buildings_data = {}
    building_floor_pairs = re.findall(
        r"RAKENNUS\s+(\w+),\s+((?:\d+\.\s+)?KERROS|VESIKATTO)", text)
    for building, floor in building_floor_pairs:
        if building not in buildings_data:
            buildings_data[building] = {floor}
        else:
            buildings_data[building].add(floor)

    # Convert sets to lists for consistency and sort the floors
    for building, floors in buildings_data.items():
        buildings_data[building] = sorted(
            list(floors), key=lambda x: (x.isdigit() == False, x))

    return buildings_data


# Function to write the extracted data to an Excel file
def write_data_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active

    headers = ["Building", "Floor"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    row = 2
    for building, floors in data.items():
        ws.cell(row=row, column=1).value = building
        for i, floor in enumerate(floors):
            if i > 0:
                row += 1
            ws.cell(row=row, column=2).value = floor

        row += 1

    wb.save(output_file)


UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)  # Add this line
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = 'secret_key'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'pdf_files' not in request.files:
            flash('No file part')
            return redirect(request.url)

        files = request.files.getlist('pdf_files')  # Get the list of files

        data = {}
        for file in files:
            # If user does not select file, browser submits an empty part without filename
            if file.filename == '':
                flash('No selected file')
                return redirect(request.url)
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                unique_id = str(uuid.uuid4())
                pdf_path = os.path.join(
                    app.config['UPLOAD_FOLDER'], unique_id + '.pdf')
                file.save(pdf_path)

                # Run the main conversion process
                text = extract_text_from_pdf(pdf_path)
                print(f"Extracted text from {file.filename}:")
                print(text)
                print("\n" + "=" * 80 + "\n")

                building_data = extract_data(text)

                # Merge building data
                for building, floors in building_data.items():
                    if building in data:
                        data[building].update(floors)
                    else:
                        data[building] = set(floors)

        # Convert sets to lists for consistency and sort the floors
        for building, floors in data.items():
            data[building] = sorted(
                list(floors), key=lambda x: (x.isdigit() == False, x))

        # Write merged data to Excel
        unique_id = str(uuid.uuid4())
        output_path = os.path.join(
            app.config['UPLOAD_FOLDER'], unique_id + '.xlsx')
        write_data_to_excel(data, output_path)

        return send_file(output_path, as_attachment=True, download_name='output.xlsx')

    return render_template('index.html')


# ... the rest of the app.py code
if __name__ == '__main__':
    app.run(debug=True)
