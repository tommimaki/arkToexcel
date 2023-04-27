from keywords import building_keywords, floor_keywords, apartment_keywords, room_keywords
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
import re
import pdfplumber
import pytesseract
from wand.image import Image
from openpyxl import Workbook
import tempfile
from PyPDF2 import PdfReader
from PIL import Image as PILImage
from werkzeug.utils import secure_filename
import uuid


# def extract_text_from_pdf(pdf_path):
#     with pdfplumber.open(pdf_path) as pdf:
#         text = ""
#         for page in pdf.pages:
#             # Extract selectable text from the current page
#             page_text = page.extract_text()
#             if page_text:
#                 text += page_text + "\n"

#             # Convert the page to an image and extract text using OCR with Tesseract
#             #  needs resolution fix prolly
#             img = page.to_image(resolution=400)

#             # Convert to black and white to improve OCR accuracy
#             with tempfile.NamedTemporaryFile(suffix=".png") as temp_img:
#                 img.save(temp_img.name, format="PNG")
#                 with Image(filename=temp_img.name) as wand_img:
#                     wand_img.type = 'bilevel'
#                     # Save wand image as PNG and open it with PIL
#                     with tempfile.NamedTemporaryFile(suffix=".png") as pil_temp_img:
#                         wand_img.save(filename=pil_temp_img.name)
#                         pil_img = PILImage.open(pil_temp_img.name)
#                         ocr_text = pytesseract.image_to_string(pil_img)
#                 text += ocr_text + "\n"
#     return text
from PyPDF2 import PdfReader


def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            # Extract selectable text from the current page
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"

            # Convert the page to an image and extract text using OCR with Tesseract
            img = page.to_image(resolution=400)

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

        # Extract annotations and form fields using PyPDF2
        with open(pdf_path, "rb") as f:
            reader = PdfReader(f)
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                if "/Annots" in page:
                    for annot in page["/Annots"]:
                        annot_obj = annot.getObject()
                        if "/Contents" in annot_obj:
                            text += annot_obj["/Contents"].replace(
                                "\n", " ").strip() + "\n"

            acroform = reader.trailer.get("/AcroForm")
            if acroform:
                form_fields = acroform["/Fields"]
                for field in form_fields:
                    field_obj = field.getObject()
                    if "/V" in field_obj:
                        value = field_obj["/V"]
                        if value not in ["", "/Off"]:
                            text += value.replace("\n", " ").strip() + "\n"

    return text


# Function to extract relevant data from the text using regular expressions


# def extract_data(text, apartment_pattern):
#     buildings_data = {}

#     building_keyword_pattern = "|".join(building_keywords)
#     # floor_keyword_pattern = "|".join(floor_keywords)
#     floor_keyword_pattern = r"(\d+)(?:\s*\.?\s*|\.)(?:" + \
#         "|".join(floor_keywords) + ")"
#     # apartment_keyword_pattern = "|".join(apartment_keywords)
#     apartment_keyword_pattern = apartment_pattern

#     building_matches = re.findall(
#         fr"(?:{building_keyword_pattern})\s+(\w+)", text)

#     for building in building_matches:
#         if building not in buildings_data:
#             buildings_data[building] = {}

#         floor_matches = re.finditer(floor_keyword_pattern, text)
#         for floor_match in floor_matches:
#             floor = floor_match.group(1)

#             if floor not in buildings_data[building]:
#                 # initialize with an empty set
#                 buildings_data[building][floor] = set()

#             apartment_matches = re.findall(
#                 fr"(?:{apartment_keyword_pattern})", text)
#             for apartment in apartment_matches:
#                 buildings_data[building][floor].add(apartment)

#     return buildings_data

def extract_data(text, apartment_pattern):
    buildings_data = {}

    building_keyword_pattern = "|".join(building_keywords)
    floor_keyword_pattern = r"(\d+)(?:\s*\.?\s*|\.)(?:" + \
        "|".join(floor_keywords) + ")"
    apartment_keyword_pattern = apartment_pattern
    room_keyword_pattern = "|".join(room_keywords)

    building_matches = re.findall(
        fr"(?:{building_keyword_pattern})\s+(\w+)", text)

    for building in building_matches:
        if building not in buildings_data:
            buildings_data[building] = {}

        floor_matches = re.finditer(floor_keyword_pattern, text)
        for floor_match in floor_matches:
            floor = floor_match.group(1)

            if floor not in buildings_data[building]:
                # initialize with an empty set
                buildings_data[building][floor] = {}

            apartment_matches = re.findall(
                fr"(?:{apartment_keyword_pattern})", text)
            for apartment in apartment_matches:
                if apartment not in buildings_data[building][floor]:
                    buildings_data[building][floor][apartment] = {}

                room_matches = re.finditer(
                    fr"(?:{room_keyword_pattern})\s*([\d.,]+)", text)
                for room_match in room_matches:
                    room_type = room_match.group(0).split()[0]
                    room_size = room_match.group(1)
                    buildings_data[building][floor][apartment][room_type] = room_size

    return buildings_data


def write_data_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active

    headers = ["Building", "Floor", "Apartment", "Room"]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    row = 2
    for building, floors in data.items():
        ws.cell(row=row, column=1).value = building
        row += 1
        for floor, apartments in floors.items():
            ws.cell(row=row, column=2).value = f"Kerros {floor}"
            row += 1
            for apartment, rooms in apartments.items():
                ws.cell(row=row, column=3).value = apartment
                row += 1
                for room in rooms.keys():
                    ws.cell(row=row, column=4).value = room
                    row += 1
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

                # Get the user's choice for the apartment format
                apartment_pattern = request.form.get('apartment_format')
                # Pass the apartment_pattern to the extract_data function
                building_data = extract_data(text, apartment_pattern)

                # Merge building data
                for building, floors in building_data.items():
                    if building in data:
                        for floor, apartments in floors.items():
                            if floor in data[building]:
                                data[building][floor].update(apartments)
                            else:
                                data[building][floor] = apartments
                    else:
                        data[building] = floors

        # Convert sets to lists for consistency and sort the floors
        # for building, floors in data.items():
        #     data[building] = sorted(
        #         list(floors), key=lambda x: (x.isdigit() == False, x))
        for building, floors in data.items():
            sorted_floors = sorted(
                floors.keys(), key=lambda x: (x.isdigit() == False, x))
            sorted_data = {floor: data[building][floor]
                           for floor in sorted_floors}
            data[building] = sorted_data

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


# def extract_data(text):
#     buildings_data = {}
#     # Update the regex pattern to capture variations in floor and apartment names
#     building_floor_pairs = re.findall(
#         r"RAKENNUS\s+(\w+),\s+((?:\d+\.\s+)?(?:KERROS|VESIKATTO|AS \d+))", text)

#     for building, floor in building_floor_pairs:
#         if building not in buildings_data:
#             buildings_data[building] = {floor}
#         else:
#             buildings_data[building].add(floor)

#     # Convert sets to lists for consistency and sort the floors
#     for building, floors in buildings_data.items():
#         buildings_data[building] = sorted(
#             list(floors), key=lambda x: (x.isdigit() == False, x))

#     return buildings_data

# Function to write the extracted data to an Excel file

# def extract_data(text):
#     buildings_data = {}

#     building_keyword_pattern = "|".join(building_keywords)
#     floor_keyword_pattern = "|".join(floor_keywords)
#     apartment_keyword_pattern = "|".join(aprtment_keywords)

#     building_matches = re.findall(
#         fr"(?:{building_keyword_pattern})\s+(\w+)", text)

#     for building in building_matches:
#         if building not in buildings_data:
#             buildings_data[building] = {}

#         floor_matches = re.findall(
#             fr"(?:{floor_keyword_pattern})\s+(\w+)", text)
#         for floor in floor_matches:
#             if floor not in buildings_data[building]:
#                 buildings_data[building][floor] = []

#             apartment_matches = re.findall(
#                 fr"(?:{apartment_keyword_pattern})", text)
#             for apartment in apartment_matches:
#                 if apartment not in buildings_data[building][floor]:
#                     buildings_data[building][floor].append(apartment)

#     return buildings_data
