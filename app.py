from keywords import building_keywords, floor_keywords, apartment_keywords, room_keywords
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
from openpyxl import Workbook
from werkzeug.utils import secure_filename
import uuid
from utils import extract_text_from_pdf, extract_data, write_data_to_excel


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
        if 'pdf_files' not in request.files:
            flash('No file part')
            return redirect(request.url)

        files = request.files.getlist('pdf_files')

        data = {}
        for file in files:
            if file.filename == '':
                flash('No selected file')
                return redirect(request.url)
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                unique_id = str(uuid.uuid4())
                pdf_path = os.path.join(
                    app.config['UPLOAD_FOLDER'], unique_id + '.pdf')
                file.save(pdf_path)

                text = extract_text_from_pdf(pdf_path)  # Util function
                print(f"Extracted text from {file.filename}:")
                print(text)
                print("\n" + "=" * 80 + "\n")

                floor_range = request.form.get('floor_range')

                special_floors = request.form.get('special_floors')

                apartment_pattern = request.form.get('apartment_format')
                print(f"apartmentpattern {apartment_pattern}")
                building_data = extract_data(
                    text, apartment_pattern, floor_range, special_floors)   # Util function

                for building, floors in building_data.items():
                    if building in data:
                        for floor, apartments in floors.items():
                            if floor in data[building]:
                                data[building][floor].update(apartments)
                            else:
                                data[building][floor] = apartments
                    else:
                        data[building] = floors

        for building, floors in data.items():
            sorted_floors = sorted(
                floors.keys(), key=lambda x: (x.isdigit() == False, x))
            sorted_data = {floor: data[building][floor]
                           for floor in sorted_floors}
            data[building] = sorted_data

        unique_id = str(uuid.uuid4())
        output_path = os.path.join(
            app.config['UPLOAD_FOLDER'], unique_id + '.xlsx')
        write_data_to_excel(data, output_path)  # Util function

        return send_file(output_path, as_attachment=True, download_name='output.xlsx')

    return render_template('index.html')


# ... the rest of the app.py code
if __name__ == '__main__':
    app.run(debug=True)
