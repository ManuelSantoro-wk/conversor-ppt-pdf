from flask import Flask, render_template, request, send_from_directory, redirect, url_for
import os
import subprocess
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    uploaded_files = request.files.getlist("files")
    converted_files = []

    for file in uploaded_files:
        if file and file.filename.endswith(('.ppt', '.pptx')):
            filename = secure_filename(file.filename)
            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)

            # Convert with LibreOffice
            subprocess.run([r"C:\Program Files\LibreOffice\program\soffice.exe",  # Caminho completo para o soffice.exe
    "--headless",
    "--convert-to", "pdf",
    "--outdir", CONVERTED_FOLDER,
    input_path
])

            pdf_filename = filename.rsplit('.', 1)[0] + '.pdf'
            converted_files.append(pdf_filename)

    return render_template('index.html', converted_files=converted_files)

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(CONVERTED_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
