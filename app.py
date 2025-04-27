from flask import Flask, request, send_file, render_template, send_from_directory, jsonify
import os
import uuid
import time
from pdf2docx import Converter
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfMerger
from PIL import Image

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --- Conversion functions ---

def pdf_to_word(input_path, output_path):
    cv = Converter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def word_to_pdf(input_path, output_folder):
    docx2pdf_convert(input_path, output_folder)

def ppt_to_pdf(input_path, output_path):
    import comtypes.client
    time.sleep(0.5)
    absolute_input_path = os.path.abspath(input_path).replace('/', '\\')
    absolute_output_path = os.path.abspath(output_path).replace('/', '\\')

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(absolute_input_path, WithWindow=False)
    deck.SaveAs(absolute_output_path, 32)
    deck.Close()
    powerpoint.Quit()

def jpg_to_pdf(input_path, output_path):
    image = Image.open(input_path)
    rgb_im = image.convert('RGB')
    rgb_im.save(output_path)

def merge_pdfs(file_list, output_path):
    merger = PdfMerger()
    for pdf in file_list:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

# --- Routes ---

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    files = request.files.getlist('file')
    operation = request.form.get('operation')

    if not files or not operation:
        return "No file or operation selected.", 400

    task_id = str(uuid.uuid4())
    task_output_folder = os.path.join(OUTPUT_FOLDER, task_id)
    os.makedirs(task_output_folder, exist_ok=True)

    converted_files = []

    if operation == "merge_pdfs":
        saved_files = []
        for file in files:
            if file.filename.lower().endswith('.pdf'):
                input_path = os.path.join(UPLOAD_FOLDER, file.filename.replace(' ', '_'))
                file.save(input_path)
                saved_files.append(input_path)
            else:
                return "Only PDF files are allowed for merging.", 400

        output_pdf = os.path.join(task_output_folder, "merged_output.pdf")
        merge_pdfs(saved_files, output_pdf)
        converted_files.append(output_pdf)

    else:
        for file in files:
            if file.filename == '':
                continue

            filename = file.filename.replace(' ', '_')
            base_name, ext = os.path.splitext(filename)

            input_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(input_path)

            output_filename = base_name + "_con"

            if operation == "pdf_to_word" and ext.lower() == '.pdf':
                output_path = os.path.join(task_output_folder, output_filename + ".docx")
                pdf_to_word(input_path, output_path)
                converted_files.append(output_path)

            elif operation == "word_to_pdf" and ext.lower() == '.docx':
                output_path = os.path.join(task_output_folder, output_filename + ".pdf")
                word_to_pdf(input_path, task_output_folder)
                converted_files.append(output_path)

            elif operation == "ppt_to_pdf" and ext.lower() in ['.ppt', '.pptx']:
                output_path = os.path.join(task_output_folder, output_filename + ".pdf")
                ppt_to_pdf(input_path, output_path)
                converted_files.append(output_path)

            elif operation == "jpg_to_pdf" and ext.lower() in ['.jpg', '.jpeg', '.png']:
                output_path = os.path.join(task_output_folder, output_filename + ".pdf")
                jpg_to_pdf(input_path, output_path)
                converted_files.append(output_path)

            else:
                return f"Unsupported file format: {file.filename}", 400

    # 返回下载链接列表
    download_links = []
    for path in converted_files:
        file_url = f"/download/{task_id}/{os.path.basename(path)}"
        download_links.append(file_url)

    return jsonify(download_links)

@app.route('/download/<task_id>/<filename>')
def download_file(task_id, filename):
    dir_path = os.path.join(OUTPUT_FOLDER, task_id)
    return send_from_directory(directory=dir_path, path=filename, as_attachment=True)

@app.route('/sitemap.xml')
def sitemap():
    return send_from_directory(directory='.', path='sitemap.xml')

@app.route('/robots.txt')
def robots():
    return send_from_directory(directory='.', path='robots.txt')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
