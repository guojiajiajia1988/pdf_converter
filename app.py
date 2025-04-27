from flask import Flask, request, send_file, render_template
import os
import uuid
import time
import zipfile
from pdf2docx import Converter
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfMerger
from PIL import Image
from flask import send_from_directory

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
    # docx2pdf要求输出目录
    docx2pdf_convert(input_path, output_folder)

def ppt_to_pdf(input_path, output_path):
    import comtypes.client

    time.sleep(0.5)
    absolute_input_path = os.path.abspath(input_path).replace('/', '\\')
    absolute_output_path = os.path.abspath(output_path).replace('/', '\\')

    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(absolute_input_path, WithWindow=False)
    deck.SaveAs(absolute_output_path, 32)  # 32 = pdf
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
        return "No file or operation selected."

    task_id = str(uuid.uuid4())
    task_output_folder = os.path.join(OUTPUT_FOLDER, task_id)
    os.makedirs(task_output_folder, exist_ok=True)

    if operation == "merge_pdfs":
        # Only for merging PDFs
        saved_files = []
        for file in files:
            if file.filename.lower().endswith('.pdf'):
                safe_filename = str(uuid.uuid4()) + "_" + file.filename.replace(' ', '_')
                input_path = os.path.join(UPLOAD_FOLDER, safe_filename)
                file.save(input_path)
                saved_files.append(input_path)
            else:
                return "Only PDF files are allowed for merging."

        output_pdf = os.path.join(task_output_folder, "merged_output.pdf")
        merge_pdfs(saved_files, output_pdf)
        return send_file(output_pdf, as_attachment=True)

    else:
        # Batch process files
        output_files = []

        for file in files:
            if file.filename == '':
                continue

            ext = file.filename.lower().split('.')[-1]
            safe_filename = str(uuid.uuid4()) + "_" + file.filename.replace(' ', '_')
            input_path = os.path.join(UPLOAD_FOLDER, safe_filename)
            file.save(input_path)

            output_filename = safe_filename.rsplit('.', 1)[0] + "_converted"

            if operation == "pdf_to_word" and ext == 'pdf':
                output_path = os.path.join(task_output_folder, output_filename + ".docx")
                pdf_to_word(input_path, output_path)
                output_files.append(output_path)
            elif operation == "word_to_pdf" and ext == 'docx':
                output_path = os.path.join(task_output_folder, output_filename + ".pdf")
                word_to_pdf(input_path, task_output_folder)
                output_files.append(os.path.join(task_output_folder, file.filename.replace('.docx', '.pdf')))
            elif operation == "ppt_to_pdf" and ext in ['ppt', 'pptx']:
                output_path = os.path.join(task_output_folder, output_filename + ".pdf")
                ppt_to_pdf(input_path, output_path)
                output_files.append(output_path)
            elif operation == "jpg_to_pdf" and ext in ['jpg', 'jpeg', 'png']:
                output_path = os.path.join(task_output_folder, output_filename + ".pdf")
                jpg_to_pdf(input_path, output_path)
                output_files.append(output_path)
            else:
                return f"Unsupported file format: {file.filename}"

        if len(output_files) == 1:
            return send_file(output_files[0], as_attachment=True)

        # Zip multiple files
        zip_filename = os.path.join(OUTPUT_FOLDER, f"{task_id}_converted.zip")
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for file_path in output_files:
                zipf.write(file_path, arcname=os.path.basename(file_path))

        return send_file(zip_filename, as_attachment=True)
        
# --- 新增：提供 sitemap.xml 文件 ---
@app.route('/sitemap.xml')
def sitemap():
    return send_from_directory(directory='.', path='sitemap.xml')
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)  # or just app.run()
