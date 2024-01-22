from flask import Flask, render_template, request, send_file
from io import BytesIO
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import os
from datetime import datetime

app = Flask(__name__)

def save_image(image, path):
    if image:
        image.save(path)

def generate_image_paths(num_images):
    return [f'static/imagen{i}.png' for i in range(1, num_images + 1)]

def generate_inline_images(document, image_paths):
    return [InlineImage(document, path, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(path) else None for path in image_paths]

def generate_date():
    date_time = datetime.now()
    return date_time.strftime("%d-%m-%Y")

def generate_time():
    date_time = datetime.now()
    return date_time.strftime("%H:%M:%S")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_word', methods=['POST'])
def generate_word():
    producto = request.form['producto']
    financiamiento = request.form['financiamiento']
    casoprueba = request.form['casoprueba']
    autor = request.form['responsable']

    opcion_radio = {f'opcion_radio{i}': request.form.get(f'opcion_radio{i}') for i in range(1, 10)}

    num_images = 9
    image_paths = generate_image_paths(num_images)

    for i in range(1, num_images + 1):
        save_image(request.files[f'imagen{i}'], image_paths[i - 1])

    template_path = 'templates/template.docx'
    document = DocxTemplate(template_path)

    context = {
        'producto': producto,
        'financiamiento': financiamiento,
        'casoprueba': casoprueba,
        'autor': autor,
        'fecha': generate_date(),
        'hora': generate_time(),
        **opcion_radio,
        **{f'imagen{i}': image for i, image in enumerate(generate_inline_images(document, image_paths), start=1)}
    }

    document.render(context)

    output_stream = BytesIO()
    document.save(output_stream)
    output_stream.seek(0)

    for image_path in image_paths:
        if os.path.exists(image_path):
            os.remove(image_path)

    return send_file(output_stream, as_attachment=True, download_name=f'APIs_CORE_P&C_EVIDENCIAQA_{producto}_{financiamiento}.docx')

if __name__ == '__main__':
    app.run(debug=True)
