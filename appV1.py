from flask import Flask, render_template, request, send_file
from io import BytesIO
from docxtpl import DocxTemplate
import os
from utils import save_image
from utils import generate_image_paths
from utils import generate_inline_images
from utils import generate_date
from utils import generate_time

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_word', methods=['POST'])
def generate_word():
    producto = request.form['producto']
    financiamiento = request.form['financiamiento']
    casoprueba = request.form['casoprueba']
    autor = request.form['responsable']

    opcion_radio = {f'opcion_radio{i}': request.form.get(f'opcion_radio{i}') for i in range(1, 12)}

    num_images = 11
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
