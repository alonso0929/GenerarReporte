from flask import Flask, render_template, request, send_file
from io import BytesIO
from docxtpl import DocxTemplate
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_word', methods=['POST'])
def generate_word():
    producto = request.form['producto']
    financiamiento = request.form['financiamiento']
    casoprueba = request.form['casoprueba']

    opcion_radio1 = request.form.get('opcion_radio1')
    opcion_radio2 = request.form.get('opcion_radio2')
    opcion_radio3 = request.form.get('opcion_radio3')
    opcion_radio4 = request.form.get('opcion_radio4')
    opcion_radio5 = request.form.get('opcion_radio5')
    opcion_radio6 = request.form.get('opcion_radio6')
    opcion_radio7 = request.form.get('opcion_radio7')
    opcion_radio8 = request.form.get('opcion_radio8')
    opcion_radio9 = request.form.get('opcion_radio9')

    def save_image(image, path):
        if image:
            image.save(path)

    imagen1 = request.files['imagen1']
    imagen_path1 = 'static/imagen1.png'
    save_image(imagen1, imagen_path1)

    imagen2 = request.files['imagen2']
    imagen_path2 = 'static/imagen2.png'
    save_image(imagen2, imagen_path2)  

    imagen3 = request.files['imagen3']
    imagen_path3 = 'static/imagen3.png'
    save_image(imagen3, imagen_path3)  

    imagen4 = request.files['imagen4']
    imagen_path4 = 'static/imagen4.png'
    save_image(imagen4, imagen_path4)  

    imagen5 = request.files['imagen5']
    imagen_path5 = 'static/imagen5.png'
    save_image(imagen5, imagen_path5)   

    imagen6 = request.files['imagen6']
    imagen_path6 = 'static/imagen6.png'
    save_image(imagen6, imagen_path6)  

    imagen7 = request.files['imagen7']
    imagen_path7 = 'static/imagen7.png'
    save_image(imagen7, imagen_path7)  

    imagen8 = request.files['imagen8']
    imagen_path8 = 'static/imagen8.png'
    save_image(imagen8, imagen_path8)  

    imagen9 = request.files['imagen9']
    imagen_path9 = 'static/imagen9.png'
    save_image(imagen9, imagen_path9) 

    template_path = 'templates/template.docx'
    document = DocxTemplate(template_path)

    context = {
        'producto': producto,
        'financiamiento': financiamiento,
        'casoprueba': casoprueba,
        'opcion_radio1': opcion_radio1,
        'opcion_radio2': opcion_radio2,
        'opcion_radio3': opcion_radio3,
        'opcion_radio4': opcion_radio4,
        'opcion_radio5': opcion_radio5,
        'opcion_radio6': opcion_radio6,
        'opcion_radio7': opcion_radio7,
        'opcion_radio8': opcion_radio8,
        'opcion_radio9': opcion_radio9, 
        'imagen1': InlineImage(document, imagen_path1, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path1) else None,
        'imagen2': InlineImage(document, imagen_path2, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path2) else None,  
        'imagen3': InlineImage(document, imagen_path3, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path3) else None,  
        'imagen4': InlineImage(document, imagen_path4, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path4) else None,  
        'imagen5': InlineImage(document, imagen_path5, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path5) else None,  
        'imagen6': InlineImage(document, imagen_path6, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path6) else None,  
        'imagen7': InlineImage(document, imagen_path7, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path7) else None,  
        'imagen8': InlineImage(document, imagen_path8, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path8) else None,
        'imagen9': InlineImage(document, imagen_path9, width=Inches(7.5), height=Inches(4.0)) if os.path.exists(imagen_path9) else None
    }

    document.render(context)

    output_stream = BytesIO()
    document.save(output_stream)
    output_stream.seek(0)

    for image_path in [imagen_path1, imagen_path2, imagen_path3, imagen_path4, imagen_path5, imagen_path6, imagen_path7, imagen_path8, imagen_path9]:
        if os.path.exists(image_path):
            os.remove(image_path)

    return send_file(output_stream, as_attachment=True, download_name=f'APIs_CORE_P&C_EVIDENCIAQA_{producto}_{financiamiento}.docx')

if __name__ == '__main__':
    app.run(debug=True)