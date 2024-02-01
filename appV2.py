from flask import Flask, render_template, request, send_file
from io import BytesIO
from docx import Document
from docx.image.exceptions import UnrecognizedImageError
from docx.shared import Inches
from bs4 import BeautifulSoup
from utils import generate_date, generate_time, configuration_word

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('indexV2.html')

@app.route('/generate_word', methods=['POST'])
def generate_word():
    responsable = request.form['responsable']
    aplicativo = request.form['aplicativo']
    producto = request.form['producto']
    indicePrueba = request.form['indiceprueba']
    descripcionPrueba = request.form['descripcionprueba']
    observaciones = request.form['observaciones']
    estado = request.form['estado']

    documento = Document()
    configuration_word(documento)

    documento.add_heading(f'REPORTE DE PRUEBAS FLAGON API CORE PYC - {producto}', 0)

    documento.add_paragraph(f'Responsable: {responsable}')
    documento.add_paragraph(f'Aplicativo: {aplicativo}')
    documento.add_paragraph(f'Producto: {producto}')
    documento.add_paragraph(f'Indice caso de prueba: {indicePrueba}')
    documento.add_paragraph(f'Descripcion caso de prueba: {descripcionPrueba}')
    documento.add_paragraph(f'Observaciones: {observaciones}')
    documento.add_paragraph(f'Estado: {estado}') 
    documento.add_paragraph(f'Fecha: {generate_date()}')

    soup = BeautifulSoup(render_template('indexv2.html'), 'html.parser')
    labels = soup.find_all('label', {'name': True})

    for i, label in enumerate(labels, start=1):
        texto = label.text.strip()
        opcion_radio = request.form.get(f'opcion_radio{i}', '')
        opcion_observaciones = request.form.get(f'observacion{i}', '')
        documento.add_paragraph(f'{texto} {opcion_radio}')
        documento.add_paragraph(f'{opcion_observaciones}')

        input_name = f'imagenes{i}[]'
        imagenes = request.files.getlist(input_name)

        if imagenes:
            for imagen in imagenes:
                try:
                    documento.add_picture(imagen, width=Inches(7.5), height=Inches(4.0))
                except UnrecognizedImageError as e:
                    pass

    output_stream = BytesIO()
    documento.save(output_stream)
    output_stream.seek(0)

    return send_file(output_stream, as_attachment=True, download_name=f'EVIDENCIA QA {aplicativo} {producto} {indicePrueba}.docx')

if __name__ == '__main__':
    app.run(debug=True)
