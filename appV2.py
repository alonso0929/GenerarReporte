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
    producto = request.form['producto']
    financiamiento = request.form['financiamiento']
    casoprueba = request.form['casoprueba']
    autor = request.form['responsable']

    documento = Document()
    configuration_word(documento)

    documento.add_heading(f'REPORTE DE PRUEBAS FLAGON API CORE PYC - {producto} {financiamiento}', 0)

    documento.add_paragraph(f'Producto: {producto}')
    documento.add_paragraph(f'Financiamiento: {financiamiento}')
    documento.add_paragraph(f'Caso de prueba: {casoprueba}')
    documento.add_paragraph(f'Autor: {autor}')
    documento.add_paragraph(f'Fecha: {generate_date()}')
    documento.add_paragraph(f'Hora: {generate_time()}')

    soup = BeautifulSoup(render_template('indexv2.html'), 'html.parser')
    labels = soup.find_all('label', {'name': True})

    for i, label in enumerate(labels, start=1):
        texto = label.text.strip()
        opcion_radio = request.form.get(f'opcion_radio{i}', '')
        documento.add_paragraph(f'{texto} {opcion_radio}')

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

    return send_file(output_stream, as_attachment=True, download_name=f'APIs_CORE_P&C_EVIDENCIA_QA_{producto}_{financiamiento}.docx')

if __name__ == '__main__':
    app.run(debug=True)
