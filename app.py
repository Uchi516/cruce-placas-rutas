#!/usr/bin/env python3
"""
Web app para procesar archivos Excel de Picking Center.
Sube Archivo 1 (Placas) y Archivo 2 (Rutas), descarga el resultado.
"""

import os
import tempfile
import uuid
from flask import Flask, render_template, request, send_file, jsonify
from procesar_excel import procesar

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/procesar', methods=['POST'])
def procesar_archivos():
    if 'archivo1' not in request.files or 'archivo2' not in request.files:
        return jsonify({'error': 'Debes subir ambos archivos'}), 400

    archivo1 = request.files['archivo1']
    archivo2 = request.files['archivo2']

    if archivo1.filename == '' or archivo2.filename == '':
        return jsonify({'error': 'Debes seleccionar ambos archivos'}), 400

    if not archivo1.filename.endswith('.xlsx') or not archivo2.filename.endswith('.xlsx'):
        return jsonify({'error': 'Los archivos deben ser .xlsx'}), 400

    # Guardar archivos temporales
    tmp_dir = tempfile.mkdtemp()
    ruta1 = os.path.join(tmp_dir, archivo1.filename)
    ruta2 = os.path.join(tmp_dir, archivo2.filename)
    archivo1.save(ruta1)
    archivo2.save(ruta2)

    # Nombre del archivo de salida
    base, ext = os.path.splitext(archivo2.filename)
    nombre_salida = f"{base} - RESULTADO{ext}"
    ruta_salida = os.path.join(tmp_dir, nombre_salida)

    try:
        procesar(ruta1, ruta2, ruta_salida)
    except Exception as e:
        # Limpiar
        for f in [ruta1, ruta2]:
            if os.path.exists(f):
                os.remove(f)
        os.rmdir(tmp_dir)
        return jsonify({'error': f'Error procesando: {str(e)}'}), 500

    # Limpiar archivos de entrada
    os.remove(ruta1)
    os.remove(ruta2)

    return send_file(
        ruta_salida,
        as_attachment=True,
        download_name=nombre_salida,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
