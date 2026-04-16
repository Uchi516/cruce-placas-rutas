"""
Vercel serverless handler for the Excel processing app.
Multi-process platform with tabs.
"""

import os
import sys
import io
import tempfile
import zipfile
from flask import Flask, request, send_file, jsonify, Response

# Add parent dir to path so we can import procesar_excel
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from procesar_excel import procesar

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

HTML_PAGE = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Procesos Excel</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #0f172a; color: #e2e8f0;
            min-height: 100vh;
        }

        /* --- NAV TABS --- */
        .nav {
            display: flex; align-items: center;
            background: #1e293b; border-bottom: 1px solid #334155;
            padding: 0 24px; gap: 4px;
        }
        .nav-brand {
            font-size: 14px; font-weight: 700; color: #94a3b8;
            margin-right: 24px; padding: 12px 0;
            letter-spacing: 0.5px;
        }
        .nav-tab {
            padding: 12px 18px; font-size: 13px; font-weight: 500;
            color: #64748b; cursor: pointer; border: none; background: none;
            border-bottom: 2px solid transparent; transition: all 0.2s;
        }
        .nav-tab:hover { color: #e2e8f0; }
        .nav-tab.active { color: #3b82f6; border-bottom-color: #3b82f6; }

        /* --- MAIN --- */
        .main { display: flex; justify-content: center; padding: 20px; }
        .tab-content { display: none; width: 100%; max-width: 580px; }
        .tab-content.active { display: block; }

        /* --- CARD --- */
        .card {
            background: #1e293b; border-radius: 14px; padding: 28px 32px;
            box-shadow: 0 25px 50px rgba(0,0,0,0.4); border: 1px solid #334155;
        }
        .card-title { text-align: center; margin-bottom: 20px; }
        .card-title h2 { font-size: 20px; font-weight: 700; color: #f8fafc; margin-bottom: 2px; }
        .card-title p { color: #94a3b8; font-size: 13px; }

        /* --- UPLOAD --- */
        .upload-section { margin-bottom: 16px; }
        .upload-section label {
            display: block; font-size: 12px; font-weight: 600; color: #94a3b8;
            text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;
        }
        .upload-area {
            position: relative; border: 2px dashed #334155; border-radius: 10px;
            padding: 16px; text-align: center; cursor: pointer;
            transition: all 0.2s; background: #0f172a;
        }
        .upload-area:hover { border-color: #3b82f6; background: rgba(59,130,246,0.05); }
        .upload-area.has-file { border-color: #22c55e; background: rgba(34,197,94,0.05); }
        .upload-area input[type="file"] {
            position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            opacity: 0; cursor: pointer;
        }
        .upload-icon { font-size: 24px; margin-bottom: 4px; }
        .upload-text { color: #64748b; font-size: 13px; }
        .upload-text .filename { color: #22c55e; font-weight: 600; }

        /* --- FORM FIELDS --- */
        .field { margin-bottom: 14px; }
        .field label {
            display: block; font-size: 12px; font-weight: 600; color: #94a3b8;
            text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px;
        }
        .field input[type="text"], .field textarea {
            width: 100%; background: #0f172a; border: 1px solid #334155;
            border-radius: 10px; padding: 10px 14px; color: #e2e8f0;
            font-size: 13px; font-family: inherit; transition: border-color 0.2s;
            outline: none;
        }
        .field input[type="text"]:focus, .field textarea:focus { border-color: #3b82f6; }
        .field textarea { min-height: 110px; resize: vertical; line-height: 1.5; }
        .field .hint { color: #475569; font-size: 11px; margin-top: 4px; }

        /* --- FILE LIST --- */
        .file-list { margin-top: 8px; }
        .file-item {
            display: flex; align-items: center; justify-content: space-between;
            background: #0f172a; border: 1px solid #334155; border-radius: 8px;
            padding: 8px 12px; margin-bottom: 6px; font-size: 12px;
        }
        .file-item .fname { color: #22c55e; font-weight: 600; }
        .file-item .fsize { color: #64748b; }
        .file-item .remove-btn {
            background: none; border: none; color: #ef4444; cursor: pointer;
            font-size: 16px; padding: 0 4px; line-height: 1;
        }
        .file-item .remove-btn:hover { color: #f87171; }

        /* --- BUTTONS --- */
        .btn {
            width: 100%; padding: 13px; border: none; border-radius: 10px;
            font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; margin-top: 6px;
        }
        .btn-primary { background: #3b82f6; color: white; }
        .btn-primary:hover {
            background: #2563eb; transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(59,130,246,0.4);
        }
        .btn-primary:disabled {
            background: #334155; color: #64748b; cursor: not-allowed;
            transform: none; box-shadow: none;
        }
        .btn-secondary { background: #334155; color: #e2e8f0; }
        .btn-secondary:hover { background: #475569; transform: translateY(-1px); }

        /* --- STATUS --- */
        .status {
            margin-top: 14px; padding: 12px; border-radius: 10px;
            font-size: 13px; display: none;
        }
        .status.processing {
            display: block; background: rgba(59,130,246,0.1);
            border: 1px solid rgba(59,130,246,0.3); color: #93c5fd;
        }
        .status.success {
            display: block; background: rgba(34,197,94,0.1);
            border: 1px solid rgba(34,197,94,0.3); color: #86efac;
        }
        .status.error {
            display: block; background: rgba(239,68,68,0.1);
            border: 1px solid rgba(239,68,68,0.3); color: #fca5a5;
        }
        .spinner {
            display: inline-block; width: 16px; height: 16px;
            border: 2px solid rgba(147,197,253,0.3); border-top-color: #93c5fd;
            border-radius: 50%; animation: spin 0.8s linear infinite;
            vertical-align: middle; margin-right: 8px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        /* --- STEPS --- */
        .steps { display: flex; justify-content: center; gap: 28px; margin-bottom: 20px; }
        .step { display: flex; align-items: center; gap: 8px; font-size: 13px; color: #64748b; }
        .step-num {
            width: 24px; height: 24px; border-radius: 50%; background: #334155;
            display: flex; align-items: center; justify-content: center;
            font-size: 12px; font-weight: 700;
        }
        .step.active .step-num { background: #3b82f6; color: white; }
        .step.done .step-num { background: #22c55e; color: white; }

        /* --- VERSION TAG --- */
        .version-tag {
            position: fixed; bottom: 10px; right: 14px;
            font-size: 10px; color: #475569; line-height: 1.4;
            text-align: right; pointer-events: none;
        }
    </style>
</head>
<body>
    <!-- NAV -->
    <nav class="nav">
        <span class="nav-brand">PROCESOS EXCEL</span>
        <button class="nav-tab active" onclick="switchTab('cruce')">Cruce Placas Rutas</button>
        <button class="nav-tab" onclick="switchTab('nuevo')">Nuevo Proceso</button>
    </nav>

    <div class="main">
        <!-- TAB 1: CRUCE PLACAS RUTAS -->
        <div class="tab-content active" id="tab-cruce">
            <div class="card">
                <div class="card-title">
                    <h2>Cruce Placas Rutas</h2>
                    <p>Picking Center - Procesamiento de Rutas</p>
                </div>
                <div class="steps">
                    <div class="step active" id="step1"><span class="step-num">1</span><span>Subir</span></div>
                    <div class="step" id="step2"><span class="step-num">2</span><span>Procesar</span></div>
                    <div class="step" id="step3"><span class="step-num">3</span><span>Descargar</span></div>
                </div>
                <form id="uploadForm">
                    <div class="upload-section">
                        <label>Archivo 1 - Programacion de Placas</label>
                        <div class="upload-area" id="area1">
                            <input type="file" name="archivo1" id="archivo1" accept=".xlsx">
                            <div class="upload-icon">&#128196;</div>
                            <div class="upload-text" id="text1">Arrastra o haz clic para seleccionar</div>
                        </div>
                    </div>
                    <div class="upload-section">
                        <label>Archivo 2 - Programacion PC</label>
                        <div class="upload-area" id="area2">
                            <input type="file" name="archivo2" id="archivo2" accept=".xlsx">
                            <div class="upload-icon">&#128196;</div>
                            <div class="upload-text" id="text2">Arrastra o haz clic para seleccionar</div>
                        </div>
                    </div>
                    <button type="submit" class="btn btn-primary" id="btnProcesar" disabled>
                        Procesar Archivos
                    </button>
                </form>
                <div class="status" id="status"></div>
            </div>
        </div>

        <!-- TAB 2: NUEVO PROCESO -->
        <div class="tab-content" id="tab-nuevo">
            <div class="card">
                <div class="card-title">
                    <h2>Nuevo Proceso</h2>
                    <p>Describe el proceso y sube archivos de ejemplo</p>
                </div>
                <form id="nuevoForm">
                    <div class="field">
                        <label>Nombre del proceso</label>
                        <input type="text" id="nombreProceso" placeholder="Ej: Consolidado de despachos, Reporte cobertura...">
                    </div>
                    <div class="field">
                        <label>Explicacion del proceso</label>
                        <textarea id="explicacion" placeholder="Describe paso a paso que hace este proceso:&#10;&#10;- Que archivos se usan como entrada?&#10;- Que columnas o datos son importantes?&#10;- Que resultado se espera?&#10;- Alguna regla o condicion especial?&#10;&#10;Mientras mas detalle, mejor..."></textarea>
                        <div class="hint">Escribe todo lo que necesites. Puedes incluir ejemplos, reglas, excepciones, etc.</div>
                    </div>
                    <div class="field">
                        <label>Archivos de ejemplo (opcional)</label>
                        <div class="upload-area" id="areaNuevo">
                            <input type="file" id="archivosEjemplo" multiple accept=".xlsx,.xls,.csv,.pdf,.txt,.png,.jpg,.jpeg">
                            <div class="upload-icon">&#128206;</div>
                            <div class="upload-text" id="textNuevo">Arrastra o haz clic para agregar archivos</div>
                        </div>
                        <div class="hint">Sube archivos de entrada de ejemplo y/o el resultado esperado. Acepta xlsx, csv, pdf, txt, imagenes.</div>
                        <div class="file-list" id="fileList"></div>
                    </div>
                    <button type="submit" class="btn btn-primary" id="btnNuevo" disabled>
                        Descargar Paquete del Proceso
                    </button>
                    <div class="hint" style="text-align:center; margin-top: 8px;">
                        Se descarga un ZIP con la explicacion + archivos para procesarlo con IA
                    </div>
                </form>
                <div class="status" id="statusNuevo"></div>
            </div>
        </div>
    </div>

    <script>
        // === TAB SWITCHING ===
        function switchTab(tab) {
            document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
            document.getElementById('tab-' + tab).classList.add('active');
            event.target.classList.add('active');
        }

        // === TAB 1: CRUCE PLACAS RUTAS ===
        const archivo1 = document.getElementById('archivo1');
        const archivo2 = document.getElementById('archivo2');
        const btnProcesar = document.getElementById('btnProcesar');
        const status = document.getElementById('status');
        const form = document.getElementById('uploadForm');

        function updateFileDisplay(input, areaId, textId) {
            const area = document.getElementById(areaId);
            const text = document.getElementById(textId);
            if (input.files.length > 0) {
                area.classList.add('has-file');
                text.innerHTML = '<span class="filename">' + input.files[0].name + '</span>';
            } else {
                area.classList.remove('has-file');
                text.textContent = 'Arrastra o haz clic para seleccionar';
            }
            checkReady();
        }

        function checkReady() {
            btnProcesar.disabled = !(archivo1.files.length > 0 && archivo2.files.length > 0);
        }

        archivo1.addEventListener('change', () => updateFileDisplay(archivo1, 'area1', 'text1'));
        archivo2.addEventListener('change', () => updateFileDisplay(archivo2, 'area2', 'text2'));

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            document.getElementById('step1').classList.remove('active');
            document.getElementById('step1').classList.add('done');
            document.getElementById('step2').classList.add('active');
            btnProcesar.disabled = true;
            btnProcesar.textContent = 'Procesando...';
            status.className = 'status processing';
            status.innerHTML = '<span class="spinner"></span> Procesando archivos, espera un momento...';

            const formData = new FormData();
            formData.append('archivo1', archivo1.files[0]);
            formData.append('archivo2', archivo2.files[0]);

            try {
                const response = await fetch('/api/procesar', { method: 'POST', body: formData });
                if (!response.ok) {
                    let msg = 'Error del servidor (HTTP ' + response.status + ')';
                    try { const data = await response.json(); msg = data.error || msg; } catch(e) {}
                    throw new Error(msg);
                }
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                const disposition = response.headers.get('Content-Disposition');
                let filename = 'resultado.xlsx';
                if (disposition) {
                    const match = disposition.match(/filename\\*?=(?:UTF-8''|"?)([^";]+)/);
                    if (match) filename = decodeURIComponent(match[1].replace(/"/g, ''));
                }
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                document.getElementById('step2').classList.remove('active');
                document.getElementById('step2').classList.add('done');
                document.getElementById('step3').classList.add('done');
                status.className = 'status success';
                status.textContent = 'Archivo procesado y descargado correctamente.';
            } catch (err) {
                document.getElementById('step2').classList.remove('active');
                status.className = 'status error';
                status.textContent = 'Error: ' + err.message;
            }
            btnProcesar.textContent = 'Procesar Archivos';
            btnProcesar.disabled = false;
        });

        // === TAB 2: NUEVO PROCESO ===
        const nombreProceso = document.getElementById('nombreProceso');
        const explicacion = document.getElementById('explicacion');
        const archivosEjemplo = document.getElementById('archivosEjemplo');
        const btnNuevo = document.getElementById('btnNuevo');
        const statusNuevo = document.getElementById('statusNuevo');
        const fileList = document.getElementById('fileList');
        const areaNuevo = document.getElementById('areaNuevo');
        let archivosAgregados = [];

        function checkNuevoReady() {
            btnNuevo.disabled = !(nombreProceso.value.trim() && explicacion.value.trim());
        }

        nombreProceso.addEventListener('input', checkNuevoReady);
        explicacion.addEventListener('input', checkNuevoReady);

        archivosEjemplo.addEventListener('change', () => {
            for (const f of archivosEjemplo.files) {
                if (!archivosAgregados.some(a => a.name === f.name && a.size === f.size)) {
                    archivosAgregados.push(f);
                }
            }
            renderFileList();
            archivosEjemplo.value = '';
        });

        function renderFileList() {
            if (archivosAgregados.length === 0) {
                fileList.innerHTML = '';
                areaNuevo.classList.remove('has-file');
                document.getElementById('textNuevo').textContent = 'Arrastra o haz clic para agregar archivos';
                return;
            }
            areaNuevo.classList.add('has-file');
            document.getElementById('textNuevo').innerHTML = '<span class="filename">' + archivosAgregados.length + ' archivo(s) agregado(s)</span>';
            fileList.innerHTML = archivosAgregados.map((f, i) =>
                '<div class="file-item">' +
                    '<span class="fname">' + f.name + '</span>' +
                    '<span class="fsize">' + (f.size / 1024).toFixed(1) + ' KB</span>' +
                    '<button type="button" class="remove-btn" onclick="removeFile(' + i + ')">&#10005;</button>' +
                '</div>'
            ).join('');
        }

        function removeFile(idx) {
            archivosAgregados.splice(idx, 1);
            renderFileList();
        }

        document.getElementById('nuevoForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            btnNuevo.disabled = true;
            btnNuevo.textContent = 'Generando paquete...';
            statusNuevo.className = 'status processing';
            statusNuevo.innerHTML = '<span class="spinner"></span> Creando ZIP con la explicacion y archivos...';

            const formData = new FormData();
            formData.append('nombre', nombreProceso.value.trim());
            formData.append('explicacion', explicacion.value.trim());
            for (const f of archivosAgregados) {
                formData.append('archivos', f);
            }

            try {
                const response = await fetch('/api/nuevo-proceso', { method: 'POST', body: formData });
                if (!response.ok) {
                    let msg = 'Error del servidor (HTTP ' + response.status + ')';
                    try { const data = await response.json(); msg = data.error || msg; } catch(e) {}
                    throw new Error(msg);
                }
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = nombreProceso.value.trim().replace(/\\s+/g, '-').toLowerCase() + '.zip';
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                statusNuevo.className = 'status success';
                statusNuevo.textContent = 'Paquete descargado correctamente.';
            } catch (err) {
                statusNuevo.className = 'status error';
                statusNuevo.textContent = 'Error: ' + err.message;
            }
            btnNuevo.textContent = 'Descargar Paquete del Proceso';
            checkNuevoReady();
        });
    </script>
    <div class="version-tag">v1.3 | Actualizado: 16/04/2026 07:55</div>
</body>
</html>"""


@app.route('/')
@app.route('/index')
def index():
    return Response(HTML_PAGE, mimetype='text/html')


@app.route('/api/procesar', methods=['POST'])
def procesar_archivos():
    if 'archivo1' not in request.files or 'archivo2' not in request.files:
        return jsonify({'error': 'Debes subir ambos archivos'}), 400

    archivo1 = request.files['archivo1']
    archivo2 = request.files['archivo2']

    if archivo1.filename == '' or archivo2.filename == '':
        return jsonify({'error': 'Debes seleccionar ambos archivos'}), 400

    if not archivo1.filename.endswith('.xlsx') or not archivo2.filename.endswith('.xlsx'):
        return jsonify({'error': 'Los archivos deben ser .xlsx'}), 400

    tmp_dir = tempfile.mkdtemp()
    ruta1 = os.path.join(tmp_dir, archivo1.filename)
    ruta2 = os.path.join(tmp_dir, archivo2.filename)
    archivo1.save(ruta1)
    archivo2.save(ruta2)

    base, ext = os.path.splitext(archivo2.filename)
    nombre_salida = f"{base} - RESULTADO{ext}"
    ruta_salida = os.path.join(tmp_dir, nombre_salida)

    try:
        procesar(ruta1, ruta2, ruta_salida)
    except BaseException as e:
        for f in [ruta1, ruta2]:
            if os.path.exists(f):
                os.remove(f)
        try:
            os.rmdir(tmp_dir)
        except OSError:
            pass
        return jsonify({'error': f'Error procesando: {str(e)}'}), 500

    os.remove(ruta1)
    os.remove(ruta2)

    return send_file(
        ruta_salida,
        as_attachment=True,
        download_name=nombre_salida,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/api/nuevo-proceso', methods=['POST'])
def nuevo_proceso():
    nombre = request.form.get('nombre', '').strip()
    explicacion = request.form.get('explicacion', '').strip()

    if not nombre or not explicacion:
        return jsonify({'error': 'Nombre y explicacion son obligatorios'}), 400

    # Build ZIP in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Add the explanation as a text file
        contenido = f"PROCESO: {nombre}\n{'=' * 60}\n\n{explicacion}\n"
        zf.writestr('explicacion.txt', contenido)

        # Add uploaded files
        archivos = request.files.getlist('archivos')
        for archivo in archivos:
            if archivo.filename:
                zf.writestr(f"archivos/{archivo.filename}", archivo.read())

    zip_buffer.seek(0)
    nombre_zip = nombre.replace(' ', '-').lower() + '.zip'

    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=nombre_zip,
        mimetype='application/zip'
    )


# For local development
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
