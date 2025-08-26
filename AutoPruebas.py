from flask import Flask, request, send_file, render_template_string, url_for
from excel_generator import generate_excel_report
from datetime import datetime
import json
import openpyxl
import os

app = Flask(__name__)

HTML_FORM = """
<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Generador de Plantilla de Pruebas</title>
  <link rel="stylesheet" href="{{ url_for('static', filename='style.css') }}">
</head>
<body>
  <!-- Formulario de generación y carga -->
  <form id="form-generate" method="post" action="/generate" enctype="multipart/form-data">
    <div style="display: flex; gap: 1rem;">
      <!-- Columna izquierda: Datos del Proyecto -->
      <div class="section" style="flex: 1;">
        <h2>Datos del Proyecto</h2>
        <label>Nombre del Proyecto</label>
        <input type="text" name="nombre" pattern="[A-Za-z0-9 ]+" required value="{{ nombre|default('') }}">
        <label>Código del Proyecto</label>
        <input type="text" name="codigo" pattern="[A-Za-z0-9.]+" required title="Formato: letras, números y puntos (ej: S.25.20.20)" value="{{ codigo|default('') }}">
        <label>Código Wrike (URL)</label>
        <input type="url" name="wrike" required value="{{ wrike|default('') }}">
        <label>Versión</label>
        <input type="text" name="version" pattern="[A-Za-z0-9.]+" required value="{{ version|default('') }}">
        <label>Fecha del Proyecto</label>
        <input type="date" name="fecha_proyecto" required value="{{ fecha_proyecto|default('') }}">
        <label>Módulo</label>
        <input type="text" name="modulo" required value="{{ modulo|default('') }}">
        <label>Fecha Planeada para Ejecución</label>
        <input type="date" name="fecha_planeada" required value="{{ fecha_planeada|default('') }}">
      </div>
      <!-- Columna derecha: carga y casos -->
      <div class="section" style="flex: 1; display: flex; flex-direction: column; gap: 1rem;">
        <!-- Zona de carga de Excel -->
        <div>
          <h2>Carga de Pruebas</h2>
          <label>Subir listado de pruebas (Excel)</label>
          <input type="file" name="test_file" accept=".xlsx">
          <button type="submit" class="btn" formaction="/load" formmethod="post">Cargar Pruebas</button>
        </div>
        <!-- Zona de casos -->
        <div>
          <h3>Casos de Prueba</h3>
          <div id="tests" class="tests-container"></div>
          <button type="button" class="btn" onclick="openModal()">Agregar Prueba</button>
          <button type="button" class="btn btn-danger" onclick="clearAll()">Limpiar</button>
        </div>
      </div>
    </div>
    <!-- Botón Generar al final -->
    <div class="section" style="text-align:center; margin-top:1rem;">
      <input type="submit" class="btn" value="Generar Excel">
    </div>
  </form>

  <!-- Modal para CRUD de pruebas -->
  <div id="modal" class="modal-overlay">
    <div class="modal">
      <h4 id="modal-title">Nuevo Caso de Prueba</h4>
      <label>Escenario</label><input id="m_caso" type="text">
      <label>Descripción/Objetivo</label><input id="m_desc" type="text">
      <label>Precondiciones</label><input id="m_pre" type="text">
      <label>Postcondiciones</label><input id="m_post" type="text">
      <label>Prioridad</label>
      <select id="m_prio"><option>Alta</option><option>Media</option><option>Baja</option></select>
      <label>Criterio de Aceptación</label>
      <select id="m_crit"><option>Pendiente</option><option>No aplica</option><option>Aprobado</option><option>No Aprobado</option></select>
      <label>Comentarios</label><input id="m_com" type="text">
      <div style="margin-top:15px;text-align:right;">
        <button type="button" class="btn" onclick="saveTest()">Guardar</button>
        <button type="button" class="btn close-btn" onclick="closeModal()">Cancelar</button>
      </div>
    </div>
  </div>

  <script>
    // Inicializar array de pruebas
    const initialTests = {{ initial_tests|default('[]')|safe }};
    let tests = [];
    if (initialTests.length) {
      tests = initialTests;
      renderTests();
    }
    let editIndex = null;

    function openModal(idx=null) {
      editIndex = idx;
      document.getElementById('modal-title').innerText = idx===null ? 'Nuevo Caso de Prueba' : 'Editar Caso de Prueba';
      if (idx!==null) {
        const t = tests[idx];
        ['m_caso','m_desc','m_pre','m_post','m_com'].forEach((id,i) => {
          document.getElementById(id).value = [t.caso,t.descripcion,t.precondiciones,t.postcondiciones,t.comentarios][i];
        });
        document.getElementById('m_prio').value = t.prioridad;
        document.getElementById('m_crit').value = t.criterio;
      } else {
        ['m_caso','m_desc','m_pre','m_post','m_com'].forEach(id=>document.getElementById(id).value='');
      }
      document.getElementById('modal').style.display='flex';
    }
    function closeModal(){document.getElementById('modal').style.display='none';}
    function saveTest(){
      const casoVal=document.getElementById('m_caso').value.trim();
      if(!casoVal){alert('Escenario es obligatorio');return;}      
      const t={caso:casoVal,descripcion:document.getElementById('m_desc').value.trim(),precondiciones:document.getElementById('m_pre').value.trim(),postcondiciones:document.getElementById('m_post').value.trim(),prioridad:document.getElementById('m_prio').value,criterio:document.getElementById('m_crit').value,comentarios:document.getElementById('m_com').value.trim()};
      if(editIndex===null) tests.push(t); else tests[editIndex]=t;
      renderTests(); closeModal();
    }
    function deleteTest(i){tests.splice(i,1);renderTests();}
    function renderTests(){
      const c=document.getElementById('tests'); c.innerHTML='';
      tests.forEach((t,i)=>{
        const card=document.createElement('div');card.className='card';
        card.innerHTML=`<h4>Prueba #${i+1}: ${t.caso}</h4><div class="card-buttons"><button type="button" class="btn" onclick="openModal(${i})">Ver más</button><button type="button" class="btn close-btn" onclick="deleteTest(${i})">Eliminar</button></div>`;
        ['caso','descripcion','precondiciones','postcondiciones','prioridad','criterio','comentarios'].forEach(k=>{const inp=document.createElement('input');inp.type='hidden';inp.name=`${k}_${i}`;inp.value=t[k];card.appendChild(inp);});
        c.appendChild(card);
      });
    }
    function clearAll(){
      // Limpiar las tarjetas
      tests = [];
      renderTests();
      // Resetear sólo los campos de datos del proyecto
      document.getElementById('form-generate').reset();
    }
  </script>
</body>
</html>
"""

@app.route('/')
def form():
    return render_template_string(HTML_FORM, initial_tests='[]', nombre='', codigo='', wrike='', version='', fecha_proyecto='', modulo='', fecha_planeada='')

@app.route('/load', methods=['POST'])
def load_tests():
    # Leer Excel y poblar tests
    file=request.files['test_file']
    wb=openpyxl.load_workbook(file, data_only=True)
    ws=wb.active
    tests=[]
    for row in ws.iter_rows(min_row=2, values_only=True):
        caso,descripcion,pre,post,prio,crit,comm=row
        tests.append({'caso':caso or '','descripcion':descripcion or '','precondiciones':pre or '','postcondiciones':post or '','prioridad':prio or '','criterio':crit or '','comentarios':comm or ''})
    # Mantener proyecto en campos
    return render_template_string(HTML_FORM, initial_tests=json.dumps(tests), nombre=request.form.get('nombre',''), codigo=request.form.get('codigo',''), wrike=request.form.get('wrike',''), version=request.form.get('version',''), fecha_proyecto=request.form.get('fecha_proyecto',''), modulo=request.form.get('modulo',''), fecha_planeada=request.form.get('fecha_planeada',''))

@app.route('/generate', methods=['POST'])
def generate():
    project_code=request.form['codigo']
    data={'nombre':request.form['nombre'],'codigo':project_code,'wrike':request.form['wrike'],'version':request.form['version'],'fecha_proyecto':request.form['fecha_proyecto'],'modulo':request.form['modulo'],'fecha_planeada':request.form['fecha_planeada'],'casos':[]}
    i=0
    while True:
        key=f"caso_{i}"
        if key not in request.form: break
        data['casos'].append({'codigo':f"{project_code}.{i}",'caso':request.form[key],'descripcion':request.form.get(f"descripcion_{i}",''),'precondiciones':request.form.get(f"precondiciones_{i}",''),'postcondiciones':request.form.get(f"postcondiciones_{i}",''),'prioridad':request.form.get(f"prioridad_{i}",''),'criterio':request.form.get(f"criterio_{i}",''),'comentarios':request.form.get(f"comentarios_{i}",'')})
        i+=1
    buf=generate_excel_report(data)
    filename=f"plantilla_pruebas_{project_code}.xlsx"
    return send_file(buf, as_attachment=True, download_name=filename)

if __name__=='__main__':
    app.run(debug=True)
