from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
from datetime import datetime
import pytz
import traceback

from procesar_pdf import (
    procesar_y_generar_excel,
    get_dashboard_data,
    registrar_correccion_nombre,
    registrar_correccion_asunto
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

def fecha_hoy_peru():
    tz = pytz.timezone("America/Lima")
    return datetime.now(tz).replace(tzinfo=None)

@app.route('/')
def index():
    return render_template('index.html')

# Generar Excel y descargar
@app.route('/procesar_pdf', methods=['POST'])
def procesar_pdf_route():
    try:
        if 'pdf_file' not in request.files:
            return jsonify({"error": "No se envió ningún PDF"}), 400
        pdf_file = request.files['pdf_file']
        pdf_bytes = pdf_file.read()
        fecha_peru = fecha_hoy_peru()
        excel_output = procesar_y_generar_excel(pdf_bytes, fecha_peru)
        return send_file(
            excel_output,
            as_attachment=True,
            download_name="Reporte_OIST.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# Mostrar dashboard (procesa y renderiza)
@app.route('/dashboard', methods=['POST'])
def dashboard():
    try:
        if 'pdf_file' not in request.files:
            return redirect(url_for('index'))
        pdf_file = request.files['pdf_file']
        pdf_bytes = pdf_file.read()
        fecha_peru = fecha_hoy_peru()
        df_dashboard, date_range, office_name = get_dashboard_data(pdf_bytes, fecha_peru)

        # Prepare structures for template
        filas = df_dashboard.to_dict('records')
        # resumen global
        resumen_global = {
            "por_recibir": int(df_dashboard['POR RECIBIR'].sum()) if 'POR RECIBIR' in df_dashboard.columns else int((df_dashboard['Estado']=='POR RECIBIR').sum()),
            "pendiente": int(df_dashboard['PENDIENTE'].sum()) if 'PENDIENTE' in df_dashboard.columns else int((df_dashboard['Estado']=='PENDIENTE').sum()),
            "total": int(len(df_dashboard)),
            "max_dias": int(df_dashboard['Dias_En_Bandeja'].max()) if not df_dashboard.empty else 0
        }

        # personas: aggregate per Nombre_Personal
        personas = []
        grafico_nombres = []
        grafico_pendientes = []
        grafico_por_recibir = []

        grouped = df_dashboard.groupby('Nombre_Personal')
        for nombre, group in grouped:
            por_recibir = int(group['Estado'].eq('POR RECIBIR').sum())
            pendiente = int(group['Estado'].eq('PENDIENTE').sum())
            total = por_recibir + pendiente
            max_dias = int(group['Dias_En_Bandeja'].max()) if not group.empty else 0

            # documento con mayor tiempo (asunto limpio)
            doc_max_row = group.sort_values(by='Dias_En_Bandeja', ascending=False).iloc[0]
            doc_max = {
                "asunto": doc_max_row['Asunto'],
                "fecha": doc_max_row['Fecha_Recepcion_Str'],
                "codigo": doc_max_row['Codigo'],
                "dias": int(doc_max_row['Dias_En_Bandeja'])
            }

            # totales por fecha
            tot_por_fecha = group.groupby(group['Fecha_Recepcion'].dt.strftime('%d/%m/%Y'))['Codigo'].count().reset_index()
            totales_fecha = [{"fecha": r[0], "cantidad": int(r[1])} for r in tot_por_fecha.values]

            personas.append({
                "nombre": nombre,
                "por_recibir": por_recibir,
                "pendiente": pendiente,
                "total": total,
                "max_dias": max_dias,
                "doc_max": doc_max,
                "totales_fecha": totales_fecha
            })

            grafico_nombres.append(nombre)
            grafico_pendientes.append(pendiente)
            grafico_por_recibir.append(por_recibir)

        grafico = {
            "nombres": grafico_nombres,
            "pendientes": grafico_pendientes,
            "por_recibir": grafico_por_recibir
        }

        return render_template(
            'dashboard.html',
            filas=filas,
            date_range=date_range,
            office_name=office_name,
            resumen_global=resumen_global,
            personas=personas,
            grafico=grafico,
            tabla_detallada=filas
        )
    except Exception as e:
        traceback.print_exc()
        return f"<h3>Error procesando PDF: {str(e)}</h3>"

# Corrección de nombre (aprendizaje)
@app.route('/corregir_nombre', methods=['POST'])
def corregir_nombre():
    data = request.json
    original = data.get("original")
    corregido = data.get("corregido")
    if not original or not corregido:
        return jsonify({"error":"datos incompletos"}), 400
    registrar_correccion_nombre(original, corregido)
    return jsonify({"mensaje": "Nombre actualizado correctamente"})

# Corrección de asunto (aprendizaje)
@app.route('/corregir_asunto', methods=['POST'])
def corregir_asunto():
    data = request.json
    original = data.get("original")
    corregido = data.get("corregido")
    if not original or not corregido:
        return jsonify({"error":"datos incompletos"}), 400
    registrar_correccion_asunto(original, corregido)
    return jsonify({"mensaje": "Asunto actualizado correctamente"})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
