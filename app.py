from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
import os
import uuid
import subprocess
import gspread
from google.oauth2.service_account import Credentials
import json 
from docxtpl import DocxTemplate
import tempfile
import shutil
from datetime import datetime
import locale

app = Flask(__name__)
CORS(app)

def numero_a_texto(numero):
    try:
        from num2words import num2words
        return num2words(numero, lang='es').capitalize() + " pesos colombianos"
    except:
        return str(numero)

def formatear_moneda(valor):
    try:
        numero = int(str(valor).replace('.', '').replace(',', '').replace('$', '').replace(' ', ''))
        return "{:,}".format(numero).replace(",", ".")
    except:
        return valor if valor is not None else ''

def guardar_en_google_sheets(data):
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    creds_dict = json.loads(os.environ.get("GOOGLE_CREDENTIALS"))
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1qWXMepGrgxjZK9QLPxcLcCDlHtsLBIH9Fo0GJDgr_Go").sheet1  
    fila = [
        data.get("nombre", ""),
        data.get("telefono", ""),
        data.get("correo", ""),
        data.get("diseno_arquitectonico", ""),
        data.get("diseno_estructural", ""),
        data.get("acompanamiento_licencias", ""),
        data.get("subtotal_etapa_1", ""),
        data.get("diseno_electrico", ""),
        data.get("diseno_hidraulico", ""),
        data.get("presupuesto_proyecto", ""),
        data.get("subtotal_etapa_2", ""),
        data.get("total_general", ""),
        data.get("total_general_texto", ""),
        data.get("costo_construccion", "")
    ]
    sheet.append_row(fila)

def convertir_word_a_pdf_libreoffice(docx_path):
    """
    Convierte un archivo Word a PDF usando LibreOffice
    """
    try:
        # Crear directorio temporal para la conversión
        temp_dir = tempfile.mkdtemp()
        
        # Comando para convertir con LibreOffice
        comando = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', temp_dir,
            docx_path
        ]
        
        # Ejecutar conversión
        resultado = subprocess.run(comando, capture_output=True, text=True, timeout=60)
        
        if resultado.returncode != 0:
            raise Exception(f"Error en LibreOffice: {resultado.stderr}")
        
        # Buscar el archivo PDF generado
        nombre_base = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(temp_dir, f"{nombre_base}.pdf")
        
        if not os.path.exists(pdf_path):
            raise Exception("No se generó el archivo PDF")
        
        return pdf_path, temp_dir
        
    except subprocess.TimeoutExpired:
        raise Exception("Timeout en la conversión PDF")
    except Exception as e:
        raise Exception(f"Error al convertir a PDF: {str(e)}")

def convertir_word_a_pdf_pandoc(docx_path):
    """
    Convierte un archivo Word a PDF usando pandoc
    """
    try:
        temp_dir = tempfile.mkdtemp()
        nombre_base = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(temp_dir, f"{nombre_base}.pdf")
        
        comando = [
            'pandoc',
            docx_path,
            '-o', pdf_path,
            '--pdf-engine=xelatex'
        ]
        
        resultado = subprocess.run(comando, capture_output=True, text=True, timeout=30)
        
        if resultado.returncode != 0:
            raise Exception(f"Error en pandoc: {resultado.stderr}")
        
        return pdf_path, temp_dir
        
    except Exception as e:
        raise Exception(f"Error al convertir con pandoc: {str(e)}")

def detectar_herramienta_conversion():
    """
    Detecta qué herramienta de conversión está disponible
    """
    # Probar LibreOffice
    try:
        resultado = subprocess.run(['libreoffice', '--version'], 
                                 capture_output=True, text=True, timeout=5)
        if resultado.returncode == 0:
            return 'libreoffice'
    except:
        pass
    
    # Probar pandoc
    try:
        resultado = subprocess.run(['pandoc', '--version'], 
                                 capture_output=True, text=True, timeout=5)
        if resultado.returncode == 0:
            return 'pandoc'
    except:
        pass
    
    return None

@app.route('/generar-documento', methods=['POST'])
def generar_documento():
    """
    Genera documento en formato Word y/o PDF usando los datos recibidos tal cual, sin cálculos ni valores por defecto. Ahora espera un JSON plano.
    """
    data = request.get_json()
    formato = data.get('formato', 'word').lower()  # 'word', 'pdf', 'ambos'

    # Mapear los datos a los nombres de variables usados en la plantilla (igual que en la imagen)
    contexto = {
        'nombre': data.get('nombre', ''),
        'telefono': data.get('telefono', ''),
        'correo': data.get('correo', ''),
        'diseno_arquitectonico': formatear_moneda(data.get('diseno_arquitectonico', data.get('Diseño_Arquitectonico', ''))),
        'diseno_estructural': formatear_moneda(data.get('diseno_estructural', data.get('Diseño_Estructural', ''))),
        'acompanamiento_licencias': formatear_moneda(data.get('acompanamiento_licencias', data.get('Acompañamiento_Licencias', ''))),
        'subtotal_etapa_1': formatear_moneda(data.get('subtotal_etapa_1', data.get('Subtotal_Etapa_I', ''))),
        'diseno_electrico': formatear_moneda(data.get('diseno_electrico', data.get('Diseño_Electrico', ''))),
        'diseno_hidraulico': formatear_moneda(data.get('diseno_hidraulico', data.get('Diseño_Hidraulico', ''))),
        'presupuesto_proyecto': formatear_moneda(data.get('presupuesto_proyecto', data.get('Presupuesto_Proyecto', ''))),
        'subtotal_etapa_2': formatear_moneda(data.get('subtotal_etapa_2', data.get('Subtotal_Etapa_II', ''))),
        'total_general': formatear_moneda(data.get('total_general', data.get('Total_General', ''))),
        'total_general_texto': data.get('total_general_texto', data.get('Total_General_Texto', '')),
        'costo_construccion': formatear_moneda(data.get('costo_construccion', data.get('Costo_Construccion', ''))),
    }

    # Guardar en Google Sheets (opcional, puedes comentar si no lo usas)
    try:
        guardar_en_google_sheets(contexto)
    except Exception as e:
        print(f"Error al guardar en Google Sheets: {e}")

    # Verificar plantilla
    plantilla_path = os.path.join(os.path.dirname(__file__), "Formato.docx")
    if not os.path.exists(plantilla_path):
        return jsonify({"error": "No se encontró la plantilla Word"}), 500

    # Generar documento Word
    doc = DocxTemplate(plantilla_path)
    doc.render(contexto)

    unique_id = str(uuid.uuid4())
    temp_dir = tempfile.mkdtemp()
    docx_path = os.path.join(temp_dir, f"cotizacion_{unique_id}.docx")
    doc.save(docx_path)

    archivos_a_limpiar = [temp_dir]

    try:
        if formato == 'word':
            # Solo Word
            response = send_file(docx_path, 
                               as_attachment=True, 
                               download_name=f"cotizacion_{unique_id}.docx")
        elif formato == 'pdf':
            herramienta = detectar_herramienta_conversion()
            if not herramienta:
                return jsonify({
                    "error": "No hay herramientas de conversión PDF disponibles",
                    "mensaje": "Instale LibreOffice o pandoc para generar PDFs"
                }), 500
            try:
                if herramienta == 'libreoffice':
                    pdf_path, pdf_temp_dir = convertir_word_a_pdf_libreoffice(docx_path)
                elif herramienta == 'pandoc':
                    pdf_path, pdf_temp_dir = convertir_word_a_pdf_pandoc(docx_path)
                archivos_a_limpiar.append(pdf_temp_dir)
                response = send_file(pdf_path, 
                                   as_attachment=True, 
                                   download_name=f"cotizacion_{unique_id}.pdf")
            except Exception as e:
                return jsonify({
                    "error": "Error al convertir a PDF",
                    "detalle": str(e)
                }), 500
        else:
            return jsonify({"error": "Formato no válido. Use 'word' o 'pdf'"}), 400
        @response.call_on_close
        def cleanup():
            for directorio in archivos_a_limpiar:
                try:
                    if os.path.exists(directorio):
                        shutil.rmtree(directorio)
                except Exception as e:
                    print(f"Error al limpiar {directorio}: {e}")
        return response
    except Exception as e:
        for directorio in archivos_a_limpiar:
            try:
                if os.path.exists(directorio):
                    shutil.rmtree(directorio)
            except:
                pass
        return jsonify({
            "error": "Error interno del servidor",
            "detalle": str(e)
        }), 500

@app.route('/herramientas-disponibles', methods=['GET'])
def herramientas_disponibles():
    """
    Endpoint para verificar qué herramientas están disponibles
    """
    herramienta = detectar_herramienta_conversion()
    return jsonify({
        "herramienta_disponible": herramienta,
        "puede_generar_pdf": herramienta is not None,
        "formatos_soportados": ['word', 'pdf'] if herramienta else ['word']
    })

# Mantener endpoint original para compatibilidad
@app.route('/generar-word', methods=['POST'])
def generar_word():
    """
    Endpoint original - mantiene compatibilidad
    """
    data = request.get_json()
    data['formato'] = 'word'
    return generar_documento()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)