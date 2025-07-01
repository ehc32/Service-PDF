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
        data.get("correo", ""),
        data.get("Subtotal_1", ""),
        data.get("Subtotal_2", ""),
        data.get("Total", ""),
        data.get("texto", "")
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
    Genera documento en formato Word y/o PDF
    """
    data = request.get_json()
    formato = data.get('formato', 'word').lower()  # 'word', 'pdf', 'ambos'
    
    # Datos predeterminados solo si no vienen del usuario
    if not data.get("Acompañamie"): data["Acompañamie"] = "1.516.141"
    if not data.get("Diseño_Calcu"): data["Diseño_Calcu"] = data.get("Diseño_Calcu", "23.918.292")
    if not data.get("Diseño_Sanitario"): data["Diseño_Sanitario"] = "20.501.393"
    # Elimina el campo alternativo si existe    
    def extraer_numero(texto):
        if not texto:
            return 0
        return int(''.join(filter(str.isdigit, str(texto))) or 0)
    
    def formatear_moneda(numero):
        return "{:,.0f}".format(numero).replace(",", ".")
    
    # Cálculos
    subtotal1 = extraer_numero(data.get("Subtotal_1")) or (
        extraer_numero(data.get("Diseño_Ar")) +
        extraer_numero(data.get("Diseño_Calcu")) +
        extraer_numero(data.get("Acompañamie"))
    )
    subtotal2 = extraer_numero(data.get("Subtotal_2")) or (
        extraer_numero(data.get("Diseño_Calcu")) +
        extraer_numero(data.get("Diseño_Sanitario")) +
        extraer_numero(data.get("Presupuesta"))
    )
    total = subtotal1 + subtotal2
    
    # Agregar la fecha actual en formato '01 de junio de 2025'
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'es_CO.UTF-8')
        except:
            locale.setlocale(locale.LC_TIME, '')  # Usa el locale por defecto si no encuentra español
    fecha_actual = datetime.now().strftime('%d de %B de %Y')
    # Asegura que el mes esté en minúsculas
    fecha_actual = fecha_actual[:6] + fecha_actual[6:].lower()
    data["fecha"] = fecha_actual
    
    data["Subtotal_1"] = formatear_moneda(subtotal1)
    data["Subtotal_2"] = formatear_moneda(subtotal2)
    data["Total"] = formatear_moneda(total)
    data["texto"] = numero_a_texto(total)
    
    # Guardar en Google Sheets
    try:
        guardar_en_google_sheets(data)
    except Exception as e:
        print(f"Error al guardar en Google Sheets: {e}")
    
    # Verificar plantilla
    plantilla_path = os.path.join(os.path.dirname(__file__), "plantilla.docx")
    if not os.path.exists(plantilla_path):
        return jsonify({"error": "No se encontró la plantilla Word"}), 500
    
    # Generar documento Word
    doc = DocxTemplate(plantilla_path)
    doc.render(data)
    
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
            # Solo PDF
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
        
        # Cleanup al cerrar la respuesta
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
        # Cleanup en caso de error
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