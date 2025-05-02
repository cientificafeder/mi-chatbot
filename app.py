import json
import os
import re
from flask import Flask, request, render_template, send_from_directory
from pypdf import PdfReader
from docx import Document
import pandas as pd
import logging
from functools import lru_cache

app = Flask(__name__)

ARCHIVO_ORDENES = "ordenes.json"
ARCHIVO_CONTEXTO = "contexto_pdf.txt"
IP_PUBLICA = "179.60.137.163"

# Configurar logging
logging.basicConfig(level=logging.DEBUG)

# Cache para el contexto
app.config['CONTEXTO_CACHE'] = None
app.config['ULTIMA_MODIFICACION'] = 0

# Mapa de carreras a valores en la columna CARRERA del archivo HORARIOS.xlsx
HORARIOS_MAP = {
    "carrera de alimentos": "ALIMENTOS",
    "gastronomía": "GASTRONOMIA",
    "ingeniería de la producción": "INGENIERÍA DE LA PRODUCCION",
    "ingeniería química": "INGENIERÍA QUÍMICA-2018"
}

# Mapa para identificar palabras clave únicas de cada carrera
CARRERA_CLAVE_MAP = {
    "carrera de alimentos": "alimentos",
    "gastronomía": "gastronomía",
    "ingeniería de la producción": "producción",
    "ingeniería química": "química"
}

# Descripciones de las carreras
DESCRIPCIONES_CARRERAS = {
    "carrera de alimentos": (
        "<p><strong>Título:</strong> Ingeniero/a de los Alimentos</p>"
        "<p><strong>Duración:</strong> 9 semestres</p>"
        "<p><strong>Modalidad:</strong> Presencial</p>"
        "<p><strong>Director(a) de la Carrera:</strong> Ing. Gonzalo Villa M., MSc.</p>"
        "<p><strong>Descripción de la carrera:</strong> Formar profesionales ingenieros/as en alimentos competentes a nivel nacional e internacional, que reúnan sólidos conocimientos humanos, científicos y tecnológicos, capacitados en el diseño, mejora de los procesos productivos relacionados con la industria alimentaria así como el manejo, transporte, almacenamiento y procesamiento de productos agrícolas, hortícolas, frutícolas, animales, granos, acuícolas y productos del mar, con el fin de obtener productos alimenticios sanos, nutritivos, de alta calidad a precios competitivos.</p>"
    ),
    "gastronomía": (
        "<p><strong>Título:</strong> Chef Internacional</p>"
        "<p><strong>Duración:</strong> 8 semestres</p>"
        "<p><strong>Modalidad:</strong> Presencial</p>"
        "<p><strong>Director(a) de la Carrera:</strong> Chef María López, MSc.</p>"
        "<p><strong>Descripción de la carrera:</strong> Formar chefs profesionales con un enfoque internacional, capaces de liderar cocinas de alta gama, innovar en técnicas culinarias y gestionar proyectos gastronómicos. Los egresados dominarán el arte de la cocina, desde la preparación de platos tradicionales hasta la creación de experiencias culinarias únicas, con énfasis en la sostenibilidad y el uso de ingredientes locales.</p>"
    ),
    "ingeniería de la producción": (
        "<p><strong>Título:</strong> Ingeniero/a en Producción Industrial</p>"
        "<p><strong>Duración:</strong> 9 semestres</p>"
        "<p><strong>Modalidad:</strong> Presencial</p>"
        "<p><strong>Director(a) de la Carrera:</strong> Ing. Carlos Mendoza, PhD.</p>"
        "<p><strong>Descripción de la carrera:</strong> Preparar ingenieros/as capaces de optimizar procesos industriales, implementar sistemas de producción eficientes y liderar proyectos de mejora continua. Los profesionales estarán capacitados para diseñar, gestionar y supervisar operaciones en plantas industriales, garantizando calidad, seguridad y sostenibilidad en los procesos productivos.</p>"
    ),
    "ingeniería química": (
        "<p><strong>Título:</strong> Ingeniero/a Químico/a</p>"
        "<p><strong>Duración:</strong> 10 semestres</p>"
        "<p><strong>Modalidad:</strong> Presencial</p>"
        "<p><strong>Director(a) de la Carrera:</strong> Ing. Laura Gómez, PhD.</p>"
        "<p><strong>Descripción de la carrera:</strong> Formar ingenieros/as químicos/as con un sólido conocimiento en procesos químicos y biotecnológicos, capaces de diseñar y operar plantas industriales, desarrollar nuevos materiales y productos, y garantizar la sostenibilidad ambiental. Los egresados estarán preparados para enfrentar desafíos en industrias como la farmacéutica, petroquímica y de energías renovables.</p>"
    )
}

# Mapa de PDFs de reseña histórica asociados a cada carrera
RESEÑAS_CARRERAS = {
    "carrera de alimentos": "Reseña_Carrera_de_Alimentos.pdf",
    "gastronomía": "Reseña_Gastronomía.pdf",
    "ingeniería de la producción": "Reseña_Ingeniería_de_la_Producción.pdf",
    "ingeniería química": "Reseña_Ingeniería_Química.pdf"
}

# Mapa de PDFs de mallas curriculares asociados a cada carrera
MALLAS_CARRERAS = {
    "carrera de alimentos": "MALLA CARRERA DE ALIMENTOS.pdf",
    "gastronomía": "MALLA CARRERA DE GASTRONOMÍA.pdf",
    "ingeniería de la producción": "MALLA CARRERA DE INGENIERÍA DE LA PRODUCCIÓN.pdf",
    "ingeniería química": "MALLA CARRERA DE INGENIERÍA QUÍMICA.pdf"
}

# Mapa de materias a PDFs de sílabos
MATERIAS_SILABOS = {
    "álgebra lineal": "ÁLGEBRA LINEAL.pdf",
    "cálculo diferencial": "CÁLCULO DIFERENCIAL.pdf",
    "democracia ciudadanía y globalización": "DEMOCRACIA, CIUDADANÍA Y GLOBALIZACIÓN.pdf",
    "física aplicada a la ingeniería": "FÍSICA APLICADA A LA INGENIERÍA.pdf",
    "física i mecánica": "FÍSICA I MECÁNICA.pdf",
    "fundamentos de matemáticas para la ingeniería": "FUNDAMENTOS DE MATEMÁTICAS PARA LA INGENIERÍA.pdf",
    "fundamentos de programación": "FUNDAMENTOS DE PROGRAMACIÓN.pdf",
    "gastronomía historia arte y ciencia": "GASTRONOMÍA, HISTORIA, ARTE Y CIENCIA.pdf",
    "herramientas digitales i": "HERRAMIENTAS DIGITALES I.pdf",
    "higiene y seguridad en la manipulación de los alimentos": "HIGIENE Y SEGURIDAD EN LA MANIPULACIÓN DE LOS ALIMENTOS.pdf",
    "lenguaje y comunicación": "LENGUAJE Y COMUNICACIÓN.pdf",
    "química general": "QUÍMICA GENERAL.pdf",
    "técnicas culinarias": "TÉCNICAS CULINARIAS.pdf"
}

def cargar_ordenes():
    return {
        "hola": (
            "Hola, bienvenido al asistente académico. Aquí tienes las carreras disponibles:<br>"
            '<div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 10px;">'
            '<a href="/carrera/carrera de alimentos" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Carrera de Alimentos</a>'
            '<a href="/carrera/ingeniería de la producción" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Ingeniería de la Producción</a>'
            '<a href="/carrera/gastronomía" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Gastronomía</a>'
            '<a href="/carrera/ingeniería química" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Ingeniería Química</a>'
            '</div><br>¡Haz clic en una carrera para ver más información!<br>'
            'También puedes hacerme preguntas generales aquí sobre los documentos cargados o preguntar por el sílabo de una materia.'
        ),
        "salir": "Hasta pronto."
    }

def leer_pdf(ruta_pdf):
    try:
        lector = PdfReader(ruta_pdf)
        texto = ""
        for pagina in lector.pages:
            texto_pagina = pagina.extract_text()
            if texto_pagina:
                texto += texto_pagina + "\n"
        if not texto.strip():
            logging.warning(f"El PDF {ruta_pdf} está vacío o no contiene texto legible.")
            return None, "El PDF está vacío o no contiene texto legible."
        logging.debug(f"Texto extraído del PDF {ruta_pdf}: {texto[:100]}...")
        return texto, None
    except Exception as e:
        logging.error(f"path {ruta_pdf}: Error al leer el PDF: {e}")
        return None, f"No se pudo leer el PDF: {e}"

def leer_docx(ruta_docx):
    try:
        doc = Document(ruta_docx)
        texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        if not texto.strip():
            logging.warning(f"El DOCX {ruta_docx} está vacío.")
            return None, "El documento Word está vacío."
        return texto, None
    except Exception as e:
        logging.error(f"Error al leer el documento Word: {e}")
        return None, f"No se pudo leer el documento Word: {e}"

def normalize_sheet_name(sheet_name):
    """Normaliza nombres de hojas para comparación (sin acentos ni mayúsculas)."""
    return sheet_name.lower().replace("í", "i").replace("á", "a").replace("é", "e").replace("ó", "o").replace("ú", "u").replace(" ", "")

def leer_excel_hoja(ruta_excel, carrera):
    try:
        # Verificar las hojas disponibles en el archivo Excel
        xl = pd.ExcelFile(ruta_excel)
        hojas_disponibles = xl.sheet_names
        logging.debug(f"Hojas disponibles en {ruta_excel}: {hojas_disponibles}")
        
        # Leer la primera hoja (asumimos que es la hoja principal)
        if not hojas_disponibles:
            logging.error(f"No se encontraron hojas en el archivo {ruta_excel}.")
            return None
        
        hoja_principal = hojas_disponibles[0]  # Tomamos la primera hoja
        df = pd.read_excel(ruta_excel, sheet_name=hoja_principal)
        
        # Filtrar las filas correspondientes a la carrera
        df_carrera = df[df['CARRERA'].str.strip() == carrera].copy()
        
        if df_carrera.empty:
            logging.error(f"No se encontraron horarios para la carrera '{carrera}' en el archivo {ruta_excel}.")
            return None
        
        df_carrera = df_carrera.fillna("")  # Reemplazar NaN con cadena vacía
        logging.debug(f"Horarios para la carrera {carrera} leídos correctamente: {df_carrera.shape}")
        return df_carrera
    except Exception as e:
        logging.error(f"Error al leer el archivo Excel {ruta_excel}: {e}")
        return None

def normalize_filename(filename):
    """Normaliza nombres de archivos para comparación (sin acentos ni mayúsculas)."""
    return filename.lower().replace("í", "i").replace("á", "a").replace("é", "e").replace("ó", "o").replace("ú", "u").replace(" ", "")

def find_file_in_uploads(expected_filename, folder="uploads"):
    """Busca un archivo en la carpeta uploads, ignorando mayúsculas y acentos."""
    if not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)  # Crear el directorio si no existe
        return None
    expected_normalized = normalize_filename(expected_filename)
    for filename in os.listdir(folder):
        filename_normalized = normalize_filename(filename)
        if filename_normalized == expected_normalized:
            return filename  # Devuelve el nombre real del archivo encontrado
    return None

def extract_syllabus_info(pdf_path):
    """Extrae información estructurada de un sílabo en formato PDF."""
    texto, error = leer_pdf(pdf_path)
    if error:
        return None, error

    # Inicializar variables para los datos
    datos_informativos = {
        "Facultad": "",
        "Dominio": "",
        "Carrera": "",
        "Asignatura": "",
        "Código": "",
        "UOC": "",
        "Campo Formación": "",
        "Plan de estudios": "",
        "N° Créditos": "",
        "Horas componente docencia": "",
        "Horas componente de práctica y experimental": "",
        "Horas componente autónomas": "",
        "Horas totales": "",
        "Prerrequisito": "",
        "Semestre": ""
    }
    justificacion = ""

    # Expresiones regulares para extraer datos informativos
    datos_regex = {
        "Facultad": r"Facultad:\s*(.*?)(?=\n|$|Dominio:)",
        "Dominio": r"Dominio:\s*(.*?)(?=\n|$|Carrera:)",
        "Carrera": r"Carrera:\s*(.*?)(?=\n|$|Asignatura:)",
        "Asignatura": r"Asignatura:\s*(.*?)(?=\n|$|Código:)",
        "Código": r"Código:\s*(.*?)(?=\n|$|UOC:)",
        "UOC": r"UOC:\s*(.*?)(?=\n|$|Campo Formación:)",
        "Campo Formación": r"Campo Formación:\s*(.*?)(?=\n|$|Plan de estudios:)",
        "Plan de estudios": r"Plan de estudios:\s*(.*?)(?=\n|$|N° Créditos:)",
        "N° Créditos": r"N° Créditos:\s*(.*?)(?=\n|$|Horas componente docencia:)",
        "Horas componente docencia": r"Horas componente docencia:\s*(.*?)(?=\n|$|Horas componente de práctica y experimental:)",
        "Horas componente de práctica y experimental": r"Horas componente de práctica y experimental:\s*(.*?)(?=\n|$|Horas componente autónomas:)",
        "Horas componente autónomas": r"Horas componente autónomas:\s*(.*?)(?=\n|$|Horas totales:)",
        "Horas totales": r"Horas totales:\s*(.*?)(?=\n|$|Prerrequisito:)",
        "Prerrequisito": r"Prerrequisito:\s*(.*?)(?=\n|$|Semestre:)",
        "Semestre": r"Semestre:\s*(.*?)(?=\n|$|B\. JUSTIFICACIÓN)"
    }

    # Extraer datos informativos
    for key, regex in datos_regex.items():
        match = re.search(regex, texto, re.DOTALL)
        if match:
            datos_informativos[key] = match.group(1).strip()

    # Extraer justificación
    justificacion_match = re.search(r"B\. JUSTIFICACIÓN DEL CONOCIMIENTO DEL SÍLABO EN EL CAMPO DE FORMACIÓN\s*Breve justificación de los contenidos del Sílabo:\s*(.*?)(?=\n\d+\s+\d+/\d+/\d+|$)", texto, re.DOTALL)
    if justificacion_match:
        justificacion = justificacion_match.group(1).strip()

    return datos_informativos, justificacion

def format_syllabus_html(datos_informativos, justificacion):
    """Formatea la información del sílabo en HTML."""
    # Crear tabla para Datos Informativos
    tabla_html = (
        "<h2>A. Datos Informativos</h2>"
        "<table class='table table-striped'>"
        "<tbody>"
        f"<tr><td><strong>Facultad:</strong></td><td>{datos_informativos['Facultad']}</td></tr>"
        f"<tr><td><strong>Dominio:</strong></td><td>{datos_informativos['Dominio']}</td></tr>"
        f"<tr><td><strong>Carrera:</strong></td><td>{datos_informativos['Carrera']}</td></tr>"
        f"<tr><td><strong>Asignatura:</strong></td><td>{datos_informativos['Asignatura']}</td></tr>"
        f"<tr><td><strong>Código:</strong></td><td>{datos_informativos['Código']}</td></tr>"
        f"<tr><td><strong>UOC:</strong></td><td>{datos_informativos['UOC']}</td></tr>"
        f"<tr><td><strong>Campo Formación:</strong></td><td>{datos_informativos['Campo Formación']}</td></tr>"
        f"<tr><td><strong>Plan de estudios:</strong></td><td>{datos_informativos['Plan de estudios']}</td></tr>"
        f"<tr><td><strong>N° Créditos:</strong></td><td>{datos_informativos['N° Créditos']}</td></tr>"
        f"<tr><td><strong>Horas componente docencia:</strong></td><td>{datos_informativos['Horas componente docencia']}</td></tr>"
        f"<tr><td><strong>Horas componente de práctica y experimental:</strong></td><td>{datos_informativos['Horas componente de práctica y experimental']}</td></tr>"
        f"<tr><td><strong>Horas componente autónomas:</strong></td><td>{datos_informativos['Horas componente autónomas']}</td></tr>"
        f"<tr><td><strong>Horas totales:</strong></td><td>{datos_informativos['Horas totales']}</td></tr>"
        f"<tr><td><strong>Prerrequisito:</strong></td><td>{datos_informativos['Prerrequisito']}</td></tr>"
        f"<tr><td><strong>Semestre:</strong></td><td>{datos_informativos['Semestre']}</td></tr>"
        "</tbody></table>"
    )

    # Agregar justificación
    justificacion_html = (
        "<h2>B. Justificación del Conocimiento del Sílabo en el Campo de Formación</h2>"
        "<p><strong>Breve justificación de los contenidos del Sílabo:</strong></p>"
        f"<p>{justificacion}</p>"
    )

    return tabla_html + justificacion_html

def buscar_silabo(materia):
    """Busca el sílabo de una materia y devuelve su información formateada."""
    materia_normalizada = normalize_filename(materia)
    for nombre_materia, pdf_filename in MATERIAS_SILABOS.items():
        if materia_normalizada in normalize_filename(nombre_materia):
            pdf_real_filename = find_file_in_uploads(pdf_filename)
            if pdf_real_filename:
                pdf_path = os.path.join("uploads", pdf_real_filename)
                datos_informativos, justificacion = extract_syllabus_info(pdf_path)
                if datos_informativos and justificacion:
                    return format_syllabus_html(datos_informativos, justificacion)
                else:
                    return f"No se pudo extraer la información del sílabo para {nombre_materia}. Asegúrate de que el PDF contenga el formato esperado."
            else:
                return f"No se encontró el archivo del sílabo para {nombre_materia} en la carpeta uploads."
    return f"No se encontró un sílabo para la materia {materia}."

def cargar_contexto_desde_carpeta(carpeta_raiz, carrera=None):
    try:
        ultima_modificacion = max([os.path.getmtime(os.path.join(carpeta_raiz, f)) 
                                  for f in os.listdir(carpeta_raiz) if os.path.isfile(os.path.join(carpeta_raiz, f))] 
                                  or [0])
        
        if carrera is None:
            if ultima_modificacion > app.config['ULTIMA_MODIFICACION'] or app.config['CONTEXTO_CACHE'] is None:
                contexto_total = ""
                if os.path.exists(carpeta_raiz):
                    for archivo in os.listdir(carpeta_raiz):
                        ruta_archivo = os.path.join(carpeta_raiz, archivo)
                        if archivo.lower().endswith(".pdf"):
                            contenido, error = leer_pdf(ruta_archivo)
                            if error:
                                logging.warning(error)
                                continue
                        elif archivo.lower().endswith(".docx"):
                            contenido, error = leer_docx(ruta_archivo)
                            if error:
                                logging.warning(error)
                                continue
                        elif archivo.lower().endswith((".xlsx", ".xls")):
                            xl = pd.ExcelFile(ruta_archivo)
                            for sheet_name in xl.sheet_names:
                                df = pd.read_excel(ruta_archivo, sheet_name=sheet_name, header=None)
                                df = df.fillna("")
                                contenido = f"\nHoja: {sheet_name}\n{df.to_string(index=False, header=False)}\n"
                        else:
                            contenido = None
                        if contenido:
                            contexto_total += f"\n--- Contenido de {archivo} ---\n{contenido}\n"
                app.config['CONTEXTO_CACHE'] = contexto_total
                app.config['ULTIMA_MODIFICACION'] = ultima_modificacion
            return app.config['CONTEXTO_CACHE']
        
        else:
            contexto_total = ""
            clave_carrera = CARRERA_CLAVE_MAP.get(carrera, "")
            if not clave_carrera:
                logging.warning(f"No se encontró palabra clave para la carrera {carrera}")
                return ""
            
            clave_normalizada = normalize_filename(clave_carrera)
            logging.debug(f"Buscando PDFs para la carrera: {carrera}, clave: {clave_normalizada}")
            
            if os.path.exists(carpeta_raiz):
                for archivo in os.listdir(carpeta_raiz):
                    archivo_normalizado = normalize_filename(archivo)
                    if archivo.lower().endswith(".pdf") and clave_normalizada in archivo_normalizado:
                        ruta_archivo = os.path.join(carpeta_raiz, archivo)
                        contenido, error = leer_pdf(ruta_archivo)
                        if contenido:
                            contexto_total += f"\n--- Contenido de {archivo} ---\n{contenido}\n"
                            logging.debug(f"PDF {archivo} agregado al contexto para {carrera}")
                        else:
                            logging.warning(f"No se pudo extraer contenido del PDF {archivo}: {error}")
            
            if not contexto_total.strip():
                logging.warning(f"No se encontró contenido para la carrera {carrera}")
            else:
                logging.debug(f"Contexto generado para {carrera}: {contexto_total[:200]}...")
            return contexto_total
    except Exception as e:
        logging.error(f"Error al cargar contexto desde {carpeta_raiz}: {e}")
        return ""

def identificar_carrera(mensaje):
    """Identifica si la pregunta menciona una carrera conocida."""
    mensaje = mensaje.lower()
    for carrera in CARRERA_CLAVE_MAP.keys():
        if carrera in mensaje:
            return carrera
    return None

def identificar_materia(mensaje):
    """Identifica si la pregunta menciona una materia conocida."""
    mensaje = mensaje.lower()
    for materia in MATERIAS_SILABOS.keys():
        if materia in mensaje:
            return materia
    return None

def extraer_titulo(descripcion):
    """Extrae el título de la descripción de la carrera."""
    match = re.search(r"<strong>Título:</strong>\s*([^<]+)", descripcion)
    if match:
        return match.group(1).strip()
    return "Título no disponible"

@lru_cache(maxsize=100)
def consultar_ia(mensaje, contenido_pdf):
    carrera_mencionada = identificar_carrera(mensaje)
    logging.debug(f"Carrera mencionada en la pregunta: {carrera_mencionada}")

    contexto_adicional = ""
    if carrera_mencionada and carrera_mencionada in DESCRIPCIONES_CARRERAS:
        contexto_adicional = DESCRIPCIONES_CARRERAS[carrera_mencionada]
        contexto_adicional = re.sub(r'<[^>]+>', '', contexto_adicional)

    contexto_completo = (contenido_pdf or "") + "\n" + contexto_adicional

    if not contexto_completo.strip():
        return "No hay contenido de documentos cargado ni información disponible para responder tu pregunta. Asegúrate de que haya PDFs cargados o menciona una carrera válida."

    # Buscar palabras clave en el contexto
    mensaje_lower = mensaje.lower()
    if "duración" in mensaje_lower:
        match = re.search(r"Duración:\s*(\d+\s*semestres)", contexto_completo, re.IGNORECASE)
        if match:
            return match.group(1)
    elif "título" in mensaje_lower:
        match = re.search(r"Título:\s*([^\n<]+)", contexto_completo, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    elif "director" in mensaje_lower:
        match = re.search(r"Director\(a\) de la Carrera:\s*([^\n<]+)", contexto_completo, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    # Si no encuentra una respuesta específica, devuelve un mensaje genérico
    if carrera_mencionada and carrera_mencionada in DESCRIPCIONES_CARRERAS:
        return "Por favor, especifica qué información deseas sobre la carrera (por ejemplo, título, duración, descripción)."
    return "No pude encontrar una respuesta útil. Intenta reformular tu pregunta o asegúrate de que los documentos cargados tengan la información necesaria."

@app.route("/", methods=["GET", "POST"])
def chatbot_web():
    ordenes = cargar_ordenes()
    respuesta = ""
    os.makedirs("uploads", exist_ok=True)
    contexto_pdf = cargar_contexto_desde_carpeta("uploads")
    pdf_cargado = bool(contexto_pdf)

    if request.method == "POST":
        if "mensaje" in request.form and request.form["mensaje"]:
            mensaje = request.form["mensaje"].lower().strip()

            if mensaje == "salir":
                respuesta = ordenes.get("salir", "Chao!")
            elif mensaje.startswith("enséñame:"):
                try:
                    parte = mensaje.replace("enséñame:", "").strip()
                    orden, resp = parte.split("=", 1)
                    ordenes[orden.strip()] = resp.strip()
                    respuesta = f"Entrenado: '{orden}' → '{resp}'"
                except ValueError:
                    respuesta = "Usa el formato 'enséñame: orden = respuesta'"
            elif mensaje in ordenes:
                respuesta = ordenes[mensaje]
            # Detectar si la pregunta es sobre un sílabo
            elif "sílabo" in mensaje or "silabo" in mensaje:
                materia = identificar_materia(mensaje)
                if materia:
                    respuesta = buscar_silabo(materia)
                else:
                    respuesta = "No reconocí la materia en tu pregunta. Por favor, menciona una materia válida, como 'Álgebra Lineal' o 'Técnicas Culinarias'."
            else:
                respuesta = consultar_ia(mensaje, contexto_pdf)

        elif "pdf" in request.files:
            archivo = request.files["pdf"]
            if archivo and archivo.filename:
                ruta = os.path.join("uploads", archivo.filename)
                archivo.save(ruta)
                contexto_pdf = cargar_contexto_desde_carpeta("uploads")
                pdf_cargado = bool(contexto_pdf)
                respuesta = "Documento cargado exitosamente. ¡Hazme preguntas!"
            else:
                respuesta = "No se seleccionó un archivo válido."

        elif "limpiar" in request.form:
            if os.path.exists(ARCHIVO_CONTEXTO):
                os.remove(ARCHIVO_CONTEXTO)
            try:
                for archivo in os.listdir("uploads"):
                    os.remove(os.path.join("uploads", archivo))
            except Exception as e:
                respuesta = f"Contexto limpiado, pero hubo un error al eliminar archivos: {e}"
            else:
                respuesta = "Contexto limpiado. Carga un nuevo documento para empezar."
            contexto_pdf = ""
            pdf_cargado = False

    if not respuesta:
        respuesta = ordenes.get("hola", "")

    return render_template("index.html", respuesta=respuesta, pdf_cargado=pdf_cargado)

@app.route("/carrera/<nombre_carrera>", methods=["GET", "POST"])
def info_carrera(nombre_carrera):
    contexto_pdf = cargar_contexto_desde_carpeta("uploads", carrera=nombre_carrera)
    pdf_cargado = bool(contexto_pdf)
    ordenes = cargar_ordenes()
    respuesta = ""

    carreras_validas = [
        "carrera de alimentos",
        "ingeniería de la producción",
        "gastronomía",
        "ingeniería química"
    ]

    if nombre_carrera not in carreras_validas:
        respuesta = (
            "<h2>Carrera no encontrada</h2>"
            "<p>Lo siento, la carrera que buscas no está disponible. Aquí tienes las carreras que puedes explorar:</p>"
            '<div style="display: flex; flex-wrap: wrap; gap: 10px; margin-top: 10px;">'
            '<a href="/carrera/carrera de alimentos" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Carrera de Alimentos</a>'
            '<a href="/carrera/ingeniería de la producción" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Ingeniería de la Producción</a>'
            '<a href="/carrera/gastronomía" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Gastronomía</a>'
            '<a href="/carrera/ingeniería química" style="padding: 10px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px;">Ingeniería Química</a>'
            '</div><br>'
            '<p><a href="/">Volver al inicio</a></p>'
        )
        return render_template("index.html", respuesta=respuesta, pdf_cargado=False)

    if request.method == "POST":
        if "mensaje" in request.form and request.form["mensaje"]:
            mensaje = request.form["mensaje"].lower().strip()

            if mensaje == "salir":
                respuesta = ordenes.get("salir", "Chao!")
            elif mensaje.startswith("enséñame:"):
                try:
                    parte = mensaje.replace("enséñame:", "").strip()
                    orden, resp = parte.split("=", 1)
                    ordenes[orden.strip()] = resp.strip()
                    respuesta = f"Entrenado: '{orden}' → '{resp}'"
                except ValueError:
                    respuesta = "Usa el formato 'enséñame: orden = respuesta'"
            elif mensaje in ordenes:
                respuesta = ordenes[mensaje]
            # Detectar si la pregunta es sobre un sílabo
            elif "sílabo" in mensaje or "silabo" in mensaje:
                materia = identificar_materia(mensaje)
                if materia:
                    respuesta = buscar_silabo(materia)
                else:
                    respuesta = "No reconocí la materia en tu pregunta. Por favor, menciona una materia válida, como 'Álgebra Lineal' o 'Técnicas Culinarias'."
            else:
                respuesta = consultar_ia(mensaje, contexto_pdf)

    if not respuesta:
        carpeta = "uploads"
        clave_carrera = CARRERA_CLAVE_MAP.get(nombre_carrera, "")
        if not clave_carrera:
            respuesta = (
                f"<h2>{nombre_carrera.upper()}</h2>"
                f"<p>No se encontró información para {nombre_carrera}.</p>"
                f'<p><a href="/">Volver al inicio</a></p>'
            )
        else:
            career_class = nombre_carrera.replace(" ", "-")
            descripcion_carrera = DESCRIPCIONES_CARRERAS.get(nombre_carrera, "Descripción no disponible.")
            reseña_pdf = RESEÑAS_CARRERAS.get(nombre_carrera)
            malla_pdf = MALLAS_CARRERAS.get(nombre_carrera)

            # Verificar y agregar enlace a la reseña histórica
            reseña_html = ""
            reseña_filename = find_file_in_uploads(reseña_pdf) if reseña_pdf else None
            if reseña_filename:
                reseña_html = f'<p><a href="/uploads/{reseña_filename}" target="_blank" class="download-btn">Descargar Reseña Histórica</a></p>'
            else:
                reseña_html = "<p>No se encontró la reseña histórica para esta carrera.</p>"

            # Verificar y agregar enlace a la malla curricular
            malla_html = ""
            malla_filename = find_file_in_uploads(malla_pdf) if malla_pdf else None
            if malla_filename:
                malla_html = f'<p><a href="/uploads/{malla_filename}" target="_blank" class="download-btn">Descargar Malla Curricular</a></p>'
            else:
                malla_html = "<p>No se encontró la malla curricular para esta carrera.</p>"

            # Verificar y agregar enlace a los horarios
            horarios_link = ""
            horarios_filename = find_file_in_uploads("HORARIOS.xlsx")
            if horarios_filename and nombre_carrera in HORARIOS_MAP:
                horarios_link = f'<li><a href="/horarios/{nombre_carrera}" style="color: #007bff; text-decoration: none;">Ver horarios de {nombre_carrera}</a></li>'
            else:
                horarios_link = "<li>No se encontraron horarios para esta carrera.</li>"

            respuesta = (
                f'<div class="carrera-{career_class}">'
                f'<div class="career-card">'
                f'{descripcion_carrera}'
                f'{reseña_html}'
                f'{malla_html}'
                f'</div>'
                f"<p><strong>🔗 Otros recursos:</strong></p>"
                f"<ul>{horarios_link}</ul>"
                f'<p><a href="/">Volver al inicio</a></p>'
                f'</div>'
            )

    return render_template("index.html", respuesta=respuesta, pdf_cargado=pdf_cargado)

@app.route("/documento/<nombre_archivo>", methods=["GET", "POST"])
def ver_documento(nombre_archivo):
    carpeta = "uploads"
    ruta_archivo = os.path.join(carpeta, nombre_archivo)
    
    if os.path.exists(ruta_archivo) and nombre_archivo.lower().endswith((".docx", ".xlsx", ".xls")):
        if nombre_archivo.lower().endswith(".docx"):
            contenido, error = leer_docx(ruta_archivo)
        elif nombre_archivo.lower().endswith((".xlsx", ".xls")) and nombre_archivo.lower() != "horarios.xlsx":
            contenido, error = leer_excel_hoja(ruta_archivo, 0)
            if contenido is not None:
                contenido = contenido.to_html(index=False, header=True, classes="table table-striped")
        else:
            contenido, error = None, "Formato no soportado aquí"
        
        if contenido:
            referrer_carrera = request.referrer.split("/")[-1] if request.referrer else "unknown"
            respuesta = (
                f"<h2>Contenido de {nombre_archivo}</h2>"
                f"<div style='white-space: pre-wrap;'>{contenido}</div>"
                f'<p><a href="/carrera/{referrer_carrera}">Volver a la carrera</a> | '
                f'<a href="/">Volver al inicio</a></p>'
            )
        else:
            respuesta = f"No se pudo leer el archivo {nombre_archivo}: {error or 'Archivo vacío o no legible'}"
    else:
        respuesta = "Archivo no encontrado o no soportado. <a href='/'>Volver al inicio</a>"

    return render_template("index.html", respuesta=respuesta)

@app.route("/horarios/<nombre_carrera>", methods=["GET", "POST"])
def ver_horarios(nombre_carrera):
    if nombre_carrera not in HORARIOS_MAP:
        return "Carrera no encontrada. <a href='/'>Volver al inicio</a>"
    
    carrera = HORARIOS_MAP[nombre_carrera]
    horarios_filename = find_file_in_uploads("HORARIOS.xlsx")
    
    if not horarios_filename:
        logging.error("Archivo HORARIOS.xlsx no encontrado en la carpeta uploads.")
        return "Archivo HORARIOS.xlsx no encontrado. <a href='/'>Volver al inicio</a>"
    
    ruta_archivo = os.path.join("uploads", horarios_filename)
    
    if os.path.exists(ruta_archivo):
        df = leer_excel_hoja(ruta_archivo, carrera)
        if df is not None:
            thead_html = "<thead><tr>" + "".join(f"<th>{col}</th>" for col in df.columns) + "</tr></thead>"
            tbody_html = "<tbody>" + "".join(
                "<tr>" + "".join(f"<td>{str(val)}</td>" for val in row) + "</tr>"
                for row in df.itertuples(index=False)
            ) + "</tbody>"
            tabla_html = f'<table class="table table-striped">{thead_html}{tbody_html}</table>'
            respuesta = (
                f"<h2>Horarios de {nombre_carrera.upper()}</h2>"
                f'<div class="table-container">{tabla_html}</div>'
                f'<p><a href="/carrera/{nombre_carrera}">Volver a la carrera</a> | '
                f'<a href="/">Volver al inicio</a></p>'
            )
        else:
            respuesta = f"No se encontraron horarios para la carrera '{carrera}' en HORARIOS.xlsx. Asegúrate de que el archivo sea válido y contenga datos para esta carrera. <a href='/'>Volver al inicio</a>"
    else:
        logging.error(f"Archivo {ruta_archivo} no encontrado en el sistema de archivos.")
        respuesta = "Archivo HORARIOS.xlsx no encontrado en el sistema de archivos. <a href='/'>Volver al inicio</a>"

    return render_template("index.html", respuesta=respuesta)

@app.route("/uploads/<filename>")
def serve_uploaded_file(filename):
    try:
        return send_from_directory("uploads", filename)
    except Exception as e:
        logging.error(f"Error al servir archivo {filename}: {e}")
        return "Archivo no encontrado. <a href='/'>Volver al inicio</a>", 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"Chat_Boot_IQ está corriendo en http://{IP_PUBLICA}:{port}")
    app.run(host="0.0.0.0", port=port, debug=False)