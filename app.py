import streamlit as st
import google.generativeai as genai
import json
import io
import smtplib
from email.message import EmailMessage
import openpyxl

# 1. Configuración de página y credenciales
st.set_page_config(page_title="Generador de BRS", page_icon="⚙️")
st.title("Generador BRS con Plantilla Personalizada 🚀")

GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
EMAIL_SENDER = st.secrets["EMAIL_SENDER"]
EMAIL_PASSWORD = st.secrets["EMAIL_PASSWORD"]

genai.configure(api_key=GEMINI_API_KEY)

# 2. Interfaz de Usuario
email_destino = st.text_input("📧 Correo electrónico destino:")
archivos_subidos = st.file_uploader("📂 Sube tus archivos SQL (.txt)", type=["txt", "sql"], accept_multiple_files=True)

# Crear un diccionario para guardar los nombres de las tablas dinámicamente
nombres_tablas = {}

if archivos_subidos:
    st.markdown("### 📝 Nombres de Tablas Destino")
    st.info("Ingresa el nombre del reporte/tabla para cada script SQL subido:")
    for archivo in archivos_subidos:
        # Genera un input text por cada archivo subido
        nombres_tablas[archivo.name] = st.text_input(f"Nombre para '{archivo.name}':", key=archivo.name)

def analizar_sql_con_gemini(sql_query):
    """Llama a Gemini para extraer la estructura exacta requerida para la plantilla."""
    model = genai.GenerativeModel('gemini-1.5-flash-001')
    
    prompt = f"""
    Eres un Analista de Datos experto. Analiza el siguiente SQL y extrae la información para un BRS.
    Devuelve ÚNICAMENTE un JSON válido con esta estructura exacta:
    {{
      "campos": [
        {{
          "columna": "nombre del campo en el select",
          "termino_negocio": "explicación en lenguaje de negocio",
          "descripcion": "explicación detallada de la regla, lógica o CASE. Si es un campo directo, repite el termino_negocio.",
          "fuente_datos": "nombre de la tabla origen de este campo"
        }}
      ],
      "tablas_origen": ["TABLA_1", "TABLA_2"]
    }}
    
    Consulta SQL:
    {sql_query}
    """
    
    response = model.generate_content(
        prompt, 
        generation_config=genai.GenerationConfig(response_mime_type="application/json")
    )
    return json.loads(response.text)

def llenar_plantilla_excel(datos_gemini, nombre_reporte, sql_texto):
    """Abre la plantilla, inyecta los datos en las hojas requeridas y retorna los bytes."""
    # Cargar la plantilla existente desde el repositorio
    wb = openpyxl.load_workbook("plantilla_brs.xlsx")
    
    campos = datos_gemini.get("campos", [])
    tablas = datos_gemini.get("tablas_origen", [])
    
    # ---------------------------------------------------------
    # HOJA: Mapeo BRS
    # ---------------------------------------------------------
    ws_mapeo = wb["Mapeo BRS"]
    
    # Si hay más de un campo, insertamos filas para empujar la información de abajo
    if len(campos) > 1:
        # Insertamos filas a partir de la fila 11. 
        ws_mapeo.insert_rows(idx=11, amount=len(campos)-1)
        
    for i, campo in enumerate(campos):
        fila_actual = 10 + i
        ws_mapeo.cell(row=fila_actual, column=1, value=nombre_reporte)           # Col A
        ws_mapeo.cell(row=fila_actual, column=2, value=campo["columna"])         # Col B
        ws_mapeo.cell(row=fila_actual, column=3, value=campo["termino_negocio"]) # Col C
        ws_mapeo.cell(row=fila_actual, column=4, value=campo["descripcion"])     # Col D
        ws_mapeo.cell(row=fila_actual, column=5, value=campo["fuente_datos"])    # Col E

    # ---------------------------------------------------------
    # HOJA: PreRequisitos
    # ---------------------------------------------------------
    if "PreRequisitos" in wb.sheetnames:
        ws_req = wb["PreRequisitos"]
        for i, tabla in enumerate(tablas):
            ws_req.cell(row=2+i, column=3, value=tabla) # C2 hacia abajo
            
    # ---------------------------------------------------------
    # HOJA: Query
    # ---------------------------------------------------------
    if "Query" in wb.sheetnames:
        ws_query = wb["Query"]
        lineas_sql = sql_texto.split('\n')
        for i, linea in enumerate(lineas_sql):
            ws_query.cell(row=2+i, column=1, value=linea) # A2 hacia abajo

    # ---------------------------------------------------------
    # HOJA: Layout
    # ---------------------------------------------------------
    if "Layout" in wb.sheetnames:
        ws_layout = wb["Layout"]
        for i, campo in enumerate(campos):
            ws_layout.cell(row=4+i, column=1, value=campo["columna"]) # A4 hacia abajo

    # Guardar en memoria
    buffer_excel = io.BytesIO()
    wb.save(buffer_excel)
    return buffer_excel.getvalue()

def enviar_correo(email_dest, adjuntos):
    msg = EmailMessage()
    msg['Subject'] = 'Tus documentos BRS generados automáticamente'
    msg['From'] = EMAIL_SENDER
    msg['To'] = email_dest
    msg.set_content('Hola,\n\nAdjunto encontrarás los archivos BRS generados a partir de tus consultas SQL y tu plantilla base.\n\nSaludos.')

    for nombre_archivo, bytes_excel in adjuntos:
        msg.add_attachment(bytes_excel, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=nombre_archivo)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

# 3. Botón de ejecución
if st.button("🚀 Procesar Plantillas y Enviar"):
    # Validaciones antes de ejecutar
    faltan_nombres = any(not nombre.strip() for nombre in nombres_tablas.values())
    
    if not email_destino:
        st.warning("Por favor, ingresa un correo destino.")
    elif not archivos_subidos:
        st.warning("Por favor, sube al menos un archivo SQL.")
    elif faltan_nombres:
        st.warning("Debes asignar un Nombre de Tabla para cada archivo subido antes de procesar.")
    else:
        with st.spinner("Leyendo consultas e inyectando datos en tu plantilla..."):
            archivos_excel_generados = []
            
            for archivo in archivos_subidos:
                nombre_reporte = nombres_tablas[archivo.name]
                st.write(f"⏳ Analizando: {archivo.name} -> Tabla: {nombre_reporte}...")
                
                contenido_sql = archivo.getvalue().decode("utf-8")
                
                try:
                    # 1. Sacar datos de Gemini
                    datos_extraidos = analizar_sql_con_gemini(contenido_sql)
                    
                    # 2. Inyectar en la plantilla Excel
                    excel_bytes = llenar_plantilla_excel(datos_extraidos, nombre_reporte, contenido_sql)
                    
                    # 3. Preparar archivo para correo
                    nombre_final = f"BRS_{nombre_reporte.replace(' ', '_')}.xlsx"
                    archivos_excel_generados.append((nombre_final, excel_bytes))
                    
                except Exception as e:
                    st.error(f"Error procesando el archivo {archivo.name}: {e}")
            
            if archivos_excel_generados:
                st.write("📧 Enviando correo con adjuntos...")
                enviar_correo(email_destino, archivos_excel_generados)
                st.success(f"¡Éxito! Se han enviado {len(archivos_excel_generados)} archivos formateados a {email_destino}.")
