import streamlit as st
import requests
import json
import io
import smtplib
from email.message import EmailMessage
import openpyxl
from copy import copy
import re
import time

# 1. Configuración de página y credenciales
st.set_page_config(page_title="Generador de BRS", page_icon="⚙️")
st.title("Generador BRS con Plantilla Personalizada 🚀")

GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"].strip()
EMAIL_SENDER = st.secrets["EMAIL_SENDER"].strip()
EMAIL_PASSWORD = st.secrets["EMAIL_PASSWORD"].strip()

# 2. Paso 1: Subir archivos fuera del formulario para detectar cuántos son
archivos_subidos = st.file_uploader(
    "📂 1. Sube tus archivos SQL (.txt o .sql)", 
    type=["txt", "sql"], 
    accept_multiple_files=True
)

# 3. Paso 2: Formulario cerrado (Evita que Streamlit recargue la página al escribir)
if archivos_subidos:
    with st.form("formulario_brs"):
        st.markdown("### 📝 2. Datos de Envío y Nombres de Tablas")
        email_destino = st.text_input("📧 Correo electrónico destino:")
        
        st.info("Ingresa el nombre del reporte/tabla para cada script SQL subido:")
        nombres_tablas = {}
        for archivo in archivos_subidos:
            nombres_tablas[archivo.name] = st.text_input(
                f"Nombre de tabla para '{archivo.name}':", 
                key=f"input_{archivo.name}"
            )
            
        # El botón de submit es el ÚNICO que dispara la ejecución
        boton_procesar = st.form_submit_button("🚀 Procesar Plantillas y Enviar")
else:
    boton_procesar = False

def obtener_modelos_activos():
    """Consulta directamente a Google qué modelos existen hoy en tu cuenta para no adivinar nombres."""
    url_modelos = f"https://generativelanguage.googleapis.com/v1beta/models?key={GEMINI_API_KEY}"
    try:
        resp = requests.get(url_modelos, timeout=15)
        if resp.status_code == 200:
            datos = resp.json().get("models", [])
            candidatos = []
            for m in datos:
                metodos = m.get("supportedGenerationMethods", [])
                nombre = m.get("name", "").replace("models/", "")
                # Filtramos solo modelos de generación de texto (excluimos embeddings, imagen, audio, o modelos viejos 1.5/2.0/2.5)
                if "generateContent" in metodos and "gemini" in nombre:
                    if not any(x in nombre for x in ["1.0", "1.5", "2.0", "2.5", "vision", "embedding", "imagen", "tts", "audio"]):
                        candidatos.append(nombre)
            
            # Ordenamos para priorizar flash/lite primero, pero dejando pro u otros como respaldo real
            candidatos.sort(key=lambda x: (0 if "flash" in x else 1, x))
            if candidatos:
                return candidatos
    except Exception:
        pass
    return ["gemini-3.8-flash"]


def analizar_sql_con_gemini(sql_query):
    """
    Llama a la API de Gemini sin Constrained Decoding (evita el 503 en SQLs largos)
    y rota entre los modelos reales activos de la cuenta si un servidor está lleno.
    """
    modelos_reales = obtener_modelos_activos()
    
    prompt = f"""
Eres un Analista de Datos experto. Analiza el siguiente SQL y extrae la información para un BRS.
Identifica únicamente los campos del SELECT principal.
Devuelve ÚNICAMENTE un objeto JSON válido con esta estructura exacta (sin texto adicional antes ni después):
{{
  "campos": [
    {{
      "columna": "nombre o alias del campo en el select principal",
      "termino_negocio": "explicación clara en lenguaje de negocio (ej. Es el ID de la Póliza)",
      "descripcion": "explicación detallada de la regla, lógica o CASE. Si es un campo directo sin lógica compleja, repite exactamente el valor de termino_negocio.",
      "fuente_datos": "nombre real de la tabla origen de donde nace este campo (resuelve los alias)"
    }}
  ],
  "tablas_origen": ["TABLA_ORIGEN_1", "TABLA_ORIGEN_2"]
}}

Consulta SQL:
{sql_query}
"""
    
    # IMPORTANTE: Quitamos "responseMimeType": "application/json" porque causa Error 503 
    # en los servidores de Google cuando el SQL de entrada es grande o complejo.
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1
        }
    }
    
    headers = {"Content-Type": "application/json"}
    
    # Mostramos en pequeño qué modelos reales detectó en tu cuenta por transparencia
    st.caption(f"🔍 Modelos disponibles detectados en tu cuenta: `{', '.join(modelos_reales)}`")
    
    max_intentos = max(4, len(modelos_reales) * 2)
    
    for intento in range(max_intentos):
        modelo_actual = modelos_reales[intento % len(modelos_reales)]
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo_actual}:generateContent?key={GEMINI_API_KEY}"
        
        response = requests.post(url, headers=headers, json=payload, timeout=120)
        
        if response.status_code == 200:
            resultado = response.json()
            texto_respuesta = resultado["candidates"][0]["content"]["parts"][0]["text"].strip()
            
            # Extraemos exactamente desde la primera llave '{' hasta la última '}'
            inicio = texto_respuesta.find("{")
            fin = texto_respuesta.rfind("}")
            if inicio != -1 and fin != -1:
                json_limpio = texto_respuesta[inicio:fin+1]
                return json.loads(json_limpio)
            else:
                raise Exception(f"La respuesta no contenía un JSON válido: {texto_respuesta[:200]}")
                
        elif response.status_code in [500, 502, 503, 504]:
            espera_servidor = 6
            st.warning(f"⚠️ `{modelo_actual}` con alta demanda (503). Cambiando al siguiente modelo disponible en {espera_servidor}s (Intento {intento+1}/{max_intentos})...")
            time.sleep(espera_servidor)
            
        elif response.status_code == 404:
            # Si algún modelo de la lista fue deprecado, lo descartamos
            continue
            
        elif response.status_code == 429:
            tiempo_espera = 22
            st.warning(f"⏳ Límite por minuto alcanzado en `{modelo_actual}`. Esperando {tiempo_espera}s...")
            time.sleep(tiempo_espera)
            
        else:
            raise Exception(f"Error HTTP {response.status_code} en {modelo_actual}: {response.text}")
            
    raise Exception("Todos los modelos disponibles están saturados en Google en este momento. Intenta de nuevo en un par de minutos.")

def copiar_estilo_fila(ws, fila_origen, fila_destino, max_col=5):
    """Copia bordes, colores, fuentes y alineación de la fila 10 a las filas nuevas."""
    for col in range(1, max_col + 1):
        celda_orig = ws.cell(row=fila_origen, column=col)
        celda_dest = ws.cell(row=fila_destino, column=col)
        if celda_orig.has_style:
            celda_dest.font = copy(celda_orig.font)
            celda_dest.border = copy(celda_orig.border)
            celda_dest.fill = copy(celda_orig.fill)
            celda_dest.number_format = copy(celda_orig.number_format)
            celda_dest.alignment = copy(celda_orig.alignment)


def llenar_plantilla_excel(datos_gemini, nombre_reporte, sql_texto):
    """Abre la plantilla, inyecta los datos conservando el formato y retorna los bytes."""
    wb = openpyxl.load_workbook("plantilla_brs.xlsx")
    
    campos = datos_gemini.get("campos", [])
    tablas = datos_gemini.get("tablas_origen", [])
    
    # ---------------------------------------------------------
    # HOJA: Mapeo BRS
    # ---------------------------------------------------------
    ws_mapeo = wb["Mapeo BRS"]
    
    if len(campos) > 1:
        ws_mapeo.insert_rows(idx=11, amount=len(campos) - 1)
        
    for i, campo in enumerate(campos):
        fila_actual = 10 + i
        if i > 0:
            copiar_estilo_fila(ws_mapeo, fila_origen=10, fila_destino=fila_actual, max_col=5)
            
        ws_mapeo.cell(row=fila_actual, column=1, value=nombre_reporte)                # Col A
        ws_mapeo.cell(row=fila_actual, column=2, value=campo.get("columna", ""))      # Col B
        ws_mapeo.cell(row=fila_actual, column=3, value=campo.get("termino_negocio", "")) # Col C
        ws_mapeo.cell(row=fila_actual, column=4, value=campo.get("descripcion", ""))  # Col D
        ws_mapeo.cell(row=fila_actual, column=5, value=campo.get("fuente_datos", "")) # Col E

    # ---------------------------------------------------------
    # HOJA: PreRequisitos
    # ---------------------------------------------------------
    if "PreRequisitos" in wb.sheetnames:
        ws_req = wb["PreRequisitos"]
        for i, tabla in enumerate(tablas):
            ws_req.cell(row=2 + i, column=3, value=tabla) # C2 hacia abajo
            
    # ---------------------------------------------------------
    # HOJA: Query
    # ---------------------------------------------------------
    if "Query" in wb.sheetnames:
        ws_query = wb["Query"]
        lineas_sql = sql_texto.splitlines()
        for i, linea in enumerate(lineas_sql):
            ws_query.cell(row=2 + i, column=1, value=linea) # A2 hacia abajo

    # ---------------------------------------------------------
    # HOJA: Layout
    # ---------------------------------------------------------
    if "Layout" in wb.sheetnames:
        ws_layout = wb["Layout"]
        for i, campo in enumerate(campos):
            ws_layout.cell(row=4 + i, column=1, value=campo.get("columna", "")) # A4 hacia abajo

    buffer_excel = io.BytesIO()
    wb.save(buffer_excel)
    return buffer_excel.getvalue()


def enviar_correo(email_dest, adjuntos):
    msg = EmailMessage()
    msg["Subject"] = "Tus documentos BRS generados automáticamente"
    msg["From"] = EMAIL_SENDER
    msg["To"] = email_dest
    msg.set_content("Hola,\n\nAdjunto encontrarás los archivos BRS generados a partir de tus consultas SQL y tu plantilla base.\n\nSaludos.")

    for nombre_archivo, bytes_excel in adjuntos:
        msg.add_attachment(
            bytes_excel, 
            maintype="application", 
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            filename=nombre_archivo
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)


# 4. Lógica al presionar el botón del formulario
if boton_procesar:
    faltan_nombres = any(not nombre.strip() for nombre in nombres_tablas.values())
    
    if not email_destino.strip():
        st.warning("⚠️ Por favor, ingresa un correo destino.")
    elif faltan_nombres:
        st.warning("⚠️ Debes asignar un Nombre de Tabla para cada archivo subido antes de procesar.")
    else:
        with st.spinner("Procesando consultas e inyectando datos en tu plantilla..."):
            archivos_excel_generados = []
            
            for index, archivo in enumerate(archivos_subidos):
                nombre_reporte = nombres_tablas[archivo.name].strip()
                st.write(f"⚙️ Procesando ({index + 1}/{len(archivos_subidos)}): **{archivo.name}** ➔ Tabla: `{nombre_reporte}`")
                
                # decode con 'utf-8-sig' o 'latin-1' de respaldo por si el .txt tiene acentos de Windows
                bytes_archivo = archivo.getvalue()
                try:
                    contenido_sql = bytes_archivo.decode("utf-8-sig")
                except UnicodeDecodeError:
                    contenido_sql = bytes_archivo.decode("latin-1")
                
                try:
                    # 1. Extraer JSON desde Gemini
                    datos_extraidos = analizar_sql_con_gemini(contenido_sql)
                    
                    # 2. Llenar plantilla Excel
                    excel_bytes = llenar_plantilla_excel(datos_extraidos, nombre_reporte, contenido_sql)
                    
                    # 3. Guardar en la lista de envío
                    nombre_final = f"BRS_{nombre_reporte.replace(' ', '_')}.xlsx"
                    archivos_excel_generados.append((nombre_final, excel_bytes))
                    st.success(f"✅ Plantilla generada para {archivo.name}")
                    
                    # Si subiste MÁS de 1 archivo, esperamos 15 seg entre cada uno para respetar el límite de 5/min
                    if index < len(archivos_subidos) - 1:
                        st.info("⏱️ Esperando 15 segundos antes del siguiente archivo para respetar el límite gratuito...")
                        time.sleep(15)
                        
                except Exception as e:
                    st.error(f"❌ Error procesando el archivo {archivo.name}: {e}")
            
            if archivos_excel_generados:
                st.write("📧 Enviando correo con los archivos adjuntos...")
                try:
                    enviar_correo(email_destino.strip(), archivos_excel_generados)
                    st.balloons()
                    st.success(f"🎉 ¡Listo! Se enviaron {len(archivos_excel_generados)} archivos BRS a **{email_destino}**.")
                except Exception as e_mail:
                    st.error(f"❌ Los archivos se generaron, pero falló el envío de correo: {e_mail}")
