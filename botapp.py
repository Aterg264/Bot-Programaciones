import streamlit as st
import json
import io
from docx import Document
from google import genai #[cite: 1]

# 1. Configuración de la interfaz web
st.title("Generador Conversacional de Programaciones")
st.write("Habla con el asistente. Cuando tenga todos los datos, generará tu documento Word.")

# 2. Conexión con Gemini
# Asegúrate de poner tu clave real o usar st.secrets si lo subes a la nube
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY) #[cite: 1]

# 3. Instrucciones del Bot (Su "Cerebro")
instrucciones = """
Eres un asistente amable. Tu objetivo es pedir al usuario los datos para un curso:
Nombre certificado, Título MF, N.º Acción, N.º Grupo, Horas Totales, Responsable, Modalidad, Fecha Inicio, Fecha Fin, Examen presencial y Objetivos.
Pregunta conversacionalmente. Cuando tengas TODOS los datos, tu ÚNICA respuesta debe ser un diccionario JSON válido con las claves:
{"{{NOMBRE_CERTIFICADO}}": "...", "{{TÍTULO_MF}}": "...", "{{ACCION}}": "...", "{{GRUPO}}": "...", "{{H_TOTAL}}": "...", "{{RESP}}": "...", "{{MODALIDAD}}": "...", "{{F_INICIO}}": "...", "{{F_FINAL}}": "...", "{{F_EXAMEN}}": "...", "{{OBJETIVOS}}": "..."}
"""

# Inicializar el chat en la web
if "mensajes" not in st.session_state:
    st.session_state.mensajes = [{"role": "model", "content": "¡Hola! ¿De qué curso vamos a hacer la programación hoy?"}]

# Mostrar el historial
for msg in st.session_state.mensajes:
    st.chat_message(msg["role"]).write(msg["content"])

# 4. Capturar respuesta del usuario
if prompt := st.chat_input("Escribe tu respuesta aquí..."):
    st.session_state.mensajes.append({"role": "user", "content": prompt})
    st.chat_message("user").write(prompt)

    # 5. Enviar el historial a Gemini
    historial = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.mensajes])
    
    response = client.models.generate_content( #[cite: 1]
        model="gemini-3.6-flash", #[cite: 1]
        contents=historial,
        config={
            "system_instruction": instrucciones #[cite: 1]
        }
    )
    
    respuesta_bot = response.text
    
    # 6. Magia: Si el bot devuelve el JSON, creamos el Word
    if "{" in respuesta_bot and "{{NOMBRE_CERTIFICADO}}" in respuesta_bot:
        st.success("¡Datos recopilados! Generando documento...")
        
        # Extraer el JSON limpio
        inicio, fin = respuesta_bot.find('{'), respuesta_bot.rfind('}') + 1
        datos_curso = json.loads(respuesta_bot[inicio:fin])
        
        try:
            doc = Document("PlantillaPython-PF.docx")
            
            # Reemplazar etiquetas
            for p in doc.paragraphs:
                for k, v in datos_curso.items():
                    if k in p.text: p.text = p.text.replace(k, str(v))
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        for p in c.paragraphs:
                            for k, v in datos_curso.items():
                                if k in p.text: p.text = p.text.replace(k, str(v))
                                
            # Preparar el archivo para descargar en la web
            archivo_descarga = io.BytesIO()
            doc.save(archivo_descarga)
            
            st.download_button(
                label="📥 Descargar Programación Word",
                data=archivo_descarga.getvalue(),
                file_name=f"Programa_{datos_curso.get('{{ACCION}}', 'Generado')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error("Error leyendo la plantilla. Asegúrate de que 'PlantillaPython-PF.docx' está en la misma carpeta.")
    else:
        # Si no es JSON, sigue la conversación normal
        st.session_state.mensajes.append({"role": "model", "content": respuesta_bot})
        st.chat_message("model").write(respuesta_bot)
