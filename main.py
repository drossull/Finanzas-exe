import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
import os
from dotenv import load_dotenv

# ==========================================
# 1. CONFIGURACIÓN Y SEGURIDAD
# ==========================================
load_dotenv()

USUARIO = os.getenv("EMAIL_USUARIO")
CLAVE_APP = os.getenv("EMAIL_CLAVE")

# ¡ACTUALIZADO! El remitente correcto según tu imagen
BANCO_SENDER = "enviodigital@bancochile.cl"

def limpiar_monto(texto_monto):
    """Convierte un texto de monto en un número entero de forma segura."""
    if texto_monto is None:
        return 0
    # Quitamos puntos y dejamos solo los números
    solo_numeros = re.sub(r'[^\d]', '', texto_monto)
    return int(solo_numeros) if solo_numeros else 0

def obtener_texto_limpio(msg):
    """Extrae el texto del correo, limpiando el formato HTML si es necesario."""
    texto = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            # A veces viene como texto plano, a veces como HTML
            if content_type == "text/plain":
                texto += part.get_payload(decode=True).decode('utf-8', errors='ignore')
            elif content_type == "text/html":
                html = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                # Quitar etiquetas HTML (<br>, <div>, etc)
                texto += re.sub(r'<[^>]+>', ' ', html)
    else:
        payload = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        if msg.get_content_type() == "text/html":
            texto = re.sub(r'<[^>]+>', ' ', payload)
        else:
            texto = payload
    
    # Reducir saltos de línea y espacios gigantes a un solo espacio
    return re.sub(r'\s+', ' ', texto)

def procesar_cuerpo_correo(cuerpo_limpio):
    """Busca los patrones específicos del correo 'Cargo en Cuenta' de Banco de Chile."""
    
    # 1. Buscar el Monto (Ej: "una compra por $13.208")
    monto_match = re.search(r'por\s*\$\s*([\d.]+)', cuerpo_limpio, re.IGNORECASE)
    monto_final = limpiar_monto(monto_match.group(1)) if monto_match else 0

    # 2. Buscar Concepto y Fecha (Ej: "en BAC TOBALABA I el 07/05/2026")
    # Busca la palabra "en", captura lo que sigue (el lugar), hasta encontrar "el" y una fecha (dd/mm/yyyy)
    concepto_match = re.search(r'en\s+(.*?)\s+el\s+([\d/]{10})', cuerpo_limpio, re.IGNORECASE)
    
    concepto = "Cargo en Cuenta"
    fecha = "Sin fecha"
    
    if concepto_match:
        concepto = concepto_match.group(1).strip()
        fecha = concepto_match.group(2).strip()

    if monto_final > 0:
        return {
            "Fecha": fecha,
            "Lugar/Concepto": concepto[:60], # Guardamos hasta 60 letras del lugar
            "Monto": monto_final,
            "Tipo": "Gasto"
        }
    return None

def extraer_correos():
    """Conexión principal y recuperación de datos."""
    registros = []

    if not USUARIO or not CLAVE_APP:
        print("❌ Error: Credenciales no encontradas en el archivo .env")
        return []

    try:
        print(f"📧 Conectando a {USUARIO}...")
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(USUARIO, CLAVE_APP)
        mail.select("inbox")

        print(f"🔍 Buscando correos de: {BANCO_SENDER}...")
        search_query = f'FROM "{BANCO_SENDER}"'
        status, mensajes = mail.search(None, search_query)
        
        if status != 'OK' or not mensajes[0]:
            print("ℹ️ No se encontraron correos de este remitente.")
            return []

        ids = mensajes[0].split()
        print(f"✅ ¡Se encontraron {len(ids)} correos!")
        
        # Procesaremos los últimos 20 correos
        cantidad_a_procesar = min(20, len(ids))
        print(f"🔎 Extrayendo datos de los últimos {cantidad_a_procesar} correos...")

        for i in ids[-cantidad_a_procesar:]:
            res, msg_data = mail.fetch(i, "(RFC822)")
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    
                    # 1. Obtener el texto limpio
                    texto_limpio = obtener_texto_limpio(msg)
                    
                    # 2. Procesar los datos
                    if texto_limpio:
                        resultado = procesar_cuerpo_correo(texto_limpio)
                        if resultado:
                            registros.append(resultado)

        mail.logout()
        print("✅ Proceso de extracción terminado.")
        return registros

    except Exception as e:
        print(f"❌ Error crítico en la conexión: {e}")
        return []

# ==========================================
# 3. EJECUCIÓN
# ==========================================
if __name__ == "__main__":
    print("🚀 Iniciando Gestor de Finanzas...")
    datos_finales = extraer_correos()

    if datos_finales:
        df = pd.DataFrame(datos_finales)
        print("\n--- 📊 RESUMEN DE COMPRAS ENCONTRADAS ---")
        print(df.to_string(index=False)) # Imprime la tabla más limpia en la consola
        
        # Guardamos en Excel local
        nombre_archivo = "Mis_Gastos_BancoChile.xlsx"
        df.to_excel(nombre_archivo, index=False)
        print(f"\n📁 Archivo '{nombre_archivo}' actualizado exitosamente.")
    else:
        print("\n⚠️ No se pudo extraer información. Revisa que haya compras registradas.")