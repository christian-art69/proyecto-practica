import os
import smtplib
import pandas as pd # Necesario para leer Excel
from email.mime.text import MIMEText
from datetime import datetime, timedelta

# --- CONFIGURACIÓN DE ENVÍO Y VARIABLES DE ENTORNO ---
# (Sin cambios)
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
# --- FIN CONFIGURACIÓN ---

# Nombre del archivo de Excel
EXCEL_FILE_PATH = 'alumnos_tareas.xlsx' 


# --- FUNCIÓN: CARGAR DATOS DESDE EXCEL ---

def load_students_from_excel(file_path):
    """Carga los datos de los alumnos desde un archivo Excel/CSV."""
    try:
        # Intentar leer el archivo Excel (.xlsx)
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        # O intentar leer el archivo CSV
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            print("Error: El archivo debe ser .xlsx o .csv")
            return []
            
        # 1. Limpiar y estandarizar nombres de columnas
        df.columns = df.columns.str.lower()
        
        # 2. Asegurarse de que las columnas críticas existan
        required_cols = ['nombre', 'email', 'vencimiento']
        if not all(col in df.columns for col in required_cols):
            print(f"Error: El archivo Excel debe contener las columnas: {required_cols}")
            return []

        alumnos_list = []
        
        # 3. Iterar sobre las filas del DataFrame y crear la estructura de datos
        for index, row in df.iterrows():
            # Asumimos que la tarea siempre es la misma ('Entrega Final del Curso (Adulto Mayor)')
            # y que la columna 'vencimiento' contiene la fecha límite.
            alumno_data = {
                "id": index + 1, # ID autogenerado
                "nombre": row['nombre'],
                "email": row['email'],
                "tareas_pendientes": [
                    {
                        "nombre": "Entrega Final del Curso (Adulto Mayor)",
                        "vencimiento": str(row['vencimiento']).split(' ')[0], # Asegura formato 'YYYY-MM-DD' si viene con hora
                        "entregado": False
                    }
                ]
            }
            alumnos_list.append(alumno_data)
            
        print(f"✅ Datos cargados exitosamente para {len(alumnos_list)} alumnos.")
        return alumnos_list
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {file_path}. Usando lista vacía.")
        return []
    except Exception as e:
        print(f"Error al procesar el archivo Excel: {e}")
        return []



# --- FUNCIÓN DE ENVÍO DE CORREO ÚNICA (sin cambios) ---

def send_email_reminder(to_email, subject, body):
    """Envía un correo electrónico a través de SMTP."""
    # ... (código send_email_reminder)
    try:
        msg = MIMEText(body, 'html')
        msg['Subject'] = subject
        msg['From'] = EMAIL_USER 
        msg['To'] = to_email

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls() 
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, to_email, msg.as_string())
        
        print(f"Correo enviado exitosamente a {to_email} desde {EMAIL_USER}")
        return True
    except Exception as e:
        print(f"Error al enviar correo a {to_email}: {e}")
        print("Verifica si el servidor SMTP, el puerto o las credenciales son correctas.")
        return False


# --- LÓGICA PRINCIPAL: Decisión por Fecha (CON CAMBIOS EN ASUNTO Y CUERPO) ---

def main_reminder_logic():
    """Itera sobre la lista de alumnos CREADA DESDE EXCEL y genera recordatorios."""
    
    # *** ESTE ES EL CAMBIO CLAVE: CARGAR LA LISTA DESDE EL ARCHIVO ***
    ALUMNOS_A_MONITOREAR = load_students_from_excel(EXCEL_FILE_PATH)
    
    if not ALUMNOS_A_MONITOREAR:
        print("No hay alumnos para monitorear. Finalizando proceso.")
        return
    
    # ... (El resto de la lógica sigue igual)
    print(f"Iniciando chequeo de recordatorios. Fecha actual: {datetime.now().strftime('%Y-%m-%d')}")
    
    hoy = datetime.now().date()

    for alumno in ALUMNOS_A_MONITOREAR:
        nombre = alumno['nombre']
        email = alumno['email']
        
        tareas_para_recordar = []

        for tarea in alumno.get('tareas_pendientes', []):
            if tarea.get('entregado', False): 
                continue

            try:
                fecha_vencimiento = datetime.strptime(tarea['vencimiento'], '%Y-%m-%d').date()
            except ValueError:
                print(f"Advertencia: Formato de fecha inválido para {tarea['nombre']}")
                continue

            # Lógica de urgencia (HOY o TARDE)
            estado = None
            if fecha_vencimiento == hoy:
                estado = "**¡VENCE HOY!**"  # Más enfático
            elif fecha_vencimiento < hoy:
                # Modificado para ser más asertivo
                estado = f"**VENCIDA** (Fecha límite: {tarea['vencimiento']})"
            
            if estado:
                tareas_para_recordar.append((tarea['nombre'], estado))
            
        if tareas_para_recordar:
            print(f"--> {nombre}: ¡Tiene {len(tareas_para_recordar)} tarea(s) pendientes/vencidas!")
            
            # --- MODIFICACIÓN DEL CUERPO DEL CORREO ---
            lista_tareas_str = "\n".join([f"- {t[0]} ({t[1]})" for t in tareas_para_recordar])

            # ASUNTO MODIFICADO: Directo y urgente.
            subject = f"⚠️ ¡Cuidado con la Fecha Límite! Tarea(s) Pendiente(s) o Vencida(s)"

            # CUERPO DEL CORREO MODIFICADO: Enfocado en no pasarse de la fecha.
            email_body = f"""
            <html><body>
                <p>Estimado(a) **{nombre}**:</p>
                <p>Este es un **AVISO URGENTE** para asegurar que no te pases de la fecha límite o para informarte que ya ha pasado. **La entrega de tu trabajo final es crítica para la aprobación del curso.**</p>
                <p>Asegúrate de enviar las siguientes tareas **INMEDIATAMENTE**: </p>
                <pre>{lista_tareas_str}</pre>
                
                <p>
                    **Si la tarea vence hoy**, no pospongas la entrega para evitar penalizaciones. 
                    **Si la tarea ya está vencida**, por favor contáctanos lo antes posible para ver cómo puedes regularizar tu situación.
                </p>
                
                <p>**Si ya realizaste la entrega, por favor, ignora este mensaje.**</p>
                
                <p>Saludos y mucho éxito,<br>El equipo de {EMAIL_USER.split('@')[1]}</p>
            </body></html>
            """
            
            send_email_reminder(email, subject, email_body)
            
    print("Proceso de recordatorios finalizado.")

# --- PUNTO DE ENTRADA DEL SCRIPT ---


if __name__ == "__main__":
    main_reminder_logic()
