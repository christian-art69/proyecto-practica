import os
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from datetime import datetime, timedelta

# --- CONFIGURACIÓN DE ENVÍO Y VARIABLES DE ENTORNO ---

SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')

# --- CONFIGURACIÓN PARA REPORTES DE ERRORES ---
ADMIN_EMAIL = os.getenv('ADMIN_EMAIL') 

# Nombre del archivo de Excel
EXCEL_FILE_PATH = 'alumnos_tareas.xlsx' 

# --- FUNCIÓN DE ALERTA DE ADMINISTRADOR ---

def send_admin_alert(subject, body):
    """Envía una notificación simple al administrador sobre un error crítico."""
    if not ADMIN_EMAIL or not EMAIL_USER or not EMAIL_PASSWORD:
        print(f"ALERTA ADMINISTRATIVA FALLIDA: ADMIN_EMAIL o credenciales no configuradas.")
        return
        
    try:
        msg = MIMEText(body, 'text')
        msg['Subject'] = f"[ALERTA CRÍTICA] {subject}"
        msg['From'] = EMAIL_USER 
        msg['To'] = ADMIN_EMAIL

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls() 
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_USER, ADMIN_EMAIL, msg.as_string())
        
        print(f"✅ Alerta de error enviada exitosamente a {ADMIN_EMAIL}")
    except Exception as e:
        print(f"🔴 ERROR FATAL: No se pudo enviar la alerta al administrador. {e}")
        
# --- FUNCIÓN: CARGAR DATOS DESDE EXCEL ---

def load_students_from_excel(file_path):
    """Carga los datos de los alumnos desde un archivo Excel/CSV."""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        elif file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            error_msg = "Error: El archivo debe ser .xlsx o .csv"
            print(error_msg)
            send_admin_alert("FORMATO DE ARCHIVO INVÁLIDO", error_msg)
            return []
            
        df.columns = df.columns.str.lower()
        required_cols = ['nombre', 'email', 'vencimiento']
        if not all(col in df.columns for col in required_cols):
            error_msg = f"Error: El archivo Excel debe contener las columnas: {required_cols}. Columnas encontradas: {list(df.columns)}"
            print(error_msg)
            send_admin_alert("COLUMNAS FALTANTES EN EXCEL/CSV", error_msg)
            return []

        alumnos_list = []
        
        for index, row in df.iterrows():
            alumno_data = {
                "id": index + 1, 
                "nombre": row['nombre'],
                "email": row['email'],
                "tareas_pendientes": [
                    {
                        "nombre": "Entrega Final del Curso (Adulto Mayor)",
                        "vencimiento": str(row['vencimiento']).split(' ')[0],
                        "entregado": False
                    }
                ]
            }
            alumnos_list.append(alumno_data)
            
        print(f"✅ Datos cargados exitosamente para {len(alumnos_list)} alumnos.")
        return alumnos_list
        
    except FileNotFoundError:
        error_msg = f"Error: No se encontró el archivo {file_path}. El script no puede continuar."
        print(error_msg)
        send_admin_alert("ARCHIVO DE DATOS NO ENCONTRADO", error_msg)
        return []
    except Exception as e:
        error_msg = f"Error crítico al procesar el archivo Excel: {e}"
        print(error_msg)
        send_admin_alert("ERROR CRÍTICO EN LA LECTURA DE EXCEL/CSV", error_msg)
        return []
        
# --- FUNCIÓN DE ENVÍO DE CORREO ÚNICA (Manejo de Errores de SMTP) ---

def send_email_reminder(to_email, subject, body):
    """Envía un correo electrónico a través de SMTP."""
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
        error_msg = f"Error al enviar correo al alumno {to_email}: {e}"
        print(error_msg)
        
        # Notificación al Administrador sobre el fallo
        admin_body = f"FALLO DE ENVÍO SMTP:\n\n{error_msg}\n\nRevisa la configuración de EMAIL_USER y EMAIL_PASSWORD."
        send_admin_alert("FALLO DE ENVÍO DE CORREO A ESTUDIANTE", admin_body)

        return False
        
# --- LÓGICA PRINCIPAL: Decisión por Fecha (CON NUEVO ENFOQUE) ---

def main_reminder_logic():
    """Itera sobre la lista de alumnos CREADA DESDE EXCEL y genera recordatorios."""
    
    ALUMNOS_A_MONITOREAR = load_students_from_excel(EXCEL_FILE_PATH)
    
    if not ALUMNOS_A_MONITOREAR:
        print("No hay alumnos para monitorear. Finalizando proceso.")
        return
    
    print(f"Iniciando chequeo de plazos. Fecha actual: {datetime.now().strftime('%Y-%m-%d')}")
    
    hoy = datetime.now().date()
    data_warnings = [] 

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
                warning_msg = f"Advertencia: Formato de fecha inválido para el alumno {nombre} en la tarea {tarea['nombre']} con valor '{tarea['vencimiento']}'. Saltando tarea."
                print(warning_msg)
                data_warnings.append(warning_msg) 
                continue

            # Lógica de Plazo Límite
            estado = None
            if fecha_vencimiento == hoy:
                estado = "**¡PLAZO FINAL HOY!**"  # Énfasis en el plazo
            elif fecha_vencimiento < hoy:
                estado = f"**PLAZO EXPIRADO** (Fecha límite: {tarea['vencimiento']})" # Énfasis en que ya pasó
            
            if estado:
                tareas_para_recordar.append((tarea['nombre'], estado))
                
        if tareas_para_recordar:
            print(f"--> {nombre}: ¡Tiene {len(tareas_para_recordar)} plazos críticos!")
            
            lista_tareas_str = "\n".join([f"- {t[0]} ({t[1]})" for t in tareas_para_recordar])

            subject = f"🚨 URGENTE: Notificación sobre el Plazo Final del Curso"
            email_body = f"""
            <html><body>
                <p>Estimado(a) **{nombre}**:</p>
                <p>Este es un **AVISO IMPORTANTE** para informarte sobre el estado de tu **Plazo Final del Curso**. La entrega de este trabajo es **CRÍTICA** para la aprobación.</p>
                <p>El estado actual de tu fecha límite es el siguiente: </p>
                <pre>{lista_tareas_str}</pre>
                <p>
                    **Si tu plazo final es hoy**, por favor, no demores la entrega para evitar la expiración del plazo. 
                    **Si el plazo ha expirado**, contáctanos de inmediato para regularizar tu situación.
                </p>
                <p>**Si ya realizaste la entrega, por favor, ignora este mensaje.**</p>
                <p>Saludos y mucho éxito,<br>El equipo de {EMAIL_USER.split('@')[1]}</p>
            </body></html>
            """
            send_email_reminder(email, subject, email_body)

    # Reporte de advertencias de datos al final del proceso
    if data_warnings:
        admin_body = "Se detectaron los siguientes problemas de formato de datos en el archivo de alumnos:\n\n"
        admin_body += "\n".join(data_warnings)
        admin_body += "\n\nPor favor, revisa el formato 'YYYY-MM-DD' en el archivo Excel o CSV."
        send_admin_alert("ADVERTENCIAS DE FORMATO DE DATOS EN EXCEL/CSV", admin_body)
        
    print("Proceso de recordatorios finalizado.")

# ----------------------------------------------------------------------
# --- PUNTO DE ENTRADA DEL SCRIPT ---
# ----------------------------------------------------------------------

if __name__ == "__main__":
    main_reminder_logic()
