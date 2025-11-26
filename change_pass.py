import pyodbc
from contextlib import closing
import random
import smtplib  # Para enviar el correo mediante SMTP
from email.mime.text import MIMEText  # Para mensajes de texto plano
from email.mime.multipart import MIMEMultipart  # Para mensajes con HTML y/o adjuntos

connection_strings = [
    {"dsn": "conexion-pixel-ambato", "uid": "DBA", "pwd": "banana1","email":"rguaman@latablitadeltartaro.com"},
    {"dsn": "conexion-pixel-colon", "uid": "DBA", "pwd": "banana1","email":"rguaman@latablitadeltartaro.com"},
    {"dsn": "conexion-pixel-floreana", "uid": "DBA", "pwd": "banana1","email":"rguaman@latablitadeltartaro.com"}
]

def execute_query_on_dsn(config):
    try:

        sql_query = """SELECT 
        EMPNUM,
        EMPNAME,
        ADRESS2,
        POSNAME,
        EMPPOSITION
        FROM employee
        WHERE EMPPOSITION = 1005 AND ISACTIVE = 1;
        """
        connection_string = f"DSN={config['dsn']};UID={config['uid']};PWD={config['pwd']}"

        with pyodbc.connect(connection_string) as conn:
            with closing(conn.cursor()) as cursor:
                cursor.execute(sql_query)
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                return {
                    "dsn": config['dsn'],
                    "columns": columns,
                    "data": rows,
                    "error": None
                }
    except pyodbc.Error as e:
        return {
            "dsn": config['dsn'],
            "columns": [],
            "data": [],
            "error": str(e)
        }


def change_password_cashier(newpass,idEmploye,config):
    try:

        sql_query = """UPDATE employee
                       SET SWIPE = ?
                       WHERE ISACTIVE = 1 and EMPNUM = ? and EMPPOSITION = 1005;"""
        connection_string = f"DSN={config['dsn']};UID={config['uid']};PWD={config['pwd']}"

        with pyodbc.connect(connection_string) as conn:
            with closing(conn.cursor()) as cursor:
                cursor.execute(sql_query,(newpass,idEmploye))
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                return {
                    "dsn": config['dsn'],
                    "columns": columns,
                    "data": rows,
                    "error": None
                }
    except pyodbc.Error as e:
        return {
            "dsn": config['dsn'],
            "columns": [],
            "data": [],
            "error": str(e)
        }

def sendCorreo():
    smtp_server = "smtp.office365.com"
    smtp_port = 587  # Para TLS
    email_remitente = "pedidocompra@latablitadeltartaro.com"
    password = "L@T@b1ita2E"
    
    # Primero agrupamos las conexiones por email
    conexiones_por_email = {}
    for config in connection_strings:
        email = config['email']
        if email not in conexiones_por_email:
            conexiones_por_email[email] = []
        conexiones_por_email[email].append(config)
    
    # Ahora procesamos cada grupo de conexiones por email
    for email, configs in conexiones_por_email.items():
        # Preparar el mensaje para este email
        mensaje = MIMEMultipart()
        mensaje["From"] = email_remitente
        mensaje["To"] = email
        mensaje["Subject"] = "Cambio Claves Mensual de Administradores"
        
        cuerpo_total = ""
        tiene_datos = False
        errores = False
        
        # Procesar cada conexión para este email
        for config in configs:
            result = execute_query_on_dsn(config)
            
            if result['error']:
                # Si hay error, agregar al cuerpo
                cuerpo_total += f"\nError en conexión {config['dsn']}:\n"
                cuerpo_total += f"{result['error']}\n"
                cuerpo_total += "-" * 50 + "\n"
                errores = True
            else:
                # Procesar datos si no hay error
                empleados_info = []
                for empleado in result['data']:
                    newPass = random.randint(100000, 999999)
                    empleado_info = {
                        'EMPNUM': empleado[0],
                        'EMPNAME': empleado[1],
                        'ADRESS2': empleado[2],
                        'POSNAME': empleado[3],
                        'newPass': newPass
                    }
                    empleados_info.append(empleado_info)
                    print(f"Administrador: {empleado[1]}, Nueva Clave: {newPass}")
                    change_password_cashier(newPass,empleado[0],config);

                
                # Agregar información al cuerpo del correo
                cuerpo_total += f"\nReporte para {config['dsn']}:\n"
                cuerpo_total += "=" * 50 + "\n"
                
                if empleados_info:
                    tiene_datos = True
                    for emp in empleados_info:
                        cuerpo_total += f"Nombre: {emp['EMPNAME']}\n"
                        cuerpo_total += f"Cedula: {emp['ADRESS2']}\n"
                        cuerpo_total += f"Nueva contraseña: {emp['newPass']}\n"
                        cuerpo_total += "-" * 50 + "\n"
                else:
                    cuerpo_total += "No se encontraron administradores para esta sucursal\n"
                
                cuerpo_total += "\n"
        
        # Determinar el asunto basado en el contenido
        if errores and not tiene_datos:
            mensaje["Subject"] = "Errores en conexiones a sucursales"
        elif errores and tiene_datos:
            mensaje["Subject"] = "Reporte de Administradores con Errores en algunas Sucursales"
        elif not tiene_datos:
            mensaje["Subject"] = "No se encontraron administradores en las sucursales"
        
        # Adjuntar el cuerpo al mensaje
        mensaje.attach(MIMEText(cuerpo_total, 'plain'))
        
        # Enviar el correo solo si hay contenido (datos o errores)
        if cuerpo_total.strip():
            try:
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(email_remitente, password)
                    server.sendmail(email_remitente, email, mensaje.as_string())
                print(f"Correo enviado exitosamente a {email} para {len(configs)} conexiones")
            except Exception as e:
                print(f"Error al enviar correo a {email}: {str(e)}")
        else:
            print(f"No hay contenido para enviar a {email}")


sendCorreo();

