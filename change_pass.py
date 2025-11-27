import pyodbc
from contextlib import closing
import random
import smtplib  # Para enviar el correo mediante SMTP
from email.mime.text import MIMEText  # Para mensajes de texto plano
from email.mime.multipart import MIMEMultipart  # Para mensajes con HTML y/o adjuntos

connection_strings = [

    #JEFE ZONAL SILVANA CURAY   
    # {"dsn": "conexion-pixel-quisur", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"quisur"},
    # {"dsn": "conexion-pixel-rpatio", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"rpatio"},
    # {"dsn": "conexion-pixel-recreotres", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"recreotres"},
    # {"dsn": "conexion-pixel-rplaza", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"rplaza"},
    # {"dsn": "conexion-pixel-sanluis", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"sanluis"},
    # {"dsn": "conexion-pixel-sanrafael", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"sanrafael"},
    # {"dsn": "conexion-pixel-villaflora", "uid": "DBA", "pwd": "banana1","email":"scuray@latablitadeltartaro.com","local":"villaflora"}
    
    # #JEFE ZONAL JAVIER QUICHIMBO
    {"dsn": "conexion-pixel-bosque", "uid": "DBA", "pwd": "banana1","email":"jquichimbo@latablitadeltartaro.com","local":"bosque"},
    {"dsn": "conexion-pixel-cci", "uid": "DBA", "pwd": "banana1","email":"jquichimbo@latablitadeltartaro.com","local":"cci"},
    {"dsn": "conexion-pixel-colon", "uid": "DBA", "pwd": "banana1","email":"jquichimbo@latablitadeltartaro.com","local":"colon"},
    {"dsn": "conexion-pixel-jardin", "uid": "DBA", "pwd": "banana1","email":"jquichimbo@latablitadeltartaro.com","local":"jardin"},
    {"dsn": "conexion-pixel-quicentro", "uid": "DBA", "pwd": "banana1","email":"jquichimbo@latablitadeltartaro.com","local":"quicentro"},
    {"dsn": "conexion-pixel-riocentroquito", "uid": "DBA", "pwd": "banana1","email":"jquichimbo@latablitadeltartaro.com","local":"riocentroquito"},

    # #JEFE ZONAL KARLA AGUIRRE
    # {"dsn": "conexion-pixel-mega", "uid": "DBA", "pwd": "banana1","email":"kaguirre@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-portaldos", "uid": "DBA", "pwd": "banana1","email":"kaguirre@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-portal", "uid": "DBA", "pwd": "banana1","email":"kaguirre@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-scala", "uid": "DBA", "pwd":  "banana1","email":"kaguirre@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-republica", "uid": "DBA", "pwd": "banana1","email":"kaguirre@latablitadeltartaro.com"},
    
    # #JEFE ZONAL ENRIQUE MUÑOZ
    # {"dsn": "conexion-pixel-floreana", "uid": "DBA", "pwd": "banana1","email":"emunioz@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-granados", "uid": "DBA", "pwd": "banana1","email":"emunioz@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-ibarra", "uid": "DBA", "pwd": "banana1","email":"emunioz@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-laguna", "uid": "DBA", "pwd": "banana1","email":"emunioz@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-prensa", "uid": "DBA", "pwd": "banana1","email":"emunioz@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-wok-quicentro", "uid": "DBA", "pwd": "banana1","email":"emunioz@latablitadeltartaro.com"},
    
    #JEFE ZONAL PABLO ROMERO
    # {"dsn": "conexion-pixel-latacunga", "uid": "DBA", "pwd": "banana1","email":"promero@latablitadeltartaro.com"},
    {"dsn": "conexion-pixel-ambato", "uid": "DBA", "pwd": "banana1","email":"emedranda@latablitadeltartaro.com"}
    # {"dsn": "conexion-pixel-riobamba", "uid": "DBA", "pwd": "banana1","email":"promero@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-santodomingo", "uid": "DBA", "pwd": "banana1","email":"promero@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-bomboli", "uid": "DBA", "pwd": "banana1","email":"promero@latablitadeltartaro.com"},

    # #JEFE ZONAL MICHELLE BENITEZ
    # {"dsn": "conexion-pixel-nueveoctubre", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-sonesta", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-dorado", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-machala", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-piazzamachala", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-urdesados", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-centenario", "uid": "DBA", "pwd": "banana1","email":"mbenitez@latablitadeltartaro.com"},

    # #JEFE ZONAL LENNY RODRIGUEZ
    # {"dsn": "conexion-pixel-mallsur", "uid": "DBA", "pwd": "banana1","email":"lrodriguez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-riosur", "uid": "DBA", "pwd": "banana1","email":"lrodriguez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-village", "uid": "DBA", "pwd": "banana1","email":"lrodriguez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-entrerios", "uid": "DBA", "pwd": "banana1","email":"lrodriguez@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-salinas", "uid": "DBA", "pwd": "banana1","email":"lrodriguez@latablitadeltartaro.com"},
    
    # #JEFE ZONAL BYAN ZAMORA
    # {"dsn": "conexion-pixel-ceibos", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-sanmarino", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-duran", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-mallnorte", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-rionorte", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-citymall", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"},
    # {"dsn": "conexion-pixel-mallpacifico", "uid": "DBA", "pwd": "banana1","email":"bzamora@latablitadeltartaro.com"}
]

def execute_query_on_dsn(config):
    try:

        sql_query = """SELECT 
        EMPNUM,
        EMPNAME,
        ADRESS2,
        POSNAME,
        EMPPOSITION,
        EmpLastName
        FROM DBA.employee
        WHERE EMPPOSITION = 1008 AND ISACTIVE = 1 AND EMPNUM =2146435125;
        """
        connection_string = f"DSN={config['dsn']};UID={config['uid']};PWD={config['pwd']}"

        with pyodbc.connect(connection_string) as conn:
            with closing(conn.cursor()) as cursor:
                cursor.execute(sql_query)
                rows = cursor.fetchall()
                return {
                    "dsn": config['dsn'],
                    "columns": [],
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


def change_password_cashier(newpass, idEmploye, config):
    try:
        sql_query = """ UPDATE DBA.employee
                        SET SWIPE = ?
                        WHERE ISACTIVE = 1 AND EMPNUM = ? AND EMPPOSITION = 1008 """

        connection_string = (
            f"DSN={config['dsn']};UID={config['uid']};PWD={config['pwd']}"
        )

        with pyodbc.connect(connection_string) as conn:
            with closing(conn.cursor()) as cursor:
                cursor.execute(sql_query, (newpass, idEmploye))

                #AQUI va el commit PARA ASEGURAR ACTUALIZACIÓN

                # rowcount = cuántas filas se actualizaron
                filas_afectadas = cursor.rowcount

                conn.commit()
                return {
                    "dsn": config["dsn"],
                    "rows_affected": filas_afectadas,
                    "error": None,
                }

    except pyodbc.Error as e:
        return {
            "dsn": config["dsn"],
            "rows_affected": 0,
            "error": str(e),
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
                        'EmpLastName': empleado[5], #SE INGRESALA COLUMNA DEL APELLIDO
                        'ADRESS2': empleado[2],
                        'POSNAME': empleado[3],
                        'newPass': newPass
                    }
                    empleados_info.append(empleado_info)
                    print(f"Administrador: {empleado[1], empleado[5]}, Nueva Clave: {newPass}") #SE INGRESA NOMBRE Y APELLIDO
                    change_password_cashier(newPass,empleado[0],config)

                
                # Agregar información al cuerpo del correo
                cuerpo_total += f"\nReporte para {config['dsn']}:\n"
                cuerpo_total += "=" * 50 + "\n"
                
                if empleados_info:
                    tiene_datos = True
                    for emp in empleados_info:
                        cuerpo_total += f"Nombre completo: {emp['EMPNAME']} {emp['EmpLastName']}\n"
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
