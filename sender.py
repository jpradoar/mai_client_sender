import pandas as pd
from email.message import EmailMessage
import smtplib
from datetime import datetime
import pytz
import openpyxl
import time
import os

# Configuracion del documento
libro_excel             = 'clients.xlsx'
hoja_de_trabajo         = 'main_page'
tiempo_entre_cada_mail  = os.environ.get('tiempo_entre_cada_mail')
pais_zona_horaria       = os.environ.get('pais_zona_horaria')

# Configurar el servidor SMTP
servidor_smtp   = os.environ.get('servidor_smtp')
puerto_smtp     = os.environ.get('puerto_smtp')
sender          = os.environ.get('sender')
password        = os.environ.get('password')

# Configuracion de fecha y zona horaria
fecha_hora_actual_utc   = datetime.now(pytz.utc)
zona_horaria            = pytz.timezone(pais_zona_horaria)
fecha_hora_argentina    = fecha_hora_actual_utc.astimezone(zona_horaria)
formato_fecha           = "%Y-%m-%d %H:%M:%S"
fecha_hora_formateada   = fecha_hora_argentina.strftime(formato_fecha)


def es_entero(valor):
    try:
        int(valor)
        return True
    except ValueError:
        print("No hubo cambios desde la ultima vez, o bien hubo un error al obtener el redondeo. Vuelva a intentar\n")
        exit()  # Cierra el programa


def enviar_mail(nombre, correo, subject, body_msg, fecha02, redondeo, fila_destino, columna_destino):
    es_entero(redondeo)
    email               = EmailMessage()
    email["From"]       = sender
    email["To"]         = correo
    email["Subject"]    = subject
    email.set_content(body_msg)
    #
    smtp                = smtplib.SMTP_SSL(servidor_smtp)
    smtp.login(sender, password)
    smtp.sendmail(sender, correo, email.as_string())
    smtp.quit()
    #
    print("Correo enviado a: "+ nombre + " (" + correo + ")\n")
    update_sent_data(fila_destino, columna_destino)
    # espero X tiempo entre cada envío de mail
    time.sleep(tiempo_entre_cada_mail)





def update_sent_data(fila_destino, columna_destino):
    # Abrir el archivo Excel
    open_excel = openpyxl.load_workbook(libro_excel)
    # Seleccionar la hoja de trabajo (puedes cambiar el nombre de la hoja según tu caso)
    hoja_trabajo = open_excel[hoja_de_trabajo]
    fila_destino = fila_destino + 1
    # Escribir la fecha y hora en la celda correspondiente
    hoja_trabajo.cell(row=fila_destino, column=columna_destino, value=fecha_hora_formateada)
    # Guardar el archivo Excel
    open_excel.save(libro_excel)
    print("Excel actualizado\n")





def leer_primera_columna_condicional(libro_excel):
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(libro_excel)
        # Verificar la condición en la segunda columna y luego imprimir la primera columna
        for indice, fila in df.iterrows():
                correo          = fila[0]
                nombre          = fila[1]
                body_msg        = fila[2]
                subject         = fila[3]
                fecha01         = str(fila[4])
                fecha02         = str(fila[5])
                costo           = str(fila[6])
                redondeo        = str(fila[7])
                fila_destino    = indice +1
                columna_destino = 10
                full_msg        = "Estimado/a cliente: " + nombre+ ". \n" + body_msg + ": " +fecha02 + " \nUn total adeudado de: $" + redondeo + "\n"
                if fila[8] == "si":
                    enviar_mail(nombre, correo, subject, full_msg, fecha02, redondeo, fila_destino, columna_destino)
                else:
                    print("\nTEST: No se envia correo.")
                    columna_destino = 11
                    print(full_msg)
                    #update_sent_data(fila_destino, columna_destino)


    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    leer_primera_columna_condicional(libro_excel)
