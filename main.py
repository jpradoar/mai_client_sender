#!/usr/bin/python3

import pandas as pd
from email.message import EmailMessage
import smtplib
from datetime import datetime
import pytz
import openpyxl
from time import sleep

# Configuracion del documento
libro_excel             = 'clientes.xlsx'
hoja_de_trabajo         = 'main_page'
tiempo_entre_cada_mail  = 5
pais_zona_horaria       = 'America/Argentina/Buenos_Aires'

servidor_smtp   = "smtp.server.com"
puerto_smtp     = 123
sender          = "user@domain.com"
password        = "superduperpassword"
Bcc_mail        = "user@domain.com"

# Configuracion de fecha y zona horaria
fecha_hora_actual_utc   = datetime.now(pytz.utc)
zona_horaria            = pytz.timezone(pais_zona_horaria)
fecha_hora_argentina    = fecha_hora_actual_utc.astimezone(zona_horaria)
formato_fecha           = "%Y-%m-%d %H:%M:%S"
fecha_hora_formateada   = fecha_hora_argentina.strftime(formato_fecha)

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

def enviar_mail(correo, MSG_SUBJECT, MAIL_FULL_MSG):
    email = EmailMessage()
    email["From"] = sender
    email["To"] = correo
    email["Subject"] = MSG_SUBJECT
    email["Bcc"] = Bcc_mail
    email.set_content(MAIL_FULL_MSG)

    try:
        smtp = smtplib.SMTP_SSL(servidor_smtp)
        smtp.login(sender, password)
        smtp.send_message(email)  # Aquí enviamos el mensaje completo
        smtp.quit()
        print("Correo enviado a: " + correo + "\n")
        return True  # Indica que el correo se envió correctamente

    except Exception as e:
        print("Error al enviar correo a " + correo + ":", e)
        return False  # Indica que hubo un error al enviar el correo

def leer_primera_columna_condicional(libro_excel):
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(libro_excel)
        for indice, fila in df.iterrows():
            EMPRESA_CORREO = str(fila['EMPRESA_CORREO'])
            correos = EMPRESA_CORREO.split(';')
            EMPRESA_NOMBRE = str(fila['EMPRESA_NOMBRE'])
            MAIL_SUBJECT = str(fila['MAIL_SUBJECT'])
            ENVIAR_MAIL = str(fila['ENVIAR_MAIL'])
            MAIL_TEXTO_A = str(fila['MAIL_TEXTO_A'])
            MAIL_ULTIMO_AJUSTE = str(fila['MAIL_ULTIMO_AJUSTE'])
            MAIL_TEXTO_B = str(fila['MAIL_TEXTO_B'])
            MAIL_PROX_AJUSTE = str(fila['MAIL_PROX_AJUSTE'])
            MAIL_TEXTO_C = str(fila['MAIL_TEXTO_C'])
            MAIL_NUEVA_TARIFA = str(fila['MAIL_NUEVA_TARIFA'])
            MAIL_IVA = str(fila['MAIL_IVA'])
            MAIL_TEXTO_D = str(fila['MAIL_TEXTO_D'])
            MAIL_FIRMA = str(fila['MAIL_FIRMA'])
            columna_destino = 14  # Asumiendo que la columna ULTIMO_ENVIO es la 14
            fila_destino = indice + 1
            MSG_SUBJECT = MAIL_SUBJECT + " " + EMPRESA_NOMBRE
            MAIL_FULL_MSG = ("Estimado/a cliente,\n\n" + MAIL_TEXTO_A + " " + MAIL_ULTIMO_AJUSTE + " " +
                             MAIL_TEXTO_B + " " + MAIL_PROX_AJUSTE + " " + MAIL_TEXTO_C + " " + MAIL_NUEVA_TARIFA +
                             " " + MAIL_IVA + "\n\n" + MAIL_TEXTO_D + "\n\nSin más, reciba un cordial saludo.\n\n" +
                             MAIL_FIRMA + " ")

            if ENVIAR_MAIL.lower() == "si":
                print(correos, MSG_SUBJECT, MAIL_FULL_MSG)
                print("\n")
                all_sent = True  # Variable para controlar si todos los correos se enviaron correctamente
                for correo in correos:
                    if enviar_mail(correo.strip(), MSG_SUBJECT, MAIL_FULL_MSG):
                        sleep(tiempo_entre_cada_mail)
                    else:
                        all_sent = False
                if all_sent:
                    update_sent_data(fila_destino, columna_destino)
                print("#----------------------------\n")

            if ENVIAR_MAIL.lower() == "test":
                print("\n#----------------------------\nTEST: No se envia correo.")
                print(MAIL_FULL_MSG)

    except Exception as e:
        print(f"Error: {e}")

leer_primera_columna_condicional(libro_excel)
