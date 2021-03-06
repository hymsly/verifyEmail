import threading
import smtplib
import imaplib
import time
import email
import os
import pandas as pd
import numpy as np
import sys

mensaje_error = str("No se encontr")

def limpiar_mensajes(user, password, IMAP):
    mail = imaplib.IMAP4_SSL(IMAP)
    mail.login(user, password)
    mail.select("inbox")
    typ, data = mail.search(None, 'ALL')
    for num in data[0].split():
        mail.store(num, '+FLAGS', r'(\Deleted)')
    mail.expunge()
    mail.close()
    mail.logout()
    return

def envia_mensaje_hilo(smtp_host,user,password,bloque_correo,mensaje,):
    server = smtplib.SMTP(smtp_host,587)
    server.starttls()
    server.login(user,password)
    for correo in bloque_correo:
        server.sendmail(user,correo,mensaje)
        print('el correo',correo,'se envio exitosamente')

    server.quit()
    return
    

def enviar_mensaje(correos_por_validar,mensaje,smtp_host,user,password):
    cantidad_correos = len(correos_por_validar)
    bloques_hilos = 5
    ind = 0
    while(ind < cantidad_correos):
        contador = 0
        bloque_correo = []
        while(ind < cantidad_correos and contador < bloques_hilos):
            bloque_correo.append(correos_por_validar[ind])
            ind = ind + 1
            contador = contador + 1
        hilo = threading.Thread(target=envia_mensaje_hilo, args=(smtp_host,user,password,bloque_correo,mensaje,))
        hilo.start()
        time.sleep(0.00001)
    return

def procesa_identificador_correos(data):
    cadena = str(data[0])
    if(cadena == "b''"):
        lista = []
        return lista
    cadena = cadena.replace('b','')
    cadena = cadena.replace("'",'')
    lista = cadena.split(' ')
    return lista

def get_body(msg):
    if(msg.is_multipart()):
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
    return


def leer_mensajes_no_validos(imap_host,user,password):
    # connect to host using SSL
    imap = imaplib.IMAP4_SSL(imap_host)

    # login to server
    imap.login(user, password)

    imap.select('Inbox')

    tmp, ids = imap.search(None, 'ALL')
    identificador = procesa_identificador_correos(ids)

    correos_no_validos = []
    for idd in identificador:
        nuevo_id = bytes(idd, "utf-8")
        result, data = imap.fetch(nuevo_id,'(RFC822)')
        raw = email.message_from_bytes(data[0][1])
        mensaje_validacion = str(get_body(raw))
        primera_aparicion = mensaje_validacion.find(mensaje_error)
        if(primera_aparicion != -1):
            primera_aparicion += 86
            correo = ""
            while(mensaje_validacion[primera_aparicion] != ' '):
                correo = correo + mensaje_validacion[primera_aparicion]
                primera_aparicion = primera_aparicion + 1
            correos_no_validos.append(correo)
    
    imap.close()
    return correos_no_validos

if __name__ == '__main__':
    directorio = 'uploads'
    ##debe ser parametro de entrada
    libro = sys.argv[1].split('.')[0]
    extension = '.xlsx'
    filename = os.path.join(directorio,libro+extension)
    fileoutput = os.path.join(directorio,libro+'_output'+extension)
    df = pd.read_excel(filename,header=None)
    df.columns = ['email']
    print(df.head())
    emails = list(df['email'])

    mensaje = 'Hi, it is message from python'
    user = 'email.python.11@gmail.com'
    password = 'emailpython11'
    smtp_host = 'smtp.gmail.com'
    imap_host = 'imap.gmail.com'

    limpiar_mensajes(user,password,imap_host)

    enviar_mensaje(emails,mensaje,smtp_host,user,password)

    time.sleep(20)
    
    correos_no_validos = leer_mensajes_no_validos(imap_host,user,password)

    resultado = []
    for correo in emails:
        if(correo in correos_no_validos):
            resultado.append('No Valido')
        else:
            resultado.append('Valido')

    print(resultado)
    df['resultado'] = np.array(resultado)
    print(df.head(20))
    df.to_excel(fileoutput)
    sys.stdout.flush()
    
    
    
    
