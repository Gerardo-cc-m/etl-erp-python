import pandas as pd
import os
import sys
from datetime import datetime
import time
import pythoncom
import win32com.client as win32

#-------------------------------------- 📌📌 Envío de mails 📌📌 #---------------------------------------------
# 1: Asunto
# 2: Destinatario
# 3: Destinatario CC
# 4: Cuerpo mensaje
# 5: Envio del mail

# Leer el archivo con el cuerpo del mensaje en espanol
with open("_301_Cuerpo_Mensaje_Espanol.html", encoding="utf-8") as f:
    template_esp = f.read()

# Leer el archivo con el cuerpo del mensaje en ingles (para Barbados)
with open("_302_Cuerpo_Mensaje_Ingles.html", encoding="utf-8") as f:
    template_ing = f.read()

# Iniciar Outlook
def inicializar_outlook():
    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass
    return win32.Dispatch("Outlook.Application")

def cerrar_outlook():
    pythoncom.CoUninitialize()

outlook = inicializar_outlook()
pausa_normal = 2 # Auxiliar para pausar despues de cada envio
max_correos_antes_pausa = 10 # Auxliar para pausar 10 seg despues de 10 envios
correos_enviados = 0 # Contador de correos enviados

# Enviar mail por cada importador que aparezca en tabla_final
for index, row in tabla_final.iterrows():
    # Consultar si existe el importador en base de mails
    contacto_info = df_envios.query(
        f'ID_PAIS == "{row["ID_PAIS"]}" and GRUPO_IMPORT == "{row["GRUPO_IMPORT"]}"')
    if contacto_info.empty:
        print(f"No se encontró importador para {row['ID_PAIS']} y {row['GRUPO_IMPORT']}")
        continue
    
     # Detectar nombres de los meses desde las columnas
    meses_trimestre = sorted(set([col.split()[1] for col in tabla_final.columns if "PEDIDOS" in col and "ACUM" not in col]),
                            key=lambda x: datetime.strptime(x, "%b"))

    mail = outlook.CreateItem(0) # Crear ventana de mensaje
    mail.SentOnBehalfOfName = correo_de_envios # Enviar desde un correo en especificado en parametrizacion
    
    #--------------------------------------------------📌 1 Asunto 📌------------------------------------------------#
    
    # nombre del importador que ira en el asunto y en el saludo
    nomb_destinatario =  contacto_info["IMPORTADOR"].values[0] + " - " + contacto_info["PAIS"].values[0]
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        mail.Subject = f"Weekly Orders Report - {nomb_destinatario} - {contacto_info["GRUPO_IMPORT"].values[0]}"
    else:    
        mail.Subject = f"Reporte Semanal de Pedidos - {nomb_destinatario} - {contacto_info["GRUPO_IMPORT"].values[0]}"

    #---------------------------------------------- 📌 2 Destinatario 📌----------------------------------------------#
    
    destinatario = contacto_info["EMAIL"].values[0]
    mail.To = destinatario

    #---------------------------------------------- 📌 3 Destinatario CC 📌-------------------------------------------#
    
    destinatario_en_copia = contacto_info["EMAIL_ZONAL"].values[0]
    mail.CC = destinatario_en_copia
    
    #------------------------------------------------📌 4 Mensaje 📌--------------------------------------------------#
    
    #------------------ Generar tabla HTML------------
    tabla_html = f"""
    <table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;'>
        <thead>
            <tr style='background-color:#f2f2f2;'>
                <th style='text-align:left;'>ITEM</th>
    """

    for mes in meses_trimestre:
        tabla_html += f"<th>{mes}</th>"
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        tabla_html += "<th>CUMULATIVE</th></tr></thead><tbody>"
    else:
        tabla_html += "<th>ACUMULADO</th></tr></thead><tbody>"

    # 📌 Fila OBJETIVO
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        tabla_html += "<tr><td style='text-align:left;'>TARGET</td>"
    else:
        tabla_html += "<tr><td style='text-align:left;'>OBJETIVO</td>"
    for mes in meses_trimestre:
        tabla_html += f"<td>{row.get(f'OBJETIVO {mes}', 0):,.0f}</td>"
    tabla_html += f"<td>{int(row['OBJETIVO_ACUM']):,}</td></tr>"

    # 📌 Fila PEDIDOS
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        tabla_html += "<tr><td style='text-align:left;'>ORDERS</td>"
    else:
        tabla_html += "<tr><td style='text-align:left;'>PEDIDOS</td>"
    for mes in meses_trimestre:
        tabla_html += f"<td>{row.get(f'PEDIDOS {mes}', 0):,.0f}</td>"
    tabla_html += f"<td>{int(row['PEDIDOS_ACUM']):,}</td></tr>"

    # 📌 Fila CUMPLIMIENTO %
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        tabla_html += "<tr><td style='text-align:left;'>ACHIEVEMENT %</td>"
    else:
        tabla_html += "<tr><td style='text-align:left;'>CUMPLIMIENTO %</td>"
    for mes in meses_trimestre:
        objetivo = row.get(f'OBJETIVO {mes}', 0)
        pedidos = row.get(f'PEDIDOS {mes}', 0)
        cumplimiento = (pedidos / objetivo) * 100 if objetivo > 0 else 0
        tabla_html += f"<td>{cumplimiento:.0f}%</td>"
    tabla_html += f"<td>{row['% CUMPLIMIENTO']}%</td></tr>"

    # 📌 Fila BONIFICACIÓN %
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        tabla_html += "<tr><td style='text-align:left;'>(**) BONIFICATION %</td>"
    else:
        tabla_html += "<tr><td style='text-align:left;'>(**) BONIFICACIÓN %</td>"
    for mes in meses_trimestre:
        objetivo = row.get(f'OBJETIVO {mes}', 0)
        pedidos = row.get(f'PEDIDOS {mes}', 0)
        cumplimiento = (pedidos / objetivo) * 100 if objetivo > 0 else 0
        bonificacion_pct = 14 if cumplimiento >= 110 else 11 if cumplimiento >= 100 else 6 if cumplimiento >= 90 else 0
        tabla_html += f"<td>{bonificacion_pct}%</td>"

    tabla_html += "</tbody></table>"

    #-----------------------------------------------

    # Cuerpo del mensaje en Espanol
    cuerpo_esp = template_esp.format(nomb_destinatario=nomb_destinatario,
                                     fecha_hoy_str=fecha_hoy_str,
                                     tabla_html=tabla_html)
    
    # Cuerpo del mensaje en Ingles   
    cuerpo_eng = template_ing.format(nomb_destinatario=nomb_destinatario,
                                     fecha_hoy_str=fecha_hoy_str,
                                     tabla_html=tabla_html)
   
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        mail.HTMLBody = cuerpo_eng # Envia en Ingles SOLO a Barbados
    else:
        mail.HTMLBody = cuerpo_esp
    
    #----------------------------------------------- 📌 5 Envio del mail 📌 -----------------------------------------#   
    mail.Send()
    print(f"Mail enviado a: {nomb_destinatario} ({destinatario})(cc: {destinatario_en_copia})")

    correos_enviados += 1
    time.sleep(pausa_normal)
    if correos_enviados % max_correos_antes_pausa == 0:
        print(f"Pausa larga. Se han enviado {correos_enviados} hasta el momento.")
        time.sleep(8)

cerrar_outlook()
print(f"Se enviaron {correos_enviados} correos exitosamente.")
