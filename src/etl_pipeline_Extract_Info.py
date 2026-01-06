import pandas as pd
import os
import sys
from datetime import datetime
import time
import pythoncom
import win32com.client as win32

#-------------------------- CONFIGURACION ---------------------------------------

ruta_directorio = r"C:\Users\sg20690\OneDrive - Stellantis" # ruta general
ruta_copia_local = r"C:\Users\sg20690\Desktop\Stellantis" # ruta para copia local
ruta_bases = os.path.join(ruta_directorio, "Bases_Datos") # ruta de las bases de datos
ruta_maestro = os.path.join(ruta_directorio, "Bases_Datos") # ruta de las bases de datos maestras
ruta_proyectos = os.path.join(ruta_directorio, "Proyectos", "Status_Comercial")
ruta_avances_pedidos = os.path.join(ruta_proyectos, "Reportes_Avance_Pedidos") # para exportar archivo de salida (mails)
ruta_contactos = r"C:\Users\sg20690\Stellantis\Importadores PV (Andina y Centroamerica) - General\07. MAESTRO IMPORTADORES"
correo_de_envios = "aftersalesimporters@stellantis.com" # mail del cual se enviara la info

# Fechas de referencia para filtrar los pedidos, objetivos y facturacion por anio y mes
fecha_hoy = datetime.now()
fecha_hoy_str = fecha_hoy.strftime('%d-%m-%Y')
anio_actual = fecha_hoy.year
mes_actual = fecha_hoy.month

# Nombre de archivos parametrizados
archivo_pedidos = "historico_pedidos.xlsx"
archivo_facturacion = "historico_facturacion.xlsx"
archivo_objetivos = "historico_objetivos_comerciales.xlsx"
archivo_contactos = "Maestro de contactos - Importadores.xlsx"
archivo_zonales = "1_maestro_jefes_zonales.xlsx"
archivo_importadores = "1_maestro_importadores.xlsx"
nombre_archivo_salida = f"Reporte_Pedidos_{datetime.now().strftime('%Y%m%d')}.xlsx"

#----------------------------------------------------------------------------------------

#------------------ 📌📌 SOLO PARA REALIZAR PRUEBAS DE ENVIOS-----------------------

question = input("¿Desea activar el modo de prueba? (si/no/cancel): ")
if question.lower() in ["si", "s"]:
    ruta_contactos = os.path.join(ruta_copia_local, "Bases_Datos") # SOLO PARA REALIZAR PRUEBA DE ENVIOS
    ruta_maestro = os.path.join(ruta_copia_local, "Bases_Datos") # SOLO PARA REALIZAR PRUEBA DE ENVIOS
    archivo_contactos = "1_maestro_contactos_test.xlsx" # SOLO PARA REALIZAR PRUEBAS DE ENVIOS
    archivo_zonales = "1_maestro_jefes_zonales_test.xlsx" # SOLO PARA REALIZAR PRUEBAS DE ENVIOS
elif question.lower() in ["cancel","cancelar","c"]:   
    print("Envio cancelado!")
    sys.exit()

#--------------------------------------------------------------------------------------

# 📌 Definir trimestre actual
trimestre = {
    1: [1, 2, 3],
    2: [4, 5, 6],
    3: [7, 8, 9],
    4: [10, 11, 12]
}
trimestre_actual = [v for k, v in trimestre.items() if mes_actual in v][0]

# 📌 Leer bases de datos
df_pedidos = pd.read_excel(os.path.join(ruta_bases, archivo_pedidos))
df_facturacion = pd.read_excel(os.path.join(ruta_bases, archivo_facturacion))
df_objetivos = pd.read_excel(os.path.join(ruta_bases, archivo_objetivos))
df_contactos = pd.read_excel(os.path.join(ruta_contactos, archivo_contactos)).dropna(subset=["STATUS AVANCE","EMAIL"])
df_zonales = pd.read_excel(os.path.join(ruta_maestro, archivo_zonales))
df_importadores = pd.read_excel(os.path.join(ruta_maestro, archivo_importadores))

# 📌 Formatear df_importador para cruzar con facturacion, obteniendo variable GRUPO_IMPORT
df_importadores = df_importadores[["ID_IMPORT","GRUPO_IMPORT"]]
df_facturacion = df_facturacion.merge(df_importadores, on=["ID_IMPORT"], how="left")

# 📌 Formatear df_contactos para luego usar en el envio de mail
df_contactos = df_contactos[df_contactos["CARGO"] != "ZONE MANAGER"] # sacar zone manager de la base
df_contactos = df_contactos[["ID_PAIS","GRUPO_IMPORT","PAIS","IMPORTADOR","EMAIL"]] # seleccionar variables
df_contactos = df_contactos[(df_contactos["EMAIL"].str.contains("@"))] # Filtrar solo los que son correos validos
# Agrupar y concatenar mails válidos con punto y coma
df_contactos_group = df_contactos.groupby(
    ["ID_PAIS", "GRUPO_IMPORT", "PAIS", "IMPORTADOR"], as_index=False).agg({"EMAIL": lambda x: "; ".join(x.dropna())})

df_zonales = df_zonales.rename(columns={"EMAIL":"EMAIL_ZONAL"}) # Renombrar variable Email    
df_zonales = df_zonales[["ID_PAIS","GRUPO_IMPORT","EMAIL_ZONAL"]] # seleccionar variables a utilizar

df_envios = pd.merge(df_contactos_group,df_zonales, how="left",on=["ID_PAIS","GRUPO_IMPORT"]) # unir bases
df_envios["EMAIL_ZONAL"] = df_envios["EMAIL_ZONAL"].fillna(correo_de_envios) # rellenar na con correo aftersales

# 📌 Formatear fechas para filtrar por mes/anio
df_pedidos["FECHA"] = pd.to_datetime(df_pedidos["FECHA"], format="%d-%m-%Y")
df_facturacion["FECHA"] = pd.to_datetime(df_facturacion["FECHA"], format="%d-%m-%Y")
df_objetivos["FECHA"] = pd.to_datetime(df_objetivos["FECHA"], format="%d-%m-%Y")

# 📌 Filtrar solo trimestre actual y anio actual
df_pedidos = df_pedidos[(df_pedidos["FECHA"].dt.month.isin(trimestre_actual)) & (df_pedidos["FECHA"].dt.year == anio_actual)]
df_facturacion = df_facturacion[(df_facturacion["FECHA"].dt.month.isin(trimestre_actual)) & (df_facturacion["FECHA"].dt.year == anio_actual)]
df_objetivos = df_objetivos[(df_objetivos["FECHA"].dt.month.isin(trimestre_actual)) & (df_objetivos["FECHA"].dt.year == anio_actual)]

# 📌 Consolidar base trimestral
def consolidar_base(base, campo_valor, nombre_columna):
    return base.groupby(["ID_PAIS", "GRUPO_IMPORT", "FECHA"]).agg({campo_valor: "sum"}).rename(columns={campo_valor: nombre_columna}).reset_index()

pedidos = consolidar_base(df_pedidos, "PEDIDOS", "PEDIDOS")
facturacion = consolidar_base(df_facturacion, "FACTURADO", "FACTURADO")
objetivos = consolidar_base(df_objetivos, "OBJETIVO", "OBJETIVO")

# 📌 Formatear valores float a int en campo de montos
pedidos["PEDIDOS"] = pedidos["PEDIDOS"].astype(int)
facturacion["FACTURADO"] = facturacion["FACTURADO"].astype(int)
objetivos["OBJETIVO"] = objetivos["OBJETIVO"].astype(int)

# 📌 Unir bases
base_merge = pedidos.merge(facturacion, on=["ID_PAIS", "GRUPO_IMPORT", "FECHA"], how="outer").merge(objetivos, on=["ID_PAIS", "GRUPO_IMPORT", "FECHA"], how="outer").fillna(0)

# 📌 Calcular acumulados trimestrales
base_merge["FECHA"] = base_merge["FECHA"].dt.strftime("%b")  # Nombre mes
acumulados = base_merge.groupby(["ID_PAIS", "GRUPO_IMPORT"]).agg({
    "PEDIDOS": "sum",
    "FACTURADO": "sum",
    "OBJETIVO": "sum"
}).rename(columns=lambda c: f"{c}_ACUM").reset_index()

# 📌 Consolidar tabla final
tabla_final = base_merge.pivot_table(index=["ID_PAIS", "GRUPO_IMPORT"],
                                     columns="FECHA",
                                     values=["PEDIDOS", "FACTURADO", "OBJETIVO"],
                                     fill_value=0).reset_index()

# 📌 Aplanar columnas
tabla_final.columns = [' '.join(col).strip() for col in tabla_final.columns.values]

# 📌 Unir acumulados
tabla_final = tabla_final.merge(acumulados, on=["ID_PAIS", "GRUPO_IMPORT"], how="left")

# 📌 Calcular % cumplimiento acumulado y bonificación
tabla_final["% CUMPLIMIENTO"] = (tabla_final["PEDIDOS_ACUM"] / tabla_final["OBJETIVO_ACUM"] * 100).round(2)
tabla_final["BONIFICACION %"] = tabla_final["% CUMPLIMIENTO"].apply(
    lambda x: 14 if x >= 110 else (11 if x >= 100 else (6 if x >= 90 else 0)))
tabla_final["BONIFICACION $"] = (tabla_final["FACTURADO_ACUM"] * (tabla_final["BONIFICACION %"]/100)).round(0)

# 📌 Exportar a Excel
ruta_salida = os.path.join(ruta_avances_pedidos, nombre_archivo_salida)
ruta_salida_local = os.path.join(ruta_copia_local,"Proyectos", "Status_Comercial","Reportes_Avance_Pedidos",nombre_archivo_salida)
tabla_final.to_excel(ruta_salida, index=False)
tabla_final.to_excel(ruta_salida_local,index=False)
print(f"Archivo exportado correctamente a: {ruta_salida}")
print(f"Archivo exportado correctamente a carpeta local: {ruta_salida_local}")

############################ FILTRO en caso de querer enviar solo a ciertos importadores
# resp_f = input("Existe un filtro activo. ¿Desea continuar? (si/no): ")
# if resp_f.lower() == "si":
#     tabla_final = tabla_final[((tabla_final["ID_PAIS"].isin(["BOL"])) & (tabla_final["GRUPO_IMPORT"] == "FCA"))]
# else:
#     sys.exit()
    
# Para revisar datos antes de enviar
pedidos_cols = [col for col in tabla_final.columns if col.startswith("PEDIDOS")]
columnas_a_mostrar = ["ID_PAIS", "GRUPO_IMPORT"] + pedidos_cols
print("Contactos a enviar\n", df_envios["EMAIL_ZONAL"].unique(),"\n" ,df_contactos["EMAIL"])
print("Datos a enviar:\n", tabla_final[columnas_a_mostrar])
respuesta = input("Revisar tabla antes de enviar! ¿Quiere continuar? (si/no): ")
if respuesta.lower() != "si":
    print("Envio cancelado!")
    sys.exit()

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
    
    ########### 📌 1 Asunto 📌###
    
    # nombre del importador que ira en el asunto y en el saludo
    nomb_destinatario =  contacto_info["IMPORTADOR"].values[0] + " - " + contacto_info["PAIS"].values[0]
    if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        mail.Subject = f"Weekly Orders Report - {nomb_destinatario} - {contacto_info["GRUPO_IMPORT"].values[0]}"
    else:    
        mail.Subject = f"Reporte Semanal de Pedidos - {nomb_destinatario} - {contacto_info["GRUPO_IMPORT"].values[0]}"

    ########### 📌 2 Destinatario 📌###
    
    destinatario = contacto_info["EMAIL"].values[0]
    mail.To = destinatario

    ########### 📌 3 Destinatario CC 📌###
    
    destinatario_en_copia = contacto_info["EMAIL_ZONAL"].values[0]
    mail.CC = destinatario_en_copia
    
    ########### 📌 4 Mensaje 📌###
    
    # 📌 Generar tabla HTML
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

    # # 📌 Fila FACTURACIÓN
    # if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        # tabla_html += "<tr><td style='text-align:left;'>(*) INVOICE</td>"
    # else:
        # tabla_html += "<tr><td style='text-align:left;'>(*) FACTURACIÓN</td>"
    # for mes in meses_trimestre:
    #     tabla_html += f"<td>{row.get(f'FACTURADO {mes}', 0):,.0f}</td>"
    # tabla_html += f"<td>{int(row['FACTURADO_ACUM']):,}</td></tr>"

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
    #tabla_html += f"<td>{row['BONIFICACION %']}%</td></tr>"

    # # 📌 Fila BONIFICACIÓN $
    # if not contacto_info.empty and contacto_info["ID_PAIS"].values[0] == "BAR":
        # tabla_html += "<tr><td style='text-align:left;'>(**) BONIFICATION $</td>"
    # else:
        # tabla_html += "<tr><td style='text-align:left;'>(**) BONIFICACIÓN $</td>"
    # bonificacion_total = 0

    # for mes in meses_trimestre:
    #     col_facturacion = f'FACTURADO {mes}'
    #     col_objetivo = f'OBJETIVO {mes}'
    #     col_pedidos = f'PEDIDOS {mes}'

    #     # Revisa si las columnas existen
    #     if col_facturacion in row and col_objetivo in row and col_pedidos in row:
    #         facturacion = row[col_facturacion]
    #         objetivo = row[col_objetivo]
    #         pedidos = row[col_pedidos]

    #         cumplimiento = (pedidos / objetivo) * 100 if objetivo > 0 else 0
    #         bonificacion_pct = 14 if cumplimiento >= 110 else 11 if cumplimiento >= 100 else 6 if cumplimiento >= 90 else 0
    #         bonificacion_monto = facturacion * (bonificacion_pct / 100) if facturacion > 0 else 0

    #         bonificacion_total += bonificacion_monto
    #         tabla_html += f"<td>${bonificacion_monto:,.0f}</td>"
    #     else:
    #         # Si no existe, muestra cero
    #         tabla_html += "<td>$0</td>"

    # tabla_html += f"<td>${bonificacion_total:,.0f}</td></tr>"

    tabla_html += "</tbody></table>"

    ###########################################################################################
    # <p> (*) Facturación del mes actual es estimada.</p>
    
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
    
    #### 📌 5 Envio del mail 📌 ###   
    mail.Send()
    print(f"Mail enviado a: {nomb_destinatario} ({destinatario})(cc: {destinatario_en_copia})")

    correos_enviados += 1
    time.sleep(pausa_normal)
    if correos_enviados % max_correos_antes_pausa == 0:
        print(f"Pausa larga. Se han enviado {correos_enviados} hasta el momento.")
        time.sleep(8)

cerrar_outlook()
print(f"Se enviaron {correos_enviados} correos exitosamente.")