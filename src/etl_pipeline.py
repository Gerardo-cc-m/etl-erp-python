import pandas as pd
import os
import pickle

# Abrir archivo con bases temporales (creadas en _1000_Generar_StatusComercial_v2)
with open("temp.pkl", "rb") as f:
    datatemp = pickle.load(f)

# abrir variables de archivo temporal
ruta_bases = datatemp["ruta_bases"] # ruta principal ubicacion bases
ruta_archivo = datatemp["ruta_archivo"] # ruta archivo de pedidos descargada de sap
nombre_archivo = datatemp["nombre_archivo"] # nombre de archivo descargado de sap
maestro_importador = datatemp["maestro_importador"]
maestro_piezas = datatemp["maestro_piezas"]
maestro_codigos = datatemp["maestro_codigos"]
pedidos_pn = datatemp["pedidos_pn"] # maestro de pedidos historicos
mes_extraccion = datatemp["mes_extraccion"]
anio_extraccion = datatemp["anio_extraccion"]
rates = datatemp["rates"] # rates con cambio de usd a eur

# Abrir base de pedidos
data = pd.read_csv(os.path.join(ruta_archivo,nombre_archivo), sep="|",skiprows=3 ,header=0, encoding="latin1")

# Renombrar columna en donde aparece si corresponde a BRP0 (peugeot) o BRC0 (citroen)
data = data.rename(columns={"    ":"br"})
# poner en minusculas, y quitar espacios y tabulaciones en nombre de columnas
data.columns = data.columns.str.strip().str.lower()

# Diccionario de palabras clave y nombres deseados para renombrar y filtrar
rename_map = {"reference": "PN",
                "qtd.ord.": "QTDE_PEDIDO",
                "dt.order": "DATA_PEDIDO",
                "vlr.pht": "VALOR_PEDIDO",
                "client": "CODIGO",
                "br":"BR"} # Este ultimo es para determinar si es citroen o peugoet en costarica y paraguay

# Generar diccionario de renombrado
col_renombradas = {col: nuevo_nombre for col in data.columns for clave, nuevo_nombre in rename_map.items() if clave in col}
# Generar variable de filtrado
key_words = list(rename_map.keys())
cols_filtradas = [col for col in data.columns if any(p in col for p in key_words)]

# Filtrar y renombrar para homologar con bases de datos
df_pedidos = data[cols_filtradas].rename(columns=col_renombradas)

# Quita los espacios y tabulaciones de todo los registros
for col in df_pedidos.select_dtypes(include="object").columns:
    df_pedidos[col] = df_pedidos[col].str.strip()

# Eliminar nulos
df_pedidos = df_pedidos.dropna()
# Transformar data pedido a variable fecha
df_pedidos["DATA_PEDIDO"] = pd.to_datetime(df_pedidos["DATA_PEDIDO"],format="%d.%m.%Y")
# Cambiar todos los dias del mes en 01
df_pedidos["DATA_PEDIDO"] = df_pedidos["DATA_PEDIDO"].values.astype("datetime64[M]")
# Transformar cantidad pedidos a int
df_pedidos["QTDE_PEDIDO"] = df_pedidos["QTDE_PEDIDO"].astype(int)
# Rellenar con ceros hasta formar cadena de 13 caracteres en variable PN
df_pedidos["PN"] = df_pedidos["PN"].astype(str).str.zfill(13)
# Valor pedido a numerico (reemplazando caracteres de puntos y comas)
df_pedidos['VALOR_PEDIDO'] = pd.to_numeric(df_pedidos['VALOR_PEDIDO'].str.replace(".","",regex=False).str.replace(",",".",regex=False).astype(float))
# Filtrar pedidos con valor mayor a cero
df_pedidos = df_pedidos[df_pedidos["VALOR_PEDIDO"] > 0]

#----- Abrir archivos maestros -------
# maestro codigo
df_maestro_codigos = pd.read_excel(os.path.join(ruta_bases,maestro_codigos),sheet_name="codigos_polo")
df_maestro_codigos= df_maestro_codigos[df_maestro_codigos["POLO"] == "PORTO REAL"] # solo porto real
# maestro de piezas pn
df_maestro_piezas = pd.read_excel(os.path.join(ruta_bases,maestro_piezas), sheet_name="porto_real")
df_maestro_piezas = df_maestro_piezas[["PN","DESC_MATERIAL","INDEX","CATEGORIA"]] # variables a utilizar
# maestro importador
df_maestro_importador = pd.read_excel(os.path.join(ruta_bases,maestro_importador))
df_maestro_importador = df_maestro_importador[["ID_IMPORT","NOM_PAIS","MARCA"]] # variables a utilizar
#------

# unir pedidos con bases maestros
df_merge = pd.merge(df_maestro_codigos,df_pedidos,how="inner",on="CODIGO")
# Determinar MARCA en Costa Rica y Paraguay (por defecto la marca viene como Peugeot en base de datos)
df_merge.loc[(df_merge["BR"]=="BRC0") & (df_merge["NOM_PAIS"].isin(["PARAGUAY","COSTA RICA"])),"MARCA"] = "CITROEN"

df_merge = pd.merge(df_merge,df_maestro_piezas,how="left",on="PN")
df_merge =pd.merge(df_merge,df_maestro_importador,how="left",on=["NOM_PAIS","MARCA"])

# Eliminar columnas que no se utilizaran
df_merge = df_merge.drop(columns=["CODIGO","BR"],errors="ignore")

# Filtrar por mes y anio de extraccion
df_nuevos_datos = df_merge[(df_merge["DATA_PEDIDO"].dt.month == mes_extraccion) & 
                        (df_merge["DATA_PEDIDO"].dt.year == anio_extraccion)]

# Agrupar todos los PN, fecha e importador
df_nuevos_datos = df_nuevos_datos.groupby(["PN","INDEX","CATEGORIA","DESC_MATERIAL","ID_IMPORT",
                                        "NOM_PAIS","MARCA","POLO","DATA_PEDIDO"]).agg({"QTDE_PEDIDO":"sum","VALOR_PEDIDO":"sum"}).reset_index()

#--------- CONVERTIR DOLAR A EURO ---------
fch = f"01-{mes_extraccion:02d}-{anio_extraccion}" # definir segun la fecha de extraccion
fecha = pd.to_datetime(fch,format="%d-%m-%Y") # crear variable fecha
df_rates = pd.read_excel(os.path.join(ruta_bases,rates)) # leer base de datos
# filtrar por fecha y tipo = Act
df_rates_filtro = df_rates[(df_rates["FECHA"] == fecha) & (df_rates["TIPO"] == "Act")]
# condicion: si base queda vacia, tomar el rates del mes anterior
if df_rates_filtro.empty:
    fecha_anterior = fecha - pd.DateOffset(months=1)
    df_rates_filtro = df_rates[(df_rates["FECHA"] == fecha_anterior) & (df_rates["TIPO"] == "Act")]

# Transformar Valor Pedido de dolares a Euros
df_nuevos_datos["VALOR_PEDIDO"] = df_nuevos_datos["VALOR_PEDIDO"] / df_rates_filtro["CONVERSION"].iloc[0]
#------------------------------------------

# Abrir archivo consolidado con historicos
df_pedidos_pn = pd.read_excel(os.path.join(ruta_bases,pedidos_pn))
# Eliminar todos los registros del mes de extraccion y polo, para reemplazar luego por los nuevos valores
df_pedidos_pn = df_pedidos_pn[~(((df_pedidos_pn["DATA_PEDIDO"].dt.month == mes_extraccion) &
                                (df_pedidos_pn["DATA_PEDIDO"].dt.year == anio_extraccion)) & 
                                (df_pedidos_pn["POLO"] == "PORTO REAL"))]

# Incorporar nuevos registros del mes extraido
df_actualizado = pd.concat([df_pedidos_pn,df_nuevos_datos])

# Guardar
df_actualizado.to_excel(os.path.join(ruta_bases,pedidos_pn),index=False)

print("Archivo con nuevos registros guardado con exito!")

# Poner fecha de actualizacion a archivo Status_Comercial
from _001_Actualizar_fecha_status import actualizar_fecha_estatus
actualizar_fecha_estatus()
