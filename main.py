import pandas as pd
import os

carpeta_excel = "C:/Users/Consultar/Desktop/Benicio y Matias/extraccion_datos_py/archivos_excel"
archivo_salida = "clientes_unificados.txt"

def encontrar_valor(df, etiqueta):
    
    for i in range(len(df)):
        for j in range(len(df.columns) - 1):
            if str(df.iat[i, j]).strip().upper() == etiqueta.upper():
                return str(df.iat[i, j + 1]).strip()
    return ""

def extraer_bloque(df, titulo):
   
    resultado = [titulo]
    for i in range(len(df)):
        row = df.iloc[i].astype(str).tolist()
        if titulo in row:
            for j in range(1, 6):  
                if i + j < len(df):
                    siguiente = df.iloc[i + j].dropna().astype(str).str.strip().tolist()
                    if any(siguiente):
                        resultado.append("\t" + " / ".join(siguiente))
            break
    return "\n".join(resultado)

def extraer_info_archivo(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo, header=None).fillna("")

        salida = []
        salida.append("DATOS DEL CLIENTE")
        salida.append(f"Nº\t{encontrar_valor(df, 'Nº')}\tEMPRESA\t{encontrar_valor(df, 'EMPRESA')}")
        salida.append(f"CÓDIGO\t{encontrar_valor(df, 'CÓDIGO')}\tCUIT\t{encontrar_valor(df, 'CUIT')}")

        salida.append(extraer_bloque(df, "INFORMACIÓN COMERCIAL"))
        salida.append(extraer_bloque(df, "INFORMACIÓN FISCAL"))
        salida.append(extraer_bloque(df, "INFORMACIÓN PLANTA"))

        return "\n".join(salida)
    except Exception as e:
        return f"[ERROR] No se pudo procesar {os.path.basename(ruta_archivo)}: {e}"


with open(archivo_salida, "w", encoding="utf-8") as f_out:
    for archivo in os.listdir(carpeta_excel):
        if archivo.endswith(".xlsx") and not archivo.startswith("~$"):
            ruta_completa = os.path.join(carpeta_excel, archivo)
            resultado = extraer_info_archivo(ruta_completa)
            f_out.write(resultado + "\n\n")

print("Extracción completada. Revisa el archivo:", archivo_salida)
