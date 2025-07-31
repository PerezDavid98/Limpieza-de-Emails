import pandas as pd
import re
import os

# --------- Funciones ---------
def tiene_errores(correo_original):
    """Detecta si el correo original tiene errores antes de corregirlo"""
    if pd.isnull(correo_original):
        return True
    
    correo = str(correo_original).strip().lower()
    
    # Verificar errores comunes
    errores = [
        correo.endswith('.'),  # Termina con punto
        correo.endswith('.co'),  # Falta 'm' en .com
        correo.endswith('.comm'),  # Doble 'm'
        correo.endswith('.con'),  # 'n' en lugar de 'm'
        correo.endswith('.cmo'),  # 'm' y 'o' invertidos
    ]
    
    return any(errores)

def limpiar_correo(correo):
    if pd.isnull(correo):
        return ""
    
    # Convertir a string y limpiar espacios
    correo = str(correo).strip().lower()
    
    # Remover puntos al final repetidamente hasta que no haya más
    while correo.endswith('.'):
        correo = correo[:-1]
    
    # Corregir extensiones comunes mal escritas
    if correo.endswith('.co'):
        correo = correo[:-3] + '.com'
    if correo.endswith('.comm'):
        correo = correo[:-5] + '.com'
    if correo.endswith('.con'):
        correo = correo[:-4] + '.com'
    if correo.endswith('.cmo'):
        correo = correo[:-4] + '.com'
    
    return correo.strip()

def es_valido(correo):
    patron = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(patron, correo) is not None

def determinar_estado(correo_original, correo_corregido):
    """Determina el estado basado en el correo original y el corregido"""
    if pd.isnull(correo_original) or str(correo_original).strip() == "":
        return "Vacío"
    
    # Si el correo original tenía errores, es inválido
    if tiene_errores(correo_original):
        return "Inválido (Corregido)"
    
    # Si no tenía errores, verificar si es válido
    if es_valido(correo_corregido):
        return "Válido"
    else:
        return "Inválido"

# --------- Rutas ---------
# Nombre del archivo de entrada (debe estar en la misma carpeta del script)
nombre_excel = "Compradores_Proyectos.xlsx"

# Carpeta donde está el script actual
carpeta_script = os.path.dirname(os.path.abspath(__file__))

# Rutas completas construidas dinámicamente
ruta_excel = os.path.join(carpeta_script, nombre_excel)
ruta_salida = os.path.join(carpeta_script, "correos_validados.xlsx")

# --------- Procesamiento ---------
df = pd.read_excel(ruta_excel)
columna_correos = "Email 1"

# Limpiar y validar
df["Correo Corregido"] = df[columna_correos].apply(limpiar_correo)
df["Estado"] = df.apply(lambda row: determinar_estado(row[columna_correos], row["Correo Corregido"]), axis=1)

# Guardar Excel en la misma carpeta del script
df_resultado = df[[columna_correos, "Correo Corregido", "Estado"]]
df_resultado.to_excel(ruta_salida, index=False)

print("Archivo generado en:", ruta_salida)
