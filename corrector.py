import pandas as pd
import re
import unicodedata
from unidecode import unidecode

REEMPLAZOS = {
    "ӎ": "ÓN",
    "Ѵ": "V",
    "Ι": "I",
    "Ӑ": "Á",
    "Ӓ": "Ä",
    "ɢ": "É",
    "ӂ": "OB",
    "Ƀ": "ÉC",
    "Ӈ": "ÓG",
    "ڂ": "ÚB",
    "я": "ÑO",

}

# 🔹 Función para limpiar caracteres raros detectados
def limpiar_caracteres(valor):
    valor = unicodedata.normalize("NFC", valor)
    # Reemplazar manualmente según el diccionario
    for raro, correcto in REEMPLAZOS.items():
        valor = valor.replace(raro, correcto)
    return valor

#  1. Leer Excel
df = pd.read_excel("Nombres.xlsx", header=None)
df = df.iloc[:1000, :12]
df = df.astype(str)

patron_normal = re.compile(r"^[A-Za-zÁÉÍÓÚÜáéíóúüÑñ0-9 _'’\"\-.,]+$")

nombres_raros = set()
cambios = set()
nombres_normalizados = set()

#  2. Procesar cada celda
for valor in df.values.flatten():
    valor = valor.strip()
    if valor == "" or valor.lower() == "nan":
        continue

    # 🔹 Limpiar caracteres extraños antes de normalizar
    valor_limpio = limpiar_caracteres(valor)

    # 🔹 Normalizar
    nombre_norm = unidecode(valor_limpio).upper().strip()

    if not patron_normal.match(valor_limpio):
        nombres_raros.add(valor)

    # Guardar cambios si el valor cambió
    if valor.upper() != nombre_norm:
        cambios.add(f"{valor} → {nombre_norm}")

    nombres_normalizados.add(nombre_norm)

#  3. Exportar resultados
with pd.ExcelWriter("Nombres_Procesados.xlsx", engine="openpyxl") as writer:
    pd.DataFrame(sorted(nombres_raros), columns=["Nombre con caracteres raros"]).to_excel(
        writer, sheet_name="Con_caracteres_raros", index=False
    )
    pd.DataFrame(sorted(cambios), columns=["Cambio"]).to_excel(
        writer, sheet_name="Cambios", index=False
    )
    pd.DataFrame(sorted(nombres_normalizados), columns=["Nombre normalizado"]).to_excel(
        writer, sheet_name="Normalizados", index=False
    )

print(" Archivo 'Nombres_Procesados.xlsx' generado correctamente.")
