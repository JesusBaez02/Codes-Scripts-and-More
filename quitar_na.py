import pandas as pd

# Cargar el archivo Excel
df = pd.read_excel('conceptos_completos.xlsx')

# Eliminar las filas donde la columna esté vacía
df = df.dropna()

# Guardar el archivo con las filas vacías eliminadas
df.to_excel('archivo_sin_filas3.xlsx', index=False)
