import pandas as pd

# Cargar el archivo Excel
file_path = 'ejemplo.xlsx'  # Reemplaza con la ruta de tu archivo
df = pd.read_excel(file_path)

# Crear un contador para identificar diferentes archivos
contador = 1
current_df = []

# Iterar sobre cada fila en el archivo
for index, row in df.iterrows():
    # Si encontramos 'ACCESO AUTOMOTRIZ', guardamos el dataframe anterior y comenzamos uno nuevo
    if 'ACCESO AUTOMOTRIZ, S.A. DE C.V.' in row.values:
        if current_df:  # Si ya tenemos datos acumulados
            # Guardar el dataframe actual en un archivo Excel separado
            output_path = f'subarchivo_{contador}.xlsx'
            pd.DataFrame(current_df).to_excel(output_path, index=False)
            print(f'Archivo guardado: {output_path}')
            contador += 1  # Incrementar el contador para el siguiente archivo
            current_df = []  # Reiniciar la lista para el siguiente grupo
    current_df.append(row)  # Añadir la fila actual al dataframe en construcción

# Asegurarse de guardar el último conjunto de datos si aún no se ha guardado
if current_df:
    output_path = f'subarchivo_{contador}.xlsx'
    pd.DataFrame(current_df).to_excel(output_path, index=False)
    print(f'Archivo guardado: {output_path}')
