import pandas as pd

file_path = './origen.xlsx'
sheet_name = 'Exportar Hoja de Trabajo' 
df = pd.read_excel(file_path, sheet_name=sheet_name)

if 'NIS Aprendiz' in df.columns:

    duplicated_nis = df[df.duplicated('NIS Aprendiz', keep=False)]['NIS Aprendiz'].unique()
    
    if len(duplicated_nis) > 0:
        print("Los siguientes NIS están repetidos:")
        for nis in duplicated_nis:
            print(f"- {nis}")
        
            print(df[df['NIS Aprendiz'] == nis][['NIS Aprendiz', 'Aprendiz']].to_string(index=False))
    else:
        print("No hay NIS repetidos en el archivo.")
else:
    print("La columna 'NIS Aprendiz' no existe en el archivo.")