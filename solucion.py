import pandas as pd

def generar_reporte(input_file, output_file):
    xls = pd.ExcelFile(input_file)
    df = pd.read_excel(xls, sheet_name="Exportar Hoja de Trabajo")

    df["Fecha Terminación Grupo"] = pd.to_datetime(df["Fecha Terminación Grupo"], dayfirst=True, errors="coerce")
    df["Fecha última competencia eval"] = pd.to_datetime(df["Fecha última competencia eval"], dayfirst=True, errors="coerce")

    df["Tiempo Transcurrido (días)"] = (df["Fecha última competencia eval"] - df["Fecha Terminación Grupo"]).dt.days

    reporte_certificacion = df.groupby("Centro").agg(
        Total_Aprendices=("NIS Aprendiz", "count"),
        A_Tiempo=("Verificacion Titulada", lambda x: (x == "A TIEMPO PARA CERTIFICAR").sum()),
        Fuera_Tiempo=("Verificacion Titulada", lambda x: (x == "FUERA DE TIEMPO PARA CERTIFICAR").sum()),
        Juicio_Emitido=("Verificacion Titulada", lambda x: (x.str.contains("JUICIO EMITIDO", na=False)).sum())
    ).reset_index()

    analisis_programa = df.groupby(["Nombre Programa", "Nivel formacion"]).agg(
        Total_Aprendices=("NIS Aprendiz", "count")
    ).reset_index()

    analisis_modalidad = df.groupby("Modalidad").agg(
        Total_Aprendices=("NIS Aprendiz", "count"),
        A_Tiempo=("Verificacion Titulada", lambda x: (x == "A TIEMPO PARA CERTIFICAR").sum()),
        Fuera_Tiempo=("Verificacion Titulada", lambda x: (x == "FUERA DE TIEMPO PARA CERTIFICAR").sum()),
        Juicio_Emitido=("Verificacion Titulada", lambda x: (x.str.contains("JUICIO EMITIDO", na=False)).sum())
    ).reset_index()

    analisis_comunidades = df.groupby("TP Apoyo Comunidades Espec").agg(
        Total_Aprendices=("NIS Aprendiz", "count"),
        A_Tiempo=("Verificacion Titulada", lambda x: (x == "A TIEMPO PARA CERTIFICAR").sum()),
        Fuera_Tiempo=("Verificacion Titulada", lambda x: (x == "FUERA DE TIEMPO PARA CERTIFICAR").sum()),
        Juicio_Emitido=("Verificacion Titulada", lambda x: (x.str.contains("JUICIO EMITIDO", na=False)).sum())
    ).reset_index()

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        reporte_certificacion.to_excel(writer, sheet_name="Estado Certificación", index=False)
        analisis_programa.to_excel(writer, sheet_name="Programa Formación", index=False)
        analisis_modalidad.to_excel(writer, sheet_name="Modalidad Formación", index=False)
        analisis_comunidades.to_excel(writer, sheet_name="Apoyo Comunidades", index=False)

generar_reporte("origen.xlsx", "reporte_certificacion.xlsx")
