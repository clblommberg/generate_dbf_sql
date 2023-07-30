import os
import csv

RUTA_ORIGEN = 'C:\\Users\\claudio\\Projects\\generate_dbf_sql\\data\\exportdesing'
RUTA_DESTINO = 'C:\\Users\\claudio\\Projects\\generate_dbf_sql\\data\\tablas_sql'

def convertir_csv_a_sql(ruta_archivo):

    nombre_tabla = ""
    columnas = []

    with open(ruta_archivo) as f:
        lector = csv.reader(f)
        primera_fila = next(lector)
        nombre_tabla = primera_fila[0].split(':')[1].strip()

        for row in lector:
            if row:
                nombre = row[0].split(' ')[0]
                tipo = row[0].split(' ')[1]
                columnas.append(f"{nombre} {tipo}")

    return {nombre_tabla: columnas}

def funcion_principal():
    archivos_csv = os.listdir(RUTA_ORIGEN)
    consultas = []

    for archivo in archivos_csv:
        if archivo.endswith('.csv'):
            ruta_archivo = os.path.join(RUTA_ORIGEN, archivo)
            columnas_por_tabla = convertir_csv_a_sql(ruta_archivo)
            for nombre_tabla, columnas in columnas_por_tabla.items():
                consulta = f"CREATE TABLE {nombre_tabla} ("
                for col in columnas:
                    consulta += f"\n {col},"
                consulta += "\n);"
                consultas.append(consulta)

    script_sql = "\n".join(consultas)

    ruta_sql_final = os.path.join(RUTA_DESTINO, "todas_las_tablas.sql")

    with open(ruta_sql_final, 'w') as f:
        f.write(script_sql)

    print("Script SQL generado con todas las tablas.")

