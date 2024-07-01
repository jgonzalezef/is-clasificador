import os
import pandas as pd

# Directorio donde se encuentran los archivos Excel
directorio = 'calificaciones/'

# Obtener la lista de archivos .xlsx en el directorio
archivos_excel = [archivo for archivo in os.listdir(directorio) if archivo.endswith('.xlsx')]

# Lista para almacenar todos los DataFrames de cada archivo
df_list = []

# Leer cada archivo y añadir su contenido al DataFrame
for archivo in archivos_excel:
    ruta_archivo = os.path.join(directorio, archivo)
    df = pd.read_excel(ruta_archivo)
    df_list.append(df)

# Concatenar todos los DataFrames en uno solo
df_completo = pd.concat(df_list, ignore_index=True)

# Función para determinar el valor final a utilizar
def obtener_valor_final(fila):
    if pd.notna(fila['Extra']):
        return fila['Extra']
    elif pd.notna(fila['Final']) and fila['Final'] != -1:
        return fila['Final']
    else:
        return -1 if pd.isna(fila['Final']) else fila['Final']

# Aplicar la función para obtener el valor final correcto
df_completo['Calificacion'] = df_completo.apply(obtener_valor_final, axis=1)

# Filtrar las filas donde 'PlanEstudiosClave' sea igual a '004'
df_filtrado = df_completo[df_completo['PlanEstudiosClave'] == 4]

# Ordenar los datos por 'Matricula', 'Materia' y 'Indice' (si es necesario)
df_filtrado = df_filtrado.sort_values(by=['Matricula', 'Materia'])

# Eliminar filas duplicadas, manteniendo la última calificación registrada por materia y matrícula
df_filtrado = df_filtrado.drop_duplicates(subset=['Matricula', 'Materia'], keep='last')

# Agrupar y sumar los valores de 'Calificacion' por 'Matricula' y 'Materia'
df_agrupado = df_filtrado.groupby(['Matricula', 'Materia'], as_index=False)['Calificacion'].sum()

# Pivotar los datos para obtener el formato deseado
df_pivotado = df_agrupado.pivot_table(index='Matricula', columns='Materia', values='Calificacion', aggfunc='sum').reset_index()

# Definir el orden específico de las materias
orden_materias = [
    'Matricula', 'Inglés I', 'Química Básica', 'Álgebra Lineal', 'Fundamentos de Computación', 'Algoritmos',
    'Matemáticas Discretas', 'Expresión Oral y Escrita I', 'Inglés II', 'Desarrollo Humano y Valores',
    'Cálculo Diferencial', 'Programación Orientada a Objetos', 'Estructura de Datos',
    'Ingeniería de Software Asistida por Computadora', 'Procesos de Desarrollo de Software',
    'Inglés III', 'Inteligencia Emocional y Manejo de Conflictos', 'Cálculo Integral',
    'Programación Visual', 'Estructura de Datos Avanzadas', 'Fundamentos de Base de Datos',
    'Ingeniería de Requerimientos de Software', 'Inglés IV', 'Habilidades Cognitivas y Creatividad',
    'Matemáticas para Ingeniería I', 'Programación Web', 'Diseño de Interfaces', 'Base de Datos',
    'Electricidad y Magnetismo', 'Inglés V', 'Ética Profesional', 'Matemáticas para Ingeniería II',
    'Programación Cliente Servidor', 'Fundamentos de Redes', 'Arquitectura de Software',
    'Sistemas Digitales', 'Inglés VI', 'Habilidades Gerenciales', 'Probabilidad y Estadística',
    'Arquitectura de Computadoras', 'Redes', 'Calidad del Software', 'Estancia I',
    'Inglés VII', 'Liderazgo de Equipos de Alto Desempeño', 'Lenguajes y Autómatas',
    'Sistemas Operativos', 'Programación Concurrente', 'Pruebas del Software', 'Estancia II',
    'Inglés VIII', 'Programación para Móviles I', 'Compiladores e Intérpretes', 'Inteligencia Artificial',
    'Análisis Financiero de Software', 'Mantenimiento de Software', 'Multimedia y Diseño Digital',
    'Inglés IX', 'Programación para Móviles II', 'Seguridad de la Información', 'Minería de Datos',
    'Administración de Proyectos de Software', 'Arquitectura Orientada a Servicios',
    'Expresión Oral y Escrita II', 'Estadía'
]

# Convertir las columnas a categorías con el orden especificado
df_pivotado.columns = pd.Categorical(df_pivotado.columns, categories=orden_materias, ordered=True)

# Ordenar las columnas según el orden específico de las materias
df_pivotado = df_pivotado.sort_index(axis=1)

# Construir la ruta absoluta para el archivo de salida
ruta_salida = os.path.join(directorio, 'resultados_completos_filtrados.xlsx')

# Guardar el DataFrame pivotado en un nuevo archivo Excel
df_pivotado.to_excel(ruta_salida, index=False)

print(f'Se ha generado el archivo "{ruta_salida}" con los datos pivotados de los archivos en el directorio, filtrando por "PlanEstudiosClave" igual a "004", y manejando las calificaciones según lo especificado.')
