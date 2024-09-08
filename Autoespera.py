import re
from striprtf.striprtf import rtf_to_text
import pandas as pd
from collections import Counter
import argparse
from datetime import datetime

def main(rtf_file_path):
    # Cargar el archivo RTF
    with open(rtf_file_path, 'r', encoding='utf-8') as file:
        rtf_content = file.read()

    # Convertir el contenido RTF a texto
    text = rtf_to_text(rtf_content)

    # Inicializar listas para almacenar los datos extraídos
    pacientes = []
    numero_historia = []
    cirujanos = []
    diagnostico = []
    cirugia = []
    fecha_inclusion = []  # Nueva lista para almacenar fechas

    # Patrón para identificar el número de historia (NºHª) que comienza con 3 o más dígitos
    historia_regex = re.compile(r'.\d{2}')
    # Expresión regular para detectar una fecha en formato dd/mm/yyyy o similar
    fecha_regex = re.compile(r'\d{2}/\d{2}/\d{4}')
    # Expresión regular para detectar una fecha en formato dd/mm/yyyy estricto
    fecha_regex_ext = re.compile(r'^\d{2}/\d{2}/\d{4}$')

    # Dividir el texto en líneas
    lines = text.split('\n')
    for i, line in enumerate(lines):
        # Filtrar las líneas que contienen solo números de historia y nombre de pacientes
        if historia_regex.match(line.strip()):
            # Añadir solo el número de historia y el paciente (siguiente línea)
            numero_historia.append(line.strip())  # NºHª
            nombre = lines[i + 1].strip()
            if (nombre.isupper() or (nombre[0].isdigit() and str(nombre[2]) != '/')) and (str(nombre) != 'SALVADOR HERRERA, CRISTINA'):
                pacientes.append(nombre)  # Nombre del paciente
                cirujanos.append('')  # Vacío para cirujanos
            else:
                if fecha_regex.search(nombre):
                    pacientes.append('')  # Vacío para pacientes
                    cirujanos.append('Null')
                else:
                    pacientes.append('')  # Vacío para pacientes
                    cirujanos.append(nombre)  # Nombre del cirujano

    # Crear un DataFrame con los datos extraídos
    df = pd.DataFrame({
        'NºHª': numero_historia,
        'PACIENTE': pacientes,
        'CIRUJANO': cirujanos
    })

    # Recorrer el DataFrame y asegurarse de que haya exactamente dos filas vacías entre cirujanos
    i = 0
    while i < len(df) - 1:
        if df['CIRUJANO'].iloc[i] != '':  # Si encontramos un cirujano
            # Contar cuántas filas vacías hay entre este cirujano y el siguiente
            empty_count = 0
            j = i + 1
            while j < len(df) and df['CIRUJANO'].iloc[j] == '':
                empty_count += 1
                j += 1
            
            # Si hay menos de dos filas vacías, insertamos las filas necesarias
            if empty_count < 2:
                # Añadir una fila vacía en todas las columnas en la posición correcta
                df = pd.concat([df.iloc[:j], pd.DataFrame([{'NºHª': '', 'PACIENTE': '', 'CIRUJANO': ''}]), df.iloc[j:]]).reset_index(drop=True)
            
            # Añadir valores a los arrays diagnostico y cirugia
            diagnostico.append(df['NºHª'].iloc[i - 1])  # Añadir el valor de la posición anterior en numero_historia
            cirugia.append(df['PACIENTE'].iloc[i - 1])  # Añadir el valor de la posición anterior en pacientes
            
            i = j  # Saltar las filas vacías ya contadas
        else:
            i += 1

    # Filtra las listas para eliminar los valores vacíos
    pacientes = [paciente for paciente in pacientes if paciente != '']
    numero_historia = [nh for nh in numero_historia if nh != '']

    # Filtrar filas donde el valor de 'PACIENTE' no empieza por una letra
    df_filtered = df[df['PACIENTE'].str.match(r'^[A-Za-z]', na=False)]

    # Actualizar las listas originales con los valores del DataFrame filtrado
    numero_historia = df_filtered['NºHª'].tolist()
    pacientes = df_filtered['PACIENTE'].tolist()

    # Limpiar la lista de cirujanos para eliminar valores vacíos
    cirujanos = [item for item in cirujanos if item]

    # Asegurarse de que cirujanos tenga la misma longitud que pacientes
    while len(cirujanos) < len(pacientes):
        cirujanos.append(None)  # Añadir valores nulos hasta que ambas listas tengan el mismo tamaño
    while len(cirugia) < len(pacientes):
        cirugia.append(None)
    while len(diagnostico) < len(pacientes): 
        diagnostico.append(None)

    # Función para eliminar la primera palabra (número) y manejar valores None
    def eliminar_numero(texto):
        if texto is None:
            return ''  # O puedes elegir otro valor para representar el caso de None
        partes = texto.split(' ', 1)  # Divide el texto en dos partes: número y resto
        return partes[1] if len(partes) > 1 else ''

    # Aplica la función a cada elemento de las listas
    diagnostico = [eliminar_numero(d) for d in diagnostico]
    cirugia = [eliminar_numero(c) for c in cirugia]

    # Agregar la lógica para encontrar las fechas de inclusión
    fecha_inclusion = [None] * len(numero_historia)  # Inicializar con None
    historia_index = {nh: i for i, nh in enumerate(numero_historia)}  # Mapa de índices

    # Para ver los pacientes que tienen dos intervenciones
    # Contar las ocurrencias de cada valor
    contador = Counter(numero_historia)

    # Filtrar los valores que aparecen más de una vez
    repetidos = [valor for valor, count in contador.items() if count > 1]

    # Recorre el texto nuevamente para encontrar fechas
    for i, line in enumerate(lines):
        # Verifica si la línea es un número de historia
        if historia_regex.match(line.strip()) and line.strip() in numero_historia and line.strip() not in repetidos:
            historia_actual = line.strip()
            fechas = []  # Lista para almacenar fechas encontradas
            # Buscar todas las fechas después del número de historia
            for j in range(i + 1, len(lines)):
                fecha_match = fecha_regex_ext.search(lines[j])
                if fecha_match and lines[j-1] in cirujanos:
                    fechas.append(fecha_match.group())
            fecha_inclusion[numero_historia.index(historia_actual)] = fechas[0]

    # Crear un nuevo DataFrame con los datos actualizados
    df_final = pd.DataFrame({
        'Fecha de Inclusion': fecha_inclusion,
        'CIRUJANO': cirujanos,
        'PACIENTE': pacientes,
        'NºHª': numero_historia,
        'DIAGNOSTICO': diagnostico,
        'CIRUGIA': cirugia,
        'CMA': None,
        'OBSERVACIONES': None,
        'FECHA PREVISTA': None
    })

    fecha_actual = datetime.today().date()
    output_path = 'lista_espera-'+str(fecha_actual)+'.xlsx'
    # Guardar el DataFrame en un archivo Excel
    df_final.to_excel(output_path, index=False)

    print(f"Archivo Excel guardado en: {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Procesa un archivo RTF y guarda los resultados en un archivo Excel.')
    parser.add_argument('rtf_file_path', type=str, help='Ruta del archivo RTF de entrada')

    args = parser.parse_args()
    main(args.rtf_file_path)
