import pymupdf  # PyMuPDF
import pandas as pd
import re
import os

def limpiar_encabezados(texto):
    """
    Quita las primeras 3 líneas (encabezado general) y el encabezado de la tabla,
    aunque esté fragmentado en varias líneas.
    """
    # Buscar el primer NRC (5 dígitos consecutivos)
    match = re.search(r"\d{5}", texto)
    if match:
        # Cortar el texto desde el primer NRC encontrado
        return texto[match.start():]
    else:
        # Si no se encuentra, regresar el texto tal cual
        return texto
    
def separar_lineas_por_nrc(texto_sin_encabeza):
    # Separa el texto en líneas, cada una iniciando con un NRC (5 dígitos)
    lineas = []
    patron = re.compile(r'(?=\d{5}\b)')
    for linea in patron.split(texto_sin_encabeza):
        linea = linea.strip()
        if linea:
            # Limpiar dobles espacios que pueden interferir con el regex
            linea_limpia = re.sub(r'\s+', ' ', linea)
            lineas.append(linea_limpia)
    return lineas


def parsear_linea_horario(linea_texto):
    """
    Utiliza una expresión regular para extraer los datos de una línea de texto.
    Está optimizada para la estructura de 9 columnas del PDF.
    """
    # Expresión regular ajustada para ser más robusta.
    # Captura: NRC, Clave, Materia, Sección, Días, Hora, Profesor, Salón, Aclaraciones
    patron = re.compile(
        r"(\d{5})\s+"                   # 1. NRC (5 dígitos exactos)
        r"([A-Z\d]+\s+[A-Z\d]+)\s+"      # 2. Clave (ej. "ICCS 261")
        r"(.+?)\s+"                     # 3. Nombre de la Materia (captura no-golosa)
        r"([A-Z\d]{3})\s+"              # 4. Sección (3 caracteres, ej. "OO1")
        r"([LMAJVSD]+)\s+"              # 5. Días de la semana
        r"(\d{4}-\d{4})\s+"              # 6. Rango de hora (ej. "1100-1159")
        # 7. Profesor (captura robusta de nombres en mayúsculas, con espacios y guiones)
        #r"([A-ZÑÁÉÍÓÚ\s\-]+[A-ZÑÁÉÍÓÚ])\s+"
        r"([A-ZÑÁÉÍÓÚ\s\-.]+)\s+"
        r"(\S+)\s*"                     # 8. Salón (cualquier caracter que no sea espacio)
        r"(.*)$"                        # 9. Aclaraciones (el resto de la línea)
    )

    # coincidencia = patron.search(linea_texto)
    # if not coincidencia:
    #     return None

    coincidencia = patron.search(linea_texto)
    if not coincidencia or len(coincidencia.groups()) < 9:
        print("Línea no válida para el regex:", repr(linea_texto))  # Descomenta para depurar
        return None

    grupos = coincidencia.groups()

    hora_inicio, hora_fin = grupos[5].split('-')

    # Diccionario con los datos limpios y extraídos
    datos = {
        "NRC": grupos[0].strip(),
        "Clave": f"{grupos[1].strip()}",
        "Materia": f"{grupos[2].strip()}",
        "Profesor": grupos[6].strip(),
        "Hora de inicio": hora_inicio.strip(),
        "Hora de fin": hora_fin.strip(),
        "Dia": grupos[4].strip(),
        "Salon": grupos[7].strip(),
        "Aclaraciones": grupos[8].strip()
    }
    return datos

def extraer_pdf_a_excel(pdf_path, excel_path):
    """
    Extrae los cursos de un PDF y los exporta a un archivo Excel.
    """
    if not os.path.exists(pdf_path):
        print(f"❌ Error: El archivo '{pdf_path}' no se encontró.")
        return False

    cursos_encontrados = []
    try:
        doc = pymupdf.open(pdf_path)
    except Exception as e:
        print(f"❌ Error al abrir el PDF: {e}")
        return False

    # Extraer texto de todas las páginas y aplicar las reglas de limpieza
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        texto_pagina = page.get_textpage().extractText()
        text_pag_sin_encabeza = limpiar_encabezados(texto_pagina)
        lineas_utiles = separar_lineas_por_nrc(text_pag_sin_encabeza)

        # Procesar y parsear cada línea útil
        for linea in lineas_utiles:
            
            # Solo procesar líneas que empiecen con un NRC
            if re.match(r"^\d{5}", linea):
                datos_curso = parsear_linea_horario(linea)
                if datos_curso:
                    if datos_curso not in cursos_encontrados:
                        cursos_encontrados.append(datos_curso)
    doc.close()

    if not cursos_encontrados:
        print("❌ No se encontraron cursos con el formato esperado en el PDF.")
        print("Verifica que el PDF no sea una imagen escaneada y que la estructura sea la correcta.")
        return

    print(f"\n✅ ¡Se encontraron {len(cursos_encontrados)} clases únicas en el PDF!")

    # --- 3. SELECCIÓN DE NRC's POR EL USUARIO ---
    nrcs_seleccionados = []
    print("\n--- Selección de Cursos ---")
    print("Ingresa el NRC de cada curso que quieras añadir a tu horario.")
    print("Cuando termines, escribe 'listo' y presiona Enter.\n")

    cursos_para_mostrar = {}
    for curso in cursos_encontrados:
        cursos_para_mostrar.setdefault(curso['NRC'], curso)
            
    for nrc, curso in cursos_para_mostrar.items():
        print(f"➡️  NRC: {curso['NRC']}, Materia: {curso['Materia']}, Profesor: {curso['Profesor']}")

    while True:
        entrada = input("\nIngresa un NRC para agregarlo (o escribe 'listo' para terminar): ").strip().lower()

        if entrada == 'listo':
            if not nrcs_seleccionados:
                print("⚠️ No seleccionaste ningún NRC. El programa terminará.")
                return
            print("\n✅ Selección finalizada. Generando el archivo de Excel...")
            break
        elif entrada.isdigit() and len(entrada) == 5:
            if entrada in nrcs_seleccionados:
                print(f"✔️  El NRC {entrada} ya había sido agregado.")
            elif entrada in cursos_para_mostrar:
                nrcs_seleccionados.append(entrada)
                print(f"👍 NRC {entrada} agregado. Seleccionados hasta ahora: {', '.join(nrcs_seleccionados)}")
            else:
                print(f"❌ El NRC {entrada} no se encontró en la lista de cursos. Intenta de nuevo.")
        else:
            print("❌ Entrada no válida. Por favor, ingresa un NRC de 5 dígitos o la palabra 'listo'.")

    # --- 4. EXPORTACIÓN A EXCEL ---
    clases_a_exportar = [curso for curso in cursos_encontrados if curso['NRC'] in nrcs_seleccionados]
    df = pd.DataFrame(clases_a_exportar)
    df = df[['NRC', 'Materia', 'Profesor', 'Hora de inicio', 'Hora de fin', 'Dia', 'Salon']]

    try:
        df.to_excel(excel_path, index=False, header=False)
        print(f"\n🎉 ¡Éxito! Se ha creado el archivo '{excel_path}' con todas las clases de los NRCs que seleccionaste.")
        return True
    except Exception as e:
        print(f"\n❌ Ocurrió un error al guardar el archivo de Excel: {e}")

    # if not cursos_encontrados:
    #     print("❌ No se encontraron cursos para exportar.")
    #     return False

    # columnas = ['NRC', 'Materia', 'Profesor', 'Hora de inicio', 'Hora de fin', 'Dia', 'Salon']
    # df = pd.DataFrame(cursos_encontrados)
    # df = df[columnas]
    # try:
    #     df.to_excel(excel_path, index=False, header=False)
    #     print(f"\n🎉 ¡Éxito! Se ha creado el archivo '{excel_path}' con todas las clases encontradas.")
    #     return True
    # except Exception as e:
    #     print(f"\n❌ Ocurrió un error al guardar el archivo de Excel: {e}")
    #     return False

# def main():
#     """
#     Función principal que orquesta la extracción, selección y exportación usando PyMuPDF.
#     """
#     # --- 1. CONFIGURACIÓN INICIAL ---
#     pdf_path = 'PA_PRIMAVERA.pdf'
#     excel_path = 'Horarios_Seleccionados.xlsx'

#     if not os.path.exists(pdf_path):
#         print(f"❌ Error: El archivo '{pdf_path}' no se encontró.")
#         print("Asegúrate de que el PDF esté en la misma carpeta que este script.")
#         return

#     # --- 2. EXTRACCIÓN Y PROCESAMIENTO CON PyMuPDF ---
#     print(f"📄 Leyendo el PDF con PyMuPDF: '{pdf_path}'...")
#     cursos_encontrados = []
    
#     try:
#         doc = pymupdf.open(pdf_path)
#     except Exception as e:
#         print(f"❌ Error al abrir el PDF. Puede que esté dañado o protegido. Error: {e}")
#         return

#     # Extraer texto de todas las páginas y aplicar las reglas de limpieza
#     for page_num in range(len(doc)):
#         page = doc.load_page(page_num)
#         texto_pagina = page.get_textpage().extractText()
#         text_pag_sin_encabeza = limpiar_encabezados(texto_pagina)
#         lineas_utiles = separar_lineas_por_nrc(text_pag_sin_encabeza)
        
#         # Procesar y parsear cada línea útil
#         for linea in lineas_utiles:
            
#             # Solo procesar líneas que empiecen con un NRC
#             if re.match(r"^\d{5}", linea):
#                 datos_curso = parsear_linea_horario(linea)
#                 if datos_curso:
#                     if datos_curso not in cursos_encontrados:
#                         cursos_encontrados.append(datos_curso)

#     doc.close()

#     if not cursos_encontrados:
#         print("❌ No se encontraron cursos con el formato esperado en el PDF.")
#         print("Verifica que el PDF no sea una imagen escaneada y que la estructura sea la correcta.")
#         return

#     print(f"\n✅ ¡Se encontraron {len(cursos_encontrados)} clases únicas en el PDF!")

#     # --- 3. SELECCIÓN DE NRC's POR EL USUARIO ---
#     nrcs_seleccionados = []
#     print("\n--- Selección de Cursos ---")
#     print("Ingresa el NRC de cada curso que quieras añadir a tu horario.")
#     print("Cuando termines, escribe 'listo' y presiona Enter.\n")

#     cursos_para_mostrar = {}
#     for curso in cursos_encontrados:
#         cursos_para_mostrar.setdefault(curso['NRC'], curso)
            
#     for nrc, curso in cursos_para_mostrar.items():
#         print(f"➡️  NRC: {curso['NRC']}, Materia: {curso['Materia']}, Profesor: {curso['Profesor']}")

#     while True:
#         entrada = input("\nIngresa un NRC para agregarlo (o escribe 'listo' para terminar): ").strip().lower()

#         if entrada == 'listo':
#             if not nrcs_seleccionados:
#                 print("⚠️ No seleccionaste ningún NRC. El programa terminará.")
#                 return
#             print("\n✅ Selección finalizada. Generando el archivo de Excel...")
#             break
#         elif entrada.isdigit() and len(entrada) == 5:
#             if entrada in nrcs_seleccionados:
#                 print(f"✔️  El NRC {entrada} ya había sido agregado.")
#             elif entrada in cursos_para_mostrar:
#                 nrcs_seleccionados.append(entrada)
#                 print(f"👍 NRC {entrada} agregado. Seleccionados hasta ahora: {', '.join(nrcs_seleccionados)}")
#             else:
#                 print(f"❌ El NRC {entrada} no se encontró en la lista de cursos. Intenta de nuevo.")
#         else:
#             print("❌ Entrada no válida. Por favor, ingresa un NRC de 5 dígitos o la palabra 'listo'.")

#     # --- 4. EXPORTACIÓN A EXCEL ---
#     clases_a_exportar = [curso for curso in cursos_encontrados if curso['NRC'] in nrcs_seleccionados]
#     df = pd.DataFrame(clases_a_exportar)
#     df = df[['NRC', 'Materia', 'Profesor', 'Hora de inicio', 'Hora de fin', 'Dia', 'Salon']]

#     try:
#         df.to_excel(excel_path, index=False, header=False)
#         print(f"\n🎉 ¡Éxito! Se ha creado el archivo '{excel_path}' con todas las clases de los NRCs que seleccionaste.")
#     except Exception as e:
#         print(f"\n❌ Ocurrió un error al guardar el archivo de Excel: {e}")

# if __name__ == '__main__':
#     main()