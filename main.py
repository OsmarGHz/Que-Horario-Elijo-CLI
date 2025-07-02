import os.path
import os
#pip install pandas
#pip install openpyxl
import openpyxl
import pandas as pd
import itertools
import math

from datetime import datetime, time
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile"
]

materias = {}
resultados = {}
archivoExcel = None

def convertirHora(stringHora):
    if isinstance(stringHora, time):
        #Si aparece un ?, favor de quitar la fecha de la celda de excel afectada (workaround momentaneo, o eso espero)
        return stringHora
    if isinstance(stringHora, float) or isinstance(stringHora, int):
        # Si viene como número (por ejemplo, 700), conviértelo a string
        stringHora = str(int(stringHora)).zfill(4)
    if isinstance(stringHora, str):
        stringHora = stringHora.strip()
        # Si es tipo '0700' o '700'
        if stringHora.isdigit() and (len(stringHora) == 4 or len(stringHora) == 3):
            stringHora = stringHora.zfill(4)
            return time(int(stringHora[:2]), int(stringHora[2:]))
        # Prueba varios formatos comunes
        for formatTime in ("%H:%M", "%H:%M:%S", "%H:%M:%S.%f", "%I:%M%p"):
            try:
                return datetime.strptime(stringHora, formatTime).time()
            except ValueError:
                continue
    # Si no se pudo interpretar, regresa None
    return None

def equivalStr(palabra1,palabra2):
    if(palabra1.lower()==palabra2.lower()):
        return True
    else: return False

def igualStr(palabra1, palabra2):
    if(palabra1==palabra2):
        return True
    else: return False

def imprimirMensajeListo():
    print("""
    Listo!
    """)

def introPrograma():
    print("""
    ¡Bienvenido a ¿Qué horario elijo?!
    En este programa, tú agregas tus opciones de clases por separado para este semestre, y nosotros haremos la magia!
    Con nuestros calendarios que puedes agregar a Google Calendar, te ahorrarás tiempo!
    """)

def mostrarAyuda(seccion): #Mostrar aiuda
    if equivalStr(seccion,"principal"):
        print("""
        Permítenos darte nuestro menú de opciones:
            account \t\t Agrega, elimina, o ve el correo de google que agregaste.
            help \t\t Muestra esta ayuda
            classes \t\t Guarda o elimina tus clases
            calendars \t\t Administra tus calendarios, y genera nuevos en base a tus clases (dentro, podrás pushear tu calendario a Calendar)
            exit \t\t Simplemente, sale del programa
        """)
    elif equivalStr(seccion,"account"):
        print("""
        Estás en principal -> account (sección para gestionar las cuentas). Aquí puedes escribir:
            connect \t\t Conecta esta app con tu cuenta de Google
            disconnect \t\t Desconecta tu cuenta de Google, de esta app
            seeInfo \t\t Ver el correo que tienes agregado
            help \t\t Muestra esta ayuda
            return \t\t Regresa a la sección anterior
            exit \t\t Simplemente, sale del programa
        """)
    elif equivalStr(seccion,"calendars"):
        print("""
        Estás en principal -> calendars (sección para gestionar los calendarios y pushearlos a Calendar). Aquí puedes escribir:
            printClasses
            generateCalendars
            chargeCals
            pushNviewCals
            help
            return
            exit
        """)

def regresar():
    print("""
    Regresando a la sección anterior... (Usa help para saber la ubicacion)
    """)
    return 0

def mostrarDespedida():
    print("""
    Nos vemos pronto!
    """)
    return -1

def mostrarError():
    print("""
    No se encuentra el comando ingresado, intente otra cosa.
    """)
    return 1

def mostrarErrorNumerico():
    print("""
    El numero ingresado está fuera de rango, o la entrada no es valida
    """)

def connect():
    print("""
    \t--- Conéctate con Google Calendar. ---\t
    """)
    creds = None
    if os.path.exists("token.json"):
        print("Ya hay una cuenta conectada.")
        resp = input("¿Quieres reemplazarla? (s/n): ").strip().lower()
        if resp != "s":
            print("""
            Operación cancelada.
            """)
            return
        else:
            os.remove("token.json")
            print("""
            Cuenta anterior desconectada.
            """)
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
        token.write(creds.to_json())
    print("""
    ¡Cuenta conectada exitosamente!
    """)

def disconnect():
    print("""
    \t--- Desconecta tu cuenta. ---\t
    """)

    if not os.path.exists("token.json"):
        print("""
        No hay cuentas conectadas.
        """)
    else:
        resp = input("¿Seguro que quieres cerrar tu sesión? (s/n): ").strip().lower()
        if resp != "s":
            print("""
            Operación cancelada.
            """)
            return
        else:
            os.remove("token.json")
            print("""
            ¡Cuenta DESconectada exitosamente!
            """)

def seeInfo():
    print("""
    \t--- Verificando tus datos... ---\t
    """)
    if not os.path.exists("token.json"):
        print("""
        No hay cuentas conectadas.
        """)
        return
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if creds and creds.valid:
        try:
            service = build('oauth2', 'v2', credentials=creds)
            user_info = service.userinfo().get().execute() # pylint: disable=no-member -> quita esto si tienes otros problemas
            email = user_info.get("email", "Correo no disponible")
            name = user_info.get("name", "Nombre no disponible")
            print(f"""
            Cuenta conectada: {name} <{email}>
            """)
        except Exception as e:
            print("No se pudo obtener la información de la cuenta.")
            print(e)
    else:
        print("""
        Lo sentimos, no pudimos obtener la información de la cuenta.
        Has ingresado correctamente tu cuenta?
        Has intentado cerrar sesion y volverla a abrir?
        """)

def selectActAccount(entrada): #Retorna 1 si no es necesario salir. 0 si regresa 1 nivel, -1 si sale abruptamente el usuario, y 2 si está regresando de connect, seeInfo, u otra funcion
    if equivalStr(entrada,"help"):
        mostrarAyuda("account")
    elif equivalStr(entrada,"connect"):
        connect()
    elif equivalStr(entrada,"disconnect"):
        disconnect()
    elif equivalStr(entrada,"seeInfo"):
        seeInfo()
    elif equivalStr(entrada,"return"):
        return regresar()
    elif equivalStr(entrada,"exit"):
        return mostrarDespedida()
    else:
        return mostrarError()
    return 2

def account():
    showHeader = True
    retorno = 1
    while retorno > 0:
        if showHeader is True:
            print("""
            \t\t-----Módulo de conexión con Google Calendar-----\t\t
            """)
            mostrarAyuda("account")
        entrada = input(" -> ")
        retorno = selectActAccount(entrada)
        if retorno == 2:
            showHeader = True
        else:
            showHeader = False

    return retorno

def limpiarMaterias():
    global materias
    print("""
        Borrando...""")
    materias = {}
    print("""
        Materias anteriores borradas!
    """)

def abrirExcel(ruta=""):
    global archivoExcel
    if not os.path.exists(ruta):
        print(f"Error: No se encontró el archivo '{ruta}'. Ingree otro nombre o intente de nuevo.")
        return 0
    try:
        archivoExcel = pd.read_excel(ruta, header=None)
        #print("Primeras filas detectadas:")
        #print(archivoExcel.head())
        return 1
    except Exception as e:
        print(f"Error al abrir el archivo: {e}")
        return 0

def definirRuta(folder="", filename=""):
    if not equivalStr(folder, ""):
        # Verifica si la carpeta existe
        if not os.path.exists(folder):
            print(f"Error: La carpeta '{folder}' no existe.")
            return 0
        # Une la ruta de forma compatible con cualquier sistema operativo
        ruta = os.path.join(folder, filename)
        return abrirExcel(ruta)

def procesarExcel():
    nrc = None
    materia = None
    profesor = None
    inicio = None
    fin = None
    dia = None
    salon = None

    if archivoExcel is None:
        print("Error: El archivo Excel no se ha cargado correctamente.")
        return
    for _, row in archivoExcel.iterrows():
        row0 = row[0]
        row1 = row[1]
        row2 = row[2]
        row3 = convertirHora(row[3])
        row4 = convertirHora(row[4])
        row5 = row[5]
        row6 = row[6]

        if pd.isna(row1):
            if pd.isna(materia):
                print("Lo sentimos, hubo un error durante la lectura del archivo")
                return
        else:
            nrc = int(row0)
            materia = row1
            profesor = row2

        inicio = convertirHora(row3)
        fin = convertirHora(row4)
        dia = str(row5).upper()
        salon = row6

        # Si la materia no está en el diccionario, agrégala
        if materia not in materias:
            materias[materia] = []

        # Busca si ya existe una opción con ese NRC y profesor
        opciones = materias[materia]
        opcion = next((o for o in opciones if o['nrc'] == nrc and o['profesor'] == profesor), None)
        if not opcion:
            opcion = {
                "nrc": nrc,
                "profesor": profesor,
                "horarios": []
            }
            opciones.append(opcion)

        # Agrega el horario a la opción correspondiente
        opcion["horarios"].append({
            "inicio": inicio,
            "fin": fin,
            "dia": dia,
            "salon": salon
        })

    imprimirMensajeListo()

def imprimirMaterias():
    print("\n")
    for materia, opciones in materias.items():
        print(f"Materia: {materia}")
        for opcion in opciones:
            print(f"  NRC: {opcion['nrc']} - Profesor: {opcion['profesor']}")
            for h in opcion["horarios"]:
                # Formatea la hora para que se vea bonito
                inicio = h['inicio'].strftime("%H:%M") if h['inicio'] else "?"
                fin = h['fin'].strftime("%H:%M") if h['fin'] else "?"
                print(f"    {h['dia']}: {inicio} - {fin} en {h['salon']}")
        print("-" * 40)

def classes():
    print("""
    \t\t-----Módulo de Manejo de Clases-----\t\t
          
    Este módulo es, de momento, de autorretroceso
    """)

    fileName = input(" Ingrese el nombre del archivo: ")

    if definirRuta("SchoolSubjectList",fileName):
        if materias != {} and input(" Existen materias ya registradas anteriormente. Desea borrarlas antes de analizar el nuevo archivo? \n(Ingrese \"si\", o cualquier otra cosa para cancelar): ").lower() == "si":
            limpiarMaterias()
        procesarExcel()


def horas_entre(t1, t2):
    """Devuelve la diferencia en horas entre dos objetos time."""
    dt1 = datetime.combine(datetime.today(), t1)
    dt2 = datetime.combine(datetime.today(), t2)
    #return (dt2 - dt1).total_seconds() / 3600
    return math.ceil((dt2 - dt1).total_seconds() / 3600)

def horarios_chocan(horarios):
    """Verifica si hay choques de horario en una lista de horarios."""
    eventos = []
    for h in horarios:
        eventos.append((h['dia'], h['inicio'], h['fin']))
    # Agrupa por día
    por_dia = {}
    for dia, inicio, fin in eventos:
        if dia not in por_dia:
            por_dia[dia] = []
        por_dia[dia].append((inicio, fin))
    # Para cada día, verifica si hay traslapes
    for bloques in por_dia.values():
        bloques.sort()
        for i in range(len(bloques)-1):
            if bloques[i][1] > bloques[i+1][0]:  # fin > inicio siguiente
                return True
    return False

def calcular_horas(combinacion):
    """Calcula horas de clase y horas de permanencia en la uni para la semana."""
    por_dia = {}
    horas_clase = 0
    for opcion in combinacion:
        for h in opcion['horarios']:
            dia = h['dia']
            inicio = h['inicio']
            fin = h['fin']
            if not (isinstance(inicio, time) and isinstance(fin, time)):
                continue
            horas_clase += horas_entre(inicio, fin)
            if dia not in por_dia:
                por_dia[dia] = []
            por_dia[dia].append((inicio, fin))
    # Para cada día, calcula permanencia
    horas_permanencia = 0
    for bloques in por_dia.values():
        bloques.sort()
        primero = min(b[0] for b in bloques)
        ultimo = max(b[1] for b in bloques)
        horas_permanencia += horas_entre(primero, ultimo)
    return horas_clase, horas_permanencia

def generar_horarios(materias,min_len_materias):
    materia_keys = list(materias.keys())
    resultados = []
    for r in range(len(materia_keys), min_len_materias - 1, -1):  # n, n-1, ..., min_len_materias
        for subconjunto in itertools.combinations(materia_keys, r):
            materia_opciones = [materias[m] for m in subconjunto]
            for combinacion in itertools.product(*materia_opciones):
                # Junta todos los horarios de la combinacion
                todos_horarios = []
                for opcion in combinacion:
                    todos_horarios.extend(opcion['horarios'])
                if horarios_chocan(todos_horarios):
                    continue  # descarta combinaciones con choques
                horas_clase, horas_permanencia = calcular_horas(combinacion)
                resultados.append({
                    "materias": subconjunto,
                    "combinacion": combinacion,
                    "horas_clase": horas_clase,
                    "horas_permanencia": horas_permanencia
                })
    return resultados

def genCalendars():
    global resultados
    max_len_materias = len(list(materias.keys()))
    min_len_materias = 0
    while not (0 < min_len_materias <= max_len_materias):
        min_len_materias = int(input(f"""Ingrese la cantidad minima de materias para generar las combinaciones de calendarios \n(Las combinaciones irán desde esa cantidad mínima, hasta {max_len_materias}): """))
        if not (0 < min_len_materias <= max_len_materias):
            mostrarErrorNumerico()
        else:
            resultados = generar_horarios(materias, min_len_materias)
            if len(resultados) == 0:
                print("""
    No se pudo generar ningún calendario sin choques de horario con los criterios dados.
    Intenta cambiar tus materias, opciones o el mínimo de materias.
""")
            else:
                print(f"""
    Se han generado {len(resultados)} calendarios en base a tus materias!
    Puedes verlos con PushNViewCals en la seccion de calendarios!
""")


def ordenar_resultados(resultados):
    # Ordena primero por número de materias (desc), luego por menor permanencia
    return sorted(resultados, key=lambda x: (-len(x["materias"]), x["horas_permanencia"]))

def vistaPushNViewCals():
    global resultados
    if not resultados:
        print("No hay calendarios generados. Usa 'generateCalendars' primero.")
        return

    resultados_ordenados = ordenar_resultados(resultados)
    for idx, r in enumerate(resultados_ordenados):
        r["id"] = idx + 1  # Asigna un id único

    pos = 0
    while True:
        r = resultados_ordenados[pos]
        print(f"\nCalendario #{r['id']} - Materias: {len(r['materias'])}")
        print(f"Materias: {', '.join(r['materias'])}")
        print(f"Horas de clase/semana: {r['horas_clase']:.2f}")
        print(f"Horas de permanencia/semana: {r['horas_permanencia']:.2f}")
        print("Detalle:")
        for i, opcion in enumerate(r["combinacion"]):
            print(f"  {r['materias'][i]} - NRC {opcion['nrc']} - {opcion['profesor']}")
            for h in opcion["horarios"]:
                ini = h['inicio'].strftime("%H:%M") if h['inicio'] else "?"
                fin = h['fin'].strftime("%H:%M") if h['fin'] else "?"
                print(f"    {h['dia']}: {ini} - {fin} en {h['salon']}")
        print("\nComandos: [d]erecha, [i]zquierda, [g]uardar excel, [p]ush a Google Calendar, [q]uitar")

        cmd = input("-> ").strip().lower()
        if cmd == "d":
            pos = (pos + 1) % len(resultados_ordenados)
        elif cmd == "i":
            pos = (pos - 1) % len(resultados_ordenados)
        elif cmd == "g":
            guardar_en_excel(r)
        elif cmd == "p":
            push_a_google_calendar(r)
        elif cmd == "q":
            print("Saliendo del carrusel de calendarios.")
            break
        else:
            print("Comando no reconocido.")

def guardar_en_excel(calendario):
    print(f"Guardando calendario #{calendario['id']} en Excel...\n")

    filas = []
    for i, opcion in enumerate(calendario["combinacion"]):
        materia = calendario["materias"][i]
        nrc = opcion["nrc"]
        profesor = opcion["profesor"]
        for h in opcion["horarios"]:
            ini = h['inicio'].strftime("%H:%M") if h['inicio'] else "?"
            fin = h['fin'].strftime("%H:%M") if h['fin'] else "?"
            filas.append({
                "Materia": materia,
                "NRC": nrc,
                "Profesor": profesor,
                "Día": h["dia"],
                "HoraInicio": ini,
                "HoraFin": fin,
                "Salón": h["salon"]
            })
    df = pd.DataFrame(filas)
    # Asegura que la carpeta exista
    folder = "Schedules"
    if not os.path.exists(folder):
        os.makedirs(folder)
    nombre_archivo = input(f"Ingrese el nombre para el archivo Excel (sin extensión) o deje vacío para usar 'Horario_{calendario['id']}': ").strip()
    if not nombre_archivo:
        nombre_archivo = f"Horario_{calendario['id']}"
    nombre_archivo = os.path.join(folder, nombre_archivo + ".xlsx")
    try:
        df.to_excel(nombre_archivo, index=False)
        print(f"¡Calendario guardado como '{nombre_archivo}'!")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")

def cargar_calendario_desde_excel():
    folder = "Schedules"
    if not os.path.exists(folder):
        print("No existe la carpeta 'Schedules'.")
        return None
    archivos = [f for f in os.listdir(folder) if f.endswith(".xlsx")]
    if not archivos:
        print("No hay archivos de calendario en 'Schedules'.")
        return None
    print("Archivos disponibles:")
    for idx, archivo in enumerate(archivos):
        print(f"{idx+1}. {archivo}")
    seleccion = input("Selecciona el número del archivo a cargar: ")
    try:
        idx = int(seleccion) - 1
        if idx < 0 or idx >= len(archivos):
            print("Selección inválida.")
            return None
        archivo_path = os.path.join(folder, archivos[idx])
        df = pd.read_excel(archivo_path)
        # Convierte el DataFrame a la estructura de resultados
        materias = list(df["Materia"].unique())
        combinacion = []
        for materia in materias:
            grupo = df[df["Materia"] == materia].iloc[0]
            opcion = {
                "nrc": grupo["NRC"],
                "profesor": grupo["Profesor"],
                "horarios": []
            }
            for _, row in df[df["Materia"] == materia].iterrows():
                inicio = convertirHora(row["HoraInicio"])
                fin = convertirHora(row["HoraFin"])
                opcion["horarios"].append({
                    "inicio": inicio,
                    "fin": fin,
                    "dia": row["Día"],
                    "salon": row["Salón"]
                })
            combinacion.append(opcion)
        # Calcula horas
        horas_clase, horas_permanencia = calcular_horas(combinacion)
        calendario = {
            "id": 0,  # Se puede reasignar después
            "materias": materias,
            "combinacion": combinacion,
            "horas_clase": horas_clase,
            "horas_permanencia": horas_permanencia
        }
        return calendario
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None

def chargeCalendars():
    global resultados
    calendario = cargar_calendario_desde_excel()
    if calendario:
        resultados = [calendario]
        print("¡Calendario cargado! Ahora puedes verlo en el carrusel con PushNViewCals.")

def push_a_google_calendar(calendario):
    # Implementa aquí la lógica para enviar el calendario a Google Calendar
    print(f"Enviando calendario #{calendario['id']} a Google Calendar... (implementa aquí)")

def selectActCalendars(entrada): #Retorna 1 si no es necesario salir. 0 si regresa 1 nivel, -1 si sale abruptamente el usuario, y 2 si está regresando de connect, seeInfo, u otra funcion
    if equivalStr(entrada,"help"):
        mostrarAyuda("calendars")
    elif equivalStr(entrada,"printClasses"):
        imprimirMaterias()
    elif equivalStr(entrada,"generateCalendars"):
        genCalendars()
    elif equivalStr(entrada,"chargeCals"):
        chargeCalendars()
    elif equivalStr(entrada,"pushNviewCals"):
        vistaPushNViewCals()
    elif equivalStr(entrada,"return"):
        return regresar()
    elif equivalStr(entrada,"exit"):
        return mostrarDespedida()
    else:
        return mostrarError()
    return 2
    

def calendars():
    showHeader = True
    retorno = 1
    while retorno > 0:
        if showHeader is True:
            print("""
            \t\t-----Módulo de Manejo de calendarios-----\t\t
            """)
            mostrarAyuda("calendars")
        entrada = input(" -> ")
        retorno = selectActCalendars(entrada)
        if retorno == 2:
            showHeader = True
        else:
            showHeader = False

    return retorno


def selectFunction(entrada):
    if equivalStr(entrada,"help"):
        mostrarAyuda("principal")
    elif equivalStr(entrada,"account"):
        if account() == -1:
            return -1
    elif equivalStr(entrada,"classes"):
        classes()
    elif equivalStr(entrada,"calendars"):
        calendars()
    elif equivalStr(entrada,"exit"):
        return mostrarDespedida()
    else:
        return mostrarError()
    return 2

def menuCiclado():
    entrada = ""
    showHeader = True
    retorno = 1
    while retorno > 0:
        if showHeader is True:
            introPrograma()
            mostrarAyuda("principal")
        entrada = input(" -> ")
        retorno = selectFunction(entrada)
        if retorno == 2:
            showHeader = True
        else:
            showHeader = False

def main():
    menuCiclado()

if __name__  == "__main__":
    main()