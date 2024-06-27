import pandas as pd
from datetime import datetime

def registrar_actividad(actividad, descripcion, urgencia, notificado_por, realizado_por):
    fecha_actual = datetime.now().strftime("%Y-%m-%d")
    hora_actual = datetime.now().strftime("%H:%M")
    registro = {
        "Fecha": fecha_actual,
        "Hora": hora_actual,
        "Urgencia": urgencia,
        "Actividad": actividad,
        "Descripción": descripcion,
        "Notificado por": notificado_por,
        "Realizado por": realizado_por
    }
    return registro

def exportar_actividades(actividades):
    archivo = "actividades.xlsx"
    try:
        df_existente = pd.read_excel(archivo)
        df_nuevo = pd.DataFrame(actividades)
        df_final = pd.concat([df_existente, df_nuevo])
    except FileNotFoundError:
        df_final = pd.DataFrame(actividades)
    df_final.sort_values(by=['Fecha', 'Hora'], inplace=True)
    df_final.to_excel(archivo, index=False)

print(
    "Ingrese uno por uno los soportes o actividades realizadas presionando ENTER después de cada uno.\nEjemplo\nFulano:falla de la impresora de la oficina tal...\nAl final del día, presione enter hasta que se cierre el programa o aparezca un mensaje de que los datos han sido exportados al archivo de excel"
)

actividades = []
while True:
    actividad = input("Ingrese la actividad: ")
    if actividad == "":
        break
    descripcion = input("Ingrese una breve descripción de la actividad: ")
    urgencia = input("Ingrese la urgencia del soporte: ")
    notificado_por = input("Ingrese quién notificó la actividad: ")
    realizado_por = input("Ingrese quién realizó la actividad: ")
    registro = registrar_actividad(actividad, descripcion, urgencia, notificado_por, realizado_por)
    actividades.append(registro)

exportar_actividades(actividades)
print(
    f"Ya se han exportado los soportes a un archivo de excel con la fecha de hoy en el nombre \n Gracias por usar este programa creado por el Pasante de Sistemas Carlos Paez")
