import sys
import shutil
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import copy
import os
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage

#-------------------- Configuración de Usuario -----------------------
# Curso
CURSO = "2021/2022"

# Ruta fichero Notas_Alumnos.xlsx
NOTAS_ALUMNOS_EXCEL_PATH = r'Inputs\Notas_Alumnos.xlsx'

# Ruta fichero Plantilla_Notas.docx
PLANTILLA_WORD_PATH = r'Inputs\Plantilla_Notas.docx'

# Ruta de salida
OUTPUT_PATH = r'Outputs'

# Ruta temporal

TEMP_PATH = r'Temp'

# Diccionario asignaturas
dict_asig = {
    'LENGUA CASTELLANA Y LITERATURA':   'Lengua Castellana y Literatura',
    'BIOLOGIA':                         'Biología',
    'GEOGRAFIA E HISTORIA':             'Geografía e Historia',
    'MATEMATICAS':                      'Matemáticas',
    'INGLES':                           'Inglés',
    'EDUCACION FISICA':                 'Educación Física',
    'ETICA':                            'Ética',
    'CULTURA CLASICA':                  'Cultura clásica',
    'MUSICA':                           'Música',
    'TECNOLOGIA':                       'Tecnología',
    'EDUCACION PLASTICA':               'Educación Plástica',
    'FRANCES':                          'Francés',
}

# COLORES
SUSPENSO_COLOR = 'ec7c7b'
APROBADO_COLOR = 'fbe083'
NOTABLE_COLOR = '4db4d7'
SOBRESALIENTE_COLOR = '48bf91'

#-------------------- /Configuración de Usuario ----------------------

# Detección de errors en ficheros Excel (alummnos con asignaturas sin asignar,
# con asignaturas repetidas y notas fuera de rango)
def DeteccionErrores(ex_df):

    err_1, err_2, err_3 = False, False, False

    # Eliminar valores duplicados de columna "ASIGNATURA" (Ordenar Alfabeticamente)
    alumnos_list = sorted(list(ex_df['NOMBRE'].astype(str).drop_duplicates()))

    # Eliminar valores duplicados de columna "ASIGNATURA" (Ordenar Alfabeticamente)
    asignatura_list = sorted(list(ex_df['ASIGNATURA'].astype(str).drop_duplicates()))

    # Iterar alumnos y asignaturas
    for al in alumnos_list:
        for asig in asignatura_list:
            # Filtrar df por alumnos y asignatura
            filt_al_as_df = ex_df[(ex_df['NOMBRE'] == al) & (ex_df['ASIGNATURA'] == asig)]

            # Error 1: El alumno no tiene la asignatura asignada
            if(len(filt_al_as_df) == 0):
                print("ERROR 1: El alumnno", al, "no tiene la asignatura", asig, "asignada!!")
                err_1 = True
            # Error 2: El alumno tiene la asignatura repetida
            elif (len(filt_al_as_df) > 1):
                print("ERROR 2: El alumno", al, "tiene la asignatura", asig, "repetida",
                      len(filt_al_as_df), "veces!!")
                err_2 = True
    # Iterar por las difentes filas del Dataframe ex_df
    for index, row in ex_df.iterrows():
        trin_list = ['NOTA T1', 'NOTA T2', 'NOTA T3']
        # Acceder al nombre del alumno y la asignatura de la fila actual
        al = row['NOMBRE']
        asig = row['ASIGNATURA']
        # Iterar sobre trimestres
        for trin in trin_list:
            if not((row[trin] >= 0.0) and (row[trin] <= 10.0)):
                print("ERROR 3: El alumno", al, "tiene el campo \""+ trin + "\" "
                "de la asignatura", asig, "fuera de rango ("+ str(row[trin])+ "!!)")
                err_3 = True

    # Si se detecta algun error. Se detiene el programa
    if (err_1 == True) or (err_2 == True) or (err_3 == True):
        print("")
        print("CORREGIR ERRORES DE EJECUCIÓN ANTES DE CONTINUAR..")
        sys.exit(1)
    #Ningun error detectado
    else:
        print("Ningún error detectado... \nComenzando a procesar")
        

# Rutina para eliminar tildes
def EliminarTildes(texto):
    # Diccionario de tildes
    tildes_dict = {
        'Á': 'A',
        'É': 'E',
        'Í': 'I',
        'Ó': 'O',
        'Ú': 'U',
    }
    #copiar texto
    texto_sin_tildes = texto

    # Iterar sobre tildes_dict
    for key in tildes_dict:
        texto_sin_tildes = texto_sin_tildes.replace(key, tildes_dict[key])
        
    return texto_sin_tildes

# Rutinas para eliminar y crear carpetas
def EliminarCrearCarpetas(ruta):
    #Si la ruta existe, Eliminar carpeta
    if os.path.exists(ruta):
        shutil.rmtree(ruta)

    #Despues de eliminar - Crear carpeta
    os.mkdir(ruta)
    
# Rutinas para eliminar carpetas
def EliminarCarpetas(ruta):
    #Si la ruta existe, Eliminar carpeta
    if os.path.exists(ruta):
        shutil.rmtree(ruta)

# Rutina para obtener calificacion 
def ObtenerCalificacion(nota_media):
    # Obtener calificacion
    if (nota_media < 5.0):
        calif = "SUSPENSO"
        color_calif = SUSPENSO_COLOR
    elif (nota_media < 7.0):
        calif = "APROBADO"
        color_calif = APROBADO_COLOR
    elif (nota_media < 9.0):
        calif = "NOTABLE"
        color_calif = NOTABLE_COLOR
    else:
        calif = "SOBRESALIENTE"
        color_calif = SOBRESALIENTE_COLOR
    
    return calif, color_calif
    
# Rutina para obtener nota final y calificación
def ObtenerNotaFinal(asignatura_dict):
    new_asignatura_dict = copy.deepcopy(asignatura_dict)
    TRIMESTRE_LIST = ['t1', 't2', 't3' ] 
    
    # Obtener nota final
    nota_media = 0
    for trim in TRIMESTRE_LIST:
        nota_media += new_asignatura_dict[trim]
    nota_media /= 3
    
    new_asignatura_dict['nota_total'] = round((nota_media),1)
    
    calif, color_calif = ObtenerCalificacion(nota_media)
    new_asignatura_dict['calificacion'] = calif    
    new_asignatura_dict['color'] = color_calif
    
    return new_asignatura_dict

# Rutina para crear grafico circular de calificaciones
def CrearGraficoCircular(asignatura_list, nombre_alumno):
    # Listas de calificaciones
    CAL_LIST = ['SUSPENSO', 'APROBADO', 'NOTABLE', 'SOBRESALIENTE']
    CAL_COLOR_LIST = [SUSPENSO_COLOR, APROBADO_COLOR, NOTABLE_COLOR, SOBRESALIENTE_COLOR]
    
    # Contar calificaciones
    cal_cont_list = [0] * len(CAL_LIST)
    for cal_idx, cal in enumerate(CAL_LIST):
        for asig in asignatura_list:
            #Evaluar calificacion y actualizar contador
            if(asig['calificacion'] == cal):
                cal_cont_list[cal_idx] += 1
                
    # Calcular porcentaje de calificaciones
    if sum(cal_cont_list) != 0:  
        cal_porc_list = [round(100 * cal_cont / sum(cal_cont_list), 2) for cal_cont in cal_cont_list]  
    else:
        cal_porc_list = [0] * len(CAL_LIST)  
    
    # Calcular porcentaje de calificaciones
    if sum(cal_cont_list) != 0:
        cal_porc_list = [round(100 * cal_cont / sum(cal_cont_list), 2) for cal_cont in cal_cont_list]
    else:
        cal_porc_list = [0] * len(CAL_LIST)         
          
    # Crear lista de colores para graficas circulares
    cal_color_list = []
    for col in CAL_COLOR_LIST:
        cal_color_list.append('#' + col)
    
    # Iterar alreves sobre CAL_LIST y Eliminar valores nulos
    cal_list = copy.deepcopy(CAL_LIST)
    for cal_idx, _ in reversed(list(enumerate(cal_list))):
        # Evaluar si hay un valores NULO y se elimina
        if (cal_porc_list[cal_idx] == 0):
            del cal_porc_list[cal_idx]
            del cal_color_list[cal_idx]
            del cal_list[cal_idx]
                 
    # Crear grafico circular en pantalla. autopct="%.2f%%" decimas y porcentaje
    if cal_porc_list:
        fig1, ax1 = plt.subplots()
        ax1.pie(cal_porc_list, labels=cal_list, colors=cal_color_list, autopct="%.2f%%", textprops={'fontsize': 12}) 
        plt.savefig(TEMP_PATH + '\\GC_' + nombre_alumno + '.png', bbox_inches = 'tight')
        plt.close(fig1)
    else:
            print(f"No se generó ningún gráfico circular para el alumno, verificar filas: {nombre_alumno}")

# Rutina para crear grafico de barras de calificaciones trimestrales
def CrearGraficosBarras(asignatura_list, nombre_alumno):
    # Listas de calificaciones y Trimestres
    TRIMESTRE_LIST = ['t1', 't2', 't3']
    TRIMESTRE_UPPER_LIST = ['T1', 'T2', 'T3']
    CAL_LIST = ['SUSPENSO', 'APROBADO', 'NOTABLE', 'SOBRESALIENTE']
    CAL_COLOR_LIST = [SUSPENSO_COLOR, APROBADO_COLOR, NOTABLE_COLOR, SOBRESALIENTE_COLOR]
    
    # Iterar sobre trimestres
    trim_nota_media_list = [0] * len(TRIMESTRE_LIST)
    trim_calif_list = []
    trim_color_list =[]
    for trim_idx,trim in enumerate(TRIMESTRE_LIST):
        # Iterar sobre asignaturas
        for asig in asignatura_list:
            trim_nota_media_list[trim_idx] += asig[trim]
        # Calcular nota media trimestral diviendo por número de asignaturas
        trim_nota_media_list[trim_idx] = round(trim_nota_media_list[trim_idx] / len(asignatura_list), 2)
    
        # Obtener calificacion trimestral y color
        calif, color_calif = ObtenerCalificacion(trim_nota_media_list[trim_idx])

        trim_calif_list.append(calif)
        trim_color_list.append('#' + color_calif)
        
    # Crear gráfico de barras
    fig1, ax1 = plt.subplots()
    ax1.bar(TRIMESTRE_UPPER_LIST, trim_nota_media_list, color=trim_color_list)
    # Asignar rango de eje de 0 a 10
    ax1.set_ylim(0, 11)
    plt.yticks(np.arange(0, 11))
    # Añadir texto encima de las barras
    for trim_nm_idx, trim_nm in enumerate(trim_nota_media_list):
        ax1.text(trim_nm_idx -0.09, trim_nm + 0.1, str(trim_nm))
    #Añadir leyenda en la parte inferior de la tabla
    handles = []
    labels = []
    for cal_index in range(len(CAL_LIST)):
        patch = mpatches.Patch(color='#' + CAL_COLOR_LIST[cal_index], label=CAL_LIST[cal_index])
        handles.append(patch)
        labels.append(CAL_LIST[cal_index])
    ax1.legend(handles=handles, labels=labels, loc='upper center', bbox_to_anchor=(0.5, -0.07), fancybox=True, shadow=True, ncol=5)
    # Agregar titulo y nombre de eje vertical
    ax1.set_ylabel("Nota Media") 
    plt.title("Calificaciones trimestrales")
    plt.savefig(TEMP_PATH + '\\GB_' + nombre_alumno + '.png', bbox_inches = 'tight')
    plt.close(fig1)

# Asignar Tags y Crear ficheros Word
def AsignarTagsCrearWord(datos_alumnos_df, excel_df):
    # Eliminar valores duplicados de columna "ASIGNATURA" (Ordenar Alfabeticamente)
    filter_asig_list = sorted(excel_df['ASIGNATURA'].astype(str).drop_duplicates())

    # Añadir tildes a las asignaturas
    filter_asig_td_list = []
    for item in filter_asig_list:
        if item in dict_asig:
            valor_td = dict_asig[item]
            filter_asig_td_list.append(valor_td.upper()) # Pasa los valores a mayusculas
        else: # Si no encuentra asignatura, reemplazara "Sin Asignar"
            filter_asig_td_list.append("Sin asignar")

    # Eliminar valores duplicados de columna "NOMBRE" (Ordenar Alfabeticamente)
    filt_nombre_alumno_list = sorted(datos_alumnos_df['NOMBRE'].astype(str).drop_duplicates())

    # Obtener nombre alumnos para extraer en el fichero
    nombre_alumno = filt_nombre_alumno_list[0]
    
    # Contador para llevar la cuenta de cuántos documentos de alumnos se han generado
    contador_documentos = 0

    # Iterar por alumno
    for nombre_alumno in filt_nombre_alumno_list:
        # Cargar plantilla Word
        docx_tpl = DocxTemplate(PLANTILLA_WORD_PATH)

        # Filtrar datos_alumnos_df y obtener clase para el alumno actual
        filt_datos_alumnos_df = datos_alumnos_df[(datos_alumnos_df['NOMBRE'] == nombre_alumno)]
        clase = filt_datos_alumnos_df.iloc[0]['CLASE']

        # Crear Tabla de notas
        asignatura_list = []
        # Iterar por indices de asignaturas
        for asig_idx in range(len(filter_asig_list)):
            # Filtrar por excel_df por alumno y asignatura
            asign = filter_asig_list[asig_idx]
            filt_al_as_excel_df = excel_df[(excel_df['NOMBRE'] == nombre_alumno) & (excel_df['ASIGNATURA'] == asign)]
            
            # Crear asignatura_dict
            asignatura_dict = {
                'nombre_asignatura': filter_asig_td_list[asig_idx],
                't1': round(filt_al_as_excel_df.iloc[0]['NOTA T1'], 1),
                't2': round(filt_al_as_excel_df.iloc[0]['NOTA T2'], 1),
                't3': round(filt_al_as_excel_df.iloc[0]['NOTA T3'], 1),
            }
            
            # Obtener nota final y calificaciones
            asignatura_dict = ObtenerNotaFinal(asignatura_dict)
            
            # Añadir imagen tabla de barra
            img_bar = TEMP_PATH + '\\GB_' + nombre_alumno + '.png'
            tab_barra = InlineImage(docx_tpl, img_bar, width=Mm(105))
            # Añadir imagen tabla circular
            img_cir = TEMP_PATH + '\\GC_' + nombre_alumno + '.png'
            tab_gcir = InlineImage(docx_tpl, img_cir, width=Mm(60))
            # Añadir logo 
            logo_img = InlineImage(docx_tpl, r"C:\Users\lalvareg\Desktop\Desarrollo-codigo\desarrolloReportes\Inputs\Logo_Montgat.png", width=Mm(5))
       
            # Añadir el diccionario asignatura_dict a asignatura_list
            asignatura_list.append(asignatura_dict)

        # Crear grafico circular de calificaciones
        CrearGraficoCircular(asignatura_list, nombre_alumno)
        
        # Crear grafico de barras trimestral
        CrearGraficosBarras(asignatura_list, nombre_alumno)
                        
        # Crear contexto
        context = {
            'curso': CURSO,
            'nombre_alumno': nombre_alumno,
            'clase': clase,
            'asignatura_list': asignatura_list,
            'logo': logo_img,
            'graf_circular': tab_gcir,
            'graf_barra': tab_barra
        }
        
        # Renderizar plantilla
        docx_tpl.render(context)

        # Crear nombre fichero Word generado
        titulo = "NOTAS_" + nombre_alumno
        titulo = titulo.upper()
        titulo = EliminarTildes(titulo)
        titulo = titulo.replace(" ", "_")
        titulo += ".docx"

        # Exportar fichero generada
        docx_tpl.save(OUTPUT_PATH + '\\' + titulo)
        # Incrementar el contador de documentos
        contador_documentos += 1
        
        print(f"Doumento para", titulo, "Generado correctamente.\nGraficos barra:",img_bar,"Generado correctamente\nGraficos circulares",img_cir, "Generado correctamente...")    

# RUTINA PRINCIPAL
def main():
    # Eliminar y volve a crear una carpeta Outputs y Temporal
    EliminarCrearCarpetas(OUTPUT_PATH)
    EliminarCrearCarpetas(TEMP_PATH)

    #Lectura de fichero "df = dataframe" EXCEL
    excel_df = pd.read_excel(NOTAS_ALUMNOS_EXCEL_PATH, sheet_name='Notas')
    datos_alumnos_df = pd.read_excel(NOTAS_ALUMNOS_EXCEL_PATH, sheet_name='Datos_Alumnos')

    # Detección Errores
    DeteccionErrores(excel_df)

    # Asignar Tags y Crear ficheros Word
    AsignarTagsCrearWord(datos_alumnos_df, excel_df)

    # Eliminar carpeta TEMP
    EliminarCarpetas(TEMP_PATH)
    print("Se ejecuto con éxito... \nAdios, recuerda dormir bien")
if __name__ == ('__main__'):
    main()