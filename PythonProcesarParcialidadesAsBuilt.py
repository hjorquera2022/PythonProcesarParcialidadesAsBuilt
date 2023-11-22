#PythonProcesarParcialidadesAsBuilt.py

import pandas as pd
import os
import subprocess

#*****
#***** ESTRUCTURA COMPLETA 
#*****


#│   ├───DOCUMENTOS ASBUILT 
#       ├───01 PDF
#       │   ├───APROBADOS
#       │   └───OBSERVADOS
#       └───02 EDITABLE
#           ├───APROBADOS
#           └───OBSERVADOS


# Ruta base donde se deben verificar los subdirectorios
ruta_base = 'R:\\01 PARCIALIDADES\\'  ### REAL

# Nombre del archivo de log
archivo_log = ruta_base + '0000-00 ADMINISTRACION\\LOG\\log_ProcesarParcialidadesAsBuilt.txt'

# Nombre del archivo Bat General
archivo_bat = ruta_base + '0000-00 ADMINISTRACION\\BAT\\Bat_ProcesarParcialidadesAsBuilt.bat'

# Planilla con la lista de parcialidades
archivo_excel = ruta_base + 'Listado de Parcialidades_AsBuilt.xlsx'

# Carga el archivo Excel en un DataFrame Hoja de Parcialidades.
df = pd.read_excel(archivo_excel, sheet_name='PARCIALIDADES')

# Filtra el DataFrame para considerar solo parcialidades a 'PROCESAR' igual a 'S'
df_parcialidades = df[df['PROCESAR'] == 'S']

# Abre el archivo de log en modo de escritura
with open(archivo_log, 'w') as log_file:

# Abre el archivo de log en modo de escritura
 with open(archivo_bat, 'w') as bat_file:

    # Itera a través de cada parcialidad y la procesa
    for parcialidad in df_parcialidades['PARCIALIDAD']:
        log_file.write(f'Parcialidad: {parcialidad}\n')

        #******* Abrir Planilla CONTROL DOCUMENTOS ING DEF con las 8 hojas para traspasar a BAT
        #******* Cargar cada una de las 8 hojas del archivo Excel en un DataFrame (DFxxxx)
        #******* Generar cada uno de los BAT 

        #******* Armar nombre de archivo del BAT por cada uno de las 3 hojas

        #ASBUILT

        #ACTUALIZA DOC VIG
        #ACTUALIZA REV LETRA PARCI APRO
        #ACTUALIZA REV NUM PARCI
        
        archivo_ACTUALIZA_EDITABLE_DOC_VIG = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ASBUILT_ACTUALIZA DOC VIG.bat'
        archivo_ACTUALIZA_EDITABLE_REV_LETRA = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ASBUILT_ACTUALIZA REV LETRA PARCI APRO.bat'        
        archivo_ACTUALIZA_EDITABLE_REV_NUM = ruta_base + '0000-00 ADMINISTRACION\\BAT\\' + parcialidad + '_BAT_ASBUILT_ACTUALIZA REV NUM PARCI.bat'
      
        #******* 
        #******* RECORRER TODAS LAS PARCIALIDADES CONTANDO LOS ARCHIVOS DE LA SIGUIENTE ESTRUCTURA de las Carpetas en REVISORES  por cada parcialidad.
        #******* 

        parcialidad_0_7_10 = parcialidad[0:7]
        if parcialidad_0_7_10   == '0029-14':
            parcialidad_0_7_10 = parcialidad[0:10]
        elif parcialidad_0_7_10 == '032ESO-':
            parcialidad_0_7_10 = parcialidad[0:9]
        elif parcialidad_0_7_10 == '032ESP-':
            parcialidad_0_7_10 = parcialidad[0:9]


                                                          #CONTROL DOCUMENTOS AS-BUILT P0001-01.xlsx
        archivo_parcialidad = ruta_base + parcialidad + '\\CONTROL DOCUMENTOS AS-BUILT P' + parcialidad_0_7_10 + '.xlsx'
        
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_EDITABLE_DOC_VIG}\" > \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}_ASBUILT.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_EDITABLE_REV_LETRA}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}_ASBUILT.log\" \n')
        bat_file.write(f'CALL \"{archivo_ACTUALIZA_EDITABLE_REV_NUM}\" >> \"{ruta_base}0000-00 ADMINISTRACION\LOG\P{parcialidad}_ASBUILT.log\" \n')
 
        if not os.path.exists(archivo_parcialidad):
              log_file.write(f'Parcialidad: {parcialidad} SIN ARCHIVO DE INGENIERIA _ASBUILT {archivo_parcialidad}\n')
        else:
                print(f'Procesando Parcialidad _ASBUILT: {parcialidad} ARCHIVO:  {archivo_parcialidad} \n')

                #ASBUILT

                #ACTUALIZA DOC VIG
                #ACTUALIZA REV LETRA PARCI APRO
                #ACTUALIZA REV NUM PARCI

                # Lee el archivo Excel para obtener los nombres de las hojas
                xl = pd.ExcelFile(archivo_parcialidad)
                nombres_hojas = xl.sheet_names

                #******* Generar cada uno de los BAT 
                
                if not 'ACTUALIZA DOC VIG' in nombres_hojas:
                        print(f'La hoja ASBUILT ACTUALIZA DOC VIG no existe en {archivo_excel}.')  
                        log_file.write(f'La hoja ASBUILT ACTUALIZA DOC VIG no existe en {archivo_excel}\n')
                else: 
                        #******* Cargar hoja del archivo Excel en un DataFrame (DF_xxxx)
                        df_ACTUALIZA_DOC_VIG = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA DOC VIG')

                        # Abre el archivo BAT en modo de escritura
                        log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_EDITABLE_DOC_VIG}\n')
                        print(f'BAT {archivo_ACTUALIZA_EDITABLE_DOC_VIG}')
                        with open(archivo_ACTUALIZA_EDITABLE_DOC_VIG, 'w') as bat_file_ACTUALIZA_EDITABLE_DOC_VIG:
                        # Itera a través de la hoja por cada linea
                            for linea in df_ACTUALIZA_DOC_VIG['RUTA']:
                                bat_file_ACTUALIZA_EDITABLE_DOC_VIG.write(f'{linea}\n')

                if not 'ACTUALIZA REV LETRA PARCI APRO' in nombres_hojas:
                        print(f'La hoja ASBUILT ACTUALIZA REV LETRA PARCI APRO no existe en {archivo_excel}.')  
                        log_file.write(f'La hoja ASBUILT ACTUALIZA REV LETRA PARCI APRO no existe en {archivo_excel}\n')
                else:
                        #******* Cargar hoja del archivo Excel en un DataFrame (DF_xxxx)
                        df_ACTUALIZA_REV_LETRA_PARCI_APRO = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA REV LETRA PARCI APRO')
 
                        # Abre el archivo BAT en modo de escritura
                        log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_EDITABLE_REV_LETRA}\n')
                        print(f'BAT {archivo_ACTUALIZA_EDITABLE_REV_LETRA}')
                        with open(archivo_ACTUALIZA_EDITABLE_REV_LETRA, 'w') as bat_file_ACTUALIZA_EDITABLE_REV_LETRA:
                        # Itera a través de la hoja por cada linea
                            for linea in df_ACTUALIZA_REV_LETRA_PARCI_APRO['RUTA']:
                                bat_file_ACTUALIZA_EDITABLE_REV_LETRA.write(f'{linea}\n')

                if not 'ACTUALIZA REV NUM PARCI' in nombres_hojas:
                        print(f'La hoja ASBUILT ACTUALIZA REV NUM PARCI no existe en {archivo_excel}.')
                        log_file.write(f'La hoja ASBUILT ACTUALIZA REV NUM PARCI no existe en {archivo_excel}\n')
                else:
                        #******* Cargar hoja del archivo Excel en un DataFrame (DF_xxxx)
                        df_ACTUALIZA_REV_NUM_PARC = pd.read_excel(archivo_parcialidad, sheet_name='ACTUALIZA REV NUM PARCI')

                        # Abre el archivo BAT en modo de escritura
                        log_file.write(f'Parcialidad: {parcialidad} BAT {archivo_ACTUALIZA_EDITABLE_REV_NUM}\n')
                        print(f'BAT {archivo_ACTUALIZA_EDITABLE_REV_NUM}')
                        with open(archivo_ACTUALIZA_EDITABLE_REV_NUM, 'w') as bat_file_ACTUALIZA_EDITABLE_REV_NUM:
                        # Itera a través de la hoja por cada linea
                            for linea in  df_ACTUALIZA_REV_NUM_PARC['RUTA']:
                                bat_file_ACTUALIZA_EDITABLE_REV_NUM.write(f'{linea}\n')
                        

print("Proceso ASBUILT finalizado. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\BAT en el archivo de log_ProcesarParcialidadesAsBuilt.")
#log_file.write(f'Proceso ASBUILT finalizado. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\LOG en el archivo de log_ProcesarParcialidadesAsBuilt.\n')
bat_file.close

#try:
#    # Ejecuta el archivo batch
#    subprocess.run(archivo_bat, shell=True)
#except Exception as e:
#    print(f"Error al ejecutar el archivo batch AsBuilt: {e}")
#    log_file.write(f'Error en ejecucion de Bat_ProcesarParcialidadesASsBuilt.bat\n')

log_file.close

