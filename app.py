#    streamlit run app_name.py --server.port 5998
#    Nada

from enum import Enum
from io import BytesIO, StringIO
from typing import Union
from datetime import datetime

import pandas as pd
from pandas import read_excel
from pandas import ExcelWriter
from pandas import read_csv
import streamlit as st
import openpyxl

import time
import sys
import base64
import uuid
import os
import pickle
import uuid
import re
import time

from PIL import Image

hora = time.strftime("%y%m%d")

# image = Image.open('imagen.png')

# st.image(image,
#           use_column_width=False)

st.title("Mini Tablero")
st.info('\nTablero de automatización TSA GARANTIAS y REVINVERSION('
                    'https://cohen.com.ar/).\n\n'
                    ) 

st.sidebar.title("Archivo TSA RESCATES")
filename = st.sidebar.file_uploader("Carga tu xlsx de Rescates", type=['xlsx'], key="tsa_file")
st.sidebar.markdown("---")

st.sidebar.title("Archivo TSA SUSCRIPCIONES")
esco = st.sidebar.file_uploader("Carga tu xlsx de Suscripciones", type=['xlsx'], key="esco_file")
st.sidebar.markdown("---")

st.sidebar.title("Conciliación SENEBI BO")
st.sidebar.header("Carga el valor del USD, luego ambos XLSX de BO")
dolar_bo = st.sidebar.text_input("Precio dolar SENEBI BO", 'dolar')
st.sidebar.markdown("---")

st.sidebar.title("Archivo ENTREGA GRTIAS TSA")
reinv = st.sidebar.file_uploader("Carga tu xlsx de Garantias", type=['xlsx'], key="gracias_tsa_reinv_file")
st.sidebar.markdown("---")

st.sidebar.title("Archivo RECEPCION GRTIAS TSA")
reinv2 = st.sidebar.file_uploader("Carga tu xlsx de Garantias", type=['xlsx'], key="gracias_tsa_reinv2_file")
st.sidebar.markdown("---")

st.sidebar.title("Reinversion")
TEST = st.sidebar.file_uploader("Carga tu xlsx de Reinversion", type=['xlsx'], key="test_file")
st.sidebar.markdown("---")

st.sidebar.title("NUEVO TSA")
bo = st.sidebar.file_uploader("Carga tu xlsx de FONDOS COHEN de BO !!!!", type=['xlsx'], key="esco_bo_conci_file")
st.sidebar.markdown("---")

st.sidebar.title("LIQUIDACIÓN TSA !!!!!!!!!!!!!!!!!!!!")
liqui_tsa = st.sidebar.file_uploader("Carga tu xlsx de Transferencias TSA de BO !!!!", type=['xlsx'], key="tsa_liqui_file")
st.sidebar.markdown("---")

def download_button(object_to_download, download_filename, button_text, pickle_it=False):
    """
    Generates a link to download the given object_to_download.
    Params:
    ------
    object_to_download:  The object to be downloaded.
    download_filename (str): filename and extension of file. e.g. mydata.csv,
    some_txt_output.txt download_link_text (str): Text to display for download
    link.
    button_text (str): Text to display on download button (e.g. 'click here to download file')
    pickle_it (bool): If True, pickle file.
    Returns:
    -------
    (str): the anchor tag to download object_to_download
    Examples:
    --------
    download_link(your_df, 'YOUR_DF.csv', 'Click to download data!')
    download_link(your_str, 'YOUR_STRING.txt', 'Click to download text!')
    """

    try:
        # some strings <-> bytes conversions necessary here
        b64 = base64.b64encode(object_to_download.encode()).decode()

    except AttributeError as e:
        b64 = base64.b64encode(object_to_download).decode()

    button_uuid = str(uuid.uuid4()).replace('-', '')
    button_id = re.sub('\d+', '', button_uuid)

    custom_css = f""" 
        <style>
            #{button_id} {{
                background-color: rgb(255, 255, 255);
                color: rgb(38, 39, 48);
                padding: 0.25em 0.38em;
                position: relative;
                text-decoration: none;
                border-radius: 4px;
                border-width: 1px;
                border-style: solid;
                border-color: rgb(230, 234, 241);
                border-image: initial;
            }} 
            #{button_id}:hover {{
                border-color: rgb(246, 51, 102);
                color: rgb(246, 51, 102);
            }}
            #{button_id}:active {{
                box-shadow: none;
                background-color: rgb(246, 51, 102);
                color: white;
                }}
        </style> """
    # print(b64)
    dl_link = custom_css + f'<a download="{download_filename}" id="{button_id}" href="data:file/txt;base64,{b64}">{button_text}</a><br></br>'
    
    return dl_link


def main():

    if filename:
        columnas = ['Comitente','CodigoCaja','Cuotas']
        tablero = pd.read_excel(filename, usecols=columnas)
        comit = tablero['Comitente']
        # st.text(comit)

        st.dataframe(tablero)
        # st.table(tablero)
    
     

        
        lista_suscri= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1 = "00Aftfaot    20"+hora+"1130560000000"
        lista_suscri.append(linea1)      

        incio = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        lista_suscri.append(incio)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie = 0
        cuotas = 0
        comitente = 0

        for valor,comit in enumerate(tablero['Comitente']):
            especie = str(tablero['CodigoCaja'][valor])
            cuotas = str(tablero['Cuotas'][valor])
            comitente = str(comit)  
            
            if especie!="nan" and cuotas!="nan" and comitente!="nan":

                #### ESPECIE ###############################################
                especie = str(int(float(especie)))
                #### COMITENTE #############################################
                comitente = str(int(float(comitente)))
                #### CUOTAS ################################################
                cuotas = str(float(cuotas))
                
                ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
                lista_suscri.append("1'I'E'0046'"+comitente+"'"+especie+"       '"+cuotas+"'0309'000050046'N'00'2111'0000'N"+"\r\n")
               # lista_suscri.append("1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"+"\r\n")       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas = len(lista_suscri)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas)))
        if len(str(num_lineas))==1:
            num_lineas = "0" + str(num_lineas)
        linea_final = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas)+"\r\n"
        lista_suscri.append(linea_final)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        lista_suscri[0] = lista_suscri[0]+str(num_lineas)

        datos=open("modelo.txt","w")
        datos.writelines(lista_suscri)
        datos.close()


        nuevo = "modelo.txt"
        with open(nuevo, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo, f'Archivo TSA {nuevo}')
        st.markdown(download_button_str, unsafe_allow_html=True)

        # os.remove("suscri_tsa1.txt")
    
    if esco:
        columnas = ['Comitente','CodigoCaja','Cuotas']
        tablero = pd.read_excel(esco, usecols=columnas)
        comit = tablero['Comitente']
        # st.text(comit)

        st.dataframe(tablero)
        # st.table(tablero)
    
     

        
        lista_suscri= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1 = "00Aftfaot    20"+hora+"1130560000000"
        lista_suscri.append(linea1)      

        incio = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        lista_suscri.append(incio)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie = 0
        cuotas = 0
        comitente = 0

        for valor,comit in enumerate(tablero['Comitente']):
            especie = str(tablero['CodigoCaja'][valor])
            cuotas = str(tablero['Cuotas'][valor])
            comitente = str(comit)  
            
            if especie!="nan" and cuotas!="nan" and comitente!="nan":

                #### ESPECIE ###############################################
                especie = str(int(float(especie)))
                #### COMITENTE #############################################
                comitente = str(int(float(comitente)))
                #### CUOTAS ################################################
                cuotas = str(float(cuotas))
                
                ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
               # lista_suscri.append("1'I'E'0046'"+comitente+"'"+especie+"       '"+cuotas+"'0309'000050046'N'00'2111'0000'N"+"\r\n")
                lista_suscri.append("1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"+"\r\n")       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas = len(lista_suscri)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas)))
        if len(str(num_lineas))==1:
            num_lineas = "0" + str(num_lineas)
        linea_final = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas)+"\r\n"
        lista_suscri.append(linea_final)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        lista_suscri[0] = lista_suscri[0]+str(num_lineas)

        datos=open("modelo.txt","w")
        datos.writelines(lista_suscri)
        datos.close()


        nuevo = "modelo.txt"
        with open(nuevo, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo, f'Archivo TSA {nuevo}')
        st.markdown(download_button_str, unsafe_allow_html=True)

        # os.remove("suscri_tsa1.txt")

    if reinv:
        columnas = ['Comitente Número','Moneda','Importe']
        tablero = pd.read_excel(reinv, usecols=columnas, engine='openpyxl')
        tablero_xls = pd.read_excel(reinv,engine='openpyxl')
        comit = tablero['Comitente Número']
        # st.text(comit)

        st.dataframe(tablero)
        # st.table(tablero)
    
        ################################ EXCEL PREPARACION #############################
        
        def crearSheet(archivo):
            archivo = archivo
            # print(archivo)

            sheet = {'Fecha Concertacion':[],
                      'Fecha Vencimiento':[],
                      'Cuenta':[],
                      'Concepto':[],
                      'Debe':[],
                      'Haber':[],
                      'Contraparte - Custodia':[],
                      'Contraparte - Depositante':[],
                      'Contraparte - Cuenta':[]}
    
            for num in archivo.index:
                # print(num)
                
                fecha = datetime.now()
                fecha = fecha.strftime("%d/%m/%Y")

                sheet['Fecha Concertacion'].append(fecha)         
                sheet['Fecha Vencimiento'].append(fecha)         
                sheet['Cuenta'].append(archivo['Comitente Número'][num])         
                sheet['Concepto'].append(archivo['Tipo'][num])        
                sheet['Debe'].append('0,00')         
                sheet['Haber'].append(archivo['Importe'][num])
                sheet['Contraparte - Custodia'].append('CAJAVAL')
                sheet['Contraparte - Depositante'].append('0046')
                sheet['Contraparte - Cuenta'].append(archivo['Comitente Número'][num])

            sheet = pd.DataFrame(sheet)
            return sheet            

        moneda_7000 = tablero_xls['Moneda'] == 'Dolar Renta Exterior - 7.000' 
        moneda_10000 = tablero_xls['Moneda'] == 'Dolar Renta Local - 10.000'
        moneda_8000 = tablero_xls['Moneda'] == 'Pesos Renta - 8.000'
        nuevo7000 = tablero_xls[moneda_7000]
        nuevo10000 = tablero_xls[moneda_10000]
        nuevo8000 = tablero_xls[moneda_8000]

        # reinversion_xls = nuevo7000.append(nuevo10000)
        reinversion_xls = pd.concat([nuevo7000, nuevo10000], ignore_index=True)
        # reinversion_xls = reinversion_xls.append(nuevo8000)
        reinversion_xls = pd.concat([reinversion_xls, nuevo8000], ignore_index=True)
      
        reinversion_xls = reinversion_xls.reindex(columns=['Número','Comitente Descripción','Fecha','Moneda','Comitente Número',
            'Importe','Tipo','Banco','Tipo de Cuenta','Sucursal','Cuenta','CBU','Tipo de identificador impositivo','Número de identificador impositivo',
            'Titular','Estado'])

        sheet_7000 = crearSheet(nuevo7000.set_index('Número'))
        sheet_10000 = crearSheet(nuevo10000.set_index('Número'))
        sheet_8000 = crearSheet(nuevo8000.set_index('Número'))

        with ExcelWriter('REINVERSION_FECHA.xlsx') as writer:
            reinversion_xls.to_excel(writer,sheet_name='Sheet1',index=False)
            sheet_7000.to_excel(writer,sheet_name='7000',index=False)  
            sheet_10000.to_excel(writer,sheet_name='10000',index=False)  
            sheet_8000.to_excel(writer,sheet_name='8000',index=False)  
        
        control_file = 'REINVERSION_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True) 


        ################### EXCEL SUBIDA A BO ####################

        ############### ESP 7000 ##################################

        with ExcelWriter('7000_FECHA.xlsx') as writer:
            sheet_7000.to_excel(writer,sheet_name='7000',index=False) 
        
        control_file = '7000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)
        

        ############### ESP 10000 ##################################

        with ExcelWriter('10000_FECHA.xlsx') as writer:
            sheet_10000.to_excel(writer,sheet_name='10000',index=False) 
        
        control_file = '10000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)

        ############### ESP 8000 ##################################

        with ExcelWriter('8000_FECHA.xlsx') as writer:
            sheet_8000.to_excel(writer,sheet_name='8000',index=False) 
        
        control_file = '8000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)   







        ################################ EXCEL PREPARACION #############################
     

        
        lista_reinv= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1 = "00Aftfaot    20"+hora+"1130560000000"
        lista_reinv.append(linea1)      

        incio = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        lista_reinv.append(incio)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie = 0
        cuotas = 0
        comitente = 0

        for valor,comit in enumerate(tablero['Comitente Número']):
            especie = str(tablero['Moneda'][valor])
            cuotas = str(tablero['Importe'][valor])
            comitente = str(comit)  
            
            if especie!="nan" and cuotas!="nan" and comitente!="nan":
                renta = especie
                #### ESPECIE ###############################################
                especie = especie
                #### COMITENTE #############################################
                comitente = str(int(float(comitente)))
                #### CUOTAS ################################################
                cuotas = str(float(cuotas))

                # renta = [["Dolar Renta Local - 10.000","10000"],["Dolar Renta Exterior - 7.000","7000"],["Pesos renta-8000","8000"]]
                # renta = {"Dolar Renta Local - 10.000":"10000","Dolar Renta Exterior - 7.000":"7000","Pesos Renta - 8.000":"8000"}
                  
                  
               # if especie in renta:
               #     especie = renta[especie]

                    ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
                lista_reinv.append("1'I'E'0046'"+comitente+"'"+especie+"       '"+cuotas+"'9046'"+comitente+"'N'00'0000'0000'N"+"\r\n")
       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas = len(lista_reinv)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas)))
        if len(str(num_lineas))==1:
            num_lineas = "0" + str(num_lineas)
        linea_final = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas)+"\r\n"
        lista_reinv.append(linea_final)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        lista_reinv[0] = lista_reinv[0]+str(num_lineas)

        datos=open("modelo_reinv.txt","w")
        datos.writelines(lista_reinv)
        datos.close()


        nuevo = "modelo_reinv.txt"
        with open(nuevo, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo, f'Archivo REINV TSA {nuevo}')
        st.markdown(download_button_str, unsafe_allow_html=True)
    
    if reinv2:
        columnas = ['Comitente Número','Moneda','Importe']
        tablero = pd.read_excel(reinv2, usecols=columnas, engine='openpyxl')
        tablero_xls = pd.read_excel(reinv2,engine='openpyxl')
        comit = tablero['Comitente Número']
        # st.text(comit)

        st.dataframe(tablero)
        # st.table(tablero)
    
        ################################ EXCEL PREPARACION #############################
        
        def crearSheet(archivo):
            archivo = archivo
            # print(archivo)

            sheet = {'Fecha Concertacion':[],
                      'Fecha Vencimiento':[],
                      'Cuenta':[],
                      'Concepto':[],
                      'Debe':[],
                      'Haber':[],
                      'Contraparte - Custodia':[],
                      'Contraparte - Depositante':[],
                      'Contraparte - Cuenta':[]}
    
            for num in archivo.index:
                # print(num)
                
                fecha = datetime.now()
                fecha = fecha.strftime("%d/%m/%Y")

                sheet['Fecha Concertacion'].append(fecha)         
                sheet['Fecha Vencimiento'].append(fecha)         
                sheet['Cuenta'].append(archivo['Comitente Número'][num])         
                sheet['Concepto'].append(archivo['Tipo'][num])        
                sheet['Debe'].append('0,00')         
                sheet['Haber'].append(archivo['Importe'][num])
                sheet['Contraparte - Custodia'].append('CAJAVAL')
                sheet['Contraparte - Depositante'].append('0046')
                sheet['Contraparte - Cuenta'].append(archivo['Comitente Número'][num])

            sheet = pd.DataFrame(sheet)
            return sheet            

        moneda_7000 = tablero_xls['Moneda'] == 'Dolar Renta Exterior - 7.000' 
        moneda_10000 = tablero_xls['Moneda'] == 'Dolar Renta Local - 10.000'
        moneda_8000 = tablero_xls['Moneda'] == 'Pesos Renta - 8.000'
        nuevo7000 = tablero_xls[moneda_7000]
        nuevo10000 = tablero_xls[moneda_10000]
        nuevo8000 = tablero_xls[moneda_8000]

        # reinversion_xls = nuevo7000.append(nuevo10000)
        reinversion_xls = pd.concat([nuevo7000, nuevo10000], ignore_index=True)
        # reinversion_xls = reinversion_xls.append(nuevo8000)
        reinversion_xls = pd.concat([reinversion_xls, nuevo8000], ignore_index=True)
      
        reinversion_xls = reinversion_xls.reindex(columns=['Número','Comitente Descripción','Fecha','Moneda','Comitente Número',
            'Importe','Tipo','Banco','Tipo de Cuenta','Sucursal','Cuenta','CBU','Tipo de identificador impositivo','Número de identificador impositivo',
            'Titular','Estado'])

        sheet_7000 = crearSheet(nuevo7000.set_index('Número'))
        sheet_10000 = crearSheet(nuevo10000.set_index('Número'))
        sheet_8000 = crearSheet(nuevo8000.set_index('Número'))

        with ExcelWriter('REINVERSION_FECHA.xlsx') as writer:
            reinversion_xls.to_excel(writer,sheet_name='Sheet1',index=False)
            sheet_7000.to_excel(writer,sheet_name='7000',index=False)  
            sheet_10000.to_excel(writer,sheet_name='10000',index=False)  
            sheet_8000.to_excel(writer,sheet_name='8000',index=False)  
        
        control_file = 'REINVERSION_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True) 


        ################### EXCEL SUBIDA A BO ####################

        ############### ESP 7000 ##################################

        with ExcelWriter('7000_FECHA.xlsx') as writer:
            sheet_7000.to_excel(writer,sheet_name='7000',index=False) 
        
        control_file = '7000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)
        

        ############### ESP 10000 ##################################

        with ExcelWriter('10000_FECHA.xlsx') as writer:
            sheet_10000.to_excel(writer,sheet_name='10000',index=False) 
        
        control_file = '10000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)

        ############### ESP 8000 ##################################

        with ExcelWriter('8000_FECHA.xlsx') as writer:
            sheet_8000.to_excel(writer,sheet_name='8000',index=False) 
        
        control_file = '8000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)   







        ################################ EXCEL PREPARACION #############################
     

        
        lista_reinv2= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1 = "00Aftfaot    20"+hora+"1130560000000"
        lista_reinv2.append(linea1)      

        incio = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        lista_reinv2.append(incio)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie = 0
        cuotas = 0
        comitente = 0

        for valor,comit in enumerate(tablero['Comitente Número']):
            especie = str(tablero['Moneda'][valor])
            cuotas = str(tablero['Importe'][valor])
            comitente = str(comit)  
            
            if especie!="nan" and cuotas!="nan" and comitente!="nan":
                renta = especie
                #### ESPECIE ###############################################
                especie = especie
                #### COMITENTE #############################################
                comitente = str(int(float(comitente)))
                #### CUOTAS ################################################
                cuotas = str(float(cuotas))

                # renta = [["Dolar Renta Local - 10.000","10000"],["Dolar Renta Exterior - 7.000","7000"],["Pesos renta-8000","8000"]]
                # renta = {"Dolar Renta Local - 10.000":"10000","Dolar Renta Exterior - 7.000":"7000","Pesos Renta - 8.000":"8000"}
                  
                  
               # if especie in renta:
               #     especie = renta[especie]

                    ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
                lista_reinv2.append("1'I'R'0046'"+comitente+"'"+especie+"       '"+cuotas+"'9046'"+comitente+"'N'00'0000'0000'N"+"\r\n")
       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas = len(lista_reinv2)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas)))
        if len(str(num_lineas))==1:
            num_lineas = "0" + str(num_lineas)
        linea_final = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas)+"\r\n"
        lista_reinv2.append(linea_final)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        lista_reinv2[0] = lista_reinv2[0]+str(num_lineas)

        datos=open("modelo_reinv.txt","w")
        datos.writelines(lista_reinv2)
        datos.close()


        nuevo = "modelo_reinv.txt"
        with open(nuevo, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo, f'Archivo reinv2 TSA {nuevo}')
        st.markdown(download_button_str, unsafe_allow_html=True)

    if dolar_bo!='dolar':
        # if control_boletos:
        #     control_bole = control_bole
        # if arancel:
        #     arancel = arancel 
        control_boletos = st.file_uploader("Carga tu xlsx BOLETOS", type=['xlsx'])
        arancel = st.file_uploader("Carga tu xlsx ARANCELES", type=['xlsx'])   
        ################################################################################################################################
        columnas = ["Tipo de Operación","Número de Boleto","Comitente - Número","Fecha de concertación","Instrumento - Símbolo","Cantidad","Moneda","Bruto"]
        

        if control_boletos and arancel:
            aranceles = pd.read_excel(arancel, engine='openpyxl')
            control = pd.read_excel(control_boletos, engine='openpyxl', usecols=columnas)
            control = control.reindex(columns=columnas)
        ################################################################################################################################




            ###### FLITRAMOS POR SOLO OPERACIONES SENEBI ####################

            # senebis = ["Compra SENEBI","Compra SENEBI Colega Pesos","Compra SENEBI CP  Letras","Compra SENEBI CP ON","Compra SENEBI Dólar Cable CP Letras",
            #            "Compra SENEBI Dolar MEP","Venta SENEBI","Venta SENEBI Cable","Venta SENEBI Colega Pesos","Venta Senebi CP Letras","Venta SENEBI Letras Dolar MEP CP",
            #            "Venta Senebi Pesos ON CP"]
            datos = []
            # print(control)
            for e in control.values:
                if "SENEBI" in e[0]:
                    if "Compra" in e[0]:
                        e[7] = 0 - e[7]
                    datos.append(e)
                elif "Senebi" in e[0]:
                    if "Compra" in e[0]:
                        e[7] = 0 - e[7]
                    datos.append(e)    

            datos = pd.DataFrame(datos, columns=columnas)


                

            # print(datos)



            ################  AGREGAMOS LA FILA "INTERES" Y LUEGO SI SON EN DOLARES MULTIPLICAMOS POR EL PRECIO DOLAR ###############3

            datos['interes'] = datos["Bruto"]
            for valor,moneda in enumerate(datos["Moneda"]):
                # print(moneda)
                if moneda!="$":
                    datos['interes'][valor] = float(datos["Bruto"][valor])*float(dolar_bo)

            # for e in datos.values:
            #     if "Compra" in e[0]:
            #         e[8] = 0 - e[8]        



            ##############  AGREGAMOS LOS ARANCELES X MANAGER SENEBI #########################
            solo_aranceles = []
            for e in aranceles.values:
                if "SENEBI" in e[9]:
                    # print(e)
                    solo_aranceles.append(e)
                elif "Senebi" in e[9]:
                    solo_aranceles.append(e)  

            datos_aranceles = pd.DataFrame(solo_aranceles, columns=aranceles.columns)
            # print(solo_aranceles["'SENEBI'"])      






            ##################### REORDENAMOS LAS COLUMNAS ##################################
            # datos = datos[["'Boleto'","'Operacion'","'Comitente'","'Nombre de la Cuenta'","'Especie'","'Imp_Bruto'","interes","'Valor_Nominal'","'Moneda'","'Total_Neto'","'Precio'"]]



            ###########   GUARDAMOS NUEVO EXCEL CON AMBAS SHEETS #######################
            with ExcelWriter('control_senebi_fecha.xlsx') as writer:
                datos.to_excel(writer,sheet_name='CONTROL',index=False)
                datos_aranceles.to_excel(writer,sheet_name='AxM',index=False)  
            control_file = 'control_senebi_fecha.xlsx'
            with open(control_file, 'rb') as f:
                s = f.read()

            download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
            st.markdown(download_button_str, unsafe_allow_html=True)  

    st.sidebar.info('\nCOHEN('
                    'https://www.cohen.com.ar/).\n\n'
                    ) 
    
    if TEST:
        columnas = ['Comitente Número','Moneda','Importe']
        tablero = pd.read_excel(TEST, usecols=columnas, engine='openpyxl')
        tablero_xls = pd.read_excel(TEST,engine='openpyxl')
        comit = tablero['Comitente Número']
        # st.text(comit)

        st.dataframe(tablero)
        # st.table(tablero)
    
        ################################ EXCEL PREPARACION #############################
        
        def crearSheet(archivo):
            archivo = archivo
            # print(archivo)

            sheet = {'Fecha Concertacion':[],
                      'Fecha Vencimiento':[],
                      'Cuenta':[],
                      'Concepto':[],
                      'Debe':[],
                      'Haber':[],
                      'Contraparte - Custodia':[],
                      'Contraparte - Depositante':[],
                      'Contraparte - Cuenta':[]}
    
            for num in archivo.index:
                # print(num)
                
                fecha = datetime.now()
                fecha = fecha.strftime("%d/%m/%Y")

                sheet['Fecha Concertacion'].append(fecha)         
                sheet['Fecha Vencimiento'].append(fecha)         
                sheet['Cuenta'].append(archivo['Comitente Número'][num])         
                sheet['Concepto'].append(archivo['Tipo'][num])        
                sheet['Debe'].append('0,00')         
                sheet['Haber'].append(archivo['Importe'][num])
                sheet['Contraparte - Custodia'].append('CAJAVAL')
                sheet['Contraparte - Depositante'].append('0046')
                sheet['Contraparte - Cuenta'].append(archivo['Comitente Número'][num])

            sheet = pd.DataFrame(sheet)
            return sheet            

        moneda_7000 = tablero_xls['Moneda'] == 'Dolar Renta Exterior - 7.000' 
        moneda_10000 = tablero_xls['Moneda'] == 'Dolar Renta Local - 10.000'
        moneda_8000 = tablero_xls['Moneda'] == 'Pesos Renta - 8.000'
        nuevo7000 = tablero_xls[moneda_7000]
        nuevo10000 = tablero_xls[moneda_10000]
        nuevo8000 = tablero_xls[moneda_8000]

        # reinversion_xls = nuevo7000.append(nuevo10000)
        reinversion_xls = pd.concat([nuevo7000, nuevo10000], ignore_index=True)
        # reinversion_xls = reinversion_xls.append(nuevo8000)
        reinversion_xls = pd.concat([reinversion_xls, nuevo8000], ignore_index=True)

        reinversion_xls = reinversion_xls.reindex(columns=['Número','Comitente Descripción','Fecha','Moneda','Comitente Número',
            'Importe','Tipo','Banco','Tipo de Cuenta','Sucursal','Cuenta','CBU','Tipo de identificador impositivo','Número de identificador impositivo',
            'Titular','Estado'])

        sheet_7000 = crearSheet(nuevo7000.set_index('Número'))
        sheet_10000 = crearSheet(nuevo10000.set_index('Número'))
        sheet_8000 = crearSheet(nuevo8000.set_index('Número'))

        with ExcelWriter('REINVERSION_FECHA.xlsx') as writer:
            reinversion_xls.to_excel(writer,sheet_name='Sheet1',index=False)
            sheet_7000.to_excel(writer,sheet_name='7000',index=False)  
            sheet_10000.to_excel(writer,sheet_name='10000',index=False)  
            sheet_8000.to_excel(writer,sheet_name='8000',index=False)  
        
        control_file = 'REINVERSION_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)
        
        

        ################### EXCEL SUBIDA A BO ####################

        ############### ESP 7000 ##################################

        with ExcelWriter('7000_FECHA.xlsx') as writer:
            sheet_7000.to_excel(writer,sheet_name='7000',index=False) 
        
        control_file = '7000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)
        

        ############### ESP 10000 ##################################

        with ExcelWriter('10000_FECHA.xlsx') as writer:
            sheet_10000.to_excel(writer,sheet_name='10000',index=False) 
        
        control_file = '10000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)

        ############### ESP 8000 ##################################

        with ExcelWriter('8000_FECHA.xlsx') as writer:
            sheet_8000.to_excel(writer,sheet_name='8000',index=False) 
        
        control_file = '8000_FECHA.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)   







        ################################ EXCEL PREPARACION #############################
     

        
        lista_reinv= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1 = "00Aftfaot    20"+hora+"1130560000000"
        lista_reinv.append(linea1)      

        incio = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        lista_reinv.append(incio)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie = 0
        cuotas = 0
        comitente = 0

        for valor,comit in enumerate(tablero['Comitente Número']):
            especie = str(tablero['Moneda'][valor])
            cuotas = str(tablero['Importe'][valor])
            comitente = str(comit)  
            
            if especie!="nan" and cuotas!="nan" and comitente!="nan":

                #### ESPECIE ###############################################
                especie = especie
                #### COMITENTE #############################################
                comitente = str(int(float(comitente)))
                #### CUOTAS ################################################
                cuotas = str(float(cuotas))

                # renta = [["Dolar Renta Local - 10.000","10000"],["Dolar Renta Exterior - 7.000","7000"],["Pesos renta-8000","8000"]]
                renta = {"Dolar Renta Local - 10.000":"10000","Dolar Renta Exterior - 7.000":"7000","Pesos Renta - 8.000":"8000"}

                if especie in renta:
                    especie = renta[especie]

                    ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
                  # lista_reinv.append("1'I'E'0046'1'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"+"\r\n")
                    lista_reinv.append("1'I'E'0046'"+comitente+"'"+especie+"       '"+cuotas+"'0046'1'N'00'0000'0000'N"+"\r\n")       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas = len(lista_reinv)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas)))
        if len(str(num_lineas))==1:
            num_lineas = "0" + str(num_lineas)
        linea_final = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas)+"\r\n"
        lista_reinv.append(linea_final)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        lista_reinv[0] = lista_reinv[0]+str(num_lineas)

        datos=open("modelo_reinv.txt","w")
        datos.writelines(lista_reinv)
        datos.close()


        nuevo = "modelo_reinv.txt"
        with open(nuevo, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo, f'Archivo REINV TSA {nuevo}')
        st.markdown(download_button_str, unsafe_allow_html=True)
    if bo:

        columnas = ['Comitente - Descripción','Instrumento - Símbolo','Instrumento - Denominación','Cuenta - Nro','Saldo Total']
        archivo_bo = pd.read_excel(bo, usecols=columnas, engine='openpyxl')
    
        # Aquí puedes añadir la lógica para cargar los archivos adicionales si es necesario
        # archivo_esco_plus = st.file_uploader("Carga tu xlsx de PLUS de ESCO !!!!!!", type=['xls'])
        # archivo_esco_crf = st.file_uploader("Carga tu xlsx de CRF de ESCO !!!!!!", type=['xls'])
        # archivo_esco_crfDOL = st.file_uploader("Carga tu xlsx de CRF DOLAR de ESCO !!!!!!", type=['xls'])
        # archivo_esco_crfPYMES = st.file_uploader("Carga tu xlsx de CRF PYMES de ESCO !!!!!!", type=['xls'])
    
        # Generar archivo .ict con encabezado especificado
        ict_header = "SourceCashAccount;ReceivingCashAccount;TransactionReference;PaymentSystem;Currency;Amount;SettlementDate;Description;CorporateActionReference;TransactionOnHoldCSD;TransactionOnHoldParticipant\n"
        ict_content = []
    
        # Aquí, itera sobre tu archivo_bo para llenar el contenido de ict_content
        for index, row in archivo_bo.iterrows():
            # Ajusta estas variables de acuerdo con tus necesidades
            SourceCashAccount = row['Cuenta - Nro']
            ReceivingCashAccount = 'SomeReceivingAccount'  # Ajusta según tu lógica
            TransactionReference = 'Ref-' + str(index)
            PaymentSystem = 'SomePaymentSystem'  # Ajusta según tu lógica
            Currency = 'USD'  # Ajusta según tu lógica
            Amount = row['Saldo Total']
            SettlementDate = datetime.now().strftime('%Y-%m-%d')
            Description = 'SomeDescription'  # Ajusta según tu lógica
            CorporateActionReference = 'CorpAct-' + str(index)
            TransactionOnHoldCSD = 'No'
            TransactionOnHoldParticipant = 'No'
    
            ict_line = f"{SourceCashAccount};{ReceivingCashAccount};{TransactionReference};{PaymentSystem};{Currency};{Amount};{SettlementDate};{Description};{CorporateActionReference};{TransactionOnHoldCSD};{TransactionOnHoldParticipant}\n"
            ict_content.append(ict_line)
    
        ict_file_content = ict_header + "".join(ict_content)
    
        # Guardar el archivo .ict
        ict_filename = 'output_file.ict'
        with open(ict_filename, 'w') as f:
            f.write(ict_file_content)
    
        # Proporcionar el enlace de descarga
        def download_button(object_to_download, download_filename, button_text):
            b64 = base64.b64encode(object_to_download.encode()).decode()
            button_id = str(uuid.uuid4()).replace('-', '')
            custom_css = f""" 
                <style>
                    #{button_id} {{
                        background-color: rgb(255, 255, 255);
                        color: rgb(38, 39, 48);
                        padding: 0.25em 0.38em;
                        position: relative;
                        text-decoration: none;
                        border-radius: 4px;
                        border-width: 1px;
                        border-style: solid;
                        border-color: rgb(230, 234, 241);
                        border-image: initial;
                    }} 
                    #{button_id}:hover {{
                        border-color: rgb(246, 51, 102);
                        color: rgb(246, 51, 102);
                    }}
                    #{button_id}:active {{
                        box-shadow: none;
                        background-color: rgb(246, 51, 102);
                        color: white;
                        }}
                </style> """
            dl_link = custom_css + f'<a download="{download_filename}" id="{button_id}" href="data:file/txt;base64,{b64}">{button_text}</a><br></br>'
            return dl_link
    
        download_button_str = download_button(ict_file_content, ict_filename, 'Descargar archivo ICT')
        st.markdown(download_button_str, unsafe_allow_html=True)
    
        st.sidebar.info('\nCOHEN (https://www.cohen.com.ar/).\n\n')

    if liqui_tsa:
        archivo = pd.read_excel(liqui_tsa, engine='openpyxl')

        nuevo_xls = []
        solo_inmediato = []
       
        for linea in archivo.values:
            
            comitente = linea[0]
            codigo = linea[1]
            tipo = linea[3]
            cantidad = linea[4]
            tratamiento = linea[5]


            if tipo == 'Venta':
                for op in archivo.values:
                    if op[0]==comitente and op[1]==codigo and op[3]=='Compra' and op[4]>=cantidad:
                        linea[5] = 'Diferido'
                    elif op[0]==comitente and op[1]==codigo and op[3]=='Compra' and op[4]<cantidad:
                        diferencia = linea[4] - op[4]
                        solo_inmediato.append([comitente,codigo,'NADA',tipo,diferencia,tratamiento])
                        linea[4] = op[4]
                        linea[5] = 'Diferido'

            nuevo_xls.append(linea)  
 
        columnas = ['Comitente - Número','Instrumento - Código caja','Instrumento - Símbolo','Transferencia - Tipo','Transferencia - Cantidad Total','Transferencia - Tratamiento'] 
        nuevo_xls = pd.DataFrame(nuevo_xls, columns=columnas)              
        solo_inmediato = pd.DataFrame(solo_inmediato, columns=columnas)              
        

        st.dataframe(nuevo_xls)
        # print(solo_inmediato)


        ################################ EXCEL PREPARACION #############################
     

        
        lista_tsa= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1 = "00Aftfaot    20"+hora+"1130560000000"
        lista_tsa.append(linea1)      

        incio = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        lista_tsa.append(incio)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie = 0
        cuotas = 0
        comitente = 0

        for valor,comit in enumerate(nuevo_xls['Comitente - Número']):
            especie = str(nuevo_xls['Instrumento - Código caja'][valor])
            cuotas = str(nuevo_xls['Transferencia - Cantidad Total'][valor])
            tipo = str(nuevo_xls['Transferencia - Tratamiento'][valor])
            lado = str(nuevo_xls['Transferencia - Tipo'][valor])
            comitente = str(comit)  
            
            if tipo=='Diferido' and lado=='Venta':

                ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
                lista_tsa.append("1'D'E'0046'"+comitente+"'"+especie+"       '"+cuotas+"'7046'10000'N'00'0000'0000'N"+"\r\n")
            elif tipo=='Inmediato' and lado=='Venta':

                ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
                lista_tsa.append("1'I'E'0046'"+comitente+"'"+especie+"       '"+cuotas+"'7046'10000'N'00'0000'0000'N"+"\r\n")    
       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas = len(lista_tsa)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas)))
        if len(str(num_lineas))==1:
            num_lineas = "0" + str(num_lineas)
        linea_final = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas)+"\r\n"
        lista_tsa.append(linea_final)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        lista_tsa[0] = lista_tsa[0]+str(num_lineas)

        datos=open("modelo_cris_tsa.txt","w")
        datos.writelines(lista_tsa)
        datos.close()


        nuevo = "modelo_cris_tsa.txt"
        with open(nuevo, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo, f'Archivo CRIS TSA {nuevo}')
        st.markdown(download_button_str, unsafe_allow_html=True)


        ################################ TSA EXTRA PREPARACION #############################
     

        
        tsa_extra= []

        # -----------------PRIMERAS DOS LINEAS OBLIGATORIAS DEL TXT------------------------------------------
        linea1_extra = "00Aftfaot    20"+hora+"1130560000000"
        tsa_extra.append(linea1_extra)      

        incio_extra = "\r\n"+"0"+hora+"FTFAOT0046"+"\r\n"
        tsa_extra.append(incio_extra)

        # -----------------AGREGAMOS LINEAS SEGUN LA CANTIDAD DE SUCRI QUE TENGAMOS-----------------------------------------

        # especie = 5 digitos 
        # cuotas = 00000000000.0000000  ( 11 y 7) 
        # comitente = 9 digitos 
        especie_extra = 0
        cuotas_extra = 0
        comitente_extra = 0

        for valor,comit in enumerate(solo_inmediato['Comitente - Número']):
            especie_extra = str(solo_inmediato['Instrumento - Código caja'][valor])
            cuotas_extra = str(solo_inmediato['Transferencia - Cantidad Total'][valor])
            # tipo = str(solo_inmediato['Transferencia - Tratamiento'][valor])
            # lado = str(solo_inmediato['Transferencia - Tipo'][valor])
            comitente_extra = str(comit)  
            
            # if tipo=='Diferido' and lado=='Venta':

            #     ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
            #     tsa_extra.append("1'D'E'0046'"+comitente_extra+"'"+especie_extra+"       '"+cuotas_extra+"'7046'1000'N'00'0000'0000'N"+"\r\n")
            # elif tipo=='Inmediato' and lado=='Venta':

                ################ AGREGO EL FORMATO A NUESTRO ARCHIVO
            tsa_extra.append("1'I'E'0046'"+comitente_extra+"'"+especie_extra+"       '"+cuotas_extra+"'7046'10000'N'00'0000'0000'N"+"\r\n")    
       

        # LINEA EJEMPLO
        #"1'I'E'0046'000000003'"+especie+"       '"+cuotas+"'0046'"+comitente+"'N'00'0000'0000'N"

        # ------------------------AGREGAMOS LINEA FINAL---------------------------------------

        # LINEA FINAL
        num_lineas_extra = len(tsa_extra)-1 # restamos la primera que no cuenta
        # print(len(str(num_lineas_extra)))
        if len(str(num_lineas_extra))==1:
            num_lineas_extra = "0" + str(num_lineas_extra)
        linea_final_extra = "99Aftfaot    20"+hora+"1130560000000"+str(num_lineas_extra)+"\r\n"
        tsa_extra.append(linea_final_extra)

        # AGREAGR NUMERO DE FILAS A LA PRIMER LINEA
        tsa_extra[0] = tsa_extra[0]+str(num_lineas_extra)

        datos_extra=open("modelo_extra_tsa.txt","w")
        datos_extra.writelines(tsa_extra)
        datos_extra.close()


        nuevo_extra = "modelo_extra_tsa.txt"
        with open(nuevo_extra, 'rb') as f:
            s = f.read()
            print(s)

        download_button_str = download_button(s, nuevo_extra, f'Archivo EXTRA TSA {nuevo_extra}')
        st.markdown(download_button_str, unsafe_allow_html=True)





        with ExcelWriter('TSA_OPS.xlsx') as writer:
                nuevo_xls.to_excel(writer,sheet_name='TSA',index=False)  
            
        control_file = 'TSA_OPS.xlsx'
        with open(control_file, 'rb') as f:
            s = f.read()

        download_button_str = download_button(s, control_file, f'EXCEL LISTO {control_file}')
        st.markdown(download_button_str, unsafe_allow_html=True)              
                        

if __name__ == '__main__':
    main()      
