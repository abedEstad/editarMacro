import streamlit as st
import io
from io import BytesIO
import os.path 
import pathlib
import openpyxl
from openpyxl import load_workbook
buffer = io.BytesIO()
st.set_page_config(page_title='***TSC - APLICACIONES WEB***',page_icon='🤡',layout='wide')
st.title(':sunglasses: :sun_with_face: :face_with_cowboy_hat: :green[Creación de Bultos por Machine Learning] :sunglasses: :sun_with_face: :face_with_cowboy_hat:')
st.write('_Esta es una version de app que permite subir un archivo excel, editarlo, guardarlo y exportarlo a tu directorio. EXCEL XLSX!_ :sunglasses:')

st.write('Iniciando la prueba ...')
archivo_subida_excel= st.file_uploader('Subir Planilla',type='xlsm',accept_multiple_files=False, label_visibility="visible",help=None)
if archivo_subida_excel is not None:
  data = archivo_subida_excel.getvalue().decode('utf-8',errors='ignore')
  parent_path = pathlib.Path(__file__).parent.parent.resolve()
  save_path = os.path.join(parent_path, "data")
  nombre =os.path.join('',archivo_subida_excel.name)
  nombreFinal = nombre.split('.')[0]
  st.write('Nombre:',nombre)

  st.write('Iniciando parte 2')
  worksheetss = openpyxl.load_workbook(archivo_subida_excel,read_only = False, keep_vba = True)
  sheetss = worksheetss['PLANILLA']
  sheetss['AP2'] = "Languages"
  sheetss['AP3'] = "Python"
  sheetss['AP4'] = "Java"
  sheetss['AP5'] = "C"
  sheetss['AP6'] = "C#"
  sheetss['AP7'] = "Swift"
  sheetss['AP8'] = "No. of articles"
  sheetss['AP9'] = 24
  sheetss['AP10'] = 45
  sheetss['AP11'] = 66
  sheetss['AP12'] = 43
  sheetss['AP13'] = 12
  worksheetss.save(nombre)

  st.write('La creación de Bultos ha sido realizada con éxito. \n Por favor, revise su archivo.\n Gracias')



