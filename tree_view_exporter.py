import xlsxwriter as xls
import tkinter as tk
from tkinter import *
from tkinter import ttk

import os
from os import *
from datetime import date
import datetime



def excel_exportar(tree,columnas,headers,title):
    dia=" Archivo Generado"
    f= filedialog.asksaveasfilename(initialfile =dia, defaultextension=".txt",filetypes=[("All Files","*.*"),("Text Documents","*.txt")])
    fd = os.open( f, os.O_RDWR|os.O_CREAT )
    os.close(fd)
  
   
    ruta=(os.path.abspath(f))
    n=str(ruta.replace('.txt','.xlsx'))

        
    workbook= xls.Workbook(str(n))
    options = {
               'format_columns': True,'format_rows':True,
               'insert_columns': False,'insert_rows': False,'insert_hyperlinks': False,'delete_columns': False,
               'delete_rows':True,'select_locked_cells':False,'sort': False,'autofilter': True
               }


    workbook.set_properties({ 'title': 'PENSA', 'subject': 'INFORMES DE TRAZABILIDAD', 'author': 'Logosi', 'manager': 'Francisco Vasquez', 'category': 'Estadistica','created': date.today(), 'comments': 'Created with Python and XlsxWriter'})


    unlocked=workbook.add_format({'locked':False})
    locked=workbook.add_format({'locked':True})
    header_format=workbook.add_format({'bold': True,   'text_wrap': True, 'valign': 'top', 'fg_color': '#D7E4BC', 'border': 1})
    locked.set_border()
    


        
    worksheet = workbook.add_worksheet('Main')
    worksheet.set_page_view()
    worksheet.set_header(title)
    worksheet.hide_gridlines()
    worksheet.protect("my_password",options)



    today=date.today()


    worksheet.write(0,0,'Fecha de Informe',locked)
    worksheet.write(0,1,str(today),locked)
    

    col_header=0
    for name in headers:
        worksheet.write(2,col_header,name,header_format)
        col_header+=1
        

    row=3
    column=0
    list_cant=columnas

    for child in tree.get_children():
        x=tree.item(child)["values"]
        for r in range(list_cant):
            worksheet.write(row,column,str(tree.item(child)["values"][r]),locked)
            column+=1
            if column>=list_cant:
                column=0
        row+=1

    worksheet.autofilter(2,0,row,columnas-1)
    

    
    workbook.close()

