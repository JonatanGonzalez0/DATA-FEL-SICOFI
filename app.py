from Item import Item
from operacion import Operacion
from datetime import datetime
import tkinter as tk
from tkinter import *
from tkinter import filedialog,messagebox,ttk
import xml.etree.ElementTree as ET
import webbrowser
import pandas as pd
import re,os
from os.path import expanduser
import xlrd
import openpyxl

#for style update pip install --upgrade xlsxwriter

home = expanduser("~")
Documents_location = os.path.join(home,'Documents')

path_DataOutput = Documents_location + "\\FEL-A-SICOFI"
path_DataXML = Documents_location + "\\FEL-A-SICOFI\\XML"
existe_path = os.path.exists(path_DataOutput)
existe_path_XML = os.path.exists(path_DataXML)

if not existe_path_XML:
    os.makedirs(path_DataXML)


if not existe_path:
    os.makedirs(path_DataOutput)
    print('No existia el directorio se ha creado')

if not existe_path:
    os.makedirs(path_DataOutput)
    print('No existia el directorio se ha creado')

nombreEmisor = ""

app = tk.Tk()
type_service_bien = ttk.Combobox(app,font=("Courier", 20))

def carga_Menu():
    Carpeta_Raiz = os.path.dirname(os.path.abspath(__file__))
    app.configure(bg='#003153')
    ancho_ventana = 600
    alto_ventana = 600
    #Centrar app dependiendo del monitor
    x_ventana = app.winfo_screenwidth() // 2 - ancho_ventana // 2
    y_ventana = app.winfo_screenheight() // 2 - alto_ventana // 2
    app.title("Data Analizer FEL-Sicofi Made by Jonatan")
    posicion = str(ancho_ventana) + "x" + str(alto_ventana) + "+" + str(x_ventana) + "+" + str(y_ventana)
    app.geometry(posicion)

    menubar = tk.Menu(app)
    filemenu = tk.Menu(menubar,tearoff=0)

    etiquetaM = tk.Label(app,text="BIENVENIDO",font=("Courier", 45))
    etiquetaM.pack(padx=30,pady=5)
    etiquetaM.configure(bg='#003153',fg='#DADEE7')
    filemenu.add_command(label="Salir", command=salir)

    menubar.add_cascade(label="Menu", menu=filemenu)
    app.config(menu=menubar)

    cargarBotones()
    app.mainloop()

def cargarBotones(): 
    global type_service_bien
    type_service_bien['values']=('Servicios afectos','Bienes afectos')
    type_service_bien.current(0)
    type_service_bien.pack(pady=40)
    combostyle = ttk.Style()
    type_service_bien['state'] = 'readonly'
    combostyle.theme_create('combostyle', parent='alt',
                         settings = {'TCombobox':
                                     {'configure':
                                      {'selectbackground': '#66b2b2',
                                       'fieldbackground': '#66b2b2',
                                       'background': '#66b2b2'
                                       }}}
                         )
    combostyle.theme_use('combostyle') 
    
    btn_cargarArchivo = tk.Button(app,text="Cargar Archivos",font=("Courier", 20),command=cargarArchivo)
    btn_cargarArchivo.configure(background='#66b2b2',fg='#fdfbfb')
    btn_cargarArchivo.pack( pady=50)

    btn_buscarCarpeta = tk.Button(app,text="Ver Carpeta Facturas",font=("Courier", 20),command=buscarCarpeta)
    btn_buscarCarpeta.configure(background='#A8CD1C',fg='#fdfbfb')
    btn_buscarCarpeta.pack( pady=20)

    btn_ExtraerInfo = tk.Button(app,text="Extraer Informacion XML",font=("Courier", 20),command=extraerInfo)
    btn_ExtraerInfo.configure(background='#ff8c00',fg='#fdfbfb')
    btn_ExtraerInfo.pack( pady=10)
    
def buscarCarpeta():
    os.startfile(path_DataOutput)
    
def salir():
    app.destroy()

def cargarArchivo():
    global type_service_bien
    try:
        Carpeta_Raiz = os.path.dirname(os.path.abspath(__file__))
        file_names = filedialog.askopenfilenames(initialdir=Carpeta_Raiz,title = "Selecciona uno o varios archivos")
        list_file_names = list(file_names)
        
        for file_name in list_file_names:
            
            #si archivo es xls convert to xlsx
            
            file_name = prepareFileExcel(file_name)
        
            df = pd.read_excel(file_name)
            serie = 'FEL'
            numero=0
            length_df = len(df.index)

            file_name_export = os.path.basename(file_name).split(".")[0]
            file = path_DataOutput + "\\"+ file_name_export + '.csv'

            with open(file,'w') as f:
                headers="serie,numero,ano,mes,dia,nit,nombre,servaf,servnaf,bienaf,biennaf,otros,inguat,status,tipodoc,concepto,descrip,nretenisr,vretenisr,tduafauca,nduafauca,tconstan,nconstan,vconstan,centro,seriefel,numerofel\n"
                f.write(headers)

                for index_df in range(length_df):
                    fechaDTE = df.loc[index_df]['Fecha de emisión']
                    matchFecha = re.search(r'\d{4}-\d{2}-\d{2}',fechaDTE).group()
                    strFecha = matchFecha.split("-")

                    ano = strFecha[0]
                    mes = strFecha[1]
                    dia = strFecha[2]

                    nit = str(df.loc[index_df]['ID del receptor'])
                    if nit =="CF":
                        nit ="C/F"
                    else:
                        size_nit = len(nit)
                        first_sub_string = nit[0:size_nit-1]
                        end_sub_string = nit[-1]
                        nit = first_sub_string + "-" + end_sub_string
                    
                    nombre = df.loc[index_df]['Nombre completo del receptor']
                    nombre = nombre.replace(",",'')

                    if type_service_bien.get()=="Servicios afectos":
                        servaf = round(df.loc[index_df]['Monto (Gran Total)'],2)
                        servnaf = 0 

                        bienaf = 0
                        biennaf = 0
                    elif type_service_bien.get()=="Bienes afectos":
                        servaf = 0
                        servnaf = 0

                        bienaf = round(df.loc[index_df]['Monto (Gran Total)'],2)
                        biennaf = 0
                    
                    otros = 0
                    inguat = 0 
                    try:
                        estado = df.loc[index_df]['Marca de anulado ']
                    except:
                            estado = df.loc[index_df]['Marca de anulado']

                    if estado =="No":
                        status = 0
                    elif estado =='Si':
                        status = 1
                    
                    tipodoc = 2
                    concepto = 1

                    descrip = "Ventas varias"

                    nretenisr = ''
                    vretenisr = ''
                    tduafauca = ''
                    nduafauca = ''
                    tconstan = ''
                    nconstan = ''
                    vconstan = '' 
                    centro =  '' 

                    seriefel = df.loc[index_df]['Serie']

                    numerofel = df.loc[index_df]['Número del DTE']

                    strFila = '{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{},{}\n'.format(serie,numero,ano,mes,dia,nit,nombre,servaf,servnaf,bienaf,biennaf,otros,inguat,status,tipodoc,concepto,descrip,nretenisr,vretenisr,tduafauca,nduafauca,tconstan,nconstan,vconstan,centro,seriefel,numerofel)
                    f.write(strFila)
                
                f.close()
        messagebox.showinfo("Correcto","Se analizaron las facturas, se encuentraen en carpeta Documentos/FEL-A-SICOFI")
        webbrowser.open(path_DataOutput)    
    except:
        messagebox.showwarning("Error","No se abrio ningun archivo") 

def extraerInfo():
    try:
        path_downloads = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')     
        file_names = filedialog.askopenfilenames(initialdir=path_downloads,title = "Selecciona uno o varios archivos")
        list_file_names = list(file_names)
        operaciones=[]
        nulls=[]
        ns = {'DTE':'http://www.sat.gob.gt/dte/fel/0.2.0'}
        for file_name in list_file_names:
            #si archivo solo contiene la palabra "null" no se procesa
            if os.path.getsize(file_name) == 4:
                #extract file name without extension
                file_name_export = os.path.basename(file_name).split(".")[0]
                nulls.append(file_name_export)
                continue
            try:
                tree = ET.parse(file_name)
                ET.indent(tree, space="\t", level=0)
                tree.write(file_name, encoding="utf-8")     
                tree = ET.parse(file_name)
                root = tree.getroot()  
                
                try:
                    DATOS_Emision = root.find('DTE:SAT',ns).find('DTE:DTE',ns).find('DTE:DatosEmision',ns)
                    DATOS_Certificacion = root.find('DTE:SAT',ns).find('DTE:DTE',ns).find('DTE:Certificacion',ns)
                except:
                    continue

                nombreEmisor = DATOS_Emision.find('DTE:Emisor',ns).get('NombreEmisor')
                nombreEmisor = nombreEmisor.replace(",","")
                nombreEmisor = nombreEmisor.replace(" ","_")
                nombreEmisor = nombreEmisor.replace(".","_")
                nombreEmisor = nombreEmisor.replace("-","_")
                nombreEmisor = nombreEmisor.replace("?","")
                nombreEmisor = nombreEmisor.replace("¿","")

                Datosgen = DATOS_Emision.find('DTE:DatosGenerales',ns)
                fecha = Datosgen.get('FechaHoraEmision')
                fecha = fecha.split("-")
                anio = fecha[0]
                mes = fecha[1]
                dia = fecha[2].split('T')[0]
                fecha = ('{}/{}/{}'.format(dia,mes,anio))
                fecha_obj = datetime.strptime(fecha, '%d/%m/%Y').date()

                #certificacion
                nitReceptor = DATOS_Emision.find('DTE:Receptor',ns).get('IDReceptor')
                NombreReceptor = DATOS_Emision.find('DTE:Receptor',ns).get('NombreReceptor')
                NombreReceptor = NombreReceptor.replace(",","")
                noDTE  = DATOS_Certificacion.find('DTE:NumeroAutorizacion',ns).get('Numero')
                serie = DATOS_Certificacion.find('DTE:NumeroAutorizacion',ns).get('Serie')
                
                #items
                items = DATOS_Emision.find('DTE:Items',ns)
                for item in items:
                    cant = item.find('DTE:Cantidad',ns).text
                    cant = int(float(cant))
                    descripcion = item.find('DTE:Descripcion',ns).text
                    descripcion = descripcion.replace('\n','')
                    descripcion = descripcion.replace('\t','')
                    descripcion = descripcion.replace(',','')
                    precioUnit = float(item.find('DTE:PrecioUnitario',ns).text)
                    total = float(item.find('DTE:Total',ns).text)       
                    temp = Item(cant,descripcion,precioUnit,total)
                    operaciones.append(Operacion(fecha_obj,nitReceptor,NombreReceptor,temp,noDTE,serie)) 
            except:
                messagebox.showwarning("Error","No se pudo extraer informacion del archivo")
                continue
        operaciones_ordenadas = sorted(operaciones,key=lambda Operacion: Operacion.fecha)
        file_export_csv = path_DataXML + "\\filetemp.csv"
        mesandyear =''
        
        with open(file_export_csv,'w') as f:
            headers = 'Fecha,Serie,Factura,NIT,Nombre,Unidad,Descripcion,Total\n'
            f.write(headers)
            for operacion in operaciones_ordenadas:
                fechastr = operacion.fecha.strftime("%d/%m/%Y")
                cant = str(operacion.item.cantidad)
                descripcion = operacion.item.descripcion
                precioUnit = str(operacion.item.precio)
                total = str(operacion.item.total)

                NombreReceptor = operacion.nombreReceptor
                nitReceptor = operacion.nitReceptor
                NoDte = operacion.noDTE
                serie = operacion.serie

                strOperacion = '{},{},{},{},{},{},{},{}\n'.format(fechastr,serie,NoDte,nitReceptor,NombreReceptor,cant,descripcion,total)
                f.write(strOperacion) 
                mesandyear = operacion.fecha.strftime("%m-%Y")     
            #cerrar archivo
            f.close()
      
        try:
            #create log file with null files
            file_log = path_DataXML + "\\" + nombreEmisor +"_"+mesandyear + "_log.txt"
            with open(file_log,'w') as f:
                if len(nulls) == 0:
                    f.write('No hay archivos nulos')
                else:
                    for null in nulls:
                        f.write('Null ' + null + '\n')
                f.close()
        except:
            messagebox.showwarning("Error","No se pudo crear archivo log")

        pf = pd.read_csv(file_export_csv,encoding='utf-8')
        #column Unidad type int
        pf['Unidad'] = pf['Unidad'].astype(int)
    
        #column Factura type text 
        pf['Factura'] = pf['Factura'].astype(str)

        #column NIT type text
        pf['NIT'] = pf['NIT'].astype(str)

        #column Total type Quetzal GTQ
        pf['Total'] = pf['Total'].astype(float)
 
        try:
        
            pf.style.set_properties(**{'text-align': 'center'})
        except Exception as e:
            print(e)
        # maneja la excepción aquí


        file_export_xlsx = path_DataXML + "\\" + nombreEmisor +"_"+mesandyear + "_DETALLE_FACTURA.xlsx"
        writer = pd.ExcelWriter(file_export_xlsx)
        pf.to_excel(writer,sheet_name='REPORTE INVENTARIO POR FACTURA',index=False,header=True,encoding='utf-8',na_rep='NaN')

        try:
            format_center = writer.book.add_format()
            format_center.set_align('center')
            format_center.set_align('vcenter')
            format_center.set_border()

            format_left = writer.book.add_format()
            format_left.set_align('left')
            format_left.set_align('vcenter')
            format_left.set_border()
        except Exception as e:
            print(e)
            
        #columna Fecha
        col_fecha = pf.columns.get_loc('Fecha')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_fecha,col_fecha,12,format_center)

        #columna Serie
        col_serie = pf.columns.get_loc('Serie')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_serie,col_serie,10.5,format_center)


        #columna Factura
        col_factura = pf.columns.get_loc('Factura')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_factura,col_factura,12,format_center)
     

        #columna NIT
        col_nit = pf.columns.get_loc('NIT')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_nit,col_nit,10,format_center)
  
        #columna Nombre
        col_nombre = pf.columns.get_loc('Nombre')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_nombre,col_nombre,60,format_left)
        
        #columna Unidad
        col_unidad = pf.columns.get_loc('Unidad')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_unidad,col_unidad,7,format_center)

        #columna Descripcion
        col_descripcion = pf.columns.get_loc('Descripcion')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_descripcion,col_descripcion,60,format_left)

        #format Quetzal GTQ
        format_gtq = writer.book.add_format({'num_format': 'Q #,##0.00'})
        format_gtq.set_align('right')
        format_gtq.set_align('vcenter')
        format_gtq.set_border()

        #columna Total
        col_total = pf.columns.get_loc('Total')
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].set_column(col_total,col_total,12,format_gtq)
        
        #Freeze pane on the top row.
        writer.sheets['REPORTE INVENTARIO POR FACTURA'].freeze_panes(1, 0)

    
        writer.save()
        
        #delete file temp
        os.remove(file_export_csv)
        messagebox.showinfo("Correcto","Se analizaron los archivos correctamente")
    except:
        messagebox.showwarning("Error","No se abrio ningun archivo")

def prepareFileExcel (file_name_xls):
    try:
        if file_name_xls.endswith('.xls'):

            file_name_xlsx = file_name_xls.replace('.xls','.xlsx')
            wb_xls = xlrd.open_workbook(file_name_xls)
            wb_xlsx = openpyxl.Workbook()
            ws_xlsx = wb_xlsx.active
            for sheet_name in wb_xls.sheet_names():
                sh_xls = wb_xls.sheet_by_name(sheet_name)
                for row in range(sh_xls.nrows):
                    for col in range(sh_xls.ncols):
                        c = sh_xls.cell(row,col)
                        c.value = str(c.value).replace('Ã¡','á')
                        c.value = str(c.value).replace('Ã©','é')
                        c.value = str(c.value).replace('Ã­','í')
                        c.value = str(c.value).replace('Ã³','ó')
                        c.value = str(c.value).replace('Ãº','ú')

                        ws_xlsx.cell(row=row+1, column=col+1).value = c.value
            
            wb_xlsx.save(file_name_xlsx)

            return file_name_xlsx
        else:
            file_name_xlsx = file_name_xls
            wb_xlsx = openpyxl.load_workbook(file_name_xlsx)
            ws_xlsx = wb_xlsx.active
            for row in range(ws_xlsx.max_row):
                for col in range(ws_xlsx.max_column):
                    c = ws_xlsx.cell(row=row+1, column=col+1)
                    c.value = str(c.value).replace('Ã¡','á')
                    c.value = str(c.value).replace('Ã©','é')
                    c.value = str(c.value).replace('Ã­','í')
                    c.value = str(c.value).replace('Ã³','ó')
                    c.value = str(c.value).replace('Ãº','ú')
            wb_xlsx.save(file_name_xlsx)
            return file_name_xlsx

    except:
        messagebox.showwarning("Error","No se pudo convertir el archivo")

carga_Menu()