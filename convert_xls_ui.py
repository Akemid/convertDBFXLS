import time
#Para la funcionalidad del UI
from tkinter import (
    messagebox,
    Tk,
    Frame,
    Label,
    Entry,
    StringVar,
    Button,
    filedialog,    
   #raiz fileDialog
)
#from tkinter.ttk import Progressbar
#para la funcionalidad de la conversion
from dbfread import DBF
import xlwt
import progressbar
from glob2 import glob

#---------------------Interfaz gráfica------------------------------------------#
raiz = Tk()
raiz.title("Convertidor DBF a XLS -- Version Beta")
raiz.resizable(False,False)
#raiz.iconbitmap("akemid.ico")
#raiz.geometry("650x350")
raiz.config(bg="#34656d")
#Variables
rutaOrigen=StringVar()
rutaDestino=StringVar()
progreso=StringVar()

miFrame=Frame()
miFrame.pack()
miFrame.config(width="800",height="350")
miFrame.config(bg="#34656d")
miFrame.config(cursor="target")
##Labels
lblRutaOrigen=Label(miFrame,text="Ruta carpeta origen:",bg="#34656d",fg="white",font=(12))
lblRutaDestino=Label(miFrame,text="Ruta carpeta destino:",bg="#34656d",fg="white",font=(12))
lblRutaOrigen.grid(row=0,column=0,sticky="e",pady=2,padx=5)
lblRutaDestino.grid(row=1,column=0,stick="e",pady=2,padx=5)
#inputs
#root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))

txtRutaOrigen=Entry(miFrame,textvariable=rutaOrigen)
txtRutaOrigen.grid(row=0,column=1,pady=2,padx=5,columnspan=3)
txtRutaOrigen.config(width="80")
txtRutaDestino=Entry(miFrame,textvariable=rutaDestino)
txtRutaDestino.grid(row=1,column=1,pady=2,padx=5,columnspan=3)
txtRutaDestino.config(width="80")
#raiz.directory = fileDialog.askdirectory()
##-----------------------------------------------PROGRESSBAR-----------------------------------------------------


#----------------------FUNCIONALIDAD CONVERSION-------------------------------------#

def convertir_dbf_xls(rutaOrigen,rutaDestino):
    print("Convertir de dbf a .xls")

    files_dbf=glob(rutaOrigen+'/*.DBF')
    #dbf_filename=files_dbf[0]

    progress = progressbar.ProgressBar(initial_value=0,max_value=len(files_dbf),min_value=0).start()
    j=0
    #nro_archivos=len(files_dbf)
    #aumento = nro_archivos / 100
    
    
    #progress.bottom.columnconfigure(0, weight=1)
    
    for dbf_filename in files_dbf:
        xls_filename = dbf_filename.replace('dbf','xls')
        #print(xls_filename)
        table = DBF(dbf_filename, encoding='unicode_escape')
        all_sheet = []
        row = 0 # control the number of rows
        write_row = 0
        sheet_list = []
        # Create an new Excel file and add a worksheet.
        book = xlwt.Workbook() # Create a new excel
        sheet = book.add_sheet('all_sheet') # Add a sheet page
        #progress.start()
        for record in table:
                col = 0
                if all_sheet == []: # This is to control only the field name once
                    sheet_dict = record.keys()
                    #print(sheet_dict)
                    # print(type(sheet_dict))
                    # <class 'odict_keys'>
                    #sheet_list = list(set(sheet_dict)) # Convert odict_keys to a list for operation
                    sheet_list = list(sheet_dict) # Convert odict_keys to a list for operation
                    #print(sheet_list)
                    all_sheet = sheet_list
                    #print(all_sheet)
                    if write_row == 0: # to control only the field name is written once
                        col = 0
                        for i in range(len(sheet_list)):
                            sheet.write(row, col, sheet_list[i])
                            col += 1
                        col = 0
                        row += 1
                        write_row += 1
                for field in record:
                    sheet.write(row, col, record[field])
                    # print(field,'=',record[field],end='')
                    col += 1
                row += 1
                #raiz.update_idletasks()
                #time.sleep(1)
        #j+=1

        book.save(r''+rutaDestino+'/'+dbf_filename.split('\\')[1].split('.')[0]+'.xls')
        j+=1
        #progreso.set(str(progress['value'])+"/"+str(nro_archivos))    
        #progress['value'] +=aumento
        #print(aumento)
        #print(progress['value'])
        #raiz.update_idletasks()
        #time.sleep(1)
        progress.update(j)
        #progress.step()
        
    progress.finish()
    #progress.stop()
    #j=0
    print("El cambio de formato ha concluido con éxito")
    messagebox.showinfo("Mensaje", "Archivos convertidos")

##----------------------------------------------------------------#
# lblProgreso=Label(miFrame,bg="#34656d",fg="white",font=(12),text="Progreso")
# lblProgreso.grid(row=2,column=0,sticky="e",pady=2,padx=5)

# progress=Progressbar(miFrame,orient='horizontal',mode='determinate',length=480)
# progress.grid(row=2,column=2,pady=5,padx=5)


# lblProgreso2=Label(miFrame,bg="#34656d",fg="white",font=(12),textvariable=progreso)

# lblProgreso2.grid(row=2,column=4,sticky="e",pady=2,padx=5)
# progreso.set("0")
#--------------------------------------------boton
def obtenerRutaOrigen():
    rutaOrigen.set(filedialog.askdirectory(title="Seleccionar ruta de origen"))
def obtenerRutaDestino():
    rutaDestino.set(filedialog.askdirectory(title="Seleccionar ruta de destino"))    
    #if valor == "Yes":
def codigoBoton():
    valor=messagebox.askquestion("Comfirmación","¿Esta seguro de convertir los archivos?")
    
    if valor == "yes":
        convertir_dbf_xls(rutaOrigen.get(),rutaDestino.get())
        #try:
        #    convertir_dbf_xls(rutaOrigen.get(),rutaDestino.get())
        #except:
        #    messagebox.showerror("Error","Error al convertir los archivos")    
    else:
        pass
def completo():    
    messagebox.showinfo("Mensaje", "Archivos convertidos")

def errorConvertir():
    messagebox.showError("Error","Error al convertir los archivos")    
    
btnConvertir=Button(miFrame,text="Convertir",command=codigoBoton,font=(12))
btnConvertir.grid(row=0,column=5,pady=15,rowspan=2)
btnConvertir.config(width="18",height="3",padx=4)

btnBuscarOrigen=Button(miFrame,text="...",command=obtenerRutaOrigen)
btnBuscarOrigen.grid(row=0,column=4,pady=15,padx=4)
btnBuscarOrigen.config(width="2",height="1")

btnBuscarDestino=Button(miFrame,text="...",command=obtenerRutaDestino)
btnBuscarDestino.grid(row=1,column=4,pady=15,padx=4)
btnBuscarDestino.config(width="2",height="1")
#Espacio en blanco
lblEspacio=Label(miFrame,text="",bg="#34656d",fg="white",font=(12))
lblEspacio.grid(row=0,column=6,sticky="w")
#btnConvertir.config(width="20",height="3",padx=2)







raiz.mainloop()