from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import xlsxwriter
import tkcalendar

path_archivo = 0

dict_estado_funcional = {
    "Encendida Buena Condición": "EB",
    "Encendida LED Incompletos": "ELI",
    "Apagada": "A",
    "No Activa en Consola ACTIA": "NAC",
    "No Aparece en Consola ACTIA": "NAP",
    "Error en Consola ACTIA": "E"
}

dict_estado_caja = {
    "Caja Apernada": "CA",
    "Caja Faltan Pernos": "CFP",
    "Caja Dañada": "CD",
    "Caja Faltante": "CF",
    "Tapa Buen Estado": "TBE",
    "Tapa Sucia": "TS",
    "Con Error en Cons ACTIA": "CE"
}


def cambiarArchivo():
    global path_archivo
    global archivo

    path_archivo = filedialog.askopenfilename()
    archivo = pd.read_excel(path_archivo)

    fecha_inicial.set_date(archivo.iloc[0][0].date())
    fecha_final.set_date(archivo.iloc[-1][0].date())


def getDateRange(df):
    dfc = df

    for i, datos in archivo.iterrows():
        fecha = datos[0].date()

        if fecha < fecha_inicial.get_date() or fecha > fecha_final.get_date():
            dfc = dfc.drop(i)

    return dfc


def estadoCaja(estado_inicial):
    lista_estado = estado_inicial.split(",")

    for i in range(len(lista_estado)):
        lista_estado[i] = lista_estado[i].strip(" ")

    estado_final = dict_estado_caja[lista_estado[0]]
    iter_lista_estado = iter(lista_estado)
    next(iter_lista_estado)

    for estado in iter_lista_estado:
        estado_final += " - " + dict_estado_caja[estado]

    return estado_final


def checkCorrecto(estado_inicial):
    lista_estado = estado_inicial.replace(" ", "").split(",")

    if "TapaSucia" in lista_estado:
        return False

    elif "ConErrorenConsACTIA" in lista_estado:
        return False

    elif "CajaFaltante" in lista_estado:
        return False

    elif estado_inicial in ["No Activa en Consola ACTIA", "No Aparece en Consola ACTIA"]:
        return False

    else:
        return True


def crearInforme():
    global archivo

    copia_archivo = getDateRange(archivo)
    copia_archivo.drop_duplicates(subset="Num Bus", keep="last", inplace=True)

    nuevo_excel = xlsxwriter.Workbook("Informe.xlsx")
    pagina = nuevo_excel.add_worksheet()

    pagina.set_column(0, 0, 7)
    pagina.set_column(1, 1, 7)
    pagina.set_column(2, 2, 15)
    pagina.set_column(3, 3, 15)
    pagina.set_column(4, 4, 10)
    pagina.set_column(5, 5, 15)

    defecto = nuevo_excel.add_format({"font_size": 9, "align": "center", "border": 1, "bold": True})
    verde = nuevo_excel.add_format({"bg_color": "#A9D08E", "font_size": 8, "align": "center", "border": 1})
    naranjo = nuevo_excel.add_format({"bg_color": "#FFD966", "font_size": 8, "align": "center", "border": 1, "bold": True})
    amarillo = nuevo_excel.add_format({"bg_color": "yellow", "font_size": 8, "align": "center", "border": 1})
    titulos = nuevo_excel.add_format({"bg_color": "#B4C6E7", "font_size": 9, "align": "center", "valign": "vcenter", "border": 1, "bold": True, "text_wrap": True})

    pagina.merge_range("A1:C1", "Consultora Mips Ltda")
    pagina.merge_range("A2:C2", "Auditorias Funcionales Sistema de Cámaras")
    pagina.merge_range("A3:C3", "PATIO SAN JOSE")
    pagina.merge_range("A5:C5", "Informe Estado Funcional Flota", nuevo_excel.add_format({"font_size": 12, "bold": True}))

    pagina.merge_range("E1:F1", "Estado Caja", defecto)

    for entrada, i in zip(dict_estado_caja.keys(), range(1, len(dict_estado_caja.keys()) + 1)):
        pagina.merge_range(f"E{i + 1}:F{i + 1}", f"{entrada} = {dict_estado_caja[entrada]}", defecto)

    pagina.merge_range("H1:J1", "Estado Funcional", defecto)

    for entrada, i in zip(dict_estado_funcional.keys(), range(1, len(dict_estado_funcional.keys()) + 1)):
        pagina.merge_range(f"H{i + 1}:J{i + 1}", f"{entrada} = {dict_estado_funcional[entrada]}", defecto)

    pagina.merge_range("A9:A11", "Nº", titulos)
    pagina.merge_range("B9:B11", "Nº BUS", titulos)
    pagina.merge_range("C10:C11", "Fecha Auditoria", titulos)
    pagina.merge_range("D10:D11", "Tipo  Auditoria", titulos)
    pagina.merge_range("E10:E11", "Sistema de Cámara", titulos)
    pagina.merge_range("F10:F11", "Funcionamiento consola ACTIA", titulos)
    pagina.merge_range("G10:G11", "Problemas", titulos)

    pagina.merge_range("C9:U9", "SISTEMA DE CÁMARAS", titulos)

    for i in range(7, 20, 2):
        pagina.write(10, i, "E. Funcional", titulos)
        pagina.write(10, i + 1, "E. Caja", titulos)

    pagina.merge_range("H10:I10", "C1", titulos)
    pagina.merge_range("J10:K10", "C2", titulos)
    pagina.merge_range("L10:M10", "C3", titulos)
    pagina.merge_range("N10:O10", "C4", titulos)
    pagina.merge_range("P10:Q10", "C5", titulos)
    pagina.merge_range("R10:S10", "C6", titulos)
    pagina.merge_range("T10:U10", "C7", titulos)

    i = 11

    for j, bus in copia_archivo.iterrows():
        fecha = bus[0].to_pydatetime()
        pagina.write(i, 0, i - (11 - 1), naranjo) # num
        pagina.write(i, 1, bus[7], verde) # NUM de bus
        pagina.write(i, 2, fecha.strftime("%d.%m"), verde) # Fecha auditoria
        pagina.write(i, 3, bus[1], verde) # Tipo auditoria
        pagina.write(i, 4, f"{bus[77]}" if f"{bus[77]}" != 'nan' else '', verde) # Sistema de camara
        pagina.write(i, 5, f"{bus[78]}" if f"{bus[78]}" != 'nan' else '', verde) # funcionamiento consola actia

        if f"{bus[77]}" != 'nan':
            pagina.write(i, 6, f"{bus[79]}" if f"{bus[79]}" != 'nan' else '', verde)  # Problemas

            pagina.write(i, 7, dict_estado_funcional[f"{bus[81]}"] if f"{bus[81]}" != 'nan' else '', verde if checkCorrecto(f"{bus[81]}") else amarillo) # E.Funcional 1
            pagina.write(i, 8, estadoCaja(f"{bus[80]}") if f"{bus[80]}" != 'nan' else '', verde if checkCorrecto(f"{bus[80]}") else amarillo) # E.Caja 1

            pagina.write(i, 9, dict_estado_funcional[f"{bus[83]}"] if f"{bus[83]}" != 'nan' else '', verde if checkCorrecto(f"{bus[83]}") else amarillo)  # E.Funcional 2
            pagina.write(i, 10, estadoCaja(f"{bus[82]}") if f"{bus[82]}" != 'nan' else '', verde if checkCorrecto(f"{bus[82]}") else amarillo)  # E.Caja 2

            pagina.write(i, 11, dict_estado_funcional[f"{bus[85]}"] if f"{bus[85]}" != 'nan' else '', verde if checkCorrecto(f"{bus[85]}") else amarillo)  # E.Funcional 3
            pagina.write(i, 12, estadoCaja(f"{bus[84]}") if f"{bus[84]}" != 'nan' else '', verde if checkCorrecto(f"{bus[84]}") else amarillo)  # E.Caja 3

            pagina.write(i, 13, dict_estado_funcional[f"{bus[87]}"] if f"{bus[87]}" != 'nan' else '', verde if checkCorrecto(f"{bus[87]}") else amarillo)  # E.Funcional 4
            pagina.write(i, 14, estadoCaja(f"{bus[86]}") if f"{bus[86]}" != 'nan' else '', verde if checkCorrecto(f"{bus[86]}") else amarillo)  # E.Caja 4

            pagina.write(i, 15, dict_estado_funcional[f"{bus[89]}"] if f"{bus[89]}" != 'nan' else '', verde if checkCorrecto(f"{bus[89]}") else amarillo)  # E.Funcional 5
            pagina.write(i, 16, estadoCaja(f"{bus[88]}") if f"{bus[88]}" != 'nan' else '', verde if checkCorrecto(f"{bus[88]}") else amarillo)  # E.Caja 5

            pagina.write(i, 17, dict_estado_funcional[f"{bus[91]}"] if f"{bus[91]}" != 'nan' else '', verde if checkCorrecto(f"{bus[91]}") else amarillo)  # E.Funcional 6
            pagina.write(i, 18, estadoCaja(f"{bus[90]}") if f"{bus[90]}" != 'nan' else '', verde if checkCorrecto(f"{bus[90]}") else amarillo)  # E.Caja 6

            pagina.write(i, 19, dict_estado_funcional[f"{bus[93]}"] if f"{bus[93]}" != 'nan' else '', verde if checkCorrecto(f"{bus[93]}") else amarillo)  # E.Funcional 7
            pagina.write(i, 20, estadoCaja(f"{bus[92]}") if f"{bus[92]}" != 'nan' else '', verde if checkCorrecto(f"{bus[92]}") else amarillo)  # E.Caja 7

        else:
            for m in range(6, 21):
                pagina.write(i, m, "", verde)

        i += 1

    try:
        nuevo_excel.close()
        messagebox.showinfo("", "Terminado con exito")
    except xlsxwriter.exceptions.FileCreateError:
        messagebox.showerror("Error", "Cierre el informe e intentelo denuevo")


ventana = Tk()
ventana.title("Sistema Subus")

label1 = Label(text="Seleccione un excel: ")
label1.grid(row=0, column=0)

label1 = Label(text="Fecha inicial: ")
label1.grid(row=2, column=0)

label1 = Label(text="Fecha final: ")
label1.grid(row=3, column=0)

fecha_inicial = tkcalendar.DateEntry(width=12, background='darkblue', foreground='white', borderwidth=2)
fecha_inicial.grid(row=2, column=1)

fecha_final = tkcalendar.DateEntry(width=12, background='darkblue', foreground='white', borderwidth=2)
fecha_final.grid(row=3, column=1)

boton_cambiar_archivo = Button(text="Buscar...", command=cambiarArchivo)
boton_cambiar_archivo.grid(row=0, column=1)

boton_crear_informe = Button(text="Generar informe", command=crearInforme)
boton_crear_informe.grid(row=4, column=1)

ventana.mainloop()