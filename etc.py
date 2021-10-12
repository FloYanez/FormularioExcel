from tkinter import *
from openpyxl import load_workbook

file_name = "Registro Datos.xlsm"
wb = load_workbook(file_name)

sheet = wb.active
target = wb.copy_worksheet(sheet)


# Limpia formulario
def limpiar():
    saludo.config(text="hola")


# Guarda info en excel
def guardar():
    saludo.config(text="hello")


# crear ventana (con barra de título)
ventana = Tk()

# Titulo
titulo = Label(ventana, text="Formulario")
titulo.pack()

# Fecha
fecha = Frame(ventana)
fecha.pack()
fecha_label = Label(fecha, text="Fecha")
fecha_label.pack(side=LEFT)
fecha_entry = Entry(fecha)
fecha_entry.pack()

# Número de operación
n_op = Frame(ventana)
n_op.pack()
n_op_label = Label(n_op, text="N° OP")
n_op_label.pack(side=LEFT)
n_op_entry = Entry(n_op)
n_op_entry.pack()

# Código
codigo = Frame(ventana)
codigo.pack()
codigo_label = Label(codigo, text="Código")
codigo_label.pack(side=LEFT)
codigo_entry = Entry(codigo)
codigo_entry.pack()

# Cantidad programada
programado = Frame(ventana)
programado.pack()
programado_label = Label(programado, text="Programado")
programado_label.pack(side=LEFT)
programado_entry = Entry(programado)
programado_entry.pack()

# Efectivo
efectivo = Frame(ventana)
efectivo.pack()
efectivo_label = Label(efectivo, text="Efectivo")
efectivo_label.pack(side=LEFT)
efectivo_entry = Entry(efectivo)
efectivo_entry.pack()

# Defectuosos
defectuosos = Frame(ventana)
defectuosos.pack()
defectuosos_label = Label(defectuosos, text="Defectuosos")
defectuosos_label.pack(side=LEFT)
defectuosos_entry = Entry(defectuosos)
defectuosos_entry.pack()

# Hora Inicio OP
hora_inicio_op = Frame(ventana)
hora_inicio_op.pack()
hora_inicio_op_label = Label(hora_inicio_op, text="Hora Inicio OP")
hora_inicio_op_label.pack(side=LEFT)
hora_inicio_op_entry = Entry(hora_inicio_op)
hora_inicio_op_entry.pack()

# Hora Término OP
hora_termino_op = Frame(ventana)
hora_termino_op.pack()
hora_termino_op_label = Label(hora_termino_op, text="Hora Término OP")
hora_termino_op_label.pack(side=LEFT)
hora_termino_op_entry = Entry(hora_termino_op)
hora_termino_op_entry.pack()

# Dotación
dotacion = Frame(ventana)
dotacion.pack()
dotacion_label = Label(dotacion, text="Dotación")
dotacion_label.pack(side=LEFT)
dotacion_entry = Entry(dotacion)
dotacion_entry.pack()

# HHEE
hhee = Frame(ventana)
hhee.pack()
hhee_label = Label(hhee, text="HHEE")
hhee_label.pack(side=LEFT)
hhee_entry = Entry(hhee)
hhee_entry.pack()

# Paros Programados
paros_programados = Frame(ventana)
paros_programados.pack()
paros_programados_label = Label(paros_programados, text="Paros Programados")
paros_programados_label.pack(side=LEFT)
paros_programados_entry = Entry(paros_programados)
paros_programados_entry.pack()

# Paros por Fallas
paros_por_fallas = Frame(ventana)
paros_por_fallas.pack()
paros_por_fallas_label = Label(paros_por_fallas, text="Paros por Fallas")
paros_por_fallas_label.pack(side=LEFT)
paros_por_fallas_entry = Entry(paros_por_fallas)
paros_por_fallas_entry.pack()

# Descripción PP
descripcion_pp = Frame(ventana)
descripcion_pp.pack()
descripcion_pp_label = Label(descripcion_pp, text="Descripción PP")
descripcion_pp_label.pack(side=LEFT)
descripcion_pp_entry = Entry(descripcion_pp)
descripcion_pp_entry.pack()

# Descripción PnP
descripcion_pnp = Frame(ventana)
descripcion_pnp.pack()
descripcion_pnp_label = Label(descripcion_pnp, text="Descripción PnP")
descripcion_pnp_label.pack(side=LEFT)
descripcion_pnp_entry = Entry(descripcion_pnp)
descripcion_pnp_entry.pack()

# marco para agrupar pregunta y respuesta
botones = Frame(ventana)
botones.pack()
limpiar = Button(botones, text="Limpiar", command=limpiar)
limpiar.pack(side=LEFT)
guardar = Button(botones, text="Guardar", command=guardar)
guardar.pack()

# mostrar ventana y esperar cierre (click en botón X)
ventana.mainloop()
