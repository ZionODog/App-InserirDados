import tkinter as tk
from tkinter import ttk
import openpyxl

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")

def load_data():
    path = "E:\VS Code\Python_Projetos\Tela_Screen2\Relatorio.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Ativos"]

    list_values = list(sheet.values)
    print(list_values)
    for col_name in list_values[0]:
        treeview.heading(col_name, text=col_name)

    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)
        
def insert_row():
    colaborador = name_entry.get()
    email = email_entry.get()
    maquina = maquina_entry.get()
    data_registro = datar_entry.get()
    ultima_data = ultimadata_entry.get()

    print(colaborador, email, maquina, data_registro, ultima_data)

    # Insert row into Excel sheet
    path = "E:\VS Code\Python_Projetos\Tela_Screen2\Relatorio.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook["Ativos"]
    row_values = [colaborador, email, maquina, data_registro, ultima_data]
    sheet.append(row_values)
    workbook.save(path)

    # Insert row into treeview
    treeview.insert('', tk.END, values=row_values)
    
    # Clear the values
    name_entry.delete(0, "end")
    name_entry.insert(0, "Colaborador")
    email_entry.delete(0, "end")
    email_entry.insert(0, "E-mail")
    maquina_entry.delete(0, "end")
    maquina_entry.insert(0, "Máquina")
    datar_entry.delete(0, "end")
    datar_entry.insert(0, "Data de Registro")
    ultimadata_entry.delete(0, "end")
    ultimadata_entry.insert(0, "Última Data")


#Criando o App
root = tk.Tk()

style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

frame = ttk.Frame(root)
frame.pack()

widgets_frame = ttk.LabelFrame(frame, text="Adicionar Máquina")
widgets_frame.grid(row=0, column=0, padx=20, pady=20)

#Nome do Colaborador
name_entry = ttk.Entry(widgets_frame)
name_entry.insert(0, "Colaborador")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete('0', 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(5, 5), sticky="ew")

#Nome do E-mail
email_entry = ttk.Entry(widgets_frame)
email_entry.insert(0, "E-mail")
email_entry.bind("<FocusIn>", lambda e: email_entry.delete('0', 'end'))
email_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew") 

#Máquina
maquina_entry = ttk.Entry(widgets_frame)
maquina_entry.insert(0, "Máquina")
maquina_entry.bind("<FocusIn>", lambda e: maquina_entry.delete('0', 'end'))
maquina_entry.grid(row=2, column=0, padx=5, pady=(0, 5), sticky="ew")

#Data de Registro
datar_entry = ttk.Entry(widgets_frame)
datar_entry.insert(0, "Data de Registro")
datar_entry.bind("<FocusIn>", lambda e: datar_entry.delete('0', 'end'))
datar_entry.grid(row=3, column=0, padx=5, pady=(0, 5), sticky="ew")

#Ultima Data
ultimadata_entry = ttk.Entry(widgets_frame)
ultimadata_entry.insert(0, "Última Data")
ultimadata_entry.bind("<FocusIn>", lambda e: ultimadata_entry.delete('0', 'end'))
ultimadata_entry.grid(row=4, column=0, padx=5, pady=(0, 5), sticky="ew")

#Botão
button = ttk.Button(widgets_frame, text="Adicionar", command=insert_row)
button.grid(row=5, column=0, padx=5, pady=5, sticky="nsew")

separator = ttk.Separator(widgets_frame)
separator.grid(row=6, column=0, padx=(20, 10), pady=10, sticky="ew")

#Trocar o modo
mode_switch = ttk.Checkbutton(widgets_frame, text="Tema", style="Switch", command=toggle_mode)
mode_switch.grid(row=7, column=0, padx=5, pady=10, sticky="nsew")

#Coluna do relatorio
treeFrame = ttk.Frame(frame)
treeFrame.grid(row=0, column=1, pady=10)
treeScroll = ttk.Scrollbar(treeFrame)
treeScroll.pack(side="right", fill="y")

#Cols
cols = ("Colaborador", "E-mail", "Máquina", "Data de Registro", "Última Data")
treeview = ttk.Treeview(treeFrame, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Colaborador", width=260)
treeview.column("E-mail", width=260)
treeview.column("Máquina", width=150)
treeview.column("Data de Registro", width=120)
treeview.column("Última Data", width=120)
treeview.pack()
treeScroll.config(command=treeview.yview)
load_data()

root.mainloop()