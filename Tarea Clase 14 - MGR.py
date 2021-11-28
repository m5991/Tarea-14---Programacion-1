import pyodbc
import pandas as pd
import tkinter as tk

from tkinter import *
from tkinter import ttk #Importamos todas las funciones que contiene tkinter
from tkinter.ttk import *
from tkinter import messagebox

# Se importa primeramente la tabla a trabajar para pasarla a SQL

carreras=pd.read_csv('C:/Users/Manuel Gastelum/Clase 13/Tarea 14/Maraton NY completo.csv', 
                  engine='python') 
carreras= carreras.fillna(value=0)
Lista_valores =  carreras.values.tolist()

for inner_list in Lista_valores:
	inner_list[5] = round(inner_list[5],2)

#print(Lista_valores)
tuplas_lista = tuple(Lista_valores) 
#print(tuplas_lista)

server = 'VINDMGASTELUMTE\MSSQLSERVER01'
conexion = pyodbc.connect('DRIVER={SQL Server};SERVER='+server, autocommit=True)


cursor = conexion.cursor()
cursor.execute("IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'Nueva') BEGIN CREATE DATABASE Nueva END")
cursor.execute("DROP DATABASE Nueva") 
cursor.execute("CREATE DATABASE Nueva")  
conexion.close()


conexion = pyodbc.connect(driver='{SQL server}', host = server, database = "Nueva")
cursor = conexion.cursor()
cursor.execute("CREATE TABLE MaratonNY_Python (Corredor INT, place INT, gender VARCHAR(25), age INT, home VARCHAR(10), time FLOAT)")
cursor.executemany("INSERT INTO MaratonNY_Python VALUES(?,?,?,?,?,?)", tuplas_lista)
cursor.commit()
conexion.close()

class General:
	def __init__(self, raiz):
		self.genero = StringVar()
		self.label_genero = Label(raiz, text = "Género")
		self.label_genero.grid(column=0, row=0)
		self.genero = Combobox(raiz, values=('Female', 'Male'), width=5)
		self.genero.grid(column=0, row=1)

		self.origen = StringVar()
		self.label_origen = Label(raiz, text = "Origen")
		self.label_origen.grid(column=0, row=5)
		self.origen = Combobox(raiz, values=("GBR", "NY", "FRA", "MI", "IRL", "GER", "Otro"), width=10)
		self.origen.grid(column=0,row=6)
        
		self.time = StringVar()
		self.label_time = Label(raiz, text = "Tiempo")
		self.label_time.grid(column=0, row=10)
		self.time = Combobox(raiz, values=("menos de 200 min", "entre 200 y 250 min", "entre 250 y 300 min", "más de 300 min", "NULL"), width=10)
		self.time.grid(column=0,row=11)

		#Creamos los botones
		#Con command le decimos cual función queremos que lleve a cabo
		self.boton_buscar= Button(raiz, text="Buscar", command=self.buscar)
		self.boton_buscar.grid(column=0, row=30)

		self.boton_borrar=Button(raiz, text="Borrar", command=self.borrar)
		self.boton_borrar.grid(column=0, row=40)

		#Tabla
		self.tabla=ttk.Treeview(raiz, column=("c1", "c2", "c3", "c4"), show='headings', height=8)
		self.tabla.column("# 1",anchor=CENTER, stretch=NO, width=100)
		self.tabla.heading("# 1", text="Corredor")
		self.tabla.column("# 2", anchor=CENTER, stretch=NO)
		self.tabla.heading("# 2", text="Género")
		self.tabla.column("# 3", anchor=CENTER, stretch=NO)
		self.tabla.heading("# 3", text="Origen")
		self.tabla.column("# 4", anchor=CENTER, stretch=NO)
		self.tabla.heading("# 4", text="Tiempo")    
		self.tabla.grid(column=0, row=50)
        
	def buscar(self):
		self.tabla.delete(*self.tabla.get_children())
		server = 'VINDMGASTELUMTE\MSSQLSERVER01'
		bd = 'Nueva'
		genero_valor = "'" + self.genero.get() + "'"
		#print(genero_valor)

		origen = "'" + self.origen.get() + "'"
		time = "'%" + self.time.get().lower() + "%'"

		conexion = pyodbc.connect(driver='{SQL server}', host = server, database = bd)
		

		#Creamos un cursor para almacenar la información en memoria
		cursor = conexion.cursor()
		if self.origen.get()=='Otro':
			instruccion = "SELECT Corredor, gender, home, time FROM MaratonNY_Python WHERE gender= " + genero_valor + " AND home <> 'GBR' AND home <> 'NY' AND home <> 'FRA' AND home <> 'MI' AND home <> 'IRL' AND home <> 'GER'"
		else:
			instruccion = "SELECT Corredor, gender, home, time FROM MaratonNY_Python WHERE gender= " + genero_valor + " AND home = " + origen
		if self.time.get() == "menos de 200 min":
			instruccion = instruccion + " AND time < 200"
		elif self.time.get() == "entre 200 y 250 min":
			instruccion = instruccion + " AND time < 250 AND time > 200"
		elif self.time.get() == "entre 250 y 300 min":
			instruccion = instruccion + " AND time < 300 AND time > 250"
		elif self.time.get() == "más de 300 min":
			instruccion = instruccion + " AND time > 300"
		cursor.execute(instruccion)
		datos_clientes = cursor.fetchall()
		print(datos_clientes)
		conexion.commit()


		#Nos aseguramos de cerrar la conexión
		conexion.close()
        
		#Mandamos la información a la tabla
		for row in datos_clientes:
			self.tabla.insert('', 'end', values=((row[0],row[1],row[2],row[3])))
		#messagebox.showinfo("Resultados", datos_clientes)

	#def desplegar_resultados(self):

	def borrar(self):
		self.tabla.delete(*self.tabla.get_children())
		self.genero.set("")
		self.origen.set("")





#Creamos el objeto que será la raiz de la aplicación
raiz = Tk()
#Le agregamos un título
raiz.title("Filtrador de tabla de corredores")
#Determinamos si se podrá cambiar su tamaño
raiz.resizable(1,1)
#Asignamos un logotipo
raiz.iconbitmap('objetos.ico')
#Asignamos un tipo de cursor, un color de background y un borde a la raiz
raiz.config(bd=8)
raiz.config(relief="ridge")
estructura = General(raiz)

raiz.mainloop()