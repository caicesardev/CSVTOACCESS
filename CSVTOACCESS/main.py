import sqlalchemy_access.pyodbc
from tkinter import Tk
from tkinter.filedialog import askopenfile
import pandas as pd
import urllib
from sqlalchemy import create_engine


root = Tk()
root.withdraw()

print("--> Selecciona archivo CSV... <--")

csv_path = askopenfile(mode='r', filetypes=[('Archivos csv', '*.csv')])
df = pd.read_csv(csv_path.name)

print(f"Archivo CSV seleccionado: {csv_path.name}")

print("--> Selecciona archivo MDB o ACCDB... <--")

acdb_path = askopenfile(mode='r', filetypes=[(
    'Archivos accdb o mdb', '*.accdb *.mdb')])

cnn_str = (
    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};"
    f"DBQ={acdb_path.name}"
)
print("--> Abriendo Access...")
cnn_url = f"access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(cnn_str)}"
acc_engine = create_engine(cnn_url)

print("--> A continuación, escribe el nombre de tu tabla... ¡Importante! (debe estar cerrada) <--")

acc_table = input("Tabla de Access: ")

print("--> Escribiendo tabla...")

df.to_sql(f"{acc_table}", acc_engine, if_exists='replace', index=False)

print("--> Escritura finalizada.")
