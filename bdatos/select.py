
import sqlite3

def conecBase():

    conexion = sqlite3.connect("C:/Users/Paul/Desktop/bdatos/BD_ERR.db")
    mi_cursor=conexion.cursor()  
    conexion.execute("""create table TOTALES_SERVIDO (
                              FECHA TIMESTAMP NOT NULL,
                              TOTALES_DIA INTEGER
                        )""")
    conexion.commit()
    conexion.close()

conecBase()