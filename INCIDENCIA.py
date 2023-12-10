import sqlite3
import pandas as pd
import numpy as np
import sys
import os
from PyQt5 import uic, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QDialog, QPushButton, QLabel, QProgressBar, QTableWidgetItem, QMainWindow, QLCDNumber
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import *
from PyQt5.QtCore import QTimer, QTime, Qt
from PyQt5.QtGui import QIcon, QPixmap
from datetime import datetime
import matplotlib.pyplot as plot   

class Consulta_incidencias(QDialog):
    def __init__(self):
        super().__init__()
        flags = Qt.WindowFlags()
        #self.setWindowFlags(Qt.WindowMinimizeButtonHint|Qt.WindowCloseButtonHint)
        self.setWindowFlags(Qt.WindowMinimizeButtonHint)

        # NUMERO 1/1
        #nombre_archivo2=self.resolver_ruta2("INCIDENCIAS_SE.ui")
        #uic.loadUi(nombre_archivo2, self)
        uic.loadUi("C:/Users/Paul/Desktop/PROYECTOS INFORMATICOS/FICHERO EXCEL_PYTHON/INCIDENCIAS_SE.ui", self)

        self.pushButton_BD.setStyleSheet("background-color : blue")
        self.pushButton_salir.setStyleSheet("background-color: red;color: white;")

        self.comboBox_error.addItem("")
        self.comboBox_error.addItem("Preparación")
        self.comboBox_error.addItem("Recepción")
        self.comboBox_error.addItem("Manipulación")
        self.comboBox_error.addItem("Otros")
        ###########################################
        self.lineEdit_indice.setDisabled(True)
        self.lineEdit_fecha.setDisabled(True)
        self.lineEdit_codproducto.setDisabled(True)
        self.lineEdit_descripcion.setDisabled(True)
        self.comboBox_error.setDisabled(True)
        self.lineEdit_unidades.setDisabled(True)
        self.lineEdit_pedidos.setDisabled(True)

        self.lineEdit_PRODUCTO.setDisabled(True)
        self.lineEdit_DESCRIPCION.setDisabled(True)

        self.lineEdit_fechaini.setDisabled(True)
        self.lineEdit_fechafin.setDisabled(True)

        self.pushButton_obtener.setDisabled(True)
        self.pushButton_exp_reg.setDisabled(True)
        self.pushButton_filtro.setDisabled(True)

        self.pushButton_ejecutar.setDisabled(True)
        self.pushButton_modif.setDisabled(True)
        self.pushButton_elim.setDisabled(True)

        self.pushButton_graf_inc.setDisabled(True)
        self.pushButton_graf_err.setDisabled(True)
        self.pushButton_exp_datfilt.setDisabled(True)
        ############################################
        self.radioButton_insertar.toggled.connect(self.onClicked)
        self.radioButton_modificar.toggled.connect(self.onClicked)
        self.radioButton_eliminar.toggled.connect(self.onClicked)
        self.radioButton_incidencia.toggled.connect(self.onClicked)
        self.radioButton_producto.toggled.connect(self.onClicked)

        #############################################
        self.pushButton_BD.clicked.connect(self.obtener_dir_db)
        self.pushButton_obtener.clicked.connect(self.ejec_obtener_reg)
        self.pushButton_limpiar.clicked.connect(self.Clear_cam)
        self.pushButton_ejecutar.clicked.connect(self.ejec_Insertar)
        self.pushButton_modif.clicked.connect(self.ejec_modificar)
        self.pushButton_elim.clicked.connect(self.ejec_eliminar)
        self.pushButton_cargarTab.clicked.connect(self.cargarTablas)
        self.pushButton_exp_reg.clicked.connect(self.exportar_registros)
        self.pushButton_filtro.clicked.connect(self.filtrado_fec)
        self.pushButton_graf_inc.clicked.connect(self.grafico_incidencia)
        self.pushButton_graf_err.clicked.connect(self.grafico_error)
        self.pushButton_exp_datfilt.clicked.connect(self.exportar_datosFiltrados)
        self.pushButton_salir.clicked.connect(self.close)

        #############################################################
        QMessageBox.about(self, "INFORMACION","Para empezar conecte la base de datos: BD_ERR")
        ################################################################

    def obtener_dir_db(self):
        try:
            ruta_archivobd = QFileDialog.getOpenFileName(self, "Buscar Archivo...")
            self.ruta_archivobd=ruta_archivobd[0]
            self.ruta_archivobd_1=ruta_archivobd[0]
            self.ruta_archivobd_2=self.ruta_archivobd_1[-9:]
            if self.ruta_archivobd_2== 'BD_ERR.db':
                QMessageBox.about(self, "INFORMACION","Conexión realizada.")
                self.pushButton_BD.setStyleSheet("background-color : gainsboro")
                self.pushButton_BD.setDisabled(True)
            else:
                QMessageBox.about(self, "INFORMACION","No se ha seleccionado ningun archivo/el archivo seleccionado no es el correcto.")
        except:
            QMessageBox.about(self, "INFORMACION","Error, no se ha podido conectar a la base de datos.")

    def conectar_base2(self):
        self.conexion2=sqlite3.connect(self.ruta_archivobd)
        self.cursor2=self.conexion2.cursor()

    def desconectar_base2(self):     
        self.conexion2.commit()
        self.conexion2.close() 

    """
    try:
        conexion2.execute(###CREATE TABLE INCIDENCIA (
                                ID_REG INTEGER PRIMARY KEY AUTOINCREMENT,
                                FECHA TEXT ,
                                COD_PRODUCTO INTEGER,
                                DESCRIPCION TEXT,
                                ERROR TEXT,
                                UNIDADES INTEGER,
                                Nº_PEDIDOS INTEGER
                            )###)
    except sqlite3.OperationalError:
        print("La tabla articulos ya existe")                    


    try:
        con_bd.execute(###CREATE TABLE PRODUCTO (
                                COD_PRODUCTO INTEGER ,
                                DESCRIPCION TEXT
                            )###)
    except sqlite3.OperationalError:
        print("La tabla productos ya existe")                    
    con_bd.close()
    
    archivo = 'c:/Users/Paul/Desktop/PROYECTOS INFORMATICOS/BASE DATOS ERRORES PYTHON/totales.xlsx'

    df = pd.read_excel(archivo, sheet_name='tot')
    list_df=df.to_numpy().tolist()
    cont=len(list_df)
    
    for i in range (0,cont):

        dato1=list_df[i][0]
        dato2=list_df[i][1]
        sql=(###INSERT INTO PRODUCTO(COD_PRODUCTO, DESCRIPCION) VALUES (?,?)###)
        cursor2.execute(sql, (dato1, dato2))
    """

    def onClicked(self):
        if self.radioButton_insertar.isChecked() and self.radioButton_incidencia.isChecked():
            self.lineEdit_fecha.setDisabled(False)
            self.lineEdit_codproducto.setDisabled(False)
            self.comboBox_error.setDisabled(False)
            self.lineEdit_unidades.setDisabled(False)
            self.lineEdit_pedidos.setDisabled(False)
            self.lineEdit_PRODUCTO.setDisabled(True)
            self.lineEdit_DESCRIPCION.setDisabled(True)
            self.pushButton_obtener.setDisabled(True)
            self.pushButton_ejecutar.setDisabled(False)
            self.pushButton_modif.setDisabled(True)
            self.pushButton_elim.setDisabled(True)
            self.Clear_cam()
        if self.radioButton_insertar.isChecked() and self.radioButton_producto.isChecked():
            self.lineEdit_indice.setDisabled(True)
            self.lineEdit_fecha.setDisabled(True)
            self.lineEdit_codproducto.setDisabled(True)
            self.lineEdit_descripcion.setDisabled(True)
            self.comboBox_error.setDisabled(True)
            self.lineEdit_unidades.setDisabled(True)
            self.lineEdit_pedidos.setDisabled(True)
            self.lineEdit_PRODUCTO.setDisabled(False)
            self.lineEdit_DESCRIPCION.setDisabled(False)
            self.pushButton_obtener.setDisabled(True)
            self.pushButton_ejecutar.setDisabled(False)
            self.pushButton_modif.setDisabled(True)
            self.pushButton_elim.setDisabled(True)
            self.Clear_cam()
        if self.radioButton_modificar.isChecked() and self.radioButton_incidencia.isChecked():
            self.lineEdit_indice.setDisabled(False)
            self.lineEdit_fecha.setDisabled(True)
            self.lineEdit_codproducto.setDisabled(True)
            self.lineEdit_descripcion.setDisabled(True)
            self.comboBox_error.setDisabled(True)
            self.lineEdit_unidades.setDisabled(True)
            self.lineEdit_pedidos.setDisabled(True)
            self.lineEdit_PRODUCTO.setDisabled(True)
            self.lineEdit_DESCRIPCION.setDisabled(True)
            self.pushButton_obtener.setDisabled(False)
            self.pushButton_ejecutar.setDisabled(True)
            self.pushButton_modif.setDisabled(False)
            self.pushButton_elim.setDisabled(True)
            self.Clear_cam()
        if  self.radioButton_modificar.isChecked() and self.radioButton_producto.isChecked():
            self.lineEdit_indice.setDisabled(True)
            self.lineEdit_fecha.setDisabled(True)
            self.lineEdit_codproducto.setDisabled(True)
            self.lineEdit_descripcion.setDisabled(True)
            self.comboBox_error.setDisabled(True)
            self.lineEdit_unidades.setDisabled(True)
            self.lineEdit_pedidos.setDisabled(True)
            self.lineEdit_PRODUCTO.setDisabled(False)
            self.lineEdit_DESCRIPCION.setDisabled(True)
            self.pushButton_obtener.setDisabled(False)
            self.pushButton_ejecutar.setDisabled(True)
            self.pushButton_modif.setDisabled(False)
            self.pushButton_elim.setDisabled(True)
            self.Clear_cam()
        if  self.radioButton_eliminar.isChecked() and self.radioButton_incidencia.isChecked():
            self.lineEdit_indice.setDisabled(False)
            self.lineEdit_fecha.setDisabled(True)
            self.lineEdit_codproducto.setDisabled(True)
            self.lineEdit_descripcion.setDisabled(True)
            self.comboBox_error.setDisabled(True)
            self.lineEdit_unidades.setDisabled(True)
            self.lineEdit_pedidos.setDisabled(True)
            self.lineEdit_PRODUCTO.setDisabled(True)
            self.lineEdit_DESCRIPCION.setDisabled(True)
            self.pushButton_obtener.setDisabled(False)
            self.pushButton_ejecutar.setDisabled(True)
            self.pushButton_modif.setDisabled(True)
            self.pushButton_elim.setDisabled(False)
            self.Clear_cam()
        if  self.radioButton_eliminar.isChecked() and self.radioButton_producto.isChecked():
            self.lineEdit_indice.setDisabled(True)
            self.lineEdit_fecha.setDisabled(True)
            self.lineEdit_codproducto.setDisabled(True)
            self.lineEdit_descripcion.setDisabled(True)
            self.comboBox_error.setDisabled(True)
            self.lineEdit_unidades.setDisabled(True)
            self.lineEdit_pedidos.setDisabled(True)
            self.lineEdit_PRODUCTO.setDisabled(False)
            self.lineEdit_DESCRIPCION.setDisabled(True)
            self.pushButton_obtener.setDisabled(False)
            self.pushButton_ejecutar.setDisabled(True)
            self.pushButton_modif.setDisabled(True)
            self.pushButton_elim.setDisabled(False)
            self.Clear_cam()
        if self.radioButton_insertar.isChecked():
            self.lineEdit_indice.setDisabled(True)
            self.lineEdit_descripcion.setDisabled(True)

    def insertar_incidencia(self):
        self.conectar_base2()
        try:    
            fecha=self.lineEdit_fecha.text()
            codigo_prod=int(self.lineEdit_codproducto.text())
            self.cursor2.execute("SELECT * FROM PRODUCTO WHERE COD_PRODUCTO=?", (codigo_prod,))
            datos = self.cursor2.fetchall()
            descripp=(datos[0][1])
            descripcion=descripp
            error=self.comboBox_error.currentText()
            unid=int(self.lineEdit_unidades.text())
            npedidos=self.lineEdit_pedidos.text()
            sql=("""INSERT INTO INCIDENCIA(FECHA, COD_PRODUCTO, DESCRIPCION, ERROR, UNIDADES, Nº_PEDIDOS) VALUES (?,?,?,?,?,?)""")
            self.cursor2.execute(sql, (fecha, codigo_prod,descripcion,error,unid,npedidos))
            self.conexion2.commit()
            QMessageBox.about(self, "INFORMACION","Registro insertado correctamente.")  
        except:
            QMessageBox.about(self, "INFORMACION","Error en la generación del registro, revise que los datos introducidos sean correctos.")
        self.conexion2.close() 

    def insertar_producto(self):
        self.conectar_base2()
        try:
            cod_prod=int(self.lineEdit_PRODUCTO.text())
            descr=self.lineEdit_DESCRIPCION.text()
            self.cursor2.execute("SELECT COD_PRODUCTO FROM PRODUCTO WHERE COD_PRODUCTO=?", (cod_prod,))
            date_producto=self.cursor2.fetchall()
            if date_producto == []:
                sql_p=("""INSERT INTO PRODUCTO(COD_PRODUCTO, DESCRIPCION) VALUES (?,?)""")
                self.cursor2.execute(sql_p, (cod_prod, descr))
                QMessageBox.about(self, "INFORMACION","Registro insertado correctamente.")
            else:
                QMessageBox.about(self, "INFORMACION","El producto ya existe")
            self.conexion2.commit() 
        except:
            QMessageBox.about(self, "INFORMACION","Error en la generación del registro, revise que los datos introducidos sean correctos.")
        self.conexion2.close()

    def modificar_incidencia(self):
        self.conectar_base2()
        try:
            fecha=self.lineEdit_fecha.text()
            codigo_prod=int(self.lineEdit_codproducto.text())
            descripcion=self.lineEdit_descripcion.text()
            error=self.comboBox_error.currentText()
            unid=int(self.lineEdit_unidades.text())
            npedidos=self.lineEdit_pedidos.text()
            id=int(self.lineEdit_indice.text())
            sql_q=('''UPDATE INCIDENCIA SET FECHA = ?, COD_PRODUCTO=?, DESCRIPCION=?, ERROR=?, UNIDADES=?, Nº_PEDIDOS=? WHERE ID_REG=?;''')
            self.cursor2.execute(sql_q, (fecha, codigo_prod,descripcion,error,unid,npedidos, id))
            self.conexion2.commit()
            QMessageBox.about(self, "INFORMACION","Registro modificado correctamente.")
        except:
            QMessageBox.about(self, "INFORMACION","Error en la modificación del registro, revise que los datos introducidos sean correctos.")
        self.conexion2.close()

    def modificar_producto(self):
        self.conectar_base2()
        try:
            codigo_prod=int(self.lineEdit_PRODUCTO.text())
            descripcion=self.lineEdit_DESCRIPCION.text()
            sql_r=('''UPDATE PRODUCTO SET DESCRIPCION=? WHERE COD_PRODUCTO=?;''')
            self.cursor2.execute(sql_r,(descripcion, codigo_prod))
            self.conexion2.commit()
            QMessageBox.about(self, "INFORMACION","Registro modificado correctamente.")
        except:
            QMessageBox.about(self, "INFORMACION","Error en la modificación del registro, revise que los datos introducidos sean correctos.")
        self.conexion2.close()

    def eliminar_incidencia(self):
        self.conectar_base2()
        try:
            iden=int(self.lineEdit_indice.text())
            sql_el=('''DELETE FROM INCIDENCIA WHERE ID_REG=?;''')
            self.cursor2.execute(sql_el, (iden, ))
            self.conexion2.commit()
            QMessageBox.about(self, "INFORMACION","Registro eliminado correctamente.")
        except:
            QMessageBox.about(self, "INFORMACION","Error, el registro no se pudo eliminar.")
        self.conexion2.close()

    def eliminar_producto(self):
        self.conectar_base2()
        try:
            codigo_prod=int(self.lineEdit_PRODUCTO.text())
            sql_ll=('''DELETE FROM PRODUCTO WHERE COD_PRODUCTO=?;''')
            self.cursor2.execute(sql_ll, (codigo_prod, ))
            self.conexion2.commit()
            QMessageBox.about(self, "INFORMACION","Registro modificado correctamente.")
        except:
            QMessageBox.about(self, "INFORMACION","Error, el registro no se pudo eliminar")
        self.conexion2.close()

    def insertar_tablewidget_incidencia(self):
        try:  
            self.tableWidget_errores.clearContents()
            self.dftb_inc = pd.read_sql_query("SELECT * FROM INCIDENCIA", self.conexion2) 
            
            lista_incidencia=self.dftb_inc.to_numpy().tolist()
            w=len(lista_incidencia)   
            for i in range (0, w):
                self.tableWidget_errores.insertRow(i)
                self.tableWidget_errores.setItem(i, 0, QTableWidgetItem(str(lista_incidencia[i][0])))
                self.tableWidget_errores.setItem(i, 1, QTableWidgetItem(str(lista_incidencia[i][1])))
                self.tableWidget_errores.setItem(i, 2, QTableWidgetItem(str(lista_incidencia[i][2])))
                self.tableWidget_errores.setItem(i, 3, QTableWidgetItem(str(lista_incidencia[i][3])))
                self.tableWidget_errores.setItem(i, 4, QTableWidgetItem(str(lista_incidencia[i][4])))
                self.tableWidget_errores.setItem(i, 5, QTableWidgetItem(str(lista_incidencia[i][5])))
                self.tableWidget_errores.setItem(i, 6, QTableWidgetItem(str(lista_incidencia[i][6])))
            self.pushButton_exp_reg.setDisabled(False)
        except:
            QMessageBox.about(self, "INFORMACION","Error, no se han podido cargar los datos.")

    def insertar_tablewidget_producto(self):
        try:   
            self.tableWidget_producto.clearContents()       
            self.dftb_prod = pd.read_sql_query("SELECT * FROM PRODUCTO", self.conexion2)     
            lista_producto=self.dftb_prod.to_numpy().tolist()
            w=len(lista_producto)         
            for i in range (0, w):
                self.tableWidget_producto.insertRow(i)
                self.tableWidget_producto.setItem(i, 0, QTableWidgetItem(str(lista_producto[i][0])))
                self.tableWidget_producto.setItem(i, 1, QTableWidgetItem(str(lista_producto[i][1])))      
        except:
            QMessageBox.about(self, "INFORMACION","Error, no se han podido cargar los datos.")

    def cargarTablas(self):
        try:
            self.conectar_base2()
            self.insertar_tablewidget_incidencia()
            self.insertar_tablewidget_producto()
            self.desconectar_base2()
            self.lineEdit_fechaini.setDisabled(False)
            self.lineEdit_fechaini.setText("")
            self.lineEdit_fechafin.setDisabled(False)
            self.lineEdit_fechafin.setText("")
            self.pushButton_filtro.setDisabled(False)
            self.pushButton_graf_inc.setDisabled(True)
            self.pushButton_graf_err.setDisabled(True)
            self.pushButton_exp_datfilt.setDisabled(True)
        except:
            QMessageBox.about(self, "INFORMACION","Error, conecte la base de datos.")

    def Clear_cam(self):
        self.lineEdit_indice.setText("")
        self.lineEdit_fecha.setText("")
        self.lineEdit_codproducto.setText("")
        self.lineEdit_descripcion.setText("")
        self.comboBox_error.setCurrentIndex(0)
        self.lineEdit_unidades.setText("")
        self.lineEdit_pedidos.setText("")
        self.lineEdit_PRODUCTO.setText("")
        self.lineEdit_DESCRIPCION.setText("")
        self.lineEdit_fechaini.setText("")
        self.lineEdit_fechafin.setText("")
        self.pushButton_filtro.setDisabled(True)
        self.pushButton_graf_inc.setDisabled(True)
        self.pushButton_graf_err.setDisabled(True)

    def Obtener_IDincidencia(self):
        try:
            self.conectar_base2()
            value_ind=int(self.lineEdit_indice.text())
            if value_ind != "":
                self.cursor2.execute("SELECT * FROM INCIDENCIA WHERE ID_REG=?", (value_ind,))
                revolver=self.cursor2.fetchall()
                self.lineEdit_fecha.setDisabled(False)
                self.lineEdit_codproducto.setDisabled(False)
                self.lineEdit_descripcion.setDisabled(False)
                self.comboBox_error.setDisabled(False)
                self.lineEdit_unidades.setDisabled(False)
                self.lineEdit_pedidos.setDisabled(False)

                self.lineEdit_fecha.setText(revolver[0][1])
                per1=str(revolver[0][2])
                self.lineEdit_codproducto.setText(per1)
                self.lineEdit_descripcion.setText(revolver[0][3])
                self.comboBox_error.setCurrentIndex(3)
                per2=str(revolver[0][5])
                self.lineEdit_unidades.setText(per2)
                per3=str(revolver[0][6])
                self.lineEdit_pedidos.setText(per3)
            self.desconectar_base2()
        except:
            QMessageBox.about(self, "INFORMACION","Error, no se ha podido acceder al registro seleccionado.")

    def Obtener_IDProducto(self):
        try:
            self.conectar_base2()
            value_indprod=int(self.lineEdit_PRODUCTO.text())
            if value_indprod != "":
                self.cursor2.execute("SELECT * FROM PRODUCTO WHERE COD_PRODUCTO=?", (value_indprod,))
                revolver_1=self.cursor2.fetchall()
                self.lineEdit_PRODUCTO.setDisabled(False)
                self.lineEdit_DESCRIPCION.setDisabled(False)
                self.lineEdit_DESCRIPCION.setText(revolver_1[0][1])
            self.desconectar_base2()
        except:
            QMessageBox.about(self, "INFORMACION","Error, no se ha podido acceder al registro seleccionado.")

    def ejec_obtener_reg(self):
        try:
            if self.lineEdit_indice.text()!= "":
                self.Obtener_IDincidencia()
            elif self.lineEdit_PRODUCTO.text()!="":
                self.Obtener_IDProducto()
            else:
                QMessageBox.about(self, "INFORMACION","No se ha podido obtener el registro.")
        except:
            QMessageBox.about(self, "INFORMACION","Error, Conecte la base de datos.")

    def ejec_Insertar(self):
        try:
            if self.lineEdit_indice.text()== "" and self.radioButton_insertar.isChecked() and self.radioButton_incidencia.isChecked():
                self.insertar_incidencia()
            elif self.lineEdit_PRODUCTO.text()!="" and self.radioButton_insertar.isChecked() and self.radioButton_producto.isChecked():
                self.insertar_producto()
            else:
                QMessageBox.about(self, "INFORMACION","Error, no se ha podido grabar el registro.")
            self.Clear_cam()
        except:
            QMessageBox.about(self, "INFORMACION","Error, Conecte la base de datos.")

    def ejec_modificar(self):
        try:
            if self.lineEdit_indice.text()!= "":
                self.modificar_incidencia()
            elif self.lineEdit_PRODUCTO.text()!="":
                self.modificar_producto()
            else:
                QMessageBox.about(self, "INFORMACION","Error, no se ha podido modificar el registro.")
            self.Clear_cam()
        except:
            QMessageBox.about(self, "INFORMACION","Conecte la base de datos.")

    def ejec_eliminar(self):
        try:
            if self.lineEdit_indice.text()!= "":
                self.eliminar_incidencia()
            elif self.lineEdit_PRODUCTO.text()!="":
                self.eliminar_producto()
            else:
                QMessageBox.about(self, "INFORMACION","Error, no se ha podido eliminar el registro.")
            self.Clear_cam()
        except:
            QMessageBox.about(self, "INFORMACION","Conecte la base de datos.")

    def filtrado_fec(self):
        try:
            self.conectar_base2()
            inicio=self.lineEdit_fechaini.text()
            final=self.lineEdit_fechafin.text()
            fdt1 = datetime.strptime(inicio, '%d-%m-%Y')
            fdt2 = datetime.strptime(final, '%d-%m-%Y')
            self.dftb_filt = pd.read_sql_query("SELECT * FROM INCIDENCIA", self.conexion2) 
            self.dftb_filt['FECHA'] = pd.to_datetime(self.dftb_filt['FECHA'], format='%d-%m-%Y')

            self.filtered_df =self.dftb_filt.loc[self.dftb_filt["FECHA"].between(fdt1, fdt2)]
            self.tableWidget_errores.clearContents()

            self.lista_incidencia_filt=self.filtered_df.to_numpy().tolist()
            w=len(self.lista_incidencia_filt)   
            for i in range (0, w):
                self.tableWidget_errores.insertRow(i)
                self.tableWidget_errores.setItem(i, 0, QTableWidgetItem(str(self.lista_incidencia_filt[i][0])))
                self.tableWidget_errores.setItem(i, 1, QTableWidgetItem(str(self.lista_incidencia_filt[i][1])))
                self.tableWidget_errores.setItem(i, 2, QTableWidgetItem(str(self.lista_incidencia_filt[i][2])))
                self.tableWidget_errores.setItem(i, 3, QTableWidgetItem(str(self.lista_incidencia_filt[i][3])))
                self.tableWidget_errores.setItem(i, 4, QTableWidgetItem(str(self.lista_incidencia_filt[i][4])))
                self.tableWidget_errores.setItem(i, 5, QTableWidgetItem(str(self.lista_incidencia_filt[i][5])))
                self.tableWidget_errores.setItem(i, 6, QTableWidgetItem(str(self.lista_incidencia_filt[i][6])))
            self.desconectar_base2()

            self.df_graf_desc=self.filtered_df[["DESCRIPCION", "UNIDADES"]]
            self.df_graf_err=self.filtered_df[["ERROR", "UNIDADES"]]
            #self.dib_inc=self.df_graf_desc.groupby(by=["DESCRIPCION"]).sum()
            #self.dib_err=self.df_graf_err.groupby(by=["ERROR"]).sum()
            self.pushButton_exp_reg.setDisabled(True)
            self.pushButton_graf_inc.setDisabled(False)
            self.pushButton_graf_err.setDisabled(False)
            self.pushButton_exp_datfilt.setDisabled(False)
        except:
            QMessageBox.about(self, "INFORMACION","Error en la generación de agrupados.")

    def grafico_incidencia(self):
        try:
            self.df_graf_desc.groupby('DESCRIPCION').sum()["UNIDADES"].plot(kind='bar')
            plot.xticks(rotation=15)
            plot.show()
        except:
            QMessageBox.about(self, "INFORMACION","Error en la generación del gráfico.")

    def grafico_error(self):
        try:
            self.df_graf_err.groupby('ERROR').sum()["UNIDADES"].plot(kind='bar')
            plot.xticks(rotation=15)
            plot.show()  
        except:
            QMessageBox.about(self, "INFORMACION","Error en la generación del gráfico.")

    def exportar_registros(self):
        try:
            fileName_save = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;excel File (*.xlsx)")
            nombre_tabla=fileName_save[0]
            self.dftb_inc.to_excel(nombre_tabla, index=False)
            QMessageBox.about(self, "INFORMACION","Registros exportados correctamente.")
        except:
            QMessageBox.about(self, "INFORMACION","Error en la exportación.")

    def exportar_datosFiltrados(self):
        try:
            fileName_save_filt = QFileDialog.getSaveFileName(self,"QFileDialog.getSaveFileName()","","All Files (*);;excel File (*.xlsx)")
            nombre_tabla_filt=fileName_save_filt[0]
            self.filtered_df.to_excel(nombre_tabla_filt, index=False)
            QMessageBox.about(self, "INFORMACION","Registros exportados correctamente.")
            self.lineEdit_fechaini.setText("")
            self.lineEdit_fechafin.setText("")

            self.pushButton_graf_inc.setDisabled(True)
            self.pushButton_graf_err.setDisabled(True)
            self.pushButton_exp_datfilt.setDisabled(True)
        except:
            QMessageBox.about(self, "INFORMACION","Error en la exportación.")


    # NUMERO 2/2

    def resolver_ruta2(self,ruta_relativa2):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, ruta_relativa2)
        return os.path.join(os.path.abspath('.'), ruta_relativa2)
    

if __name__ == "__main__":
    app=QApplication(sys.argv)
    QApplication.setAttribute(Qt.AA_DisableWindowContextHelpButton)  
    GUIde=Consulta_incidencias()
    #GUIa.setModal(True) #inhbilita la ventana principal
    GUIde.show()
    sys.exit(app.exec_())

                                    