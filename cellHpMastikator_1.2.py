# -*- coding: utf-8 -*-
"""
Created on Fri May 31 11:50:54 2019
Author: Pablo La Grutta
mail: pablo.lg.unlam@gmail.com
"""

import tkinter as tk
from tkinter import ttk,scrolledtext as st
from tkinter.ttk import Style
from tkinter.filedialog import askopenfilename, askdirectory
import datetime
import os
import pandas as pd
import sys

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
    print('application_path de Py', application_path)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
    

class GUI(tk.Tk):
    
    def __init__(self, window): 
                # parameters that you want to send through the Frame class. 
        self.path_resultados = ''
        self.input_text_resultados= ''
        etiq_bt1 = "Archivo EQDB3-Celda-HP"
        etiq_bt2 = "Resultados"
        etiq_bt3 = "Descargar en..."
        window.title("EQDB3 Mastikator 1.2")
        window.config(bg='#321899')
        window.resizable(0, 0) 
        window.geometry("800x300")  #ancho largo
        style = Style() 
        style.configure('W.TButton', font =('calibri', 10, 'bold', 'underline'), background = "orange", foreground = 'blue')               
        ttk.Button(window, text = etiq_bt1,style = 'W.TButton', command = lambda: self.set_path_cellHp()).grid(row = 0) 
        self.scrolledtext0=st.ScrolledText(window, width=80, height=2)
        self.scrolledtext0.grid(column=1,row=0, padx=0, pady=10)
        ttk.Button(window, text = etiq_bt3,style = 'W.TButton', command = lambda: self.set_path_resultados()).grid(row = 2) 
        self.scrolledtext1=st.ScrolledText(window, width=80, height=2)
        self.scrolledtext1.grid(column=1,row=2, padx=0, pady=10)      
        ttk.Button(window, text = etiq_bt2, style = 'W.TButton',command = lambda: self.Proceso1()).grid(row = 3) 
    def Proceso1(self):  
        ##--diccionario cosmico
        dfDict = pd.DataFrame({'Región':["SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","SUR","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","Norte Lito","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","AMBA","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi","Norte Medi"],

                               'celdaXYZ':["A","B","C","R","S","T","D","E","F","U","V","W","G","H","I","X","Y","Z","J","K","L","M","N","O","U11","U12","U13","U14","U15","U16","U21","U22","U23","U24","U25","U26","U31","U32","U33","U34","U35","U36","U41","U42","U43","U44","U45","U46","W11","W12","W13","A","B","C","D","E","F","M","N","O","P","Q","R","S","T","U","V","W","X","G","H","I","J","K","L","V11","V12","V13","V14","V15","V16","U11","U12","U13","U14","U15","U16","U21","U22","U23","U24","U25","U26","V21","V22","V23","V24","V25","V26","A","B","C","D","E","F","A2","B2","C2","D2","E2","F2","A3","B3","C3","D3","E3","F3","A4","B4","C4","D4","E4","F4","U11","U12","U13","U14","U15","U16","U21","U22","U23","U24","U25","U26","V11","V12","V13","V14","V15","V16","V21","V22","V23","V24","V25","V26","A","B","C","D","E","F","M","N","O","P","Q","R","S","T","U","V","W","X","G","H","I","J","K","L","V11","V12","V13","V14","V15","V16","U11","U12","U13","U14","U15","U16","U21","U22","U23","U24","U25","U26","V21","V22","V23","V24","V25","V26"],
                               'Banda':[1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,900,900,900,850,850,850,850,850,850,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,850,850,850,850,850,850,850,850,850,850,850,850,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,850,850,850,850,850,850,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,850,850,850,850,850,850,850,850,850,850,850,850,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,850,850,850,850,850,850,850,850,850,850,850,850,850,850,850,850,850,850,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,850,850,850,850,850,850,850,850,850,850,850,850,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,1900,850,850,850,850,850,850],
                               'Portadora':['1','1','1','1','1','1','2','2','2','2','2','2','3','3','3','3','3','3','4','4','4','4','4','4','1','1','1','1','1','1','2','2','2','2','2','2','3','3','3','3','3','3','4','4','4','4','4','4','1','1','1','1','1','1','1','1','1','1','1','1','1','1','1','2','2','2','2','2','2','2','2','2','2','2','2','1','1','1','1','1','1','1','1','1','1','1','1','2','2','2','2','2','2','2','2','2','2','2','2','1','1','1','1','1','1','2','2','2','2','2','2','1','1','1','1','1','1','2','2','2','2','2','2','1','1','1','1','1','1','2','2','2','2','2','2','1','1','1','1','1','1','2','2','2','2','2','2','1','1','1','1','1','1','1','1','1','1','1','1','2','2','2','2','2','2','2','2','2','2','2','2','1','1','1','1','1','1','1','1','1','1','1','1','2','2','2','2','2','2','2','2','2','2','2','2'],
                               'Sector': ['1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3','1','2','3']
                                   })
        self.df_cellHp['Sitio'] = [x[0:6] if (len(x)==9 or len(x)==7) 
                                                   else x for x in self.df_cellHp['Celda']]   #-=SI(LARGO(Celda)=9;Celda[0:6]);A2)

        self.df_cellHp['celdaXYZ'] = [ (str(x[-3:]) if (len(x)==9  or len(x)==5) 
                                                else (str(x[-1:]) if len(x)==7 
                                                      else(str(x[-2:]) if len(x)==6 
                                                           else str(x)))) for x in self.df_cellHp['Celda']]

        self.df_cellHp =  self.df_cellHp[['Región','Celda','Sitio','RNC','celdaXYZ','Banda',"Tráfico Erl Voz", "Usuarios HSDPA"]]
        self.df_cellHp = pd.merge(dfDict,self.df_cellHp,on=['Región','Banda','celdaXYZ'])

        self.df_cellHp['sitioSector'] = self.df_cellHp['Sitio'] + self.df_cellHp['Sector'].astype(str)
        self.df_cellHp[["Tráfico Erl Voz", "Usuarios HSDPA"]] = self.df_cellHp[["Tráfico Erl Voz", "Usuarios HSDPA"]].apply(pd.to_numeric)
        self.df_cellHp.set_index('sitioSector',inplace=True)
        self.df_pivot_bp = self.df_cellHp.groupby(['sitioSector','Banda','Portadora']).sum()  ##--agregación por sitioSector y Banda
        self.df_pivot_b = self.df_cellHp.groupby(['sitioSector','Banda']).sum()
        self.df_pivot_total = self.df_pivot_b.groupby(['sitioSector']).sum()
        self.df_pivot_total['Potencia']= 10.47 + self.df_pivot_total['Tráfico Erl Voz']*0.27 + self.df_pivot_total['Usuarios HSDPA']*0.26
        ct = str(datetime.datetime.now() ).replace(' ','_').replace(':','').replace('.','')[0:17]

        with pd.ExcelWriter(self.path_res + '/'+ self.path_cellHpSep[-1][0:-4] + '_Resultado_' + ct + '.xlsx') as writer:  # el argumento de ExcelWriter es el path al archivo, ahi tengo que cargar el path deseado
            self.df_pivot_bp.to_excel(writer,sheet_name='Banda&Portadora',index=True)
            self.df_pivot_b.to_excel(writer,sheet_name='Banda',index=True)
            self.df_pivot_total.to_excel(writer,sheet_name='Total',index=True) 

        gui.__init__(window)

    def set_path_cellHp(self):
        self.path_cellHp = askopenfilename( filetypes = [("Archivo EQDB3","*.csv"),("Todos los archivos","*.*")], title = "Seleccionar archivo CSV")
        self.path_cellHpSep = str(self.path_cellHp).split('/')
        self.scrolledtext0.insert("1.0", self.path_cellHpSep[-1])
        self.df_cellHp= pd.read_csv(self.path_cellHp,  sep=';' ,decimal=',' , engine='python') #  , utf-16, mbcs,chardet, utf-8, chardet

    def set_path_resultados(self):
        self.path_res = askdirectory( )
        self.scrolledtext1.insert("1.0",self.path_res)


if __name__ == '__main__':
    window = tk.Tk()
    gui = GUI(window)
    window.mainloop()
 





       

    
    
