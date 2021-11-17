# -*- coding: utf-8 -*-
from builtins import str
# from qgis.core import QgsProject, QgsMessageLog

# from PyQt5.QtCore import *
# from PyQt5 import QtCore 

# from PyQt5.QtGui import QDesktopServices
# from PyQt5 import QtGui #Paquetes requeridos para crear ventanas de diálogo e interfaz gráfica.
# from PyQt5 import QtWidgets
# from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog, QMessageBox
# import traceback

# from qgis.core import *  # Paquetes requeridos para crer el registro de eventos.
# from qgis.gui import *  # Paquete requerido para desplegar mensajes en la ventana principal de QGIS.

def renameVoltage(MVCode,LVCode): #Para monofasicos, tension L-N, para trifasicos tension L-L
        MV={'LN':0,'LL':0}
        LV={'LN':0,'LL':0}
        if MVCode==270:
            MV={'LN':7.967,'LL':13.8}
        if MVCode==380:
            MV= {'LN':19.92,'LL':34.5}
        if MVCode==230:
            MV= {'LN':7.62,'LL':13.2}
        if MVCode==340:
            MV= {'LN':14.38,'LL':24.9}
        if LVCode ==20:
            LV = {'LN':0.12,'LL':0.208}
        if LVCode ==30:
            LV = {'LN':0.12,'LL':0.24}
        if LVCode ==35:
            LV = {'LN':0.254,'LL':0.44}
        if LVCode ==40:
            LV = {'LN':0.240,'LL':0.48}
        if LVCode ==50:
            LV = {'LN':0.277,'LL':0.48}
        if LVCode ==60:
            LV = {'LN':0.48,'LL':0.48} #Conexión delta 3 conductores
        if LVCode ==70:
            LV = {'LN':0.24,'LL':0.416}
        voltageCodes={'MVCode':MV,'LVCode':LV}
        return voltageCodes

react_list1F3W={'13.8':{      # List of reactances for single phase three winding transformers
                    '0.12': {
                            '3': 'Xhl=0.82 Xht=0.82 Xlt=0.55',
                            '5': 'Xhl=0.82 Xht=0.82 Xlt=0.55',
                            '10':'Xhl=1.54 Xht=1.54 Xlt=1.03',
                            '15':'Xhl=1.19 Xht=1.19 Xlt=0.80',
                            '25':'Xhl=1.49 Xht=1.49 Xlt=1.00',
                            '37':'Xhl=1.71 Xht=1.71 Xlt=1.14',
                            '50':'Xhl=1.86 Xht=1.86 Xlt=1.24',
                            '75':'Xhl=1.80 Xht=1.80 Xlt=1.20',
                            '100':'Xhl=2.14 Xht=2.14 Xlt=1.43',   
                            '167':'Xhl=2.35 Xht=2.35 Xlt=1.57', 
                            '250':'Xhl=3.17 Xht=3.17 Xlt=2.11',
                            '333':'Xhl=3.52 Xht=3.52 Xlt=2.35',
                            '500':'Xhl=3.61 Xht=3.61 Xlt=2.40'},
                    '0.24':{
                            '3':  'Xhl=1.31 Xht=1.31 Xlt=0.87',
                            '5':  'Xhl=1.31 Xht=1.31 Xlt=0.87',
                            '10': 'Xhl=1.40 Xht=1.40 Xlt=0.94',
                            '15': 'Xhl=1.19 Xht=1.19 Xlt=0.79',
                            '25': 'Xhl=1.49 Xht=1.49 Xlt=0.99',
                            '37': 'Xhl=1.70 Xht=1.70 Xlt=1.14',
                            '50': 'Xhl=1.86 Xht=1.86 Xlt=1.24',
                            '75': 'Xhl=1.80 Xht=1.80 Xlt=1.20',
                            '100':'Xhl=1.94 Xht=1.94 Xlt=1.30',
                            '167':'Xhl=2.15 Xht=2.15 Xlt=1.43',
                            '250':'Xhl=2.69 Xht=2.69 Xlt=1.79',
                            '333':'Xhl=2.66 Xht=2.66 Xlt=1.78',
                            '500':'Xhl=2.76 Xht=2.76 Xlt=1.84'}},
              '24.9':{
                    '0.12': {
                            '3': 'Xhl=0.82 Xht=0.82 Xlt=0.55',
                            '5': 'Xhl=0.82 Xht=0.82 Xlt=0.55',
                            '10':'Xhl=1.54 Xht=1.54 Xlt=1.03',
                            '15':'Xhl=1.19 Xht=1.19 Xlt=0.80',
                            '25':'Xhl=1.49 Xht=1.49 Xlt=1.00',
                            '37':'Xhl=1.71 Xht=1.71 Xlt=1.14',
                            '50':'Xhl=1.86 Xht=1.86 Xlt=1.24',
                            '75':'Xhl=1.80 Xht=1.80 Xlt=1.20',
                            '100':'Xhl=2.14 Xht=2.14 Xlt=1.43',   
                            '167':'Xhl=2.35 Xht=2.35 Xlt=1.57', 
                            '250':'Xhl=3.17 Xht=3.17 Xlt=2.11',
                            '333':'Xhl=3.52 Xht=3.52 Xlt=2.35',
                            '500':'Xhl=3.61 Xht=3.61 Xlt=2.40'},
                    '0.24':{
                            '3':  'Xhl=1.31 Xht=1.31 Xlt=0.87',
                            '5':  'Xhl=1.31 Xht=1.31 Xlt=0.87',
                            '10': 'Xhl=1.4 Xht=1.4 Xlt=0.94',
                            '15': 'Xhl=1.19 Xht=1.19 Xlt=0.79',
                            '25': 'Xhl=1.49 Xht=1.49 Xlt=0.99',
                            '37': 'Xhl=1.7 Xht=1.7 Xlt=1.14',
                            '50': 'Xhl=1.86 Xht=1.86 Xlt=1.24',
                            '75': 'Xhl=1.8 Xht=1.8 Xlt=1.2',
                            '100':'Xhl=1.94 Xht=1.94 Xlt=1.3',
                            '167':'Xhl=2.15 Xht=2.15 Xlt=1.43',
                            '250':'Xhl=2.69 Xht=2.69 Xlt=1.79',
                            '333':'Xhl=2.66 Xht=2.66 Xlt=1.78',
                            '500':'Xhl=2.76 Xht=2.76 Xlt=1.84'}},
              '34.5':{
                    '0.12': {
                            '3':  'Xhl=4.92 Xht=4.92 Xlt=3.28',
                            '5': 'Xhl=4.92 Xht=4.92 Xlt=3.28', 
                            '10':'Xhl=5.34 Xht=5.34 Xlt=3.56', 
                            '15':'Xhl=5.52 Xht=5.52 Xlt=3.68',    
                            '25':'Xhl=5.65 Xht=5.65 Xlt=3.77',
                            '37':'Xhl=5.80 Xht=5.80 Xlt=3.87',   
                            '50':'Xhl=5.90 Xht=5.90 Xlt=3.93',
                            '75':'Xhl=5.94 Xht=5.94 Xlt=3.96',    
                            '100':'Xhl=5.97 Xht=5.97 Xlt=3.98',
                            '167':'Xhl=6.04 Xht=6.04 Xlt=4.03', 
                            '250':'Xhl=6.08 Xht=6.08 Xlt=4.05', 
                            '333':'Xhl=6.10 Xht=6.10 Xlt=4.07',   
                            '500':'Xhl=6.12 Xht=6.12 Xlt=4.08'},
                    '0.24':{
                            '3':  'Xhl=4.92 Xht=4.92 Xlt=3.28',
                            '5':  'Xhl=4.92 Xht=4.92 Xlt=3.28',
                            '10': 'Xhl=5.34 Xht=5.34 Xlt=3.56',
                            '15': 'Xhl=5.52 Xht=5.52 Xlt=3.68',
                            '25': 'Xhl=5.65 Xht=5.65 Xlt=3.77',
                            '37': 'Xhl=5.8 Xht=5.8 Xlt=3.86',
                            '50': 'Xhl=5.89 Xht=5.89 Xlt=3.93',
                            '75': 'Xhl=5.94 Xht=5.94 Xlt=3.96',
                            '100':'Xhl=5.98 Xht=5.98 Xlt=3.98',
                            '167':'Xhl=6.04 Xht=6.04 Xlt=4.02',
                            '250':'Xhl=6.07 Xht=6.07 Xlt=4.05',
                            '333':'Xhl=6.1 Xht=6.1 Xlt=4.06',
                            '500':'Xhl=6.12 Xht=6.12 Xlt=4.08'}}}
res_list1F3W={'13.8':{    # List of resistances for single phase three winding transformers
                    '0.12': {
                            '3': '%Rs=[1.15 2.30 2.30]',
                            '5': '%Rs=[1.15 2.30 2.30]',
                            '10':'%Rs=[0.70 1.40 1.40]', 
                            '15':'%Rs=[0.75 1.50 1.50]',
                            '25':'%Rs=[0.65 1.30 1.30]',
                            '37':'%Rs=[0.55 1.10 1.10]',
                            '50':'%Rs=[0.55 1.10 1.10] ',
                            '75':'%Rs=[0.50 1.00 1.00]',
                            '100':'%Rs=[0.45 0.90 0.90]',                
                            '167':'%Rs=[0.50 1.00 1.00]', 
                            '250':'%Rs=[0.60 1.20 1.20]', 
                            '333':'%Rs=[0.50 1.00 1.00]',
                            '500':'%Rs=[0.55 1.10 1.10]'},
                    '0.24':{
                            '3':  '%Rs=[0.74 1.49 1.49]',
                            '5':  '%Rs=[0.74 1.49 1.49]',
                            '10': '%Rs=[0.75 1.50 1.50]',
                            '15': '%Rs=[0.75 1.5 1.5]',
                            '25': '%Rs=[0.65 1.3 1.3]',
                            '37': '%Rs=[0.55 1.1 1.1]',
                            '50': '%Rs=[0.55 1.1 1.1]',
                            '75': '%Rs=[0.5 1.0 1.0]',
                            '100':'%Rs=[0.5 1.0 1.0]',
                            '167':'%Rs=[0.45 0.9 0.9]',
                            '250':'%Rs=[0.55 1.1 1.1]',
                            '333':'%Rs=[0.45 0.9 0.9]',
                            '500':'%Rs=[0.35 0.7 0.7]'}},
                            
            '24.9':{
                    '0.12': {
                            '3': '%Rs=[1.15 2.30 2.30]',
                            '5': '%Rs=[1.15 2.30 2.30]',
                            '10':'%Rs=[0.70 1.40 1.40]', 
                            '15':'%Rs=[0.75 1.50 1.50]',
                            '25':'%Rs=[0.65 1.30 1.30]',
                            '37':'%Rs=[0.55 1.10 1.10]',
                            '50':'%Rs=[0.55 1.10 1.10] ',
                            '75':'%Rs=[0.50 1.00 1.00]',
                            '100':'%Rs=[0.45 0.90 0.90]',                
                            '167':'%Rs=[0.50 1.00 1.00]', 
                            '250':'%Rs=[0.60 1.20 1.20]', 
                            '333':'%Rs=[0.50 1.00 1.00]',
                            '500':'%Rs=[0.55 1.10 1.10]'},
                    '0.24':{
                            '3':  '%Rs=[0.74 1.49 1.49]',
                            '5':  '%Rs=[0.74 1.49 1.49]',
                            '10': '%Rs=[0.75 1.5 1.5]',
                            '15': '%Rs=[0.75 1.5 1.5]',
                            '25': '%Rs=[0.65 1.3 1.3]',
                            '37': '%Rs=[0.55 1.1 1.1]',
                            '50': '%Rs=[0.55 1.1 1.1]',
                            '75': '%Rs=[0.5 1.0 1.0]',
                            '100':'%Rs=[0.5 1.0 1.0]',
                            '167':'%Rs=[0.45 0.9 0.9]',
                            '250':'%Rs=[0.55 1.1 1.1]',
                            '333':'%Rs=[0.45 0.9 0.9]',
                            '500':'%Rs=[0.35 0.7 0.7]'}},
            '34.5':{
                    '0.12': {
                            '3': '%Rs=[1.60 3.20 3.20]',
                            '5': '%Rs=[1.60 3.20 3.20]',
                            '10':'%Rs=[1.34 2.68 2.68]', 
                            '15':'%Rs=[1.21 2.42 2.42]',
                            '25':'%Rs=[1.10 2.20 2.20]',
                            '37':'%Rs=[0.96 1.92 1.92]',
                            '50':'%Rs=[0.85 1.70 1.70]',
                            '75':'%Rs=[0.80 1.61 1.61]',
                            '100':'%Rs=[0.75 1.50 1.50]',
                            '167':'%Rs=[0.65 1.31 1.31]', 
                            '250':'%Rs=[0.59 1.18 1.18]', 
                            '333':'%Rs=[0.55 1.10 1.10]',
                            '500':'%Rs=[0.50 1.00 1.00]'},
                    '0.24':{
                            '3':  '%Rs=[1.6 3.2 3.2]',
                            '5':  '%Rs=[1.6 3.2 3.2]',
                            '10': '%Rs=[1.34 2.68 2.68]',
                            '15': '%Rs=[1.21 2.42 2.42]',
                            '25': '%Rs=[1.1 2.2 2.2]',
                            '37': '%Rs=[0.96 1.92 1.92]',
                            '50': '%Rs=[0.85 1.7 1.7]',
                            '75': '%Rs=[0.81 1.61 1.61]',
                            '100':'%Rs=[0.75 1.5 1.5]',
                            '167':'%Rs=[0.66 1.31 1.31]',
                            '250':'%Rs=[0.59 1.18 1.18]',
                            '333':'%Rs=[0.55 1.1 1.1]',
                            '500':'%Rs=[0.5 1.0 1.0]'}}}

imp_list1F2W={'13.8':{# List of impedances for single phase two winding transformers [MV][LV][Power]
                    '0.12': {
                            '3':  'Xhl=0.69 %Rs=[1.15 1.15]',
                            '5':  'Xhl=0.69 %Rs=[1.15 1.15]',
                            '10': 'Xhl=1.28 %Rs=[0.70 0.70]',
                            '15': 'Xhl=0.99 %Rs=[0.75 0.75]',
                            '25': 'Xhl=1.24 %Rs=[0.65 0.65]',
                            '37': 'Xhl=1.42 %Rs=[0.55 0.55]',
                            '50': 'Xhl=1.55 %Rs=[0.55 0.55]',
                            '75': 'Xhl=1.50 %Rs=[0.50 0.50]',
                            '100':'Xhl=1.79 %Rs=[0.45 0.45]',
                            '167':'Xhl=1.96 %Rs=[0.50 0.50]', 
                            '250':'Xhl=2.64 %Rs=[0.60 0.60]',
                            '333':'Xhl=2.93 %Rs=[0.50 0.50]',
                            '500':'Xhl=3.00 %Rs=[0.55 0.55]'},
                    '0.24':{
                            '3':  'Xhl=1.09 %Rs=[0.74 0.74]',
                            '5':  'Xhl=1.09 %Rs=[0.74 0.74]',
                            '10': 'Xhl=1.17 %Rs=[0.75 0.75]',
                            '15': 'Xhl=0.99 %Rs=[0.75 0.75]',
                            '25': 'Xhl=1.24 %Rs=[0.65 0.65]',
                            '37': 'Xhl=1.42 %Rs=[0.55 0.55]',
                            '50': 'Xhl=1.55 %Rs=[0.55 0.55]',
                            '75': 'Xhl=1.50 %Rs=[0.50 0.50]',
                            '100':'Xhl=1.62 %Rs=[0.50 0.50]',
                            '167':'Xhl=1.79 %Rs=[0.45 0.45]',
                            '250':'Xhl=2.24 %Rs=[0.55 0.55]',
                            '333':'Xhl=2.22 %Rs=[0.45 0.45]',
                            '500':'Xhl=2.3 %Rs=[0.35 0.35]'}},
                            
              '24.9':{
                    '0.12': {
                            '3':  'Xhl=0.69 %Rs=[1.15 1.15]',
                            '5':  'Xhl=0.69 %Rs=[1.15 1.15]',
                            '10': 'Xhl=1.28 %Rs=[0.70 0.70]',
                            '15': 'Xhl=0.99 %Rs=[0.75 0.75]',
                            '25': 'Xhl=1.24 %Rs=[0.65 0.65]',
                            '37': 'Xhl=1.42 %Rs=[0.55 0.55]',
                            '50': 'Xhl=1.55 %Rs=[0.55 0.55]',
                            '75': 'Xhl=1.50 %Rs=[0.50 0.50]',
                            '100':'Xhl=1.79 %Rs=[0.45 0.45]',   
                            '167':'Xhl=1.96 %Rs=[0.50 0.50]', 
                            '250':'Xhl=2.64 %Rs=[0.60 0.60]',
                            '333':'Xhl=2.93 %Rs=[0.50 0.50]',
                            '500':'Xhl=3.00 %Rs=[0.55 0.55]'},
                    '0.24':{
                            '3':  'Xhl=1.09 %Rs=[0.74 0.74]',
                            '5':  'Xhl=1.09 %Rs=[0.74 0.74]',
                            '10': 'Xhl=1.17 %Rs=[0.75 0.75]',
                            '15': 'Xhl=0.99 %Rs=[0.75 0.75]',
                            '25': 'Xhl=1.24 %Rs=[0.65 0.65]',
                            '37': 'Xhl=1.42 %Rs=[0.55 0.55]',
                            '50': 'Xhl=1.55 %Rs=[0.55 0.55]',
                            '75': 'Xhl=1.50 %Rs=[0.50 0.50]',
                            '100':'Xhl=1.62 %Rs=[0.50 0.50]',
                            '167':'Xhl=1.79 %Rs=[0.45 0.45]',
                            '250':'Xhl=2.24 %Rs=[0.55 0.55]',
                            '333':'Xhl=2.22 %Rs=[0.45 0.45]',
                            '500':'Xhl=2.3 %Rs=[0.35 0.35]'}},
              '34.5':{
                    '0.12': {
                            '3':  'Xhl=4.10 %Rs=[1.601 1.601]', 
                            '5':  'Xhl=4.10 %Rs=[1.601 1.601]', 
                            '10': 'Xhl=4.45 %Rs=[1.342 1.342]', 
                            '15': 'Xhl=4.60 %Rs=[1.21 1.21]',    
                            '25': 'Xhl=4.71 %Rs=[1.10 1.10]',
                            '37': 'Xhl=4.83 %Rs=[0.958 0.958]',   
                            '50': 'Xhl=4.91 %Rs=[0.85 0.85]',
                            '75': 'Xhl=4.95 %Rs=[0.803 0.803]',    
                            '100':'Xhl=4.98 %Rs=[0.75 0.75]',
                            '167':'Xhl=5.03 %Rs=[0.654 0.654]', 
                            '250':'Xhl=5.06 %Rs=[0.591 0.591]', 
                            '333':'Xhl=5.08 %Rs=[0.55 0.55]',   
                            '500':'Xhl=5.10 %Rs=[0.50 0.50]'},
                    '0.24':{
                            '3':  'Xhl=4.10 %Rs=[1.60 1.60]',
                            '5':  'Xhl=4.10 %Rs=[1.60 1.60]',
                            '10': 'Xhl=4.45 %Rs=[1.34 1.34]',
                            '15': 'Xhl=4.60 %Rs=[1.21 1.21]',
                            '25': 'Xhl=4.71 %Rs=[1.10 1.10]',
                            '37': 'Xhl=4.83 %Rs=[0.96 0.96]',
                            '50': 'Xhl=4.91 %Rs=[0.85 0.85]',
                            '75': 'Xhl=4.95 %Rs=[0.81 0.81]',
                            '100':'Xhl=4.98 %Rs=[0.75 0.75]',
                            '167':'Xhl=5.03 %Rs=[0.66 0.66]',
                            '250':'Xhl=5.06 %Rs=[0.59 0.59]',
                            '333':'Xhl=5.08 %Rs=[0.55 0.55]',
                            '500':'Xhl=5.10 %Rs=[0.50 0.50]'}}}
				
noloadloss_list1F={ 
                    '3': '%noloadloss=0.84',
                    '5': '%noloadloss=0.84',
                    '10':'%noloadloss=0.73',
                    '15':'%noloadloss=0.56',
                    '25':'%noloadloss=0.47',
                    '37':'%noloadloss=0.44',
                    '50':'%noloadloss=0.37',
                    '75':'%noloadloss=0.38',
                    '100':'%noloadloss=0.36',   
                    '167':'%noloadloss=0.30', 
                    '250':'%noloadloss=0.24',
                    '333':'%noloadloss=0.25',
                    '500':'%noloadloss=0.23'}
               				
imag_list1F={ 
              '3':  '%imag=2.40',
              '5':  '%imag=2.40',
              '10': '%imag=1.60',
              '15': '%imag=1.40',
              '25': '%imag=1.30',
              '37': '%imag=1.10',
              '50': '%imag=1.00',
              '75': '%imag=1.40',
              '100':'%imag=1.30',   
              '167':'%imag=1.00', 
              '250':'%imag=1.00',
              '333':'%imag=1.00',
              '500':'%imag=1.00'}				
							
imp_list3F={'13.8':{
                '9':'Xhl=2.84 %Rs=[1.316 1.316]', 
                '30':'Xhl=2.28 %Rs=[1.05 1.05]',
                '45':'Xhl=2.08 %Rs=[0.868 0.868]',
                '75':'Xhl=1.65 %Rs=[0.65 0.65]',
                '112':'Xhl=1.30 %Rs=[0.55 0.55]',
                '150':'Xhl=1.55 %Rs=[0.55 0.55]',
                '225':'Xhl=1.55 %Rs=[0.55 0.55]', 
                '300':'Xhl=1.67 %Rs=[0.55 0.55]',
                '500':'Xhl=2.07 %Rs=[0.50 0.50]',
                '750':'Xhl=5.59 %Rs=[0.55 0.55]',
                '1000':'Xhl=5.61 %Rs=[0.50 0.50]',  
                '1500':'Xhl=5.59 %Rs=[0.55 0.55]',
                '2500':'Xhl=5.59 %Rs=[0.55 0.55]',
                '3000':'Xhl=5.59 %Rs=[0.55 0.55]',				
                '3750':'Xhl=5.59 %Rs=[0.55 0.55]',
                '5000':'Xhl=5.59 %Rs=[0.53 0.53]',
                '7500':'Xhl=5.59 %Rs=[0.52 0.52]',
                '10000':'Xhl=5.59 %Rs=[0.51 0.51]',
                '15000':'Xhl=5.59 %Rs=[0.49 0.49]',
                '20000':'Xhl=5.59 %Rs=[0.47 0.47]',
                '25000':'Xhl=5.59 %Rs=[0.45 0.45]',
                '30000':'Xhl=5.59 %Rs=[0.43 0.43]',
                '40000':'Xhl=5.59 %Rs=[0.39 0.39]',
                '50000':'Xhl=5.59 %Rs=[0.35 0.35]'
                },
            '24.9':{
                '9':'Xhl=2.84 %Rs=[1.316 1.316]', 
                '30':'Xhl=2.28 %Rs=[1.05 1.05]',
                '45':'Xhl=2.08 %Rs=[0.868 0.868]',
                '75':'Xhl=1.65 %Rs=[0.65 0.65]',
                '112':'Xhl=1.30 %Rs=[0.55 0.55]',
                '150':'Xhl=1.55 %Rs=[0.55 0.55]',
                '225':'Xhl=1.55 %Rs=[0.55 0.55]', 
                '300':'Xhl=1.67 %Rs=[0.55 0.55]',
                '500':'Xhl=2.07 %Rs=[0.50 0.50]',
                '750':'Xhl=5.59 %Rs=[0.55 0.55]',
                '1000':'Xhl=5.49 %Rs=[0.455 0.455]',  
                '1500':'Xhl=5.50 %Rs=[0.40 0.40]',
                '2500':'Xhl=5.51 %Rs=[0.36 0.36]',				
                '3000':'Xhl=5.51 %Rs=[0.36 0.36]',				
                '3750':'Xhl=4.95 %Rs=[0.345 0.345]',
                '5000':'Xhl=5.62 %Rs=[0.30 0.30]',
                '7500':'Xhl=5.66 %Rs=[0.27 0.27]',
                '10000':'Xhl=5.70 %Rs=[0.24 0.24]',
                '15000':'Xhl=5.75 %Rs=[0.22 0.22]',
                '20000':'Xhl=5.78 %Rs=[0.20 0.20]',
                '25000':'Xhl=5.81 %Rs=[0.19 0.19]',
                '30000':'Xhl=5.86 %Rs=[0.16 0.16]',
                '40000':'Xhl=5.94 %Rs=[0.12 0.12]',
                '50000':'Xhl=6.03 %Rs=[0.08 0.08]'
                },				
            '34.5':{
                '9':'Xhl=3.96 %Rs=[1.91 1.91]', 
                '30':'Xhl=4.80 %Rs=[1.339 1.339]',
                '45':'Xhl=4.96 %Rs=[1.188 1.188]',
                '75':'Xhl=5.11 %Rs=[1.022 1.022]',
                '112':'Xhl=5.17 %Rs=[0.939 0.939]',
                '150':'Xhl=5.24 %Rs=[0.833 0.833]',
                '225':'Xhl=5.30 %Rs=[0.739 0.739]', 
                '300':'Xhl=5.32 %Rs=[0.70 0.70]',
                '500':'Xhl=5.37 %Rs=[0.60 0.60]',
                '750':'Xhl=5.40 %Rs=[0.518 0.518]',
                '1000':'Xhl=5.49 %Rs=[0.455 0.455]',  
                '1500':'Xhl=5.50 %Rs=[0.40 0.40]',
                '2500':'Xhl=5.51 %Rs=[0.36 0.36]',				
                '3000':'Xhl=5.51 %Rs=[0.36 0.36]',				
                '3750':'Xhl=4.95 %Rs=[0.345 0.345]',
                '5000':'Xhl=5.62 %Rs=[0.30 0.30]',
                '7500':'Xhl=5.66 %Rs=[0.27 0.27]',
                '10000':'Xhl=5.70 %Rs=[0.24 0.24]',
                '15000':'Xhl=5.75 %Rs=[0.22 0.22]',
                '20000':'Xhl=5.78 %Rs=[0.20 0.20]',
                '25000':'Xhl=5.81 %Rs=[0.19 0.19]',
                '30000':'Xhl=5.86 %Rs=[0.16 0.16]',
                '40000':'Xhl=5.94 %Rs=[0.12 0.12]',
                '50000':'Xhl=6.03 %Rs=[0.08 0.08]'
                }}
        
noloadloss_list3F={ '9': '%noloadloss=1.06',
                    '30':'%noloadloss=0.71',
                    '45':'%noloadloss=0.62',
                    '75':'%noloadloss=0.48',
                    '112':'%noloadloss=0.53',
                    '150':'%noloadloss=0.37',   
                    '225':'%noloadloss=0.39', 
                    '300':'%noloadloss=0.35',
                    '500':'%noloadloss=0.32', 
                    '750':'%noloadloss=0.24',
                    '1000':'%noloadloss=0.21',  
                    '1500':'%noloadloss=0.19',
                    '2500':'%noloadloss=0.17',				
                    '3000':'%noloadloss=0.17',				
                    '3750':'%noloadloss=0.15',
                    '5000':'%noloadloss=0.14',
                    '7500':'%noloadloss=0.12',
                    '10000':'%noloadloss=0.11',
                    '15000':'%noloadloss=0.10',
                    '20000':'%noloadloss=0.09',
                    '25000':'%noloadloss=0.08',
                    '30000':'%noloadloss=0.07',
                    '40000':'%noloadloss=0.05',
                    '50000':'%noloadloss=0.03'}
               				
imag_list3F={ '9':  '%imag=2.48',
             '30':  '%imag=1.72',
             '45':  '%imag=1.52',
             '75':  '%imag=1.50',
             '112': '%imag=1.00',
             '150': '%imag=1.00',   
             '225': '%imag=1.00', 
             '300': '%imag=1.00',
             '500': '%imag=1.00', 
             '750': '%imag=1.00',
             '1000':'%imag=1.00',  
             '1500':'%imag=1.00',
             '2500':'%imag=1.00',				
             '3000':'%imag=1.00',				
             '3750':'%imag=1.00',
             '5000': '%imag=1.00', 
             '7500': '%imag=1.00',
             '10000':'%imag=1.00',  
             '15000':'%imag=1.00',
             '20000':'%imag=1.00',
             '25000':'%imag=1.00',
             '30000':'%imag=1.00',
             '40000':'%imag=1.00',
             '50000':'%imag=1.00'}			
		
		
		
def impedanceSingleUnit(phase,voltageMV,voltageLV, rated_power): 
    phase=str(phase)
    voltageMV = str(voltageMV)
    if voltageMV == "13.2":
        voltageMV="13.8"
    if voltageLV == "0.277":
        voltageLV="0.24"
    voltageLV = str(voltageLV)
    power = str(int(float(rated_power)))
    
    if phase =='3':
        if (power in list(imp_list3F[voltageMV].keys())):
            imp_out = imp_list3F[voltageMV][power]
            imag_out = imag_list3F[power]
            noloadloss_out=noloadloss_list3F[power]
            imp = {'Z':imp_out,'Im':imag_out,'Pnoload':noloadloss_out}
        else:
            aviso= u'No existen valores de reactancia y resistencia para transformadores trifásicos de '+power+' kVA operando en '+voltageMV+' kV.'
            # QgsMessageLog.logMessage(aviso,u'Transformadores',Qgis.Warning)
            imp_out = ' %R=UNKNOWN Xhl=UNKNOWN '
            imp = {'Z':imp_out,'Im':'%imag=1.00','Pnoload':'%noloadloss=0.20'}
    elif phase =='1':
        if power == "37.5" or power == "38":
            power = "37"
        if (power in list(react_list1F3W[voltageMV][voltageLV].keys())) and (power in list(res_list1F3W[voltageMV][voltageLV].keys())):#####
            rea_out = react_list1F3W[voltageMV][voltageLV][power]#####
            res_out = res_list1F3W[voltageMV][voltageLV][power]####
            imag_out = imag_list1F[power]	
            noloadloss_out=noloadloss_list1F[power]			
            imp = {'X':rea_out,'R':res_out,'Im':imag_out,'Pnoload':noloadloss_out}
        else:
            aviso= u'No existen valores de reactancia y resistencia para transformadores monofásicos de '+power+' kVA operando en '+voltageMV+' kV.'
            # QgsMessageLog.logMessage(aviso,u'Transformadores',Qgis.Warning)
            rea_out = ' Xhl=UNKNOWN Xht=UNKNOWN Xlt=UNKNOWN '
            res_out = ' %Rs=[UNKNOWN] '
            imp = {'X':rea_out,'R':res_out,'Im':'%imag=1.00','Pnoload':'%noloadloss=0.20'}
    return imp

def impedanceMultiUnit(voltageMV,voltageLV, rated_powerA,rated_powerB,rated_powerC, phase):
    voltageMV = str(voltageMV)
    if voltageMV == "13.2":
        voltageMV = "13.8"
    voltageLV = str(voltageLV)
    powerA=str(int(float(rated_powerA)))
    powerB=str(int(float(rated_powerB)))
    powerC=str(int(float(rated_powerC)))
    if powerA == "38":
        powerA = "37"
    if powerB == "38":
        powerB = "37"
    if powerC == "38":
        powerC = "37"


    if voltageLV =="0.277":
        voltageLV="0.24"
    if (powerA in list(imp_list1F2W[voltageMV][voltageLV].keys())):####
        imp_outA=imp_list1F2W[voltageMV][voltageLV][powerA]#####
        imag_outA = imag_list1F[powerA]	
        noloadloss_outA=noloadloss_list1F[powerA]		
        impA = {'Za':imp_outA,'ImA':imag_outA,'PnoloadA':noloadloss_outA}
    else:
        aviso= 'No existen valores de reactancia y resistencia para transformadores monofasicos de '+powerA+' kVA operando en '+voltageMV+' kV.'
        # QgsMessageLog.logMessage(aviso, QCoreApplication.translate('dialog', 'Transformadores'), Qgis.Warning)        
        imp_outA = '%R=UNKNOWN Xhl=UNKNOWN'
        impA = {'Za':imp_outA,'ImA':'%imag=1.00','PnoloadA':'%noloadloss=0.20'}

    if (powerB in list(react_list1F3W[voltageMV][voltageLV].keys())) and (powerB in list(res_list1F3W[voltageMV][voltageLV].keys())):#####
        rea_outB = react_list1F3W[voltageMV][voltageLV][powerB]####3
        res_outB = res_list1F3W[voltageMV][voltageLV][powerB]#####
        imp_outB = imp_list1F2W[voltageMV][voltageLV][powerB]  #####
        imag_outB = imag_list1F[powerB]
        noloadloss_outB=noloadloss_list1F[powerB]
        impB =  {'Xb':rea_outB,'Rb':res_outB,'ImB':imag_outB,'PnoloadB':noloadloss_outB,'Zb':imp_outB}
    else:
        aviso= 'No existen valores de reactancia y resistencia para transformadores monofasicos de '+powerB+' kVA operando en '+voltageMV+' kV.'
        # QgsMessageLog.logMessage(aviso,u'Transformadores',Qgis.Warning)
        rea_outB = 'Xhl=UNKNOWN Xht=UNKNOWN Xlt=UNKNOWN'
        res_outB = '%Rs=[UNKNOWN]'
        imp_outB = '%R=UNKNOWN Xhl=UNKNOWN'
        impB =  {'Xb':rea_outB,'Rb':res_outB,'ImB':'%imag=1.00','PnoloadB':'%noloadloss=0.20','Zb':imp_outB}

    if phase !='.2.3' and phase !='.1.3'and phase !='.1.2':
        if (powerC in list(imp_list1F2W[voltageMV][voltageLV].keys())):####
            imp_outC=imp_list1F2W[voltageMV][voltageLV][powerC]#####
            imag_outC = imag_list1F[powerC]	
            noloadloss_outC=noloadloss_list1F[powerC]			
            impC = {'Zc':imp_outC,'ImC':imag_outC,'PnoloadC':noloadloss_outC}
        else:
            aviso= 'No existen valores de reactancia y resistencia para transformadores monofasicos de '+powerC+' kVA operando en '+voltageMV+' kV.'
            # QgsMessageLog.logMessage(aviso,u'Transformadores',Qgis.Warning)
            imp_outC='%R=UNKNOWN Xhl=UNKNOWN'
            impC = {'Zc':imp_outC,'ImC':'%imag=1.00','PnoloadC':'%noloadloss=0.20'}
    else:
        imp_outC='%R=UNKNOWN Xhl=UNKNOWN'
        impC = {'Zc':imp_outC,'ImC':'%imag=1.00','PnoloadC':'%noloadloss=0.20'}
    impBanco={'impA':impA,'impB':impB,'impC':impC}
    return impBanco
