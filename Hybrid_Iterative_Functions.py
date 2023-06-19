# -*- coding: utf-8 -*-
"""
Created on Wed Oct  7 13:59:32 2020

@author: Orlando Pereira y María José Parajeles 
"""
#%% Packages
from . import auxiliary_functions as auxfcns
import csv
import io
import networkx as nx
import numpy as np
import pandas as pd
pd.options.mode.chained_assignment = None
import pickle
import os
import time
import sys
import traceback
import copy
from qgis.PyQt.QtWidgets import QFileDialog, QMessageBox
from matplotlib import pyplot as plt
from . import trafoOperations_spyder as trafOps


#%% INITIALIZE OPENDSS - Initializes com interface to use opendss from python
# Input: NONE
# Output: DSSobj, DSSstart, DSStext, DSScircuit, DSSprogress

def setUpCOMInterface(): #Función plugin rojo inicializar OpenDSS

    import comtypes.client as cc  # import comtypes library

    DSSobj = cc.CreateObject("OpenDSSEngine.DSS")  # create OpenDSS object
    DSSstart = DSSobj.Start(0)  # start the object
    DSStext = DSSobj.Text  # DSS command introduction
    DSScircuit = DSSobj.ActiveCircuit  # DSS active circuit calling
    DSSprogress = DSSobj.DSSProgress  # DSS progress object
    return DSSobj, DSSstart, DSStext, DSScircuit, DSSprogress

#%% LOADS FEEDER'S LOAD CURVE - Loads the feeder substation load curve - ideally a low load, high load day
# Input: load_curve_path, load_curve_name
# Output: circuit_demand: dict({time instant (i.e. 0,1,2,..,96): [dd/mm/yyyy, hh:mm, P, Q ]})

def loadFeederLoadCurve(load_curve_circuit):
        
    # Primero verificar la extensión del archivo
    file_ext = load_curve_circuit.split(".")[-1].upper() # Extensión del archivo. Debe de ser .csv
    if file_ext == "CSV":

        df_cd = pd.read_csv(load_curve_circuit, sep=None, engine='python').astype(str)
        df_cd.columns = df_cd.columns.str.upper()
        
        # check that the required columns are present in the CSV file
        # Potencia Activa
        p_options = ["P (KW)", "P (MW)", "P(KW)", "P(MW)"]
        if not df_cd.columns.isin(p_options).any():
            msg = "Existe un error en el nombre de la columna de la potencia activa o esta no existe en el archivo de curva de demanda y tensión. \n"
            msg += "Recuerde que el nombre de la columna debe de ser 'P (kW)', 'P (MW)', 'P(kW)' o 'P(MW)', sin importar mayúsculas o minúsculas. \n" 
            msg += "Favor corregir el archivo."
            title = "Error en el archivo de curva de demanda y tensión"
            QMessageBox.information(None, title, msg)
            return None
        else:
            p_name = df_cd.columns[df_cd.columns.isin(p_options)][0]
            if "KW" in p_name:
                factor_p = 1    # Si el valor de la curva está en kW
            else:
                factor_p = 1000 # Si el valor de la curva está en MW
        
        # Potencia Reactiva
        q_options = ["Q (KVAR)", "Q (MVAR)", "Q(KVAR)", "Q(MVAR)"]
        if not df_cd.columns.isin(q_options).any():
            msg = "Existe un error en el nombre de la columna de la potencia reactiva o esta no existe en el archivo de curva de demanda y tensión. \n"
            msg += "Recuerde que el nombre de la columna debe de ser 'Q (kVAr)', 'Q (MVAr)', 'Q(kVAr)' o 'Q(MVAr)', sin importar mayúsculas o minúsculas. \n" 
            msg += "Favor corregir el archivo."
            title = "Error en el archivo de curva de demanda y tensión"
            QMessageBox.information(None, title, msg)
            return None
        else:
            q_name = df_cd.columns[df_cd.columns.isin(q_options)][0]
            if "KVAR" in q_name:
                factor_q = 1    # Si el valor de la curva está en kVAr
            else:
                factor_q = 1000 # Si el valor de la curva está en MVAr
        
        # Tensión
        v_options = ["V (PU)", "V(PU)"]
        if not df_cd.columns.isin(v_options).any():
            msg = "Existe un error en el nombre de la columna de la tensión en la cabecera o esta no existe en el archivo de curva de demanda y tensión. \n"
            msg += "Recuerde que el nombre de la columna debe de ser 'V (pu)' o 'V(pu)', sin importar mayúsculas o minúsculas. \n "
            msg += "Favor corregir el archivo."
            title = "Error en el archivo de curva de demanda y tensión"
            QMessageBox.information(None, title, msg)
            return None
        else:
            v_name = df_cd.columns[df_cd.columns.isin(v_options)][0]
        
        # Verificar que los valores de la columna estén entre 0.8 y 1.5
        if not df_cd[v_name].astype(float).between(0.75, 1.25).all():
            msg = "Los valores asignados para la columna de tensión no son valores en por unidad (pu). \n"
            msg += "Recuerde que el nombre de la columna debe de ser 'V (pu)' o 'V(pu)', sin importar mayúsculas o minúsculas. \n" 
            msg += "Favor corregir el archivo."
            title = "Error en el archivo de curva de demanda y tensión"
            QMessageBox.information(None, title, msg)
            return None
            
        # Día
        d_options = ["DIA", "DAY", "DATE"]
        # Verificación de la existencia de la columna
        if not df_cd.columns.isin(d_options).any():
            msg = "Existe un error en el nombre de la columna de los días medidos o esta no existe en el archivo de curva de demanda y tensión. \n"
            msg += "Recuerde que el nombre de la columna debe de ser 'Dia' (SIN TILDE), sin importar mayúsculas o minúsculas. \n" 
            msg += "Favor corregir el archivo."
            title = "Error en el archivo de curva de demanda y tensión"
            QMessageBox.information(None, title, msg)
            return None
        else:
            d_name = df_cd.columns[df_cd.columns.isin(d_options)][0]
        
        # Hora
        h_options = ["HORA", "HOUR"]
        if not df_cd.columns.isin(d_options).any():
            msg = "Existe un error en el nombre de la columna de las horas medidos o esta no existe en el archivo de curva de demanda y tensión. \n"
            msg += "Recuerde que el nombre de la columna debe de ser 'Hora', sin importar mayúsculas o minúsculas. \n "
            msg += "Favor corregir el archivo."
            title = "Error en el archivo de curva de demanda y tensión"
            QMessageBox.information(None, title, msg)
            return None
        else:
            h_name = df_cd.columns[df_cd.columns.isin(h_options)][0]
        
        # reorder the columns as required
        df_cd = df_cd.loc[:, [d_name, h_name, p_name, q_name, v_name]]
        
        circuit_demand = df_cd.values.tolist()
        
        return circuit_demand
    
    else:
        msg = "La extensión del archivo no es válida. \n"
        msg += "Recuerde que la extensión requerida es de tipo .csv"
        title = "Error en el archivo de curva de demanda y tensión"
        QMessageBox.information(None, title, msg)
        return None  
        
    
    
    

#%% BUS LV GROUP- dataframe.apply function
# Input: DSScircuit, g: bus name, lv_groups
# Output: bus group

def buslvgroup(DSScircuit, g, lv_groups ):
    try: 
        group=int(lv_groups.loc[g.split('.')[0].upper()]['LV_GROUP'])
    except: 
        group=np.nan
    return group

#%% TRANSFORMER NAME ACCORDING TO GROUP NUMBER - dataframe.apply function
# Input: g: group,line_lv_groups
# Output: Trafo's name or nan for nan groups

def txname(g, tx_groups):
    try: 
        txn=tx_groups[tx_groups['LV_GROUP']==int(g)].index.values[-1]
    except: 
        txn=np.nan
    return txn

#%% SNAPSHOT TO GET LN VOLTAGE BASES - Does a basic snapshot simulation without loads and DERS just to find the voltage profile in the middle of the night, and assign the voltage base to every node in every phase
# Input: DSStext, DSScircuit, dss_network: dss path file
# Output: Base_V: Dataframe([all buses base list], index=[buses names by phase: bus.1, bus.2,.. ], columns=['base'] )

def getbases(DSStext, DSScircuit, dss_network, firstLine, lv_groups, tx_groups):
    DSStext.Command = 'clear'  # clean previous circuits
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=snapshot'  # Type of Simulation
    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time                
    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
    DSStext.Command = 'batchedit load..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit storage..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit PVSystem..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit Generator..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit RegControl..* enabled = no' # No RegControls
    DSStext.Command = 'batchedit CapControl..* enabled = no' # No CapControls
    
    DSStext.Command = "VSource.Source.pu = 1.00"
    
    DSScircuit.Solution.Solve()  # Solve the circuit

    VBuses_b = pd.DataFrame(list(DSScircuit.AllBusVmag), index=list(DSScircuit.AllNodeNames), columns=['VOLTAGEV'])
    VBuses_b=VBuses_b[~VBuses_b.index.str.contains('aftermeter')]; VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_der')]; VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_swt')]

    base_vals = [120, 138, 208, 240, 254, 277, 416, 440, 480,
                        2402, 4160, 7620, 7967, 13200, 13800, 14380,
                        19920, 24900, 34500, 79670, 132790]

    Base_V = pd.DataFrame()
    Base_V['BASE'] = VBuses_b['VOLTAGEV'].apply(lambda v : base_vals[[abs(v-i) for i in base_vals].index(min([abs(v-i) for i in base_vals]))])
    Base_V['BUSNAME'] = Base_V.index
    Base_V['LV_GROUP'] = Base_V.BUSNAME.apply(lambda g : buslvgroup(DSScircuit,g, lv_groups) )
    Base_V['TX']= Base_V.LV_GROUP.apply(lambda g : txname(g, tx_groups) )
    del Base_V['BUSNAME']

    Base_V.to_csv('bases.csv')

    return VBuses_b, Base_V

def getbases_simple(DSStext, DSScircuit, dss_network, firstLine):
    DSStext.Command = 'clear'  # clean previous circuits
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=snapshot'  # Type of Simulation
    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time                
    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
    DSStext.Command = 'batchedit load..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit storage..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit PVSystem..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit Generator..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit RegControl..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit CapControl..* enabled = no' # No load simulation
    
    DSStext.Command = "VSource.Source.pu = 1.00"
    
    DSScircuit.Solution.Solve()  # Solve the circuit

    VBuses_b = pd.DataFrame(list(DSScircuit.AllBusVmag), index=list(DSScircuit.AllNodeNames), columns=['VOLTAGEV'])
    VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_der')]; VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_swt')]

    base_vals = [120, 138, 208, 240, 254, 277, 416, 440, 480,
                        2402, 4160, 7620, 7967, 13200, 13800, 14380,
                        19920, 24900, 34500, 79670, 132790]

    Base_V = pd.DataFrame()
    Base_V['BASE'] = VBuses_b['VOLTAGEV'].apply(lambda v : base_vals[[abs(v-i) for i in base_vals].index(min([abs(v-i) for i in base_vals]))])
    Base_V['BUSNAME'] = Base_V.index
    del Base_V['BUSNAME']

    return Base_V

#%% CIRCUIT BASE CASE SNAPSHOT RUN - does a simulation of the circuits base case (only already installed DERs). Gets the circuit's base characteristics for comparison with increased DERs levels in the HC analysis
# Input: DSStext, DSScircuit, snapshotdate, snapshottime, firstLine, tx_modelling, substation_type, line_tx_definition, circuit_demand, Base_V
# Output: No_DERs_run_Vbuses: DataFrame(buses_pu, index = [buses names], columns = ['voltage'] ), kW_sim: load allocation P multiplier, kVAr_sim: load allocation Q multiplier

def base_Case_Run(DSStext, DSScircuit, DSSobj,DSSprogress, snapshotdate,
                  snapshottime, firstLine, tx_modelling, substation_type,
                  line_tx_definition, circuit_demand, Base_V,
                  CircuitBreakDvFF_BFC, CircuitBreakDvRoR, faulttypes,
                  dss_network, lv_loads_layer, mv_loads_layer,
                  tx_layer, FF_analysis, BFC_analysis, RoR_analysis): # time: hh:mm
    t1= time.time()

    #%% Calculate the hour in the simulation
    h, m = snapshottime.split(':')
    if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
        if int(m) <= 7:
            m = '00'
        elif int(m) <= 22:
            m = '15'
        elif int(m) <= 37:
            m = '30'
        elif int(m) <= 52:
            m = '45'
        else:
            m = '00'
            h = str(int(h) + 1)
            if int(h) == 24:  # last round on 23:45
                h = '23'
                m = '45'

    snapshottime = h + ':' + m
    day_ = snapshotdate.replace('/', '')
    day_ = day_.replace('-', '')
    daily_strtime = str(day_ + snapshottime.replace(':', ''))         
    hora_sec = snapshottime.split(':')

    # P and Q to match
    P_to_be_matched = 0
    Q_to_be_matched = 0
    for ij in range(len(circuit_demand)):
        temp_a = circuit_demand[ij][0]  # day
        temp_b = circuit_demand[ij][1]  # hour                    
        if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
            P_to_be_matched = circuit_demand[ij][2]  # Active power
            Q_to_be_matched = circuit_demand[ij][3]  # Reactive power
            V_to_be_matched = circuit_demand[ij][4]  # Voltage
            break

    #%% LoadAllocation Simulation
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time                
    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
    
    # Modify the vpu from source according to circuit_demand voltage curve:
    
    DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched)
    
    # Run the daily power flow for a particular moment
    DSScircuit.Solution.Solve()  # Initialization solution                                            
    
    errorP = 0.003  # Maximum desired correction error for active power
    errorQ = 0.01  # Maximum desired correction error for reactive power
    max_it_correction = 10  # Maximum number of allowed iterations
    study = 'snapshot'  # Study type for PQ_corrector
    gen_powers = np.zeros(1)
    gen_rpowers = np.zeros(1)
    # De acá en adelante hasta proximo # se comenta si no se quiere tomar encuenta la simulación de GD ya instalada
    gen_p = 0
    gen_q = 0
    GenNames = DSScircuit.Generators.AllNames
    if GenNames[0] != 'NONE':
        for i in GenNames: # extract power from generators

            DSScircuit.setActiveElement('generator.' + i)
            p = DSScircuit.ActiveElement.Powers

            for w in range(0, len(p), 2):
                gen_p += -p[w] # P
                gen_q += -p[w + 1] # Q
        gen_powers[0] += gen_p
        gen_rpowers[0] += gen_q

    #DSStext.Command = 'batchedit storage..* enabled = no' # No storage simulation
    # load allocation algorithm
    DSSobj.AllowForms = 0
    [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim,
      kVAr_sim] = auxfcns.PQ_corrector(DSSprogress, DSScircuit, DSStext, errorP, errorQ, max_it_correction,
                                      P_to_be_matched, Q_to_be_matched, V_to_be_matched, hora_sec, study,
                                      dss_network, tx_modelling, 1, firstLine, substation_type,
                                      line_tx_definition, gen_powers, gen_rpowers)

    DSSobj.AllowForms = 1
    t2= time.time()
    print('Tiempo Load Allocation : ' + str(t2 - t1))
    
    
    #%% Post load allocation simulation - base without DERs for Voltage analysis
    
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
    
    DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched)
    DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
    DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
    
    DSStext.Command = "batchedit generator..* enabled=no" # Apaga todos los generadores existentes previo al estudio
    
    DSScircuit.Solution.Solve()
    
    DSScircuit.setActiveElement('line.' + firstLine)
    temp_powers = DSScircuit.ActiveElement.Powers
    print('P_wo_Sub: '+ str(temp_powers[0] + temp_powers[2] + temp_powers[4]))
    print(temp_powers)

    #Vpu
    VBuses = pd.DataFrame(list(DSScircuit.AllBusVmag),
                          index=list(DSScircuit.AllNodeNames),
                          columns=['VOLTAGEV'])
    VBuses=VBuses[~VBuses.index.str.contains('aftermeter')]
    VBuses=VBuses[~VBuses.index.str.contains('_der')]
    VBuses=VBuses[~VBuses.index.str.contains('_swt')]
    VBuses=VBuses[~VBuses.index.str.endswith('.4')]
    No_DERs_run_Vbuses = pd.DataFrame()
    No_DERs_run_Vbuses['VOLTAGE'] = VBuses.VOLTAGEV/Base_V.BASE
    
    #LV layer update 
    lv_loads_layer['kVA_snap'] = lv_loads_layer.DSSNAME.apply(lambda x : get_kva_load(DSScircuit, x))
    tx_layer['kVA_snap'] = tx_layer.DSSNAME.apply(lambda x : get_kva_trafos(DSScircuit, x))
    if mv_loads_layer.empty  is False:
        mv_loads_layer['kVA_snap'] = mv_loads_layer.DSSNAME.apply(lambda x : get_kva_load(DSScircuit, x))
    
    # Retornar el valor de los taps del regulador (RegControl)
    reg_taps = []
    for tx in DSScircuit.Transformers.AllNames:
        if "reg" in tx:
            DSScircuit.Transformers.Name = tx
            tap_val = DSScircuit.ActiveElement.Properties("tap").val 
            reg_taps.append((tx,tap_val))
            
    cap_steps = []
    for cap in DSScircuit.Capacitors.AllNames:
        if DSScircuit.Capacitors.AllNames[0] != "NONE":
            DSScircuit.Capacitors.Name = cap
            step_val = DSScircuit.ActiveElement.Properties("states").val 
            cap_steps.append((cap, step_val))
            
    t3= time.time()
    print('Basecase snapshot time: ' + str(t3 - t2))
    
    
    #%% Fault type studies
    # defines the terminals necessary for each type of fault.
    # Dictionary containing: Fault type: [terminal
    # conection for bus 1, terminal conection for bus2, number of phases]
    faulttypeterminals = {'ABC':['.1.2','.3.3', '2'],
                          'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'],
                          'BC':['.2','.3','1'], 'AC':['.1','.3','1'],
                          'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'],
                          'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'],
                          'BG':['.2','.0','1'], 'CG':['.3','.0','1']}

    if (FF_analysis or BFC_analysis) is True:
        ### FORWARD CURRENT
        zp = zip(CircuitBreakDvFF_BFC['Element'],
                 CircuitBreakDvFF_BFC['BusInstalled'])
        mltidx = pd.MultiIndex.from_tuples(zp, names=['Element', 'FaultedBus'])
        No_DER_FFCurrents = pd.DataFrame(index=mltidx, columns=list(faulttypes))
        #BFC
        temp=CircuitBreakDvFF_BFC[CircuitBreakDvFF_BFC['RecloserElement'].notnull()]
        zp = zip(temp['Element'], temp['RecloserElement'], temp['BusInstalled'])
        mltidx = pd.MultiIndex.from_tuples(zp, names=['FuseElement', 'BreakerElement',
                                           'FaultedBus'])
        No_DER_BFCCurrents = pd.DataFrame(index=mltidx, columns=list(faulttypes))
        del temp
        
        for element, faultedbus in list(No_DER_FFCurrents.index):
            print('FF: '+element)
            ftlist=faultedbus.split('.')
            if len(ftlist) == 2:
                currentspositions = {'AB': [0, 2], 'BC': [0, 2], 'AC': [0, 2],
                                     'ABG': [0, 2], 'BCG': [2, 4],
                                     'ACG': [0, 2], 'AG': [0],
                                     'BG': [0], 'CG': [0]}
                if ftlist[1] == '1':
                    faultlabels = ['AG']
                elif ftlist[1] == '2':
                    faultlabels = ['BG']
                else: 
                    faultlabels = ['CG']
            elif len(ftlist) == 3:
                currentspositions = {'AB': [0, 2], 'BC': [0, 2], 'AC': [0, 2],
                                     'ABG': [0, 2], 'BCG': [2, 4],
                                     'ACG': [0, 2], 'AG': [0],
                                     'BG': [0], 'CG': [0]}
                if ftlist[1] == '1' and ftlist[2] == '2':
                    faultlabels = ['AB', 'ABG','AG','BG']
                elif ftlist[1] == '2' and ftlist[2] == '3':
                    faultlabels = ['BC', 'BCG', 'CG', 'BG']
                else: 
                    faultlabels = ['AC', 'ACG', 'AG', 'CG']
            else:
                currentspositions = {'ABC': [0, 2, 4], 'ABCG': [0, 2, 4],
                                     'AB': [0, 2], 'BC': [2, 4],
                                     'AC': [0, 4], 'ABG': [0, 2],
                                     'BCG': [2, 4], 'ACG': [0, 4], 'AG': [0],
                                     'BG': [2], 'CG': [4]}
                faultlabels = list(faulttypeterminals.keys())
    
            faultlabels = list(set(faulttypes)&set(faultlabels)) # removes those types of faults not being evaluated
    
            for faulttype in faultlabels:
                print('FF: '+element+ ', Fault: '+faulttype)
                # Machines inicialization simulation
                DSStext.Command = 'clear'
                DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
                DSStext.Command = 'Set mode=daily'  # Type of Simulation
                DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched) # Tensión en cabecera
                DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
                DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                DSScircuit.Solution.Solve()
    
                #%% Fault study (dynamics) simulation
                # type of fault
                term1 = faulttypeterminals[faulttype][0]
                term2 = faulttypeterminals[faulttype][1]
                phasessc = faulttypeterminals[faulttype][2]
    
                DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                DSStext.Command = 'new fault.fault_'+ faulttype+' bus1='+ftlist[0]+term1+' bus2='+ftlist[0]+term2+' phases='+phasessc
                DSScircuit.Solution.Solve()
    
                DSScircuit.setActiveElement(element) # faulted element
                temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                No_DER_FFCurrents[faulttype][element, faultedbus] = [temp_currents[i] for i in currentspositions[faulttype]]
    
                #BFC
                breakerelement = CircuitBreakDvFF_BFC.loc[CircuitBreakDvFF_BFC['Element']==element, 'RecloserElement'].iloc[-1]
                if type(breakerelement) is str:
                    No_DER_BFCCurrents[faulttype][element, breakerelement, faultedbus] = []
                    for element_i in (element, breakerelement):
                        DSScircuit.setActiveElement(element_i) # faulted element
                        temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                        No_DER_BFCCurrents[faulttype][element, breakerelement, faultedbus].append([temp_currents[i] for i in currentspositions[faulttype]])
    
    else:             
        No_DER_FFCurrents = pd.DataFrame()
        No_DER_BFCCurrents = pd.DataFrame()
        
        ### REDUCTION OF REACH
    if RoR_analysis is True:
            
        elements = []; FaultedBus = []
        for element in list(CircuitBreakDvRoR['Element']):
            if CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusZone'].iloc[-1] == CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusBackUp'].iloc[-1]:
                elements.append(element)
                FaultedBus.append(CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusZone'].iloc[-1])
            else:
                elements.append(element)
                elements.append(element)
                FaultedBus.append(CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusZone'].iloc[-1])
                FaultedBus.append(CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusBackUp'].iloc[-1])
        
       
        mltidx = pd.MultiIndex.from_tuples(zip(elements,FaultedBus), names=['Element', 'FaultedBus'])
        No_DER_RoRCurrents = pd.DataFrame(index=mltidx, columns=list(faulttypes))
    
        for element, faultedbus in list(No_DER_RoRCurrents.index):
            print('ROR: '+element)
            ftlist=faultedbus.split('.')
            if len(ftlist) == 2:
                currentspositions = {'AB': [0, 2], 'BC': [0, 2], 'AC': [0, 2],
                                     'ABG': [0, 2], 'BCG': [2, 4],
                                     'ACG': [0, 2], 'AG': [0], 'BG': [0],
                                     'CG':[0]}
                if ftlist[1] == '1':
                    faultlabels = ['AG']
                elif ftlist[1] == '2':
                    faultlabels = ['BG']
                else: 
                    faultlabels = ['CG']
            elif len(ftlist) == 3:
                currentspositions = {'AB': [0, 2], 'BC': [0, 2], 'AC': [0, 2],
                                     'ABG': [0, 2], 'BCG': [2, 4],
                                     'ACG': [0, 2], 'AG': [0], 'BG': [0],
                                     'CG':[0]}
                if ftlist[1] == '1' and ftlist[2] == '2':
                    faultlabels = ['AB', 'ABG','AG','BG']
                elif ftlist[1] == '2' and ftlist[2] == '3':
                    faultlabels = ['BC', 'BCG','CG','BG']
                else: 
                    faultlabels = ['AC', 'ACG','AG','CG']
            else:
                currentspositions = {'ABC': [0, 2, 4], 'ABCG': [0, 2, 4],
                                     'AB': [0, 2], 'BC': [2, 4], 'AC': [0, 4],
                                     'ABG': [0, 2], 'BCG': [2, 4],
                                     'ACG': [0, 4], 'AG': [0],
                                     'BG': [2], 'CG': [4]}
                faultlabels = list(faulttypeterminals.keys())
    
            faultlabels = list(set(faulttypes) & set(faultlabels)) # removes those types of faults not being evaluated
            for faulttype in faultlabels:
                print('ROR: '+element+ ', Fault: '+faulttype)
                # Machines inicialization simulation
                DSStext.Command = 'clear'
                DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
                DSStext.Command = 'Set mode=daily'  # Type of Simulation
                DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched) # Tensión en cabecera
                DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
                DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                DSScircuit.Solution.Solve()
    
        
                #%% Fault study (dynamics) simulation
                # type of fault
                term1 = faulttypeterminals[faulttype][0]
                term2 = faulttypeterminals[faulttype][1]
                phasessc = faulttypeterminals[faulttype][2]
    
                line_com = 'new fault.fault_'+ faulttype +' bus1=' + ftlist[0]
                line_com += term1 + ' bus2=' + ftlist[0] + term2
                line_com += ' phases=' + phasessc
                DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                DSStext.Command = line_com
                DSScircuit.Solution.Solve()
    
                DSScircuit.setActiveElement(element) # faulted element
                temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                tmp = [temp_currents[i] for i in currentspositions[faulttype]]
                No_DER_RoRCurrents[faulttype][element, faultedbus] = tmp
    
        t4= time.time()
        print('Basecase fault time: '+str(t4-t3))
        
    else:
        No_DER_RoRCurrents = pd.DataFrame()

    return No_DERs_run_Vbuses, No_DER_FFCurrents, No_DER_RoRCurrents, No_DER_BFCCurrents, kW_sim, kVAr_sim, reg_taps, cap_steps, lv_loads_layer

#%% FIND THE LINES NOMINAL CAPACITY (already exixts in plug in rojo) - Gets the nominal capacity of lines for thermal analysis in HC
# Input: DSScircuit
# Output: normalAmpsDic: normalAmpsDic[LINE NAME] = lineNormalAmps

def normalAmps(DSScircuit):
    normalAmpsDic = {}
    lineNames = list(DSScircuit.Lines.AllNames)
    for name in lineNames:
        DSScircuit.SetActiveElement('line.' + name)
        lineNormalAmps = DSScircuit.ActiveCktElement.NormalAmps  # read normal Amps for each line
        normalAmpsDic[str(name).upper()] = lineNormalAmps  # add normal amps in dictionary
    return normalAmpsDic

#%% FIND LOADS BUSES (no existe en plug in rojo) - Creates two lists with all the LV and MV loads names 
# Input: loadsLV_file_name, loadsMV_file_name
# Output: loadslv_buses: list(busname.1, busname.2, ...), loadsmv_buses=list(busname.1, busname.2, ...)

def getloadbuses(loadsLV_file_name, loadsMV_file_name, dss_network):

    name = os.path.join(dss_network, loadsLV_file_name)
    if ".dss" not in name:
        name += ".dss"
    with open(name,'r') as f:
        loadslv_file = f.readlines()
    try:
        name = os.path.join(dss_network, loadsMV_file_name)
        if ".dss" not in name:
            name += ".dss"
        with open(name, 'r') as f:
            loadsmv_file = f.readlines()
    except:
        loadsmv_file=list()
    
    loadslv_buses = list() 
    for loadline in loadslv_file:
        busld = loadline.split(' ')[2].split('=')[1].split('.')
        bus_list = [ busld[0].lower()+'.'+ busld[i] for i in range(1,len(busld))]
        loadslv_buses = loadslv_buses + bus_list
    loadsmv_buses = list() 
    for loadline in loadsmv_file:
        busld = loadline.split(' ')[2].split('=')[1].split('.')
        bus_list = [ busld[0].lower()+'.'+ busld[i] for i in range(1,len(busld))]
        loadsmv_buses = loadsmv_buses + bus_list
    return loadslv_buses, loadsmv_buses

#%% FINDS GD CAPACITY ALREADY INSTALLED
# Input: name_file_created - GD file 
# Output: DERinstalled_buses : list(busname)

def getGDinstalled(der_ss_layer):
    try:
        DERinstalled_buses = list(der_ss_layer["bus1"].values) # forbidden buses --> buses with DER already installed
        DERinstalled = der_ss_layer["KVA"].sum()
    except:
        DERinstalled_buses = []  # forbidden buses --> buses with DER already installed
        DERinstalled=0
    return DERinstalled_buses, DERinstalled

#%% DERs SNAPSHOT RUN 
# Input: DSStext, DSScircuit, snapshotdate, snapshottime, firstLine, tx_modelling, substation_type, line_tx_definition, circuit_demand, Base_V, kW_sim, kVar_sim, DERs
# Output: NONE, just runs the simulation and updates OPENDSS instances

def DERs_Run(DSStext, DSScircuit, snapshotdate, snapshottime, firstLine,
            tx_modelling, substation_type, line_tx_definition, circuit_demand,
            Base_V, kW_sim, kVAr_sim, DERs, Trafos, dss_network, reg_taps, cap_steps): # time: hh:mm
    #%% Calculate the hour in the simulation
    h, m = snapshottime.split(':')
    if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
        if int(m) <= 7:
            m = '00'
        elif int(m) <= 22:
            m = '15'
        elif int(m) <= 37:
            m = '30'
        elif int(m) <= 52:
            m = '45'
        else:
            m = '00'
            h = str(int(h) + 1)
            if int(h) == 24:  # last round on 23:45
                h = '23'
                m = '45'
    
    snapshottime = h + ':' + m
    day_ = snapshotdate.replace('/', '')
    day_ = day_.replace('-', '')
    daily_strtime = str(day_ + snapshottime.replace(':', ''))
    hora_sec = snapshottime.split(':')
    
    # V to match
    V_to_be_matched = 0
    for ij in range(len(circuit_demand)):
        temp_a = circuit_demand[ij][0]  # day
        temp_b = circuit_demand[ij][1]  # hour                    
        if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
            V_to_be_matched = circuit_demand[ij][4]  # Voltage
            break
    
    #%% simulation
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
    DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched) # Tensión en cabecera
    DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
    DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
    DSStext.Command = 'batchedit RegControl..* enabled = no' # No RegControls
    DSStext.Command = 'batchedit CapControl..* enabled = no' # No CapControls
    
    for tx in Trafos:
        DSStext.Command = tx
    for der in DERs:
        DSStext.Command = der
    
    for reg_t in reg_taps:
        reg = reg_t[0]
        tap = reg_t[1]
        # Fija la posición del tap del regulador con respecto al caso base
        DSStext.Command = "Transformer."+reg+".tap ="+tap
    
    for cap_s in cap_steps:
        cap = cap_s[0]
        step = cap_s[1]
        # Fija la posición del tap del regulador con respecto al caso base
        DSStext.Command = "Capacitor."+cap+".states ="+step
        
    DSScircuit.Solution.Solve()

#%% THREE PHASE MV BUSES TO ALLOCATE LARGE SCALE DER - finds the buses every 100 m to intall DER of large scale
# Input: nodos_mt, lineas_mt, trafos, fixed_distance
# Output: chosen_buses_dict: dict with distance every 100 m as keys and a tuple as value (BUSMV, NOMVOLTMV), chosen_buses_list: list of all buses for every distance
def distance_nodes_MV(G, nodes_mv, lineas_mt, trafos, fixed_distance):
    # Empieza algoritmo
    #rand_bus = list(G.nodes.keys())[0]
    #ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    # first_bus = 'BUSMV' + ckt_name + '1'
    first_bus = 'AFTERMETER'
    

    mt_3bus_list = nx.get_node_attributes(G,'num_phases') #list of 3phase mv buses on the whole circuit
    remove = [k for k in mt_3bus_list if mt_3bus_list[k] != '3']
    for k in remove: del mt_3bus_list[k]
    
    df_distancia = pd.DataFrame(np.nan, index=list(G.nodes), columns=['nodes', 'distance'])
    df_distancia['nodes'] = list(G.nodes)
    print("first_bus = ", first_bus)
    
    for idx, yield_ in df_distancia.iterrows():
        node = yield_['nodes']
        try:
            dist = nx.shortest_path_length(G, first_bus, node,
                                           weight='distance')
        except:
            dist = 0
            print("Error nodo ", node)
        df_distancia.loc[idx, 'distance']= dist
    
    # df_distancia['distance']=df_distancia['nodes'].apply(lambda b: shortest_path_handling_errors(first_bus, b))
    # df_distancia['distance'] = df_distancia['distance'].astype(float)
    
    df_distancia = df_distancia.loc[list(mt_3bus_list)] #filtered dataframe with 3phase buses    
    
    n_it = int(np.ceil(max(df_distancia['distance'])/fixed_distance)) #number of iterations
    dist_dict = {}
    # n_it=3
    for n in range(n_it):
        
        #initialize variables
        dist_dict[(n+1)*fixed_distance] = {}
        
        temp_df = pd.DataFrame()
        #delimited dataframe
        temp_df = df_distancia.loc[(df_distancia['distance'] >= (n)*fixed_distance) & (df_distancia['distance'] < (n+1)*fixed_distance)]
        
        if temp_df.empty  is False:
            temp_df['distance'] = temp_df['distance'].astype(float)
    
            source = []
            number_paths = 0
            
            if n == 0: #caso inicial
                source.append(temp_df.loc[temp_df['distance'].idxmin(), 'nodes'])
            else:
                for tgts in dist_dict[(n)*fixed_distance]: 
                    source.append(dist_dict[(n)*fixed_distance][tgts]['target'])
                    
            #identified sources
            for u_source in source:
                downstream_source_allnodes = nx.descendants(G, u_source)
                downstream_source_dfnodes = list(set(downstream_source_allnodes).intersection(temp_df.index.values)) #just nodes of interest
                
                for node in downstream_source_dfnodes:
                    #if it has exclusively nodes that are not part from the downstream_df_nodes
                    downstream_node_studied = [y for y in G.neighbors(node)]
                    downstream_node_3ph_studied = [z for z in set(downstream_node_studied) & set(mt_3bus_list)]
                    
                    if len(downstream_node_3ph_studied) == 0:
                        number_paths+=1
                        
                        dist_dict[(n+1)*fixed_distance][number_paths] = {}
                        dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                        dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node
                        
                        edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]
                        
                        try:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                        except:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]
                        
                        dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
                    
                    else:
                        excluded_from_df = list(set(downstream_node_3ph_studied) - set(temp_df.index.values))
                        
                        if len(excluded_from_df) > 0:
                            # target_list.append(node)
                            number_paths+=1
                            
                            dist_dict[(n+1)*fixed_distance][number_paths] = {}
                            dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                            dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node
                            
                            edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]
                            
                            try:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                            except:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]
                            
                            dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
                            
                            

    chosen_buses_dict = {}
    chosen_buses_df = pd.DataFrame(np.nan, index=[], columns=['distance', 'nomvolt'])
    chosen_buses_list = []
    
    for step in dist_dict:
        chosen_buses_dict[step] = []
        step_pd = pd.DataFrame(np.nan, index=[], columns = ['source', 'target', 'distance'])
        source_list = []
        for key in dist_dict[step]:
            step_pd.loc[key,'source'] = dist_dict[step][key]['source']
            step_pd.loc[key,'target'] = dist_dict[step][key]['target']
        
        for idx, yield_ in step_pd.iterrows():
            source = yield_['source']
            target = yield_['target']
            try:
                dist = nx.shortest_path_length(G, source, target,
                                                weight='distance')
            except:
                dist = 0
                print("Error nodo origen ", first_bus, " nodo destino ", node)
            step_pd.loc[idx, 'distance']= dist
        # step_pd['distance'] = step_pd.apply(lambda b: nx.shortest_path_length(G, b.source, b.target, weight='distance'), axis=1)
        source_list = list(set(step_pd['source'].values))
        
        
        
        for s in source_list:
            id_max = step_pd.loc[step_pd['source']==s]['distance'].idxmax()
            info_tuple = (step_pd.loc[id_max,'target'], nodes_mv.loc[step_pd.loc[id_max,'target'], 'NOMVOLT'])
            chosen_buses_list.append(info_tuple)
            chosen_buses_dict[step].append(info_tuple)
            chosen_buses_df.loc[info_tuple[0], 'distance'] = step
            chosen_buses_df.loc[info_tuple[0], 'nomvolt'] = info_tuple[1]
        
    return chosen_buses_dict, chosen_buses_list, chosen_buses_df

def tri_ph_dist(G, bus):
    
    #Upstream buses calculation 
    us_bus = list(nx.ancestors(G, bus)) + [bus]
    mt_3bus_list = [x for x,y in G.nodes(data=True, default=None) if y and y['num_phases']=='3']
    mt_3_upstream = list(set(us_bus).intersection(mt_3bus_list))
    
    bus_study = pd.DataFrame(np.nan, columns=['u_bus', 'distance'], index = mt_3_upstream)
    bus_study['u_bus'] = mt_3_upstream
    bus_study['distance'] = bus_study['u_bus'].apply(lambda b: float(nx.shortest_path_length(G, b, bus, weight='distance')))
            
    # print(bus)
    return bus_study.loc[bus_study['distance'] == bus_study['distance'].min()].index.values[0]
        
    
def data_grouping(G, MV_hist_df , fixed_distance, final_der,
                  lines_mv_oh_layer_original, lines_mv_ug_layer_original, lim_kVA):
    

    first_bus = 'AFTERMETER'
    
    ###############################################################################################
    # AGGRUPATION CODE
    
    # find the first 3ph upstream bus per node
    distance_to_3ph_bus = MV_hist_df[['DSSNAME', 'bus1', final_der]]
    distance_to_3ph_bus.loc[:,'min_bus'] = distance_to_3ph_bus.loc[:,'bus1'].apply(lambda b: tri_ph_dist(G, b))
    # Getting the list of the selected 3ph buses
    list_3phase_nodes = list(set(distance_to_3ph_bus.loc[distance_to_3ph_bus[final_der]>0]['min_bus']))
    
    accumulated_kVA = pd.DataFrame(np.nan, index=[], columns=['bus', 'kVA_val'])
    accumulated_kVA['bus'] = list_3phase_nodes
    accumulated_kVA['kVA_val'] = accumulated_kVA['bus'].apply(lambda x: distance_to_3ph_bus.loc[distance_to_3ph_bus['min_bus'] == x][final_der].sum())
    accumulated_kVA.index = accumulated_kVA['bus']
    

    mt_3bus_list = [x for x,y in G.nodes(data=True, default=None) if y and y['num_phases']=='3']
    
    df_distancia = pd.DataFrame(np.nan, index=mt_3bus_list, columns=['nodes', 'distance'])
    df_distancia['nodes'] = mt_3bus_list
    for idx, yield_ in df_distancia.iterrows():
        node = yield_['nodes']
        try:
            dist = nx.shortest_path_length(G, first_bus, node,
                                           weight='distance')
        except:
            dist = 0
            print("Error nodo ", node)
        df_distancia.loc[idx, 'distance']= dist
    # df_distancia['distance'] = df_distancia['nodes'].apply(lambda b: nx.shortest_path_length(G, first_bus, b, weight='distance'))
    # df_distancia['distance'] = df_distancia['distance'].astype(float)
    
    # df_distancia = df_distancia.loc[mt_3bus_list] #filtered dataframe with 3phase buses
    
    
    n_it = int(np.ceil(max(df_distancia['distance'])/fixed_distance)) #number of iterations
    dist_dict = {}
    # n_it=3
    for n in range(n_it):
        
        #initialize variables
        dist_dict[(n+1)*fixed_distance] = {}
        
        temp_df = pd.DataFrame()
        #delimited dataframe
        temp_df = df_distancia.loc[(df_distancia['distance'] >= (n)*fixed_distance) & (df_distancia['distance'] < (n+1)*fixed_distance)]
        
        if temp_df.empty  is False:
            temp_df['distance'] = temp_df['distance'].astype(float)
    
            source = []
            number_paths = 0
            
            if n == 0: #caso inicial
                source.append(temp_df.loc[temp_df['distance'].idxmin(), 'nodes'])
            else:
                for tgts in dist_dict[(n)*fixed_distance]: 
                    source.append(dist_dict[(n)*fixed_distance][tgts]['target'])
                    
            #identified sources
            for u_source in source:
                downstream_source_allnodes = nx.descendants(G, u_source)
                downstream_source_dfnodes = list(set(downstream_source_allnodes).intersection(temp_df.index.values)) #just nodes of interest
                for u_source2 in list(set(source) - set([u_source])): #deletes possible intersection between zones
                    if u_source2 in downstream_source_allnodes:
                        usource2_group = [u_source2] + list(nx.descendants(G, u_source2))
                        downstream_source_dfnodes = list(set(downstream_source_dfnodes) - set(usource2_group))
                
                for node in downstream_source_dfnodes:
                    #if it has exclusively nodes that are not part from the downstream_df_nodes
                    downstream_node_studied = [y for y in G.neighbors(node)]
                    downstream_node_3ph_studied = [z for z in set(downstream_node_studied) & set(mt_3bus_list)]
                    
                    if len(downstream_node_3ph_studied) == 0:
                        number_paths+=1
                        
                        dist_dict[(n+1)*fixed_distance][number_paths] = {}
                        dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                        dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node
                        
                        edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]
                        
                        try:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                        except:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]
                        
                        dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
                    
                    else:
                        excluded_from_df = list(set(downstream_node_3ph_studied) - set(temp_df.index.values))
                        
                        if len(excluded_from_df) > 0:
                            # target_list.append(node)
                            number_paths+=1
                            
                            dist_dict[(n+1)*fixed_distance][number_paths] = {}
                            dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                            dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node
                            
                            edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]
                            
                            try:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                            except:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]
                            
                            dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
        
        
        path_dict = {}
        
        for dist in dist_dict:
            for key in dist_dict[dist]:
                path_dict[dist_dict[dist][key]['source']] = {}
                path_dict[dist_dict[dist][key]['source']]['nodes'] = []
                path_dict[dist_dict[dist][key]['source']]['edges'] = []
                path_dict[dist_dict[dist][key]['source']]['kVA'] = 0
              
            for key in dist_dict[dist]:
                path_dict[dist_dict[dist][key]['source']]['nodes'] = list(set(path_dict[dist_dict[dist][key]['source']]['nodes'] + dist_dict[dist][key]['nodes']))
                path_dict[dist_dict[dist][key]['source']]['edges'] = list(set(path_dict[dist_dict[dist][key]['source']]['edges'] + dist_dict[dist][key]['edges']))
            
        for source_key in path_dict:
            nodes_in_accKVAdf = list(set(path_dict[source_key]['nodes']).intersection(accumulated_kVA.index.values))
            try:
                path_dict[source_key]['kVA'] += np.round(accumulated_kVA.loc[nodes_in_accKVAdf, 'kVA_val'].sum(),1)
            except:
                pass
        
        lines_mv_oh_layer_original.loc[:, "HC_RES_SS"] = np.nan
        try:
            lines_mv_ug_layer_original.loc[:, "HC_RES_SS"] = np.nan
        except:
            pass

        for source_key in path_dict:
            nodes_in_oh_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_oh_layer_original['DSSNAME'].values))
            idx_oh_list = lines_mv_oh_layer_original.loc[lines_mv_oh_layer_original['DSSNAME'].isin(nodes_in_oh_layer)].index.values
            lines_mv_oh_layer_original.loc[idx_oh_list, "HC_RES_SS"] = path_dict[source_key]['kVA']
            
            try:
                nodes_in_ug_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_ug_layer_original['DSSNAME'].values))
                idx_ug_list = lines_mv_ug_layer_original.loc[lines_mv_ug_layer_original['DSSNAME'].isin(nodes_in_ug_layer)].index.values
                lines_mv_ug_layer_original.loc[idx_ug_list, "HC_RES_SS"] = path_dict[source_key]['kVA']
            
            except:
                pass
            
    return lines_mv_oh_layer_original, lines_mv_ug_layer_original

#%%
def data_grouping_iterative(G, chosen_buses_df , fixed_distance,
                            mv_lines_layer, lines_mv_oh_layer_original,
                            lines_mv_ug_layer_original):
    
    first_bus = 'AFTERMETER'

    mt_3bus_list = nx.get_node_attributes(G,'num_phases') #list of 3phase mv buses on the whole circuit
    remove = [k for k in mt_3bus_list if mt_3bus_list[k] != '3']
    for k in remove: del mt_3bus_list[k]
    
    df_distancia = pd.DataFrame(np.nan, index=list(G.nodes), columns=['nodes', 'distance'])
    df_distancia['nodes'] = list(G.nodes)
    for idx, yield_ in df_distancia.iterrows():
        node = yield_['nodes']
        try:
            dist = nx.shortest_path_length(G, first_bus, node,
                                           weight='distance')
        except:
            dist = 0
            print("Error nodo ", node)
        df_distancia.loc[idx, 'distance']= dist
    # df_distancia['distance']=df_distancia['nodes'].apply(lambda b: nx.shortest_path_length(G, first_bus, b, weight='distance'))
    # df_distancia['distance'] = df_distancia['distance'].astype(float)
    
    df_distancia = df_distancia.loc[mt_3bus_list] #filtered dataframe with 3phase buses
    
    
    n_it = int(np.ceil(max(df_distancia['distance'])/fixed_distance)) #number of iterations
    dist_dict = {}
    # n_it=3
    for n in range(n_it):
        
        #initialize variables
        dist_dict[(n+1)*fixed_distance] = {}
        
        temp_df = pd.DataFrame()
        #delimited dataframe
        temp_df = df_distancia.loc[(df_distancia['distance'] >= (n)*fixed_distance) & (df_distancia['distance'] < (n+1)*fixed_distance)]
        
        if temp_df.empty  is False:
            temp_df['distance'] = temp_df['distance'].astype(float)
    
            source = []
            number_paths = 0
            
            if n == 0: #caso inicial
                source.append(temp_df.loc[temp_df['distance'].idxmin(), 'nodes'])
            else:
                for tgts in dist_dict[(n)*fixed_distance]: 
                    source.append(dist_dict[(n)*fixed_distance][tgts]['target'])
                    
            #identified sources
            for u_source in source:
                downstream_source_allnodes = nx.descendants(G, u_source)
                downstream_source_dfnodes = list(set(downstream_source_allnodes).intersection(temp_df.index.values)) #just nodes of interest
                for u_source2 in list(set(source) - set([u_source])): #deletes possible intersection between zones
                    if u_source2 in downstream_source_allnodes:
                        usource2_group = [u_source2] + list(nx.descendants(G, u_source2))
                        downstream_source_dfnodes = list(set(downstream_source_dfnodes) - set(usource2_group))
                
                for node in downstream_source_dfnodes:
                    #if it has exclusively nodes that are not part from the downstream_df_nodes
                    downstream_node_studied = [y for y in G.neighbors(node)]
                    downstream_node_3ph_studied = [z for z in set(downstream_node_studied) & set(mt_3bus_list)]
                    
                    if len(downstream_node_3ph_studied) == 0:
                        number_paths+=1
                        
                        dist_dict[(n+1)*fixed_distance][number_paths] = {}
                        dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                        dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node
                        
                        edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]
                        
                        try:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                        except:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]
                        
                        dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
                    
                    else:
                        excluded_from_df = list(set(downstream_node_3ph_studied) - set(temp_df.index.values))
                        
                        if len(excluded_from_df) > 0:
                            # target_list.append(node)
                            number_paths+=1
                            
                            dist_dict[(n+1)*fixed_distance][number_paths] = {}
                            dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                            dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node
                            
                            edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]
                            
                            try:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                            except:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]
                            
                            dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
        
        
        path_dict = {}
        
        for dist in dist_dict:
            for key in dist_dict[dist]:
                path_dict[dist_dict[dist][key]['source']] = {}
                path_dict[dist_dict[dist][key]['source']]['nodes'] = []
                path_dict[dist_dict[dist][key]['source']]['edges'] = []
                path_dict[dist_dict[dist][key]['source']]['kVA'] = 0
              
            for key in dist_dict[dist]:
                path_dict[dist_dict[dist][key]['source']]['nodes'] = list(set(path_dict[dist_dict[dist][key]['source']]['nodes'] + dist_dict[dist][key]['nodes']))
                path_dict[dist_dict[dist][key]['source']]['edges'] = list(set(path_dict[dist_dict[dist][key]['source']]['edges'] + dist_dict[dist][key]['edges']))
            
        for source_key in path_dict:
            nodes_in_accKVAdf = list(set(path_dict[source_key]['nodes']).intersection(chosen_buses_df.index.values))
            try:
                path_dict[source_key]['kVA'] += chosen_buses_df.loc[nodes_in_accKVAdf, 'max_kVA'].sum()
            except:
                pass
        
        #####################################################################################################

        lines_mv_oh_layer_original.loc[:, "HC_RES_LS"] = np.nan
        try:
            lines_mv_ug_layer_original.loc[:, "HC_RES_LS"] = np.nan
        except:
            pass

        for source_key in path_dict:
            nodes_in_oh_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_oh_layer_original['DSSNAME'].values))
            idx_oh_list = lines_mv_oh_layer_original.loc[lines_mv_oh_layer_original['DSSNAME'].isin(nodes_in_oh_layer)].index.values
            lines_mv_oh_layer_original.loc[idx_oh_list, "HC_RES_LS"] = path_dict[source_key]['kVA']
            
            try:
                nodes_in_ug_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_ug_layer_original['DSSNAME'].values))
                idx_ug_list = lines_mv_ug_layer_original.loc[lines_mv_ug_layer_original['DSSNAME'].isin(nodes_in_ug_layer)].index.values
                lines_mv_ug_layer_original.loc[idx_ug_list, "HC_RES_LS"] = path_dict[source_key]['kVA']
            
            except:
                pass
            
                
    return lines_mv_oh_layer_original, lines_mv_ug_layer_original

#%% WRITE DER AND STEPUP TRANSFORMER OPENDSS SENTENCES - allocates DER in bus list and assigns corresponding MV-LV transformer
# Input: bus_list, installed_capacity: DER size
# Output: Trafos: tx sentences, Trafos_Monitor, DERs: DERs'sentences
def trafos_and_DERs_text_command(bus_list, installed_capacity, der_type, der_type_val):

    DERs = []
    DERs_Monitor = []
    Trafos = []
    Trafos_Monitor =[]
    
    # tomar el valor de la capcidad y asociarla con el KVA de la lista de trafos, tal que este
    # sea el más cercano al 15% o 20% de la capacidad que se quiere instalar

    for item in range(len(bus_list)):
        
        bus_name = bus_list[item][0] #nombre del bus
        bus_nomvolt = int(bus_list[item][1]) #nomvolt code
        
        for kva_rating in [float(i) for i in list(trafOps.imag_list3F.keys())]:
            if kva_rating > 1.20*installed_capacity:
                kVA = str(kva_rating)
                break

        cantFases = '3' 
        normhkva = " normhkva=" + str(kVA)
        kV_LowLL = str(trafOps.renameVoltage(bus_nomvolt, 50)['LVCode']['LL']) #El lado de baja es 480 (50)
        kV_LowLN = str(trafOps.renameVoltage(bus_nomvolt, 50)['LVCode']['LN'])
        kV_MedLL = str(trafOps.renameVoltage(bus_nomvolt, 50)['MVCode']['LL'])
        kV_MedLN = str(trafOps.renameVoltage(bus_nomvolt, 50)['MVCode']['LN'])
        busMV = bus_name
        busLV = 'BUSLV' + busMV.replace('BUSMV', '') + '_DER'
        confMV = 'wye'
        confLV = 'wye'
        
        tap = '1'
    
        impedance = trafOps.impedanceSingleUnit(cantFases, kV_MedLL, kV_LowLN, kVA).get('Z')
        noloadloss = trafOps.impedanceSingleUnit(cantFases, kV_MedLL, kV_LowLN, kVA).get('Pnoload')
        imag = trafOps.impedanceSingleUnit(cantFases, kV_MedLL, kV_LowLN, kVA).get('Im')
        trafName =  'TX_DER_' + busMV.replace('BUSMV', '')   
        derName = 'DER_' + busMV.replace('BUSMV', '')

        trafo_line = 'new transformer.' + trafName + ' phases=3 windings=2 '
        trafo_line += noloadloss + " " + imag + ' buses=[' + busMV + '.1.2.3 '
        trafo_line += busLV + '.1.2.3] conns=[' + confMV + ' ' + confLV + ']'
        trafo_line += ' kvs=[' + kV_MedLL + " " +  kV_LowLL + ']'
        trafo_line += ' kvas=[' + kVA + " " + kVA + '] ' + impedance
        trafo_line += ' Taps=[' + tap + ', 1]' + normhkva
        
        Trafos.append(trafo_line) # Se añade a la lista de Trafos
        
        trafo_line_monitor = "new monitor.Mon" + trafName
        trafo_line_monitor += " Element=transformer." + trafName
        trafo_line_monitor += " Terminal=1 Mode=1\n"
        
        Trafos_Monitor.append(trafo_line_monitor) # Se añade a la lista de monitores de trafos
        
        if der_type == "INV":
            der_line = 'new generator.DER' + derName + ' bus1=' + busLV + '.1.2.3'
            der_line += ' kV=' + kV_LowLL + ' phases=3 kW='
            der_line += str(float(installed_capacity)) + ' PF=1 conn=wye kVA='
            der_line += str(np.round(float(installed_capacity)*1.2,2))
            der_line + ' Model=7 Vmaxpu=1.5 Vminpu='+str(1/der_type_val)+' Balanced=yes Enabled=yes' 
        
        else:
            der_line = 'new generator.DER' + derName + ' bus1=' + busLV + '.1.2.3'
            der_line += ' kV=' + kV_LowLL + ' phases=3 kW='
            der_line += str(float(installed_capacity)) + ' PF=1 conn=wye kVA='
            der_line += str(np.round(float(installed_capacity)*1.2,2))
            der_line + ' Model=1 Xdp='+str(der_type_val)+' Xdpp=0.020 Balanced=no Enabled=yes'
    
        DERs.append(der_line)
        
    return Trafos, Trafos_Monitor, DERs

#%% POWER QUALITY DER INTEGRATION EFFECT (overvoltage and voltage deviation) - verification, in all load buses, by phase. It updates the results data frames for every monte carlo and every generation capacity installation
# Overvoltage means magnitude in any phase goes above 1.05 p.u
# Maximum voltage deviation in MT is 3% (based on IEEE Std. 1453-2015) and in BT is 5% (based on EPRI studies)
# Input: DSScircuit, DSStext, dss_network,  Base_V, loadslv_buses, loadsmv_buses, capacity_i, NoDERsPF_Vbuses, Overvoltage_loads_df, Overvoltage_rest_df, Voltagedeviation_loads_df, Voltagedeviation_rest_df
# Output: Overvoltage_loads_df: pd.DataFrame({ 'DER_kWp':[capacity_i],'BUS':[busi_l],'Voltage' : [vbusi_l], 'POVERVOLTAGE' : [povervi_l] } ), Voltagedeviation_loads_df, Overvoltage_rest_df, Voltagedeviation_rest_df, all have the same format

#%% FIND THREE PHASE VOLTAGE UNBALANCE
def voltageunbalance(DSScircuit, busname):
    DSScircuit.setActiveBus(busname.split('.')[0])
    
    if len(DSScircuit.ActiveBus.Nodes) == 3:
        unbt = np.round((DSScircuit.ActiveBus.SeqVoltages[2] / DSScircuit.ActiveBus.SeqVoltages[1]),4)
    else:
        unbt = np.nan
    return unbt
#%%

def pq_voltage(DSScircuit, DSStext, dss_network, Base_V, loadslv_buses,
               loadsmv_buses, RegDevices, CapDevices, capacity_i, lv_groups, tx_groups,
               NoDERsPF_Vbuses, Overvoltage_loads_df, Overvoltage_rest_df,
               Voltagedeviation_loads_df, Voltagedeviation_rest_df,
               Voltagedeviation_reg_df, Voltageunbalance_df,
               Overvoltage_analysis, VoltageDeviation_analysis,
               VoltageRegulation_analysis, VoltageUnbalance,
               max_v, max_lv_dev, max_mv_dev, max_v_unb, 
               redo_sim, last_blacklisted_ov, last_blacklisted_vd, voltage_vals, overvoltage_vals, report_txt): 

    #%% PU CALCULATION
    t1i=time.time()
    VBuses = pd.DataFrame(list(DSScircuit.AllBusVmag),
                          index=list(DSScircuit.AllNodeNames),
                          columns=['VOLTAGEV'])
    VBuses = VBuses[~VBuses.index.str.contains('aftermeter')]
    VBuses = VBuses[~VBuses.index.str.contains('_der')]
    VBuses=VBuses[~VBuses.index.str.contains('_swt')]
    VBuses=VBuses[~VBuses.index.str.endswith('.4')]
    
    V_buses = pd.DataFrame()
    V_buses['VOLTAGE'] = VBuses.VOLTAGEV/Base_V.BASE
    V_buses['LV_GROUP'] = Base_V.LV_GROUP
    V_buses['TX'] = Base_V.TX
    t1f=time.time()

    print('V_buses_pu: ' + str(t1f - t1i))
    
    #%% OVERVOLTAGE MONITORING - Loads nodes and rest of nodes
    if Overvoltage_analysis is True:
        t2i=time.time()
    
        #loads
        
        V_buses_loads = V_buses[V_buses.index.isin(loadslv_buses + loadsmv_buses)]; #pu loads buses dataframe
        vbusi_l = V_buses_loads.VOLTAGE.max()
        busi_l = V_buses_loads.VOLTAGE.idxmax()
        povervi_l = V_buses_loads.VOLTAGE.gt(max_v).mean()
        #rest
        V_buses_rest = V_buses[~V_buses.index.isin(loadslv_buses + loadsmv_buses)] #pu rest of buses dataframe
        vbusi_r = V_buses_rest.VOLTAGE.max()
        busi_r = V_buses_rest.VOLTAGE.idxmax()
        povervi_r = V_buses_rest.VOLTAGE.gt(max_v).mean()
        
        OVl_List = list(set(V_buses_loads.loc[V_buses_loads['VOLTAGE'].gt(max_v), 'TX'].dropna()))
        OVr_List = list(set(V_buses_rest.loc[V_buses_rest['VOLTAGE'].gt(max_v), 'TX'].dropna()))
        
        temp_OVlr_list = list(set(OVl_List + OVr_List))
        
        # loadsmv_buses
        if not Overvoltage_loads_df.empty:
            OVl_List = list(set(list(Overvoltage_loads_df['BLACKLIST_TX'].values[-1]) + OVl_List))
            OVr_List = list(set(list(Overvoltage_rest_df['BLACKLIST_TX'].values[-1]) + OVr_List))
                
        if V_buses[V_buses['VOLTAGE'].gt(max_v)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if any(V_buses[V_buses['VOLTAGE'].gt(1.05)].index.str.startswith('busmv')) else False  
        
        # Actualización Mayo 2023, paro de simulación debido a problemas en otros secundarios
        if redo_sim:
            for tx in temp_OVlr_list:
                if tx in last_blacklisted_ov:
                    HCSTOP = True
        
        overvoltage_vals[capacity_i] = V_buses
        
        pd_tmp = pd.DataFrame({'DER_kWp':[capacity_i], 'BUS':[busi_l],
                               'VOLTAGE': [vbusi_l],
                               'BLACKLIST_TX' : [OVl_List],
                               'HCSTOP':[HCSTOP]})
        Overvoltage_loads_df = pd.concat([Overvoltage_loads_df, pd_tmp],
                                         ignore_index=True)
        pd_tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'BUS': [busi_r],
                               'VOLTAGE' : [vbusi_r],
                               'BLACKLIST_TX' : [OVr_List],
                               'HCSTOP':[HCSTOP]})
        Overvoltage_rest_df = pd.concat([Overvoltage_rest_df, pd_tmp],
                                        ignore_index=True)
        t2f=time.time()
        msg = 'Análisis de sobretensión: ' +str(np.round(t2f-t2i,2)) + " s"
        print(msg)
        report_txt.append(msg)
    
    #%% VOLTAGE DEVIATION MONITORING (5% in LV, 3% in MV). For regulation nodes, a HCSTOP in 1 is reported if V excceeds 1/2 the bandwith of the regulator
    
    if VoltageDeviation_analysis is True:
        t3i=time.time()
        #loads
        NoDERsPF_MVbuses_loads = NoDERsPF_Vbuses[NoDERsPF_Vbuses.index.isin(loadsmv_buses)] # pu loads buses dataframe at base power flow (no DER)
        NoDERsPF_LVbuses_loads = NoDERsPF_Vbuses[NoDERsPF_Vbuses.index.isin(loadslv_buses)]
        MV_buses_loads = V_buses[V_buses.index.isin(loadsmv_buses)]  # pu loads buses dataframe
        LV_buses_loads = V_buses[V_buses.index.isin(loadslv_buses)]  # pu loads buses dataframe
        devLV_loads = pd.DataFrame(zip(abs((LV_buses_loads.VOLTAGE-NoDERsPF_LVbuses_loads.VOLTAGE)/NoDERsPF_LVbuses_loads.VOLTAGE), Base_V.LV_GROUP, Base_V.TX),
                                   index=NoDERsPF_LVbuses_loads.index,
                                   columns=['VDEVIATION', 'LV_GROUP', 'TX'])
        devMV_loads = pd.DataFrame(zip(abs((MV_buses_loads.VOLTAGE-NoDERsPF_MVbuses_loads.VOLTAGE)/NoDERsPF_MVbuses_loads.VOLTAGE), Base_V.LV_GROUP, Base_V.TX),
                                   index=NoDERsPF_MVbuses_loads.index,
                                   columns=['VDEVIATION', 'LV_GROUP', 'TX'])
    
        #rest of buses
        NoDERsPF_MVbuses_rest = NoDERsPF_Vbuses[~NoDERsPF_Vbuses.index.str.contains('lv')]
        NoDERsPF_MVbuses_rest = NoDERsPF_MVbuses_rest[~NoDERsPF_MVbuses_rest.index.isin(loadsmv_buses)] # pu rest of buses dataframe at base power flow (no DER)
        NoDERsPF_LVbuses_rest = NoDERsPF_Vbuses[NoDERsPF_Vbuses.index.str.contains('lv')]
        NoDERsPF_LVbuses_rest = NoDERsPF_LVbuses_rest[~NoDERsPF_LVbuses_rest.index.isin(loadslv_buses)]
        MV_buses_rest = V_buses[~V_buses.index.str.contains('lv')]
        MV_buses_rest = MV_buses_rest[~MV_buses_rest.index.isin(loadsmv_buses)]  # pu rest of buses dataframe
        LV_buses_rest = V_buses[V_buses.index.str.contains('lv')]
        LV_buses_rest = LV_buses_rest[~LV_buses_rest.index.isin(loadslv_buses)]  # pu rest of buses dataframe
        devLV_rest = pd.DataFrame(zip(abs((LV_buses_rest.VOLTAGE-NoDERsPF_LVbuses_rest.VOLTAGE)/NoDERsPF_LVbuses_rest.VOLTAGE ), Base_V.LV_GROUP, Base_V.TX ),
                                  index=NoDERsPF_LVbuses_rest.index,
                                  columns=['VDEVIATION', 'LV_GROUP', 'TX'])
        devMV_rest = pd.DataFrame(zip(abs((MV_buses_rest.VOLTAGE-NoDERsPF_MVbuses_rest.VOLTAGE)/NoDERsPF_MVbuses_rest.VOLTAGE ), Base_V.LV_GROUP, Base_V.TX ),
                                  index=NoDERsPF_MVbuses_rest.index,
                                  columns=['VDEVIATION', 'LV_GROUP', 'TX'])
        
        VDl_List = list(set(devLV_loads.loc[devLV_loads['VDEVIATION'].gt(max_lv_dev), 'TX'].dropna()))
        VDr_List = list(set(devLV_rest.loc[devLV_rest['VDEVIATION'].gt(max_lv_dev), 'TX'].dropna()))
        temp_VDlr_list = list(set(VDl_List+ VDr_List)) # Secundarios blacklisteados de esta iteración
        
        if not Voltagedeviation_loads_df.empty:
            VDl_List = list(set(list(Voltagedeviation_loads_df['BLACKLIST_TX'].values[-1]) + VDl_List))
            VDr_List = list(set(list(Voltagedeviation_rest_df['BLACKLIST_TX'].values[-1]) + VDr_List))
        
        devLV_all = pd.concat([devLV_loads, devLV_rest])
        devMV_all = pd.concat([devMV_loads, devMV_rest])
        
        if devLV_all[devLV_all['VDEVIATION'].gt(max_lv_dev)].index.empty and devMV_all[devMV_all['VDEVIATION'].gt(max_mv_dev)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if not devMV_all[devMV_all['VDEVIATION'].gt(max_mv_dev)].index.empty else False  
        
        # 2023: se agrega esta verificación para contener el problema de desviación de tensión
        # en las siguientes iteraciones.
        if redo_sim:
            for tx in temp_VDlr_list:
                if tx in last_blacklisted_vd:
                    HCSTOP = True
                    
        voltage_vals[capacity_i] = devLV_all            
        max_lvdev_l = devLV_loads.VDEVIATION.max()
        buslvi_l=devLV_loads.VDEVIATION.idxmax()
        pdelvvi_l = devLV_loads.VDEVIATION.gt(max_lv_dev).mean()
        if not devLV_rest.empty:
            max_lvdev_r = devLV_rest.VDEVIATION.max()
            buslvi_r=devLV_rest.VDEVIATION.idxmax()
            pdelvvi_r = devLV_rest.VDEVIATION.gt(max_lv_dev).mean()
        else:
            max_lvdev_r = np.nan
            buslvi_r=np.nan
            pdelvvi_r = np.nan
        max_mvdev_r = devMV_rest.VDEVIATION.max()
        busmvi_r =devMV_rest.VDEVIATION.idxmax()
        pdemvvi_r = devMV_rest.VDEVIATION.gt(max_mv_dev).mean()
        
        if not devMV_loads.empty:
            max_mvdev_l = devMV_loads.VDEVIATION.max()
            busmvi_l = devMV_loads.VDEVIATION.idxmax()
            pdemvvi_l = devMV_loads.VDEVIATION.gt(max_mv_dev).mean()
            tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'BUSLV': [buslvi_l],
                                'VDEVIATIONLV' : [max_lvdev_l],
                                'BUSMV':[busmvi_l],
                                'VDEVIATIONMV' : [max_mvdev_l],
                                'BLACKLIST_TX' : [VDl_List],
                                'HCSTOP':[HCSTOP]})
            Voltagedeviation_loads_df = pd.concat([Voltagedeviation_loads_df, tmp],
                                                  ignore_index=True)
        else:
            tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'BUSLV':[buslvi_l],
                                'VDEVIATIONLV': [max_lvdev_l],
                                'BLACKLIST_TX': [VDl_List],
                                'HCSTOP': [HCSTOP]})
            Voltagedeviation_loads_df = pd.concat([Voltagedeviation_loads_df, tmp],
                                                  ignore_index=True)
    
        tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'BUSLV': [buslvi_r],
                            'VDEVIATIONLV' : [max_lvdev_r],
                            'VDEVIATIONMV' : [max_mvdev_r],
                            'BLACKLIST_TX' : [VDr_List],
                            'HCSTOP':[HCSTOP] } )
        Voltagedeviation_rest_df = pd.concat([Voltagedeviation_rest_df, tmp],
                                             ignore_index=True)
        
        t3f=time.time()
        msg = 'Análisis de desviación de tensión: ' +str(np.round(t3f-t3i,2))+ " s"
        print(msg)
        report_txt.append(msg)
        
    #%% Voltage deviation at regulator devices (autotransformers and capacitors)
    if VoltageRegulation_analysis is True:
        t4i = time.time()
        
        try:
            max_v_dif = 0
            max_bus = ""
            HCSTOP = False
            
            try:
                for reg in list(RegDevices['BUSINST']):
                    buslist = [reg.split('.')[0].lower()+'.'+ph for ph in reg.split('.')[1:]]
                    bndwdth = RegDevices.loc[RegDevices['BUSINST']==reg, 'BANDWIDTH'].iloc[-1] 
                    vreg = RegDevices.loc[RegDevices['BUSINST']==reg, 'VREG'].iloc[-1]
                    temp_max_v_dif_reg = max([abs(NoDERsPF_Vbuses['VOLTAGE'][busph] - V_buses['VOLTAGE'][busph]) for busph in buslist])
                    if  temp_max_v_dif_reg > ((vreg + bndwdth/2)/vreg)-1 : 
                        HCSTOP= True
                    else: 
                        HCSTOP= False
                    
                    if temp_max_v_dif_reg  > max_v_dif:
                        max_v_dif = temp_max_v_dif_reg
                        max_bus = reg
            except:
                pass
                
            try:
                for cap in list(CapDevices['BUSINST']):
                    buslist = [cap.split('.')[0].lower()+'.'+ph for ph in cap.split('.')[1:]]
                    bndwdth = CapDevices.loc[CapDevices['BUSINST']==cap, "OBJ_MAX"].iloc[-1] - CapDevices.loc[CapDevices['BUSINST']==cap, "OBJ_MIN"].iloc[-1] # Los dos valores están en PU
                    temp_max_v_dif_cap = max([abs(NoDERsPF_Vbuses['VOLTAGE'][busph] - V_buses['VOLTAGE'][busph]) for busph in buslist])
                    if  temp_max_v_dif_cap > (bndwdth/2): 
                        HCSTOP= True
                    else: 
                        HCSTOP= False
                    
                    if temp_max_v_dif_cap  > max_v_dif:
                        max_v_dif = temp_max_v_dif_cap
                        max_bus = cap
            except:
                pass

            tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'BUSREG': [max_bus],
                                'VDIF':[max_v_dif], 'HCSTOP': [HCSTOP]})
            
            Voltagedeviation_reg_df = pd.concat([Voltagedeviation_reg_df, tmp],
                                                ignore_index=True)
        except:
            pass
        
        t4f=time.time()
        msg = 'Análisis de regulación de tensión: ' +str(np.round(t4f-t4i,2))+ " s"
        print(msg)
        report_txt.append(msg)
        
    
    #%% VOLTAGE UNBALANCE
    if VoltageUnbalance is True:
        t5i = time.time()
        # Greater of all three phase nodes
        buseslist = pd.DataFrame(columns=['BUS','LV_GROUP', 'TX', 'UNBALANCE'])
        buseslist['BUS'] = list(V_buses.index)
        buseslist['LV_GROUP'] = list(V_buses.LV_GROUP)
        buseslist['TX'] = list( V_buses['TX'])
        buseslist.index = buseslist['BUS']
        buseslist['UNBALANCE'] = buseslist['BUS'].apply(lambda b: voltageunbalance(DSScircuit, b))
        maxunb=buseslist.UNBALANCE.max()
        busunb=buseslist.UNBALANCE.idxmax()
        buseslist = buseslist[buseslist.UNBALANCE.ge(max_v_unb)]
    
        VU_List = list(set(buseslist.TX.dropna()))
        if buseslist.empty:
            HCSTOP = False
            if not Voltageunbalance_df.empty: 
                VU_List = list(set(list(Voltageunbalance_df['BLACKLIST_TX'].values[-1]) + VU_List))
        else:
            if any(buseslist.BUS.str.startswith('busmv')):
                HCSTOP=True
            else:
                if not Voltageunbalance_df.empty:
                    VU_List = list(set(list(Voltageunbalance_df['BLACKLIST_TX'].values[-1]) + VU_List))
                    HCSTOP = False
    
        tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'BUS': [busunb],
                            'UNBALANCE': [maxunb],'BLACKLIST_TX': [VU_List],
                            'HCSTOP': [HCSTOP]})
        Voltageunbalance_df = pd.concat([Voltageunbalance_df, tmp],
                                        ignore_index=True)
        t5f=time.time()
        msg = 'Procesamiento de desbalance de tensión: ' +str(np.round(t5f-t5i,2))+ " s"
        print(msg)
        report_txt.append(msg)
    
    if Overvoltage_analysis and VoltageDeviation_analysis:
        pass
    elif Overvoltage_analysis and not VoltageDeviation_analysis:
        temp_VDlr_list = []
        voltage_vals = pd.DataFrame()
    elif not Overvoltage_analysis and VoltageDeviation_analysis:
        temp_OVlr_list = []
        overvoltage_vals = pd.DataFrame()
    else:
        temp_OVlr_list = []
        temp_VDlr_list = []
        voltage_vals = pd.DataFrame()
        overvoltage_vals = pd.DataFrame()
        
    return VBuses, Overvoltage_loads_df, Voltagedeviation_loads_df, Overvoltage_rest_df, Voltagedeviation_rest_df, Voltagedeviation_reg_df, Voltageunbalance_df,  temp_OVlr_list, temp_VDlr_list, voltage_vals, overvoltage_vals, report_txt

#%% LINE LENGTH ASSIGNATION - dataframe.apply function
# Input: DSScircuit, name, normalAmpsDic
# Output: line's length
def line_length(DSScircuit, name):
    DSScircuit.SetActiveElement('line.' + name)
    return float(DSScircuit.Lines.length)

#%% LINE LENGTH ASSIGNATION - dataframe.apply function
# Input: DSScircuit, name, normalAmpsDic
# Output: line's length

def line_current(DSScircuit, name, normalAmpsDic):
    DSScircuit.SetActiveElement('line.' + name)
    lineAmps = DSScircuit.ActiveCktElement.CurrentsMagAng  # read line currents
    meanCurrent = np.max([lineAmps[x] for x in range(0, int(len(lineAmps)/2), 2)]) / normalAmpsDic[str(name.upper())]
    return float(meanCurrent)

#%% LINE LV GROUP ASSIGNATION - dataframe.apply function
# Input: g: line name,line_lv_groups
# Output: line's group or nan for MV lines

def linelvgroup(g,line_lv_groups):
    try:
        group=int(line_lv_groups.loc[str(g).upper()]['LV_GROUP'])
    except: 
        group=np.nan
    return group

#%% THERMAL IMPACT IN LINES AND TRANSFORMERS - Checks if any line of transformer has a loading over 1.00 p.u 
# Input: DSScircuit, normalAmpsDic,  capacity_i, Thermal_loading_lines_df, Thermal_loading_tx_df
# Output: Thermal_loading_lines_df : pd.DataFrame({ 'DER_kWp':[capacity_i],'Line':[txi],'LOADING' : [max_loading], 'POVERLOADTX' : [ptloadingi] } ), Thermal_loading_tx_df : pd.DataFrame({ 'DER_kWp':[capacity_i],'Line':[txi],'LOADING' : [max_loading], 'POVERLOADTX' : [ptloadingi] } )

def thermal_Lines_Tx(DSScircuit, DSStext, normalAmpsDic,
                     capacity_i, line_lv_groups, tx_groups,
                     Thermal_loading_lines_df, Thermal_loading_tx_df,
                     name_file_created, linelvgroups, Thermal_analysis, report_txt): # Most code from auxfns.lineCurrents
    
    if Thermal_analysis is True: 
        # LINES
        start = time.time()
        
        tmp = 'mv3p' + name_file_created.split('_')[0].lower() + '00'
        CurrentDF = pd.DataFrame()
        CurrentDF['DSSNAME'] = DSScircuit.Lines.AllNames
        CurrentDF=CurrentDF[~CurrentDF.DSSNAME.str.contains('swt')]
        CurrentDF=CurrentDF[~CurrentDF.DSSNAME.str.contains(tmp)]
        CurrentDF['DSSNAME'].astype(str)
        CurrentDF['CURRENT'] = CurrentDF.DSSNAME.apply(lambda c : line_current(DSScircuit, c, normalAmpsDic))
        CurrentDF['LENGTH'] = CurrentDF.DSSNAME.apply(lambda c : line_length(DSScircuit, c))
    
        if len(linelvgroups) == 0: 
            CurrentDF['LV_GROUP'] = CurrentDF.DSSNAME.apply(lambda g : linelvgroup(g, line_lv_groups) )
            linelvgroups=list(CurrentDF['LV_GROUP'])
        CurrentDF['LV_GROUP'] = linelvgroups
        CurrentDF['TX']= CurrentDF.LV_GROUP.apply(lambda g : txname(g,tx_groups) )
        CurrentDF.index = CurrentDF['DSSNAME']#; del CurrentDF['DSSNAME']
        
        max_loading = CurrentDF.CURRENT.max()
        linei=CurrentDF.CURRENT.idxmax()
        lenloadingi = CurrentDF.loc[CurrentDF['CURRENT'] > 1.00, 'LENGTH'].sum()
        
        LO_List = list(set(CurrentDF.loc[CurrentDF['CURRENT'].gt(1.00), 'TX'].dropna()))
        if not Thermal_loading_lines_df.empty:
            LO_List = list(set(list(Thermal_loading_lines_df['BLACKLIST_TX'].values[-1]) + LO_List))
    
        if CurrentDF[CurrentDF['CURRENT'].gt(1.00)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if any(CurrentDF[CurrentDF['CURRENT'].gt(1.00)].index.str.startswith('mv')) else False
    
        tmp = pd.DataFrame({'DER_kWp':[capacity_i], 'LINE': [linei],
                            'CURRENT' : [max_loading],
                            'LENGTHOVERLOADED' : [lenloadingi],
                            'BLACKLIST_TX' : [LO_List], 'HCSTOP':[HCSTOP]})
        Thermal_loading_lines_df = pd.concat([Thermal_loading_lines_df, tmp],
                                             ignore_index=True)
    
        # TX
        trafosPowers = {}
        trafosDict = {}
    
        DSScircuit.Transformers.First
        nextTx = 1
        while nextTx != 0:  # results writing by transformer
            temp = 0
            trafo = DSScircuit.Transformers.Name # get transformer name
            trafo = str(trafo).upper()
            DSScircuit.SetActiveElement('transformer.' + trafo)  # set active transformer
            capacity = int(DSScircuit.Transformers.kva)  # transformer capacity
            if 'auto' in trafo.lower():
                capacity = capacity * 3
            if len(DSScircuit.ActiveElement.Powers) >= 16: # 3ph trafo of two or three terminals
                temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
                p=temp_powers[0] + temp_powers[2] + temp_powers[4]
                q=temp_powers[1] + temp_powers[3] + temp_powers[5]
                valor = np.sqrt(p**2+q**2)
                temp = np.abs(valor) / capacity
            else:
                temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
                p=temp_powers[0] 
                q=temp_powers[1] 
                valor = np.sqrt(p**2+q**2)
                temp = np.abs(valor) / capacity
    
            # for 3-units 3ph / 2-units 3ph transformers there is only one point representing it
            # this code writes the worst case for the unit
            tmp = name_file_created.split('_')[0].upper()
            if 'U' in trafo.replace(tmp, '') and 'auto' not in str(trafo.lower()):
                trafo = trafo.split('_')[0] + '_' + trafo.split('_')[2]
                try:
                    if np.mean(trafosDict[trafo]) > np.mean(temp):
                        temp = trafosDict[trafo]
                except KeyError:
                    pass
            trafosPowers.update({trafo:[]})
            trafosPowers[trafo]=temp_powers
        
            trafosDict[trafo] = float(temp)
            nextTx = DSScircuit.Transformers.Next  # set active next transformer
        DSStext.Command = 'export powers KVA elem'
        TxDF = pd.DataFrame(list(trafosDict.values()), index=list(trafosDict.keys()), columns=['LOADING'])
    
        max_loading = TxDF.LOADING.max(); txi=TxDF.LOADING.idxmax(); ptloadingi = TxDF.LOADING.gt(1.00).mean()
        
        TO_List = list(set(TxDF.loc[TxDF['LOADING'].gt(1.00)].index))
        if not Thermal_loading_tx_df.empty:
            TO_List = list(set(list(Thermal_loading_tx_df['BLACKLIST_TX'].values[-1]) + TO_List))
    
        if TxDF[TxDF['LOADING'].gt(1.00)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if len(TO_List) == len(set(tx_groups.LV_GROUP)) else False
    
        tmp = pd.DataFrame({'DER_kWp': [capacity_i], 'TX':[txi], 'LOADING' : [max_loading], 'BLACKLIST_TX': [TO_List],
               'HCSTOP': [HCSTOP]})
        Thermal_loading_tx_df = pd.concat([Thermal_loading_tx_df, tmp],
                                          ignore_index=True)
        
        end = time.time()
        sim_time = end - start
        msg = 'Análisis térmico: '+str(np.round(sim_time,2))+' s'
        print(msg)
        report_txt.append(msg)
    
    else:
        CurrentDF = pd.DataFrame()
    
    return Thermal_loading_lines_df, Thermal_loading_tx_df, CurrentDF, report_txt

#%% MV GRAPH CREATION
# Input: nodos_mt, lineas_mt, trafos
# Output: G -> graph
def circuit_graph(nodes_mv, lines_mv):
    
    start = time.time()
    
    G = nx.DiGraph()
    
    #circuit name identification from layer info
    # rand_bus = lines_mv.loc[0,'bus1']
    # ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    
    # Información y atributos de nodos
    for idx in nodes_mv.index:
        
        bus = nodes_mv.loc[idx,'BUS']
        phase_num = nodes_mv.loc[idx,'PHASES']
        node_conn = nodes_mv.loc[idx,'NODES']
            
        G.add_node(bus, num_phases=phase_num, conn=node_conn)
    
    nodes_mv.index = nodes_mv.loc[:,'BUS'] #ordena el dataframe por buses
    
    #Información de líneas
    for idx in lines_mv.index:
        bus1 = lines_mv.loc[idx,'bus1']
        bus2 = lines_mv.loc[idx,'bus2']
        dist = lines_mv.loc[idx, 'LENGTH']
        name = lines_mv.loc[idx, 'DSSNAME']
        G.add_edge(bus1, bus2, distance=dist, line_name=name)
    
    end = time.time()
    sim_time = end - start
    print('Graph time: '+str(np.round(sim_time,2))+' sec.')
    
    return G


#%% MV GRAPH CREATION
# Input: nodos_mt, lineas_mt, trafos
# Output: G -> graph
def circuit_graph_mv(nodes_mv, lines_mv):    
    start = time.time()
    
    G = nx.DiGraph()
    
    #circuit name identification from layer info
    # rand_bus = lines_mv.loc[0,'bus1']
    # ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    
    # Información y atributos de nodos
    for idx in nodes_mv.index:        
        bus = nodes_mv.loc[idx, 'BUS']
        node_conn = nodes_mv.loc[idx, 'NODES']
        #point = nodes_mv.loc[idx, 'geometry']
        volts = nodes_mv.loc[idx, 'BASEKV_LN']
        phases = nodes_mv.loc[idx, 'PHASES']
        #x_ = point.x
        #y_ = point.y
        # posic = (x_, y_)
        G.add_node(bus, conn=node_conn, voltage=volts, num_phases=phases)
    
    nodes_mv.index = nodes_mv.loc[:, 'BUS'] #ordena el dataframe por buses
    
    #Información de líneas
    for idx in lines_mv.index:
        bus1 = lines_mv.loc[idx, 'bus1']
        bus2 = lines_mv.loc[idx, 'bus2']
        nomvolt = lines_mv.loc[idx, 'NOMVOLT']
        dist = lines_mv.loc[idx, 'LENGTH']
        name = lines_mv.loc[idx, 'DSSNAME']
        G.add_edge(bus1, bus2, distance=dist, line_name=name, bus1=bus1, bus2=bus2)
        nodes_mv.loc[bus1, 'NOMVOLT'] = nomvolt
        nodes_mv.loc[bus2, 'NOMVOLT'] = nomvolt
        G.nodes[bus1].update({'NOMVOLT': nomvolt})
        G.nodes[bus2].update({'NOMVOLT': nomvolt})
    
    end = time.time()
    sim_time = end - start
    print('Graph lv time: '+ str(np.round(sim_time,2)) +' sec.')
    
    return G
#%% FINDS THE ELEMENT OF THE PROTECTION DEVICE - dataframe.apply function
# Input: pdevice, ini_bus: substation bus, G
# Output: element name

def find_element(pdevice, ini_bus, firstLine, G):

    us_bus = [n for n in G.predecessors(pdevice)][0] #Upstream bus
    
    if us_bus == ini_bus:
        # ckt_name = ini_bus.split(ini_bus[min([i for i, c in enumerate(ini_bus) if c.isdigit()])])[0].split('BUSMV')[1]
        ckt_name = firstLine.split("MV3P")[1].split("0")[0]
        elem = 'line.MV3P'+ckt_name+'0'
    else:
        elem = 'line.'+G.edges[us_bus, pdevice]['line_name']
    return elem

#%% FINDS PHASE CONNECTION FOR EVERY PROTECTION DEVICE BUS - dataframe.apply function
# Input: pdevice, ini_bus, G
# Output: phase connection e.g. 1.2.3

def find_conn(pdevice, ini_bus, G):
    if pdevice in ['AFTERMETER', ini_bus]:
        conn = '.1.2.3'
    else:
        conn_bus = G.nodes[pdevice]['conn']
        conn = ''
        for x in conn_bus:
            conn += '.'+x
    
    return conn

#%% FINDS RECLOSER ELEMENT WHEN THERE'S A FUSE-RECLOSER COORDINATION - dataframe.apply function
# Input: pdevice, fusibles
# Output: recloser bus 

def find_recloserelem(G, pdevice, fusibles, reclosers):
    if pdevice.split(".")[0] in list(fusibles['bus2']):
        fuse_idx = fusibles[fusibles['bus2']==pdevice.split(".")[0] ].index.values[0] #index to loc
        
        if (fusibles.loc[fuse_idx, 'SAVE'].upper() == 'SI') or (fusibles.loc[fuse_idx, 'SAVE'].upper() == 'YES'):
            recloser_ID = fusibles.loc[fuse_idx, 'COORDINATE']
            recloser_idx = reclosers.loc[reclosers['DSSNAME'] == recloser_ID].index.values[0]
            recloser_bus1 = reclosers.loc[recloser_idx, 'bus1']
            recloser_bus2 = reclosers.loc[recloser_idx, 'bus2']
            recloser_elem = 'line.'+G.edges[recloser_bus1, recloser_bus2]['line_name']
        else:
            recloser_elem = np.nan
    else:
        recloser_elem = np.nan
    
    return recloser_elem

#%% PROTECTION DEVICES DATAFRAME CREATION
# Input: pdevice, ini_bus, G
# Output: CircuitBreakDvRoR, CircuitBreakDvFF_BFC

def pDevices(G, firstLine, fusibles, reclosers, FF_analysis, BFC_analysis): 

    # Se modifica internamente el dataframe de fusibles para que solo considere el fusible que tiene recloser asociado
    if (BFC_analysis is True) and (FF_analysis is False):
        fusibles = fusibles[~fusibles["SAVE"].isin([None, "No", "NO", "NULL", "", " "])]
    else:
        pass

    bus_list_FF_BFC = list(set([(fusibles.loc[x, 'bus1'], fusibles.loc[x, 'bus2']) for x in fusibles.index] + [(reclosers.loc[y, 'bus1'], reclosers.loc[y, 'bus2']) for y in reclosers.index]))
    bus_list_DvRoR = [(reclosers.loc[y, 'bus1'], reclosers.loc[y, 'bus2']) for y in reclosers.index]

    #bus de inicio: 
    ini_bus = 'AFTERMETER'
    
    #Dataframe initialization 
    # pdevice_df = pd.DataFrame(np.nan, index=devices_bus_list, columns=['Element', 'Installed_bus', 'Upstream_bus', 'Downstream_bus', 'Furthest_bus', , 'Furthest_zone_bus','Furthest_zone_bus_dist'])

    p_RoR_list = [bus_list_DvRoR[x][1] for x in range(len(bus_list_DvRoR))]
    CircuitBreakDvRoR = pd.DataFrame(np.nan, columns=['Element', 'BusInstalled', 'FurthestBusZone', 'Furthest_zone_bus_dist', 'FurthestBusBackUp', 'Furthest_bus_dist'], index=p_RoR_list) # Reclosers nada más
    
    
    #%% DvRoR dataframe ###############################################
    
    mt_3bus_list = nx.get_node_attributes(G,'num_phases') #list of 3phase mv buses on the whole circuit
    remove3 = [k for k in mt_3bus_list if mt_3bus_list[k] != '3']
    for k in remove3: del mt_3bus_list[k]
    
    mt_1bus_list = nx.get_node_attributes(G,'num_phases') #list of 3phase mv buses on the whole circuit
    remove1 = [k for k in mt_1bus_list if mt_1bus_list[k] != '1']
    for k in remove1: del mt_1bus_list[k]
    
    #First, recognize which buses are downstream to the pdevice bus
    downstream_buses_pdevice = {}
    start = time.time()
    start1 = time.time()
    #element definition    
    
    for pdevice in p_RoR_list:
        downstream_buses_pdevice[pdevice] = {}
        # distance and descendants dataframe per pdevice
        downstream_buses_pdevice[pdevice]['dataframe'] = pd.DataFrame()#np.nan, index=[0], columns=['descendants', 'distance']) 
        if G.nodes[pdevice]['num_phases'] == '3':
            downstream_buses_pdevice[pdevice]['dataframe']['descendants'] = list(set(mt_3bus_list).intersection(nx.descendants(G,pdevice)))
        elif G.nodes[pdevice]['num_phases'] == '1':
            downstream_buses_pdevice[pdevice]['dataframe']['descendants'] = list(set(mt_1bus_list).intersection(nx.descendants(G,pdevice)))
        
        downstream_buses_pdevice[pdevice]['dataframe']['distance']=downstream_buses_pdevice[pdevice]['dataframe']['descendants'].apply(lambda b: nx.shortest_path_length(G, pdevice, b, weight='distance'))
        
        #Furthest_general_bus
        max_gnal_idx = downstream_buses_pdevice[pdevice]['dataframe']['distance'].idxmax()
        back_up_bus = downstream_buses_pdevice[pdevice]['dataframe'].loc[max_gnal_idx, 'descendants']
        
        downstream_buses_pdevice[pdevice]['FurthestBusBackUp'] = back_up_bus + find_conn(pdevice, ini_bus, G)
        downstream_buses_pdevice[pdevice]['Furthest_bus_dist'] = downstream_buses_pdevice[pdevice]['dataframe'].loc[max_gnal_idx, 'distance']   

        # pdevice element characteristics
        downstream_buses_pdevice[pdevice]['Element'] = find_element(pdevice, ini_bus, firstLine, G)
        downstream_buses_pdevice[pdevice]['BusInstalled'] = pdevice + find_conn(pdevice, ini_bus, G)
            
    # Furthest_bus_zone 
    for pdevice in p_RoR_list:
        for pdevice_2 in list(set(p_RoR_list) - set(pdevice)):
            if pdevice_2 in downstream_buses_pdevice[pdevice]['dataframe']['descendants'].values:
                # print(pdevice, pdevice_2)
                downstream_buses_pdevice[pdevice]['dataframe'] = downstream_buses_pdevice[pdevice]['dataframe'][~downstream_buses_pdevice[pdevice]['dataframe']['descendants'].isin(downstream_buses_pdevice[pdevice_2]['dataframe']['descendants'])]
            
        max_zone_idx = downstream_buses_pdevice[pdevice]['dataframe']['distance'].idxmax()
        fur_zone_bus = downstream_buses_pdevice[pdevice]['dataframe'].loc[max_zone_idx, 'descendants']
                
        downstream_buses_pdevice[pdevice]['FurthestBusZone'] = fur_zone_bus+find_conn(pdevice, ini_bus, G)
        downstream_buses_pdevice[pdevice]['Furthest_zone_bus_dist'] = downstream_buses_pdevice[pdevice]['dataframe'].loc[max_zone_idx, 'distance']
    
    CircuitBreakDvRoR = pd.DataFrame.from_dict(downstream_buses_pdevice, orient='index'); del CircuitBreakDvRoR['dataframe']
    
    
    end = time.time()
    sim_time = end - start1
    print('Primera parte: '+str(np.round(sim_time,5))+' sec.')
        
    #%% FF_BFC dataframe ###############################################  
      
    start2 = time.time()
    p_FF_list = [bus_list_FF_BFC[x][1] for x in range(len(bus_list_FF_BFC))]
    CircuitBreakDvFF_BFC = pd.DataFrame(np.nan, columns=['Element', 'BusInstalled', 'RecloserElement'], index=range(len(p_FF_list))) #Si de fusibles y reclosers
    CircuitBreakDvFF_BFC['BusInstalled'] = p_FF_list
    
    CircuitBreakDvFF_BFC['Element'] = CircuitBreakDvFF_BFC['BusInstalled'].apply(lambda b : find_element(b, ini_bus, firstLine, G))
    if not fusibles.empty:
        CircuitBreakDvFF_BFC['RecloserElement'] = CircuitBreakDvFF_BFC['BusInstalled'].apply(lambda b : find_recloserelem(G, b, fusibles, reclosers))
    
    CircuitBreakDvFF_BFC['BusInstalled'] = CircuitBreakDvFF_BFC['BusInstalled'].apply(lambda b : b+find_conn(b, ini_bus, G))
    
    end = time.time()
    sim_time = end - start2
    print('Segunda parte: '+str(np.round(sim_time,2))+' sec.')
        
    return CircuitBreakDvRoR, CircuitBreakDvFF_BFC
#%% PROTECTION DEVICES VERIFICATION - Forward fault current increase and Breaker Fuse descoordination
# Iverter contribution to fault current: 1.2 (120% according to SANDIA Protections report)
# Records the fault current increase for each protection device (Checks if fault current increases by 10% respecting the no DER scenario)
# Checks breakers (excluding substation one), reclosers and fuses. All located in MV.
# Checks if there's a scheme of breaker fuse coordination, finds the difference between change in breaker current and fuse current
# Input: DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, kW_sim, kVAr_sim,  capacity_i, Trafos, DERs, No_DER_FFCurrents, FFCurrents, No_DER_BFCCurrents, BFCCurrents, faulttypes
# Output: FFCurrents, BFCCurrents

def FF_BFC_Current(DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, circuit_demand, kW_sim, kVAr_sim,  
                    capacity_i, Trafos, DERs, No_DER_FFCurrents, FFCurrents, No_DER_BFCCurrents, BFCCurrents, faulttypes, 
                    FF_analysis, BFC_analysis, max_increase_ff, max_increase_bfc, report_txt):
    
    if (FF_analysis or BFC_analysis) is True:
        start = time.time()
        #%% Calculate the hour and date in the simulation
        h, m = snapshottime.split(':')
        if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
            if int(m) <= 7:
                m = '00'
            elif int(m) <= 22:
                m = '15'
            elif int(m) <= 37:
                m = '30'
            elif int(m) <= 52:
                m = '45'
            else:
                m = '00'
                h = str(int(h) + 1)
                if int(h) == 24:  # last round on 23:45
                    h = '23'
                    m = '45'
    
        snapshottime = h + ':' + m
        day_ = snapshotdate.replace('/', '')
        day_ = day_.replace('-', '')
        daily_strtime = str(day_ + snapshottime.replace(':', ''))         
        hora_sec = snapshottime.split(':')

        # P and Q to match
        V_to_be_matched = 0
        for ij in range(len(circuit_demand)):
            temp_a = circuit_demand[ij][0]  # day
            temp_b = circuit_demand[ij][1]  # hour                    
            if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
                V_to_be_matched = circuit_demand[ij][4]  # Voltage
                break
        
        #%% Fault type studies
        # defines the terminals necessary for each type of fault. Dictionary containing: Fault type: [terminal conection for bus 1, terminal conection for bus2, number of phases]
        faulttypeterminals = {'ABC':['.1.2','.3.3', '2'], 'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'], 'BC':['.2','.3','1'], 'AC':['.1','.3','1'], 'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'], 'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'], 'BG':['.2','.0','1'], 'CG':['.3','.0','1']}
        HCSTOP = False
        if FFCurrents.empty:
            columnslist=list([ 'DER_kWp', 'Element', 'FaultedBus']+faulttypes+['HCSTOP'])
            FFCurrents = pd.DataFrame(columns=columnslist)
    
        if BFCCurrents.empty:
                columnslist=list([ 'DER_kWp', 'FuseElement', 'BreakerElement','FaultedBus']+faulttypes+['HCSTOP'])
                BFCCurrents = pd.DataFrame(columns=columnslist)
    
        for element, faultedbus in list(No_DER_FFCurrents.index):
            ftlist=faultedbus.split('.')
            if len(ftlist) == 2:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2], 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,2], 'AG':[0], 'BG':[0], 'CG':[0]}
                if ftlist[1] == '1':
                    faultlabels = ['AG']
                elif ftlist[1] == '2':
                    faultlabels = ['BG']
                else: 
                    faultlabels = ['CG']
            elif len(ftlist) == 3:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2], 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,2], 'AG':[0], 'BG':[0], 'CG':[0]}
                if ftlist[1] == '1' and ftlist[2] == '2':
                    faultlabels = ['AB', 'ABG','AG','BG']
                elif ftlist[1] == '2' and ftlist[2] == '3':
                    faultlabels = ['BC', 'BCG','CG','BG']
                else: 
                    faultlabels = ['AC', 'ACG','AG','CG']
            else:
                currentspositions = {'ABC':[0,2,4], 'ABCG':[0,2,4], 'AB':[0,2], 'BC':[2,4], 'AC':[0,4], 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,4], 'AG':[0], 'BG':[2], 'CG':[4]}
                faultlabels = list(faulttypeterminals.keys())
    
            faultlabels = list(set(faulttypes)&set(faultlabels)) # removes those types of faults not being evaluated
    
            temp_row_FF = []
            temp_row_FF.append(capacity_i)
            temp_row_FF.append(element)
            temp_row_FF.append(faultedbus)
    
            if element in No_DER_BFCCurrents.index.get_level_values('FuseElement'):
                breakerelement = No_DER_BFCCurrents.xs(element).index.get_level_values('BreakerElement')[-1]
                temp_row_BFC = []
                temp_row_BFC.append(capacity_i)
                temp_row_BFC.append(element)
                temp_row_BFC.append(breakerelement)
                temp_row_BFC.append(faultedbus)
            else:
                breakerelement = np.nan
            
            for faulttype in faulttypes:
                if faulttype not in faultlabels:
                    temp_row_FF.append(np.nan)
                    try:
                        temp_row_BFC.append(np.nan)
                    except:
                        pass
                else:
                    # Machines inicialization simulation
                    DSStext.Command = 'clear'
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=daily'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                    DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
                    DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                    DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched) # Tensión en cabecera
                    for tx in Trafos:
                        DSStext.Command = tx
                    for der in DERs:
                        DSStext.Command = der
    
                    DSScircuit.Solution.Solve()
    
                    #%% Fault study (dynamics) simulation
                    # type of fault
                    term1 = faulttypeterminals[faulttype][0]
                    term2 = faulttypeterminals[faulttype][1]
                    phasessc = faulttypeterminals[faulttype][2]
                    
                    DSStext.Command = 'Solve mode=dynamic stepsize=0.00002'
                    
                    DSStext.Command = 'new fault.fault_'+ faulttype+' bus1='+ftlist[0]+term1+' bus2='+ftlist[0]+term2+' phases='+phasessc
    
                    DSScircuit.Solution.Solve()
    
                    DSScircuit.setActiveElement(element) # faulted element
                    temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                    maxIchange = max(list(np.divide([temp_currents[i] for i in currentspositions[faulttype]], No_DER_FFCurrents[faulttype][element,faultedbus])))
                    temp_row_FF.append(maxIchange)
                    HCSTOP = False if maxIchange < max_increase_ff else True
                    #BFC
                    
                    if type(breakerelement) is str:
                        fuseD = np.array([temp_currents[i] for i in currentspositions[faulttype]]) - np.array(No_DER_BFCCurrents[faulttype][element,breakerelement,faultedbus][0])
                        DSScircuit.setActiveElement(breakerelement)
                        breaker_currents = DSScircuit.ActiveElement.CurrentsMagAng
                        breakerD = np.array([breaker_currents[i] for i in currentspositions[faulttype]]) - np.array(No_DER_BFCCurrents[faulttype][element,breakerelement,faultedbus][1])
                        maxIdelta=round(max(abs(fuseD-breakerD)),2)
                        temp_row_BFC.append(maxIdelta)
                        HCSTOP = False if maxIdelta < max_increase_bfc else True
            
            if type(breakerelement) is str:
                tmp = pd.DataFrame([temp_row_BFC+[HCSTOP]], index=[0],
                                   columns=list(BFCCurrents.columns))
                
                BFCCurrents = pd.concat([BFCCurrents, tmp], ignore_index=True)
            
            tmp = pd.DataFrame([temp_row_FF+[HCSTOP]], index=[0],
                               columns=list(FFCurrents.columns))
            FFCurrents = pd.concat([FFCurrents, tmp], ignore_index=True)
            
        end = time.time()
        sim_time = end - start
        msg = "Análisis de aumento de corriente de falla y coordinación recloser/fusible " +str(np.round(sim_time,2))+' s'
        print(msg)
        report_txt.append(msg)
        
    
    return FFCurrents, BFCCurrents, report_txt

#%% PROTECTION DEVICES VERIFICATION - Reduction of reach verification
# Iverter contribution to fault current: 1.2 (120% according to SANDIA Protections report)
# Records the current decrease for each breaker or recloser element to a fault far in its protection zone and farthest in the circuit. If the decrease is greater that 10%, there's a ROR problem.
# Input: DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, kW_sim, kVAr_sim,  capacity_i, Trafos, DERs, No_DER_RoRCurrents, RoRCurrents, faulttypes
# Output: RORCurrents
def ReductionReach(DSScircuit, DSStext, dss_network, firstLine,
                   snapshottime, snapshotdate, kW_sim, kVAr_sim,
                   circuit_demand, capacity_i, Trafos, DERs, No_DER_RoRCurrents,
                   RoRCurrents, faulttypes, RoR_analysis, max_reduction, report_txt):
    
    if RoR_analysis is True:
        start = time.time()
        #%% Calculate the hour and date in the simulation
        h, m = snapshottime.split(':')
        if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
            if int(m) <= 7:
                m = '00'
            elif int(m) <= 22:
                m = '15'
            elif int(m) <= 37:
                m = '30'
            elif int(m) <= 52:
                m = '45'
            else:
                m = '00'
                h = str(int(h) + 1)
                if int(h) == 24:  # last round on 23:45
                    h = '23'
                    m = '45'
        
        snapshottime = h + ':' + m
        day_ = snapshotdate.replace('/', '')
        day_ = day_.replace('-', '')
        daily_strtime = str(day_ + snapshottime.replace(':', ''))         
        hora_sec = snapshottime.split(':')

        # P and Q to match
        V_to_be_matched = 0
        for ij in range(len(circuit_demand)):
            temp_a = circuit_demand[ij][0]  # day
            temp_b = circuit_demand[ij][1]  # hour                    
            if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
                V_to_be_matched = circuit_demand[ij][4]  # Voltage
                break
        
        #%% Fault type studies
        # defines the terminals necessary for each type of fault. Dictionary containing: Fault type: [terminal conection for bus 1, terminal conection for bus2, number of phases]
        faulttypeterminals = {'ABC':['.1.2','.3.3', '2'], 'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'], 'BC':['.2','.3','1'], 'AC':['.1','.3','1'], 'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'], 'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'], 'BG':['.2','.0','1'], 'CG':['.3','.0','1']}
    
        if RoRCurrents.empty:
            columnslist=list([ 'DER_kWp', 'Element', 'FaultedBus']+faulttypes+['HCSTOP'])
            RoRCurrents = pd.DataFrame(columns=columnslist)
    
        for element, faultedbus in list(No_DER_RoRCurrents.index):
            ftlist=faultedbus.split('.')
            if len(ftlist) == 2:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2], 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,2], 'AG':[0], 'BG':[0], 'CG':[0]}
                if ftlist[1] == '1':
                    faultlabels = ['AG']
                elif ftlist[1] == '2':
                    faultlabels = ['BG']
                else: 
                    faultlabels = ['CG']
            elif len(ftlist) == 3:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2], 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,2], 'AG':[0], 'BG':[0], 'CG':[0]}
                if ftlist[1] == '1' and ftlist[2] == '2':
                    faultlabels = ['AB', 'ABG','AG','BG']
                elif ftlist[1] == '2' and ftlist[2] == '3':
                    faultlabels = ['BC', 'BCG','CG','BG']
                else: 
                    faultlabels = ['AC', 'ACG','AG','CG']
            else:
                currentspositions = {'ABC':[0,2,4], 'ABCG':[0,2,4], 'AB':[0,2], 'BC':[2,4], 'AC':[0,4], 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,4], 'AG':[0], 'BG':[2], 'CG':[4]}
                faultlabels = list(faulttypeterminals.keys())
    
            faultlabels = list(set(faulttypes) & set(faultlabels)) # removes those types of faults not being evaluated
    
            temp_row = []
            temp_row.append(capacity_i)
            temp_row.append(element)
            temp_row.append(faultedbus)
    
            for faulttype in faulttypes:
                
                if faulttype not in faultlabels:
                    temp_row.append(np.nan)
                else:
                    # Machines inicialization simulation
                    DSStext.Command = 'clear'
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=daily'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                    DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
                    DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                    DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched) # Tensión en cabecera
                    for tx in Trafos:
                        DSStext.Command = tx
                    for der in DERs:
                        DSStext.Command = der
    
                    DSScircuit.Solution.Solve()
    
                    #%% Fault study (dynamics) simulation
                    # type of fault
                    term1 = faulttypeterminals[faulttype][0]
                    term2 = faulttypeterminals[faulttype][1]
                    phasessc = faulttypeterminals[faulttype][2]
                    
                    DSStext.Command = 'Solve mode=dynamic stepsize=0.00002'
                    
                    DSStext.Command = 'new fault.fault_'+ faulttype+' bus1='+ftlist[0]+term1+' bus2='+ftlist[0]+term2+' phases='+phasessc
    
    
                    DSScircuit.Solution.Solve()
    
                    DSScircuit.setActiveElement(element) # faulted element
                    temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                    minIchange=min(np.divide([temp_currents[i] for i in currentspositions[faulttype]], No_DER_RoRCurrents[faulttype][element,faultedbus]))
                    temp_row.append(minIchange)
                    HCSTOP = False if minIchange > max_reduction else True
    
            tmp = pd.DataFrame([temp_row+[HCSTOP]], index=[0],
                               columns=list(RoRCurrents.columns))
            RoRCurrents = pd.concat([RoRCurrents, tmp], ignore_index=True)
        
        end = time.time()
        sim_time = end - start
        msg = "Análisis de reducción de alcance " +str(np.round(sim_time,2))+' s'
        print(msg)
        report_txt.append(msg)
        
    return RoRCurrents, report_txt

#%% PROTECTION DEVICES VERIFICATION - Sympathetic tripping
# Iverter contribution to fault current: 1.2 (120% according to SANDIA Protections report)
# For faults upstream the circuit breaker (simulating faults in different circuits of the same HV bus in the power substation), checks the breaker's zero sequence current, has to bellow a certain value, e.g., 150 A.
# ActiveElement.seqCurrents yields for the first line of the circuit (IO_t1, I1_t1, I2_t1, IO_t2, I1_t2, I2_t2)
# Input: DSStext, DSScircuit, dss_network, firstLine, hora_sec, kW_sim, kVAr_sim, capacity_i, SympatheticTripping_analysis, faulttypes, Izero_trip, SympatheticTripping_df
# Output: SympatheticTripping_df
def SympatheticTripping(DSStext, DSScircuit, dss_network, firstLine,
                        snapshotdate, snapshottime, kW_sim, kVAr_sim, circuit_demand,
                        capacity_i, Trafos, DERs, SympatheticTripping_analysis,
                        faulttypes, I_51_p_trip, I_51_g_trip, SympatheticTripping_df, report_txt):
    faulttypeterminals = {'ABC':['.1.2','.3.3', '2'], 'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'], 'BC':['.2','.3','1'], 'AC':['.1','.3','1'], 'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'], 'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'], 'BG':['.2','.0','1'], 'CG':['.3','.0','1']}

    if SympatheticTripping_analysis:
        start = time.time()
        #%% Calculate the hour and date in the simulation
        h, m = snapshottime.split(':')
        if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
            if int(m) <= 7:
                m = '00'
            elif int(m) <= 22:
                m = '15'
            elif int(m) <= 37:
                m = '30'
            elif int(m) <= 52:
                m = '45'
            else:
                m = '00'
                h = str(int(h) + 1)
                if int(h) == 24:  # last round on 23:45
                    h = '23'
                    m = '45'
        
        snapshottime = h + ':' + m
        day_ = snapshotdate.replace('/', '')
        day_ = day_.replace('-', '')
        daily_strtime = str(day_ + snapshottime.replace(':', ''))         
        hora_sec = snapshottime.split(':')

        # P and Q to match
        V_to_be_matched = 0
        for ij in range(len(circuit_demand)):
            temp_a = circuit_demand[ij][0]  # day
            temp_b = circuit_demand[ij][1]  # hour                    
            if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
                V_to_be_matched = circuit_demand[ij][4]  # Voltage
                break
        
        element = firstLine
        ftlist = 'symptrip.1.2.3'.split('.')
        ip_max = 0
        ig_max = 0
        
        for faulttype in faulttypes:
            print('ST: '+element+ ', Fault: '+faulttype)
            # Machines inicialization simulation
            DSStext.Command = 'clear'
            DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
            DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
            DSStext.Command = 'Set mode=daily'  # Type of Simulation
            DSStext.Command = 'Set number=1'  # Number of steps to be simulated
            DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
            DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
            DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
            DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
            DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched) # Tensión en cabecera
            # Integra los trafos y DERs generados por el código
            for tx in Trafos:
                DSStext.Command = tx
            for der in DERs:
                DSStext.Command = der
            # Integra la línea extra (100m de la subestación/cabecera)
            DSStext.Command = "new line.symptrip bus1=Sourcebus.1.2.3 bus2=symptrip.1.2.3 geometry=3FMV336AAAC1/0AAAC_H length=100 units=m "
                        
            DSScircuit.Solution.Solve()

            #%% Fault study (dynamics) simulation
            # type of fault
            term1 = faulttypeterminals[faulttype][0]
            term2 = faulttypeterminals[faulttype][1]
            phasessc = faulttypeterminals[faulttype][2]

            DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
            DSStext.Command = 'new fault.fault_'+ faulttype+' bus1='+ftlist[0]+term1+' bus2='+ftlist[0]+term2+' phases='+phasessc
            print('faulttype: '+faulttype)
            print('bus1: '+ftlist[0]+term1)
            print('bus2: '+ftlist[0]+term2)
            DSScircuit.Solution.Solve()
            
            DSScircuit.SetActiveElement('line.'+firstLine)
            print('currents: '+str(DSScircuit.ActiveCktElement.CurrentsMagAng))
            
            i_pA = DSScircuit.ActiveCktElement.CurrentsMagAng[0] # Magnitud Neutro
            print("IP_A = "+str(i_pA))
            i_pB = DSScircuit.ActiveCktElement.CurrentsMagAng[2] # Magnitud Neutro
            print("IP_B = "+str(i_pB))
            i_pC = DSScircuit.ActiveCktElement.CurrentsMagAng[4] # Magnitud Neutro
            print("IP_C = "+str(i_pC))
            
            for i_fase in [i_pA, i_pB, i_pC]:
                print(i_fase, ip_max)
                if i_fase > ip_max:
                    ip_max = i_fase
                    faulttype_max_ip = faulttype
            
            if "G" in faulttype: # Solo fallas a tierra
                i_g = DSScircuit.ActiveCktElement.Residuals[0]# Magnitud residual
                print("IG = "+str(i_g))
            
                if i_g > ig_max:
                    ig_max = i_g
                    faulttype_max_ig = faulttype
                    
        HCSTOP = True if ((ip_max >= I_51_p_trip) or (ig_max >= I_51_g_trip)) else False
        
        tmp = pd.DataFrame({'DER_kWp':[capacity_i],'FAULT_IP': [faulttype_max_ip],
                            'IP':[i_fase], 'FAULT_IG': [faulttype_max_ig],
                            'IG':[i_g], 'HCSTOP' : [HCSTOP]})
        
        SympatheticTripping_df = pd.concat([SympatheticTripping_df, tmp],
                                           ignore_index=True)
        
        end = time.time()
        sim_time = end - start
        msg = "Análisis de disparo indebido " +str(np.round(sim_time,2))+' s'
        print(msg)
        report_txt.append(msg)
        
    return SympatheticTripping_df, report_txt
#%% INITIAL TX AND MV LOADS INFORMATION FOR FUTURE HISTOGRAM 
#Input: DSScircuit, DSSobj, DSStext, G, tx_layer, mv_loads_layer
#Output: LoadTrafos_MVLoads -> Dataframe 

def base_info_tx_and_mvloads(DSScircuit, DSSobj, DSStext, G, tx_layer, der_ss_layer, mv_loads_layer):
    start = time.time()
    
    first_bus = 'AFTERMETER'
    
    LoadTrafos_MVLoads_1 = tx_layer[['DSSNAME', 'LV_GROUP', 'bus1', 'KVAPHASEA', 'KVAPHASEB', 'KVAPHASEC', 'kVA_snap']]
    LoadTrafos_MVLoads_1 = LoadTrafos_MVLoads_1.assign(Rating = lambda x : x.KVAPHASEA.astype(float) + x.KVAPHASEB.astype(float) + x.KVAPHASEC.astype(float))
    LoadTrafos_MVLoads_1 = LoadTrafos_MVLoads_1.drop(columns=['KVAPHASEA', 'KVAPHASEB', 'KVAPHASEC'])  
    LoadTrafos_MVLoads_1.index = LoadTrafos_MVLoads_1['DSSNAME']
    
    if mv_loads_layer.empty  is False:
        LoadTrafos_MVLoads_2 = mv_loads_layer[['DSSNAME', 'bus1', 'kVA_snap']]
        LoadTrafos_MVLoads_2.index = LoadTrafos_MVLoads_2['DSSNAME']
        
        LoadTrafos_MVLoads = pd.concat([LoadTrafos_MVLoads_1, LoadTrafos_MVLoads_2], sort=True)
    
    else: 
        LoadTrafos_MVLoads = LoadTrafos_MVLoads_1.copy()
    
    LoadTrafos_MVLoads['distance_m'] = LoadTrafos_MVLoads['bus1'].apply(lambda b: (nx.shortest_path_length(G, first_bus, b, weight='distance'))/1000)
    if not der_ss_layer.empty:
        LoadTrafos_MVLoads["Av_Rating"] = LoadTrafos_MVLoads.apply(lambda row: row["Rating"] - der_ss_layer[der_ss_layer["LV_GROUP"] == str(row["LV_GROUP"])]["KVA"].sum(), axis=1)
    else:
        LoadTrafos_MVLoads["Av_Rating"] = LoadTrafos_MVLoads["Rating"]
    
    end = time.time()
    sim_time = end - start
    print('LoadMV: '+str(np.round(sim_time,2))+' sec.')
    
    return LoadTrafos_MVLoads

#%% GET LOAD KW VALUES AFTER BASE CASE RUN

def get_kva_load(DSScircuit, elem_name):
    
    DSScircuit.setActiveElement('load.' + elem_name)

    if len(DSScircuit.ActiveElement.Powers) > 4: # 3ph load
        temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
        p=temp_powers[0] + temp_powers[2] + temp_powers[4]
        q=temp_powers[1] + temp_powers[3] + temp_powers[5]
        s = np.sqrt(p**2+q**2)
    else:
        temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
        p=temp_powers[0] + temp_powers[2]
        q=temp_powers[1] + temp_powers[3]
        s = np.sqrt(p**2+q**2)
        
    return s #It can be modified to return P, Q or S.

def get_kva_trafos(DSScircuit, elem_name):
    
    ckt_name = elem_name.split(elem_name[min([i for i, c in enumerate(elem_name) if c.isdigit()])])[0]
    tx_num_phase =  elem_name.split(ckt_name)[1][elem_name.split(ckt_name)[1].index('P')-1] #Phase number 
                    
    #Number of tx units
    
    if str(tx_num_phase) == '3':
        try:
            tx_units = int(elem_name.split(ckt_name)[1][elem_name.split(ckt_name)[1].index('U')-1])
        except:
            tx_units=1
    else:
        tx_units=1
    
    if tx_units == 1:
        
        DSScircuit.setActiveElement('transformer.'+elem_name)
        
        if len(DSScircuit.ActiveElement.Powers) >= 16: # 3ph load
            temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
            p=temp_powers[0] + temp_powers[2] + temp_powers[4]
            q=temp_powers[1] + temp_powers[3] + temp_powers[5]
            s = np.sqrt(p**2+q**2)
        else:
            temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
            p=temp_powers[0]
            q=temp_powers[1]
            s = np.sqrt(p**2+q**2)
        
    else:
        tx_id = elem_name.split('_')[0]
        tx_num = elem_name.split('_')[1]
        
        s=0
        
        for unit in range(1, tx_units+1):
            elem_name = tx_id+'_'+str(unit)+'_'+tx_num
            
            DSScircuit.setActiveElement('transformer.'+elem_name)
            
            temp_powers = DSScircuit.ActiveElement.Powers  # extract power from tx
            p=temp_powers[0]
            q=temp_powers[1]
            s += np.sqrt(p**2+q**2)
    
    
    return s #It can be modified to return P, Q or S.

#%%INITIAL LV LOADS INFORMATION FOR FUTURE HISTOGRAM 
#Input: DSScircuit, DSSobj, DSStext, G, tx_layer, mv_loads_layer
#Output: Load_lvloads -> Dataframe blacklist

def base_info_lvloads(lv_loads_layer, LoadTrafos_MVLoads, DSScircuit):
    start = time.time()
    
    Load_lvloads = {} #dict where all the data will be placed
    blacklist = []
    for group in list(set(LoadTrafos_MVLoads['LV_GROUP'])):
        
        group_loads = lv_loads_layer.loc[lv_loads_layer['LV_GROUP']==group] #filtered dataframe
        if group_loads.empty is True:
            #append trafo
            blacklist += list(LoadTrafos_MVLoads.loc[LoadTrafos_MVLoads['LV_GROUP']==group].index.values) # appended to blacklist and will not be part of dict
        
        else:
            Load_lvloads[int(group)] = group_loads[['bus1','DSSNAME', 'NOMVOLT', 'SERVICE', 'kVA_snap']]
            
            val_sum = Load_lvloads[int(group)]['kVA_snap'].sum()
            Load_lvloads[int(group)] = Load_lvloads[int(group)].assign(kVA_norm = lambda x : x.kVA_snap/val_sum)
            
            Load_lvloads[int(group)] = Load_lvloads[int(group)].assign(DER_previo = lambda x : np.nan)
            Load_lvloads[int(group)] = Load_lvloads[int(group)].assign(DER_actual = lambda x : np.nan)
        
    end = time.time()
    sim_time = end - start
    print('LV dataframe time: '+str(np.round(sim_time,2))+' sec.')
    
    return Load_lvloads, blacklist


#%% Step calculation 

def step_calc(MV_hist_df, blacklist, blacklist_dict, max_kVA_step,
              max_kVA_mvloads, lim_kVA, der_level_actual):
    
    not_in_blacklist = list(set(MV_hist_df.index.values) - set(blacklist))
    mv_load_list = list(MV_hist_df.loc[MV_hist_df['Rating'].isnull()].index.values)
    # update kVA_snap_norm:
    val_sum = MV_hist_df.loc[not_in_blacklist, 'kVA_snap'].sum()
    
    # first estimation
    if len(list(set(not_in_blacklist) - set(mv_load_list))) > 0:
        ref_val = max_kVA_step
        study_list = list(set(not_in_blacklist) - set(mv_load_list))
        der_step = int(max_kVA_step/(MV_hist_df.loc[list(set(not_in_blacklist)-set(mv_load_list)), 'kVA_snap'].max()/val_sum))
    
    else:
        ref_val = max_kVA_mvloads
        study_list = mv_load_list
        der_step = int(max_kVA_mvloads/(MV_hist_df.loc[mv_load_list, 'kVA_snap'].max()/val_sum))
    
    ##################################################################################
    
    der_level_actual += der_step #UPDATES THE VALUE
    
    if der_level_actual > der_step:  # filter for the first iteration
        MV_hist_df['DER_previo'] = MV_hist_df.copy()[der_level_actual - der_step]
        
    # setting blacklisted values
    MV_hist_df.loc[blacklist, 'DER_actual'] = MV_hist_df.loc[blacklist, 'DER_previo']
    #then, makes 0 the kVA_snap attribute to exclude the correspondent tx or mvload from analysis
    MV_hist_df.loc[blacklist, 'kVA_snap'] = 0 #it's automatically updated
    
    # update kVA_snap_norm: (general)
    MV_hist_df['kVA_snap_norm'] = MV_hist_df['kVA_snap']/max(MV_hist_df.loc[study_list, 'kVA_snap'])
   
    # Finally, Update DER actual
    MV_hist_df.loc[not_in_blacklist, 'DER_actual'] = MV_hist_df.loc[not_in_blacklist, 'DER_previo'] + (ref_val) * MV_hist_df.loc[not_in_blacklist, 'kVA_snap_norm']
    
    #Checks if mvloads didn't excede the fixed limit
    if not (MV_hist_df.loc[mv_load_list,:].loc[(MV_hist_df['DER_actual'] - MV_hist_df['DER_previo'] > max_kVA_mvloads)]).empty:
        
        mv_topped = MV_hist_df.loc[mv_load_list,:].loc[(MV_hist_df['DER_actual'] - MV_hist_df['DER_previo'] > max_kVA_mvloads)].index.values
        #updates values
        MV_hist_df.loc[mv_topped,'DER_actual'] = MV_hist_df.loc[mv_topped,'DER_previo'] + max_kVA_mvloads
        der_level_actual = int(MV_hist_df['DER_actual'].sum())
        der_step = int(MV_hist_df['DER_actual'].sum())- int(MV_hist_df['DER_previo'].sum())
    
    if lim_kVA is True:
        #filtered data
        lim_df = MV_hist_df.loc[MV_hist_df['DER_actual'] > MV_hist_df['Av_Rating']]
        temp_blacklist = list(lim_df.index.values)
        #Updates values
        MV_hist_df.loc[temp_blacklist, 'DER_actual'] = MV_hist_df.loc[temp_blacklist, 'Av_Rating'].copy()
        der_level_actual = int(MV_hist_df['DER_actual'].sum())
        der_step = int(MV_hist_df['DER_actual'].sum())- int(MV_hist_df['DER_previo'].sum())
        
    else:
        temp_blacklist = []
        
            
    return der_step, der_level_actual, MV_hist_df, temp_blacklist
        
#%% UPDATE HISTOGRAMS
# Input: der_level_actual, LoadTrafos_MVLoads, blacklist, LV_hist_df, MV_hist_df=0
# Output: MV_hist_df, LV_hist_df

def DER_calc(der_level_actual, der_step, LoadTrafos_MVLoads, blacklist,
             LV_hist_df, MV_hist_df=0):
    
    if der_level_actual == 0: #initial step
        # first, MV histogram 
        
        MV_hist_df = LoadTrafos_MVLoads[['DSSNAME', 'bus1', 'LV_GROUP', 'Rating', 'Av_Rating' ,'kVA_snap']]
        MV_hist_df = MV_hist_df.assign(base_kVA_val = lambda x: x.kVA_snap)
        
        MV_hist_df.loc[blacklist, 'kVA_snap'] = 0 # make the txs without loads get a value of 0 (excluded from analysis)
        val_sum = MV_hist_df['kVA_snap'].sum()
        # MV_hist_df['kVA_snap_norm'] = MV_hist_df['kVA_snap']/val_sum
        MV_hist_df = MV_hist_df.assign(kVA_snap_norm = lambda x : x.kVA_snap/val_sum)
        
        MV_hist_df['DER_previo'] = 0
        MV_hist_df['DER_actual'] = 0
        # MV_hist_df['blacklist_info'] =''
        MV_hist_df.loc[blacklist, 'blacklist_info'] = 'Secundario sin cargas'
        
        # LV histogram 
        for group in LV_hist_df:
            LV_hist_df[int(group)]['DER_previo'].values[:] = 0
            LV_hist_df[int(group)]['DER_actual'].values[:] = 0
            
    else:
        # # MV histogram
        ###### Register of Installed Capacity per Iteration 
        
        MV_hist_df[der_level_actual] = np.nan
        MV_hist_df[der_level_actual] = MV_hist_df['DER_actual']
        
        # LV histogram

        for group in LV_hist_df:
            group_idx = MV_hist_df.loc[MV_hist_df['LV_GROUP']==group].index[0]
            base_val = float(MV_hist_df.loc[group_idx, 'DER_actual']) #VALOR YA ACTUALIZADO
            # print(group, base_val)
            if der_level_actual > der_step:
                LV_hist_df[int(group)]['DER_previo'] = LV_hist_df[int(group)].copy()[der_level_actual - der_step]
            
            LV_hist_df[int(group)]['DER_actual'] = LV_hist_df[int(group)]['kVA_norm'] * base_val
            
            idx_zero = LV_hist_df[int(group)].loc[LV_hist_df[int(group)]['DER_actual'] <= 0.005].index.values
            LV_hist_df[int(group)].loc[idx_zero, 'DER_actual'] = 0
            
            ###### Register of Installed Capacity per Iteration 
            
            LV_hist_df[int(group)][der_level_actual] = np.nan
            LV_hist_df[int(group)][der_level_actual] = LV_hist_df[int(group)]['DER_actual']
            
    return MV_hist_df, LV_hist_df

#%% HYBRID DER ALLOCATION ALGORITHM FOR ITERATIVE HOSTING CAPACITY (HHC - EPER) - assigns the defined capacity in MV_hist_df and LV_hist_df
# Input: LV_hist_df, MV_hist_df, mv_loads_layer
# Output: DERs_LV, DERs_MV, Trafos_DERs_MV

def DER_allocation_HHC(LV_hist_df, MV_hist_df, mv_loads_layer, trafo_df_f, icc):
    
    DERs_LV = []
    DERs_MV = []
    Trafos_DERs_MV = []
    
    for elem in MV_hist_df.loc[MV_hist_df['DER_actual'] > 0].index.values:
        group_id = MV_hist_df.loc[elem, 'LV_GROUP']
            
        if np.isnan(group_id) == False: #si es un trafo
            for idx_lv_load in LV_hist_df[int(group_id)].loc[LV_hist_df[int(group_id)]['DER_actual'] > 0].index: # recorre el df
                service = ''
                n_phases = ''
                bus_conn = ''
                nomvolt = ''
                busLV = ''
                derName = ''
                installed_capacity = ''
                der_line = []
                
                #Retrieve number of phases depending on service data from layer
                service = str(int(LV_hist_df[int(group_id)].loc[idx_lv_load, 'SERVICE']))
                if service == '12':
                    n_phases = '1'
                    bus_conn = '.1.2'
                elif service == '123':
                    n_phases = '3'
                    bus_conn = '.1.2.3'
                
                nomvolt = int(LV_hist_df[int(group_id)].loc[idx_lv_load, 'NOMVOLT'])
                kV_LowLL = str(trafOps.renameVoltage(380, nomvolt)['LVCode']['LL']) #nivel de tensión
                busLV = LV_hist_df[int(group_id)].loc[idx_lv_load, 'bus1']
                derName = LV_hist_df[int(group_id)].loc[idx_lv_load, 'DSSNAME']
                installed_capacity = str(LV_hist_df[int(group_id)].loc[idx_lv_load, 'DER_actual'])
    
                der_line = 'new generator.DER_' + derName + ' phases='
                der_line += n_phases + ' bus1=' + busLV + bus_conn
                der_line += ' kV=' + kV_LowLL + ' kW=' + str(float(installed_capacity))
                der_line += ' PF=1 conn=wye kVA='
                der_line+= str(np.round(float(installed_capacity)*1.2,2))
                der_line += ' Model=7 Vmaxpu=1.5 Vminpu='+str(1/icc)+' Balanced=yes Enabled=yes'                    
        ###############################################
                DERs_LV.append(der_line)
        
        elif mv_loads_layer.empty is False: # si es una carga de media
            # inicialización de variables
            service = ''
            n_phases = ''
            bus_conn = ''
            nomvolt = ''
            busLV = ''
            busMV = ''
            derName = '' 
            installed_capacity = ''
            kVA = ''
            
            trafo_line = []
            der_line = []
            
            #index en el layer
            idx_mv_load = mv_loads_layer.loc[mv_loads_layer['DSSNAME']==elem].index[0]
            
            #hacer la diferenciación de casos
            service = str(mv_loads_layer.loc[idx_mv_load, 'PHASEDESIG'])
                
            if service in ['7', 'ABC', 'RST']:
                n_phases = '3'
                bus_conn = '.1.2.3'
            elif service in ['6', 'AB', 'RS']:
                n_phases = '2'
                bus_conn = '.1.2'
            elif service in ['5', 'AC', 'RT']:
                n_phases = '2'
                bus_conn = '.1.3'
            elif service in ['4', 'A', 'R']:
                n_phases = '1'
                bus_conn = '.1'
            elif service in ['3', 'BC', 'ST']:
                n_phases = '2'
                bus_conn = '.1.2'
            elif service in ['2', 'B', 'S']:
                n_phases = '1'
                bus_conn = '.2'
            elif service in ['1', 'C', 'T']:
                n_phases = '1'
                bus_conn = '.3'
                    
            busMV = mv_loads_layer.loc[idx_mv_load, 'bus1']
            busLV = 'BUSLV' + busMV.split('MV')[1][0:3] + '_DER_' + elem
            
            derName = elem
            installed_capacity = str(MV_hist_df.loc[elem, 'DER_actual'])
            
            ### tx parameters calculation
            
            #kVA trafo
            if n_phases == '3':
                for kva_rating in [float(i) for i in list(trafOps.imag_list3F.keys())]:
                    if kva_rating > 1.20*float(installed_capacity):
                        kVA = str(kva_rating)
                        break
            
            else:
                for kva_rating in [float(i) for i in list(trafOps.imag_list1F.keys())]:
                    if kva_rating > 1.20*float(installed_capacity):
                        kVA = str(kva_rating)
                        break
            
            normhkva = " normhkva=" + str(kVA)
            
            # voltage level
            nomvolt = int(mv_loads_layer.loc[idx_mv_load, 'NOMVOLT'])
            
            if n_phases == '1':
                kV_MedLL = str(trafOps.renameVoltage(nomvolt, 40)['MVCode']['LL']) # media tensión trifásico
                kV_MedLN = str(trafOps.renameVoltage(nomvolt, 40)['MVCode']['LN']) # media tensión para monofásico
                kV_LowLL = str(trafOps.renameVoltage(nomvolt, 40)['LVCode']['LL']) # baja tensión trifásico (480 V)
                kV_LowLN = str(trafOps.renameVoltage(nomvolt, 40)['LVCode']['LN']) # baja tensión para monofásico
             
                impedance_X = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('X')
                impedance_R = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('R')
                noloadloss = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Pnoload')
                imag = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Im')
                trafName = elem.split('MV_')[1] +'_DER'
                
                confMV = 'wye'; confLV_1 = 'wye'; confLV_2 = 'wye'
                tap = '1'
                
                # OpenDSS tx sentence
                trafo_line = 'new transformer.' + trafName + ' phases='
                trafo_line += n_phases+ ' windings=3 ' + noloadloss + " "
                trafo_line += imag + ' buses=[' + busMV +bus_conn+' '
                trafo_line += busLV + '.1.0 '+ busLV + '.0.2 ] conns=['
                trafo_line += confMV + ' ' + confLV_1 + ' '+confLV_2+'] kvs=['
                trafo_line += kV_MedLN + ' ' +  kV_LowLN + ' '+ kV_LowLN + ']'
                trafo_line += ' kvas=[' + kVA + ' ' + kVA + ' '+ kVA +'] '
                trafo_line += impedance_X +' '+ impedance_R + ' Taps=[' + tap
                trafo_line += ', 1, 1 ]' + normhkva
                
                # OpenDSS DER sentence
                der_line = 'new generator.DER_'+ derName + ' phases='
                der_line += n_phases + ' bus1=' + busLV + '.1.2' + ' kV='
                der_line += kV_LowLL + ' kW=' + str(float(installed_capacity))
                der_line += ' PF=1 conn=wye kVA='
                der_line += str(np.round(float(installed_capacity)*1.2,2))
                der_line += ' Model=7 Vmaxpu=1.5 Vminpu='+str(1/icc)+' Balanced=yes Enabled=yes'  


            elif n_phases == '2' or n_phases == '3':
                kV_MedLL = str(trafOps.renameVoltage(nomvolt, 50)['MVCode']['LL']) # media tensión trifásico
                kV_MedLN = str(trafOps.renameVoltage(nomvolt, 50)['MVCode']['LN']) # media tensión para monofásico
                kV_LowLL = str(trafOps.renameVoltage(nomvolt, 50)['LVCode']['LL']) # baja tensión trifásico (480 V)
                kV_LowLN = str(trafOps.renameVoltage(nomvolt, 50)['LVCode']['LN']) # baja tensión para monofásico
                impedance = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Z')
                noloadloss = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Pnoload')
                imag = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Im')
                trafName = elem.split('MV')[1][0:3] + n_phases + 'DERs_' + elem
                
                confMV = 'wye'
                confLV = 'wye'
                tap = '1'
                
                # OpenDSS tx sentence
                trafo_line = 'new transformer.' + trafName + ' phases=3 windings=2 '
                trafo_line += noloadloss + " " + imag + ' buses=[' + busMV + '.1.2.3 '
                trafo_line += busLV + '.1.2.3]' + ' conns=[' + confMV + ' '
                trafo_line += confLV + ']' + ' kvs=[' + kV_MedLL + " " +  kV_LowLL + ']'
                trafo_line += ' kvas=[' + kVA + " " + kVA + '] ' + impedance
                trafo_line += ' Taps=[' + tap + ', 1]' + normhkva
                
                # OpenDSS DER sentence
                der_line = 'new generator.DER_'+ derName + ' phases='
                der_line += n_phases + ' bus1=' + busLV + bus_conn + ' kV='
                der_line += kV_LowLL + ' kW=' + str(float(installed_capacity))
                der_line += ' PF=1 conn=wye kVA='
                der_line += str(np.round(float(installed_capacity)*1.2,2))
                der_line += ' Model=7' + ' Vmaxpu=1.5 Vminpu='+str(1/icc)+' Balanced=yes Enabled=yes'

        ###############################################
            DERs_MV.append(der_line)
            Trafos_DERs_MV.append(trafo_line)
      
    print("************** For DER_allocation ************** ")
    print("len(DERs_LV) = ", len(DERs_LV), " len(DERs_MV) = ", len(DERs_MV), "len(Trafos_DERs_MV) = ", len(Trafos_DERs_MV))
    return DERs_LV, DERs_MV, Trafos_DERs_MV
#%%
def flag_and_blacklist_calc(Overvoltage_loads_df, Voltagedeviation_loads_df, 
                            Overvoltage_rest_df, Voltagedeviation_rest_df, 
                            Voltagedeviation_reg_df, Voltageunbalance_df, 
                            Thermal_loading_lines_df, Thermal_loading_tx_df,
                            FFCurrents, BFCCurrents, RoRCurrents, SympatheticTripping_df,
                            Overvoltage_analysis, VoltageDeviation_analysis, 
                            VoltageRegulation_analysis, VoltageUnbalance, Thermal_analysis, 
                            FF_analysis, BFC_analysis, RoR_analysis, SympatheticTripping_analysis, MV_hist_df, 
                            capacity_i, flag, blacklist, blacklist_dict):
    
    criteria_dict = {'Sobretensión en cargas':{'Criteria': Overvoltage_analysis, 'data': Overvoltage_loads_df},
                     'Sobretensión en líneas':{'Criteria': Overvoltage_analysis, 'data': Overvoltage_rest_df},
                     'Desviación de tensión en cargas':{'Criteria': VoltageDeviation_analysis, 'data': Voltagedeviation_loads_df},
                     'Desviación de tensión en líneas':{'Criteria': VoltageDeviation_analysis, 'data': Voltagedeviation_rest_df},
                     'Regulación de tensión':{'Criteria': VoltageRegulation_analysis, 'data': Voltagedeviation_reg_df},
                     'Desbalance de tensión':{'Criteria': VoltageUnbalance, 'data': Voltageunbalance_df},
                     'Cargabilidad en transformadores': {'Criteria': Thermal_analysis, 'data': Thermal_loading_tx_df},
                     'Cargabilidad en líneas': {'Criteria': Thermal_analysis, 'data': Thermal_loading_lines_df}, 
                     'Aumento de corriente de falla': {'Criteria': FF_analysis, 'data': FFCurrents},
                     'Coordinación Recloser/Fusible':{'Criteria': BFC_analysis, 'data': BFCCurrents},
                     'Reducción de alcance':{'Criteria': RoR_analysis, 'data': RoRCurrents},
                     'Disparo indebido':{'Criteria': SympatheticTripping_analysis, 'data': SympatheticTripping_df}}
    
    hc_stop_list = []
    hc_blacklist_list = []
    flag_list = []
    for crit in criteria_dict:
        if criteria_dict[crit]['Criteria'] == True:
            #HCSTOP INFO
            hc_stop_list.append(criteria_dict[crit]['data'])
            try:
                if criteria_dict[crit]['data']['HCSTOP'].any():
                    flag_list.append(crit)
            except:
                pass
            #BLACKLIST INFO
            try:
                hc_blacklist_list.append(criteria_dict[crit]['data'].iloc[-1,:].to_frame().T)
                bl_elem = criteria_dict[crit]['data'].iloc[-1,:].to_frame().T['BLACKLIST_TX'].tolist()[0]
                MV_hist_df.loc[bl_elem, 'blacklist_info'] = crit
            except:
                pass
            
    
    hc_stop = pd.DataFrame(pd.concat(hc_stop_list, sort=True)['HCSTOP'])
    
    try:
        hc_blacklist = pd.DataFrame(pd.concat(hc_blacklist_list, sort=True)['BLACKLIST_TX']).dropna()
    except:
        hc_blacklist = pd.DataFrame([],columns=["BLACKLIST_TX"], index=[])
    
    updated_blacklist = [] #sum of tx who entered in the blacklist at the evaluated DER level
 
    for index, rows in hc_blacklist.iterrows(): # Create list for the current row 
        updated_blacklist += rows.BLACKLIST_TX
        
    updated_blacklist =list(set(updated_blacklist)) # the new ones from the voltage and thermal criteria at this DER level
    
    new_blacklist = list(set(updated_blacklist) - set(blacklist)) #the new topped tx at this DER level
    #save blacklist per installed capacity
    try: #if blacklist_dict[capacity] it's already called
        blacklist_dict[capacity_i] = list(set(blacklist_dict[capacity_i] + new_blacklist))
    except:
        blacklist_dict[capacity_i] =  new_blacklist #the new topped tx at this DER level
    
    blacklist += list(set(updated_blacklist + blacklist)); blacklist = list(set(blacklist)) # the accumulative blacklist 
    
    # Flag info
    if (hc_stop['HCSTOP'].any()):
        flag = True
        
    elif (len(blacklist) == len(MV_hist_df.index)):
        flag = True
        flag_list = ['Problemas térmicos o de tensión'] 
    
    redo_sim = False # Varible whose function is to determine if it's necessary to repeat the simulation at the same DER level, but with updated blacklist.
    if len(new_blacklist) > 0:
        # print(new_blacklist)
        print('blacklisted txs: ' + str(len(blacklist)) + '/' +str(len(MV_hist_df.index)))
        redo_sim = True
    
    return flag, flag_list, blacklist, blacklist_dict, redo_sim
        
#%% Save and upload variables in pickle files
# Input: path,name,variable
# Output: NONE, just saves the variables 

def pickle_dump(path,name,variable):
    with open(os.path.join(path, name+'.pkl'), 'wb') as f:
        pickle.dump(variable, f)
    f.close()

#%% Save and upload variables in pickle files
# Input: path,name
# Output: variable loaded

def pickle_load(path,name):
    with open(os.path.join(path, name+'.pkl'),'rb') as f:
        variable = pickle.load(f)
    return variable

#%% 
def save_criteria_data_iterative(data_dict, Overvoltage_loads_df, Voltagedeviation_loads_df, 
                                 Overvoltage_rest_df, Voltagedeviation_rest_df, Voltagedeviation_reg_df, 
                                 Voltageunbalance_df, Thermal_loading_lines_df, Thermal_loading_tx_df, 
                                 FFCurrents, BFCCurrents, RoRCurrents, SympatheticTripping_df):
    
    data_dict['Overvoltage_loads'] = Overvoltage_loads_df
    data_dict['Overvoltage_rest'] = Overvoltage_rest_df
    data_dict['Voltage_deviation_loads'] = Voltagedeviation_loads_df
    data_dict['Voltage_deviation_rest'] = Voltagedeviation_rest_df
    data_dict['Voltage_deviation_reg'] = Voltagedeviation_reg_df
    data_dict['Voltage_unbalance'] = Voltageunbalance_df
    data_dict['Thermal_lines'] = Thermal_loading_lines_df
    data_dict['Thermal_txs'] = Thermal_loading_tx_df
    data_dict['RoR'] = RoRCurrents
    data_dict['FF'] = FFCurrents
    data_dict['BFC'] = BFCCurrents
    data_dict['SympatheticTripping'] = SympatheticTripping_df

    return data_dict
    
def GraphSnapshotVoltages(dataframe_voltages, fig_name):
    try:
        
        nodes = list(dataframe_voltages.index)
        
        mv_nodes_A = []
        mv_nodes_B = []
        mv_nodes_C = []
        
        lv_nodes_A = []
        lv_nodes_B = []
        lv_nodes_C = []
        
        for node in nodes:
            if "mv" in str(node).lower() or "source" in str(node).lower():
                if str(node)[-1] == '1':
                    mv_nodes_A.append(node)
                elif str(node)[-1] == '2':
                    mv_nodes_B.append(node)
                elif str(node)[-1] == '3':
                    mv_nodes_C.append(node)
            
            if "lv" in str(node).lower():
                if str(node)[-1] == '1':
                    lv_nodes_A.append(node)
                elif str(node)[-1] == '2':
                    lv_nodes_B.append(node)
                elif str(node)[-1] == '3':
                    lv_nodes_C.append(node)
            
        V_mv_A = pd.DataFrame(index = mv_nodes_A, columns = ['voltage', 'distance'])
        V_mv_B = pd.DataFrame(index = mv_nodes_B, columns = ['voltage', 'distance'])
        V_mv_C = pd.DataFrame(index = mv_nodes_C, columns = ['voltage', 'distance'])
        
        V_lv_A = pd.DataFrame(index = lv_nodes_A, columns = ['voltage', 'distance'])
        V_lv_B = pd.DataFrame(index = lv_nodes_B, columns = ['voltage', 'distance'])
        V_lv_C = pd.DataFrame(index = lv_nodes_C, columns = ['voltage', 'distance'])
        
        #Asignar valores
        for node in list(dataframe_voltages.index):
            #FASE A MEDIA TENSIÓN
            if str(node)[-1] == '1' and node in mv_nodes_A:
                V_mv_A.loc[node, 'voltage'] = dataframe_voltages.loc[node, 'VSnap_pu']
                V_mv_A.loc[node, 'distance'] = dataframe_voltages.loc[node, 'Distancia']
            #FASE A BAJA TENSIÓN
            elif str(node)[-1] == '1' and node in lv_nodes_A:
                V_lv_A.loc[node, 'voltage'] = dataframe_voltages.loc[node, 'VSnap_pu']
                V_lv_A.loc[node, 'distance'] = dataframe_voltages.loc[node, 'Distancia']
            #FASE B MEDIA TENSIÓN
            elif str(node)[-1] == '2' and node in mv_nodes_B:
                V_mv_B.loc[node, 'voltage'] = dataframe_voltages.loc[node, 'VSnap_pu']
                V_mv_B.loc[node, 'distance'] = dataframe_voltages.loc[node, 'Distancia']
            #FASE B BAJA TENSIÓN
            elif str(node)[-1] == '2' and node in lv_nodes_B:
                V_lv_B.loc[node, 'voltage'] = dataframe_voltages.loc[node, 'VSnap_pu']
                V_lv_B.loc[node, 'distance'] = dataframe_voltages.loc[node, 'Distancia']
            #FASE C MEDIA TENSIÓN
            elif str(node)[-1] == '3' and node in mv_nodes_C:
                V_mv_C.loc[node, 'voltage'] = dataframe_voltages.loc[node, 'VSnap_pu']
                V_mv_C.loc[node, 'distance'] = dataframe_voltages.loc[node, 'Distancia']
            #FASE C BAJA TENSIÓN
            elif str(node)[-1] == '3' and node in lv_nodes_C:
                V_lv_C.loc[node, 'voltage'] = dataframe_voltages.loc[node, 'VSnap_pu']
                V_lv_C.loc[node, 'distance'] = dataframe_voltages.loc[node, 'Distancia']
        
        #PLOT 
        
        leyenda = []
        fign = plt.figure(fig_name)
        plt.title("Tensiones pu con respecto a la distancia de la subestación")
        if V_mv_A.empty == False:	
            plt.plot(V_mv_A['distance'], V_mv_A['voltage'], 'ro', markersize = 3.5)
            leyenda.append("MV fase A")
        if V_mv_B.empty == False:
            plt.plot(V_mv_B['distance'], V_mv_B['voltage'], 'ko', markersize = 3.5)
            leyenda.append("MV fase B")
        if V_mv_C.empty == False:
            plt.plot(V_mv_C['distance'], V_mv_C['voltage'], 'bo', markersize = 3.5)
            leyenda.append("MV fase C")
        if V_lv_A.empty == False:
            plt.plot(V_lv_A['distance'], V_lv_A['voltage'], 'r.', markersize = 2.5)
            leyenda.append("LV fase A (vivo 1)")
        if V_lv_B.empty == False:
            plt.plot(V_lv_B['distance'], V_lv_B['voltage'], 'k.', markersize = 2.5)
            leyenda.append("LV fase B (vivo 2)")
        if V_lv_C.empty == False:
            plt.plot( V_lv_C['distance'], V_lv_C['voltage'], 'b.', markersize = 2.5)
            leyenda.append("LV fase C")
        

        titulo_graph = "Tensión pu por bus"
       
        plt.legend(leyenda)
        plt.ylabel("Tensión (pu)")
        plt.xlabel("Distancia (km)")
        
        plt.title(titulo_graph)
        mng = plt.get_current_fig_manager()
        mng.window.showMaximized()
        
        
        print("Graficación de tensiones exitosa")
        return 1
    except:
        exc_info = sys.exc_info()
        print("\nError: ", exc_info)
        print("*************************  Información detallada del error ********************")
            
        for tb in traceback.format_tb(sys.exc_info()[2]):
            print(tb)
        return 0
            
## Función de estudio base 1: verificación de tensiones para evitar malas conexiones

def base_study_1(DSStext, DSScircuit, dss_network, firstLine):
    
    DSStext.Command = 'clear'  # clean previous circuits
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=snapshot'  # Type of Simulation
    DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time                
    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
    DSStext.Command = 'batchedit load..* enabled = no' # No load simulation
    DSStext.Command = 'batchedit storage..* enabled = no' # No Storage
    DSStext.Command = 'batchedit PVSystem..* enabled = no' # No PVsystems
    DSStext.Command = 'batchedit Generator..* enabled = no' # No generators
    DSStext.Command = 'batchedit Capacitor..* enabled = no' # No capacitors
    DSStext.Command = 'batchedit RegControl..* enabled = no' # No RegControls
    DSStext.Command = 'batchedit CapControl..* enabled = no' # No CapControls
    
    for meter in DSScircuit.Meters.AllNames:
        DSScircuit.Meters.Name = meter
        if not "sub" in meter:
            DSStext.Command = 'EnergyMeter.'+meter+'.enabled = no'
    
    DSScircuit.Solution.Solve()  # Solve the circuit

    VBuses_b = pd.DataFrame(list(DSScircuit.AllBusVmag), index=list(DSScircuit.AllNodeNames), columns=['VOLTAGEV'])
    VBuses_b=VBuses_b[~VBuses_b.index.str.contains('aftermeter')]; VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_der')]; VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_swt')]

    base_vals = [120, 138, 208, 240, 254, 277, 416, 440, 480,
                        2402, 4160, 7620, 7967, 13200, 13800, 14380,
                        19920, 24900, 34500, 79670, 132790]

    Base_V = pd.DataFrame()
    Base_V['BASE'] = VBuses_b['VOLTAGEV'].apply(lambda v : base_vals[[abs(v-i) for i in base_vals].index(min([abs(v-i) for i in base_vals]))])
    
    Profile_Voltage = pd.DataFrame(np.nan, index=list(DSScircuit.AllNodeNames), columns=['Distancia', 'Base', 'VSnap_Mag', 'VSnap_pu'])
    Profile_Voltage['Distancia'] = list(DSScircuit.AllNodeDistances)
    Profile_Voltage['Base'] = Base_V['BASE']
    Profile_Voltage['VSnap_Mag'] = list(DSScircuit.AllBusVmag)
    Profile_Voltage['VSnap_pu'] = Profile_Voltage['VSnap_Mag']/Profile_Voltage['Base']
    
    GraphSnapshotVoltages(Profile_Voltage, "Estudio base 1: Verificación de conectividad de red")

    return 0

## Funciones de estudio base 2: verificación de snapshot sin DER existentes, pero modificando
## la curva del alimentador

def sum_curves(df_ss, df_ls, curve_dir):
    curve_dir_ss = os.path.join(curve_dir, "DG")
    curve_dir_ls = os.path.join(curve_dir, "LSDG")

    # Create an empty list to store the sum of all curves
    # Small scale DG:
    sum_list_p_ss = [0] * 96
    sum_list_q_ss = [0] * 96
    
    # Large scale DG:
    sum_list_p_ls = [0] * 96
    sum_list_q_ls = [0] * 96    

    # Part 1: Small Scale:
    
    if not df_ss.empty: 
        # Iterate over each row in the DataFrame
        for _, row in df_ss.iterrows():
            # Get the filename specified in the "CURVE1" column
            curve_file_p = row["CURVE1"]
            curve_file_q = row["CURVE2"]
            
            if curve_file_p == curve_file_q:
                Flag_same_csv = True
            else:
                Flag_same_csv = False
            
            # Create the full path to the curve file
            # P:
            curve_path_p = os.path.join(curve_dir_ss, curve_file_p)
            ext_p = os.path.splitext(curve_path_p)[1].lower()
            # Q:
            curve_path_q = os.path.join(curve_dir_ss, curve_file_q)
            ext_q = os.path.splitext(curve_path_p)[1].lower()
            
            # Load the data from the curve file
            if ((ext_p == ".csv") and (ext_q == ".csv") and (Flag_same_csv is True)):
                curve_data_p = pd.read_csv(curve_path_p, usecols=[0], header=None).values.flatten()
                curve_data_q = pd.read_csv(curve_path_q, usecols=[1], header=None).values.flatten()
            
            elif ((ext_p == ".csv") and (ext_q == ".csv") and (Flag_same_csv is False)):
                curve_data_p = pd.read_csv(curve_path_p, usecols=[0], header=None).values.flatten()
                curve_data_q = pd.read_csv(curve_path_q, usecols=[0], header=None).values.flatten()
            
            elif ((ext_p == ".csv") and (ext_q != ".csv") and (Flag_same_csv is False)):
                curve_data_p = pd.read_csv(curve_path_p, usecols=[0], header=None).values.flatten()
                curve_data_q = pd.read_csv(curve_path_q, header=None).values.flatten().tolist()
            
            elif ((ext_p != ".csv") and (ext_q == ".csv") and (Flag_same_csv is False)):
                curve_data_p = pd.read_csv(curve_path_p, header=None).values.flatten().tolist()
                curve_data_q = pd.read_csv(curve_path_q, usecols=[0], header=None).values.flatten()
            
            elif ((ext_p != ".csv") and (ext_q != ".csv") and (Flag_same_csv is False)):
                curve_data_p = pd.read_csv(curve_path_p, header=None).values.flatten().tolist()
                curve_data_q = pd.read_csv(curve_path_q, header=None).values.flatten().tolist()
            
            # Add the data to the sum_list
            sum_list_p_ss = [sum(x) for x in zip(sum_list_p_ss, curve_data_p)]
            sum_list_q_ss = [sum(x) for x in zip(sum_list_q_ss, curve_data_q)]
    
    # Part 2: Large Scale:
    
    if not df_ls.empty: 
        # Iterate over each row in the DataFrame
        for _, row in df_ls.iterrows():
            # Get the filename specified in the "DAILY" column
            curve_file = row["DAILY"]
            curve_path = os.path.join(curve_dir_ls, curve_file)
            
            # Load the data from the curve file
            curve_data_p_ls = pd.read_csv(curve_path, usecols=[0], header=None).values.flatten()
            curve_data_q_ls = pd.read_csv(curve_path, usecols=[1], header=None).values.flatten()
            
            sum_list_p_ls = [sum(x) for x in zip(sum_list_p_ls, curve_data_p_ls)]
            sum_list_q_ls = [sum(x) for x in zip(sum_list_q_ls, curve_data_q_ls)]
    
    # Part 3: Sum all:
    
    sum_list_p = [sum(x) for x in zip(sum_list_p_ss, sum_list_p_ls)]
    sum_list_q = [sum(x) for x in zip(sum_list_q_ss, sum_list_q_ls)]
    
    return sum_list_p, sum_list_q

def LoadAllocation_Run(DSSprogress, DSStext, DSScircuit, DSSobj, s_date, s_time, study, firstLine, circuit_demand, dss_network, tx_modelling, substation_type, line_tx_definition):
    
    t1= time.time()
      # time: hh:mm
    #%% Calculate the hour in the simulation
    h, m = s_time.split(':')
    if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
        if int(m) <= 7:
            m = '00'
        elif int(m) <= 22:
            m = '15'
        elif int(m) <= 37:
            m = '30'
        elif int(m) <= 52:
            m = '45'
        else:
            m = '00'
            h = str(int(h) + 1)
            if int(h) == 24:  # last round on 23:45
                h = '23'
                m = '45'

    s_time = h + ':' + m
    day_ = s_date.replace('/', '')
    day_ = day_.replace('-', '')
    daily_strtime = str(day_ + s_time.replace(':', ''))         
    hora_sec = s_time.split(':')
    
    if study.lower() == 'snapshot': 
    
        # P and Q to match
        P_to_be_matched = 0
        Q_to_be_matched = 0
        V_to_be_matched = 0
        for ij in range(len(circuit_demand)):
            temp_a = circuit_demand[ij][0]  # day
            temp_b = circuit_demand[ij][1]  # hour                    
            if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
                P_to_be_matched = circuit_demand[ij][2]  # Active power
                Q_to_be_matched = circuit_demand[ij][3]  # Reactive power
                V_to_be_matched = circuit_demand[ij][4]  # Reactive power
                break
                
        # LoadAllocation Simulation
        
        DSStext.Command = 'clear'
        DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
        DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
        DSStext.Command = 'Set mode=daily'  # Type of Simulation
        DSStext.Command = 'Set number=1'  # Number of steps to be simulated
        DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
        DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time                
        DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
        
        # Modify the vpu from source according to circuit_demand voltage curve:
        DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched)
    
        # Run the daily power flow for a particular moment
        DSScircuit.Solution.Solve()  # Initialization solution                                            
        
        errorP = 0.003  # Maximum desired correction error for active power
        errorQ = 0.01  # Maximum desired correction error for reactive power
        max_it_correction = 100  # Maximum number of allowed iterations
        # study = 'snapshot'  # Study type for PQ_corrector
        gen_powers = np.zeros(1)
        gen_rpowers = np.zeros(1)
        # De acá en adelante hasta proximo # se comenta si no se quiere tomar encuenta la simulación de GD ya instalada
        gen_p = 0
        gen_q = 0
        GenNames = DSScircuit.Generators.AllNames
        
        if GenNames[0] != 'NONE':
            for i in GenNames: # extract power from generators
                DSScircuit.setActiveElement('generator.' + i)
                if DSScircuit.ActiveElement.Enabled is True: # Solo en generadores activos
                    p = DSScircuit.ActiveElement.Powers
                    for w in range(0, len(p), 2):
                        gen_p += -p[w] # P
                        gen_q += -p[w + 1] # Q
                        
            gen_powers[0] += gen_p
            gen_rpowers[0] += gen_q
        
        print("Potencia activa generada por generadores: " +str(gen_powers[0]) +" kW")
        print("Potencia reactiva generada por generadores: " +str(gen_rpowers[0]) +" kVAr")
        DSSobj.AllowForms = 0
        
        [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim, kVAr_sim] = auxfcns.PQ_corrector(DSSprogress, DSScircuit, DSStext, errorP, errorQ, max_it_correction,
                                      P_to_be_matched, Q_to_be_matched, V_to_be_matched, hora_sec, study,
                                      dss_network, tx_modelling, 1, firstLine, substation_type,
                                      line_tx_definition, gen_powers, gen_rpowers)
                                      
    
        DSSobj.AllowForms = 1
        t2= time.time()
        print('Tiempo del load allocation: '+str(t2-t1))
    
    #%%
    elif study== 'daily':   
        # P and Q to match
        P_to_be_matched = []
        Q_to_be_matched = []
        for ij in range(len(circuit_demand)):
            temp_a = circuit_demand[ij][0]  # day             
            if str(temp_a.replace('/', '')) == str(s_date.replace('/','')):                        
                P_to_be_matched.append(circuit_demand[ij][2])  # Active power
                Q_to_be_matched.append(circuit_demand[ij][3])  # Reactive power
                V_to_be_matched.append(circuit_demand[ij][4])  # Voltage
        
        DSStext.Command = 'clear'  # clean previous circuits
        DSStext.Command = 'New Circuit.Circuito_Distribucion_Daily'  # create a new circuit
        DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # master file compilation
        DSStext.Command = 'Set mode = daily'  # daily simulation mode
        DSStext.Command = 'Set number= 1'  # steps by solve
        DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s
        DSStext.Command = 'Set time=(0,0)'  # Set the start simulation time
        DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the first line to monitor P and Q
        # generators powers
        gen_powers = np.zeros(96)
        gen_rpowers = np.zeros(96)
        GenNames = DSScircuit.Generators.AllNames
        
        # solver for-loop
        for t in range(96):
            DSScircuit.Solution.Solve()
        
            if GenNames[0] != 'NONE':
                gen_p = 0
                gen_q = 0
                for i in GenNames:  # extract power from existing generators
                    DSScircuit.setActiveElement('generator.' + i)
                    if DSScircuit.ActiveElement.Enabled is True: # Solo en generadores activos
                        p = DSScircuit.ActiveElement.Powers
                        for w in range(0, len(p), 2):
                            gen_p += -p[w]
                            gen_q += -p[w+1]
                gen_powers[t] += gen_p
                gen_rpowers[t] += gen_q

        errorP = 0.003  # Maximum desired correction error for active power
        errorQ = 0.01  # Maximum desired correction error for reactive power
        max_it_correction = 100  # Maximum number of allowed iterations

        # load allocation algorithm            
        
        [DSScircuit, errorP_i, errorQ_i, temp_powersP, temp_powersQ, kW_sim, kVAr_sim] = auxfcns.PQ_corrector(DSSprogress, DSScircuit, DSStext, errorP, errorQ, max_it_correction,
                                      P_to_be_matched, Q_to_be_matched, hora_sec, study,
                                      dss_network, tx_modelling, 1, firstLine, substation_type,
                                      line_tx_definition, gen_powers, gen_rpowers)
        
        
        auxfcns.PQ_corrector(DSScircuit, DSStext, errorP, errorQ, max_it_correction, P_to_be_matched, Q_to_be_matched, V_to_be_matched, hora_sec, study, dss_network, 1, firstLine, gen_powers, gen_rpowers)
        
        t2= time.time()
        print('Tiempo del Load Allocation: '+str(t2-t1))
    
    return kW_sim, kVAr_sim
    
def base_study_2(DSSprogress, DSStext, DSScircuit, DSSobj, snapshotdate, snapshottime, dir_network, tx_modelling, firstLine, substation_type, line_tx_definition, der_ss, der_ls, circuit_demand):

    # Paso 1: Se procede con el load allocation
    study = "snapshot"
    
    #%% Calculate the hour in the simulation
    h, m = snapshottime.split(':')
    if m != '00' or m != '15' or m != '30' or m != '45':  # round sim minutes
        if int(m) <= 7:
            m = '00'
        elif int(m) <= 22:
            m = '15'
        elif int(m) <= 37:
            m = '30'
        elif int(m) <= 52:
            m = '45'
        else:
            m = '00'
            h = str(int(h) + 1)
            if int(h) == 24:  # last round on 23:45
                h = '23'
                m = '45'

    snapshottime = h + ':' + m
    day_ = snapshotdate.replace('/', '')
    day_ = day_.replace('-', '')
    daily_strtime = str(day_ + snapshottime.replace(':', ''))         
    hora_sec = snapshottime.split(':')
    
    V_to_be_matched = 0
    for ij in range(len(circuit_demand)):
        temp_a = circuit_demand[ij][0]  # day
        temp_b = circuit_demand[ij][1]  # hour                    
        if str(temp_a.replace('/', '') + temp_b.replace(':', '')) == daily_strtime:                        
            V_to_be_matched = circuit_demand[ij][4]  # Voltage
            break
    
    kW_sim, kVAr_sim = LoadAllocation_Run(DSSprogress, DSStext, DSScircuit, DSSobj, snapshotdate, snapshottime, study, firstLine, circuit_demand, dir_network, tx_modelling, substation_type, line_tx_definition)
    
    # Paso 2: Verificación de tensiones base:
    Base_V = getbases_simple(DSStext, DSScircuit, dir_network, firstLine)
    
    # Paso 3: Simulación
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dir_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
    
    for meter in DSScircuit.Meters.AllNames:
        DSScircuit.Meters.Name = meter
        if not "sub" in meter:
            DSStext.Command = 'EnergyMeter.'+meter+'.enabled = no'
    
    DSStext.Command = 'batchedit load..* kW=' + str(kW_sim[0]) # kW corrector
    DSStext.Command = 'batchedit load..* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
    
    # Modify the vpu from source according to circuit_demand voltage curve:
    DSStext.Command = "VSource.Source.pu =" +str(V_to_be_matched)
        
    DSScircuit.Solution.Solve()
   
    Profile_Voltage = pd.DataFrame(np.nan, index=list(DSScircuit.AllNodeNames), columns=['Distancia', 'Base', 'VSnap_Mag', 'VSnap_pu'])
    Profile_Voltage['Distancia'] = list(DSScircuit.AllNodeDistances)
    Profile_Voltage['Base'] = Base_V['BASE']
    Profile_Voltage['VSnap_Mag'] = list(DSScircuit.AllBusVmag)
    Profile_Voltage['VSnap_pu'] = Profile_Voltage['VSnap_Mag']/Profile_Voltage['Base']
    
    #Profile_Voltage.to_csv("hola.csv")
    
    GraphSnapshotVoltages(Profile_Voltage, "Estudio base 2: Resultado estudio Snapshot caso base con DER")
    
    return 0

def update_final_dataframe(df, add_value, max_value):
    if df.empty:
        return df
    
    df_updated = df.copy()  # Create a copy of the original dataframe
    
    # Add the input value to the "DER_kWp" column
    df_updated["DER_kWp"] += add_value
    
    # Filter out rows where "DER_kWp" value is greater than the input max_value
    df_updated = df_updated[df_updated["DER_kWp"] <= max_value]
    
    return df_updated
    
        
    
    
    
    