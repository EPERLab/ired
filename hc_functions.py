# -*- coding: utf-8 -*-
"""
Created on Wed Oct  7 13:59:32 2020

@author: Orlando Pereira y María José Parajeles 
"""
#%% Packages
from . import auxiliary_functions as auxfcns
import csv
from collections import OrderedDict
import glob
import io
import math
from matplotlib import pyplot as plt
import networkx as nx
import numpy as np
import pandas as pd
pd.options.mode.chained_assignment = None
import pickle
import os
import random
import time
from . import trafoOperations_spyder as trafOps
from qgis.PyQt.QtWidgets import QMessageBox
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

def loadFeederLoadCurve(load_curve_path, load_curve_name):
    with io.open(os.path.join(load_curve_path,load_curve_name+'.csv'), 'rt', encoding = "ascii") as workbook:
        reader = csv.reader(workbook)                    
        next(reader)                    
        circuit_demand = [[row[3], row[2], row[0], row[1]] for row in reader]  #day, hour, P (kW), Q (kVAr)
        workbook.closed
    return circuit_demand

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
    print("Directorio master = ", str(dss_network + '/Master.dss'))
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
    
    DSScircuit.Solution.Solve()  # Solve the circuit

    VBuses_b = pd.DataFrame(list(DSScircuit.AllBusVmag),
                            index=list(DSScircuit.AllNodeNames),
                            columns=['VOLTAGEV'])
    VBuses_b=VBuses_b[~VBuses_b.index.str.contains('aftermeter')]
    VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_der')]
    VBuses_b=VBuses_b[~VBuses_b.index.str.contains('_swt')]

    base_vals = [120, 138, 208, 240, 254, 277, 416, 440, 480,
                        2402, 4160, 7620, 7967, 13200, 13800, 14380,
                        19920, 24900, 34500, 79670, 132790]

    Base_V = pd.DataFrame()
    Base_V['BASE'] = VBuses_b['VOLTAGEV'].apply(lambda v : base_vals[[abs(v-i) for i in base_vals].index(min([abs(v-i) for i in base_vals]))])
    Base_V['BUSNAME'] = Base_V.index
    Base_V['LV_GROUP'] = Base_V.BUSNAME.apply(lambda g : buslvgroup(DSScircuit,g, lv_groups) )
    Base_V['TX']= Base_V.LV_GROUP.apply(lambda g : txname(g, tx_groups) )
    del Base_V['BUSNAME']
    path = dss_network + '/bases.csv'
    # Base_V.to_csv(path)

    return VBuses_b, Base_V

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

    #%% LoadAllocation Simulation
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time                
    DSStext.Command = 'New Monitor.HVMV_PQ_vs_Time line.' + firstLine + ' 1 Mode=1 ppolar=0'  # Monitor in the transformer secondary side to monitor P and Q
    
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
    PVNames = DSScircuit.PVSystems.AllNames
    if GenNames[0] != 'NONE':
        for i in GenNames: # extract power from generators

            DSScircuit.setActiveElement('generator.' + i)
            p = DSScircuit.ActiveElement.Powers

            for w in range(0, len(p), 2):
                gen_p += -p[w] # P
                gen_q += -p[w + 1] # Q
        gen_powers[0] += gen_p
        gen_rpowers[0] += gen_q
    if PVNames[0] != 'NONE':
        for i in PVNames: # extract power from PVSystems
            DSScircuit.setActiveElement('PVSystem.' + i)
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
                                      P_to_be_matched, Q_to_be_matched, hora_sec, study,
                                      dss_network, tx_modelling, 1, firstLine, substation_type,
                                      line_tx_definition, gen_powers, gen_rpowers)

    if DSScircuit == 0 or DSScircuit == -1:
        msg = "Favor verifique que la fecha introducida"
        msg += " esté en la curva de demanda"
        title = "Error"
        QMessageBox.critical(None, title, msg)
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(),\
		   pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    DSSobj.AllowForms = 1
    t2= time.time()
    print('Loadallocationtime: '+str(t2-t1))
    #%% Post load allocation simulation
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time

    DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
    DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
    DSScircuit.Solution.Solve()
    
    DSScircuit.setActiveElement('line.' + firstLine)
    temp_powers = DSScircuit.ActiveElement.Powers
    numb = str(temp_powers[0] + temp_powers[2] + temp_powers[4])
    print('P_wo_Sub: ' + numb)
    print(temp_powers)
    
    # DSScircuit.setActiveElement('line.MV3PGUA3189')
    # temp_powers = DSScircuit.ActiveElement.Powers
    # print('P_GEN_3189: '+ str(temp_powers[0]+temp_powers[2]+temp_powers[4]))

    # DSScircuit.setActiveElement('line.MV3PGUA1061')
    # temp_powers = DSScircuit.ActiveElement.Powers
    # print('GEN_1061: '+ str(temp_powers[0]+temp_powers[2]+temp_powers[4]))

    #Vpu
    VBuses = pd.DataFrame(list(DSScircuit.AllBusVmag),
						  index = list(DSScircuit.AllNodeNames),
						  columns=['VOLTAGEV'])
    VBuses=VBuses[~VBuses.index.str.contains('aftermeter')]
    VBuses=VBuses[~VBuses.index.str.contains('_der')]
    VBuses=VBuses[~VBuses.index.str.contains('_swt')]
    No_DERs_run_Vbuses = pd.DataFrame()
    No_DERs_run_Vbuses['VOLTAGE'] = VBuses.VOLTAGEV/Base_V.BASE
    
    #LV layer update 
    lv_loads_layer['kVA_snap'] = lv_loads_layer.DSSName.apply(lambda x : get_kva_load(DSScircuit, x))
    tx_layer['kVA_snap'] = tx_layer.DSSName.apply(lambda x : get_kva_trafos(DSScircuit, x))
    if mv_loads_layer.empty is False:
        mv_loads_layer['kVA_snap'] = mv_loads_layer.DSSName.apply(lambda x : get_kva_load(DSScircuit, x))

            
    t3= time.time()
    print('Basecase snapshot time: '+str(t3-t2))
    
    
        #%% Fault type studies
        # defines the terminals necessary for each type of fault. Dictionary containing: Fault type: [terminal conection for bus 1, terminal conection for bus2, number of phases]
    faulttypeterminals = {'ABC':['.1.2','.3.3', '2'],
                          'ABCG':['.1.2.3','.0.0.0','3'],
                          'AB':['.1','.2','1'], 'BC':['.2','.3','1'],
                          'AC':['.1','.3','1'],
                          'ABG':['.1.1','.2.0','2'],
                          'BCG':['.2.2','.3.0','2'],
                          'ACG':['.1.1','.3.0','2'],
                          'AG':['.1','.0','1'], 'BG':['.2','.0','1'],
                          'CG':['.3','.0','1']}

    if (FF_analysis or BFC_analysis) is True:
        ### FORWARD CURRENT
        
        if CircuitBreakDvRoR.empty or CircuitBreakDvFF_BFC.empty:  # Revisar ?
            mltidx = pd.DataFrame()
        else:
            mltidx = pd.MultiIndex.from_tuples(zip(CircuitBreakDvFF_BFC['Element'],
                                               CircuitBreakDvFF_BFC['BusInstalled']),
                                               names=['Element', 'FaultedBus'])
        No_DER_FFCurrents = pd.DataFrame(index=mltidx, columns=list(faulttypes))
        #BFC
        temp=CircuitBreakDvFF_BFC[CircuitBreakDvFF_BFC['RecloserElement'].notnull()]
        mltidx = pd.MultiIndex.from_tuples(zip(temp['Element'],
										   temp['RecloserElement'],
										   temp['BusInstalled']),
										   names=['FuseElement','BreakerElement', 'FaultedBus'])
        No_DER_BFCCurrents = pd.DataFrame(index=mltidx, columns=list(faulttypes))
        del temp
        
        for element, faultedbus in list(No_DER_FFCurrents.index):
            print('FF: '+element)
            ftlist=faultedbus.split('.')
            if len(ftlist) == 2:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2],
									 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,2],
									 'AG':[0], 'BG':[0], 'CG':[0]}
                if ftlist[1] == '1':
                    faultlabels = ['AG']
                elif ftlist[1] == '2':
                    faultlabels = ['BG']
                else: 
                    faultlabels = ['CG']
            elif len(ftlist) == 3:
                currentspositions = {'AB': [0,2], 'BC': [0,2], 'AC': [0,2],
									 'ABG': [0,2], 'BCG': [2,4], 'ACG': [0,2],
									 'AG':[0], 'BG':[0], 'CG':[0]}
                if ftlist[1] == '1' and ftlist[2] == '2':
                    faultlabels = ['AB', 'ABG','AG','BG']
                elif ftlist[1] == '2' and ftlist[2] == '3':
                    faultlabels = ['BC', 'BCG','CG','BG']
                else: 
                    faultlabels = ['AC', 'ACG','AG','CG']
            else:
                currentspositions = {'ABC':[0,2,4], 'ABCG':[0,2,4],
									 'AB':[0,2], 'BC':[2,4], 'AC':[0,4],
									 'ABG':[0,2], 'BCG':[2,4],
									 'ACG':[0,4], 'AG':[0], 'BG':[2],
									 'CG':[4]}
                faultlabels = list(faulttypeterminals.keys())
    
            faultlabels = list(set(faulttypes)&set(faultlabels)) # removes those types of faults not being evaluated
    
            for faulttype in faultlabels:
                print('FF: ' + element + ', Fault: ' + faulttype)
                # Machines inicialization simulation
                DSStext.Command = 'clear'
                DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
                DSStext.Command = 'Set mode=daily'  # Type of Simulation
                DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
                DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                DSScircuit.Solution.Solve()
    
                #%% Fault study (dynamics) simulation
                # type of fault
                term1 = faulttypeterminals[faulttype][0]
                term2 = faulttypeterminals[faulttype][1]
                phasessc = faulttypeterminals[faulttype][2]
    
                DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                com = 'new fault.fault_' + faulttype + ' bus1='
                com += ftlist[0] + term1 + ' bus2=' + ftlist[0]
                com += term2 + ' phases=' + phasessc
                DSStext.Command = com
                DSScircuit.Solution.Solve()
    
                DSScircuit.setActiveElement(element) # faulted element
                temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                tmp = [temp_currents[i] for i in currentspositions[faulttype]]
                No_DER_FFCurrents[faulttype][element, faultedbus] = tmp
    
                #BFC
                breakerelement = CircuitBreakDvFF_BFC.loc[CircuitBreakDvFF_BFC['Element'] == element,
														  'RecloserElement'].iloc[-1]
                if not np.isnan(breakerelement):
                    No_DER_BFCCurrents[faulttype][element, breakerelement, faultedbus] = []
                    for element_i in (element, breakerelement):
                        DSScircuit.setActiveElement(element_i) # faulted element
                        temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                        tmp = [temp_currents[i] for i in currentspositions[faulttype]]
                        No_DER_BFCCurrents[faulttype][element, breakerelement, faultedbus].append(tmp)

    else:             
        No_DER_FFCurrents = pd.DataFrame()
        No_DER_BFCCurrents = pd.DataFrame()

        ### REDUCTION OF REACH
    if RoR_analysis is True:
        elements = []; FaultedBus = []
        for element in list(CircuitBreakDvRoR['Element']):
            if CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element'] == element, 'FurthestBusZone'].iloc[-1] == CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusBackUp'].iloc[-1]:
                elements.append(element)
                FaultedBus.append(CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element'] == element,
													    'FurthestBusZone'].iloc[-1])
            else:
                elements.append(element); elements.append(element)
                FaultedBus.append(CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element'] == element, 'FurthestBusZone'].iloc[-1])
                FaultedBus.append(CircuitBreakDvRoR.loc[CircuitBreakDvRoR['Element']==element, 'FurthestBusBackUp'].iloc[-1])

        mltidx = pd.MultiIndex.from_tuples(zip(elements,FaultedBus),
										   names=['Element', 'FaultedBus'])
        No_DER_RoRCurrents = pd.DataFrame(index=mltidx, columns=list(faulttypes))

        for element, faultedbus in list(No_DER_RoRCurrents.index):
            print('ROR: '+element)
            ftlist=faultedbus.split('.')
            if len(ftlist) == 2:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2],
									 'ABG':[0,2], 'BCG':[2,4],
									 'ACG':[0,2], 'AG':[0], 'BG':[0],
									 'CG':[0]}
                if ftlist[1] == '1':
                    faultlabels = ['AG']
                elif ftlist[1] == '2':
                    faultlabels = ['BG']
                else: 
                    faultlabels = ['CG']
            elif len(ftlist) == 3:
                currentspositions = {'AB':[0,2], 'BC':[0,2], 'AC':[0,2],
									 'ABG':[0,2], 'BCG':[2,4],
									 'ACG':[0,2], 'AG':[0], 'BG':[0],
									 'CG':[0]}
                if ftlist[1] == '1' and ftlist[2] == '2':
                    faultlabels = ['AB', 'ABG','AG','BG']
                elif ftlist[1] == '2' and ftlist[2] == '3':
                    faultlabels = ['BC', 'BCG','CG','BG']
                else: 
                    faultlabels = ['AC', 'ACG','AG','CG']
            else:
                currentspositions = {'ABC':[0,2,4], 'ABCG':[0,2,4],
									 'AB':[0,2], 'BC':[2,4], 'AC':[0,4],
									 'ABG':[0,2], 'BCG':[2,4], 'ACG':[0,4],
									 'AG':[0], 'BG':[2], 'CG':[4]}
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
                DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
                DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
                DSScircuit.Solution.Solve()

                #%% Fault study (dynamics) simulation
                # type of fault
                term1 = faulttypeterminals[faulttype][0]
                term2 = faulttypeterminals[faulttype][1]
                phasessc = faulttypeterminals[faulttype][2]
    
                DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
                com = 'new fault.fault_' + faulttype + ' bus1='
                com += ftlist[0] + term1 + ' bus2=' + ftlist[0]
                com += term2 + ' phases=' + phasessc
                DSStext.Command = com
                DSScircuit.Solution.Solve()
    
                DSScircuit.setActiveElement(element) # faulted element
                temp_currents = DSScircuit.ActiveElement.CurrentsMagAng
                tmp = [temp_currents[i] for i in currentspositions[faulttype]]
                No_DER_RoRCurrents[faulttype][element, faultedbus] = tmp

        t4= time.time()
        print('Basecase fault time: '+str(t4-t3))
        
    else:
        No_DER_RoRCurrents = pd.DataFrame()

    return No_DERs_run_Vbuses, No_DER_FFCurrents, No_DER_RoRCurrents,\
		   No_DER_BFCCurrents, kW_sim, kVAr_sim, lv_loads_layer

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
    dss_name = loadsLV_file_name + '.dss'
    filename = os.path.join(dss_network, dss_name)
    with open(filename, 'r') as f:
        loadslv_file = f.readlines()
    try:
        filename = os.path.join(dss_network, loadsMV_file_name + '.dss')
        with open(filename,'r') as f:
            loadsmv_file = f.readlines()
    except:
        loadsmv_file=list()
    
    loadslv_buses = list() 
    for loadline in loadslv_file:
        busld = loadline.split(' ')[2].split('=')[1].split('.')
        bus_list = [busld[0].lower() + '.' + busld[i] for i in range(1,len(busld))]
        loadslv_buses = loadslv_buses + bus_list
    loadsmv_buses = list() 
    for loadline in loadsmv_file:
        busld = loadline.split(' ')[2].split('=')[1].split('.')
        bus_list = [busld[0].lower() + '.' + busld[i] for i in range(1, len(busld))]
        loadsmv_buses = loadsmv_buses + bus_list
    return loadslv_buses, loadsmv_buses

#%% FINDS GD CAPACITY ALREADY INSTALLED
# Input: name_file_created - GD file 
# Output: DERinstalled_buses : list(busname)

def getGDinstalled(name_file_created):
    DERinstalled_buses = []  # forbidden buses --> buses with DER already installed
    DERinstalled=0
    try:
        gd_file = open(name_file_created.split('_')[0] + '_DG.dss')  # DG file reading
        gds = gd_file.read().split('\n')
        for i in gds:
            try:
                DERinstalled_buses.append(i.split(" ")[2].replace("bus1=", "").split(".")[0])  # forbidden bus
                DERinstalled += float(i.split('kW=')[1].split(" ")[0])
            except IndexError:
                pass
    except IOError:
        pass
    return DERinstalled_buses, DERinstalled

#%% DERs SNAPSHOT RUN 
# Input: DSStext, DSScircuit, snapshotdate, snapshottime, firstLine, tx_modelling, substation_type, line_tx_definition, circuit_demand, Base_V, kW_sim, kVar_sim, DERs
# Output: NONE, just runs the simulation and updates OPENDSS instances

def DERs_Run(DSStext, DSScircuit, snapshotdate, snapshottime,
			 firstLine, tx_modelling, substation_type, line_tx_definition,
			 circuit_demand, Base_V, kW_sim, kVAr_sim, DERs,
			 Trafos, dss_network): # time: hh:mm
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
    hora_sec = snapshottime.split(':')
    
    #%% simulation
    DSStext.Command = 'clear'
    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
    DSStext.Command = 'Set mode=daily'  # Type of Simulation
    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                          
    DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
    DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
    for tx in Trafos:
        DSStext.Command = tx
    for der in DERs:
        DSStext.Command = der
    
    DSScircuit.Solution.Solve()

#%% THREE PHASE MV BUSES TO ALLOCATE LARGE SCALE DER - finds the buses every 100 m to intall DER of large scale
# Input: nodos_mt, lineas_mt, trafos, fixed_distance
# Output: chosen_buses_dict: dict with distance every 100 m as keys and a tuple as value (BUSMV, NOMVOLTMV), chosen_buses_list: list of all buses for every distance
def distance_nodes_MV(G, nodes_mv, lineas_mt, trafos, fixed_distance):
    # Empieza algoritmo
    
    rand_bus = list(G.nodes.keys())[0]
    ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    first_bus = 'BUSMV' + ckt_name + '1'
    
    mt_3bus_list = nx.get_node_attributes(G,'num_phases') #list of 3phase mv buses on the whole circuit
    remove = [k for k in mt_3bus_list if mt_3bus_list[k] != '3']
    for k in remove: del mt_3bus_list[k]
    
    df_distancia = pd.DataFrame(np.nan, index=list(G.nodes), columns=['nodes', 'distance'])
    df_distancia['nodes'] = list(G.nodes)
    df_distancia['distance'] = df_distancia['nodes'].apply(lambda b: nx.shortest_path_length(G, first_bus, b, weight='distance'))
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
        
        if temp_df.empty == False:
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
                            tmp = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = tmp
                        except:
                            dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]

                        tmp = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
                        dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = tmp

                    else:
                        excluded_from_df = list(set(downstream_node_3ph_studied) - set(temp_df.index.values))

                        if len(excluded_from_df) > 0:
                            # target_list.append(node)
                            number_paths += 1

                            dist_dict[(n+1)*fixed_distance][number_paths] = {}
                            dist_dict[(n+1)*fixed_distance][number_paths]['source'] = u_source
                            dist_dict[(n+1)*fixed_distance][number_paths]['target'] = node

                            edges_calc = [x for x in nx.all_simple_edge_paths(G, u_source, node)]

                            try:
                                tmp = [G.get_edge_data(x[0],x[1])['line_name'] for x in edges_calc[0]]
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] = tmp
                            except:
                                dist_dict[(n+1)*fixed_distance][number_paths]['edges'] =[]

                            lst = list(set(nx.shortest_path(G, u_source,node)) - set([u_source]))
                            dist_dict[(n+1)*fixed_distance][number_paths]['nodes'] = lst

    chosen_buses_dict = {}
    chosen_buses_df = pd.DataFrame(np.nan, index=[],
								   columns=['distance', 'nomvolt'])
    chosen_buses_list = []
    
    for step in dist_dict:
        chosen_buses_dict[step] = []
        step_pd = pd.DataFrame(np.nan, index=[],
							   columns = ['source', 'target', 'distance'])
        source_list = []
        for key in dist_dict[step]:
            step_pd.loc[key,'source'] = dist_dict[step][key]['source']
            step_pd.loc[key,'target'] = dist_dict[step][key]['target']

        step_pd['distance'] = step_pd.apply(lambda b: nx.shortest_path_length(G, b.source, b.target, weight='distance'), axis=1)
        source_list = list(set(step_pd['source'].values))

        for s in source_list:
            id_max = step_pd.loc[step_pd['source']==s]['distance'].idxmax()
            info_tuple = (step_pd.loc[id_max,'target'], nodes_mv.loc[step_pd.loc[id_max,'target'], 'NOMVOLT'])
            chosen_buses_list.append(info_tuple)
            chosen_buses_dict[step].append(info_tuple)
            chosen_buses_df.loc[info_tuple[0], 'distance'] = step
            chosen_buses_df.loc[info_tuple[0], 'nomvolt'] = info_tuple[1]

    return chosen_buses_dict, chosen_buses_list, chosen_buses_df

#%% THREE PHASE MV BUSES TO ALLOCATE LARGE SCALE DER - finds the buses every 100 m to intall DER of large scale
# Input: nodos_mt, lineas_mt, trafos, fixed_distance
# Output: chosen_buses_dict: dict with distance every 100 m as keys and a tuple as value (BUSMV, NOMVOLTMV), chosen_buses_list: list of all buses for every distance

# def tri_ph_dist(G, bus):
    
#     #Upstream buses calculation 
#     us_bus = list(nx.ancestors(G, bus)) + [bus]
    
#     mt_3_upstream = {}
#     for u_bus in us_bus:
#         mt_3_upstream[u_bus] = [x for x,y in G.nodes(data=True) if y['num_phases']=='3']
#     #     mt_3_upstream[u_bus] = G.nodes[u_bus]['num_phases']
    
#     # remove = [k for k in mt_3_upstream if mt_3_upstream[k] != '3']
#     # for k in remove: del mt_3_upstream[k]
        
#     bus_study = pd.DataFrame(np.nan, columns=['distance'], index = list(mt_3_upstream.keys()))
#     for tri_bus in mt_3_upstream:
#         try:
#            bus_study.loc[tri_bus,'distance'] = float(nx.shortest_path_length(G, tri_bus, bus, weight='distance'))
#         except:
#             bus_study.loc[tri_bus, 'distance'] = np.nan
            
#     # print()
#     return bus_study.loc[bus_study['distance'] == bus_study['distance'].min()].index.values[0]

def tri_ph_dist(G, bus):
    
    #Upstream buses calculation 
    us_bus = list(nx.ancestors(G, bus)) + [bus]
    mt_3bus_list = [x for x,y in G.nodes(data=True) if y['num_phases']=='3']
    mt_3_upstream = list(set(us_bus).intersection(mt_3bus_list))
    
    bus_study = pd.DataFrame(np.nan, columns=['u_bus', 'distance'], index = mt_3_upstream)
    bus_study['u_bus'] = mt_3_upstream
    bus_study['distance'] = bus_study['u_bus'].apply(lambda b: float(nx.shortest_path_length(G, b, bus, weight='distance')))
            
    # print(bus)
    return bus_study.loc[bus_study['distance'] == bus_study['distance'].min()].index.values[0]


def data_grouping(G, MV_hist_df , fixed_distance, final_der,
				  mv_lines_layer, lines_mv_oh_layer_original,
				  lines_mv_ug_layer_original, Voltage_comp,
				  Thermal_analysis, Prot_comp, lim_kVA):
    if final_der == [0] or final_der == 0:
        msg = "Por favor disminuya el aumento de instalaciones "
        msg += "(step) seleccionado "
        title = "Error"
        QMessageBox.critical(None, title, msg)
        pd.DataFrame(), pd.DataFrame(), ""
        
    rand_bus = list(G.nodes.keys())[0]
    ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    first_bus = 'BUSMV' + ckt_name + '1'
    
    ###############################################################################################
    # AGGRUPATION CODE
    
    # find the first 3ph upstream bus per node
    distance_to_3ph_bus = MV_hist_df[['DSSName', 'bus1', final_der]]
    distance_to_3ph_bus.loc[:,'min_bus'] = distance_to_3ph_bus.loc[:,'bus1'].apply(lambda b: tri_ph_dist(G, b))
    # Getting the list of the selected 3ph buses
    list_3phase_nodes = list(set(distance_to_3ph_bus.loc[distance_to_3ph_bus[final_der]>0]['min_bus']))
    
    accumulated_kVA = pd.DataFrame(np.nan, index=[], columns=['bus', 'kVA_val'])
    accumulated_kVA['bus'] = list_3phase_nodes
    accumulated_kVA['kVA_val'] = accumulated_kVA['bus'].apply(lambda x: distance_to_3ph_bus.loc[distance_to_3ph_bus['min_bus'] == x][final_der].sum())
    accumulated_kVA.index = accumulated_kVA['bus']
    

    mt_3bus_list = [x for x,y in G.nodes(data=True) if y['num_phases']=='3']
    
    df_distancia = pd.DataFrame(np.nan, index=mt_3bus_list, columns=['nodes', 'distance'])
    df_distancia['nodes'] = mt_3bus_list
    df_distancia['distance']=df_distancia['nodes'].apply(lambda b: nx.shortest_path_length(G, first_bus, b, weight='distance'))
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
        
        if temp_df.empty == False:
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
                path_dict[source_key]['kVA'] += accumulated_kVA.loc[nodes_in_accKVAdf, 'kVA_val'].sum()
            except:
                pass

        # for source_key in path_dict:
        #     nodes_in_oh_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_oh_layer_original['DSSName'].values))
        #     idx_oh_list = lines_mv_oh_layer_original.loc[lines_mv_oh_layer_original['DSSName'].isin(nodes_in_oh_layer)].index.values
        #     lines_mv_oh_layer_original.loc[idx_oh_list, 'kVA_sep'+str(fixed_distance)] = path_dict[source_key]['kVA']
            
        #     nodes_in_ug_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_ug_layer_original['DSSName'].values))
        #     # print(source_key, nodes_in_ug_layer)
        #     idx_ug_list = lines_mv_ug_layer_original.loc[lines_mv_ug_layer_original['DSSName'].isin(nodes_in_ug_layer)].index.values
        #     lines_mv_ug_layer_original.loc[idx_ug_list, 'kVA_sep'+str(fixed_distance)] = path_dict[source_key]['kVA']
            
        if lim_kVA== True:    
            if (Voltage_comp == True and Thermal_analysis== True and Prot_comp== True):
                name_col = 'HB_M-VTP'
            elif (Voltage_comp == True and Thermal_analysis== True and Prot_comp== False):
                name_col = 'HB_M-VT'
            elif (Voltage_comp == True and Thermal_analysis== False and Prot_comp== False):
                name_col = 'HB_M-V'
        
        else: 
            if (Voltage_comp == True and Thermal_analysis== True and Prot_comp== True):
                name_col = 'HB_VTP'
            elif (Voltage_comp == True and Thermal_analysis== True and Prot_comp== False):
                name_col = 'HB_VT'
            elif (Voltage_comp == True and Thermal_analysis== False and Prot_comp== False):
                name_col = 'HB_V'

        lines_mv_oh_layer_original.loc[:, name_col] = np.nan
        try:
            lines_mv_ug_layer_original.loc[:, name_col] = np.nan
        except:
            pass

        for source_key in path_dict:
            nodes_in_oh_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_oh_layer_original['DSSName'].values))
            idx_oh_list = lines_mv_oh_layer_original.loc[lines_mv_oh_layer_original['DSSName'].isin(nodes_in_oh_layer)].index.values
            lines_mv_oh_layer_original.loc[idx_oh_list, name_col] = path_dict[source_key]['kVA']
            
            try:
                nodes_in_ug_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_ug_layer_original['DSSName'].values))
                # print(source_key, nodes_in_ug_layer)
                idx_ug_list = lines_mv_ug_layer_original.loc[lines_mv_ug_layer_original['DSSName'].isin(nodes_in_ug_layer)].index.values
                lines_mv_ug_layer_original.loc[idx_ug_list, name_col] = path_dict[source_key]['kVA']

            except:
                pass

    return lines_mv_oh_layer_original, lines_mv_ug_layer_original, name_col

#%%
def data_grouping_iterative(G, chosen_buses_df, fixed_distance,
							mv_lines_layer, lines_mv_oh_layer_original,
							lines_mv_ug_layer_original, Voltage_comp,
							Thermal_analysis, Prot_comp):

    rand_bus = list(G.nodes.keys())[0]
    ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    first_bus = 'BUSMV'+ckt_name+'1'

    mt_3bus_list = nx.get_node_attributes(G,'num_phases') #list of 3phase mv buses on the whole circuit
    remove = [k for k in mt_3bus_list if mt_3bus_list[k] != '3']
    for k in remove: del mt_3bus_list[k]

    df_distancia = pd.DataFrame(np.nan, index=list(G.nodes), columns=['nodes', 'distance'])
    df_distancia['nodes'] = list(G.nodes)
    df_distancia['distance']=df_distancia['nodes'].apply(lambda b: nx.shortest_path_length(G, first_bus, b, weight='distance'))
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
        
        if temp_df.empty == False:
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
            
        name_col = ""
        if (Voltage_comp == True and Thermal_analysis== True and Prot_comp== True):
            name_col = 'IT_DER_VTP'
        elif (Voltage_comp == True and Thermal_analysis== True and Prot_comp== False):
            name_col = 'IT_DER_VT'
        elif (Voltage_comp == True and Thermal_analysis== False and Prot_comp== False):
            name_col = 'IT_DER_V'

        lines_mv_oh_layer_original.loc[:, name_col] = np.nan
        try:
            lines_mv_ug_layer_original.loc[:, name_col] = np.nan
        except:
            pass

        for source_key in path_dict:
            nodes_in_oh_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_oh_layer_original['DSSName'].values))
            idx_oh_list = lines_mv_oh_layer_original.loc[lines_mv_oh_layer_original['DSSName'].isin(nodes_in_oh_layer)].index.values
            lines_mv_oh_layer_original.loc[idx_oh_list, name_col] = path_dict[source_key]['kVA']
            
            try:
                nodes_in_ug_layer = list(set(path_dict[source_key]['edges']).intersection(lines_mv_ug_layer_original['DSSName'].values))
                # print(source_key, nodes_in_ug_layer)
                idx_ug_list = lines_mv_ug_layer_original.loc[lines_mv_ug_layer_original['DSSName'].isin(nodes_in_ug_layer)].index.values
                lines_mv_ug_layer_original.loc[idx_ug_list, name_col] = path_dict[source_key]['kVA']
            
            except:
                pass

    return lines_mv_oh_layer_original, lines_mv_ug_layer_original, name_col

#%% WRITE DER AND STEPUP TRANSFORMER OPENDSS SENTENCES - allocates DER in bus list and assigns corresponding MV-LV transformer
# Input: bus_list, installed_capacity: DER size
# Output: Trafos: tx sentences, Trafos_Monitor, DERs: DERs'sentences
def trafos_and_DERs_text_command(bus_list, installed_capacity):

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
        busLV = 'BUSLV'+busMV.replace('BUSMV', '')+'_DER'
        confMV = 'wye'
        confLV = 'wye'

        tap = '1'

        impedance = trafOps.impedanceSingleUnit(cantFases, kV_MedLL, kV_LowLN, kVA).get('Z')
        noloadloss = trafOps.impedanceSingleUnit(cantFases, kV_MedLL, kV_LowLN, kVA).get('Pnoload')
        imag = trafOps.impedanceSingleUnit(cantFases, kV_MedLL, kV_LowLN, kVA).get('Im')
        trafName =  'TX_DER_'+busMV.replace('BUSMV', '')   
        derName = 'DER_'+busMV.replace('BUSMV', '')

        trafo_line = 'new transformer.' + trafName + ' phases=3 windings=2 ' + noloadloss + " " + imag + ' buses=[' + busMV + '.1.2.3 '
        trafo_line += busLV + '.1.2.3]' + ' conns=[' + confMV + ' ' + confLV + ']' + ' kvs=[' + kV_MedLL + " " +  kV_LowLL + ']'
        trafo_line += ' kvas=[' + kVA + " " + kVA + '] ' + impedance + ' Taps=[' + tap + ', 1]' + normhkva

        Trafos.append(trafo_line) # Se añade a la lista de Trafos

        trafo_line_monitor = "new monitor.Mon" + trafName + " Element=transformer." + trafName + " Terminal=1 Mode=1\n"

        Trafos_Monitor.append(trafo_line_monitor) # Se añade a la lista de monitores de trafos

        der_line = 'new generator.DER'+ derName + ' bus1='+ busLV 
        der_line += '.1.2.3 kV='+ kV_LowLL + ' phases=3 kW=' + str(float(installed_capacity)) 
        der_line += ' PF=1 conn=wye kVA=' +str(round(float(installed_capacity)*1.2,2))
        der_line += ' Model=7 Vmaxpu=1.5 Vminpu=0.83 Balanced=no Enabled=yes' 

        DERs.append(der_line)

    return Trafos, Trafos_Monitor, DERs

#%% POWER QUALITY DER INTEGRATION EFFECT (overvoltage and voltage deviation) - verification, in all load buses, by phase. It updates the results data frames for every monte carlo and every generation capacity installation
# Overvoltage means magnitude in any phase goes above 1.05 p.u
# Maximum voltage deviation in MT is 3% (based on IEEE Std. 1453-2015) and in BT is 5% (based on EPRI studies)
# Input: DSScircuit, DSStext, dss_network,  Base_V, loadslv_buses, loadsmv_buses, capacity_i, NoDERsPF_Vbuses, Overvoltage_loads_df, Overvoltage_rest_df, Voltagedeviation_loads_df, Voltagedeviation_rest_df
# Output: Overvoltage_loads_df: pd.DataFrame({ 'GDKWP':[capacity_i],'BUS':[busi_l],'Voltage' : [vbusi_l], 'POVERVOLTAGE' : [povervi_l] } ), Voltagedeviation_loads_df, Overvoltage_rest_df, Voltagedeviation_rest_df, all have the same format

#%% FIND THREE PHASE VOLTAGE UNBALANCE
def voltageunbalance(DSScircuit, busname):
    DSScircuit.setActiveBus(busname.split('.')[0])
    
    if len(DSScircuit.ActiveBus.Nodes) == 3:
        # print(busname, DSScircuit.ActiveBus.SeqVoltages[1])
        unbt = round((DSScircuit.ActiveBus.SeqVoltages[2] / DSScircuit.ActiveBus.SeqVoltages[1]),4)
    else:
        unbt = np.nan
    return unbt
#%%

def pq_voltage(DSScircuit, DSStext, dss_network, Base_V, loadslv_buses,
			   loadsmv_buses, RegDevices, capacity_i, lv_groups, tx_groups,
			   NoDERsPF_Vbuses, Overvoltage_loads_df, Overvoltage_rest_df,
			   Voltagedeviation_loads_df, Voltagedeviation_rest_df,
			   Voltagedeviation_reg_df, Voltageunbalance_df,
			   Overvoltage_analysis, VoltageDeviation_analysis,
			   VoltageRegulation_analysis, VoltageUnbalance): 

    #%% PU CALCULATION
    t1=time.time()
    VBuses = pd.DataFrame(list(DSScircuit.AllBusVmag),
						  index=list(DSScircuit.AllNodeNames),
						  columns=['VOLTAGEV'])
    VBuses=VBuses[~VBuses.index.str.contains('aftermeter')]
    VBuses=VBuses[~VBuses.index.str.contains('_der')]
    VBuses=VBuses[~VBuses.index.str.contains('_swt')]
    V_buses = pd.DataFrame()
    V_buses['VOLTAGE'] = VBuses.VOLTAGEV/Base_V.BASE
    V_buses['LV_GROUP']=Base_V.LV_GROUP
    V_buses['TX']=Base_V.TX
    t2=time.time()

    print('V_buses_pu: ' +str(t2-t1))
    
    #%% OVERVOLTAGE MONITORING - Loads nodes and rest of nodes
    if Overvoltage_analysis == True:
        #loads

        V_buses_loads = V_buses[V_buses.index.isin(loadslv_buses+loadsmv_buses)]; #pu loads buses dataframe
        vbusi_l=V_buses_loads.VOLTAGE.max()
        busi_l=V_buses_loads.VOLTAGE.idxmax()
        povervi_l=V_buses_loads.VOLTAGE.gt(1.05).mean()
        #rest
        V_buses_rest = V_buses[~V_buses.index.isin(loadslv_buses+loadsmv_buses)]; #pu rest of buses dataframe
        vbusi_r=V_buses_rest.VOLTAGE.max(); busi_r=V_buses_rest.VOLTAGE.idxmax()
        povervi_r=V_buses_rest.VOLTAGE.gt(1.05).mean()
        
        OVl_List = list(set(V_buses_loads.loc[V_buses_loads['VOLTAGE'].gt(1.05), 'TX'].dropna()))
        OVr_List = list(set(V_buses_rest.loc[V_buses_rest['VOLTAGE'].gt(1.05), 'TX'].dropna()))
        # loadsmv_buses
        if not Overvoltage_loads_df.empty:
            OVl_List = list(set(list(Overvoltage_loads_df['BLACKLIST_TX'].values[-1]) + OVl_List))
            OVr_List = list(set(list(Overvoltage_rest_df['BLACKLIST_TX'].values[-1]) + OVr_List))
                
        if V_buses[V_buses['VOLTAGE'].gt(1.05)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if any(V_buses[V_buses['VOLTAGE'].gt(1.05)].index.str.startswith('busmv')) else False  
        
        Overvoltage_loads_df = Overvoltage_loads_df.append(pd.DataFrame({'GDKWP':[capacity_i],'BUS':[busi_l],'VOLTAGE' : [vbusi_l], 'POVERVOLTAGE' : [povervi_l],  'BLACKLIST_TX' : [OVl_List], 'HCSTOP':[HCSTOP] } ), ignore_index=True)
        Overvoltage_rest_df = Overvoltage_rest_df.append(pd.DataFrame({'GDKWP':[capacity_i],'BUS':[busi_r],'VOLTAGE' : [vbusi_r], 'POVERVOLTAGE' : [povervi_r], 'BLACKLIST_TX' : [OVr_List], 'HCSTOP':[HCSTOP] } ), ignore_index=True)
        t3=time.time()
        print('Overvoltage processing: ' +str(t3-t2))
    
    #%% VOLTAGE DEVIATION MONITORING (5% in LV, 3% in MV). For regulation nodes, a HCSTOP in 1 is reported if V excceeds 1/2 the bandwith of the regulator
    
    if VoltageDeviation_analysis == True:
        #loads
        NoDERsPF_MVbuses_loads = NoDERsPF_Vbuses[NoDERsPF_Vbuses.index.isin(loadsmv_buses)] # pu loads buses dataframe at base power flow (no DER)
        NoDERsPF_LVbuses_loads = NoDERsPF_Vbuses[NoDERsPF_Vbuses.index.isin(loadslv_buses)]
        MV_buses_loads = V_buses[V_buses.index.isin(loadsmv_buses)]  # pu loads buses dataframe
        LV_buses_loads = V_buses[V_buses.index.isin(loadslv_buses)]  # pu loads buses dataframe
        devLV_loads = pd.DataFrame( zip( abs( (LV_buses_loads.VOLTAGE-NoDERsPF_LVbuses_loads.VOLTAGE)/NoDERsPF_LVbuses_loads.VOLTAGE ), Base_V.LV_GROUP, Base_V.TX ), index=NoDERsPF_LVbuses_loads.index, columns=['VDEVIATION', 'LV_GROUP', 'TX'])
        devMV_loads = pd.DataFrame( zip( abs( (MV_buses_loads.VOLTAGE-NoDERsPF_MVbuses_loads.VOLTAGE)/NoDERsPF_MVbuses_loads.VOLTAGE ), Base_V.LV_GROUP, Base_V.TX ), index=NoDERsPF_MVbuses_loads.index, columns=['VDEVIATION', 'LV_GROUP', 'TX'])
    
        #rest of buses
        NoDERsPF_MVbuses_rest = NoDERsPF_Vbuses[~NoDERsPF_Vbuses.index.str.contains('lv')]; NoDERsPF_MVbuses_rest = NoDERsPF_MVbuses_rest[~NoDERsPF_MVbuses_rest.index.isin(loadsmv_buses)] # pu rest of buses dataframe at base power flow (no DER)
        NoDERsPF_LVbuses_rest = NoDERsPF_Vbuses[NoDERsPF_Vbuses.index.str.contains('lv')]; NoDERsPF_LVbuses_rest = NoDERsPF_LVbuses_rest[~NoDERsPF_LVbuses_rest.index.isin(loadslv_buses)]
        MV_buses_rest = V_buses[~V_buses.index.str.contains('lv')]; MV_buses_rest = MV_buses_rest[~MV_buses_rest.index.isin(loadsmv_buses)]  # pu rest of buses dataframe
        LV_buses_rest = V_buses[V_buses.index.str.contains('lv')]; LV_buses_rest = LV_buses_rest[~LV_buses_rest.index.isin(loadslv_buses)]  # pu rest of buses dataframe
        devLV_rest = pd.DataFrame( zip( abs( (LV_buses_rest.VOLTAGE-NoDERsPF_LVbuses_rest.VOLTAGE)/NoDERsPF_LVbuses_rest.VOLTAGE ), Base_V.LV_GROUP, Base_V.TX ), index=NoDERsPF_LVbuses_rest.index, columns=['VDEVIATION', 'LV_GROUP', 'TX'])
        devMV_rest = pd.DataFrame( zip( abs( (MV_buses_rest.VOLTAGE-NoDERsPF_MVbuses_rest.VOLTAGE)/NoDERsPF_MVbuses_rest.VOLTAGE ), Base_V.LV_GROUP, Base_V.TX ), index=NoDERsPF_MVbuses_rest.index, columns=['VDEVIATION', 'LV_GROUP', 'TX'])
        
        VDl_List = list(set(devLV_loads.loc[devLV_loads['VDEVIATION'].gt(0.05), 'TX'].dropna()))
        VDr_List = list(set(devLV_rest.loc[devLV_rest['VDEVIATION'].gt(0.05), 'TX'].dropna()))
        if not Voltagedeviation_loads_df.empty:
            VDl_List = list(set(list(Voltagedeviation_loads_df['BLACKLIST_TX'].values[-1]) + VDl_List))
            VDr_List = list(set(list(Voltagedeviation_rest_df['BLACKLIST_TX'].values[-1]) + VDr_List))
        
        devLV_all = pd.concat([devLV_loads, devLV_rest])
        devMV_all = pd.concat([devMV_loads, devMV_rest])
        if devLV_all[devLV_all['VDEVIATION'].gt(0.05)].index.empty and devMV_all[devMV_all['VDEVIATION'].gt(0.03)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if not devMV_all[devMV_all['VDEVIATION'].gt(0.03)].index.empty else False  
    
        max_lvdev_l = devLV_loads.VDEVIATION.max()
        buslvi_l=devLV_loads.VDEVIATION.idxmax()
        pdelvvi_l = devLV_loads.VDEVIATION.gt(0.05).mean()
        if not devLV_rest.empty:
            max_lvdev_r = devLV_rest.VDEVIATION.max()
            buslvi_r=devLV_rest.VDEVIATION.idxmax()
            pdelvvi_r = devLV_rest.VDEVIATION.gt(0.05).mean()
        else:
            max_lvdev_r = np.nan
            buslvi_r = np.nan
            pdelvvi_r = np.nan
        max_mvdev_r = devMV_rest.VDEVIATION.max()
        busmvi_r = devMV_rest.VDEVIATION.idxmax()
        pdemvvi_r = devMV_rest.VDEVIATION.gt(0.03).mean()
        
        if not devMV_loads.empty:
            max_mvdev_l = devMV_loads.VDEVIATION.max()
            busmvi_l=devMV_loads.VDEVIATION.idxmax()
            pdemvvi_l = devMV_loads.VDEVIATION.gt(0.03).mean()
            dict_ = {'GDKWP': [capacity_i], 'BUSLV': [buslvi_l],
                     'VDEVIATIONLV': [max_lvdev_l], 'PDEVIATIONLV': [pdelvvi_l],
                     'BUSMV': [busmvi_l], 'VDEVIATIONMV': [max_mvdev_l],
                     'PDEVIATIONMV': [pdemvvi_l], 'BLACKLIST_TX': [VDl_List],
                     'HCSTOP': [HCSTOP]}
            Voltagedeviation_loads_df = Voltagedeviation_loads_df.append(pd.DataFrame(dict_),
                                                                         ignore_index=True)
        else:
            dict_ = {'GDKWP': [capacity_i], 'BUSLV': [buslvi_l],
                     'VDEVIATIONLV': [max_lvdev_l],
                     'PDEVIATIONLV': [pdelvvi_l],
                     'BLACKLIST_TX': [VDl_List], 'HCSTOP': [HCSTOP]}
            Voltagedeviation_loads_df = Voltagedeviation_loads_df.append(pd.DataFrame(dict_),
                                                                         ignore_index=True)
    
        dict_ = {'GDKWP': [capacity_i], 'BUSLV': [buslvi_r],
                 'VDEVIATIONLV': [max_lvdev_r], 'PDEVIATIONLV': [pdelvvi_r],
                 'BUSMV': [busmvi_r], 'VDEVIATIONMV': [max_mvdev_r],
                 'PDEVIATIONMV': [pdemvvi_r], 'BLACKLIST_TX': [VDr_List],
                 'HCSTOP':[HCSTOP]}
        Voltagedeviation_rest_df = Voltagedeviation_rest_df.append(pd.DataFrame(dict_),
                                                                   ignore_index=True)

    #%% Voltage deviation at regulator devices (autotransformers and capacitors)
    if VoltageRegulation_analysis == True:
        try:
            col = 'BUSINSTALLED'
            for reg in list(RegDevices[col]):
                buslist = [reg.split('.')[0].lower() + '.' + ph for ph in reg.split('.')[1:]]
                bndwdth = RegDevices.loc[RegDevices[col] == reg, 'BANDWIDTH'].iloc[-1]
                vreg_ = RegDevices.loc[RegDevices[col] == reg, 'VREG'].iloc[-1]
                bndwdth = float(bndwdth)/float(vreg_)
                if max([abs(NoDERsPF_Vbuses['voltage'][busph] - V_buses['VOLTAGE'][busph]) for busph in buslist]) > bndwdth/2:
                    HCSTOP = True
                else: 
                    HCSTOP = False
            Voltagedeviation_reg_df = Voltagedeviation_reg_df.append(pd.DataFrame({ 'GDKWP':[capacity_i], 'BUSREG':[reg] ,'HCSTOP' : [HCSTOP] } ), ignore_index=True)
        except:
            pass
        t4=time.time()
        print('Voltage deviation processing: ' +str(t4-t3))
    
    #%% VOLTAGE UNBALANCE
    if VoltageUnbalance == True:
        # Greater of all three phase nodes
        buseslist = pd.DataFrame(columns=['BUS','LV_GROUP', 'TX', 'UNBALANCE'])
        buseslist['BUS'] = list(V_buses.index); buseslist['LV_GROUP'] = list(V_buses.LV_GROUP); buseslist['TX'] = list( V_buses['TX'])
        buseslist.index = buseslist['BUS']
        buseslist['UNBALANCE'] = buseslist['BUS'].apply(lambda b: voltageunbalance(DSScircuit, b))
        maxunb=buseslist.UNBALANCE.max(); busunb=buseslist.UNBALANCE.idxmax()
        buseslist = buseslist[buseslist.UNBALANCE.ge(0.03)]
    
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
    
        Voltageunbalance_df = Voltageunbalance_df.append(pd.DataFrame({ 'GDKWP':[capacity_i],'BUS':[busunb],'UNBALANCE':[maxunb],'BLACKLIST_TX': [VU_List], 'HCSTOP' : [HCSTOP] } ), ignore_index=True)

    return VBuses, Overvoltage_loads_df, Voltagedeviation_loads_df, Overvoltage_rest_df, Voltagedeviation_rest_df, Voltagedeviation_reg_df, Voltageunbalance_df
    # return V_buses, Overvoltage_loads_df, Overvoltage_rest_df

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
# Output: Thermal_loading_lines_df : pd.DataFrame({ 'GDKWP':[capacity_i],'Line':[txi],'LOADING' : [max_loading], 'POVERLOADTX' : [ptloadingi] } ), Thermal_loading_tx_df : pd.DataFrame({ 'GDKWP':[capacity_i],'Line':[txi],'LOADING' : [max_loading], 'POVERLOADTX' : [ptloadingi] } )

def thermal_Lines_Tx(DSScircuit,DSStext, normalAmpsDic,  capacity_i,
                     line_lv_groups, tx_groups, Thermal_loading_lines_df,
                     Thermal_loading_tx_df, name_file_created, linelvgroups,
                     Thermal_analysis): # Most code from auxfns.lineCurrents
    
    if Thermal_analysis == True: 
        # LINES
        start = time.time()
        
        CurrentDF = pd.DataFrame()
        CurrentDF['DSSName'] = DSScircuit.Lines.AllNames; CurrentDF=CurrentDF[~CurrentDF.DSSName.str.contains('swt')]; CurrentDF=CurrentDF[~CurrentDF.DSSName.str.contains('mv3p'+name_file_created.split('_')[0].lower()+'00')]
        CurrentDF['DSSName'].astype(str)
        CurrentDF['CURRENT'] = CurrentDF.DSSName.apply(lambda c : line_current(DSScircuit, c, normalAmpsDic))
        CurrentDF['LENGTH'] = CurrentDF.DSSName.apply(lambda c : line_length(DSScircuit, c))
    
        if len(linelvgroups) == 0: 
            CurrentDF['LV_GROUP'] = CurrentDF.DSSName.apply(lambda g : linelvgroup(g, line_lv_groups) )
            linelvgroups=list(CurrentDF['LV_GROUP'])
        CurrentDF['LV_GROUP'] = linelvgroups
        CurrentDF['TX']= CurrentDF.LV_GROUP.apply(lambda g : txname(g,tx_groups) )
        CurrentDF.index = CurrentDF['DSSName']#; del CurrentDF['DSSName']
        
        max_loading = CurrentDF.CURRENT.max(); linei=CurrentDF.CURRENT.idxmax(); lenloadingi = CurrentDF.loc[CurrentDF['CURRENT'] > 1.00, 'LENGTH'].sum()
        
        LO_List = list(set(CurrentDF.loc[CurrentDF['CURRENT'].gt(1.00), 'TX'].dropna()))
        if not Thermal_loading_lines_df.empty:
            LO_List = list(set(list(Thermal_loading_lines_df['BLACKLIST_TX'].values[-1]) + LO_List))
    
        if CurrentDF[CurrentDF['CURRENT'].gt(1.00)].index.empty:
            HCSTOP = False
        else: 
            HCSTOP = True if any(CurrentDF[CurrentDF['CURRENT'].gt(1.00)].index.str.startswith('mv')) else False
    
        Thermal_loading_lines_df = Thermal_loading_lines_df.append(pd.DataFrame({'GDKWP':[capacity_i],'LINE':[linei],'CURRENT' : [max_loading], 'LENGTHOVERLOADED' : [lenloadingi],'BLACKLIST_TX' : [LO_List], 'HCSTOP':[HCSTOP] } ), ignore_index=True)
    
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
            if 'U' in trafo.replace(name_file_created.split('_')[0].upper(), '') and 'auto' not in str(trafo.lower()):
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
    
        Thermal_loading_tx_df = Thermal_loading_tx_df.append(pd.DataFrame({ 'GDKWP':[capacity_i],'TX':[txi],'LOADING' : [max_loading], 'POVERLOADTX' : [ptloadingi], 'BLACKLIST_TX': [TO_List],'HCSTOP': [HCSTOP]  } ), ignore_index=True)
        
        end = time.time()
        sim_time = end - start
        print('Thermal analysis: '+str(round(sim_time,2))+' sec.')
    
    else:
        CurrentDF = pd.DataFrame()
    
    return Thermal_loading_lines_df, Thermal_loading_tx_df, CurrentDF

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
    hola = {}
    for idx in lines_mv.index:
        bus1 = lines_mv.loc[idx,'bus1']
        bus2 = lines_mv.loc[idx,'bus2']
        nomvolt = lines_mv.loc[idx, 'NOMVOLT']
        dist = float(lines_mv.loc[idx,'LENGTH'])
        name = lines_mv.loc[idx, 'DSSName']
        G.add_edge(bus1,bus2,distance=dist, line_name=name)
        nodes_mv.loc[bus1, 'NOMVOLT'] = nomvolt
        nodes_mv.loc[bus2, 'NOMVOLT'] = nomvolt
        hola[idx] = [bus1, bus2, name]
    
    end = time.time()
    sim_time = end - start
    print('Graph time: '+str(round(sim_time,2))+' sec.')
    
    return G

#%% FINDS THE ELEMENT OF THE PROTECTION DEVICE - dataframe.apply function
# Input: pdevice, ini_bus: substation bus, G
# Output: element name

def find_element(pdevice, ini_bus, firstLine, G):
    # if pdevice in ['AFTERMETER', ini_bus]:
    #     elem = 'line.'+firstLine
    us_bus = [n for n in G.predecessors(pdevice)][0] #Upstream bus
    
    if us_bus == ini_bus:
        ckt_name = ini_bus.split(ini_bus[min([i for i, c in enumerate(ini_bus) if c.isdigit()])])[0].split('BUSMV')[1]
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
    if pdevice in fusibles['bus1']:
        fuse_idx = fusibles[fusibles['bus1']==pdevice].index.values[0] #index to loc
            
        if (fusibles.loc[fuse_idx, 'SAVE'] == 'SI') or (fusibles.loc[fuse_idx, 'SAVE'] == 'YES'):
            recloser_ID = fusibles.loc[fuse_idx, 'COORDINATE']
        else:
            recloser_ID = np.nan
        recloser_idx = reclosers.loc[reclosers['PDID'] == recloser_ID].index.values
        recloser_bus1 = reclosers.loc[recloser_idx, 'bus1']
        recloser_bus2 = reclosers.loc[recloser_idx, 'bus2']
        recloser_elem = 'line.'+G.edges[recloser_bus1, recloser_bus2]['line_name']
    else:
        recloser_elem = np.nan
    
    return recloser_elem

#%% PROTECTION DEVICES DATAFRAME CREATION
# Input: pdevice, ini_bus, G
# Output: CircuitBreakDvRoR, CircuitBreakDvFF_BFC

def pDevices(G, firstLine, fusibles, reclosers): 
    bus_list_FF_BFC = list(set([(fusibles.loc[x,'bus1'], fusibles.loc[x,'bus2']) for x in fusibles.index.values] + [(reclosers.loc[y,'bus1'], reclosers.loc[y,'bus2']) for y in reclosers.index.values]))
    bus_list_DvRoR = [(reclosers.loc[y,'bus1'], reclosers.loc[y,'bus2']) for y in reclosers.index.values]

    #bus de inicio: 
    rand_bus = bus_list_FF_BFC[0][0]
    ini_bus = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0]+'1' #BUSMV+nombrecircuito+1
    
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
        downstream_buses_pdevice[pdevice]['BusInstalled'] = pdevice+find_conn(pdevice, ini_bus, G)
            
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
    print('Primera parte: '+str(round(sim_time,5))+' sec.')
        
    #%% FF_BFC dataframe ###############################################  
      
    start2 = time.time()
    p_FF_list = [bus_list_FF_BFC[x][1] for x in range(len(bus_list_FF_BFC))]
    CircuitBreakDvFF_BFC = pd.DataFrame(np.nan, columns=['Element', 'BusInstalled', 'RecloserElement'], index=range(len(p_FF_list))) #Si de fusibles y reclosers
    CircuitBreakDvFF_BFC['BusInstalled'] = p_FF_list
    
    # try:
    #     idx_bus1 = CircuitBreakDvFF_BFC.loc[CircuitBreakDvFF_BFC['BusInstalled'] == ini_bus].index.values[0]
        
    #     ############## MAKE THE CHANGE
    #     CircuitBreakDvFF_BFC.loc[idx_bus1,'BusInstalled'] = 'AFTERMETER' #################
    #     # print('hola', CircuitBreakDvFF_BFC.loc[idx_bus1,'BusInstalled'])
    # except:
    #     pass
    
    CircuitBreakDvFF_BFC['Element'] = CircuitBreakDvFF_BFC['BusInstalled'].apply(lambda b : find_element(b, ini_bus, firstLine, G))
    if not fusibles.empty:
        CircuitBreakDvFF_BFC['RecloserElement'] = CircuitBreakDvFF_BFC['BusInstalled'].apply(lambda b : find_recloserelem(G, b, fusibles, reclosers))
    
    CircuitBreakDvFF_BFC['BusInstalled'] = CircuitBreakDvFF_BFC['BusInstalled'].apply(lambda b : b+find_conn(b, ini_bus, G))
    
    end = time.time()
    sim_time = end - start2
    print('Segunda parte: '+str(round(sim_time,5))+' sec.')
        
    return CircuitBreakDvRoR, CircuitBreakDvFF_BFC
#%% PROTECTION DEVICES VERIFICATION - Forward fault current increase and Breaker Fuse descoordination
# Iverter contribution to fault current: 1.2 (120% according to SANDIA Protections report)
# Records the fault current increase for each protection device (Checks if fault current increases by 10% respecting the no DER scenario)
# Checks breakers (excluding substation one), reclosers and fuses. All located in MV.
# Checks if there's a scheme of breaker fuse coordination, finds the difference between change in breaker current and fuse current
# Input: DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, kW_sim, kVAr_sim,  capacity_i, Trafos, DERs, No_DER_FFCurrents, FFCurrents, No_DER_BFCCurrents, BFCCurrents, faulttypes
# Output: FFCurrents, BFCCurrents

def FF_BFC_Current(DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, kW_sim, kVAr_sim,  capacity_i, Trafos, DERs, No_DER_FFCurrents, FFCurrents, No_DER_BFCCurrents, BFCCurrents, faulttypes, FF_analysis, BFC_analysis):
    
    if (FF_analysis or BFC_analysis) == True:
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
        hora_sec = snapshottime.split(':')
        
        #%% Fault type studies
        # defines the terminals necessary for each type of fault. Dictionary containing: Fault type: [terminal conection for bus 1, terminal conection for bus2, number of phases]
        faulttypeterminals = {'ABC':['.1.2','.3.3', '2'], 'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'], 'BC':['.2','.3','1'], 'AC':['.1','.3','1'], 'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'], 'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'], 'BG':['.2','.0','1'], 'CG':['.3','.0','1']}
        HCSTOP = False
        if FFCurrents.empty:
            columnslist=list([ 'GDKWP', 'Element', 'FaultedBus']+faulttypes+['HCSTOP'])
            FFCurrents = pd.DataFrame(columns=columnslist)
    
        if BFCCurrents.empty:
                columnslist=list([ 'GDKWP', 'FuseElement', 'BreakerElement','FaultedBus']+faulttypes+['HCSTOP'])
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
    
            temp_row_FF = [];  temp_row_FF.append(capacity_i); temp_row_FF.append(element); temp_row_FF.append(faultedbus)
    
            if element in No_DER_BFCCurrents.index.get_level_values('FuseElement'):
                breakerelement = No_DER_BFCCurrents.xs(element).index.get_level_values('BreakerElement')[-1]
                temp_row_BFC = [];  temp_row_BFC.append(capacity_i); temp_row_BFC.append(element); temp_row_BFC.append(breakerelement); temp_row_BFC.append(faultedbus)
            else:
                breakerelement = np.nan
            for faulttype in faulttypes:
                if faulttype not in faultlabels:
                    temp_row_FF.append(np.nan)
                else:
                    # Machines inicialization simulation
                    DSStext.Command = 'clear'
                    DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
                    DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
                    DSStext.Command = 'Set mode=daily'  # Type of Simulation
                    DSStext.Command = 'Set number=1'  # Number of steps to be simulated
                    DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
                    DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
                    DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
                    DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
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
                    HCSTOP = False if maxIchange < 1.1 else True
                    #BFC
                    if not np.isnan(breakerelement):
                        fuseD = np.array([temp_currents[i] for i in currentspositions[faulttype]]) - np.array(No_DER_BFCCurrents[faulttype][element,breakerelement,faultedbus][0])
                        DSScircuit.setActiveElement(breakerelement)
                        breaker_currents = DSScircuit.ActiveElement.CurrentsMagAng
                        breakerD = np.array([breaker_currents[i] for i in currentspositions[faulttype]]) - np.array(No_DER_BFCCurrents[faulttype][element,breakerelement,faultedbus][1])
                        maxIdelta=round(max(fuseD-breakerD),2)
                        temp_row_BFC.append(maxIdelta)
                        HCSTOP = False if maxIdelta < 100 else True
            if not np.isnan(breakerelement):
                BFCCurrents = BFCCurrents.append(pd.DataFrame([temp_row_BFC+[HCSTOP]], index=[0], columns=list(BFCCurrents.columns)), ignore_index=True)
            FFCurrents = FFCurrents.append(pd.DataFrame([temp_row_FF+[HCSTOP]], index=[0], columns=list(FFCurrents.columns)), ignore_index=True)
    
    return FFCurrents, BFCCurrents

#%% PROTECTION DEVICES VERIFICATION - Reduction of reach verification
# Iverter contribution to fault current: 1.2 (120% according to SANDIA Protections report)
# Records the current decrease for each breaker or recloser element to a fault far in its protection zone and farthest in the circuit. If the decrease is greater that 10%, there's a ROR problem.
# Input: DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, kW_sim, kVAr_sim,  capacity_i, Trafos, DERs, No_DER_RoRCurrents, RoRCurrents, faulttypes
# Output: RORCurrents
def ReductionReach(DSScircuit, DSStext, dss_network,firstLine, snapshottime, snapshotdate, kW_sim, kVAr_sim,  capacity_i, Trafos, DERs, No_DER_RoRCurrents, RoRCurrents, faulttypes, RoR_analysis):
    
    if RoR_analysis == True:
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
        hora_sec = snapshottime.split(':')
        
        #%% Fault type studies
        # defines the terminals necessary for each type of fault. Dictionary containing: Fault type: [terminal conection for bus 1, terminal conection for bus2, number of phases]
        faulttypeterminals = {'ABC':['.1.2','.3.3', '2'], 'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'], 'BC':['.2','.3','1'], 'AC':['.1','.3','1'], 'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'], 'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'], 'BG':['.2','.0','1'], 'CG':['.3','.0','1']}
    
        if RoRCurrents.empty:
            columnslist=list([ 'GDKWP', 'Element', 'FaultedBus']+faulttypes+['HCSTOP'])
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
    
            temp_row = [];  temp_row.append(capacity_i); temp_row.append(element); temp_row.append(faultedbus)
    
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
                    DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
                    DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
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
                    HCSTOP = False if minIchange > 0.90 else True
    
            RoRCurrents = RoRCurrents.append(pd.DataFrame([temp_row+[HCSTOP]], index=[0], columns=list(RoRCurrents.columns)), ignore_index=True)
    return RoRCurrents

#%% PROTECTION DEVICES VERIFICATION - Sympathetic tripping
# Iverter contribution to fault current: 1.2 (120% according to SANDIA Protections report)
# For faults upstream the circuit breaker (simulating faults in different circuits of the same HV bus in the power substation), checks the breaker's zero sequence current, has to bellow a certain value, e.g., 150 A.
# ActiveElement.seqCurrents yields for the first line of the circuit (IO_t1, I1_t1, I2_t1, IO_t2, I1_t2, I2_t2)
# Input: DSStext, DSScircuit, dss_network, firstLine, hora_sec, kW_sim, kVAr_sim, capacity_i, SympatheticTripping_analysis, faulttypes, Izero_trip, SympatheticTripping_df
# Output: SympatheticTripping_df
def SympatheticTripping(DSStext, DSScircuit, dss_network, firstLine, snapshotdate, snapshottime, kW_sim, kVAr_sim, capacity_i, Trafos, DERs, SympatheticTripping_analysis, faulttypes, Izero_trip, SympatheticTripping_df):
    faulttypeterminals = {'ABC':['.1.2','.3.3', '2'], 'ABCG':['.1.2.3','.0.0.0','3'], 'AB':['.1','.2','1'], 'BC':['.2','.3','1'], 'AC':['.1','.3','1'], 'ABG':['.1.1','.2.0','2'], 'BCG':['.2.2','.3.0','2'], 'ACG':['.1.1','.3.0','2'], 'AG':['.1','.0','1'], 'BG':['.2','.0','1'], 'CG':['.3','.0','1']}

    if SympatheticTripping_analysis:
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
        hora_sec = snapshottime.split(':')
        
        element = firstLine
        ftlist = 'sourcebus.1.2.3'.split('.')
        Izero_max=0
        for faulttype in faulttypes:
            print('FF: '+element+ ', Fault: '+faulttype)
            # Machines inicialization simulation
            DSStext.Command = 'clear'
            DSStext.Command = 'New Circuit.Circuito_Distribucion_Snapshot'
            DSStext.Command = 'Compile ' + dss_network + '/Master.dss'  # Compile the OpenDSS Master file
            DSStext.Command = 'Set mode=daily'  # Type of Simulation
            DSStext.Command = 'Set number=1'  # Number of steps to be simulated
            DSStext.Command = 'Set stepsize=15m'  # Stepsize of the simulation (se usa 1m = 60s)
            DSStext.Command = 'Set time=(' + hora_sec[0] + ',' + hora_sec[1] + ')'  # Set the start simulation time
            DSStext.Command = 'batchedit load.n_.* kW=' + str(kW_sim[0]) # kW corrector
            DSStext.Command = 'batchedit load.n_.* kVAr=' + str(kVAr_sim[0]) # kVAr corrector
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

            DSStext.Command = 'Solve mode=dynamics stepsize=0.00002'
            com = 'new fault.fault_' + faulttype + ' bus1=' + ftlist[0]
            com += term1 + ' bus2=' + ftlist[0] + term2 + ' phases=' + phasessc
            DSStext.Command = com
            print('faulttype: '+faulttype)
            print('bus1: '+ftlist[0]+term1)
            print('bus2: '+ftlist[0]+term2)
            DSScircuit.Solution.Solve()
            
            DSScircuit.SetActiveElement('line.'+firstLine)
            print('seq: '+str(DSScircuit.ActiveCktElement.SeqCurrents))
            print('currents: '+str(DSScircuit.ActiveCktElement.CurrentsMagAng))
            print('Powers: '+str(DSScircuit.ActiveCktElement.Powers))
            if Izero_max < DSScircuit.ActiveCktElement.SeqCurrents[0]:
                Izero_max = DSScircuit.ActiveCktElement.SeqCurrents[0]
                faulttype_max = faulttype
        HCSTOP = True if Izero_max >= Izero_trip else False
        dict_ = {'GDKWP': [capacity_i], 'FAULT': [faulttype_max],
                 'I0': [Izero_max], 'HCSTOP': [HCSTOP]}
        SympatheticTripping_df = SympatheticTripping_df.append(pd.DataFrame(dict_),
                                                               ignore_index=True)

    return SympatheticTripping_df
#%% INITIAL TX AND MV LOADS INFORMATION FOR FUTURE HISTOGRAM 
#Input: DSScircuit, DSSobj, DSStext, G, tx_layer, mv_loads_layer
#Output: LoadTrafos_MVLoads -> Dataframe 

def base_info_tx_and_mvloads(DSScircuit, DSSobj, DSStext, G, tx_layer, mv_loads_layer):
    start = time.time()
    
    rand_bus = tx_layer.loc[0,'bus1']
    ckt_name = rand_bus.split(rand_bus[min([i for i, c in enumerate(rand_bus) if c.isdigit()])])[0].split('BUSMV')[1]
    first_bus = 'BUSMV'+ckt_name+'1'
    vect = ['DSSName', 'LV_GROUP', 'bus1', 'KVAPHASEA',
            'KVAPHASEB', 'KVAPHASEC', 'kVA_snap']
    LoadTrafos_MVLoads_1 = tx_layer[vect]
    LoadTrafos_MVLoads_1 = LoadTrafos_MVLoads_1.assign(Rating = lambda x : x.KVAPHASEA.astype(float) + x.KVAPHASEB.astype(float) + x.KVAPHASEC.astype(float))
    cols = ['KVAPHASEA', 'KVAPHASEB', 'KVAPHASEC']
    LoadTrafos_MVLoads_1 = LoadTrafos_MVLoads_1.drop(columns=cols)  
    LoadTrafos_MVLoads_1.index = LoadTrafos_MVLoads_1['DSSName']
    
    if mv_loads_layer.empty == False:
        LoadTrafos_MVLoads_2 = mv_loads_layer[['DSSName', 'bus1', 'kVA_snap']]
        LoadTrafos_MVLoads_2.index = LoadTrafos_MVLoads_2['DSSName']
        LoadTrafos_MVLoads = pd.concat([LoadTrafos_MVLoads_1, LoadTrafos_MVLoads_2], sort=True)
    else: 
        LoadTrafos_MVLoads = LoadTrafos_MVLoads_1.copy()
        LoadTrafos_MVLoads['distance_m'] = LoadTrafos_MVLoads['bus1'].apply(lambda b: (nx.shortest_path_length(G, first_bus, b, weight='distance'))/1000)
    end = time.time()
    sim_time = end - start
    print('LoadMV: '+str(round(sim_time,2))+' sec.')
    
    return LoadTrafos_MVLoads

#%% GET LOAD KW VALUES AFTER BASE CASE RUN

def get_kva_load(DSScircuit, elem_name):
    
    DSScircuit.setActiveElement('load.'+elem_name)

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
        DSScircuit.setActiveElement('transformer.' + elem_name)
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
        if group_loads.empty == True:
            #append trafo
            blacklist += list(LoadTrafos_MVLoads.loc[LoadTrafos_MVLoads['LV_GROUP']==group].index.values) # appended to blacklist and will not be part of dict
        
        else:
            Load_lvloads[int(group)] = group_loads[['bus1','DSSName', 'NOMVOLT', 'SERVICE', 'kVA_snap']]
            
            val_sum = Load_lvloads[int(group)]['kVA_snap'].sum()
            Load_lvloads[int(group)] = Load_lvloads[int(group)].assign(kVA_norm = lambda x : x.kVA_snap/val_sum)
            
            # Load_lvloads[int(group)].loc['kVA_norm'] = Load_lvloads[int(group)]['kVA_snap']/val_sum
            
            Load_lvloads[int(group)] = Load_lvloads[int(group)].assign(DER_previo = lambda x : np.nan)
            Load_lvloads[int(group)] = Load_lvloads[int(group)].assign(DER_actual = lambda x : np.nan)
            
    end = time.time()
    sim_time = end - start
    print('LV dataframe time: ' + str(round(sim_time, 2)) + ' sec.')
    return Load_lvloads, blacklist

#%% Step calculation
def step_calc(MV_hist_df, blacklist, blacklist_dict, max_kVA_step,
              max_kVA_mvloads, lim_kVA, der_level_actual):
    not_in_blacklist = list(set(MV_hist_df.index.values) - set(blacklist))
    mv_load_list = list(MV_hist_df.loc[MV_hist_df['Rating'].isnull()].index.values)
    # update kVA_snap_norm:
    val_sum = MV_hist_df.loc[not_in_blacklist, 'kVA_snap'].sum()
    
    #first estimation
    if len(list(set(not_in_blacklist)-set(mv_load_list))) >0:
        ref_val = max_kVA_step
        study_list = list(set(not_in_blacklist)-set(mv_load_list))
        der_step = int(max_kVA_step/(MV_hist_df.loc[list(set(not_in_blacklist)-set(mv_load_list)), 'kVA_snap'].max()/val_sum))
    
    else:
        ref_val = max_kVA_mvloads
        study_list = mv_load_list
        der_step = int(max_kVA_mvloads/(MV_hist_df.loc[mv_load_list, 'kVA_snap'].max()/val_sum))
    
    ##################################################################################
    der_level_actual += der_step #UPDATES THE VALUE
    
    if der_level_actual > der_step: #filter for the first iteration
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
    
    if lim_kVA == True:
        #filtered data
        lim_df = MV_hist_df.loc[MV_hist_df['DER_actual'] > MV_hist_df['Rating']]
        temp_blacklist = list(lim_df.index.values)
        #Updates values
        MV_hist_df.loc[temp_blacklist, 'DER_actual'] = MV_hist_df.loc[temp_blacklist, 'Rating'].copy()
        der_level_actual = int(MV_hist_df['DER_actual'].sum())
        der_step = int(MV_hist_df['DER_actual'].sum())- int(MV_hist_df['DER_previo'].sum())
        
    else:
        temp_blacklist = []
            
    return der_step, der_level_actual, MV_hist_df, temp_blacklist
        
#%% UPDATE HISTOGRAMS
# Input: der_level_actual, LoadTrafos_MVLoads, blacklist, LV_hist_df, MV_hist_df=0
# Output: MV_hist_df, LV_hist_df

def DER_calc(der_level_actual, der_step, LoadTrafos_MVLoads, blacklist, LV_hist_df, MV_hist_df=0):
    if der_level_actual == 0: #initial step
        # first, MV histogram 
        MV_hist_df = LoadTrafos_MVLoads[['DSSName', 'bus1', 'LV_GROUP', 'Rating' ,'kVA_snap']]
        # MV_hist_df = MV_hist_df.rename(columns={'base_snap': 'kVA_snap'})
        MV_hist_df = MV_hist_df.assign(base_kVA_val = lambda x: x.kVA_snap)
        # MV_hist_df['base_kVA_val'] = MV_hist_df.copy()['kVA_snap']
        
        MV_hist_df.loc[blacklist, 'kVA_snap'] = 0 # make the txs without loads get a value of 0 (excluded from analysis)
        val_sum = MV_hist_df['kVA_snap'].sum()
        # MV_hist_df['kVA_snap_norm'] = MV_hist_df['kVA_snap']/val_sum
        MV_hist_df = MV_hist_df.assign(kVA_snap_norm = lambda x : x.kVA_snap/val_sum)
        
        MV_hist_df['DER_previo'] = 0
        MV_hist_df['DER_actual'] = 0
        # MV_hist_df['blacklist_info'] =''
        MV_hist_df.loc[blacklist, 'blacklist_info'] = 'no load tx'
        
        # LV histogram 
        for group in LV_hist_df:
            LV_hist_df[int(group)]['DER_previo'].values[:] = 0
            LV_hist_df[int(group)]['DER_actual'].values[:] = 0
            
    else:
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

def DER_allocation_HHC(LV_hist_df, MV_hist_df, mv_loads_layer):
    
    DERs_LV = []
    DERs_MV = []
    Trafos_DERs_MV = []
    
    for elem in MV_hist_df.loc[MV_hist_df['DER_actual'] > 0].index.values:
        group_id = MV_hist_df.loc[elem, 'LV_GROUP']
        ########################################################
        if np.isnan(group_id) == False: #si es un trafo
            for idx_lv_load in LV_hist_df[int(group_id)].loc[LV_hist_df[int(group_id)]['DER_actual'] > 0].index: #recorre el df
                
                service = ''; n_phases = ''; bus_conn = ''; nomvolt = ''; busLV= ''; derName = ''; installed_capacity = ''
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
                derName = LV_hist_df[int(group_id)].loc[idx_lv_load, 'DSSName']
                installed_capacity = str(LV_hist_df[int(group_id)].loc[idx_lv_load, 'DER_actual'])
    
                der_line = 'new generator.DER_'+ derName + ' phases='+n_phases+' bus1='+ busLV + bus_conn + ' kV='+ kV_LowLL + ' kW=' + str(float(installed_capacity))
                der_line += ' PF=1' + ' conn=wye' + ' kVA=' +str(round(float(installed_capacity)*1.2,2)) +' Model=7' + ' Vmaxpu=1.5 Vminpu=0.80 Balanced=no Enabled=yes'                    
        ###############################################
                DERs_LV.append(der_line)
        
        else: #si es una carga de media
            
            #inicialización de variables
            service = ''; n_phases = ''; bus_conn = ''; nomvolt = ''; busLV= ''; busMV= ''; derName = '' 
            installed_capacity = ''; kVA = ''
            
            trafo_line = []; der_line = []
            
            #index en el layer
            idx_mv_load = mv_loads_layer.loc[mv_loads_layer['DSSName']==elem].index[0]
            
            #hacer la diferenciación de casos
            service = str(mv_loads_layer.loc[idx_mv_load, 'PHASEDESIG'])
            # if type(mv_loads_layer.loc[idx_mv_load, 'PHASEDESIG']) != str:
                
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
            busLV = 'BUSLV'+busMV.split('MV')[1][0:3]+'_DER_'+elem
            
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
                trafo_line = 'new transformer.' + trafName + ' phases='+n_phases+ ' windings=3 ' + noloadloss + " " + imag + ' buses=[' + busMV +bus_conn+' '
                trafo_line += busLV + '.1.0 '+ busLV + '.0.2 ]' + ' conns=[' + confMV + ' ' + confLV_1 + ' '+confLV_2+']' + ' kvs=[' + kV_MedLN + ' ' +  kV_LowLN + ' '+ kV_LowLN+']'
                trafo_line += ' kvas=[' + kVA + ' ' + kVA + ' '+ kVA +'] ' + impedance_X +' '+ impedance_R + ' Taps=[' + tap + ', 1, 1 ]' + normhkva
                
                # OpenDSS DER sentence
                der_line = 'new generator.DER_'+ derName + ' phases='+n_phases+' bus1='+ busLV + '.1.2' + ' kV='+ kV_LowLL + ' kW=' + str(float(installed_capacity))
                der_line += ' PF=1' + ' conn=wye' + ' kVA=' +str(round(float(installed_capacity)*1.2,2)) +' Model=7' + ' Vmaxpu=1.5 Vminpu=0.80 Balanced=no Enabled=yes'  


            elif n_phases == '2' or n_phases == '3':
                kV_MedLL = str(trafOps.renameVoltage(nomvolt, 50)['MVCode']['LL']) # media tensión trifásico
                kV_MedLN = str(trafOps.renameVoltage(nomvolt, 50)['MVCode']['LN']) # media tensión para monofásico
                kV_LowLL = str(trafOps.renameVoltage(nomvolt, 50)['LVCode']['LL']) # baja tensión trifásico (480 V)
                kV_LowLN = str(trafOps.renameVoltage(nomvolt, 50)['LVCode']['LN']) # baja tensión para monofásico
                # print(idx_mv_load, elem, kV_MedLL, kVA)
                impedance = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Z')
                noloadloss = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Pnoload')
                imag = trafOps.impedanceSingleUnit(n_phases, kV_MedLL, kV_LowLN, kVA).get('Im')
                trafName = elem.split('MV')[1][0:3] + n_phases + 'DERs_' + elem
                
                confMV = 'wye'
                confLV = 'wye'
                tap = '1'
                
                # OpenDSS tx sentence
                trafo_line = 'new transformer.' + trafName + ' phases=3 windings=2 ' + noloadloss + " " + imag + ' buses=[' + busMV + '.1.2.3 '
                trafo_line += busLV + '.1.2.3]' + ' conns=[' + confMV + ' ' + confLV + ']' + ' kvs=[' + kV_MedLL + " " +  kV_LowLL + ']'
                trafo_line += ' kvas=[' + kVA + " " + kVA + '] ' + impedance + ' Taps=[' + tap + ', 1]' + normhkva
                
                # OpenDSS DER sentence
                der_line = 'new generator.DER_'+ derName + ' phases='+n_phases+' bus1='+ busLV + bus_conn + ' kV='+ kV_LowLL + ' kW=' + str(float(installed_capacity))
                der_line += ' PF=1' + ' conn=wye' + ' kVA=' +str(round(float(installed_capacity)*1.2,2)) +' Model=7' + ' Vmaxpu=1.5 Vminpu=0.80 Balanced=no Enabled=yes'  
            

        ###############################################
            DERs_MV.append(der_line)
            Trafos_DERs_MV.append(trafo_line)
      
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
    
    criteria_dict = {'Overvoltage_loads':{'Criteria': Overvoltage_analysis, 'data': Overvoltage_loads_df},
                     'Overvoltage_rest':{'Criteria': Overvoltage_analysis, 'data': Overvoltage_rest_df},
                     'VoltageDeviation_loads':{'Criteria': VoltageDeviation_analysis, 'data': Voltagedeviation_loads_df},
                     'VoltageDeviation_rest':{'Criteria': VoltageDeviation_analysis, 'data': Voltagedeviation_rest_df},
                     'VoltageRegulators':{'Criteria': VoltageRegulation_analysis, 'data': Voltagedeviation_reg_df},
                     'VoltageUnbalance':{'Criteria': VoltageUnbalance, 'data': Voltageunbalance_df},
                     'ThermalLoadingTx': {'Criteria': Thermal_analysis, 'data': Thermal_loading_tx_df},
                     'ThermalLoadingLines': {'Criteria': Thermal_analysis, 'data': Thermal_loading_lines_df}, 
                     'ProtectionFF': {'Criteria': FF_analysis, 'data': FFCurrents},
                     'ProtectionBFC':{'Criteria': BFC_analysis, 'data': BFCCurrents},
                     'ProtectionRoR':{'Criteria': RoR_analysis, 'data': RoRCurrents},
                     'SympatheticTripping':{'Criteria': SympatheticTripping_analysis, 'data': SympatheticTripping_df}}
    
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
    
    hc_blacklist = pd.DataFrame(pd.concat(hc_blacklist_list, sort=True)['BLACKLIST_TX']).dropna()
    
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
        flag_list = ['Voltage or Thermal problems'] 
    
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

