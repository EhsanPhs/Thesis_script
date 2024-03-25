# -*- coding: utf-8 -*-
"""
Created on Wed Oct 12 11:48:14 2022

@author: Matteo
"""
import sys
sys.path.append(r'D:\apps\DIgSILENT\PowerFactory 2023 SP2\Python\3.9')#change the path to the folder in which digsilent is stored
import powerfactory as pf
import pandas as pd
import os


# manually create project in digsilent and insert the name here
project_name='Aosta_1'
try:
    app
    app.PrintInfo('Adding Nodes')
except NameError:
    app = pf.GetApplication()
    app.ClearOutputWindow()
    app.PrintPlain('Single call to the function')
    # Identifying the network I'm working on
    ProjectActivation = app.ActivateProject(project_name)
    MyProject = app.GetActiveProject()
    nome_progetto = MyProject.loc_name
    nome_rete = nome_progetto.split()[0] #split divides the words in a string into a list. It was set to 1
    app.PrintPlain(nome_rete)
    app.PrintPlain(nome_rete + ' Network input applied')
else:
    pass

app.PrintPlain('Reading the file for Nodes...')
try:
    location_folder=r'D:\Politecnico\MSC Thesis\Files\input'
    Nodes = pd.read_excel(os.path.join(location_folder,'Nodes_super_fixed.xlsx'))
except FileNotFoundError:
    location_folder=r'C:\Users\Matteo\Politecnico di Milano\Marco Merlo - ZZZ_Lipari\geo-referenced points'
    Nodes = pd.read_pickle(os.path.join(location_folder,'powerfactory_nodes.pkl'))


# Seleziono l'oggetto rete
MyGrid = app.GetCalcRelevantObjects(nome_rete + '.ElmNet')[0]


# Aggiungo i nodi ad uno ad uno
for i,row in Nodes.iterrows():
    name =row['Codice Nodo']
    #name_new =row['Codice Nodo Tipo']
    # if len(name) == 13:
    #     node_type = int(name[5])
    # else:
    #     node_type = int(4)
    
    node_type=int(row['Tipo nodo'])
    
    #node_type = int(name[5]) # the fifth number of the cname describes the node type
    #node_type = int(row['Tipo nodo']) #only in Nodes_fixed.xlsx file this column exists I added it manually
    voltage=15
    latitudine = float(row['Latitudine (numerico)'])
    longitudine = float(row['Longitudine (numerico)'])
    if node_type == 2 : #secondary substation
        MySubstation = MyGrid.CreateObject('ElmTrfstat', name)
        MySubstation.SetAttribute("sType", "Transformer Station: LV Distribution")
        MySubstation.SetAttribute("chr_name", name)
        MySubstation.SetAttribute("sShort",'SS')
        app.PrintPlain(MySubstation)

        # Creo il nodo MT
        MyMV_busbar = MySubstation.CreateObject('ElmTerm', name)
        MyMV_busbar.SetAttribute('uknom', voltage)
        MySubstation.SetAttribute("GPSlat", latitudine)
        MySubstation.SetAttribute("GPSlon", longitudine)
    elif node_type == 'generator':
        pass
        # find out what happens with the coordinates, are they a property of the object generator or what
    elif node_type == 3: #switching substation
        MySubstation = MyGrid.CreateObject('ElmTrfstat', name)
        MySubstation.SetAttribute("sType", "Transformer Station: LV Distribution")
        MySubstation.SetAttribute("chr_name", name)
        MySubstation.SetAttribute("sShort",'SW')
        app.PrintPlain(MySubstation)

        # Creo il nodo MT
        MyMV_busbar = MySubstation.CreateObject('ElmTerm', name)
        MyMV_busbar.SetAttribute('uknom', voltage)
        MySubstation.SetAttribute("GPSlat", latitudine)
        MySubstation.SetAttribute("GPSlon", longitudine)
        MyMV_busbar.SetAttribute('uknom', voltage)
        
        pass
    elif node_type == 4: #junction
        # Creo il nodo MT
        MyMV_busbar = MyGrid.CreateObject('ElmTerm', name)
        MyMV_busbar.SetAttribute("iUsage", 1)  # 1 = Junction Node
        MyMV_busbar.SetAttribute('uknom', voltage)

        app.PrintPlain(MyMV_busbar)
        MyMV_busbar.SetAttribute("GPSlat", latitudine)
        MyMV_busbar.SetAttribute("GPSlon", longitudine)
    elif node_type ==1: #primary substation node
        MySubstation = MyGrid.CreateObject('ElmSubstat', name)
        MySubstation.SetAttribute("sType", "Primary Substation")
        MySubstation.SetAttribute("chr_name", name)
        app.PrintPlain(MySubstation)
        MySubstation.SetAttribute("sShort",'PS')
        MySubstation.SetAttribute("GPSlat", latitudine)
        MySubstation.SetAttribute("GPSlon", longitudine)

        # Creo il nodo MT - in the excel file only one row with name V-1-380296 exists, here in the code i manually
        # seperated and created 2 busbars in PS. red and green
        MyMV_busbar1 = MySubstation.CreateObject('ElmTerm', name+'-green')
        MyMV_busbar1.SetAttribute('uknom', voltage)
        MyMV_busbar2 = MySubstation.CreateObject('ElmTerm', name+'-red')
        MyMV_busbar2.SetAttribute('uknom', voltage)
        


    #if math.isnan(latitudine) or math.isnan(longitudine):
        #app.PrintWarn("Warning! " + name + ' does not have geographical coordinates')


