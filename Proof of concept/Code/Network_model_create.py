# -*- coding: utf-8 -*-
"""
Created on Wed Jan 18 14:16:11 2023

@author: h_jet
"""

from tkinter import Tk
from tkinter.filedialog import askdirectory
import os
import pandas as pd
import glob
import networkx as nx

inputsheet = 'F_Inputs'
outputsheet = 'F_Outputs'

path = askdirectory(title='Select Folder') # shows dialog box and return the path
print(path)

os.chdir(path)

# create lists of all inputs and outputs
xlsxfiles = []
for file in glob.glob("*.xlsx"):
    xlsxfiles.append(file)

finputs = []
foutputs = []

for file in xlsxfiles:
    try:
        dfinput = pd.read_excel(file, sheet_name=inputsheet)
        dfinput = dfinput.rename(columns=dfinput.iloc[0]).drop(dfinput.index[0])
        dfinput = dfinput.iloc[1:]
        finput = dfinput.Reference.tolist()
        finputs.append(finput)
    except ValueError:
        finputs.append([])
 
    try:
        dfoutput = pd.read_excel(file, sheet_name=outputsheet)
        dfoutput = dfoutput.rename(columns=dfoutput.iloc[0]).drop(dfoutput.index[0])
        dfoutput = dfoutput.iloc[1:]
        foutput = dfoutput.Reference.tolist()
        foutputs.append(foutput)
    except ValueError:
        foutputs.append([])
        
# create digraph, nodes and edges
num_models = len(xlsxfiles)
G = nx.DiGraph()
for i in range(num_models): G.add_node(xlsxfiles[i], F_Inputs = finputs[i], F_Outputs = foutputs[i])

for i in range(num_models): 
    for j in range(num_models):
        data_trans = [k for k in finputs[i] if k in foutputs[j]]
        if data_trans:
            G.add_edge(xlsxfiles[j], xlsxfiles[i], Data_transfered = data_trans)
        
# draw network
nx.draw_networkx(G)

# check for cycles
try:
    cycles = nx.find_cycle(G)
    print("cycles found")
    print(*cycles, sep="/n")
except nx.exception.NetworkXNoCycle:
    print("no cycles found")
    
# work out batch order
'''
def find_order(G):
    nodes = list(G.nodes)
    for node in nodes:
'''      


