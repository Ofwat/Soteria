# -*- coding: utf-8 -*-

"""
Created on Wed Jan 18 14:16:11 2023

@author: h_jet
"""

#Next steps (not ordered)
# Add in scraping for QA1,2 and 3. - PARTIALLY DONE
# Functionalise the code for creating the newtork. Make this for singluar repetition
    # in case we need to add nodes and edges individually in the future
# Functionalise the code for creating the xlsxfiles list
    # also record the state of F_INputs, F_Outputs and CLEAR_SHEETs whilst testing the files
# convert all the lists into a dictionary of dictionaries instead - much more robust
# create run order creator - needs to decide what run order actually means
    # most accurate case would be to have run order per model (ie based on dependency)
    # can calculate concurrent run order which looks better but not as efficient
    # can also calculate linear run order but that is so innefficient it likely won't be used
# create fucntion to check order
    # pull use input timestamps, output timestamps and QA1 from Fountain (not from output sheet)
    # create function to check that models have been run in order individually (input -> calc -> output)
    # create function to check that all models have been run (inputs) after all predecessors have been run: 1) correctly (passes previous check) 2) afer predecessors have outputed
# Check models have been calcualted
    # function to pull output report from Fountain, the F_Input timestamp from graph and QA3 code from graph
    # scrape QA3 value from report and compare against F_Input timestamp: If Fountain QA 3 value = Finput timestamp then calc done else calc not done.
# Compare df F_Inputs and df F_Outputs to report on Fountain - MARIA working on this
    # pull dfs from graph and compare against reports on Fountain
    # output differences and flag if passed
# Check models are connected to correct run
    # Compare input runid and output runid against defined correct run
    # output flag if correct or incorrect
# Functionalise circularity checks - not much else to do on this as its a fairly standard network test
# Create function to check lists of unused items and output. ALso add to graph and return graph
# 
# visualise. Possibly use pyvis. 
    #   Come up with a way to procedurally generate node positions
    #   Visualise error reports (e.g red nodes have errors, Green ok, Amber not tested)
    #   POssibly export graph to pbi compatible format to visualise individual model results there

from tkinter import Tk
from tkinter.filedialog import askdirectory
import os
import pandas as pd
import glob
import networkx as nx
#import xlrd
#from openpyxl import load_workbook
#from pyxlsb import open_workbook
import time
import matplotlib.pyplot as plt
import pickle
import fnmatch 

pd.set_option('display.max_columns', 40)
pd.set_option('display.width', 2000)

start = time.time()

# temporary list of common names. In practice will be stored centrally somewhere
inputsheet = 'F_Inputs'
outputsheet = 'F_Outputs'
clearsheet = 'CLEAR_SHEET'

clearsheet_items = {'finputid':'F_Inputs_Report_ID', 'foutputid': 'F_Outputs_Report_ID',
                    'finputrun':'inputRunId', 'foutputrun': 'outputRunId',
                    'foutputs_timestamp': 'outputSheetLastSent', 'companyid': 'companyId',
                    'tagid': 'tagId'}

QA_codes = {'QA1':{'QA_pattern' : '*_OUT1', 'QA_code':''},
            'QA2': {'QA_pattern' : '*_OUT2', 'QA_code':''}, 
            'QA3': {'QA_pattern' : '*_OUT3', 'QA_code':''},
            'QA4': {'QA_pattern' : '*_OUT4', 'QA_code':''}}


#  Select the parent folder for the model files.
path = askdirectory(title='Select Folder')  # This shows the dialog box and return the path.
print(path, "\n")

#os.chdir(path)  # Make the parent folder the working directory

# Create list of all xls* files in root folder and subfolders and adds to models if an "F_Inputs" or "F_Outputs" sheet is present
#models = {}



def scrape_folder(path): 
    os.chdir(path)
    models = {} # Create list of all xls* files in root folder and subfolders and adds to models if an "F_Inputs" or "F_Outputs" sheet is present
    for file in glob.glob("**/*.xls*", recursive=True): # If recursive is true, the pattern “**” will match any files and zero or more directories, subdirectories and symbolic links to directories.
        filename, file_extension = os.path.splitext(file) 
        # check if files contain an F_Inputs or F_Outputs sheet    
    
        try:
            if inputsheet in pd.ExcelFile(file).sheet_names or outputsheet in pd.ExcelFile(file).sheet_names:
                models[file] = {}
        except:
            print("error with", file)
        
            # this section checks if the F_Input, F_Output and CLEAR_SHEET are present and save a flag in the xlsxfile dictionary
            if inputsheet in pd.ExcelFile(file).sheet_names:
                models[file]['finput_present'] = ['True']
            else: models[file]['finput_present'] = ['False']
            
            if outputsheet in pd.ExcelFile(file).sheet_names:
                models[file]['foutput_present'] = ['True']
            else: models[file]['foutput_present'] = ['False']
        
            if clearsheet in pd.ExcelFile(file).sheet_names:
                models[file]['clearsheet_present'] = ['True']
            else: models[file]['clearsheet_present'] = ['False']
        
    return models

# TOD Check if this code segment is needed. No need to define now if exceptions define later
# Add keys to dictionary to prepare for defining later. this insures all keys are present even if blank

#for model in list(models.keys()):
    
#    models[model]['finputs_codes'] = []  # empty list of F_Inputs boncodes
#    models[model]['foutsputs_codes'] = [] # empty list of F_Outputs boncodes
#    models[model]['finputs_dfs'] = pd.DataFrame() # empty dataframes of F_Inputs sheets
#    models[model]['foutputs_dfs'] = pd.DataFrame() # empty dataframes of F_Outputs sheets
#    models[model]['finputs_timestamps'] = '' # empty string of F_Inputs timestamps
#    models[model]['foutputs_timestamps'] = '' # empty string of F_Outputs timestamps
#    models[model]['finputs_runs'] = '' # empty string of F_Input runs
#    models[model]['foutputs_runs'] = '' # empty string of F_Output runs
#    models[model]['finputs_reportids'] = '' # empty string of F_inputs run IDs
#    models[model]['foutputs_reportids'] = '' # empty string of F_Outputs run IDs
#    models[model]['companyids'] = '' # empty string of F_Inputs company IDs


# The try except code needs to be replaced with more specific ways of capturing if sheet isn't present
def scrape_model(models, inputsheet, outputsheet, clearsheet, clearsheet_items, QA_codes):
    for model in list(models.keys()):
        try:
            dffinput = pd.read_excel(model, sheet_name=inputsheet)  #  Create a dataframe from the F_Input sheet
        except:
            dffinput = pd.DataFrame() # create a blank dataframe if something goes wrong in the reading
            
        try:
            finput = dffinput.rename(columns=dffinput.iloc[0]).drop(dffinput.index[0]) # drop the extra index column the API provides
            finput = finput.iloc[1:] 
            finput = finput.Reference.tolist()  # Get a list of the BON codes
            finput = [x for x in finput if str(x) !='nan']  # Drop NANs from the list of BON codes
        except:
            finput = []
            
        try:
            finput_timestamp = dffinput.columns[4] # finds F_Inputs timestamp as the column header of the dataframe
        except:
            finput_timestamp = '' # blank string if F_INputs sheet isn't present
         
        try:
            dfoutput = pd.read_excel(model, sheet_name=outputsheet)  # Create a dataframe from the F_Output sheet
        except:
            dfoutput = pd.DataFrame() # create a blank dataframe if something goes wrong in the reading
            
        try:
            foutput = dfoutput.rename(columns=dfoutput.iloc[0]).drop(dfoutput.index[0])
            foutput = foutput.iloc[1:]
            foutput = foutput.Reference.tolist()   # Get the list of BON codes
            foutput = [x for x in foutput if str(x) !='nan']  # Drop NANs from the list of BON codes
        except:
            foutput = []
            
        # read clearsheet items passed and create a dataframe
        clearsheet_metadata = pd.DataFrame.from_dict(clearsheet_items, orient = 'index', columns =['fields'])
        
        if clearsheet in pd.ExcelFile(model).sheet_names:
            dfclearsheet = pd.read_excel(model, sheet_name=clearsheet, index_col=0, header= None, names=['metadata'])  #  Create a dataframe from the clearsheet
                
            clearsheet_metadata = clearsheet_metadata.merge(dfclearsheet, left_on='fields', right_index=True, how = 'left') #m
            clearsheet_metadata.fillna('', inplace=True) # replaces nans from a failed merge with empty strings
        else: 
            print("clear sheet not found in", model)
            clearsheet_metadata['metadata'] = ''
            
        for QA in list(QA_codes.keys()): # Loops through QA codes in lst of F_Outputs codes
            pattern = QA_codes[QA]['QA_pattern'] # pulls out the pattern from the QA_code dict
            matching = fnmatch.filter(foutput,pattern) # creates a list of all QA_codes that match the pattern
            if len(matching) >0:  # if QA codes have been found, use the first one (QA codes in the list should be identical)
                QA_codes[QA]['QA_code'] = matching[0]
            else: QA_codes[QA]['QA_code'] = '' # if no QA codes have been found return blank string
                    
        models[model]['finputs_codes'] = finput  # Add the list of bon codes.
        models[model]['foutputs_codes'] = foutput  # Add the list of bon codes.
        models[model]['finputs_dfs'] = dffinput # Add the F_Inputs dataframe
        models[model]['foutputs_dfs'] =(dfoutput) # Add the F_Outputs dataframes
        models[model]['finputs_timestamps'] = finput_timestamp #Add straing timestamp to list of timestamps

        models[model]['finputs_runs'] = clearsheet_metadata.loc['finputrun']['metadata']  # Add input run IDs
        models[model]['foutputs_runs'] = clearsheet_metadata.loc['foutputrun']['metadata'] # Add output run IDs
        models[model]['finputs_reportids'] = clearsheet_metadata.loc['finputid']['metadata']  # Add F_Input report IDs
        models[model]['foutputs_reportids'] = clearsheet_metadata.loc['foutputid']['metadata']  # Add F_Outputs report IDs
        models[model]['companyids'] = clearsheet_metadata.loc['companyid']['metadata']  # Add company IDs
        models[model]['QA1_code'] = QA_codes['QA1']['QA_code'] # Add QA1 code
        models[model]['QA2_code'] = QA_codes['QA2']['QA_code'] # Add QA2 code
        models[model]['QA3_code'] = QA_codes['QA3']['QA_code'] # Add QA3 code
        models[model]['QA4_code'] = QA_codes['QA4']['QA_code'] # Add QA4 code       
    return models

models = scrape_folder(path)    

#for model in list(models.keys()):  # For each model file we are working with...
models = scrape_model(models, inputsheet, outputsheet, clearsheet, clearsheet_items, QA_codes)  # ...get the list of BON codes in the f_input and f_output sheets
#    models[model]['finputs_codes'] = finput  # Add the list of bon codes.
#    models[model]['foutputs_codes'] = foutput  # Add the list of bon codes.
#    models[model]['finputs_dfs'] = dffinput # Add the F_Inputs dataframe
#    models[model]['foutputs_dfs'] =(dfoutput) # Add the F_Outputs dataframes
#    models[model]['finputs_timestamps'] = finput_timestamp #Add straing timestamp to list of timestamps

#    models[model]['finputs_runs'] = clearsheet_metadata.loc['finputrun']['metadata']  # Add input run IDs
#    models[model]['foutputs_runs'] = clearsheet_metadata.loc['foutputrun']['metadata'] # Add output run IDs
#    models[model]['finputs_reportids'] = clearsheet_metadata.loc['finputid']['metadata']  # Add F_Input report IDs
#    models[model]['foutputs_reportids'] = clearsheet_metadata.loc['foutputid']['metadata']  # Add F_Outputs report IDs
#    models[model]['companyids'] = clearsheet_metadata.loc['companyid']['metadata']  # Add company IDs
#    models[model]['QA1_code'] = QA_codes['QA1']['QA_code'] # Add QA1 code
#    models[model]['QA2_code'] = QA_codes['QA2']['QA_code'] # Add QA2 code
#    models[model]['QA3_code'] = QA_codes['QA3']['QA_code'] # Add QA3 code
#    models[model]['QA4_code'] = QA_codes['QA4']['QA_code'] # Add QA4 code

    
# create digraph, nodes and edges
G = nx.DiGraph()  # Create an empty digraph with no nodes and no edges.

# Add models as nodes
attributes = [v2 for v1 in models.values() for v2 in v1]
for model in list(models.keys()):  # For each of the model files...
    G.add_node(model) # add the model as a node
    
# Loop through every node and attribute and assign accordingly
for attribute in attributes: # for each attribute taken from the models
    for model in list(G.nodes): # for each model in the graph
        G.nodes[model][attribute] = models[model][attribute]  # assign an attribute from the model dictionary
              
#        attribute = models[model][attribute],  # ...associate it with its list of F_Input BON codes and ...
#               F_Outputs_codes = models[model]['foutputs_codes'],   # ...associate it with its list of F_Input BON codes
#               F_Inputs = models[model]['finputs_dfs'],
#               F_Outputs = models[model]['foutputs_dfs'],
#               F_Inputs_timestamp = models[model]['foutputs_dfs'],
#               )

for i in list(models.keys()):  # For each of the model files...
    for j in list(models.keys()):  # For each of the model files...
        data_trans = [k for k in models[i]['finputs_codes'] if k in models[j]['foutputs_codes']]
        if data_trans:
            # Add edge between node j and i, associate with BON codes
            G.add_edge(j, i, Data_transfered=data_trans)

# had to replace some code since some variables aren't available after functionalising and conversion to dict use instead of lists for model variables
#all_F_inputs = sum(finputs_codes, []) #  This is a way to flatten the list of lists
#BONs_nowhere = {}  # This dictionary will hold the models and its f_Output bons that go nowhere
#for i in range(G.number_of_nodes):  # For each of the model files...
#    data_not_trans = list(set(foutputs_codes[i]) - set(all_F_inputs))  # The data not transferred are the f_output BONs in model i that are not in the flat list of all f_input BONS
#    BONs_nowhere.update({xlsxfiles[i]: data_not_trans})  # Update the dictionary


all_F_inputs = sum(list(nx.get_node_attributes(G, 'finputs_codes').values()),[]) # Extract F_Inputs from graph and flatten into single list
BONs_nowhere = {}  # This dictionary will hold the models and its f_Output bons that go nowhere
for model in list(G.nodes):  # For each of the model files...
    data_not_trans = list(set(nx.get_node_attributes(G, 'foutputs_codes')[model]) - set(all_F_inputs))  # The data not transferred are the f_output BONs in model i that are not in the flat list of all f_input BONS
    nx.set_node_attributes(G, {model:{'BONs_nowhere' : data_not_trans}})  # Update the Graph


for model in list(G.nodes):
    y = sum((k for i,j,k in list(G.out_edges(model, 'Data_transfered'))), [])

# Export the results (orient is required to overcome the need for each column (model) to have the same number of entries. It is transposed back later
pd.DataFrame.from_dict(BONs_nowhere, orient='index').T.to_excel("Outputs/BONS_nowhere.xlsx", index=False)


# Draw network
nx.draw_networkx(G)
plt.show()

# Check for cycles
try:
    cycles = nx.find_cycle(G)
    print("Cycles found")
    print(*cycles, sep="/n")
except nx.exception.NetworkXNoCycle:
    print("No cycles found")
    
# Work out batch order
'''
def find_order(G):
    nodes = list(G.nodes)
    for node in nodes:
'''

#  Export results
nx.write_gpickle(G, 'graph2.pkl')
pickle.dump(G, open('graph.pkl', 'wb'))
#nx.write_gml(G, "graph.graphml")
#nx.write_graphml(G, "graph.graphml")


end = time.time()
print("Time: ", round(end-start, 1))

