import pickle
from tkinter.filedialog import askdirectory
import networkx as nx
import matplotlib.pyplot as plt
import os

import pandas as pd

path = askdirectory(title='Select Folder')  # This shows the dialog box and return the path.
print(path, "\n")

os.chdir(path)  # Make the parent folder the working directory

file_set = []  # Create a list to hold the relative file paths
for dirpath, dirnames, files in os.walk(path): # Walk through the directory stuff
    for file_name in files:  #  For each file it comes across
        rel_dir = os.path.relpath(dirpath, path)  # Get the relative path
        rel_file = os.path.join(rel_dir, file_name)  # Join the relative path with the file name
        file_set.append(rel_file) # Add the item to the list

df_files = pd.DataFrame(file_set, columns=["Files"])
df_files = df_files["Files"].str.split(pat="\\", expand=True, n=2)
df_files.columns = "Folder", "Filename"
df_files.to_excel("Outputs/File Paths.xlsx", index=False)




#  LOAD DATA
G = pickle.load(open('graph.pkl', 'rb'))

# PRINT OUTPUTS
#print("The model files used are: ", xlsxfiles, "\n")  # Print the files used
print("The nodes are: ", G.nodes, "\n")
print("The notes data is as follows: ", G.nodes.data(), "\n")
print("The edges are: ", G.edges, "\n")
print("The edges data is as follows: ", G.edges.data(), "\n")

#print("Node 0................", dict(G.nodes.data())keys())


#  EXPORT RESULTS
with open('Outputs/Nodes_data.txt', 'w') as f:
    f.write(str(G.nodes.data()))

with open('Outputs/Edges_data.txt', 'w') as f:
    f.write(str(G.edges.data()))


nx.to_pandas_edgelist(G).to_excel('Outputs/edgelist.xlsx', index=False)
nx.to_pandas_adjacency(G).to_excel('Outputs/adjacency_matrix.xlsx')


#  PRINT GRAPH
nx.draw_networkx(G)
#plt.show()
plt.savefig("Outputs/graph.png")


