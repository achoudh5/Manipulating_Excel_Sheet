# To visualise the data from Hacktoberfest_Inputt, we plot the occcurences of 
# each relation

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from pprint import pprint
from collections import Counter



def visualiser(fileName):
    df = pd.read_excel(fileName).to_dict()
    # pprint(df)
    rows = list(df['Source Color'].keys())

    column_A = list(df['Source Color'].values())    #Source colors
    column_B = list(df['Dest color'].values())      #Destination colors

    # Zone relations
    relations = []

    # Relations from to column 1 to column 2 for the directed graph
    relG = {'from':[],'to':[]}

    # Colors available
    colors=['blue','brown','green','grey','red','orange','yellow']

    for i in rows:
        if (str(column_A[i])).lower() in colors:
            # relations.append((str(valuesSc[i])+' -> '+str(valuesDc[i])))
            
            relG['from'].append(str(column_A[i]))
            relG['to'].append(str(column_B[i]))
    # print(relations)

    # # Associating the relations with their occurences
    # relationFlow = Counter(relations)
    # # pprint(relationFlow)

    # # A scatter plot to show the occurences
    # x,y = zip(*relationFlow.items())
    # plt.scatter(x,y)
    # plt.plot(x,y)
    # plt.show()

    # Creating a directed graph with column1 and column2 as nodes
    G = nx.from_pandas_edgelist(relG,'from','to',create_using=nx.DiGraph())
    nodeColors=[]
    myPos = nx.spring_layout(G, seed=2)
    for node in G:
        nodeColors.append(node)
    nx.draw_networkx(G,pos=myPos, node_color= nodeColors,  with_labels=True)
    plt.show()

if __name__ == "__main__":
    visualiser('Hacktoberfest_Inputt.xlsx')