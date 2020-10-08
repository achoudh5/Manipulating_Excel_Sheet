# To visualise the data from Hacktoberfest_Inputt, we plot the occcurences of 
# each relation

import pandas as pd
import matplotlib.pyplot as plt
from pprint import pprint
from collections import Counter

df = pd.read_excel('Hacktoberfest_Inputt.xlsx').to_dict()

rows = list(df['Source Color'].keys())

valuesSc = list(df['Source Color'].values())    #Source colors
valuesDc = list(df['Dest color'].values())      #Destination colors

# Zone relations
relations = []

for i in rows:
    relations.append((str(valuesSc[i])+' -> '+str(valuesDc[i])))
# print(relations)

# Associating the relations with their occurences
relationFlow = Counter(relations)
# pprint(relationFlow)

# A scatter plot to show the occurences
x,y = zip(*relationFlow.items())
plt.scatter(x,y)
# plt.plot(x,y)
plt.show()
