import pandas as pd

dataframe = pd.read_excel('Hacktoberfest_Inputt.xlsx')

sorted_dataframe = df.sort_values(['Source', 'Destination'], ascending=[1, 1])

sorted_dataframe.to_excel('sorted_output_file.xlsx')
