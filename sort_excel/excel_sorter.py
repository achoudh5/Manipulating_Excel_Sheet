"""
@author: Ishita Mehta
"""

import pandas as pd

def sort_file(file_path):
    dataframe = pd.read_excel(file_path)
    sorted_dataframe = dataframe.sort_values(['Source', 'Destination'], ascending=[1, 1])
    sorted_dataframe.to_excel('sorted_output_file.xlsx')


if __name__ == '__main__':
    sort_file('Hacktoberfest_Inputt.xlsx')    
