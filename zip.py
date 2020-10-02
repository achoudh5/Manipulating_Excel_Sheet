#sample functions, correct errors
#read zone color from col A and put it in same cell of column B, similarly for B->F
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles import *

def main():
    def parse_user_input(ss): #parses the source and destination ip's from the user's spreadsheet
        load_spreadsheet = openpyxl.load_workbook(ss)
        sheet = load_spreadsheet['Sheet1']  # Get a sheet from the workbook
        for row in sheet.iter_cols(min_col=4, min_row=2, max_col=6,
                                   max_row=100):  # iterate through all rows in specific column
            # correct this loop, eliminate checking the column E
            for i in range(2,101):
                if row==('Sheet1.E'+str(i)):
                    pass
            for cell in row:
                try:
                    if (cell.value)==None:
                        pass
                    elif ',' in cell.value:
                        if cell.value.split(','):
                            for i in cell.value.split(','):
                                i = i.strip()
                                #color the zone as per following mapping:- column A color for Column D and Column B color for Column F
                                # input read_color_from_column by reading column A color for Column D and Column B color for Column F
                                color_ip(read_color_from_column,cell) #calling the function color_ip
                    elif cell.value==None:
                        pass

                    else:
                        if cell.value.split('\n'):
                            for i in cell.value.split('\n'):
                                i = i.strip()
                                #color the zone as per following mapping:- column A color for Column D and Column B color for Column F
                                #input read_color_from_column by reading column A color for Column D and Column B color for Column F
                                color_ip(read_color_from_column,cell) #calling the function color_ip
                except AttributeError:
                    # color the zone as per following mapping:- column A color for Column D and Column B color for Column F
                    # input read_color_from_column by reading column A color for Column D and Column B color for Column F
                    color_ip(read_color_from_column, cell)  # calling the function color_ip
                except ValueError:
                    break
        load_spreadsheet.save(ss)
    def color_ip(prevsymbol ,cell):  # colors the ip based on zone
        if prevsymbol != None:
            if (prevsymbol.lower()) == "green":
                cell.fill = PatternFill(fgColor="0080FF00", fill_type="solid")
            if (prevsymbol.lower()) == "blue":
                cell.fill = PatternFill(fgColor="000080FF", fill_type="solid")
            if (prevsymbol.lower()) == "yellow":
                cell.fill = PatternFill(fgColor="00FFFF33", fill_type="solid")
            if (prevsymbol.lower()) == "red":
                cell.fill = PatternFill(fgColor="00FF3333", fill_type="solid")
            if (prevsymbol.lower()) == "grey":
                cell.fill = PatternFill(fgColor="00A0A0A0", fill_type="solid")
            if (prevsymbol.lower()) == "brown":
                cell.fill = PatternFill(fgColor="00994C00", fill_type="solid")
            if (prevsymbol.lower()) == "orange":
                cell.fill = PatternFill(fgColor="00FF9933", fill_type="solid")

    ss_inp = input('Enter the path to excel sheet\n')
    parse_user_input(ss_inp)
if __name__== "__main__":
    main()