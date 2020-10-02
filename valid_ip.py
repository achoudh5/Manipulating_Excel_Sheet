import openpyxl
import ipaddress  # included in Python 3

wb = openpyxl.load_workbook('C:/Users/Lenovo/Downloads/Hacktoberfest.xlsx')
ws = wb['Sheet1']

# function to validate a IP address.
def is_valid(ip):
    try:
        ipaddress.IPv4Network(ip)
        return True
    except ValueError:
        return False

# for processing the string
def process(s):
    
    # removes everything except full stops and numbers.
    for idx, k in enumerate(s):

        if not (k == '.' or k.isnumeric()):

            s = s[:idx] + " " + s[idx + 1:]

    s = s.replace(" ","")
    
    return(s)

# Driving Code
if __name__ == "__main__":
    
    col=[0,2]   # INPUT. A list containing the indices of the
                # columns, which needs to be checked for the
                # IP adresses.

    lis=[]      # for storing the valid cell values.
                # It will be a list of dictionaries with keys as
                # the respective column headings.

    # 2 is so as to ignore the headings row.
    for i in range(2, ws.max_row+1):
        
        # found variable checks if any of the column consists of
        # non IP address format values.
        found=False
        
        for j in col:

            if ws[i][j].value:

                s = process(ws[i][j].value)

                if not is_valid(s):
                    found = True
                    break
            else:
                found = True
 
        if not found:
            dic = {}
            for j in col:
                dic[ws[1][j].value] = ws[i][j].value
            lis.append(dic)

    print(lis)
