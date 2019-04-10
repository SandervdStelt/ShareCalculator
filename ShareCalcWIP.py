
import pandas as pd

### Import fudara timeseries

# Script to read excel containing timeseries pasting into excel with timeseries
excel_file = pd.read_excel (r'D:\Decab\EbaseWorksheetTest\ShareCalCulator\Fudura Meetgegevens\Opwek_Maart.xlsx',sheet_name = 1,header = 0,parse_cols = "A,C") 
excel_file.columns = ["Date","Total"]

# define empty cell
empty = ' '
# Find first empty row in timeseries to determine enddate
first_empty_row = excel_file.index[excel_file.iloc[:,0] == empty][0] -1
# Get total consumption
totaal_verbruik = excel_file.iloc[0:first_empty_row][0,1]
# multiply by -1
totaal_verbruik.iloc[:,1] *= -1
# convert date column to datetimeformat
totaal_verbruik["Date"] = pd.to_datetime(totaal_verbruik["Date"])
# Set datetime as index
totaal_verbruik.index = totaal_verbruik["Date"]
# delete datecolumn
del totaal_verbruik["Date"]


###
test_database = pd.read_excel(r'C:\Projects\myfirstpie\TestDatabase.xlsx',sheet_name = 0,header = 0,parse_cols = "A,B") 
test_database.columns = ["Date","Total"]
test_database["Date"] = pd.to_datetime(test_database["Date"])
test_database.index = test_database["Date"]
del test_database["Date"]

# use concat to paste new dataset into old
new_database = pd.concat([test_database,totaal_verbruik]).sort_index()

# Store Excel file
writer = pd.ExcelWriter(r'C:\Projects\myfirstpie\TestDatabase.xlsx', engine = 'xlsxwriter')
new_database.to_excel(writer)
writer.save()



