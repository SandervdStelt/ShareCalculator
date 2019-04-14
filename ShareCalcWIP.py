
import pandas as pd
import glob
### Import fudara timeseries

# define some directories
production_database_path = r'C:\Projects\ShareCalculator\LocatieProductie.xlsx'
Test_database_path = r'C:\Projects\ShareCalculator\LocatieProductie_test.xlsx'
fudura_import_path = glob.glob(r'C:\Projects\ShareCalculator\Fudura Import\*.xlsx')
# Script to read excel containing timeseries pasting into excel with timeseries
excel_file = pd.read_excel (fudura_import_path[0],sheet_name = 1,header = 0,parse_cols = "A,C") 
excel_file.columns = ["Date","Total"]

# Modify data to store in database
empty = ' ' # define empty cell
first_empty_row = excel_file.index[excel_file.iloc[:,0] == empty][0] # Find first empty row in timeseries to determine enddate
totaal_verbruik = excel_file.iloc[0:first_empty_row,[0,1]] # Get total consumption
totaal_verbruik.iloc[:,1] *= -1 # multiply by -1
totaal_verbruik["Date"] = pd.to_datetime(totaal_verbruik["Date"],dayfirst = True, format = '%Y-%m-%d %H:%M:%S') # convert date column to datetimeformat
totaal_verbruik.index = totaal_verbruik["Date"] # Set datetime as index
del totaal_verbruik["Date"] # delete datecolum

# Get database (currently an excel file)
test_database = pd.read_excel(production_database_path,sheet_name = 0,header = 0,parse_cols = "A,B") 
test_database.columns = ["Date","Total"]
test_database["Date"] = pd.to_datetime(test_database["Date"],dayfirst = True, format = '%d-%m-%Y') # convert date column to datetimeformat
test_database.index = test_database["Date"] # Set datetime as index
del test_database["Date"] # delete datecolumn


# use concat to paste new dataset into old
#new_database = pd.concat([test_database, totaal_verbruik])
#new_database = test_database.merge(totaal_verbruik,how = 'left',right_index = True,left_index = True)
new_database = test_database.combine_first(totaal_verbruik)
#new_database = new_database[~new_database.index.duplicated(keep = 'last')]
print(new_database)
# Store Excel file

writer = pd.ExcelWriter(Test_database_path, engine = 'xlsxwriter')
new_database.to_excel(writer)
writer.save()



#



