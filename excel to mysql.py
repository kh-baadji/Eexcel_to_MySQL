import pyodbc
import pandas as pd_lib
from tqdm import tqdm

# Connection to the databases server.
mySql_Connection_String = "DRIVER={MySQL ODBC 9.0 ANSI Driver}; SERVER=localhost;DATABASE=olympics; UID=root;"
connection = pyodbc.connect(mySql_Connection_String)
cursor = connection.cursor()

excel_file = 'olympics.xlsx'
excel_sheet = 'events'


for i in tqdm(range(28)):

    # Get data from file (Excel).
    excel_data = pd_lib.read_excel(excel_file, sheet_name=excel_sheet, nrows = 10000, skiprows=(10000 * i))

    for index, row in excel_data.iterrows():

        # Remplace empty cells by NULL values
        row = row.where(pd_lib.notnull(row), 'NULL')
        insert_request = f"INSERT INTO olympics VALUES(\"{row.iloc[0]}\", \"{row.iloc[1]}\", \"{row.iloc[2]}\", {row.iloc[3]}, {row.iloc[4]}, {row.iloc[5]}, \"{row.iloc[6]}\", \"{row.iloc[7]}\", \"{row.iloc[8]}\", {row.iloc[9]}, \"{row.iloc[10]}\", \"{row.iloc[11]}\", \"{row.iloc[12]}\", \"{row.iloc[13]}\", \"{row.iloc[14]}\")"
        cursor.execute(insert_request)

    connection.commit()

connection.close()
