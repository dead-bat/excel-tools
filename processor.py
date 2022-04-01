# Process D1 Historical Data

import os
import pandas as pd
from datetime import date

TODAY = str(date.today())
ROOT_DIR = os.getcwd()


files_list = []

for file in os.listdir(path=ROOT_DIR):
    if '.xls' in file or '.xlsx' in file:
        files_list.append(file)

for file in files_list:
    print("PROCESSING FILE \'%s\'" % file)
    out_file_root = file[:-5] if '.xlsx' in file else file[:-4]
    out_file = out_file_root + "_" + TODAY + ".csv"
    xl = pd.ExcelFile(file)
    df = None
    for sheet in xl.sheet_names:
        if df is None:
            print("Parsing sheet \'%s\' in file \'%s\'..." % (sheet, file))
            df = pd.read_excel(xl, sheet, header=0)
        else:
            # header=0, skiprows=None
            print("Adding data from sheet, \'%s\', to dataframe for file \'%s\'..." % (
                sheet, file))
            df_add = pd.read_excel(file, sheet, skiprows=0)
            df.append(df_add)
    print("Writing out aggregated data to file: \'%s\'..." % out_file)
    df.to_csv(out_file, index=None)
# EOF
