import pandas as pd

from pandas import ExcelWriter

source = r"\\ACCURATEMATE\projects\Test Slips\Sep 19.xlsx"
destination = r"C:\Users\Office\Desktop\pp.xlsx"
dd = pd.read_excel(source, None)
b = list(dd.keys())

for x in range(0, len(b)):
    globals()[f"df{x}"] = pd.read_excel(source, x)

    print(x)

writer = pd.ExcelWriter(destination, engine='xlsxwriter')

# write each DataFrame to a specific sheet
for i in range(0, len(b)):
    globals()[f"df{i}"].to_excel(writer, sheet_name=b[i])

# close the Pandas Excel writer and output the Excel file
writer.save()