import pandas as pd

dd = pd.read_excel(r"\\ACCURATEMATE\projects\Test Slips\2020\Test.xlsx",None)
b=list(dd.keys())

for x in range(1,len(b)):
    dd=pd.read_excel(r"\\ACCURATEMATE\projects\Test Slips\2020\Test.xlsx")
    globals()[f"df{x}"] = pd.read_excel(r"\\ACCURATEMATE\projects\Test Slips\2020\Test.xlsx",x)




    print(x)
