import pandas as pd
import numpy as np
file_loc = "path.xlsx"
df = pd.read_excel(r"C:\Users\Office\Desktop\Sep 20.xlsx", index_col=None, na_values=['NA'], usecols="A,C:AA")
print(df)
