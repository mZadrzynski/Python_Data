import pandas as pd
import numpy as np

#df = (pd.read_excel("OA - Porownanie_20220420_1220.xlsx", "Export Worksheet", usecols="G, H"))
df = (pd.read_excel("numtest.xlsx", "Arkusz1"))

#wpisuje wszystkie wyrazy z komórki excela w tablice i porównuje je z drugą
for i in range (0, 9005):
    try:
        a = df.iloc[i, 1]
        b = df.iloc[i, 2]
        x = a.split()
        y = b.split()

        data_1 = {'TYPE': x}
        data_2 = {'TYPE': y} 

        if np.array_equal(data_1, data_2):
            df.iloc[i, 3] = 'well'
        else:
            df.iloc[i, 3] = 'bad'
    except:
        df.iloc[i, 3] = 'very bad'
df.to_excel("numtest2.xlsx", sheet_name="Sheet3")
