import pandas as pd
import numpy as np

#df = (pd.read_excel("OA - Porownanie_20220420_1220.xlsx", "Export Worksheet", usecols="G, H"))
df = (pd.read_excel("sql.xlsx", "Export Worksheet"))

#wpisuje wszystkie wyrazy z komórki excela w tablice i porównuje je z drugą


for i in range (0, 1032):
    try:
        a = df.iloc[i, 0]
        b = df.iloc[i, 1]
        z = ["NO.1", "NO.2", "NO.3", "NO.4", "NO.5", "NO.6", "NO.7","NO.8", "NO.9"]

        #data_1 = a.split()
        #data_2 = b.split()
        for j in range (0, 8):
            if a.find(z[j]) > 0:
                df.iloc[i, 2] = z[j]
                break
        for j in range (0, 8):    
            if b.find(z[j]) > 0:
                df.iloc[i, 3] = z[j] 
                break
        if np.array_equal(df.iloc[i, 2], df.iloc[i,3]):
            df.iloc[i, 4] = 'jest ok'
    except:
    #   df.iloc[i, 4] = 'zle'
        print(i) 
        print("ten nie pasuje")
df.to_excel("test2.xlsx", sheet_name="Sheet3")
