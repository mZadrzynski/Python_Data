import pandas as pd
import numpy as np

#df = (pd.read_excel("OA - Porownanie_20220420_1220.xlsx", "Export Worksheet", usecols="G, H"))
df = (pd.read_excel("doors.xlsx", "Export Worksheet"))

#print/read rows
#print(df.head(5))

#printing first colums (samo wczytanie bez printa)
#print(df.columns)

#print/read column (pierwsze 5 - 0:5)
#print(df['FUNCDESCR_(newG)','columnn name...'][0:5]) 

#print specific location ([row, column])
#print(df.iloc[2, 6])

#wpisuje wszystkie wyrazy z komórki excela w tablice i porównuje je z drugą
for i in range(0, 1483):
    try:
        a = df.iloc[i, 10]
        b = df.iloc[i, 13]
        x = a.split()
        y = b.split()

        data_1 = {'TYPE': x}
        data_2 = {'TYPE': y} 
        if np.array_equal(data_1, data_2):
            df.iloc[i, 14] = 'jest ok'
    except:
            print(i)
            #df.iloc[i, 14] = 'jest zle'
    
df.to_excel("doors2.xlsx", sheet_name="Sheet3")
