import pandas as pd
import numpy as np
from pandas import ExcelWriter, ExcelFile

df = pd.read_excel('toBeParsed.xlsx', sheet_name='Kas')
df = df[np.isfinite(df['NIK'])]  # dropping nonexistent NIKs
df['Masuk'].fillna(0, inplace=True)  # replace all NaN with 0
df['Keluar'].fillna(0, inplace=True)  # replace all NaN with 0


# Map<NIK, <Map<Tanggal, List<Map<Type, Value>>>> map;

nik_dict = dict()  # dictionary that contains all the transaction information
for index, row, in df.iterrows():
    curr_nik_dict = None
    if row['NIK'] not in nik_dict.keys() and row['Tanggal'] not in nik_dict.values():
        nik_dict[int(row['NIK'])] = {row['Tanggal']: []}
    elif row['Tanggal'] not in nik_dict[int(row['NIK'])].keys():
        nik_dict[int(row['NIK'])][row['Tanggal']] = [];

    curr_nik_dict = nik_dict[int(row['NIK'])][row['Tanggal']]
    jumlah = abs(row['Masuk'] + row['Keluar'])  # merge these columns to one
    curr_nik_dict.append({'Jumlah': jumlah})
    curr_nik_dict.append({'Transaksi': row['Transaksi']})


print(nik_dict)

# for NIK in NIK_column:
#     if NIK is not in dict: create a new excel file first,
#     else:
#         add or subtract from "saldo" (grouping them to dates)
