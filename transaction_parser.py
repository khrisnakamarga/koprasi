import pandas as pd
import numpy as np
import xlsxwriter

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


# print(nik_dict)

workbook = xlsxwriter.Workbook('nik_report.xlsx')

for nik in nik_dict.keys():
    curr_sheet = workbook.add_worksheet("NIK {}".format(nik))
    curr_sheet.write(0, 0, "Tanggal")
    curr_sheet.write(0, 1, "Jumlah Transaksi")
    curr_sheet.write(0, 2, "Tipe Transaksi")
    row = 0
    col = 0
    for date in nik_dict[nik].keys():
        row += 1
        curr_sheet.write(row, col, date)
        item = nik_dict[nik][date]
        print(item)
        curr_sheet.write(row, col + 1, item[0]["Jumlah"])
        curr_sheet.write(row, col + 2, item[1]["Transaksi"])
        row += 1

workbook.close()
