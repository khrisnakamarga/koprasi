import pandas as pd
import numpy as np
import xlsxwriter


class Transaksi:
    def __init__(self, jumlah, tipe):
        # TO DO: Privatize
        self.jumlah = jumlah
        self.tipe = tipe

    def __repr__(self):
        return "({} sejumlah Rp{})".format(self.tipe, self.jumlah)


def load_file(data_frame):
    # Map<NIK, <Map<Tanggal, List<Transaksi(jumlah, tipe)>>> map;

    nik_dict = dict()  # dictionary that contains all the transaction information
    for index, row, in data_frame.iterrows():
        if row['NIK'] not in nik_dict.keys() and row['Tanggal'] not in nik_dict.values():
            nik_dict[int(row['NIK'])] = {row['Tanggal']: []}
        elif row['Tanggal'] not in nik_dict[int(row['NIK'])].keys():
            nik_dict[int(row['NIK'])][row['Tanggal']] = []

        # Filters recorded transactions to only "Tarikan" or "Tabungan"
        curr_transaksi = nik_dict[int(row['NIK'])][row['Tanggal']]
        jumlah = abs(row['Masuk'] + row['Keluar'])  # merge these columns to one
        curr_transaksi.append(Transaksi(jumlah, row['Transaksi']))
    print(nik_dict)

    return nik_dict


def generate_report(nik_dict):
    workbook = xlsxwriter.Workbook('nik_report.xlsx')

    for nik in nik_dict.keys():
        curr_sheet = workbook.add_worksheet("NIK {}".format(nik))
        curr_sheet.write(0, 0, "Tanggal")
        curr_sheet.write(0, 1, "Jumlah Transaksi")
        curr_sheet.write(0, 2, "Tipe Transaksi")
        row = 0
        col = 0
        for date in nik_dict[nik].keys():
            item_list = nik_dict[nik][date]
            # print(item)
            for item in item_list:
                row += 1
                curr_sheet.write(row, col, date)
                curr_sheet.write(row, col + 1, item.jumlah)
                curr_sheet.write(row, col + 2, item.tipe)

    workbook.close()


if __name__ == "__main__":
    df = pd.read_excel('toBeParsed.xlsx', sheet_name='Kas')
    df = df[np.isfinite(df['NIK'])]  # dropping nonexistent NIKs
    df['Masuk'].fillna(0, inplace=True)  # replace all NaN with 0
    df['Keluar'].fillna(0, inplace=True)  # replace all NaN with 0

    nik_dict = load_file(df)
    # print(nik_dict)

    generate_report(nik_dict)
